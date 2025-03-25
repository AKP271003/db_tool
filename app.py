from flask import Flask, request, send_file, jsonify
import pymysql
import pandas as pd
import zipfile
from werkzeug.utils import secure_filename
from io import BytesIO
import re

app = Flask(__name__)

# Database connection settings
host = 'localhost'
user = 'root'
password = 'password'
port = 3307
database = 'HugeDatabase'

# Connect to the database
def connect_db():
    try:
        connection = pymysql.connect(
            host=host,
            user=user,
            password=password,
            port=port,
            database=database,
            charset='utf8mb4',
            collation='utf8mb4_general_ci'
        )
        return connection
    except pymysql.MySQLError as err:
        print(f"Error connecting to database: {err}")
        return None

# Execute SQL queries and return results as a DataFrame
def execute_queries(script, case_number):
    connection = connect_db()
    if connection is None:
        return None, ["failed to connect to the database"]
    
    results=[]
    messages=[]
    
    try:
        cursor = connection.cursor()
        
        # Replace ? with case_number in the script
        script = script.replace('?', str(case_number))
        
        statements = re.split(r';\s*', script)
        for statement in statements:
            statement = statement.strip()
            if not statement:
                continue

            # Check query execution plan
            explain_query = f"EXPLAIN {statement}"
            cursor.execute(explain_query)
            explain_result = cursor.fetchall()
            
            # Log the EXPLAIN result for debugging
            messages.append(f"EXPLAIN result for '{statement}': {explain_result}")
            
            # Check if query reads more than 10,000 rows
            for row in explain_result:
                try:
                    # Check if rows_examined is available and numeric
                    if len(row) > 9 and isinstance(row[9], int):
                        rows_examined = row[9]
                        messages.append(f"Rows examined for '{statement}': {rows_examined}")
                        if rows_examined > 10000:
                            messages.append(f"Query aborted: Reads more than 10,000 rows for '{statement}'.")
                            return results, messages
                    else:
                        messages.append(f"Error parsing rows examined for '{statement}': {row}")
                except Exception as e:
                    messages.append(f"Error parsing rows examined for '{statement}': {e}")
            
            try:
                cursor.execute(statement)
                result = cursor.fetchall()
                if result:
                    columns = [desc[0] for desc in cursor.description]
                    df = pd.DataFrame(result, columns=columns)
                    results.append(df)
                    messages.append(f"Query executed successfully with results for '{statement}'.")
                else:
                    messages.append(f"Query executed successfully with no results for '{statement}'.")
            except pymysql.MySQLError as err:
                messages.append(f"Error executing statement '{statement}': {err}")

    finally:
        connection.close()

    return results, messages


# API endpoint to upload queries and download results
@app.route('/execute_queries', methods=['POST'])
def handle_execute_queries():
    # Get the uploaded file and case_number
    file = request.files.get('file')
    case_number = request.form.get('case_number')
    if file is None or case_number is None:
        return jsonify({'error': 'Missing required parameters'}), 400
    
    filename = secure_filename(file.filename)
    
    # Read queries from the file
    script = file.read().decode('utf-8')
    
    results, messages = execute_queries(script, case_number)

    if not results:
        return jsonify({'messsage': 'No results', "execution_log": messages}), 200
    
    # Check if results are empty
    if all(df.empty for df in results):
        return jsonify({'messsage': 'No data in results', "execution_log": messages}), 200
    
    excel_buffer = BytesIO()

    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        for i, df in enumerate(results, 1):
            sheet_name = f'Result_{i}'
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    excel_buffer.seek(0)

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr('query_results.xlsx', excel_buffer.getvalue())

    zip_buffer.seek(0)

    return send_file(
        zip_buffer,
        mimetype='application/zip',
        as_attachment=True,
        download_name='query_results.zip'
    )

if __name__ == '__main__':
    app.run(debug=True)

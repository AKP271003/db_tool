from flask import Flask, request, send_file, jsonify
import pymysql
import pandas as pd
import zipfile
from werkzeug.utils import secure_filename
from io import BytesIO, StringIO
import re
import os
import mysql.connector.pooling

app = Flask(__name__)

# Database connection settings
host = 'localhost'
user = 'root'
password = 'password'
port = 3307
database = 'HugeDatabase'

# Create a connection pool
pool = mysql.connector.pooling.MySQLConnectionPool(
    pool_name="my_pool",
    pool_size=5,
    host=host,
    user=user,
    password=password,
    database=database,
    port=port
)

QUERIES = """
SELECT * FROM huge_table WHERE case_number = ?;
SELECT * FROM huge_table WHERE case_number < ?;
SELECT * FROM huge_table WHERE case_number > ?;
"""

# Connect to the database
def connect_db():
    try:
        connection = pool.get_connection()
        return connection
    except Exception as err:
        print(f"Error connecting to database: {err}")
        return None

# Execute SQL queries and return results as a DataFrame
def execute_queries(case_number):
    connection = connect_db()
    if connection is None:
        return None, ["failed to connect to the database"]
    
    results=[]
    execution_log=[]
    ROW_LIMIT = 10000
    
    try:
        cursor = connection.cursor()
        case_number_str = str(case_number)
        script = QUERIES.replace('?', str(case_number))
        statements = re.split(r';\s*', script)

        for statement in statements:
            statement = statement.strip()
            if not statement:
                continue

            query_result = None
            query_error = None
            explain_rows = None

            try:

                # Check query execution plan
                explain_query = f"EXPLAIN {statement}"
                execution_log.append(f"\nAttempting EXPLAIN with: {explain_query}")
                cursor.execute(explain_query)
                explain_result = cursor.fetchall()
                execution_log.append(f"Raw EXPLAIN result: {explain_result}")
            
                if not explain_result:
                    raise ValueError("Empty EXPLAIN result")
                
                explain_rows=0
                for row in explain_result:
                    if len(row) > 8:
                        try:
                            row_estimate = int(row[8]) if row[8] is not None else 0
                            explain_rows += row_estimate
                        except (ValueError, TypeError) as e:
                            execution_log.append(f"Warning: Could not convert row estimate '{row[8]}' to number")
                            explain_rows = float('inf')
                
                execution_log.append(f"EXPLAIN for '{statement}': Estimated rows = {explain_rows}")

                if explain_rows > ROW_LIMIT:
                    query_error = f"Query aborted: Estimated rows ({explain_rows}) > {ROW_LIMIT}"
                    execution_log.append(query_error)
                    results.append({"error": query_error, "query":statement})
                    continue
            except Exception as e:
                query_error = f"EXPLAIN failed for '{statement}': {str(e)}"
                execution_log.append(query_error)
                results.append({"error": query_error, "query": statement})
                continue
            try:
                execution_log.append(f"Executing: {statement}")
                cursor.execute(statement)
                result = cursor.fetchall()

                if result:
                    columns = [desc[0] for desc in cursor.description]
                    df = pd.DataFrame(result, columns=columns)
                    actual_rows = len(df)
                    execution_log.append(f"Actual rows returned: {actual_rows}")

                    results.append({"data": df, "query": statement})
                    execution_log.append(f"Executed successfully: '{statement}'")
                else:
                    results.append({"data": None, "query": statement})
                    execution_log.append(f"No results for: '{statement}'")
            except Exception as e:
                query_error = f"Query execution failed for '{statement}': {str(e)}"
                execution_log.append(query_error)
                results.append({"error": query_error, "query": statement})
    
    finally:
            connection.close()
    
    return results, execution_log

# API endpoint to upload queries and download results
@app.route('/execute_queries', methods=['POST'])
def handle_execute_queries():
    if 'case_number' not in request.form:
        return jsonify({'error': 'Missing required parameter case_number'}), 400

    case_number = request.form['case_number']

    try:
        results, messages = execute_queries(case_number)

        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for i, result in enumerate(results, 1):
                filename = f'query_{i}'

                if 'data' in result and result['data'] is not None:
                    csv_buffer = StringIO()
                    result['data'].to_csv(csv_buffer, index=False)
                    zip_file.writestr(f'{filename}.csv', csv_buffer.getvalue())
                elif 'error' in result:
                    error_content = f"Error: {result['error']}\nQuery: {result.get('query', '')[:500]}"
                    zip_file.writestr(f'{filename}_error.txt', error_content)
                else:
                    zip_file.writestr(f'{filename}_empty.txt', "No results for this query")

            zip_file.writestr('execution_log.txt', "\n".join(messages))

        zip_buffer.seek(0)

        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name=f'query_results_{case_number}.zip'
        )
    except Exception as e:
        return jsonify({
            'error': f'Processing failed: {str(e)}',
            'execution_log': messages if 'messages' in locals() else[]
        }), 500

if __name__ == '__main__':
    app.run(debug=True)

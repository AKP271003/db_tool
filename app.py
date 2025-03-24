from flask import Flask, request, send_file
import mysql.connector
import pandas as pd
import zipfile
from werkzeug.utils import secure_filename

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
        connection = mysql.connector.connect(
            host=host,
            user=user,
            password=password,
            port=port,
            database=database,
            charset='utf8mb4',
            collation='utf8mb4_general_ci'
        )
        return connection
    except mysql.connector.Error as err:
        print(f"Error connecting to database: {err}")
        return None

# Execute SQL queries and return results as a DataFrame
def execute_query(query):
    connection = connect_db()
    if connection is None:
        return None
    
    try:
        cursor = connection.cursor()
        
        # Check if query is a procedure creation
        if query.upper().startswith('CREATE PROCEDURE'):
            cursor.execute(query, multi=True)
            connection.commit()
            connection.close()
            return None  # No results for procedure creation
        
        # Check if query is a procedure call
        elif query.upper().startswith('CALL'):
            cursor.execute(query)
            results = cursor.fetchall()
            columns = [desc[0] for desc in cursor.description]
            df = pd.DataFrame(results, columns=columns)
            connection.close()
            return df
        else:
            df = pd.read_sql_query(query, connection)
            connection.close()
            return df
    except mysql.connector.Error as err:
        print(f"Error executing query: {err}")
        connection.close()
        return None

# Save DataFrame to Excel and zip it
def save_and_zip(df, filename):
    df.to_excel(filename, index=False)
    
    zip_filename = f"{filename}.zip"
    with zipfile.ZipFile(zip_filename, 'w') as zip_file:
        zip_file.write(filename)
    
    return zip_filename

# API endpoint to upload queries and download results
@app.route('/execute_queries', methods=['POST'])
def execute_queries():
    # Get the uploaded file
    file = request.files['file']
    filename = secure_filename(file.filename)
    
    # Read queries from the file
    queries = file.read().decode('utf-8')
    
    # Split queries based on DELIMITER
    delimiter = '//'
    queries_list = queries.split(delimiter)
    
    # Process each query block
    results = []
    
    for query_block in queries_list:
        query_block = query_block.strip()
        if query_block:
            # Check if it's a procedure creation
            if query_block.upper().startswith('CREATE PROCEDURE'):
                # Execute procedure creation directly
                connection = connect_db()
                cursor = connection.cursor()
                cursor.execute(query_block, multi=True)
                connection.commit()
                connection.close()
            else:
                # Split by semicolon and execute each query
                for query in query_block.split(';'):
                    query = query.strip()
                    if query:
                        df = execute_query(query)
                        if df is not None:
                            results.append(df)
    
    if results:
        combined_result = pd.concat(results, ignore_index=True)
        
        # Save to Excel and zip
        excel_filename = 'query_results.xlsx'
        zip_filename = save_and_zip(combined_result, excel_filename)
        
        # Return the zip file for download
        return send_file(zip_filename, as_attachment=True)
    else:
        return "No results to download."

if __name__ == '__main__':
    app.run(debug=True)

'''import mysql.connector

# Connect to MariaDB with explicit charset and collation
try:
    connection = mysql.connector.connect(
        host='localhost',
        user='root',
        password='password',
        port=3307,
        charset='utf8mb4',  # Force MariaDB-compatible charset
        collation='utf8mb4_general_ci'  # Force compatible collation
    )
    cursor = connection.cursor()

    # Create a new database with a compatible collation
    cursor.execute("CREATE DATABASE IF NOT EXISTS HugeDatabase CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci")
    print("Database 'HugeDatabase' created successfully.")

    # Create a table with 90 columns, changing some to TEXT
    columns = ['case_number INT PRIMARY KEY']
    for i in range(1, 91):
        if i <= 10:  # Keep first 10 columns as VARCHAR
            columns.append(f'column_{i} VARCHAR(255)')
        else:  # Change remaining columns to TEXT
            columns.append(f'column_{i} TEXT')

    cursor.execute('USE HugeDatabase')

    create_table_query = f"CREATE TABLE IF NOT EXISTS huge_table ({', '.join(columns)})"
    cursor.execute(create_table_query)
    print("Table 'huge_table' created successfully.")

    cursor.close()
    connection.close()
except mysql.connector.Error as err:
    print(f"Error: {err}")
'''
import mysql.connector


print("Starting script...")  # Check if the script starts executing

try:
    connection = mysql.connector.connect(
        host='localhost',
        user='root',
        password='password',
        port=3307,
        database='HugeDatabase',
        charset='utf8mb4',
        collation='utf8mb4_general_ci'
    )
    print("Connected to database successfully.")  # Ensure connection is established
except Exception as e:
    print(e)

import numpy as np
import pandas as pd
import mysql.connector

num_records = 1000000  # Total records to insert
chunksize = 10000  # Insert in chunks of 1000 rows

try:
    connection = mysql.connector.connect(
        host='localhost',
        user='root',
        password='password',
        port=3307,
        database='HugeDatabase',
        charset='utf8mb4',
        collation='utf8mb4_general_ci'
    )
    print("Connected to database successfully.")
    
    connection.autocommit = True  # Ensure auto-commit is enabled
    cursor = connection.cursor()

    # Check if table exists
    cursor.execute("SHOW TABLES LIKE 'huge_table'")
    if cursor.fetchone() is None:
        print("Table 'huge_table' does not exist. Creating it...")
        
        # Create table if it doesn't exist
        columns = ['case_number INT PRIMARY KEY']
        for i in range(1, 91):
            if i <= 10:  # Keep first 10 columns as VARCHAR
                columns.append(f'column_{i} VARCHAR(255)')
            else:  # Change remaining columns to TEXT
                columns.append(f'column_{i} TEXT')

        create_table_query = f"CREATE TABLE huge_table ({', '.join(columns)})"
        cursor.execute(create_table_query)
        print("Table 'huge_table' created successfully.")

    for i in range(0, num_records, chunksize):
        print(f"Generating chunk {i // chunksize + 1}...")
        
        chunk_data = {
            'case_number': np.arange(i + 1, min(i + chunksize + 1, num_records + 1)),
        }
        for j in range(1, 91):
            chunk_data[f'column_{j}'] = np.random.choice(
                ['value1', 'value2', 'value3', ''],  # Avoid NULL values
                min(chunksize, num_records - i)
            )

        chunk_df = pd.DataFrame(chunk_data)

        insert_query = "INSERT INTO huge_table ({}) VALUES ({})".format(
            ', '.join(chunk_df.columns),
            ', '.join(['%s'] * len(chunk_df.columns))
        )

        print(f"Inserting chunk {i // chunksize + 1}...")
        print(insert_query)  # Debugging output
        print(chunk_df.values.tolist()[0])  # Print first row for validation

        try:
            cursor.executemany(insert_query, chunk_df.values.tolist())
            connection.commit()
            print(f"Inserted chunk {i // chunksize + 1} of {num_records // chunksize + 1}")
        except mysql.connector.Error as err:
            print(f"Error inserting chunk: {err}")

    cursor.close()
    connection.close()
except mysql.connector.Error as err:
    print(f"Database connection error: {err}")
except Exception as err:
    print(f"Unexpected error: {err}")

import pyodbc
import os
from dotenv import load_dotenv

load_dotenv()

def test_connection():
    print("Available ODBC Drivers:")
    for d in pyodbc.drivers():
        print(f" - {d}")
    print("-" * 30)

    driver = os.getenv("DB_DRIVER", "ODBC Driver 17 for SQL Server")
    server = os.getenv("DB_SERVER", "localhost")
    port = os.getenv("DB_PORT")
    database = os.getenv("DB_NAME")
    username = os.getenv("DB_USER")
    password = os.getenv("DB_PASSWORD")

    print(f"Testing connection to {server} / {database} as {username}...")

    # First try connecting to master to verify credentials
    conn_str_master = f"DRIVER={{{driver}}};SERVER={server};"
    if port:
        conn_str_master += f"PORT={port};"
    conn_str_master += f"DATABASE=master;UID={username};PWD={password};Encrypt=no;"

    print(f"Testing connection to {server} / master as {username}...")
    try:
        conn = pyodbc.connect(conn_str_master, timeout=10)
        print("Successfully connected to master!")
        
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sys.databases WHERE name = ?", database)
        row = cursor.fetchone()
        if row:
            print(f"Database '{database}' exists.")
        else:
            print(f"Database '{database}' DOES NOT EXIST.")

        cursor.execute("SELECT @@VERSION")
        row = cursor.fetchone()
        print(f"SQL Server Version: {row[0]}")
        
        conn.close()
    except Exception as e:
        print(f"Connection to master failed: {str(e)}")
        return

    # Now try connecting to target database
    conn_str = f"DRIVER={{{driver}}};SERVER={server};"
    if port:
        conn_str += f"PORT={port};"
    conn_str += f"DATABASE={database};UID={username};PWD={password};Encrypt=no;"

    print(f"Testing connection to {server} / {database} as {username}...")

    try:
        conn = pyodbc.connect(conn_str, timeout=10)
        print(f"Successfully connected to {database}!")
        
        cursor = conn.cursor()
        cursor.execute("SELECT count(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE'")
        row = cursor.fetchone()
        print(f"Table count: {row[0]}")
        
        conn.close()
    except Exception as e:
        print(f"Connection to {database} failed: {str(e)}")

if __name__ == "__main__":
    test_connection()
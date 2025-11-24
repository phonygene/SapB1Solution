import pyodbc
import os
import re
from dotenv import load_dotenv

# Load configuration from mcp-sqlserver/.env
env_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'mcp-sqlserver', '.env')
load_dotenv(env_path)

def get_connection(database="master"):
    driver = os.getenv("DB_DRIVER", "ODBC Driver 17 for SQL Server")
    server = os.getenv("DB_SERVER", "localhost")
    user = os.getenv("DB_USER")
    password = os.getenv("DB_PASSWORD")
    
    # Strip quotes if present
    if server.startswith('"') and server.endswith('"'):
        server = server[1:-1]
        
    conn_str = f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};UID={user};PWD={password};Encrypt=no;"
    return pyodbc.connect(conn_str, timeout=10)

def execute_sql_file(file_path):
    print(f"Processing {os.path.basename(file_path)}...")
    
    with open(file_path, 'r', encoding='utf-8') as f:
        sql_content = f.read()

    # Determine initial database context
    target_db = "master"
    if "jtdb" in os.path.basename(file_path) and "CreateDatabase" not in os.path.basename(file_path):
         target_db = "jtdb"
    if "CreateTable" in os.path.basename(file_path) or "AlterTable" in os.path.basename(file_path):
        target_db = "jtdb"
    if "MDR" in os.path.basename(file_path) and "CreateDatabase" not in os.path.basename(file_path):
        target_db = "MDR"
        
    print(f"Connecting to {target_db}...")
    
    try:
        conn = get_connection(target_db)
        cursor = conn.cursor()
        
        # Split by GO command
        batches = re.split(r'^\s*GO\s*$', sql_content, flags=re.MULTILINE | re.IGNORECASE)
        
        for batch in batches:
            if batch.strip():
                # Remove USE statements to avoid context switching issues if we are already connected to target
                # batch = re.sub(r'USE\s+\w+', '', batch, flags=re.IGNORECASE)
                
                try:
                    cursor.execute(batch)
                    conn.commit()
                except Exception as e:
                    print(f"Error executing batch in {os.path.basename(file_path)}: {e}")
                    # print(f"Batch content: {batch[:100]}...")
                    
        print(f"Successfully processed {os.path.basename(file_path)}")
        conn.close()
    except Exception as e:
        print(f"Failed to process {os.path.basename(file_path)}: {e}")

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    sql_dir = os.path.join(base_dir, "SqlQuery")
    
    scripts = [
        "00_CreateDatabase_jtdb.sql",
        "05_CreateTable_addr.sql",
        "06_CreateTable_expense_category.sql",
        "07_AlterTable_jOPCH_Add_ApprovalComments.sql",
        "08_AlterTable_User_Add_CanApproveExpense.sql",
        "10_CreateDatabase_MDR_Local.sql"
    ]
    
    for script in scripts:
        full_path = os.path.join(sql_dir, script)
        if os.path.exists(full_path):
            execute_sql_file(full_path)
        else:
            print(f"File not found: {full_path}")

if __name__ == "__main__":
    main()
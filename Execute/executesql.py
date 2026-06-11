import pyodbc

def get_connection():
    try:
        conn = pyodbc.connect(
            
        )
        print("Connection Successful")
        return conn
    except Exception as e:
        print("Database connection failed:", e)
        return None
    
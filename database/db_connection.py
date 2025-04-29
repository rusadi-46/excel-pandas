import mysql.connector
from mysql.connector import Error

def create_connection():
    try:
        connection = mysql.connector.connect(
            host='localhost',
            user='root',
            password='',
            database='tstyle_salon'
        )
        if connection.is_connected():
            return connection
    except Error as e:
        print(f"Koneksi gagal: {e}")
        return None



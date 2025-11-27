import psycopg2
from psycopg2 import extras
import pandas as pd
from datetime import datetime

print("Pandas imported successfully")

def get_db_connection_debug():
    try:
        conn = psycopg2.connect(
            dbname="aplus_com_test",
            user="app_user",
            password="cailfornia123",
            host="192.168.1.51",
            port="5432",
            connect_timeout=5
        )
        return conn
    except Exception as e:
        print(f"Connection Error: {e}")
        return None

conn = get_db_connection_debug()
if not conn:
    print("Failed to connect to database")
    exit()

try:
    with conn.cursor() as cursor:
        print("\n--- Employees ---")
        cursor.execute("SELECT emp_id, fname, lname, status FROM employees LIMIT 5")
        rows = cursor.fetchall()
        for row in rows:
            print(row)

        print("\n--- Time Attendance Logs (Min/Max Date) ---")
        cursor.execute("SELECT MIN(scan_timestamp), MAX(scan_timestamp) FROM time_attendance_logs")
        min_ts, max_ts = cursor.fetchone()
        print(f"Min: {min_ts}, Max: {max_ts}")

        print("\n--- Recent Logs (Top 5) ---")
        cursor.execute("SELECT * FROM time_attendance_logs ORDER BY scan_timestamp DESC LIMIT 5")
        for row in cursor.fetchall():
            print(row)
finally:
    if conn:
        conn.close()

import mysql.connector

conn = mysql.connector.connect(
    host="localhost",
    user="root",
    password=""
)

cur = conn.cursor()
cur.execute("CREATE DATABASE IF NOT EXISTS exam_timetable_db")

print("Database created")

cur.close()
conn.close()

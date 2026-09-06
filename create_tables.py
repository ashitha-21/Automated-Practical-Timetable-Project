import mysql.connector

conn = mysql.connector.connect(
    host="localhost",
    user="root",
    password="",
    database="exam_timetable_db"
)

cur = conn.cursor()

# ---------------- STUDENTS ----------------
cur.execute("""
CREATE TABLE IF NOT EXISTS students(
id INT AUTO_INCREMENT PRIMARY KEY,
regno VARCHAR(20) UNIQUE,
class VARCHAR(10),
sem VARCHAR(10)
)
""")

# ---------------- SUBJECTS ----------------
cur.execute("""
CREATE TABLE IF NOT EXISTS subjects(
id INT AUTO_INCREMENT PRIMARY KEY,
class VARCHAR(5),
subject_order INT,
subject_code VARCHAR(20) UNIQUE,
subject_name VARCHAR(100)
)
""")

# ---------------- LABS ----------------
cur.execute("""
CREATE TABLE IF NOT EXISTS labs(
id INT AUTO_INCREMENT PRIMARY KEY,
lab_name VARCHAR(50) UNIQUE,
capacity INT
)
""")

# ---------------- EXAM DAYS ----------------
cur.execute("""
CREATE TABLE IF NOT EXISTS exam_days(
id INT AUTO_INCREMENT PRIMARY KEY,
day_no INT,
exam_date DATE
)
""")

# ---------------- SESSIONS ----------------
cur.execute("""
CREATE TABLE IF NOT EXISTS sessions(
id INT AUTO_INCREMENT PRIMARY KEY,
session_no INT,
session_time VARCHAR(30)
)
""")

# ---------------- TIMETABLE ----------------
cur.execute("""
CREATE TABLE IF NOT EXISTS timetable(
id INT AUTO_INCREMENT PRIMARY KEY,
batch_id VARCHAR(20),

regno_from VARCHAR(20),
regno_to VARCHAR(20),

subject_id INT,
lab_id INT,
day_id INT,
session_id INT,

FOREIGN KEY (subject_id) REFERENCES subjects(id) ON DELETE CASCADE,
FOREIGN KEY (lab_id) REFERENCES labs(id) ON DELETE CASCADE,
FOREIGN KEY (day_id) REFERENCES exam_days(id) ON DELETE CASCADE,
FOREIGN KEY (session_id) REFERENCES sessions(id) ON DELETE CASCADE
)
""")

# ---------------- STAFF ALLOCATION ----------------
cur.execute("""
CREATE TABLE IF NOT EXISTS staff_allocation(
id INT AUTO_INCREMENT PRIMARY KEY,
timetable_id INT,
internal_staff VARCHAR(100),
external_staff VARCHAR(100),

FOREIGN KEY (timetable_id) REFERENCES timetable(id) ON DELETE CASCADE
)
""")

# ---------------- RESET TIMETABLE DATA ----------------
# Whenever new timetable is generated, old timetable is cleared

cur.execute("SET FOREIGN_KEY_CHECKS=0")

cur.execute("TRUNCATE TABLE staff_allocation")
cur.execute("TRUNCATE TABLE timetable")
cur.execute("TRUNCATE TABLE exam_days")
cur.execute("TRUNCATE TABLE sessions")

cur.execute("SET FOREIGN_KEY_CHECKS=1")

conn.commit()
cur.close()
conn.close()

print("Tables created successfully and timetable data reset")
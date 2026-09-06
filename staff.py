from flask import Blueprint, render_template, request, redirect
from db import get_db_connection
from openpyxl import Workbook
from flask import send_file

staff_bp = Blueprint("staff", __name__)

INTERNAL_STAFF_OPTIONS = [
    "Mrs Premalatha R Shetty",
    "Dr Ravindra Swami K",
    "Mr Naveen Mascarenhas",
    "Mr Ashok M Prasad",
    "Ms Renita Caroline Menezes",
    "Ms Vanaja A",
    "Dr Vidya Kumari",
    "Ms Rashmi",
    "Ms Vinaya Durga M",
    "Dr Royal Praveen Dsouza",
    "Dr Archana Yashodhar",
    "Ms Achala Nagesh B",
    "Ms Himani B S",
    "Mr Harshith S Kottary",
    "Ms Rishal Sharal Noronha",
    "Ms Bhoomika A",
    "Ms Jasmine Maria Serrao",
    "Ms Jyothi Priya Dsouza",
    "Ms Bindya",
    "Ms Anne Neha Vaz",
    "Dr Priyadarshini P"
]


@staff_bp.route("/staff_menu")
def staff_menu():
    saved = request.args.get("saved")
    error = request.args.get("error")
    return render_template("staff_menu.html", saved=saved, error=error)

@staff_bp.route("/internal_staff")
def internal_staff():

    error = request.args.get("error")

    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)

    cur.execute("""
        SELECT 
        t.id,
        t.batch_id,
        l.lab_name,
        se.session_time AS exam_time,
        s.subject_name AS subject,
        d.exam_date
        FROM timetable t
        JOIN subjects s ON t.subject_id = s.id
        JOIN labs l ON t.lab_id = l.id
        JOIN exam_days d ON t.day_id = d.id
        JOIN sessions se ON t.session_id = se.id
        ORDER BY 
        exam_date,
        CASE 
            WHEN exam_time = '09.00-12.00' THEN 1
            WHEN exam_time = '01.00-04.00' THEN 2
            ELSE 3
        END,
        lab_name, t.batch_id
    """)

    batches = cur.fetchall()

    # already saved internal staff if exists
    for row in batches:
        cur.execute("""
            SELECT internal_staff
            FROM staff_allocation
            WHERE timetable_id = %s
        """, (row["id"],))
        existing = cur.fetchone()
        row["selected_staff"] = existing["internal_staff"] if existing and existing["internal_staff"] else ""

    cur.close()
    conn.close()

    return render_template(
        "staff_form.html",
        batches=batches,
        staff_type="INTERNAL",
        internal_staff_options=INTERNAL_STAFF_OPTIONS,
        error=error
    )

@staff_bp.route("/external_staff")
def external_staff():

    error = request.args.get("error")

    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)

    cur.execute("""
        SELECT 
        t.id,
        t.batch_id,
        l.lab_name,
        se.session_time AS exam_time,
        s.subject_name AS subject,
        d.exam_date
        FROM timetable t
        JOIN subjects s ON t.subject_id = s.id
        JOIN labs l ON t.lab_id = l.id
        JOIN exam_days d ON t.day_id = d.id
        JOIN sessions se ON t.session_id = se.id
        ORDER BY 
        exam_date,
        CASE 
            WHEN exam_time = '09.00-12.00' THEN 1
            WHEN exam_time = '01.00-04.00' THEN 2
            ELSE 3
        END,
        lab_name, t.batch_id
    """)

    batches = cur.fetchall()

    # already saved external staff if exists
    for row in batches:
        cur.execute("""
            SELECT external_staff
            FROM staff_allocation
            WHERE timetable_id = %s
        """, (row["id"],))
        existing = cur.fetchone()
        row["selected_staff"] = existing["external_staff"] if existing and existing["external_staff"] else ""

    cur.close()
    conn.close()

    return render_template(
        "staff_form.html",
        batches=batches,
        staff_type="EXTERNAL",
        error=error
    )

@staff_bp.route("/save_staff", methods=["POST"])
def save_staff():

    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)

    staff_type = request.form["staff_type"]

    # read timetable details
    cur.execute("""
        SELECT 
        t.id,
        t.batch_id,
        s.subject_name AS subject,
        l.lab_name,
        d.exam_date,
        se.session_time AS exam_time
        FROM timetable t
        JOIN subjects s ON t.subject_id = s.id
        JOIN labs l ON t.lab_id = l.id
        JOIN exam_days d ON t.day_id = d.id
        JOIN sessions se ON t.session_id = se.id
        ORDER BY 
        exam_date,
        CASE 
            WHEN exam_time = '09.00-12.00' THEN 1
            WHEN exam_time = '01.00-04.00' THEN 2
            ELSE 3
        END,
        lab_name, t.batch_id
    """)
    timetable = cur.fetchall()

    if staff_type == "INTERNAL":
        used_in_this_submission = {}

        for row in timetable:
            timetable_id = row["id"]
            exam_date = row["exam_date"]
            exam_time = row["exam_time"]

            staff_name = request.form.get(f"staff_{timetable_id}")

            if not staff_name:
                continue

            slot_key = f"{exam_date}_{exam_time}"

            if slot_key not in used_in_this_submission:
                used_in_this_submission[slot_key] = set()

            # duplicate inside same form submission
            if staff_name in used_in_this_submission[slot_key]:
                cur.close()
                conn.close()
                return redirect(
                    f"/internal_staff?error=Same internal staff cannot be assigned twice for {exam_date} {exam_time}"
                )

            used_in_this_submission[slot_key].add(staff_name)

        # duplicate against already saved rows in DB
        for row in timetable:
            timetable_id = row["id"]
            exam_date = row["exam_date"]
            exam_time = row["exam_time"]

            staff_name = request.form.get(f"staff_{timetable_id}")

            if not staff_name:
                continue

            cur.execute("""
                SELECT sa.internal_staff, t.id
                FROM staff_allocation sa
                JOIN timetable t ON sa.timetable_id = t.id
                JOIN exam_days d ON t.day_id = d.id
                JOIN sessions se ON t.session_id = se.id
                WHERE sa.internal_staff = %s
                  AND d.exam_date = %s
                  AND se.session_time = %s
                  AND t.id != %s
            """, (staff_name, exam_date, exam_time, timetable_id))

            conflict = cur.fetchone()

            if conflict:
                cur.close()
                conn.close()
                return redirect(
                    f"/internal_staff?error={staff_name} is already assigned on {exam_date} at {exam_time}"
                )

    # ---------------- SAVE DATA ----------------
    for row in timetable:
        timetable_id = row["id"]

        staff_name = request.form.get(f"staff_{timetable_id}")

        if not staff_name:
            continue

        cur.execute(
            "SELECT id FROM staff_allocation WHERE timetable_id=%s",
            (timetable_id,)
        )
        existing = cur.fetchone()

        if existing:
            if staff_type == "INTERNAL":
                cur.execute("""
                    UPDATE staff_allocation
                    SET internal_staff=%s
                    WHERE timetable_id=%s
                """, (staff_name, timetable_id))
            else:
                cur.execute("""
                    UPDATE staff_allocation
                    SET external_staff=%s
                    WHERE timetable_id=%s
                """, (staff_name, timetable_id))
        else:
            cur.execute("""
                INSERT INTO staff_allocation
                (timetable_id, internal_staff, external_staff)
                VALUES (%s, %s, %s)
            """, (
                timetable_id,
                staff_name if staff_type == "INTERNAL" else None,
                staff_name if staff_type == "EXTERNAL" else None
            ))

    conn.commit()
    cur.close()
    conn.close()

    return redirect("/staff_menu?saved=1")


@staff_bp.route("/view_staff")
def view_staff():

    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)

    cur.execute("""
    SELECT 
    t.batch_id,
    s.subject_name AS subject,
    l.lab_name,
    d.exam_date,
    se.session_time,
    sa.internal_staff,
    sa.external_staff
    FROM staff_allocation sa
    JOIN timetable t ON sa.timetable_id = t.id
    JOIN subjects s ON t.subject_id = s.id
    JOIN labs l ON t.lab_id = l.id
    JOIN exam_days d ON t.day_id = d.id
    JOIN sessions se ON t.session_id = se.id
    ORDER BY
        d.exam_date,
        CASE
            WHEN se.session_time = '09.00-12.00' THEN 1
            WHEN se.session_time = '01.00-04.00' THEN 2
            ELSE 3
        END,
        l.lab_name
    """)

    data = cur.fetchall()

    cur.close()
    conn.close()

    return render_template("view_staff.html", staff=data)


@staff_bp.route("/export_staff_excel")
def export_staff_excel():

    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)

    cur.execute("""
    SELECT 
    t.batch_id,
    s.subject_name AS subject,
    l.lab_name,
    d.exam_date,
    se.session_time,
    sa.internal_staff,
    sa.external_staff
    FROM staff_allocation sa
    JOIN timetable t ON sa.timetable_id = t.id
    JOIN subjects s ON t.subject_id = s.id
    JOIN labs l ON t.lab_id = l.id
    JOIN exam_days d ON t.day_id = d.id
    JOIN sessions se ON t.session_id = se.id
    ORDER BY
        d.exam_date,
        CASE
            WHEN se.session_time = '09.00-12.00' THEN 1
            WHEN se.session_time = '01.00-04.00' THEN 2
            ELSE 3
        END,
        l.lab_name
    """)

    rows = cur.fetchall()

    cur.close()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Staff Allocation"

    ws.append([
        "Batch",
        "Subject",
        "Lab",
        "Exam Date",
        "Time",
        "Internal Staff",
        "External Staff"
    ])

    for r in rows:
        ws.append([
            r["batch_id"],
            r["subject"],
            r["lab_name"],
            str(r["exam_date"]),
            r["session_time"],
            r["internal_staff"],
            r["external_staff"]
        ])

    file_path = "staff_allocation.xlsx"
    wb.save(file_path)

    return send_file(file_path, as_attachment=True)
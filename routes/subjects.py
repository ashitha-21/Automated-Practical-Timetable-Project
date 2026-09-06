from flask import Blueprint, request, jsonify
from db import get_db_connection

subjects_bp = Blueprint("subjects", __name__)

@subjects_bp.route("/add_subjects", methods=["POST"])
def add_subjects():

    data = request.json
    subjects = data["subjects"]

    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("DELETE FROM subjects")
    conn.commit() 

    order_no = 1   # subject numbering

    year_map = {
        "ty": "TY",
        "sy": "SY",
        "fy": "FY"
    }

    for year, subs in subjects.items():

        subject_order = 1   #  reset for each year

        for s in subs:
            cur.execute("""
                INSERT INTO subjects
                (class, subject_order, subject_code, subject_name)
                VALUES (%s,%s,%s,%s)
            """, (
                year_map[year],
                subject_order,
                s["code"],
                s["name"]
            ))

            subject_order += 1

            order_no += 1

    conn.commit()
    cur.close()
    conn.close()

    return jsonify({"message": "Subjects saved"})

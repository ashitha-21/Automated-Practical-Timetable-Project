from flask import Blueprint, request, jsonify
from datetime import timedelta
import pandas as pd
from db import get_db_connection

step1_bp = Blueprint("step1", __name__)

@step1_bp.route("/save_step1", methods=["POST"])
def save_step1():

    data = request.get_json()

    conn = get_db_connection()
    cur = conn.cursor()

    # clear old
    cur.execute("DELETE FROM exam_days")
    cur.execute("DELETE FROM sessions")

    start = data["start_date"]
    end = data["end_date"]
    from datetime import datetime, date

    start_date = datetime.strptime(start, "%Y-%m-%d").date()
    end_date = datetime.strptime(end, "%Y-%m-%d").date()

    today = date.today()

    #  block today & past dates
    if start_date <= today or end_date <= today:
        return jsonify({
            "message": "Start and End date must be after today"
        }), 400
    excluded = data.get("excluded_dates", [])

    d = pd.to_datetime(start)
    end_d = pd.to_datetime(end)

    day_no = 1

    while d <= end_d:
        if str(d.date()) not in excluded:
            cur.execute(
                "INSERT INTO exam_days (day_no, exam_date) VALUES (%s,%s)",
                (day_no, d.date())
            )
            day_no += 1

        d += timedelta(days=1)

    # SAVE SESSIONS
    sessions = data.get("sessions", [])

    session_no = 1
    for s in sessions:
        cur.execute(
            "INSERT INTO sessions(session_no, session_time) VALUES (%s,%s)",
            (session_no, s)
        )
        session_no += 1

    conn.commit()
    cur.close()
    conn.close()

    return jsonify({"message": "Step 1 saved"})

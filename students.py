from flask import Blueprint, jsonify, request
import pandas as pd
from db import get_db_connection

students_bp = Blueprint("students", __name__)


@students_bp.route("/upload_students", methods=["POST"])
def upload_students():

    try:

        # -----------------------------
        # GET SEMESTER
        # -----------------------------
        semester = request.form.get("semester", "Odd")

        # -----------------------------
        # GET EXCEL FILE
        # -----------------------------
        file = request.files.get("file")

        if not file:
            return jsonify({
                "message": "No Excel file uploaded"
            }), 400

        # -----------------------------
        # READ EXCEL
        # -----------------------------
        df = pd.read_excel(file)

        if df.empty:
            return jsonify({
                "message": "Excel file is empty"
            }), 400

        # -----------------------------
        # NORMALIZE COLUMN NAMES
        # -----------------------------
        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
            .str.lower()
            .str.replace("_", " ", regex=False)
        )

        print("Excel columns:", list(df.columns))

        # -----------------------------
        # FIND REGISTER NUMBER COLUMN
        # -----------------------------
        possible_names = [
            "regno",
            "reg no",
            "register no",
            "register number",
            "registration no",
            "registration number",
            "reg number",
            "register_no",
            "registration_no"
        ]

        regno_column = None

        for col in df.columns:

            normalized_col = (
                col.strip()
                .lower()
                .replace("_", " ")
            )

            if normalized_col in possible_names:
                regno_column = col
                break

        if regno_column is None:

            return jsonify({
                "message":
                "Excel must contain a Register Number column. "
                "Example: Register Number"
            }), 400

        print("Register number column:", regno_column)

        # -----------------------------
        # DATABASE CONNECTION
        # -----------------------------
        conn = get_db_connection()
        cur = conn.cursor()

        # Remove old students
        cur.execute("DELETE FROM students")

        inserted = 0

        # -----------------------------
        # INSERT STUDENTS
        # -----------------------------
        for _, row in df.iterrows():

            value = row[regno_column]

            if pd.isna(value):
                continue

            # Avoid values like 2351001.0
            if isinstance(value, float) and value.is_integer():
                regno = str(int(value))
            else:
                regno = str(value).strip()

            if not regno:
                continue

            # -----------------------------
            # DETERMINE CLASS
            # -----------------------------
            if regno.startswith("25"):
                cls = "FY"

            elif regno.startswith("24"):
                cls = "SY"

            elif regno.startswith("23"):
                cls = "TY"

            else:
                cls = "UNKNOWN"

            cur.execute(
                """
                INSERT INTO students
                (regno, class, sem)
                VALUES (%s, %s, %s)
                """,
                (
                    regno,
                    cls,
                    semester
                )
            )

            inserted += 1

        # -----------------------------
        # CHECK STUDENTS
        # -----------------------------
        if inserted == 0:

            conn.rollback()

            cur.close()
            conn.close()

            return jsonify({
                "message":
                "No valid register numbers found in Excel"
            }), 400

        # -----------------------------
        # SAVE
        # -----------------------------
        conn.commit()

        cur.close()
        conn.close()

        print(
            f"{inserted} students uploaded successfully"
        )

        return jsonify({
            "message":
            f"Students uploaded successfully ({inserted} students)"
        }), 200

    except Exception as e:

        import traceback

        print("====================================")
        print("STUDENT UPLOAD ERROR:")
        traceback.print_exc()
        print("====================================")

        try:
            conn.rollback()
            cur.close()
            conn.close()
        except:
            pass

        return jsonify({
            "message": f"Student upload failed: {str(e)}"
        }), 500
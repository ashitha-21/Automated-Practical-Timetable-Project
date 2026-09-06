from flask import Blueprint, request, jsonify
from db import get_db_connection

labs_bp = Blueprint("labs", __name__)

@labs_bp.route("/add_labs", methods=["POST"])
def add_labs():

    data = request.json

    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("DELETE FROM labs")
    conn.commit() 
    
    for lab in data["labs"]:
        cur.execute(
            "INSERT INTO labs(lab_name,capacity) VALUES (%s,%s)",
            (lab, data["capacity"])
        )

    conn.commit()
    cur.close()
    conn.close()

    return jsonify({"message": "Labs saved"})

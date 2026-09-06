from flask import Blueprint, jsonify, render_template, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.page import PageMargins
from db import get_db_connection
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth
import io
from datetime import datetime

timetable_bp = Blueprint("timetable", __name__)


# =========================================================
# COMMON HELPERS
# =========================================================

def format_date(date_value):
    """Return date as dd-mm-yyyy."""
    if hasattr(date_value, "strftime"):
        return date_value.strftime("%d-%m-%Y")
    return str(date_value)


def format_day_date(date_value):
    """Return:
       Wednesday
       25-03-2026
    """
    if hasattr(date_value, "strftime"):
        return f"{date_value.strftime('%A')}\n{date_value.strftime('%d-%m-%Y')}"
    return str(date_value)


def format_time(time_value):
    """
    Make the stored session time look like the sample timetable.
    Examples:
        09.00-12.00 -> 9:00 – 12:00
        12.30-03.30 -> 12:30 – 3:30
        01.00-04.00 -> 1:00 – 4:00
    """
    if not time_value:
        return ""

    value = str(time_value).strip()
    parts = value.replace("–", "-").replace("—", "-").split("-")

    if len(parts) != 2:
        return value

    def clean_time(part):
        part = part.strip().replace(".", ":")
        try:
            hour, minute = part.split(":")
            hour = int(hour)
            minute = int(minute)
            return f"{hour}:{minute:02d}"
        except Exception:
            return part

    return f"{clean_time(parts[0])} – {clean_time(parts[1])}"


def class_display(class_name):
    """
    Convert the internal FY/SY/TY class names into the semester-style
    wording used in the sample timetable.
    """
    mapping = {
        "FY": "II SEM BCA",
        "SY": "IV SEM BCA",
        "TY": "VI SEM BCA"
    }
    return mapping.get(class_name, class_name or "")


def get_timetable_rows(cur):
    """Fetch timetable rows in the required Date -> Time -> Batch order."""
    cur.execute("""
        SELECT
            t.id,
            t.batch_id,
            t.regno_from,
            t.regno_to,
            s.subject_name AS subject,
            s.class AS class_name,
            l.lab_name,
            d.exam_date,
            se.session_time AS exam_time,
            se.session_no
        FROM timetable t
        JOIN subjects s ON t.subject_id = s.id
        JOIN labs l ON t.lab_id = l.id
        JOIN exam_days d ON t.day_id = d.id
        JOIN sessions se ON t.session_id = se.id
        ORDER BY
            d.exam_date,
            se.session_no,
            CAST(
                CASE
                    WHEN t.batch_id LIKE 'BM-%' THEN SUBSTRING(t.batch_id, 4)
                    WHEN t.batch_id LIKE 'BE-%' THEN SUBSTRING(t.batch_id, 4)
                    WHEN t.batch_id LIKE 'BA-%' THEN SUBSTRING(t.batch_id, 4)
                    WHEN t.batch_id LIKE 'BM%' THEN SUBSTRING(t.batch_id, 3)
                    WHEN t.batch_id LIKE 'BA%' THEN SUBSTRING(t.batch_id, 3)
                    ELSE SUBSTRING(t.batch_id, 1)
                END AS UNSIGNED
            ),
            l.lab_name
    """)
    return cur.fetchall()


def compress_register_numbers(numbers):
    """
    Convert individual register numbers into the compact format used
    in the sample:
        2351001,2351002,...,2351019
        -> 2351001 – 1019

    If there are gaps, preserve them as separate ranges:
        2351020-1025, 1027-1039
    """
    if not numbers:
        return ""

    # Keep only valid numeric register numbers and remove duplicates.
    values = []
    for number in numbers:
        try:
            values.append(int(str(number).strip()))
        except Exception:
            pass

    values = sorted(set(values))
    if not values:
        return ""

    groups = []
    start = previous = values[0]

    for value in values[1:]:
        if value == previous + 1:
            previous = value
        else:
            groups.append((start, previous))
            start = previous = value

    groups.append((start, previous))

    def compact_range(start_value, end_value):
        start_text = str(start_value)
        end_text = str(end_value)

        if start_value == end_value:
            return start_text

        # Display the complete register number for both values.
        return f"{start_text} – {end_text}"

    result = []
    for start_value, end_value in groups:
        result.append(compact_range(start_value, end_value))

    return ", ".join(result)


def prepare_display_rows(cur, rows):
    """
    Add:
      - sample-style class
      - formatted date/time
      - actual candidate register numbers
      - total candidate count
      - rowspan for Day & Date
    """
    # Fetch all students once. This avoids one database query per batch.
    cur.execute("""
        SELECT regno, class
        FROM students
        ORDER BY CAST(regno AS UNSIGNED)
    """)
    all_students = cur.fetchall()

    students_by_class = {"FY": [], "SY": [], "TY": []}

    for student in all_students:
        if student["class"] in students_by_class:
            try:
                students_by_class[student["class"]].append(
                    int(str(student["regno"]).strip())
                )
            except Exception:
                continue

    display_rows = []

    for row in rows:
        class_name = row["class_name"]

        try:
            reg_from = int(str(row["regno_from"]).strip())
            reg_to = int(str(row["regno_to"]).strip())
        except Exception:
            reg_from = None
            reg_to = None

        candidates = []

        if reg_from is not None and reg_to is not None:
            for regno in students_by_class.get(class_name, []):
                if reg_from <= regno <= reg_to:
                    candidates.append(regno)

        # Fallback: if the class could not be matched, use the stored
        # From/To values so that the row is never left empty.
        if not candidates and reg_from is not None:
            candidates = [reg_from]
            if reg_to != reg_from and reg_to is not None:
                candidates = list(range(reg_from, reg_to + 1))

        display_rows.append({
            "batch_id": row["batch_id"],
            "class_name": class_display(class_name),
            "subject": row["subject"],
            "lab_name": row["lab_name"],
            "exam_date": row["exam_date"],
            "day_date": format_day_date(row["exam_date"]),
            "exam_time": format_time(row["exam_time"]),
            "reg_numbers": compress_register_numbers(candidates),
            "total": len(candidates),
        })

    # Merge Day & Date vertically for each date.
    # Also assign alternating colours to the complete day's rows.
    index = 0
    day_index = 0

    while index < len(display_rows):

        current_date = display_rows[index]["exam_date"]
        end = index

        while (
            end + 1 < len(display_rows)
            and display_rows[end + 1]["exam_date"] == current_date
        ):
            end += 1

        rowspan = end - index + 1

        # Same day_index for every row belonging to this date
        for i in range(index, end + 1):
            display_rows[i]["date_rowspan"] = rowspan if i == index else 0
            display_rows[i]["day_index"] = day_index

        day_index += 1
        index = end + 1

    return display_rows


# =========================================================
# GENERATE TIMETABLE
# =========================================================

@timetable_bp.route("/generate_timetable", methods=["POST"])
def generate_timetable():

    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)

    try:
        # Remove old generated timetable before creating a new one.
        cur.execute("DELETE FROM timetable")
        conn.commit()

        cur.execute("SELECT * FROM students ORDER BY regno")
        students = cur.fetchall()

        cur.execute("""
            SELECT *
            FROM subjects
            ORDER BY subject_order,
            FIELD(class,'TY','SY','FY')
        """)
        subjects = cur.fetchall()

        cur.execute("SELECT * FROM labs ORDER BY id")
        labs = cur.fetchall()

        cur.execute("SELECT * FROM exam_days ORDER BY day_no")
        days = cur.fetchall()

        cur.execute("SELECT * FROM sessions ORDER BY session_no")
        sessions = cur.fetchall()

        if not students:
            return jsonify({"message": "Upload students first"}), 400

        if not labs:
            return jsonify({"message": "Add labs first"}), 400

        if not days:
            return jsonify({"message": "No exam dates found"}), 400

        if not sessions:
            return jsonify({"message": "No sessions configured"}), 400

        # ---------------- GROUP STUDENTS ----------------
        classes = {"FY": [], "SY": [], "TY": []}

        for student in students:
            if student["class"] in classes:
                classes[student["class"]].append(student)

        for cls in classes:
            classes[cls] = sorted(
                classes[cls],
                key=lambda x: int(x["regno"])
            )

        # ---------------- CREATE SLOTS ----------------
        # Order:
        # Date 1 -> Morning labs -> Afternoon labs
        # Date 2 -> Morning labs -> Afternoon labs
        slots = []

        for day in days:
            for session in sessions:
                for lab in labs:
                    slots.append({
                        "day_id": day["id"],
                        "session_id": session["id"],
                        "lab_id": lab["id"]
                    })

        slot_index = 0
        batch_no = 1

        # Existing project design uses one common capacity for all labs.
        capacity = int(labs[0]["capacity"])

        course_prefix = "B"

        for subject in subjects:

            cls_students = classes.get(subject["class"], [])

            if not cls_students:
                continue

            total_students = len(cls_students)

            # =====================================================
            # SMART BATCH ALLOCATION
            # If the remaining students are less than 70% of
            # capacity, redistribute them among the existing batches
            # instead of creating a very small final batch.
            # =====================================================

            full_batches = total_students // capacity
            remaining = total_students % capacity

            batches = []

            if remaining == 0:
                # Exact division - normal batches
                number_of_batches = full_batches

            elif remaining < (capacity * 0.70):
                # Small remainder - distribute all students evenly
                number_of_batches = full_batches

            else:
                # Remainder is large enough - create another batch
                number_of_batches = full_batches + 1

            # Safety: at least one batch
            number_of_batches = max(1, number_of_batches)

            # Distribute students as evenly as possible
            base_size = total_students // number_of_batches
            extra = total_students % number_of_batches

            start = 0

            for i in range(number_of_batches):

                # First 'extra' batches get one additional student
                batch_size = base_size + (1 if i < extra else 0)

                batches.append(
                    cls_students[start:start + batch_size]
                )

                start += batch_size

            for group in batches:

                if slot_index >= len(slots):
                    return jsonify({
                        "message": "Not enough exam days/sessions for allocation"
                    }), 400

                slot = slots[slot_index]

                # Morning = BM, Afternoon = BA.
                # The table/export layout is independent of this naming.
                if slot["session_id"] == 1:
                    session_prefix = "M"
                else:
                    session_prefix = "A"

                batch_id = f"{course_prefix}{session_prefix}{batch_no:03}"

                cur.execute("""
                    INSERT INTO timetable
                    (
                        batch_id,
                        regno_from,
                        regno_to,
                        subject_id,
                        lab_id,
                        day_id,
                        session_id
                    )
                    VALUES (%s,%s,%s,%s,%s,%s,%s)
                """, (
                    batch_id,
                    group[0]["regno"],
                    group[-1]["regno"],
                    subject["id"],
                    slot["lab_id"],
                    slot["day_id"],
                    slot["session_id"]
                ))

                batch_no += 1
                slot_index += 1

        conn.commit()

        return jsonify({
            "message": "Timetable generated successfully"
        })

    except Exception as e:
        conn.rollback()
        return jsonify({
            "message": f"Error generating timetable: {str(e)}"
        }), 500

    finally:
        cur.close()
        conn.close()


# =========================================================
# VIEW TIMETABLE
# =========================================================

@timetable_bp.route("/view_timetable")
def view_timetable():

    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)

    try:
        rows = get_timetable_rows(cur)
        timetable = prepare_display_rows(cur, rows)

        return render_template(
            "view_timetable.html",
            timetable=timetable
        )

    finally:
        cur.close()
        conn.close()


# =========================================================
# EXPORT EXCEL - SAME PATTERN AS PDF/SCREEN
# =========================================================

@timetable_bp.route("/export_excel")
def export_excel():

    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)

    try:
        rows = get_timetable_rows(cur)
        timetable = prepare_display_rows(cur, rows)

        wb = Workbook()
        ws = wb.active
        ws.title = "Practical Timetable"

        # ---------------- PAGE SETUP ----------------
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True

        ws.page_margins = PageMargins(
            left=0.25,
            right=0.25,
            top=0.35,
            bottom=0.35,
            header=0.15,
            footer=0.15
        )

        ws.print_title_rows = "1:1"

        # ---------------- COLUMN WIDTHS ----------------
        widths = {
            "A": 16,   # Day & Date
            "B": 15,   # Time
            "C": 18,   # Class
            "D": 30,   # Practical Subject
            "E": 13,   # Batch
            "F": 48,   # Register No.
            "G": 9,    # Total
            "H": 12    # Lab
        }

        for col, width in widths.items():
            ws.column_dimensions[col].width = width

        # ---------------- HEADER ----------------
        headers = [
            "Day & Date",
            "Time",
            "Class",
            "Practical Subject",
            "Batch",
            "Register No. of the candidates assigned",
            "Total",
            "Lab"
        ]

        ws.append(headers)

        header_fill = PatternFill("solid", fgColor="808080")
        header_font = Font(
            name="Arial",
            size=9,
            bold=True,
            color="FFFFFF"
        )

        thin_black = Side(style="thin", color="000000")
        border = Border(
            left=thin_black,
            right=thin_black,
            top=thin_black,
            bottom=thin_black
        )

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = border

        ws.row_dimensions[1].height = 35

        # ---------------- DATA ----------------
        for row in timetable:
            ws.append([
                row["day_date"] if row["date_rowspan"] else "",
                row["exam_time"],
                row["class_name"],
                row["subject"],
                row["batch_id"],
                row["reg_numbers"],
                row["total"],
                row["lab_name"]
            ])

        # Style data rows.
        for row_cells in ws.iter_rows(
            min_row=2,
            max_row=ws.max_row,
            min_col=1,
            max_col=8
        ):
            for cell in row_cells:
                cell.font = Font(
                    name="Arial",
                    size=9
                )
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True
                )
                cell.border = border

        # ---------------- MERGE DATE CELLS + ALTERNATING DAY COLORS ----------------
        excel_row = 2
        index = 0

        while index < len(timetable):

            rowspan = timetable[index]["date_rowspan"]
            day_index = timetable[index]["day_index"]

            # Day 1 = grey, Day 2 = white, Day 3 = grey, etc.
            if day_index % 2 == 0:
                fill = PatternFill("solid", fgColor="D9D9D9")
            else:
                fill = PatternFill("solid", fgColor="FFFFFF")

            # Apply the colour to ALL columns and ALL rows of this day
            for r in range(excel_row, excel_row + rowspan):
                for c in range(1, 9):
                    ws.cell(r, c).fill = fill

            # Merge the Day & Date column
            if rowspan > 1:
                ws.merge_cells(
                    start_row=excel_row,
                    start_column=1,
                    end_row=excel_row + rowspan - 1,
                    end_column=1
                )

            # Format the Day & Date cell
            date_cell = ws.cell(excel_row, 1)

            date_cell.font = Font(
                name="Arial",
                size=9,
                bold=False
            )

            date_cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

            # Row height
            for r in range(excel_row, excel_row + rowspan):
                ws.row_dimensions[r].height = 30

            excel_row += rowspan
            index += rowspan

        # Print area.
        ws.print_area = f"A1:H{ws.max_row}"

        # Center horizontally on printed page.
        ws.print_options.horizontalCentered = True

        # Footer page numbering.
        ws.oddFooter.right.text = "Page &P of &N"
        ws.oddFooter.right.size = 8
        ws.oddFooter.right.font = "Arial"

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        return send_file(
            buffer,
            as_attachment=True,
            download_name="practical_exam_timetable.xlsx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

    finally:
        cur.close()
        conn.close()


# =========================================================
# EXPORT PDF - LANDSCAPE A4, SAME PATTERN
# =========================================================

@timetable_bp.route("/export_pdf")
def export_pdf():

    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)

    try:
        rows = get_timetable_rows(cur)
        timetable = prepare_display_rows(cur, rows)

        buffer = io.BytesIO()

        # A4 LANDSCAPE
        pdf = SimpleDocTemplate(
            buffer,
            pagesize=landscape(A4),
            rightMargin=8 * mm,
            leftMargin=8 * mm,
            topMargin=8 * mm,
            bottomMargin=10 * mm
        )

        # ---------------- TABLE DATA ----------------
        data = [[
            "Day & Date",
            "Time",
            "Class",
            "Practical Subject",
            "Batch",
            "Register No. of the candidates assigned",
            "Total",
            "Lab"
        ]]

        for row in timetable:
            data.append([
                row["day_date"] if row["date_rowspan"] else "",
                row["exam_time"],
                row["class_name"],
                row["subject"],
                row["batch_id"],
                row["reg_numbers"],
                str(row["total"]),
                row["lab_name"]
            ])

        # A4 landscape usable width is approximately 825 pt.
        col_widths = [
            67,   # Day & Date
            58,   # Time
            67,   # Class
            125,  # Practical Subject
            55,   # Batch
            330,  # Register numbers
            38,   # Total
            55    # Lab
        ]

        table = Table(
            data,
            colWidths=col_widths,
            repeatRows=1,
            hAlign="CENTER"
        )

        style_commands = [
            # Header
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#808080")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 7.5),

            # Body
            ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 1), (-1, -1), 7.2),

            # Alignment
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),

            # Grid
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),

            # Header row height
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#808080")),
        ]

        # Merge Day & Date cells + alternating colours for each day
        pdf_row = 1  # first body row

        index = 0

        while index < len(timetable):

            rowspan = timetable[index]["date_rowspan"]
            day_index = timetable[index]["day_index"]

            # Merge Day & Date column
            if rowspan > 1:
                style_commands.append(
                    (
                        "SPAN",
                        (0, pdf_row),
                        (0, pdf_row + rowspan - 1)
                    )
                )

            # Day 1 = grey, Day 2 = white, Day 3 = grey, etc.
            if day_index % 2 == 0:
                day_color = colors.HexColor("#D9D9D9")
            else:
                day_color = colors.white

            # Apply colour to ALL columns and ALL rows of this day
            style_commands.append(
                (
                    "BACKGROUND",
                    (0, pdf_row),
                    (-1, pdf_row + rowspan - 1),
                    day_color
                )
            )

            pdf_row += rowspan
            index += rowspan

        table.setStyle(TableStyle(style_commands))

        # ---------------- PAGE FOOTER ----------------
        def add_page_number(canvas, document):
            canvas.saveState()
            canvas.setFont("Helvetica", 7)
            canvas.drawRightString(
                landscape(A4)[0] - 8 * mm,
                5 * mm,
                f"Page {document.page}"
            )
            canvas.restoreState()

        pdf.build(
            [table],
            onFirstPage=add_page_number,
            onLaterPages=add_page_number
        )

        buffer.seek(0)

        return send_file(
            buffer,
            as_attachment=True,
            download_name="practical_exam_timetable.pdf",
            mimetype="application/pdf"
        )

    finally:
        cur.close()
        conn.close()

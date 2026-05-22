import os
import sqlite3
from datetime import datetime, date, timedelta
from functools import wraps
from io import BytesIO
from email.message import EmailMessage
import smtplib

from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file, Response
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

APP_TITLE = "Smart Late Students System"
DB_PATH = os.environ.get("DB_PATH", "late_students_dynamic.db")

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-this-secret-key")

DEFAULT_ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "admin")
DEFAULT_ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin123")

SAMPLE_CLASSES = {
    "Grade 9": ["9A", "9B", "9C", "9D"],
    "Grade 10": ["10A", "10B", "10C"],
    "Grade 11": ["11A", "11B"],
    "Grade 12": ["12A", "12B"],
}

SAMPLE_STUDENTS = {
    "9A": ["Ahmed Al Mansoori", "Saeed Al Kaabi", "Hamad Al Nuaimi"],
    "9B": ["Omar Al Shamsi", "Mohammed Al Ali", "Khalifa Al Ameri"],
    "10A": ["Rashid Al Ketbi", "Sultan Al Darmaki", "Ali Al Hammadi"],
    "11A": ["Abdullah Al Zaabi", "Majid Al Suwaidi", "Yousef Al Blooshi"],
    "12A": ["Jassim Al Shamsi", "Hamdan Al Muhairi", "Mubarak Al Neyadi"],
}


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_conn() as conn:
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            full_name TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin','user')),
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        )
        """)
        conn.execute("""
        CREATE TABLE IF NOT EXISTS grades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            sort_order INTEGER NOT NULL DEFAULT 0
        )
        """)
        conn.execute("""
        CREATE TABLE IF NOT EXISTS sections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            grade_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            UNIQUE(grade_id, name),
            FOREIGN KEY(grade_id) REFERENCES grades(id) ON DELETE CASCADE
        )
        """)
        conn.execute("""
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            grade_id INTEGER NOT NULL,
            section_id INTEGER NOT NULL,
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            UNIQUE(name, grade_id, section_id),
            FOREIGN KEY(grade_id) REFERENCES grades(id) ON DELETE CASCADE,
            FOREIGN KEY(section_id) REFERENCES sections(id) ON DELETE CASCADE
        )
        """)
        conn.execute("""
        CREATE TABLE IF NOT EXISTS late_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            late_date TEXT NOT NULL,
            late_time TEXT NOT NULL,
            recorded_by INTEGER,
            created_at TEXT NOT NULL,
            UNIQUE(student_id, late_date),
            FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
            FOREIGN KEY(recorded_by) REFERENCES users(id) ON DELETE SET NULL
        )
        """)
        ensure_seed_data(conn)
        conn.commit()


def ensure_seed_data(conn):
    now = datetime.now().isoformat(timespec="seconds")
    if not conn.execute("SELECT id FROM users WHERE username=?", (DEFAULT_ADMIN_USERNAME,)).fetchone():
        conn.execute(
            "INSERT INTO users(username,password_hash,full_name,role,active,created_at) VALUES (?,?,?,?,?,?)",
            (DEFAULT_ADMIN_USERNAME, generate_password_hash(DEFAULT_ADMIN_PASSWORD), "System Admin", "admin", 1, now),
        )
    for idx, grade_name in enumerate(SAMPLE_CLASSES.keys(), start=9):
        conn.execute("INSERT OR IGNORE INTO grades(name, sort_order) VALUES (?, ?)", (grade_name, idx))
    for grade_name, sections in SAMPLE_CLASSES.items():
        grade = conn.execute("SELECT id FROM grades WHERE name=?", (grade_name,)).fetchone()
        for sec in sections:
            conn.execute("INSERT OR IGNORE INTO sections(grade_id,name) VALUES (?,?)", (grade["id"], sec))
    for sec_name, names in SAMPLE_STUDENTS.items():
        sec = conn.execute("SELECT s.id, s.grade_id FROM sections s WHERE s.name=?", (sec_name,)).fetchone()
        if sec:
            for name in names:
                conn.execute(
                    "INSERT OR IGNORE INTO students(name,grade_id,section_id,active,created_at) VALUES (?,?,?,?,?)",
                    (name, sec["grade_id"], sec["id"], 1, now),
                )


def current_user():
    uid = session.get("user_id")
    if not uid:
        return None
    with get_conn() as conn:
        return conn.execute("SELECT * FROM users WHERE id=? AND active=1", (uid,)).fetchone()


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not current_user():
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper


def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        u = current_user()
        if not u or u["role"] != "admin":
            flash("Admin access required.", "error")
            return redirect(url_for("dashboard"))
        return fn(*args, **kwargs)
    return wrapper


@app.context_processor
def inject_user():
    return {"app_title": APP_TITLE, "me": current_user()}


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        with get_conn() as conn:
            u = conn.execute("SELECT * FROM users WHERE username=? AND active=1", (username,)).fetchone()
        if u and check_password_hash(u["password_hash"], password):
            session.clear()
            session["user_id"] = u["id"]
            flash("Logged in successfully.", "success")
            return redirect(url_for("dashboard"))
        flash("Invalid username or password.", "error")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def dashboard():
    selected_grade = request.args.get("grade_id", type=int)
    selected_section = request.args.get("section_id", type=int)
    today_text = parse_date_or_today(request.args.get("date"))
    with get_conn() as conn:
        grades = conn.execute("SELECT * FROM grades ORDER BY sort_order, name").fetchall()
        if not selected_grade and grades:
            selected_grade = grades[0]["id"]
        sections = conn.execute("SELECT * FROM sections WHERE grade_id=? ORDER BY name", (selected_grade,)).fetchall() if selected_grade else []
        if not selected_section and sections:
            selected_section = sections[0]["id"]
        students = conn.execute("""
            SELECT st.*, g.name AS grade_name, s.name AS section_name
            FROM students st JOIN grades g ON g.id=st.grade_id JOIN sections s ON s.id=st.section_id
            WHERE st.active=1 AND st.grade_id=? AND st.section_id=?
            ORDER BY st.name
        """, (selected_grade, selected_section)).fetchall() if selected_grade and selected_section else []
        today_records = conn.execute("SELECT student_id, late_time FROM late_records WHERE late_date=?", (today_text,)).fetchall()
    record_map = {r["student_id"]: r["late_time"] for r in today_records}
    data = []
    for s in students:
        streak = warning_streak(s["id"], today_text)
        data.append({**dict(s), "late_time": record_map.get(s["id"]), "streak": streak, "color": status_color(streak)})
    return render_template("dashboard.html", grades=grades, sections=sections, students=data, selected_grade=selected_grade, selected_section=selected_section, today=today_text)


def consecutive_late_days(student_id, end_day):
    with get_conn() as conn:
        rows = conn.execute("SELECT late_date FROM late_records WHERE student_id=? AND late_date<=? ORDER BY late_date DESC", (student_id, end_day)).fetchall()
    days = {r["late_date"] for r in rows}
    d = datetime.fromisoformat(end_day).date()
    count = 0
    while d.isoformat() in days:
        count += 1
        d -= timedelta(days=1)
    return count


def parse_date_or_today(value):
    """Safely read yyyy-mm-dd from forms/query strings."""
    try:
        return datetime.strptime(value or "", "%Y-%m-%d").date().isoformat()
    except ValueError:
        return date.today().isoformat()


def total_late_days(student_id, end_day=None):
    """Total unique late days for a student up to the selected date.
    This is used for color status and the total shown on today's dashboard.
    """
    with get_conn() as conn:
        if end_day:
            row = conn.execute(
                "SELECT COUNT(DISTINCT late_date) AS c FROM late_records WHERE student_id=? AND late_date<=?",
                (student_id, end_day),
            ).fetchone()
        else:
            row = conn.execute(
                "SELECT COUNT(DISTINCT late_date) AS c FROM late_records WHERE student_id=?",
                (student_id,),
            ).fetchone()
    return row["c"] or 0


def warning_streak(student_id, today_text):
    # Backward-compatible name: now returns TOTAL late days up to selected date, not only consecutive days.
    return total_late_days(student_id, today_text)


def status_color(count):
    if count >= 4:
        return "red"
    if count >= 1:
        return "yellow"
    return "normal"


@app.post("/mark_late/<int:student_id>")
@login_required
def mark_late(student_id):
    selected_date = parse_date_or_today(request.form.get("date"))
    selected_grade = request.form.get("grade_id", type=int)
    selected_section = request.form.get("section_id", type=int)
    now = datetime.now()
    with get_conn() as conn:
        s = conn.execute("SELECT grade_id, section_id FROM students WHERE id=?", (student_id,)).fetchone()
        # The late_date can be a previous calendar date, while late_time records the real entry time.
        conn.execute("""
            INSERT OR REPLACE INTO late_records(student_id, late_date, late_time, recorded_by, created_at)
            VALUES (?, ?, ?, ?, ?)
        """, (student_id, selected_date, now.strftime("%H:%M:%S"), session["user_id"], now.isoformat(timespec="seconds")))
        conn.commit()
    flash(f"Late record saved for {selected_date}.", "success")
    return redirect(url_for("dashboard", grade_id=selected_grade or s["grade_id"], section_id=selected_section or s["section_id"], date=selected_date))


@app.post("/unmark_late/<int:student_id>")
@login_required
def unmark_late(student_id):
    selected_date = parse_date_or_today(request.form.get("date"))
    selected_grade = request.form.get("grade_id", type=int)
    selected_section = request.form.get("section_id", type=int)
    with get_conn() as conn:
        s = conn.execute("SELECT grade_id, section_id FROM students WHERE id=?", (student_id,)).fetchone()
        conn.execute("DELETE FROM late_records WHERE student_id=? AND late_date=?", (student_id, selected_date))
        conn.commit()
    flash(f"Late record removed for {selected_date}.", "success")
    return redirect(url_for("dashboard", grade_id=selected_grade or s["grade_id"], section_id=selected_section or s["section_id"], date=selected_date))


@app.route("/admin")
@login_required
@admin_required
def admin_home():
    with get_conn() as conn:
        grades = conn.execute("SELECT * FROM grades ORDER BY sort_order, name").fetchall()
        sections = conn.execute("SELECT s.*, g.name AS grade_name FROM sections s JOIN grades g ON g.id=s.grade_id ORDER BY g.sort_order, s.name").fetchall()
        students = conn.execute("""
            SELECT st.*, g.name AS grade_name, s.name AS section_name
            FROM students st JOIN grades g ON g.id=st.grade_id JOIN sections s ON s.id=st.section_id
            ORDER BY st.active DESC, g.sort_order, s.name, st.name
        """).fetchall()
        users = conn.execute("SELECT id, username, full_name, role, active, created_at FROM users ORDER BY role, username").fetchall()
    return render_template("admin.html", grades=grades, sections=sections, students=students, users=users)


@app.post("/admin/grades")
@login_required
@admin_required
def add_grade():
    with get_conn() as conn:
        conn.execute("INSERT OR IGNORE INTO grades(name, sort_order) VALUES (?, ?)", (request.form["name"].strip(), request.form.get("sort_order", type=int) or 0))
        conn.commit()
    return redirect(url_for("admin_home"))


@app.post("/admin/sections")
@login_required
@admin_required
def add_section():
    with get_conn() as conn:
        conn.execute("INSERT OR IGNORE INTO sections(grade_id, name) VALUES (?, ?)", (request.form.get("grade_id", type=int), request.form["name"].strip()))
        conn.commit()
    return redirect(url_for("admin_home"))


@app.post("/admin/students")
@login_required
@admin_required
def add_student():
    now = datetime.now().isoformat(timespec="seconds")
    section_id = request.form.get("section_id", type=int)
    with get_conn() as conn:
        sec = conn.execute("SELECT grade_id FROM sections WHERE id=?", (section_id,)).fetchone()
        conn.execute("INSERT OR IGNORE INTO students(name, grade_id, section_id, active, created_at) VALUES (?, ?, ?, 1, ?)", (request.form["name"].strip(), sec["grade_id"], section_id, now))
        conn.commit()
    return redirect(url_for("admin_home"))


@app.post("/admin/students/<int:student_id>/delete")
@login_required
@admin_required
def delete_student(student_id):
    with get_conn() as conn:
        conn.execute("UPDATE students SET active=0 WHERE id=?", (student_id,))
        conn.commit()
    return redirect(url_for("admin_home"))


@app.post("/admin/users")
@login_required
@admin_required
def add_user():
    now = datetime.now().isoformat(timespec="seconds")
    with get_conn() as conn:
        conn.execute("INSERT INTO users(username,password_hash,full_name,role,active,created_at) VALUES (?,?,?,?,1,?)", (
            request.form["username"].strip(), generate_password_hash(request.form["password"]), request.form["full_name"].strip(), request.form["role"], now
        ))
        conn.commit()
    return redirect(url_for("admin_home"))


@app.post("/admin/users/<int:user_id>/toggle")
@login_required
@admin_required
def toggle_user(user_id):
    if user_id == session.get("user_id"):
        flash("You cannot deactivate your own account.", "error")
        return redirect(url_for("admin_home"))
    with get_conn() as conn:
        conn.execute("UPDATE users SET active = CASE active WHEN 1 THEN 0 ELSE 1 END WHERE id=?", (user_id,))
        conn.commit()
    return redirect(url_for("admin_home"))


@app.route("/report")
@login_required
def report():
    today_iso = date.today().isoformat()
    from_date = parse_date_or_today(request.args.get("from_date") or request.args.get("date") or today_iso)
    to_date = parse_date_or_today(request.args.get("to_date") or from_date)
    if from_date > to_date:
        from_date, to_date = to_date, from_date
    records = get_records_range(from_date, to_date)
    return render_template("report.html", records=records, from_date=from_date, to_date=to_date, today=today_iso)


def get_records_range(from_date, to_date):
    with get_conn() as conn:
        rows = conn.execute("""
            SELECT r.*, st.name AS student_name, g.name AS grade_name, sec.name AS section_name,
                   u.full_name AS recorder_name, u.username AS recorder_username
            FROM late_records r
            JOIN students st ON st.id=r.student_id
            JOIN grades g ON g.id=st.grade_id
            JOIN sections sec ON sec.id=st.section_id
            LEFT JOIN users u ON u.id=r.recorded_by
            WHERE r.late_date BETWEEN ? AND ?
            ORDER BY r.late_date, g.sort_order, sec.name, st.name, r.late_time
        """, (from_date, to_date)).fetchall()
    result = []
    for r in rows:
        item = dict(r)
        item["total_days"] = total_late_days(r["student_id"], to_date)
        item["day_name"] = datetime.strptime(r["late_date"], "%Y-%m-%d").strftime("%A")
        result.append(item)
    return result


def get_daily_records(day_text):
    return get_records_range(day_text, day_text)


@app.route("/export/late_report.xlsx")
@login_required
def export_late_report_excel():
    today_iso = date.today().isoformat()
    from_date = parse_date_or_today(request.args.get("from_date") or request.args.get("date") or today_iso)
    to_date = parse_date_or_today(request.args.get("to_date") or from_date)
    if from_date > to_date:
        from_date, to_date = to_date, from_date
    records = get_records_range(from_date, to_date)
    return build_excel_response(records, from_date, to_date)


@app.route("/export/daily.xlsx")
@login_required
def export_daily_excel():
    # Kept for old buttons/links; now supports date or from_date/to_date.
    return export_late_report_excel()


def sanitize_sheet_name(name):
    """Clean invalid Excel sheet-name characters and keep length <= 31."""
    bad_chars = ['\\', '/', '?', '*', '[', ']', ':']
    clean = str(name or "Sheet")
    for ch in bad_chars:
        clean = clean.replace(ch, '-')
    return clean[:31] or "Sheet"


def each_date_between(from_date, to_date):
    start = datetime.strptime(from_date, "%Y-%m-%d").date()
    end = datetime.strptime(to_date, "%Y-%m-%d").date()
    while start <= end:
        yield start.isoformat(), start.strftime("%A")
        start += timedelta(days=1)


def all_grades_for_excel():
    with get_conn() as conn:
        return conn.execute("SELECT id, name, sort_order FROM grades ORDER BY sort_order, name").fetchall()


def build_excel_response(records, from_date, to_date):
    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    grades = all_grades_for_excel()
    records_by_grade_date = {}
    for r in records:
        records_by_grade_date.setdefault(r["grade_name"], {}).setdefault(r["late_date"], []).append(r)

    # Professional colors
    dark_blue = "1F4E78"
    light_blue = "5B9BD5"
    pale_blue = "D9EAF7"
    yellow = "FFF2CC"
    red = "F4CCCC"
    white = "FFFFFF"
    grey = "F2F2F2"
    border_color = "B7B7B7"

    headers = [
        "No.",
        "Student Name",
        "Grade",
        "Section",
        "Late/Absent Day",
        "Date",
        "Late Time",
        "Total Late Days",
        "Recorded By",
        "Username",
        "Entry Created At",
    ]

    if not grades:
        # Fallback sheet in the unlikely case the admin has removed all grades.
        grades = [{"name": "No Grades", "sort_order": 0}]

    used_sheet_names = set()
    for grade in grades:
        grade_name = grade["name"]
        base_name = sanitize_sheet_name(grade_name)
        sheet_name = base_name
        n = 2
        while sheet_name in used_sheet_names:
            suffix = f" {n}"
            sheet_name = sanitize_sheet_name(base_name[:31-len(suffix)] + suffix)
            n += 1
        used_sheet_names.add(sheet_name)

        ws = wb.create_sheet(sheet_name)
        ws.sheet_view.rightToLeft = False
        title = f"{grade_name} - Late/Absent Interval Report"
        period = f"Period: {from_date}" if from_date == to_date else f"Period: {from_date} to {to_date}"

        ws.merge_cells("A1:K1")
        ws["A1"] = title
        ws["A1"].font = Font(size=16, bold=True, color=white)
        ws["A1"].fill = PatternFill("solid", fgColor=dark_blue)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28

        ws.merge_cells("A2:K2")
        total_records_for_grade = sum(len(v) for v in records_by_grade_date.get(grade_name, {}).values())
        ws["A2"] = f"{period} | Total records in this grade: {total_records_for_grade}"
        ws["A2"].font = Font(size=11, bold=True, color="1F1F1F")
        ws["A2"].fill = PatternFill("solid", fgColor=pale_blue)
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

        header_row = 4
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(header_row, col)
            cell.value = header
            cell.font = Font(bold=True, color=white)
            cell.fill = PatternFill("solid", fgColor=dark_blue)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[header_row].height = 24
        ws.freeze_panes = "A5"

        counter = 1
        records_for_grade = records_by_grade_date.get(grade_name, {})
        for day_iso, day_name in each_date_between(from_date, to_date):
            group_row = ws.max_row + 1
            ws.merge_cells(start_row=group_row, start_column=1, end_row=group_row, end_column=11)
            group_cell = ws.cell(group_row, 1)
            group_cell.value = f"{day_name} - {day_iso}"
            group_cell.font = Font(bold=True, color=white)
            group_cell.fill = PatternFill("solid", fgColor=light_blue)
            group_cell.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[group_row].height = 21

            day_records = records_for_grade.get(day_iso, [])
            if day_records:
                for r in day_records:
                    ws.append([
                        counter,
                        r["student_name"],
                        r["grade_name"],
                        r["section_name"],
                        r["day_name"],
                        r["late_date"],
                        r["late_time"],
                        r["total_days"],
                        r["recorder_name"] or "Unknown",
                        r["recorder_username"] or "-",
                        r["created_at"],
                    ])
                    data_row = ws.max_row
                    total_cell = ws.cell(data_row, 8)
                    if (r["total_days"] or 0) >= 4:
                        total_cell.fill = PatternFill("solid", fgColor=red)
                    elif (r["total_days"] or 0) >= 1:
                        total_cell.fill = PatternFill("solid", fgColor=yellow)
                    counter += 1
            else:
                ws.append(["", "No late/absent students recorded", grade_name, "", day_name, day_iso, "", "", "", "", ""])
                no_row = ws.max_row
                for col in range(1, 12):
                    ws.cell(no_row, col).fill = PatternFill("solid", fgColor=grey)
                ws.cell(no_row, 2).font = Font(italic=True, color="666666")

        # Formatting and print readiness
        thin = Side(style="thin", color=border_color)
        for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=1, max_col=11):
            for cell in row:
                cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
                cell.alignment = Alignment(vertical="center", wrap_text=True)

        widths = [7, 30, 13, 12, 18, 14, 13, 17, 24, 16, 22]
        for i, width in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = width
        ws.auto_filter.ref = f"A4:K{ws.max_row}"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_title_rows = "1:4"
        ws.page_margins.left = 0.25
        ws.page_margins.right = 0.25
        ws.page_margins.top = 0.5
        ws.page_margins.bottom = 0.5

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    suffix = from_date if from_date == to_date else f"{from_date}_to_{to_date}"
    return send_file(
        bio,
        as_attachment=True,
        download_name=f"late_absent_interval_{suffix}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )



# =========================
# AI ANALYTICS + CHATBOT
# =========================

def grade_number_from_name(name):
    digits = ''.join(ch for ch in str(name or '') if ch.isdigit())
    return int(digits) if digits else 999


def get_ai_risk_report(from_date, to_date):
    """Rule-based AI engine for lateness risk.
    It detects: high total lateness, consecutive days, and repeated weekday patterns.
    No paid AI API is required, so it works on Render immediately.
    """
    with get_conn() as conn:
        students = conn.execute("""
            SELECT st.id, st.name AS student_name, g.name AS grade_name, sec.name AS section_name
            FROM students st
            JOIN grades g ON g.id=st.grade_id
            JOIN sections sec ON sec.id=st.section_id
            WHERE st.active=1
            ORDER BY g.sort_order, sec.name, st.name
        """).fetchall()
        rows = conn.execute("""
            SELECT r.student_id, r.late_date, r.late_time, u.full_name AS recorder_name, u.username AS recorder_username
            FROM late_records r
            LEFT JOIN users u ON u.id=r.recorded_by
            WHERE r.late_date BETWEEN ? AND ?
            ORDER BY r.student_id, r.late_date
        """, (from_date, to_date)).fetchall()

    by_student = {}
    for r in rows:
        by_student.setdefault(r['student_id'], []).append(dict(r))

    report = []
    for st in students:
        records = by_student.get(st['id'], [])
        dates = sorted({r['late_date'] for r in records})
        total = len(dates)
        weekday_counts = {}
        for d in dates:
            day_name = datetime.strptime(d, "%Y-%m-%d").strftime("%A")
            weekday_counts[day_name] = weekday_counts.get(day_name, 0) + 1
        repeated_weekdays = {k: v for k, v in weekday_counts.items() if v >= 2}

        max_consecutive = 0
        current = 0
        previous_day = None
        for d in dates:
            current_day = datetime.strptime(d, "%Y-%m-%d").date()
            if previous_day and (current_day - previous_day).days == 1:
                current += 1
            else:
                current = 1
            max_consecutive = max(max_consecutive, current)
            previous_day = current_day

        reasons = []
        score = 0
        if total > 5:
            score += 45
            reasons.append(f"More than 5 late records in the selected period ({total}).")
        elif total >= 3:
            score += 25
            reasons.append(f"Repeated lateness: {total} records in the selected period.")
        elif total >= 1:
            score += 10
            reasons.append(f"Recorded late {total} time(s) in the selected period.")

        if max_consecutive >= 2:
            score += 30
            reasons.append(f"Consecutive lateness detected for {max_consecutive} day(s).")

        if repeated_weekdays:
            score += 25
            repeated_text = ', '.join([f"{day} ({count} times)" for day, count in repeated_weekdays.items()])
            reasons.append(f"Repeated weekday pattern: {repeated_text}.")

        if total == 0:
            risk = "Low"
            recommendation = "No action required. Continue normal monitoring."
            reasons.append("No lateness recorded in this period.")
        elif score >= 70:
            risk = "High"
            recommendation = "Immediate follow-up is recommended. Review morning arrival pattern and assign targeted monitoring."
        elif score >= 35:
            risk = "Medium"
            recommendation = "Monitor closely for the next week and remind the student about punctual arrival."
        else:
            risk = "Low"
            recommendation = "Keep monitoring. No urgent intervention is required now."

        report.append({
            "student_id": st['id'],
            "student_name": st['student_name'],
            "grade_name": st['grade_name'],
            "section_name": st['section_name'],
            "total_late_days": total,
            "max_consecutive": max_consecutive,
            "repeated_weekdays": repeated_weekdays,
            "risk": risk,
            "score": min(score, 100),
            "reasons": reasons,
            "recommendation": recommendation,
            "records": records,
        })
    risk_order = {"High": 0, "Medium": 1, "Low": 2}
    report.sort(key=lambda x: (risk_order.get(x['risk'], 3), -x['total_late_days'], x['grade_name'], x['section_name'], x['student_name']))
    return report


def ai_summary_text(report, from_date, to_date):
    total_students = len(report)
    high = sum(1 for r in report if r['risk'] == 'High')
    medium = sum(1 for r in report if r['risk'] == 'Medium')
    low = sum(1 for r in report if r['risk'] == 'Low')
    total_records = sum(r['total_late_days'] for r in report)
    top = [r for r in report if r['total_late_days'] > 0][:5]
    lines = [
        f"AI Summary for {from_date} to {to_date}",
        f"Total students analyzed: {total_students}.",
        f"Total late records: {total_records}.",
        f"Risk distribution: High {high}, Medium {medium}, Low {low}.",
    ]
    if top:
        lines.append("Top students needing follow-up: " + "; ".join([f"{r['student_name']} ({r['risk']}, {r['total_late_days']} days)" for r in top]) + ".")
    else:
        lines.append("No late records found in this period.")
    return " ".join(lines)


@app.route('/ai-report')
@login_required
def ai_report():
    today_iso = date.today().isoformat()
    from_date = parse_date_or_today(request.args.get('from_date') or today_iso)
    to_date = parse_date_or_today(request.args.get('to_date') or from_date)
    if from_date > to_date:
        from_date, to_date = to_date, from_date
    report = get_ai_risk_report(from_date, to_date)
    summary = ai_summary_text(report, from_date, to_date)
    return render_template('ai_report.html', report=report, summary=summary, from_date=from_date, to_date=to_date, today=today_iso)


def is_arabic(text):
    return any('\u0600' <= ch <= '\u06FF' for ch in text or '')


def chatbot_answer(question):
    q = (question or '').strip()
    q_low = q.lower()
    arabic = is_arabic(q)
    today_iso = date.today().isoformat()
    default_from = (date.today() - timedelta(days=30)).isoformat()
    report = get_ai_risk_report(default_from, today_iso)

    # Find a student name mentioned in the question.
    mentioned = None
    for r in report:
        if r['student_name'].lower() in q_low:
            mentioned = r
            break
        parts = [p.lower() for p in r['student_name'].split() if len(p) > 2]
        if parts and any(p in q_low for p in parts):
            mentioned = r
            break

    def en_student(r):
        weekdays = ', '.join([f"{d} ({c})" for d, c in r['repeated_weekdays'].items()]) or 'No repeated weekday pattern'
        reasons = ' '.join(r['reasons'])
        return (f"{r['student_name']} - {r['grade_name']} {r['section_name']}\n"
                f"Risk Level: {r['risk']} ({r['score']}/100)\n"
                f"Total Late Days in last 30 days: {r['total_late_days']}\n"
                f"Max Consecutive Late Days: {r['max_consecutive']}\n"
                f"Repeated Weekdays: {weekdays}\n"
                f"AI Reasons: {reasons}\n"
                f"Recommendation: {r['recommendation']}")

    def ar_student(r):
        risk_ar = {'High': 'مرتفع', 'Medium': 'متوسط', 'Low': 'منخفض'}.get(r['risk'], r['risk'])
        weekdays = '، '.join([f"{d} ({c})" for d, c in r['repeated_weekdays'].items()]) or 'لا يوجد نمط يوم متكرر'
        reasons = ' '.join(r['reasons'])
        return (f"الطالب: {r['student_name']} - {r['grade_name']} {r['section_name']}\n"
                f"مستوى الخطورة: {risk_ar} ({r['score']}/100)\n"
                f"مجموع أيام التأخير آخر 30 يومًا: {r['total_late_days']}\n"
                f"أعلى تأخير متتالي: {r['max_consecutive']} يوم\n"
                f"الأيام المتكررة: {weekdays}\n"
                f"أسباب التحليل: {reasons}\n"
                f"التوصية: {r['recommendation']}")

    if mentioned:
        return ar_student(mentioned) if arabic else en_student(mentioned)

    if any(w in q_low for w in ['highest', 'most', 'top', 'أكثر', 'اعلى', 'أعلى']):
        top = [r for r in report if r['total_late_days'] > 0][:5]
        if not top:
            return 'لا توجد سجلات تأخير خلال آخر 30 يومًا.' if arabic else 'No late records were found in the last 30 days.'
        if arabic:
            return 'أكثر الطلاب تأخيرًا آخر 30 يومًا:\n' + '\n'.join([f"- {r['student_name']} | {r['grade_name']} {r['section_name']} | {r['total_late_days']} أيام | خطورة {r['risk']}" for r in top])
        return 'Top late students in the last 30 days:\n' + '\n'.join([f"- {r['student_name']} | {r['grade_name']} {r['section_name']} | {r['total_late_days']} days | {r['risk']} risk" for r in top])

    if any(w in q_low for w in ['risk', 'danger', 'خطورة', 'ريسك', 'خطر']):
        high = [r for r in report if r['risk'] == 'High']
        medium = [r for r in report if r['risk'] == 'Medium']
        if arabic:
            return f"تحليل الخطورة آخر 30 يومًا: مرتفع {len(high)} طالب، متوسط {len(medium)} طالب.\n" + ('الطلاب عالي الخطورة:\n' + '\n'.join([f"- {r['student_name']} ({r['total_late_days']} أيام)" for r in high[:10]]) if high else 'لا يوجد طلاب عالي الخطورة.')
        return f"Risk analysis for the last 30 days: High risk {len(high)} students, Medium risk {len(medium)} students.\n" + ('High-risk students:\n' + '\n'.join([f"- {r['student_name']} ({r['total_late_days']} days)" for r in high[:10]]) if high else 'No high-risk students found.')

    if any(w in q_low for w in ['summary', 'report', 'ملخص', 'تقرير']):
        return ai_summary_text(report, default_from, today_iso)

    # Grade questions such as Grade 9 / الصف التاسع.
    grade_map = {
        '9': ['grade 9', 'تاسع', 'التاسع', 'صف 9'],
        '10': ['grade 10', 'عاشر', 'العاشر', 'صف 10'],
        '11': ['grade 11', 'حادي عشر', 'الحادي عشر', 'صف 11'],
        '12': ['grade 12', 'ثاني عشر', 'الثاني عشر', 'صف 12'],
    }
    for num, keys in grade_map.items():
        if any(k in q_low for k in keys):
            grade_rows = [r for r in report if str(grade_number_from_name(r['grade_name'])) == num]
            total = sum(r['total_late_days'] for r in grade_rows)
            high = sum(1 for r in grade_rows if r['risk'] == 'High')
            if arabic:
                return f"ملخص Grade {num} آخر 30 يومًا: مجموع سجلات التأخير {total}، وعدد الطلاب عالي الخطورة {high}."
            return f"Grade {num} summary for the last 30 days: {total} total late records, {high} high-risk students."

    if arabic:
        return "يمكنك أن تسألني مثل: ما خطورة الطالب Ahmed Al Mansoori؟ من أكثر الطلاب تأخيرًا؟ أعطني ملخص Grade 10؟ من الطلاب عالي الخطورة؟"
    return "You can ask me: What is Ahmed Al Mansoori's risk? Who is late the most? Give me Grade 10 summary. Who are the high-risk students?"


@app.route('/chatbot', methods=['GET', 'POST'])
@login_required
def chatbot():
    answer = None
    question = ''
    if request.method == 'POST':
        question = request.form.get('question', '')
        answer = chatbot_answer(question)
    return render_template('chatbot.html', question=question, answer=answer)


def build_email_body(day_text):
    records = get_daily_records(day_text)
    if not records:
        return f"Daily Late Arrival Report - {day_text}\n\nNo late students recorded today."
    lines = [f"Daily Late Arrival Report - {day_text}", "", "Student | Grade | Section | Date | Time | Total Days | Recorded By", "-"*90]
    for r in records:
        lines.append(f"{r['student_name']} | {r['grade_name']} | {r['section_name']} | {r['late_date']} | {r['late_time']} | {r['total_days']} | {r['recorder_name'] or 'Unknown'}")
    return "\n".join(lines)


def send_daily_email():
    recipients = [x.strip() for x in os.environ.get("REPORT_RECIPIENTS", "").split(",") if x.strip()]
    if not recipients:
        print("REPORT_RECIPIENTS is empty. Email skipped.")
        return
    smtp_host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
    smtp_port = int(os.environ.get("SMTP_PORT", "587"))
    smtp_user = os.environ.get("SMTP_USER")
    smtp_password = os.environ.get("SMTP_PASSWORD")
    sender = os.environ.get("EMAIL_FROM", smtp_user)
    if not smtp_user or not smtp_password or not sender:
        print("SMTP settings missing. Email skipped.")
        return
    day_text = date.today().isoformat()
    msg = EmailMessage()
    msg["Subject"] = f"Daily Late Arrival Report - {day_text}"
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg.set_content(build_email_body(day_text))
    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)


@app.route("/send-report-now")
def send_report_now():
    if request.args.get("token") != os.environ.get("CRON_SECRET", "change-me"):
        return Response("Unauthorized", status=401)
    send_daily_email()
    return "Report email sent or skipped. Check logs."


if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
else:
    init_db()

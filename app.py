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


def build_excel_response(records, from_date, to_date):
    wb = Workbook()
    ws = wb.active
    ws.title = "Late Students"
    ws.merge_cells("A1:J1")
    title = f"Late/Absent Interval Report - {from_date}" if from_date == to_date else f"Late/Absent Interval Report - {from_date} to {to_date}"
    ws["A1"] = title
    ws["A1"].font = Font(size=16, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.append([])
    headers = ["No.", "Student Name", "Grade", "Section", "Late/Absent Day", "Late Date", "Late Time", "Total Late Days", "Recorded By", "Entry Created At"]
    ws.append(headers)
    for cell in ws[3]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center")

    current_date = None
    counter = 1
    if records:
        for r in records:
            if r["late_date"] != current_date:
                current_date = r["late_date"]
                ws.append([])
                group_row = ws.max_row
                ws.merge_cells(start_row=group_row, start_column=1, end_row=group_row, end_column=10)
                ws.cell(group_row, 1).value = f"{r['day_name']} - {r['late_date']}"
                ws.cell(group_row, 1).font = Font(bold=True, color="FFFFFF")
                ws.cell(group_row, 1).fill = PatternFill("solid", fgColor="5B9BD5")
                ws.cell(group_row, 1).alignment = Alignment(horizontal="left")
            ws.append([
                counter, r["student_name"], r["grade_name"], r["section_name"], r["day_name"], r["late_date"], r["late_time"],
                r["total_days"], f"{r['recorder_name'] or 'Unknown'} ({r['recorder_username'] or '-'})", r["created_at"]
            ])
            counter += 1
    else:
        ws.append(["No late/absent students recorded for this selected interval."])

    thin = Side(style="thin", color="B7B7B7")
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=10):
        for cell in row:
            cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
            cell.alignment = Alignment(vertical="center")
    widths = [8, 28, 14, 12, 18, 14, 14, 18, 30, 22]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    suffix = from_date if from_date == to_date else f"{from_date}_to_{to_date}"
    return send_file(bio, as_attachment=True, download_name=f"late_absent_interval_{suffix}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


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

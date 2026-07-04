from flask import Flask, render_template, request, redirect, url_for
from werkzeug.security import check_password_hash, generate_password_hash
import mysql.connector
from mysql.connector import Error
import random
from openpyxl import Workbook
from flask import send_file
import csv
from flask import Response,flash


app = Flask(
    __name__,
    template_folder="Templates",
    static_folder="static"
)
app.secret_key = "campussync123"


@app.context_processor
def inject_admin():
    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)
    cursor.execute("SELECT * FROM users WHERE role='admin' LIMIT 1")
    admin = cursor.fetchone()
    cursor.close()
    connection.close()
    return {"admin": admin}


DB_CONFIG = {
    "host": "localhost",
    "user": "root",
    "password": "22@Yuva@sri",
    "database": "campussync",
}

DASHBOARD_ROUTES = {
    "admin": "dashboard_admin",
    "teacher": "dashboard_teacher",
    "student": "dashboard_student",
}


def get_db_connection(use_database=True):
    config = DB_CONFIG.copy()
    if not use_database:
        config.pop("database")
    return mysql.connector.connect(**config)


def ensure_column(connection, table_name, column_name, definition):
    cursor = connection.cursor()
    cursor.execute("SHOW COLUMNS FROM `%s` LIKE %%s" % table_name, (column_name,))
    if cursor.fetchone() is None:
        cursor.execute(f"ALTER TABLE `{table_name}` ADD COLUMN {definition}")
    cursor.close()


def initialize_database():
    connection = get_db_connection(use_database=False)
    cursor = connection.cursor()
    cursor.execute("CREATE DATABASE IF NOT EXISTS campussync")
    cursor.close()
    connection.close()

    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INT AUTO_INCREMENT PRIMARY KEY,
            name VARCHAR(100) NOT NULL,
            email VARCHAR(100) NOT NULL UNIQUE,
            password_hash VARCHAR(255) NOT NULL,
            role ENUM('student', 'teacher', 'admin') NOT NULL,
            notice_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS students (
            id INT AUTO_INCREMENT PRIMARY KEY,
            name VARCHAR(100) NOT NULL,
            email VARCHAR(100) NOT NULL UNIQUE,
            enrollment_id VARCHAR(30) NOT NULL UNIQUE,
            class_name VARCHAR(50) NOT NULL,
            status VARCHAR(20) NOT NULL DEFAULT 'Active'
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS teachers (
            id INT AUTO_INCREMENT PRIMARY KEY,
            name VARCHAR(100) NOT NULL,
            email VARCHAR(100) NOT NULL UNIQUE,
            department VARCHAR(80) NOT NULL,
            designation VARCHAR(100) DEFAULT NULL,
            classes_assigned INT NOT NULL DEFAULT 0,
            status VARCHAR(20) NOT NULL DEFAULT 'Active'
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS subjects (
            id INT AUTO_INCREMENT PRIMARY KEY,
            name VARCHAR(100) NOT NULL,
            code VARCHAR(20) NOT NULL UNIQUE,
            instructor VARCHAR(100) NOT NULL,
            credits INT NOT NULL,
            students_enrolled INT NOT NULL DEFAULT 0,
            status VARCHAR(20) NOT NULL DEFAULT 'Active'
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS assignments (
            id INT AUTO_INCREMENT PRIMARY KEY,
            title VARCHAR(120) NOT NULL,
            class_name VARCHAR(100) NOT NULL,
            due_date DATE NOT NULL,
            submissions VARCHAR(20) NOT NULL,
            status VARCHAR(20) NOT NULL DEFAULT 'Active'
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS attendance (
            id INT AUTO_INCREMENT PRIMARY KEY,
            student_name VARCHAR(100) NOT NULL,
            class_name VARCHAR(100) NOT NULL,
            attendance_date DATE NOT NULL,
            status VARCHAR(20) NOT NULL
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS grades (
            id INT AUTO_INCREMENT PRIMARY KEY,
            course VARCHAR(100) NOT NULL,
            code VARCHAR(20) NOT NULL,
            instructor VARCHAR(100) NOT NULL,
            credits INT NOT NULL,
            grade VARCHAR(5) NOT NULL,
            points INT NOT NULL
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS reports (
            id INT AUTO_INCREMENT PRIMARY KEY,
            report_name VARCHAR(120) NOT NULL,
            report_type VARCHAR(50) NOT NULL,
            generated_date DATE NOT NULL,
            generated_by VARCHAR(100) NOT NULL
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS notices (
            id INT AUTO_INCREMENT PRIMARY KEY,
            title VARCHAR(120) NOT NULL,
            body TEXT NOT NULL,
            notice_date DATE NOT NULL
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS student_dashboard_items (
            id INT AUTO_INCREMENT PRIMARY KEY,
            title VARCHAR(120) NOT NULL,
            detail TEXT NOT NULL,
            due_date DATE NOT NULL,
            status VARCHAR(30) NOT NULL DEFAULT 'Due Soon'
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS student_preferences (
            id INT AUTO_INCREMENT PRIMARY KEY,
            name VARCHAR(80) NOT NULL,
            value VARCHAR(80) NOT NULL,
            description TEXT NOT NULL
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS classes (
            id INT AUTO_INCREMENT PRIMARY KEY,
            class_name VARCHAR(100) NOT NULL UNIQUE
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS subject_allocation (
            id INT AUTO_INCREMENT PRIMARY KEY,
            class_id INT NOT NULL,
            subject_id INT NOT NULL,
            teacher_id INT NOT NULL
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS timetable (
            id INT AUTO_INCREMENT PRIMARY KEY,
            class_name VARCHAR(100) NOT NULL,
            day VARCHAR(20) NOT NULL,
            period INT NOT NULL,
            subject VARCHAR(120) NOT NULL,
            teacher VARCHAR(120) NOT NULL,
            room VARCHAR(20) NOT NULL
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS student_submissions (
            id INT AUTO_INCREMENT PRIMARY KEY,
            assignment_id INT NOT NULL,
            student_name VARCHAR(100) NOT NULL,
            submitted_on DATE NOT NULL,
            status VARCHAR(20) NOT NULL DEFAULT 'Pending',
            notes TEXT NOT NULL
        )
        """
    )
    ensure_column(connection, "teachers", "designation", "designation VARCHAR(100) DEFAULT NULL")
    ensure_column(connection, "teachers", "classes_assigned", "classes_assigned INT NOT NULL DEFAULT 0")
    ensure_column(connection, "subjects", "students_enrolled", "students_enrolled INT NOT NULL DEFAULT 0")
    ensure_column(connection, "assignments", "submissions", "submissions VARCHAR(20) NOT NULL DEFAULT '0/0'")
    ensure_column(connection, "assignments", "status", "status VARCHAR(20) NOT NULL DEFAULT 'Active'")
    seed_database(cursor)
    connection.commit()
    cursor.close()
    connection.close()


def seed_database(cursor):
    cursor.executemany(
        """
        INSERT IGNORE INTO students (name, email, enrollment_id, class_name, status)
        VALUES (%s, %s, %s, %s, %s)
        """,
        [
            ("Aarav Patel", "aarav.patel@campus.edu", "CS2024001", "Class 12 A", "Active"),
            ("Zara Khan", "zara.khan@campus.edu", "CS2024002", "Class 12 B", "Active"),
            ("Ravi Singh", "ravi.singh@campus.edu", "CS2024003", "Class 12 A", "Active"),
            ("Priya Sharma", "priya.sharma@campus.edu", "CS2024004", "Class 11 C", "Inactive"),
            ("Neha Verma", "neha.verma@campus.edu", "CS2024005", "Class 12 B", "Active"),
        ],
    )
    cursor.executemany(
        """
        INSERT IGNORE INTO teachers (name, email, department, classes_assigned, status)
        VALUES (%s, %s, %s, %s, %s)
        """,
        [
            ("Dr. Rajesh Kumar", "rajesh.kumar@campus.edu", "Mathematics", 3, "Active"),
            ("Ms. Kavya Rao", "kavya.rao@campus.edu", "Physics", 2, "Active"),
            ("Prof. Anjali Singh", "anjali.singh@campus.edu", "English", 4, "Active"),
            ("Dr. Vikram Patel", "vikram.patel@campus.edu", "Chemistry", 2, "Inactive"),
            ("Ms. Priya Nair", "priya.nair@campus.edu", "Computer Science", 3, "Active"),
        ],
    )
    cursor.executemany(
        """
        INSERT IGNORE INTO subjects (name, code, instructor, credits, students_enrolled, status)
        VALUES (%s, %s, %s, %s, %s, %s)
        """,
        [
            ("Mathematics 101", "M101", "Dr. Rajesh Kumar", 4, 45, "Active"),
            ("Physics 201", "P201", "Ms. Kavya Rao", 4, 30, "Active"),
            ("English Literature", "E150", "Prof. Anjali Singh", 3, 35, "Active"),
            ("Chemistry Basics", "C120", "Dr. Vikram Patel", 4, 28, "Inactive"),
            ("Computer Science Fundamentals", "CS101", "Ms. Priya Nair", 3, 42, "Active"),
        ],
    )
    cursor.executemany(
        """
        INSERT IGNORE INTO assignments (id, title, class_name, due_date, submissions, status)
        VALUES (%s, %s, %s, %s, %s, %s)
        """,
        [
            (1, "Literature Essay", "English 150", "2026-06-22", "32/35", "Active"),
            (2, "Lab Report", "Physics 201", "2026-06-24", "28/30", "Active"),
            (3, "Midterm Exam", "Mathematics 101", "2026-06-18", "45/45", "Closed"),
            (4, "Project Proposal", "Chemistry 120", "2026-06-27", "15/28", "Active"),
        ],
    )
    cursor.executemany(
        """
        INSERT IGNORE INTO attendance (id, student_name, class_name, attendance_date, status)
        VALUES (%s, %s, %s, %s, %s)
        """,
        [
            (1, "Aarav Patel", "Mathematics 101", "2026-06-21", "Present"),
            (2, "Zara Khan", "Physics 201", "2026-06-21", "Absent"),
            (3, "Ravi Singh", "English 150", "2026-06-21", "Late"),
            (4, "Priya Sharma", "Chemistry 120", "2026-06-21", "Present"),
        ],
    )
    cursor.executemany(
        """
        INSERT IGNORE INTO grades (id, course, code, instructor, credits, grade, points)
        VALUES (%s, %s, %s, %s, %s, %s, %s)
        """,
        [
            (1, "Mathematics 101", "M101", "Dr. Rajesh Kumar", 4, "A", 89),
            (2, "Physics 201", "P201", "Ms. Kavya Rao", 4, "A", 85),
            (3, "English Literature", "E150", "Prof. Anjali Singh", 3, "A", 91),
            (4, "Chemistry Basics", "C120", "Dr. Vikram Patel", 4, "B", 87),
            (5, "Computer Science", "CS101", "Ms. Priya Nair", 3, "A", 93),
        ],
    )
    cursor.executemany(
        """
        INSERT IGNORE INTO reports (id, report_name, report_type, generated_date, generated_by)
        VALUES (%s, %s, %s, %s, %s)
        """,
        [
            (1, "Monthly Attendance Summary", "Attendance", "2026-06-20", "Admin User"),
            (2, "Semester Performance Analysis", "Academic", "2026-06-15", "Admin User"),
            (3, "Student Enrollment Report", "Enrollment", "2026-06-10", "Admin User"),
        ],
    )
    cursor.executemany(
        """
        INSERT IGNORE INTO notices (id, title, body, notice_date)
        VALUES (%s, %s, %s, %s)
        """,
        [
            (1, "Exam Schedule Published", "Semester exam dates are now available.", "2026-06-18"),
            (2, "Library Hours Updated", "The library will remain open until 8 PM on weekdays.", "2026-06-19"),
            (3, "Project Submission Reminder", "Final project files are due this week.", "2026-06-20"),
        ],
    )


def fetch_all(query, params=()):
    try:
        connection = get_db_connection()
        cursor = connection.cursor(dictionary=True)
        cursor.execute(query, params)
        rows = cursor.fetchall()
        cursor.close()
        connection.close()
        return rows
    except Error as error:
        print(f"MySQL query failed: {error}")
        return []


def fetch_count(table_name):
    rows = fetch_all(f"SELECT COUNT(*) AS total FROM {table_name}")
    return rows[0]["total"] if rows else 0


def dashboard_for_role(role):
    return DASHBOARD_ROUTES.get(role.lower(), "welcome")


def get_first_user_by_role(role):
    rows = fetch_all("SELECT * FROM users WHERE role = %s LIMIT 1", (role,))
    if rows:
        return rows[0]
    return None


try:
    initialize_database()
except Error as error:
    print(f"MySQL setup skipped: {error}")

@app.route("/")
def welcome():
    return render_template(
        "welcome.html",
        notices=fetch_all("SELECT * FROM notices ORDER BY notice_date DESC LIMIT 3"),
        teacher_count=fetch_count("teachers"),
        student_count=fetch_count("students"),
        subject_count=fetch_count("subjects"),
    )

from flask import session

app.secret_key = "campussync123"

@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        email = request.form.get("email")
        password = request.form.get("password")
        role = request.form.get("role", "").lower()

        connection = get_db_connection()
        cursor = connection.cursor(dictionary=True)

        cursor.execute(
            "SELECT * FROM users WHERE email=%s AND role=%s",
            (email, role)
        )

        user = cursor.fetchone()

        if user and check_password_hash(user["password_hash"], password):

            session["user_id"] = user["id"]
            session["name"] = user["name"]
            session["email"] = user["email"]
            session["role"] = user["role"]

            cursor.close()
            connection.close()

            return redirect(url_for(dashboard_for_role(user["role"])))

        cursor.close()
        connection.close()

        return "Invalid Email or Password"

    return render_template("login.html")
@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        name = request.form.get("name")
        email = request.form.get("email")
        password = request.form.get("password")
        confirm = request.form.get("confirm")
        role = request.form.get("role", "").lower()

        if password != confirm:
            return "Passwords do not match. Please go back and try again.", 400

        if role not in DASHBOARD_ROUTES:
            return "Please select a valid role.", 400

        try:
            connection = get_db_connection()
            cursor = connection.cursor()
            cursor.execute(
                """
                INSERT INTO users (name, email, password_hash, role)
                VALUES (%s, %s, %s, %s)
                """,
                (name, email, generate_password_hash(password), role),
            )
            connection.commit()
            cursor.close()
            connection.close()
        except Error as error:
            if getattr(error, "errno", None) == 1062:
                return "This email is already registered. Please log in.", 409
            return f"Database error: {error}", 500

        return redirect(url_for(dashboard_for_role(role)))
    return render_template("register.html")


@app.route("/dashboard_teacher")
def dashboard_teacher():
    return render_template(
        "teacher/dashboard.html",
        assignments=fetch_all("SELECT * FROM assignments ORDER BY due_date LIMIT 4"),
        attendance=fetch_all("SELECT * FROM attendance ORDER BY attendance_date DESC LIMIT 4"),
    )


@app.route("/dashboard_student")
def dashboard_student():

    if "email" not in session:
        return redirect(url_for("login"))

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    # Logged-in student
    cursor.execute("""
        SELECT *
        FROM students
        WHERE email=%s
    """, (session["email"],))

    student = cursor.fetchone()

    if not student:
        cursor.close()
        connection.close()
        return "Student not found."

    # Total Subjects
    cursor.execute("""
        SELECT COUNT(*) AS total
        FROM subject_allocation
        WHERE class_id=(
            SELECT id
            FROM classes
            WHERE class_name=%s
        )
    """, (student["class_name"],))

    total_subjects = cursor.fetchone()["total"]

    # Attendance Percentage
    cursor.execute("""
        SELECT
            COUNT(*) AS total_classes,
            SUM(CASE WHEN status='Present' THEN 1 ELSE 0 END) AS present
        FROM attendance
        WHERE student_name=%s
    """, (student["name"],))

    attendance = cursor.fetchone()

    if attendance["total_classes"] == 0:
        attendance_percent = 0
    else:
        attendance_percent = round(
            attendance["present"] * 100 / attendance["total_classes"], 1
        )

    # Pending Assignments
    cursor.execute("""
        SELECT COUNT(*) AS total
        FROM assignments
        WHERE class_name=%s
        AND status='Pending'
    """, (student["class_name"],))

    pending = cursor.fetchone()["total"]

    # Today's Timetable
    import datetime

    today = datetime.datetime.today().strftime("%A")

    cursor.execute("""
        SELECT *
        FROM timetable
        WHERE class_name=%s
        AND day=%s
        ORDER BY period
    """,
    (
        student["class_name"],
        today
    ))

    timetable = cursor.fetchall()

    # Latest Notices
    cursor.execute("""
        SELECT *
        FROM notices
        ORDER BY id DESC
        LIMIT 5
    """)

    notices = cursor.fetchall()

    cursor.close()
    connection.close()

    return render_template(
        "student/dashboard.html",
        student=student,
        total_subjects=total_subjects,
        attendance_percent=attendance_percent,
        pending=pending,
        timetable=timetable,
        notices=notices
    )

@app.route("/student_add_dashboard_item", methods=["POST"])
def student_add_dashboard_item():
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute(
        """
        INSERT INTO student_dashboard_items (title, detail, due_date, status)
        VALUES (%s, %s, %s, %s)
        """,
        (
            request.form.get("title"),
            request.form.get("detail"),
            request.form.get("due_date"),
            request.form.get("status", "Due Soon"),
        ),
    )
    connection.commit()
    cursor.close()
    connection.close()
    return redirect(url_for("dashboard_student"))


@app.route("/student_edit_dashboard_item/<int:id>", methods=["GET", "POST"])
def student_edit_dashboard_item(id):
    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)
    if request.method == "POST":
        cursor.execute(
            """
            UPDATE student_dashboard_items
            SET title=%s, detail=%s, due_date=%s, status=%s
            WHERE id=%s
            """,
            (
                request.form.get("title"),
                request.form.get("detail"),
                request.form.get("due_date"),
                request.form.get("status", "Due Soon"),
                id,
            ),
        )
        connection.commit()
        cursor.close()
        connection.close()
        return redirect(url_for("dashboard_student"))

    cursor.execute("SELECT * FROM student_dashboard_items WHERE id=%s", (id,))
    item = cursor.fetchone()
    cursor.execute("SELECT * FROM student_dashboard_items ORDER BY due_date")
    items = cursor.fetchall()
    cursor.close()
    connection.close()
    return render_template(
        "student/dashboard.html",
        assignments=fetch_all("SELECT * FROM assignments ORDER BY due_date LIMIT 4"),
        dashboard_items=items,
        edit_dashboard_item=item,
        user_name=get_first_user_by_role("student").get("name") if get_first_user_by_role("student") else "Student",
    )


@app.route("/student_delete_dashboard_item/<int:id>")
def student_delete_dashboard_item(id):
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute("DELETE FROM student_dashboard_items WHERE id=%s", (id,))
    connection.commit()
    cursor.close()
    connection.close()
    return redirect(url_for("dashboard_student"))


@app.route("/classes_teacher")
def classes_teacher():
    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)
    cursor.execute("SELECT * FROM subjects ORDER BY name")
    subjects = cursor.fetchall()
    cursor.close()
    connection.close()
    return render_template("teacher/classes.html", subjects=subjects, edit_subject=None)


@app.route("/teacher_add_subject", methods=["POST"])
def teacher_add_subject():
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute(
        """
        INSERT INTO subjects (code, name, instructor, credits, students_enrolled, status)
        VALUES (%s, %s, %s, %s, %s, %s)
        """,
        (
            request.form.get("code"),
            request.form.get("name"),
            request.form.get("instructor"),
            request.form.get("credits"),
            request.form.get("students_enrolled"),
            request.form.get("status", "Active"),
        ),
    )
    connection.commit()
    cursor.close()
    connection.close()
    return redirect(url_for("classes_teacher"))


@app.route("/teacher_edit_subject/<int:id>", methods=["GET", "POST"])
def teacher_edit_subject(id):
    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)
    if request.method == "POST":
        cursor.execute(
            """
            UPDATE subjects
            SET code=%s, name=%s, instructor=%s, credits=%s, students_enrolled=%s, status=%s
            WHERE id=%s
            """,
            (
                request.form.get("code"),
                request.form.get("name"),
                request.form.get("instructor"),
                request.form.get("credits"),
                request.form.get("students_enrolled"),
                request.form.get("status", "Active"),
                id,
            ),
        )
        connection.commit()
        cursor.close()
        connection.close()
        return redirect(url_for("classes_teacher"))

    cursor.execute("SELECT * FROM subjects WHERE id=%s", (id,))
    subject = cursor.fetchone()
    cursor.execute("SELECT * FROM subjects ORDER BY name")
    subjects = cursor.fetchall()
    cursor.close()
    connection.close()
    return render_template("teacher/classes.html", subjects=subjects, edit_subject=subject)


@app.route("/teacher_delete_subject/<int:id>")
def teacher_delete_subject(id):
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute("DELETE FROM subjects WHERE id=%s", (id,))
    connection.commit()
    cursor.close()
    connection.close()
    return redirect(url_for("classes_teacher"))


@app.route("/assignments_teacher")
def assignments_teacher():
    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)
    cursor.execute("SELECT * FROM assignments ORDER BY due_date")
    assignments = cursor.fetchall()
    cursor.close()
    connection.close()
    return render_template("teacher/assignments.html", assignments=assignments, edit_assignment=None)


@app.route("/teacher_add_assignment", methods=["POST"])
def teacher_add_assignment():
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute(
        """
        INSERT INTO assignments (title, class_name, due_date, submissions, status)
        VALUES (%s, %s, %s, %s, %s)
        """,
        (
            request.form.get("title"),
            request.form.get("class_name"),
            request.form.get("due_date"),
            request.form.get("submissions", "0/0"),
            request.form.get("status", "Active"),
        ),
    )
    connection.commit()
    cursor.close()
    connection.close()
    return redirect(url_for("assignments_teacher"))


@app.route("/teacher_edit_assignment/<int:id>", methods=["GET", "POST"])
def teacher_edit_assignment(id):
    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)
    if request.method == "POST":
        cursor.execute(
            """
            UPDATE assignments
            SET title=%s, class_name=%s, due_date=%s, submissions=%s, status=%s
            WHERE id=%s
            """,
            (
                request.form.get("title"),
                request.form.get("class_name"),
                request.form.get("due_date"),
                request.form.get("submissions", "0/0"),
                request.form.get("status", "Active"),
                id,
            ),
        )
        connection.commit()
        cursor.close()
        connection.close()
        return redirect(url_for("assignments_teacher"))

    cursor.execute("SELECT * FROM assignments WHERE id=%s", (id,))
    assignment = cursor.fetchone()
    cursor.execute("SELECT * FROM assignments ORDER BY due_date")
    assignments = cursor.fetchall()
    cursor.close()
    connection.close()
    return render_template("teacher/assignments.html", assignments=assignments, edit_assignment=assignment)


@app.route("/teacher_delete_assignment/<int:id>")
def teacher_delete_assignment(id):
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute("DELETE FROM assignments WHERE id=%s", (id,))
    connection.commit()
    cursor.close()
    connection.close()
    return redirect(url_for("assignments_teacher"))


@app.route("/attendance_teacher")
def attendance_teacher():
    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)
    cursor.execute("SELECT * FROM attendance ORDER BY attendance_date DESC")
    attendance = cursor.fetchall()
    cursor.close()
    connection.close()
    return render_template("teacher/attendance.html", attendance=attendance, edit_attendance=None)


@app.route("/teacher_add_attendance", methods=["POST"])
def teacher_add_attendance():
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute(
        """
        INSERT INTO attendance (student_name, class_name, attendance_date, status)
        VALUES (%s, %s, %s, %s)
        """,
        (
            request.form.get("student_name"),
            request.form.get("class_name"),
            request.form.get("attendance_date"),
            request.form.get("status"),
        ),
    )
    connection.commit()
    cursor.close()
    connection.close()
    return redirect(url_for("attendance_teacher"))


@app.route("/teacher_edit_attendance/<int:id>", methods=["GET", "POST"])
def teacher_edit_attendance(id):
    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)
    if request.method == "POST":
        cursor.execute(
            """
            UPDATE attendance
            SET student_name=%s, class_name=%s, attendance_date=%s, status=%s
            WHERE id=%s
            """,
            (
                request.form.get("student_name"),
                request.form.get("class_name"),
                request.form.get("attendance_date"),
                request.form.get("status"),
                id,
            ),
        )
        connection.commit()
        cursor.close()
        connection.close()
        return redirect(url_for("attendance_teacher"))

    cursor.execute("SELECT * FROM attendance WHERE id=%s", (id,))
    attendance = cursor.fetchone()
    cursor.execute("SELECT * FROM attendance ORDER BY attendance_date DESC")
    attendance_rows = cursor.fetchall()
    cursor.close()
    connection.close()
    return render_template("teacher/attendance.html", attendance=attendance_rows, edit_attendance=attendance)


@app.route("/teacher_delete_attendance/<int:id>")
def teacher_delete_attendance(id):
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute("DELETE FROM attendance WHERE id=%s", (id,))
    connection.commit()
    cursor.close()
    connection.close()
    return redirect(url_for("attendance_teacher"))


@app.route("/schedule_teacher")
def schedule_teacher():
    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)
    cursor.execute("SELECT * FROM timetable ORDER BY FIELD(day,'Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'), period")
    schedule_items = cursor.fetchall()
    cursor.close()
    connection.close()
    return render_template("teacher/schedule.html", schedule_items=schedule_items, edit_schedule=None)


@app.route("/teacher_add_schedule", methods=["POST"])
def teacher_add_schedule():
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute(
        """
        INSERT INTO timetable (class_name, day, period, subject, teacher, room)
        VALUES (%s, %s, %s, %s, %s, %s)
        """,
        (
            request.form.get("class_name"),
            request.form.get("day"),
            request.form.get("period"),
            request.form.get("subject"),
            request.form.get("teacher"),
            request.form.get("room"),
        ),
    )
    connection.commit()
    cursor.close()
    connection.close()
    return redirect(url_for("schedule_teacher"))


@app.route("/teacher_edit_schedule/<int:id>", methods=["GET", "POST"])
def teacher_edit_schedule(id):
    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)
    if request.method == "POST":
        cursor.execute(
            """
            UPDATE timetable
            SET class_name=%s, day=%s, period=%s, subject=%s, teacher=%s, room=%s
            WHERE id=%s
            """,
            (
                request.form.get("class_name"),
                request.form.get("day"),
                request.form.get("period"),
                request.form.get("subject"),
                request.form.get("teacher"),
                request.form.get("room"),
                id,
            ),
        )
        connection.commit()
        cursor.close()
        connection.close()
        return redirect(url_for("schedule_teacher"))

    cursor.execute("SELECT * FROM timetable WHERE id=%s", (id,))
    schedule = cursor.fetchone()
    cursor.execute("SELECT * FROM timetable ORDER BY FIELD(day,'Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'), period")
    schedule_items = cursor.fetchall()
    cursor.close()
    connection.close()
    return render_template("teacher/schedule.html", schedule_items=schedule_items, edit_schedule=schedule)


@app.route("/teacher_delete_schedule/<int:id>")
def teacher_delete_schedule(id):
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute("DELETE FROM timetable WHERE id=%s", (id,))
    connection.commit()
    cursor.close()
    connection.close()
    return redirect(url_for("schedule_teacher"))


@app.route("/messages_teacher")
def messages_teacher():
    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)
    cursor.execute("SELECT * FROM notices ORDER BY notice_date DESC")
    notices = cursor.fetchall()
    cursor.close()
    connection.close()
    return render_template("teacher/messages.html", notices=notices, edit_notice=None)


@app.route("/teacher_add_notice", methods=["POST"])
def teacher_add_notice():
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute(
        """
        INSERT INTO notices (title, body, notice_date)
        VALUES (%s, %s, %s)
        """,
        (
            request.form.get("title"),
            request.form.get("body"),
            request.form.get("notice_date"),
        ),
    )
    connection.commit()
    cursor.close()
    connection.close()
    return redirect(url_for("messages_teacher"))


@app.route("/teacher_edit_notice/<int:id>", methods=["GET", "POST"])
def teacher_edit_notice(id):
    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)
    if request.method == "POST":
        cursor.execute(
            """
            UPDATE notices
            SET title=%s, body=%s, notice_date=%s
            WHERE id=%s
            """,
            (
                request.form.get("title"),
                request.form.get("body"),
                request.form.get("notice_date"),
                id,
            ),
        )
        connection.commit()
        cursor.close()
        connection.close()
        return redirect(url_for("messages_teacher"))

    cursor.execute("SELECT * FROM notices WHERE id=%s", (id,))
    notice = cursor.fetchone()
    cursor.execute("SELECT * FROM notices ORDER BY notice_date DESC")
    notices = cursor.fetchall()
    cursor.close()
    connection.close()
    return render_template("teacher/messages.html", notices=notices, edit_notice=notice)


@app.route("/teacher_delete_notice/<int:id>")
def teacher_delete_notice(id):
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute("DELETE FROM notices WHERE id=%s", (id,))
    connection.commit()
    cursor.close()
    connection.close()
    return redirect(url_for("messages_teacher"))


@app.route("/settings_teacher")
def settings_teacher():
    user = get_first_user_by_role("teacher") or {}
    return render_template("teacher/settings.html", user_name=user.get("name") if user else "Teacher")


@app.route("/dashboard_admin")
def dashboard_admin():
    return render_template(
        "admin/dashboard.html",
        teacher_count=fetch_count("teachers"),
        student_count=fetch_count("students"),
        subject_count=fetch_count("subjects"),
        reports=fetch_all("SELECT * FROM reports ORDER BY generated_date DESC LIMIT 3"),
    )

@app.route("/add_student", methods=["GET", "POST"])
def add_student():

    if request.method == "POST":

        name = request.form["name"]
        email = request.form["email"]
        enrollment_id = request.form["enrollment_id"]
        class_name = request.form["class_name"]
        status = request.form["status"]

        connection = get_db_connection()
        cursor = connection.cursor()

        cursor.execute("""
        INSERT INTO students
        (name,email,enrollment_id,class_name,status)
        VALUES(%s,%s,%s,%s,%s)
        """,
        (name,email,enrollment_id,class_name,status))

        connection.commit()

        cursor.close()
        connection.close()

        return redirect(url_for("students_admin"))

    return render_template("admin/add_student.html")

@app.route("/edit_student/<int:id>", methods=["GET","POST"])
def edit_student(id):

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    if request.method == "POST":

        name = request.form["name"]
        email = request.form["email"]
        enrollment_id = request.form["enrollment_id"]
        class_name = request.form["class_name"]
        status = request.form["status"]

        cursor.execute("""
        UPDATE students
        SET
        name=%s,
        email=%s,
        enrollment_id=%s,
        class_name=%s,
        status=%s
        WHERE id=%s
        """,

        (
            name,
            email,
            enrollment_id,
            class_name,
            status,
            id
        ))

        connection.commit()

        cursor.close()
        connection.close()

        return redirect(url_for("students_admin"))

    cursor.execute(
        "SELECT * FROM students WHERE id=%s",
        (id,)
    )

    student = cursor.fetchone()

    cursor.close()
    connection.close()

    return render_template(
        "admin/edit_student.html",
        student=student
    )

@app.route("/delete_student/<int:id>")
def delete_student(id):

    connection = get_db_connection()
    cursor = connection.cursor()

    cursor.execute(
        "DELETE FROM students WHERE id=%s",
        (id,)
    )

    connection.commit()

    cursor.close()
    connection.close()

    return redirect(url_for("students_admin"))

@app.route("/students_admin")
def students_admin():
    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)
    cursor.execute("SELECT * FROM students")
    students = cursor.fetchall()
    cursor.close()
    connection.close()
    return render_template(
        "admin/students.html",
        students=students
    )

@app.route("/add_teacher", methods=["GET", "POST"])
def add_teacher():

    if request.method == "POST":

        name = request.form["name"]
        email = request.form["email"]
        department = request.form["department"]
        designation = request.form["designation"]
        status = request.form["status"]

        connection = get_db_connection()
        cursor = connection.cursor()

        cursor.execute("""
        INSERT INTO teachers
        (name,email,department,designation,status)
        VALUES(%s,%s,%s,%s,%s)
        """,(name,email,department,designation,status))

        connection.commit()

        cursor.close()
        connection.close()

        return redirect(url_for("teachers_admin"))

    return render_template("admin/add_teacher.html")

@app.route("/edit_teacher/<int:id>", methods=["GET","POST"])
def edit_teacher(id):

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    if request.method == "POST":

        name = request.form["name"]
        email = request.form["email"]
        department = request.form["department"]
        designation = request.form["designation"]
        status = request.form["status"]

        cursor.execute("""
        UPDATE teachers
        SET
        name=%s,
        email=%s,
        department=%s,
        designation=%s,
        status=%s
        WHERE id=%s
        """,
        (
            name,
            email,
            department,
            designation,
            status,
            id
        ))

        connection.commit()

        cursor.close()
        connection.close()

        return redirect(url_for("teachers_admin"))

    cursor.execute(
        "SELECT * FROM teachers WHERE id=%s",
        (id,)
    )

    teacher = cursor.fetchone()

    cursor.close()
    connection.close()

    return render_template(
        "admin/edit_teacher.html",
        teacher=teacher
    )

@app.route("/delete_teacher/<int:id>")
def delete_teacher(id):

    connection = get_db_connection()
    cursor = connection.cursor()

    # Delete related subject allocations
    cursor.execute(
        "DELETE FROM subject_allocation WHERE teacher_id=%s",
        (id,)
    )

    # Delete teacher
    cursor.execute(
        "DELETE FROM teachers WHERE id=%s",
        (id,)
    )

    connection.commit()

    cursor.close()
    connection.close()

    flash("Teacher deleted successfully.", "success")

    return redirect(url_for("teachers_admin"))

@app.route("/teachers_admin")
def teachers_admin():
    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)
    cursor.execute("SELECT * FROM teachers")
    teachers = cursor.fetchall()
    cursor.close()
    connection.close()
    return render_template(
        "admin/teachers.html",
        teachers=teachers
    )

@app.route("/add_subject", methods=["GET", "POST"])
def add_subject():

    if request.method == "POST":

        code = request.form["code"]
        name = request.form["name"]
        instructor = request.form["instructor"]
        credits = request.form["credits"]
        students_enrolled = request.form["students_enrolled"]
        status = request.form["status"]

        connection = get_db_connection()
        cursor = connection.cursor()

        cursor.execute("""
        INSERT INTO subjects
        (code,name,instructor,credits,students_enrolled,status)
        VALUES(%s,%s,%s,%s,%s,%s)
        """,
        (
            code,
            name,
            instructor,
            credits,
            students_enrolled,
            status
        ))

        connection.commit()

        cursor.close()
        connection.close()

        return redirect(url_for("subjects_admin"))

    return render_template("admin/add_subject.html")

@app.route("/edit_subject/<int:id>", methods=["GET","POST"])
def edit_subject(id):

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    if request.method=="POST":

        code=request.form["code"]
        name=request.form["name"]
        instructor=request.form["instructor"]
        credits=request.form["credits"]
        students_enrolled=request.form["students_enrolled"]
        status=request.form["status"]

        cursor.execute("""
        UPDATE subjects
        SET
        code=%s,
        name=%s,
        instructor=%s,
        credits=%s,
        students_enrolled=%s,
        status=%s
        WHERE id=%s
        """,
        (
            code,
            name,
            instructor,
            credits,
            students_enrolled,
            status,
            id
        ))

        connection.commit()

        cursor.close()
        connection.close()

        return redirect(url_for("subjects_admin"))

    cursor.execute(
        "SELECT * FROM subjects WHERE id=%s",
        (id,)
    )

    subject=cursor.fetchone()

    cursor.close()
    connection.close()

    return render_template(
        "admin/edit_subject.html",
        subject=subject
    )

@app.route("/delete_subject/<int:id>")
def delete_subject(id):

    connection=get_db_connection()

    cursor=connection.cursor()

    cursor.execute(
        "DELETE FROM subjects WHERE id=%s",
        (id,)
    )

    connection.commit()

    cursor.close()

    connection.close()

    return redirect(url_for("subjects_admin"))

@app.route("/subjects_admin")
def subjects_admin():
    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)
    cursor.execute("SELECT * FROM subjects")
    subjects = cursor.fetchall()
    cursor.close()
    connection.close()
    return render_template(
        "admin/subjects.html",
        subjects=subjects
    )

@app.route("/subject_allocation_admin")
def subject_allocation_admin():

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    cursor.execute("SELECT * FROM classes")
    classes = cursor.fetchall()

    cursor.execute("SELECT * FROM subjects")
    subjects = cursor.fetchall()

    cursor.execute("SELECT * FROM teachers")
    teachers = cursor.fetchall()

    cursor.execute("""
    SELECT
        sa.id,
        c.class_name,
        s.name AS subject,
        t.name AS teacher
    FROM subject_allocation sa
    JOIN classes c
        ON sa.class_id=c.id
    JOIN subjects s
        ON sa.subject_id=s.id
    JOIN teachers t
        ON sa.teacher_id=t.id
    """)

    allocations = cursor.fetchall()

    cursor.close()
    connection.close()

    return render_template(
        "admin/subject_allocation.html",
        classes=classes,
        subjects=subjects,
        teachers=teachers,
        allocations=allocations
    )


@app.route("/add_subject_allocation", methods=["POST"])
def add_subject_allocation():

    class_id = request.form["class_id"]
    subject_id = request.form["subject_id"]
    teacher_id = request.form["teacher_id"]

    connection = get_db_connection()
    cursor = connection.cursor()

    cursor.execute("""
    INSERT INTO subject_allocation
    (class_id,subject_id,teacher_id)
    VALUES(%s,%s,%s)
    """,
    (
        class_id,
        subject_id,
        teacher_id
    ))

    connection.commit()

    cursor.close()
    connection.close()

    return redirect(url_for("subject_allocation_admin"))


@app.route("/delete_subject_allocation/<int:id>")
def delete_subject_allocation(id):

    connection = get_db_connection()
    cursor = connection.cursor()

    cursor.execute(
        "DELETE FROM subject_allocation WHERE id=%s",
        (id,)
    )

    connection.commit()

    cursor.close()
    connection.close()

    return redirect(url_for("subject_allocation_admin"))

@app.route("/classes_admin")
def classes_admin():

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    cursor.execute("SELECT * FROM classes")
    classes = cursor.fetchall()

    total_students = sum(c["strength"] or 0 for c in classes)

    cursor.close()
    connection.close()

    return render_template(
        "admin/classes_admin.html",
        classes=classes,
        total_students=total_students
    )

@app.route("/add_class", methods=["GET","POST"])
def add_class():

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    if request.method == "POST":

        cursor.execute("""
            INSERT INTO classes
            (class_name,department,year,semester,section,strength,class_teacher,status)
            VALUES(%s,%s,%s,%s,%s,%s,%s,%s)
        """,(

            request.form["class_name"],
            request.form["department"],
            request.form["year"],
            request.form["semester"],
            request.form["section"],
            request.form["strength"],
            request.form["class_teacher"],
            request.form["status"]

        ))

        connection.commit()

        flash("Class Added Successfully")

        cursor.close()
        connection.close()

        return redirect(url_for("classes_admin"))

    cursor.close()
    connection.close()

    return render_template("admin/add_class.html")

@app.route("/edit_class/<int:id>", methods=["GET","POST"])
def edit_class(id):

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    if request.method == "POST":

        cursor.execute("""
            UPDATE classes
            SET
            class_name=%s,
            department=%s,
            year=%s,
            semester=%s,
            section=%s,
            strength=%s,
            class_teacher=%s,
            status=%s
            WHERE id=%s
        """,(

            request.form["class_name"],
            request.form["department"],
            request.form["year"],
            request.form["semester"],
            request.form["section"],
            request.form["strength"],
            request.form["class_teacher"],
            request.form["status"],
            id

        ))

        connection.commit()

        flash("Class Updated Successfully")

        cursor.close()
        connection.close()

        return redirect(url_for("classes_admin"))

    cursor.execute("SELECT * FROM classes WHERE id=%s",(id,))
    edit_class=cursor.fetchone()

    cursor.close()
    connection.close()

    return render_template(
        "admin/edit_class.html",
        edit_class=edit_class
    )

@app.route("/delete_class/<int:id>")
def delete_class(id):

    connection=get_db_connection()
    cursor=connection.cursor()

    cursor.execute("DELETE FROM classes WHERE id=%s",(id,))

    connection.commit()

    flash("Class Deleted Successfully")

    cursor.close()
    connection.close()

    return redirect(url_for("classes_admin"))

@app.route("/timetable_admin")
def timetable_admin():

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    # Load all classes
    cursor.execute("""
        SELECT class_name
        FROM classes
        ORDER BY class_name
    """)
    classes = cursor.fetchall()

    selected_class = request.args.get("class_name")

    if selected_class:

        cursor.execute("""
            SELECT *
            FROM timetable
            WHERE class_name=%s
            ORDER BY
            FIELD(day,
            'Monday',
            'Tuesday',
            'Wednesday',
            'Thursday',
            'Friday',
            'Saturday'),
            period
        """, (selected_class,))

    else:

        # Show all timetables if no class selected
        cursor.execute("""
            SELECT *
            FROM timetable
            ORDER BY
            class_name,
            FIELD(day,
            'Monday',
            'Tuesday',
            'Wednesday',
            'Thursday',
            'Friday',
            'Saturday'),
            period
        """)

    timetable = cursor.fetchall()

    cursor.close()
    connection.close()

    return render_template(
        "admin/timetable.html",
        classes=classes,
        timetable=timetable,
        selected_class=selected_class
    )

@app.route("/generate_timetable", methods=["POST"])
def generate_timetable():

    class_name = request.form["class_name"]

    working_days = int(request.form["working_days"])

    periods = int(request.form["periods"])


    if working_days == 5:

        days = [
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday"
        ]

    else:

        days = [
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday",
            "Saturday"
        ]


    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    # Subjects allocated for selected class

    cursor.execute("""

        SELECT
            s.name,
            t.name AS instructor

        FROM subject_allocation sa

        JOIN subjects s
        ON sa.subject_id=s.id

        JOIN teachers t
        ON sa.teacher_id=t.id

        JOIN classes c
        ON sa.class_id=c.id

        WHERE c.class_name=%s

    """,(class_name,))

    subjects = cursor.fetchall()


    if len(subjects)==0:

        cursor.close()
        connection.close()

        return "No Subjects Allocated."


    # Delete only selected class timetable

    cursor.execute("""

        DELETE FROM timetable
        WHERE class_name=%s

    """,(class_name,))

    connection.commit()

    random.shuffle(subjects)

    room=101

    subject_index=0


    for day in days:

        for period in range(1,periods+1):

            current_subject=subjects[
                subject_index % len(subjects)
            ]

            cursor.execute("""

                INSERT INTO timetable
                (
                class_name,
                day,
                period,
                subject,
                teacher,
                room
                )

                VALUES(%s,%s,%s,%s,%s,%s)

            """,(

                class_name,
                day,
                period,
                current_subject["name"],
                current_subject["instructor"],
                "A"+str(room)

            ))

            subject_index+=1

            room+=1

            if room>110:
                room=101


    connection.commit()

    cursor.close()
    connection.close()

    return redirect(
        url_for(
            "timetable_admin",
            class_name=class_name
        )
    )

@app.route("/export_timetable")
def export_timetable():

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    cursor.execute("""
        SELECT
        class_name,
        day,
        period,
        subject,
        teacher,
        room
        FROM timetable
        ORDER BY
        class_name,
        FIELD(day,'Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'),
        period
    """)

    timetable = cursor.fetchall()

    cursor.close()
    connection.close()

    wb = Workbook()

    ws = wb.active

    ws.title = "Timetable"

    ws.append([
        "Class",
        "Day",
        "Period",
        "Subject",
        "Teacher",
        "Room"
    ])

    for row in timetable:

        ws.append([
            row["class_name"],
            row["day"],
            row["period"],
            row["subject"],
            row["teacher"],
            row["room"]
        ])

    filename = "CampusSync_Timetable.xlsx"

    wb.save(filename)

    return send_file(
        filename,
        as_attachment=True
    )

@app.route("/attendance_admin")
def attendance_admin():

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    cursor.execute("""
        SELECT *
        FROM attendance
        ORDER BY attendance_date DESC
    """)

    attendance = cursor.fetchall()

    cursor.close()
    connection.close()

    return render_template(
        "admin/attendance.html",
        attendance=attendance
    )

@app.route("/add_attendance", methods=["GET","POST"])
def add_attendance():

    if request.method == "POST":

        student_name = request.form["student_name"]
        class_name = request.form["class_name"]
        attendance_date = request.form["attendance_date"]
        status = request.form["status"]

        connection = get_db_connection()
        cursor = connection.cursor()

        cursor.execute("""
            INSERT INTO attendance
            (student_name,class_name,attendance_date,status)
            VALUES(%s,%s,%s,%s)
        """,
        (
            student_name,
            class_name,
            attendance_date,
            status
        ))

        connection.commit()

        cursor.close()
        connection.close()

        return redirect(url_for("attendance_admin"))

    return render_template("admin/add_attendance.html")

@app.route("/edit_attendance/<int:id>", methods=["GET","POST"])
def edit_attendance(id):

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    if request.method == "POST":

        student_name = request.form["student_name"]
        class_name = request.form["class_name"]
        attendance_date = request.form["attendance_date"]
        status = request.form["status"]

        cursor.execute("""
            UPDATE attendance
            SET
            student_name=%s,
            class_name=%s,
            attendance_date=%s,
            status=%s
            WHERE id=%s
        """,
        (
            student_name,
            class_name,
            attendance_date,
            status,
            id
        ))

        connection.commit()

        cursor.close()
        connection.close()

        return redirect(url_for("attendance_admin"))

    cursor.execute(
        "SELECT * FROM attendance WHERE id=%s",
        (id,)
    )

    attendance = cursor.fetchone()

    cursor.close()
    connection.close()

    return render_template(
        "admin/edit_attendance.html",
        attendance=attendance
    )

@app.route("/delete_attendance/<int:id>")
def delete_attendance(id):

    connection = get_db_connection()
    cursor = connection.cursor()

    cursor.execute(
        "DELETE FROM attendance WHERE id=%s",
        (id,)
    )

    connection.commit()

    cursor.close()
    connection.close()

    return redirect(url_for("attendance_admin"))

@app.route("/export_attendance")
def export_attendance():

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    cursor.execute("SELECT * FROM attendance")

    attendance = cursor.fetchall()

    cursor.close()
    connection.close()

    def generate():

        data = csv.writer(open("attendance.csv", "w", newline=""))

        yield "Student Name,Class,Date,Status\n"

        for row in attendance:

            yield f"{row['student_name']},{row['class_name']},{row['attendance_date']},{row['status']}\n"

    return Response(
        generate(),
        mimetype="text/csv",
        headers={
            "Content-Disposition":"attachment; filename=attendance.csv"
        }
    )

@app.route("/reports_admin")
def reports_admin():

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    # Dashboard Counts
    cursor.execute("SELECT COUNT(*) AS total FROM students")
    total_students = cursor.fetchone()["total"]

    cursor.execute("SELECT COUNT(*) AS total FROM teachers")
    total_teachers = cursor.fetchone()["total"]

    cursor.execute("SELECT COUNT(*) AS total FROM subjects")
    total_subjects = cursor.fetchone()["total"]

    cursor.execute("SELECT COUNT(*) AS total FROM classes")
    total_classes = cursor.fetchone()["total"]

    # Reports List
    cursor.execute("""
        SELECT *
        FROM reports
        ORDER BY generated_date DESC
    """)
    reports = cursor.fetchall()

    cursor.close()
    connection.close()

    return render_template(
        "admin/reports.html",
        reports=reports,
        total_students=total_students,
        total_teachers=total_teachers,
        total_subjects=total_subjects,
        total_classes=total_classes
    )

@app.route("/add_report", methods=["GET", "POST"])
def add_report():

    if request.method == "POST":

        report_name = request.form["report_name"]
        report_type = request.form["report_type"]
        generated_date = request.form["generated_date"]
        generated_by = request.form["generated_by"]

        connection = get_db_connection()
        cursor = connection.cursor()

        cursor.execute("""
            INSERT INTO reports
            (
                report_name,
                report_type,
                generated_date,
                generated_by
            )
            VALUES(%s,%s,%s,%s)
        """,(
            report_name,
            report_type,
            generated_date,
            generated_by
        ))

        connection.commit()

        cursor.close()
        connection.close()

        return redirect(url_for("reports_admin"))

    return render_template("admin/add_report.html")

@app.route("/delete_report/<int:id>")
def delete_report():

    connection = get_db_connection()
    cursor = connection.cursor()

    cursor.execute(
        "DELETE FROM reports WHERE id=%s",
        (id,)
    )

    connection.commit()

    cursor.close()
    connection.close()

    return redirect(url_for("reports_admin"))

from openpyxl import Workbook
from flask import send_file

@app.route("/export_reports")
def export_reports():

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    cursor.execute("""
        SELECT
        report_name,
        report_type,
        generated_date,
        generated_by
        FROM reports
        ORDER BY generated_date DESC
    """)

    reports = cursor.fetchall()

    cursor.close()
    connection.close()

    wb = Workbook()

    ws = wb.active
    ws.title = "Reports"

    ws.append([
        "Report Name",
        "Report Type",
        "Generated Date",
        "Generated By"
    ])

    for report in reports:

        ws.append([
            report["report_name"],
            report["report_type"],
            str(report["generated_date"]),
            report["generated_by"]
        ])

    filename = "CampusSync_Reports.xlsx"

    wb.save(filename)

    return send_file(
        filename,
        as_attachment=True
    )

@app.route("/settings_admin", methods=["GET", "POST"])
def settings_admin():

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    # Get settings
    cursor.execute("SELECT * FROM settings LIMIT 1")
    settings = cursor.fetchone()

    # Get admin user
    cursor.execute("SELECT * FROM users WHERE role='admin' LIMIT 1")
    admin = cursor.fetchone()

    if request.method == "POST":

        campus_name = request.form["campus_name"]
        academic_year = request.form["academic_year"]
        session_timeout = request.form["session_timeout"]

        admin_name = request.form["admin_name"]
        admin_email = request.form["admin_email"]
        password = request.form["password"]

        # Update settings table
        cursor.execute("""
            UPDATE settings
            SET
            campus_name=%s,
            academic_year=%s,
            session_timeout=%s
            WHERE id=1
        """,
        (
            campus_name,
            academic_year,
            session_timeout
        ))

        # Update admin profile
        if password.strip() != "":

            password_hash = generate_password_hash(password)

            cursor.execute("""
                UPDATE users
                SET
                name=%s,
                email=%s,
                password_hash=%s
                WHERE id=%s
            """,
            (
                admin_name,
                admin_email,
                password_hash,
                admin["id"]
            ))

        else:

            cursor.execute("""
                UPDATE users
                SET
                name=%s,
                email=%s
                WHERE id=%s
            """,
            (
                admin_name,
                admin_email,
                admin["id"]
            ))

        connection.commit()

        cursor.execute("SELECT * FROM settings LIMIT 1")
        settings = cursor.fetchone()

        cursor.execute("SELECT * FROM users WHERE role='admin' LIMIT 1")
        admin = cursor.fetchone()

    cursor.close()
    connection.close()

    return render_template(
        "admin/settings.html",
        settings=settings,
        admin=admin
    )

@app.route("/backup_database")
def backup_database():

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    cursor.execute("SELECT * FROM students")
    students = cursor.fetchall()

    cursor.close()
    connection.close()

    filename = "backup.txt"

    with open(filename, "w") as f:
        f.write("CampusSync Backup\n\n")

        for student in students:
            f.write(str(student) + "\n")

    return send_file(
        filename,
        as_attachment=True
    )
@app.route("/clear_cache")
def clear_cache():

    flash("Cache Cleared Successfully!", "success")

    return redirect(url_for("settings_admin"))

@app.route("/courses_student")
def courses_student():

    if "email" not in session:
        return redirect(url_for("login"))

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    # Logged-in student
    cursor.execute("""
        SELECT *
        FROM students
        WHERE email=%s
    """, (session["email"],))

    student = cursor.fetchone()

    if not student:
        cursor.close()
        connection.close()
        return "Student not found."

    # Get class id
    cursor.execute("""
        SELECT id
        FROM classes
        WHERE class_name=%s
    """, (student["class_name"],))

    cls = cursor.fetchone()

    # Subjects allocated for the student's class
    cursor.execute("""
        SELECT
            s.id,
            s.name,
            s.code,
            s.credits,
            s.students_enrolled,
            t.name AS teacher_name
        FROM subject_allocation sa
        JOIN subjects s
            ON sa.subject_id = s.id
        JOIN teachers t
            ON sa.teacher_id = t.id
        WHERE sa.class_id=%s
        ORDER BY s.name
    """, (cls["id"],))

    subjects = cursor.fetchall()

    cursor.close()
    connection.close()

    return render_template(
        "student/courses.html",
        student=student,
        subjects=subjects
    )

@app.route("/assignments_student")
def assignments_student():

    if "email" not in session:
        return redirect(url_for("login"))

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    # Logged-in student
    cursor.execute("""
        SELECT *
        FROM students
        WHERE email=%s
    """, (session["email"],))

    student = cursor.fetchone()

    if not student:
        cursor.close()
        connection.close()
        return "Student not found"

    # Assignments for student's class
    cursor.execute("""
        SELECT *
        FROM assignments
        WHERE class_name=%s
        ORDER BY due_date ASC
    """, (student["class_name"],))

    assignments = cursor.fetchall()

    cursor.close()
    connection.close()

    return render_template(
        "student/assignments.html",
        student=student,
        assignments=assignments
    )

@app.route("/student_add_submission", methods=["POST"])
def student_add_submission():
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute(
        """
        INSERT INTO student_submissions (assignment_id, student_name, submitted_on, status, notes)
        VALUES (%s, %s, %s, %s, %s)
        """,
        (
            request.form.get("assignment_id"),
            request.form.get("student_name"),
            request.form.get("submitted_on"),
            request.form.get("status", "Pending"),
            request.form.get("notes", ""),
        ),
    )
    connection.commit()
    cursor.close()
    connection.close()
    return redirect(url_for("assignments_student"))


@app.route("/student_edit_submission/<int:id>", methods=["GET", "POST"])
def student_edit_submission(id):
    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)
    if request.method == "POST":
        cursor.execute(
            """
            UPDATE student_submissions
            SET assignment_id=%s, student_name=%s, submitted_on=%s, status=%s, notes=%s
            WHERE id=%s
            """,
            (
                request.form.get("assignment_id"),
                request.form.get("student_name"),
                request.form.get("submitted_on"),
                request.form.get("status", "Pending"),
                request.form.get("notes", ""),
                id,
            ),
        )
        connection.commit()
        cursor.close()
        connection.close()
        return redirect(url_for("assignments_student"))

    cursor.execute("SELECT * FROM student_submissions WHERE id=%s", (id,))
    submission = cursor.fetchone()
    cursor.execute("SELECT * FROM assignments ORDER BY due_date")
    assignments = cursor.fetchall()
    cursor.execute("SELECT * FROM student_submissions ORDER BY submitted_on DESC")
    submissions = cursor.fetchall()
    cursor.close()
    connection.close()
    return render_template("student/assignments.html", assignments=assignments, submissions=submissions, edit_submission=submission)


@app.route("/student_delete_submission/<int:id>")
def student_delete_submission(id):
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute("DELETE FROM student_submissions WHERE id=%s", (id,))
    connection.commit()
    cursor.close()
    connection.close()
    return redirect(url_for("assignments_student"))

@app.route("/grades_student")
def grades_student():

    if "email" not in session:
        return redirect(url_for("login"))

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    # Logged-in student
    cursor.execute("""
        SELECT *
        FROM students
        WHERE email=%s
    """, (session["email"],))

    student = cursor.fetchone()

    # Student grades
    cursor.execute("""
        SELECT *
        FROM grades
        WHERE student_email=%s
        ORDER BY course
    """, (session["email"],))

    grades = cursor.fetchall()

    # Calculate SGPA
    total_points = 0
    total_credits = 0

    for row in grades:
        total_points += row["points"] * row["credits"]
        total_credits += row["credits"]

    sgpa = round(total_points / total_credits, 2) if total_credits > 0 else 0

    cursor.close()
    connection.close()

    return render_template(
        "student/grades.html",
        student=student,
        grades=grades,
        sgpa=sgpa
    )

@app.route("/timetable_student")
def timetable_student():

    if "email" not in session:
        return redirect(url_for("login"))

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    # Logged-in student
    cursor.execute("""
        SELECT *
        FROM students
        WHERE email=%s
    """, (session["email"],))

    student = cursor.fetchone()

    if student:

        cursor.execute("""
            SELECT *
            FROM timetable
            WHERE class_name=%s
            ORDER BY
            FIELD(day,'Monday','Tuesday','Wednesday','Thursday','Friday'),
            period
        """,(student["class_name"],))

        timetable = cursor.fetchall()

    else:
        timetable=[]

    cursor.close()
    connection.close()

    timetable_dict={}

    for row in timetable:

        key=(row["day"],row["period"])

        timetable_dict[key]=row

    return render_template(
        "student/timetable.html",
        student=student,
        timetable=timetable_dict
    )

@app.route("/notices_student")
def notices_student():

    if "email" not in session:
        return redirect(url_for("login"))

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    cursor.execute(
        "SELECT * FROM students WHERE email=%s",
        (session["email"],)
    )

    student = cursor.fetchone()

    cursor.execute("""
        SELECT *
        FROM notices
        ORDER BY notice_date DESC
    """)

    notices = cursor.fetchall()

    cursor.close()
    connection.close()

    return render_template(
        "student/notices.html",
        student=student,
        notices=notices
    )

@app.route("/settings_student", methods=["GET", "POST"])
def settings_student():

    if "email" not in session:
        return redirect(url_for("login"))

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    cursor.execute(
        "SELECT * FROM students WHERE email=%s",
        (session["email"],)
    )

    student = cursor.fetchone()

    if request.method == "POST":

        name = request.form["name"]
        email = request.form["email"]

        cursor.execute("""
            UPDATE students
            SET name=%s,
                email=%s
            WHERE id=%s
        """,(name,email,student["id"]))

        # Keep session updated
        session["email"] = email
        session["name"] = name

        connection.commit()

        flash("Profile Updated Successfully","success")

        return redirect(url_for("settings_student"))

    cursor.close()
    connection.close()

    return render_template(
        "student/settings.html",
        student=student
    )

@app.route("/student/change_password", methods=["POST"])
def student_change_password():

    if "email" not in session:
        return redirect(url_for("login"))

    current = request.form["current_password"]
    new = request.form["new_password"]

    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)

    cursor.execute(
        "SELECT * FROM users WHERE email=%s",
        (session["email"],)
    )

    user = cursor.fetchone()

    if user and check_password_hash(user["password_hash"], current):

        new_hash = generate_password_hash(new)

        cursor.execute("""
            UPDATE users
            SET password_hash=%s
            WHERE id=%s
        """,(new_hash,user["id"]))

        connection.commit()

        flash("Password Changed Successfully","success")

    else:

        flash("Current Password is Incorrect","danger")

    cursor.close()
    connection.close()

    return redirect(url_for("settings_student"))

@app.route("/student_add_preference", methods=["POST"])
def student_add_preference():
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute(
        """
        INSERT INTO student_preferences (name, value, description)
        VALUES (%s, %s, %s)
        """,
        (
            request.form.get("name"),
            request.form.get("value"),
            request.form.get("description"),
        ),
    )
    connection.commit()
    cursor.close()
    connection.close()
    return redirect(url_for("settings_student"))


@app.route("/student_edit_preference/<int:id>", methods=["GET", "POST"])
def student_edit_preference(id):
    connection = get_db_connection()
    cursor = connection.cursor(dictionary=True)
    if request.method == "POST":
        cursor.execute(
            """
            UPDATE student_preferences
            SET name=%s, value=%s, description=%s
            WHERE id=%s
            """,
            (
                request.form.get("name"),
                request.form.get("value"),
                request.form.get("description"),
                id,
            ),
        )
        connection.commit()
        cursor.close()
        connection.close()
        return redirect(url_for("settings_student"))

    cursor.execute("SELECT * FROM student_preferences WHERE id=%s", (id,))
    preference = cursor.fetchone()
    cursor.execute("SELECT * FROM student_preferences ORDER BY id")
    preferences = cursor.fetchall()
    cursor.close()
    connection.close()
    user = get_first_user_by_role("student") or {}
    return render_template(
        "student/settings.html",
        user_name=user.get("name") if user else "Student",
        preferences=preferences,
        edit_preference=preference,
    )


@app.route("/student_delete_preference/<int:id>")
def student_delete_preference(id):
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute("DELETE FROM student_preferences WHERE id=%s", (id,))
    connection.commit()
    cursor.close()
    connection.close()
    return redirect(url_for("settings_student"))

@app.route("/logout")
def logout():
    return redirect(url_for("login"))

if __name__ == "__main__":
    app.run(debug=True)

"""
نظام تربية الطلاب - Nizam Tarbiyyatud Dullab
Madarasatu Tarbiyyatud Dullab Wa Ta'alimul Qur'an (Unguwar Gwaza, Gusau, Zamfara State)
Single-File Production Backend, Role-Separated RTL Web Portal & Result Engine
"""

import csv
import datetime
import io
import json
import os
import random
import re
import sqlite3
import string
from typing import List, Optional

from fastapi import FastAPI, File, Form, HTTPException, Request, Response, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    from google.cloud import storage as gcs_storage
except ImportError:
    gcs_storage = None

try:
    import psycopg2
except ImportError:
    psycopg2 = None

app = FastAPI(
    title="نظام تربية الطلاب - Nizam Tarbiyyatud Dullab",
    description="مدرسة تربية الطلاب وتعليم القرآن (Unguwar Gwaza, Gusau, Zamfara State)",
    version="2.1.0",
)

DB_FILE = os.environ.get("DB_FILE", "islamiyya_results.db")
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
IS_POSTGRES = bool(
    DATABASE_URL
    and (
        DATABASE_URL.startswith("postgres://")
        or DATABASE_URL.startswith("postgresql://")
    )
)

MASTER_STUDENTS_EXCEL = os.environ.get(
    "MASTER_STUDENTS_EXCEL", "students_master_database.xlsx"
)
GCS_BUCKET_NAME = os.environ.get("GCS_BUCKET_NAME", "")

DB_INTEGRITY_ERRORS = (
    (sqlite3.IntegrityError, psycopg2.IntegrityError)
    if psycopg2 is not None
    else (sqlite3.IntegrityError,)
)


def _normalize_pg_url(url: str) -> str:
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://") :]
    if "sslmode=" not in url and "localhost" not in url and "127.0.0.1" not in url:
        sep = "&" if "?" in url else "?"
        url = f"{url}{sep}sslmode=require"
    return url


def _translate_sql_for_pg(sql: str) -> str:
    """Translates SQLite-flavored SQL (`?` placeholders & `INSERT OR IGNORE`) to PostgreSQL SQL."""
    s = sql.strip()
    if re.search(r"INSERT\s+OR\s+IGNORE\s+INTO", s, flags=re.IGNORECASE):
        s = re.sub(
            r"INSERT\s+OR\s+IGNORE\s+INTO",
            "INSERT INTO",
            s,
            flags=re.IGNORECASE,
        )
        if "ON CONFLICT" not in s.upper():
            s = s.rstrip(";").rstrip() + " ON CONFLICT DO NOTHING"
    s = s.replace("?", "%s")
    return s


class DBCursor:
    """Unified cursor adapter supporting both Cloud PostgreSQL (Neon/Supabase) and local SQLite."""

    def __init__(self, raw_cursor, is_postgres: bool):
        self._cur = raw_cursor
        self.is_postgres = is_postgres

    def execute(self, sql: str, params: tuple = ()):
        if self.is_postgres:
            sql = _translate_sql_for_pg(sql)
        if params:
            return self._cur.execute(sql, tuple(params))
        return self._cur.execute(sql)

    def executemany(self, sql: str, seq_of_params):
        if self.is_postgres:
            sql = _translate_sql_for_pg(sql)
        return self._cur.executemany(sql, seq_of_params)

    def fetchone(self):
        return self._cur.fetchone()

    def fetchall(self):
        return self._cur.fetchall()

    @property
    def lastrowid(self):
        return getattr(self._cur, "lastrowid", None)

    @property
    def rowcount(self):
        return self._cur.rowcount


class DBConnection:
    """Unified connection wrapper for PostgreSQL and SQLite."""

    def __init__(self, raw_conn, is_postgres: bool):
        self._conn = raw_conn
        self.is_postgres = is_postgres

    def cursor(self) -> DBCursor:
        return DBCursor(self._conn.cursor(), self.is_postgres)

    def commit(self):
        return self._conn.commit()

    def rollback(self):
        return self._conn.rollback()

    def close(self):
        return self._conn.close()


def get_db_connection() -> DBConnection:
    if IS_POSTGRES:
        if psycopg2 is None:
            raise RuntimeError(
                "psycopg2-binary is required when DATABASE_URL is set."
            )
        pg_conn = psycopg2.connect(_normalize_pg_url(DATABASE_URL))
        return DBConnection(pg_conn, is_postgres=True)
    sq_conn = sqlite3.connect(DB_FILE)
    sq_conn.execute("PRAGMA foreign_keys = ON")
    return DBConnection(sq_conn, is_postgres=False)


def insert_and_get_id(
    cursor: DBCursor, sql: str, params: tuple = (), pk_col: str = "id"
) -> int:
    """Executes an INSERT statement and returns the newly generated primary key ID on both PG and SQLite."""
    if cursor.is_postgres:
        pg_sql = sql.strip().rstrip(";") + f" RETURNING {pk_col}"
        cursor.execute(pg_sql, params)
        row = cursor.fetchone()
        return int(row[0]) if row else 0
    cursor.execute(sql, params)
    return int(cursor.lastrowid)


# ==========================================
# 0. GOOGLE CLOUD STORAGE ARCHIVING
# ==========================================
def archive_to_gcs(
    file_bytes: bytes, destination_blob_name: str, content_type: str
) -> Optional[str]:
    """Archives generated report cards, broadsheets, and master databases to GCS if configured."""
    if not GCS_BUCKET_NAME or gcs_storage is None:
        return None
    try:
        client = gcs_storage.Client()
        bucket = client.bucket(GCS_BUCKET_NAME)
        blob = bucket.blob(destination_blob_name)
        blob.upload_from_string(file_bytes, content_type=content_type)
        return f"gs://{GCS_BUCKET_NAME}/{destination_blob_name}"
    except Exception:
        return None


def format_td_code(seq_number: int) -> str:
    """Formats a permanent student sequence number as TD0001 (strictly monotonic)."""
    return f"TD{int(seq_number):04d}"


# ==========================================
# 1. DATABASE SETUP, MIGRATION & SEEDING
# ==========================================
def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    auto_pk = (
        "SERIAL PRIMARY KEY"
        if conn.is_postgres
        else "INTEGER PRIMARY KEY AUTOINCREMENT"
    )

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS admin (
            id INTEGER PRIMARY KEY,
            username TEXT UNIQUE,
            password TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS academic_settings (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            active_session TEXT NOT NULL DEFAULT '2026/2027',
            active_term INTEGER NOT NULL DEFAULT 1 CHECK (active_term IN (1, 2, 3)),
            is_open INTEGER NOT NULL DEFAULT 1 CHECK (is_open IN (0, 1))
        )
    """)

    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS classes (
            id {auto_pk},
            name TEXT UNIQUE NOT NULL
        )
    """)

    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS subjects (
            id {auto_pk},
            name_ar TEXT UNIQUE NOT NULL,
            name_en TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS class_subjects (
            class_id INTEGER,
            subject_id INTEGER,
            PRIMARY KEY (class_id, subject_id),
            FOREIGN KEY (class_id) REFERENCES classes(id) ON DELETE CASCADE,
            FOREIGN KEY (subject_id) REFERENCES subjects(id) ON DELETE CASCADE
        )
    """)

    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS teachers (
            id {auto_pk},
            full_name TEXT NOT NULL,
            ref_key TEXT UNIQUE NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS teacher_assignments (
            teacher_id INTEGER,
            class_id INTEGER,
            subject_id INTEGER,
            PRIMARY KEY (teacher_id, class_id, subject_id),
            FOREIGN KEY (teacher_id) REFERENCES teachers(id) ON DELETE CASCADE,
            FOREIGN KEY (class_id) REFERENCES classes(id) ON DELETE CASCADE,
            FOREIGN KEY (subject_id) REFERENCES subjects(id) ON DELETE CASCADE
        )
    """)

    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS students (
            id {auto_pk},
            student_code TEXT UNIQUE,
            full_name TEXT NOT NULL,
            class_id INTEGER,
            status TEXT DEFAULT 'active',
            FOREIGN KEY (class_id) REFERENCES classes(id) ON DELETE SET NULL
        )
    """)

    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS student_index_registry (
            seq_num {auto_pk},
            student_code TEXT UNIQUE NOT NULL,
            student_id INTEGER,
            full_name TEXT NOT NULL,
            class_name TEXT,
            status TEXT DEFAULT 'active',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS scores (
            student_id INTEGER,
            subject_id INTEGER,
            term_id INTEGER,
            session_year TEXT,
            score REAL CHECK(score >= 0 AND score <= 100),
            PRIMARY KEY (student_id, subject_id, term_id, session_year),
            FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE,
            FOREIGN KEY (subject_id) REFERENCES subjects(id) ON DELETE CASCADE
        )
    """)

    if conn.is_postgres:
        cursor.execute(
            "ALTER TABLE students ADD COLUMN IF NOT EXISTS student_code TEXT UNIQUE"
        )
        cursor.execute(
            "ALTER TABLE students ADD COLUMN IF NOT EXISTS status TEXT DEFAULT 'active'"
        )
    else:
        cursor.execute("PRAGMA table_info(students)")
        existing_cols = {row[1] for row in cursor.fetchall()}
        if "student_code" not in existing_cols:
            cursor.execute("ALTER TABLE students ADD COLUMN student_code TEXT")
            cursor.execute(
                "CREATE UNIQUE INDEX IF NOT EXISTS idx_students_code ON students(student_code)"
            )
        if "status" not in existing_cols:
            cursor.execute(
                "ALTER TABLE students ADD COLUMN status TEXT DEFAULT 'active'"
            )

    # Default Admin
    cursor.execute(
        "INSERT OR IGNORE INTO admin (id, username, password) VALUES (1, 'admin', 'admin123')"
    )

    # Default Active Session & Lock State
    cursor.execute("""
        INSERT OR IGNORE INTO academic_settings (id, active_session, active_term, is_open)
        VALUES (1, '2026/2027', 1, 1)
    """)

    conn.commit()
    seed_initial_data(cursor, conn)
    backfill_student_codes(cursor, conn)
    conn.close()
    sync_master_student_excel()


def seed_initial_data(cursor, conn):
    classes_list = [
        "روضة",
        "فصل الأول",
        "فصل الثاني",
        "فصل الثالث",
        "فصل الرابع أ",
        "فصل الرابع ب",
        "فصل الخامس",
        "فصل السادس",
    ]

    for c in classes_list:
        cursor.execute("INSERT OR IGNORE INTO classes (name) VALUES (?)", (c,))

    subjects_list = [
        ("القرآن", "Al-Qur'an"),
        ("الحديث", "Al-Hadith"),
        ("التجويد", "At-Tajweed"),
        ("التوحيد", "At-Tauhid"),
        ("الفقه", "Al-Fiqh"),
        ("العربية", "Al-Arabiyyah"),
        ("السيرة", "As-Seerah"),
        ("النحو", "An-Nahw"),
        ("الأذكار", "Al-Azkar"),
        ("الحروف", "Al-Huruf"),
    ]

    for ar, en in subjects_list:
        cursor.execute(
            "INSERT OR IGNORE INTO subjects (name_ar, name_en) VALUES (?, ?)",
            (ar, en),
        )

    conn.commit()

    cursor.execute("SELECT name_ar, id FROM subjects")
    subj_map = dict(cursor.fetchall())

    cursor.execute("SELECT name, id FROM classes")
    class_map = dict(cursor.fetchall())

    # 1. Senior Classes (4A, 4B, 5, 6) -> 9 Subjects
    senior_subjs = [
        "القرآن",
        "الحديث",
        "التجويد",
        "التوحيد",
        "الفقه",
        "العربية",
        "السيرة",
        "النحو",
        "الأذكار",
    ]
    for c_name in [
        "فصل الرابع أ",
        "فصل الرابع ب",
        "فصل الخامس",
        "فصل السادس",
    ]:
        if c_name in class_map:
            for s_name in senior_subjs:
                if s_name in subj_map:
                    cursor.execute(
                        "INSERT OR IGNORE INTO class_subjects (class_id, subject_id) VALUES (?, ?)",
                        (class_map[c_name], subj_map[s_name]),
                    )

    # 2. Class 3 -> 9 Subjects
    c3_subjs = [
        "القرآن",
        "الحديث",
        "التجويد",
        "التوحيد",
        "الفقه",
        "العربية",
        "السيرة",
        "الحروف",
        "الأذكار",
    ]
    if "فصل الثالث" in class_map:
        for s_name in c3_subjs:
            if s_name in subj_map:
                cursor.execute(
                    "INSERT OR IGNORE INTO class_subjects (class_id, subject_id) VALUES (?, ?)",
                    (class_map["فصل الثالث"], subj_map[s_name]),
                )

    # 3. Rawdah, Class 1 & Class 2 -> 6 Subjects
    junior_subjs = [
        "القرآن",
        "الحديث",
        "التوحيد",
        "الفقه",
        "الحروف",
        "الأذكار",
    ]
    for c_name in ["روضة", "فصل الأول", "فصل الثاني"]:
        if c_name in class_map:
            for s_name in junior_subjs:
                if s_name in subj_map:
                    cursor.execute(
                        "INSERT OR IGNORE INTO class_subjects (class_id, subject_id) VALUES (?, ?)",
                        (class_map[c_name], subj_map[s_name]),
                    )

    conn.commit()


def backfill_student_codes(cursor, conn):
    """Ensures every existing student has a permanent, unique TDxxxx code recorded in the registry."""
    cursor.execute("""
        SELECT st.id, st.student_code, st.full_name, c.name, COALESCE(st.status, 'active')
        FROM students st
        LEFT JOIN classes c ON st.class_id = c.id
        ORDER BY st.id ASC
    """)
    rows = cursor.fetchall()
    for st_id, st_code, full_name, class_name, status in rows:
        if not st_code:
            st_code = format_td_code(st_id)
            cursor.execute(
                "UPDATE students SET student_code = ?, status = ? WHERE id = ?",
                (st_code, status, st_id),
            )
        cursor.execute(
            """
            INSERT OR IGNORE INTO student_index_registry
                (seq_num, student_code, student_id, full_name, class_name, status)
            VALUES (?, ?, ?, ?, ?, ?)
        """,
            (st_id, st_code, st_id, full_name, class_name or "", status),
        )
    conn.commit()


def create_student_with_permanent_td(
    cursor, full_name: str, class_id: int, class_name: str = ""
) -> tuple:
    """Allocates a strictly monotonic, never-reused TD index and inserts the student."""
    cursor.execute("SELECT COALESCE(MAX(seq_num), 0) FROM student_index_registry")
    max_reg = cursor.fetchone()[0] or 0
    cursor.execute("SELECT COALESCE(MAX(id), 0) FROM students")
    max_st = cursor.fetchone()[0] or 0
    next_num = max(max_reg, max_st) + 1

    while True:
        td_code = format_td_code(next_num)
        cursor.execute(
            "SELECT 1 FROM student_index_registry WHERE student_code = ? UNION SELECT 1 FROM students WHERE student_code = ?",
            (td_code, td_code),
        )
        if not cursor.fetchone():
            break
        next_num += 1

    cursor.execute(
        """
        INSERT INTO students (id, student_code, full_name, class_id, status)
        VALUES (?, ?, ?, ?, 'active')
    """,
        (next_num, td_code, full_name, class_id),
    )
    cursor.execute(
        """
        INSERT INTO student_index_registry
            (seq_num, student_code, student_id, full_name, class_name, status)
        VALUES (?, ?, ?, ?, ?, 'active')
    """,
        (next_num, td_code, next_num, full_name, class_name),
    )
    return next_num, td_code


def sync_master_student_excel() -> bytes:
    """Generates and updates the Master Student Database Excel workbook (`students_master_database.xlsx`).

    - Sheet 1: Complete Permanent Registry (All TD numbers ever issued, active & former)
    - One sheet per Class: All active students in that class ordered by TD number.
    """
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT id, name FROM classes ORDER BY id")
    classes = cursor.fetchall()

    cursor.execute("""
        SELECT st.id, st.student_code, st.full_name, st.class_id, COALESCE(c.name, ''), COALESCE(st.status, 'active')
        FROM students st
        LEFT JOIN classes c ON st.class_id = c.id
        ORDER BY st.id ASC
    """)
    all_active_students = cursor.fetchall()

    cursor.execute("""
        SELECT seq_num, student_code, full_name, COALESCE(class_name, ''), COALESCE(status, 'active'), COALESCE(created_at, '')
        FROM student_index_registry
        ORDER BY seq_num ASC
    """)
    registry_rows = cursor.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    header_fill = PatternFill(
        start_color="1A4D2E", end_color="1A4D2E", fill_type="solid"
    )
    sub_fill = PatternFill(
        start_color="E8F5E9", end_color="E8F5E9", fill_type="solid"
    )
    thin_border = Border(
        left=Side(style="thin", color="1A4D2E"),
        right=Side(style="thin", color="1A4D2E"),
        top=Side(style="thin", color="1A4D2E"),
        bottom=Side(style="thin", color="1A4D2E"),
    )

    # Sheet 1: Complete Permanent Student Registry
    ws_all = wb.active
    ws_all.title = "السجل العام الدائم (TD)"
    ws_all.sheet_view.rightToLeft = True

    ws_all.merge_cells("A1:F1")
    ws_all["A1"] = (
        "مدرسة تربية الطلاب وتعليم القرآن — السجل العام الدائم لجميع أرقام الطلاب (TD)"
    )
    ws_all["A1"].font = Font(name="Calibri", size=14, bold=True, color="1A4D2E")
    ws_all["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_all.append([])

    reg_headers = [
        "م",
        "الرقم الدائم (TD ID)",
        "اسم الطالب الكامل",
        "الفصل",
        "الحالة",
        "تاريخ التسجيل",
    ]
    ws_all.append(reg_headers)
    for col in range(1, len(reg_headers) + 1):
        c = ws_all.cell(row=3, column=col)
        c.fill = header_fill
        c.font = Font(color="FFFFFF", bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border

    r_idx = 4
    for idx, (_, st_code, fname, cname, status, created_at) in enumerate(
        registry_rows, start=1
    ):
        status_ar = "نشط (Active)" if status == "active" else "غادر المدرسة / غير نشط"
        ws_all.append([idx, st_code, fname, cname, status_ar, created_at])
        for col in range(1, 7):
            cell = ws_all.cell(row=r_idx, column=col)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        r_idx += 1

    ws_all.column_dimensions["A"].width = 8
    ws_all.column_dimensions["B"].width = 18
    ws_all.column_dimensions["C"].width = 32
    ws_all.column_dimensions["D"].width = 20
    ws_all.column_dimensions["E"].width = 22
    ws_all.column_dimensions["F"].width = 22

    # Separate Sheet for Each Class
    used_titles = {ws_all.title}
    for cid, cname in classes:
        clean_title = re.sub(r"[\\/*?:\[\]]", "_", cname)[:28].strip() or f"Class_{cid}"
        if clean_title in used_titles:
            clean_title = f"{clean_title}_{cid}"
        used_titles.add(clean_title)

        ws_cls = wb.create_sheet(title=clean_title)
        ws_cls.sheet_view.rightToLeft = True

        ws_cls.merge_cells("A1:D1")
        ws_cls["A1"] = (
            f"مدرسة تربية الطلاب وتعليم القرآن — قائمة طلاب الفصل: {cname}"
        )
        ws_cls["A1"].font = Font(
            name="Calibri", size=14, bold=True, color="1A4D2E"
        )
        ws_cls["A1"].alignment = Alignment(
            horizontal="center", vertical="center"
        )
        ws_cls.append([])

        cls_headers = ["م", "رقم التسجيل الدائم (TD)", "اسم الطالب الكامل", "الفصل"]
        ws_cls.append(cls_headers)
        for col in range(1, len(cls_headers) + 1):
            c = ws_cls.cell(row=3, column=col)
            c.fill = header_fill
            c.font = Font(color="FFFFFF", bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin_border

        cls_students = [
            s for s in all_active_students if s[3] == cid and s[5] == "active"
        ]
        row_num = 4
        for idx, (st_id, st_code, fname, _, _, _) in enumerate(
            cls_students, start=1
        ):
            ws_cls.append([idx, st_code or format_td_code(st_id), fname, cname])
            for col in range(1, 5):
                cell = ws_cls.cell(row=row_num, column=col)
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )
                cell.border = thin_border
            row_num += 1

        ws_cls.column_dimensions["A"].width = 8
        ws_cls.column_dimensions["B"].width = 22
        ws_cls.column_dimensions["C"].width = 34
        ws_cls.column_dimensions["D"].width = 20

    buf = io.BytesIO()
    wb.save(buf)
    file_bytes = buf.getvalue()

    try:
        with open(MASTER_STUDENTS_EXCEL, "wb") as f:
            f.write(file_bytes)
    except Exception:
        pass

    archive_to_gcs(
        file_bytes,
        "database/students_master_database.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    return file_bytes


init_db()


def get_academic_settings(cursor) -> dict:
    cursor.execute(
        "SELECT active_session, active_term, is_open FROM academic_settings WHERE id = 1"
    )
    row = cursor.fetchone()
    if not row:
        return {
            "active_session": "2026/2027",
            "active_term": 1,
            "is_open": True,
        }
    return {
        "active_session": row[0],
        "active_term": int(row[1]),
        "is_open": bool(row[2]),
    }


# ==========================================
# 2. CALCULATION & RANKING UTILITIES
# ==========================================
def get_remark(score: float) -> str:
    if score >= 80:
        return "ممتاز"
    elif score >= 70:
        return "جيد جدا"
    elif score >= 55:
        return "جميل"
    elif score >= 45:
        return "مقبول"
    else:
        return "راسب"


def calculate_positions(student_totals: list) -> dict:
    """Calculates standard competition rank ('1224' tie algorithm)."""
    sorted_students = sorted(
        student_totals, key=lambda x: x["total"], reverse=True
    )
    positions = {}
    current_rank = 1

    for i, s in enumerate(sorted_students):
        if i > 0 and s["total"] < sorted_students[i - 1]["total"]:
            current_rank = i + 1
        positions[s["student_id"]] = current_rank

    return positions


def generate_ref_key(cursor) -> str:
    while True:
        key = "UST-" + "".join(random.choices(string.digits, k=4))
        cursor.execute("SELECT 1 FROM teachers WHERE ref_key = ?", (key,))
        if not cursor.fetchone():
            return key


# ==========================================
# 3. ADMIN AUTH, SESSION CONTROL & METADATA
# ==========================================
@app.post("/admin/login")
def admin_login(username: str = Form(...), password: str = Form(...)):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, username FROM admin WHERE username = ? AND password = ?",
        (username.strip(), password),
    )
    row = cursor.fetchone()
    conn.close()
    if not row:
        raise HTTPException(
            status_code=401, detail="اسم المستخدم أو كلمة المرور غير صحيحة"
        )
    return {
        "status": "success",
        "role": "admin",
        "admin_id": row[0],
        "username": row[1],
    }


@app.post("/admin/credentials")
def update_admin_credentials(
    current_password: str = Form(...),
    new_username: str = Form(...),
    new_password: str = Form(...),
):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT password FROM admin WHERE id = 1")
    row = cursor.fetchone()
    if not row or row[0] != current_password:
        conn.close()
        raise HTTPException(
            status_code=401, detail="كلمة المرور الحالية غير صحيحة"
        )
    if not new_username.strip() or not new_password:
        conn.close()
        raise HTTPException(
            status_code=400, detail="اسم المستخدم وكلمة المرور الجديدة مطلوبان"
        )
    cursor.execute(
        "UPDATE admin SET username = ?, password = ? WHERE id = 1",
        (new_username.strip(), new_password),
    )
    conn.commit()
    conn.close()
    return {
        "status": "success",
        "message": "تم تحديث بيانات دخول الإدارة بنجاح",
    }


@app.get("/api/session-settings")
def read_session_settings():
    conn = get_db_connection()
    cursor = conn.cursor()
    settings = get_academic_settings(cursor)
    conn.close()
    return settings


@app.post("/admin/session-settings")
def update_session_settings(
    active_session: str = Form(...),
    active_term: int = Form(...),
    is_open: int = Form(...),
):
    """Enables/updates the current session & term, or locks it (is_open=0) after compilation."""
    if active_term not in (1, 2, 3):
        raise HTTPException(
            status_code=400, detail="الفترة الدراسية يجب أن تكون 1 أو 2 أو 3"
        )
    clean_session = active_session.strip()
    if not clean_session:
        raise HTTPException(status_code=400, detail="يرجى إدخال العام الدراسي")

    open_flag = 1 if int(is_open) else 0
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        """
        UPDATE academic_settings
        SET active_session = ?, active_term = ?, is_open = ?
        WHERE id = 1
    """,
        (clean_session, active_term, open_flag),
    )
    conn.commit()
    settings = get_academic_settings(cursor)
    conn.close()

    state_label = (
        "مفتوح لرصد الدرجات"
        if settings["is_open"]
        else "مغلق (مقفل بعد اعتماد النتائج)"
    )
    return {
        "status": "success",
        "message": f"تم تحديث إعدادات الدورة ({clean_session} - الفترة {active_term}): {state_label}",
        "settings": settings,
    }


@app.get("/api/metadata")
def get_metadata():
    """Returns all classes, subjects, class_subjects, teachers, students, and active session settings."""
    conn = get_db_connection()
    cursor = conn.cursor()

    settings = get_academic_settings(cursor)

    cursor.execute("SELECT id, name FROM classes ORDER BY id")
    classes = [{"id": r[0], "name": r[1]} for r in cursor.fetchall()]

    cursor.execute("SELECT id, name_ar, name_en FROM subjects ORDER BY id")
    subjects = [
        {"id": r[0], "name_ar": r[1], "name_en": r[2] or ""}
        for r in cursor.fetchall()
    ]

    cursor.execute("""
        SELECT cs.class_id, s.id, s.name_ar, s.name_en
        FROM class_subjects cs
        JOIN subjects s ON cs.subject_id = s.id
        ORDER BY cs.class_id, s.id
    """)
    class_subjects = {}
    for cid, sid, nar, nen in cursor.fetchall():
        class_subjects.setdefault(str(cid), []).append(
            {"id": sid, "name_ar": nar, "name_en": nen or ""}
        )

    cursor.execute("SELECT id, full_name, ref_key FROM teachers ORDER BY id")
    teachers = [
        {"id": r[0], "full_name": r[1], "ref_key": r[2]}
        for r in cursor.fetchall()
    ]

    cursor.execute("""
        SELECT ta.teacher_id, t.full_name, t.ref_key, c.id, c.name, s.id, s.name_ar
        FROM teacher_assignments ta
        JOIN teachers t ON ta.teacher_id = t.id
        JOIN classes c ON ta.class_id = c.id
        JOIN subjects s ON ta.subject_id = s.id
        ORDER BY ta.teacher_id, c.id, s.id
    """)
    assignments = [
        {
            "teacher_id": r[0],
            "teacher_name": r[1],
            "ref_key": r[2],
            "class_id": r[3],
            "class_name": r[4],
            "subject_id": r[5],
            "subject_name": r[6],
        }
        for r in cursor.fetchall()
    ]

    cursor.execute("""
        SELECT st.id, st.student_code, st.full_name, st.class_id, c.name, COALESCE(st.status, 'active')
        FROM students st
        LEFT JOIN classes c ON st.class_id = c.id
        WHERE COALESCE(st.status, 'active') = 'active'
        ORDER BY st.class_id, st.id
    """)
    students = [
        {
            "id": r[0],
            "student_code": r[1] or format_td_code(r[0]),
            "full_name": r[2],
            "class_id": r[3],
            "class_name": r[4] or "",
            "status": r[5],
        }
        for r in cursor.fetchall()
    ]

    cursor.execute("SELECT COUNT(*) FROM student_index_registry")
    total_ever_registered = cursor.fetchone()[0]

    conn.close()
    return {
        "settings": settings,
        "classes": classes,
        "subjects": subjects,
        "class_subjects": class_subjects,
        "teachers": teachers,
        "assignments": assignments,
        "students": students,
        "total_ever_registered": total_ever_registered,
    }


# ==========================================
# 3A. CLASSES & SUBJECTS MANAGEMENT
# ==========================================
@app.post("/admin/classes")
def add_class(
    name: str = Form(...), subject_ids_json: Optional[str] = Form(None)
):
    clean_name = name.strip()
    if not clean_name:
        raise HTTPException(status_code=400, detail="اسم الفصل مطلوب")

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        class_id = insert_and_get_id(
            cursor, "INSERT INTO classes (name) VALUES (?)", (clean_name,)
        )
        if subject_ids_json:
            try:
                sids = json.loads(subject_ids_json)
                for sid in sids:
                    cursor.execute(
                        "INSERT OR IGNORE INTO class_subjects (class_id, subject_id) VALUES (?, ?)",
                        (class_id, int(sid)),
                    )
            except Exception:
                pass
        conn.commit()
    except DB_INTEGRITY_ERRORS:
        conn.rollback()
        conn.close()
        raise HTTPException(status_code=400, detail="هذا الفصل موجود بالفعل")
    conn.close()
    sync_master_student_excel()
    return {
        "status": "success",
        "class_id": class_id,
        "message": f"تمت إضافة الفصل '{clean_name}' بنجاح",
    }


@app.post("/admin/classes/{class_id}/subjects")
def set_class_subjects_checkboxes(
    class_id: int, subject_ids_json: str = Form(...)
):
    """Updates the subjects mapped to a class via checkboxes."""
    try:
        subject_ids = [int(x) for x in json.loads(subject_ids_json)]
    except Exception:
        raise HTTPException(status_code=400, detail="بيانات المواد غير صالحة")

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM classes WHERE id = ?", (class_id,))
    if not cursor.fetchone():
        conn.close()
        raise HTTPException(status_code=404, detail="الفصل غير موجود")

    cursor.execute("DELETE FROM class_subjects WHERE class_id = ?", (class_id,))
    for sid in subject_ids:
        cursor.execute(
            "INSERT OR IGNORE INTO class_subjects (class_id, subject_id) VALUES (?, ?)",
            (class_id, sid),
        )
    conn.commit()
    conn.close()
    return {
        "status": "success",
        "message": "تم تحديث المواد المقررة للفصل بنجاح",
    }


@app.post("/admin/subjects")
def add_subject(name_ar: str = Form(...), name_en: str = Form("")):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        subject_id = insert_and_get_id(
            cursor,
            "INSERT INTO subjects (name_ar, name_en) VALUES (?, ?)",
            (name_ar.strip(), name_en.strip()),
        )
        conn.commit()
        return {
            "status": "success",
            "subject_id": subject_id,
            "message": f"تمت إضافة المادة '{name_ar.strip()}' بنجاح",
        }
    except DB_INTEGRITY_ERRORS:
        conn.rollback()
        raise HTTPException(status_code=400, detail="هذه المادة موجودة بالفعل")
    finally:
        conn.close()


@app.post("/admin/class-subjects")
def assign_subject_to_class(
    class_id: int = Form(...), subject_id: int = Form(...)
):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "INSERT OR IGNORE INTO class_subjects (class_id, subject_id) VALUES (?, ?)",
        (class_id, subject_id),
    )
    conn.commit()
    conn.close()
    return {"status": "success", "message": "تم ربط المادة بالفصل بنجاح"}


# ==========================================
# 3B. STUDENT MANAGEMENT (PERMANENT TD ID, BULK UPLOAD, MULTI-DELETE, PROMOTION)
# ==========================================
@app.post("/admin/students")
def register_student(full_name: str = Form(...), class_id: int = Form(...)):
    clean_name = full_name.strip()
    if not clean_name:
        raise HTTPException(status_code=400, detail="اسم الطالب مطلوب")

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, name FROM classes WHERE id = ?", (class_id,))
    cls_row = cursor.fetchone()
    if not cls_row:
        conn.close()
        raise HTTPException(status_code=404, detail="الفصل غير موجود")

    student_id, td_code = create_student_with_permanent_td(
        cursor, clean_name, class_id, cls_row[1]
    )
    conn.commit()
    conn.close()
    sync_master_student_excel()
    return {
        "status": "success",
        "student_id": student_id,
        "student_code": td_code,
        "message": f"تم تسجيل الطالب '{clean_name}' بالرقم الدائم ({td_code}) بنجاح",
    }


@app.post("/admin/students/{student_id}/edit")
def edit_student(
    student_id: int, full_name: str = Form(...), class_id: int = Form(...)
):
    """Updates student name or promotes/transfers class while strictly preserving their permanent TD code."""
    clean_name = full_name.strip()
    if not clean_name:
        raise HTTPException(status_code=400, detail="اسم الطالب مطلوب")

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, student_code FROM students WHERE id = ?", (student_id,)
    )
    st_row = cursor.fetchone()
    if not st_row:
        conn.close()
        raise HTTPException(status_code=404, detail="الطالب غير موجود")

    cursor.execute("SELECT id, name FROM classes WHERE id = ?", (class_id,))
    cls_row = cursor.fetchone()
    if not cls_row:
        conn.close()
        raise HTTPException(status_code=404, detail="الفصل غير موجود")

    cursor.execute(
        "UPDATE students SET full_name = ?, class_id = ? WHERE id = ?",
        (clean_name, class_id, student_id),
    )
    cursor.execute(
        "UPDATE student_index_registry SET full_name = ?, class_name = ? WHERE student_id = ?",
        (clean_name, cls_row[1], student_id),
    )
    conn.commit()
    conn.close()
    sync_master_student_excel()
    return {
        "status": "success",
        "student_code": st_row[1],
        "message": f"تم تحديث بيانات الطالب '{clean_name}' ({st_row[1]}) مع الاحتفاظ برقمه الدائم",
    }


@app.post("/admin/students/{student_id}/delete")
def delete_student(student_id: int):
    """Deletes a student from active classes while keeping their TD index permanently reserved in registry."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT full_name, student_code FROM students WHERE id = ?",
        (student_id,),
    )
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="الطالب غير موجود")

    cursor.execute(
        "UPDATE student_index_registry SET status = 'left' WHERE student_id = ?",
        (student_id,),
    )
    cursor.execute("DELETE FROM students WHERE id = ?", (student_id,))
    conn.commit()
    conn.close()
    sync_master_student_excel()
    return {
        "status": "success",
        "message": f"تم حذف الطالب '{row[0]}' (الرقم {row[1]} محجوز باسمه دائماً ولن يُمنح لغيره)",
    }


@app.post("/admin/students/bulk-delete")
def bulk_delete_students(student_ids_json: str = Form(...)):
    """Deletes multiple selected students at once while keeping their TD numbers permanently reserved."""
    try:
        student_ids = [int(x) for x in json.loads(student_ids_json)]
    except Exception:
        raise HTTPException(status_code=400, detail="قائمة الطلاب غير صالحة")

    if not student_ids:
        raise HTTPException(
            status_code=400, detail="يرجى تحديد طالب واحد على الأقل للحذف"
        )

    conn = get_db_connection()
    cursor = conn.cursor()
    placeholders = ",".join("?" for _ in student_ids)
    cursor.execute(
        f"UPDATE student_index_registry SET status = 'left' WHERE student_id IN ({placeholders})",
        student_ids,
    )
    cursor.execute(
        f"DELETE FROM students WHERE id IN ({placeholders})",
        student_ids,
    )
    deleted_count = cursor.rowcount
    conn.commit()
    conn.close()
    sync_master_student_excel()
    return {
        "status": "success",
        "deleted_count": deleted_count,
        "message": f"تم حذف {deleted_count} طالباً دفعة واحدة بنجاح (مع حفظ أرقامهم المرجعية TD في السجل الدائم)",
    }


@app.post("/admin/students/delete-by-class")
def delete_entire_class_students(class_id: int = Form(...)):
    """Deletes all students in a given class at once while preserving their TD numbers in the permanent registry."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM classes WHERE id = ?", (class_id,))
    cls_row = cursor.fetchone()
    if not cls_row:
        conn.close()
        raise HTTPException(status_code=404, detail="الفصل غير موجود")

    cursor.execute("SELECT id FROM students WHERE class_id = ?", (class_id,))
    ids = [r[0] for r in cursor.fetchall()]
    if not ids:
        conn.close()
        return {
            "status": "success",
            "deleted_count": 0,
            "message": f"لا يوجد طلاب في '{cls_row[0]}' لحذفهم",
        }

    placeholders = ",".join("?" for _ in ids)
    cursor.execute(
        f"UPDATE student_index_registry SET status = 'left' WHERE student_id IN ({placeholders})",
        ids,
    )
    cursor.execute("DELETE FROM students WHERE class_id = ?", (class_id,))
    deleted_count = cursor.rowcount
    conn.commit()
    conn.close()
    sync_master_student_excel()
    return {
        "status": "success",
        "deleted_count": deleted_count,
        "message": f"تم حذف جميع طلاب '{cls_row[0]}' ({deleted_count} طالب) دفعة واحدة بنجاح",
    }


@app.post("/admin/students/bulk-promote")
def bulk_promote_students(
    student_ids_json: str = Form(...), target_class_id: int = Form(...)
):
    """Promotes/transfers multiple selected students to a new class while keeping their TD codes unchanged."""
    try:
        student_ids = [int(x) for x in json.loads(student_ids_json)]
    except Exception:
        raise HTTPException(status_code=400, detail="قائمة الطلاب غير صالحة")

    if not student_ids:
        raise HTTPException(
            status_code=400, detail="يرجى تحديد طالب واحد على الأقل للترقية/النقل"
        )

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM classes WHERE id = ?", (target_class_id,))
    cls_row = cursor.fetchone()
    if not cls_row:
        conn.close()
        raise HTTPException(status_code=404, detail="الفصل المستهدف غير موجود")

    placeholders = ",".join("?" for _ in student_ids)
    cursor.execute(
        f"UPDATE students SET class_id = ? WHERE id IN ({placeholders})",
        [target_class_id] + student_ids,
    )
    moved_count = cursor.rowcount
    cursor.execute(
        f"UPDATE student_index_registry SET class_name = ? WHERE student_id IN ({placeholders})",
        [cls_row[0]] + student_ids,
    )
    conn.commit()
    conn.close()
    sync_master_student_excel()
    return {
        "status": "success",
        "moved_count": moved_count,
        "message": f"تم نقل / ترقية {moved_count} طالباً إلى '{cls_row[0]}' مع الاحتفاظ بأرقامهم الدائمة (TD)",
    }


@app.post("/admin/students/bulk")
async def bulk_upload_students(
    default_class_id: Optional[int] = Form(None),
    names_text: Optional[str] = Form(None),
    file: Optional[UploadFile] = File(None),
):
    """Bulk registers students via pasted multi-line text OR uploaded CSV / Excel (.xlsx) file."""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT id, name FROM classes")
    class_rows = cursor.fetchall()
    id_to_class = {r[0]: r[1] for r in class_rows}
    name_to_class_id = {r[1].strip(): r[0] for r in class_rows}

    records_to_insert = []

    if names_text and names_text.strip():
        if not default_class_id or default_class_id not in id_to_class:
            conn.close()
            raise HTTPException(
                status_code=400,
                detail="يرجى اختيار فصل صحيح عند الإدخال النصي الجماعي",
            )
        for raw_line in names_text.splitlines():
            line = raw_line.strip()
            if not line:
                continue
            if "," in line or "\t" in line:
                sep = "\t" if "\t" in line else ","
                parts = [p.strip() for p in line.split(sep, 1)]
                st_name = parts[0]
                cls_token = parts[1] if len(parts) > 1 else ""
                cid = name_to_class_id.get(cls_token)
                if cid is None and cls_token.isdigit():
                    cid = (
                        int(cls_token)
                        if int(cls_token) in id_to_class
                        else None
                    )
                records_to_insert.append((st_name, cid or default_class_id))
            else:
                records_to_insert.append((line, default_class_id))

    if file and file.filename:
        fname = file.filename.lower()
        content = await file.read()

        rows = []
        if fname.endswith(".xlsx"):
            try:
                wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
                ws = wb.active
                for r in ws.iter_rows(values_only=True):
                    if r and any(
                        cell is not None and str(cell).strip() for cell in r
                    ):
                        rows.append([
                            str(cell).strip() if cell is not None else ""
                            for cell in r
                        ])
            except Exception as e:
                conn.close()
                raise HTTPException(
                    status_code=400, detail=f"تعذر قراءة ملف Excel: {e}"
                )
        elif fname.endswith(".csv") or fname.endswith(".txt"):
            text_data = None
            for enc in ("utf-8-sig", "utf-8", "cp1256", "latin-1"):
                try:
                    text_data = content.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            if text_data is None:
                conn.close()
                raise HTTPException(
                    status_code=400, detail="تعذر فك ترميز ملف CSV"
                )
            reader = csv.reader(io.StringIO(text_data))
            for r in reader:
                if r and any(cell.strip() for cell in r):
                    rows.append([cell.strip() for cell in r])
        else:
            conn.close()
            raise HTTPException(
                status_code=400,
                detail="صيغة الملف غير مدعومة. يرجى رفع ملف Excel (.xlsx) أو CSV (.csv)",
            )

        header_tokens = {
            "full_name",
            "name",
            "student_name",
            "student",
            "اسم الطالب",
            "الاسم",
            "اسم",
        }
        for idx, row in enumerate(rows):
            first_col = row[0].strip() if row else ""
            if not first_col:
                continue
            if idx == 0 and first_col.lower() in header_tokens:
                continue

            resolved_cid = None
            if len(row) > 1 and row[1]:
                cls_val = row[1].strip()
                if cls_val in name_to_class_id:
                    resolved_cid = name_to_class_id[cls_val]
                elif cls_val.isdigit() and int(cls_val) in id_to_class:
                    resolved_cid = int(cls_val)

            if resolved_cid is None:
                if default_class_id and default_class_id in id_to_class:
                    resolved_cid = default_class_id
                else:
                    conn.close()
                    raise HTTPException(
                        status_code=400,
                        detail=f"لم يتم تحديد فصل صالح للطالب '{first_col}' في السطر {idx + 1}",
                    )

            records_to_insert.append((first_col, resolved_cid))

    if not records_to_insert:
        conn.close()
        raise HTTPException(
            status_code=400,
            detail="لم يتم العثور على بيانات طلاب للإضافة. أدخل الأسماء أو ارفع ملفاً.",
        )

    assigned_codes = []
    for st_name, cid in records_to_insert:
        _, td_code = create_student_with_permanent_td(
            cursor, st_name, cid, id_to_class.get(cid, "")
        )
        assigned_codes.append(td_code)

    conn.commit()
    conn.close()
    sync_master_student_excel()

    return {
        "status": "success",
        "inserted_count": len(records_to_insert),
        "first_code": assigned_codes[0] if assigned_codes else "",
        "last_code": assigned_codes[-1] if assigned_codes else "",
        "message": f"تم استيراد وتسجيل {len(records_to_insert)} طالباً بنجاح (الأرقام الدائمة: {assigned_codes[0]} إلى {assigned_codes[-1]})",
    }


@app.get("/admin/students/template")
def download_student_upload_template():
    """Generates an RTL Excel template for bulk student upload."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "نموذج رفع الطلاب"
    ws.sheet_view.rightToLeft = True

    ws.append(["اسم الطالب", "الفصل"])
    ws.append(["إبراهيم موسى", "فصل الأول"])
    ws.append(["فاطمة عبد الله", "فصل الثالث"])

    header_fill = PatternFill(
        start_color="1A4D2E", end_color="1A4D2E", fill_type="solid"
    )
    for col in (1, 2):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 22

    buf = io.BytesIO()
    wb.save(buf)
    return Response(
        content=buf.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": 'attachment; filename="Students_Bulk_Upload_Template.xlsx"'
        },
    )


@app.get("/export/students-database")
def export_master_students_database():
    """Downloads the complete, auto-updated multi-sheet Master Student Database Excel file."""
    file_bytes = sync_master_student_excel()
    return Response(
        content=file_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": 'attachment; filename="Nizam_Tarbiyyatud_Dullab_Students_Master_DB.xlsx"'
        },
    )


# ==========================================
# 3C. TEACHER CRUD, CHECKBOX ASSIGNMENT & BULK EXCEL UPLOAD
# ==========================================
@app.post("/admin/register-teacher")
def register_teacher(full_name: str = Form(...)):
    clean_name = full_name.strip()
    if not clean_name:
        raise HTTPException(status_code=400, detail="اسم المعلم مطلوب")

    conn = get_db_connection()
    cursor = conn.cursor()
    ref_key = generate_ref_key(cursor)
    try:
        teacher_id = insert_and_get_id(
            cursor,
            "INSERT INTO teachers (full_name, ref_key) VALUES (?, ?)",
            (clean_name, ref_key),
        )
        conn.commit()
        return {
            "status": "success",
            "teacher_id": teacher_id,
            "message": f"تم تسجيل المعلم {clean_name} بنجاح",
            "ref_key": ref_key,
        }
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()


@app.post("/admin/teachers/{teacher_id}/edit")
def edit_teacher(
    teacher_id: int,
    full_name: str = Form(...),
    regenerate_key: int = Form(0),
):
    clean_name = full_name.strip()
    if not clean_name:
        raise HTTPException(status_code=400, detail="اسم المعلم مطلوب")

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, ref_key FROM teachers WHERE id = ?", (teacher_id,)
    )
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="المعلم غير موجود")

    new_key = generate_ref_key(cursor) if int(regenerate_key) else row[1]
    cursor.execute(
        "UPDATE teachers SET full_name = ?, ref_key = ? WHERE id = ?",
        (clean_name, new_key, teacher_id),
    )
    conn.commit()
    conn.close()
    return {
        "status": "success",
        "teacher_id": teacher_id,
        "ref_key": new_key,
        "message": f"تم تحديث بيانات المعلم '{clean_name}' بنجاح",
    }


@app.post("/admin/teachers/{teacher_id}/delete")
def delete_teacher(teacher_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT full_name FROM teachers WHERE id = ?", (teacher_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="المعلم غير موجود")

    cursor.execute(
        "DELETE FROM teacher_assignments WHERE teacher_id = ?", (teacher_id,)
    )
    cursor.execute("DELETE FROM teachers WHERE id = ?", (teacher_id,))
    conn.commit()
    conn.close()
    return {
        "status": "success",
        "message": f"تم حذف المعلم '{row[0]}' وإلغاء جميع إسناداته بنجاح",
    }


@app.post("/admin/assign-teacher")
def assign_teacher(
    teacher_id: int = Form(...),
    class_id: int = Form(...),
    subject_id: int = Form(...),
):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "INSERT OR IGNORE INTO class_subjects (class_id, subject_id) VALUES (?, ?)",
        (class_id, subject_id),
    )
    cursor.execute(
        """
        INSERT OR IGNORE INTO teacher_assignments (teacher_id, class_id, subject_id)
        VALUES (?, ?, ?)
    """,
        (teacher_id, class_id, subject_id),
    )
    conn.commit()
    conn.close()
    return {"status": "success", "message": "تم تعيين المادة للمعلم بنجاح"}


@app.post("/admin/assign-teacher-checkboxes")
def assign_teacher_checkboxes(
    teacher_id: int = Form(...),
    assignments_json: str = Form(...),
    replace_existing: int = Form(1),
):
    """Saves multiple (class_id, subject_id) checkbox selections for a teacher in a single click.

    assignments_json: JSON list of objects [{"class_id": 1, "subject_id": 2}, ...]
    """
    try:
        pairs = json.loads(assignments_json)
    except Exception:
        raise HTTPException(status_code=400, detail="صيغة الإسنادات غير صالحة")

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT full_name FROM teachers WHERE id = ?", (teacher_id,)
    )
    t_row = cursor.fetchone()
    if not t_row:
        conn.close()
        raise HTTPException(status_code=404, detail="المعلم غير موجود")

    if int(replace_existing):
        cursor.execute(
            "DELETE FROM teacher_assignments WHERE teacher_id = ?",
            (teacher_id,),
        )

    count = 0
    for item in pairs:
        cid = int(item["class_id"])
        sid = int(item["subject_id"])
        cursor.execute(
            "INSERT OR IGNORE INTO class_subjects (class_id, subject_id) VALUES (?, ?)",
            (cid, sid),
        )
        cursor.execute(
            "INSERT OR IGNORE INTO teacher_assignments (teacher_id, class_id, subject_id) VALUES (?, ?, ?)",
            (teacher_id, cid, sid),
        )
        count += 1

    conn.commit()
    conn.close()
    return {
        "status": "success",
        "assigned_count": count,
        "message": f"تم حفظ {count} إسناداً (فصل/مادة) للمعلم '{t_row[0]}' دفعة واحدة بنجاح",
    }


@app.post("/admin/teachers/bulk-upload")
async def bulk_upload_teacher_assignments(
    assignments_text: Optional[str] = Form(None),
    file: Optional[UploadFile] = File(None),
):
    """Bulk uploads teachers and their class/subject assignments via Excel (.xlsx), CSV, or pasted text.

    Columns supported:
      Column 1: Teacher Full Name (اسم المعلم) - auto-registered with UST-XXXX if new!
      Column 2: Class Name (الفصل) - e.g. فصل الأول (can also be multiple classes separated by / or ،)
      Column 3: Subject(s) (المادة / المواد) - e.g. القرآن or multiple separated by ، / , (القرآن، الحديث)
    """
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT id, name FROM classes")
    name_to_class = {r[1].strip(): r[0] for r in cursor.fetchall()}

    cursor.execute("SELECT id, name_ar, COALESCE(name_en, '') FROM subjects")
    subj_lookup = {}
    for sid, nar, nen in cursor.fetchall():
        subj_lookup[nar.strip()] = sid
        if nen.strip():
            subj_lookup[nen.strip().lower()] = sid

    cursor.execute("SELECT id, full_name FROM teachers")
    teacher_lookup = {r[1].strip(): r[0] for r in cursor.fetchall()}

    raw_rows = []

    if assignments_text and assignments_text.strip():
        for line in assignments_text.splitlines():
            line = line.strip()
            if not line:
                continue
            sep = "\t" if "\t" in line else ("," if "," in line else "،")
            parts = [p.strip() for p in line.split(sep)]
            if len(parts) >= 3:
                raw_rows.append([parts[0], parts[1], ",".join(parts[2:])])

    if file and file.filename:
        fname = file.filename.lower()
        content = await file.read()
        if fname.endswith(".xlsx"):
            try:
                wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
                ws = wb.active
                for r in ws.iter_rows(values_only=True):
                    if r and any(
                        cell is not None and str(cell).strip() for cell in r
                    ):
                        raw_rows.append([
                            str(cell).strip() if cell is not None else ""
                            for cell in r
                        ])
            except Exception as e:
                conn.close()
                raise HTTPException(
                    status_code=400, detail=f"تعذر قراءة ملف Excel: {e}"
                )
        elif fname.endswith(".csv") or fname.endswith(".txt"):
            text_data = None
            for enc in ("utf-8-sig", "utf-8", "cp1256", "latin-1"):
                try:
                    text_data = content.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            if text_data is None:
                conn.close()
                raise HTTPException(
                    status_code=400, detail="تعذر فك ترميز ملف CSV"
                )
            for r in csv.reader(io.StringIO(text_data)):
                if r and any(cell.strip() for cell in r):
                    raw_rows.append([cell.strip() for cell in r])
        else:
            conn.close()
            raise HTTPException(
                status_code=400,
                detail="يرجى رفع ملف Excel (.xlsx) أو CSV (.csv)",
            )

    if not raw_rows:
        conn.close()
        raise HTTPException(
            status_code=400,
            detail="لم يتم العثور على بيانات في الملف أو النص المرفوع",
        )

    header_tokens = {
        "teacher",
        "teacher_name",
        "full_name",
        "اسم المعلم",
        "المعلم",
    }
    new_teachers_created = 0
    assignments_added = 0

    for idx, row in enumerate(raw_rows):
        if len(row) < 3:
            continue
        t_name = row[0].strip()
        c_cell = row[1].strip()
        s_cell = row[2].strip()
        if not t_name or not c_cell or not s_cell:
            continue
        if idx == 0 and t_name.lower() in header_tokens:
            continue

        # 1. Resolve or auto-create Teacher
        if t_name not in teacher_lookup:
            ref_key = generate_ref_key(cursor)
            tid = insert_and_get_id(
                cursor,
                "INSERT INTO teachers (full_name, ref_key) VALUES (?, ?)",
                (t_name, ref_key),
            )
            teacher_lookup[t_name] = tid
            new_teachers_created += 1
        else:
            tid = teacher_lookup[t_name]

        # 2. Resolve Class(es)
        class_tokens = [
            x.strip() for x in re.split(r"[,،/|+]", c_cell) if x.strip()
        ]
        resolved_cids = []
        for c_tok in class_tokens:
            if c_tok in name_to_class:
                resolved_cids.append(name_to_class[c_tok])
            else:
                # Auto-create class if it doesn't exist
                cursor.execute(
                    "INSERT OR IGNORE INTO classes (name) VALUES (?)", (c_tok,)
                )
                cursor.execute(
                    "SELECT id FROM classes WHERE name = ?", (c_tok,)
                )
                cid = cursor.fetchone()[0]
                name_to_class[c_tok] = cid
                resolved_cids.append(cid)

        # 3. Resolve Subject(s)
        subj_tokens = [
            x.strip() for x in re.split(r"[,،/|+]", s_cell) if x.strip()
        ]
        resolved_sids = []
        for s_tok in subj_tokens:
            sid = subj_lookup.get(s_tok) or subj_lookup.get(s_tok.lower())
            if sid is None:
                cursor.execute(
                    "INSERT OR IGNORE INTO subjects (name_ar, name_en) VALUES (?, '')",
                    (s_tok,),
                )
                cursor.execute(
                    "SELECT id FROM subjects WHERE name_ar = ?", (s_tok,)
                )
                sid = cursor.fetchone()[0]
                subj_lookup[s_tok] = sid
            resolved_sids.append(sid)

        for cid in resolved_cids:
            for sid in resolved_sids:
                cursor.execute(
                    "INSERT OR IGNORE INTO class_subjects (class_id, subject_id) VALUES (?, ?)",
                    (cid, sid),
                )
                cursor.execute(
                    "INSERT OR IGNORE INTO teacher_assignments (teacher_id, class_id, subject_id) VALUES (?, ?, ?)",
                    (tid, cid, sid),
                )
                assignments_added += 1

    conn.commit()
    conn.close()
    return {
        "status": "success",
        "new_teachers_created": new_teachers_created,
        "assignments_added": assignments_added,
        "message": f"تم استيراد {assignments_added} إسناداً بنجاح (وتسجيل {new_teachers_created} معلماً جديداً تلقائياً)",
    }


@app.get("/admin/teachers/template")
def download_teacher_assignment_template():
    """Generates an RTL Excel template for bulk teacher & subject/class assignment upload."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "نموذج إسناد المعلمين"
    ws.sheet_view.rightToLeft = True

    ws.append(["اسم المعلم", "الفصل", "المواد (يمكن فصل عدة مواد بفاصلة)"])
    ws.append(["الأستاذ محمد إبرهيم", "فصل الرابع ب", "النحو، الفقه"])
    ws.append(["الأستاذ ثالث غربى", "فصل الأول", "القرآن، الحديث"])
    ws.append(["الأستاذ رابع شيخ", "فصل الثالث", "التجويد"])

    header_fill = PatternFill(
        start_color="1A4D2E", end_color="1A4D2E", fill_type="solid"
    )
    for col in (1, 2, 3):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 36

    buf = io.BytesIO()
    wb.save(buf)
    return Response(
        content=buf.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": 'attachment; filename="Teachers_Bulk_Assignments_Template.xlsx"'
        },
    )


@app.post("/admin/reassign-teacher")
def reassign_teacher(
    old_teacher_id: int = Form(...),
    old_class_id: int = Form(...),
    old_subject_id: int = Form(...),
    new_teacher_id: int = Form(...),
    new_class_id: int = Form(...),
    new_subject_id: int = Form(...),
):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        """
        DELETE FROM teacher_assignments
        WHERE teacher_id = ? AND class_id = ? AND subject_id = ?
    """,
        (old_teacher_id, old_class_id, old_subject_id),
    )
    cursor.execute(
        "INSERT OR IGNORE INTO class_subjects (class_id, subject_id) VALUES (?, ?)",
        (new_class_id, new_subject_id),
    )
    cursor.execute(
        """
        INSERT OR IGNORE INTO teacher_assignments (teacher_id, class_id, subject_id)
        VALUES (?, ?, ?)
    """,
        (new_teacher_id, new_class_id, new_subject_id),
    )
    conn.commit()
    conn.close()
    return {
        "status": "success",
        "message": "تمت إعادة تعيين الفصل / المادة بنجاح",
    }


@app.post("/admin/unassign-teacher")
def unassign_teacher(
    teacher_id: int = Form(...),
    class_id: int = Form(...),
    subject_id: int = Form(...),
):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        """
        DELETE FROM teacher_assignments
        WHERE teacher_id = ? AND class_id = ? AND subject_id = ?
    """,
        (teacher_id, class_id, subject_id),
    )
    conn.commit()
    conn.close()
    return {"status": "success", "message": "تم إلغاء إسناد المادة من المعلم"}


# ==========================================
# 3D. TEACHER LOGIN & SCORE ENTRY (SESSION-LOCKED)
# ==========================================
@app.post("/teacher/login")
def teacher_login(ref_key: str = Form(...)):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, full_name, ref_key FROM teachers WHERE UPPER(ref_key) = UPPER(?)",
        (ref_key.strip(),),
    )
    teacher = cursor.fetchone()

    if not teacher:
        conn.close()
        raise HTTPException(
            status_code=401, detail="مفتاح المرجع غير صحيح. تواصل مع الإدارة."
        )

    teacher_id, name, canonical_key = teacher
    cursor.execute(
        """
        SELECT c.id, c.name, s.id, s.name_ar
        FROM teacher_assignments ta
        JOIN classes c ON ta.class_id = c.id
        JOIN subjects s ON ta.subject_id = s.id
        WHERE ta.teacher_id = ?
        ORDER BY c.id, s.id
    """,
        (teacher_id,),
    )
    assignments = cursor.fetchall()
    settings = get_academic_settings(cursor)
    conn.close()

    return {
        "status": "success",
        "role": "teacher",
        "teacher_id": teacher_id,
        "name": name,
        "ref_key": canonical_key,
        "assignments": assignments,
        "assignments_detailed": [
            {
                "class_id": r[0],
                "class_name": r[1],
                "subject_id": r[2],
                "subject_name": r[3],
            }
            for r in assignments
        ],
        "session_settings": settings,
    }


@app.get("/api/scores")
def get_class_subject_scores(
    class_id: int,
    subject_id: int,
    term_id: Optional[int] = None,
    session_year: Optional[str] = None,
):
    conn = get_db_connection()
    cursor = conn.cursor()
    settings = get_academic_settings(cursor)

    effective_session = (
        session_year.strip()
        if session_year and session_year.strip()
        else settings["active_session"]
    )
    effective_term = (
        term_id if term_id in (1, 2, 3) else settings["active_term"]
    )

    cursor.execute(
        """
        SELECT st.id, st.student_code, st.full_name, sc.score
        FROM students st
        LEFT JOIN scores sc
            ON st.id = sc.student_id
            AND sc.subject_id = ?
            AND sc.term_id = ?
            AND sc.session_year = ?
        WHERE st.class_id = ? AND COALESCE(st.status, 'active') = 'active'
        ORDER BY st.id
    """,
        (subject_id, effective_term, effective_session, class_id),
    )
    rows = [
        {
            "student_id": r[0],
            "student_code": r[1] or format_td_code(r[0]),
            "full_name": r[2],
            "score": r[3] if r[3] is not None else "",
            "remark": get_remark(r[3]) if r[3] is not None else "-",
        }
        for r in cursor.fetchall()
    ]
    conn.close()
    return {
        "students": rows,
        "session_year": effective_session,
        "term_id": effective_term,
        "is_open": settings["is_open"],
    }


@app.post("/scores/submit")
def submit_score(
    student_id: int = Form(...),
    subject_id: int = Form(...),
    term_id: Optional[int] = Form(None),
    session_year: Optional[str] = Form(None),
    score: float = Form(...),
    ref_key: Optional[str] = Form(None),
):
    if score < 0 or score > 100:
        raise HTTPException(
            status_code=400, detail="يجب أن تكون الدرجة بين 0 و 100"
        )

    conn = get_db_connection()
    cursor = conn.cursor()
    settings = get_academic_settings(cursor)

    if not settings["is_open"]:
        conn.close()
        raise HTTPException(
            status_code=403,
            detail="تم إغلاق رصد وتعديل الدرجات من قبل الإدارة بعد اعتماد النتائج.",
        )

    effective_session = settings["active_session"]
    effective_term = settings["active_term"]

    if session_year and session_year.strip() != effective_session:
        conn.close()
        raise HTTPException(
            status_code=403,
            detail=f"لا يمكن الرصد إلا في العام الدراسي المفعل حالياً ({effective_session}).",
        )
    if term_id is not None and int(term_id) != effective_term:
        conn.close()
        raise HTTPException(
            status_code=403,
            detail=f"لا يمكن الرصد إلا في الفترة الدراسية المفعلة حالياً (الفترة {effective_term}).",
        )

    cursor.execute("SELECT class_id FROM students WHERE id = ?", (student_id,))
    st_row = cursor.fetchone()
    if not st_row:
        conn.close()
        raise HTTPException(status_code=404, detail="الطالب غير موجود")
    student_class_id = st_row[0]

    if ref_key and ref_key.strip():
        cursor.execute(
            "SELECT id FROM teachers WHERE UPPER(ref_key) = UPPER(?)",
            (ref_key.strip(),),
        )
        t_row = cursor.fetchone()
        if not t_row:
            conn.close()
            raise HTTPException(
                status_code=401, detail="مفتاح المرجع للمعلم غير صالح"
            )
        teacher_id = t_row[0]
        cursor.execute(
            """
            SELECT 1 FROM teacher_assignments
            WHERE teacher_id = ? AND class_id = ? AND subject_id = ?
        """,
            (teacher_id, student_class_id, subject_id),
        )
        if not cursor.fetchone():
            conn.close()
            raise HTTPException(
                status_code=403,
                detail="غير مصرح لك برصد درجات هذه المادة أو هذا الفصل",
            )

    cursor.execute(
        """
        INSERT INTO scores (student_id, subject_id, term_id, session_year, score)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(student_id, subject_id, term_id, session_year)
        DO UPDATE SET score = excluded.score
    """,
        (student_id, subject_id, effective_term, effective_session, score),
    )
    conn.commit()
    conn.close()
    return {
        "status": "success",
        "message": "تم حفظ الدرجة بنجاح",
        "session_year": effective_session,
        "term_id": effective_term,
        "remark": get_remark(score),
    }


@app.post("/scores/submit-batch")
def submit_scores_batch(
    subject_id: int = Form(...),
    scores_json: str = Form(...),
    ref_key: Optional[str] = Form(None),
):
    """Allows a teacher to save all entered scores for a class/subject in one click."""
    try:
        items = json.loads(scores_json)
    except Exception:
        raise HTTPException(status_code=400, detail="بيانات الدرجات غير صالحة")

    conn = get_db_connection()
    cursor = conn.cursor()
    settings = get_academic_settings(cursor)

    if not settings["is_open"]:
        conn.close()
        raise HTTPException(
            status_code=403,
            detail="تم إغلاق رصد وتعديل الدرجات من قبل الإدارة بعد اعتماد النتائج.",
        )

    effective_session = settings["active_session"]
    effective_term = settings["active_term"]

    teacher_id = None
    if ref_key and ref_key.strip():
        cursor.execute(
            "SELECT id FROM teachers WHERE UPPER(ref_key) = UPPER(?)",
            (ref_key.strip(),),
        )
        t_row = cursor.fetchone()
        if not t_row:
            conn.close()
            raise HTTPException(
                status_code=401, detail="مفتاح المرجع للمعلم غير صالح"
            )
        teacher_id = t_row[0]

    saved_count = 0
    for entry in items:
        st_id = int(entry["student_id"])
        sc_val = float(entry["score"])
        if sc_val < 0 or sc_val > 100:
            conn.close()
            raise HTTPException(
                status_code=400, detail="يجب أن تكون جميع الدرجات بين 0 و 100"
            )
        cursor.execute("SELECT class_id FROM students WHERE id = ?", (st_id,))
        st_row = cursor.fetchone()
        if not st_row:
            continue
        if teacher_id is not None:
            cursor.execute(
                "SELECT 1 FROM teacher_assignments WHERE teacher_id = ? AND class_id = ? AND subject_id = ?",
                (teacher_id, st_row[0], subject_id),
            )
            if not cursor.fetchone():
                conn.close()
                raise HTTPException(
                    status_code=403,
                    detail="غير مصرح لك برصد درجات هذه المادة أو هذا الفصل",
                )
        cursor.execute(
            """
            INSERT INTO scores (student_id, subject_id, term_id, session_year, score)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(student_id, subject_id, term_id, session_year)
            DO UPDATE SET score = excluded.score
        """,
            (st_id, subject_id, effective_term, effective_session, sc_val),
        )
        saved_count += 1

    conn.commit()
    conn.close()
    return {
        "status": "success",
        "saved_count": saved_count,
        "message": f"تم حفظ درجات {saved_count} طالباً دفعة واحدة بنجاح",
    }


# ==========================================
# 4. PRINTABLE ARABIC HTML REPORT CARDS (INDIVIDUAL & WHOLE-CLASS BATCH)
# ==========================================
def _build_single_report_card_block(
    student_code: str,
    student_name: str,
    class_name: str,
    term_id: int,
    session_year: str,
    total_students_in_class: int,
    student_pos,
    subject_rows: list,
) -> str:
    max_obtainable = len(subject_rows) * 100
    obtained_total = sum(r[2] for r in subject_rows)
    avg_pct = (
        round((obtained_total / max_obtainable) * 100, 2)
        if max_obtainable > 0
        else 0.0
    )
    overall_remark = get_remark(avg_pct)

    rows_html = ""
    for sub_ar, sub_en, sc, remark in subject_rows:
        label = f"{sub_ar} ({sub_en})" if sub_en else sub_ar
        sc_display = int(sc) if sc == int(sc) else round(sc, 2)
        rows_html += f"""
        <tr>
            <td>{label}</td>
            <td>100</td>
            <td>{sc_display}</td>
            <td>{remark}</td>
        </tr>
        """

    obtained_display = (
        int(obtained_total)
        if obtained_total == int(obtained_total)
        else round(obtained_total, 2)
    )

    return f"""
    <div class="report-card">
        <div class="header">
            <h2>مدرسة تربية الطلاب وتعليم القرآن</h2>
            <p>أنغور غوازا غسو ولاية زمفرا</p>
            <h3>MADARASATU TARBIYYATUD DULLAB WA TA'ALIMUL QUR'AN</h3>
            <p>Unguwar Gwaza, Gusau, Zamfara State</p>
        </div>
        <div class="title-banner">نتيجة الإمتحان - نظام تربية الطلاب</div>
        <div class="info-grid">
            <div class="info-item">رقم الطالب (ID): <span>{student_code}</span></div>
            <div class="info-item">اسم الطالب: <span>{student_name}</span></div>
            <div class="info-item">الفصل: <span>{class_name}</span></div>
            <div class="info-item">الفترة الدراسية: <span>الفترة {term_id} ({session_year})</span></div>
            <div class="info-item">عدد الطلاب: <span>{total_students_in_class}</span></div>
            <div class="info-item">الموقف (الترتيب): <span>{student_pos}</span></div>
        </div>
        <table>
            <thead>
                <tr>
                    <th>المواد المقررة<br>Subject</th>
                    <th>الدرجة المطلوبة<br>Obtainable</th>
                    <th>الدرجة المحصولة<br>Obtained</th>
                    <th>التقدير<br>Remarks</th>
                </tr>
            </thead>
            <tbody>
                {rows_html}
                <tr style="font-weight: bold; background-color: #f9f9f9;">
                    <td>إجمالي (Total)</td>
                    <td>{max_obtainable}</td>
                    <td>{obtained_display}</td>
                    <td>المعدل: {avg_pct}% ({overall_remark})</td>
                </tr>
            </tbody>
        </table>
        <div class="grading-key">
            <strong>مفتاح التقديرات:</strong>
            80–100%: ممتاز | 70–79%: جيد جدا | 55–69%: جميل | 45–54%: مقبول | 0–44%: راسب
        </div>
        <div class="footer-section">
            <p><strong>ملاحظات (Remarks):</strong></p>
            <div class="remarks-box">نوصي بالإهتمام والمواظبة على الدروس.</div>
            <div class="signatures">
                <div>راعي الفصل: _________________</div>
                <div>التوقيع: _________________</div>
                <div>تاريخ العودة: ___ / ___ / 2026م</div>
            </div>
        </div>
    </div>
    """


def _wrap_report_cards_document(title: str, cards_html: str, count: int = 1) -> str:
    return f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>{title}</title>
    <style>
        body {{ font-family: 'Amiri', 'Traditional Arabic', Tahoma, sans-serif; margin: 0; padding: 20px; background: #f4f4f4; color: #000; }}
        .report-card {{ width: 210mm; min-height: 290mm; margin: 0 auto 25px auto; background: #fff; padding: 25px; box-sizing: border-box; border: 3px double #1a4d2e; page-break-after: always; }}
        .report-card:last-child {{ page-break-after: auto; }}
        .header {{ text-align: center; border-bottom: 2px solid #1a4d2e; padding-bottom: 10px; margin-bottom: 15px; }}
        .header h2 {{ margin: 0; color: #1a4d2e; font-size: 26px; }}
        .header h3 {{ margin: 5px 0; font-size: 18px; color: #333; }}
        .header p {{ margin: 3px 0; font-size: 14px; color: #444; }}
        .title-banner {{ background-color: #1a4d2e; color: white; text-align: center; font-size: 20px; font-weight: bold; padding: 6px; margin-bottom: 15px; border-radius: 4px; }}
        .info-grid {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 10px; margin-bottom: 15px; font-size: 15px; font-weight: bold; }}
        .info-item {{ border-bottom: 1px dashed #666; padding-bottom: 4px; }}
        table {{ width: 100%; border-collapse: collapse; margin-bottom: 15px; }}
        th, td {{ border: 1px solid #1a4d2e; padding: 8px; text-align: center; font-size: 15px; }}
        th {{ background-color: #e8f5e9; color: #1a4d2e; }}
        .footer-section {{ margin-top: 20px; font-size: 15px; }}
        .remarks-box {{ border: 1px solid #1a4d2e; padding: 10px; margin-bottom: 15px; min-height: 40px; background: #fafafa; }}
        .grading-key {{ font-size: 13px; color: #333; margin-bottom: 15px; text-align: center; background: #f1f8e9; padding: 6px; border: 1px solid #c5e1a5; border-radius: 4px; }}
        .signatures {{ display: flex; justify-content: space-between; margin-top: 30px; font-weight: bold; }}
        @media print {{
            body {{ background: none; padding: 0; }}
            .report-card {{ border: 2px solid #000; width: 100%; margin: 0; page-break-after: always; }}
            .report-card:last-child {{ page-break-after: auto; }}
            .no-print {{ display: none; }}
        }}
    </style>
</head>
<body>
<div class="no-print" style="text-align: center; margin-bottom: 18px;">
    <button onclick="window.print()" style="padding: 12px 28px; font-size: 16px; font-weight: bold; background: #1a4d2e; color: white; border: none; border-radius: 6px; cursor: pointer;">
        🖨️ طباعة الشهادات ({count} شهادة) / حفظ PDF
    </button>
</div>
{cards_html}
</body>
</html>"""


@app.get("/report-card/{student_id}", response_class=HTMLResponse)
def generate_report_card(
    student_id: int,
    term_id: Optional[int] = None,
    session_year: Optional[str] = None,
):
    """Generates an individual printable RTL HTML report card."""
    conn = get_db_connection()
    cursor = conn.cursor()
    settings = get_academic_settings(cursor)

    effective_session = (
        session_year.strip()
        if session_year and session_year.strip()
        else settings["active_session"]
    )
    effective_term = (
        term_id if term_id in (1, 2, 3) else settings["active_term"]
    )

    cursor.execute(
        """
        SELECT s.student_code, s.full_name, c.id, c.name
        FROM students s JOIN classes c ON s.class_id = c.id
        WHERE s.id = ?
    """,
        (student_id,),
    )
    st_row = cursor.fetchone()
    if not st_row:
        conn.close()
        raise HTTPException(status_code=404, detail="طالب غير موجود")

    student_code, student_name, class_id, class_name = st_row
    student_code = student_code or format_td_code(student_id)

    cursor.execute(
        "SELECT id FROM students WHERE class_id = ? AND COALESCE(status, 'active') = 'active' ORDER BY id",
        (class_id,),
    )
    class_student_ids = [r[0] for r in cursor.fetchall()]
    total_students_in_class = len(class_student_ids)

    cursor.execute(
        """
        SELECT s.id, s.name_ar, s.name_en FROM subjects s
        JOIN class_subjects cs ON s.id = cs.subject_id
        WHERE cs.class_id = ?
        ORDER BY s.id
    """,
        (class_id,),
    )
    class_subjects = cursor.fetchall()

    student_totals = []
    target_student_scores = []

    for sid in class_student_ids:
        st_sum = 0.0
        for sub_id, sub_ar, sub_en in class_subjects:
            cursor.execute(
                """
                SELECT score FROM scores
                WHERE student_id = ? AND subject_id = ? AND term_id = ? AND session_year = ?
            """,
                (sid, sub_id, effective_term, effective_session),
            )
            sc_res = cursor.fetchone()
            sc = float(sc_res[0]) if sc_res and sc_res[0] is not None else 0.0
            st_sum += sc

            if sid == student_id:
                target_student_scores.append(
                    (sub_ar, sub_en or "", sc, get_remark(sc))
                )

        student_totals.append({"student_id": sid, "total": st_sum})

    conn.close()

    positions = calculate_positions(student_totals)
    student_pos = positions.get(student_id, "-")

    card_block = _build_single_report_card_block(
        student_code=student_code,
        student_name=student_name,
        class_name=class_name,
        term_id=effective_term,
        session_year=effective_session,
        total_students_in_class=total_students_in_class,
        student_pos=student_pos,
        subject_rows=target_student_scores,
    )
    html_content = _wrap_report_cards_document(
        f"نظام تربية الطلاب - نتيجة الطالب {student_name} ({student_code})",
        card_block,
        count=1,
    )

    safe_session = effective_session.replace("/", "_")
    archive_to_gcs(
        html_content.encode("utf-8"),
        f"reports/{safe_session}/term_{effective_term}/student_{student_code}.html",
        "text/html; charset=utf-8",
    )
    return html_content


@app.get("/report-cards/class/{class_id}", response_class=HTMLResponse)
def generate_class_report_cards_batch(
    class_id: int,
    term_id: Optional[int] = None,
    session_year: Optional[str] = None,
):
    """Generates a single printable RTL HTML document containing report cards for ALL students in a class."""
    conn = get_db_connection()
    cursor = conn.cursor()
    settings = get_academic_settings(cursor)

    effective_session = (
        session_year.strip()
        if session_year and session_year.strip()
        else settings["active_session"]
    )
    effective_term = (
        term_id if term_id in (1, 2, 3) else settings["active_term"]
    )

    cursor.execute("SELECT name FROM classes WHERE id = ?", (class_id,))
    cls_row = cursor.fetchone()
    if not cls_row:
        conn.close()
        raise HTTPException(status_code=404, detail="الفصل غير موجود")
    class_name = cls_row[0]

    cursor.execute(
        """
        SELECT id, student_code, full_name
        FROM students
        WHERE class_id = ? AND COALESCE(status, 'active') = 'active'
        ORDER BY id
    """,
        (class_id,),
    )
    students = cursor.fetchall()
    if not students:
        conn.close()
        raise HTTPException(
            status_code=404, detail=f"لا يوجد طلاب مسجلون في '{class_name}'"
        )

    cursor.execute(
        """
        SELECT s.id, s.name_ar, s.name_en FROM subjects s
        JOIN class_subjects cs ON s.id = cs.subject_id
        WHERE cs.class_id = ?
        ORDER BY s.id
    """,
        (class_id,),
    )
    class_subjects = cursor.fetchall()

    # Load all scores for this class/term/session in a single query for speed
    cursor.execute(
        """
        SELECT sc.student_id, sc.subject_id, sc.score
        FROM scores sc
        JOIN students st ON sc.student_id = st.id
        WHERE st.class_id = ? AND sc.term_id = ? AND sc.session_year = ?
    """,
        (class_id, effective_term, effective_session),
    )
    score_map = {}
    for sid, sub_id, sc_val in cursor.fetchall():
        score_map[(sid, sub_id)] = float(sc_val) if sc_val is not None else 0.0

    conn.close()

    student_totals = []
    per_student_rows = {}
    for sid, _, _ in students:
        st_sum = 0.0
        rows = []
        for sub_id, sub_ar, sub_en in class_subjects:
            sc = score_map.get((sid, sub_id), 0.0)
            st_sum += sc
            rows.append((sub_ar, sub_en or "", sc, get_remark(sc)))
        per_student_rows[sid] = rows
        student_totals.append({"student_id": sid, "total": st_sum})

    positions = calculate_positions(student_totals)
    total_students = len(students)

    cards_blocks = []
    for sid, st_code, st_name in students:
        cards_blocks.append(
            _build_single_report_card_block(
                student_code=st_code or format_td_code(sid),
                student_name=st_name,
                class_name=class_name,
                term_id=effective_term,
                session_year=effective_session,
                total_students_in_class=total_students,
                student_pos=positions.get(sid, "-"),
                subject_rows=per_student_rows[sid],
            )
        )

    html_content = _wrap_report_cards_document(
        f"شهادات الفصل: {class_name} — الفترة {effective_term} ({effective_session})",
        "\n".join(cards_blocks),
        count=total_students,
    )

    safe_session = effective_session.replace("/", "_")
    archive_to_gcs(
        html_content.encode("utf-8"),
        f"reports/{safe_session}/term_{effective_term}/class_{class_id}_all_cards.html",
        "text/html; charset=utf-8",
    )
    return html_content


# ==========================================
# 5. PER-TERM & ANNUAL EXCEL BROADSHEETS
# ==========================================
@app.get("/export/broadsheet/term")
def export_term_broadsheet_excel(
    class_id: int, term_id: int = 1, session_year: str = "2026/2027"
):
    """Generates Per-Term Broadsheet XLSX with Right-To-Left (RTL) layout and TD IDs."""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT name FROM classes WHERE id = ?", (class_id,))
    class_row = cursor.fetchone()
    if not class_row:
        conn.close()
        raise HTTPException(status_code=404, detail="الفصل غير موجود")
    class_name = class_row[0]

    cursor.execute(
        """
        SELECT s.id, s.name_ar FROM subjects s
        JOIN class_subjects cs ON s.id = cs.subject_id
        WHERE cs.class_id = ?
        ORDER BY s.id
    """,
        (class_id,),
    )
    subjects = cursor.fetchall()

    cursor.execute(
        """
        SELECT id, student_code, full_name
        FROM students
        WHERE class_id = ? AND COALESCE(status, 'active') = 'active'
        ORDER BY id
    """,
        (class_id,),
    )
    students = cursor.fetchall()

    wb = openpyxl.Workbook()
    ws = wb.active
    safe_session = session_year.replace("/", "_")
    ws.title = f"كشف الفترة {term_id} {safe_session}"
    ws.sheet_view.rightToLeft = True

    total_cols = 2 + len(subjects) + 4
    last_col_letter = get_column_letter(max(total_cols, 6))
    ws.merge_cells(f"A1:{last_col_letter}1")
    ws["A1"] = (
        f"مدرسة تربية الطلاب وتعليم القرآن - كشف درجات الفترة {term_id}"
        f" ({session_year}م) - الفصل: {class_name}"
    )
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="1A4D2E")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.append([])

    headers = ["رقم الطالب (TD)", "اسم الطالب"]
    for _, sub_ar in subjects:
        headers.append(f"{sub_ar} (100)")
    headers.extend(["المجموع الكلي", "المعدل (%)", "التقدير", "الترتيب"])
    ws.append(headers)

    header_fill = PatternFill(
        start_color="1A4D2E", end_color="1A4D2E", fill_type="solid"
    )
    thin_border = Border(
        left=Side(style="thin", color="1A4D2E"),
        right=Side(style="thin", color="1A4D2E"),
        top=Side(style="thin", color="1A4D2E"),
        bottom=Side(style="thin", color="1A4D2E"),
    )

    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=3, column=col)
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    student_totals = []
    student_records = {}

    for st_id, st_code, st_name in students:
        subj_scores = {}
        total_score = 0.0
        for sub_id, _ in subjects:
            cursor.execute(
                """
                SELECT score FROM scores
                WHERE student_id = ? AND subject_id = ? AND term_id = ? AND session_year = ?
            """,
                (st_id, sub_id, term_id, session_year),
            )
            r = cursor.fetchone()
            sc = float(r[0]) if r and r[0] is not None else 0.0
            subj_scores[sub_id] = sc
            total_score += sc

        max_possible = len(subjects) * 100
        avg_pct = (
            round((total_score / max_possible) * 100, 2)
            if max_possible > 0
            else 0.0
        )
        student_records[st_id] = {
            "code": st_code or format_td_code(st_id),
            "name": st_name,
            "scores": subj_scores,
            "total": total_score,
            "avg": avg_pct,
            "remark": get_remark(avg_pct),
        }
        student_totals.append({"student_id": st_id, "total": total_score})

    conn.close()

    positions = calculate_positions(student_totals)

    row_idx = 4
    for st_id, _, _ in students:
        rec = student_records[st_id]
        row_data = [rec["code"], rec["name"]]
        for sub_id, _ in subjects:
            row_data.append(rec["scores"][sub_id])
        row_data.extend(
            [rec["total"], f"{rec['avg']}%", rec["remark"], positions[st_id]]
        )
        ws.append(row_data)

        for c in range(1, len(row_data) + 1):
            cell = ws.cell(row=row_idx, column=c)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        row_idx += 1

    buffer = io.BytesIO()
    wb.save(buffer)
    file_bytes = buffer.getvalue()

    filename = f"Term_{term_id}_Broadsheet_Class_{class_id}_{safe_session}.xlsx"
    archive_to_gcs(
        file_bytes,
        f"broadsheets/{safe_session}/term_{term_id}/{filename}",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    return Response(
        content=file_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/export/broadsheet/annual")
def export_annual_broadsheet_excel(
    class_id: int, session_year: str = "2026/2027"
):
    """Generates Annual Cumulative Broadsheet XLSX with Right-To-Left layout and TD IDs."""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT name FROM classes WHERE id = ?", (class_id,))
    class_row = cursor.fetchone()
    if not class_row:
        conn.close()
        raise HTTPException(status_code=404, detail="الفصل غير موجود")
    class_name = class_row[0]

    cursor.execute(
        """
        SELECT s.id, s.name_ar FROM subjects s
        JOIN class_subjects cs ON s.id = cs.subject_id
        WHERE cs.class_id = ?
        ORDER BY s.id
    """,
        (class_id,),
    )
    subjects = cursor.fetchall()

    cursor.execute(
        """
        SELECT id, student_code, full_name
        FROM students
        WHERE class_id = ? AND COALESCE(status, 'active') = 'active'
        ORDER BY id
    """,
        (class_id,),
    )
    students = cursor.fetchall()

    wb = openpyxl.Workbook()
    ws = wb.active
    safe_session = session_year.replace("/", "_")
    ws.title = f"كشف السنوي {safe_session}"
    ws.sheet_view.rightToLeft = True

    total_cols = 2 + len(subjects) * 4 + 4
    last_col_letter = get_column_letter(max(total_cols, 13))
    ws.merge_cells(f"A1:{last_col_letter}1")
    ws["A1"] = (
        f"مدرسة تربية الطلاب وتعليم القرآن - كشف الدرجات السنوي التراكمي"
        f" ({session_year}م) - الفصل: {class_name}"
    )
    ws["A1"].font = Font(name="Calibri", size=15, bold=True, color="1A4D2E")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.append([])

    row3 = ["رقم الطالب (TD)", "اسم الطالب"]
    row4 = ["", ""]

    for sub_id, sub_ar in subjects:
        row3.extend([f"{sub_ar} (300)", "", "", ""])
        row4.extend(["ف1", "ف2", "ف3", "المجموع"])

    row3.extend(
        ["المجموع الكلي", "المعدل السنوي (%)", "التقدير السنوي", "الترتيب السنوي"]
    )
    row4.extend(["", "", "", ""])

    ws.append(row3)
    ws.append(row4)

    start_col = 3
    for _ in subjects:
        ws.merge_cells(
            start_row=3,
            start_column=start_col,
            end_row=3,
            end_column=start_col + 3,
        )
        start_col += 4

    header_fill = PatternFill(
        start_color="1A4D2E", end_color="1A4D2E", fill_type="solid"
    )
    sub_fill = PatternFill(
        start_color="E8F5E9", end_color="E8F5E9", fill_type="solid"
    )
    thin_border = Border(
        left=Side(style="thin", color="1A4D2E"),
        right=Side(style="thin", color="1A4D2E"),
        top=Side(style="thin", color="1A4D2E"),
        bottom=Side(style="thin", color="1A4D2E"),
    )

    for col in range(1, len(row4) + 1):
        c3 = ws.cell(row=3, column=col)
        c3.fill = header_fill
        c3.font = Font(color="FFFFFF", bold=True)
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c3.border = thin_border

        c4 = ws.cell(row=4, column=col)
        c4.fill = sub_fill
        c4.font = Font(color="1A4D2E", bold=True)
        c4.alignment = Alignment(horizontal="center", vertical="center")
        c4.border = thin_border

    student_totals = []
    student_records = {}

    for st_id, st_code, st_name in students:
        st_subjs = {}
        grand_total = 0.0

        for sub_id, _ in subjects:
            term_scores = {}
            subj_sum = 0.0
            for t in (1, 2, 3):
                cursor.execute(
                    """
                    SELECT score FROM scores
                    WHERE student_id = ? AND subject_id = ? AND term_id = ? AND session_year = ?
                """,
                    (st_id, sub_id, t, session_year),
                )
                r = cursor.fetchone()
                sc = float(r[0]) if r and r[0] is not None else 0.0
                term_scores[t] = sc
                subj_sum += sc

            st_subjs[sub_id] = {"t": term_scores, "total": subj_sum}
            grand_total += subj_sum

        max_possible = len(subjects) * 300
        annual_avg = (
            round((grand_total / max_possible) * 100, 2)
            if max_possible > 0
            else 0.0
        )

        student_records[st_id] = {
            "code": st_code or format_td_code(st_id),
            "name": st_name,
            "subjects": st_subjs,
            "grand_total": grand_total,
            "annual_avg": annual_avg,
            "remark": get_remark(annual_avg),
        }
        student_totals.append({"student_id": st_id, "total": grand_total})

    conn.close()

    positions = calculate_positions(student_totals)

    row_idx = 5
    for st_id, _, _ in students:
        rec = student_records[st_id]
        r_data = [rec["code"], rec["name"]]

        for sub_id, _ in subjects:
            sdata = rec["subjects"][sub_id]
            r_data.extend([
                sdata["t"][1],
                sdata["t"][2],
                sdata["t"][3],
                sdata["total"],
            ])

        r_data.extend([
            rec["grand_total"],
            f"{rec['annual_avg']}%",
            rec["remark"],
            positions[st_id],
        ])
        ws.append(r_data)

        for c in range(1, len(r_data) + 1):
            cell = ws.cell(row=row_idx, column=c)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        row_idx += 1

    buffer = io.BytesIO()
    wb.save(buffer)
    file_bytes = buffer.getvalue()

    filename = f"Annual_Broadsheet_Class_{class_id}_{safe_session}.xlsx"
    archive_to_gcs(
        file_bytes,
        f"broadsheets/{safe_session}/annual/{filename}",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    return Response(
        content=file_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ==========================================
# 6. ROLE-SEPARATED RTL ARABIC WEB PORTAL
# ==========================================
@app.get("/", response_class=HTMLResponse)
@app.get("/admin", response_class=HTMLResponse)
@app.get("/teacher", response_class=HTMLResponse)
def web_portal():
    return r"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>نظام تربية الطلاب - مدرسة تربية الطلاب وتعليم القرآن</title>
    <style>
        :root {
            --primary: #1a4d2e;
            --primary-dark: #123520;
            --primary-light: #e8f5e9;
            --accent: #d4af37;
            --danger: #b91c1c;
            --info: #0284c7;
            --bg: #f4f7f5;
            --card: #ffffff;
            --border: #c8e6c9;
        }
        * { box-sizing: border-box; font-family: 'Segoe UI', Tahoma, 'Traditional Arabic', sans-serif; }
        body { margin: 0; background: var(--bg); color: #1f2937; }
        header {
            background: linear-gradient(135deg, #1a4d2e, #2e7d32);
            color: #fff;
            padding: 16px 24px;
            text-align: center;
            border-bottom: 4px solid var(--accent);
            position: relative;
        }
        header h1 { margin: 0; font-size: 25px; }
        header p { margin: 4px 0 0; font-size: 14px; opacity: 0.92; }
        .logout-btn {
            position: absolute;
            left: 20px;
            top: 50%;
            transform: translateY(-50%);
            background: rgba(255,255,255,0.16);
            color: #fff;
            border: 1px solid rgba(255,255,255,0.4);
            padding: 7px 14px;
            border-radius: 6px;
            cursor: pointer;
            font-weight: bold;
            font-size: 13px;
            display: none;
        }
        .logout-btn:hover { background: rgba(255,255,255,0.28); }
        .session-banner {
            background: #fff;
            border-bottom: 2px solid var(--border);
            padding: 9px 20px;
            text-align: center;
            font-weight: bold;
            font-size: 14px;
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 14px;
            flex-wrap: wrap;
        }
        .status-pill {
            padding: 4px 12px;
            border-radius: 999px;
            font-size: 13px;
            color: #fff;
        }
        .status-open { background: #15803d; }
        .status-closed { background: #b91c1c; }
        .container { max-width: 1280px; margin: 20px auto; padding: 0 16px; }
        .auth-wrapper {
            max-width: 820px;
            margin: 35px auto;
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(330px, 1fr));
            gap: 22px;
        }
        .tabs { display: flex; gap: 10px; margin-bottom: 20px; flex-wrap: wrap; }
        .tab-btn {
            padding: 10px 18px;
            border: 1px solid var(--primary);
            background: #fff;
            color: var(--primary);
            font-weight: bold;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
        }
        .tab-btn.active { background: var(--primary); color: #fff; }
        .panel { display: none; }
        .panel.active { display: block; }
        .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(350px, 1fr)); gap: 18px; }
        .card {
            background: var(--card);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 18px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.04);
            margin-bottom: 18px;
        }
        .card h3 {
            margin-top: 0;
            color: var(--primary);
            border-bottom: 2px solid var(--primary-light);
            padding-bottom: 8px;
            font-size: 17px;
        }
        label { display: block; margin: 8px 0 4px; font-weight: 600; font-size: 14px; }
        input[type="text"], input[type="password"], input[type="number"], input[type="file"], select, textarea {
            width: 100%;
            padding: 9px 10px;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            font-size: 14px;
        }
        button.action {
            margin-top: 12px;
            width: 100%;
            padding: 10px 14px;
            background: var(--primary);
            color: #fff;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: bold;
            cursor: pointer;
        }
        button.action:hover { opacity: 0.93; }
        button.btn-sm {
            padding: 5px 10px;
            font-size: 12px;
            border-radius: 4px;
            border: none;
            cursor: pointer;
            font-weight: bold;
            margin: 2px;
        }
        .btn-edit { background: var(--info); color: #fff; }
        .btn-danger { background: var(--danger); color: #fff; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 14px; }
        th, td { border: 1px solid #cbd5e1; padding: 8px; text-align: center; }
        th { background: var(--primary-light); color: var(--primary); }
        .badge {
            display: inline-block;
            background: var(--primary-light);
            color: var(--primary);
            padding: 3px 8px;
            border-radius: 4px;
            font-weight: bold;
            font-family: monospace;
        }
        .td-badge {
            display: inline-block;
            background: #fef3c7;
            color: #92400e;
            border: 1px solid #fde68a;
            padding: 2px 8px;
            border-radius: 4px;
            font-weight: bold;
            font-family: monospace;
        }
        .checkbox-matrix {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(230px, 1fr));
            gap: 12px;
            max-height: 400px;
            overflow-y: auto;
            padding: 6px;
            border: 1px solid #e2e8f0;
            border-radius: 6px;
            background: #f8fafc;
        }
        .matrix-class-box {
            background: #fff;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            padding: 10px;
        }
        .matrix-class-header {
            font-weight: bold;
            color: var(--primary);
            border-bottom: 1px solid #e2e8f0;
            padding-bottom: 6px;
            margin-bottom: 6px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .chk-item {
            display: flex;
            align-items: center;
            gap: 6px;
            font-size: 13px;
            margin: 4px 0;
            cursor: pointer;
            font-weight: normal;
        }
        .chk-item input[type="checkbox"] { width: 16px; height: 16px; cursor: pointer; }
        .bulk-toolbar {
            display: flex;
            gap: 10px;
            align-items: center;
            flex-wrap: wrap;
            background: #f1f5f9;
            padding: 10px 12px;
            border-radius: 6px;
            margin: 10px 0;
            border: 1px solid #cbd5e1;
        }
        .toast {
            position: fixed;
            bottom: 20px;
            left: 20px;
            background: #1a4d2e;
            color: #fff;
            padding: 12px 20px;
            border-radius: 6px;
            display: none;
            z-index: 1000;
            box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        }
        .locked-alert {
            background: #fef2f2;
            border: 1px solid #fecaca;
            color: #991b1b;
            padding: 12px;
            border-radius: 6px;
            margin-bottom: 14px;
            font-weight: bold;
            text-align: center;
            display: none;
        }
    </style>
</head>
<body>
<header>
    <button id="logout-btn" class="logout-btn" onclick="logoutPortal()">🚪 تسجيل الخروج (Logout)</button>
    <h1>نظام تربية الطلاب (Nizam Tarbiyyatud Dullab)</h1>
    <p>مدرسة تربية الطلاب وتعليم القرآن — Unguwar Gwaza, Gusau, Zamfara State</p>
</header>

<div class="session-banner">
    <span>الدورة الدراسية الحالية للنظام: <strong id="banner-session">2026/2027</strong> — <strong id="banner-term">الفترة 1</strong></span>
    <span id="banner-status" class="status-pill status-open">مفتوح لرصد الدرجات</span>
</div>

<div class="container">
    <!-- ========================================== -->
    <!-- 0. LOGIN GATE (SEPARATES ADMIN & TEACHER)  -->
    <!-- ========================================== -->
    <div id="view-login">
        <div class="auth-wrapper">
            <!-- Teacher Login Card -->
            <div class="card" style="border-top: 5px solid var(--primary);">
                <h3>👨‍🏫 بوابة المعلمين (Teacher Portal)</h3>
                <p style="font-size:13px; color:#4b5563;">
                    خاص بالمعلمين فقط لرصد درجات الفصول والمواد المسندة إليهم في الدورة المفعلة.
                </p>
                <form onsubmit="handleTeacherLogin(event)">
                    <label>مفتاح المرجع للمعلم (Teacher Reference Key):</label>
                    <input type="text" id="login-ref-key" placeholder="مثال: UST-1806" required>
                    <button type="submit" class="action">دخول بوابة المعلم</button>
                </form>
            </div>

            <!-- Admin Login Card -->
            <div class="card" style="border-top: 5px solid var(--accent);">
                <h3>🛡️ بوابة الإدارة (Admin Portal)</h3>
                <p style="font-size:13px; color:#4b5563;">
                    خاص بالإدارة لإعداد الدورة والفصول، إسناد المواد، إدارة الطلاب، واستخراج الشهادات والكشوفات.
                </p>
                <form onsubmit="handleAdminLogin(event)">
                    <label>اسم مستخدم الإدارة:</label>
                    <input type="text" id="admin-user" placeholder="admin" required>
                    <label>كلمة المرور:</label>
                    <input type="password" id="admin-pass" placeholder="••••••••" required>
                    <button type="submit" class="action" style="background:#123520;">دخول لوحة الإدارة</button>
                </form>
            </div>
        </div>
    </div>

    <!-- ========================================== -->
    <!-- 1. TEACHER-ONLY WORKSPACE                  -->
    <!-- ========================================== -->
    <div id="view-teacher" style="display:none;">
        <div class="card">
            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                <div>
                    <h3 style="margin:0; border:none; padding:0;">👨‍🏫 بوابة رصد درجات المعلم: <span id="t-portal-name"></span> (<span id="t-portal-key" class="badge"></span>)</h3>
                </div>
                <button class="action" style="width:auto; margin:0; background:#0284c7;" onclick="saveAllTeacherScores()">💾 حفظ جميع الدرجات المدخلة دفعة واحدة</button>
            </div>

            <div id="teacher-locked-banner" class="locked-alert" style="margin-top:12px;">
                🔒 تم إغلاق رصد وتعديل الدرجات لهذه الدورة من قبل الإدارة بعد اعتماد النتائج. يمكنك استعراض الدرجات فقط.
            </div>

            <div class="grid" style="margin-top:12px;">
                <div>
                    <label>اختر الفصل والمادة المسندة إليك:</label>
                    <select id="t-assignment-select" onchange="loadTeacherStudents()"></select>
                </div>
                <div>
                    <label>الفترة الدراسية المفعلة من الإدارة:</label>
                    <input type="text" id="t-term-display" disabled style="background:#f1f5f9; font-weight:bold;">
                </div>
                <div>
                    <label>العام الدراسي المفعل من الإدارة:</label>
                    <input type="text" id="t-session-display" disabled style="background:#f1f5f9; font-weight:bold;">
                </div>
            </div>

            <table>
                <thead>
                    <tr>
                        <th>م</th>
                        <th>الرقم الدائم (TD)</th>
                        <th>اسم الطالب</th>
                        <th>الدرجة (0 - 100)</th>
                        <th>التقدير</th>
                        <th>حفظ فردي</th>
                    </tr>
                </thead>
                <tbody id="score-entry-tbody"></tbody>
            </table>
        </div>
    </div>

    <!-- ========================================== -->
    <!-- 2. ADMIN-ONLY DASHBOARD                    -->
    <!-- ========================================== -->
    <div id="view-admin" style="display:none;">
        <div class="tabs">
            <button class="tab-btn active" onclick="switchAdminTab('teachers', this)">👨‍🏫 المعلمون وإسناد المواد (Checkboxes & Excel)</button>
            <button class="tab-btn" onclick="switchAdminTab('students', this)">👨‍🎓 سجل الطلاب والرفع والحذف الجماعي (TD ID)</button>
            <button class="tab-btn" onclick="switchAdminTab('reports', this)">🖨️ الشهادات الجماعية والكشوفات وقاعدة البيانات</button>
            <button class="tab-btn" onclick="switchAdminTab('settings', this)">⚙️ إعدادات الدورة والفصول والمواد</button>
        </div>

        <!-- TAB A: TEACHERS, CHECKBOX MATRIX ASSIGNMENT & BULK EXCEL UPLOAD -->
        <div id="admin-tab-teachers" class="panel active">
            <div class="grid">
                <div class="card">
                    <h3>1. تسجيل معلم جديد يدوياً (توليد مفتاح UST-XXXX)</h3>
                    <form onsubmit="registerTeacher(event)">
                        <label>اسم المعلم الكامل:</label>
                        <input type="text" id="t-name" placeholder="مثال: الأستاذ عبد الله محمد" required>
                        <button type="submit" class="action">تسجيل المعلم وتوليد المفتاح</button>
                    </form>
                    <div id="new-teacher-key" style="margin-top:10px; font-weight:bold; color:#1a4d2e;"></div>
                </div>

                <div class="card">
                    <h3>2. رفع المعلمين وإسناد الفصول والمواد جماعياً (Excel / CSV)</h3>
                    <form onsubmit="bulkUploadTeachers(event)">
                        <label>خيار 1: ارفع ملف Excel (.xlsx) أو CSV يحتوي على (اسم المعلم | الفصل | المواد):</label>
                        <input type="file" id="t-bulk-file" accept=".xlsx,.csv,.txt">
                        <label>خيار 2: أو الصق السطور مباشرة (اسم المعلم، الفصل، المواد):</label>
                        <textarea id="t-bulk-text" rows="3" placeholder="الأستاذ محمد إبرهيم, فصل الرابع ب, النحو، الفقه"></textarea>
                        <div style="margin-top:6px; font-size:13px;">
                            <a href="/admin/teachers/template" style="color:var(--primary); font-weight:bold;">⬇️ تحميل نموذج Excel لإسناد المعلمين والمواد</a>
                        </div>
                        <button type="submit" class="action">رفع وتطبيق إسنادات المعلمين جماعياً</button>
                    </form>
                </div>
            </div>

            <!-- Checkbox Matrix Assignment Card -->
            <div class="card" style="border: 2px solid var(--primary);">
                <h3>✅ نافذة إسناد الفصول والمواد للمعلم بالمربعات (Checkbox Matrix Assignment)</h3>
                <p style="margin:0 0 10px; font-size:13px; color:#4b5563;">
                    اختر المعلم، ثم حدد جميع الفصول والمواد التي يدرسها دفعة واحدة عبر المربعات (Checkboxes) واضغط حفظ مرة واحدة فقط.
                </p>
                <div style="display:flex; gap:12px; align-items:flex-end; flex-wrap:wrap; margin-bottom:12px;">
                    <div style="flex:1; min-width:260px;">
                        <label>اختر المعلم لتحرير إسناداته:</label>
                        <select id="matrix-teacher-select" onchange="syncTeacherMatrixCheckboxes()"></select>
                    </div>
                    <button type="button" class="btn-sm" style="background:#e2e8f0; color:#1e293b; padding:9px 14px;" onclick="toggleAllMatrixCheckboxes(false)">إلغاء تحديد الكل</button>
                    <button type="button" class="action" style="width:auto; margin:0; padding:10px 24px;" onclick="saveTeacherMatrixAssignments()">💾 حفظ جميع الإسنادات المحددة للمعلم دفعة واحدة</button>
                </div>
                <div id="teacher-checkbox-matrix" class="checkbox-matrix"></div>
            </div>

            <div class="card">
                <h3>قائمة المعلمين ومفاتيح الدخول (Teacher Reference Keys)</h3>
                <table>
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>اسم المعلم</th>
                            <th>مفتاح الدخول (Ref Key)</th>
                            <th>الفصول والمواد المسندة حالياً</th>
                            <th>إجراءات المعلم</th>
                        </tr>
                    </thead>
                    <tbody id="teachers-table-body"></tbody>
                </table>
            </div>
        </div>

        <!-- TAB B: STUDENTS MANAGEMENT (PERMANENT TD ID, BULK UPLOAD, MULTI-DELETE, WHOLE CLASS DELETE, PROMOTION) -->
        <div id="admin-tab-students" class="panel">
            <div class="card" style="background:#f0fdf4; border:2px solid #16a34a;">
                <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px;">
                    <div>
                        <h3 style="margin:0; border:none; padding:0;">📊 قاعدة بيانات الطلاب الكاملة (مقسمة حسب الفصول في ملف Excel واحد)</h3>
                        <p style="margin:4px 0 0; font-size:13px; color:#166534;">
                            يتم تحديث ملف Excel الرئيسي تلقائياً مع كل إضافة أو تعديل أو ترقية. كل طالب يحتفظ برقمه المرجعي الدائم (<strong>TDxxxx</strong>) للأبد ولا يُمنح لأي طالب آخر.
                        </p>
                    </div>
                    <button class="action" style="width:auto; margin:0; padding:10px 22px;" onclick="window.location.href='/export/students-database'">
                        📥 تحميل قاعدة بيانات الطلاب الكاملة (Excel متعدد الأوراق)
                    </button>
                </div>
            </div>

            <div class="grid">
                <div class="card">
                    <h3>➕ تسجيل طالب يدوياً (يرصد له رقم TD دائم تلقائياً)</h3>
                    <form onsubmit="registerStudent(event)">
                        <label>اسم الطالب الكامل:</label>
                        <input type="text" id="st-name" placeholder="مثال: إبراهيم موسى" required>
                        <label>الفصل الدراسي:</label>
                        <select id="st-class-id" required></select>
                        <button type="submit" class="action">تسجيل الطالب</button>
                    </form>
                </div>

                <div class="card">
                    <h3>📥 رفع وتسجيل الطلاب جماعياً (Bulk Upload: Excel / CSV / Text)</h3>
                    <form onsubmit="bulkUploadStudents(event)">
                        <label>الفصل الافتراضي للطلاب المرفوعين:</label>
                        <select id="bulk-class-id" required></select>
                        <label>خيار 1: الصق قائمة أسماء الطلاب (كل طالب في سطر):</label>
                        <textarea id="bulk-names-text" rows="3" placeholder="أحمد محمد&#10;عبد الله إبراهيم"></textarea>
                        <label>خيار 2: أو ارفع ملف Excel (.xlsx) أو CSV (.csv):</label>
                        <input type="file" id="bulk-file" accept=".xlsx,.csv,.txt">
                        <div style="margin-top:6px; font-size:13px;">
                            <a href="/admin/students/template" style="color:var(--primary); font-weight:bold;">⬇️ تحميل نموذج Excel لرفع الطلاب</a>
                        </div>
                        <button type="submit" class="action">رفع وتسجيل الطلاب جماعياً</button>
                    </form>
                </div>
            </div>

            <div class="card">
                <h3>قائمة الطلاب المسجلين (تحديد متعدد / حذف فصل كامل / ترقية فصل مع بقاء رقم TD)</h3>

                <div class="bulk-toolbar">
                    <div style="min-width:210px;">
                        <label style="margin:0 0 3px;">تصفية حسب الفصل:</label>
                        <select id="filter-student-class" onchange="renderStudentsTable()">
                            <option value="0">-- جميع الفصول --</option>
                        </select>
                    </div>

                    <div style="min-width:220px;">
                        <label style="margin:0 0 3px;">بحث بالاسم أو رقم TD:</label>
                        <input type="text" id="search-student-input" placeholder="ابحث عن TD0185 أو اسم الطالب..." oninput="renderStudentsTable()">
                    </div>

                    <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin-top:18px;">
                        <span id="selected-students-count" class="badge">المحدد: 0</span>
                        <button type="button" class="btn-sm btn-danger" style="padding:8px 12px;" onclick="deleteSelectedStudents()">
                            🗑️ حذف الطلاب المحددين
                        </button>
                        <button type="button" class="btn-sm btn-danger" style="padding:8px 12px; background:#7f1d1d;" onclick="deleteWholeClassStudents()">
                            🗑️ حذف جميع طلاب الفصل المصفى بالكامل
                        </button>
                    </div>

                    <div style="display:flex; gap:6px; align-items:center; flex-wrap:wrap; margin-top:18px; border-right:2px solid #cbd5e1; padding-right:10px;">
                        <select id="promote-target-class" style="width:170px;"></select>
                        <button type="button" class="btn-sm btn-edit" style="padding:8px 12px;" onclick="promoteSelectedStudents()">
                            ⬆️ ترقية / نقل المحددين إلى هذا الفصل (مع بقاء رقم TD)
                        </button>
                    </div>
                </div>

                <table>
                    <thead>
                        <tr>
                            <th style="width:42px;"><input type="checkbox" id="select-all-students-chk" onchange="toggleSelectAllStudents(this.checked)"></th>
                            <th>الرقم الدائم (TD ID)</th>
                            <th>اسم الطالب الكامل</th>
                            <th>الفصل الحالي</th>
                            <th>إجراءات فردية</th>
                        </tr>
                    </thead>
                    <tbody id="students-table-body"></tbody>
                </table>
            </div>
        </div>

        <!-- TAB C: REPORTS (WHOLE CLASS & INDIVIDUAL), BROADSHEETS & MASTER DB -->
        <div id="admin-tab-reports" class="panel">
            <div class="grid">
                <!-- Batch Class Report Cards -->
                <div class="card" style="border:2px solid var(--primary);">
                    <h3>🖨️ طباعة شهادات فصل كامل دفعة واحدة (Whole-Class Report Cards)</h3>
                    <p style="font-size:13px; color:#4b5563;">
                        يفتح صفحة طباعة واحدة تحتوي على شهادات جميع طلاب الفصل المختار (كل طالب في صفحة مستقلة جاهزة للطباعة أو الحفظ كـ PDF).
                    </p>
                    <label>اختر الفصل:</label>
                    <select id="batch-rep-class-id"></select>
                    <label>الفترة الدراسية:</label>
                    <select id="batch-rep-term-id">
                        <option value="1">الفترة الأولى (1)</option>
                        <option value="2">الفترة الثانية (2)</option>
                        <option value="3">الفترة الثالثة (3)</option>
                    </select>
                    <label>العام الدراسي:</label>
                    <input type="text" id="batch-rep-session" value="2026/2027">
                    <button class="action" onclick="openWholeClassReportCards()">🖨️ عرض وطباعة جميع شهادات الفصل دفعة واحدة</button>
                </div>

                <!-- Individual Student Report Card -->
                <div class="card">
                    <h3>📄 طباعة شهادة طالب مفرد (Individual Student Report Card)</h3>
                    <p style="font-size:13px; color:#4b5563;">
                        لاستخراج وطباعة شهادة طالب واحد عند الحاجة.
                    </p>
                    <label>اختر الفصل:</label>
                    <select id="rep-class-id" onchange="filterReportStudents()"></select>
                    <label>اختر الطالب (مع رقمه الدائم TD):</label>
                    <select id="rep-student-id"></select>
                    <label>الفترة الدراسية:</label>
                    <select id="rep-term-id">
                        <option value="1">الفترة الأولى (1)</option>
                        <option value="2">الفترة الثانية (2)</option>
                        <option value="3">الفترة الثالثة (3)</option>
                    </select>
                    <label>العام الدراسي:</label>
                    <input type="text" id="rep-session" value="2026/2027">
                    <button class="action" style="background:#0284c7;" onclick="openReportCard()">📄 عرض وطباعة شهادة الطالب المحدد</button>
                </div>

                <!-- Excel Broadsheets & Master DB -->
                <div class="card">
                    <h3>📊 تصدير كشوفات الدرجات وقاعدة البيانات (Excel RTL)</h3>
                    <label>الفصل الدراسي:</label>
                    <select id="bs-class-id"></select>
                    <label>الفترة الدراسية (لكشف الفترة):</label>
                    <select id="bs-term-id">
                        <option value="1">الفترة الأولى (1)</option>
                        <option value="2">الفترة الثانية (2)</option>
                        <option value="3">الفترة الثالثة (3)</option>
                    </select>
                    <label>العام الدراسي:</label>
                    <input type="text" id="bs-session" value="2026/2027">
                    <button class="action" onclick="downloadTermBroadsheet()">📥 تحميل كشف الفترة للفصل (Per-Term XLSX)</button>
                    <button class="action" style="background:#2e7d32;" onclick="downloadAnnualBroadsheet()">📥 تحميل الكشف السنوي التراكمي (Annual XLSX)</button>
                    <button class="action" style="background:#0f766e;" onclick="window.location.href='/export/students-database'">📥 تحميل قاعدة بيانات جميع الطلاب (Excel مقسم بالفصول)</button>
                </div>
            </div>
        </div>

        <!-- TAB D: SESSION CONTROL, CLASSES/ARMS & SUBJECTS -->
        <div id="admin-tab-settings" class="panel">
            <div class="card" style="border: 2px solid var(--primary);">
                <h3>🔒 التحكم في الدورة الدراسية وقفل رصد الدرجات (Session & Compilation Lock)</h3>
                <p style="margin:0 0 10px; font-size:13px; color:#4b5563;">
                    المعلمون يمكنهم فقط الوصول إلى العام الدراسي والفترة المفعلة هنا. عند الانتهاء من تجميع النتائج، قم بإغلاق الدورة لمنع أي تعديل إضافي.
                </p>
                <form onsubmit="saveSessionSettings(event)">
                    <div class="grid">
                        <div>
                            <label>العام الدراسي المفعل (Active Session):</label>
                            <input type="text" id="set-session" value="2026/2027" required>
                        </div>
                        <div>
                            <label>الفترة الدراسية المفعلة (Active Term):</label>
                            <select id="set-term">
                                <option value="1">الفترة الأولى (1)</option>
                                <option value="2">الفترة الثانية (2)</option>
                                <option value="3">الفترة الثالثة (3)</option>
                            </select>
                        </div>
                        <div>
                            <label>حالة رصد الدرجات للمعلمين:</label>
                            <select id="set-is-open">
                                <option value="1">🟢 مفتوح (السماح للمعلمين برصد وتعديل الدرجات)</option>
                                <option value="0">🔴 مغلق / مقفل بعد التجميع (منع أي تعديل)</option>
                            </select>
                        </div>
                    </div>
                    <button type="submit" class="action">حفظ وتطبيق إعدادات الدورة الدراسية</button>
                </form>
            </div>

            <div class="grid">
                <div class="card">
                    <h3>➕ إضافة فصل جديد أو شعبة (Create Class / Arm) مع تحديد المواد</h3>
                    <form onsubmit="addClassWithSubjects(event)">
                        <label>اسم الفصل / الشعبة الجديدة:</label>
                        <input type="text" id="new-class-name" placeholder="مثال: فصل الخامس ب" required>
                        <label>المواد المقررة لهذا الفصل (حدد بالمربعات):</label>
                        <div id="new-class-subjects-box" style="max-height:180px; overflow-y:auto; border:1px solid #cbd5e1; padding:8px; border-radius:6px; background:#f8fafc;"></div>
                        <button type="submit" class="action">إضافة الفصل وربط المواد</button>
                    </form>
                </div>

                <div class="card">
                    <h3>📚 إضافة مادة دراسية جديدة</h3>
                    <form onsubmit="addSubject(event)">
                        <label>اسم المادة بالعربية:</label>
                        <input type="text" id="new-subj-ar" placeholder="مثال: الخط العربي" required>
                        <label>اسم المادة بالإنجليزية (اختياري):</label>
                        <input type="text" id="new-subj-en" placeholder="Calligraphy">
                        <button type="submit" class="action">إضافة المادة</button>
                    </form>

                    <hr style="margin:18px 0; border:none; border-top:1px solid #e2e8f0;">
                    <h3>🔑 تغيير كلمة مرور الإدارة</h3>
                    <form onsubmit="changeAdminPassword(event)">
                        <label>كلمة المرور الحالية:</label>
                        <input type="password" id="adm-old-pass" required>
                        <label>اسم المستخدم الجديد:</label>
                        <input type="text" id="adm-new-user" value="admin" required>
                        <label>كلمة المرور الجديدة:</label>
                        <input type="password" id="adm-new-pass" required>
                        <button type="submit" class="action" style="background:#334155;">تحديث بيانات دخول الإدارة</button>
                    </form>
                </div>
            </div>
        </div>
    </div>
</div>

<div id="toast" class="toast"></div>

<script>
let meta = {
    settings: { active_session: '2026/2027', active_term: 1, is_open: true },
    classes: [],
    subjects: [],
    class_subjects: {},
    teachers: [],
    assignments: [],
    students: []
};
let currentRole = null; // 'admin' or 'teacher'
let activeTeacher = null;
let selectedStudentIds = new Set();

function showToast(msg, isError=false) {
    const t = document.getElementById('toast');
    t.textContent = msg;
    t.style.background = isError ? '#b91c1c' : '#1a4d2e';
    t.style.display = 'block';
    setTimeout(() => { t.style.display = 'none'; }, 4000);
}

function showPortalView(role) {
    currentRole = role;
    document.getElementById('view-login').style.display = role ? 'none' : 'block';
    document.getElementById('view-admin').style.display = role === 'admin' ? 'block' : 'none';
    document.getElementById('view-teacher').style.display = role === 'teacher' ? 'block' : 'none';
    document.getElementById('logout-btn').style.display = role ? 'inline-block' : 'none';
}

function logoutPortal() {
    activeTeacher = null;
    currentRole = null;
    sessionStorage.removeItem('ntd_role');
    sessionStorage.removeItem('ntd_teacher_key');
    showPortalView(null);
}

async function handleAdminLogin(e) {
    e.preventDefault();
    const fd = new FormData();
    fd.append('username', document.getElementById('admin-user').value);
    fd.append('password', document.getElementById('admin-pass').value);
    const res = await fetch('/admin/login', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) {
        showToast(data.detail || 'بيانات الدخول غير صحيحة', true);
        return;
    }
    sessionStorage.setItem('ntd_role', 'admin');
    document.getElementById('admin-pass').value = '';
    await loadMetadata();
    showPortalView('admin');
    showToast('تم تسجيل الدخول إلى لوحة الإدارة بنجاح');
}

async function handleTeacherLogin(e) {
    if (e) e.preventDefault();
    const key = document.getElementById('login-ref-key').value.trim();
    const fd = new FormData();
    fd.append('ref_key', key);
    const res = await fetch('/teacher/login', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) {
        showToast(data.detail || 'مفتاح المعلم غير صحيح', true);
        return;
    }
    activeTeacher = data;
    sessionStorage.setItem('ntd_role', 'teacher');
    sessionStorage.setItem('ntd_teacher_key', data.ref_key);
    if (data.session_settings) {
        meta.settings = data.session_settings;
        updateSessionBanner();
    }
    document.getElementById('t-portal-name').textContent = data.name;
    document.getElementById('t-portal-key').textContent = data.ref_key;
    const sel = document.getElementById('t-assignment-select');
    sel.innerHTML = data.assignments_detailed.map(a =>
        `<option value="${a.class_id}:${a.subject_id}">${a.class_name} — ${a.subject_name}</option>`
    ).join('');
    showPortalView('teacher');
    loadTeacherStudents();
}

function switchAdminTab(name, btn) {
    document.querySelectorAll('#view-admin .panel').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('#view-admin .tab-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('admin-tab-' + name).classList.add('active');
    if (btn) btn.classList.add('active');
}

function updateSessionBanner() {
    const s = meta.settings;
    document.getElementById('banner-session').textContent = s.active_session;
    document.getElementById('banner-term').textContent = `الفترة ${s.active_term}`;
    const pill = document.getElementById('banner-status');
    if (s.is_open) {
        pill.textContent = '🟢 مفتوح لرصد الدرجات';
        pill.className = 'status-pill status-open';
    } else {
        pill.textContent = '🔴 مغلق (مقفل بعد اعتماد النتائج)';
        pill.className = 'status-pill status-closed';
    }

    document.getElementById('set-session').value = s.active_session;
    document.getElementById('set-term').value = String(s.active_term);
    document.getElementById('set-is-open').value = s.is_open ? '1' : '0';

    document.getElementById('batch-rep-term-id').value = String(s.active_term);
    document.getElementById('batch-rep-session').value = s.active_session;
    document.getElementById('rep-term-id').value = String(s.active_term);
    document.getElementById('rep-session').value = s.active_session;
    document.getElementById('bs-term-id').value = String(s.active_term);
    document.getElementById('bs-session').value = s.active_session;

    document.getElementById('t-term-display').value = `الفترة ${s.active_term}`;
    document.getElementById('t-session-display').value = s.active_session;
    document.getElementById('teacher-locked-banner').style.display = s.is_open ? 'none' : 'block';
}

async function loadMetadata() {
    const res = await fetch('/api/metadata');
    meta = await res.json();
    updateSessionBanner();

    const classSelects = ['st-class-id', 'bulk-class-id', 'promote-target-class', 'batch-rep-class-id', 'rep-class-id', 'bs-class-id'];
    classSelects.forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        const val = el.value;
        el.innerHTML = meta.classes.map(c => `<option value="${c.id}">${c.name}</option>`).join('');
        if (val) el.value = val;
    });

    const filterEl = document.getElementById('filter-student-class');
    const fVal = filterEl.value || '0';
    filterEl.innerHTML = '<option value="0">-- جميع الفصول (' + meta.students.length + ' طالب) --</option>' +
        meta.classes.map(c => {
            const cnt = meta.students.filter(s => s.class_id === c.id).length;
            return `<option value="${c.id}">${c.name} (${cnt} طالب)</option>`;
        }).join('');
    filterEl.value = fVal;

    const matrixTeacherEl = document.getElementById('matrix-teacher-select');
    const prevT = matrixTeacherEl.value;
    matrixTeacherEl.innerHTML = meta.teachers.map(t => `<option value="${t.id}">${t.full_name} (${t.ref_key})</option>`).join('');
    if (prevT && meta.teachers.some(t => String(t.id) === String(prevT))) {
        matrixTeacherEl.value = prevT;
    }

    renderTeacherCheckboxMatrix();
    syncTeacherMatrixCheckboxes();

    const newClassSubjBox = document.getElementById('new-class-subjects-box');
    if (newClassSubjBox) {
        newClassSubjBox.innerHTML = meta.subjects.map(s => `
            <label class="chk-item">
                <input type="checkbox" class="new-cls-subj-chk" value="${s.id}" checked>
                <span>${s.name_ar} (${s.name_en})</span>
            </label>
        `).join('');
    }

    const tbody = document.getElementById('teachers-table-body');
    tbody.innerHTML = meta.teachers.map(t => {
        const assigned = meta.assignments
            .filter(a => a.teacher_id === t.id)
            .map(a => `<span style="display:inline-block; background:#f1f5f9; padding:3px 8px; border-radius:4px; margin:2px;">
                ${a.class_name}: ${a.subject_name}
                <button class="btn-sm btn-danger" style="padding:1px 5px;" title="إلغاء هذا الإسناد" onclick="unassignTeacher(${t.id}, ${a.class_id}, ${a.subject_id})">✖</button>
            </span>`)
            .join(' ');
        return `<tr>
            <td>${t.id}</td>
            <td>${t.full_name}</td>
            <td><span class="badge">${t.ref_key}</span></td>
            <td>${assigned || '—'}</td>
            <td>
                <button class="btn-sm btn-edit" onclick="selectTeacherInMatrix(${t.id})">☑️ تحديد مواده</button>
                <button class="btn-sm btn-edit" onclick="promptEditTeacher(${t.id})">✏️ تعديل</button>
                <button class="btn-sm btn-danger" onclick="deleteTeacher(${t.id})">🗑️ حذف</button>
            </td>
        </tr>`;
    }).join('');

    renderStudentsTable();
    filterReportStudents();
}

function renderTeacherCheckboxMatrix() {
    const container = document.getElementById('teacher-checkbox-matrix');
    if (!container) return;
    container.innerHTML = meta.classes.map(c => {
        const clsSubjs = meta.class_subjects[String(c.id)] || meta.subjects;
        const checkboxesHtml = clsSubjs.map(s => `
            <label class="chk-item">
                <input type="checkbox" class="matrix-assign-chk" data-class-id="${c.id}" data-subject-id="${s.id}">
                <span>${s.name_ar}</span>
            </label>
        `).join('');
        return `
            <div class="matrix-class-box">
                <div class="matrix-class-header">
                    <span>🏫 ${c.name}</span>
                    <button type="button" class="btn-sm" style="background:#e8f5e9; color:#1a4d2e;" onclick="toggleClassMatrix(${c.id})">الكل</button>
                </div>
                ${checkboxesHtml}
            </div>
        `;
    }).join('');
}

function syncTeacherMatrixCheckboxes() {
    const tid = parseInt(document.getElementById('matrix-teacher-select').value || '0');
    const assignedSet = new Set(
        meta.assignments
            .filter(a => a.teacher_id === tid)
            .map(a => `${a.class_id}:${a.subject_id}`)
    );
    document.querySelectorAll('.matrix-assign-chk').forEach(chk => {
        const key = `${chk.dataset.classId}:${chk.dataset.subjectId}`;
        chk.checked = assignedSet.has(key);
    });
}

function selectTeacherInMatrix(teacherId) {
    const sel = document.getElementById('matrix-teacher-select');
    sel.value = String(teacherId);
    syncTeacherMatrixCheckboxes();
    sel.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function toggleClassMatrix(classId) {
    const boxes = Array.from(document.querySelectorAll(`.matrix-assign-chk[data-class-id="${classId}"]`));
    const allChecked = boxes.every(b => b.checked);
    boxes.forEach(b => { b.checked = !allChecked; });
}

function toggleAllMatrixCheckboxes(state) {
    document.querySelectorAll('.matrix-assign-chk').forEach(b => { b.checked = state; });
}

async function saveTeacherMatrixAssignments() {
    const tid = document.getElementById('matrix-teacher-select').value;
    if (!tid) {
        showToast('يرجى اختيار معلم أولاً', true);
        return;
    }
    const pairs = [];
    document.querySelectorAll('.matrix-assign-chk:checked').forEach(chk => {
        pairs.push({
            class_id: parseInt(chk.dataset.classId),
            subject_id: parseInt(chk.dataset.subjectId)
        });
    });
    const fd = new FormData();
    fd.append('teacher_id', tid);
    fd.append('assignments_json', JSON.stringify(pairs));
    fd.append('replace_existing', '1');

    const res = await fetch('/admin/assign-teacher-checkboxes', { method: 'POST', body: fd });
    const data = await res.json();
    if (res.ok) {
        showToast(data.message);
        loadMetadata();
    } else {
        showToast(data.detail || 'فشل حفظ الإسنادات', true);
    }
}

async function bulkUploadTeachers(e) {
    e.preventDefault();
    const fd = new FormData();
    const txt = document.getElementById('t-bulk-text').value;
    if (txt.trim()) fd.append('assignments_text', txt);
    const fileInput = document.getElementById('t-bulk-file');
    if (fileInput.files.length > 0) fd.append('file', fileInput.files[0]);

    const res = await fetch('/admin/teachers/bulk-upload', { method: 'POST', body: fd });
    const data = await res.json();
    if (res.ok) {
        document.getElementById('t-bulk-text').value = '';
        fileInput.value = '';
        showToast(data.message);
        loadMetadata();
    } else {
        showToast(data.detail || 'فشل رفع ملف المعلمين', true);
    }
}

function getFilteredStudents() {
    const cid = parseInt(document.getElementById('filter-student-class').value || '0');
    const q = (document.getElementById('search-student-input').value || '').trim().toLowerCase();
    return meta.students.filter(s => {
        if (cid !== 0 && s.class_id !== cid) return false;
        if (q && !s.full_name.toLowerCase().includes(q) && !s.student_code.toLowerCase().includes(q)) return false;
        return true;
    });
}

function updateSelectedStudentsCount() {
    document.getElementById('selected-students-count').textContent = `المحدد: ${selectedStudentIds.size}`;
}

function renderStudentsTable() {
    const list = getFilteredStudents();
    const tbody = document.getElementById('students-table-body');
    document.getElementById('select-all-students-chk').checked = false;

    if (!list.length) {
        tbody.innerHTML = '<tr><td colspan="5">لا يوجد طلاب مطابقون للبحث</td></tr>';
        updateSelectedStudentsCount();
        return;
    }
    tbody.innerHTML = list.map(s => `
        <tr>
            <td>
                <input type="checkbox" class="student-row-chk" value="${s.id}" ${selectedStudentIds.has(s.id) ? 'checked' : ''} onchange="toggleStudentSelection(${s.id}, this.checked)">
            </td>
            <td><span class="td-badge">${s.student_code}</span></td>
            <td>${s.full_name}</td>
            <td>${s.class_name}</td>
            <td>
                <button class="btn-sm btn-edit" onclick="promptEditStudent(${s.id})">✏️ تعديل / نقل</button>
                <button class="btn-sm btn-danger" onclick="deleteStudent(${s.id})">🗑️ حذف</button>
            </td>
        </tr>
    `).join('');
    updateSelectedStudentsCount();
}

function toggleStudentSelection(studentId, checked) {
    if (checked) selectedStudentIds.add(studentId);
    else selectedStudentIds.delete(studentId);
    updateSelectedStudentsCount();
}

function toggleSelectAllStudents(checked) {
    const list = getFilteredStudents();
    list.forEach(s => {
        if (checked) selectedStudentIds.add(s.id);
        else selectedStudentIds.delete(s.id);
    });
    document.querySelectorAll('.student-row-chk').forEach(chk => { chk.checked = checked; });
    updateSelectedStudentsCount();
}

async function deleteSelectedStudents() {
    if (selectedStudentIds.size === 0) {
        showToast('يرجى تحديد طالب واحد على الأقل من الجدول', true);
        return;
    }
    if (!confirm(`هل أنت متأكد من حذف ${selectedStudentIds.size} طالباً دفعة واحدة؟\n(سيبقى رقم TD الخاص بكل طالب محفوظاً باسمه في السجل الدائم ولن يُمنح لغيره)`)) return;

    const fd = new FormData();
    fd.append('student_ids_json', JSON.stringify(Array.from(selectedStudentIds)));
    const res = await fetch('/admin/students/bulk-delete', { method: 'POST', body: fd });
    const data = await res.json();
    if (res.ok) {
        selectedStudentIds.clear();
        showToast(data.message);
        loadMetadata();
    } else {
        showToast(data.detail || 'فشل الحذف الجماعي', true);
    }
}

async function deleteWholeClassStudents() {
    const cid = parseInt(document.getElementById('filter-student-class').value || '0');
    if (cid === 0) {
        showToast('يرجى اختيار فصل محدد من قائمة "تصفية حسب الفصل" أولاً لحذف طلابه بالكامل', true);
        return;
    }
    const cls = meta.classes.find(c => c.id === cid);
    const clsName = cls ? cls.name : cid;
    if (!confirm(`⚠️ تحذير: هل أنت متأكد من حذف جميع طلاب "${clsName}" بالكامل دفعة واحدة؟`)) return;

    const fd = new FormData();
    fd.append('class_id', cid);
    const res = await fetch('/admin/students/delete-by-class', { method: 'POST', body: fd });
    const data = await res.json();
    if (res.ok) {
        selectedStudentIds.clear();
        showToast(data.message);
        loadMetadata();
    } else {
        showToast(data.detail || 'فشل حذف طلاب الفصل', true);
    }
}

async function promoteSelectedStudents() {
    if (selectedStudentIds.size === 0) {
        showToast('يرجى تحديد الطلاب المراد ترقيتهم / نقلهم أولاً عبر المربعات', true);
        return;
    }
    const targetCid = document.getElementById('promote-target-class').value;
    const cls = meta.classes.find(c => String(c.id) === String(targetCid));
    if (!confirm(`هل تريد نقل / ترقية ${selectedStudentIds.size} طالباً إلى "${cls ? cls.name : targetCid}" مع الاحتفاظ بأرقامهم الدائمة (TD)؟`)) return;

    const fd = new FormData();
    fd.append('student_ids_json', JSON.stringify(Array.from(selectedStudentIds)));
    fd.append('target_class_id', targetCid);
    const res = await fetch('/admin/students/bulk-promote', { method: 'POST', body: fd });
    const data = await res.json();
    if (res.ok) {
        selectedStudentIds.clear();
        showToast(data.message);
        loadMetadata();
    } else {
        showToast(data.detail || 'فشل النقل الجماعي', true);
    }
}

function filterReportStudents() {
    const cid = parseInt(document.getElementById('rep-class-id').value || '0');
    const stSelect = document.getElementById('rep-student-id');
    const list = meta.students.filter(s => s.class_id === cid);
    stSelect.innerHTML = list.map(s => `<option value="${s.id}">${s.student_code} — ${s.full_name}</option>`).join('');
}

async function saveSessionSettings(e) {
    e.preventDefault();
    const fd = new FormData();
    fd.append('active_session', document.getElementById('set-session').value);
    fd.append('active_term', document.getElementById('set-term').value);
    fd.append('is_open', document.getElementById('set-is-open').value);
    const res = await fetch('/admin/session-settings', { method: 'POST', body: fd });
    const data = await res.json();
    if (res.ok) {
        showToast(data.message);
        loadMetadata();
    } else {
        showToast(data.detail || 'خطأ في حفظ الإعدادات', true);
    }
}

async function registerTeacher(e) {
    e.preventDefault();
    const fd = new FormData();
    fd.append('full_name', document.getElementById('t-name').value);
    const res = await fetch('/admin/register-teacher', { method: 'POST', body: fd });
    const data = await res.json();
    if (res.ok) {
        document.getElementById('new-teacher-key').textContent = `مفتاح الدخول: ${data.ref_key}`;
        document.getElementById('t-name').value = '';
        showToast(data.message);
        loadMetadata();
    } else {
        showToast(data.detail || 'خطأ', true);
    }
}

async function promptEditTeacher(teacherId) {
    const t = meta.teachers.find(x => x.id === teacherId);
    const currentName = t ? t.full_name : '';
    const newName = prompt('أدخل الاسم الجديد للمعلم:', currentName);
    if (!newName || !newName.trim()) return;
    const regen = confirm('هل ترغب أيضاً في توليد مفتاح دخول (UST-XXXX) جديد لهذا المعلم؟');
    const fd = new FormData();
    fd.append('full_name', newName.trim());
    fd.append('regenerate_key', regen ? '1' : '0');
    const res = await fetch(`/admin/teachers/${teacherId}/edit`, { method: 'POST', body: fd });
    const data = await res.json();
    if (res.ok) {
        showToast(data.message);
        loadMetadata();
    } else {
        showToast(data.detail || 'خطأ', true);
    }
}

async function deleteTeacher(teacherId) {
    const t = meta.teachers.find(x => x.id === teacherId);
    const name = t ? t.full_name : teacherId;
    if (!confirm(`هل أنت متأكد من حذف المعلم "${name}" وإلغاء جميع المواد المسندة إليه؟`)) return;
    const res = await fetch(`/admin/teachers/${teacherId}/delete`, { method: 'POST' });
    const data = await res.json();
    if (res.ok) {
        showToast(data.message);
        loadMetadata();
    } else {
        showToast(data.detail || 'خطأ', true);
    }
}

async function unassignTeacher(teacherId, classId, subjectId) {
    const fd = new FormData();
    fd.append('teacher_id', teacherId);
    fd.append('class_id', classId);
    fd.append('subject_id', subjectId);
    const res = await fetch('/admin/unassign-teacher', { method: 'POST', body: fd });
    const data = await res.json();
    if (res.ok) {
        showToast(data.message);
        loadMetadata();
    } else {
        showToast(data.detail || 'خطأ', true);
    }
}

async function registerStudent(e) {
    e.preventDefault();
    const fd = new FormData();
    fd.append('full_name', document.getElementById('st-name').value);
    fd.append('class_id', document.getElementById('st-class-id').value);
    const res = await fetch('/admin/students', { method: 'POST', body: fd });
    const data = await res.json();
    if (res.ok) {
        document.getElementById('st-name').value = '';
        showToast(data.message);
        loadMetadata();
    } else {
        showToast(data.detail || 'خطأ', true);
    }
}

async function bulkUploadStudents(e) {
    e.preventDefault();
    const fd = new FormData();
    fd.append('default_class_id', document.getElementById('bulk-class-id').value);
    const textVal = document.getElementById('bulk-names-text').value;
    if (textVal.trim()) {
        fd.append('names_text', textVal);
    }
    const fileInput = document.getElementById('bulk-file');
    if (fileInput.files.length > 0) {
        fd.append('file', fileInput.files[0]);
    }
    const res = await fetch('/admin/students/bulk', { method: 'POST', body: fd });
    const data = await res.json();
    if (res.ok) {
        document.getElementById('bulk-names-text').value = '';
        fileInput.value = '';
        showToast(data.message);
        loadMetadata();
    } else {
        showToast(data.detail || 'فشل الرفع الجماعي', true);
    }
}

async function promptEditStudent(studentId) {
    const s = meta.students.find(x => x.id === studentId);
    if (!s) return;
    const newName = prompt(`تعديل اسم الطالب (${s.student_code}):`, s.full_name);
    if (!newName || !newName.trim()) return;
    const fd = new FormData();
    fd.append('full_name', newName.trim());
    fd.append('class_id', s.class_id);
    const res = await fetch(`/admin/students/${studentId}/edit`, { method: 'POST', body: fd });
    const data = await res.json();
    if (res.ok) {
        showToast(data.message);
        loadMetadata();
    } else {
        showToast(data.detail || 'خطأ', true);
    }
}

async function deleteStudent(studentId) {
    const s = meta.students.find(x => x.id === studentId);
    const name = s ? `${s.full_name} (${s.student_code})` : studentId;
    if (!confirm(`هل أنت متأكد من حذف الطالب "${name}"؟\nسيبقى رقمه (${s ? s.student_code : ''}) محجوزاً باسمه للأبد ولن يُمنح لطالب آخر.`)) return;
    const res = await fetch(`/admin/students/${studentId}/delete`, { method: 'POST' });
    const data = await res.json();
    if (res.ok) {
        selectedStudentIds.delete(studentId);
        showToast(data.message);
        loadMetadata();
    } else {
        showToast(data.detail || 'خطأ', true);
    }
}

async function addClassWithSubjects(e) {
    e.preventDefault();
    const sids = Array.from(document.querySelectorAll('.new-cls-subj-chk:checked')).map(c => parseInt(c.value));
    const fd = new FormData();
    fd.append('name', document.getElementById('new-class-name').value);
    fd.append('subject_ids_json', JSON.stringify(sids));
    const res = await fetch('/admin/classes', { method: 'POST', body: fd });
    const data = await res.json();
    if (res.ok) {
        document.getElementById('new-class-name').value = '';
        showToast(data.message);
        loadMetadata();
    } else {
        showToast(data.detail || 'خطأ', true);
    }
}

async function addSubject(e) {
    e.preventDefault();
    const fd = new FormData();
    fd.append('name_ar', document.getElementById('new-subj-ar').value);
    fd.append('name_en', document.getElementById('new-subj-en').value);
    const res = await fetch('/admin/subjects', { method: 'POST', body: fd });
    const data = await res.json();
    if (res.ok) {
        document.getElementById('new-subj-ar').value = '';
        document.getElementById('new-subj-en').value = '';
        showToast(data.message);
        loadMetadata();
    } else {
        showToast(data.detail || 'خطأ', true);
    }
}

async function changeAdminPassword(e) {
    e.preventDefault();
    const fd = new FormData();
    fd.append('current_password', document.getElementById('adm-old-pass').value);
    fd.append('new_username', document.getElementById('adm-new-user').value);
    fd.append('new_password', document.getElementById('adm-new-pass').value);
    const res = await fetch('/admin/credentials', { method: 'POST', body: fd });
    const data = await res.json();
    if (res.ok) {
        document.getElementById('adm-old-pass').value = '';
        document.getElementById('adm-new-pass').value = '';
        showToast(data.message);
    } else {
        showToast(data.detail || 'فشل تحديث بيانات الإدارة', true);
    }
}

async function loadTeacherStudents() {
    const val = document.getElementById('t-assignment-select').value;
    if (!val) {
        document.getElementById('score-entry-tbody').innerHTML = '<tr><td colspan="6">لا توجد مواد مسندة لهذا المعلم بعد</td></tr>';
        return;
    }
    const [cid, sid] = val.split(':');
    const res = await fetch(`/api/scores?class_id=${cid}&subject_id=${sid}`);
    const data = await res.json();

    meta.settings.is_open = data.is_open;
    meta.settings.active_session = data.session_year;
    meta.settings.active_term = data.term_id;
    updateSessionBanner();

    const tbody = document.getElementById('score-entry-tbody');
    if (!data.students.length) {
        tbody.innerHTML = '<tr><td colspan="6">لا يوجد طلاب مسجلون في هذا الفصل</td></tr>';
        return;
    }
    const disabledAttr = data.is_open ? '' : 'disabled';
    tbody.innerHTML = data.students.map((st, idx) => `
        <tr>
            <td>${idx + 1}</td>
            <td><span class="td-badge">${st.student_code}</span></td>
            <td>${st.full_name}</td>
            <td><input type="number" min="0" max="100" step="0.5" class="teacher-score-input" data-student-id="${st.student_id}" id="sc-${st.student_id}" value="${st.score}" ${disabledAttr} style="width:105px; text-align:center;"></td>
            <td id="rem-${st.student_id}">${st.remark}</td>
            <td>
                <button class="action" ${disabledAttr} style="margin:0; padding:6px 12px; ${data.is_open ? '' : 'background:#9ca3af; cursor:not-allowed;'}" onclick="saveSingleScore(${st.student_id}, ${sid})">
                    ${data.is_open ? 'حفظ' : 'مقفل 🔒'}
                </button>
            </td>
        </tr>
    `).join('');
}

async function saveSingleScore(studentId, subjectId) {
    if (!meta.settings.is_open) {
        showToast('تم إغلاق رصد الدرجات من قبل الإدارة', true);
        return;
    }
    const scoreVal = document.getElementById(`sc-${studentId}`).value;
    if (scoreVal === '') {
        showToast('أدخل الدرجة أولاً', true);
        return;
    }
    const fd = new FormData();
    fd.append('student_id', studentId);
    fd.append('subject_id', subjectId);
    fd.append('score', scoreVal);
    if (activeTeacher) fd.append('ref_key', activeTeacher.ref_key);

    const res = await fetch('/scores/submit', { method: 'POST', body: fd });
    const data = await res.json();
    if (res.ok) {
        document.getElementById(`rem-${studentId}`).textContent = data.remark;
        showToast(data.message);
    } else {
        showToast(data.detail || 'فشل الحفظ', true);
    }
}

async function saveAllTeacherScores() {
    if (!meta.settings.is_open) {
        showToast('تم إغلاق رصد الدرجات من قبل الإدارة', true);
        return;
    }
    const val = document.getElementById('t-assignment-select').value;
    if (!val) return;
    const [, sid] = val.split(':');
    const items = [];
    document.querySelectorAll('.teacher-score-input').forEach(inp => {
        if (inp.value !== '') {
            items.push({
                student_id: parseInt(inp.dataset.studentId),
                score: parseFloat(inp.value)
            });
        }
    });
    if (!items.length) {
        showToast('لم يتم إدخال أي درجات للحفظ', true);
        return;
    }
    const fd = new FormData();
    fd.append('subject_id', sid);
    fd.append('scores_json', JSON.stringify(items));
    if (activeTeacher) fd.append('ref_key', activeTeacher.ref_key);

    const res = await fetch('/scores/submit-batch', { method: 'POST', body: fd });
    const data = await res.json();
    if (res.ok) {
        showToast(data.message);
        loadTeacherStudents();
    } else {
        showToast(data.detail || 'فشل حفظ الدرجات', true);
    }
}

function openWholeClassReportCards() {
    const cid = document.getElementById('batch-rep-class-id').value;
    const term = document.getElementById('batch-rep-term-id').value;
    const session = encodeURIComponent(document.getElementById('batch-rep-session').value);
    window.open(`/report-cards/class/${cid}?term_id=${term}&session_year=${session}`, '_blank');
}

function openReportCard() {
    const sid = document.getElementById('rep-student-id').value;
    if (!sid) {
        showToast('اختر طالباً أولاً', true);
        return;
    }
    const term = document.getElementById('rep-term-id').value;
    const session = encodeURIComponent(document.getElementById('rep-session').value);
    window.open(`/report-card/${sid}?term_id=${term}&session_year=${session}`, '_blank');
}

function downloadTermBroadsheet() {
    const cid = document.getElementById('bs-class-id').value;
    const term = document.getElementById('bs-term-id').value;
    const session = encodeURIComponent(document.getElementById('bs-session').value);
    window.location.href = `/export/broadsheet/term?class_id=${cid}&term_id=${term}&session_year=${session}`;
}

function downloadAnnualBroadsheet() {
    const cid = document.getElementById('bs-class-id').value;
    const session = encodeURIComponent(document.getElementById('bs-session').value);
    window.location.href = `/export/broadsheet/annual?class_id=${cid}&session_year=${session}`;
}

// Initial startup: load session banner and restore active session if present
(async function initPortal() {
    try {
        const sRes = await fetch('/api/session-settings');
        if (sRes.ok) {
            meta.settings = await sRes.json();
            updateSessionBanner();
        }
        const savedRole = sessionStorage.getItem('ntd_role');
        if (savedRole === 'admin') {
            await loadMetadata();
            showPortalView('admin');
        } else if (savedRole === 'teacher') {
            const savedKey = sessionStorage.getItem('ntd_teacher_key');
            if (savedKey) {
                document.getElementById('login-ref-key').value = savedKey;
                await handleTeacherLogin(null);
            }
        }
    } catch (err) {
        console.error(err);
    }
})();
</script>
</body>
</html>"""


if __name__ == "__main__":
    import uvicorn

    port = int(os.environ.get("PORT", 8000))
    host = os.environ.get("HOST", "0.0.0.0")
    uvicorn.run("app:app", host=host, port=port, reload=False)


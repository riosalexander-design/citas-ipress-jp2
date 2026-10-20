"""
database.py
Base de datos SQLite para el Aplicativo de Citas en Línea del Servicio de Medicina
del C.S. Su Santidad Juan Pablo II, sincronizado con la programación asistencial
de https://turnos-ipress.streamlit.app/
"""

import sqlite3
import hashlib
import os
from datetime import datetime

DB_FILE = os.path.join(os.path.dirname(__file__), "citas_jp2_v2.db")

# Ruta local de la base de datos de turnos-ipress (cuando se ejecuta en el mismo equipo)
EXTERNAL_TURNOS_DB_CANDIDATES = [
    r"C:\Users\Alexander\OneDrive - EduCorpPERU (1)\2026\2026-2\app_turnos_ipress\ipress_turnos.db",
    os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "app_turnos_ipress", "ipress_turnos.db")),
]


def hash_pin(pin: str) -> str:
    """Hashea un PIN o contraseña con SHA-256."""
    return hashlib.sha256(pin.encode("utf-8")).hexdigest()


def _remove_corrupted_db(target: str):
    """Elimina archivos SQLite corruptos para permitir su recreación limpia."""
    for suffix in ("", "-wal", "-shm", "-journal"):
        p = target + suffix
        if os.path.exists(p):
            try:
                os.remove(p)
            except OSError:
                pass


def get_connection(db_path: str = None):
    """Retorna una conexión SQLite segura con auto-recuperación ante archivos corruptos."""
    target = db_path or DB_FILE
    try:
        conn = sqlite3.connect(target, timeout=30.0)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode = DELETE")
        # Verificar que el archivo sea una base de datos SQLite válida
        conn.execute("SELECT name FROM sqlite_master LIMIT 1").fetchall()
        return conn
    except sqlite3.DatabaseError:
        try:
            conn.close()
        except Exception:
            pass
        _remove_corrupted_db(target)
        conn = sqlite3.connect(target, timeout=30.0)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode = DELETE")
        return conn


def find_external_turnos_db() -> str | None:
    """Retorna la ruta de ipress_turnos.db si está disponible localmente."""
    for path in EXTERNAL_TURNOS_DB_CANDIDATES:
        if os.path.exists(path):
            return path
    return None


def init_db():
    """Inicializa las tablas de médicos, turnos sincronizados, citas médicas y auditoría."""
    try:
        conn = get_connection()
    except sqlite3.DatabaseError:
        _remove_corrupted_db(DB_FILE)
        conn = get_connection()

    with conn:
        cursor = conn.cursor()

        # 1. Tabla de Médicos Cirujanos (idéntica a turnos-ipress)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS doctors (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                condition TEXT NOT NULL, -- 'CAS' o 'Nombrado'
                cmp TEXT UNIQUE NOT NULL,
                pin_hash TEXT NOT NULL,
                role TEXT DEFAULT 'doctor', -- 'doctor' o 'admin'
                is_active INTEGER DEFAULT 1,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # 2. Tabla de Turnos Asistenciales Sincronizados desde turnos-ipress
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS shifts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                year INTEGER NOT NULL,
                month INTEGER NOT NULL,
                day INTEGER NOT NULL,
                shift_period TEXT NOT NULL, -- 'M', 'T', '12h', 'DES', 'LIBRE'
                service TEXT NOT NULL,      -- 'TU', 'CE', 'TS', 'TU+CE', 'CE+TU', 'CE+CE', etc.
                doctor_id INTEGER NOT NULL REFERENCES doctors(id),
                is_proposal INTEGER DEFAULT 1,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(year, month, day, doctor_id)
            )
        """)

        # 3. Tabla de Asignación Titular o Contingencia de Consultorio CE (Mañana / Tarde)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS clinic_schedule_overrides (
                year INTEGER NOT NULL,
                month INTEGER NOT NULL,
                day INTEGER NOT NULL,
                shift_period TEXT NOT NULL, -- 'M' o 'T'
                doctor_id INTEGER REFERENCES doctors(id),
                is_enabled INTEGER DEFAULT 1, -- 1=Consultorio habilitado, 0=Cerrado excepcionalmente
                notes TEXT,
                updated_by TEXT,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY(year, month, day, shift_period)
            )
        """)

        # 4. Tabla de Citas Médicas (16 cupos en Mañana y 16 cupos en Tarde)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS appointments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_code TEXT UNIQUE NOT NULL,
                year INTEGER NOT NULL,
                month INTEGER NOT NULL,
                day INTEGER NOT NULL,
                appointment_date TEXT NOT NULL, -- YYYY-MM-DD
                shift_period TEXT NOT NULL,     -- 'M' (Mañana 08:00-12:00) o 'T' (Tarde 14:00-18:00)
                slot_number INTEGER NOT NULL,   -- 1 al 16
                slot_time TEXT NOT NULL,        -- Ej: '08:00 - 08:15'
                doctor_id INTEGER NOT NULL REFERENCES doctors(id),
                doctor_name TEXT NOT NULL,
                doctor_cmp TEXT NOT NULL,
                patient_doc_type TEXT DEFAULT 'DNI', -- 'DNI' o 'CE'
                patient_dni TEXT NOT NULL,
                patient_names TEXT NOT NULL,
                patient_age INTEGER,
                patient_phone TEXT,
                insurance_type TEXT DEFAULT 'SIS', -- 'SIS', 'Pagante', 'Exonerado / Otro'
                reason TEXT,
                status TEXT DEFAULT 'reservada',   -- 'reservada', 'atendida', 'no_asistio', 'cancelada'
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Migración automática si la tabla ya existía sin patient_doc_type
        cols_info = cursor.execute("PRAGMA table_info(appointments)").fetchall()
        col_names = [c["name"] for c in cols_info]
        if "patient_doc_type" not in col_names:
            cursor.execute("ALTER TABLE appointments ADD COLUMN patient_doc_type TEXT DEFAULT 'DNI'")

        # Índices para búsquedas rápidas por fecha, turno y DNI
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_appointments_date_shift
            ON appointments(appointment_date, shift_period, status)
        """)
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_appointments_dni
            ON appointments(patient_dni, status)
        """)

        # 5. Tabla de Auditoría
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS audit_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_name TEXT NOT NULL,
                action TEXT NOT NULL,
                details TEXT,
                timestamp TEXT NOT NULL
            )
        """)

        # 6. Metadatos de Sincronización
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sync_metadata (
                key TEXT PRIMARY KEY,
                value TEXT,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        conn.commit()

        # Precargar la nómina oficial de médicos del C.S. Su Santidad Juan Pablo II y Jefatura
        initial_doctors = [
            ("Caldas Pacsi, Lily Katherine", "CAS", "71201", "4182", "doctor"),
            ("Chunga Torres, Tuly Magnolia", "CAS", "68432", "6935", "doctor"),
            ("Laureano Varillas, Elizabeth Miriam", "CAS", "75319", "5274", "doctor"),
            ("Sanchez Quintanilla, Cinthya", "CAS", "80145", "8319", "doctor"),
            ("Solano Ricaldi, Edward Jhonny", "Nombrado", "54982", "7461", "doctor"),
            ("Jefatura Médica / Admisión C.S. Juan Pablo II", "Nombrado", "ADMIN", "admin2026", "admin"),
        ]

        for name, cond, cmp_num, pin, role in initial_doctors:
            cursor.execute("SELECT id FROM doctors WHERE cmp = ?", (cmp_num,))
            row = cursor.fetchone()
            if not row:
                cursor.execute(
                    """
                    INSERT INTO doctors (name, condition, cmp, pin_hash, role)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    (name, cond, cmp_num, hash_pin(pin), role),
                )
            else:
                cursor.execute(
                    """
                    UPDATE doctors
                    SET condition = ?, name = ?
                    WHERE cmp = ?
                    """,
                    (cond, name, cmp_num),
                )

        conn.commit()
    conn.close()


def log_audit(user_name: str, action: str, details: str = "", conn=None):
    """Registra un evento en el log de auditoría del sistema de citas."""
    now_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if conn is not None:
        conn.execute(
            """
            INSERT INTO audit_logs (user_name, action, details, timestamp)
            VALUES (?, ?, ?, ?)
            """,
            (user_name, action, details, now_iso),
        )
    else:
        with get_connection() as c:
            c.execute(
                """
                INSERT INTO audit_logs (user_name, action, details, timestamp)
                VALUES (?, ?, ?, ?)
                """,
                (user_name, action, details, now_iso),
            )
            c.commit()

import sqlite3
from pathlib import Path

DB_PATH = Path(__file__).resolve().parent.parent / "data" / "sgsst.db"

_SCHEMA = """
CREATE TABLE IF NOT EXISTS trabajadores (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    dni              TEXT UNIQUE NOT NULL,
    nombre           TEXT NOT NULL,
    apellido         TEXT,
    cargo            TEXT,
    fecha_nacimiento TEXT,
    correo           TEXT,
    celular          TEXT,
    estado           TEXT NOT NULL DEFAULT 'ACTIVO',
    created_at       TEXT
);

CREATE TABLE IF NOT EXISTS documentos_generados (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    dni_trabajador   TEXT NOT NULL,
    tipo_documento   TEXT NOT NULL,
    fecha_generacion TEXT,
    ruta_archivo     TEXT NOT NULL,
    FOREIGN KEY (dni_trabajador) REFERENCES trabajadores(dni)
);
"""


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with get_connection() as conn:
        conn.executescript(_SCHEMA)
        # Migration: add apellido column to existing databases
        try:
            conn.execute("ALTER TABLE trabajadores ADD COLUMN apellido TEXT")
            conn.commit()
        except Exception:
            pass  # column already exists

import sqlite3

from src.config import DB_PATH

_SCHEMA = """
PRAGMA journal_mode = WAL;
PRAGMA foreign_keys = ON;

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
    created_at       TEXT NOT NULL DEFAULT (strftime('%Y-%m-%d %H:%M:%S', 'now'))
);

CREATE TABLE IF NOT EXISTS documentos_generados (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    dni_trabajador   TEXT NOT NULL,
    tipo_documento   TEXT NOT NULL,
    fecha_generacion TEXT NOT NULL DEFAULT (strftime('%Y-%m-%d %H:%M:%S', 'now')),
    ruta_archivo     TEXT NOT NULL,
    FOREIGN KEY (dni_trabajador) REFERENCES trabajadores(dni)
);
"""


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with get_connection() as conn:
        conn.executescript(_SCHEMA)
        try:
            conn.execute("ALTER TABLE trabajadores ADD COLUMN apellido TEXT")
            conn.commit()
        except sqlite3.OperationalError:
            pass  # columna ya existe

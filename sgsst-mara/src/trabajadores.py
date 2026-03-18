from datetime import datetime

from src.database import get_connection

_ESTADOS_VALIDOS = {"ACTIVO", "INACTIVO"}


def buscar_por_dni(dni: str) -> dict | None:
    with get_connection() as conn:
        row = conn.execute(
            "SELECT * FROM trabajadores WHERE dni = ?", (dni.strip(),)
        ).fetchone()
    return dict(row) if row else None


def listar_activos() -> list[dict]:
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT * FROM trabajadores WHERE estado = 'ACTIVO' ORDER BY nombre ASC"
        ).fetchall()
    return [dict(r) for r in rows]


def cambiar_estado(dni: str, nuevo_estado: str) -> bool:
    if nuevo_estado not in _ESTADOS_VALIDOS:
        raise ValueError(f"Estado inválido '{nuevo_estado}'. Válidos: {_ESTADOS_VALIDOS}")
    with get_connection() as conn:
        cursor = conn.execute(
            "UPDATE trabajadores SET estado = ? WHERE dni = ?", (nuevo_estado, dni.strip())
        )
    return cursor.rowcount > 0


def registrar_trabajador(datos: dict) -> bool:
    created_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with get_connection() as conn:
        cursor = conn.execute(
            """INSERT OR IGNORE INTO trabajadores
               (dni, nombre, cargo, fecha_nacimiento, correo, celular, estado, created_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                datos["dni"],
                datos["nombre"],
                datos.get("cargo"),
                datos.get("fecha_nacimiento"),
                datos.get("correo"),
                datos.get("celular"),
                datos.get("estado", "ACTIVO"),
                created_at,
            ),
        )
    return cursor.rowcount > 0

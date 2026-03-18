from src.database import get_connection


def buscar_por_dni(dni: str):
    with get_connection() as conn:
        row = conn.execute(
            "SELECT * FROM trabajadores WHERE dni = ?", (dni.strip(),)
        ).fetchone()
    return dict(row) if row else None


def listar_activos() -> list:
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT * FROM trabajadores WHERE estado = 'ACTIVO' ORDER BY nombre"
        ).fetchall()
    return [dict(r) for r in rows]

import csv

import requests

from src.config import GOOGLE_SHEET_ID
from src.database import get_connection

_UPSERT = """
INSERT INTO trabajadores
    (dni, nombre, apellido, cargo, fecha_nacimiento, correo, celular, estado,
     created_at)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, strftime('%Y-%m-%d %H:%M:%S', 'now'))
ON CONFLICT(dni) DO UPDATE SET
    nombre           = excluded.nombre,
    apellido         = excluded.apellido,
    cargo            = excluded.cargo,
    fecha_nacimiento = excluded.fecha_nacimiento,
    correo           = excluded.correo,
    celular          = excluded.celular,
    estado           = excluded.estado
"""


def sincronizar_desde_sheets() -> dict:
    """
    Descarga la hoja 'Registro' del Google Sheet y sincroniza contra SQLite.
    Retorna: {"nuevos": int, "actualizados": int, "omitidos": int, "errores": list}
    """
    url = (
        f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}"
        f"/gviz/tq?tqx=out:csv&sheet=Registro"
    )
    response = requests.get(url, timeout=15)
    response.raise_for_status()

    lineas = list(csv.reader(response.content.decode('utf-8').splitlines()))
    if not lineas:
        return {"nuevos": 0, "actualizados": 0, "omitidos": 0, "errores": ["Hoja vacía"]}

    encabezados = lineas[0]

    def _idx(col: str) -> int | None:
        return encabezados.index(col) if col in encabezados else None

    # Columnas obligatorias
    try:
        idx_nombre   = encabezados.index('Nombre')
        idx_apellido = encabezados.index('Apellido')
        idx_dni      = encabezados.index('DNI')
        idx_cargo    = encabezados.index('Cargo')
        idx_estado   = encabezados.index('Estado')
    except ValueError as exc:
        return {"nuevos": 0, "actualizados": 0, "omitidos": 0,
                "errores": [f"Columna faltante en el Sheet: {exc}"]}

    # Columnas opcionales
    idx_fn      = _idx('Fecha de nacimiento')
    idx_correo  = _idx('Correo electrónico')
    idx_celular = _idx('Celular')

    nuevos       = 0
    actualizados = 0
    omitidos     = 0
    errores      = []

    def _get(fila, idx):
        return fila[idx].strip() if idx is not None and idx < len(fila) else None

    with get_connection() as conn:
        for i, fila in enumerate(lineas[1:], start=2):
            try:
                if len(fila) <= idx_estado:
                    omitidos += 1
                    continue

                dni = _get(fila, idx_dni)
                if not dni:
                    omitidos += 1
                    continue

                apellido = _get(fila, idx_apellido) or ''
                nombre_p = _get(fila, idx_nombre) or ''
                if not apellido and not nombre_p:
                    omitidos += 1
                    continue

                cargo    = _get(fila, idx_cargo)
                estado   = (_get(fila, idx_estado) or 'ACTIVO').upper()
                fecha_nac = _get(fila, idx_fn)
                correo    = _get(fila, idx_correo)
                celular   = _get(fila, idx_celular)

                existe = conn.execute(
                    "SELECT 1 FROM trabajadores WHERE dni = ?", (dni,)
                ).fetchone()

                conn.execute(
                    _UPSERT,
                    (dni, nombre_p, apellido, cargo, fecha_nac,
                     correo, celular, estado),
                )

                if existe:
                    actualizados += 1
                else:
                    nuevos += 1

            except Exception as exc:
                errores.append(f"Fila {i}: {exc}")

    return {
        "nuevos":       nuevos,
        "actualizados": actualizados,
        "omitidos":     omitidos,
        "errores":      errores,
    }

import csv
import requests
from datetime import datetime

from src.database import get_connection

GOOGLE_SHEET_ID = '1N0i-sG3UBH2UMVo_DF4Ivcwxyr2MvgIfAa4h18oQ5oc'

_UPSERT = """
INSERT INTO trabajadores
    (dni, nombre, apellido, cargo, fecha_nacimiento, correo, celular, estado, created_at)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
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
    - Nuevos DNI   → INSERT
    - DNI ya existe → UPDATE (nombre, cargo, estado, etc.)
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

    idx_nombre   = encabezados.index('Nombre')
    idx_apellido = encabezados.index('Apellido')
    idx_dni      = encabezados.index('DNI')
    idx_cargo    = encabezados.index('Cargo')
    idx_estado   = encabezados.index('Estado')
    idx_fn       = _idx('Fecha de nacimiento')
    idx_correo   = _idx('Correo electrónico')
    idx_celular  = _idx('Celular')

    nuevos      = 0
    actualizados = 0
    omitidos    = 0
    errores     = []

    conn = get_connection()
    try:
        for i, fila in enumerate(lineas[1:], start=2):
            try:
                if len(fila) <= idx_estado:
                    omitidos += 1
                    continue

                dni = str(fila[idx_dni]).strip()
                if not dni:
                    omitidos += 1
                    continue

                apellido  = fila[idx_apellido].strip()
                nombre_p  = fila[idx_nombre].strip()
                if not apellido and not nombre_p:
                    omitidos += 1
                    continue

                # Canonical format: APELLIDO NOMBRE
                nombre_completo = f"{apellido} {nombre_p}".strip()
                cargo    = fila[idx_cargo].strip() if len(fila) > idx_cargo else None
                estado   = fila[idx_estado].strip().upper() if len(fila) > idx_estado else 'ACTIVO'
                fecha_nac = fila[idx_fn].strip() if idx_fn is not None and len(fila) > idx_fn else None
                correo    = fila[idx_correo].strip() if idx_correo is not None and len(fila) > idx_correo else None
                celular   = fila[idx_celular].strip() if idx_celular is not None and len(fila) > idx_celular else None
                created_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                # Check whether DNI already exists to distinguish new vs updated
                existe = conn.execute(
                    "SELECT 1 FROM trabajadores WHERE dni = ?", (dni,)
                ).fetchone()

                conn.execute(
                    _UPSERT,
                    (dni, nombre_completo, apellido, cargo, fecha_nac,
                     correo, celular, estado, created_at),
                )

                if existe:
                    actualizados += 1
                else:
                    nuevos += 1

            except Exception as exc:
                errores.append(f"Fila {i}: {exc}")

        conn.commit()
    finally:
        conn.close()

    return {
        "nuevos": nuevos,
        "actualizados": actualizados,
        "omitidos": omitidos,
        "errores": errores,
    }

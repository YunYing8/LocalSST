import openpyxl
from pathlib import Path

from src.database import get_connection

_INSERT = """
INSERT OR IGNORE INTO trabajadores
    (dni, nombre, apellido, cargo, fecha_nacimiento, correo, celular, estado, created_at)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, strftime('%Y-%m-%d %H:%M:%S', 'now'))
"""

# Mapeo: nombre de columna en el Excel → clave interna
_COL_MAP = {
    'nombre':           'nombre',
    'apellido':         'apellido',
    'dni':              'dni',
    'cargo':            'cargo',
    'fecha_nacimiento': 'fecha_nacimiento',
    'fecha nacimiento': 'fecha_nacimiento',
    'correo':           'correo',
    'correo electronico': 'correo',
    'correo electrónico': 'correo',
    'celular':          'celular',
    'telefono':         'celular',
    'teléfono':         'celular',
    'estado':           'estado',
}


def _cell(row: tuple, indices: dict, key: str) -> str | None:
    idx = indices.get(key)
    if idx is None or idx >= len(row):
        return None
    val = row[idx]
    return str(val).strip() if val is not None else None


def importar_trabajadores(ruta_excel: str | Path) -> dict:
    wb = openpyxl.load_workbook(str(ruta_excel), read_only=True, data_only=True)
    ws = wb.worksheets[0]

    importados = 0
    omitidos   = 0
    errores    = []

    # Leer encabezados de la primera fila para mapear columnas dinámicamente
    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if header_row is None:
        wb.close()
        return {"importados": 0, "omitidos": 0, "errores": ["El archivo está vacío"]}

    indices: dict[str, int] = {}
    for col_idx, cell in enumerate(header_row):
        if cell is not None:
            key = str(cell).strip().lower()
            if key in _COL_MAP:
                indices[_COL_MAP[key]] = col_idx

    if 'dni' not in indices or 'nombre' not in indices:
        wb.close()
        return {
            "importados": 0,
            "omitidos":   0,
            "errores":    ["No se encontraron las columnas obligatorias 'DNI' y 'Nombre'"],
        }

    conn = get_connection()
    try:
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            try:
                dni = _cell(row, indices, 'dni')
                if not dni or dni.lower() == 'none':
                    omitidos += 1
                    continue

                nombre_raw   = _cell(row, indices, 'nombre') or ''
                apellido_str = _cell(row, indices, 'apellido') or ''
                if not nombre_raw and not apellido_str:
                    omitidos += 1
                    continue

                nombre = nombre_raw

                cargo     = _cell(row, indices, 'cargo')
                fecha_nac = _cell(row, indices, 'fecha_nacimiento')
                correo    = _cell(row, indices, 'correo')
                celular   = _cell(row, indices, 'celular')
                estado    = (_cell(row, indices, 'estado') or 'ACTIVO').upper()

                cursor = conn.execute(
                    _INSERT,
                    (dni, nombre, apellido_str, cargo, fecha_nac, correo, celular, estado),
                )
                if cursor.rowcount == 1:
                    importados += 1
                else:
                    omitidos += 1

            except Exception as exc:
                errores.append(f"Fila {row_idx}: {exc}")
                omitidos += 1

        conn.commit()
    finally:
        conn.close()
        wb.close()

    return {"importados": importados, "omitidos": omitidos, "errores": errores}

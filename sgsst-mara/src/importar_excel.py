import openpyxl
from datetime import datetime

from src.database import get_connection

_INSERT = """
INSERT OR IGNORE INTO trabajadores
    (dni, nombre, apellido, cargo, fecha_nacimiento, correo, celular, estado, created_at)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
"""


def importar_trabajadores(ruta_excel: str) -> dict:
    wb = openpyxl.load_workbook(ruta_excel, read_only=True, data_only=True)
    ws = wb.worksheets[0]

    importados = 0
    omitidos = 0
    errores = []

    conn = get_connection()
    try:
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            try:
                nombre_col, apellido_col, dni_raw, cargo, fecha_nac, _edad, correo, celular, estado = row

                # Validate DNI
                dni = str(dni_raw).strip() if dni_raw is not None else None
                if not dni or dni == "None":
                    omitidos += 1
                    continue

                # Validate nombre
                nombre_raw = str(nombre_col).strip() if nombre_col else ""
                if not nombre_raw or nombre_raw == "None":
                    omitidos += 1
                    continue

                # Store as "APELLIDO NOMBRE" — matches constancia format
                apellido_str = str(apellido_col).strip() if apellido_col else ""
                nombre = f"{apellido_str} {nombre_raw}".strip()

                estado_val  = str(estado).strip().upper() if estado else "ACTIVO"
                fecha_str   = str(fecha_nac).strip() if fecha_nac is not None else None
                cargo_str   = str(cargo).strip() if cargo is not None else None
                correo_str  = str(correo).strip() if correo is not None else None
                celular_str = str(celular).strip() if celular is not None else None
                created_at  = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                cursor = conn.execute(
                    _INSERT,
                    (dni, nombre, apellido_str, cargo_str, fecha_str,
                     correo_str, celular_str, estado_val, created_at),
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

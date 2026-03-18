import openpyxl

from src.database import get_connection

_INSERT = """
INSERT OR IGNORE INTO trabajadores (dni, nombre, cargo, fecha_nacimiento, correo, celular, estado)
VALUES (?, ?, ?, ?, ?, ?, ?)
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
                nombre_col, apellido, dni_raw, cargo, fecha_nac, _edad, correo, celular, estado = row

                dni = str(dni_raw).strip() if dni_raw is not None else None
                if not dni or dni == "None":
                    errores.append(f"Fila {row_idx}: DNI vacío, omitida")
                    omitidos += 1
                    continue

                nombre_raw = nombre_col or ""
                apellido_raw = apellido or ""
                nombre = f"{nombre_raw} {apellido_raw}".strip()
                if not nombre:
                    errores.append(f"Fila {row_idx}: Nombre vacío, omitida")
                    omitidos += 1
                    continue

                estado_val = str(estado).strip().upper() if estado else "ACTIVO"
                fecha_str = str(fecha_nac).strip() if fecha_nac is not None else None
                cargo_str = str(cargo).strip() if cargo is not None else None
                correo_str = str(correo).strip() if correo is not None else None
                celular_str = str(celular).strip() if celular is not None else None

                cursor = conn.execute(_INSERT, (dni, nombre, cargo_str, fecha_str, correo_str, celular_str, estado_val))
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

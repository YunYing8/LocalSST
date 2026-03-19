import sys
import tempfile
import os
import openpyxl

from src.database import init_db, get_connection, DB_PATH
from src.trabajadores import (
    registrar_trabajador,
    buscar_por_dni,
    listar_activos,
    cambiar_estado,
)
from src.importar_excel import importar_trabajadores
from src.trabajadores import nombre_completo

passed = 0
failed = 0


def check(label: str, number: int):
    import contextlib

    @contextlib.contextmanager
    def _ctx():
        global passed, failed
        try:
            yield
            print(f"[{number}] {label} PASS")
            passed += 1
        except Exception as exc:
            print(f"[{number}] {label} FAIL  ({exc})")
            failed += 1

    return _ctx()


if __name__ == "__main__":
    # TEST 1 - init_db()
    with check("init_db() ...........       ", 1):
        init_db()
        assert DB_PATH.exists(), f"DB file not found at {DB_PATH}"
        with get_connection() as conn:
            tables = {
                row[0]
                for row in conn.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                ).fetchall()
            }
        assert "trabajadores" in tables, "Table 'trabajadores' missing"
        assert "documentos_generados" in tables, "Table 'documentos_generados' missing"

        # Verify apellido column exists
        with get_connection() as conn:
            cols = {
                row[1]
                for row in conn.execute(
                    "PRAGMA table_info(trabajadores)"
                ).fetchall()
            }
        assert "apellido" in cols, "Column 'apellido' missing"

    # TEST 2 - registrar_trabajador()
    with check("registrar_trabajador() ...", 2):
        result = registrar_trabajador(
            {
                "dni":             "99999999",
                "nombre":          "TRABAJADOR TEST",
                "apellido":        "TRABAJADOR",
                "cargo":           "PRUEBA",
                "fecha_nacimiento": "01/01/1990",
                "correo":          "test@test.com",
                "celular":         "999000000",
                "estado":          "ACTIVO",
            }
        )
        assert result is True, f"Expected True, got {result}"

    # TEST 3 - buscar_por_dni()
    with check("buscar_por_dni() ........", 3):
        t = buscar_por_dni("99999999")
        assert t is not None, "Expected a dict, got None"
        assert t["nombre"] == "TRABAJADOR TEST", f"Unexpected nombre: {t['nombre']}"
        assert t["apellido"] == "TRABAJADOR", f"Unexpected apellido: {t['apellido']}"
        assert t["created_at"] is not None, "created_at missing"

    # TEST 4 - listar_activos()
    with check("listar_activos() ........", 4):
        activos = listar_activos()
        assert len(activos) >= 1, "Expected at least 1 active worker"
        assert all(r["estado"] == "ACTIVO" for r in activos), \
            "Some records have estado != ACTIVO"

    # TEST 5 - cambiar_estado()
    with check("cambiar_estado() ........", 5):
        ok = cambiar_estado("99999999", "INACTIVO")
        assert ok is True, f"Expected True, got {ok}"
        t = buscar_por_dni("99999999")
        assert t["estado"] == "INACTIVO", f"Expected INACTIVO, got {t['estado']}"

    # TEST 6 - cambiar_estado() invalid input
    with check("cambiar_estado() validación", 6):
        raised = False
        try:
            cambiar_estado("99999999", "SUSPENDIDO")
        except ValueError:
            raised = True
        assert raised, "Expected ValueError but none was raised"

    # TEST 7 - importar_trabajadores() with synthetic Excel
    with check("importar_trabajadores() ..", 7):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append([
            "Nombre", "Apellido", "DNI", "Cargo",
            "Fecha de nacimiento", "Edad",
            "Correo electrónico", "Celular", "Estado",
        ])
        # Valid row — nombre stored as first name, apellido stored separately
        ws.append([
            "JUAN", "PEREZ LOPEZ", "11111111", "RIGGER",
            "01/01/1985", 40, "juan@test.com", "911000000", "ACTIVO",
        ])
        # Invalid row — DNI and Nombre are None
        ws.append([None, None, None, None, None, None, None, None, "ACTIVO"])

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name
        try:
            wb.save(tmp_path)
            res = importar_trabajadores(tmp_path)
        finally:
            os.unlink(tmp_path)

        assert res["importados"] == 1, f"Expected importados=1, got {res['importados']}"
        assert res["omitidos"]   == 1, f"Expected omitidos=1, got {res['omitidos']}"

        t2 = buscar_por_dni("11111111")
        assert t2 is not None, "Imported worker not found"
        assert t2["nombre"]  == "JUAN",       f"Unexpected nombre: {t2['nombre']}"
        assert t2["apellido"] == "PEREZ LOPEZ", f"Unexpected apellido: {t2['apellido']}"
        assert nombre_completo(t2) == "PEREZ LOPEZ JUAN", \
            f"Unexpected nombre_completo: {nombre_completo(t2)}"
        assert t2["created_at"] is not None, "created_at missing"

    # TEST 8 - cleanup
    with check("cleanup .................", 8):
        with get_connection() as conn:
            conn.execute(
                "DELETE FROM trabajadores WHERE dni IN ('99999999', '11111111')"
            )
        assert buscar_por_dni("99999999") is None, "99999999 still present after cleanup"
        assert buscar_por_dni("11111111") is None, "11111111 still present after cleanup"

    print()
    print("=============================")
    print(f"TOTAL: {passed} passed, {failed} failed")
    print("=============================")
    sys.exit(0 if failed == 0 else 1)

from src.database import init_db
from src.importar_excel import importar_trabajadores
from src.trabajadores import buscar_por_dni, listar_activos

_MENU = """
========================================
       SISTEMA SGSST - MARA
========================================
[1] Importar trabajadores desde Excel
[2] Buscar trabajador por DNI
[3] Listar trabajadores activos
[4] Generar documento Word
[5] Salir
----------------------------------------
Seleccione una opción: """

_FIELDS = [
    ("ID",               "id"),
    ("DNI",              "dni"),
    ("Nombre",           "nombre"),
    ("Cargo",            "cargo"),
    ("Fecha nacimiento", "fecha_nacimiento"),
    ("Correo",           "correo"),
    ("Celular",          "celular"),
    ("Estado",           "estado"),
    ("Creado",           "created_at"),
]


def _print_trabajador(t: dict) -> None:
    label_w = max(len(label) for label, _ in _FIELDS)
    for label, key in _FIELDS:
        value = t.get(key) or "-"
        print(f"  {label.ljust(label_w)} : {value}")


def _print_activos(trabajadores: list) -> None:
    col_dni    = max((len(str(t["dni"]))    for t in trabajadores), default=3)
    col_nombre = max((len(str(t["nombre"])) for t in trabajadores), default=6)
    col_cargo  = max((len(str(t["cargo"] or "")) for t in trabajadores), default=5)

    col_dni    = max(col_dni,    3)
    col_nombre = max(col_nombre, 6)
    col_cargo  = max(col_cargo,  5)

    sep = f"+-{'-' * col_dni}-+-{'-' * col_nombre}-+-{'-' * col_cargo}-+"
    header = (
        f"| {'DNI'.ljust(col_dni)} "
        f"| {'NOMBRE'.ljust(col_nombre)} "
        f"| {'CARGO'.ljust(col_cargo)} |"
    )
    print(sep)
    print(header)
    print(sep)
    for t in trabajadores:
        dni    = str(t["dni"]).ljust(col_dni)
        nombre = str(t["nombre"]).ljust(col_nombre)
        cargo  = str(t["cargo"] or "-").ljust(col_cargo)
        print(f"| {dni} | {nombre} | {cargo} |")
    print(sep)
    print(f"Total: {len(trabajadores)} trabajadores activos")


def _opcion_1() -> None:
    ruta = input("Ingrese ruta del archivo Excel: ").strip()
    result = importar_trabajadores(ruta)
    print(f"  Importados : {result['importados']}")
    print(f"  Omitidos   : {result['omitidos']}")
    if result["errores"]:
        print("  Errores:")
        for e in result["errores"]:
            print(f"    - {e}")


def _opcion_2() -> None:
    dni = input("Ingrese DNI: ").strip()
    trabajador = buscar_por_dni(dni)
    if trabajador is None:
        print("  Trabajador no encontrado")
    else:
        _print_trabajador(trabajador)


def _opcion_3() -> None:
    activos = listar_activos()
    if not activos:
        print("  No hay trabajadores activos registrados.")
        return
    _print_activos(activos)


def main() -> None:
    init_db()
    while True:
        opcion = input(_MENU).strip()
        print()
        try:
            if opcion == "1":
                _opcion_1()
            elif opcion == "2":
                _opcion_2()
            elif opcion == "3":
                _opcion_3()
            elif opcion == "4":
                print("Módulo pendiente - disponible en siguiente versión")
            elif opcion == "5":
                print("Saliendo...")
                break
            else:
                print("  Opción no válida. Intente de nuevo.")
        except Exception as e:
            print(f"  Error: {e}")
        print()


if __name__ == "__main__":
    main()

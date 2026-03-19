"""
Generador de Cargo de Entrega de EPP en Excel.

Salida: documentos_generados/EPP_{DNI}_{YYYYMMDD_HHMMSS}.xlsx
"""
from datetime import datetime

from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from src.config import CARPETA_FIRMAS, CARPETA_SALIDA, FIRMA_ALTO_PX

EPPS_BASICOS = [
    "Casco de seguridad",
    "Lentes de seguridad",
    "Protector auditivo",
    "Mascarilla / Respirador",
    "Guantes de seguridad",
    "Zapatos de seguridad",
    "Chaleco reflectivo",
]

_AZUL      = "1F4E79"
_AZUL_MED  = "2E75B6"
_AZUL_CLAR = "BDD7EE"
_BLANCO    = "FFFFFF"


def _side(style="thin"):
    return Side(style=style)


def _border():
    s = _side()
    return Border(left=s, right=s, top=s, bottom=s)


def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def _font(bold=False, color="000000", size=10, name="Century Gothic"):
    return Font(bold=bold, color=color, size=size, name=name)


def _aln(h="center", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def generar_registro_epp(trabajador: dict, epps: list[str], fecha: str, area: str = "Taller") -> str:
    """
    Genera el Excel de Cargo de Entrega de EPP para un trabajador.

    trabajador : dict con claves dni, nombre, apellido, cargo
    epps       : lista de nombres de equipos entregados
    fecha      : fecha de entrega (dd/mm/yyyy)
    area       : área del trabajador (por defecto "Taller")
    Retorna    : ruta del archivo generado
    """
    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)

    apellido = (trabajador.get("apellido") or "").strip()
    nombre_p = (trabajador.get("nombre")   or "").strip()
    nombre   = f"{apellido} {nombre_p}".strip() if apellido else nombre_p
    dni      = trabajador.get("dni", "")
    cargo    = trabajador.get("cargo") or ""

    wb = Workbook()
    ws = wb.active
    ws.title = "ENTREGA"

    # ── column widths ─────────────────────────────────────────────────────────
    anchos = {
        "A": 5,  "B": 13, "C": 20, "D": 18,
        "E": 18, "F": 18, "G": 10, "H": 10,
        "I": 12, "J": 15, "K": 10,
    }
    for col, w in anchos.items():
        ws.column_dimensions[col].width = w

    def _c(ref, value="", bold=False, bg=None, fg="000000",
           h="center", v="center", wrap=True, size=10):
        c = ws[ref]
        c.value     = value
        c.font      = _font(bold=bold, color=fg, size=size)
        c.alignment = _aln(h=h, v=v, wrap=wrap)
        c.border    = _border()
        if bg:
            c.fill = _fill(bg)
        return c

    def _merge(r1, c1, r2, c2):
        ws.merge_cells(start_row=r1, start_column=c1,
                       end_row=r2,   end_column=c2)

    # ── FILA 1-3: cabecera ────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 35
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 18

    _merge(1, 1, 3, 2)
    ws["A1"].border    = _border()
    ws["A1"].alignment = _aln()

    _merge(1, 3, 3, 8)
    _c("C1", "CARGO DE ENTREGA DE EPP", bold=True, size=14,
       fg=_BLANCO, bg=_AZUL_MED)

    _merge(1, 9, 1, 11)
    _c("I1", "Código: SGSST-P-19-F-01", size=9, bg=_AZUL_CLAR)
    _merge(2, 9, 2, 11)
    _c("I2", "Versión: Actualizada el 17.02.2023", size=9, bg=_AZUL_CLAR)
    _merge(3, 9, 3, 11)
    _c("I3", f"Fecha: {fecha}", size=9, bg=_AZUL_CLAR)

    # ── FILA 5: N° Registro ───────────────────────────────────────────────────
    ws.row_dimensions[5].height = 18
    _merge(5, 1, 5, 2)
    _c("A5", "N° Registro:", bold=True, bg=_AZUL_CLAR, h="right")
    _merge(5, 3, 5, 5)
    _c("C5", "")

    # ── FILA 7-9: datos del empleador ─────────────────────────────────────────
    ws.row_dimensions[7].height = 16
    ws.row_dimensions[8].height = 30
    ws.row_dimensions[9].height = 52

    _merge(7, 1, 7, 11)
    _c("A7", "DATOS DEL EMPLEADOR:", bold=True, fg=_BLANCO, bg=_AZUL, size=11)

    _merge(8, 1, 8, 2)
    _c("A8", "Razón\nSocial",          bold=True, bg=_AZUL_CLAR, size=9)
    _c("C8", "RUC",                    bold=True, bg=_AZUL_CLAR, size=9)
    _merge(8, 4, 8, 6)
    _c("D8", "Domicilio",              bold=True, bg=_AZUL_CLAR, size=9)
    _merge(8, 7, 8, 9)
    _c("G8", "Actividad económica",    bold=True, bg=_AZUL_CLAR, size=9)
    _merge(8, 10, 8, 11)
    _c("J8", "N° trabajadores",        bold=True, bg=_AZUL_CLAR, size=9)

    _merge(9, 1, 9, 2)
    _c("A9", "GRUAS MARA S.A.C.",      bold=True, size=9)
    _c("C9", "20525068162",            size=9)
    _merge(9, 4, 9, 6)
    _c("D9", "Av. Elmer Faucett 5068, Urb. Las Fresas, Callao", size=9, h="left")
    _merge(9, 7, 9, 9)
    _c("G9", "Alquiler de grúas móviles, camiones grúa, "
             "montacargas y equipos de elevación.", size=9)
    _merge(9, 10, 9, 11)
    _c("J9", "25", size=9)

    # ── FILA 11-13: datos del trabajador ──────────────────────────────────────
    ws.row_dimensions[11].height = 16
    ws.row_dimensions[12].height = 30
    ws.row_dimensions[13].height = 22

    _merge(11, 1, 11, 11)
    _c("A11", "DATOS DEL TRABAJADOR", bold=True, fg=_BLANCO, bg=_AZUL, size=11)

    _merge(12, 1, 12, 2)
    _c("A12", "NOMBRE DEL\nTRABAJADOR", bold=True, bg=_AZUL_CLAR, size=9)
    _merge(12, 3, 12, 8)
    _c("C12", nombre, bold=True, size=11, h="left")
    _c("I12", "DNI", bold=True, bg=_AZUL_CLAR, size=9)
    _merge(12, 10, 12, 11)
    _c("J12", dni, bold=True, size=11)

    _merge(13, 1, 13, 2)
    _c("A13", "ÁREA", bold=True, bg=_AZUL_CLAR, size=9)
    _merge(13, 3, 13, 5)
    _c("C13", area, size=10)
    _c("F13", "PUESTO", bold=True, bg=_AZUL_CLAR, size=9)
    _merge(13, 7, 13, 11)
    _c("G13", cargo, size=10, h="left")

    # ── FILA 15: cabecera de tabla ────────────────────────────────────────────
    ws.row_dimensions[15].height = 28

    _c("A15", "#",                bold=True, fg=_BLANCO, bg=_AZUL_MED, size=9)
    _c("B15", "Fecha",            bold=True, fg=_BLANCO, bg=_AZUL_MED, size=9)
    _merge(15, 3, 15, 4)
    _c("C15", "Equipo Entregado", bold=True, fg=_BLANCO, bg=_AZUL_MED, size=9)
    _merge(15, 5, 15, 7)
    _c("E15", "Firma (RECIBÍ)",   bold=True, fg=_BLANCO, bg=_AZUL_MED, size=9)
    _merge(15, 8, 15, 9)
    _c("H15", "Firma (ENTREGUÉ)", bold=True, fg=_BLANCO, bg=_AZUL_MED, size=9)
    _merge(15, 10, 15, 11)
    _c("J15", "Observaciones",    bold=True, fg=_BLANCO, bg=_AZUL_MED, size=9)

    # ── FILAS DE EPPs ─────────────────────────────────────────────────────────
    firma_path = CARPETA_FIRMAS / f"firma_{dni}.png"
    row_h      = max(FIRMA_ALTO_PX, 50) * 0.75 + 8   # px → pts aprox.

    for i, epp in enumerate(epps):
        fila = 16 + i
        ws.row_dimensions[fila].height = row_h

        _c(f"A{fila}", str(i + 1), size=10)
        _c(f"B{fila}", fecha, size=9)
        _merge(fila, 3, fila, 4)
        _c(f"C{fila}", epp, size=10, h="left")
        _merge(fila, 5, fila, 7)
        ws[f"E{fila}"].border = _border()
        _merge(fila, 8, fila, 9)
        ws[f"H{fila}"].border = _border()
        _merge(fila, 10, fila, 11)
        ws[f"J{fila}"].border = _border()

        if firma_path.exists():
            img        = ExcelImage(str(firma_path))
            img.height = FIRMA_ALTO_PX
            img.width  = int(FIRMA_ALTO_PX * 2.5)
            img.anchor = f"E{fila}"
            ws.add_image(img)

    # ── guardar ───────────────────────────────────────────────────────────────
    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"EPP_{dni}_{ts}.xlsx"
    ruta     = CARPETA_SALIDA / filename
    wb.save(str(ruta))
    return str(ruta)

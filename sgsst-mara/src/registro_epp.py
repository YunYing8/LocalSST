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

_AZUL      = "1D428A"   # azul oscuro (headers principales)
_AZUL_CLAR = "BDD7EE"   # azul claro (labels de datos)
_BLANCO    = "FFFFFF"

_RESPONSABLE_NOMBRE = "Alonso Homero Mancilla Tamara"
_RESPONSABLE_CARGO  = "Prevencionista de riesgos"


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

    # ── anchos de columna (igual que el template) ──────────────────────────────
    anchos = {
        "A": 6.57,  "B": 19.14, "C": 15.86, "D": 8.5,
        "E": 8.5,   "F": 19.86, "G": 8.5,   "H": 24.43,
        "I": 15.57, "J": 14.0,  "K": 16.0,
    }
    for col, w in anchos.items():
        ws.column_dimensions[col].width = w

    # ── helpers locales ────────────────────────────────────────────────────────
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

    def _sep(fila, altura=7.5):
        """Fila separadora sin contenido."""
        ws.row_dimensions[fila].height = altura

    # ── FILAS 1-3: cabecera ────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 35
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 18

    # espacio para logo (A1:B3)
    _merge(1, 1, 3, 2)
    ws["A1"].border    = _border()
    ws["A1"].alignment = _aln()

    # título
    _merge(1, 3, 3, 8)
    _c("C1", "CARGO DE ENTREGA DE EPP", bold=True, size=14,
       fg=_BLANCO, bg=_AZUL)

    # código / versión / fecha (label en I, valor en J:K)
    _merge(1, 10, 1, 11)
    _c("I1", "Código",    bold=False, size=9, bg=_AZUL_CLAR, h="left")
    _c("J1", "SGSST-P-19-F-01", bold=False, size=9, bg=_AZUL_CLAR)

    _merge(2, 10, 2, 11)
    _c("I2", "Versión",   bold=False, size=9, bg=_AZUL_CLAR, h="left")
    _c("J2", "Actualizada el 17.02.2023", bold=False, size=9, bg=_AZUL_CLAR)

    _merge(3, 10, 3, 11)
    _c("I3", "Fecha",     bold=False, size=9, bg=_AZUL_CLAR, h="left")
    _c("J3", fecha,       bold=False, size=9, bg=_AZUL_CLAR)

    # ── FILA 4: separador ─────────────────────────────────────────────────────
    _sep(4)

    # ── FILA 5: N° Registro ───────────────────────────────────────────────────
    ws.row_dimensions[5].height = 20.25
    _merge(5, 1, 5, 2)
    _c("A5", "N° Registro:", bold=True, bg=_AZUL, fg=_BLANCO, h="right")
    _merge(5, 3, 5, 4)
    _c("C5", "")

    # ── FILA 6: separador ─────────────────────────────────────────────────────
    _sep(6)

    # ── FILAS 7-9: datos del empleador ────────────────────────────────────────
    ws.row_dimensions[7].height = 18.75
    ws.row_dimensions[8].height = 26.25
    ws.row_dimensions[9].height = 61.9

    _merge(7, 1, 7, 11)
    _c("A7", "DATOS DEL EMPLEADOR:", bold=True, fg=_BLANCO, bg=_AZUL, size=11)

    # labels fila 8
    _merge(8, 1, 8, 2)
    _c("A8", "Razón\nSocial",       bold=True, bg=_AZUL_CLAR, size=9)
    _merge(8, 3, 8, 4)
    _c("C8", "RUC",                 bold=True, bg=_AZUL_CLAR, size=9)
    _merge(8, 5, 8, 8)
    _c("E8", "Domicilio",           bold=True, bg=_AZUL_CLAR, size=9)
    _merge(8, 9, 8, 10)
    _c("I8", "Actividad económica", bold=True, bg=_AZUL_CLAR, size=9)
    _c("K8", "N° trabajadores",     bold=True, bg=_AZUL_CLAR, size=9)

    # valores fila 9
    _merge(9, 1, 9, 2)
    _c("A9", "GRUAS MARA S.A.C.",   bold=True, size=9)
    _merge(9, 3, 9, 4)
    _c("C9", "20525068162",         size=9)
    _merge(9, 5, 9, 8)
    _c("E9", "Av. Elmer Faucett 5068, Urb. Las Fresas, Callao", size=9, h="left")
    _merge(9, 9, 9, 10)
    _c("I9", "Alquiler de grúas móviles, camiones grúa, "
             "montacargas y equipos de elevación.", size=9)
    _c("K9", "25", size=9)

    # ── FILA 10: separador ────────────────────────────────────────────────────
    _sep(10, 10.9)

    # ── FILAS 11-13: datos del trabajador ─────────────────────────────────────
    ws.row_dimensions[11].height = 18.75
    ws.row_dimensions[12].height = 27.0
    ws.row_dimensions[13].height = 27.0

    _merge(11, 1, 11, 11)
    _c("A11", "DATOS DEL TRABAJADOR", bold=True, fg=_BLANCO, bg=_AZUL, size=11)

    # fila 12: nombre + dni
    _merge(12, 1, 12, 2)
    _c("A12", "NOMBRE DEL\nTRABAJADOR", bold=True, bg=_AZUL_CLAR, size=9)
    _merge(12, 3, 12, 8)
    _c("C12", nombre, bold=True, size=11, h="left")
    _c("I12", "DNI", bold=True, bg=_AZUL_CLAR, size=9)
    _merge(12, 10, 12, 11)
    _c("J12", dni, bold=True, size=11)

    # fila 13: área + puesto
    _merge(13, 1, 13, 2)
    _c("A13", "ÁREA",   bold=True, bg=_AZUL_CLAR, size=9)
    _merge(13, 3, 13, 6)
    _c("C13", area,     size=10, bg=_BLANCO)
    _c("G13", "PUESTO", bold=True, bg=_AZUL_CLAR, size=9)
    _merge(13, 8, 13, 11)
    _c("H13", cargo,    size=10, h="left")

    # ── FILA 14: separador ────────────────────────────────────────────────────
    _sep(14)

    # ── FILA 15: cabecera de tabla ────────────────────────────────────────────
    ws.row_dimensions[15].height = 18.75

    _c("A15", "#",                bold=True, fg=_BLANCO, bg=_AZUL, size=9)
    _c("B15", "Fecha",            bold=True, fg=_BLANCO, bg=_AZUL, size=9)
    _merge(15, 3, 15, 5)
    _c("C15", "Equipo Entregado", bold=True, fg=_BLANCO, bg=_AZUL, size=9)
    _merge(15, 6, 15, 7)
    _c("F15", "Firma (RECIBÍ)",   bold=True, fg=_BLANCO, bg=_AZUL, size=9)
    _merge(15, 8, 15, 9)
    _c("H15", "Firma (ENTREGUÉ)", bold=True, fg=_BLANCO, bg=_AZUL, size=9)
    _merge(15, 10, 15, 11)
    _c("J15", "Observaciones",    bold=True, fg=_BLANCO, bg=_AZUL, size=9)

    # ── FILAS DE EPPs (siempre 20 filas) ──────────────────────────────────────
    firma_path = CARPETA_FIRMAS / f"firma_{dni}.png"
    row_h      = 62.25   # alto de fila igual al template

    for i in range(20):
        fila = 16 + i
        ws.row_dimensions[fila].height = row_h

        epp = epps[i] if i < len(epps) else ""
        epp_fecha = fecha if epp else ""

        _c(f"A{fila}", str(i + 1), size=10, bold=True, bg=_BLANCO)
        _c(f"B{fila}", epp_fecha, size=9, bg=_BLANCO)
        _merge(fila, 3, fila, 5)
        _c(f"C{fila}", epp, size=10, h="left", bg=_BLANCO)
        _merge(fila, 6, fila, 7)
        ws[f"F{fila}"].border = _border()
        _merge(fila, 8, fila, 9)
        ws[f"H{fila}"].border = _border()
        _merge(fila, 10, fila, 11)
        ws[f"J{fila}"].border = _border()

        # firma del trabajador en columna F (RECIBÍ)
        if epp and firma_path.exists():
            img        = ExcelImage(str(firma_path))
            img.height = FIRMA_ALTO_PX
            img.width  = int(FIRMA_ALTO_PX * 2.5)
            img.anchor = f"F{fila}"
            ws.add_image(img)

    # ── FILA 35: separador ────────────────────────────────────────────────────
    _sep(35)

    # ── FILAS 36-39: responsable del registro ─────────────────────────────────
    ws.row_dimensions[36].height = 18.75
    ws.row_dimensions[37].height = 23.25
    ws.row_dimensions[38].height = 30.0
    ws.row_dimensions[39].height = 57.75

    _merge(36, 1, 36, 11)
    _c("A36", "RESPONSABLE DEL REGISTRO", bold=True, fg=_BLANCO, bg=_AZUL, size=11)

    # fila 37-38: nombre + firma
    _merge(37, 1, 38, 2)
    _c("A37", "Nombre", bold=True, bg=_AZUL_CLAR, size=9)
    _merge(37, 3, 38, 7)
    _c("C37", _RESPONSABLE_NOMBRE, bold=True, size=10, h="left")
    _merge(37, 8, 37, 11)
    _c("H37", "Firma", bold=True, bg=_AZUL_CLAR, size=9)
    _merge(38, 8, 39, 11)
    ws["H38"].border = _border()

    # fila 39: cargo
    _merge(39, 1, 39, 2)
    _c("A39", "Cargo", bold=True, bg=_AZUL_CLAR, size=9)
    _merge(39, 3, 39, 7)
    _c("C39", _RESPONSABLE_CARGO, bold=True, size=10, h="left")

    # ── guardar ───────────────────────────────────────────────────────────────
    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"EPP_{dni}_{ts}.xlsx"
    ruta     = CARPETA_SALIDA / filename
    wb.save(str(ruta))
    return str(ruta)

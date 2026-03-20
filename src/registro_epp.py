"""
Generador de Cargo de Entrega de EPP en Excel.

Carga la plantilla EPPs.xlsx y rellena los datos del trabajador y EPPs.
Salida: documentos_generados/EPP_{DNI}_{YYYYMMDD_HHMMSS}.xlsx
"""
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage

from src.config import CARPETA_FIRMAS, CARPETA_SALIDA, PROJECT_ROOT

PLANTILLA_EPP         = PROJECT_ROOT / "data" / "plantillas" / "EPPs.xlsx"
# DNI del prevencionista (Alonso Mancilla) — firma siempre la columna ENTREGUÉ
DNI_PREVENCIONISTA    = "48578094"

EPPS_BASICOS = [
    "Casco de seguridad",
    "Lentes de seguridad",
    "Protector auditivo",
    "Mascarilla / Respirador",
    "Guantes de seguridad",
    "Zapatos de seguridad",
    "Chaleco reflectivo",
]


def _insertar_firma(ws, ruta_img, celda: str) -> None:
    """Inserta la imagen en la celda indicada sin alterar su tamaño original."""
    if ruta_img.exists():
        img        = ExcelImage(str(ruta_img))
        img.anchor = celda
        ws.add_image(img)


def generar_registro_epp(trabajador: dict, epps: list[str], fecha: str) -> str:
    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)

    apellido = (trabajador.get("apellido") or "").strip()
    nombre_p = (trabajador.get("nombre")   or "").strip()
    nombre   = f"{apellido} {nombre_p}".strip() if apellido else nombre_p
    dni      = trabajador.get("dni", "")
    cargo    = trabajador.get("cargo") or ""

    wb = load_workbook(str(PLANTILLA_EPP))
    ws = wb.active

    # ── Fecha ────────────────────────────────────────────────────────────────
    ws["I3"] = f"Fecha: {fecha}"

    # ── Datos del trabajador ─────────────────────────────────────────────────
    ws["C12"] = nombre
    ws["J12"] = dni
    ws["H13"] = cargo          # cargo en H13

    # ── Rutas de firma ───────────────────────────────────────────────────────
    firma_trabajador    = CARPETA_FIRMAS / f"firma_{dni}.png"
    firma_prevencionista = CARPETA_FIRMAS / f"firma_{DNI_PREVENCIONISTA}.png"

    # ── Filas de EPP (desde fila 16) ─────────────────────────────────────────
    for i, epp in enumerate(epps):
        fila = 16 + i

        ws[f"A{fila}"] = i + 1
        ws[f"B{fila}"] = fecha
        ws[f"C{fila}"] = epp

        # Firma (RECIBÍ)   — columna E — firma del trabajador
        _insertar_firma(ws, firma_trabajador,    f"E{fila}")
        # Firma (ENTREGUÉ) — columna H — firma de Alonso Mancilla
        _insertar_firma(ws, firma_prevencionista, f"H{fila}")

    # ── Firma del prevencionista en H38 ──────────────────────────────────────
    _insertar_firma(ws, firma_prevencionista, "H38")

    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"EPP_{dni}_{ts}.xlsx"
    ruta     = CARPETA_SALIDA / filename
    wb.save(str(ruta))
    return str(ruta)

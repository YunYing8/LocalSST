"""
Configuración centralizada del proyecto.
Todos los módulos deben importar rutas y constantes desde aquí.
"""
import os
from pathlib import Path

# ── raíz del proyecto ────────────────────────────────────────────────────────
PROJECT_ROOT = Path(__file__).resolve().parent.parent

# ── rutas de datos ───────────────────────────────────────────────────────────
DATA_DIR              = PROJECT_ROOT / "data"
DB_PATH               = DATA_DIR / "sgsst.db"
CARPETA_FIRMAS_ORIG   = DATA_DIR / "firmas_originales"
CARPETA_FIRMAS        = DATA_DIR / "firmas"
PLANTILLA_RISST       = DATA_DIR / "plantillas" / "plantilla_risst.docx"
PLANTILLA_ASISTENCIA  = DATA_DIR / "plantillas" / "plantilla_asistencia.xlsx"
PLANTILLA_EPPS        = DATA_DIR / "plantillas" / "EPPs.xlsx"
CARPETA_SALIDA        = PROJECT_ROOT / "documentos_generados"

# ── Google Sheets ─────────────────────────────────────────────────────────────
GOOGLE_SHEET_ID = os.environ.get(
    "SGSST_SHEET_ID", "1N0i-sG3UBH2UMVo_DF4Ivcwxyr2MvgIfAa4h18oQ5oc"
)

# ── remove.bg ────────────────────────────────────────────────────────────────
REMOVEBG_API_KEY = os.environ.get("REMOVEBG_API_KEY", "TW85KXjQCpmyVS35yYoeybso")
REMOVEBG_URL     = "https://api.remove.bg/v1.0/removebg"

# ── firma en constancias ─────────────────────────────────────────────────────
FIRMA_ALTO_PX = 65

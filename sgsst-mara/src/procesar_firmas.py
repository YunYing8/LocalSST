"""
Procesador de firmas usando rembg (local, sin API key, sin costo).

Estructura esperada:
    data/firmas_originales/firma_XXXXXXXX.png   ← archivos de entrada
    data/firmas/firma_XXXXXXXX.png              ← archivos procesados

El nombre del archivo debe seguir el patrón: firma_{DNI}.{ext}
"""
from pathlib import Path
from PIL import Image
from rembg import remove
import io

from src.database import DB_PATH

PROJECT_ROOT      = DB_PATH.parent.parent
CARPETA_ENTRADA   = PROJECT_ROOT / "data" / "firmas_originales"
CARPETA_SALIDA    = PROJECT_ROOT / "data" / "firmas"
ALTO_FINAL        = 65
_EXTENSIONES      = {'.png', '.jpg', '.jpeg'}


def procesar_firma(ruta_entrada: Path, ruta_salida: Path) -> None:
    """Quita el fondo y redimensiona una firma. Lanza excepción si falla."""
    with open(ruta_entrada, 'rb') as f:
        datos = f.read()

    output = remove(datos)
    imagen = Image.open(io.BytesIO(output)).convert("RGBA")

    ancho_orig, alto_orig = imagen.size
    proporcion  = ALTO_FINAL / alto_orig
    ancho_final = int(ancho_orig * proporcion)
    imagen_redim = imagen.resize((ancho_final, ALTO_FINAL), Image.Resampling.LANCZOS)

    ruta_salida.parent.mkdir(parents=True, exist_ok=True)
    imagen_redim.save(str(ruta_salida), format='PNG')


def procesar_todas_firmas() -> dict:
    """
    Procesa todas las imágenes en CARPETA_ENTRADA que aún no tienen
    su versión procesada en CARPETA_SALIDA.
    Retorna {"procesadas": int, "omitidas": int, "errores": list[str]}
    """
    if not CARPETA_ENTRADA.exists():
        return {
            "procesadas": 0,
            "omitidas":   0,
            "errores":    [f"Carpeta no encontrada: {CARPETA_ENTRADA}"],
        }

    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)

    procesadas = 0
    omitidas   = 0
    errores: list[str] = []

    archivos = [
        f for f in sorted(CARPETA_ENTRADA.iterdir())
        if f.suffix.lower() in _EXTENSIONES
    ]

    for archivo in archivos:
        # Normalizar a firma_{dni}.png sin importar el nombre original
        stem = archivo.stem
        if not stem.startswith('firma_'):
            stem = f'firma_{stem}'
        ruta_salida = CARPETA_SALIDA / (stem + '.png')

        if ruta_salida.exists():
            omitidas += 1
            continue

        try:
            procesar_firma(archivo, ruta_salida)
            procesadas += 1
        except Exception as exc:
            errores.append(f"{archivo.name}: {exc}")

    return {"procesadas": procesadas, "omitidas": omitidas, "errores": errores}

"""
Procesador de firmas usando la API de remove.bg.

Estructura esperada:
    data/firmas_originales/firma_XXXXXXXX.png   ← archivos de entrada
    data/firmas/firma_XXXXXXXX.png              ← archivos procesados
"""
from io import BytesIO
from pathlib import Path

import requests
from PIL import Image

from src.config import (
    CARPETA_FIRMAS,
    CARPETA_FIRMAS_ORIG,
    FIRMA_ALTO_PX,
    REMOVEBG_API_KEY,
    REMOVEBG_URL,
)

_EXTENSIONES = {'.png', '.jpg', '.jpeg'}


def procesar_firma(ruta_entrada: Path, ruta_salida: Path) -> None:
    """Quita el fondo vía remove.bg y redimensiona una firma."""
    with open(ruta_entrada, 'rb') as f:
        datos = f.read()

    response = requests.post(
        REMOVEBG_URL,
        files={'image_file': datos},
        headers={'X-Api-Key': REMOVEBG_API_KEY},
        data={'size': 'auto'},
        timeout=30,
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"remove.bg error {response.status_code}: {response.text[:200]}"
        )

    imagen = Image.open(BytesIO(response.content))
    ancho_orig, alto_orig = imagen.size
    proporcion   = FIRMA_ALTO_PX / alto_orig
    ancho_final  = int(ancho_orig * proporcion)
    imagen_redim = imagen.resize((ancho_final, FIRMA_ALTO_PX), Image.Resampling.LANCZOS)

    ruta_salida.parent.mkdir(parents=True, exist_ok=True)
    imagen_redim.save(str(ruta_salida), format='PNG')


def procesar_todas_firmas() -> dict:
    """
    Procesa las imágenes en CARPETA_FIRMAS_ORIG que aún no están procesadas.
    Retorna {"procesadas": int, "omitidas": int, "errores": list[str]}
    """
    if not CARPETA_FIRMAS_ORIG.exists():
        return {
            "procesadas": 0,
            "omitidas":   0,
            "errores":    [f"Carpeta no encontrada: {CARPETA_FIRMAS_ORIG}"],
        }

    CARPETA_FIRMAS.mkdir(parents=True, exist_ok=True)

    procesadas = 0
    omitidas   = 0
    errores: list[str] = []

    for archivo in sorted(CARPETA_FIRMAS_ORIG.iterdir()):
        if archivo.suffix.lower() not in _EXTENSIONES:
            continue

        stem = archivo.stem
        if not stem.startswith('firma_'):
            stem = f'firma_{stem}'
        ruta_salida = CARPETA_FIRMAS / (stem + '.png')

        if ruta_salida.exists():
            omitidas += 1
            continue

        try:
            procesar_firma(archivo, ruta_salida)
            procesadas += 1
        except Exception as exc:
            errores.append(f"{archivo.name}: {exc}")

    return {"procesadas": procesadas, "omitidas": omitidas, "errores": errores}

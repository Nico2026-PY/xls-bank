import re
import sys
from pathlib import Path


def ruta_recurso(ruta_relativa):
    """
    Obtiene la ruta correcta al ejecutar desde Python o desde un .exe.

    Soporta:
    - Desarrollo: proyecto/assets
    - Desarrollo desde app_src/app.py
    - PyInstaller: carpeta temporal _MEIPASS
    """
    ruta_relativa = Path(ruta_relativa)

    if hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS) / ruta_relativa

    archivo_actual = Path(__file__).resolve()

    # utils.py está en app_src/xlsbank/utils.py
    carpeta_xlsbank = archivo_actual.parent
    carpeta_app_src = carpeta_xlsbank.parent
    carpeta_proyecto = carpeta_app_src.parent

    posibles = [
        carpeta_proyecto / ruta_relativa,
        carpeta_app_src / ruta_relativa,
        carpeta_xlsbank / ruta_relativa,
        Path.cwd() / ruta_relativa,
    ]

    for ruta in posibles:
        if ruta.exists():
            return ruta

    return carpeta_proyecto / ruta_relativa


def norm(txt):
    txt = '' if txt is None else str(txt)
    txt = txt.strip().lower()
    txt = txt.replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u').replace('ü', 'u')
    return re.sub(r'\s+', ' ', txt)


def limpiar_numero(valor):
    if valor is None:
        return 0.0

    try:
        import pandas as pd
        if isinstance(valor, float) and pd.isna(valor):
            return 0.0
        if pd.isna(valor):
            return 0.0
    except Exception:
        pass

    if isinstance(valor, (int, float)):
        return float(valor)

    s = str(valor).strip()
    if s == '' or s.lower() in ['nan', 'none']:
        return 0.0

    s = s.replace('$', '').replace(' ', '')

    # Formato argentino: 1.234.567,89
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')

    try:
        return float(s)
    except Exception:
        return 0.0
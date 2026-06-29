import re
import sys
import pandas as pd

from datetime import datetime
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
    
def limpiar_fecha(valor):
    """
    Convierte fechas de bancos a fecha real.

    Corrige especialmente Patagonia, que a veces trae la fecha como texto
    tipo 14/6/26 o 14/05/26. Si se ordena eso como texto, Excel/Python
    mezcla meses. Con esta función siempre se ordena como fecha real.
    """
    if valor is None:
        return pd.NaT

    try:
        if pd.isna(valor):
            return pd.NaT
    except Exception:
        pass

    if isinstance(valor, pd.Timestamp):
        return valor.to_pydatetime()

    if isinstance(valor, datetime):
        return valor

    # Por si viene como date de Python, sin ser datetime.
    if hasattr(valor, 'year') and hasattr(valor, 'month') and hasattr(valor, 'day') and not isinstance(valor, str):
        try:
            return datetime(valor.year, valor.month, valor.day)
        except Exception:
            pass

    # Fechas seriales de Excel.
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        try:
            numero = float(valor)
            if 20000 <= numero <= 80000:
                fecha = pd.to_datetime('1899-12-30') + pd.to_timedelta(numero, unit='D')
                return fecha.to_pydatetime()
        except Exception:
            pass

    s = str(valor).strip()
    if s == '' or s.lower() in ['nan', 'none', 'nat']:
        return pd.NaT

    # Argentina: día/mes/año. Soporta 14/6/26, 14/06/2026, 14-06-26, etc.
    fecha = pd.to_datetime(s, errors='coerce', dayfirst=True)

    if pd.isna(fecha):
        return pd.NaT

    return fecha.to_pydatetime()


def normalizar_fechas_movimientos(df):
    """Normaliza columnas de fecha para que todas se ordenen igual."""
    df = df.copy()

    for col in ['Fecha Valor', 'Fecha Operación']:
        if col in df.columns:
            df[col] = df[col].apply(limpiar_fecha)

    return df


def ordenar_por_fecha_real(df, descendente=True):
    """Ordena movimientos por fecha real, no por texto."""
    df = normalizar_fechas_movimientos(df)

    if 'Fecha Valor' not in df.columns:
        return df

    fecha_orden = pd.to_datetime(df['Fecha Valor'], errors='coerce')

    if 'Fecha Operación' in df.columns:
        fecha_operacion = pd.to_datetime(df['Fecha Operación'], errors='coerce')
        fecha_orden = fecha_orden.fillna(fecha_operacion)

    df['_FechaOrden'] = fecha_orden
    df = df.sort_values(
        '_FechaOrden',
        ascending=not descendente,
        na_position='last',
        kind='mergesort'
    )

    return df.drop(columns=['_FechaOrden'])
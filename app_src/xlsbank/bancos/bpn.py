import pandas as pd

from xlsbank.config import COLUMNAS_SALIDA
from xlsbank.utils import limpiar_numero


def procesar_bpn(path, read_excel_any):
    """
    Procesa extractos BPN con estructura tabular estándar.

    Se recibe read_excel_any como parámetro para no duplicar todavía
    la función de lectura de Excel/CSV.
    """
    df = read_excel_any(path, header=0)
    df.columns = [str(c).strip() for c in df.columns]

    cuenta = (
        df['Cuenta'].dropna().iloc[0]
        if 'Cuenta' in df.columns and not df['Cuenta'].dropna().empty
        else ''
    )

    out = pd.DataFrame()
    out['Cuenta'] = df.get('Cuenta', cuenta)
    out['Fecha Valor'] = df.get('Fecha Valor', '')
    out['Fecha Operación'] = df.get('Fecha Operación', '')
    out['Movimiento Fecha-Valor'] = df.get('Movimiento Fecha-Valor', '')
    out['Descripción'] = df.get('Descripción', '')
    out['Detalle'] = df.get('Detalle', '')
    out['Importe'] = df['Importe'].apply(limpiar_numero) if 'Importe' in df.columns else ''
    out['Saldo'] = df['Saldo'].apply(limpiar_numero) if 'Saldo' in df.columns else ''
    out['Categoria'] = df.get('Categoria', '')
    out['Referencia'] = df.get('Referencia', '')
    out['Etiquetas'] = df.get('Etiquetas', '')

    return out[COLUMNAS_SALIDA]
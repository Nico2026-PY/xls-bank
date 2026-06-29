import pandas as pd

from xlsbank.config import COLUMNAS_SALIDA
from xlsbank.utils import limpiar_numero


def procesar_galicia(path, read_excel_any):
    """
    Procesa extractos Galicia con columnas Fecha, Descripción,
    Débitos, Créditos, Saldo y datos adicionales si existen.
    """
    df = read_excel_any(path, header=0)
    df.columns = [str(c).strip() for c in df.columns]

    cuenta = 'GALICIA'

    deb = df['Débitos'].apply(limpiar_numero) if 'Débitos' in df.columns else 0
    cre = df['Créditos'].apply(limpiar_numero) if 'Créditos' in df.columns else 0

    out = pd.DataFrame()
    out['Cuenta'] = cuenta
    out['Fecha Valor'] = df.get('Fecha', '')
    out['Fecha Operación'] = df.get('Fecha', '')
    out['Movimiento Fecha-Valor'] = ''
    out['Descripción'] = df.get('Descripción', '')

    partes = []
    for c in [
        'Observaciones Cliente',
        'Leyendas Adicionales 1',
        'Leyendas Adicionales 2',
        'Leyendas Adicionales 3'
    ]:
        if c in df.columns:
            partes.append(df[c].fillna('').astype(str))

    if partes:
        detalle = partes[0]
        for p in partes[1:]:
            detalle = detalle.str.cat(p, sep=' ', na_rep='')
    else:
        detalle = ''

    out['Detalle'] = detalle
    out['Importe'] = cre - deb
    out['Saldo'] = df['Saldo'].apply(limpiar_numero) if 'Saldo' in df.columns else 0
    out['Categoria'] = ''
    out['Referencia'] = df.get('Número de Comprobante', '')
    out['Etiquetas'] = ''

    return out[COLUMNAS_SALIDA]
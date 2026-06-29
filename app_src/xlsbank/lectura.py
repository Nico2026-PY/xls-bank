from pathlib import Path

import pandas as pd

from xlsbank.utils import norm


def read_excel_any(path, header=None):
    """
    Lee .xlsx, .xls y .csv.

    Para CSV usa separador ; porque Mercado Pago exporta así.
    """
    ext = Path(path).suffix.lower()

    if ext == '.csv':
        try:
            return pd.read_csv(
                path,
                header=header,
                dtype=object,
                sep=';',
                encoding='utf-8-sig'
            )
        except UnicodeDecodeError:
            return pd.read_csv(
                path,
                header=header,
                dtype=object,
                sep=';',
                encoding='latin1'
            )

    if ext == '.xls':
        return pd.read_excel(path, header=header, dtype=object, engine='xlrd')

    return pd.read_excel(path, header=header, dtype=object, engine='openpyxl')


def encontrar_fila_encabezado(df_raw, nombres):
    objetivos = [norm(x) for x in nombres]

    for i in range(len(df_raw)):
        vals = [norm(x) for x in df_raw.iloc[i].tolist()]
        aciertos = sum(1 for o in objetivos if o in vals)

        if aciertos >= max(2, min(3, len(objetivos))):
            return i

    return None
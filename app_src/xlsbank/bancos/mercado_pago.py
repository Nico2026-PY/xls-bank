import pandas as pd

from xlsbank.config import COLUMNAS_SALIDA
from xlsbank.utils import norm, limpiar_numero

def col_mp(df, nombre, default=''):
    """Devuelve columna de Mercado Pago si existe, si no devuelve default."""
    if nombre in df.columns:
        return df[nombre]
    return default


def describir_tipo_mp(transaction_type):
    """Traduce tipos de operación de Mercado Pago a textos más claros."""
    t = norm(transaction_type)

    traducciones = {
        'settlement': 'Liquidación',
        'payment': 'Pago',
        'refund': 'Devolución',
        'chargeback': 'Contracargo',
        'withdrawal': 'Retiro de dinero',
        'payout': 'Retiro de dinero',
        'transfer': 'Transferencia',
    }

    return traducciones.get(t, str(transaction_type).strip() or 'Movimiento')


def describir_medio_mp(payment_method_type):
    """Traduce medios de pago de Mercado Pago."""
    t = norm(payment_method_type)

    traducciones = {
        'available_money': 'Dinero disponible',
        'account_money': 'Dinero en cuenta',
        'debit_card': 'Tarjeta débito',
        'credit_card': 'Tarjeta crédito',
        'bank_transfer': 'Transferencia bancaria',
        'ticket': 'Efectivo / cupón',
        'digital_currency': 'Moneda digital',
    }

    return traducciones.get(t, str(payment_method_type).strip() or '')


def describir_subunidad_mp(sub_unit):
    """Traduce subunidad/canal de Mercado Pago."""
    t = norm(sub_unit)

    traducciones = {
        'qr': 'QR',
        'point': 'Point',
        'checkout': 'Checkout',
        'available_money': 'Dinero disponible',
        'debit_card': 'Tarjeta débito',
        'credit_card': 'Tarjeta crédito',
        'bank_transfer': 'Transferencia bancaria',
    }

    return traducciones.get(t, str(sub_unit).strip() or '')

def procesar_mercado_pago(path, read_excel_any):
    """
    Procesa reportes de Mercado Pago Argentina.

    Soporta:
    1) Formato viejo/en español.
    2) CSV nuevo con columnas SOURCE_ID, TRANSACTION_TYPE, REAL_AMOUNT, etc.
    """
    df = read_excel_any(path, header=0)
    df.columns = [str(c).strip() for c in df.columns]

    columnas = set(df.columns)

    # ============================================================
    # FORMATO NUEVO CSV MERCADO PAGO
    # ============================================================
    columnas_mp_csv = {
        'SOURCE_ID',
        'PAYMENT_METHOD_TYPE',
        'TRANSACTION_TYPE',
        'TRANSACTION_AMOUNT',
        'TRANSACTION_DATE',
        'REAL_AMOUNT'
    }

    if columnas_mp_csv.issubset(columnas):
        out = pd.DataFrame(index=df.index)

        fecha_operacion = pd.to_datetime(
            df['TRANSACTION_DATE'],
            errors='coerce',
            utc=True
        ).dt.tz_convert(None)

        if 'SETTLEMENT_DATE' in df.columns:
            fecha_valor = pd.to_datetime(
                df['SETTLEMENT_DATE'],
                errors='coerce',
                utc=True
            ).dt.tz_convert(None)
        elif 'MONEY_RELEASE_DATE' in df.columns:
            fecha_valor = pd.to_datetime(
                df['MONEY_RELEASE_DATE'],
                errors='coerce',
                utc=True
            ).dt.tz_convert(None)
        else:
            fecha_valor = fecha_operacion

        importe_neto = df['REAL_AMOUNT'].apply(limpiar_numero)

        tipo_raw = df['TRANSACTION_TYPE'].fillna('').astype(str)
        medio_raw = df['PAYMENT_METHOD_TYPE'].fillna('').astype(str)

        tipo_claro = tipo_raw.apply(describir_tipo_mp)
        medio_claro = medio_raw.apply(describir_medio_mp)

        if 'SUB_UNIT' in df.columns:
            sub_raw = df['SUB_UNIT'].fillna('').astype(str)
            sub_claro = sub_raw.apply(describir_subunidad_mp)
        else:
            sub_raw = pd.Series('', index=df.index)
            sub_claro = pd.Series('', index=df.index)

        descripcion = tipo_claro

        for extra in [medio_claro, sub_claro]:
            extra = extra.fillna('').astype(str).str.strip()
            descripcion = descripcion.where(
                extra.eq(''),
                descripcion + ' - ' + extra
            )

        detalle_partes = []

        for c in ['BUSINESS_UNIT', 'SUB_UNIT', 'PAYMENT_METHOD_TYPE']:
            if c in df.columns:
                detalle_partes.append(df[c].fillna('').astype(str))

        if detalle_partes:
            detalle = detalle_partes[0]
            for parte in detalle_partes[1:]:
                detalle = detalle.str.cat(parte, sep=' | ', na_rep='')
        else:
            detalle = pd.Series('', index=df.index)

        out['Cuenta'] = 'Mercado Pago Argentina'
        out['Fecha Valor'] = fecha_valor
        out['Fecha Operación'] = fecha_operacion
        out['Movimiento Fecha-Valor'] = ''
        out['Descripción'] = descripcion
        out['Detalle'] = detalle
        out['Importe'] = importe_neto

        # Mercado Pago no informa saldo de cuenta en este CSV.
        out['Saldo'] = ''

        out['Categoria'] = ''
        out['Referencia'] = df['SOURCE_ID'].fillna('').astype(str)
        out['Etiquetas'] = ''

        return out[COLUMNAS_SALIDA]

    # ============================================================
    # FORMATO VIEJO / ESPAÑOL MERCADO PAGO
    # ============================================================
    columnas_necesarias = [
        'FECHA DE APROBACIÓN',
        'TIPO DE OPERACIÓN',
        'MONTO NETO DE LA OPERACIÓN QUE IMPACTÓ TU DINERO'
    ]

    faltantes = [c for c in columnas_necesarias if c not in df.columns]
    if faltantes:
        raise ValueError(
            'No se reconoce formato Mercado Pago. Faltan columnas: '
            + ', '.join(faltantes)
        )

    out = pd.DataFrame(index=df.index)

    fecha_aprobacion = pd.to_datetime(
        df['FECHA DE APROBACIÓN'],
        errors='coerce',
        utc=True
    ).dt.tz_convert(None)

    fecha_liquidacion = (
        pd.to_datetime(
            df['FECHA DE LIQUIDACIÓN DEL DINERO'],
            errors='coerce',
            utc=True
        ).dt.tz_convert(None)
        if 'FECHA DE LIQUIDACIÓN DEL DINERO' in df.columns
        else fecha_aprobacion
    )

    importe_neto = df['MONTO NETO DE LA OPERACIÓN QUE IMPACTÓ TU DINERO'].apply(limpiar_numero)

    tipo_operacion = df['TIPO DE OPERACIÓN'].fillna('').astype(str)

    if 'DETALLE DE LA VENTA' in df.columns:
        detalle_venta = df['DETALLE DE LA VENTA'].fillna('').astype(str)
        descripcion = tipo_operacion + ' - ' + detalle_venta
    else:
        descripcion = tipo_operacion

    partes = []
    for c in ['PAGADOR', 'MEDIO DE PAGO', 'BANCO DE ORIGEN', 'NOMBRE DE LOCAL', 'CANAL DE VENTA']:
        if c in df.columns:
            partes.append(df[c].fillna('').astype(str))

    if partes:
        detalle = partes[0]
        for p in partes[1:]:
            detalle = detalle.str.cat(p, sep=' | ', na_rep='')
    else:
        detalle = pd.Series('', index=df.index)

    if 'ID DE OPERACIÓN EN MERCADO PAGO' in df.columns:
        referencia = df['ID DE OPERACIÓN EN MERCADO PAGO'].fillna('').astype(str)
    elif 'NÚMERO DE IDENTIFICACIÓN' in df.columns:
        referencia = df['NÚMERO DE IDENTIFICACIÓN'].fillna('').astype(str)
    else:
        referencia = pd.Series('', index=df.index)

    out['Cuenta'] = 'Mercado Pago Argentina'
    out['Fecha Valor'] = fecha_liquidacion
    out['Fecha Operación'] = fecha_aprobacion
    out['Movimiento Fecha-Valor'] = ''
    out['Descripción'] = descripcion
    out['Detalle'] = detalle
    out['Importe'] = importe_neto
    out['Saldo'] = ''
    out['Categoria'] = ''
    out['Referencia'] = referencia
    out['Etiquetas'] = ''

    return out[COLUMNAS_SALIDA]
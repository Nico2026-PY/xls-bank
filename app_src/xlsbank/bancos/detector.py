from pathlib import Path

from xlsbank.config import BANCOS_CLAVES
from xlsbank.utils import norm


def detectar_banco_detalle(path, df_raw=None):
    """Devuelve banco detectado y método de detección."""
    n = norm(Path(path).stem)

    for banco, claves in BANCOS_CLAVES.items():
        if any(c in n for c in claves):
            return banco, 'nombre_archivo'

    if df_raw is not None:
        try:
            texto = ' '.join(df_raw.head(30).astype(str).fillna('').values.ravel())
            t = norm(texto)

            for banco, claves in BANCOS_CLAVES.items():
                if any(c in t for c in claves):
                    return banco, 'contenido_archivo'
        except Exception:
            pass

    return 'BANCO', 'desconocido'


def detectar_banco(path, df_raw=None):
    banco, _ = detectar_banco_detalle(path, df_raw)
    return banco


def detectar_empresa_detalle(path, empresas_claves, df_raw=None):
    """Devuelve empresa detectada y método de detección."""
    n = norm(Path(path).stem)

    for emp, claves in empresas_claves.items():
        if any(c in n for c in claves):
            return emp, 'nombre_archivo'

    if df_raw is not None:
        try:
            texto = ' '.join(df_raw.head(25).astype(str).fillna('').values.ravel())
            t = norm(texto)

            for emp, claves in empresas_claves.items():
                if any(c in t for c in claves):
                    return emp, 'contenido_archivo'
        except Exception:
            pass

    fallback = Path(path).stem.split()[0].upper()
    return fallback, 'fallback'


def detectar_empresa(path, empresas_claves, df_raw=None):
    empresa, _ = detectar_empresa_detalle(path, empresas_claves, df_raw)
    return empresa


def detectar_tipo_cuenta(cuenta, archivo):
    txt = norm(str(cuenta) + ' ' + Path(archivo).stem)

    if 'caja' in txt or ' ca ' in f' {txt} ' or txt.startswith('ca') or 'ca ' in txt:
        return 'Caja de Ahorro'

    if 'corriente' in txt or 'cta cte' in txt or 'cc ' in txt or txt.startswith('cc'):
        return 'Cuenta Corriente'

    if 'mercado pago' in txt or 'mercadopago' in txt or ' mp ' in f' {txt} ':
        return 'Billetera Virtual'

    return 'Cuenta'
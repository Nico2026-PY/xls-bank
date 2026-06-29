import re
import pandas as pd

from pathlib import Path

from xlsbank.config import COLUMNAS_SALIDA
from xlsbank.utils import limpiar_numero, limpiar_celda_texto, norm

def es_codigo_cuenta_patagonia(texto):
    """Detecta cuentas Patagonia tipo CC$ 385-123012668-000 o CA$ 385-..."""
    if not texto:
        return False

    t = limpiar_celda_texto(texto)
    tn = norm(t)

    # Evita devolver etiquetas del encabezado.
    etiquetas = ['cuenta:', 'cuenta', 'titularidad:', 'titularidad', 'movimientos de cuenta']
    if tn in etiquetas:
        return False

    patrones = [
        r'\b(?:cc|ca|cc\$|ca\$)\s*\$?\s*\d{2,4}[- ]\d{5,12}[- ]\d{1,4}\b',
        r'\b\d{2,4}[- ]\d{5,12}[- ]\d{1,4}\b',
    ]
    return any(re.search(p, tn, flags=re.IGNORECASE) for p in patrones)


def extraer_valor_despues_de_etiqueta(raw, etiqueta, filas_busqueda=25, columnas_busqueda=10):
    """
    Busca una etiqueta en las primeras filas y devuelve el valor cercano.

    Banco Patagonia suele traer:
    Cuenta:      | CC$ 385-123012668-000
    Titularidad: | EMPRESA EJEMPLO

    En algunos Excel el valor queda en la misma celda, en la celda de al lado
    o en una fila cercana. Esta función cubre esos casos sin depender de una
    columna fija.
    """
    etiqueta_norm = norm(etiqueta).replace(':', '')
    max_filas = min(filas_busqueda, len(raw))
    max_cols = min(columnas_busqueda, raw.shape[1])

    for i in range(max_filas):
        for j in range(max_cols):
            celda = limpiar_celda_texto(raw.iat[i, j])
            if not celda:
                continue

            celda_norm = norm(celda).replace(':', '')
            if etiqueta_norm not in celda_norm:
                continue

            # Caso 1: "Cuenta: CC$ 385-..." todo en la misma celda.
            partes = re.split(r':', celda, maxsplit=1)
            if len(partes) == 2 and limpiar_celda_texto(partes[1]):
                valor = limpiar_celda_texto(partes[1])
                if norm(valor) not in [etiqueta_norm, 'cuenta', 'titularidad']:
                    return valor

            # Caso 2: valor en celdas cercanas hacia la derecha.
            for jj in range(j + 1, min(j + 6, raw.shape[1])):
                valor = limpiar_celda_texto(raw.iat[i, jj])
                if valor and norm(valor).replace(':', '') not in [etiqueta_norm, 'cuenta', 'titularidad']:
                    return valor

            # Caso 3: valor en la fila siguiente o subsiguiente.
            for ii in range(i + 1, min(i + 4, len(raw))):
                for jj in range(j, min(j + 6, raw.shape[1])):
                    valor = limpiar_celda_texto(raw.iat[ii, jj])
                    if valor and norm(valor).replace(':', '') not in [etiqueta_norm, 'cuenta', 'titularidad']:
                        return valor

    return ''


def extraer_cuenta_patagonia(raw):
    """Devuelve número de cuenta y titularidad de extractos Banco Patagonia."""
    cuenta = extraer_valor_despues_de_etiqueta(raw, 'Cuenta:')
    titularidad = extraer_valor_despues_de_etiqueta(raw, 'Titularidad:')

    # Si la búsqueda por etiqueta encontró una celda incorrecta, hace un barrido
    # por patrón de cuenta dentro del encabezado.
    if not es_codigo_cuenta_patagonia(cuenta):
        cuenta_detectada = ''
        max_filas = min(25, len(raw))
        max_cols = min(12, raw.shape[1])
        for i in range(max_filas):
            for j in range(max_cols):
                valor = limpiar_celda_texto(raw.iat[i, j])
                if es_codigo_cuenta_patagonia(valor):
                    cuenta_detectada = valor
                    break
            if cuenta_detectada:
                break
        cuenta = cuenta_detectada or cuenta

    # Limpieza final por si quedó texto extra al lado.
    if cuenta:
        m = re.search(
            r'((?:CC|CA)\s*\$?\s*\d{2,4}[- ]\d{5,12}[- ]\d{1,4}|\d{2,4}[- ]\d{5,12}[- ]\d{1,4})',
            cuenta,
            flags=re.IGNORECASE
        )
        if m:
            cuenta = m.group(1).strip()

    return cuenta, titularidad



def extraer_textos_binarios_xls(path):
    """
    Extrae cadenas legibles de archivos .XLS antiguos.

    Algunos extractos de Banco Patagonia guardan el número de cuenta como
    texto dentro del archivo, pero pandas/xlrd puede devolver la plantilla
    o celdas vacías. Este fallback revisa el binario para encontrar valores
    reales como CC$ 385-123012668-000.
    """
    textos = []
    try:
        data = Path(path).read_bytes()
    except Exception:
        return textos

    # Cadenas ANSI/CP1252 típicas de .xls BIFF.
    patron = re.compile(rb'[\x20-\x7E\x80-\xFF]{3,}')
    for match in patron.finditer(data):
        try:
            s = match.group(0).decode('cp1252', errors='ignore')
        except Exception:
            continue
        s = limpiar_celda_texto(s)
        if s and s not in textos:
            textos.append(s)

    # También intenta UTF-16LE por si el .xls guarda strings Unicode.
    try:
        texto_wide = data.decode('utf-16le', errors='ignore')
        for s in re.findall(r'[A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9$.,:/() _\-]{3,}', texto_wide):
            s = limpiar_celda_texto(s)
            if s and s not in textos:
                textos.append(s)
    except Exception:
        pass

    return textos


def extraer_cuenta_patagonia_desde_archivo(path):
    """Fallback para leer cuenta/titularidad Patagonia desde el binario .XLS."""
    textos = extraer_textos_binarios_xls(path)
    cuenta = ''
    titularidad = ''

    patron_cuenta = re.compile(
        r'((?:CC|CA)\s*\$?\s*\d{2,4}[- ]\d{5,12}[- ]\d{1,4}|\d{2,4}[- ]\d{5,12}[- ]\d{1,4})',
        flags=re.IGNORECASE
    )

    indice_cuenta = None
    for idx, texto in enumerate(textos):
        m = patron_cuenta.search(texto)
        if m:
            cuenta = m.group(1).strip().rstrip('#').strip()
            indice_cuenta = idx
            break

    if indice_cuenta is not None:
        # Suele venir justo después del número de cuenta.
        for texto in textos[indice_cuenta + 1: indice_cuenta + 8]:
            tn = norm(texto)
            if not texto or es_codigo_cuenta_patagonia(texto):
                continue
            if any(x in tn for x in ['cuenta', 'titularidad', 'movimientos de cuenta', 'banco patagonia']):
                continue
            if re.search(r'[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]', texto) and len(texto) >= 3:
                titularidad = texto.strip(' #\t')
                break

    return cuenta, titularidad

def procesar_patagonia(path, read_excel_any, encontrar_fila_encabezado):
    raw = read_excel_any(path, header=None)
    fila = encontrar_fila_encabezado(raw, ['Fecha', 'Descripción', 'Débito', 'Crédito', 'Saldo'])
    if fila is None:
        raise ValueError('No se encontró encabezado Patagonia con Fecha/Descripción/Débito/Crédito/Saldo')
    headers = raw.iloc[fila].tolist()
    data = raw.iloc[fila+1:].copy()
    data.columns = [str(h).strip() if str(h) != 'nan' else '' for h in headers]
    data = data.dropna(how='all')
    data = data[data.get('Fecha').notna()] if 'Fecha' in data.columns else data

    cuenta, titularidad = extraer_cuenta_patagonia(raw)

    # Fallback especial para Patagonia .XLS: en algunos archivos xlrd lee
    # placeholders o celdas vacías, aunque el número real está guardado en el binario.
    if not es_codigo_cuenta_patagonia(cuenta):
        cuenta_bin, titular_bin = extraer_cuenta_patagonia_desde_archivo(path)
        if cuenta_bin:
            cuenta = cuenta_bin
        if not titularidad and titular_bin:
            titularidad = titular_bin

    deb = data['Débito'].apply(limpiar_numero) if 'Débito' in data.columns else 0
    cre = data['Crédito'].apply(limpiar_numero) if 'Crédito' in data.columns else 0
    out = pd.DataFrame()
    out['Cuenta'] = cuenta or 'PATAGONIA'
    out['Fecha Valor'] = data.get('Fecha', '')
    out['Fecha Operación'] = data.get('Fecha', '')
    out['Movimiento Fecha-Valor'] = ''
    out['Descripción'] = data.get('Descripción', '')
    out['Detalle'] = data.get('Referencia', '')
    out['Importe'] = cre - deb
    out['Saldo'] = data['Saldo'].apply(limpiar_numero) if 'Saldo' in data.columns else 0
    out['Categoria'] = ''  # Patagonia no trae una categoría real; se deja vacío para no repetir la titularidad.
    out['Referencia'] = data.get('Referencia', '')
    out['Etiquetas'] = ''
    return out[COLUMNAS_SALIDA]

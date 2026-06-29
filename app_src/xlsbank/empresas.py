import os
import sys
import json

from pathlib import Path

from xlsbank.config import CONFIG_EMPRESAS_ARCHIVO


def rutas_posibles_config_empresas():
    """
    Devuelve ubicaciones posibles del archivo privado de empresas.

    No se incluye este archivo en GitHub ni en el ZIP de release.
    Orden recomendado:
    1) Variable de entorno BANCOS_EMPRESAS_CONFIG.
    2) Carpeta del usuario en Windows: %APPDATA%\\Procesador_Bancario\\empresas_config.json.
    3) Carpeta del .exe o del código, útil para pruebas locales.
    """
    rutas = []

    ruta_env = os.environ.get('BANCOS_EMPRESAS_CONFIG', '').strip()
    if ruta_env:
        rutas.append(Path(ruta_env))

    for variable in ['APPDATA', 'LOCALAPPDATA']:
        base = os.environ.get(variable, '').strip()
        if base:
            rutas.append(Path(base) / 'Procesador_Bancario' / CONFIG_EMPRESAS_ARCHIVO)

    try:
        rutas.append(Path(sys.executable).resolve().parent / CONFIG_EMPRESAS_ARCHIVO)
    except Exception:
        pass

    try:
        rutas.append(Path(__file__).resolve().parent / CONFIG_EMPRESAS_ARCHIVO)
    except Exception:
        pass

    rutas.append(Path.cwd() / CONFIG_EMPRESAS_ARCHIVO)

    # Quita repetidos conservando orden.
    vistas = []
    unicas = []

    for ruta in rutas:
        clave = str(ruta).lower()
        if clave not in vistas:
            vistas.append(clave)
            unicas.append(ruta)

    return unicas


def normalizar_config_empresas(data):
    """
    Acepta estos formatos privados:

    Formato recomendado:
    {
      "empresas": {
        "EMPRESA_1": ["clave archivo", "razon social"],
        "EMPRESA_2": {"claves": ["otra clave"], "cuit": "opcional"}
      }
    }

    También acepta lista:
    {
      "empresas": [
        {"nombre": "EMPRESA_1", "claves": ["clave"], "cuit": "opcional"}
      ]
    }

    El CUIT puede estar en el archivo privado para futuras reglas, pero esta
    versión solo usa nombre y claves para detectar empresa.
    """
    if not isinstance(data, dict):
        return {}

    empresas = data.get('empresas', data)
    salida = {}

    if isinstance(empresas, list):
        for item in empresas:
            if not isinstance(item, dict):
                continue

            nombre = str(item.get('nombre', '')).strip().upper()
            claves = item.get('claves', [])

            if isinstance(claves, str):
                claves = [claves]

            claves = [str(c).strip() for c in claves if str(c).strip()]

            if nombre and claves:
                salida[nombre] = claves

        return salida

    if isinstance(empresas, dict):
        for nombre, valor in empresas.items():
            nombre_limpio = str(nombre).strip().upper()
            claves = []

            if isinstance(valor, dict):
                claves = valor.get('claves', [])
            elif isinstance(valor, str):
                claves = [valor]
            elif isinstance(valor, list):
                claves = valor

            claves = [str(c).strip() for c in claves if str(c).strip()]

            if nombre_limpio and claves:
                salida[nombre_limpio] = claves

    return salida


def cargar_empresas_claves():
    """Carga palabras clave privadas de empresas desde empresas_config.json."""
    for ruta in rutas_posibles_config_empresas():
        try:
            if not ruta.exists():
                continue

            data = json.loads(ruta.read_text(encoding='utf-8-sig'))
            empresas = normalizar_config_empresas(data)

            if empresas:
                return empresas

        except Exception:
            # No se detiene la app por una configuración privada dañada.
            continue

    # Sin config privada, la app sigue funcionando por fallback:
    # usa la primera palabra del archivo como empresa y muestra advertencia.
    return {}
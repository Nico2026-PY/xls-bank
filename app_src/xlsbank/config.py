APP_NAME = "XlsBank"
APP_VERSION = "v0.2.16-dev"
APP_AUTHOR = "Nicolás Mellado"
APP_COPYRIGHT = "Copyright © 2026 Nicolás Mellado. Todos los derechos reservados."
APP_SUBTITLE = "Procesador Bancario"


# Las empresas NO se dejan escritas en el código para poder publicar
# el repositorio sin exponer nombres reales, CUITs ni datos internos.
# Se cargan desde un archivo privado externo: empresas_config.json
CONFIG_EMPRESAS_ARCHIVO = "empresas_config.json"

BANCOS_CLAVES = {
    "BPN": ["bpn", "provincia del neuquen", "neuquen"],
    "GALICIA": ["galicia"],
    "PATAGONIA": ["patagonia"],
    "MERCADO_PAGO": ["mercado pago", "mercadopago", "mp "],
    "NACION": ["nacion", "bna"],
    "MACRO": ["macro"],
    "SANTANDER": ["santander"],
    "BBVA": ["bbva", "frances"],
    "ICBC": ["icbc"],
}

COLUMNAS_SALIDA = [
    "Cuenta",
    "Fecha Valor",
    "Fecha Operación",
    "Movimiento Fecha-Valor",
    "Descripción",
    "Detalle",
    "Importe",
    "Saldo",
    "Categoria",
    "Referencia",
    "Etiquetas",
]
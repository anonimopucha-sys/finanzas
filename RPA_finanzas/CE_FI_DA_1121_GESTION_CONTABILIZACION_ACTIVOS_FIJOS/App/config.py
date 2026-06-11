from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
RPA_DIR = BASE_DIR.parent


def _resolver_ruta(ruta_windows: str, ruta_local: Path) -> str:
    """Prioriza rutas reales del sistema y cae a rutas locales del proyecto cuando no están disponibles."""
    ruta = Path(ruta_windows)
    if ruta.exists():
        return str(ruta)
    return str(ruta_local)


URL_BASE = "https://epm-vws20a.corp.epm.com.co/jde/E1Menu.maf"
WEB_DRIVER = None  # Se resuelve dinámicamente con webdriver-manager.
ARCHIVOS_SALIDA = _resolver_ruta(r"C:\CE_FI_DA_1121_Archivos_de_Salida", BASE_DIR / "Data" / "salida")
ARCHIVOS_ENTRADA = _resolver_ruta(r"C:\CE_FI_DA_1121_Archivos_de_Entrada", BASE_DIR / "Data" / "entrada")
EXCEL_FILTROS = _resolver_ruta(
    r"C:\CE_FI_DA_1121_Archivos_de_Entrada\Relacion.Cuentas.Dependencias.xlsx",
    Path(ARCHIVOS_ENTRADA) / "Relacion.Cuentas.Dependencias.xlsx",
)
LOGS_PATH = BASE_DIR / "Data" / "logs"
PUNTO_DE_CONTROL = BASE_DIR / "Data" / "json" / "checkpoint.json"
JSON_FILTROS = BASE_DIR / "Data" / "json" / "reglas.json"
NUMEROS_COMPANIA = ["00533","00534","00598"]
OTES = ["2C","2K","2E"]   
LATENCIA_MAXIMA = 700  # En milisegundos

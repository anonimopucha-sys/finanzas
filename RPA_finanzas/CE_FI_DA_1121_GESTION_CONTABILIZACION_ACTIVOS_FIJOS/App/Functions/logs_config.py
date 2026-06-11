import logging
import os
from datetime import datetime

# Nombre único del logger para el proyecto (evita contaminar logs de Selenium/urllib3)
LOGGER_NAME = "CE1121"

def configurar_logger(nombre: str = LOGGER_NAME) -> logging.Logger:
    """
    Devuelve el logger nombrado del proyecto.
    Si ya tiene handlers configurados, no los duplica.
    Usar 'CE1121' como nombre base garantiza que las librerías externas
    (selenium, urllib3, requests) NO escriban en el archivo de log del RPA.
    """
    carpeta_logs = os.path.abspath(
        os.path.join(os.path.dirname(__file__), os.pardir, "Data", "logs")
    )
    os.makedirs(carpeta_logs, exist_ok=True)

    logger = logging.getLogger(nombre)

    # Evitar duplicar handlers si ya fue configurado anteriormente
    if logger.hasHandlers():
        return logger

    formato = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    fecha_actual = datetime.now().strftime("%Y-%m-%d")
    nombre_archivo_log = os.path.join(carpeta_logs, f"log_{fecha_actual}.log")

    # Handler de archivo
    file_handler = logging.FileHandler(nombre_archivo_log, encoding="utf-8")
    file_handler.setFormatter(formato)

    # Handler de consola para facilitar depuración
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formato)
    console_handler.setLevel(logging.WARNING)  # Solo WARNING+ en consola

    logger.setLevel(logging.INFO)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    # Evitar propagación al logger raíz (impide que selenium escriba aquí)
    logger.propagate = False

    return logger


# Logger de conveniencia para importación directa: from .logs_config import logger
logger = configurar_logger()
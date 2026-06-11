import os
import requests
import time
from .logs_config import configurar_logger

logger = configurar_logger()

# URL de respaldo para verificar conectividad general
_URL_RESPALDO = "https://www.google.com"


def verificar_archivo_y_conexion(
    archivo_entrada: str,
    latencia_max: int,
    url_prueba: str | None = None,
) -> bool:
    """
    Verifica que:
      1. El archivo de entrada exista y sea legible.
      2. Haya conectividad de red con latencia aceptable.

    CAMBIO: si se recibe url_prueba (la URL de JDE), se prueba primero contra
    ese servidor. Si falla (red corporativa inaccesible), se reporta el error
    real en lugar de dar un falso positivo probando contra Google.
    Si url_prueba es None, se usa Google como fallback (comportamiento original).
    """
    # ── 1. Verificar archivo ──────────────────────────────────────────────────
    if not os.path.exists(archivo_entrada):
        logger.error(
            "El archivo de entrada '%s' no existe o no es accesible. "
            "Verifique la ruta configurada en config.py (EXCEL_FILTROS).",
            archivo_entrada,
        )
        return False

    if not os.path.isfile(archivo_entrada):
        logger.error(
            "'%s' existe pero no es un archivo (puede ser un directorio).",
            archivo_entrada,
        )
        return False

    if os.path.getsize(archivo_entrada) == 0:
        logger.error(
            "El archivo de entrada '%s' está vacío (0 bytes).",
            archivo_entrada,
        )
        return False

    logger.info("Archivo de entrada verificado: '%s'", archivo_entrada)

    # ── 2. Verificar conectividad ─────────────────────────────────────────────
    # Priorizar la URL del servidor JDE si se proporcionó
    urls_a_probar: list[tuple[str, str]] = []
    if url_prueba:
        urls_a_probar.append((url_prueba, "servidor JDE"))
    urls_a_probar.append((_URL_RESPALDO, "Google (verificación general)"))

    for url, descripcion in urls_a_probar:
        try:
            inicio = time.time()
            respuesta = requests.get(url, timeout=10)
            latencia = round((time.time() - inicio) * 1000, 2)

            if respuesta.status_code < 500:
                logger.info(
                    "Conectividad OK con %s. Latencia: %.0f ms.",
                    descripcion, latencia,
                )
                if latencia > latencia_max:
                    logger.error(
                        "Latencia con %s demasiado alta: %.0f ms (límite: %d ms). "
                        "La automatización puede fallar por timeouts. "
                        "Verifique la red o aumente LATENCIA_MAXIMA en config.py.",
                        descripcion, latencia, latencia_max,
                    )
                    return False
                return True
            else:
                logger.warning(
                    "El servidor %s respondió con código %d. Se intenta con el siguiente.",
                    descripcion, respuesta.status_code,
                )

        except requests.exceptions.ConnectionError:
            logger.warning(
                "No se pudo conectar con %s (ConnectionError). "
                "Verifique que esté en la red corporativa si es el servidor JDE.",
                descripcion,
            )
        except requests.exceptions.Timeout:
            logger.warning(
                "Timeout al conectar con %s en 10 segundos.",
                descripcion,
            )
        except requests.exceptions.RequestException as e:
            logger.warning("Error de red al probar %s: %s", descripcion, e)

    logger.error(
        "No se pudo verificar conectividad con ninguna URL. "
        "Compruebe la conexión a internet y/o la VPN corporativa."
    )
    return False
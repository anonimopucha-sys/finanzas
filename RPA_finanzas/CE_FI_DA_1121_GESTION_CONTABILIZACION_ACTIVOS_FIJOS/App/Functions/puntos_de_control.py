import json
import os
from pathlib import Path
from .logs_config import configurar_logger

logger = configurar_logger()


def guardar_estado(estado: dict, archivo: str, variables: dict | None = None) -> bool:
    """
    Guarda el estado actual en el archivo de checkpoint.
    Retorna True si guardó correctamente, False si hubo error.
    NUNCA silencia errores: los registra siempre en el log.
    """
    try:
        if estado is None:
            estado = {}
        if variables:
            estado.update(variables)

        # Crear directorio si no existe
        archivo_path = Path(archivo)
        directorio = archivo_path.parent
        directorio.mkdir(parents=True, exist_ok=True)

        # Escritura atómica: escribir en temporal y luego renombrar
        # Evita dejar el checkpoint corrupto si el proceso muere a mitad de escritura
        ruta_temporal = archivo_path.with_suffix(archivo_path.suffix + ".tmp")
        with open(ruta_temporal, "w", encoding="utf-8") as f:
            json.dump(estado, f, indent=4, ensure_ascii=False)
        os.replace(ruta_temporal, archivo_path)

        logger.debug("Checkpoint guardado correctamente en %s", archivo)
        return True

    except OSError as e:
        logger.error(
            "OSError al guardar el checkpoint en '%s': %s. "
            "Verifique espacio en disco y permisos de escritura.",
            archivo, e
        )
        return False
    except Exception as e:
        logger.error("Error inesperado al guardar el checkpoint en '%s': %s", archivo, e)
        return False


def cargar_estado(archivo: str) -> dict | None:
    """
    Carga el estado desde el archivo de checkpoint.
    Retorna el diccionario de estado, {} si el archivo no existe o está vacío,
    o None si el archivo existe pero está corrupto/ilegible.
    """
    try:
        archivo_path = Path(archivo)
        if not archivo_path.exists():
            logger.info("No existe checkpoint en '%s'. Se inicia desde cero.", archivo)
            archivo_path.parent.mkdir(parents=True, exist_ok=True)
            with open(archivo_path, "w", encoding="utf-8") as f:
                json.dump({}, f, indent=4)
            return {}

        with open(archivo_path, "r", encoding="utf-8") as f:
            contenido = f.read().strip()

        if not contenido:
            logger.info("Checkpoint vacío en '%s'. Se inicia desde cero.", archivo)
            return {}

        estado = json.loads(contenido)
        logger.info("Checkpoint cargado desde '%s'. Paso actual: %s", archivo, estado.get("paso_actual", "inicio"))
        return estado

    except json.JSONDecodeError as e:
        logger.error(
            "El archivo de checkpoint '%s' está corrupto y no se puede leer: %s. "
            "Elimínelo manualmente o use 'Reiniciar Proceso' para comenzar desde cero.",
            archivo, e
        )
        return None
    except OSError as e:
        logger.error("OSError al leer el checkpoint '%s': %s", archivo, e)
        return None
    except Exception as e:
        logger.error("Error inesperado al cargar el checkpoint '%s': %s", archivo, e)
        return None


def vaciar_json(ruta: str) -> bool:
    """Vacía completamente el checkpoint. Retorna True si tuvo éxito."""
    try:
        archivo_path = Path(ruta)
        archivo_path.parent.mkdir(parents=True, exist_ok=True)
        with open(archivo_path, "w", encoding="utf-8") as f:
            json.dump({}, f, indent=4)
        logger.info("Checkpoint '%s' reiniciado correctamente.", archivo_path)
        logger.info("Checkpoint vaciado: %s", archivo_path)
        return True
    except Exception as e:
        print(f"Error al vaciar el checkpoint: {e}")
        logger.error("Error al vaciar el checkpoint '%s': %s", ruta, e)
        return False


def vaciar_json_manteniendo_excel(ruta: str) -> bool:
    """
    Vacía el checkpoint pero conserva 'excel_referencia' si existe.
    Retorna True si tuvo éxito.
    """
    try:
        archivo_path = Path(ruta)
        archivo_path.parent.mkdir(parents=True, exist_ok=True)

        datos = {}
        if archivo_path.exists():
            with open(archivo_path, "r", encoding="utf-8") as f:
                contenido = f.read().strip()
            if contenido:
                try:
                    datos = json.loads(contenido)
                except json.JSONDecodeError:
                    logger.warning("Checkpoint corrupto al intentar conservar referencia. Se vaciará completamente.")
                    datos = {}

        nuevo_contenido: dict = {}
        if "excel_referencia" in datos:
            nuevo_contenido["excel_referencia"] = datos["excel_referencia"]
            print(f"Se conservó 'excel_referencia': {datos['excel_referencia']}")

        with open(archivo_path, "w", encoding="utf-8") as f:
            json.dump(nuevo_contenido, f, indent=4)

        logger.info("Checkpoint '%s' reiniciado conservando referencia.", archivo_path)
        logger.info("Checkpoint vaciado conservando excel_referencia: %s", archivo_path)
        return True

    except Exception as e:
        print(f"Error al vaciar el checkpoint: {e}")
        logger.error("Error al vaciar el checkpoint '%s' manteniendo referencia: %s", ruta, e)
        return False
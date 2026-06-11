"""Paquete de funciones compartidas para la automatización."""

from . import (
    descarga_informes,
    envio_correos,
    filtros_simples,
    logs_config,
    puntos_de_control,
    tabla_dinamica,
    unificacion,
    usar_referencia,
    verificacion,
)

from .descarga_informes import (
    iniciar_driver,
    cerrar_driver,
    login,
    navegacion,
    verificar_2A,
    cerrar_sesion,
    verficacion_carpetas,
    sesion_activa,
)
from .unificacion import exportar_excel
from .filtros_simples import aplicar_todos_los_filtros
from .verificacion import verificar_archivo_y_conexion
from .filtros_complejos import lanzar_instancias
from .logs_config import configurar_logger
from .puntos_de_control import (
    cargar_estado,
    guardar_estado,
    vaciar_json,
    vaciar_json_manteniendo_excel,
)
from .usar_referencia import usar_excel_referencia, eliminar_informes
from .tabla_dinamica import crear_tabla_dinamica

__all__ = [
    "descarga_informes",
    "envio_correos",
    "filtros_complejos",
    "filtros_simples",
    "logs_config",
    "puntos_de_control",
    "tabla_dinamica",
    "unificacion",
    "usar_referencia",
    "verificacion",
    "iniciar_driver",
    "cerrar_driver",
    "login",
    "navegacion",
    "verificar_2A",
    "cerrar_sesion",
    "verficacion_carpetas",
    "sesion_activa",
    "exportar_excel",
    "aplicar_todos_los_filtros",
    "verificar_archivo_y_conexion",
    "lanzar_instancias",
    "configurar_logger",
    "cargar_estado",
    "guardar_estado",
    "vaciar_json",
    "vaciar_json_manteniendo_excel",
    "usar_excel_referencia",
    "eliminar_informes",
    "crear_tabla_dinamica",
]

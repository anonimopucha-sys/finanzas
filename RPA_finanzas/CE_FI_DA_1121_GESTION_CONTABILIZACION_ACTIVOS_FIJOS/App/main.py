from mainencriptador import obtener_credenciales
from Functions.descarga_informes import (
    iniciar_driver, cerrar_driver, login, navegacion,
    verificar_2A, cerrar_sesion, verficacion_carpetas, sesion_activa,
)
from Functions.unificacion import exportar_excel
from Functions.filtros_simples import aplicar_todos_los_filtros
from Functions.verificacion import verificar_archivo_y_conexion
from Functions.filtros_complejos import lanzar_instancias
from Functions.logs_config import configurar_logger
from Functions.puntos_de_control import (
    cargar_estado, guardar_estado, vaciar_json, vaciar_json_manteniendo_excel,
)
from Functions.usar_referencia import usar_excel_referencia, eliminar_informes
from Functions.tabla_dinamica import crear_tabla_dinamica
from config import (
    URL_BASE, NUMEROS_COMPANIA, ARCHIVOS_SALIDA,
    JSON_FILTROS, EXCEL_FILTROS, PUNTO_DE_CONTROL, LATENCIA_MAXIMA,
)

logger = configurar_logger()


def reiniciar_ejecucion():
    """Vacía completamente el archivo JSON de puntos de control."""
    print(PUNTO_DE_CONTROL)
    vaciar_json(PUNTO_DE_CONTROL)


def reiniciar_ejecucion_con_referencia():
    """Vacía el JSON pero mantiene la referencia del Excel anterior."""
    vaciar_json_manteniendo_excel(PUNTO_DE_CONTROL)


def ejecutar_proceso():
    try:
        # Las credenciales se obtienen al inicio, antes de cualquier paso,
        # para que estén disponibles incluso al reanudar desde un checkpoint.
        try:
            USER, PASS = obtener_credenciales()
        except Exception as e:
            logger.error("Error cargando credenciales: %s", e)
            print(f"Error cargando credenciales: {e}")
            return

        # Carga el estado actual desde el archivo de puntos de control
        estado = cargar_estado(PUNTO_DE_CONTROL) or {}
        paso = estado.get("paso_actual")

        # ── PASO 1: INICIO – Verificación de requisitos y gestión de 2A ────────
        if paso is None or paso == "inicio":
            try:
                # Verifica que existan los archivos necesarios y haya conexión
                if not verificar_archivo_y_conexion(
                    EXCEL_FILTROS, LATENCIA_MAXIMA, url_prueba=URL_BASE
                ):
                    logger.error("No se cumplen los requisitos mínimos para iniciar.")
                    return

                print("Se inicia el proceso de ignorar 2A")
                # iniciar_driver ya no recibe driver_path (se resuelve con webdriver-manager)
                driver, ruta_archivos, nombre_documento = iniciar_driver(
                    ARCHIVOS_SALIDA, URL_BASE
                )

                if driver is None:
                    logger.error("No se pudo iniciar el navegador. Proceso detenido.")
                    print("ERROR: No se pudo iniciar el navegador.")
                    return

                # AUTO-LOGIN INTELIGENTE
                if sesion_activa(driver):
                    logger.info("Sesión activa detectada. Se omite login.")
                    print("Sesión activa detectada. No se realizó login.")
                else:
                    logger.info("No hay sesión activa. Realizando login automático.")
                    print("No hay sesión activa. Realizando login.")
                    if not login(driver, "incorrectos", USER, PASS):
                        logger.error("Login fallido en el paso de verificación 2A.")
                        cerrar_driver(driver)
                        return

                # Inicia proceso de omitir 2A
                verificar_2A(driver)
                # Cierra la sesión y el driver
                cerrar_sesion(driver)
                cerrar_driver(driver)

                # Avanza al siguiente paso
                paso = "descargar_informes"
                variables = {
                    "paso_actual": "descargar_informes",
                    "ruta_archivos": ruta_archivos,
                    "nombre_documento": nombre_documento,
                }
                guardar_estado(estado, PUNTO_DE_CONTROL, variables)
                logger.info("Paso 1 completado: verificación de 2A.")
                print("Paso 1 completado: verificación de 2A.")

            except Exception as e:
                logger.error("Error al ejecutar el paso 'inicio': %s", e)
                print(f"Error al ejecutar el paso de verificación de 2A: {e}")
                return

        # ── PASO 2: DESCARGA DE INFORMES ────────────────────────────────────────
        if paso == "descargar_informes":
            print("Se inicia paso de descarga de informes")
            try:
                ruta_archivos = estado.get("ruta_archivos") if estado else None

                # navegacion ya no recibe web_driver (eliminado en esta versión)
                navegacion(NUMEROS_COMPANIA, ruta_archivos, URL_BASE, USER, PASS)

                # Verifica que se hayan creado todas las carpetas y archivos
                if verficacion_carpetas(ruta_archivos, NUMEROS_COMPANIA):
                    paso = "unificar_excel"
                    variables = {"paso_actual": "unificar_excel"}
                    guardar_estado(estado, PUNTO_DE_CONTROL, variables)
                else:
                    logger.error(
                        "Error: no se encontraron los informes; archivo final incompleto."
                    )
                    print("Error: no se encontraron los informes, archivo final incompleto.")
                    return

                print("Proceso de descarga de informes completado.")

            except Exception as e:
                logger.error("Error al ejecutar el paso 'descargar_informes': %s", e)
                print(f"Error al ejecutar el paso de descarga de informes: {e}")
                return

        # ── PASO 3: UNIFICACIÓN DE EXCEL Y FILTROS SIMPLES ──────────────────────
        if paso == "unificar_excel":
            print("Se inicia el proceso de unificar informes")
            try:
                ruta_archivos = estado.get("ruta_archivos") if estado else None
                nombre_documento = estado.get("nombre_documento") if estado else None

                # Unifica todos los informes descargados en un solo Excel
                excel = exportar_excel(ruta_archivos, nombre_documento, NUMEROS_COMPANIA)

                if not excel:
                    logger.error(
                        "No se pudo generar el Excel unificado. Proceso detenido."
                    )
                    print("Error: No se pudo generar el Excel unificado.")
                    return

                # Aplica filtros simples al Excel unificado
                # aplicar_todos_los_filtros ahora retorna la ruta o None
                excel_final = aplicar_todos_los_filtros(EXCEL_FILTROS, JSON_FILTROS, excel)

                if not excel_final:
                    logger.error(
                        "Los filtros simples fallaron. Proceso detenido."
                    )
                    print("Error: No se pudieron aplicar los filtros simples.")
                    return

                paso = "filtros_complejos"
                variables = {
                    "paso_actual": "filtros_complejos",
                    "excel_final": excel_final,
                }
                guardar_estado(estado, PUNTO_DE_CONTROL, variables)
                print("Se completó correctamente el paso de unificación y filtros simples.")

                # Si existe un Excel de referencia previo, lo usa para comparaciones
                if estado and estado.get("excel_referencia"):
                    excel_referencia = estado.get("excel_referencia")
                    usar_excel_referencia(excel_referencia, excel_final)

                # Elimina los informes individuales para liberar espacio
                eliminar_informes(ruta_archivos, NUMEROS_COMPANIA)

            except Exception as e:
                logger.error("Error al ejecutar el paso 'unificar_excel': %s", e)
                print(f"Error al ejecutar el paso de unificación de informes: {e}")
                return

        # ── PASO 4: FILTROS COMPLEJOS Y FINALIZACIÓN ────────────────────────────
        if paso == "filtros_complejos":
            print("Se inicia el proceso de filtros complejos")
            try:
                excel_final = estado.get("excel_final") if estado else None
                if not excel_final:
                    logger.error(
                        "No se encontró el Excel procesado en el estado."
                    )
                    print("Error: No se encontró el Excel procesado.")
                    return

                # lanzar_instancias ya no recibe web_driver en esta versión
                lanzar_instancias(excel_final, EXCEL_FILTROS, URL_BASE, PASS, USER)

                # Crea tabla dinámica con los datos finales
                crear_tabla_dinamica(excel_final)

                # Guarda el Excel final como referencia para futuras ejecuciones
                variables = {"excel_referencia": excel_final}
                guardar_estado(estado, PUNTO_DE_CONTROL, variables)

            except Exception as e:
                logger.error("Error al ejecutar el paso 'filtros_complejos': %s", e)
                print(f"Error al aplicar filtros complejos: {e}")
                return

            print(
                "Se finalizó el proceso de filtros complejos. "
                "Recuerde reiniciar los puntos de control si el resultado fue satisfactorio."
            )

    except Exception as e:
        logger.error("Error en el proceso general: %s", e)


if __name__ == "__main__":
    ejecutar_proceso()

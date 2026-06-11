from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    WebDriverException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
    MoveTargetOutOfBoundsException,
)
from selenium.webdriver.common.keys import Keys
import time
import contextlib
import threading
import os
from typing import Any
from selenium.webdriver.common.by import By
from datetime import datetime, timedelta
from selenium.webdriver.common.action_chains import ActionChains
from .driver_manager import (
    crear_driver,
    cerrar_driver as cerrar_driver_manager,
    DriverFatalError,
    DriverRetryableError,
)
from .logs_config import configurar_logger
logger = configurar_logger()
### Funciones del driver
@contextlib.contextmanager
def suppress_stderr():
    with open(os.devnull, 'w') as devnull:
        old_stderr = os.dup(2)
        os.dup2(devnull.fileno(), 2)
        try:
            yield
        finally:
            os.dup2(old_stderr, 2)
            
def sesion_activa(driver: Any, timeout: int = 5) -> bool:
    """
    Retorna True si JDE ya está logueado (iframe principal existe).
    """
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.ID, "e1menuAppIframe"))
        )
        return True
    except TimeoutException:
        return False


def sesion_valida(driver: Any) -> bool:
    """Retorna False si la sesión Selenium ya no es válida."""
    try:
        if driver is None:
            return False
        return bool(getattr(driver, "session_id", None)) and driver.current_url is not None
    except Exception:
        return False


def iniciar_driver(
    base_dir: str | None,
    url: str,
    numero_compañia: str | None = None,
    headless: bool = True,
) -> tuple[webdriver.Chrome, str | None, str | None]:
    # Configurar directorio de descargas
    if base_dir is not None:
        ruta_libros, archivo_final = crear_carpeta_libros(base_dir, numero_compañia)
    else:
        ruta_libros = os.path.expanduser("~")
        archivo_final = None

    driver = crear_driver(download_dir=ruta_libros, headless=headless)
    logger.info("Driver Selenium creado correctamente con webdriver-manager")
    logger.debug("URL recibida: %s", url)
    driver.get(url)
    return driver, ruta_libros, archivo_final


def cerrar_driver(driver: Any) -> None:
    try:
        cerrar_driver_manager(driver)
    except WebDriverException as e:
        logger.error("Error al cerrar el driver: %s", e)

def ingresar_texto_jde_real(driver: Any, by: Any, identificador: Any, texto: str, timeout: int = 10) -> bool:
    """
    Wrapper de compatibilidad para código antiguo.
    """
    try:
        wait = WebDriverWait(driver, timeout)
        campo = wait.until(EC.visibility_of_element_located((by, identificador)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", campo)
        time.sleep(0.3)

        try:
            campo.click()
        except (ElementClickInterceptedException, MoveTargetOutOfBoundsException, StaleElementReferenceException):
            driver.execute_script("arguments[0].click();", campo)
            time.sleep(0.3)

        try:
            campo.clear()
        except WebDriverException:
            driver.execute_script("arguments[0].value = '';", campo)

        time.sleep(0.2)
        for caracter in texto:
            campo.send_keys(caracter)
            time.sleep(0.05)

        campo.send_keys(Keys.TAB)
        time.sleep(0.3)
        return True

    except TimeoutException:
        logger.error(f"Timeout: no se encontró el campo con identificador '{identificador}' en {timeout} segundos")
        return False
    except WebDriverException as e:
        logger.error(f"Error escribiendo texto en JDE ({identificador}): {e}")
        return False
    except Exception as e:
        logger.error(f"Error inesperado escribiendo texto en JDE ({identificador}): {e}")
        return False

def ingresar_texto_jde(driver: Any, by: Any, identificador: Any, texto: str, timeout: int = 10) -> bool:
    return ingresar_texto_jde_real(driver, by, identificador, texto, timeout)

def ingresar_texto_jde_con_copiar_pegar(driver: Any, by: Any, identificador: Any, texto: str, timeout: int = 10) -> None:
    try:
        wait = WebDriverWait(driver, timeout)
        wait.until(EC.visibility_of_element_located((by, identificador)))
        campo = driver.find_element(by, identificador)
        # Hacer click y esperar estabilización
        campo.click()
        time.sleep(0.5)
        # Seleccionar todo el contenido y reemplazar
        campo.send_keys(Keys.CONTROL + "a")
        campo.send_keys(texto)
    except TimeoutException:
        logger.error(f"Timeout: no se encontró el campo con identificador '{identificador}' en {timeout} segundos")
    except WebDriverException as e:
        logger.error(f"Error al ingresar texto en el campo '{identificador}': {e}")

def ingresar_texto_jde_ignorar_2A(driver: Any, by: Any, identificador: Any, texto: str, timeout: int = 10) -> None:
    try:
        wait = WebDriverWait(driver, timeout)
        campo = wait.until(EC.visibility_of_element_located((by, identificador)))
    
        campo.clear()
        # Usar ActionChains para mayor control de la interacción
        acciones = ActionChains(driver)
        
        acciones.move_to_element(campo).click().send_keys(Keys.CONTROL + 'P').send_keys(Keys.DELETE)
        acciones.send_keys(texto).perform()
    except TimeoutException:
        logger.error(f"Timeout: no se encontró el campo con identificador '{identificador}' en {timeout} segundos")
    except WebDriverException as e:
        logger.error(f"Error al ingresar texto en el campo '{identificador}': {e}")

def hacer_click(driver: Any, by: Any, identificador: Any, timeout: int = 10) -> bool:
    if not sesion_valida(driver):
        logger.error("Sesión de Selenium inválida o navegador cerrado antes de hacer click en '%s'.", identificador)
        raise WebDriverException("Sesión inválida antes de hacer click.")

    try:
        wait = WebDriverWait(driver, timeout)
        boton = wait.until(EC.element_to_be_clickable((by, identificador)))
        boton.click()
        return True
    except TimeoutException as exc:
        logger.error(f"Timeout: No se encontró o no fue clickeable el elemento '{identificador}' en {timeout} segundos.")
        raise WebDriverException(f"Timeout haciendo click en '{identificador}'", exc)
    except WebDriverException as e:
        logger.error(f"Error al hacer click en el elemento '{identificador}': {e}")
        raise


def hacer_click_elementos(driver: Any, by: Any, identificador: Any, timeout: int = 10) -> bool:
    if not sesion_valida(driver):
        logger.error("Sesión de Selenium inválida o navegador cerrado antes de hacer click en '%s'.", identificador)
        raise WebDriverException("Sesión inválida antes de hacer click en elemento.")

    wait = WebDriverWait(driver, timeout)
    try:
        # Esperar a que el elemento esté visible y clickeable
        elemento = wait.until(EC.visibility_of_element_located((by, identificador)))
        wait.until(EC.element_to_be_clickable((by, identificador)))

        # Asegurarse de que esté en el viewport
        driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", elemento)
        time.sleep(0.5)

        try:
            elemento.click()
            logger.debug(f"Se hizo clic en el elemento '{identificador}'.")
            return True
        except (ElementClickInterceptedException, MoveTargetOutOfBoundsException, StaleElementReferenceException) as exc:
            logger.warning(
                "Click directo falló en '%s' con %s; intentando click por JavaScript.",
                identificador,
                type(exc).__name__,
            )
            driver.execute_script("arguments[0].click();", elemento)
            logger.debug(f"Se hizo clic por JavaScript en el elemento '{identificador}'.")
            return True

    except TimeoutException as exc:
        logger.error(
            f"Timeout: El elemento con identificador '{identificador}' no es visible o clickeable en {timeout} segundos."
        )
        raise WebDriverException(f"Timeout en hacer_click_elementos '{identificador}'", exc)
    except WebDriverException:
        raise


def verificar_error_login(driver: Any, by: Any, identificador_error: Any, texto_esperado: str, timeout: int = 5) -> bool:
    try:
        wait = WebDriverWait(driver, timeout)
        elemento_error = wait.until(EC.visibility_of_element_located((by, identificador_error)))
        texto_error = elemento_error.text.lower().strip()
        claves = [parte.strip().lower() for parte in str(texto_esperado).split('|') if parte.strip()]
        return any(clave in texto_error for clave in claves) if claves else False
    except TimeoutException:
        return False
    
def login(driver: Any, texto_error: str, usuariow: str, passw: str, max_intentos: int = 2) -> bool:
    """Intenta iniciar sesión y retorna True si fue exitoso."""
    intentos = 0
    while intentos < max_intentos:
        if not sesion_valida(driver):
            logger.error("Sesión de Selenium inválida antes de iniciar login.")
            return False

        usuario = usuariow
        contraseña = passw

        logger.debug("Usuario enviado a JDE: [%s]", usuario)
        logger.debug("Longitud usuario: %s", len(usuario))
        logger.debug("Contraseña longitud: %s", len(contraseña))

        # Intenta realizar el login
        if not ingresar_texto_jde(driver, By.ID, "User", usuario):
            logger.error("No se pudo escribir el usuario en el login.")
            return False

        if not ingresar_texto_jde(driver, By.ID, "Password", contraseña):
            logger.error("No se pudo escribir la contraseña en el login.")
            return False

        try:
            hacer_click_elementos(driver, By.XPATH, '//*[@id="mainLoginTable"]/tbody/tr[4]/td/input', timeout=30)
        except WebDriverException as e:
            logger.warning("No se pudo hacer click en el botón de login con click directo: %s", e)
            try:
                password_field = driver.find_element(By.ID, "Password")
                password_field.send_keys(Keys.ENTER)
                logger.debug("Intento de login enviado con ENTER en el campo Password.")
            except Exception as inner:
                logger.warning("No se pudo enviar login con ENTER: %s", inner)
                try:
                    boton_login = driver.find_element(By.XPATH, '//*[@id="mainLoginTable"]/tbody/tr[4]/td/input')
                    driver.execute_script("arguments[0].click();", boton_login)
                    logger.debug("Intento de login enviado con click por JavaScript.")
                except Exception as inner2:
                    logger.error("No se pudo enviar el login por ninguna vía: %s", inner2)
                    intentos += 1
                    time.sleep(2)
                    continue

        if verificar_error_login(driver, By.ID, "SignInError", texto_error):
            logger.warning("Login fallido, intento %s de %s.", intentos + 1, max_intentos)
            intentos += 1
            continue

        iframe_xpath = '//*[@id="e1menuAppIframe"]'
        try:
            driver.switch_to.default_content()
            if WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, iframe_xpath))):
                logger.debug("Login exitoso.")
                time.sleep(2)
                return True
        except TimeoutException:
            logger.warning(
                "No se detectó el iframe de JDE tras el login en 60 segundos. "
                "URL actual: %s",
                driver.current_url,
            )
            intentos += 1
            continue
        except WebDriverException as e:
            logger.error("Error durante la verificación del login: %s", e)
            return False

    logger.error("Máximo de intentos alcanzado. No se pudo iniciar sesión.")
    return False

#Navegacion
def verificar_ventanas(driver: Any, timeout: int = 10) -> bool:
    ventanas = driver.window_handles
    logger.debug(f"Se ha abierto {len(ventanas)} ventanas")
    if len(ventanas) > 1:
        for idx, ventana in enumerate(ventanas):
            logger.debug(f"ventana {idx + 1}: {ventana}")
        return True
    else: 
        logger.debug("Solo hay una ventana abierta.")
        return False

def cambiar_a_iframe(driver: Any, iframes_xpaths: Any) -> None:
    try:
        # Recorre la lista de XPaths y cambia al contexto de cada iframe
        for idx, xpath in enumerate(iframes_xpaths):
            logger.debug(f"Cambiando al iframe {idx + 1} con XPath: {xpath}")
            driver.switch_to.frame(driver.find_element(By.XPATH, xpath))  # Cambiar al iframe
        logger.debug(f"Cambiado al último iframe de la ruta.")
    except Exception as e:
        logger.error(f"Error al cambiar a uno de los iframes en la ruta: {e}")

def cambiar_ventana(driver: Any, ventana: Any = None) -> None:
    ventanas = driver.window_handles
    if ventana is None:
        # Cambiar a la ventana más reciente (útil para popups nuevos)
        ventanas = driver.window_handles
        driver.switch_to.window(ventanas[-1])
        logger.debug(f"Cambiado a la nueva ventana con handle: {ventanas[-1]}")
    else:
        if ventana > 0 and ventana <= len(ventanas):
            driver.switch_to.window(ventanas[ventana - 1])
            logger.debug(f"Cambiado a la ventana número {ventana} con handle: {ventanas[ventana - 1]}")
        else:
            # Cambiar a ventana específica (numeración 1-indexada)
            logger.error(f"Índice de ventana inválido: {ventana}. Solo hay {len(ventanas)} ventana(s) abiertas.")

def regresar_al_contexto_principal(driver: Any) -> None:
    driver.switch_to.default_content()  # Regresar al contexto principal
    logger.debug("Regresado al contexto principal fuera del iframe.")

def obtener_fecha_libro_mayor():

    hoy = datetime.today()
    if hoy.day < 5:
        # Primer día del mes actual -> mes pasado
        primer_dia_mes_pasado = hoy.replace(day=1) - timedelta(days=1)
        # Primer día del mes antepasado
        primer_dia_mes_antepasado = primer_dia_mes_pasado.replace(day=1)
        # Último día del mes antepasado
        fecha_objetivo = primer_dia_mes_antepasado - timedelta(days=1)
    else:
        # Primer día del mes actual -> mes pasado
        primer_dia_mes_pasado = hoy.replace(day=1)
        # Último día del mes pasado
        fecha_objetivo = primer_dia_mes_pasado - timedelta(days=1)          
    return f">{fecha_objetivo.year}/{fecha_objetivo.month:02d}/{fecha_objetivo.day:02d}"

def abrir_menu(driver,xpath_elemento,xpath_opcion, timeout=5):
    try:
        # Localiza el elemento sobre el que se hará clic derecho usando su XPath
        elemento = driver.find_element(By.XPATH, xpath_elemento)
        # Crea una cadena de acciones con Selenium para simular el clic derecho
        acciones = ActionChains(driver)
        acciones.context_click(elemento).perform()
        # Espera breve para que el menú contextual aparezca completamente
        time.sleep(2)
        # Llama a una función que hace clic en la opción del menú
        hacer_click_elementos(driver, By.XPATH, xpath_opcion)
       
    except Exception as e:
        logger.error(f"Error al hacer click derecho: {e}")

# CORRECCIÓN 1: timeout y timeout_scroll aumentados para manejar tablas grandes (ej. compañía 00533)
def verificar_tabla(driver, timeout=30, timeout_scroll=5):
    try:
        # XPath del cuerpo de la tabla y del contenedor de scroll virtual
        tabla_xpath = '//*[@id="jdeGridData0_1"]/tbody'
        scroll_xpath ='//*[@id="jdeGridVirtualAbove0_1"]'
        # Espera hasta que el cuerpo de la tabla esté presente en el DOM
        wait = WebDriverWait(driver, timeout)
        wait.until(EC.presence_of_element_located((By.XPATH, tabla_xpath)))
        
        tr_count= 0 # Contador inicial de filas
        start_time = time.time()
        # Repetir hasta que se supere el timeout
        while time.time() - start_time < timeout:
            # Obtener todas las filas actuales de la tabla
            trs = driver.find_elements(By.XPATH, f"{tabla_xpath}/tr")
            nuevo_tr_count = len(trs)
            # Si no se han cargado más filas, se detiene el scroll
            if nuevo_tr_count == tr_count:
                logger.debug("No se han cargado más filas.")
                break
            # Mostrar cuántas filas se han cargado hasta ahora
            logger.debug(f"Filas cargadas: {nuevo_tr_count}")
            tr_count = nuevo_tr_count
            # Hacer scroll hacia abajo en el contenedor virtual para cargar más filas
            scroll_container = driver.find_element(By.XPATH, scroll_xpath)
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", scroll_container)
            # Esperar un tiempo antes de verificar nuevamente
            time.sleep(timeout_scroll)
        # Si no se encontró ninguna fila, se considera que la tabla está vacía o falló la carga
        if tr_count == 0:
            logger.debug("No se encontraron filas en la tabla.")
            return False
        # Revisar si la última fila contiene la palabra 'total'
        ultimo_tr = driver.find_element(By.XPATH, f"{tabla_xpath}/tr[{tr_count}]")
        if "total" in ultimo_tr.get_attribute("innerText").lower():
            logger.debug("Se encontró la palabra 'total' en el último tr.")
            return True
        else:
            logger.error("No se encontró la palabra 'total' en el último tr.")
            return False
    except Exception as e:
        logger.error(f"Error al procesar la tabla: {e}")
        return False

def crear_carpeta_libros(base_dir, numero_compañia):
    ahora = datetime.now()
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
             "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    nombre_mes = meses[ahora.month - 1]
     # Obtener día, hora y am/pm
    dia = ahora.day
    hora = ahora.strftime("%I")  # Hora en formato 12h con cero inicial
    am_pm = ahora.strftime("%p").lower()  # 'am' o 'pm'

    # Nombre de la subcarpeta: 'DD-HH(am|pm)'
    carpeta_fecha_hora = f"{dia}-{hora}{am_pm}"
    nombre_archivo_final = f"Transacciones pendientes por contabilizar en el filtro_{dia}-{nombre_mes}-{ahora.year}_{hora}{am_pm}.xlsx"

    # Ruta completa
    if numero_compañia:
        ruta_completa = os.path.join(base_dir, numero_compañia)
    else:
        ruta_completa = os.path.join(base_dir, nombre_mes, carpeta_fecha_hora)

    # Crear la ruta si no existe
    os.makedirs(ruta_completa, exist_ok=True)

    logger.debug(f"Carpeta creada: {ruta_completa}")
    return ruta_completa, nombre_archivo_final

def verificar_2A(driver, timeout=60):
    try:
        time.sleep(1)
        # Volver al contexto principal del documento (por si está dentro de un iframe)
        driver.switch_to.default_content()
        span_xpath = '//*[@id="pageContainer"]/table/tbody[6]/tr/td/div[2]/table/tbody/tr/td[2]/div/table/tbody/tr[2]/td[2]/div/table/tr[2]/td/span'
        iframes = ['//*[@id="e1menuAppIframe"]']
        cambiar_a_iframe(driver, iframes)
        hacer_click_elementos(driver, By.XPATH, '//*[@id="tab1"]')
        time.sleep(1)
        driver.switch_to.default_content()
        iframes= ['//*[@id="e1menuAppIframe"]','//*[@id="wcFrame1"]','//*[@id="RIPaneIFRAME1"]']
        cambiar_a_iframe(driver, iframes)
        if not click_y_esperar_nueva_ventana(driver, By.XPATH, span_xpath, timeout=max(timeout, 60), reintentos=3):
            logger.error("No se abrió la ventana secundaria de 2A en el tiempo esperado.")
            return False
        verificar_ventanas(driver)
        cambiar_ventana(driver)
    
        iframes = ['//*[@id="e1menuAppIframe"]']
        cambiar_a_iframe(driver, iframes)

        ingresar_texto_jde(driver, By.XPATH, '//*[@id="qbeRow0_1"]/td[5]/div/nobr/input', "2A")
        fecha_libro_mayor = obtener_fecha_libro_mayor()
        ingresar_texto_jde(driver, By.XPATH,'//*[@id="qbeRow0_1"]/td[7]/div/nobr/input', fecha_libro_mayor)
        hacer_click_elementos(driver, By.ID, "C0_100")
        WebDriverWait(driver, timeout=360).until(EC.text_to_be_present_in_element((By.ID,"GridLabel0_1.Records"), "Registros"))
        try:
            WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "GOTOLAST0_1")))
            hacer_click_elementos(driver, By.ID, "GOTOLAST0_1")
            time.sleep(5)
            # Esperar a que los registros estén cargados
            registros_cargados=esperar_registros(driver)
            if not registros_cargados:
                #caso funcional normal: no hay registros 2A
                logger.info("no se encontraron registros 2A. Se continua el proceso.")
                return False
            verificar_tabla(driver)
            hacer_click_elementos(driver,By.ID,"selectAll0_1")
            time.sleep(2)
             # Abrir menú contextual sobre las filas seleccionadas y seleccionar opción
            abrir_menu(driver,'//*[@id="jdeGridData0_1"]/tbody','//*[@id="HE0_97"]/tbody/tr/td[2]/span/nobr')
            time.sleep(2)
            hacer_click_elementos(driver, By.XPATH,'//*[@id="C0_17"]')
                # Reintentar escribir "P" tres veces por si hay validación o errores
            for i in range(3):
                ingresar_texto_jde_ignorar_2A(driver, By.XPATH, '//*[@id="C0_17"]', "P")
            hacer_click_elementos(driver, By.XPATH,'//*[@id="C0_24"]')
            time.sleep(2)
            hacer_click_elementos(driver, By.XPATH,'//*[@id="C0_11"]')
            hacer_click_elementos(driver, By.XPATH,'//*[@id="C0_11"]')
            time.sleep(30)
        except Exception:
            logger.info("Pantalla de 2A sin registros.")
            return False
        
        hacer_click_elementos(driver, By.ID, "C0_89")
        WebDriverWait(driver, 30).until(lambda d: len(d.window_handles) == 1)
        verificar_ventanas(driver)
        cambiar_ventana(driver)
        return True
    except Exception as e:
        logger.error(f"error al ignorar 2A en el documento: {e}")
        return False

def cerrar_sesion(driver: Any) -> None:
    # Cerrar sesion de JDEdwards
    hacer_click_elementos(driver, By.XPATH, '//*[@id="userSessionDropdownArrow"]')
    hacer_click_elementos(driver, By.XPATH, '//*[@id="e1LogoutLink"]')

# CORRECCIÓN 2: timeout aumentado a 120s para manejar la carga de tablas grandes (ej. compañía 00533 con +15MB)
def esperar_registros(driver, timeout=120, intervalo=0.5):
    """
    Espera a que los registros estén completamente cargados.
    Retorna True si cargan, False si no.
    NUNCA retorna None.
    """
    valor_esperado = 'FormDivScrollHandler.syncToolBar(true)'
    def comparacion(driver):
            elemento = driver.find_element(By.XPATH,'//*[@id="e1formDiv"]')
            return elemento.get_attribute('onscroll') == valor_esperado
    try:  
         WebDriverWait(driver, timeout, poll_frequency=intervalo).until(comparacion)
         return True
    except TimeoutException:
        return False


def esperar_nueva_ventana(driver, identificadores_anteriores, timeout=90, intervalo=2) -> bool:
    """Espera a que se abra una nueva ventana tras una acción que la dispara."""
    try:
        WebDriverWait(driver, timeout, poll_frequency=intervalo).until(
            EC.new_window_is_opened(identificadores_anteriores)
        )
        return True
    except TimeoutException:
        return False


def click_y_esperar_nueva_ventana(driver, by, identificador, timeout=90, reintentos=2, intervalo_click=2) -> bool:
    """Hace click en un elemento y espera a que se abra una nueva ventana, con reintentos."""
    for intento in range(1, reintentos + 1):
        try:
            ventanas_antes = list(driver.window_handles)
            hacer_click_elementos(driver, by, identificador)
            if esperar_nueva_ventana(driver, ventanas_antes, timeout=timeout):
                return True
            logger.warning(
                "Intento %s/%s: no se abrió nueva ventana tras hacer click en %s.",
                intento,
                reintentos,
                identificador,
            )
        except WebDriverException as e:
            logger.warning(
                "Intento %s/%s fallido al abrir nueva ventana en %s: %s",
                intento,
                reintentos,
                identificador,
                e,
            )
        if intento < reintentos:
            time.sleep(intervalo_click)
    return False


def descargar_informes(driver, numeros_compania, timeout=60):
    time.sleep(1)
    driver.switch_to.default_content()
    iframes = ['//*[@id="e1menuAppIframe"]']
    cambiar_a_iframe(driver, iframes)
    hacer_click_elementos(driver, By.XPATH, '//*[@id="tab1"]')
    time.sleep(3)
    driver.switch_to.default_content()
    span_xpath = '//*[@id="pageContainer"]/table/tbody[6]/tr/td/div[2]/table/tbody/tr/td[2]/div/table/tbody/tr[2]/td[2]/div/table/tr[2]/td/span'
    iframes= ['//*[@id="e1menuAppIframe"]','//*[@id="wcFrame1"]','//*[@id="RIPaneIFRAME1"]']
    cambiar_a_iframe(driver, iframes)
    if not click_y_esperar_nueva_ventana(driver, By.XPATH, span_xpath, timeout=max(timeout, 60), reintentos=3):
        logger.error("No se abrió la ventana secundaria del informe en el tiempo esperado.")
        return
    verificar_ventanas(driver)
    cambiar_ventana(driver)
    ###
    iframes = ['//*[@id="e1menuAppIframe"]']
    cambiar_a_iframe(driver, iframes)
    time.sleep(2)
    ingresar_texto_jde(driver, By.ID, 'C0_23', numeros_compania)
    fecha_libro_mayor = obtener_fecha_libro_mayor()
    ingresar_texto_jde(driver, By.XPATH,'//*[@id="qbeRow0_1"]/td[7]/div/nobr/input', fecha_libro_mayor)
    # FIX: verificar si reintentos_boton tuvo éxito antes de continuar
    if not reintentos_boton(driver):
        logger.error(f"No se pudo cargar la búsqueda para compañía {numeros_compania}. Se cancela la descarga.")
        return

    # FIX: timeout aumentado a 480s para compañías grandes como 00533
    WebDriverWait(driver, timeout=480).until(EC.text_to_be_present_in_element((By.ID,"GridLabel0_1.Records"), "Registros"))
    try:
        WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "GOTOLAST0_1")))
        hacer_click_elementos(driver, By.ID, "GOTOLAST0_1")
    except TimeoutException:
        logger.warning("No se pudo encontrar o esperar a que el elemento GOTOLAST0_1 sea visible.")
    except Exception as e:
        logger.debug(f"error en compañia {numeros_compania}: {e}")
    try:
        time.sleep(3)
        WebDriverWait(driver, 120).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="GridHeader0_1"]/tbody/tr/td[2]/table/tbody/tr/td[3]/a'))
        )
        WebDriverWait(driver, 60).until_not(
            EC.visibility_of_element_located((By.ID, "GOTOLAST0_1"))
        )
    except Exception as e:
        logger.warning("posible falla en espera de elementos para la descarga del archivo")
    try:
        if esperar_registros(driver):
            verificar_tabla(driver)
            hacer_click_elementos(driver, By.ID, "selectAll0_1")
            hacer_click_elementos(driver, By.XPATH, '//*[@id="GridHeader0_1"]/tbody/tr/td[2]/table/tbody/tr/td[3]/a', timeout=30)
            hacer_click_elementos(driver, By.XPATH, '//*[@id="WebMenuBar"]/tbody/tr/td[1]/table/tbody/tr[2]/td', timeout=30)
            time.sleep(5)
            hacer_click_elementos(driver, By.ID, "C0_89")
        else: 
            logger.warning(f"Error en esperar registros para la compañia {numeros_compania}")
    except TimeoutException:
        logger.error(f"Error en tiempo de espera de descarga del archivo {numeros_compania}")

def reintentos_boton(driver, intentos=3, espera=10):
    # Reintenta clicar el botón de búsqueda y espera que la tabla empiece a mostrar registros.
    for i in range(1, intentos + 1):
        try:
            if not hacer_click_elementos(driver, By.ID, "C0_100", timeout=30):
                raise WebDriverException("No se pudo clicar el botón C0_100.")

            WebDriverWait(driver, 30).until(
                EC.text_to_be_present_in_element((By.ID, "GridLabel0_1.Records"), "Registros")
            )
            return True
        except (ElementClickInterceptedException, MoveTargetOutOfBoundsException, StaleElementReferenceException, WebDriverException) as e:
            logger.warning(
                "reintentos_boton intento %s/%s fallido (%s): %s",
                i,
                intentos,
                type(e).__name__,
                e,
            )
            if i < intentos:
                time.sleep(espera)
        except Exception as e:
            logger.warning(
                "reintentos_boton intento %s/%s fallido: %s",
                i,
                intentos,
                e,
            )
            if i < intentos:
                time.sleep(espera)
    logger.error("reintentos_boton: se agotaron todos los intentos sin éxito.")
    return False

def comprobacion_archivos(carpeta):
    if not carpeta or not isinstance(carpeta, (str, bytes, os.PathLike)):
        return False
    if not os.path.exists(carpeta):
        return False
    # Verificar que es realmente una carpeta
    if not os.path.isdir(carpeta):
        return False
    try:
        elementos = os.listdir(carpeta)
        archivos= []
        for elemento in elementos:
            ruta_completa = os.path.join(carpeta, elemento)
            if not ruta_completa or not isinstance(ruta_completa, (str, bytes, os.PathLike)):
                continue
            if os.path.isfile(ruta_completa):
                try:
                    if os.path.getsize(ruta_completa) >= 0:  # Permite archivos de 0 bytes
                        archivos.append(elemento)
                except (OSError, IOError):
                    # Si no se puede acceder al archivo, lo omitimos
                    continue
        return len(archivos) > 0
    except (OSError, PermissionError, TypeError):
        return False


def verficacion_carpetas(carpeta_base,numeros_compañia):
    resultados = {}
    for numero in numeros_compañia:
        carpeta = os.path.join(carpeta_base, numero)
        if os.path.exists(carpeta):
            archivos = os.listdir(carpeta)
            archivos = [archivo for archivo in archivos if os.path.isfile(os.path.join(carpeta, archivo))]

            if len(archivos) > 0:
                resultados[numero] = True  # La carpeta contiene archivos
            else:
                resultados[numero] = False  # La carpeta está vacía
        else:
            resultados[numero] = False  # La carpeta no existe
    
    return all(resultados.values())

# CORRECCIÓN 3: el bloque finally ahora siempre cierra y limpia el driver correctamente,
# forzando la re-creación de una sesión limpia en cada reintento.
def descargar_para_compania(numero_compania, archivo_salida, url, user, passw, semaforo, max_intentos=3):
    if not archivo_salida:
        logger.error("No se especificó un directorio de salida para la descarga de la compañía %s", numero_compania)
        raise DriverFatalError("Directorio de salida inválido para la descarga.")

    adquirido = False
    if semaforo:
        semaforo.acquire()
        adquirido = True

    driver = None
    intentos = 0
    carpeta = os.path.join(archivo_salida, numero_compania)

    try:
        while intentos < max_intentos:
            try:
                if comprobacion_archivos(carpeta):
                    logger.info("Descarga completada para compañía %s. Archivos encontrados.", numero_compania)
                    logger.info("Descarga completada para el tipo de compañía %s", numero_compania)
                    return

                if driver is None or not sesion_valida(driver):
                    if driver is not None:
                        logger.info("La sesión actual de Selenium no es válida. Reiniciando navegador para compañía %s.", numero_compania)
                        cerrar_driver(driver)
                        driver = None
                    driver, carpeta, *_ = iniciar_driver(archivo_salida, url, numero_compania)
                else:
                    try:
                        driver.get(url)
                    except WebDriverException as e:
                        logger.warning(
                            "No se pudo recargar la URL inicial en el driver existente: %s. Reiniciando el navegador.",
                            e,
                        )
                        cerrar_driver(driver)
                        driver = None
                        continue

                if not login(driver, "incorrectos|error de usuario|credenciales", user, passw):
                    raise DriverRetryableError("No se pudo iniciar sesión en JDE para la compañía %s." % numero_compania)
                descargar_informes(driver, numero_compania)

                if comprobacion_archivos(carpeta):
                    logger.info(f"Descarga completada para #compañía {numero_compania}. Archivos encontrados.")
                    logger.info("Descarga completada para el tipo de compañía %s", numero_compania)
                    return
                else:
                    intentos += 1
                    logger.warning(f"No se encontraron archivos para #compañía {numero_compania}. Intento {intentos}/{max_intentos}.")
                    if intentos < max_intentos:
                        logger.info(f"Reintentando descarga para #compañía {numero_compania}...")
                        if driver is not None and not sesion_valida(driver):
                            cerrar_driver(driver)
                            driver = None
                        time.sleep(20)  # FIX: aumentado a 20s para dar más tiempo entre reintentos
                    else:
                        logger.error(f"Se alcanzó el máximo de intentos ({max_intentos}) para #compañía {numero_compania}. No se pudo completar la descarga.")
                        logger.error("Se alcanzó el máximo de intentos (%s) para #compañía %s. No se pudo completar la descarga.", max_intentos, numero_compania)

            except DriverFatalError as e:
                logger.error("Error crítico en descarga del informe para compañía %s: %s", numero_compania, e)
                logger.error("ERROR crítico al iniciar el navegador para la compañía %s. Revise los logs.", numero_compania)
                return
            except DriverRetryableError as e:
                intentos += 1
                logger.warning("Error transitorio en descarga del informe de compañía %s: %s", numero_compania, e)
                if intentos < max_intentos:
                    logger.info("Reintentando para compañía %s... (Intento %s/%s)", numero_compania, intentos, max_intentos)
                    time.sleep(10)
                else:
                    logger.error("Se alcanzó el máximo de intentos (%s) para compañía %s. No se pudo completar la descarga.", max_intentos, numero_compania)
            except Exception as e:
                intentos += 1
                logger.error("Error en descarga del informe de compañía %s: %s", numero_compania, e)
                if intentos < max_intentos:
                    logger.info("Reintentando para compañía %s... (Intento %s/%s)", numero_compania, intentos, max_intentos)
                    time.sleep(10)
                else:
                    logger.error("Se alcanzó el máximo de intentos (%s) para compañía %s. No se pudo completar la descarga.", max_intentos, numero_compania)
                    logger.error("Se alcanzó el máximo de intentos (%s) para compañía %s. No se pudo completar la descarga.", max_intentos, numero_compania)

    finally:
        # Siempre cerrar el driver al finalizar cada intento, sin importar si fue exitoso o no.
        # Esto garantiza que el próximo intento empiece con un navegador y sesión completamente nuevos.
        if driver is not None:
            try:
                cerrar_driver(driver)
            except Exception as e:
                logger.error(f"Error al cerrar driver de #compañía {numero_compania}: {e}")
            driver = None  # Forzar re-creación en el próximo intento

        if adquirido and semaforo:
            try:
                semaforo.release()
            except Exception as e:
                logger.error(f"Error liberando el semáforo para #compañía {numero_compania}: {e}")


    
def navegacion(numeros_compania, archivo_salida, url, user, passw, max_hilos=2):
    # FIX: max_hilos reducido a 2 para evitar saturación de sesiones JDE simultáneas
    semaforo = threading.Semaphore(max_hilos)
    hilos = []
    for numero in numeros_compania:
        hilo = threading.Thread(target=descargar_para_compania, args=(numero, archivo_salida, url, user, passw, semaforo,))
        hilo.start()
        hilos.append(hilo)
        time.sleep(5)  # FIX: aumentado a 5s para dar tiempo a que cada sesión se establezca
    for hilo in hilos:
        hilo.join()

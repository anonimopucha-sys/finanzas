from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.keys import Keys
import time
import openpyxl
import unicodedata
from datetime import datetime
import threading
import pandas as pd
from typing import Any, cast
from .logs_config import configurar_logger
from .hojas_excel import obtener_hoja_excel
from .descarga_informes import (
    iniciar_driver, login, cambiar_a_iframe, hacer_click_elementos, verificar_ventanas, cambiar_ventana, ingresar_texto_jde,
    regresar_al_contexto_principal, cerrar_driver, ingresar_texto_jde_con_copiar_pegar
)
from .driver_manager import DriverFatalError

logger = configurar_logger()

def verificar_boton(driver: Any, handle: str, boton_xpath: str) -> bool:
    driver.switch_to.window(handle)
    try:
        hacer_click_elementos(driver, By.XPATH, boton_xpath)
        return True
    except (TimeoutException, WebDriverException):
        return False

def cerrar_ventanas(driver: Any) -> None:
    WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) == 4)
    ventanas = driver.window_handles
    ventana_4 = ventanas[3]
    ventana_3 = ventanas[2]
    try:
        if verificar_boton(driver, ventana_4, '//*[@id="C0_52"]'):
            WebDriverWait(driver, 10).until(lambda driver: len(driver.window_handles) == 3)
    except Exception:
        logger.warning("pasar ventana")

    try:
        if verificar_boton(driver, ventana_3, '//*[@id="C0_52"]'):
            WebDriverWait(driver, 10).until(lambda driver: len(driver.window_handles) == 3)
    except Exception:
        logger.warning("error en todas las ventanas")
    iframes = ['//*[@id="e1menuAppIframe"]']
    cambiar_ventana(driver, 2)
    cambiar_a_iframe(driver, iframes)
    hacer_click_elementos(driver, By.XPATH, '//*[@id="C0_5"]')
    WebDriverWait(driver, 10).until(lambda driver: len(driver.window_handles) == 2)
    cambiar_ventana(driver)

def extraer_batchs(ruta_excel: str) -> dict[str, list[str]]:
    hoja = obtener_hoja_excel(ruta_excel, ("Transac. Pend. Por Contabilizar", "Transacciones pendientes por contabilizar"))
    df = cast(
        pd.DataFrame,
        pd.read_excel(
            ruta_excel,
            header=None,
            engine="openpyxl",
            sheet_name=hoja,
        ),
    )  # type: ignore[reportUnknownMemberType]
    filtrado = df.loc[df[4].isin(["PU", "JE", "A5"])]
    subset_columns: list[Any] = [21]
    filtrado = filtrado.dropna(subset=subset_columns)
    batchs_dict: dict[str, list[str]] = {}
    for _, row in filtrado.iterrows():
        tipo_doc = str(row[4])
        batch = str(row[21]).strip()
        responsable = row.iloc[0]
        if pd.isna(responsable) or str(responsable) == "":
            batchs_dict.setdefault(tipo_doc, [])
            if batch not in batchs_dict[tipo_doc]:
                batchs_dict[tipo_doc].append(batch)
    return batchs_dict

def extraer_ct(ruta_excel: str) -> pd.DataFrame:
    hoja = obtener_hoja_excel(ruta_excel, ("Transac. Pend. Por Contabilizar", "Transacciones pendientes por contabilizar"))
    df = cast(
        pd.DataFrame,
        pd.read_excel(
            ruta_excel,
            engine="openpyxl",
            sheet_name=hoja,
        ),
    )  # type: ignore[reportUnknownMemberType]
    df = df.loc[df.iloc[:, 4] == "CT"]
    df_seleccionado = df.iloc[:, [2, 5, 6, 15, 26]].copy()
    columna_original = df_seleccionado.iloc[:, 0]
    columna_texto = columna_original.astype(str)
    df_seleccionado["Numero PV"] = columna_texto.str.extract(r"PV-(\d+).")
    # Mantener solo los registros con número PV extraído
    mask = columna_original.notna() & (columna_texto != "")
    df_seleccionado = df_seleccionado[mask]
    return df_seleccionado

def mover_scroll_ct(driver: Any, colindex: int, acumulado: int = 5000, max_intentos: int = 10) -> str | None:
    intentos = 0
    desplazamiento = 0
    while intentos < max_intentos:
        try:
            driver.execute_script("""
                var elem = document.querySelector('#jdeGridBack0_1');
                if (elem) {
                    elem.scrollLeft = arguments[0];
                }
            """, desplazamiento)

            xpath_elemento = f"//*[@id='G0_1_R0']/td[@colindex='{colindex}']/div"
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, xpath_elemento)))
            elemento = driver.find_element(By.XPATH, xpath_elemento)
            valor = elemento.text
            return valor
        except Exception as e:
            desplazamiento += acumulado
            time.sleep(1)
            intentos += 1
    return None

def scroll_pu(driver: Any, colindex: int = 96, acumulado: int = 2000, max_intentos: int = 20) -> str | None:
    intentos = 0
    desplazamiento = 0
    while intentos < max_intentos:
        try:
            driver.execute_script("""
                var elem = document.querySelector('#jdeGridBack0_209');
                if (elem) {
                    elem.scrollLeft = arguments[0];
                }
            """, desplazamiento)

            xpath_elemento = f"//*[@id='G0_209_R0']/td[@colindex='{colindex}']/div"
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, xpath_elemento)))
            elemento = driver.find_element(By.XPATH, xpath_elemento)
            valor = elemento.text
            driver.execute_script("""
                var elem = document.querySelector('#jdeGridBack0_209');
                if (elem) {
                    elem.scrollLeft = 0;
                }
            """)
            return valor
        except Exception as e:
            desplazamiento += acumulado
            time.sleep(1)
            intentos += 1
    return None

def analizar_pu(driver: Any, batches: dict[str, list[str]], responsables_pu: dict[str, str | None], user: str, passw: str, url_base: str) -> None:
    max_intentos = 3
    intento = 0
    while intento < max_intentos:
        try:
            login(driver, "incorrectos", user, passw)
            time.sleep(3)
            iframes = ['//*[@id="e1menuAppIframe"]']
            cambiar_a_iframe(driver, iframes)
            hacer_click_elementos(driver, By.XPATH, '//*[@id="tab1"]')
            driver.switch_to.default_content()
            time.sleep(2)
            iframes = ['//*[@id="e1menuAppIframe"]', '//*[@id="wcFrame1"]', '//*[@id="RIPaneIFRAME1"]']
            cambiar_a_iframe(driver, iframes)
            registro_activos_xpath = '//*[@id="pageContainer"]/table/tbody[1]/tr/td/div[2]/table/tbody/tr/td[1]/div/table/tbody/tr[2]/td[2]/div/table/tr[2]/td/span'
            hacer_click_elementos(driver, By.XPATH, registro_activos_xpath)
            WebDriverWait(driver, timeout=10).until(lambda driver: len(driver.window_handles) == 2)
            verificar_ventanas(driver)
            cambiar_ventana(driver)
            iframes = ['//*[@id="e1menuAppIframe"]']
            cambiar_a_iframe(driver, iframes)
            hacer_click_elementos(driver, By.XPATH, '//*[@id="WebMenuBar"]/tbody/tr/td[13]/table/tbody/tr/td[3]/a')
            hacer_click_elementos(driver, By.XPATH, '//*[@id="Examinador_de_datos"]/tbody/tr/td[2]/span/nobr')
            cerrar_ventanas(driver)
            hacer_click_elementos(driver, By.XPATH, '//*[@id="table"]')
            ingresar_texto_jde(driver, By.XPATH, '//*[@id="tableName"]', "F0911")
            hacer_click_elementos(driver, By.XPATH, '//*[@id="tableName"]')
            hacer_click_elementos(driver, By.XPATH, '/html/body/table[5]/tbody/tr/td[2]/form/table/tbody/tr/td/table/tbody/tr[11]/td[2]/input')
            hacer_click_elementos(driver, By.XPATH, '/html/body/table[8]/tbody/tr/td[2]/ul/li[1]/a')
            WebDriverWait(driver, timeout=10).until(lambda driver: len(driver.window_handles) == 3)
            cambiar_ventana(driver)
            break
        except DriverFatalError as e:
            logger.error("Error crítico al intentar buscar PU: %s", e)
            cerrar_driver(driver)
            return
        except Exception as e:
            logger.error("Error al intentar buscar PU: %s", e)
            cerrar_driver(driver)
            try:
                driver, *_ = iniciar_driver(None, url_base)
            except DriverFatalError as fatal_error:
                logger.error("Error crítico al reintentar buscar PU: %s", fatal_error)
                return
            intento += 1
            if intento == 3:
                logger.error("Fallo total en analizar PU")
                return

    hacer_click_elementos(driver, By.XPATH, '//*[@id="gtab0_209"]')
    hacer_click_elementos(driver, By.XPATH, '//*[@id="gtab0_209"]/option')
    intento = 0
    for tipo_doc, n_batch in batches.items():
        for batch in n_batch:
            while intento < max_intentos:
                try:
                    ingresar_texto_jde(driver, By.XPATH, "//*[@id='qbeRow0_209']/td[@colindex='1']/div/nobr/input", tipo_doc)
                    scroll_n_batch(driver, batch)
                    hacer_click_elementos(driver, By.XPATH, '//*[@id="C0_51"]')
                    WebDriverWait(driver, timeout=60).until(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="GridLabel0_209.Records"]'), "Registros"))
                    responsable = scroll_pu(driver)
                    responsables_pu[batch] = responsable
                    logger.debug("responsable del numero de batch: %s, es el usuario %s", batch, responsable)
                    break
                except Exception as e:
                    logger.error(f"Error en analizar batch: {e}")
                    intento += 1
                    driver.refresh()
                    if intento == 3:
                        logger.error(f"fallo tottal en analizar pu")
                        return

    cerrar_driver(driver)


def normalizar_texto(texto: str) -> str:
    return ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    ).upper()

def scroll_ordenes(driver: Any, texto: str, retroceder: bool = True, colindex: int = 22, acumulado: int = 500, max_intentos: int = 10) -> str | None:
    intentos = 0
    desplazamiento = 0
    while intentos < max_intentos:
        try:
            driver.execute_script("""
                var elem = document.querySelector('#jdeGridBack0_1');
                if (elem) {
                    elem.scrollLeft = arguments[0];
                }
            """, desplazamiento)

            xpath_elemento = f"//*[@id='qbeRow0_1']/td[@colindex='{colindex}']/div/nobr/input"
            cuenta_visible = WebDriverWait(driver, timeout=2).until(
                EC.element_to_be_clickable((By.XPATH, xpath_elemento))
            )
            if cuenta_visible:
                ingresar_texto_jde_con_copiar_pegar(driver, By.XPATH, xpath_elemento, texto)
                time.sleep(1)
                if retroceder == True:
                    driver.execute_script("""
                    var elem = document.querySelector('#jdeGridBack0_1');
                    if (elem) {
                        elem.scrollLeft = 0;
                    }
                    """)
                break
        except Exception as e:
            desplazamiento += acumulado
            time.sleep(1)
            intentos += 1
    return None

def scroll_n_batch(driver: Any, texto: str, retroceder: bool = True, colindex: int = 7, acumulado: int = 500, max_intentos: int = 10) -> None:
    intentos = 0
    desplazamiento = 0
    while intentos < max_intentos:
        try:
            driver.execute_script("""
                var elem = document.querySelector('#jdeGridBack0_209');
                if (elem) {
                    elem.scrollLeft = arguments[0];
                }
            """, desplazamiento)

            xpath_elemento = f"//*[@id='qbeRow0_209']/td[@colindex='{colindex}']/div/nobr/input"
            batch_visible = WebDriverWait(driver, timeout=2).until(
                EC.element_to_be_clickable((By.XPATH, xpath_elemento))
            )
            if batch_visible:
                ingresar_texto_jde_con_copiar_pegar(driver, By.XPATH, xpath_elemento, texto)
                time.sleep(1)
                if retroceder == True:
                    driver.execute_script("""
                    var elem = document.querySelector('#jdeGridBack0_209');
                    if (elem) {
                        elem.scrollLeft = 0;
                    }
                    """)
                break
        except Exception as e:
            desplazamiento += acumulado
            time.sleep(1)
            intentos += 1
    return None

def scroll_ordenes_origen(driver: Any, colindex: int = 71, acumulado: int = 1000, max_intentos: int = 10) -> str | None:
    intentos = 0
    desplazamiento = 0
    while intentos < max_intentos:
        try:
            driver.execute_script("""
                var elem = document.querySelector('#jdeGridBack0_1');
                if (elem) {
                    elem.scrollLeft = arguments[0];
                }
            """, desplazamiento)

            xpath_elemento = f"//*[@id='G0_1_R0']/td[@colindex='{colindex}']/div"
            origen_visible = WebDriverWait(driver, timeout=2).until(
                EC.element_to_be_clickable((By.XPATH, xpath_elemento))
            )
            if origen_visible:
                elemento = driver.find_element(By.XPATH, xpath_elemento)
                valor = elemento.text
                driver.execute_script("""
                var elem = document.querySelector('#jdeGridBack0_1');
                if (elem) {
                    elem.scrollLeft = 0;
                }
                """)
                return valor
        except Exception as e:
            desplazamiento += acumulado
            time.sleep(1)
            intentos += 1
    return None

def scroll_interventor(driver: Any, colindex: int = 33, acumulado: int = 1000, max_intentos: int = 10) -> str | None:
    intentos = 0
    desplazamiento = 0
    while intentos < max_intentos:
        try:
            driver.execute_script("""
                var elem = document.querySelector('#jdeGridBack0_1');
                if (elem) {
                    elem.scrollLeft = arguments[0];
                }
            """, desplazamiento)

            xpath_elemento = f"//*[@id='G0_1_R0']/td[@colindex='{colindex}']/div"
            interventor_visible = WebDriverWait(driver, timeout=2).until(
                EC.element_to_be_clickable((By.XPATH, xpath_elemento))
            )
            if interventor_visible:
                elemento = driver.find_element(By.XPATH, xpath_elemento)
                valor = elemento.text
                driver.execute_script("""
                var elem = document.querySelector('#jdeGridBack0_1');
                if (elem) {
                    elem.scrollLeft = 0;
                }
                """)
                return valor
        except Exception as e:
            desplazamiento += acumulado
            time.sleep(1)
            intentos += 1
    return None

def scroll_n_comprador(driver: Any, colindex: int, colindex_env: int, acumulado: int = 800, max_intentos: int = 5) -> tuple[str | None, str | None, str | None]:
    intentos = 0
    desplazamiento = 0
    while intentos < max_intentos:
        try:
            driver.execute_script("""
                var elem = document.querySelector('#jdeGridBack0_1');
                if (elem) {
                    elem.scrollLeft = arguments[0];
                }
            """, desplazamiento)

            xpath_comprador = f"//*[@id='G0_1_R0']/td[@colindex='{colindex}']/div/input"
            xpath_envio = f"//*[@id='G0_1_R0']/td[@colindex='{colindex_env}']/div/input"

            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, xpath_envio)))
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, xpath_comprador)))

            comprador = driver.find_element(By.XPATH, xpath_comprador)
            envio = driver.find_element(By.XPATH, xpath_envio)

            valor_comprador = comprador.get_attribute("value")
            valor_envio = envio.get_attribute("value")

            return valor_comprador, valor_envio, xpath_comprador
        except Exception as e:
            desplazamiento += acumulado
            time.sleep(1)
            intentos += 1
    return None, None, None

def extraer_pv_ov(ruta_excel: str) -> tuple[dict[str, list[str | None]], dict[str, list[str | None]], dict[str, list[str | None]]]:
    hoja = obtener_hoja_excel(ruta_excel, ("Transac. Pend. Por Contabilizar", "Transacciones pendientes por contabilizar"))
    df = cast(
        pd.DataFrame,
        pd.read_excel(ruta_excel, sheet_name=hoja),
    )  # type: ignore[reportUnknownMemberType]
    if max(15, 26) >= len(df.columns):
        raise IndexError("Índice de columna fuera de rango.")

    col_filtro = df.columns[4]
    df[col_filtro] = df[col_filtro].astype(str).str.strip()
    df = df.loc[df[col_filtro].isin(["OV", "PV"])]

    col_orden = df.columns[26]
    col_cuenta = df.columns[15]
    subset_columns: list[Any] = [col_orden, col_cuenta]
    df = df.dropna(subset=subset_columns)
    df[col_orden] = df[col_orden].astype(str).str.strip()
    df[col_cuenta] = df[col_cuenta].astype(str)
    combinaciones = df[[col_orden, col_cuenta]].drop_duplicates()
    resultado: dict[str, list[str | None]] = {}
    for _, row in combinaciones.iterrows():
        orden = str(row[col_orden]).strip()
        cuenta = str(row[col_cuenta]).strip()

        registros_combinacion = df[(df[col_orden] == orden) & (df[col_cuenta] == cuenta)]
        if registros_combinacion.iloc[:, 0].isna().any():
            resultado[f"{orden}-{cuenta}"] = [orden, cuenta, "", "", ""]

    if not resultado:
        logger.warning("No se encontraron combinaciones válidas.")
        return {}, {}, {}
    logger.info("Se harán %s búsqueda(s) de registros PV y OV", len(resultado))
    dict1, dict2, dict3 = dividir_dict(resultado)
    return dict1, dict2, dict3

def dividir_dict(diccionario: dict[str, list[str | None]]) -> tuple[dict[str, list[str | None]], dict[str, list[str | None]], dict[str, list[str | None]]]:
    items = list(diccionario.items())
    if not items:
        return {}, {}, {}
    if len(items) == 1:
        return {items[0][0]: items[0][1]}, {}, {}
    elif len(items) == 2:
        return {items[0][0]: items[0][1]}, {items[1][0]: items[1][1]}, {}
    tercio = len(items) // 3
    dict1 = dict(items[:tercio])
    dict2 = dict(items[tercio:2 * tercio])
    dict3 = dict(items[2 * tercio:])
    return dict1, dict2, dict3

# FIX: driver.close() envuelto en try/except para manejar ventana ya cerrada
def reintento_orden(driver: Any):
    try:
        driver.close()
    except Exception as e:
        logger.warning(f"No se pudo cerrar la ventana (posiblemente ya estaba cerrada): {e}")

    try:
        WebDriverWait(driver, 10).until(
            lambda driver: len(driver.window_handles) == 1
        )
    except Exception as e:
        logger.warning(f"Timeout esperando que quede 1 ventana: {e}")

    try:
        cambiar_ventana(driver)
    except Exception as e:
        logger.warning(f"No se pudo cambiar de ventana: {e}")
        return

    time.sleep(1)
    iframes = ['//*[@id="e1menuAppIframe"]', '//*[@id="wcFrame1"]', '//*[@id="RIPaneIFRAME1"]']
    cambiar_a_iframe(driver, iframes)
    xpath = '//*[@id="pageContainer"]/table/tbody[37]/tr/td/div[2]/table/tbody/tr/td[1]/div/table/tbody/tr[2]/td[2]/div/table/tr[2]/td/span'
    hacer_click_elementos(driver, By.XPATH, xpath)
    WebDriverWait(driver, 10).until(lambda driver: len(driver.window_handles) == 2)
    cambiar_ventana(driver)
    iframes = ['//*[@id="e1menuAppIframe"]']
    cambiar_a_iframe(driver, iframes)

def extraer_empleados(excel_reglas: str, hoja: str = "PERSONAL_CENS") -> pd.DataFrame:
    df = cast(
        pd.DataFrame,
        pd.read_excel(excel_reglas, sheet_name=hoja, skiprows=1),
    )  # type: ignore[reportUnknownMemberType]
    df_fltrado = df.iloc[:, [1, 4]].copy()
    return df_fltrado

def comprobacion_empleados(df: pd.DataFrame, empleado: str | None) -> str | None:
    df.iloc[:, 0] = df.iloc[:, 0].astype(str)
    resultado = df[df.iloc[:, 0].str.contains(str(empleado), na=False, case=False)]

    if not resultado.empty:
        return str(resultado.iloc[0, 1])
    df_temp = df.copy()
    df_temp.iloc[:, 0] = df_temp.iloc[:, 0].apply(normalizar_texto)
    empleado_normalizado = normalizar_texto(str(empleado))
    resultado = df_temp[df_temp.iloc[:, 0].str.contains(empleado_normalizado, na=False)]

    if not resultado.empty:
        index_original = resultado.index[0]
        return str(df.iloc[index_original, 1])
    else:
        logger.info(f"el empleado {empleado} no fue encontrado en el archivo.")
        return None

def consultar_orden_ct(driver: Any, dict_ct: dict[str, list[str | None]], empleados: pd.DataFrame, responsables_ct: dict[str, list[str | None]]) -> None:
    try:
        cambiar_ventana(driver)
        xpath = '//*[@id="pageContainer"]/table/tbody[37]/tr/td/div[2]/table/tbody/tr/td[1]/div/table/tbody/tr[2]/td[2]/div/table/tr[2]/td/span'
        iframes = ['//*[@id="e1menuAppIframe"]', '//*[@id="wcFrame1"]', '//*[@id="RIPaneIFRAME1"]']
        cambiar_a_iframe(driver, iframes)
        hacer_click_elementos(driver, By.XPATH, xpath)
        WebDriverWait(driver, timeout=10).until(lambda driver: len(driver.window_handles) == 2)
        cambiar_ventana(driver)
        iframes = ['//*[@id="e1menuAppIframe"]']
        cambiar_a_iframe(driver, iframes)
        max_intentos = 3
        hacer_click_elementos(driver, By.XPATH, '//*[@id="gtab0_1"]')
        hacer_click_elementos(driver, By.XPATH, '//*[@id="gtab0_1"]/option')
        for orden, lista in dict_ct.items():
            intentos = 0
            while intentos < max_intentos:
                try:
                    texto_orden = str(lista[0]) if lista[0] is not None else ""
                    texto_cuenta = str(lista[1]) if lista[1] is not None else ""
                    ingresar_texto_jde(driver, By.XPATH, '//*[@id="qbeRow0_1"]/td[2]/div/nobr/input', texto_orden)
                    scroll_ordenes(driver, texto_cuenta)
                    hacer_click_elementos(driver, By.XPATH, '//*[@id="C0_8"]')
                    time.sleep(2)
                    hacer_click_elementos(driver, By.XPATH, '//*[@id="G0_1_R0"]/td[3]/div/a')
                    time.sleep(2)
                    break
                except Exception as e:
                    logger.error(f"error en escribir orden de compra: {e}")
                    intentos += 1
                    reintento_orden(driver)
                    if intentos == 3:
                        logger.error("error al analizar CT")
                        return
            try:
                responsable = driver.find_element(By.XPATH, '//*[@id="divTC0_769"]/span/i')
                responsable = responsable.text
                dependencia = comprobacion_empleados(empleados, responsable)
                if dependencia is None:
                    numero_comprador, numero_envio, xpath_comprador = scroll_n_comprador(driver, 34, 33)
                    input_comprador = driver.find_element(By.XPATH, xpath_comprador)
                    input_comprador.send_keys(Keys.F2)
                    time.sleep(3)
                    iframes = ['//*[@id="modalIframe1"]']
                    cambiar_a_iframe(driver, iframes)
                    texto_a_ingresar = numero_comprador if numero_comprador is not None else (numero_envio or "")
                    ingresar_texto_jde(driver, By.XPATH, '//*[@id="qbeRow0_1"]/td[2]/div/nobr/input', texto_a_ingresar)

                    hacer_click_elementos(driver, By.XPATH, '//*[@id="C0_13"]')
                    WebDriverWait(driver, timeout=20).until(EC.text_to_be_present_in_element((By.ID, "GridLabel0_1.Records"), "Registros"))

                    direccion = driver.find_element(By.XPATH, '//*[@id="G0_1_R0"]/td[2]/div')
                    empleado = driver.find_element(By.XPATH, '//*[@id="G0_1_R0"]/td[3]/div')
                    valor_empleado = empleado.text

                    n_dependencia = comprobacion_empleados(empleados, valor_empleado)
                    if n_dependencia is None:
                        n_dependencia = 'DESCONOCIDA'  # FIX: era `==` en lugar de `=`

                    responsables_ct[orden] = [
                        lista[0], lista[1], valor_empleado, n_dependencia, lista[3], lista[2]
                    ]
                    hacer_click_elementos(driver, By.XPATH, '//*[@id="C0_3"]')
                    regresar_al_contexto_principal(driver)
                    iframes = ['//*[@id="e1menuAppIframe"]']
                    cambiar_a_iframe(driver, iframes)
                else:
                    responsables_ct[orden] = [
                        lista[0], lista[1], responsable, dependencia, lista[3], lista[2]
                    ]
            except Exception as e:
                logger.error(f"error en ingreso de cts: {e}")
            hacer_click_elementos(driver, By.XPATH, '//*[@id="C0_5"]')
            time.sleep(2)
    except Exception as e:
        logger.error(f"error en ct {e}")
    cerrar_driver(driver)

def analizar_pv_ov(driver: Any, diccionario: dict[str, list[str | None]], resultado: dict[str, list[str | None]], numeros_origen: set[str], user: str, passw: str, url_base: str) -> None:
    max_intentos = 3
    intento = 0
    while intento < max_intentos:
        try:
            logger.info("Se inicia busqueda de PV y OV")
            login(driver, "incorrectos", user, passw)
            time.sleep(3)
            iframes = ['//*[@id="e1menuAppIframe"]']
            cambiar_a_iframe(driver, iframes)
            hacer_click_elementos(driver, By.XPATH, '//*[@id="tab1"]')
            driver.switch_to.default_content()
            time.sleep(2)
            iframes = ['//*[@id="e1menuAppIframe"]', '//*[@id="wcFrame1"]', '//*[@id="RIPaneIFRAME1"]']
            cambiar_a_iframe(driver, iframes)
            xpath = '//*[@id="pageContainer"]/table/tbody[37]/tr/td/div[2]/table/tbody/tr/td[1]/div/table/tbody/tr[2]/td[2]/div/table/tr[2]/td/span'
            hacer_click_elementos(driver, By.XPATH, xpath)
            WebDriverWait(driver, timeout=10).until(lambda driver: len(driver.window_handles) == 2)
            cambiar_ventana(driver)
            iframes = ['//*[@id="e1menuAppIframe"]']
            cambiar_a_iframe(driver, iframes)
            break
        except DriverFatalError as e:
            logger.error("Error crítico al analizar PV y OV: %s", e)
            cerrar_driver(driver)
            return
        except Exception as e:
            logger.error("Error al analizar PV y OV: %s", e)
            intento += 1
            cerrar_driver(driver)
            try:
                driver, *_ = iniciar_driver(None, url_base)
            except DriverFatalError as fatal_error:
                logger.error("Error crítico al reintentar analizar PV y OV: %s", fatal_error)
                return
            if intento == 3:
                logger.error("Error total al analizar PV y OV")
                return

    hacer_click_elementos(driver, By.XPATH, '//*[@id="gtab0_1"]')
    hacer_click_elementos(driver, By.XPATH, '//*[@id="gtab0_1"]/option')
    max_intentos = 3
    for combinacion, lista in diccionario.items():
        intentos = 0
        while intentos < max_intentos:
            try:
                texto_numero = str(lista[0]) if lista[0] is not None else ""
                texto_cuenta = str(lista[1]) if lista[1] is not None else ""
                ingresar_texto_jde(driver, By.XPATH, "//*[@id='qbeRow0_1']/td[@colindex='0']/div/nobr/input", texto_numero)
                scroll_ordenes(driver, texto_cuenta)
                hacer_click_elementos(driver, By.XPATH, '//*[@id="C0_8"]')
                numero_origen = scroll_ordenes_origen(driver) or ""
                logger.debug("para la combinacion: %s su numero origen es: %s", combinacion, numero_origen)
                time.sleep(1)
                resultado[combinacion] = [
                    texto_numero, texto_cuenta, "", "", numero_origen
                ]
                break
            except Exception as e:
                logger.error(f"error en ingreso de pv_ov {combinacion} : se hara reintento {intentos + 1}")
                intentos += 1
                reintento_orden(driver)
                if intentos == 3:
                    logger.error("Error al analizar PV y OV")
                    return
        time.sleep(1)
    cerrar_driver(driver)
    for valores in resultado.values():
        numero_origen = valores[4]
        if numero_origen:
            numeros_origen.add(numero_origen)
    return

def analizar_numeros_origen(driver: Any, empleados: pd.DataFrame, dict_origen: dict[str, list[str | None]], numeros_origen: set[str], user: str, passw: str, url_base: str) -> None:
    login(driver, "incorrectos", user, passw)
    hacer_click_elementos(driver, By.XPATH, '//*[@id="drop_mainmenu"]')
    hacer_click_elementos(driver, By.XPATH, '/html/body/div[3]/div/div[2]/div[3]/div/div[2]/div[4]/div/div/div[2]/div[1]/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td[1]/span')
    hacer_click_elementos(driver, By.XPATH, '/html/body/div[10]/table/tbody/tr/td/div/div[1]/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td[1]/span')
    hacer_click_elementos(driver, By.XPATH, '/html/body/div[11]/table/tbody/tr/td/div/div/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td[1]/span')
    hacer_click_elementos(driver, By.XPATH, '/html/body/div[12]/table/tbody/tr/td/div/div[5]/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td[1]/span')
    hacer_click_elementos(driver, By.XPATH, '/html/body/div[13]/table/tbody/tr/td/div/div[2]/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td[1]/span')
    hacer_click_elementos(driver, By.XPATH, '/html/body/div[14]/table/tbody/tr/td/div/div[3]/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td[1]/span')
    hacer_click_elementos(driver, By.XPATH, '/html/body/div[15]/table/tbody/tr/td/div/div[13]/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td[1]/a')
    WebDriverWait(driver, timeout=10).until(lambda driver: len(driver.window_handles) == 2)
    cambiar_ventana(driver)
    iframes = ['//*[@id="e1menuAppIframe"]']
    cambiar_a_iframe(driver, iframes)
    hacer_click_elementos(driver, By.XPATH, '//*[@id="gtab0_1"]')
    hacer_click_elementos(driver, By.XPATH, '//*[@id="gtab0_1"]/option')
    for n_origen in numeros_origen:
            ingresar_texto_jde_con_copiar_pegar(driver, By.XPATH, '//*[@id="C0_36"]', n_origen)
            for i in range(2):
                hacer_click_elementos(driver, By.XPATH, '//*[@id="C0_15"]')
            interventor = scroll_interventor(driver)
            responsable = comprobacion_empleados(empleados, interventor)
            dict_origen[n_origen] = [
                interventor, responsable
            ]
            logger.debug("para el numero: %s su interventor es: %s:(%s)", n_origen, interventor, responsable)

    cerrar_driver(driver)
    return

def analizar_ct(driver: Any, registros_ct: pd.DataFrame, empleados: pd.DataFrame, responsables_ct: dict[str, list[str | None]], user: str, passw: str) -> None:
    try:
        login(driver, "incorrectos", user, passw)
        time.sleep(3)
        iframes = ['//*[@id="e1menuAppIframe"]']
        cambiar_a_iframe(driver, iframes)
        hacer_click_elementos(driver, By.XPATH, '//*[@id="tab1"]')
        driver.switch_to.default_content()
        time.sleep(2)
        iframes = ['//*[@id="e1menuAppIframe"]', '//*[@id="wcFrame1"]', '//*[@id="RIPaneIFRAME1"]']
        cambiar_a_iframe(driver, iframes)

        hacer_click_elementos(driver, By.XPATH, '//*[@id="pageContainer"]/table/tbody[36]/tr/td/div[2]/table/tbody[1]/tr/td[5]/div/table/tbody/tr[2]/td[2]/div/table/tr[2]/td/span')
        time.sleep(3)
        WebDriverWait(driver, timeout=10).until(lambda driver: len(driver.window_handles) == 2)
        verificar_ventanas(driver)
        cambiar_ventana(driver)

        iframes = ['//*[@id="e1menuAppIframe"]']
        cambiar_a_iframe(driver, iframes)
        fecha_hoy = datetime.today().strftime('%y/%m/%d')
        dict_ct: dict[str, list[str | None]] = {}
        for idx, fila in registros_ct.iterrows():
            numero_pv = fila.iloc[5]
            fecha = fila.iloc[2].strftime('%Y/%m/%d')
            if pd.isna(fila.iloc[4]):  # Detectar NaN y None correctamente
                max_intentos = 10
                intentos = 0
                while intentos < max_intentos:
                    try:
                        ingresar_texto_jde(driver, By.XPATH, '//*[@id="C0_47"]', fila.iloc[3])
                        hacer_click_elementos(driver, By.XPATH, '//*[@id="C0_51"]')
                        time.sleep(1)
                        campo = driver.find_element(By.XPATH, '//*[@id="C0_47"]')
                        if fila.iloc[3] in campo.get_attribute('value'):
                            logger.debug("Texto ingresado correctamente.")
                            break
                        else:
                            logger.debug(f"Texto no coincidente en intento {intentos + 1}. Reintentando...")
                    except Exception as e:
                        logger.error(f"Error durante el ingreso de texto")
                    intentos += 1
                ingresar_texto_jde(driver, By.XPATH, '//*[@id="C0_51"]', fecha)
                ingresar_texto_jde(driver, By.XPATH, '//*[@id="C0_53"]', fecha_hoy)
                ingresar_texto_jde(driver, By.XPATH, '//*[@id="qbeRow0_1"]/td[6]/div/nobr/input', 'PV')
                ingresar_texto_jde(driver, By.XPATH, '//*[@id="qbeRow0_1"]/td[8]/div/nobr/input', numero_pv)
                for i in range(2):
                    hacer_click_elementos(driver, By.XPATH, '//*[@id="C0_15"]')
                WebDriverWait(driver, timeout=10).until(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="GridLabel0_1.Records"]'), "Registros"))
                WebDriverWait(driver, timeout=10).until(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="G0_1_R0"]/td[9]/div/a'), numero_pv))
                valor = mover_scroll_ct(driver, 59)
                time.sleep(3)
                driver.execute_script("""
                var elem = document.querySelector('#jdeGridBack0_1');
                if (elem) {
                    elem.scrollLeft = 0;
                }
                """)
                dict_ct[f"{valor}-{str(fila.iloc[3]).strip()}"] = [
                    valor, str(fila.iloc[3]), numero_pv, "PV"
                ]
            else:
                orden = fila.iloc[4]
                dict_ct[f"{orden}-{str(fila.iloc[3]).strip()}"] = [
                    orden, str(fila.iloc[3]), numero_pv, "CT"
                ]
        hacer_click_elementos(driver, By.XPATH, '//*[@id="C0_16"]')
        time.sleep(2)
        WebDriverWait(driver, timeout=10).until(lambda driver: len(driver.window_handles) == 1)
        consultar_orden_ct(driver, dict_ct, empleados, responsables_ct)

    except Exception:
        logger.error("error en registro de ct")

def completar_pendientes(pendientes: dict[str, list[str | None]], responsables_pv_ov: dict[str, list[str | None]], empleados: pd.DataFrame, user: str, passw: str, url_base: str) -> None:
    try:
        driver, *_ = iniciar_driver(None, url_base)
    except DriverFatalError as e:
        logger.error("Error crítico al iniciar el navegador para completar pendientes: %s", e)
        return
    driver = cast(Any, driver)
    driver = cast(Any, driver)
    login(driver, "incorrectos", user, passw)
    time.sleep(3)
    iframes = ['//*[@id="e1menuAppIframe"]']
    cambiar_a_iframe(driver, iframes)
    hacer_click_elementos(driver, By.XPATH, '//*[@id="tab1"]')
    driver.switch_to.default_content()
    time.sleep(2)
    xpath = '//*[@id="pageContainer"]/table/tbody[37]/tr/td/div[2]/table/tbody/tr/td[1]/div/table/tbody/tr[2]/td[2]/div/table/tr[2]/td/span'
    iframes = ['//*[@id="e1menuAppIframe"]', '//*[@id="wcFrame1"]', '//*[@id="RIPaneIFRAME1"]']
    cambiar_a_iframe(driver, iframes)
    hacer_click_elementos(driver, By.XPATH, xpath)
    WebDriverWait(driver, timeout=20).until(lambda driver: len(driver.window_handles) == 2)
    cambiar_ventana(driver)
    iframes = ['//*[@id="e1menuAppIframe"]']
    cambiar_a_iframe(driver, iframes)
    max_intentos = 3
    hacer_click_elementos(driver, By.XPATH, '//*[@id="gtab0_1"]')
    hacer_click_elementos(driver, By.XPATH, '//*[@id="gtab0_1"]/option')

    for combinacion, valores in pendientes.items():
        intentos = 0
        while intentos < max_intentos:
            try:
                texto_orden = str(valores[0]) if valores[0] is not None else ""
                texto_cuenta = str(valores[1]) if valores[1] is not None else ""
                ingresar_texto_jde(driver, By.XPATH, '//*[@id="qbeRow0_1"]/td[2]/div/nobr/input', texto_orden)
                scroll_ordenes(driver, texto_cuenta)
                hacer_click_elementos(driver, By.XPATH, '//*[@id="C0_8"]')
                time.sleep(2)
                hacer_click_elementos(driver, By.XPATH, '//*[@id="G0_1_R0"]/td[3]/div/a')
                time.sleep(2)
                break
            except Exception:
                logger.error("error en escribir orden de compra")
                intentos += 1
                reintento_orden(driver)
                if intentos == 3:
                    logger.error("error al analizar Pendiente")
                    return
        try:
            responsable = driver.find_element(By.XPATH, '//*[@id="divTC0_769"]/span/i')
            responsable = responsable.text
            dependencia = comprobacion_empleados(empleados, responsable)
            logger.debug("%s %s", responsable, dependencia)
            if dependencia is None:  # FIX: usar `is None`
                numero_comprador, numero_envio, xpath_comprador = scroll_n_comprador(driver, 34, 33)
                input_comprador = driver.find_element(By.XPATH, xpath_comprador)
                input_comprador.send_keys(Keys.F2)
                time.sleep(3)
                iframes = ['//*[@id="modalIframe1"]']
                cambiar_a_iframe(driver, iframes)
                texto_a_ingresar = numero_comprador if numero_comprador is not None else (numero_envio or "")
                ingresar_texto_jde(driver, By.XPATH, '//*[@id="qbeRow0_1"]/td[2]/div/nobr/input', texto_a_ingresar)

                hacer_click_elementos(driver, By.XPATH, '//*[@id="C0_13"]')
                WebDriverWait(driver, timeout=20).until(EC.text_to_be_present_in_element((By.ID, "GridLabel0_1.Records"), "Registros"))

                direccion = driver.find_element(By.XPATH, '//*[@id="G0_1_R0"]/td[2]/div')
                empleado = driver.find_element(By.XPATH, '//*[@id="G0_1_R0"]/td[3]/div')
                valor_empleado = empleado.text

                n_dependencia = comprobacion_empleados(empleados, valor_empleado)
                if n_dependencia is None:  # FIX: usar `is None` y asignación correcta
                    n_dependencia = 'DESCONOCIDA'
                logger.info("para la combinacion: %s se encontro como responsable: %s:(%s)", combinacion, valor_empleado, n_dependencia)
                responsables_pv_ov[combinacion] = [
                    valores[0], valores[1], valor_empleado, n_dependencia, valores[4]
                ]
                hacer_click_elementos(driver, By.XPATH, '//*[@id="C0_3"]')
                regresar_al_contexto_principal(driver)
                iframes = ['//*[@id="e1menuAppIframe"]']
                cambiar_a_iframe(driver, iframes)
            else:
                logger.info("para la combinacion: %s se encontro como responsable: %s:(%s)", combinacion, responsable, dependencia)
                responsables_pv_ov[combinacion] = [
                    valores[0], valores[1], responsable, dependencia, valores[4]
                ]
        except Exception as e:
            logger.error(f"error en ingreso de pendientes: {e}")
        hacer_click_elementos(driver, By.XPATH, '//*[@id="C0_5"]')
        time.sleep(2)
    cerrar_driver(driver)

def iniciar_hilos_pv_ov(dicts: list[dict[str, list[str | None]]], drivers: list[Any], resultados: list[dict[str, list[str | None]]], numeros_origen: list[set[str]], user: str, passw: str, url_base: str) -> None:
    hilos: list[threading.Thread] = []
    for diccionario, driver, resultado, numero_origen in zip(dicts, drivers, resultados, numeros_origen):
        if diccionario:
            hilo = threading.Thread(target=analizar_pv_ov, args=(driver, diccionario, resultado, numero_origen, user, passw, url_base,))
            hilo.start()
            hilos.append(hilo)
    for hilo in hilos:
        hilo.join()

def iniciar_hilos_numeros_origen(dicts_origen: list[dict[str, list[str | None]]], drivers: list[Any], empleados: pd.DataFrame, numeros_origen: list[set[str]], user: str, passw: str, url_base: str) -> None:
    hilos: list[threading.Thread] = []
    for dict_origen, driver, numero_origen in zip(dicts_origen, drivers, numeros_origen):
        if numero_origen:
            hilo = threading.Thread(target=analizar_numeros_origen, args=(driver, empleados, dict_origen, numero_origen, user, passw, url_base,))
            hilo.start()
            hilos.append(hilo)
    for hilo in hilos:
        hilo.join()

def dividir_sets(sets: set[str]) -> tuple[set[str], set[str], set[str]]:
    list_sets = list(sets)
    tercio = len(list_sets) // 3
    set1 = set(list_sets[:tercio])
    set2 = set(list_sets[tercio:2 * tercio])
    set3 = set(list_sets[2 * tercio:])
    return set1, set2, set3

def lanzar_instancias(ruta_excel: str, excel_reglas: str, url_base: str, passw: str, user: str) -> None:
    responsables_pu: dict[str, str | None] = {}
    batchs = extraer_batchs(ruta_excel)
    logger.info("Se harán %s búsqueda(s) sobre registros PU, JE y A5", len(batchs))
    hilo_pu: threading.Thread | None = None
    driver_pu = None
    driver_ct = None
    all_drivers: list[Any] = []
    try:
        if batchs:
            try:
                driver_pu, *_ = iniciar_driver(None, url_base)
                if driver_pu is not None:
                    all_drivers.append(driver_pu)
                    hilo_pu = threading.Thread(target=analizar_pu, args=(driver_pu, batchs, responsables_pu, user, passw, url_base,))
                else:
                    logger.error("No se pudo iniciar el navegador para analizar PU/JE/A5.")
            except DriverFatalError as e:
                logger.error("Error crítico al iniciar el navegador para analizar PU/JE/A5: %s", e)

        registros_ct = extraer_ct(ruta_excel)
        responsables_ct: dict[str, list[str | None]] = {}
        empleados = extraer_empleados(excel_reglas)
        hilo_ct: threading.Thread | None = None
        logger.info("Se harán %s búsqueda(s) sobre registros CT", len(registros_ct))
        if not registros_ct.empty:
            try:
                driver_ct, *_ = iniciar_driver(None, url_base)
                if driver_ct is not None:
                    all_drivers.append(driver_ct)
                    hilo_ct = threading.Thread(target=analizar_ct, args=(driver_ct, registros_ct, empleados, responsables_ct, user, passw,))
                else:
                    logger.error("No se pudo iniciar el navegador para analizar CT.")
            except DriverFatalError as e:
                logger.error("Error crítico al iniciar el navegador para analizar CT: %s", e)

        if hilo_pu:
            hilo_pu.start()
        if hilo_ct:
            hilo_ct.start()
        if hilo_pu:
            hilo_pu.join()
        if hilo_ct:
            hilo_ct.join()

        dict1, dict2, dict3 = extraer_pv_ov(ruta_excel)
        resultado1: dict[str, list[str | None]] = {}
        resultado2: dict[str, list[str | None]] = {}
        resultado3: dict[str, list[str | None]] = {}
        drivers: list[Any] = []

        for diccionario in [dict1, dict2, dict3]:
            if diccionario:
                try:
                    driver, *_ = iniciar_driver(None, url_base)
                    drivers.append(driver)
                    if driver is not None:
                        all_drivers.append(driver)
                except Exception as e_drv:
                    logger.error("Error iniciando driver para PV/OV: %s", e_drv)
                    drivers.append(None)
            else:
                drivers.append(None)

        n_origen1: set[str] = set()
        n_origen2: set[str] = set()
        n_origen3: set[str] = set()
        iniciar_hilos_pv_ov([dict1, dict2, dict3], drivers, [resultado1, resultado2, resultado3], [n_origen1, n_origen2, n_origen3], user, passw, url_base)
        union = n_origen1.union(n_origen2, n_origen3)
        n_origen1, n_origen2, n_origen3 = dividir_sets(union)
        drivers = []
        for diccionario in [n_origen1, n_origen2, n_origen3]:
            if diccionario:
                try:
                    driver, *_ = iniciar_driver(None, url_base)
                    if driver is not None:
                        drivers.append(driver)
                        all_drivers.append(driver)
                    else:
                        drivers.append(None)
                except Exception as e_drv:
                    logger.error("Error iniciando driver para números de origen: %s", e_drv)
                    drivers.append(None)
            else:
                drivers.append(None)

        dict_origen1: dict[str, list[str | None]] = {}
        dict_origen2: dict[str, list[str | None]] = {}
        dict_origen3: dict[str, list[str | None]] = {}
        iniciar_hilos_numeros_origen([dict_origen1, dict_origen2, dict_origen3], drivers, empleados, [n_origen1, n_origen2, n_origen3], user, passw, url_base)
        dicts_origen: dict[str, list[str | None]] = {**dict_origen1, **dict_origen2, **dict_origen3}
        responsables_pv_ov: dict[str, list[str | None]] = {**resultado1, **resultado2, **resultado3}
        pendientes: dict[Any, list[str | None]] = {}
    finally:
        for driver in all_drivers:
            if driver:
                try:
                    cerrar_driver(driver)
                except Exception as e:
                    logger.warning("Error cerrando driver al finalizar lanzamiento de instancias: %s", e)

    # CORRECCIÓN 4: Registrar un resumen único de orígenes no encontrados en lugar de
    # un error por cada registro. Antes se generaban decenas de líneas de error idénticas
    # (una por cada fila del Excel) lo que saturaba el log y ocultaba errores reales.
    # Causas comunes: el archivo Relacion.Cuentas.Dependencias.xlsx no tiene mapeados
    # esos códigos de origen, o los informes de JDE traen códigos nuevos que aún no
    # están en la tabla de referencia. Solución: agregar los códigos faltantes al archivo
    # de referencia o contactar al administrador de JDE para validarlos.
    origenes_no_encontrados: set[str] = set()

    for key, value in responsables_pv_ov.items():
        if len(value) < 5:
            logger.error(f"Lista muy corta: {value}")
            continue
        origen = value[4]

        if origen not in dicts_origen:
            # Acumular el origen faltante en lugar de loguear uno a uno
            origenes_no_encontrados.add(str(origen))
            # Marcar como pendiente para que completar_pendientes intente resolverlo
            pendientes[key] = [
                value[0], value[1], None, None, value[4]
            ]
            continue
        nuevos_valores = dicts_origen[origen]
        value[2] = nuevos_valores[0]
        value[3] = nuevos_valores[1]
        if len(value) > 3 and value[2] is None and value[3] is None:
            pendientes[key] = [
                value[0], value[1], value[2], value[3], value[4]
            ]

    # Loguear resumen único de orígenes no encontrados
    if origenes_no_encontrados:
        logger.warning(
            f"Los siguientes {len(origenes_no_encontrados)} orígenes no se encontraron en dict_origen "
            f"y sus registros pasaron a pendientes: {sorted(origenes_no_encontrados)}. "
            f"Verifique que estén en el archivo Relacion.Cuentas.Dependencias.xlsx."
        )

    if pendientes:
        logger.info("Se inicia el proceso de consulta de registros pendientes")
        completar_pendientes(pendientes, responsables_pv_ov, empleados, user, passw, url_base)
    df = aplicar_filtros(ruta_excel, responsables_pu, responsables_ct, responsables_pv_ov, empleados, excel_reglas)
    actalizar_excel(ruta_excel, df)

def extraer_dataframe(ruta_excel: str) -> pd.DataFrame:
    hoja = obtener_hoja_excel(ruta_excel, ("Transac. Pend. Por Contabilizar", "Transacciones pendientes por contabilizar"))
    df = cast(
        pd.DataFrame,
        pd.read_excel(ruta_excel, header=0, sheet_name=hoja),
    )  # type: ignore[reportUnknownMemberType]
    columnas = [0, 2, 4, 15, 21, 26]
    df_extraido = df.iloc[:, columnas]
    return df_extraido

def comprobacion_empleados_pu(df: pd.DataFrame, diccionario_actualizado: dict[str, str | None]) -> dict[str, str | None]:
    df_temp = df.copy()
    df_temp.iloc[:, 0] = df_temp.iloc[:, 0].apply(normalizar_texto)
    for clave, valor in diccionario_actualizado.items():
        valor_str = str(valor) if valor is not None else ""
        nombre_normalizado = sorted(normalizar_texto(valor_str).lower().strip().split())
        nombre_normalizado = ' '.join(nombre_normalizado)
        for _, fila in df_temp.iterrows():
            empleado = sorted(fila.iloc[0].lower().strip().split())
            empleado = ' '.join(empleado)
            if nombre_normalizado in empleado:
                logger.info("para el numero de batch '%s' se asigno como responsable '%s' (%s)", clave, valor, fila.iloc[1])
                diccionario_actualizado[clave] = str(fila.iloc[1])
    return diccionario_actualizado

def determinar_dependencia_pu(excel: str, responsables_pu: dict[str, str | None], empleados: pd.DataFrame) -> dict[str, str | None]:
    hoja = obtener_hoja_excel(excel, ("UsuariosAD_21072025", "UsuariosAD", "Usuarios"), nombre_default="UsuariosAD_21072025")
    df = cast(
        pd.DataFrame,
        pd.read_excel(excel, sheet_name=hoja),
    )  # type: ignore[reportUnknownMemberType]
    df_filtrado = df.iloc[:, [0, 2]].copy()
    diccionario_actualizado = responsables_pu.copy()
    for clave, valor in diccionario_actualizado.items():
        valor_normalizado = str(valor).strip().lower()
        columna_normalizada = df_filtrado.iloc[:, 1].astype(str).str.strip().str.lower()
        coincidencias = df_filtrado[columna_normalizada == valor_normalizado]
        if not coincidencias.empty:
            nuevo_valor = coincidencias.iloc[0, 0]
            diccionario_actualizado[clave] = nuevo_valor
        else:
            logger.warning("No se encontro ningun usuario con el nombre: %s", valor)
    diccionario_actualizado = comprobacion_empleados_pu(empleados, diccionario_actualizado)
    return diccionario_actualizado

def filtro_pu(df: pd.DataFrame, responsables_pu: dict[str, str | None], excel_reglas: str, empleados: pd.DataFrame) -> None:
    filtro_pu_je = (df.iloc[:, 2] == 'PU') | (df.iloc[:, 2] == 'JE') | (df.iloc[:, 2] == 'A5')
    df_pu_je = df[filtro_pu_je].copy()
    responsables_pu = determinar_dependencia_pu(excel_reglas, responsables_pu, empleados)
    for idx, _ in enumerate(df_pu_je.index):
        bact = df_pu_je.iloc[idx, 4]
        try:
            responsable = responsables_pu.get(str(int(bact)), df_pu_je.iloc[idx, 0])
            df_pu_je.iloc[idx, 0] = responsable
        except Exception as e:
            logger.warning(f"No se pudo asignar responsable para fila {idx}: {e}")
            continue
    df.loc[filtro_pu_je, df.columns[0]] = df_pu_je.iloc[:, 0].values

def filtro_ct(df: pd.DataFrame, responsable_ct: dict[str, list[str | None]]) -> None:
    filtro_ct = df.iloc[:, 2] == 'CT'
    df_ct = df[filtro_ct].copy()

    for idx, _ in enumerate(df_ct.index):
        n_orden = df_ct.iloc[idx, 5]
        n_cuenta = str(df_ct.iloc[idx, 3])
        explicacion = str(df_ct.iloc[idx, 1]).strip()

        for datos in responsable_ct.values():
            tipo_doc = str(datos[4]).strip()
            orden = str(int(datos[0])) if datos[0] is not None else ""
            orden = orden.strip()
            cuenta = str(datos[1])
            responsable = datos[3]
            valor_pv = datos[5]
            if pd.isna(n_orden):
                if tipo_doc == 'PV' and (valor_pv is not None and valor_pv in explicacion) and (n_cuenta.strip() in cuenta):
                    df_ct.iloc[idx, 0] = responsable
                    break
            else:
                if tipo_doc == 'CT' and (n_orden == orden) and (n_cuenta.strip() in cuenta):
                    df_ct.iloc[idx, 0] = responsable
                    break
    df.loc[filtro_ct, df.columns[0]] = df_ct.iloc[:, 0].values


def filtro_pv_ov(df: pd.DataFrame, responsables_pv_ov: dict[str, list[str | None]]) -> None:
    filtro_pv_ov = (df.iloc[:, 2] == 'PV') | (df.iloc[:, 2] == 'OV')
    df_pv_ov = df[filtro_pv_ov].copy()

    for idx, _ in enumerate(df_pv_ov.index):
        n_orden = df_pv_ov.iloc[idx, 5]
        n_cuenta = str(df_pv_ov.iloc[idx, 3])
        for datos in responsables_pv_ov.values():
            orden = str(datos[0])
            cuenta = str(datos[1])
            responsable = datos[3]

            if (n_orden == orden) and (n_cuenta.strip() in cuenta):
                df_pv_ov.iloc[idx, 0] = responsable
                break
    df.loc[filtro_pv_ov, df.columns[0]] = df_pv_ov.iloc[:, 0].values


def aplicar_filtros(ruta_excel: str, responsables_pu: dict[str, str | None], responsables_ct: dict[str, list[str | None]], responsables_pv_ov: dict[str, list[str | None]], empleados: pd.DataFrame, excel_reglas: str) -> pd.DataFrame:
    df = extraer_dataframe(ruta_excel)
    filtro_pu(df, responsables_pu, excel_reglas, empleados)
    filtro_ct(df, responsables_ct)
    filtro_pv_ov(df, responsables_pv_ov)
    return df

def actalizar_excel(ruta_excel: str, df: pd.DataFrame) -> None:
    wb = openpyxl.load_workbook(ruta_excel)
    sheet = cast(Any, wb.active)
    data = df.values.tolist()
    for i, row in enumerate(data, start=2):
        sheet.cell(row=i, column=1, value=row[0])
    wb.save(ruta_excel)
    logger.info("archivo guardado correctamente, por favor revisar la ruta: %s", ruta_excel)
import pandas as pd
import os
import json
from openpyxl import load_workbook
from .logs_config import configurar_logger

OTES = ["2C", "2K", "2E"]
logger = configurar_logger()


# ─────────────────────────────────────────────────────────────────────────────
# Lectura y procesamiento del Excel de reglas
# ─────────────────────────────────────────────────────────────────────────────

def leer_excel(excel_reglas: str) -> dict:
    """Lee todas las hojas del archivo Excel sin encabezado."""
    return pd.read_excel(excel_reglas, sheet_name=None, header=None)


def guardar_json(datos: dict, ruta_json: str) -> None:
    """Guarda los datos procesados en un archivo JSON."""
    directorio = os.path.dirname(ruta_json)
    if directorio:
        os.makedirs(directorio, exist_ok=True)
    with open(ruta_json, "w", encoding="utf-8") as f:
        json.dump(datos, f, indent=4, ensure_ascii=False)
    logger.info("Archivo de reglas guardado en '%s'.", ruta_json)


def procesar_datos(
    ruta_excel: str,
    ruta_json: str,
    hojas_a_ignorar: list[str] | None = None,
) -> dict:
    """
    Procesa las hojas del Excel de reglas y devuelve un diccionario.
    Si falla, registra el error y retorna {} (nunca None).
    """
    if hojas_a_ignorar is None:
        hojas_a_ignorar = ["PERSONAL_CENS", "CORREOS", "Reglas"]

    if not os.path.exists(ruta_excel):
        logger.error(
            "El archivo de reglas '%s' no existe. "
            "Verifique la ruta EXCEL_FILTROS en config.py.",
            ruta_excel,
        )
        return {}

    try:
        excel_data = leer_excel(ruta_excel)
    except Exception as e:
        logger.error("No se pudo leer el archivo de reglas '%s': %s", ruta_excel, e)
        return {}

    if not excel_data:
        logger.error("El archivo de reglas '%s' está vacío.", ruta_excel)
        return {}

    # Cargar JSON existente para comparar
    reglas_json: dict = {}
    if os.path.exists(ruta_json):
        try:
            with open(ruta_json, "r", encoding="utf-8") as f:
                reglas_json = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            logger.warning(
                "El archivo JSON de reglas '%s' está corrupto o ilegible: %s. "
                "Se regenerará desde el Excel.",
                ruta_json, e,
            )

    reglas: dict = {}
    try:
        for nombre_hoja, df in excel_data.items():
            if nombre_hoja in hojas_a_ignorar:
                continue
            df = df.reset_index(drop=True)
            for _, row in df.iterrows():
                clave = str(row.iloc[0]).strip() if row.iloc[0] is not None else ""
                valor = str(row.iloc[1]).strip() if row.iloc[1] is not None else ""
                if nombre_hoja not in reglas:
                    reglas[nombre_hoja] = {}
                reglas[nombre_hoja][clave] = valor

        if reglas != reglas_json:
            logger.info("El Excel de reglas fue modificado. Actualizando JSON.")
        guardar_json(reglas, ruta_json)
        return reglas

    except Exception as e:
        logger.error("Error al procesar las hojas del Excel de reglas: %s", e)
        # Retornar lo que se haya procesado hasta ahora
        return reglas


def separar_tres_hojas(
    ruta_excel: str, ruta_json: str
) -> tuple[dict, dict, dict]:
    """
    Extrae las tres reglas principales del Excel de configuración.
    Si alguna no existe, retorna {} para esa regla y registra advertencia.
    """
    json_reglas = procesar_datos(ruta_excel, ruta_json)

    if not json_reglas:
        logger.error(
            "No se pudieron cargar las reglas desde '%s'. "
            "Los filtros no se aplicarán correctamente.",
            ruta_excel,
        )
        return {}, {}, {}

    tipo_doc = json_reglas.get("Tipo doc.", {})
    cuentas_historial = json_reglas.get("Cuentas historial", {})
    otes = json_reglas.get("OTES(2C,2E,2K)", {})

    if not tipo_doc:
        logger.warning("Hoja 'Tipo doc.' no encontrada o vacía en el Excel de reglas.")
    if not cuentas_historial:
        logger.warning("Hoja 'Cuentas historial' no encontrada o vacía en el Excel de reglas.")
    if not otes:
        logger.warning("Hoja 'OTES(2C,2E,2K)' no encontrada o vacía en el Excel de reglas.")

    return tipo_doc, cuentas_historial, otes


# ─────────────────────────────────────────────────────────────────────────────
# Extracción y transformación del DataFrame principal
# ─────────────────────────────────────────────────────────────────────────────

def extraer_dataframe(
    ruta_excel: str,
    columnas: list[str] | None = None,
) -> pd.DataFrame:
    """
    Lee el Excel de salida y extrae solo las columnas necesarias.
    Retorna DataFrame vacío si falla.
    """
    if columnas is None:
        columnas = ["A", "C", "D", "E", "F", "K", "P", "V"]

    try:
        df = pd.read_excel(ruta_excel)
    except Exception as e:
        logger.error("No se pudo leer el Excel de salida '%s': %s", ruta_excel, e)
        return pd.DataFrame()

    indices = {letter: idx for idx, letter in enumerate("ABCDEFGHIJKLMNOPQRSTUVWXYZ")}
    indices_col = [indices[i] for i in columnas if i in indices]
    nombre_col = [df.columns[x] for x in indices_col if x < len(df.columns)]

    if not nombre_col:
        logger.error(
            "Ninguna de las columnas %s existe en el Excel '%s'.",
            columnas, ruta_excel,
        )
        return pd.DataFrame()

    df_seleccionado = df.loc[:, nombre_col].copy()
    for col in df_seleccionado.columns:
        if df_seleccionado[col].dtype == object:
            df_seleccionado[col] = df_seleccionado[col].astype(str).str.strip()

    return df_seleccionado


# ─────────────────────────────────────────────────────────────────────────────
# Filtros individuales
# ─────────────────────────────────────────────────────────────────────────────

def filtro_tipo_doc(df: pd.DataFrame, dict_tipo_doc: dict) -> pd.DataFrame:
    if "Responsable" not in df.columns or "Tipo doc" not in df.columns:
        logger.error("Faltan columnas 'Responsable' o 'Tipo doc' para aplicar filtro_tipo_doc.")
        return df
    if not dict_tipo_doc:
        return df

    df["Responsable"] = df["Responsable"].astype(object).str.strip()
    df["Tipo doc"] = df["Tipo doc"].str.strip()

    for key, valor in dict_tipo_doc.items():
        mask = df["Tipo doc"].str.contains(key, na=False)
        df.loc[mask, "Responsable"] = valor

    return df


def filtro_otes(df: pd.DataFrame, dict_otes: dict, tipo_doc: list[str] | None = None) -> pd.DataFrame:
    if tipo_doc is None:
        tipo_doc = OTES
    if not dict_otes:
        return df

    df["Cuenta"] = df["Cuenta"].astype(str).str.strip()
    df["Responsable"] = df["Responsable"].astype(object).str.strip()
    df.columns = df.columns.str.strip()

    mask_tipo = df["Tipo doc"].isin(tipo_doc)
    for key, valor in dict_otes.items():
        mask = mask_tipo & df["Cuenta"].str.contains(key, na=False)
        df.loc[mask, "Responsable"] = valor

    # Propagación de responsables dentro del mismo número de documento
    df_filtrado = df[mask_tipo].copy()
    responsables = df_filtrado["Responsable"].dropna().unique()

    for responsable in responsables:
        if not responsable or responsable == "nan":
            continue
        docs = df_filtrado[df_filtrado["Responsable"] == responsable]["Número documento"].unique()
        pares_mask = mask_tipo & df["Número documento"].isin(docs)
        df.loc[pares_mask, "Responsable"] = responsable

    return df


def filtro_historial(df: pd.DataFrame, dict_historial: dict) -> pd.DataFrame:
    if not dict_historial:
        return df

    df["Cuenta"] = df["Cuenta"].astype(str).str.strip()
    df["Responsable"] = df["Responsable"].astype(object).str.strip()

    for key, valor in dict_historial.items():
        mask = df["Cuenta"].str.contains(key, na=False)
        df.loc[mask, "Responsable"] = valor

    return df


# ─────────────────────────────────────────────────────────────────────────────
# Guardado en Excel
# ─────────────────────────────────────────────────────────────────────────────

def guardar_en_excel(df: pd.DataFrame, ruta_salida: str) -> bool:
    """
    Actualiza la columna 'Responsable' en el Excel existente.
    Retorna True si guardó correctamente.
    """
    try:
        book = load_workbook(ruta_salida)
        sheet = book.active
        if sheet is None:
            logger.error("No se pudo acceder a la hoja activa del libro '%s'.", ruta_salida)
            return False

        filas_actualizadas = 0
        filas_sin_responsable = 0
        for idx, row in df.iterrows():
            excel_row = idx + 2
            if "Responsable" in row:
                sheet[f"A{excel_row}"] = row["Responsable"]
                filas_actualizadas += 1
            else:
                filas_sin_responsable += 1

        if filas_sin_responsable > 0:
            logger.warning(
                "%d filas no tenían columna 'Responsable' y no fueron actualizadas.",
                filas_sin_responsable,
            )

        book.save(ruta_salida)
        logger.info(
            "Excel actualizado: %d filas con 'Responsable' en '%s'.",
            filas_actualizadas, ruta_salida,
        )
        return True

    except PermissionError:
        logger.error(
            "Permiso denegado al guardar '%s'. "
            "Cierre el archivo si lo tiene abierto en Excel.",
            ruta_salida,
        )
        return False
    except Exception as e:
        logger.error("Error al guardar en Excel '%s': %s", ruta_salida, e)
        return False


# ─────────────────────────────────────────────────────────────────────────────
# Función principal
# ─────────────────────────────────────────────────────────────────────────────

def aplicar_todos_los_filtros(
    ruta_excel: str, ruta_json: str, ruta_salida: str
) -> str | None:
    """
    Aplica todos los filtros simples al Excel de salida.
    Retorna la ruta del archivo actualizado, o None si hubo un error crítico.
    """
    df = extraer_dataframe(ruta_salida)
    if df.empty:
        logger.error(
            "No se pudo extraer datos de '%s'. Los filtros simples no se aplicarán.",
            ruta_salida,
        )
        return None

    tipo_doc, cuentas_historial, otes = separar_tres_hojas(ruta_excel, ruta_json)

    # Los filtros se aplican aunque algún diccionario esté vacío
    # (registran advertencia internamente pero no abortan)
    df = filtro_tipo_doc(df, tipo_doc)
    df = filtro_historial(df, cuentas_historial)
    df = filtro_otes(df, otes)

    if not guardar_en_excel(df, ruta_salida):
        logger.error("No se pudo guardar el resultado de los filtros simples.")
        return None

    return ruta_salida
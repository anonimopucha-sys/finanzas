import pandas as pd


def obtener_hoja_excel(ruta_excel: str, candidatos: tuple[str, ...], nombre_default: str | None = None) -> str:
    """Devuelve una hoja existente del Excel, usando candidatos preferidos o la primera hoja disponible."""
    try:
        excel_file = pd.ExcelFile(ruta_excel)
        hojas = list(excel_file.sheet_names)
    except Exception as exc:
        raise ValueError(f"No se pudo leer el archivo Excel '{ruta_excel}': {exc}") from exc

    for nombre in candidatos:
        if nombre in hojas:
            return nombre

    if nombre_default and nombre_default in hojas:
        return nombre_default

    if hojas:
        return hojas[0]

    raise ValueError(f"El archivo Excel '{ruta_excel}' no contiene hojas válidas para procesamiento.")

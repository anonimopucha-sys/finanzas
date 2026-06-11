"""
Módulo de envío de correos para CE1121.

Soporta dos backends:
  - Outlook (Windows, vía win32com) — predeterminado en Windows
  - SMTP    (cross-platform, vía smtplib + EmailMessage)

Parámetros clave de procesar_envio():
  dry_run     : True → solo loguea, no envía. Útil para pruebas.
  use_outlook : True/False para forzar backend. None = auto (Outlook en Windows).
  smtp_config : dict con 'server', 'port', 'use_tls', 'username', 'password', 'from_addr'.
  delay       : segundos de espera entre correos (evita bloqueos en servidores SMTP).
"""

import mimetypes
import os
import platform
import smtplib
import time
from email.message import EmailMessage
from typing import Optional, Dict

import pandas as pd

from .logs_config import configurar_logger

logger = configurar_logger()


# ─────────────────────────────────────────────────────────────────────────────
# Generación de Excel por responsable
# ─────────────────────────────────────────────────────────────────────────────

def generar_excel(df: pd.DataFrame, responsable: str, ruta: str) -> str:
    """
    Genera un Excel filtrado por responsable en <ruta>/<responsable>.xlsx.
    Retorna la ruta del archivo generado.
    """
    df_responsable = df[df["Responsable"] == responsable]
    carpeta = os.path.join(ruta, responsable)
    os.makedirs(carpeta, exist_ok=True)
    archivo = os.path.join(carpeta, f"{responsable}.xlsx")
    df_responsable.to_excel(archivo, index=False)
    logger.info("Archivo generado para '%s': %s", responsable, archivo)
    return archivo


# ─────────────────────────────────────────────────────────────────────────────
# Backends de envío
# ─────────────────────────────────────────────────────────────────────────────

def _send_via_outlook(
    destinatario: str, asunto: str, mensaje: str, archivo_adjunto: str
) -> bool:
    """Envía un correo usando Outlook (solo Windows)."""
    try:
        import win32com.client as win32  # import local: no falla en Linux
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = asunto
        mail.HTMLBody = mensaje
        mail.Attachments.Add(archivo_adjunto)
        mail.Send()
        logger.info("Correo enviado vía Outlook a '%s'.", destinatario)
        return True
    except ImportError:
        logger.warning("win32com no disponible. Outlook no puede usarse en este sistema.")
        return False
    except Exception:
        logger.exception("Fallo al enviar vía Outlook a '%s'.", destinatario)
        return False


def _send_via_smtp(
    destinatario: str,
    asunto: str,
    mensaje: str,
    archivo_adjunto: str,
    smtp_config: Optional[Dict] = None,
) -> bool:
    """Envía un correo usando SMTP con adjunto MIME."""
    smtp_config = smtp_config or {}
    servidor  = smtp_config.get("server", "localhost")
    puerto    = int(smtp_config.get("port", 25))
    remitente = smtp_config.get("from_addr", f"noreply@{platform.node()}")
    use_tls   = smtp_config.get("use_tls", False)
    username  = smtp_config.get("username")
    password  = smtp_config.get("password")

    try:
        msg = EmailMessage()
        msg["From"]    = remitente
        msg["To"]      = destinatario
        msg["Subject"] = asunto
        msg.set_content(mensaje, subtype="html")

        # Adjuntar archivo si existe
        if archivo_adjunto and os.path.exists(archivo_adjunto):
            ctype, _ = mimetypes.guess_type(archivo_adjunto)
            if ctype is None:
                ctype = "application/octet-stream"
            maintype, subtype = ctype.split("/", 1)
            with open(archivo_adjunto, "rb") as f:
                data = f.read()
            msg.add_attachment(
                data,
                maintype=maintype,
                subtype=subtype,
                filename=os.path.basename(archivo_adjunto),
            )
        else:
            logger.warning(
                "Archivo adjunto '%s' no encontrado. Se enviará sin adjunto.", archivo_adjunto
            )

        with smtplib.SMTP(servidor, puerto, timeout=30) as s:
            s.ehlo()
            if use_tls:
                s.starttls()
                s.ehlo()
            if username and password:
                s.login(username, password)
            s.send_message(msg)

        logger.info(
            "Correo enviado vía SMTP a '%s' (servidor=%s:%s).", destinatario, servidor, puerto
        )
        return True

    except smtplib.SMTPAuthenticationError:
        logger.error(
            "Error de autenticación SMTP para '%s'. Verifique usuario/contraseña en smtp_config.",
            destinatario,
        )
        return False
    except smtplib.SMTPConnectError:
        logger.error(
            "No se pudo conectar al servidor SMTP %s:%s. Verifique la configuración.",
            servidor, puerto,
        )
        return False
    except Exception:
        logger.exception("Fallo inesperado al enviar vía SMTP a '%s'.", destinatario)
        return False


# ─────────────────────────────────────────────────────────────────────────────
# Función pública de envío
# ─────────────────────────────────────────────────────────────────────────────

def enviar_correo(
    destinatario: str,
    asunto: str,
    mensaje: str,
    archivo_adjunto: str,
    *,
    dry_run: bool = False,
    use_outlook: Optional[bool] = None,
    smtp_config: Optional[Dict] = None,
) -> bool:
    """
    Envía un correo con adjunto.

    - dry_run=True: solo loguea, útil para pruebas sin servidor real.
    - use_outlook=None: usa Outlook en Windows, SMTP en otros sistemas.
    - smtp_config: configuración del servidor SMTP (ver DEPLOYMENT_SMTP.md).
    """
    if dry_run:
        logger.info(
            "[dry_run] Simularía envío a '%s' | Asunto: '%s' | Adjunto: '%s'",
            destinatario, asunto, archivo_adjunto,
        )
        return True

    prefer_outlook = use_outlook if use_outlook is not None else (platform.system() == "Windows")

    if prefer_outlook:
        ok = _send_via_outlook(destinatario, asunto, mensaje, archivo_adjunto)
        if ok:
            return True
        logger.warning("Outlook falló. Intentando fallback SMTP.")

    return _send_via_smtp(destinatario, asunto, mensaje, archivo_adjunto, smtp_config=smtp_config)


# ─────────────────────────────────────────────────────────────────────────────
# Lectura del archivo de reglas (mapeo responsable → correo)
# ─────────────────────────────────────────────────────────────────────────────

def analizar_excel(
    ruta_archivo: str, hoja: Optional[str | int] = 0
) -> Optional[Dict[str, str]]:
    """
    Lee el archivo de reglas y retorna un dict {nombre_responsable: correo}.
    Detecta automáticamente si el archivo tiene o no fila de encabezado.
    Retorna None si no se encuentran correos válidos.
    """
    try:
        # Intentar sin encabezado primero
        df_none = pd.read_excel(ruta_archivo, sheet_name=hoja, header=None)
        vals_none = df_none.iloc[:, 1].astype(str).str.strip()
        if vals_none.str.contains("@").any():
            keys = df_none.iloc[:, 0].astype(str).str.strip()
            return dict(zip(keys, vals_none))

        # Con encabezado
        df_hdr = pd.read_excel(ruta_archivo, sheet_name=hoja, header=0)

        if df_hdr.shape[1] >= 2 and df_hdr.iloc[:, 1].astype(str).str.contains("@").any():
            keys = df_hdr.iloc[:, 0].astype(str).str.strip()
            vals = df_hdr.iloc[:, 1].astype(str).str.strip()
            return dict(zip(keys, vals))

        # Buscar la columna con más correos
        best_col, best_count = None, 0
        for col in df_hdr.columns:
            try:
                cnt = int(df_hdr[col].astype(str).str.contains("@").sum())
                if cnt > best_count:
                    best_count = cnt
                    best_col = col
            except Exception:
                continue

        if best_col is not None and best_count > 0:
            keys = df_hdr.iloc[:, 0].astype(str).str.strip()
            vals = df_hdr[best_col].astype(str).str.strip()
            return dict(zip(keys, vals))

        logger.warning(
            "No se detectaron correos en '%s'. Verifique que la segunda columna "
            "tenga direcciones de correo con '@'.",
            ruta_archivo,
        )
        return None

    except FileNotFoundError:
        logger.error("El archivo de reglas '%s' no existe.", ruta_archivo)
        return None
    except Exception:
        logger.exception("Error al procesar el archivo de reglas '%s'.", ruta_archivo)
        return None


# ─────────────────────────────────────────────────────────────────────────────
# Proceso principal de envío
# ─────────────────────────────────────────────────────────────────────────────

def procesar_envio(
    ruta_excel: str,
    reglas: str,
    hoja: Optional[str | int],
    ruta_base: str,
    *,
    dry_run: bool = False,
    smtp_config: Optional[Dict] = None,
    use_outlook: Optional[bool] = None,
    delay: float = 1.0,
) -> None:
    """
    Procesa el Excel principal, genera un Excel por responsable y envía correos.

    Parámetros:
      ruta_excel  : Excel con columna 'Responsable'.
      reglas      : Excel con mapeo nombre → correo.
      hoja        : Hoja del Excel de reglas (nombre o índice, default 0).
      ruta_base   : Carpeta base donde se guardan los Excel individuales.
      dry_run     : Si True, simula el envío sin enviar realmente.
      smtp_config : Configuración SMTP (ver DEPLOYMENT_SMTP.md).
      use_outlook : Forzar backend (True=Outlook, False=SMTP, None=auto).
      delay       : Segundos entre correos.
    """
    try:
        responsables_dict = analizar_excel(reglas, hoja)
        if not responsables_dict:
            logger.error(
                "No se pudo cargar el diccionario de responsables desde '%s'. "
                "Verifique el archivo de reglas.",
                reglas,
            )
            return

        try:
            df = pd.read_excel(ruta_excel)
        except Exception as e:
            logger.error("No se pudo leer el Excel de datos '%s': %s", ruta_excel, e)
            return

        if "Responsable" not in df.columns:
            logger.error(
                "La columna 'Responsable' no existe en '%s'. "
                "Ejecute primero el proceso de filtros.",
                ruta_excel,
            )
            return

        responsables = df["Responsable"].dropna().unique()
        logger.info("Procesando %d responsable(s) únicos.", len(responsables))
        enviados, fallidos, omitidos = 0, 0, 0

        for responsable in responsables:
            correo = responsables_dict.get(str(responsable))
            if not correo or "@" not in str(correo):
                logger.warning(
                    "No se encontró correo válido para '%s'. Se omite.", responsable
                )
                omitidos += 1
                continue

            asunto = f"Registros pendientes para {responsable}"
            archivo_adjunto = generar_excel(df, str(responsable), ruta_base)
            mensaje = f"""
<html>
<body>
    <h2>Estimado {responsable},</h2>
    <p>Adjunto encontrarás el archivo con los registros asignados a ti.</p>
    <p>Saludos,<br>Proceso automatizado CE1121</p>
</body>
</html>
"""
            ok = enviar_correo(
                correo, asunto, mensaje, archivo_adjunto,
                dry_run=dry_run,
                use_outlook=use_outlook,
                smtp_config=smtp_config,
            )
            if ok:
                enviados += 1
            else:
                fallidos += 1
                logger.error("No se pudo enviar el correo a '%s' (%s).", responsable, correo)

            time.sleep(delay)

        logger.info(
            "Envío completado: %d enviados, %d fallidos, %d omitidos (sin correo).",
            enviados, fallidos, omitidos,
        )

    except Exception:
        logger.exception("Error inesperado en procesar_envio.")
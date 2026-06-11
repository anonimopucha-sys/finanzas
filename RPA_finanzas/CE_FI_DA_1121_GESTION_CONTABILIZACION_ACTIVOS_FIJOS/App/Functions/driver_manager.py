import os
import platform
import re
import shutil
import stat
import subprocess
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

import urllib3
from selenium import webdriver
from selenium.common.exceptions import SessionNotCreatedException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Permitir a webdriver-manager resolver descargas en entornos con SSL restringido.
# Esto corresponde a: os.environ['WDM_SSL_VERIFY'] = '0'
os.environ.setdefault("WDM_SSL_VERIFY", "0")
if os.environ.get("WDM_SSL_VERIFY") == "0":
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from .logs_config import configurar_logger

logger = configurar_logger("CE1121.driver")
CHROME_DOWNLOAD_URL = "https://www.google.com/chrome/"
CHROMEDRIVER_DOWNLOAD_URL = "https://chromedriver.chromium.org/downloads"


class DriverException(Exception):
    """Base exception for driver creation and validation errors."""


class DriverFatalError(DriverException):
    """Non-recoverable error that should abort processing."""


class DriverRetryableError(DriverException):
    """Transient Selenium error that may be retried."""


class EnvironmentValidationError(DriverFatalError):
    """Environment is not ready for browser automation."""


def _parse_version(output: str) -> str | None:
    match = re.search(r"(\d+\.\d+\.\d+\.\d+)", output)
    return match.group(1) if match else None


def _run_process(command: list[str]) -> str | None:
    try:
        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            check=True,
            timeout=10,
        )
        return result.stdout.strip()
    except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired) as exc:
        logger.debug("Error ejecutando comando %s: %s", command, exc)
        return None


def _is_windows() -> bool:
    return platform.system() == "Windows"


def _ensure_executable(path: str) -> None:
    if not os.path.exists(path):
        return
    if _is_windows():
        return
    try:
        current_mode = os.stat(path).st_mode
        if not (current_mode & stat.S_IXUSR):
            os.chmod(path, current_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)
            logger.debug("Se agregaron permisos de ejecución a ChromeDriver: %s", path)
    except OSError as exc:
        logger.warning("No se pudieron ajustar permisos de ejecución para ChromeDriver %s: %s", path, exc)


def _local_chromedriver_path() -> str | None:
    project_root = Path(__file__).resolve().parents[2]
    candidates: list[str] = []
    env_driver = os.environ.get("CHROMEDRIVER_PATH") or os.environ.get("CHROME_DRIVER_PATH")
    if env_driver:
        candidates.append(env_driver)

    candidates.extend([
        r"C:\RPA_finanzas\drivers\chromedriver.exe",
        r"C:\RPA_finanzas\chromedriver-win64\chromedriver.exe",
        r"C:\RPA\chromedriver-win64\chromedriver.exe",
        str(project_root / "drivers" / "chromedriver.exe"),
        str(project_root / "chromedriver-win64" / "chromedriver.exe"),
        str(project_root / "chromedriver.exe"),
    ])

    user_cache = Path.home() / ".wdm" / "drivers" / "chromedriver"
    if user_cache.exists():
        for candidate in user_cache.rglob("chromedriver.exe"):
            candidates.append(str(candidate))

    for candidate in project_root.rglob("chromedriver*"):
        if candidate.is_file():
            candidates.append(str(candidate))

    valid_candidates: list[tuple[str, str]] = []
    for candidate in candidates:
        if candidate and os.path.isfile(candidate):
            if not candidate.lower().endswith(".exe"):
                logger.debug("Ignorando candidato ChromeDriver sin extensión .exe: %s", candidate)
                continue

            try:
                file_size = os.path.getsize(candidate)
            except OSError:
                logger.debug("No se pudo obtener tamaño del archivo ChromeDriver: %s", candidate)
                continue

            if file_size < 1024:
                logger.debug("Ignorando candidato ChromeDriver inválido o vacío: %s (tamaño %s bytes)", candidate, file_size)
                continue

            version = obtener_version_chromedriver(candidate)
            if not version:
                logger.debug("Ignorando candidato ChromeDriver sin versión válida: %s", candidate)
                continue

            valid_candidates.append((version, os.path.abspath(candidate)))

    if not valid_candidates:
        return None

    valid_candidates.sort(key=lambda item: tuple(int(part) for part in item[0].split(".") if part.isdigit()), reverse=True)
    chosen = valid_candidates[0][1]
    logger.debug("Seleccionado ChromeDriver local: %s (versión %s)", chosen, valid_candidates[0][0])
    return chosen


def _find_cached_chromedriver(chrome_version: str | None = None) -> str | None:
    user_cache = Path.home() / ".wdm" / "drivers" / "chromedriver"
    if not user_cache.exists():
        return None

    candidates: list[tuple[str, str]] = []
    for candidate in user_cache.rglob("chromedriver.exe"):
        if candidate.is_file():
            try:
                file_size = candidate.stat().st_size
            except OSError:
                continue
            if file_size < 1024:
                continue
            version = obtener_version_chromedriver(str(candidate))
            if version:
                candidates.append((version, str(candidate)))

    if not candidates:
        return None

    if chrome_version:
        chrome_major = chrome_version.split(".")[0]
        filtered = [item for item in candidates if item[0].startswith(f"{chrome_major}.")]
        if filtered:
            candidates = filtered

    candidates.sort(key=lambda item: tuple(int(part) for part in item[0].split(".") if part.isdigit()), reverse=True)
    return candidates[0][1]


def _install_chromedriver(chrome_version: str | None = None) -> str:
    try:
        if chrome_version:
            manager = ChromeDriverManager()
        else:
            manager = ChromeDriverManager()

        driver_path = manager.install()
        _ensure_executable(driver_path)
        logger.info("ChromeDriver resuelto con webdriver-manager: %s", driver_path)
        return driver_path
    except Exception as exc:
        logger.exception("Error instalando ChromeDriver con webdriver-manager")
        cached_path = _find_cached_chromedriver(chrome_version)
        if cached_path:
            logger.warning(
                "Se encontró ChromeDriver cacheado en %s tras fallo de descarga; usando ese driver.",
                cached_path,
            )
            return cached_path
        raise DriverFatalError(
            "No se pudo resolver ChromeDriver automáticamente. "
            "Verifique la conexión a Internet o proporcione CHROMEDRIVER_PATH. "
            f"Si no tiene ChromeDriver, descárguelo desde {CHROMEDRIVER_DOWNLOAD_URL} "
            "y use la versión que coincida con su Chrome instalado."
        ) from exc


def _find_chrome_executable() -> str | None:
    candidates: list[str] = []
    env_chrome = os.environ.get("CHROME_BIN") or os.environ.get("GOOGLE_CHROME_BIN")
    if env_chrome:
        candidates.append(env_chrome)

    candidates.extend([
        shutil.which("chrome"),
        shutil.which("google-chrome"),
        shutil.which("chromium-browser"),
        shutil.which("chromium"),
        os.path.join(os.environ.get("PROGRAMFILES", "C:\\Program Files"), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", "C:\\Program Files (x86)"), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
    ])

    for candidate in candidates:
        if candidate and os.path.isfile(candidate):
            return os.path.abspath(candidate)

    logger.debug("No se encontró Google Chrome. Candidatos evaluados: %s", candidates)
    return None


def _windows_chrome_file_version(chrome_path: str) -> str | None:
    try:
        escaped_path = chrome_path.replace("'", "''")
        command = [
            "powershell",
            "-NoProfile",
            "-Command",
            f"(Get-Item -LiteralPath '{escaped_path}').VersionInfo.ProductVersion",
        ]
        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            shell=False,
            timeout=15,
        )
        salida = result.stdout.strip() or result.stderr.strip()
        logger.info("Versión detectada de Chrome vía PowerShell: %s", salida)
        return _parse_version(salida)
    except Exception as exc:
        logger.debug("No se pudo obtener versión de Chrome vía PowerShell: %s", exc)
        return None


def obtener_version_chrome() -> str | None:
    chrome_path = _find_chrome_executable()

    if not chrome_path:
        return None

    for command in ([chrome_path, "--version"], [chrome_path, "--product-version"]):
        try:
            result = subprocess.run(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                shell=False,
                timeout=10,
            )
            salida = result.stdout.strip() or result.stderr.strip()
            logger.info("Versión detectada de Chrome (comando %s): %s", command, salida)
            version = _parse_version(salida)
            if version:
                return version
        except Exception as e:
            logger.warning("Error ejecutando Chrome para versión con comando %s: %s", command, e)

    if _is_windows():
        version = _windows_chrome_file_version(chrome_path)
        if version:
            return version

    logger.warning("No se pudo determinar la versión de Chrome a partir de la salida del binario.")
    return None


def obtener_version_chromedriver(driver_path: str) -> str | None:
    salida = _run_process([driver_path, "--version"])
    if salida:
        return _parse_version(salida)
    return None


def _crear_opciones(download_dir: str | None = None, headless: bool = False) -> Options:
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-background-networking")
    options.add_argument("--disable-default-apps")
    options.add_argument("--remote-allow-origins=*")

    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")

    if download_dir:
        options.add_experimental_option("prefs", {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        })

    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    return options


def _es_incompatibilidad(error: WebDriverException) -> bool:
    mensaje = str(error).lower()
    return (
        "only supports chrome version" in mensaje
        or "current browser version is" in mensaje
        or "session not created" in mensaje
        or "driver executable" in mensaje
        or "executable may have wrong architecture" in mensaje
    )


def validar_entorno() -> bool:
    chrome_path = _find_chrome_executable()

    if not chrome_path:
        raise EnvironmentValidationError(
            "No se encontró Google Chrome instalado."
        )

    chrome_version = obtener_version_chrome()

    if not chrome_version:
        raise EnvironmentValidationError(
            "No se pudo determinar la versión de Chrome."
        )

    driver_path = _local_chromedriver_path()
    if driver_path is None:
        logger.warning(
            "No se encontró ChromeDriver local; intentando resolver con webdriver-manager. "
            f"Si no tiene ChromeDriver, descárguelo desde {CHROMEDRIVER_DOWNLOAD_URL} "
            "con la versión compatible con su Chrome instalado."
        )
        driver_path = _install_chromedriver()

    _ensure_executable(driver_path)

    chromedriver_version = obtener_version_chromedriver(driver_path)

    logger.info(
        "Validación OK → Chrome=%s | Driver=%s",
        chrome_version,
        chromedriver_version
    )

    return True



def _crear_driver_con_path(driver_path: str, download_dir: str | None, headless: bool) -> webdriver.Chrome:
    if not driver_path or not os.path.isfile(driver_path):
        raise DriverFatalError(
            f"No se encontró ChromeDriver en la ruta: {driver_path}"
        )

    _ensure_executable(driver_path)

    service = Service(driver_path)
    options = _crear_opciones(download_dir=download_dir, headless=headless)

    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(120)
    logger.debug("Driver iniciado correctamente con ChromeDriver %s", driver_path)
    return driver


def crear_driver(download_dir: str | None = None, headless: bool = False) -> webdriver.Chrome:
    driver_path = _local_chromedriver_path()
    if driver_path is None:
        logger.warning("No se encontró ChromeDriver local; intentando resolver con webdriver-manager.")
        driver_path = _install_chromedriver()
    else:
        current_driver_version = obtener_version_chromedriver(driver_path)
        logger.debug(
            "ChromeDriver local seleccionado: %s (versión %s)",
            driver_path,
            current_driver_version or "desconocida",
        )

    try:
        return _crear_driver_con_path(driver_path, download_dir, headless)

    except (SessionNotCreatedException, WebDriverException) as exc:
        chrome_version = obtener_version_chrome()
        chromedriver_version = obtener_version_chromedriver(driver_path)
        compat_message = (
            f"Chrome={chrome_version or 'desconocido'} "
            f"ChromeDriver={chromedriver_version or 'desconocido'}"
        )

        if _es_incompatibilidad(exc):
            logger.warning(
                "Driver incompatible o versión incorrecta: %s. Intentando resolver con webdriver-manager.",
                compat_message,
            )
            logger.warning(
                "Ruta de ChromeDriver en uso antes del reintento: %s (versión %s)",
                driver_path,
                chromedriver_version or "desconocida",
            )
            try:
                driver_path = _install_chromedriver(chrome_version)
                chromedriver_version = obtener_version_chromedriver(driver_path)
                logger.debug(
                    "Reintentando con ChromeDriver %s",
                    chromedriver_version,
                )
                logger.debug(
                    "ChromeDriver resuelto después de incompatibilidad: %s",
                    driver_path,
                )
                return _crear_driver_con_path(driver_path, download_dir, headless)
            except Exception as install_exc:
                logger.exception("No se pudo reintentar con webdriver-manager después de incompatibilidad")
                raise DriverFatalError(
                    "No se pudo iniciar Chrome con una versión compatible de ChromeDriver. "
                    f"{compat_message}. "
                    f"{CHROMEDRIVER_DOWNLOAD_URL}"
                ) from install_exc

        if isinstance(exc, SessionNotCreatedException) or _es_incompatibilidad(exc):
            logger.error("Incompatibilidad detectada: %s", compat_message)
            raise DriverFatalError(
                "Incompatibilidad entre Chrome y ChromeDriver detectada. "
                f"{compat_message}. "
                f"{CHROMEDRIVER_DOWNLOAD_URL}"
            ) from exc

        logger.exception("Error transitorio creando el Chrome driver")
        raise DriverRetryableError(
            "Error transitorio al iniciar el driver de Selenium."
        ) from exc

    except Exception as exc:
        logger.exception("Error inesperado creando el Chrome driver")
        raise DriverRetryableError(
            "Error inesperado al crear el driver de Selenium."
        ) from exc


def cerrar_driver(driver: webdriver.Chrome | None) -> None:
    if not driver:
        return
    try:
        driver.quit()
    except WebDriverException as exc:
        logger.warning("Error cerrando el driver de Selenium: %s", exc)


@contextmanager
def driver_context(download_dir: str | None = None, headless: bool = False) -> Iterator[webdriver.Chrome]:
    driver = crear_driver(download_dir=download_dir, headless=headless)
    try:
        yield driver
    finally:
        cerrar_driver(driver)

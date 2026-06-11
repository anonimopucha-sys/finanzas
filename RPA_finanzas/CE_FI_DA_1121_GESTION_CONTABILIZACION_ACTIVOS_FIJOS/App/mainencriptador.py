import sys
import os
from cryptography.fernet import Fernet
from cryptography.fernet import InvalidToken
from dotenv import load_dotenv

# Resolver ruta base tanto para .py como para exe
BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(__file__)

# Cargar variables de entorno
dotenv_path = os.path.join(BASE_DIR, "config", ".env")
load_dotenv(dotenv_path, override=True)


def desencriptarCreden(usuario_encriptado: str, contrasena_encriptada: str) -> tuple[str, str]:
    """
    Desencripta usuario y contraseña usando Fernet.
    Retorna (usuario, contraseña) en texto plano.
    """

    llave = os.getenv("LLAVE_MAESTRA")
    if not llave:
        raise ValueError("No se pudo cargar LLAVE_MAESTRA desde config/.env")

    try:
        key = llave.encode("utf-8")
        fernet = Fernet(key)

        user = fernet.decrypt(usuario_encriptado.encode("utf-8")).decode()
        password = fernet.decrypt(contrasena_encriptada.encode("utf-8")).decode()

        return user, password

    except InvalidToken:
        raise RuntimeError(
            "Las credenciales no coinciden con la llave de encriptación.\n"
            "Por favor vuelva a guardarlas."
        )


def obtener_credenciales() -> tuple[str, str]:
    """
    Lee USER y PASSWORD del .env y los desencripta.
    """

    user_enc = os.getenv("USER")
    pass_enc = os.getenv("PASSWORD")

    if not user_enc or not pass_enc:
        raise ValueError("USER o PASSWORD no existen en el archivo .env")

    return desencriptarCreden(user_enc, pass_enc)


# 🛑 MUY IMPORTANTE:
# NO ejecutar absolutamente nada al importar este módulo

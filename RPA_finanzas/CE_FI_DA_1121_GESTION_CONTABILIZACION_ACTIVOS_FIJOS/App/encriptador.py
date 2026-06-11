import os
from cryptography.fernet import Fernet
from dotenv import load_dotenv


# BASE FIJA EN /App
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
DOTENV_PATH = os.path.join(CONFIG_DIR, ".env")

# Asegurar carpeta App/config
os.makedirs(CONFIG_DIR, exist_ok=True)

def guardar_credenciales(user: str, password: str) -> str:
    """
    Borra cualquier .env previo y guarda nuevas credenciales encriptadas.
    Retorna la ruta del archivo .env creado.
    """


    # Cargar env (por si hubiera algo previo)
    load_dotenv(DOTENV_PATH, override=True)

    # Obtener o generar llave
    llave = os.getenv("LLAVE_MAESTRA")
    if not llave:
        llave = Fernet.generate_key().decode("utf-8")

    fernet = Fernet(llave.encode("utf-8"))

    user_enc = fernet.encrypt(user.encode()).decode()
    pass_enc = fernet.encrypt(password.encode()).decode()

    with open(DOTENV_PATH, "w", encoding="utf-8") as f:
        f.write(f"USER={user_enc}\n")
        f.write(f"PASSWORD={pass_enc}\n")
        f.write(f"LLAVE_MAESTRA={llave}\n")

    return DOTENV_PATH
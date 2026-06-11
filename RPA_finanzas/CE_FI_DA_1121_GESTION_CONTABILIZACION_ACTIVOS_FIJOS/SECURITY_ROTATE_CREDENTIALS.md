Instrucciones para rotar credenciales y proteger la llave maestra

1) Rotación inmediata (acción requerida):
- Asuma que las credenciales actuales están comprometidas porque la llave maestra (`LLAVE_MAESTRA`) estaba en el repositorio.
- Generar nuevas credenciales de JDE y anular las anteriores en el sistema de JDE.
- Re-encriptar las nuevas credenciales localmente usando el script de encriptación del repositorio (ej. `App/encriptador.py`) y guardar solo el valor en un lugar seguro (no en el repositorio).

2) Quitar la llave maestra del repositorio (ya aplicado):
- `App/config/.env` fue eliminado del índice git y se añadió a `.gitignore`.
- Verifique el historial de Git por exposiciones previas y coordine rotación de claves si se detecta acceso externo.

3) Buenas prácticas de manejo de secretos:
- La `LLAVE_MAESTRA` debe almacenarse en una variable de entorno del sistema o en un gestor de secretos (Azure Key Vault, AWS Secrets Manager, HashiCorp Vault, Windows Credential Manager).
- No subir `.env` al repositorio. Mantener un archivo `README_ENV.md` con el esquema de variables que deben definirse en cada entorno.

4) Pasos operativos para el desarrollador / operador:
- Local (Windows): exportar variable de entorno antes de ejecutar la RPA:

```powershell
$env:LLAVE_MAESTRA = '...'
$env:USER = 'gAAAAAB...'
$env:PASSWORD = 'gAAAAAB...'
python App/gui.py
```

- Linux/macOS:

```bash
export LLAVE_MAESTRA='...'
export USER='gAAAAAB...'
export PASSWORD='gAAAAAB...'
python3 App/gui.py
```

5) Recomendación a mediano plazo:
- Integrar almacenamiento de secrets en el pipeline (CI/CD) y usar variables de entorno del runner para despliegues programados.
- Añadir un script `scripts/rotate_credentials.sh` que automatice la generación, encriptación y subida segura de nuevos secretos a un vault.

6) Auditoría Git:
- Buscar en el historial con `git log --all -- App/config/.env` o usar `git-secrets`/`truffleHog` para detectar otras fugas.

7) Contacto y seguimiento:
- Informe a seguridad interna sobre la exposición y documente la rotación.

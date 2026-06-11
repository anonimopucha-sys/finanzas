# Manual de Instalación Fácil

Este documento explica cómo instalar y ejecutar el proyecto `CE_FI_DA_1121_GESTION_CONTABILIZACION_ACTIVOS_FIJOS` en Windows.

## 1. ¿Qué hace este proyecto?

Este proyecto abre una interfaz de usuario en Windows y ejecuta una automatización con Selenium para el proceso JDE.

## 2. Requisitos mínimos

- Windows 10 o Windows 11
- Python 3.11 o 3.12 de 64 bits
- Conexión a Internet (solo para descargar Python si no está instalado)
- Carpetas locales para archivos de entrada y salida
- Chrome compatible con el `chromedriver.exe` instalado

## 3. Preparar el equipo

### 3.1 Instalar Python

1. Descarga Python 3.11 o 3.12 desde https://www.python.org/downloads/windows/
2. Durante la instalación, marca la casilla "Add Python to PATH".
3. Comprueba la instalación con el Símbolo del sistema:

```bat
python --version
pip --version
```

### 3.2 Verificar que el proyecto esté disponible

El proyecto debe estar en la carpeta:

```bat
RPA\CE_FI_DA_1121_GESTION_CONTABILIZACION_ACTIVOS_FIJOS
```

Dentro de esta carpeta debe existir:

- `CE1121.bat`
- `requirements.txt`
- `App\gui.py` (interfaz Tkinter)
- `LIbs\` (o `libs` con las ruedas locales de Python)

## 4. Configuración inicial necesaria

### 4.1 Crear las carpetas locales requeridas

Crea estas carpetas en Windows:

- `C:\RPA\chromedriver-win64`
- `C:\CE_FI_DA_1121_Archivos_de_Entrada`
- `C:\CE_FI_DA_1121_Archivos_de_Salida`
- `RPA\CE_FI_DA_1121_GESTION_CONTABILIZACION_ACTIVOS_FIJOS\App\Data\json`

### 4.2 ChromeDriver y webdriver-manager

A partir de esta versión, el proyecto utiliza `webdriver-manager` para resolver automáticamente la versión correcta de ChromeDriver.

Esto elimina la necesidad de copiar manualmente `chromedriver.exe` a una ruta fija.

Si Google Chrome está instalado y la máquina tiene acceso a internet, el driver se descargará y gestionará automáticamente.

En caso de error, comprueba que:

- Google Chrome está instalado
- la versión de Chrome es compatible con el driver descargado
- el equipo puede acceder a `https://chromedriver.chromium.org/downloads`

Si necesitas forzar una ruta personalizada, modifica `App\config.py` y usa un valor distinto en `WEB_DRIVER`.

### 4.3 Colocar archivos de entrada

Coloca tu Excel de filtros aquí:

```bat
C:\CE_FI_DA_1121_Archivos_de_Entrada\Relacion.Cuentas.Dependencias.xlsx
```

Si el archivo está en otra ubicación, actualiza esta línea en `App\config.py`:

```python
EXCEL_FILTROS = r"C:\CE_FI_DA_1121_Archivos_de_Entrada\Relacion.Cuentas.Dependencias.xlsx"
```

## 5. Instalar bibliotecas Python

### 5.1 Usar el instalador automático del proyecto

Abre el símbolo de sistema (CMD) de Windows:

1. Presiona `Win + R`.
2. Escribe `cmd` y presiona `Enter`.
3. Cambia al directorio del proyecto:

```bat
cd \\ruta\a\Finansas\RPA\CE_FI_DA_1121_GESTION_CONTABILIZACION_ACTIVOS_FIJOS
```

Por ejemplo:

```bat
cd C:\Users\TuUsuario\Documents\Finansas\RPA\CE_FI_DA_1121_GESTION_CONTABILIZACION_ACTIVOS_FIJOS
```

Ejecuta:

```bat
CE1121.bat
```

Este archivo hace dos cosas:

1. Comprueba si `selenium`, `requests`, `pandas` y `openpyxl` están instalados.
2. Si falta alguna dependencia, instala todas desde la carpeta local `LIbs` usando `requirements.txt`.

### 5.2 Instalar manualmente desde CMD

Si prefieres instalar todas las dependencias tú mismo, ejecuta estos comandos desde el mismo directorio:

```bat
cd C:\Users\TuUsuario\Documents\Finansas\RPA\CE_FI_DA_1121_GESTION_CONTABILIZACION_ACTIVOS_FIJOS
python -m pip install --upgrade pip
python -m pip install --no-index --find-links=LIbs -r requirements.txt
```

> Nota: usa la ruta correcta donde esté tu proyecto.

### 5.3 Verificar la instalación

Después de instalar, revisa que `pip` y Python funcionen:

```bat
python --version
pip --version
python -c "import selenium, pandas, openpyxl; print('OK')"
```

## 6. Ejecutar la aplicación

Una vez completada la instalación, ejecuta de nuevo:

```bat
CE1121.bat
```

Esto abrirá la interfaz gráfica con Tkinter y lanzará la automatización.

## 7. Configuración adicional opcional

### 7.1 Ajustar rutas personalizadas

Abre `App\config.py` y cambia cualquiera de estas rutas si no deseas usar las rutas por defecto:

- `WEB_DRIVER`
- `ARCHIVOS_SALIDA`
- `ARCHIVOS_ENTRADA`
- `EXCEL_FILTROS`
- `URL_BASE`

### 7.2 Crear el directorio JSON

Si `App\Data\json` no existe, créalo manualmente. El proyecto usa este directorio para guardar:

- `checkpoint.json`
- `reglas.json`

## 8. Qué hacer si hay errores

- Si el instalador no encuentra `pip`, comprueba que Python se instaló con "Add Python to PATH".
- Si falta `chromedriver.exe`, descarga el driver compatible con tu versión de Chrome y colócalo en `C:\RPA\chromedriver-win64`.
- Si la aplicación no encuentra archivos, revisa las rutas en `App\config.py`.
- Revisa logs en `App\Data\logs` para diagnosticar problemas.

## 9. Resumen rápido

1. Instala Python 3.11/3.12 y añade PATH.
2. Crea carpetas: `C:\RPA\chromedriver-win64`, `C:\CE_FI_DA_1121_Archivos_de_Entrada`, `C:\CE_FI_DA_1121_Archivos_de_Salida`, y `App\Data\json`.
3. Copia `chromedriver.exe` al directorio correcto.
4. Coloca tu Excel de filtros en la carpeta de entrada.
5. Ejecuta `CE1121.bat` desde la carpeta del proyecto.
6. Si es necesario, ajusta `App\config.py`.

¡Listo! Con estos pasos deberías tener el proyecto instalado y listo para ejecutarse en Windows.

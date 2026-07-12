@echo off
REM Ejecuta la app directamente con Python, sin generar un .exe.
REM Util si no queres compilar nada, o mientras probas la app.

python -c "print(1)" >nul 2>nul
if %errorlevel% neq 0 (
    echo ERROR: No se encontro una instalacion real de Python.
    echo Si Windows menciona la "Microsoft Store", desactiva el alias en
    echo Configuracion - Aplicaciones - Alias de ejecucion de aplicaciones,
    echo y despues instala Python desde https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Verificando dependencias...
python -m pip install -r requirements.txt -q

python app.py

@echo off
echo ============================================
echo  Renombrador de Libros PDF - Generador de EXE
echo ============================================
echo.
echo Este script va a instalar las dependencias necesarias
echo y generar un archivo ejecutable (.exe) para Windows.
echo Puede tardar unos minutos la primera vez.
echo.
pause

REM Verifica que Python este instalado de verdad (evita el alias falso
REM de Microsoft Store que Windows agrega por defecto)
python --version >nul 2>nul
if %errorlevel% neq 0 (
    goto :sin_python
)
python -c "print(1)" >nul 2>nul
if %errorlevel% neq 0 (
    goto :sin_python
)
goto :con_python

:sin_python
echo.
echo ERROR: No se encontro una instalacion real de Python.
echo.
echo Si Windows te muestra un mensaje sobre "Microsoft Store", es porque
echo NO tenes Python instalado: ese es un alias falso de Windows.
echo.
echo Solucion:
echo  1. Anda a Configuracion - Aplicaciones - Configuracion avanzada de
echo     aplicaciones - Alias de ejecucion de aplicaciones
echo  2. Desactiva "python.exe" y "python3.exe"
echo  3. Instala Python desde https://www.python.org/downloads/
echo     (tildando "Add python.exe to PATH" durante la instalacion)
echo  4. Volve a ejecutar este script
echo.
pause
exit /b 1

:con_python

echo.
echo Instalando dependencias...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller

echo.
echo Limpiando compilaciones anteriores (para que tome el icono nuevo)...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist CanonicaDL.spec del /q CanonicaDL.spec

echo.
echo Generando el ejecutable (esto puede tardar 1-3 minutos)...
if exist icono.ico (
    python -m PyInstaller --onefile --windowed --name CanonicaDL --icon=icono.ico --collect-all tkinterdnd2 --collect-all customtkinter app.py
) else (
    echo (No se encontro icono.ico, se genera sin icono personalizado en el archivo .exe)
    python -m PyInstaller --onefile --windowed --name CanonicaDL --collect-all tkinterdnd2 --collect-all customtkinter app.py
)

echo.
if exist dist\CanonicaDL.exe (
    echo ============================================
    echo  LISTO! Tu ejecutable esta en:
    echo  dist\CanonicaDL.exe
    echo.
    echo  Podes mover ese archivo a donde quieras y
    echo  usarlo con doble clic, sin necesitar Python.
    echo.
    echo  Si el icono en el Explorador de Windows sigue
    echo  mostrando el icono viejo (una plumita), es el
    echo  cache de iconos de Windows, no la app: mové o
    echo  renombra el .exe a otra carpeta, o reiniciá el
    echo  Explorador de Windows. El icono de la BARRA DE
    echo  TAREAS al abrir la app sí debería verse bien.
    echo ============================================
) else (
    echo Hubo un problema generando el ejecutable.
    echo Revisa los mensajes de arriba para mas detalles.
)
echo.
pause

@echo off
cd /d %~dp0

echo ===============================
echo COMPILANDO ARCHIVO DOCUMENTAL
echo ===============================

echo.
echo [1] Instalando dependencias...
python -m pip install --upgrade pip
python -m pip install pyinstaller openpyxl python-docx pillow

echo.
echo [2] Limpiando compilaciones anteriores...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo [3] Compilando...
pyinstaller --clean --noconfirm ArchivoDocumental.spec

echo.
echo ===============================
echo PROCESO FINALIZADO
echo ===============================
echo.

pause
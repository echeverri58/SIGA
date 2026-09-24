@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Probar conversion Excel -> PDF

echo ============================================================
echo   PRUEBA DE CONVERSION  (Excel -^> PDF)
echo   Sirve para verificar que Excel puede convertir la planilla.
echo ============================================================
echo.

set "XLSX=%~1"
if not "%XLSX%"=="" goto :tengoarchivo

echo Buscando el ultimo Excel generado en Descargas...
for /f "delims=" %%f in ('dir /b /o-d "%USERPROFILE%\Downloads\SIGA_FORMATO_PLANILLA_ASISTENCIA_*.xlsx" 2^>nul') do (
    set "XLSX=%USERPROFILE%\Downloads\%%f"
    goto :tengoarchivo
)

echo No se encontro ninguno en Descargas. Se usara la plantilla.
set "XLSX=%~dp0Siga_planilla.xlsx"

:tengoarchivo
if not exist "%XLSX%" (
    echo [ERROR] No existe el archivo: %XLSX%
    echo         Genera primero un Excel desde la aplicacion, o pasa la ruta:
    echo         PROBAR_CONVERSION.bat "C:\ruta\mi_archivo.xlsx"
    pause
    exit /b 1
)

echo Archivo de entrada:
echo   %XLSX%
echo.
echo Convirtiendo... (puede tardar unos segundos)
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0convertir.ps1" "%XLSX%" "%~dp0_prueba_conversion.pdf"
set "RES=%ERRORLEVEL%"

echo.
echo ------------------------------------------------------------
if exist "%~dp0_prueba_conversion.pdf" (
    echo  RESULTADO: OK - Excel SI convirtio el archivo.
    echo  PDF: %~dp0_prueba_conversion.pdf
    echo.
    echo  Abrelo para comparar con el Excel. Debe verse igual.
    start "" "%~dp0_prueba_conversion.pdf"
) else (
    echo  RESULTADO: FALLO - Excel NO pudo convertir ^(codigo %RES%^).
    echo  El mensaje de error esta arriba.
)
echo ------------------------------------------------------------
echo.
pause

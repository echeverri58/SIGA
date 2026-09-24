@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Compartir SIGA (con PDF identico al Excel)

echo ============================================================
echo   COMPARTIR SIGA CON TUS COMPANEROS
echo   Genera un enlace publico temporal a esta aplicacion.
echo   Como aqui SI esta Excel, el PDF sale identico al Excel.
echo ============================================================
echo.

REM --- Verificar Node ---
where node >nul 2>&1
if errorlevel 1 (
    echo [ERROR] No se encontro Node.js. Instalalo desde https://nodejs.org
    pause
    exit /b 1
)

REM --- Descargar cloudflared si hace falta ---
if not exist "cloudflared.exe" (
    echo Descargando Cloudflare Tunnel ^(solo la primera vez^)...
    curl -L -o "cloudflared.exe" "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe"
    if not exist "cloudflared.exe" (
        echo [ERROR] No se pudo descargar cloudflared. Revisa tu conexion a internet.
        pause
        exit /b 1
    )
)

REM --- Arrancar el servidor local (con conversion a PDF) ---
echo Iniciando el servidor local...
start "SIGA - servidor" cmd /c "node server.js"
timeout /t 4 >nul

echo.
echo ------------------------------------------------------------
echo   Copia el enlace que aparece abajo (termina en
echo   .trycloudflare.com) y compartelo en el grupo.
echo   Manten esta ventana ABIERTA mientras tus companeros usan la app.
echo   Para terminar: cierra esta ventana y la del servidor.
echo ------------------------------------------------------------
echo.

cloudflared.exe tunnel --url http://localhost:8080

pause

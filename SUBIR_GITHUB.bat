@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
title Subir SIGA a GitHub

REM ================================================================
REM   SUBIR SIGA A GITHUB
REM   Repositorio: https://github.com/echeverri58/SIGA.git
REM   Uso: doble clic sobre este archivo.
REM        (opcional) SUBIR_GITHUB.bat "mensaje del commit"
REM ================================================================

cd /d "%~dp0"

REM ---------- 1) Verificar que Git esta instalado ----------
git --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] No se encontro Git.
    echo         Instalalo desde https://git-scm.com y vuelve a ejecutar.
    echo.
    pause
    exit /b 1
)

REM ---------- 2) Inicializar el repositorio si hace falta ----------
if not exist ".git" (
    echo Inicializando repositorio local...
    git init
)

REM ---------- 3) Usar la rama principal "main" ----------
git branch -M main 2>nul

REM ---------- 4) Configurar el remoto origin ----------
git remote remove origin 2>nul
git remote add origin https://github.com/echeverri58/SIGA.git

REM ---------- 5) Agregar y confirmar los archivos ----------
git add -A

set "MSG=%~1"
if "%MSG%"=="" set "MSG=Actualizacion SIGA"

git commit -m "%MSG%" >nul 2>&1
if errorlevel 1 (
    echo No hay cambios nuevos para confirmar.
) else (
    echo Cambios confirmados: %MSG%
)

REM ---------- 6) Subir a GitHub ----------
echo.
echo Subiendo a GitHub (la primera vez te pedira usuario y token)...
echo.
git push -u origin main
if errorlevel 1 (
    echo.
    echo [AVISO] No se pudo subir directamente. Intentando integrar
    echo         cambios que ya existan en el repositorio remoto...
    git pull origin main --rebase --allow-unrelated-histories
    if not errorlevel 1 (
        git push -u origin main
        if errorlevel 1 (
            echo.
            echo [ERROR] No se pudo subir. Revisa con: git status
        )
    ) else (
        echo.
        echo [ERROR] No se pudo subir automaticamente.
        echo         Ejecuta "git status" para revisar los conflictos.
    )
)

echo.
echo ============================================================
echo  Listo. Revisa tu repositorio en:
echo  https://github.com/echeverri58/SIGA
echo ============================================================
echo.
pause

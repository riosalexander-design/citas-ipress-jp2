@echo off
title Subir Citas C.S. Juan Pablo II a GitHub (riosalexander-design)
cd /d "%~dp0"
set "PATH=%LOCALAPPDATA%\Programs\Git\cmd;%LOCALAPPDATA%\Programs\Git\gcm;%PATH%"
set "GCM_INTERACTIVE=always"
echo ========================================================
echo   PUBLICANDO PROYECTO CITAS MEDICINA JP2 EN GITHUB
echo   Usuario: riosalexander-design
echo   Repositorio: citas-ipress-jp2
echo   Web Streamlit: https://citascsjp2.streamlit.app/
echo ========================================================
echo.
echo 1. Guardando cambios locales...
git add -A
git commit -m "Actualizacion del sistema de citas C.S. Su Santidad Juan Pablo II"
echo.
echo 2. Subiendo rama main a GitHub...
git -c credential.helper="%LOCALAPPDATA%/Programs/Git/gcm/git-credential-manager.exe" push -u origin main
echo.
if %ERRORLEVEL% equ 0 (
    echo ========================================================
    echo   EXITO: Tu codigo ya esta publicado en:
    echo   GitHub:    https://github.com/riosalexander-design/citas-ipress-jp2
    echo   Streamlit: https://citascsjp2.streamlit.app/
    echo ========================================================
) else (
    echo.
    echo Ocurrio un inconveniente al subir. Revisa tus credenciales de GitHub.
)
echo.
pause

@echo off
title C.S. Su Santidad Juan Pablo II - Citas de Medicina en Linea (Streamlit)
cd /d "%~dp0"
echo ====================================================================
echo   CENTRO DE SALUD SU SANTIDAD JUAN PABLO II
echo   SISTEMA DE CITAS EN LINEA - SERVICIO DE MEDICINA (CE)
echo   Sincronizado con: https://turnos-ipress.streamlit.app/
echo ====================================================================
echo.
echo Iniciando servidor Streamlit...
streamlit run streamlit_app.py
pause

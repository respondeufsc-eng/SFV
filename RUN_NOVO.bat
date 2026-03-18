@echo off
color 0A
title SISTEMA ANALISE MODULOS FV
echo ========================================
echo   INICIANDO SISTEMA DE ANALISE
echo ========================================
echo.
call venv\Scripts\activate
echo Ambiente virtual ativado!
echo.
echo Iniciando servidor Flask...
echo.
echo Acesse: http://localhost:5000
echo Pressione CTRL+C para parar
echo.
python app.py
pause

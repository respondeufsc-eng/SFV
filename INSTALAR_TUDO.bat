@echo off
title INSTALADOR COMPLETO - SISTEMA ANÁLISE MÓDULOS FV
color 0A
echo =====================================================
echo   INSTALADOR COMPLETO - SISTEMA ANÁLISE MÓDULOS FV
echo =====================================================
echo.

REM Verificar Python
echo [1/5] Verificando Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: Python nao encontrado!
    echo Baixe Python 3.8+ em: https://www.python.org/downloads/
    pause
    exit /b 1
)
python --version
echo.

REM Criar ambiente virtual
echo [2/5] Criando ambiente virtual...
if exist "venv" (
    echo Removendo ambiente antigo...
    rmdir /s /q venv
)
python -m venv venv
echo OK
echo.

REM Instalar TODAS as dependências
echo [3/5] Instalando dependencias...
call venv\Scripts\activate

echo Instalando Flask...
venv\Scripts\pip install flask --quiet

echo Instalando Flask-Session...
venv\Scripts\pip install flask-session --quiet

echo Instalando Pandas...
venv\Scripts\pip install pandas --quiet

echo Instalando OpenPyXL...
venv\Scripts\pip install openpyxl --quiet

echo Instalando ReportLab...
venv\Scripts\pip install reportlab --quiet

echo Instalando NumPy...
venv\Scripts\pip install numpy --quiet
echo OK
echo.

REM Verificar instalação
echo [4/5] Verificando instalacao...
venv\Scripts\python -c "import flask; print('✓ Flask OK')" 2>nul || echo "✗ Flask FALHOU"
venv\Scripts\python -c "import flask_session; print('✓ Flask-Session OK')" 2>nul || echo "✗ Flask-Session FALHOU"
venv\Scripts\python -c "import pandas; print('✓ Pandas OK')" 2>nul || echo "✗ Pandas FALHOU"
venv\Scripts\python -c "import openpyxl; print('✓ OpenPyXL OK')" 2>nul || echo "✗ OpenPyXL FALHOU"
venv\Scripts\python -c "import reportlab; print('✓ ReportLab OK')" 2>nul || echo "✗ ReportLab FALHOU"
venv\Scripts\python -c "import numpy; print('✓ NumPy OK')" 2>nul || echo "✗ NumPy FALHOU"
echo.

REM Criar RUN.bat atualizado
echo [5/5] Criando arquivos de execucao...
(
echo @echo off
echo color 0A
echo title SISTEMA ANALISE MODULOS FV
echo echo ========================================
echo echo   INICIANDO SISTEMA DE ANALISE
echo echo ========================================
echo echo.
echo call venv\Scripts\activate
echo echo Ambiente virtual ativado!
echo echo.
echo echo Iniciando servidor Flask...
echo echo.
echo echo Acesse: http://localhost:5000
echo echo Pressione CTRL+C para parar
echo echo.
echo python app.py
echo pause
) > RUN_NOVO.bat

echo.
echo =====================================================
echo   INSTALACAO CONCLUIDA COM SUCESSO!
echo =====================================================
echo.
echo PARA INICIAR O SISTEMA:
echo   1. Execute: RUN_NOVO.bat
echo.
echo PARA ACESSAR:
echo   Abra o navegador em: http://localhost:5000
echo.
pause
@echo off
chcp 65001 >nul
title Mapeamento de Dores - Servidor Local
cd /d "%~dp0"

echo ===================================================================
echo     PAINEL EXECUTIVO & OPERACIONAL - MAPEAMENTO DE DORES
echo ===================================================================
echo.
echo [1/2] Verificando dependencias do Python...
echo [2/2] Iniciando aplicacao e abrindo o navegador em http://127.0.0.1:5000 ...
echo.
echo Para fechar o sistema, basta fechar esta janela.
echo.

python app.py

if %errorlevel% neq 0 (
    echo.
    echo [AVISO] Ocorreu um problema ao executar o Python.
    echo Certifique-se de que o Python esteja instalado e marcado na opcao 'Add python.exe to PATH'.
    pause
)

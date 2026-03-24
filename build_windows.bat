@echo off
chcp 65001 >nul
title Builder - Digitador de Notas

echo ============================================================
echo   BUILDER DO DIGITADOR DE NOTAS - VERSAO OFFLINE
echo ============================================================
echo.

:: Verifica se o Python está instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado!
    echo Por favor, instale o Python em: https://www.python.org/downloads/
    echo Marque a opcao "Add Python to PATH" durante a instalacao.
    pause
    exit /b 1
)

echo [OK] Python encontrado.
echo.

:: Atualiza o pip
echo [1/4] Atualizando o pip...
python -m pip install --upgrade pip --quiet

:: Instala as dependencias
echo [2/4] Instalando dependencias (pode demorar alguns minutos)...
pip install customtkinter selenium webdriver-manager pyinstaller --quiet

if %errorlevel% neq 0 (
    echo [ERRO] Falha ao instalar dependencias!
    pause
    exit /b 1
)
echo [OK] Dependencias instaladas.
echo.

:: Gera o executavel
echo [3/4] Gerando o executavel...
echo       Isso pode levar de 1 a 3 minutos, aguarde...
echo.
pyinstaller digitador-off.spec --clean --noconfirm

if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Falha ao gerar o executavel!
    echo Verifique as mensagens de erro acima.
    pause
    exit /b 1
)

echo.
echo [4/4] Concluido!
echo.
echo ============================================================
echo   EXECUTAVEL GERADO COM SUCESSO!
echo   Pasta: dist\DigitadorDeNotas.exe
echo ============================================================
echo.
echo O arquivo "DigitadorDeNotas.exe" esta na pasta "dist".
echo Voce pode copiar esse .exe para qualquer computador Windows!
echo (O Chrome precisa estar instalado no computador de destino)
echo.
pause

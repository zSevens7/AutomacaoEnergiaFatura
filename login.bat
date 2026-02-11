@echo off
TITLE Assistente de Login - Sistema Equatorial
SETLOCAL EnableDelayedExpansion

:: --- CONFIGURACOES ---
set "PYTHON_URL=https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
set "INSTALLER_NAME=python_installer.exe"
set "CMD_PYTHON=python"

echo ========================================================
echo   SISTEMA EQUATORIAL - MODULO DE LOGIN
echo ========================================================
echo.

:: ----------------------------------------------------------
:: 1. VERIFICACAO INICIAL
:: ----------------------------------------------------------

:: 1.1 Tenta no PATH global (Melhor cenario)
python --version >nul 2>&1
if %errorlevel% equ 0 goto :VERIFICAR_LIBS

:: 1.2 Tenta encontrar executaveis ja existentes
if exist "C:\Program Files\Python311\python.exe" (
    set "CMD_PYTHON=C:\Program Files\Python311\python.exe"
    goto :VERIFICAR_LIBS
)
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    set "CMD_PYTHON=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    goto :VERIFICAR_LIBS
)

:: ----------------------------------------------------------
:: 2. INSTALACAO AUTOMATICA
:: ----------------------------------------------------------
echo [ALERTA] Python nao encontrado.
echo.
echo --------------------------------------------------------
echo   INICIANDO INSTALACAO AUTOMATICA...
echo   (Isso pode levar alguns minutos. Aguarde.)
echo --------------------------------------------------------
echo.

:: 2.1 Download
echo [1/3] Baixando instalador...
curl -o %INSTALLER_NAME% %PYTHON_URL%

if not exist %INSTALLER_NAME% (
    color 0C
    echo [ERRO] Falha no download. Verifique a internet.
    pause
    exit
)

:: 2.2 Instalacao Silenciosa
echo [2/3] Instalando... (Aceite a janela de Administrador)
start /wait %INSTALLER_NAME% /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
del %INSTALLER_NAME%

:: ----------------------------------------------------------
:: 3. DETECCAO FINAL DO AMBIENTE (O PULO DO GATO)
:: ----------------------------------------------------------
echo [3/3] Configurando ambiente...

:: TENTATIVA 1: Pasta Padrao (Program Files)
if exist "C:\Program Files\Python311\python.exe" (
    set "CMD_PYTHON=C:\Program Files\Python311\python.exe"
    goto :VERIFICAR_LIBS
)

:: TENTATIVA 2: Pasta de Usuario (AppData)
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    set "CMD_PYTHON=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    goto :VERIFICAR_LIBS
)

:: TENTATIVA 3: Usar o 'py launcher' (O Windows instala isso na pasta do sistema)
py --version >nul 2>&1
if %errorlevel% equ 0 (
    set "CMD_PYTHON=py"
    goto :VERIFICAR_LIBS
)

:: TENTATIVA 4: Procurar onde diabos o Python foi instalado
echo       Procurando executavel...
for /f "delims=" %%F in ('dir /b /s "C:\Program Files\Python311\python.exe" 2^>nul') do (
    set "CMD_PYTHON=%%F"
    goto :VERIFICAR_LIBS
)

:: Se falhar tudo, definimos como 'python' e rezamos para o PATH ter atualizado
set "CMD_PYTHON=python"

:VERIFICAR_LIBS
:: ----------------------------------------------------------
:: 4. INSTALACAO DAS BIBLIOTECAS
:: ----------------------------------------------------------
echo [STATUS] Ambiente Python detectado.
echo [1/2] Verificando dependencias...

:: Instala usando o caminho que achamos
"%CMD_PYTHON%" -m pip install --upgrade pip --quiet
"%CMD_PYTHON%" -m pip install pandas openpyxl xlsxwriter pyperclip pymupdf selenium webdriver-manager undetected-chromedriver --quiet

:: ----------------------------------------------------------
:: 5. EXECUCAO DO SCRIPT
:: ----------------------------------------------------------
echo.
echo [2/2] Iniciando Assistente de Login...
echo ---------------------------------------
cd /d "%~dp0"

:: Executa o script Python
"%CMD_PYTHON%" src/app_hibrido.py

if %errorlevel% neq 0 (
    color 0C
    echo.
    echo [ERRO] O sistema fechou com um erro.
    echo Se acabou de instalar o Python, tente fechar e abrir novamente.
    pause
) else (
    echo.
    echo [FIM] Processo finalizado.
    pause
)
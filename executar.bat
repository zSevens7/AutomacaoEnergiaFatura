@echo off
TITLE Gerador de Relatorios - Sistema Equatorial
SETLOCAL EnableDelayedExpansion

:: --- CONFIGURACOES ---
set "PYTHON_URL=https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
set "INSTALLER_NAME=python_installer.exe"
set "CMD_PYTHON=python"

echo ========================================================
echo   SISTEMA EQUATORIAL - GERADOR DE RELATORIOS
echo ========================================================
echo.

:: ----------------------------------------------------------
:: 1. VERIFICACAO INICIAL (Busca agressiva pelo Python)
:: ----------------------------------------------------------

:: 1.1 Tenta no PATH global
python --version >nul 2>&1
if %errorlevel% equ 0 goto :VERIFICAR_LIBS

:: 1.2 Tenta encontrar executaveis ja existentes (Admin ou User)
if exist "C:\Program Files\Python311\python.exe" (
    set "CMD_PYTHON=C:\Program Files\Python311\python.exe"
    goto :VERIFICAR_LIBS
)
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    set "CMD_PYTHON=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    goto :VERIFICAR_LIBS
)

:: ----------------------------------------------------------
:: 2. INSTALACAO AUTOMATICA (Se chegou aqui, nao tem Python)
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

:: Tenta localizar onde instalou para usar AGORA (sem reiniciar)
if exist "C:\Program Files\Python311\python.exe" (
    set "CMD_PYTHON=C:\Program Files\Python311\python.exe"
    goto :VERIFICAR_LIBS
)
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    set "CMD_PYTHON=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    goto :VERIFICAR_LIBS
)

:: Tenta usar o 'py launcher' (padrao do Windows)
py --version >nul 2>&1
if %errorlevel% equ 0 (
    set "CMD_PYTHON=py"
    goto :VERIFICAR_LIBS
)

:: Ultima tentativa: Busca manual
for /f "delims=" %%F in ('dir /b /s "C:\Program Files\Python311\python.exe" 2^>nul') do (
    set "CMD_PYTHON=%%F"
    goto :VERIFICAR_LIBS
)

:: Se falhar, usa o generico e torce para funcionar
set "CMD_PYTHON=python"

:VERIFICAR_LIBS
:: ----------------------------------------------------------
:: 4. INSTALACAO DAS BIBLIOTECAS
:: ----------------------------------------------------------
echo [STATUS] Ambiente Python detectado.
echo [1/2] Verificando dependencias...

:: Verifica fitz (PyMuPDF) usando o caminho correto
"%CMD_PYTHON%" -c "import fitz" >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] Instalando bibliotecas necessarias...
    "%CMD_PYTHON%" -m pip install --upgrade pip --quiet
    "%CMD_PYTHON%" -m pip install pandas openpyxl xlsxwriter pyperclip pymupdf selenium webdriver-manager undetected-chromedriver --quiet
    echo [INFO] Bibliotecas instaladas.
) else (
    echo [INFO] Dependencias ja instaladas.
)

:: ----------------------------------------------------------
:: 5. EXECUCAO DO SCRIPT (src/main.py)
:: ----------------------------------------------------------
echo.
echo [2/2] Processando Faturas PDF...
echo ---------------------------------------
cd /d "%~dp0"

:: Executa o script Python correto
"%CMD_PYTHON%" src/main.py

if %errorlevel% neq 0 (
    color 0C
    echo.
    echo [ERRO] Ocorreu um problema durante a execucao.
    echo DICA: Verifique se ha arquivos PDF na pasta 'output/faturas'.
    pause
) else (
    color 0A
    echo.
    echo ---------------------------------------
    echo [SUCESSO] Relatorio gerado com sucesso!
    echo Verifique a pasta de saida.
    pause
)
@echo off
TITLE Instalador de Dependencias - Sistema Equatorial
SETLOCAL EnableDelayedExpansion

:: Cores para o console (Fundo preto, texto verde)
color 0A

echo ========================================================
echo   SISTEMA EQUATORIAL - CONFIGURACAO DE AMBIENTE
echo ========================================================
echo.

:: 1. Verificar se o Python esta instalado
echo [PASSO 1/9] Verificando instalacao do Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    color 0C
    echo [AVISO] Python nao encontrado no sistema.
    echo Iniciando download do instalador oficial...
    
    :: Comando PowerShell para baixar o instalador silenciosamente
    powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.5/python-3.11.5-amd64.exe' -OutFile 'python_installer.exe'"
    
    echo [INFO] Instalador baixado. Por favor, instale o Python manualmente.
    echo IMPORTANTE: Marque a opcao "Add Python to PATH" durante a instalacao.
    start python_installer.exe
    pause
    exit
)
echo OK: Python detectado com sucesso.
echo.

:: 2. Atualizar o PIP
echo [PASSO 2/9] Atualizando o gerenciador de pacotes (PIP)...
python -m pip install --upgrade pip
echo.

:: 3. Instalar Pandas
echo [PASSO 3/9] Instalando biblioteca: Pandas (Processamento de Dados)...
pip install pandas
echo.

:: 4. Instalar Openpyxl
echo [PASSO 4/9] Instalando biblioteca: Openpyxl (Leitura de Excel)...
pip install openpyxl
echo.

:: 5. Instalar XlsxWriter
echo [PASSO 5/9] Instalando biblioteca: XlsxWriter (Escrita de Excel)...
pip install xlsxwriter
echo.

:: 6. Instalar Selenium
echo [PASSO 6/9] Instalando biblioteca: Selenium (Automacao Web)...
pip install selenium
echo.

:: 7. Instalar Webdriver-Manager
echo [PASSO 7/9] Instalando biblioteca: Webdriver-Manager (Gestao de Drivers)...
pip install webdriver-manager
echo.

:: 8. Instalar Undetected-Chromedriver
echo [PASSO 8/9] Instalando biblioteca: Undetected-Chromedriver (Anti-Bot)...
pip install undetected-chromedriver
echo.

:: 9. Instalar Pyperclip
echo [PASSO 9/9] Instalando biblioteca: Pyperclip (Area de Transferencia)...
pip install pyperclip
echo.

echo ========================================================
echo   INSTALACAO CONCLUIDA COM SUCESSO!
echo ========================================================
echo O ambiente esta pronto para rodar o Sistema Equatorial.
echo.
pause
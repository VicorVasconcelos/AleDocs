@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "BASE_DIR=%~dp0"
cd /d "%BASE_DIR%"

set "DRY_RUN=0"
if /I "%~1"=="--check" set "DRY_RUN=1"

set "PYTHON_EXE=%BASE_DIR%.venv\Scripts\python.exe"
if not exist "%PYTHON_EXE%" (
  echo [ERRO] Python da venv nao encontrado em:
  echo        %PYTHON_EXE%
  echo        Ative/crie a venv antes de iniciar.
  exit /b 1
)

set "NGROK_EXE="

if exist "%BASE_DIR%ngrok.exe" set "NGROK_EXE=%BASE_DIR%ngrok.exe"

if not defined NGROK_EXE (
  for %%I in (ngrok.exe) do (
    if exist "%%~$PATH:I" set "NGROK_EXE=%%~$PATH:I"
  )
)

if not defined NGROK_EXE (
  for /d %%D in ("%LOCALAPPDATA%\Microsoft\WinGet\Packages\Ngrok.Ngrok_*") do (
    if exist "%%~fD\ngrok.exe" set "NGROK_EXE=%%~fD\ngrok.exe"
  )
)

if not defined NGROK_EXE (
  echo [ERRO] Nao foi possivel localizar o ngrok.exe.
  echo        Instale o ngrok e tente novamente.
  exit /b 1
)

echo [OK] Pasta base: %BASE_DIR%
echo [OK] Python:    %PYTHON_EXE%
echo [OK] Ngrok:     %NGROK_EXE%

if "%DRY_RUN%"=="1" (
  echo.
  echo Modo validacao (^--check^): nenhuma janela sera aberta.
  echo Para iniciar de verdade, execute sem parametros.
  exit /b 0
)

echo.
echo Iniciando AleDocs e tunel publico...
echo.

start "AleDocs App" cmd /k "cd /d ""%BASE_DIR%"" && ""%PYTHON_EXE%"" app.py"
start "AleDocs Ngrok" cmd /k "cd /d ""%BASE_DIR%"" && ""%NGROK_EXE%"" http 5000"

echo Pronto!
echo 1) A janela "AleDocs App" executa o sistema local na porta 5000.
echo 2) A janela "AleDocs Ngrok" mostra o link publico https em "Forwarding".
echo.
echo Dica: salve um atalho deste .bat na Area de Trabalho para iniciar com 2 cliques.

exit /b 0

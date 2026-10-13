@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title MD2Word Local - gerar executavel e instalador
echo.
echo ===============================================================
echo   MD2Word Local - geracao do programa (so precisa fazer 1 vez)
echo ===============================================================
echo.

rem ---------- localizar o Python ----------
set "PY="
where py >nul 2>&1 && set "PY=py -3"
if not defined PY where python >nul 2>&1 && set "PY=python"
if not defined PY goto :sempython
%PY% -c "import sys; sys.exit(0 if sys.version_info >= (3,9) else 1)" >nul 2>&1
if errorlevel 1 goto :sempython

echo [1/4] Criando ambiente isolado em .venv-build ...
if not exist ".venv-build\Scripts\python.exe" %PY% -m venv .venv-build
if not exist ".venv-build\Scripts\python.exe" goto :erro
set "VPY=.venv-build\Scripts\python.exe"

echo [2/4] Instalando dependencias (a internet e usada SOMENTE nesta etapa, via PyPI) ...
"%VPY%" -m pip install --disable-pip-version-check --upgrade pip >nul
"%VPY%" -m pip install --disable-pip-version-check -r requirements.txt pyinstaller
if errorlevel 1 goto :erro

echo [3/4] Gerando dist\MD2Word\MD2Word.exe ...
"%VPY%" -m PyInstaller --noconfirm --clean --windowed --onedir --name MD2Word ^
  --icon "assets\md2word.ico" --add-data "assets;assets" --paths src ^
  --collect-data latex2mathml --collect-data docx --collect-all tkinterdnd2 ^
  src\app.py
if errorlevel 1 goto :erro
if not exist "dist\MD2Word\MD2Word.exe" goto :erro

echo Gerando LEIA-ME.docx com o proprio conversor ...
"%VPY%" src\conversor.py LEIA-ME.md >nul 2>&1

echo [4/4] Procurando o Inno Setup 6 para criar o instalador ...
call :acharinno
if defined ISCC goto :compilar
echo.
echo   O Inno Setup 6 (gratuito) nao foi encontrado. Ele so e necessario para gerar
echo   o instalador; o programa ja funciona em dist\MD2Word\MD2Word.exe
echo.
choice /C SN /M "  Instalar o Inno Setup agora pelo winget"
if errorlevel 2 goto :seminno
winget install --id JRSoftware.InnoSetup -e --accept-source-agreements --accept-package-agreements
call :acharinno
if not defined ISCC goto :seminno

:compilar
"%ISCC%" /Qp instalador.iss
if errorlevel 1 goto :erro
echo.
echo ===============================================================
echo   PRONTO! Instalador criado na pasta "instalador":
dir /b instalador\*.exe
echo   Execute-o para instalar o MD2Word Local.
echo ===============================================================
start "" explorer "%~dp0instalador"
goto :fim

:seminno
echo.
echo ===============================================================
echo   Programa gerado (versao portatil, sem instalador):
echo      %~dp0dist\MD2Word\MD2Word.exe
echo   A pasta dist\MD2Word pode ser copiada para qualquer lugar.
echo ===============================================================
start "" explorer "%~dp0dist\MD2Word"
goto :fim

:acharinno
set "ISCC="
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if exist "%LocalAppData%\Programs\Inno Setup 6\ISCC.exe" set "ISCC=%LocalAppData%\Programs\Inno Setup 6\ISCC.exe"
exit /b 0

:sempython
echo.
echo   Python 3.9 ou superior nao foi encontrado.
echo   Instale pelo site https://www.python.org/downloads/ (marque "Add python.exe to PATH")
echo   ou execute no terminal:   winget install Python.Python.3.13
echo   Depois rode este build.bat novamente.
goto :fim

:erro
echo.
echo   *** Ocorreu um erro. Veja as mensagens acima. ***

:fim
echo.
pause
endlocal

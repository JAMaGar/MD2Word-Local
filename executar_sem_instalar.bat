@echo off
setlocal EnableExtensions
cd /d "%~dp0"
rem Roda o MD2Word direto pelo Python, sem gerar o .exe (util para testar).
set "PY="
where py >nul 2>&1 && set "PY=py -3"
if not defined PY where python >nul 2>&1 && set "PY=python"
if not defined PY goto :sempython
if not exist ".venv-run\Scripts\pythonw.exe" call :preparar
start "" ".venv-run\Scripts\pythonw.exe" "src\app.py" %*
goto :eof

:preparar
echo Preparando o ambiente na primeira execucao (usa a internet so agora)...
%PY% -m venv .venv-run
".venv-run\Scripts\python.exe" -m pip install --disable-pip-version-check -r requirements.txt
exit /b 0

:sempython
echo Python nao encontrado. Instale em https://www.python.org/downloads/
pause

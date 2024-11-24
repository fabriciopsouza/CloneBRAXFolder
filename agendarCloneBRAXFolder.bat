@echo off
setlocal EnableDelayedExpansion

:: Define possíveis locais do script Python
set "DESKTOP_PATH=%USERPROFILE%\Desktop\CloneBRAXFolder.py"
set "DOWNLOADS_PATH=%USERPROFILE%\Downloads\CloneBRAXFolder.py"
set "SCRIPT_PATH="

:: Verifica se o script existe no Desktop ou Downloads
if exist "%DESKTOP_PATH%" (
    set "SCRIPT_PATH=%DESKTOP_PATH%"
    echo Script encontrado no Desktop
) else if exist "%DOWNLOADS_PATH%" (
    set "SCRIPT_PATH=%DOWNLOADS_PATH%"
    echo Script encontrado em Downloads
) else (
    echo Script CloneBRAXFolder.py nao encontrado!
    echo Procurei em:
    echo - %DESKTOP_PATH%
    echo - %DOWNLOADS_PATH%
    echo.
    echo Por favor, verifique se o arquivo existe em uma dessas pastas.
    pause
    exit /b 1
)

set "TASK_NAME=CopyExcelFiles_Daily"

:: Verifica se Python está instalado
python --version > nul 2>&1
if errorlevel 1 (
    echo Python nao encontrado! Por favor, instale o Python primeiro.
    echo Voce pode baixar em: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Cria a tarefa agendada
echo Criando tarefa agendada para executar diariamente as 09:00...
schtasks /create /tn "%TASK_NAME%" /tr "python \"%SCRIPT_PATH%\"" /sc daily /st 09:00 /ru "%USERNAME%" /f

if errorlevel 1 (
    echo Erro ao criar a tarefa agendada!
    echo Tente executar este script como administrador.
) else (
    echo.
    echo Tarefa agendada criada com sucesso!
    echo.
    echo Nome da tarefa: %TASK_NAME%
    echo Script Python: %SCRIPT_PATH%
    echo Horario de execucao: 09:00 diariamente
    echo.
)

echo.
echo Pressione qualquer tecla para sair...
pause > nul

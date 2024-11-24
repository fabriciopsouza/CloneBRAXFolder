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

:: Remove a tarefa se já existir
schtasks /query /tn "%TASK_NAME%" > nul 2>&1
if not errorlevel 1 (
    echo Removendo agendamento anterior...
    schtasks /delete /tn "%TASK_NAME%" /f > nul
)

:: Cria o XML da tarefa com retry e recuperação
set "XML_FILE=%TEMP%\task_definition.xml"
echo ^<?xml version="1.0" encoding="UTF-16"?^> > "%XML_FILE%"
echo ^<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task"^> >> "%XML_FILE%"
echo   ^<RegistrationInfo^> >> "%XML_FILE%"
echo     ^<Description^>Copia arquivos BRAX diariamente com retry em caso de falha^</Description^> >> "%XML_FILE%"
echo   ^</RegistrationInfo^> >> "%XML_FILE%"
echo   ^<Triggers^> >> "%XML_FILE%"
echo     ^<CalendarTrigger^> >> "%XML_FILE%"
echo       ^<StartBoundary^>2024-01-01T09:00:00^</StartBoundary^> >> "%XML_FILE%"
echo       ^<Enabled^>true^</Enabled^> >> "%XML_FILE%"
echo       ^<ScheduleByDay^> >> "%XML_FILE%"
echo         ^<DaysInterval^>1^</DaysInterval^> >> "%XML_FILE%"
echo       ^</ScheduleByDay^> >> "%XML_FILE%"
echo     ^</CalendarTrigger^> >> "%XML_FILE%"
echo   ^</Triggers^> >> "%XML_FILE%"
echo   ^<Principals^> >> "%XML_FILE%"
echo     ^<Principal id="Author"^> >> "%XML_FILE%"
echo       ^<LogonType^>InteractiveToken^</LogonType^> >> "%XML_FILE%"
echo       ^<RunLevel^>LeastPrivilege^</RunLevel^> >> "%XML_FILE%"
echo     ^</Principal^> >> "%XML_FILE%"
echo   ^</Principals^> >> "%XML_FILE%"
echo   ^<Settings^> >> "%XML_FILE%"
echo     ^<MultipleInstancesPolicy^>IgnoreNew^</MultipleInstancesPolicy^> >> "%XML_FILE%"
echo     ^<DisallowStartIfOnBatteries^>false^</DisallowStartIfOnBatteries^> >> "%XML_FILE%"
echo     ^<StopIfGoingOnBatteries^>false^</StopIfGoingOnBatteries^> >> "%XML_FILE%"
echo     ^<AllowHardTerminate^>true^</AllowHardTerminate^> >> "%XML_FILE%"
echo     ^<StartWhenAvailable^>true^</StartWhenAvailable^> >> "%XML_FILE%"
echo     ^<RunOnlyIfNetworkAvailable^>true^</RunOnlyIfNetworkAvailable^> >> "%XML_FILE%"
echo     ^<IdleSettings^> >> "%XML_FILE%"
echo       ^<StopOnIdleEnd^>true^</StopOnIdleEnd^> >> "%XML_FILE%"
echo       ^<RestartOnIdle^>false^</RestartOnIdle^> >> "%XML_FILE%"
echo     ^</IdleSettings^> >> "%XML_FILE%"
echo     ^<AllowStartOnDemand^>true^</AllowStartOnDemand^> >> "%XML_FILE%"
echo     ^<Enabled^>true^</Enabled^> >> "%XML_FILE%"
echo     ^<Hidden^>false^</Hidden^> >> "%XML_FILE%"
echo     ^<RunOnlyIfIdle^>false^</RunOnlyIfIdle^> >> "%XML_FILE%"
echo     ^<RestartOnFailure^> >> "%XML_FILE%"
echo       ^<Interval^>PT1H^</Interval^> >> "%XML_FILE%"
echo       ^<Count^>2^</Count^> >> "%XML_FILE%"
echo     ^</RestartOnFailure^> >> "%XML_FILE%"
echo     ^<WakeToRun^>false^</WakeToRun^> >> "%XML_FILE%"
echo     ^<ExecutionTimeLimit^>PT12H^</ExecutionTimeLimit^> >> "%XML_FILE%"
echo     ^<Priority^>7^</Priority^> >> "%XML_FILE%"
echo   ^</Settings^> >> "%XML_FILE%"
echo   ^<Actions Context="Author"^> >> "%XML_FILE%"
echo     ^<Exec^> >> "%XML_FILE%"
echo       ^<Command^>python^</Command^> >> "%XML_FILE%"
echo       ^<Arguments^>"%SCRIPT_PATH%"^</Arguments^> >> "%XML_FILE%"
echo     ^</Exec^> >> "%XML_FILE%"
echo   ^</Actions^> >> "%XML_FILE%"
echo ^</Task^> >> "%XML_FILE%"

:: Cria a tarefa usando o XML
echo Criando tarefa agendada...
schtasks /create /xml "%XML_FILE%" /tn "%TASK_NAME%" /f

if errorlevel 1 (
    echo Erro ao criar a tarefa agendada!
    echo Tentando metodo alternativo...
    
    :: Tenta criar usando comando direto do schtasks
    schtasks /create /tn "%TASK_NAME%" /tr "python \"%SCRIPT_PATH%\"" /sc daily /st 09:00 /ru "%USERNAME%" /f /rl HIGHEST ^
             /V1 /Z /ed "31/12/2025" /IT ^
             /RI 60 /RP 2

    if errorlevel 1 (
        echo Falha ao criar tarefa agendada!
    ) else (
        echo Tarefa criada com sucesso usando metodo alternativo!
    )
) else (
    echo.
    echo Tarefa agendada criada com sucesso!
    echo.
    echo Nome da tarefa: %TASK_NAME%
    echo Script Python: %SCRIPT_PATH%
    echo Horario de execucao: 09:00 diariamente
    echo.
    echo Configuracoes adicionais:
    echo - Tenta novamente 2 vezes se falhar
    echo - Espera 1 hora entre tentativas
    echo - Executa tarefas perdidas quando possivel
    echo - Requer conexao de rede
    echo.
)

:: Limpa o arquivo XML temporário
del "%XML_FILE%" > nul 2>&1

echo.
echo Pressione qualquer tecla para sair...
pause > nul

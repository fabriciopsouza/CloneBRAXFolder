# Caminhos e configurações
$SourcePath = "Y:\OPER_BAP_ESO\NP-1\01 - Carimbo AIVI (YSFECHAMENTO)\BRAX RPA"
$DestinationPath = "\\172.31.9.186\interfaces\BW0924\OPER-ESO"
$FilePattern = "YSFECHAMENTO AIVI (ysvi_homolog_* oficial).xlsx"
$TaskName = "CopyLatestExcelFile"
$VPNProcessName = "FortiClient"

# Obter o caminho da pasta "Meus Documentos"
$MyDocuments = [Environment]::GetFolderPath("MyDocuments")
$LogFile = Join-Path $MyDocuments "CopyLatestExcelFile.log"
$LastRunTimeFile = Join-Path $MyDocuments "CopyLatestExcelFile_LastRunTime.txt"

# Função para verificar a conexão VPN
function Wait-ForVPN {
    Write-Output "Verificando conexão VPN..."
    while (-not (Get-Process -Name $VPNProcessName -ErrorAction SilentlyContinue)) {
        Write-Output "VPN não conectada. Aguardando..."
        Start-Sleep -Seconds 10
    }
    Write-Output "VPN conectada."
}

# Verificar se é a primeira execução
if (!(Test-Path $LastRunTimeFile)) {
    $FirstExecution = $true
    $LastRunTime = Get-Date "01/01/1970"
} else {
    $FirstExecution = $false
    $LastRunTime = Get-Content $LastRunTimeFile | Out-String
    $LastRunTime = [datetime]::Parse($LastRunTime)
}

# Esperar pela conexão VPN
Wait-ForVPN

# Obter os arquivos para copiar
if ($FirstExecution) {
    # Primeira execução: copiar todos os arquivos correspondentes
    $Files = Get-ChildItem -Path $SourcePath -Filter $FilePattern -File -Recurse
} else {
    # Execuções subsequentes: copiar somente arquivos modificados desde a última execução
    $Files = Get-ChildItem -Path $SourcePath -Filter $FilePattern -File -Recurse | Where-Object { $_.LastWriteTime -gt $LastRunTime }
}

# Copiar arquivos e registrar no log
foreach ($File in $Files) {
    try {
        $DestinationFile = Join-Path $DestinationPath $File.Name
        Copy-Item -Path $File.FullName -Destination $DestinationFile -Force
        $LogEntry = "$(Get-Date) - Copiado: $($File.FullName) para $DestinationFile"
        Add-Content -Path $LogFile -Value $LogEntry
    } catch {
        $LogEntry = "$(Get-Date) - Erro ao copiar $($File.FullName): $_"
        Add-Content -Path $LogFile -Value $LogEntry
    }
}

# Atualizar o arquivo com a data/hora da última execução
$CurrentTime = Get-Date
$CurrentTime | Out-File -FilePath $LastRunTimeFile -Force

# Agendar a tarefa no Task Scheduler
# Verificar se a tarefa já existe
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    # Remover a tarefa existente
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

# Criar a ação e o gatilho para a tarefa agendada
$ScriptPath = $MyInvocation.MyCommand.Definition
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File `"$ScriptPath`" -WindowStyle Hidden"
$TriggerTime = Get-Date
$Trigger = New-ScheduledTaskTrigger -Daily -At $TriggerTime.TimeOfDay

# Registrar a tarefa agendada
Register-ScheduledTask -Action $Action -Trigger $Trigger -TaskName $TaskName -Description "Copia arquivos Excel modificados" -RunLevel Highest -Force

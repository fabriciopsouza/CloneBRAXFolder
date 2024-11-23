# Caminhos e configurações
$SourcePath = "Y:\OPER_BAP_ESO\NP-1\01 - Carimbo AIVI (YSFECHAMENTO)\BRAX RPA"
$DestinationPath = "\\172.31.9.186\interfaces\BW0924\OPER-ESO"
$FilePattern = "YSFECHAMENTO AIVI (ysvi_homolog_* oficial).xlsx"
$TaskName = "CopyLatestExcelFile"
$VPNProcessName = "FortiClient"

# Obter o caminho da pasta "Meus Documentos"
$MyDocuments = [Environment]::GetFolderPath("MyDocuments")
$LogFile = Join-Path $MyDocuments "CopyLatestExcelFile\CopyLatestExcelFile.log"
$LastRunTimeFile = Join-Path $MyDocuments "CopyLatestExcelFile\CopyLatestExcelFile_LastRunTime.txt"

# Criar a pasta para os logs se não existir
if (!(Test-Path (Split-Path $LogFile))) {
    New-Item -Path (Split-Path $LogFile) -ItemType Directory | Out-Null
}

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
    $LastRunTimeContent = Get-Content $LastRunTimeFile | Out-String
    $LastRunTime = [datetime]::Parse($LastRunTimeContent)
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
$CurrentTime.ToString() | Out-File -FilePath $LastRunTimeFile -Force

# Agendar a tarefa usando SCHTASKS
# Verificar se a tarefa já existe
$TaskExists = schtasks /Query /TN $TaskName 2>&1 | Select-String $TaskName

if ($TaskExists) {
    # Remover a tarefa existente
    schtasks /Delete /TN $TaskName /F > $null 2>&1
}

# Obter o caminho completo do script
$ScriptPath = $MyInvocation.MyCommand.Path

# Agendar a nova tarefa
$TriggerTime = (Get-Date).ToString("HH:mm")
schtasks /Create /SC DAILY /TN $TaskName /TR "PowerShell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`"" /ST $TriggerTime /F > $null 2>&1

# Caminhos e configurações
$SourcePath = "Y:\OPER_BAP_ESO\NP-1\01 - Carimbo AIVI (YSFECHAMENTO)\BRAX RPA"
$DestinationPath = "\\172.31.9.186\interfaces\BW0924\OPER-ESO"
$FilePattern = "YSFECHAMENTO AIVI (ysvi_homolog_* oficial).xlsx"
$VPNProcessName = "FortiClient"

# Obter o caminho da pasta "Meus Documentos"
$MyDocuments = [Environment]::GetFolderPath("MyDocuments")
$LogFolder = Join-Path $MyDocuments "CopyLatestExcelFile"
$LogFile = Join-Path $LogFolder "CopyLatestExcelFile.log"
$LastRunTimeFile = Join-Path $LogFolder "CopyLatestExcelFile_LastRunTime.txt"

# Criar a pasta para os logs se não existir
if (!(Test-Path $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory | Out-Null
}

# Função para verificar a conexão VPN
function Wait-ForVPN {
    Write-Host "Verificando conexão VPN..."
    while (-not (Get-Process -Name $VPNProcessName -ErrorAction SilentlyContinue)) {
        Write-Host "VPN não conectada. Aguardando..."
        Start-Sleep -Seconds 10
    }
    Write-Host "VPN conectada."
}

# Verificar se é a primeira execução
if (!(Test-Path $LastRunTimeFile)) {
    $FirstExecution = $true
    $LastRunTime = Get-Date "01/01/1970"
} else {
    $FirstExecution = $false
    $LastRunTimeContent = Get-Content $LastRunTimeFile
    $LastRunTime = [datetime]::Parse($LastRunTimeContent)
}

# Esperar pela conexão VPN
Wait-ForVPN

# Obter os arquivos para copiar
if ($FirstExecution) {
    # Primeira execução: copiar todos os arquivos correspondentes
    $Files = Get-ChildItem -Path $SourcePath -Filter $FilePattern -File
} else {
    # Execuções subsequentes: copiar somente arquivos modificados desde a última execução
    $Files = Get-ChildItem -Path $SourcePath -Filter $FilePattern -File | Where-Object { $_.LastWriteTime -gt $LastRunTime }
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

Write-Host "Processo concluído."

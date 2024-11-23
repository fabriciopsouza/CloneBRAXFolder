# Caminhos e configurações
$SourcePath = "Y:\OPER_BAP_ESO\NP-1\01 - Carimbo AIVI (YSFECHAMENTO)\BRAX RPA"
$DestinationPath = "\\172.31.9.186\interfaces\BW0924\OPER-ESO"
$FilePattern = "YSFECHAMENTO AIVI (ysvi_homolog_* oficial).xlsx"
$VPNProcessName = "FortiClient"

# Configuração de logs
$LogFolder = Join-Path $env:USERPROFILE "Documents\CopyLatestExcelFile"
$LogFile = Join-Path $LogFolder "CopyLatestExcelFile.log"
$LastRunTimeFile = Join-Path $LogFolder "LastRunTime.txt"

# Criar pasta de logs se não existir
if (!(Test-Path $LogFolder)) {
    try {
        New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
        Write-Host "Pasta de logs criada em: $LogFolder"
    } catch {
        Write-Host "Erro ao criar pasta de logs: $_"
        exit 1
    }
}

# Função para verificar a conexão VPN
function Test-VPNConnection {
    try {
        $vpnProcess = Get-Process -Name $VPNProcessName -ErrorAction SilentlyContinue
        return $null -ne $vpnProcess
    } catch {
        Write-Host "Erro ao verificar VPN: $_"
        return $false
    }
}

# Função para registrar log
function Write-Log {
    param($Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    try {
        Add-Content -Path $LogFile -Value $logMessage -ErrorAction Stop
        Write-Host $logMessage
    } catch {
        Write-Host "Erro ao escrever log: $_"
    }
}

# Verificar VPN
Write-Host "Verificando conexão VPN..."
if (-not (Test-VPNConnection)) {
    Write-Host "ERRO: VPN não está conectada (FortiClient não está em execução)"
    exit 1
}

# Verificar última execução
try {
    if (Test-Path $LastRunTimeFile) {
        $LastRunTime = Get-Content $LastRunTimeFile | Get-Date
    } else {
        $LastRunTime = Get-Date "1970-01-01"
        Write-Log "Primeira execução detectada"
    }
} catch {
    Write-Log "Erro ao ler arquivo de última execução: $_"
    $LastRunTime = Get-Date "1970-01-01"
}

# Verificar acesso às pastas
if (!(Test-Path $SourcePath)) {
    Write-Log "ERRO: Pasta de origem não encontrada: $SourcePath"
    exit 1
}

if (!(Test-Path $DestinationPath)) {
    Write-Log "ERRO: Pasta de destino não encontrada: $DestinationPath"
    exit 1
}

# Buscar arquivos para copiar
try {
    $Files = Get-ChildItem -Path $SourcePath -Filter $FilePattern -File |
             Where-Object { $_.LastWriteTime -gt $LastRunTime }
    
    Write-Log "Encontrados $($Files.Count) arquivos para copiar"
} catch {
    Write-Log "Erro ao buscar arquivos: $_"
    exit 1
}

# Copiar arquivos
foreach ($File in $Files) {
    try {
        $DestinationFile = Join-Path $DestinationPath $File.Name
        Copy-Item -Path $File.FullName -Destination $DestinationFile -Force
        Write-Log "Arquivo copiado com sucesso: $($File.Name)"
    } catch {
        Write-Log "ERRO ao copiar arquivo $($File.Name): $_"
    }
}

# Atualizar última execução
try {
    Get-Date -Format "yyyy-MM-dd HH:mm:ss" | Out-File -FilePath $LastRunTimeFile -Force
    Write-Log "Horário da última execução atualizado"
} catch {
    Write-Log "Erro ao atualizar horário da última execução: $_"
}

Write-Log "Processo concluído"

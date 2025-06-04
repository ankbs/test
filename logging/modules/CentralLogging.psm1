function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet("INFO","SUCCESS","ERROR","DEBUG")]
        [string]$Level = "INFO",
        [string]$LogFile = $script:CentralLogFile
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $prefix = switch ($Level) {
        "INFO"    { "ℹ️" }
        "SUCCESS" { "✅" }
        "ERROR"   { "❌" }
        "DEBUG"   { "🐛" }
        default   { "🔹" }
    }
    $entry = "$timestamp $prefix $Message"
    try {
        if ($LogFile) {
            Add-Content -Path $LogFile -Value $entry -Encoding utf8
        } else {
            Write-Host "WARNUNG: Kein LogFile gesetzt!" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "FEHLER beim Schreiben ins Logfile: $entry" -ForegroundColor Red
    }
    Write-Host $entry -Encoding utf8
}

function Set-LogFile {
    param(
        [string]$LogFolder = "$PSScriptRoot\Logs"
    )
    if (-not (Test-Path $LogFolder)) { New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null }
    $script:CentralLogFile = Join-Path $LogFolder ("CentralLog_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}

function Write-LogError {
    param(
        [string]$Message,
        [object]$ErrorObject
    )
    $msg = if ($ErrorObject -is [System.Management.Automation.ErrorRecord]) { $ErrorObject.Exception.Message } else { $ErrorObject.ToString() }
    $fullMsg = "${Message}: $msg"
    try {
        Write-Log -Message $fullMsg -Level "ERROR"
    } catch {
        Write-Host "FEHLER beim Logging (Write-LogError): $fullMsg" -ForegroundColor Red
    }
    throw $fullMsg
}
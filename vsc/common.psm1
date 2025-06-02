# Zentrale Funktionen für Logging, Error-Handling und Modulprüfung

function Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $prefix = switch ($Level) {
        "INFO"     { "ℹ️" }
        "SUCCESS"  { "✅" }
        "ERROR"    { "❌" }
        "DEBUG"    { "🐛" }
        default    { "🔹" }
    }
    $logEntry = "$timestamp $prefix $Message"
    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $logEntry -Encoding utf8
    }
    Write-Host $logEntry -Encoding utf8
}

function Handle-Error {
    param (
        [string]$Message,
        [object]$ErrorObject
    )
    $exceptionMessage = switch ($ErrorObject) {
        { $_ -is [System.Exception] } { $_.Message }
        { $_ -is [System.Management.Automation.ErrorRecord] } { $_.Exception.Message }
        default { "$ErrorObject" }
    }
    $fullMessage = "${Message}: $exceptionMessage"
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "$timestamp ❌ $fullMessage"
    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $logEntry -Encoding utf8
    }
    Write-Host $logEntry -ForegroundColor Red
}

function Ensure-Module {
    param([string]$ModuleName)
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Log "📦 Modul '$ModuleName' nicht gefunden – versuche Installation..." "INFO"
        if (-not (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue)) {
            Register-PSRepository -Default
        }
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber
            Log "✅ Modul '$ModuleName' installiert." "SUCCESS"
        } catch {
            Handle-Error "Modulinstallation fehlgeschlagen" $_
        }
    } else {
        Log "✅ Modul '$ModuleName' ist bereits installiert." "DEBUG"
    }
}

Export-ModuleMember -Function Log,Handle-Error,Ensure-Module
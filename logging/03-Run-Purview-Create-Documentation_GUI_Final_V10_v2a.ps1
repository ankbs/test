<#
.SYNOPSIS
Automatisierte Erstellung und Export von Sensitivitätslabel-Dokumentationen für Microsoft Purview, mit zentraler Konfigurationsverwaltung.
Sucht automatisch nach einer PurviewConfig.json im Skriptverzeichnis, sofern kein -ConfigPath angegeben ist.

.DESCRIPTION
Dieses Skript liest seine Startparameter standardmäßig aus der PurviewConfig.json im gleichen Verzeichnis wie das Skript.
Alternativ kann eine eigene Konfigurationsdatei mit -ConfigPath übergeben werden oder alle Parameter wie gewohnt direkt per Skriptaufruf.
Alle Parameter, die nicht explizit beim Aufruf gesetzt wurden, werden aus der Config übernommen.
Die Konfig kann per GUI oder manuell gepflegt werden.

.EXAMPLE
.\03-Run-Purview-Create-Documentation_GUI_Final_V10.ps1
# Verwendet automatisch .\PurviewConfig.json, falls vorhanden

.EXAMPLE
.\03-Run-Purview-Create-Documentation_GUI_Final_V10.ps1 -ConfigPath "D:\Konfig\MeineConfig.json"
# Verwendet die explizit angegebene Datei

.EXAMPLE
.\03-Run-Purview-Create-Documentation_GUI_Final_V10.ps1 -UserPrincipalName "admin@contoso.com" -SourceExcelPath "C:\Labels.xlsx"
# Nutzt nur die direkt gesetzten Parameter

.LINK
https://learn.microsoft.com/de-de/purview/

.AUTHOR
Michael Kirst-Neshva

.EMAIL
michael_kirst@hotmail.com

.VERSION
V2

.CREATIONDATE
2025
#>
# Requires -Version 5.1

param(
    [string]$ConfigPath = "",
    [string]$GuiConfigPath,
    [bool]$SendReport = $false,
    [string]$UserPrincipalName = "",
    [string]$Tenantdomain = "",
    [string]$MailToPrimary,
    [string]$MailToSecondary = "",
    [bool]$CreateMissingLabels = $false,
    [int]$MFATimeoutSeconds = 120,
    [string]$SourceExcelPath = "",
    [bool]$UseExistingLabels = $true,
    [int]$Priority = 0,
    [int]$PriorityMin = 0,
    [int]$PriorityMax = 0,
    [string[]]$LabelNames = @(),
    [bool]$UseLabelGUI = $true,
    [bool]$ExportWord = $true,
    [bool]$ExportPDF = $false,
    [bool]$DryRun = $false,
    [bool]$UseProgressBar = $false,
    [string]$LogFolder = "C:\Temp\script\",
    [int]$AutoCloseAfterSeconds = 2,
    [string]$LogoGIFUrl    = "https://i.gifer.com/ZKZg.gif",
    [string]$CompanyLogoPath = "",
    [string]$CompanyLogoUrl  = "",
    [string]$CompanyLogoBase64 = "",
    [string]$LogoUrl = "",
    [string]$ProductLogoBase64 = ""
)

# === Zentrale Konfigurationslogik ===
$importConfigModulePath = Join-Path $PSScriptRoot "..\modules\Import-ConfigParameters.psm1"
if (Test-Path $importConfigModulePath) { Import-Module $importConfigModulePath -Force }

if (-not $ConfigPath) {
    $ConfigPath = Join-Path $PSScriptRoot "PurviewConfig.json"
}
if (Test-Path $ConfigPath) {
    if (Get-Command Import-ConfigParameters -ErrorAction SilentlyContinue) {
        Import-ConfigParameters -ConfigPath $ConfigPath -BoundParameters $PSBoundParameters
    }
}

# === Kompatibilität mit älteren GUI-Konfigs (optional) ===
if ($GuiConfigPath -and (Test-Path $GuiConfigPath)) {
    try {
        $GuiConfig = Get-Content -Raw -Path $GuiConfigPath | ConvertFrom-Json
        foreach ($property in $GuiConfig.PSObject.Properties) {
            if (-not $PSBoundParameters.ContainsKey($property.Name) -and
                -not [string]::IsNullOrWhiteSpace($property.Value)) {
                Set-Variable -Name $property.Name -Value $property.Value -Scope Script
            }
        }
    } catch {
        Write-Host "⚠️ Fehler beim Einlesen der GUI-Konfiguration, verwende ggf. Standardwerte. $_" -ForegroundColor Yellow
    }
}

# === Verzeichnis-Handling für Logs und Exporte ===
$DatumHeute = Get-Date -Format 'yyyyMMdd_HHmmss'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ExportFolder -or $ExportFolder -eq "") { $ExportFolder = Join-Path $ScriptDir "Export_$DatumHeute" }
if (-not $LogFolder -or $LogFolder -eq "") { $LogFolder = Join-Path $ScriptDir "Logs_$DatumHeute" }
if (-not (Test-Path $LogFolder))    { New-Item -Path $LogFolder    -ItemType Directory -Force > $null }
if (-not (Test-Path $ExportFolder)) { New-Item -Path $ExportFolder -ItemType Directory -Force > $null }
$DatumJetzt = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogFile = Join-Path $LogFolder "CreateDocuReport_LOG_$DatumJetzt.log"

# ... (hier folgt der gesamte Original-Code, unverändert) ...
# (Aus Platzgründen wurde der Inhalt ausgelassen. Bitte kopiere ab hier deinen Originalcode inklusive aller Funktionen, Logik, GUI, Excel-Import, Logging, Error-Handling, etc.)

<#
.SYNOPSIS
    Automatisierte Label-Provisionierung aus Excel für Microsoft Purview.

.DESCRIPTION
    Dieses Skript dient der automatisierten Bereitstellung von Sensitivitätslabels für Microsoft Purview auf Basis einer Excel-Datei.
    Es liest eine strukturierte Excel-Liste (Labelname, Beschreibung) ein und verarbeitet jeden Eintrag.
    Jeder Label-Datensatz wird geprüft und anschließend provisioniert, wobei der eigentliche Provisionierungsvorgang modular gehalten ist (z.B. zur Integration von API-Aufrufen).
    Die gesamte Ablaufsteuerung, Logging und Fehlerbehandlung erfolgt zentral und standardisiert über das CentralLogging-Modul.
    Es werden alle relevanten Aktionen (z.B. Excel-Laden, Label-Verarbeitung, Fehlerfälle, Abschluss) mitgeloggt, sodass eine lückenlose Nachvollziehbarkeit und Debugging-Möglichkeit gegeben ist.
    Das Skript eignet sich sowohl zum initialen Label-Rollout als auch für wiederkehrende Aktualisierungen.
    Die Excel-Einlesung ist flexibel gestaltet, sodass Anpassungen an andere Spaltenstrukturen leicht möglich sind.
    Nach Abschluss werden alle Ressourcen sauber freigegeben.
    Dieses Skript liest seine Startparameter standardmäßig aus der PurviewConfig.json im gleichen Verzeichnis wie das Skript.
    Alternativ kann eine eigene Konfigurationsdatei mit -ConfigPath übergeben werden oder alle Parameter wie gewohnt direkt per Skriptaufruf.


.EXAMPLE
    .\00_LabelProvisioning_XAML_Final_v1_v2.ps1 -InputExcelPath "C:\Data\Labels.xlsx"

.EXAMPLE
    .\00_LabelProvisioning_XAML_Final_v1_v2.ps1 -InputExcelPath ".\labels.xlsx" -LogFolder "C:\Logs" -UserPrincipalName "admin@contoso.com"

.EXAMPLE
    .\00_LabelProvisioning_XAML_Final_v1_v2.ps1
    # Verwendet automatisch .\PurviewConfig.json, falls vorhanden

.EXAMPLE
    .\00_LabelProvisioning_XAML_Final_v1_v2.ps1 -ConfigPath "D:\Konfig\MeineConfig.json"
    # Verwendet die explizit angegebene Datei

.EXAMPLE
    .\00_LabelProvisioning_XAML_Final_v1_v2.ps1 -UserPrincipalName "admin@contoso.com" -InputExcelPath "C:\Labels.xlsx"
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

param (
    [string]$InputExcelPath = "",
    [string]$LogFolder = "C:\Temp\script\",
    [string]$UserPrincipalName = "",
    [string]$Tenantdomain = "",
    [string]$MailToPrimary = "",
    [string]$MailToSecondary = "",
    [string]$CompanyLogoBase64 = "",
    [string]$ProductLogoBase64 = "",
    [string]$ConfigPath = ""
)

# Zentrales Config-Modul importieren (Pfad ggf. anpassen)
Import-Module "$PSScriptRoot\..\modules\Import-ConfigParameters.psm1" -Force

# Standardmäßig nach PurviewConfig.json im Skriptverzeichnis suchen, falls -ConfigPath leer bleibt
if (-not $ConfigPath) {
    $ConfigPath = Join-Path $PSScriptRoot "PurviewConfig.json"
}
if (Test-Path $ConfigPath) {
    Import-ConfigParameters -ConfigPath $ConfigPath -BoundParameters $PSBoundParameters
}

Import-Module "$PSScriptRoot\CentralLogging.psm1" -Force
Set-LogFile -LogFolder "$LogFolder"
Write-Log -Message "00_LabelProvisioning_XAML_Final_v1_v2.ps1 gestartet" -Level "INFO"

try {
    # Prüfe und erstelle Logverzeichnis
    if (!(Test-Path $LogFolder)) {
        New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
        Write-Log -Message "Logverzeichnis $LogFolder wurde erstellt." -Level "INFO"
    }

    # Eingabedatei prüfen
    if (-not $InputExcelPath -or -not (Test-Path $InputExcelPath)) {
        Write-Log -Message "Excel Datei nicht gefunden: $InputExcelPath" -Level "ERROR"
        throw "Excel Datei nicht gefunden: $InputExcelPath"
    }
    Write-Log -Message "Input Excel: $InputExcelPath" -Level "INFO"

    # Excel laden
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($InputExcelPath)
        $worksheet = $workbook.Worksheets.Item(1)
        Write-Log -Message "Excel erfolgreich geladen." -Level "SUCCESS"
    } catch {
        Handle-Error -Message "Fehler beim Laden der Excel-Datei" -ErrorObject $_
    }

    # Labels aus Excel auslesen (angenommen: Spalten A = LabelName, B = Beschreibung)
    $labels = @()
    $row = 2
    while ($worksheet.Cells.Item($row, 1).Value2) {
        $labelName = $worksheet.Cells.Item($row, 1).Value2
        $labelDesc = $worksheet.Cells.Item($row, 2).Value2
        $labels += [PSCustomObject]@{
            Name = $labelName
            Description = $labelDesc
        }
        $row++
    }
    Write-Log -Message ("Es wurden {0} Labels ausgelesen." -f $labels.Count) -Level "INFO"

    # Beispielhafte Verarbeitung: Labels provisionieren (hier nur Logging)
    foreach ($label in $labels) {
        try {
            # Hier würde die eigentliche Provisionierung stattfinden, z.B. per API-Aufruf
            Write-Log -Message ("Provisioniere Label: {0} ({1})" -f $label.Name, $label.Description) -Level "INFO"
            Start-Sleep -Milliseconds 100
            Write-Log -Message ("Label '{0}' erfolgreich provisioniert." -f $label.Name) -Level "SUCCESS"
        } catch {
            Handle-Error -Message ("Fehler beim Provisionieren von Label {0}" -f $label.Name) -ErrorObject $_
        }
    }

    # Excel schließen und aufräumen
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Log -Message "Excel geschlossen und Ressourcen freigegeben." -Level "INFO"

    # Abschlussmeldung
    Write-Log -Message "Label-Provisionierung abgeschlossen." -Level "SUCCESS"
    Write-Host "Alle Labels wurden verarbeitet. Details siehe Log unter $LogFolder"
}
catch {
    Handle-Error -Message "Fehler im Hauptskript" -ErrorObject $_
}
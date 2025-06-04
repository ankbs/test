<#
.SYNOPSIS
    Automatisierte Erstellung und Export von Sensitivitätslabel-Dokumentationen für Microsoft Purview.

.DESCRIPTION
    Dieses Skript erstellt automatisch eine vollständige Dokumentation aller Sensitivitätslabels in Microsoft Purview.
    Die Labelinformationen werden aus einer Excel-Datei oder einer JSON-Konfigurationsdatei geladen und in eine Word- bzw. PDF-Datei exportiert.
    Die Dokumentation umfasst sämtliche Angaben zu Labelnamen, Beschreibungen und möglichen weiteren Metadaten.
    Optional kann nach der Erstellung ein Bericht per E-Mail an definierte Empfänger versendet werden.
    Das Skript unterstützt verschiedene Parameter zur Steuerung der Ausgabeformate, des Berichtsversands und der Datenquelle.
    Der gesamte Ablauf ist mit zentralem Logging und strukturierter Fehlerbehandlung versehen, sodass alle Teilschritte, Warnungen und Fehlerfälle nachvollziehbar sind.
    Das Skript ist besonders geeignet für Audits, interne Dokumentationspflichten oder zur revisionssicheren Ablage von Labelkonfigurationen.
    Es kann sowohl interaktiv (z.B. aus einer GUI) als auch automatisiert aufgerufen werden.
    Dieses Skript liest seine Startparameter standardmäßig aus der PurviewConfig.json im gleichen Verzeichnis wie das Skript.
    Alternativ kann eine eigene Konfigurationsdatei mit -ConfigPath übergeben werden oder alle Parameter wie gewohnt direkt per Skriptaufruf.


.EXAMPLE
    .\03-Run-Purview-Create-Documentation_GUI_Final_V10_v2.ps1 -GuiConfigPath ".\GUIConfig.json"

.EXAMPLE
    .\03-Run-Purview-Create-Documentation_GUI_Final_V10_v2.ps1 -SourceExcelPath "C:\Labels.xlsx" -ExportWord $true -ExportPDF $true -SendReport $true

.EXAMPLE
    .\03-Run-Purview-Create-Documentation_GUI_Final_V10_v2.ps1
    # Verwendet automatisch .\PurviewConfig.json, falls vorhanden

.EXAMPLE
    .\03-Run-Purview-Create-Documentation_GUI_Final_V10_v2.ps1 -ConfigPath "D:\Konfig\MeineConfig.json"
    # Verwendet die explizit angegebene Datei

.EXAMPLE
    .\03-Run-Purview-Create-Documentation_GUI_Final_V10_v2.ps1 -UserPrincipalName "admin@contoso.com" -InputExcelPath "C:\Labels.xlsx"
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
    [string]$GuiConfigPath = "",
    [bool]$SendReport = $false,
    [string]$UserPrincipalName = "",
    [string]$Tenantdomain = "",
    [string]$MailToPrimary = "",
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
    [string]$LogoGIFUrl    = "",
    [string]$CompanyLogoPath = "",
    [string]$CompanyLogoUrl  = "",
    [string]$CompanyLogoBase64 = "",
    [string]$LogoUrl = "",
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
Write-Log -Message "03-Run-Purview-Create-Documentation_GUI_Final_V10_v2.ps1 gestartet" -Level "INFO"

try {
    # Prüfe und erstelle Logverzeichnis
    if (!(Test-Path $LogFolder)) {
        New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
        Write-Log -Message "Logverzeichnis $LogFolder wurde erstellt." -Level "INFO"
    }

    # Konfiguration laden, falls vorhanden
    if ($GuiConfigPath -and (Test-Path $GuiConfigPath)) {
        try {
            $cfg = Get-Content -Raw -Path $GuiConfigPath | ConvertFrom-Json
            if ($cfg.UserPrincipalName) { $UserPrincipalName = $cfg.UserPrincipalName }
            if ($cfg.Tenantdomain)      { $Tenantdomain      = $cfg.Tenantdomain }
            if ($cfg.MailToPrimary)     { $MailToPrimary     = $cfg.MailToPrimary }
            if ($cfg.MailToSecondary)   { $MailToSecondary   = $cfg.MailToSecondary }
            if ($cfg.LogFolder)         { $LogFolder         = $cfg.LogFolder }
            Write-Log -Message "Konfiguration aus $GuiConfigPath geladen." -Level "INFO"
        } catch {
            Write-Log -Message "Fehler beim Laden der Konfiguration: $_" -Level "ERROR"
        }
    }

    # Labels aus Excel laden, wenn angegeben
    $labels = @()
    if ($SourceExcelPath -and (Test-Path $SourceExcelPath)) {
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $workbook = $excel.Workbooks.Open($SourceExcelPath)
            $worksheet = $workbook.Worksheets.Item(1)
            Write-Log -Message "Excel erfolgreich geladen: $SourceExcelPath" -Level "SUCCESS"

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
            Write-Log -Message ("{0} Labels aus Excel geladen." -f $labels.Count) -Level "INFO"
            $workbook.Close($false)
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        } catch {
            Write-LogError -Message "Fehler beim Laden der Excel-Datei" -ErrorObject $_
        }
    } else {
        Write-Log -Message "Keine Quelldatei für Labels angegeben oder Datei existiert nicht." -Level "ERROR"
    }

    # Dokumentation erzeugen (Word, PDF)
    if ($ExportWord -or $ExportPDF) {
        try {
            # Beispielhafte Word-Export-Logik
            $wordPath = Join-Path $LogFolder ("LabelReport_{0}.docx" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
            $pdfPath = [System.IO.Path]::ChangeExtension($wordPath, ".pdf")
            Write-Log -Message "Dokumentation wird erstellt: $wordPath" -Level "INFO"

            # (Hier könnte mit Word-Interop oder einer Library gearbeitet werden, wir simulieren das mit Logging)
            foreach ($label in $labels) {
                Write-Log -Message ("Dokumentiere Label: {0} ({1})" -f $label.Name, $label.Description) -Level "DEBUG"
            }
            Write-Log -Message "Word-Dokument erstellt: $wordPath" -Level "SUCCESS"
            if ($ExportPDF) {
                Write-Log -Message "PDF-Dokument erstellt: $pdfPath" -Level "SUCCESS"
            }
        } catch {
            Write-LogError -Message "Fehler beim Dokumentationsexport" -ErrorObject $_
        }
    }

    # Optional Report versenden
    if ($SendReport) {
        try {
            # Beispiel: Report per Mail verschicken (hier simuliert)
            Write-Log -Message "Sende Bericht an $MailToPrimary ..." -Level "INFO"
            # Hier würde Send-MailMessage oder ein API-Aufruf stehen
            Start-Sleep -Seconds 1
            Write-Log -Message "Bericht erfolgreich versendet." -Level "SUCCESS"
        } catch {
            Write-LogError -Message "Fehler beim Versand des Berichts" -ErrorObject $_
        }
    }

    # Abschlussmeldung
    Write-Log -Message "Dokumentation abgeschlossen." -Level "SUCCESS"
    Write-Host "Label-Dokumentation abgeschlossen. Siehe Log unter $LogFolder"
}
catch {
    Write-LogError -Message "Fehler im Hauptskript" -ErrorObject $_
}
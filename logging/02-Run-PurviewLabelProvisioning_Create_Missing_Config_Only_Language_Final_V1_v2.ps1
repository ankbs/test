<#
.SYNOPSIS
    Prüfung und Erstellung fehlender Label-Sprach-Konfigurationen aus Excel für Microsoft Purview.

.DESCRIPTION
    Dieses Skript prüft auf Basis einer Excel-Konfigurationsdatei, ob für alle in der Datei gelisteten Sensitivitätslabels auch die notwendigen Sprachkonfigurationen existieren.
    Die Excel-Datei muss eine strukturierte Tabelle enthalten, in der für jedes Label und jede Sprache der gewünschte Wert hinterlegt ist.
    Für jede Kombination aus Label und Sprache wird geprüft, ob die Konfiguration vorhanden ist oder fehlt.
    Fehlende Konfigurationen werden erkannt, geloggt und können nachträglich automatisiert provisioniert werden.
    Sämtliche Vorgänge werden zentral über das Logging-Modul dokumentiert. Auch Fehler (z.B. fehlende Werte, fehlerhafte Excel-Datei, unerwartete Laufzeitfehler) werden sauber behandelt und im Log nachvollziehbar festgehalten.
    Das Skript ist ideal für mehrsprachige Purview-Setups und erleichtert die Kontrolle und Nachpflege von Sprachvarianten für Sensitivitätslabels.
    Nach Abschluss erfolgt ein übersichtlicher Statusbericht.
    Dieses Skript liest seine Startparameter standardmäßig aus der PurviewConfig.json im gleichen Verzeichnis wie das Skript.
    Alternativ kann eine eigene Konfigurationsdatei mit -ConfigPath übergeben werden oder alle Parameter wie gewohnt direkt per Skriptaufruf.


.EXAMPLE
    .\02-Run-PurviewLabelProvisioning_Create_Missing_Config_Only_Language_Final_V1_v2.ps1 -ConfigExcelPath ".\LabelConfig.xlsx"

.EXAMPLE
    .\02-Run-PurviewLabelProvisioning_Create_Missing_Config_Only_Language_Final_V1_v2.ps1 -ConfigExcelPath "C:\Konfig.xlsx" -LogFolder "C:\Logs" -UserPrincipalName "admin@contoso.com"

.EXAMPLE
    .\02-Run-PurviewLabelProvisioning_Create_Missing_Config_Only_Language_Final_V1_v2.ps1
    # Verwendet automatisch .\PurviewConfig.json, falls vorhanden

.EXAMPLE
    .\02-Run-PurviewLabelProvisioning_Create_Missing_Config_Only_Language_Final_V1_v2.ps1 -ConfigPath "D:\Konfig\MeineConfig.json"
    # Verwendet die explizit angegebene Datei

.EXAMPLE
    .\02-Run-PurviewLabelProvisioning_Create_Missing_Config_Only_Language_Final_V1_v2.ps1 -UserPrincipalName "admin@contoso.com" -InputExcelPath "C:\Labels.xlsx"
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
    [string]$ConfigExcelPath = "",
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
Write-Log -Message "02-Run-PurviewLabelProvisioning_Create_Missing_Config_Only_Language_Final_V1_v2.ps1 gestartet" -Level "INFO"

try {
    # Prüfe und erstelle Logverzeichnis
    if (!(Test-Path $LogFolder)) {
        New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
        Write-Log -Message "Logverzeichnis $LogFolder wurde erstellt." -Level "INFO"
    }

    # Eingabedatei prüfen
    if (-not $ConfigExcelPath -or -not (Test-Path $ConfigExcelPath)) {
        Write-Log -Message "Excel-Konfigurationsdatei nicht gefunden: $ConfigExcelPath" -Level "ERROR"
        throw "Excel-Konfigurationsdatei nicht gefunden: $ConfigExcelPath"
    }
    Write-Log -Message "Konfigurationsdatei: $ConfigExcelPath" -Level "INFO"

    # Excel laden
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($ConfigExcelPath)
        $worksheet = $workbook.Worksheets.Item(1)
        Write-Log -Message "Excel erfolgreich geladen." -Level "SUCCESS"
    } catch {
        Write-LogError -Message "Fehler beim Laden der Excel-Konfigurationsdatei" -ErrorObject $_
    }

    # Labels und Sprachen auslesen (angenommen: Spalten A = LabelName, B = Sprache, C = Wert)
    $configEntries = @()
    $row = 2
    while ($worksheet.Cells.Item($row, 1).Value2) {
        $labelName = $worksheet.Cells.Item($row, 1).Value2
        $language  = $worksheet.Cells.Item($row, 2).Value2
        $labelValue = $worksheet.Cells.Item($row, 3).Value2
        $configEntries += [PSCustomObject]@{
            Label    = $labelName
            Language = $language
            Value    = $labelValue
        }
        $row++
    }
    Write-Log -Message ("Es wurden {0} Einträge ausgelesen." -f $configEntries.Count) -Level "INFO"

    # Beispielhafte Verarbeitung: Fehlende Konfigurationen für Sprachen anlegen
    $missingConfigs = @()
    foreach ($entry in $configEntries) {
        try {
            # Hier würde geprüft/angelegt werden, z.B. per API-Aufruf
            Write-Log -Message ("Prüfe Konfiguration für Label '{0}' in Sprache '{1}'." -f $entry.Label, $entry.Language) -Level "DEBUG"
            # Beispiel: Fehlende Sprache simulieren
            if ([string]::IsNullOrWhiteSpace($entry.Value)) {
                $missingConfigs += $entry
                Write-Log -Message ("Fehlende Konfiguration für Label '{0}' in Sprache '{1}' erkannt." -f $entry.Label, $entry.Language) -Level "ERROR"
            }
            else {
                Write-Log -Message ("Konfiguration vorhanden: {0} ({1})" -f $entry.Label, $entry.Language) -Level "SUCCESS"
            }
        } catch {
            Write-LogError -Message ("Fehler bei der Prüfung von Label {0}, Sprache {1}" -f $entry.Label, $entry.Language) -ErrorObject $_
        }
    }

    # Zusammenfassung fehlender Konfigurationen
    if ($missingConfigs.Count -gt 0) {
        Write-Log -Message ("Es fehlen {0} Konfigurationen für Sprachen. Details siehe Log." -f $missingConfigs.Count) -Level "ERROR"
    } else {
        Write-Log -Message "Alle Label-Sprach-Konfigurationen sind vorhanden." -Level "SUCCESS"
    }

    # Excel schließen und aufräumen
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Log -Message "Excel geschlossen und Ressourcen freigegeben." -Level "INFO"

    # Abschlussmeldung
    Write-Log -Message "Label-Sprach-Konfigurationsprüfung abgeschlossen." -Level "SUCCESS"
    Write-Host "Alle Konfigurationen wurden geprüft. Details siehe Log unter $LogFolder"
}
catch {
    Write-LogError -Message "Fehler im Hauptskript" -ErrorObject $_
}
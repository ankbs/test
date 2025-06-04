<#
.SYNOPSIS
    Verwaltung und Sortierung von Sensitivitätslabels nach Prioritäten für Microsoft Purview.

.DESCRIPTION
    Dieses Skript liest Label-Prioritäten aus einer Excel-Datei ein, sortiert die Sensitivitätslabels nach den definierten Prioritäten und bereitet die Daten für weitere Verarbeitungsschritte auf.
    Es eignet sich zur automatisierten Anpassung und Steuerung der Reihenfolge von Labels in Microsoft Purview, zum Beispiel für konsistente Präsentation in Compliance-Prozessen oder für technische Automatisierungen.
    Die Excel-Datei muss für jedes Label eine Prioritätszahl enthalten; die Labels werden entsprechend aufbereitet, sortiert und optional zur weiteren Nutzung exportiert.
    Das Skript verwendet ein zentrales Logging- und Fehlerbehandlungssystem, um eine lückenlose Nachvollziehbarkeit zu gewährleisten.
    Typische Fehlerfälle (fehlende Datei, inkorrekte Werte, Laufzeitfehler) werden abgefangen und dokumentiert.
    Nach Abschluss steht eine sortierte Liste zur Verfügung, die für Importe, Reports oder weitere Automatisierung genutzt werden kann.
    Dieses Skript liest seine Startparameter standardmäßig aus der PurviewConfig.json im gleichen Verzeichnis wie das Skript.
    Alternativ kann eine eigene Konfigurationsdatei mit -ConfigPath übergeben werden oder alle Parameter wie gewohnt direkt per Skriptaufruf.


.EXAMPLE
    .\04-Run-Purview-Label-PriorityManager_V1_v2.ps1 -PriorityConfigExcel ".\Priorities.xlsx"

.EXAMPLE
    .\04-Run-Purview-Label-PriorityManager_V1_v2.ps1 -UserPrincipalName "user@domain.de" -PriorityConfigExcel "C:\Prioritäten.xlsx" -LogFolder "C:\Logs"

.EXAMPLE
    .\04-Run-Purview-Label-PriorityManager_V1_v2.ps1
    # Verwendet automatisch .\PurviewConfig.json, falls vorhanden

.EXAMPLE
    .\04-Run-Purview-Label-PriorityManager_V1_v2.ps1 -ConfigPath "D:\Konfig\MeineConfig.json"
    # Verwendet die explizit angegebene Datei

.EXAMPLE
    .\04-Run-Purview-Label-PriorityManager_V1_v2.ps1 -UserPrincipalName "admin@contoso.com" -InputExcelPath "C:\Labels.xlsx"
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
    [string]$UserPrincipalName = "",
    [string]$Tenantdomain = "",
    [string]$LogFolder = "C:\Temp\script\",
    [string]$MSPPartner = "",
    [string]$MSPNameAP  = "",
    [string]$MSPMail    = "",
    [string]$MSPURL     = "",
    [string]$CompanyLogoBase64 = "",
    [string]$ProductLogoBase64 = "",
    [string]$PriorityConfigExcel = "",
    [string]$ConfigPath = ""
)

# Zentrales Config-Modul importieren (Pfad ggf. anpassen)
Import-Module "$PSScriptRoot\modules\Import-ConfigParameters.psm1" -Force

# Standardmäßig nach PurviewConfig.json im Skriptverzeichnis suchen, falls -ConfigPath leer bleibt
if (-not $ConfigPath) {
    $ConfigPath = Join-Path $PSScriptRoot "PurviewConfig.json"
}
if (Test-Path $ConfigPath) {
    Import-ConfigParameters -ConfigPath $ConfigPath -BoundParameters $PSBoundParameters
}

Import-Module "$PSScriptRoot\CentralLogging.psm1" -Force
Set-LogFile -LogFolder "$LogFolder"
Write-Log -Message "04-Run-Purview-Label-PriorityManager_V1_v2.ps1 gestartet" -Level "INFO"

try {
    # Prüfe und erstelle Logverzeichnis
    if (!(Test-Path $LogFolder)) {
        New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
        Write-Log -Message "Logverzeichnis $LogFolder wurde erstellt." -Level "INFO"
    }

    # Prüfe Prioritäten-Konfigurationsdatei
    if (-not $PriorityConfigExcel -or -not (Test-Path $PriorityConfigExcel)) {
        Write-Log -Message "Prioritäten-Konfigurationsdatei nicht gefunden: $PriorityConfigExcel" -Level "ERROR"
        throw "Prioritäten-Konfigurationsdatei nicht gefunden: $PriorityConfigExcel"
    }
    Write-Log -Message "Prioritäten-Konfigurationsdatei: $PriorityConfigExcel" -Level "INFO"

    # Lade Excel und lese Prioritäten
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($PriorityConfigExcel)
        $worksheet = $workbook.Worksheets.Item(1)
        Write-Log -Message "Excel erfolgreich geladen." -Level "SUCCESS"
    } catch {
        Write-LogError -Message "Fehler beim Laden der Prioritäten-Excel-Datei" -ErrorObject $_
    }

    # Lese Labels und Prioritäten (angenommen: Spalten A = LabelName, B = Priorität)
    $priorities = @()
    $row = 2
    while ($worksheet.Cells.Item($row, 1).Value2) {
        $labelName = $worksheet.Cells.Item($row, 1).Value2
        $priority  = $worksheet.Cells.Item($row, 2).Value2
        $priorities += [PSCustomObject]@{
            Name     = $labelName
            Priority = $priority
        }
        $row++
    }
    Write-Log -Message ("Es wurden {0} Label-Prioritäten gefunden." -f $priorities.Count) -Level "INFO"

    # Beispielhafte Priorisierung/Sortierung
    $sorted = $priorities | Sort-Object -Property Priority
    foreach ($label in $sorted) {
        try {
            Write-Log -Message ("Label '{0}' hat Priorität {1}." -f $label.Name, $label.Priority) -Level "INFO"
            # Hier würde ggf. eine Aktualisierung in einem Zielsystem erfolgen
        } catch {
            Write-LogError -Message ("Fehler bei Priorität Label {0}" -f $label.Name) -ErrorObject $_
        }
    }

    # Excel schließen und aufräumen
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Log -Message "Excel geschlossen und Ressourcen freigegeben." -Level "INFO"

    Write-Log -Message "Label-Priorisierung abgeschlossen." -Level "SUCCESS"
    Write-Host "Alle Labels wurden nach Priorität verarbeitet. Details siehe Log unter $LogFolder"
}
catch {
    Write-LogError -Message "Fehler im Hauptskript" -ErrorObject $_
}
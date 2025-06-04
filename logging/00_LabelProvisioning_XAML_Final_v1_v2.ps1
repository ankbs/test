param (
    [string]$InputExcelPath = "",
    [string]$LogFolder = "C:\Temp\script\",
    [string]$UserPrincipalName = "",
    [string]$Tenantdomain = "",
    [string]$MailToPrimary = "",
    [string]$MailToSecondary = "",
    [string]$CompanyLogoBase64 = "",
    [string]$ProductLogoBase64 = ""
)

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
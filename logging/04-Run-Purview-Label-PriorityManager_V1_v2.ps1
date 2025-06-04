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
    [string]$PriorityConfigExcel = ""
)

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
        Handle-Error -Message "Fehler beim Laden der Prioritäten-Excel-Datei" -ErrorObject $_
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
            Handle-Error -Message ("Fehler bei Priorität Label {0}" -f $label.Name) -ErrorObject $_
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
    Handle-Error -Message "Fehler im Hauptskript" -ErrorObject $_
}
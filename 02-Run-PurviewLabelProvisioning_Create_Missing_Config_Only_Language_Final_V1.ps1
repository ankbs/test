<#
.SYNOPSIS
    Microsoft Purview Label Language Updater GUI – prüft, erstellt und aktualisiert Labels und Label-Sprachen.
.DESCRIPTION
    PowerShell 5-kompatibles WPF-Skript für die Verwaltung und Übersetzung von Purview Labels.
    - Import von Labeldaten aus Excel oder DeepL-JSON (mehrsprachig, Mapping für IPPS).
    - GUI mit Statusanzeige, Typ (P/C) editierbar, Refresh, Import, Export und Update-Button.
    - Parent/Child-Logik anhand der Typ-Spalte (P=Parent, C=Child).
    - Sprach-/Übersetzungs-Update übernimmt alle Sprachen korrekt in die IPPS-Labels.
    - Logging und Fehlerbehandlung.
    - Kürzt DisplayName auf max. 64 Zeichen und entfernt ggf. Klammerausdrücke.
    - Kein Import weiterer Parameter aus JSON/Excel außer Name, DisplayName, Tooltip und Übersetzungen.
#>

param(
    [string]$ExportFolder = "",
    [string]$LogFolder = "",
    [string]$UserPrincipalName = "",
    [string]$Tenantdomain = "",
    [string]$CompanyLogoBase64 = "",
    [string]$ProductLogoBase64 = ""
)

# === Modulprüfung und Import ===
function Install-RequiredModule {
    param([string]$ModuleName)
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "📦 Modul '$ModuleName' wird installiert..." -ForegroundColor Yellow
        Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $ModuleName
}
Install-RequiredModule -ModuleName "ExchangeOnlineManagement"
Install-RequiredModule -ModuleName "ImportExcel"

# === Verzeichnis-Handling für Logs und Exporte ===
$DatumHeute = Get-Date -Format 'yyyyMMdd_HHmmss'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ExportFolder -or $ExportFolder -eq "") { $ExportFolder = Join-Path $ScriptDir "Export_$DatumHeute" }
if (-not $LogFolder -or $LogFolder -eq "") { $LogFolder = Join-Path $ScriptDir "Logs_$DatumHeute" }
if (-not (Test-Path $LogFolder))    { New-Item -Path $LogFolder    -ItemType Directory -Force > $null }
if (-not (Test-Path $ExportFolder)) { New-Item -Path $ExportFolder -ItemType Directory -Force > $null }
$DatumJetzt = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogFile = Join-Path $LogFolder "PurviewLabelProvisioning_LOG_$DatumJetzt.log"
# $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path


# === Logging + Fehler-Handling ===
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $prefix = switch ($Level) {
        "INFO" { "ℹ️" }
        "SUCCESS" { "✅" }
        "ERROR" { "❌" }
        "DEBUG" { "🐛" }
        default { "🔹" }
    }
    $entry = "$timestamp $prefix $Message"
    Add-Content -Path $LogFile -Value $entry -Encoding utf8
    Write-Host $entry -Encoding utf8
}
function Write-LogError {
    param([string]$Message, [object]$ErrorObject)
    $msg = if ($ErrorObject -is [System.Management.Automation.ErrorRecord]) { $ErrorObject.Exception.Message } else { $ErrorObject.ToString() }
    $fullMsg = "${Message}: $msg"
    Log $fullMsg "ERROR"
    exit 1
}

# === MFA Connect ===
if (-not $UserPrincipalName) { $UserPrincipalName = Read-Host "🔑 Bitte geben Sie den UserPrincipalName ein" }
try {
    Connect-IPPSSession -UserPrincipalName $UserPrincipalName
    Write-Log "✅ IPPS verbunden" "SUCCESS"
} catch {
    Write-LogError "❌ IPPS Verbindung fehlgeschlagen" $_
}

# === Logos vorbereiten ===
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

function Create-ImageFromBase64 {
    param($Base64)
    try {
        $clean = $Base64 -replace '^data:image\/[a-z]+;base64,', ''
        $bytes = [Convert]::FromBase64String($clean)
        $ms = New-Object IO.MemoryStream (,[byte[]]$bytes)
        return [System.Drawing.Image]::FromStream($ms)
    } catch {
        Write-Log "❌ Fehler beim Konvertieren von Base64-Logo: $_" "ERROR"
        return $null
    }
}

$CompanyLogo = $null
if ($CompanyLogoBase64 -and $CompanyLogoBase64.Length -gt 100) {
    $CompanyLogo = Create-ImageFromBase64 $CompanyLogoBase64
}
$ProductLogo = $null
if ($ProductLogoBase64 -and $ProductLogoBase64.Length -gt 100) {
    $ProductLogo = Create-ImageFromBase64 $ProductLogoBase64
}

# === Label-Farbe auf Existenz prüfen (PowerShell 5.x kompatibel mit Add-Member) ===
function Check-Labels-Existence {
    param([System.Collections.IList]$LabelData)
    foreach ($label in $LabelData) {
        $LabelName = $label.Name
        $LabelOnline = Get-Label -Identity $LabelName -ErrorAction SilentlyContinue
        if ($LabelOnline) {
            $label | Add-Member -NotePropertyName StatusColor -NotePropertyValue "LightGreen" -Force
            $label | Add-Member -NotePropertyName StatusText  -NotePropertyValue "Vorhanden" -Force
            Write-Log "Label '$LabelName' existiert bereits." "SUCCESS"
        } else {
            $label | Add-Member -NotePropertyName StatusColor -NotePropertyValue "Salmon" -Force
            $label | Add-Member -NotePropertyName StatusText  -NotePropertyValue "Fehlt" -Force
            Write-Log "Label '$LabelName' fehlt." "INFO"
        }
    }
}

# === Funktion zum Erkennen von Parent/Child anhand Spalte "Typ" (editierbar in DataGrid)
function Detect-LabelType {
    param($Label)
    if ($Label.PSObject.Properties['Typ']) {
        return $Label.Typ
    }
    return "P"
}

# === DisplayName-Validierung/Kürzung ===
function Get-ValidDisplayName {
    param($text)
    if ($text.Length -le 64) { return $text }
    # Kürze Klammerausdruck weg
    $ohneKlammer = $text -replace '\s*\([^)]+\)', ''
    if ($ohneKlammer.Length -le 64) { return $ohneKlammer.Trim() }
    # Wenn immer noch zu lang, hart auf 64 Zeichen abschneiden
    return $ohneKlammer.Substring(0, 64).Trim()
}

# === Tooltip Helper ===
function Get-Tooltip($label) {
    if ($label.PSObject.Properties['Tool1'] -and $label.Tool1) { return $label.Tool1 }
    if ($label.PSObject.Properties['Tooltip'] -and $label.Tooltip) { return $label.Tooltip }
    if ($label.PSObject.Properties['ToolTip'] -and $label.ToolTip) { return $label.ToolTip }
    if ($label.PSObject.Properties['DisplayName'] -and $label.DisplayName) { return $label.DisplayName }
    return ""
}

# === Labels anlegen (nur Name, DisplayName und Tooltip) ===
function Create-Missing-Labels {
    param([System.Collections.IList]$LabelData)
    $created = 0

    # PARENT-Labels anlegen
    foreach ($label in $LabelData) {
        if ($label.StatusText -eq "Fehlt") {
            $LabelName = $label.Name
            $LabelType = Detect-LabelType $label

            if ($LabelType -eq "P") {
                try {
                    $validDisplayName = Get-ValidDisplayName $label.DisplayName
                    New-Label -Name $label.Name -DisplayName $validDisplayName -Tooltip (Get-Tooltip $label)
                    Write-Log "✅ Parent-Label '$LabelName' erstellt." "SUCCESS"
                    $label | Add-Member -NotePropertyName StatusColor -NotePropertyValue "LightGreen" -Force
                    $label | Add-Member -NotePropertyName StatusText  -NotePropertyValue "Vorhanden" -Force
                    $label.DisplayName = $validDisplayName
                    $created++
                } catch {
                    Write-Log "❌ Fehler beim Erstellen von Parent-Label '$LabelName': $($_.Exception.Message)" "ERROR"
                }
            }
        }
    }

    # CHILD-Labels anlegen (immer ParentId setzen, nie als Parent!)
    foreach ($label in $LabelData) {
        if ($label.StatusText -eq "Fehlt") {
            $LabelName = $label.Name
            $LabelType = Detect-LabelType $label
            if ($LabelType -eq "C") {
                $ParentName = ($LabelName -replace " \(unverschlüsselt\)", "" -replace " \(verschlüsselt\)", "").Trim()
                $ParentLabel = Get-Label -Identity $ParentName -ErrorAction SilentlyContinue
                if ($ParentLabel) {
                    $ParentId = $ParentLabel.ExchangeObjectId
                    try {
                        $validDisplayName = Get-ValidDisplayName $label.DisplayName
                        New-Label -Name $label.Name -DisplayName $validDisplayName -Tooltip (Get-Tooltip $label) -ParentId $ParentId
                        Write-Log "✅ Child-Label '$LabelName' erstellt (ParentId: $ParentId)." "SUCCESS"
                        $label | Add-Member -NotePropertyName StatusColor -NotePropertyValue "LightGreen" -Force
                        $label | Add-Member -NotePropertyName StatusText  -NotePropertyValue "Vorhanden" -Force
                        $label.DisplayName = $validDisplayName
                        $created++
                    } catch {
                        Write-Log "❌ Fehler beim Erstellen von Child-Label '$LabelName': $($_.Exception.Message)" "ERROR"
                    }
                } else {
                    Write-Log "❌ Parent-Label '$ParentName' nicht gefunden für Child-Label '$LabelName'." "ERROR"
                }
            }
        }
    }

    if ($created -gt 0) {
        [System.Windows.MessageBox]::Show("$created fehlende Labels wurden erfolgreich erstellt.", "Labels erstellt", "OK", "Info")
    } else {
        [System.Windows.MessageBox]::Show("Keine fehlenden Labels zu erstellen.", "Info", "OK", "Info")
    }
    # Nach Anlage neu prüfen und markieren und DataGrid refreshen:
    Check-Labels-Existence -LabelData $LabelData
    $dgLabelData.ItemsSource = $null
    $dgLabelData.ItemsSource = $LabelData
}

# === Sprach-/Übersetzungs-Update (nur Name, DisplayName, Tooltip, Übersetzungen) ===
function Start-LabelUpdate {
    $labelsWithMissingTranslations = @()
    $langOrder = @('de-de','en-us','fr-fr','it-it','hr-hr','pl-pl','ro-ro','sk-sk','cs-cz','uk-ua','hu-hu','pt-br')
    foreach ($label in $global:ExcelData) {
        $LabelName = $label.Display1
        $Languages = @()
        $DisplayNames = @()
        $Tooltips = @()
        $labelHadMissingTranslations = $false

        for ($i=1; $i -le 12; $i++) {
            $lang = $label.("Language$i")
            $disp = $label.("Display$i")
            $tip = $label.("Tool$i")
            if ($lang -and (-not $disp -or -not $tip)) {
                $labelHadMissingTranslations = $true
                continue 2 # Überspringe dieses Label, keine Übersetzung möglich
            }
            if ($lang -and $disp -and $tip) {
                $validDisp = Get-ValidDisplayName $disp
                $Languages += $lang
                $DisplayNames += $validDisp
                $Tooltips += $tip
                # Schreibe korrigierten Wert zurück (wichtig für Export)
                $label.("Display$i") = $validDisp
            }
        }

        if ($labelHadMissingTranslations) {
            $labelsWithMissingTranslations += $LabelName
        }

        if ($Languages.Count -eq 0) {
            Write-Log "❌ Keine gültigen Übersetzungen für Label '$LabelName'. Überspringe..." "ERROR"
            continue
        }

        $DisplayNameLocaleSettings = [PSCustomObject]@{ LocaleKey = 'DisplayName'; Settings = @() }
        $TooltipLocaleSettings = [PSCustomObject]@{ LocaleKey = 'Tooltip'; Settings = @() }
        for ($j=0; $j -lt $Languages.Count; $j++) {
            $DisplayNameLocaleSettings.Settings += @{ key = $Languages[$j]; value = $DisplayNames[$j] }
            $TooltipLocaleSettings.Settings += @{ key = $Languages[$j]; value = $Tooltips[$j] }
        }

        try {
            Set-Label -Identity $LabelName `
                -Tooltip $Tooltips[0] `
                -LocaleSettings (
                    (ConvertTo-Json $DisplayNameLocaleSettings -Depth 4 -Compress),
                    (ConvertTo-Json $TooltipLocaleSettings -Depth 4 -Compress)
                )
            Write-Log "✅ Label '$LabelName' Sprachen aktualisiert." "SUCCESS"
        } catch {
            Write-Log "❌ Fehler beim Update '$LabelName': $($_.Exception.Message)" "ERROR"
        }
        Start-Sleep -s 2
    }

    # Export CSV/JSON
    $exportCsv = Join-Path $LogFolder "Translated_Labels_$($DatumJetzt).csv"
    $exportJson = Join-Path $LogFolder "Translated_Labels_$($DatumJetzt).json"
    $global:ExcelData | Export-Csv -Path $exportCsv -Delimiter ";" -Encoding UTF8 -NoTypeInformation
    $global:ExcelData | ConvertTo-Json -Depth 5 | Out-File -FilePath $exportJson -Encoding utf8

    if ($labelsWithMissingTranslations.Count -gt 0) {
        Write-Log "✅ Übersetzungen ergänzt & exportiert: CSV=$exportCsv, JSON=$exportJson" "SUCCESS"
        Write-Log "ℹ️ Labels mit fehlenden Übersetzungen (vorher): $($labelsWithMissingTranslations -join ', ')" "INFO"
        [System.Windows.MessageBox]::Show("⚠️ Einige Labels hatten fehlende Übersetzungen und wurden ergänzt!\n\nDateien:\n- $exportCsv\n- $exportJson", "Übersetzungen ergänzt", "OK", "Warning")
    } else {
        Write-Log "✅ Keine fehlenden Übersetzungen – alles aktuell." "SUCCESS"
        [System.Windows.MessageBox]::Show("✅ Keine fehlenden Übersetzungen – alles aktuell.", "Info", "OK", "Info")
    }
}

# === DeepL-JSON-Import: Mapping zu ExcelData-Struktur (nur Name, DisplayName, Tooltip, Übersetzungen) ===
function Load-LabelData-FromDeepLJson ($jsonFile) {
    $langOrder = @('de-de','en-us','fr-fr','it-it','hr-hr','pl-pl','ro-ro','sk-sk','cs-cz','uk-ua','hu-hu','pt-br')
    try {
        $jsonRaw = Get-Content -Raw -Path $jsonFile -Encoding UTF8
        if ($jsonRaw -and $jsonRaw.Trim().Length -gt 0) {
            $jsonData = $jsonRaw | ConvertFrom-Json
            if ($null -eq $jsonData -or $jsonData.Count -eq 0) {
                Write-Log "❌ DeepL JSON-Datei leer oder fehlerhaft!" "ERROR"
                [System.Windows.MessageBox]::Show("❌ DeepL JSON-Datei leer oder fehlerhaft!", "Fehler", "OK", "Error")
                return $false
            }
            $result = @()
            foreach ($item in $jsonData) {
                $obj = [PSCustomObject]@{}
                if ($item.PSObject.Properties['Name']) {
                    $obj | Add-Member -NotePropertyName Name -NotePropertyValue $item.Name
                }
                for ($i=0; $i -lt $langOrder.Count; $i++) {
                    $lang = $langOrder[$i]
                    $langCol = "Language$([int]($i+1))"
                    $dispCol = "Display$([int]($i+1))"
                    $toolCol = "Tool$([int]($i+1))"
                    $obj | Add-Member -NotePropertyName $langCol -NotePropertyValue $lang
                    if ($i -eq 0) {
                        # Standardwert ohne Sprachcode
                        $obj | Add-Member -NotePropertyName $dispCol -NotePropertyValue ($item.DisplayName)
                        $obj | Add-Member -NotePropertyName $toolCol -NotePropertyValue ($item.Tooltip)
                    } else {
                        $obj | Add-Member -NotePropertyName $dispCol -NotePropertyValue ($item."DisplayName_$lang")
                        $obj | Add-Member -NotePropertyName $toolCol -NotePropertyValue ($item."Tooltip_$lang")
                    }
                }
                # Typ-Spalte initialisieren falls nicht vorhanden
                if (-not $obj.PSObject.Properties['Typ']) {
                    if ($obj.Name -match "\(unverschlüsselt\)" -or $obj.Name -match "\(verschlüsselt\)") {
                        $obj | Add-Member -NotePropertyName Typ -NotePropertyValue "C" -Force
                    } else {
                        $obj | Add-Member -NotePropertyName Typ -NotePropertyValue "P" -Force
                    }
                }
                # DisplayName als Property für DataGrid
                $obj | Add-Member -NotePropertyName DisplayName -NotePropertyValue ($item.DisplayName)
                $obj | Add-Member -NotePropertyName Tooltip -NotePropertyValue ($item.Tooltip)
                $result += $obj
            }
            $global:ExcelData = $result
            Check-Labels-Existence -LabelData $global:ExcelData
            $dgLabelData.ItemsSource = $null
            $dgLabelData.ItemsSource = $global:ExcelData
            Write-Log "📂 DeepL JSON geladen und gemappt: $jsonFile" "INFO"
            return $true
        } else {
            Write-Log "❌ DeepL JSON-Datei leer oder fehlerhaft!" "ERROR"
            [System.Windows.MessageBox]::Show("❌ DeepL JSON-Datei leer oder fehlerhaft!", "Fehler", "OK", "Error")
            return $false
        }
    } catch {
        Write-Log "❌ Fehler beim Einlesen der DeepL-JSON-Datei: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show("❌ Fehler beim Einlesen der DeepL-JSON-Datei!", "Fehler", "OK", "Error")
        return $false
    }
}

# === XAML GUI (PowerShell 5-kompatibel, keine x:Array etc.) ===
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Label Updater" Height="700" Width="1400">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header mit Logos -->
        <DockPanel Grid.Row="0" LastChildFill="False" Margin="0,0,0,10">
            <Image Name="imgCompanyLogo" Width="48" Height="48" Margin="0,0,10,0" />
            <TextBlock Text="Microsoft Purview Automation" FontWeight="Bold" FontSize="20" VerticalAlignment="Center" />
            <Image Name="imgProductLogo" Width="48" Height="48" Margin="10,0,0,0" HorizontalAlignment="Right" />
        </DockPanel>

        <!-- DataGrid -->
        <DataGrid Name="dgLabelData" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="False" RowHeight="32" HeadersVisibility="Column" CanUserAddRows="False">
            <DataGrid.Resources>
                <Style TargetType="DataGridRow">
                    <Setter Property="Background" Value="{Binding StatusColor}" />
                </Style>
            </DataGrid.Resources>
            <DataGrid.Columns>
                <DataGridTextColumn Header="Typ (P/C)" Binding="{Binding Typ, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="60"/>
                <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="*"/>
                <DataGridTextColumn Header="DisplayName" Binding="{Binding DisplayName}" Width="*"/>
                <DataGridTextColumn Header="Tooltip" Binding="{Binding Tooltip}" Width="*"/>
                <DataGridTextColumn Header="Status" Binding="{Binding StatusText}" Width="Auto"/>
            </DataGrid.Columns>
        </DataGrid>

        <!-- Button-Reihe -->
        <StackPanel Name="spButtons" Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Left" Margin="0,10,0,0">
            <Button Name="btnRefresh" Content="Refresh" Width="120" Margin="0,0,10,0"/>
            <Button Name="btnLoadJson" Content="Daten Laden" Width="130" Margin="0,0,10,0"/>
            <Button Name="btnCreateMissingLabels" Content="Fehlende Labels erstellen" Width="200" Margin="0,0,10,0"/>
            <Button Name="btnLoadDeepLJson" Content="DeepL Sprachen JSON laden" Width="200" Margin="0,0,10,0"/>
            <Button Name="btnLoadExcel" Content="Excel Import" Width="130" Margin="0,0,10,0"/>
            <Button Name="btnUpdateLanguages" Content="Sprachen/Übersetzungen aktualisieren" Width="220" Margin="0,0,10,0"/>
            <Button Name="btnCreateReport" Content="Report(s) erstellen" Width="160" Margin="0,0,10,0"/>
            <Button Name="btnClose" Content="Abbrechen" Width="120"/>
        </StackPanel>

        <!-- Footer -->
        <StackPanel Grid.Row="3" Orientation="Vertical" VerticalAlignment="Bottom" Margin="0,10,0,0">
            <TextBlock Text="Bei Fragen wenden Sie sich an:" FontSize="10"/>
            <TextBlock Name="tbFooterContact" FontSize="10"/>
        </StackPanel>
    </Grid>
</Window>
"@

# === XAML laden und Controls referenzieren ===
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
$dgLabelData = $window.FindName("dgLabelData")
$btnRefresh = $window.FindName("btnRefresh")
$btnLoadJson = $window.FindName("btnLoadJson")
$btnLoadExcel = $window.FindName("btnLoadExcel")
$btnLoadDeepLJson = $window.FindName("btnLoadDeepLJson")
$btnCreateMissingLabels = $window.FindName("btnCreateMissingLabels")
$btnUpdateLanguages = $window.FindName("btnUpdateLanguages")
$btnClose = $window.FindName("btnClose")
$imgCompanyLogo = $window.FindName("imgCompanyLogo")
$imgProductLogo = $window.FindName("imgProductLogo")
$tbFooterContact = $window.FindName("tbFooterContact")
$btnCreateReport = $window.FindName("btnCreateReport")

# Logos setzen
if ($CompanyLogo) {
    $imgCompanyLogo.Source = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHBitmap(
        $CompanyLogo.GetHbitmap(),
        [IntPtr]::Zero,
        ([System.Windows.Int32Rect]::Empty),
        [System.Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions()
    )
}
if ($ProductLogo) {
    $imgProductLogo.Source = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHBitmap(
        $ProductLogo.GetHbitmap(),
        [IntPtr]::Zero,
        ([System.Windows.Int32Rect]::Empty),
        [System.Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions()
    )
}
$tbFooterContact.Text = "Cloud Security & Compliance Services, Michael Kirst-Neshva"

# === Globale Variable für Daten ===
$global:ExcelData = @()

# === Automatisch letzte JSON aus Exportfolder beim Start laden ===
function Get-LastExportJson {
    $jsonFiles = Get-ChildItem -Path $ExportFolder -Filter "provision_labels*.json" | Sort-Object LastWriteTime -Descending
    if ($jsonFiles.Count -gt 0) { return $jsonFiles[0].FullName }
    else { return $null }
}

function Load-LabelData-FromJson ($jsonFile) {
    try {
        $jsonRaw = Get-Content -Raw -Path $jsonFile -Encoding UTF8
        if ($jsonRaw -and $jsonRaw.Trim().Length -gt 0) {
            $jsonData = $jsonRaw | ConvertFrom-Json
            if ($null -eq $jsonData) {
                Write-Log "❌ JSON konnte nicht konvertiert werden (null)" "ERROR"
                [System.Windows.MessageBox]::Show("❌ JSON konnte nicht konvertiert werden!", "Fehler", "OK", "Error")
                return $false
            } elseif ($jsonData.Count -eq 0) {
                Write-Log "❌ JSON-Datei ist leer oder enthält keine Labels" "ERROR"
                [System.Windows.MessageBox]::Show("❌ JSON-Datei ist leer oder enthält keine Labels!", "Fehler", "OK", "Error")
                return $false
            } else {
                if ($jsonData -isnot [System.Collections.IEnumerable]) {
                    $global:ExcelData = @($jsonData)
                } else {
                    $global:ExcelData = $jsonData
                }
                foreach ($label in $global:ExcelData) {
                    if (-not $label.PSObject.Properties['Typ']) {
                        if ($label.Name -match "\(unverschlüsselt\)" -or $label.Name -match "\(verschlüsselt\)") {
                            $label | Add-Member -NotePropertyName Typ -NotePropertyValue "C" -Force
                        } else {
                            $label | Add-Member -NotePropertyName Typ -NotePropertyValue "P" -Force
                        }
                    }
                }
                Check-Labels-Existence -LabelData $global:ExcelData
                $dgLabelData.ItemsSource = $null
                $dgLabelData.ItemsSource = $global:ExcelData
                Write-Log "📂 JSON geladen: $jsonFile" "INFO"
                return $true
            }
        } else {
            Write-Log "❌ JSON-Datei leer oder fehlerhaft!" "ERROR"
            [System.Windows.MessageBox]::Show("❌ JSON-Datei leer oder fehlerhaft!", "Fehler", "OK", "Error")
            return $false
        }
    } catch {
        Write-Log "❌ Fehler beim Einlesen der JSON-Datei: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show("❌ Fehler beim Einlesen der JSON-Datei!", "Fehler", "OK", "Error")
        return $false
    }
}

# === Event Handler für Buttons ===

# "Refresh"
$btnRefresh.Add_Click({
    if ($global:ExcelData.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Keine Daten geladen.", "Fehler", "OK", "Error")
        return
    }
    Check-Labels-Existence -LabelData $global:ExcelData
    $dgLabelData.ItemsSource = $null
    $dgLabelData.ItemsSource = $global:ExcelData
})

# "Daten Laden" (JSON laden)
$btnLoadJson.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "JSON Dateien (*.json)|*.json"
    $ofd.InitialDirectory = $ExportFolder
    if ($ofd.ShowDialog() -eq 'OK') {
        Load-LabelData-FromJson $ofd.FileName | Out-Null
    }
})

# "Fehlende Labels erstellen"
$btnCreateMissingLabels.Add_Click({
    if ($global:ExcelData.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Keine Daten geladen.", "Fehler", "OK", "Error")
        return
    }
    Create-Missing-Labels -LabelData $global:ExcelData
    Check-Labels-Existence -LabelData $global:ExcelData
    $dgLabelData.ItemsSource = $null
    $dgLabelData.ItemsSource = $global:ExcelData
})

# "DeepL JSON Import"
$btnLoadDeepLJson.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "JSON Dateien (*.json)|*.json"
    $ofd.InitialDirectory = $ExportFolder
    if ($ofd.ShowDialog() -eq 'OK') {
        if (Load-LabelData-FromDeepLJson $ofd.FileName) {
            $msg = "DeepL Übersetzungsdatei wurde erfolgreich geladen.`n`nMöchten Sie jetzt die Übersetzungen sofort auf die Labels anwenden?"
            $res = [System.Windows.MessageBox]::Show($msg, "DeepL JSON geladen", [System.Windows.MessageBoxButton]::OKCancel, [System.Windows.MessageBoxImage]::Information)
            if ($res -eq [System.Windows.MessageBoxResult]::OK) {
                Start-LabelUpdate
            }
        }
    }
})


# "Excel Import"
$btnLoadExcel.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Excel-Dateien (*.xls;*.xlsx)|*.xls;*.xlsx"
    if ($ofd.ShowDialog() -eq 'OK') {
        try {
            $excelData = Import-Excel -Path $ofd.FileName
            $global:ExcelData = $excelData | ForEach-Object {
                if (-not $_.PSObject.Properties['Typ']) {
                    if ($_.Name -match "\(unverschlüsselt\)" -or $_.Name -match "\(verschlüsselt\)") {
                        $_ | Add-Member -NotePropertyName Typ -NotePropertyValue "C" -Force
                    } else {
                        $_ | Add-Member -NotePropertyName Typ -NotePropertyValue "P" -Force
                    }
                }
                $_ | Add-Member -NotePropertyName StatusColor -NotePropertyValue "White" -Force
                $_ | Add-Member -NotePropertyName StatusText  -NotePropertyValue "" -Force
                $_
            }
            Check-Labels-Existence -LabelData $global:ExcelData
            $dgLabelData.ItemsSource = $null
            $dgLabelData.ItemsSource = $global:ExcelData
            Write-Log "📂 Excel geladen: $($ofd.FileName)" "INFO"
        } catch {
            Write-Log "❌ Fehler beim Excel-Import: $($_.Exception.Message)" "ERROR"
            [System.Windows.MessageBox]::Show("❌ Fehler beim Excel-Import!", "Fehler", "OK", "Error")
        }
    }
})

# "Sprachen/Übersetzungen aktualisieren"
$btnUpdateLanguages.Add_Click({
    if ($global:ExcelData.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Keine Daten geladen.", "Fehler", "OK", "Error")
        return
    }
    Start-LabelUpdate
})

# Reports erstellen


$btnCreateReport.Add_Click({
    $mainScript = ".\03-Run-Purview-Create-Documentation_GUI_Final_V10.ps1"
    # $arguments = @("-ExecutionPolicy Bypass", "-STA", "-File `"$mainScript`"", "-GuiConfigPath `"$GuiConfigPath`" -UserPrincipalName `"$UserPrincipalName`" -Tenantdomain `"$Tenantdomain`"") -join " "
    $arguments = @("-ExecutionPolicy Bypass", "-STA", "-File `"$mainScript`"", "-UserPrincipalName `"$UserPrincipalName`" -Tenantdomain `"$Tenantdomain`"") -join " "
    Write-Host "Starte: $arguments" -ForegroundColor Cyan
    # Add-Content -Path $DeepLLog -Value ("[{0}] [INFO] Button 'Neue Labels erstellen' gedrückt, Script: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $scriptName)
    try {
        Start-Process powershell.exe -ArgumentList $arguments -WindowStyle Normal
        $window.Close()
    } catch {
        [System.Windows.MessageBox]::Show("Fehler beim Starten von $scriptName`n$_")
    }
})


# "Abbrechen"
$btnClose.Add_Click({ $window.Close() })

# === Automatisch letzte Export-JSON beim Start laden ===
$lastJson = Get-LastExportJson
if ($lastJson) { Load-LabelData-FromJson $lastJson | Out-Null }

$window.ShowDialog() | Out-Null
Write-Log "⚡ Script wurde beendet." "INFO"
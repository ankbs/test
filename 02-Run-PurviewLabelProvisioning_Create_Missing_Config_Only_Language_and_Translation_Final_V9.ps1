# 02-Run-PurviewLabelProvisioning_Only_Language_GUI_Version_Final.ps1
param(
    [string]$LogFolder = "C:\Temp\PurviewExport",
    [string]$UserPrincipalName = "",
    [string]$ExcelPath = "C:\Temp\labels.xlsx",
    [string]$DeepLApiKey = ""
    # [string]$DeepLApiKey = "b7941202-f333-4f6e-ac28-3929ef648a2f:fx"

)

# === Modulprüfung und Import ===
function Ensure-Module {
    param([string]$ModuleName)
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "📦 Modul '$ModuleName' wird installiert..." -ForegroundColor Yellow
        Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $ModuleName
}
Ensure-Module -ModuleName "ExchangeOnlineManagement"
Ensure-Module -ModuleName "ImportExcel"

# === Logging ===
if (!(Test-Path $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null }
$DatumJetzt = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogFile = Join-Path $LogFolder "PurviewLabelProvisioning_LOG_$DatumJetzt.log"

function Log {
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
function Handle-Error {
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
    Log "✅ IPPS verbunden" "SUCCESS"
} catch {
    Handle-Error "❌ IPPS Verbindung fehlgeschlagen" $_
}

# === DeepL Übersetzungsfunktion (UTF-8 abgesichert) ===
function Translate-DeepL {
    param($Text, $TargetLang)
    if (-not $DeepLApiKey) { return $Text }
    try {
        $responseBytes = Invoke-WebRequest -Uri "https://api-free.deepl.com/v2/translate" `
            -Method Post `
            -Headers @{ "Authorization" = "DeepL-Auth-Key $DeepLApiKey" } `
            -Body @{
                text = $Text
                target_lang = $TargetLang
            } `
            -OutFile ([System.IO.Path]::GetTempFileName())

        $utf8String = Get-Content -Path $responseBytes -Raw -Encoding UTF8 | ConvertFrom-Json
        return $utf8String.translations[0].text
    } catch {
        Log "❌ Fehler bei der Übersetzung: $_" "ERROR"
        return $Text
    }
}

# === Funktion: Labels prüfen & erstellen ===
function Ensure-Labels-Existence {
    param([System.Collections.IList]$ExcelData)

    foreach ($label in $ExcelData) {
        $LabelName = $label.Display1.Trim()
        $Level = $label.Level

        $LabelOnline = Get-Label -Identity $LabelName -ErrorAction SilentlyContinue

        if ($LabelOnline -and $Level -eq "P") {
            Log "ℹ️ Parent-Label '$LabelName' bereits vorhanden." "INFO"
        }

        if (-not $LabelOnline -and $Level -eq "P") {
            $choice = [System.Windows.MessageBox]::Show("Parent-Label '$LabelName' nicht gefunden. Erstellen?", "Frage", "YesNo", "Question")
            if ($choice -eq "Yes") {
                try {
                    if ($LabelName -like "*(verschlüsselt)*" -or $LabelName -like "*Streng Vertraulich*") {
                        $MIPRightsA = $label.MIPIdentity1, $label.MIPRights1 -join ":"
                        $MIPRightsB = $label.MIPIdentity2, $label.MIPRights2 -join ":"
                        $MIPRightsC = $label.MIPIdentity3, $label.MIPRights3 -join ":"
                        $MIPRightsR = $MIPRightsA, $MIPRightsB, $MIPRightsC -join ";"

                        New-Label -Name $LabelName -DisplayName $LabelName -Tooltip $label.Tool1 `
                            -EncryptionContentExpiredOnDateInDaysOrNever $label.EncryptionContentExpiredOnDateInDaysOrNever `
                            -EncryptionDoNotForward $label.EncryptionDoNotForward `
                            -EncryptionEnabled $label.EncryptionEnabled `
                            -EncryptionEncryptOnly $label.EncryptionEncryptOnly `
                            -EncryptionOfflineAccessDays $label.EncryptionOfflineAccessDays `
                            -EncryptionPromptUser $label.EncryptionPromptUser `
                            -EncryptionProtectionType $label.EncryptionProtectionType `
                            -EncryptionRightsDefinitions $MIPRightsR `
                            -ApplyContentMarkingFooterAlignment Left `
                            -ApplyContentMarkingFooterEnabled $label.ApplyContentMarkingFooterEnabled `
                            -ApplyContentMarkingFooterFontColor $label.ApplyContentMarkingFooterFontColor `
                            -ApplyContentMarkingFooterFontSize $label.ApplyContentMarkingFooterFontSize `
                            -ApplyContentMarkingFooterMargin $label.ApplyContentMarkingFooterMargin `
                            -ApplyContentMarkingFooterText $label.ApplyContentMarkingFooterText `
                            -ApplyContentMarkingHeaderAlignment Center `
                            -ApplyContentMarkingHeaderEnabled $label.ApplyContentMarkingHeaderEnabled `
                            -ApplyContentMarkingHeaderFontColor $label.ApplyContentMarkingHeaderFontColor `
                            -ApplyContentMarkingHeaderFontSize $label.ApplyContentMarkingHeaderFontSize `
                            -ApplyContentMarkingHeaderMargin $label.ApplyContentMarkingHeaderMargin `
                            -ApplyContentMarkingHeaderText $label.ApplyContentMarkingHeaderText `
                            -ApplyWaterMarkingEnabled $label.ApplyWaterMarkingEnabled `
                            -ApplyWaterMarkingFontColor $label.ApplyWaterMarkingFontColor `
                            -ApplyWaterMarkingFontSize $label.ApplyWaterMarkingFontSize `
                            -ApplyWaterMarkingLayout $label.ApplyWaterMarkingLayout `
                            -ApplyWaterMarkingText $label.ApplyWaterMarkingText
                    } else {
                        New-Label -Name $LabelName -DisplayName $LabelName -Tooltip $label.Tool1
                    }
                    Log "✅ Parent-Label '$LabelName' erstellt." "SUCCESS"
                } catch {
                    Log "❌ Fehler beim Erstellen von '$LabelName': $($_.Exception.Message)" "ERROR"
                }
            }
        }

        if ($Level -eq "C") {
            $ParentName = ($LabelName -replace " \(unverschlüsselt\)", "" -replace " \(verschlüsselt\)", "").Trim()
            $ParentLabel = Get-Label -Identity $ParentName -ErrorAction SilentlyContinue

            if ($ParentLabel) {
                $ParentId = $ParentLabel.ExchangeObjectId
                $ChildOnline = Get-Label -Identity $LabelName -ErrorAction SilentlyContinue

                if (-not $ChildOnline) {
                    $choice = [System.Windows.MessageBox]::Show("Child-Label '$LabelName' nicht gefunden. Erstellen?", "Frage", "YesNo", "Question")
                    if ($choice -eq "Yes") {
                        try {
                            if ($LabelName -like "*(verschlüsselt)*" -or $LabelName -like "*Streng Vertraulich*") {
                                $MIPRightsA = $label.MIPIdentity1, $label.MIPRights1 -join ":"
                                $MIPRightsB = $label.MIPIdentity2, $label.MIPRights2 -join ":"
                                $MIPRightsC = $label.MIPIdentity3, $label.MIPRights3 -join ":"
                                $MIPRightsR = $MIPRightsA, $MIPRightsB, $MIPRightsC -join ";"

                                New-Label -Name $LabelName -DisplayName $LabelName -Tooltip $label.Tool1 -ParentId $ParentId `
                                    -EncryptionContentExpiredOnDateInDaysOrNever $label.EncryptionContentExpiredOnDateInDaysOrNever `
                                    -EncryptionDoNotForward $label.EncryptionDoNotForward `
                                    -EncryptionEnabled $label.EncryptionEnabled `
                                    -EncryptionEncryptOnly $label.EncryptionEncryptOnly `
                                    -EncryptionOfflineAccessDays $label.EncryptionOfflineAccessDays `
                                    -EncryptionPromptUser $label.EncryptionPromptUser `
                                    -EncryptionProtectionType $label.EncryptionProtectionType `
                                    -EncryptionRightsDefinitions $MIPRightsR `
                                    -ApplyContentMarkingFooterAlignment Left `
                                    -ApplyContentMarkingFooterEnabled $label.ApplyContentMarkingFooterEnabled `
                                    -ApplyContentMarkingFooterFontColor $label.ApplyContentMarkingFooterFontColor `
                                    -ApplyContentMarkingFooterFontSize $label.ApplyContentMarkingFooterFontSize `
                                    -ApplyContentMarkingFooterMargin $label.ApplyContentMarkingFooterMargin `
                                    -ApplyContentMarkingFooterText $label.ApplyContentMarkingFooterText `
                                    -ApplyContentMarkingHeaderAlignment Center `
                                    -ApplyContentMarkingHeaderEnabled $label.ApplyContentMarkingHeaderEnabled `
                                    -ApplyContentMarkingHeaderFontColor $label.ApplyContentMarkingHeaderFontColor `
                                    -ApplyContentMarkingHeaderFontSize $label.ApplyContentMarkingHeaderFontSize `
                                    -ApplyContentMarkingHeaderMargin $label.ApplyContentMarkingHeaderMargin `
                                    -ApplyContentMarkingHeaderText $label.ApplyContentMarkingHeaderText `
                                    -ApplyWaterMarkingEnabled $label.ApplyWaterMarkingEnabled `
                                    -ApplyWaterMarkingFontColor $label.ApplyWaterMarkingFontColor `
                                    -ApplyWaterMarkingFontSize $label.ApplyWaterMarkingFontSize `
                                    -ApplyWaterMarkingLayout $label.ApplyWaterMarkingLayout `
                                    -ApplyWaterMarkingText $label.ApplyWaterMarkingText
                            } else {
                                New-Label -Name $LabelName -DisplayName $LabelName -Tooltip $label.Tool1 -ParentId $ParentId
                            }
                            Log "✅ Child-Label '$LabelName' erstellt (ParentId: $ParentId)." "SUCCESS"
                        } catch {
                            Log "❌ Fehler beim Erstellen von '$LabelName': $($_.Exception.Message)" "ERROR"
                        }
                    }
                } else {
                    Log "ℹ️ Child-Label '$LabelName' bereits vorhanden." "INFO"
                }
            } else {
                Log "❌ Parent-Label '$ParentName' nicht gefunden für Child-Label '$LabelName'." "ERROR"
            }
        }
    }

    # Zusätzlicher Log- und Konsolenhinweis nach Abschluss der Prüfung aller Labels:
    Log "✅ Alle Labels überprüft. Nun kann mit der Übersetzung der Sprachen fortgefahren werden. Oder Abbrechen" "INFO"
    Write-Host "✅ Alle Labels überprüft. Nun kann mit der Übersetzung der Sprachen fortgefahren werden. Oder Abbrechen" -ForegroundColor Green
}


# === Label-Update (Sprachen) ===
# === Label-Update (Sprachen + Übersetzungen + Encryption-Parameter bei "(verschlüsselt)") ===
function Start-LabelUpdate {
    $labelsWithMissingTranslations = @()
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
                if (-not $DeepLApiKey) {
                    Log "❌ Fehlende Übersetzung bei Label '$LabelName' ($lang). Kein API-Key – Überspringe..." "ERROR"
                    continue 2
                }
                try {
                    $disp = Translate-DeepL -Text $label.Display1 -TargetLang $lang
                    $toolBase = "$($label.Display1) - ($($label.MIPCompanyName))"
                    $tip = Translate-DeepL -Text $toolBase -TargetLang $lang
                    $label.("Display$i") = $disp
                    $label.("Tool$i") = $tip
                } catch {
                    Log "❌ Fehler bei der Übersetzung für Label '$LabelName' ($lang): $($_.Exception.Message)" "ERROR"
                    continue 2
                }
            }
            if ($lang -and $disp -and $tip) {
                $Languages += $lang
                $DisplayNames += $disp
                $Tooltips += $tip
            }
        }

        if ($labelHadMissingTranslations) {
            $labelsWithMissingTranslations += $LabelName
        }

        if ($Languages.Count -eq 0) {
            Log "❌ Keine gültigen Übersetzungen für Label '$LabelName'. Überspringe..." "ERROR"
            continue
        }

        $DisplayNameLocaleSettings = [PSCustomObject]@{ LocaleKey = 'DisplayName'; Settings = @() }
        $TooltipLocaleSettings = [PSCustomObject]@{ LocaleKey = 'Tooltip'; Settings = @() }
        for ($j=0; $j -lt $Languages.Count; $j++) {
            $DisplayNameLocaleSettings.Settings += @{ key = $Languages[$j]; value = $DisplayNames[$j] }
            $TooltipLocaleSettings.Settings += @{ key = $Languages[$j]; value = $Tooltips[$j] }
        }

        try {
            # Log "🔍 DisplayNames=$($DisplayNames -join ', ')"
            # Log "🔍 Tooltips=$($Tooltips -join ', ')"

            # === Encryption-Parameter bei "(verschlüsselt)" oder "Streng vertraulich" Labels ===
            if ($LabelName -like "*(verschlüsselt)*" -or $LabelName -like "*Streng Vertraulich*") {

                $MIPRightsA = $label.MIPIdentity1, $label.MIPRights1 -join ":"
                $MIPRightsB = $label.MIPIdentity2, $label.MIPRights2 -join ":"
                $MIPRightsC = $label.MIPIdentity3, $label.MIPRights3 -join ":"
                $MIPRightsR = $MIPRightsA, $MIPRightsB, $MIPRightsC -join ";"

                Set-Label -Identity $LabelName `
                    -Tooltip $Tooltips[0] `
                    -LocaleSettings (
                        (ConvertTo-Json $DisplayNameLocaleSettings -Depth 4 -Compress),
                        (ConvertTo-Json $TooltipLocaleSettings -Depth 4 -Compress)
                    ) `
                    -EncryptionContentExpiredOnDateInDaysOrNever $label.EncryptionContentExpiredOnDateInDaysOrNever `
                    -EncryptionDoNotForward $label.EncryptionDoNotForward `
                    -EncryptionEnabled $label.EncryptionEnabled `
                    -EncryptionEncryptOnly $label.EncryptionEncryptOnly `
                    -EncryptionOfflineAccessDays $label.EncryptionOfflineAccessDays `
                    -EncryptionPromptUser $label.EncryptionPromptUser `
                    -EncryptionProtectionType $label.EncryptionProtectionType `
                    -EncryptionRightsDefinitions $MIPRightsR `
                    -ApplyContentMarkingFooterAlignment Left `
                    -ApplyContentMarkingFooterEnabled $label.ApplyContentMarkingFooterEnabled `
                    -ApplyContentMarkingFooterFontColor $label.ApplyContentMarkingFooterFontColor `
                    -ApplyContentMarkingFooterFontSize $label.ApplyContentMarkingFooterFontSize `
                    -ApplyContentMarkingFooterMargin $label.ApplyContentMarkingFooterMargin `
                    -ApplyContentMarkingFooterText $label.ApplyContentMarkingFooterText `
                    -ApplyContentMarkingHeaderAlignment Center `
                    -ApplyContentMarkingHeaderEnabled $label.ApplyContentMarkingHeaderEnabled `
                    -ApplyContentMarkingHeaderFontColor $label.ApplyContentMarkingHeaderFontColor `
                    -ApplyContentMarkingHeaderFontSize $label.ApplyContentMarkingHeaderFontSize `
                    -ApplyContentMarkingHeaderMargin $label.ApplyContentMarkingHeaderMargin `
                    -ApplyContentMarkingHeaderText $label.ApplyContentMarkingHeaderText `
                    -ApplyWaterMarkingEnabled $label.ApplyWaterMarkingEnabled `
                    -ApplyWaterMarkingFontColor $label.ApplyWaterMarkingFontColor `
                    -ApplyWaterMarkingFontSize $label.ApplyWaterMarkingFontSize `
                    -ApplyWaterMarkingLayout $label.ApplyWaterMarkingLayout `
                    -ApplyWaterMarkingText $label.ApplyWaterMarkingText
            } else {
                # Normales Label (kein Encryption)
                Set-Label -Identity $LabelName `
                    -Tooltip $Tooltips[0] `
                    -LocaleSettings (
                        (ConvertTo-Json $DisplayNameLocaleSettings -Depth 4 -Compress),
                        (ConvertTo-Json $TooltipLocaleSettings -Depth 4 -Compress)
                    )
            }
            Log "✅ Label '$LabelName' Sprachen aktualisiert." "SUCCESS"
        } catch {
            Log "❌ Fehler beim Update '$LabelName': $($_.Exception.Message)" "ERROR"
        }
        Start-Sleep -s 2
    }

    # Export CSV/JSON
    $exportCsv = Join-Path $LogFolder "Translated_Labels_$($DatumJetzt).csv"
    $exportJson = Join-Path $LogFolder "Translated_Labels_$($DatumJetzt).json"
    $global:ExcelData | Export-Csv -Path $exportCsv -Delimiter ";" -Encoding UTF8 -NoTypeInformation
    $global:ExcelData | ConvertTo-Json -Depth 5 | Out-File -FilePath $exportJson -Encoding utf8

    if ($labelsWithMissingTranslations.Count -gt 0) {
        Log "✅ Übersetzungen ergänzt & exportiert: CSV=$exportCsv, JSON=$exportJson" "SUCCESS"
        Log "ℹ️ Labels mit fehlenden Übersetzungen (vorher): $($labelsWithMissingTranslations -join ', ')" "INFO"
        [System.Windows.MessageBox]::Show("⚠️ Einige Labels hatten fehlende Übersetzungen und wurden ergänzt!\n\nDateien:\n- $exportCsv\n- $exportJson", "Übersetzungen ergänzt", "OK", "Warning")
    } else {
        Log "✅ Keine fehlenden Übersetzungen – alles aktuell." "SUCCESS"
        [System.Windows.MessageBox]::Show("✅ Keine fehlenden Übersetzungen – alles aktuell.", "Info", "OK", "Info")
    }
}


# === GUI ===
Add-Type -AssemblyName PresentationFramework
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Label Updater" Height="600" Width="900">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Text="Label Language Updater" FontWeight="Bold" FontSize="20" Margin="0,0,0,10"/>
        <DataGrid Name="dgExcelData" Grid.Row="1" AutoGenerateColumns="True" IsReadOnly="True"/>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button Name="btnLoadExcel" Content="Laden" Width="100" Margin="0,0,10,0"/>
            <Button Name="btnStartUpdate" Content="Sprachen aktualisieren" Width="150" Margin="0,0,10,0"/>
            <Button Name="btnClose" Content="Abbrechen" Width="100"/>
        </StackPanel>
    </Grid>
</Window>
"@
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
$btnLoadExcel = $window.FindName("btnLoadExcel")
$btnStartUpdate = $window.FindName("btnStartUpdate")
$btnClose = $window.FindName("btnClose")
$dgExcelData = $window.FindName("dgExcelData")

$btnLoadExcel.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Excel-Dateien (*.xlsx)|*.xlsx"
    if ($ofd.ShowDialog() -eq 'OK') {
        $global:ExcelData = Import-Excel -Path $ofd.FileName
        $dgExcelData.ItemsSource = $global:ExcelData
        Log "📂 Excel geladen: $($ofd.FileName)" "INFO"
        Ensure-Labels-Existence -ExcelData $global:ExcelData
    }
})
$btnStartUpdate.Add_Click({
    if (-not $global:ExcelData) {
        [System.Windows.MessageBox]::Show("Bitte zuerst eine Excel-Datei laden!", "Fehler", "OK", "Error")
        return
    }
    $window.Close()
    Start-LabelUpdate
})
$btnClose.Add_Click({ $window.Close() })
$window.ShowDialog() | Out-Null

Log "⚡ Script wurde beendet." "INFO"

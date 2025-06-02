<#
.SYNOPSIS
 Start-GUI für Microsoft Purview Label Tools – WPF-Variante
#>

# === Parameter definieren ===
param (
    [string]$UserPrincipalName = "user@domain.tld",
    [string]$Tenantdomain      = "domain.tld",
    [string]$CompanyLogoPath = "",
    [string]$CompanyLogoUrl = "",
    [string]$ProductLogoPath = "",
    [string]$ProductLogoUrl = "",
    [string]$LogoUrl = "",
    [string]$LogoGIFUrl = "",
    [string]$MailToPrimary = "michael.kirst-neshva@bdo-digital.eu",
    [string]$MailToSecondary = "",
    [string]$LogFolder,
    [string]$MSPPartner,
    [string]$MSPNameAP,
    [string]$MSPMail,
    [string]$MSPURL,
    [string]$MSPNameEU,
    [string]$CompanyLogoBase64,
    [string]$ProductLogoBase64
)

# === Zentrale Module laden ===
$commonModule = Join-Path $PSScriptRoot "common.psm1"
$configFile   = Join-Path $PSScriptRoot "config.ps1"

if (Test-Path $commonModule) {
    Import-Module $commonModule -Force
} else {
    Write-Host "❌ Modul common.psm1 nicht gefunden im Pfad $commonModule" -ForegroundColor Red
    exit 1
}

if (Test-Path $configFile) {
    . $configFile
} else {
    Write-Host "❌ Konfigurationsdatei config.ps1 nicht gefunden im Pfad $configFile" -ForegroundColor Red
    exit 1
}

# Standardwerte aus config.ps1 übernehmen, falls Parameter nicht gesetzt
$LogFolder      = $LogFolder      -or $script:DefaultLogFolder
$MSPPartner     = $MSPPartner     -or $script:DefaultMSPPartner
$MSPNameAP      = $MSPNameAP      -or $script:DefaultMSPNameAP
$MSPMail        = $MSPMail        -or $script:DefaultMSPMail
$MSPURL         = $MSPURL         -or $script:DefaultMSPURL
$CompanyLogoBase64 = $CompanyLogoBase64 -or $script:DefaultCompanyLogoBase64
$ProductLogoBase64  = $ProductLogoBase64  -or $script:DefaultProductLogoBase64


# === Logging
if (!(Test-Path $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null }
$DatumJetzt = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogFile = Join-Path $LogFolder "StartGUI_$DatumJetzt.log"
$GuiConfigPath = Join-Path $PSScriptRoot "GUIConfig.json"

# Logging-Funktion aus common.psm1 verwenden
$script:LogFile = $LogFile

# === Konfiguration aus JSON laden (falls vorhanden)
if (Test-Path $GuiConfigPath) {
    try {
        $cfg = Get-Content -Raw -Path $GuiConfigPath | ConvertFrom-Json
        if ($cfg.UserPrincipalName) { $UserPrincipalName = $cfg.UserPrincipalName }
        if ($cfg.Tenantdomain)      { $Tenantdomain      = $cfg.Tenantdomain }
        if ($cfg.MailToPrimary)     { $MailToPrimary     = $cfg.MailToPrimary }
        if ($cfg.MailToSecondary)   { $MailToSecondary   = $cfg.MailToSecondary }
        if ($cfg.LogFolder)         { $LogFolder         = $cfg.LogFolder }
        Log "📥 Konfiguration aus $GuiConfigPath geladen." "INFO"
    } catch {
        Log "⚠️ Fehler beim Laden der Konfiguration: $_" "ERROR"
    }
}

# === XAML GUI
Add-Type -AssemblyName PresentationFramework
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Microsoft Purview – Startkonfiguration"
        Height="780" Width="580" WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" Topmost="True">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Orientation="Horizontal" VerticalAlignment="Top">
            <Image Name="imgCompanyLogo" Height="48" Width="48" Margin="0,0,10,0"/>
            <TextBlock Text="Microsoft Purview Automation" FontSize="20" FontWeight="Bold" VerticalAlignment="Center"/>
            <Image Name="imgProductLogo" Height="48" Width="48" HorizontalAlignment="Right"/>
        </StackPanel>

        <StackPanel Grid.Row="1" Margin="0,10,0,10">
            <TextBlock Text="🔐 Anmeldung für Purview Compliance Portal" FontWeight="Bold"/>
            <StackPanel Orientation="Horizontal" Margin="0,5">
                <TextBlock Text="Benutzer-Email (UPN):" Width="180"/>
                <TextBox Name="txtUPN" Width="300"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,5">
                <TextBlock Text="Tenant Domain:" Width="180"/>
                <TextBox Name="txtTenant" Width="300"/>
            </StackPanel>

            <TextBlock Text="📧 eMail Empfänger" FontWeight="Bold" Margin="0,10,0,0"/>
            <StackPanel Orientation="Horizontal" Margin="0,5">
                <TextBlock Text="Primärer Empfänger:" Width="180"/>
                <TextBox Name="txtMailPrimary" Width="300"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,5">
                <TextBlock Text="Sekundärer Empfänger:" Width="180"/>
                <TextBox Name="txtMailSecondary" Width="300"/>
            </StackPanel>

            <TextBlock Text="💾 Speicherorte" FontWeight="Bold" Margin="0,10,0,0"/>
            <StackPanel Orientation="Horizontal" Margin="0,5">
                <TextBlock Text="Log-/Bericht-Ordner:" Width="180"/>
                <TextBox Name="txtLogFolder" Width="250"/>
                <Button Name="btnBrowse" Content="..." Width="30" Margin="5,0"/>
            </StackPanel>

            <TextBlock Text="Aktion" FontWeight="Bold" Margin="0,10,0,0"/>
            <StackPanel Orientation="Horizontal" Margin="0,5">
                <RadioButton Name="radReport" Content="Label-Report" IsChecked="True" Margin="0,0,10,0"/>
                <RadioButton Name="radAddLanguage" Content="Label-Sprachen" Margin="0,0,10,0"/>
                <RadioButton Name="radSortieren" Content="Label-Sortieren"/>
            </StackPanel>

            <TextBlock Text="Bei Fragen wenden Sie sich an:" FontSize="10" Margin="0,10,0,0"/>
            <TextBlock Text="$MSPPartner, $MSPNameAP" FontSize="10"/>
        </StackPanel>

        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnStart" Content="Starten" Width="120" Margin="0,0,10,0"/>
            <Button Name="btnCancel" Content="Abbrechen" Width="100"/>
        </StackPanel>
    </Grid>
</Window>
"@
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Vorbelegen
$window.FindName("txtUPN").Text = $UserPrincipalName
$window.FindName("txtTenant").Text = $Tenantdomain
$window.FindName("txtMailPrimary").Text = $MailToPrimary
$window.FindName("txtMailSecondary").Text = $MailToSecondary
$window.FindName("txtLogFolder").Text = $LogFolder

# Logos
function Set-Image ($ctrl, $base64) {
    if ($base64 -and $base64.Length -gt 100) {
        try {
            $bytes = [Convert]::FromBase64String(($base64 -replace '^data:image\/[a-z]+;base64,', ''))
            $stream = New-Object IO.MemoryStream (,[byte[]]$bytes)
            $img = [System.Windows.Media.Imaging.BitmapImage]::new()
            $img.BeginInit()
            $img.StreamSource = $stream
            $img.EndInit()
            $ctrl.Source = $img
        } catch { Log "❌ Logo-Fehler: $_" "ERROR" }
    }
}
Set-Image ($window.FindName("imgCompanyLogo")) $CompanyLogoBase64
Set-Image ($window.FindName("imgProductLogo")) $ProductLogoBase64

# Aktionen
$window.FindName("btnBrowse").Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dialog.ShowDialog() -eq 'OK') { $window.FindName("txtLogFolder").Text = $dialog.SelectedPath }
})
$window.FindName("btnCancel").Add_Click({ Log "Abgebrochen." "INFO"; $window.Close(); exit 1 })
$window.FindName("btnStart").Add_Click({
    $script:Result = @{
        UserPrincipalName = $window.FindName("txtUPN").Text
        Tenantdomain      = $window.FindName("txtTenant").Text
        MailToPrimary     = $window.FindName("txtMailPrimary").Text
        MailToSecondary   = $window.FindName("txtMailSecondary").Text
        LogFolder         = $window.FindName("txtLogFolder").Text
        Aktion            = if ($window.FindName("radReport").IsChecked) { "Report" }
                             elseif ($window.FindName("radAddLanguage").IsChecked) { "AddLanguage" }
                             else { "Sortieren" }
    }
    $script:Result | ConvertTo-Json -Depth 2 | Set-Content -Path $GuiConfigPath -Encoding UTF8
    Log "✅ Konfiguration gespeichert: $GuiConfigPath" "SUCCESS"
    $window.Close()
})

$window.ShowDialog() | Out-Null
if (-not $script:Result) { Log "Abbruch durch Benutzer." "ERROR"; exit 1 }

# Hauptskript
$mainScript = switch ($script:Result.Aktion) {
    "Report"     { ".\03-Run-Purview-Create-Documentation_GUI_Final_V9.ps1" }
    "AddLanguage"{ ".\02-Run-PurviewLabelProvisioning_Create_Missing_Config_Only_Language_and_Translation_Final_V9.ps1" }
    "Sortieren"  { ".\03-Run-Purview-Sort-Labels.ps1" }
    default      { Write-Error "Keine gültige Aktion!"; exit 1 }
}
if (-not (Test-Path $mainScript)) { Write-Error "Hauptskript fehlt: $mainScript"; exit 1 }

$arguments = @("-ExecutionPolicy Bypass", "-STA", "-File `"$mainScript`"", "-GuiConfigPath `"$GuiConfigPath`"") -join " "
Start-Process powershell.exe -ArgumentList $arguments -WindowStyle Normal
Log "✅ Hauptskript gestartet: $mainScript" "SUCCESS"
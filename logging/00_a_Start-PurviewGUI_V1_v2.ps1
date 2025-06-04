
param (
    [string]$UserPrincipalName = "",
    [string]$Tenantdomain      = "",
    [string]$CompanyLogoPath = "",
    [string]$CompanyLogoUrl = "",
    [string]$ProductLogoPath = "",
    [string]$ProductLogoUrl = "",
    [string]$LogoUrl = "",
    [string]$LogoGIFUrl = "",
    [string]$MailToPrimary = "",
    [string]$MailToSecondary = "",
    [string]$LogFolder = "C:\Temp\script\",
    [string]$MSPPartner = "",
    [string]$MSPNameAP  = "",
    [string]$MSPMail    = "",
    [string]$MSPURL     = "",
    [string]$MSPNameEU  = "",
    [string]$CompanyLogoBase64 = "",
    [string]$ProductLogoBase64 = ""
)

Import-Module "$PSScriptRoot\CentralLogging.psm1" -Force
Set-LogFile -LogFolder "$PSScriptRoot\Logs"
Write-Log -Message "00_a_Start-PurviewGUI_V1_v2.ps1 gestartet" -Level "INFO"



try {
    if (!(Test-Path $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null }
    $DatumJetzt = Get-Date -Format 'yyyyMMdd_HHmmss'
    $LogFile = Join-Path $LogFolder "StartGUI_$DatumJetzt.log"
    $GuiConfigPath = Join-Path $PSScriptRoot "GUIConfig.json"

    function Log {
        param([string]$Message, [string]$Level = "INFO")
        Write-Log -Message $Message -Level $Level -LogFile $LogFile
    }

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

    # ==== Demo Daten
    $UserPrincipalName = "mkn@ankbs.de"
    $Tenantdomain      = "ankbs.de"
    $MailToPrimary = "michael.kirst-neshva@bdo-digital.eu"
    $MailToSecondary = ""
    $LogFolder = "C:\Temp\script\"
    $MSPPartner = "Any MSP Partner"
    $MSPNameAP  = "Contactname"
    $MSPMail    = "Support eMail"
    $MSPURL     = "Support URL"
    $MSPNameEU  = "Contactname"
    $CompanyLogoBase64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAKoAAACqCAIAAACyFEPVAAAAIGNIUk0AAHomAACAhAAA+gAAAIDoAAB1MAAA6mAAADqYAAAXcJy6UTwAAAAGYktHRAD/AP8A/6C9p5MAAB3eSURBVHja7Z15mFTFvfe/..."
    # === XAML GUI
    Add-Type -AssemblyName PresentationFramework
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Microsoft Purview – Startkonfiguration"
        Height="780" Width="740" WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" Topmost="True">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/> <!-- Footer-Bereich -->
        </Grid.RowDefinitions>
        <!-- Kopfzeile -->
        <StackPanel Grid.Row="0" Orientation="Horizontal" VerticalAlignment="Top">
            <Image Name="imgCompanyLogo" Height="48" Width="48" Margin="0,0,10,0"/>
            <TextBlock Text="Microsoft Purview Automation" FontSize="20" FontWeight="Bold" VerticalAlignment="Center"/>
            <Image Name="imgProductLogo" Height="48" Width="48" HorizontalAlignment="Right"/>
        </StackPanel>
        <!-- Hauptbereich -->
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
                <RadioButton Name="radReport" Content="Label Report" IsChecked="True" Margin="0,0,10,0"/>
                <RadioButton Name="radAddLanguage" Content="Label bearbeiten" Margin="0,0,10,0"/>
                <RadioButton Name="radSortieren" Content="Label sortieren"/>
            </StackPanel>
            <TextBlock Text="Bei Fragen wenden Sie sich an:" FontSize="10" Margin="0,10,0,0"/>
            <TextBlock Text="$MSPPartner, $MSPNameAP" FontSize="10"/>
        </StackPanel>
        <!-- Buttons -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button Name="btnStart" Content="Starten" Width="120" Margin="0,0,10,0"/>
            <Button Name="btnCancel" Content="Abbrechen" Width="100"/>
        </StackPanel>
        <!-- Footer mit MSP Infos -->
        <DockPanel Grid.Row="3" Margin="0,10,0,0" LastChildFill="False">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Left">
                <TextBlock Text="MSP: " FontWeight="Bold"/>
                <TextBlock Name="txtMSP" Margin="5,0,0,0"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <TextBlock Text="🔗 " FontWeight="Bold"/>
                <TextBlock Name="txtURL" Foreground="Blue" TextDecorations="Underline" Cursor="Hand" Margin="0,0,10,0"/>
                <TextBlock Text="✉️ " FontWeight="Bold"/>
                <TextBlock Name="txtMail" Foreground="Blue" TextDecorations="Underline" Cursor="Hand"/>
            </StackPanel>
        </DockPanel>
    </Grid>
</Window>
"@
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $window.FindName("txtUPN").Text = $UserPrincipalName
    $window.FindName("txtTenant").Text = $Tenantdomain
    $window.FindName("txtMailPrimary").Text = $MailToPrimary
    $window.FindName("txtMailSecondary").Text = $MailToSecondary
    $window.FindName("txtLogFolder").Text = $LogFolder

    # Footer-Elemente befüllen
    $txtMSP = $window.FindName("txtMSP")
    $txtURL = $window.FindName("txtURL")
    $txtMail = $window.FindName("txtMail")
    $txtMSP.Text = "$MSPPartner - $MSPNameAP"
    $txtURL.Text = $MSPURL
    $txtMail.Text = $MSPMail

    # Klickbare Links
    $txtURL.Add_MouseLeftButtonUp({ Start-Process $MSPURL })
    $txtMail.Add_MouseLeftButtonUp({ Start-Process "mailto:$MSPMail" })

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
        "Report"     { ".\03-Run-Purview-Create-Documentation_GUI_Final_V10.ps1" }
        "AddLanguage"{ ".\02-Run-PurviewLabelProvisioning_Create_Missing_Config_Only_Language_Final_V1.ps1" }
        "Sortieren"  { ".\04-Run-Purview-Label-PriorityManager_V1.ps1" }
        default      { Write-Log -Message "Keine gültige Aktion!" -Level "ERROR"; exit 1 }
    }
    if (-not (Test-Path $mainScript)) { Write-Log -Message "Hauptskript fehlt: $mainScript" -Level "ERROR"; exit 1 }

    $arguments = @("-ExecutionPolicy Bypass", "-STA", "-File `"$mainScript`"", "-GuiConfigPath `"$GuiConfigPath`"") -join " "
    Start-Process powershell.exe -ArgumentList $arguments -WindowStyle Normal
    Log "✅ Hauptskript gestartet: $mainScript" "SUCCESS"
}
catch {
    Handle-Error -Message "Fehler im Start-GUI Skript" -ErrorObject $_
}
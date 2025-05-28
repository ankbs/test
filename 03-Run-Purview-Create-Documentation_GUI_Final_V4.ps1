# === Vorbereitungen ===
param(
    [bool]$SendReport = $false,
    [string]$UserPrincipalName,
    [string]$Tenantdomain,
    [string]$MailToPrimary,
    [string]$MailToSecondary = "",
    [bool]$CreateMissingLabels = $false,
    [int]$MFATimeoutSeconds = 120,
    [string]$SourceExcelPath = "",
    [bool]$UseExistingLabels = $true,
    [int]$Priority = 0,
    [int]$PriorityMin = 0,
    [int]$PriorityMax = 0,
    # [string]$Priorities = "",
    [string[]]$LabelNames = @(),  # Alternative Labelnamen angeben (z. B. -LabelNames "Confidential", "Public")
    [bool]$UseLabelGUI = $true,
    [bool]$ExportWord = $true,
    [bool]$ExportPDF = $false,
    [bool]$DryRun = $false,
    [bool]$UseProgressBar = $false,
    [string]$LogFolder = "C:\Temp\script\",
    [int]$AutoCloseAfterSeconds = 60,
    [string]$LogoGIFUrl    = "https://i.gifer.com/ZKZg.gif", # "", # "https://i.gifer.com/ZKZg.gif",  # animiertes Zahnrad
    # [string]$LogoUrl = "https://upload.wikimedia.org/wikipedia/commons/thumb/e/e5/Microsoft_Purview_Logo.svg/250px-Microsoft_Purview_Logo.svg.png",  # Produkt Logo 🔁 Ersetze durch tatsächliche Bild-URL
    [string]$CompanyLogoPath = "", # ".\logo.png",
    [string]$CompanyLogoUrl  = "", # "https://example.com/logo.png",
    [string]$CompanyLogoBase64 = "data:image/png;base64,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r1+fSqUmTZo0evRoQkhLS8uCBQtGjhw5YsSIDes3zJ8/f9asWWeddVYikWhtbV22bNl99913880367qez+cJIZZlLV68uL29/fLLL29qaspms6tWrXrmmWcopZqmpdPpVatWXXzxxblc7vXXX+/u7p46deqcOXMYYxs2bHj44YdFUTz1lFOXPb9syZIlc+bMGTVqlKIou3btWrly5csvv3z22We7SgOBgNtNVKaBgZO/rJiWefpZU2dcd5X7oKw0/hNzrrj8v/+wAJo1JBgbHE0AYASC2/HbtEZQFVDkjCUPLbrrpttqPn8e08o0nsS4MXdEGu+67rqutzbG4zXduQze1TN72sYOboqdPg5dBsLKeV/+4glKTUoJYcvOZX/562133lZ/2hTksojHTzz19Ht+dOePv/eD2edfrDhoGtLw1au/lhg7CgEBNk6++MLR/7Ni+TPPfnboMInQEeOa59xxAwAwDof837k//sU3rv3LQw9fcvU1j/36oWuv+MqIr12BCMAxvib542Fjbr32m+sWPTnxk9Pat+78/A3XT77qMoQF2PyOr3/zzFOnfPo710HUEVFHnXHKqHETE13G2tdXo6uMsACLOYRbtmXbdjAYnD17dlNTUyaTCQQC06ZN27Bhw+rVq4cNG/anP/1p9uzZl156qWmasiyPHz++ubn55ptvXrFixcyZM+vr62VZ3rhx48qVK+fOnZtKpWRZjkQijY2Nix9Z3NXVRQiRehFFURTFSy65ZFjjMEEUAJx77rkbN25sbW2dMmXKU089ddlll02bNs2xnUw2M2nSpGHDhrW2th5iIuuI+v72zmS8ZsTE8WAA53AYapT42KZMPgcuJOSgAgpeMfijNBQMUSogly2bxo63t7x2x5ssKKvBADRrEA3s6mjb+s47NXWpkBqEqk7/l0/f8oNbOmZ/ZcKJExyHDR/eNOacc5BIZF55GY7T8sqq1lXrisVioiahKIphGMVMbueaNXX1dWoskhjVBEUAAN2ETNVYxDZMxBStVD535gxQGJahCDJkQLZHnHrS1jdbkDeK+Xx7e7ux4NF9xWypWASQqq/LF4sb1785ceq0kWNGT5p2JkKCxqxA3jK78+dffSWigCWUiwXR5jKTJ0+evHLVKigEIZlxxjjjnLuddH19vWM7NTU1nHPHdjjnbgdsmuZpp51WLBbD4bCu65xz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", # Companylogo als eingebettetes Base64
    [string]$LogoUrl = "",  # Produktlogo per URL
    [string]$ProductLogoBase64 = "data:image/png;base64,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" # Produktlogo als eingebettetes Base64
)


    if (-not $MailToPrimary -or $MailToPrimary -eq "") {
    Write-Host "❌ Es wurde keine primäre E-Mail-Adresse für den Versand angegeben." -ForegroundColor Red
    exit 1
}

    # === Zufalls-GIF fallback wenn kein GIF übergeben wurde
    if (-not $LogoGIFUrl -or $LogoGIFUrl.Trim() -eq "") {
        $StandardGIFs = @(
            "https://media.tenor.com/I6kN-6X7nhAAAAAi/loading-buffering.gif",
            "https://media.tenor.com/On7kvXhzml4AAAAj/loading-gif.gif",
            "https://media.tenor.com/2uyENRmiUt0AAAAC/circle-loading.gif"
        )

        $LogoGIFUrl = Get-Random -InputObject $StandardGIFs
        Log "🎞️ Animiertes GIF für Splash: $LogoGIFUrl" "INFO"
        Write-Host "🔄 Kein GIF übergeben – zufälliges Standard-GIF ausgewählt: $LogoGIFUrl" -ForegroundColor Cyan
    }

    if ($LogoUrl -and $LogoUrl.Trim() -ne "") {
        Log "🏷️ Produktlogo wird geladen: $LogoUrl" "DEBUG"
        }



Clear-Host

# === Konfiguration ===
# $UserPrincipalName = "mkn@huehodevbdo.onmicrosoft.com"
# $Tenantdomain = "huehodevbdo.onmicrosoft.com"

# === MSP Kontaktinformationen ===
$MSPPartner = "Cloud Security && Compliance Services"
$MSPNameAP  = "Michael Kirst-Neshva"
$MSPMail    = "support@domaine.io"
$MSPURL     = "https://www.domaine.io"
# $LogoUrl    = "https://upload.wikimedia.org/wikipedia/commons/thumb/e/e5/Microsoft_Purview_Logo.svg/250px-Microsoft_Purview_Logo.svg.png"  # Produkt Logo 🔁 Ersetze durch tatsächliche Bild-URL
# $LogoGIFUrl    = "https://i.gifer.com/ZKZg.gif"  # animiertes Zahnrad
# $LogoMSUrl   = "https://upload.wikimedia.org/wikipedia/commons/4/44/Microsoft_logo.svg"  # alternativ PNG verwenden
# $LogoMSUrl     = "https://upload.wikimedia.org/wikipedia/commons/thumb/e/e5/Microsoft_Purview_Logo.svg/250px-Microsoft_Purview_Logo.svg.png"  # 🔁 Ersetze durch tatsächliche Bild-URL

# https://i.gifer.com/ZKZg.gif → animiertes Zahnrad
# https://upload.wikimedia.org/wikipedia/commons/4/4e/Microsoft_logo.svg → statisches MS-Logo (muss konvertiert werden)


if (!(Test-Path $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null }

# === Zeitstempel / Logdateien ===
$DatumJetzt = Get-Date -Format 'yyyyMMdd_HHmmss'
$DatumAnzeige = Get-Date -Format 'dd.MM.yyyy'
$cleanUser = $UserPrincipalName -replace '[^a-zA-Z0-9]', '_'
$LogFolder = $LogFolder.TrimEnd('\')
$LogFile = [System.IO.Path]::Combine($LogFolder, "Label_Status_LOG_$DatumJetzt.log")
$CreatedLabelsCsv = [System.IO.Path]::Combine($LogFolder, "Erstellte_Labels_$DatumJetzt.csv")
$StatusReportCsv = [System.IO.Path]::Combine($LogFolder, "Label_Status_$DatumJetzt.csv")

# === Logging + Error-Handling ===
function Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $prefix = switch ($Level) {
        "INFO"     { "ℹ️" }
        "SUCCESS"  { "✅" }
        "ERROR"    { "❌" }
        "DEBUG"    { "🐛" }
        default    { "🔹" }
    }
    $logEntry = "$timestamp $prefix $Message"
    Add-Content -Path $LogFile -Value $logEntry -Encoding utf8
    Write-Host $logEntry -Encoding utf8
}

function Handle-Error {
    param([string]$Message, [System.Exception]$ErrorObject)
    Log "$Message $($ErrorObject.Exception.Message)" "ERROR"
    exit 1
}

# === PRIORITÄTENPARSER ===
function Read-PriorityInfo {
    param(
        [string]$Input,
        [int[]]$AvailablePriorities
    )

    $result = @()

    if ([string]::IsNullOrWhiteSpace($Input)) {
        return @()
    }

    $entries = $Input -split ','

    foreach ($entryRaw in $entries) {
    Log "🔍 Starte Prüfung von Eintrag: '$entryRaw'" "DEBUG"
        $entry = $entryRaw.Trim()

        try {
            if ($entry -match '^\d+-\d+$') {
                # Bereich z. B. 100-120
                Log "➕ Bereich erkannt: $start-$end" "DEBUG"
                Log "➕ Einzelwert erkannt: $val" "DEBUG"
                Log "➕ Bereich ab erkannt: $start → $($matched.Count) Treffer" "DEBUG"
                $rangeParts = $entry -split '-'
                $start = [int]$rangeParts[0]
                $end   = [int]$rangeParts[1]

                if ($start -le $end) {
                    $result += $start..$end
                } else {
                    Log "⚠️ Bereich ignoriert (Start > Ende): '$entry'" "WARNING"
                }
            }
            elseif ($entry -match '^\d+-$') {
                # Bereich ab z. B. 150-
                Log "➕ Bereich erkannt: $start-$end" "DEBUG"
                Log "➕ Einzelwert erkannt: $val" "DEBUG"
                Log "➕ Bereich ab erkannt: $start → $($matched.Count) Treffer" "DEBUG"
                $start = [int]($entry -replace '-$', '')
                $matched = $AvailablePriorities | Where-Object { $_ -ge $start }
                $result += $matched
            }
            elseif ($entry -match '^\d+$') {
                # Einzelwert
                Log "➕ Bereich erkannt: $start-$end" "DEBUG"
                Log "➕ Einzelwert erkannt: $val" "DEBUG"
                Log "➕ Bereich ab erkannt: $start → $($matched.Count) Treffer" "DEBUG"
                $val = [int]$entry
                $result += $val
            }
            else {
                Log "⚠️ Ungültiger Prioritätseintrag ignoriert: '$entry'" "WARNING"
            }
        }
        catch {
            Log "⚠️ Fehler beim Parsen von Prioritätseintrag '$entry': $($_.Exception.Message)" "ERROR"
        }
    }

    return ($result | Sort-Object -Unique)
}




# === Modulprüfung ...
# (Keine Änderungen in Modulprüfung, GUI, Connect-Session etc.)

# === Modulprüfung

function Test-ModuleInstalled {
    param([string]$ModuleName)
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Log "📦 Modul '$ModuleName' nicht gefunden – versuche Installation..." "INFO"
        if (-not (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue)) {
            Register-PSRepository -Default
        }
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber
            Log "✅ Modul '$ModuleName' installiert." "SUCCESS"
        } catch {
            Handle-Error "Modulinstallation fehlgeschlagen" $_
        }
    } else {
        Log "✅ Modul '$ModuleName' ist bereits installiert." "DEBUG"
    }
}

Ensure-Module -ModuleName "ExchangeOnlineManagement"
Import-Module ExchangeOnlineManagement
Ensure-Module -ModuleName "ImportExcel"
Import-Module ImportExcel

if ($DryRun) {
    Log "🧪 DryRun-Modus aktiviert – es werden keine Dateien gespeichert oder Mails gesendet." "WARNING"
}

# Log "🐛 [DEBUG] Inhalt von Priorities: $Priorities" "DEBUG"

# =====================================================================##

# === Import Excel Data

function Import-LabelsFromExcel {
    param (
        [string]$FilePath
    )

    if (-not (Test-Path $FilePath)) {
        Log "❌ Excel-Datei nicht gefunden: $FilePath" "ERROR"
        return
    }

    try {
        $data = Import-Excel -Path $FilePath
        $importedLabels = @()
        $notFound = @()

        foreach ($row in $data) {
            $labelName = $row.Display1
            $fallbackName = $row.LabelNameOLD
            $label = $null

            if ($labelName -and $labelName.Trim() -ne "") {
                try {
                    $label = Get-Label -Identity $labelName.Trim()
                    Log "✅ Label gefunden: $($label.DisplayName)" "SUCCESS"
                } catch {
                    Log "⚠️ Label nicht gefunden mit Display1: $labelName – versuche LabelNameOLD..." "WARNING"
                }
            }

            # Fallback auf LabelNameOLD
            if (-not $label -and $fallbackName -and $fallbackName.Trim() -ne "") {
                try {
                    $label = Get-Label -Identity $fallbackName.Trim()
                    Log "✅ Fallback-Label gefunden: $($label.DisplayName)" "SUCCESS"
                } catch {
                    $notFound += "$labelName / $fallbackName"
                    Log "❌ Label nicht gefunden: $labelName (Fallback: $fallbackName)" "ERROR"
                }
            }

            if ($label) {
                $importedLabels += $label
            }
        }

        if ($importedLabels.Count -gt 0) {
            $script:allLabels = $importedLabels
            Update-LabelList -source $script:allLabels
            Log "📥 $($importedLabels.Count) gültige Labels aus Excel übernommen." "INFO"
        } else {
            Log "⚠️ Keine gültigen Labels aus Excel importiert." "WARNING"
        }

        if ($notFound.Count -gt 0) {
            Log "⚠️ Nicht gefundene Labels (Excel): $($notFound -join ', ')" "DEBUG"
        }
    } catch {
        Handle-Error "Fehler beim Import oder Verarbeiten der Excel-Datei" $_
    }
}



# =====================================================================##


# === SplashScreen ===


function Start-SplashInThread {
    param (
        [bool]$UseProgressBar = $false,
        [int]$AutoCloseAfterSeconds = 60,
        [string]$CompanyLogoPath,
        [string]$CompanyLogoUrl,
        [string]$CompanyLogoBase64,
        [string]$LogoGIFUrl,
        [string]$ProductLogoBase64 = "",
        [string]$LogoUrl = ""
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $script:splashRunspace = [runspacefactory]::CreateRunspace()
    $script:splashRunspace.ApartmentState = "STA"
    $script:splashRunspace.ThreadOptions = "ReuseThread"
    $script:splashRunspace.Open()

    $ps = [powershell]::Create()
    $ps.Runspace = $script:splashRunspace

    $ps.AddScript({
        param(
            $CompanyLogoPath, $CompanyLogoUrl, $CompanyLogoBase64,
            $UseProgressBar, $AutoCloseAfterSeconds,
            $LogoGIFUrl, $ProductLogoBase64, $LogoUrl
        )

        function Log {
            param([string]$Message, [string]$Level = "INFO")
            $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            $prefix = switch ($Level) {
                "INFO"     { "ℹ️" }
                "SUCCESS"  { "✅" }
                "ERROR"    { "❌" }
                "DEBUG"    { "🐛" }
                default    { "🔹" }
            }
            $logEntry = "$timestamp $prefix $Message"
            $global:LogFile = "$env:TEMP\PurviewGUI.log"
            Add-Content -Path $global:LogFile -Value $logEntry -Encoding utf8
            Write-Host $logEntry
        }

        function Handle-Error {
            param([string]$Message, [Parameter(ValueFromPipeline=$true)]$ErrorObject)
            $msg = if ($ErrorObject.Exception) { $ErrorObject.Exception.Message } else { "$ErrorObject" }
            Log "$Message – $msg" "ERROR"
            throw "$Message – $msg"
        }

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $form = New-Object Windows.Forms.Form
        $form.Text = "Microsoft Purview GUI lädt..."
        $form.Size = New-Object Drawing.Size(500, 260)
        $form.StartPosition = "CenterScreen"
        $form.FormBorderStyle = "FixedDialog"
        $form.ControlBox = $false
        $form.TopMost = $true
        $form.BackColor = [System.Drawing.Color]::White

        $wc = New-Object Net.WebClient
        $timer = $null

        try {
            $image = $null
            if ($CompanyLogoPath -and (Test-Path $CompanyLogoPath)) {
                $image = [System.Drawing.Image]::FromFile((Resolve-Path $CompanyLogoPath))
                Log "📦 Firmenlogo aus Datei geladen: $CompanyLogoPath"
            }
            elseif ($CompanyLogoUrl -and $CompanyLogoUrl.StartsWith("http")) {
                $imgBytes = $wc.DownloadData($CompanyLogoUrl)
                $ms = New-Object IO.MemoryStream (,[byte[]]$imgBytes)
                $image = [System.Drawing.Image]::FromStream($ms)
                Log "🌐 Firmenlogo von URL geladen: $CompanyLogoUrl"
            }
            elseif ($CompanyLogoBase64 -and $CompanyLogoBase64.Length -gt 100) {
                $cleanBase64 = $CompanyLogoBase64 -replace '^data:image\/[a-z]+;base64,', ''
                $bytes = [Convert]::FromBase64String($cleanBase64)
                $ms = New-Object IO.MemoryStream (,[byte[]]$bytes)
                $image = [System.Drawing.Image]::FromStream($ms)
                Log "🧬 Firmenlogo aus eingebettetem Base64 geladen."
            }

            if ($image) {
                $pic1 = New-Object Windows.Forms.PictureBox
                $pic1.Image = $image
                $pic1.SizeMode = "Zoom"
                $pic1.Location = New-Object Drawing.Point(30, 30)
                $pic1.Size = New-Object Drawing.Size(100, 100)
                $form.Controls.Add($pic1)
            }
        } catch {
            Handle-Error "⚠️ Fehler beim Laden des Firmenlogos" $_
        }

        try {
            $imageProduct = $null
            if ($ProductLogoBase64 -and $ProductLogoBase64.Length -gt 100) {
                $cleanBase64 = $ProductLogoBase64 -replace '^data:image\/[a-z]+;base64,', ''
                $bytes = [Convert]::FromBase64String($cleanBase64)
                $ms = New-Object IO.MemoryStream (,[byte[]]$bytes)
                $imageProduct = [System.Drawing.Image]::FromStream($ms)
                Log "🧬 Produktlogo aus Base64 geladen."
            }
            elseif ($LogoUrl -and $LogoUrl.StartsWith("http")) {
                $imgBytes = $wc.DownloadData($LogoUrl)
                $ms = New-Object IO.MemoryStream (,[byte[]]$imgBytes)
                $imageProduct = [System.Drawing.Image]::FromStream($ms)
                Log "🌐 Produktlogo von URL geladen: $LogoUrl"
            }

            if ($imageProduct) {
                $picProduct = New-Object Windows.Forms.PictureBox
                $picProduct.Image = $imageProduct
                $picProduct.SizeMode = "Zoom"
                $picProduct.Location = New-Object Drawing.Point(360, 30)
                $picProduct.Size = New-Object Drawing.Size(100, 100)
                $form.Controls.Add($picProduct)
            }
        } catch {
            Handle-Error "⚠️ Fehler beim Laden des Produktlogos" $_
        }

        if (-not $UseProgressBar) {
            try {
                $imgBytes = $wc.DownloadData($LogoGIFUrl)
                $ms = New-Object IO.MemoryStream (,[byte[]]$imgBytes)
                $gif = [System.Drawing.Image]::FromStream($ms)
                $pic3 = New-Object Windows.Forms.PictureBox
                $pic3.Image = $gif
                $pic3.SizeMode = "Zoom"
                $pic3.Location = New-Object Drawing.Point(195, 30)
                $pic3.Size = New-Object Drawing.Size(100, 100)
                $form.Controls.Add($pic3)
                Log "⚙️ Animiertes GIF geladen: $LogoGIFUrl"
            } catch {
                Handle-Error "⚠️ Fehler beim GIF-Download" $_
            }
        } else {
            $progress = New-Object Windows.Forms.ProgressBar
            $progress.Style = 'Marquee'
            $progress.MarqueeAnimationSpeed = 30
            $progress.Location = New-Object Drawing.Point(30, 100)
            $progress.Size = New-Object Drawing.Size(430, 20)
            $form.Controls.Add($progress)

            $timer = New-Object Windows.Forms.Timer
            $timer.Interval = 1000
            $timer.Add_Tick({ try { $null = $progress.Value } catch {} })
            $timer.Start()
        }

        $lbl = New-Object Windows.Forms.Label
        $lbl.Text = "⏳ Die Microsoft Purview GUI wird vorbereitet..."
        $lbl.Font = New-Object Drawing.Font("Segoe UI", 10)
        $lbl.AutoSize = $true
        $lbl.Location = New-Object Drawing.Point(30, 150)
        $form.Controls.Add($lbl)

        $closeTimer = New-Object Windows.Forms.Timer
        $closeTimer.Interval = ($AutoCloseAfterSeconds * 1000)
        $closeTimer.Add_Tick({
            try {
                $closeTimer.Stop()
                if ($form -and !$form.IsDisposed) {
                    $form.Invoke({ $form.Close() })
                }
            } catch {
                Handle-Error "⚠️ Fehler beim automatischen Schließen des Splashscreens" $_
            }
        })
        $closeTimer.Start()

        $form.Add_FormClosing({
            try {
                if ($timer) { $timer.Stop(); $timer.Dispose() }
                if ($closeTimer) { $closeTimer.Stop(); $closeTimer.Dispose() }
            } catch {
                Handle-Error "⚠️ Fehler im FormClosing" $_
            }
            [System.Windows.Forms.Application]::Exit()
        })

        $global:SplashForm = $form
        [System.Windows.Forms.Application]::Run($form)
    }) | Out-Null

    $null = $ps.AddArgument($CompanyLogoPath).
                 AddArgument($CompanyLogoUrl).
                 AddArgument($CompanyLogoBase64).
                 AddArgument($UseProgressBar).
                 AddArgument($AutoCloseAfterSeconds).
                 AddArgument($LogoGIFUrl).
                 AddArgument($ProductLogoBase64).
                 AddArgument($LogoUrl)

    $script:splashJob = $ps.BeginInvoke()
}


function Stop-SplashThread {
    if ($global:SplashForm -and !$global:SplashForm.IsDisposed) {
        $global:SplashForm.Invoke({ $global:SplashForm.Close() })
    }

    if ($script:splashRunspace) {
        $script:splashRunspace.Close()
        $script:splashRunspace.Dispose()
    }
}


# ====================


# === GUI für Labelauswahl + Optionen

function Show-LabelSelectionForm {
    param (
        [string]$CompanyLogoPath,
        [string]$CompanyLogoUrl,
        [string]$CompanyLogoBase64,
        [string]$LogoUrl,
        [string]$ProductLogoBase64
    )

    # === Abhängigkeiten laden (Windows Forms, Drawing)
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # === Initialisierung von Variablen
    $script:WasAbgebrochen = $false
    $script:labelMap = @{}
    $script:allLabels = Get-Label

    # === Formular definieren
    $form = New-Object Windows.Forms.Form
    $form.Text = "Microsoft Purview – Labelauswahl"
    $form.Size = New-Object Drawing.Size(540, 820)
    $form.StartPosition = "CenterScreen"
    $form.Topmost = $true
    $form.Font = New-Object Drawing.Font("Segoe UI", 9)

    # === Tooltip-Unterstützung (Balloon, immer anzeigen)
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.IsBalloon = $true
    $toolTip.ShowAlways = $true

    # === WebClient für Logo-Download
    $wc = New-Object System.Net.WebClient

# === Firmenlogo laden (Pfad > URL > Base64)
try {
    $companyImage = $null

    if ($CompanyLogoPath -and (Test-Path $CompanyLogoPath)) {
        # Logo von lokaler Datei
        $companyImage = [System.Drawing.Image]::FromFile((Resolve-Path $CompanyLogoPath))
    }
    elseif ($CompanyLogoUrl -and $CompanyLogoUrl.StartsWith("http")) {
        # Logo von externer URL
        $bytes = $wc.DownloadData($CompanyLogoUrl)
        $ms = New-Object IO.MemoryStream (,[byte[]]$bytes)
        $companyImage = [System.Drawing.Image]::FromStream($ms)
    }
    elseif ($CompanyLogoBase64 -and $CompanyLogoBase64.Length -gt 100) {
        # Logo aus eingebettetem Base64 (vorher bereinigen!)
        $cleanBase64 = $CompanyLogoBase64 -replace '^data:image\/[a-z]+;base64,', ''
        $bytes = [Convert]::FromBase64String($cleanBase64)
        $ms = New-Object IO.MemoryStream (,[byte[]]$bytes)
        $companyImage = [System.Drawing.Image]::FromStream($ms)
    }

    if ($companyImage) {
        $picCompany = New-Object Windows.Forms.PictureBox
        $picCompany.Image = $companyImage
        $picCompany.SizeMode = "Zoom"
        $picCompany.Size = New-Object Drawing.Size(48, 48)
        $picCompany.Location = New-Object Drawing.Point(10, 10)
        $form.Controls.Add($picCompany)
    }
}
catch {
    Write-Host "⚠️ Fehler beim Laden des Firmenlogos: $_"
}

# === Produktlogo laden (URL > Base64)
try {
    $productImage = $null

    if ($LogoUrl -and $LogoUrl.Trim() -ne "") {
        # Produktlogo von URL
        $bytes = $wc.DownloadData($LogoUrl)
        $ms = New-Object IO.MemoryStream (,[byte[]]$bytes)
        $productImage = [System.Drawing.Image]::FromStream($ms)
    }
    elseif ($ProductLogoBase64 -and $ProductLogoBase64.Length -gt 100) {
        # Produktlogo aus Base64
        $cleanBase64 = $ProductLogoBase64 -replace '^data:image\/[a-z]+;base64,', ''
        $bytes = [Convert]::FromBase64String($cleanBase64)
        $ms = New-Object IO.MemoryStream (,[byte[]]$bytes)
        $productImage = [System.Drawing.Image]::FromStream($ms)
    }

    if ($productImage) {
        $picProduct = New-Object Windows.Forms.PictureBox
        $picProduct.Image = $productImage
        $picProduct.SizeMode = "Zoom"
        $picProduct.Size = New-Object Drawing.Size(48, 48)
        $picProduct.Location = New-Object Drawing.Point(430, 10)
        $form.Controls.Add($picProduct)
    }
}
catch {
    Write-Host "⚠️ Fehler beim Laden des Produktlogos: $_"
}

    # === Titel in der Mitte
    $title = New-Object Windows.Forms.Label
    $title.Text = "Purview Security && Compliance Automation"
    $title.Font = New-Object Drawing.Font("Segoe UI", 12, [Drawing.FontStyle]::Bold)
    $title.Location = New-Object Drawing.Point(70, 20)
    $title.AutoSize = $true
    $form.Controls.Add($title)

    # === Filter
    $lblFilter = New-Object Windows.Forms.Label
    $lblFilter.Text = "Filter nach Endziffer (z.B. 0345):"
    $lblFilter.Location = New-Object Drawing.Point(10, 70)
    $lblFilter.AutoSize = $true
    $form.Controls.Add($lblFilter)

    $txtFilterCode = New-Object Windows.Forms.TextBox
    $txtFilterCode.Width = 100
    $txtFilterCode.Location = New-Object Drawing.Point(200, 90)
    $form.Controls.Add($txtFilterCode)
    $toolTip.SetToolTip($txtFilterCode, "Nur Labels anzeigen, die auf diese 4 Ziffern (Buchungscode) enden")

    $btnApplyFilter = New-Object Windows.Forms.Button
    $btnApplyFilter.Text = "Anwenden"
    $btnApplyFilter.Location = New-Object Drawing.Point(310, 90)
    $btnApplyFilter.Width = 80
    $form.Controls.Add($btnApplyFilter)

    # === Sortierung
    $chkSortByCode = New-Object Windows.Forms.CheckBox
    $chkSortByCode.Text = "Nach Endziffer sortieren"
    $chkSortByCode.Width = 480
    $chkSortByCode.Location = New-Object Drawing.Point(10, 120)
    $form.Controls.Add($chkSortByCode)

    $rbSortAsc = New-Object Windows.Forms.RadioButton
    $rbSortAsc.Text = "Aufsteigend"
    $rbSortAsc.Location = New-Object Drawing.Point(30, 145)
    $rbSortAsc.Size = New-Object Drawing.Size(120, 20)
    $form.Controls.Add($rbSortAsc)

    $rbSortDesc = New-Object Windows.Forms.RadioButton
    $rbSortDesc.Text = "Absteigend"
    $rbSortDesc.Location = New-Object Drawing.Point(160, 145)
    $rbSortDesc.Size = New-Object Drawing.Size(120, 20)
    $rbSortDesc.Checked = $true
    $form.Controls.Add($rbSortDesc)

    # === Label-Liste
    $listbox = New-Object Windows.Forms.CheckedListBox
    $listbox.Size = New-Object Drawing.Size(500, 280)
    $listbox.Location = New-Object Drawing.Point(10, 175)
    $form.Controls.Add($listbox)
    $toolTip.SetToolTip($listbox, "Liste der Sensitivity Labels")

    # === Optionen
    $chkWord = New-Object Windows.Forms.CheckBox
    $chkWord.Text = "Word-Bericht erzeugen"
    $chkWord.Checked = $true
    $chkWord.AutoSize = $false
    $chkWord.Width = 500  # vorher z. B. 480
    $chkWord.Location = New-Object Drawing.Point(10, 470)
    $form.Controls.Add($chkWord)
    $toolTip.SetToolTip($chkWord, "Erstelle einen Word Bericht (Word muss lokal installiert sein!")

    $chkPDF = New-Object Windows.Forms.CheckBox
    $chkPDF.Text = "PDF-Bericht erzeugen"
    $chkPDF.Checked = $false
    $chkPDF.AutoSize = $false
    $chkPDF.Width = 500
    $chkPDF.Location = New-Object Drawing.Point(10, 500)
    $form.Controls.Add($chkPDF)
    $toolTip.SetToolTip($chkPDF, "Erstelle einen PDF Bericht (Word Bericht muss vorher erstellt worden sein!")

    $chkMail = New-Object Windows.Forms.CheckBox
    $chkMail.Text = "Bericht per Mail senden"
    $chkMail.Checked = $false
    $chkMail.AutoSize = $false
    $chkMail.Width = 500
    $chkMail.Location = New-Object Drawing.Point(10, 530)
    $form.Controls.Add($chkMail)
    $toolTip.SetToolTip($chkMail, "Versende einen eMail Report mit allen Dateien (Aufruf muss mit optionalen SMTP Sendgrid Daten erfolgt sein!")

    $txtFolder = New-Object Windows.Forms.TextBox
    $txtFolder.Text = $LogFolder
    $txtFolder.Width = 360
    $txtFolder.Location = New-Object Drawing.Point(10, 560)
    $form.Controls.Add($txtFolder)


    $btnBrowseLog = New-Object Windows.Forms.Button
#    $btnBrowseLog.Text = "📁"
    $btnBrowseLog.Text = "..."
    $btnBrowseLog.Width = 30
    $btnBrowseLog.Location = New-Object Drawing.Point(380, 560)
    $form.Controls.Add($btnBrowseLog)

    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $btnBrowseLog.Add_Click({
        if ($folderDialog.ShowDialog() -eq 'OK') {
            $txtFolder.Text = $folderDialog.SelectedPath
        }
    })

    $btnLoadExcel = New-Object Windows.Forms.Button
    $btnLoadExcel.Text = "📄 Excel laden"
    $btnLoadExcel.Width = 100
    $btnLoadExcel.Location = New-Object Drawing.Point(420, 560)
    $form.Controls.Add($btnLoadExcel)
    $toolTip.SetToolTip($btnLoadExcel, "Alternativ kann eine lokale Excel Datei die Auswahl eingrenzen")

    $btnLoadExcel.Add_Click({
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Filter = "Excel-Dateien (*.xlsx)|*.xlsx"
        $fileDialog.Title = "Excel-Datei mit Labels laden (optional)"
        $fileDialog.CheckFileExists = $true
        $fileDialog.Multiselect = $false

        if ($fileDialog.ShowDialog() -eq 'OK') {
            Import-LabelsFromExcel -FilePath $fileDialog.FileName
            Update-LabelList -source $script:allLabels
            Log "✅ Excel-Datei geladen: $($fileDialog.FileName)" "SUCCESS"
        }
    })

# === Footer
$footerText = @"
Bei Fragen oder Problemen, wenden Sie sich bitte an:
$MSPPartner
$MSPNameAP
$MSPMail
$MSPURL
"@

$footer = New-Object Windows.Forms.Label
$footer.Text = $footerText
$footer.Font = New-Object Drawing.Font("Segoe UI", 8)
$footer.Location = New-Object Drawing.Point(10, 610)
$footer.Size = New-Object Drawing.Size(510, 100)  # ⬅️ angepasst (war 480)
$footer.AutoSize = $false
$form.Controls.Add($footer)

# === Buttons
$btnOK = New-Object Windows.Forms.Button
$btnOK.Text = "OK"
$btnOK.Width = 90
$btnOK.Location = New-Object Drawing.Point(10, 730)
$form.Controls.Add($btnOK)

$btnCancel = New-Object Windows.Forms.Button
$btnCancel.Text = "Abbrechen"
$btnCancel.Width = 90
$btnCancel.Location = New-Object Drawing.Point(110, 730)
$form.Controls.Add($btnCancel)

$btnCancel.Add_Click({
    $script:WasAbgebrochen = $true
    $form.Close()
})

# === List aktualisieren
function Update-LabelList {
    param([System.Collections.IEnumerable]$source)

    $listbox.Items.Clear()
    $script:labelMap.Clear()

    $filter = $txtFilterCode.Text.Trim()
    if ($filter -match "^\d{4}$") {
        $source = $source | Where-Object { $_.DisplayName -match "$filter$" }
    }

    if ($chkSortByCode.Checked) {
        $source = $source | Sort-Object {
            if ($_ -and $_.DisplayName -match "(\d{4})$") {
                [int]$matches[1]
            } else {
                0
            }
        } -Descending:($rbSortDesc.Checked)
    }

    foreach ($lbl in $source) {
        $display = "$($lbl.Priority): $($lbl.DisplayName)"
        $listbox.Items.Add($display) | Out-Null
        $script:labelMap[$display] = $lbl.Name
    }
}

$btnApplyFilter.Add_Click({
    Update-LabelList -source $script:allLabels
})

# === Auswahl übernehmen
$btnOK.Add_Click({
    $selectedNames = @()
    foreach ($item in $listbox.CheckedItems) {
        $labelName = $script:labelMap["$item"]
        if ($labelName) {
            $selectedNames += $labelName
        }
    }

    $form.Tag = @{
        Labels      = [string[]]$selectedNames
        ExportWord  = $chkWord.Checked
        ExportPDF   = $chkPDF.Checked
        SendReport  = $chkMail.Checked
        LogFolder   = $txtFolder.Text
    }

    $form.Close()
})

$form.Add_FormClosing({
    if (-not $form.Tag) {
        $script:WasAbgebrochen = $true
    }
})

# === Erster Listaufbau bei Start
Update-LabelList -source $script:allLabels
$form.ShowDialog() | Out-Null

if ($script:WasAbgebrochen) {
    Write-Host "❌ Abbruch durch Benutzer" -ForegroundColor Red
    exit 1
}

return $form.Tag
}



# === Verbindung aufbauen
function Connect-MFASessions {
    try {
        Log "🔐 Verbinde IPPS Session via MFA..." "INFO"
        Connect-IPPSSession -UserPrincipalName $UserPrincipalName
        Log "✅ IPPS verbunden" "SUCCESS"
    } catch {
        Handle-Error "❌ IPPS Verbindung fehlgeschlagen" $_
    }
    if (-not (Get-Command Get-Label -ErrorAction SilentlyContinue)) {
        Handle-Error "❌ Cmdlet 'Get-Label' ist nicht verfügbar" ([System.Exception]::new("Get-Label fehlt"))
    }
}

Connect-MFASessions

# IPPS-Connect erfolgreich:
    Log "✅ IPPS verbunden" "SUCCESS"



    Log "🧭⌛ Lade GUI Ansicht, einen Moment bitte ..." "INFO"
    # Optional: kannst du auch einen "⏳" oder "⌛" Emoji statt "🧭" verwenden
    # ..dann öffnet sich die GUI 🖥️




# === GUI nutzen? Dann Werte überschreiben
if ($UseLabelGUI) {
    Log "🧭 Lade GUI Ansicht, bitte einen Moment Geduld..." "INFO"
    Start-SplashInThread -UseProgressBar:$UseProgressBar `
                         -AutoCloseAfterSeconds:$AutoCloseAfterSeconds `
                         -CompanyLogoPath $CompanyLogoPath `
                         -CompanyLogoUrl $CompanyLogoUrl `
                         -CompanyLogoBase64 $CompanyLogoBase64 `
                         -LogoGIFUrl $LogoGIFUrl `
                         -ProductLogoBase64 $ProductLogoBase64 `
                         -LogoUrl $LogoUrl # <-- Produktlogo übergeben
                         

    try {
        # GUI starten (blockiert Hauptthread – Splash bleibt responsiv)
        $guiResult = Show-LabelSelectionForm `
            -CompanyLogoPath $CompanyLogoPath `
            -CompanyLogoUrl $CompanyLogoUrl `
            -CompanyLogoBase64 $CompanyLogoBase64 `
            -LogoUrl $LogoUrl `
            -ProductLogoBase64 $ProductLogoBase64
    } finally {
        # Splash garantieren schließen
        Stop-SplashThread
    }

    # Benutzerabbruch prüfen
    if (-not $guiResult -or $script:WasAbgebrochen) {
        Log "❌ GUI-Abbruch durch Benutzer – Skript wird beendet." "ERROR"
        exit 1
    }

    # GUI-Werte übernehmen
    $LabelNames = @()
    if ($guiResult.Labels -is [System.Collections.IEnumerable]) {
        foreach ($ln in $guiResult.Labels) {
            if ($ln -is [string]) {
                $LabelNames += $ln
            } else {
                Log "⚠️ Ungültiger Labelname ignoriert: $ln" "DEBUG"
            }
        }
    }

    $ExportWord  = $guiResult.ExportWord
    $ExportPDF   = $guiResult.ExportPDF
    $SendReport  = $guiResult.SendReport
    $LogFolder   = $guiResult.LogFolder

    # Zielpfade aktualisieren
    $DatumJetzt       = Get-Date -Format 'yyyyMMdd_HHmmss'
    $CreatedLabelsCsv = [System.IO.Path]::Combine($LogFolder, "Erstellte_Labels_$DatumJetzt.csv")
    $StatusReportCsv  = [System.IO.Path]::Combine($LogFolder, "Label_Status_$DatumJetzt.csv")
    $LogFile          = [System.IO.Path]::Combine($LogFolder, "LabelReport_$DatumJetzt.log")
}




# (SKRIPT FÄHRT HIER FORT MIT LABEL-LADEN, EXPORT, WORD/PDF, MAIL usw.)

# === Labels laden (Name, GUI oder Priorität) ===
# Log "ℹ️ Prioritätsfilter wird geprüft (LabelNames leer, Priorities gesetzt: '$Priorities')" "DEBUG"


# === Fallback prüfen: Ist überhaupt eine Labelquelle definiert?
if (
    (-not $UseLabelGUI) -and
    ($LabelNames.Count -eq 0 -or $null -eq $LabelNames) -and
    (-not $Priorities) -and
    (-not $UseExistingLabels)
) {
    Handle-Error "❌ Keine gültigen Label-Quellen angegeben. Verwende -UseLabelGUI, -LabelNames oder -Priorities." ([System.Exception]::new("Keine Eingabequelle definiert"))
}


# === Labels laden (GUI, Namen, Prioritäten)
$labels = @()

if ($UseExistingLabels -or $LabelNames.Count -gt 0 -or $UseLabelGUI) {
    Log "📥 Lade Sensitivity Labels..." "INFO"

        if ($UseExistingLabels -and
        $LabelNames.Count -eq 0 -and
        -not $UseLabelGUI -and
        $Priority -eq 0 -and
        $PriorityMin -eq 0 -and
        $PriorityMax -eq 0
    ) {
        Log "⚠️ Es wurden keine Prioritätsfilter oder Labelnamen angegeben – alle Labels werden geladen." "WARNING"
    }


    if ($LabelNames.Count -gt 0) {
        foreach ($name in $LabelNames) {
            try {
                $label = Get-Label -Identity $name
                $labels += $label
                Log "✅ Label geladen: $name" "SUCCESS"
            } catch {
                Log "❌ Fehler beim Laden von Label '$name': $_" "ERROR"
            }
        }
    } else {
        $allLabels = Get-Label
        $filtered = @()

        if ($Priority -gt 0) {
            $filtered = $allLabels | Where-Object { $_.Priority -eq $Priority }
            Log "🎯 Filter: Priorität = $Priority" "DEBUG"
        } elseif ($PriorityMin -gt 0 -and $PriorityMax -gt 0) {
            $filtered = $allLabels | Where-Object { $_.Priority -ge $PriorityMin -and $_.Priority -le $PriorityMax }
            Log "🎯 Filter: Prioritäten von $PriorityMin bis $PriorityMax" "DEBUG"
        } elseif ($PriorityMin -gt 0) {
            $filtered = $allLabels | Where-Object { $_.Priority -ge $PriorityMin }
            Log "🎯 Filter: Prioritäten ab $PriorityMin" "DEBUG"
        } else {
            $filtered = $allLabels
            Log "🎯 Keine Prioritätsfilter – alle Labels geladen" "DEBUG"
        }

        foreach ($entry in $filtered) {
            try {
                $label = Get-Label -Identity $entry.Name
                $labels += $label
                Log "✅ Label geladen: $($entry.Name)" "SUCCESS"
            } catch {
                Log "❌ Fehler bei Get-Label -Identity '$($entry.Name)': $_" "ERROR"
            }
        Log "📄🖥️ Erzeuge Word-Bericht... Bitte warten...." "INFO"
        }
    }

    if (-not $labels -or $labels.Count -eq 0) {
        Handle-Error "⚠️ Keine Labels gefunden" ([System.Exception]::new("Leere Ergebnismenge"))
    }
}




# === Label-Metadaten exportieren als CSV/Excel ===
try {
    $exportData = foreach ($label in $labels) {
        [PSCustomObject]@{
            Name = $label.DisplayName
            Tooltip = $label.Tooltip
            Priority = $label.Priority
            Description = $label.Description
            ContentType = $label.ContentType
            Enabled = $label.Enabled
            Id = $label.ImmutableId
            PublisherName = $label.PublisherName
            LastModified = $label.LastModifiedDateTime
        }
    }

    if (-not $DryRun) {
    $exportData | Export-Csv -Path $StatusReportCsv -NoTypeInformation -Encoding UTF8
    $exportData | Export-Excel -Path ($StatusReportCsv -replace '.csv$', '.xlsx') -AutoSize
    Log "📄 CSV und Excel mit Label-Metadaten exportiert." "SUCCESS"
} else {
    Log "🧪 DryRun – CSV/Excel-Export übersprungen." "DEBUG"
}

    # $exportData | Export-Csv -Path $StatusReportCsv -NoTypeInformation -Encoding UTF8
    # $exportData | Export-Excel -Path ($StatusReportCsv -replace '.csv$', '.xlsx') -AutoSize

    Log "📄 CSV und Excel mit Label-Metadaten exportiert." "SUCCESS" -Encoding utf8
} catch {
    Handle-Error "Fehler beim Export der Label-Metadaten" $_
}


Log "📑 Export in Word/PDF gestartet mit $($labels.Count) Label(s)." "INFO"
foreach ($l in $labels) {
    Log "➡️  Label: $($l.DisplayName)" "DEBUG" -Encoding utf8
}

# === Prüfung bzgl. leerer Labels ===
if (-not $labels -or $labels.Count -eq 0) {
    Log "⚠️ WARNUNG: Keine Labels im Speicher – Testlabel wird eingefügt." "WARNING" -Encoding utf8

    $labels = @(
        [PSCustomObject]@{
            DisplayName = "Test Label"
            Tooltip = "Tooltip-Test"
            Priority = 1
            Description = "Testbeschreibung"
            ContentType = "File"
            Enabled = $true
            ImmutableId = "test-id"
            PublisherName = "TestPublisher"
            LastModifiedDateTime = (Get-Date)
            LabelActionsJson = '[{"Type":"Protect","Settings":{"Encryption":"Enabled"}}]'
        }
    )
}


# === Dokumentenerzeugung (Word/PDF) ===
if ($ExportWord -or $ExportPDF) {

#    Start-ConsoleSpinner -Message "📄 Erzeuge Word-Bericht..."

    if ($DryRun) {
        Log "🧪 DryRun – Word/PDF-Erstellung übersprungen." "WARNING"
    } else {

    try {
        # === Pfade vorbereiten ===
        if (-not (Test-Path -Path $LogFolder)) {
            New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
        }

        # $WordPath = [System.IO.Path]::Combine($LogFolder, "PurviewExport_${cleanUser}_$DatumJetzt.docx")
        # $PdfPath  = [System.IO.Path]::Combine($LogFolder, "PurviewExport_${cleanUser}_$DatumJetzt.pdf")
        $WordPath = [System.IO.Path]::Combine($LogFolder, "PurviewExport_$Tenantdomain_$DatumJetzt.docx")
        $PdfPath  = [System.IO.Path]::Combine($LogFolder, "PurviewExport_$Tenantdomain_$DatumJetzt.pdf")

        Write-Host "📄 WordPath = $WordPath"
        Write-Host "📄 PdfPath  = $PdfPath"

                # === Word starten ===
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true

#        $doc = $word.Documents.Add()
#        $selection = $word.Selection

#        # === Kopfbereich ===
#        $selection.Style = "Titel"
#        $selection.TypeText("M365 Purview Sensitivity Labels – $Tenantdomain")
#        $selection.TypeParagraph()
#        $selection.TypeText("Erstellt am: $DatumAnzeige")
#        $selection.TypeParagraph()
#        $selection.TypeText("Autor: Microsoft Purview Automation - BDO Digital GmbH")
#        $selection.InsertNewPage()

#        # === Inhaltsverzeichnis ===
#        $tocRange = $selection.Range
#        # $doc.TablesOfContents.Add($tocRange, $true, 1, 3) | Out-Null

        $doc = $word.Documents.Add()
        $selection = $word.Selection

    # === Deckblatt: Abschnitt 1 - Company Logo (oben zentriert) ===
    if ($CompanyLogoPath -and (Test-Path $CompanyLogoPath)) {
        $selection.ParagraphFormat.Alignment = 1
        $selection.InlineShapes.AddPicture($CompanyLogoPath, $false, $true)
        $selection.TypeParagraph()
        $selection.TypeParagraph()
    }
    elseif ($CompanyLogoBase64 -and $CompanyLogoBase64.Length -gt 100) {
        $bytes = [Convert]::FromBase64String(($CompanyLogoBase64 -replace '^data:image\/[a-z]+;base64,', ''))
        $LogoFile = [System.IO.Path]::Combine($LogFolder, "CompanyLogo_$DatumJetzt.png")
        [IO.File]::WriteAllBytes($LogoFile, $bytes)
        $selection.ParagraphFormat.Alignment = 1
        $selection.InlineShapes.AddPicture($LogoFile, $false, $true)
        $selection.TypeParagraph()
        $selection.TypeParagraph()
    }

    # === Abschnitt 2 - Titel und Metadaten ===
        $selection.Style = "Titel"
        $selection.TypeText("M365 Purview Sensitivity Labels – $Tenantdomain")
        $selection.TypeParagraph()
        $selection.TypeText("Erstellt am: $DatumAnzeige")
        $selection.TypeParagraph()
        $selection.TypeText("Autor: Microsoft Purview Automation - BDO Digital GmbH")
        $selection.TypeParagraph()
        $selection.TypeParagraph()
    
    # === Abschnitt 3 - Produktlogo (unten zentriert) ===
    if ($LogoUrl -and $LogoUrl.Trim() -ne "") {
        $LogoFile = [System.IO.Path]::Combine($LogFolder, "ProductLogo_$DatumJetzt.png")
        $wc = New-Object Net.WebClient
        $wc.DownloadFile($LogoUrl, $LogoFile)
        $selection.ParagraphFormat.Alignment = 1
        $shape = $selection.InlineShapes.AddPicture($LogoFile, $false, $true)
        $shape.Width  = $shape.Width * 0.5
        $shape.Height = $shape.Height * 0.5
    }
    elseif ($ProductLogoBase64 -and $ProductLogoBase64.Length -gt 100) {
        $bytes = [Convert]::FromBase64String(($ProductLogoBase64 -replace '^data:image\/[a-z]+;base64,', ''))
        $LogoFile = [System.IO.Path]::Combine($LogFolder, "ProductLogo_$DatumJetzt.png")
        [IO.File]::WriteAllBytes($LogoFile, $bytes)
        $selection.ParagraphFormat.Alignment = 1
        $shape = $selection.InlineShapes.AddPicture($LogoFile, $false, $true)
        $shape.Width  = $shape.Width * 0.5
        $shape.Height = $shape.Height * 0.5
    }


    # === Neue Seite für Inhaltsverzeichnis ===
        $selection.InsertNewPage()


        # === Inhaltsverzeichnis ===
        $tocRange = $selection.Range


        # Inhaltsverzeichnis an den Anfang setzen
        # $range = $doc.Range(0,0)

        # TOC mit Hyperlinks, Ebene 1–3
        # Beispiel: Fügt ein Inhaltsverzeichnis hinzu mit den gewünschten Optionen
        $missing = [System.Type]::Missing

        $doc.TablesOfContents.Add(
            $tocRange,
            $true,      # UseHeadingStyles
            1,          # UpperHeadingLevel
            3,          # LowerHeadingLevel
            $true,      # UseFields
            "",         # TableID
            $true,      # RightAlignPageNumbers
            $true,      # IncludePageNumbers
            $missing,   # AddedStyles
            $true,      # UseHyperlinks
            $false      # HidePageNumbersInWeb
            )

        $selection.InsertNewPage()

            # === Fußzeile korrekt formatieren ===
                $section = $doc.Sections.Item(1)
                $section.PageSetup.DifferentFirstPageHeaderFooter = $true
                $footer = $section.Footers.Item([Microsoft.Office.Interop.Word.WdHeaderFooterIndex]::wdHeaderFooterPrimary)
                $footerRange = $footer.Range

            # Leeren und Tabstop hinzufügen
                $footerRange.Text = ""
                $footerRange.ParagraphFormat.TabStops.ClearAll()
                $footerRange.ParagraphFormat.TabStops.Add(16 * 28.35, [Microsoft.Office.Interop.Word.WdTabAlignment]::wdAlignTabRight)

            # Linker Text: "Seite: [Feld]"
                $leftRange = $footerRange.Duplicate
                $leftRange.InsertAfter("Seite: ")
                $leftRange.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
                $leftRange.Fields.Add($leftRange, [Microsoft.Office.Interop.Word.WdFieldType]::wdFieldPage)

            # Rechter Text mit Tab + Datum
                $rightRange = $footerRange.Duplicate
                $rightRange.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
                $rightRange.InsertAfter("`terstellt am: $DatumAnzeige")

            # Formatierung
                $footer.Range.Font.Name = "Arial"
                $footer.Range.Font.Size = 9

        
# === Labels einfügen ===
foreach ($label in $labels | Sort-Object Priority) {
    $selection.Style = "Überschrift 1"
    $selection.TypeText($label.DisplayName)
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = "Standard"
    $selection.Font.Size = 11
    $selection.Font.Bold = $false

    # === Basisinformationen einfügen ===
    $properties = @{
        "Tooltip"         = $label.Tooltip
        "Priority"        = $label.Priority
        "Description"     = $label.Description
        "ContentType"     = $label.ContentType
        "Enabled"         = $label.Enabled
        "Id"              = $label.ImmutableId
        "PublisherName"   = $label.PublisherName
        "LastModified"    = $label.LastModifiedDateTime
    }

    foreach ($key in $properties.Keys) {
        $selection.TypeText("${key}: ")
        $selection.Font.Bold = $true
        $selection.TypeText("$($properties[$key])")
        $selection.Font.Bold = $false
        $selection.TypeParagraph()
    }


# === LocaleSettings (strukturierte Sprachwerte aus JSON) ===
# === LocaleSettings (strukturierte Sprachwerte bei nicht-standardisiertem JSON) ===
foreach ($prop in $label.PSObject.Properties) {
    if ($prop.Name -eq "LocaleSettings") {
        try {
            $selection.Style = "Überschrift 2"
            $selection.TypeText("Locale Settings")
            $selection.TypeParagraph()

            # Einige Backends liefern mehrere JSON-Objekte hintereinander als String
            $raw = $prop.Value -replace '}\s*{','}|SPLIT|{'
            $jsonParts = $raw -split '\|SPLIT\|'

            $locales = @{}

            foreach ($part in $jsonParts) {
                $obj = $part | ConvertFrom-Json -ErrorAction Stop
                $localeKey = $obj.LocaleKey  # z. B. displayName, tooltip

                foreach ($entry in $obj.Settings) {
                    $lang = $entry.Key
                    $value = $entry.Value

                    if (-not $locales.ContainsKey($lang)) {
                        $locales[$lang] = @{}
                    }
                    $locales[$lang][$localeKey] = $value
                }
            }

            # Ausgabe pro Sprache
            foreach ($lang in ($locales.Keys | Sort-Object)) {
                    $selection.Style = "Überschrift 3"
                    $selection.TypeText("Sprache: $lang")
                    $selection.TypeParagraph()
                    $selection.Style = "Standard"

                $entry = $locales[$lang]

                if ($entry.displayName) {
                    $selection.TypeText("   • Name: $($entry.displayName)")
                    $selection.TypeParagraph()
                }
                if ($entry.tooltip) {
                    $selection.TypeText("   • Tooltip: $($entry.tooltip)")
                    $selection.TypeParagraph()
                }
                if ($entry.description) {
                    $selection.TypeText("   • Description: $($entry.description)")
                    $selection.TypeParagraph()
                }

#                $selection.TypeParagraph()
            }
        }
        catch {
            $selection.TypeText("⚠️ Fehler beim Verarbeiten von LocaleSettings")
            $selection.TypeParagraph()
            $selection.TypeText($_.Exception.Message)
            $selection.TypeParagraph()
        }
    }
}

    # === LabelActionsJson (falls vorhanden) ===
# === LabelActionsJson (strukturierte Verarbeitung mehrerer JSON-Objekte im String) ===
# if ($label.PSObject.Properties["LabelActionsJson"] -and $label.LabelActionsJson) {
if ($label.PSObject.Properties["LabelActions"] -and $label.LabelActions) {
    try {
        $selection.Style = "Überschrift 2"
        $selection.TypeText("Label Actions")
        $selection.TypeParagraph()

        # Schritt 1: Rohdaten auslesen
        $raw = $label.LabelActions

        # Schritt 2: JSON-Objekte voneinander trennen
        $normalized = $raw -replace '}\s*{','}|SPLIT|{'
        $jsonParts = $normalized -split '\|SPLIT\|'

        foreach ($part in $jsonParts) {
            $action = $part | ConvertFrom-Json -ErrorAction Stop

            $type = $action.Type
            $subtype = if ($action.SubType) { $action.SubType } else { "-" }

            $selection.Style = "Überschrift 3"
            $selection.TypeText("Aktion: $type")
            $selection.TypeParagraph()
            $selection.Style = "Standard"
            $selection.TypeText("Untertyp: $subtype")
            $selection.TypeParagraph()

            foreach ($setting in $action.Settings) {
                $key = $setting.Key
                $value = $setting.Value

                # Spezialfall: rightsdefinitions (JSON-Array als String)
                if ($key -eq "rightsdefinitions" -and $value -match '^\s*\[') {
                    $selection.TypeText(" - ${key}:")
                    $selection.TypeParagraph()
                    try {
                        $rights = $value | ConvertFrom-Json -ErrorAction Stop
                        foreach ($entry in $rights) {
                            $selection.TypeText("     • Identity: $($entry.Identity)")
                            $selection.TypeParagraph()
                            $selection.TypeText("       Rights: $($entry.Rights)")
                            $selection.TypeParagraph()
                        }
                    } catch {
                        $selection.TypeText("     ⚠️ Fehler beim Parsen von rightsdefinitions")
                        $selection.TypeParagraph()
                    }
                }
                else {
                    $selection.TypeText(" - ${key}: $value")
                    $selection.TypeParagraph()
                }
            }

#            $selection.TypeParagraph()
        }
    }
    catch {
        $selection.TypeText("⚠️ Fehler beim Verarbeiten von LabelActionsJson")
        $selection.TypeParagraph()
        $selection.TypeText($_.Exception.Message)
        $selection.TypeParagraph()
    }
}
    $selection.InsertNewPage()
}
       
        # === Inhaltsverzeichnis aktualisieren ===
        $doc.TablesOfContents.Item(1).Update()

        # === Speichern ===
        if ($ExportWord) {
            $doc.SaveAs($WordPath, 16)
            Log "📄 Word gespeichert: $WordPath" "SUCCESS"
        }

        if ($ExportPDF) {
            $doc.SaveAs($PdfPath, 17)
            Log "📄 PDF gespeichert: $PdfPath" "SUCCESS"
        }

        # === Dokument schließen ===
        $doc.Close($false)
        $word.Quit()

        # === COM Cleanup ===
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selection) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
    catch {
        Write-Host "❌ Fehler beim Word-/PDF-Export" -ForegroundColor Red
        Write-Host "   Fehler: $($_.Exception.GetType().FullName)" -ForegroundColor DarkRed
        Write-Host "   Meldung: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "   StackTrace: $($_.Exception.StackTrace)" -ForegroundColor Gray

        $logPath = "$env:TEMP\PurviewExportError.log"
        $_.Exception | Out-File -FilePath $logPath -Append -Encoding UTF8
        Write-Host "Fehlerdetails gespeichert unter: $logPath"
    }
}
# Stop-ConsoleSpinner
}

# === Mailversand via SendGrid (UTF-8 gesichert)
if ($SendReport) {
    $SMTPServer = "smtp.sendgrid.net"
    $SMTPPort = 587
    $SMTPUsername = "apikey"
    $SMTPPassword = ""
    $MailFrom = "support-bdodigital@bdo-digital.eu"
    $MailTo = $MailToPrimary  # <-- Übergabe aus Parametern!
    $MailCC = $MailToSecondary # <-- Optional für CC

    $Subject = "Report ueber Aktualisierung im Tenant $Tenantdomain vom $DatumAnzeige"

    $BodyRaw = @"
Diese Mail wurde versendet, weil entweder:<br>
A) Sensitivity Labels im Tenant <b>$Tenantdomain</b> erstellt oder aktualisiert wurden. (oder)<br>
B) Ein Report ueber ausgewaehlte Labels angefordert wurde. <br>
Weitere Informationen finden Sie im Dateianhang (Word, PDF, CSV).<br><br>
Mit freundlichen Gruessen<br>
Microsoft Security & Compliance Automation
"@

    # Body temporär UTF-8 enkodieren (als Workaround)
    $utf8BodyPath = [System.IO.Path]::GetTempFileName()
    [System.IO.File]::WriteAllText($utf8BodyPath, $BodyRaw, [System.Text.Encoding]::UTF8)
    $BodyEncoded = Get-Content -Path $utf8BodyPath -Encoding UTF8 | Out-String

    # Anhänge sammeln
    $attachments = @()
    if ($ExportWord -and (Test-Path $WordPath)) { $attachments += $WordPath }
    if ($ExportPDF -and (Test-Path $PdfPath)) { $attachments += $PdfPath }
    if (Test-Path $StatusReportCsv) { $attachments += $StatusReportCsv }

    try {
        # Hauptempfänger und optional CC für Send-MailMessage
        $mailParams = @{
            From        = $MailFrom
            To          = $MailTo
            Subject     = $Subject
            Body        = $BodyEncoded
            BodyAsHtml  = $true
            SmtpServer  = $SMTPServer
            Port        = $SMTPPort
            Credential  = (New-Object PSCredential($SMTPUsername, (ConvertTo-SecureString $SMTPPassword -AsPlainText -Force)))
            Attachments = $attachments
            UseSsl      = $true
        }
        if ($MailCC -and $MailCC -ne "") {
            $mailParams.Cc = $MailCC
        }

        Send-MailMessage @mailParams

        Log "✅ Mail erfolgreich gesendet an $MailTo$(if ($MailCC) { " (CC: $MailCC)" })" "SUCCESS"
    }
    catch {
        Handle-Error "Fehler beim Mailversand" $_
    }
    finally {
        Remove-Item $utf8BodyPath -Force -ErrorAction SilentlyContinue
    }
}


Disconnect-ExchangeOnline -Confirm:$false
Log "✅ Exchange Online Sitzung beendet." "INFO"

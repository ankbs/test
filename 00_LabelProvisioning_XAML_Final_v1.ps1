﻿<#
.SYNOPSIS
    Microsoft Purview Label Provisioning GUI mit DeepL-Übersetzung und flexibler Excel-Importfunktion (lokal, URL, Dialog).
.DESCRIPTION
    Dieses PowerShell-Skript stellt eine WPF-GUI bereit, mit der Microsoft Purview Labels aus Excel-Dateien (Buchungskreis) provisioniert, angezeigt, exportiert und in verschiedene Sprachen (DeepL) übersetzt werden können.
    Die Excel-Datei kann
      - automatisch im lokalen Pfad gesucht werden (Dateiname: "Buchungskreis*.xls*"),
      - per Startparameter (-ExcelFilePath oder -ExcelFileUrl) übergeben werden
      - oder bei Bedarf per Dateiauswahldialog ausgewählt werden.
    Logos (Base64) können als Startparameter übergeben werden.
    Excel wird über das ImportExcel-Modul eingelesen (keine lokale Office-Installation notwendig).

    Deepl API: Bitte die Zeile am Anfang suchen
    # ======= DeepL API Key (bitte setzen) =======
    und dann den eigenen (kostenfreien) API Key eintragen.
    Sign Up: https://www.deepl.com/en/pro#developer
    API Key Access: https://www.deepl.com/en/your-account/keys

    Sensitive Label Variablen:
    Diese werden im Block
    # ======= Vorschau- und Export-Logik =======
    gesetzt und können bei Bedarf angepasst werden

    Logdaten / Exportdaten:
    In dem ausführenden Verzeichnis werden jeweils entsprechende Unterordner erstellt.
    Nach der Übersetzung öffnet sich der Dateiexplorer entsprechend

.EXAMPLE
    # Automatische Dateisuche im Skriptverzeichnis (Standard)
    .\LabelProvisioning_XAML_ImportExcel.ps1

    # Mit explizitem Excel-Pfad
    .\LabelProvisioning_XAML_ImportExcel.ps1 -ExcelFilePath "C:\Temp\Buchungskreis2024.xlsx"

    # Mit Download von URL
    .\LabelProvisioning_XAML_ImportExcel.ps1 -ExcelFileUrl "https://server/meineBuchungskreise.xlsx"

    # Mit Base64-Logos
    .\LabelProvisioning_XAML_ImportExcel.ps1 -CompanyLogoBase64 (Get-Content .\logo1.b64 -Raw) -ProductLogoBase64 (Get-Content .\logo2.b64 -Raw)

.AUTHOR
    Michael Kirst-Neshva - michael_kirst@hotmail.com
#>

[CmdletBinding()]
param(
    [string]$MSPPartner = "",
    [string]$MSPNameAP  = "",
    [string]$MSPMail    = "",
    [string]$MSPURL     = "",
    [string]$MSPNameEU  = "",
    [string]$CompanyLogoBase64 = "data:image/png;base64,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r1+fSqUmTZo0evRoQkhLS8uCBQtGjhw5YsSIDes3zJ8/f9asWWeddVYikWhtbV22bNl99913880367qez+cJIZZlLV68uL29/fLLL29qaspms6tWrXrmmWcopZqmpdPpVatWXXzxxblc7vXXX+/u7p46deqcOXMYYxs2bHj44YdFUTz1lFOXPb9syZIlc+bMGTVqlKIou3btWrly5csvv3z22We7SgOBgNtNVKaBgZO/rJiWefpZU2dcd5X7oKw0/hNzrrj8v/+wAJo1JBgbHE0AYASC2/HbtEZQFVDkjCUPLbrrpttqPn8e08o0nsS4MXdEGu+67rqutzbG4zXduQze1TN72sYOboqdPg5dBsLKeV/+4glKTUoJYcvOZX/562133lZ/2hTksojHTzz19Ht+dOePv/eD2edfrDhoGtLw1au/lhg7CgEBNk6++MLR/7Ni+TPPfnboMInQEeOa59xxAwAwDof837k//sU3rv3LQw9fcvU1j/36oWuv+MqIr12BCMAxvib542Fjbr32m+sWPTnxk9Pat+78/A3XT77qMoQF2PyOr3/zzFOnfPo710HUEVFHnXHKqHETE13G2tdXo6uMsACLOYRbtmXbdjAYnD17dlNTUyaTCQQC06ZN27Bhw+rVq4cNG/anP/1p9uzZl156qWmasiyPHz++ubn55ptvXrFixcyZM+vr62VZ3rhx48qVK+fOnZtKpWRZjkQijY2Nix9Z3NXVRQiRehFFURTFSy65ZFjjMEEUAJx77rkbN25sbW2dMmXKU089ddlll02bNs2xnUw2M2nSpGHDhrW2th5iIuuI+v72zmS8ZsTE8WAA53AYapT42KZMPgcuJOSgAgpeMfijNBQMUSogly2bxo63t7x2x5ssKKvBADRrEA3s6mjb+s47NXWpkBqEqk7/l0/f8oNbOmZ/ZcKJExyHDR/eNOacc5BIZF55GY7T8sqq1lXrisVioiahKIphGMVMbueaNXX1dWoskhjVBEUAAN2ETNVYxDZMxBStVD535gxQGJahCDJkQLZHnHrS1jdbkDeK+Xx7e7ux4NF9xWypWASQqq/LF4sb1785ceq0kWNGT5p2JkKCxqxA3jK78+dffSWigCWUiwXR5jKTJ0+evHLVKigEIZlxxjjjnLuddH19vWM7NTU1nHPHdjjnbgdsmuZpp51WLBbD4bCu65xzWZLHjRu3e/duRVE0Tcvn8y0tLVOmTBkyZIimaX0pevIpk1e+tlIQBFmWZVmWJMn9S319vcMcx3TcUYKqquVyecuWLbFYrLm52bIsSZJSqRSAZDI5duxYHJaBp31UtVgqyaoCGU6xAApQtjPbqYgSZEFiKHfnUPkENCUBVRUpNQvFsKIGRLkumUomktlCnqiSLuJ7d9xy2pln2LZDBIqO9poppz2w+JHZF8y2bbutbd+SJUtu/NK/bX3qSaOkpRJJVVZUVW1oaBBEwWEsFI1cf8N3myZP7M5mRUWGgIKpaaYBRYYkhkKhQqGArAnLjoci7ljMKpXAgdqwFAlyAthm0TKYLHRmM4qiJGNxWZSy+dxll19++Re/iEymbBlMIJquMc4RDFpFrZjNosTNQjEYCMmJKIKqLMtEEZEMQCAQBXfhQBRFwzAkUSqWirqul0ol0zKLxaKu69lsVlGUzs5OxlihUFBVNRAICKIQjUbdVWZBEILBoGVZsiybpikIguuecz5q5CjHcQghbtD35YC+FWpJkpjD3LFCIpEwDEPXdXfICcAd9kuSdPjZBXHAaTjLrkmmYDjo1IVQBCoBIMqqAorOAtMtEQQiEWwGBigUWWPXrl2TTFNO1iKvjT3nXDQPRliALMKwkSmvXf5CY01UCqoW4TDNvc8tG3L2/xl9+aWjVQKbw2Av3/er11e8fNlXv2qa5pkXzcK4EXAAxpErQFI2L1+O0U18w4ayoTsydSQqOYBugcp2vqQMSYEgHomuWvrCueediQgoFaA7EITtW7Yqw+vRlMpRZ+yUSYnZZ4MZAIUqwcDWpctBKURetAxBUQJchCOAYujgIZtfeG34OZ+SAzEYNrI6gtFXX36FcgDQuWNrRjgYlGUZlKjBQNnQQpGISHvCKR6PU0IbGhps206n0xMmTKis2q1bt9bU1Kiq6g4Vx44d++ijj8qy7KYN9/bvxRdflCSJc57NZt27O865YRh9Q30AjDNCiK7rdXV1lNJdu3Y1NDS4d/mc80KhsHnz5jFjxhyY2StGAANEPwEourKZUCiMoIogAWA5XDeNvZ0dABt9xuQd7fv2rliFTBmcYnfXorlzu7LdRJURlE89b9q873+Xd6fBOEQC23nkod/9/o9/AGc24UVuoSby5DN/e+IX82Fb0AxIBAJty6aFSACnThgyfszcu+aZrTshAyECrTz/9lsffnQxRNgirR0yyCH9Gi2lAhMIBFIyjeX/88KGPz2OLk2AAE6X3v3rt9at/9evXglqTb/80lvunde+cQOoiLCEzsxfHvzVj+/9GSKqJQKKBIEgKMF2oOkXz/nXZ5Y+t+43v8PeDmTL+b0dv/nuTS+88IIcCgBQiOBOVzjgpml2ZzIAREpN28rlcpZldXd393XPTz31VEtLiyiKblAuWrRo586dM2fOzOfztbW1jLFx48bJsrxw4cKyVnZ17tq1a/ny5aZpEkKi0ajbxTiOc3Dnreu6O000c+bMJ5544rXXXuvq6iqXy9u3b1+4cGE2m+2L/kOu7Q088o+FlZpot14c4RaxwOBEEvHGCZ9AiOCsEz/z1SvmfevGkUMb5US0qGuzZ83qNIqaTBAgn777ltLcu67/0tcaGxsTtcmWlk1DBw+ef8/9kOS96c7gsHrUqP/5k1vn3nTru9deJ3ISDAbbujp1OLfe+1MErCvv+N493731tm9c35Cssy1rT65rxLixc+/+GYoFXYRBOSGUAsSdeBBBw6oQUiHQmsbBZ51x5qvr1jy24tm6VG2mraPY1vWdG29Ijh3JwS/+9n9kiXX/LT8anKxlIu0sZAOJ2J333Y26uCPwWEN9e1dn/aghUESL2U2zpl43/yd/+PVvH1/xrKIogiKfN+3sTzVP/tWiBVamUAiQgCjbhsk5VxRleFOTqgYMy7RtJxaLMcbi8bg7bp81axaAxx57bNmyZaqq7t69OxwOX3311aNHj85ms5ZlRSKRcDh8xRVXPP300w8++KAgCIwxXdfPP//8p59+OpPJFItFd16BMVYZ+i6mabpTCDNnzgTw97//nVJKKVUU5aKLLopEIn3J4JD6Sd9KaD8cwIa1u01KJSBSSIJuamo0WM6VlLIpyBJTRcoFtOzasW6jLmLMyeOF0SOhFwqCE0nUoKSDqtiwdeOaN8SgWjdkcM0pkyFwhASrWOgu5GqTKSdblAKRrtfX723d4TisYdTw2skTUBs2KSt0Z5JipGvVm7u3tCqKMurk8cqIRogEjCFfYgKsoTUmhcygWIAjYPdeDE6BCkgXQQWEpNZ1b7y7bWdj3aARU07HINUp60JIZYxRSrEt/c6rqzqLuUEjh42cejrCglkoM8NErqzW1CAgGpahKIFiW6dStqRQDbrzKJYwuA6p2KZFj/3hb0/c+bdFeUuTRZEwrkiyYZmarscjUdabSx3bSXeno9Gou5nM1bl9+/ZyuVxXVxePxyVJikajuq7ncrlgMNg3J7h69ep0Oh2Px5ubm1VVNU3TVV4oFFKpVFtbWyAQUCtmnCiluVxOVdVQKLRx48YJEyak0+lMJhOJROLxuKIoCxYsIIRceeWVnHPmMHLQ4vWh9POKH7L/NSYAgNC3Cs4Am+9flhFJTzn0Lhf1LsBA6N/J0N4yvHcvposACD2rNRIqfj8TqTgrBlCYAnMIJBDRAWwCziEQOAyM9hyH9r8EueIgfcs8pOf3N3FwwgFO9p+8TP5+96+3vPbGt+/5OWQJigDO+Y5t3/7O9TMuvehT//Z5IRoSB15UZoz1VG7/m64Dpt4Ov1bEOT9ADeOsskDlESih8/9rfigUuuqqqwgh7ge99NJLjzzyyO233+52MQfG/YD60VtrvPKzXf09p+BevANQQOAA4LgLeKzi/+I9C4ZWb0mZ9bzISM8nUL7fMatoak5v43Fbi+A67S1pg7PelUih71ist7XSXsGVS5SkX2t2LRPO0Xv5BGT/wyEiQZf+4Pdv3rltx6hxnwjURPPl4iurXz/trDO/fv23xNr44dfJ+9LssVrTG2hLTuXr6XT6z3/+s2mabr/T2dlZKpWmTp06Y8aMQy4RHVY/OfBj3EpmbjVV7J6jQL+11b4gI9ABAG6qMl2FJsB7ftnSgc2F9IQu5QDpaTG9Te3AQKP9z/SARWMcsjA/6CgMqLz2yhxjMzgMHNs3bCgbejaXC8ajwXBo6PCmcCIOAvf2+lBCeo5z8Pzah0A6ne7s7Eyn04IgpFKpQYMGhcPhQ7rH4fSTykrkvHc7R69+gVZWemWO7Y1gCjhAGaBACACHRSAAtEJ/T8lD6gccAmf/6eyvWxfhwNzUL7UfEJlkoNubSv39VvoJGOe2xQgEUQKBe7ulm4bJWTQQPJyB46c/n8+LoqjIiiAKpmkahhGJRDDA6jCOQD9HT20ygPTN77nRv3/fBT8ox/bWrdObt9FnpaLf3b+Vg1d+aL+BwUAcplL5od49TKo+ZKKuTKp9InvaK++pgfc4ifc80WNN5dj+yJvde+71Y+hvol9VHryxrvcVesjRHj3U0clBr7xX1dH+JSs5qgpn+zeX9L+Iin+4VUp63VcZA+ln7miK7v9nTwWxyujvy7kfLv/sL+utgHwQEUqOQ53gffU1B+kn6NlO1e/P478rzeeD4BD7/Fm/J2T2Qw/4L1CtzaIyu1QuagqomNv46F36++j7q/bhRZ8joeoe8/Aw76PvP9Lop0dZ3udjwWGiv3ekD1RY5/1awEG94CEe7eADHLii/EAM2NYG+GKGY0X/zz3oAw6+3znkcy8fB44omv2Qr1b8vv9Y8LGdDnpP/R9A5PMjyo1+yvkQENkA/S8dqP4HaOlHNBNX+TzoAC1gwEllfthjoqLkh8PHpHc/PH6MeZqj7/uPttVXrteS93uQA452vCK+6vgQo/+fd+9zrOkX/fR9tYYD1ukPxwB3/O/xqQffSfsN6Bjh9/2e5qj7/qNdaz+Ga/M4zFY+n/fFMZ728WV8vBDf4/6eHPj3I+rjKzj8NtwBP9fnQ+FYh6sHvvS8mvCztafx9Xua/nv9jiRvf5j33B+rtfOPI370expfv6fx9Xsa8SN9n+b3+h8wfvR7moEnff3I8wB+9HuaQ0V/v1X5A+fp39+eAJ+PJr5LTzNg398X9xyMVLSSynzgZ4KPO++5z+rYbtfw+WjxHts9iB/fVY1v19P4+j2Nr9/TvMd3+LrjfH+EX634Xj2Nr9/T+Po9zXv0/T7VjR/9nsbX72l8/Z7G1+9pfP2extfvaXz9nsbX72l8/Z7G1+9pfP2extfvaXz9nsbX72l8/Z7G1+9pfP2extfvaXz9nsbX72l8/Z7mfwHq188OqEy/RgAAACV0RVh0ZGF0ZTpjcmVhdGUAMjAyNS0wNS0yNlQwNzoyNzoxNyswMDowMDD8f1sAAAAldEVYdGRhdGU6bW9kaWZ5ADIwMjUtMDUtMjZUMDc6Mjc6MTcrMDA6MDBBocfnAAAAKHRFWHRkYXRlOnRpbWVzdGFtcAAyMDI1LTA1LTI2VDA3OjI3OjIxKzAwOjAw++vU4QAAAABJRU5ErkJggg==", # Companylogo als eingebettetes Base64
    [string]$ProductLogoBase64 = "data:image/png;base64,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", # Produktlogo als eingebettetes Base64
    [string]$ExcelFilePath = "",
    [string]$ExcelFileUrl  = ""
)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Web



# ======= Universelle URL-Encoder-Funktion =======
function UrlEncode([string]$str) {
    [System.Uri]::EscapeDataString($str)
}

# ======= MSP-Daten (Fallback auf Defaults, falls leer) =======
if (-not $MSPPartner) { $MSPPartner = "Some MSP Partner" }
if (-not $MSPNameAP)  { $MSPNameAP  = "Contactname" }
if (-not $MSPMail)    { $MSPMail    = "support@domainen.io" }
if (-not $MSPURL)     { $MSPURL     = "https://www.domainen.io" }
if (-not $MSPNameEU)  { $MSPNameEU  = "EU Contact Name" }

# ======= Zeitstempel / Logdateien =======
# $DatumJetzt = Get-Date -Format 'yyyyMMdd_HHmmss'
# $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
# $LogFolder = Join-Path $ScriptDir "Logs"
# if (-not (Test-Path $LogFolder)) { New-Item -Path $LogFolder -ItemType Directory -Force > $null }
# $DeepLLog = Join-Path $LogFolder "deepl_api_traffic_$DatumJetzt.log"

# ======= Verzeichnis-Handling für Logs und Exporte =======
# $DatumHeute = Get-Date -Format 'yyyyMMdd'
$DatumHeute = Get-Date -Format 'yyyyMMdd_HHmmss'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogFolder = Join-Path $ScriptDir "Logs_$DatumHeute"
$ExportFolder = Join-Path $ScriptDir "Export_$DatumHeute"

if (-not (Test-Path $LogFolder))    { New-Item -Path $LogFolder    -ItemType Directory -Force > $null }
if (-not (Test-Path $ExportFolder)) { New-Item -Path $ExportFolder -ItemType Directory -Force > $null }

$DatumJetzt = Get-Date -Format 'yyyyMMdd_HHmmss'
$DeepLLog   = Join-Path $LogFolder   "deepl_api_traffic_$DatumJetzt.log"

# ======= Logging + Error-Handling =======
function Write-Log-DeepLRequest {
    param(
        [string]$Stage,
        [string]$Lang,
        [string]$Text,
        [string]$Body,
        [string]$Response,
        [string]$Error
    )
    $logEntry = @()
    $logEntry += ("----- {0} {1} {2} -----" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Stage, $Lang)
    if ($Text)     { $logEntry += "Text:`n$Text" }
    if ($Body)     { $logEntry += "Body:`n$Body" }
    if ($Response) { $logEntry += "Response:`n$Response" }
    if ($Error)    { $logEntry += "Error:`n$Error" }
    $logEntry += ""
    Add-Content -Path $DeepLLog -Encoding UTF8 -Value ($logEntry -join "`n")
}
function To-Utf8String($text) {
    if ($null -eq $text) { return "" }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$text)
    return [System.Text.Encoding]::UTF8.GetString($bytes)
}

# ======= DeepL unterstützte Sprachen =======
$DeepLSupported = @("BG","CS","DA","DE","EL","EN-GB","EN-US","ES","ET","FI","FR","HU","ID","IT","JA","KO","LT","LV","NB","NL","PL","PT-PT","PT-BR","RO","RU","SK","SL","SV","TR","UK","ZH")

# ======= DeepL API Key (bitte setzen) =======
$script:DeepLApiKey = "b7941202-f333-4f6e-ac28-3929ef648a2f:fx"

# ======= Zielsprachen-Mapping =======
$targetLangs = @{
    "en-us" = "EN-US"
    "fr-fr" = "FR"
    "it-it" = "IT"
    "hr-hr" = "HR"
    "pl-pl" = "PL"
    "ro-ro" = "RO"
    "sk-sk" = "SK"
    "cs-cz" = "CS"
    "uk-ua" = "UK"
    "hu-hu" = "HU"
    "pt-br" = "PT-BR"    
}

# ======= DeepL Translation =======
function Translate-DeepL {
    param(
        [string]$Text,
        [string]$TargetLang,   # z.B. "HR", "FR", "EN-US" etc.
        [int]$MaxRetries = 4,
        [int]$SleepMs = 1000
    )
    if (-not $script:DeepLApiKey) { return $Text }
    $deeplLang = $TargetLang.ToUpper()
    if (-not ($DeepLSupported -contains $deeplLang)) {
        Log-DeepLRequest "FALLBACK" $deeplLang $Text $null "Fallback to EN-US" "not supported by DeepL"
        $deeplLang = "EN-US"
    }
    $tries = 0
    while ($tries -lt $MaxRetries) {
        try {
            $body = "text=$(UrlEncode $Text)&target_lang=$deeplLang&source_lang=DE"
            Log-DeepLRequest "REQUEST" $deeplLang $Text $body $null $null
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
            $url = "https://api-free.deepl.com/v2/translate"
            $request = [System.Net.WebRequest]::Create($url)
            $request.Method = "POST"
            $request.ContentType = "application/x-www-form-urlencoded"
            $request.Headers.Add("Authorization", "DeepL-Auth-Key $script:DeepLApiKey")
            $request.ContentLength = $bytes.Length
            $requestStream = $request.GetRequestStream()
            $requestStream.Write($bytes, 0, $bytes.Length)
            $requestStream.Close()
            $response = $request.GetResponse()
            $stream = $response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
            $result = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()
            $response.Close()
            Log-DeepLRequest "RESPONSE" $deeplLang $Text $body $result $null
            $json = $result | ConvertFrom-Json
            return $json.translations[0].text
        } catch {
            $errorMsg = $_.Exception.Message
            if ($_.Exception.Response) {
                try {
                    $respStream = $_.Exception.Response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($respStream, [System.Text.Encoding]::UTF8)
                    $errorBody = $reader.ReadToEnd()
                    $reader.Close()
                } catch { $errorBody = "n/a" }
                Log-DeepLRequest "ERROR" $deeplLang $Text $body $errorBody $errorMsg
            } else {
                Log-DeepLRequest "ERROR" $deeplLang $Text $body $null $errorMsg
            }
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode.value__ -eq 429) {
                $tries++
                Start-Sleep -Seconds 3
            } else {
                Write-Host "❌ Fehler bei der Übersetzung ($deeplLang): $errorMsg"
                return $Text
            }
        }
    }
    Write-Host "❌ Übersetzung nach $MaxRetries Versuchen fehlgeschlagen!"
    return $Text
}

$utf8NoBom = New-Object System.Text.UTF8Encoding($False)

# === Modulprüfung und Import ===
function Install-RequiredModule {
    param([string]$ModuleName)
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "📦 Modul '$ModuleName' nicht gefunden – versuche Installation..." -ForegroundColor Yellow
        if (-not (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue)) {
            Register-PSRepository -Default
        }
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber
            Write-Host "✅ Modul '$ModuleName' installiert." -ForegroundColor Green
        } catch {
            Write-Host "❌ Modulinstallation fehlgeschlagen: $ModuleName" -ForegroundColor Red
            exit
        }
    } else {
        Write-Host "✅ Modul '$ModuleName' ist bereits installiert." -ForegroundColor Green
    }
}
Install-RequiredModule -ModuleName "ImportExcel"
Import-Module ImportExcel -Force

# ======= Excel Datei bestimmen und einlesen (lokal/URL/Dialog) =======
function Select-ExcelFileDialog {
    Add-Type -AssemblyName System.Windows.Forms
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Excel-Dateien (*.xls;*.xlsx)|*.xls;*.xlsx|Alle Dateien (*.*)|*.*"
    $ofd.Title = "Bitte Buchungskreis-Excel auswählen"
    $null = $ofd.ShowDialog()
    return $ofd.FileName
}

if ($ExcelFilePath -and (Test-Path $ExcelFilePath)) {
    $excelFile = $ExcelFilePath
    Write-Host "Excel-Datei per Pfad: $excelFile" -ForegroundColor Green
}
elseif ($ExcelFileUrl -and $ExcelFileUrl -match "^https?://") {
    $tempFile = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.xlsx'
    Invoke-WebRequest -Uri $ExcelFileUrl -OutFile $tempFile
    $excelFile = $tempFile
    Write-Host "Excel-Datei per URL geladen: $excelFile" -ForegroundColor Green
}
else {
    $excelFile = Get-ChildItem .\ | Where-Object { $_.Name -like "buchungskreis*.xls*" } | Select-Object -First 1
    if ($excelFile) {
        $excelFile = $excelFile.FullName
        Write-Host "Excel-Datei automatisch gefunden: $excelFile" -ForegroundColor Green
    }
    else {
        [System.Windows.MessageBox]::Show("Keine passende Buchungskreis Excel Datei gefunden. Bitte wählen Sie die Datei aus.", "Datei auswählen", "OK", "Warning")
        $selectedPath = Select-ExcelFileDialog
        if ($selectedPath -and (Test-Path $selectedPath)) {
            $excelFile = $selectedPath
            Write-Host "Excel-Datei per Dialog gewählt: $excelFile" -ForegroundColor Green
        } else {
            [System.Windows.MessageBox]::Show("Es wurde keine Datei ausgewählt. Das Script wird beendet.")
            exit
        }
    }
}

try {
    $excelData = Import-Excel -Path $excelFile
    Write-Host "Excel-Datei erfolgreich eingelesen." -ForegroundColor Green
} catch {
    Write-Host "Excel-Datei konnte nicht eingelesen werden: $_" -ForegroundColor Red
    exit
}

# ======= Spaltennamen automatisch finden =======
$colBKreis   = ($excelData | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match "Buchungskreis" })[0].Name
$colCompany  = ($excelData | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match "Bezeichnung" })[0].Name
$colMailDom  = ($excelData | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match "E.?Mail.*Domain" -or $_.Name -match "Mail.*Domain" })[0].Name

# ======= Buchungskreis-ComboBox befüllen =======
$bkList = $excelData | Select-Object -ExpandProperty $colBKreis | ForEach-Object {
    $_.ToString().PadLeft(4,'0')
} | Sort-Object -Unique

# ======= XAML-Definition inkl. ProgressBar & Status + neue Buttons =======
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Label Provisioning" Height="900" Width="1100" WindowStartupLocation="CenterScreen">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
        </Grid.RowDefinitions>

        <!-- Header mit Logos und Titel -->
        <StackPanel Grid.Row="0" Orientation="Horizontal" VerticalAlignment="Top" Margin="10">
            <Image Name="imgCompanyLogo" Height="48" Width="48" Margin="0,0,10,0"/>
            <TextBlock Text="Microsoft Purview Automation" FontSize="20" FontWeight="Bold" VerticalAlignment="Center"/>
            <Image Name="imgProductLogo" Height="48" Width="48" HorizontalAlignment="Right" Margin="10,0,0,0"/>
        </StackPanel>

        <!-- Hauptbereich -->
        <StackPanel Grid.Row="1" Margin="10,10,10,10">
            <TextBlock Text="Buchungskreis auswählen:" FontWeight="Bold" Margin="0,0,0,5"/>
            <ComboBox Name="cbBkreis" Width="180" Height="28" Margin="0,0,0,10"/>
            <TextBlock Text="Vorschau:" FontWeight="Bold" Margin="0,0,0,5"/>
            <TextBox Name="tbPreview" Height="400" Width="1060" Margin="0,0,0,10" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" AcceptsReturn="True" FontFamily="Consolas" FontSize="13" IsReadOnly="True"/>
            <StackPanel Orientation="Horizontal" Margin="0,10,0,10">
                <ProgressBar Name="pbTranslate" Width="350" Height="20" Minimum="0" Maximum="100" Value="0" Margin="0,0,10,0"/>
                <TextBlock Name="tbStatus" VerticalAlignment="Center" FontWeight="Bold"/>
            </StackPanel>
            <Grid Margin="0,0,0,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Orientation="Horizontal" Grid.Column="0">
                    <Button Name="btnExportCSV" Width="140" Height="32" Margin="0,0,10,0" Content="Exportiere als CSV"/>
                    <Button Name="btnExportJSON" Width="140" Height="32" Content="Exportiere als JSON" Margin="0,0,10,0"/>
                    <Button Name="btnTranslate" Width="170" Height="32" Content="Übersetzen (DeepL)" Margin="0,0,10,0"/>
                </StackPanel>
                <Button Name="btnCancel" Grid.Column="2" Width="140" Height="32" Content="Abbrechen" HorizontalAlignment="Right" />
            </Grid>
        </StackPanel>

        <!-- Neue Zeile für weitere Buttons -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Left" Margin="10,0,10,10">
            <Button Name="btnCreateLabels" Width="220" Height="36" Margin="0,0,15,0" Content="Neue Labels erstellen"/>
            <Button Name="btnMoreOptions" Width="220" Height="36" Content="weitere Optionen"/>
        </StackPanel>

        <!-- Footer -->
        <StackPanel Grid.Row="3" Orientation="Vertical" VerticalAlignment="Bottom" Margin="10,0,10,10">
            <TextBlock Text="Bei Fragen wenden Sie sich an:" FontSize="10" Margin="0,10,0,0"/>
            <TextBlock Name="tbFooterContact" FontSize="10"/>
        </StackPanel>
    </Grid>
</Window>
"@

# ======= XAML laden und Controls referenzieren =======
$reader = (New-Object System.Xml.XmlNodeReader ([xml]$xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)

$cbBkreis   = $window.FindName('cbBkreis')
$tbPreview  = $window.FindName('tbPreview')
$btnExportCSV = $window.FindName('btnExportCSV')
$btnExportJSON = $window.FindName('btnExportJSON')
$btnCancel = $window.FindName('btnCancel')
$btnTranslate = $window.FindName('btnTranslate')
$imgCompanyLogo = $window.FindName('imgCompanyLogo')
$imgProductLogo = $window.FindName('imgProductLogo')
$tbFooterContact = $window.FindName('tbFooterContact')
$pbTranslate = $window.FindName('pbTranslate')
$tbStatus    = $window.FindName('tbStatus')
$btnCreateLabels  = $window.FindName('btnCreateLabels')
$btnMoreOptions   = $window.FindName('btnMoreOptions')

if (-not $cbBkreis -or -not $tbPreview -or -not $btnExportCSV -or -not $btnExportJSON -or -not $btnTranslate -or -not $btnCancel -or -not $pbTranslate -or -not $tbStatus -or -not $btnCreateLabels -or -not $btnMoreOptions) {
    [System.Windows.MessageBox]::Show("Fehler beim Laden der Oberfläche! (Einige Controls konnten nicht gefunden werden.)")
    return
}

$cbBkreis.ItemsSource = $bkList
$tbFooterContact.Text = "$MSPPartner, $MSPNameAP"

function Set-Image ($ctrl, $base64) {
    try {
        if ($ctrl -and $base64 -and $base64.Length -gt 100) {
            $base64 = $base64 -replace '^data:image\/[a-z]+;base64,', ''
            $bytes = [Convert]::FromBase64String($base64)
            $stream = New-Object IO.MemoryStream (,[byte[]]$bytes)
            $img = [System.Windows.Media.Imaging.BitmapImage]::new()
            $img.BeginInit()
            $img.StreamSource = $stream
            $img.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $img.EndInit()
            $img.Freeze()
            $ctrl.Source = $img
        }
    } catch {
        Write-Host "❌ Logo-Fehler: $_"
    }
}
Set-Image $imgCompanyLogo $CompanyLogoBase64
Set-Image $imgProductLogo $ProductLogoBase64

$labelTemplates = @(
    "Private Informationen",
    "Geschäftliche Informationen",
    "Nur für internen Dienstgebrauch",
    "Nur für internen Dienstgebrauch (unverschlüsselt)",
    "Nur für internen Dienstgebrauch (verschlüsselt)",
    "Vertraulich",
    "Vertraulich (unverschlüsselt)",
    "Vertraulich (verschlüsselt)",
    "Streng Vertraulich"
)

# ======= Vorschau- und Export-Logik =======
function Build-LabelObjects ($bk, $company, $maildomain) {
    $bk4 = $bk.ToString().PadLeft(4, '0')
    $companyClean = $company.TrimEnd()
    $maildomainClean = $maildomain.Trim()
    $out = @()
    foreach ($label in $labelTemplates) {
        $obj = [PSCustomObject]@{
            Name        = "$label - $bk4"
            DisplayName = "$label - $bk4"
            Tooltip     = "$label - $companyClean"
            EncryptionProtectionType = $(if ($label -eq "Streng Vertraulich") {"UserDefined"} else {"Template"})
            EncryptionOfflineAccessDays = -1
            LabelContentType = "File, Email,SchematizedData"
        }
        if ($label -eq "Nur für internen Dienstgebrauch (verschlüsselt)") {
            $obj | Add-Member NoteProperty MIPIdentity1 "AuthenticatedUsers"
            $obj | Add-Member NoteProperty MIPRights1 "VIEW,VIEWRIGHTSDATA,DOCEDIT,EDIT,PRINT,REPLY,REPLYALL"
            $obj | Add-Member NoteProperty MIPIdentity2 "huelsenbergholding.onmicrosoft.com"
            $obj | Add-Member NoteProperty MIPRights2 "VIEW,VIEWRIGHTSDATA,DOCEDIT,EDIT,PRINT,REPLY,REPLYALL,EXTRACT,FORWARD,EXPORT,OBJMODEL"
            $obj | Add-Member NoteProperty MIPIdentity3 $maildomainClean
            $obj | Add-Member NoteProperty MIPRights3 "VIEW,VIEWRIGHTSDATA,DOCEDIT,EDIT,PRINT,REPLY,REPLYALL,EXTRACT,FORWARD,EDITRIGHTSDATA,EXPORT,OBJMODEL"
        }
        elseif ($label -eq "Vertraulich (verschlüsselt)" -or $label -eq "Streng Vertraulich") {
            $obj | Add-Member NoteProperty MIPIdentity1 "huelsenbergholding.de"
            $obj | Add-Member NoteProperty MIPRights1 "VIEW,VIEWRIGHTSDATA,DOCEDIT,EDIT,PRINT,REPLY,REPLYALL,EXTRACT,FORWARD,EXPORT,OBJMODEL"
            $obj | Add-Member NoteProperty MIPIdentity2 "huelsenbergholding.onmicrosoft.com"
            $obj | Add-Member NoteProperty MIPRights2 "VIEW,VIEWRIGHTSDATA,DOCEDIT,EDIT,PRINT,REPLY,REPLYALL,EXTRACT,FORWARD,EXPORT,OBJMODEL"
            $obj | Add-Member NoteProperty MIPIdentity3 $maildomainClean
            $obj | Add-Member NoteProperty MIPRights3 "VIEW,VIEWRIGHTSDATA,DOCEDIT,EDIT,PRINT,REPLY,REPLYALL,EXTRACT,FORWARD,EDITRIGHTSDATA,EXPORT,OBJMODEL"
        }
        if ($label -like "*verschlüsselt*" -or $label -eq "Streng Vertraulich") { $obj | Add-Member NoteProperty EncryptionEnabled 1 }
        if ($label -like "Vertraulich*" -or $label -eq "Streng Vertraulich") {
            $obj | Add-Member NoteProperty EncryptionEncryptOnly 0
            $obj | Add-Member NoteProperty EncryptionDoNotForward 1
        }
        if ($label -like "Nur für internen Dienstgebrauch*") {
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterAlignment ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterEnabled ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterFontColor ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterFontName ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterFontSize ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterMargin ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterPlacement ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterText ""
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderAlignment "Center"
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderEnabled 1
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderFontColor "#000000"
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderFontName ""
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderFontSize 8
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderMargin 5
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderPlacement "Header"
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderText "Nur für internen Dienstgebrauch - $companyClean - internal use only"
            $obj | Add-Member NoteProperty ApplyWaterMarkingEnabled ""
            $obj | Add-Member NoteProperty ApplyWaterMarkingFontColor ""
            $obj | Add-Member NoteProperty ApplyWaterMarkingFontName ""
            $obj | Add-Member NoteProperty ApplyWaterMarkingFontSize ""
            $obj | Add-Member NoteProperty ApplyWaterMarkingLayout ""
            $obj | Add-Member NoteProperty ApplyWaterMarkingText ""
        }
        elseif ($label -like "Vertraulich*") {
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterAlignment ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterEnabled ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterFontColor ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterFontName ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterFontSize ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterMargin ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterPlacement ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterText ""
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderAlignment "Center"
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderEnabled 1
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderFontColor "#000000"
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderFontName ""
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderFontSize 8
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderMargin 5
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderPlacement "Header"
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderText "Vertraulich - $companyClean - confidential"
            $obj | Add-Member NoteProperty ApplyWaterMarkingEnabled 1
            $obj | Add-Member NoteProperty ApplyWaterMarkingFontColor "#a2a2a2"
            $obj | Add-Member NoteProperty ApplyWaterMarkingFontName ""
            $obj | Add-Member NoteProperty ApplyWaterMarkingFontSize 28
            $obj | Add-Member NoteProperty ApplyWaterMarkingLayout "Diagonal"
            $obj | Add-Member NoteProperty ApplyWaterMarkingText "vertraulich / confidential"
            $obj | Add-Member NoteProperty EncryptionContentExpiredOnDateInDaysOrNever "Never"
        }
        elseif ($label -eq "Streng Vertraulich") {
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterAlignment ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterEnabled ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterFontColor ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterFontName ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterFontSize ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterMargin ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterPlacement ""
            $obj | Add-Member NoteProperty ApplyContentMarkingFooterText ""
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderAlignment "Center"
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderEnabled 1
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderFontColor "#000000"
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderFontName ""
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderFontSize 8
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderMargin 5
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderPlacement "Header"
            $obj | Add-Member NoteProperty ApplyContentMarkingHeaderText "Streng Vertraulich - $companyClean - strictly confidential"
            $obj | Add-Member NoteProperty ApplyWaterMarkingEnabled 1
            $obj | Add-Member NoteProperty ApplyWaterMarkingFontColor "#a2a2a2"
            $obj | Add-Member NoteProperty ApplyWaterMarkingFontName ""
            $obj | Add-Member NoteProperty ApplyWaterMarkingFontSize 28
            $obj | Add-Member NoteProperty ApplyWaterMarkingLayout "Diagonal"
            $obj | Add-Member NoteProperty ApplyWaterMarkingText "Streng vertraulich / strictly confidential"
            $obj | Add-Member NoteProperty EncryptionContentExpiredOnDateInDaysOrNever "Never"
            $obj | Add-Member NoteProperty EncryptionPromptUser 1
        }
        $out += $obj
    }
    return $out
}
function Build-PreviewText ($labels) {
    $lines = @()
    foreach ($item in $labels) {
        foreach ($prop in $item.PSObject.Properties) {
            $lines += "{0,-38}: {1}" -f $prop.Name, $prop.Value
        }
        $lines += ""
    }
    return ($lines -join "`r`n")
}

$updatePreview = {
    $selBK = $cbBkreis.SelectedItem
    if ($selBK) {
        $entry = $excelData | Where-Object { $_.$colBKreis.ToString().PadLeft(4,'0') -eq $selBK }
        if ($entry) {
            $company = $entry.$colCompany
            $maildomain = $entry.$colMailDom
            $labels = Build-LabelObjects $selBK $company $maildomain
            $tbPreview.Text = Build-PreviewText $labels
            $window.Tag = $labels
        }
    } else {
        $tbPreview.Text = ""
        $window.Tag = $null
    }
}
$null = $cbBkreis.Add_SelectionChanged($updatePreview)

# Exportiere wirklich alle Felder ins CSV (wie JSON)
$btnExportCSV.Add_Click({
    $labels = $window.Tag
    $selBK = $cbBkreis.SelectedItem
    if ($labels -and $selBK) {
        $bk4 = $selBK.ToString().PadLeft(4,'0')
        $filename = Join-Path $ExportFolder "provision_labels_${bk4}.csv"
        # Alle Properties extrahieren (vereinigt über alle Objekte)
        $allProps = @{}
        foreach ($x in $labels) {
            foreach ($p in $x.PSObject.Properties.Name) { $allProps[$p] = $true }
        }
        # Sortierung: Name, DisplayName, Tooltip, dann alle weiteren alphabetisch
        $mainCols = @("Name", "DisplayName", "Tooltip")
        $otherCols = $allProps.Keys | Where-Object { $mainCols -notcontains $_ } | Sort-Object
        $orderedProps = $mainCols + $otherCols

        $labels | Select-Object $orderedProps | Export-Csv -Encoding UTF8 -NoTypeInformation $filename
        [System.Windows.MessageBox]::Show("CSV exportiert: $filename")
    } else {
        [System.Windows.MessageBox]::Show("Bitte Buchungskreis auswählen!")
    }
})

$btnExportJSON.Add_Click({
    $labels = $window.Tag
    $selBK = $cbBkreis.SelectedItem
    if ($labels -and $selBK) {
        $bk4 = $selBK.ToString().PadLeft(4,'0')
        $filename = Join-Path $ExportFolder "provision_labels_${bk4}.json"
        $jsonString = $labels | ConvertTo-Json -Depth 4
        [System.IO.File]::WriteAllText($filename, $jsonString, $utf8NoBom)
        [System.Windows.MessageBox]::Show("JSON exportiert: $filename")
    } else {
        [System.Windows.MessageBox]::Show("Bitte Buchungskreis auswählen!")
    }
})

$btnTranslate.Add_Click({
    $btnTranslate.IsEnabled = $false
    $btnTranslate.Content = "Übersetze..."
    $pbTranslate.Value = 0
    $tbStatus.Text = "Übersetzung läuft..."
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $labels = $window.Tag
        $selBK = $cbBkreis.SelectedItem
        if (-not $labels -or -not $selBK) {
            [System.Windows.MessageBox]::Show("Bitte Buchungskreis auswählen!")
            return
        }
        if (-not $script:DeepLApiKey -or $script:DeepLApiKey.Length -lt 10) {
            [System.Windows.MessageBox]::Show("Bitte trage deinen DeepL API-Key im Skript ein!")
            return
        }
        $bk4 = $selBK.ToString().PadLeft(4,'0')
        $filename = Join-Path $ExportFolder "provision_labels_translated_${bk4}.json"
        $translated = @()
        $total = $labels.Count * ($targetLangs.Count * 2)
        $done = 0
        foreach ($label in $labels) {
            $item = [ordered]@{ }
            $label.PSObject.Properties | ForEach-Object { $item[$_.Name] = $_.Value }
            foreach ($lang in $targetLangs.Keys) {
                $deeplLang = $targetLangs[$lang]
                $item["DisplayName_$lang"] = To-Utf8String(Translate-DeepL $label.DisplayName $deeplLang)
                $done++; $pbTranslate.Value = [int](($done/$total)*100)
                $tbStatus.Text = "Übersetze: $label.DisplayName → $lang"
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 500
                $item["Tooltip_$lang"] = To-Utf8String(Translate-DeepL $label.Tooltip $deeplLang)
                $done++; $pbTranslate.Value = [int](($done/$total)*100)
                $tbStatus.Text = "Übersetze: $label.Tooltip → $lang"
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 500
            }
            $translated += [PSCustomObject]$item
        }
        $pbTranslate.Value = 100
        $tbStatus.Text = "Fertig!"
        $jsonString = $translated | ConvertTo-Json -Depth 4
        [System.IO.File]::WriteAllText($filename, $jsonString, $utf8NoBom)
        [System.Windows.MessageBox]::Show("Übersetzung abgeschlossen und als '${filename}' gespeichert.")
        Start-Process explorer.exe "/select,`"$filename`""
    } catch {
        $tbStatus.Text = "FEHLER: $_"
        [System.Windows.MessageBox]::Show("Fehler: $_")
    } finally {
        $btnTranslate.IsEnabled = $true
        $btnTranslate.Content = "Übersetzen (DeepL)"
        Start-Sleep 1
        $pbTranslate.Value = 0
        $tbStatus.Text = ""
    }
})

$btnCancel.Add_Click({ $window.Close() })

# Neue Buttons: "Neue Labels erstellen" und "weitere Optionen"
$btnCreateLabels.Add_Click({
    $scriptName = "02-Run-PurviewLabelProvisioning_Create_Missing_Config_Only_Language_Final_V1.ps1"
    Write-Host "Starte: $scriptName" -ForegroundColor Cyan
    Add-Content -Path $DeepLLog -Value ("[{0}] [INFO] Button 'Neue Labels erstellen' gedrückt, Script: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $scriptName)
    try {
        Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptDir\$scriptName`"" -WindowStyle Normal
        $window.Close()
    } catch {
        [System.Windows.MessageBox]::Show("Fehler beim Starten von $scriptName`n$_")
    }
})

$btnMoreOptions.Add_Click({
    $scriptName = "00_q_Start-PurviewGUI_V10_280525d_chat.ps1"
    Write-Host "Starte: $scriptName" -ForegroundColor Cyan
    Add-Content -Path $DeepLLog -Value ("[{0}] [INFO] Button 'weitere Optionen' gedrückt, Script: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $scriptName)
    try {
        Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptDir\$scriptName`"" -WindowStyle Normal
        $window.Close()
    } catch {
        [System.Windows.MessageBox]::Show("Fehler beim Starten von $scriptName`n$_")
    }
})

# Ctrl+C-Handler für sauberes GUI-Schließen
try {
    if ($Host.Name -eq "ConsoleHost" -and [Console]::CancelKeyPress) {
        [Console]::CancelKeyPress.Add({
            try {
                if ($window -and $window.Dispatcher.CheckAccess()) { $window.Close() }
            } catch {}
            Write-Host "`nAbbruch durch Benutzer (Strg+C). Fenster wird geschlossen."
            $_.Cancel = $true
        }) > $null
    }
} catch {
    # kein Problem, wenn das Event nicht existiert (z.B. im ISE oder WPF-Host)
}

try {
    $window.ShowDialog() > $null
} catch {
    Write-Host "Kritischer Fehler im GUI: $_"
    try { if ($window -and $window.Dispatcher.CheckAccess()) { $window.Close() } } catch {}
} finally {
    # Excel aufräumen (ImportExcel benötigt kein explizites Aufräumen)
}
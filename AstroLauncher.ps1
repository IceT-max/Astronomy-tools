#Requires -Version 5.1
# ============================================================
#  AstroLauncher.ps1
#  Launcher per gli script astronomici - Jean Meeus Suite
#  PowerShell 5.1 - Windows Forms - ASCII standard
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# --- Cartella dello script ---
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
if ([string]::IsNullOrEmpty($ScriptDir)) { $ScriptDir = $PWD.Path }

# ============================================================
#  DEFINIZIONE DEGLI SCRIPT  (6 voci)
# ============================================================
$Scripts = @(
    [PSCustomObject]@{
        Label       = "  [1]  AstroCalc"
        File        = "AstroCalc.ps1"
        Title       = "Calcolatore Astronomico"
        Description = "Posizione Sole, Luna e pianeti." + "`r`n" +
                      "Sorgere/tramonto, fase lunare," + "`r`n" +
                      "distanze e magnitudini."
        Color       = [System.Drawing.Color]::FromArgb(255, 200,  80,  20)
    },
    [PSCustomObject]@{
        Label       = "  [2]  Fasi Lunari"
        File        = "FasiLunari.ps1"
        Title       = "Fasi Lunari"
        Description = "Calcolo fasi lunari (Meeus cap.32)." + "`r`n" +
                      "Luna nuova, primo quarto," + "`r`n" +
                      "luna piena, ultimo quarto."
        Color       = [System.Drawing.Color]::FromArgb(255, 120, 120, 200)
    },
    [PSCustomObject]@{
        Label       = "  [3]  Apogeo & Perigeo"
        File        = "LunaApogeoPerigeo.ps1"
        Title       = "Apogeo e Perigeo Lunare"
        Description = "Calcolo apogeo/perigeo lunare" + "`r`n" +
                      "(Meeus cap.50). Data, ora e" + "`r`n" +
                      "distanza in km."
        Color       = [System.Drawing.Color]::FromArgb(255,  60, 160, 120)
    },
    [PSCustomObject]@{
        Label       = "  [4]  Conversione Coordinate"
        File        = "ConversioneCoordinate.ps1"
        Title       = "Conversione Coordinate"
        Description = "Conversione equatoriali <-> alt-az." + "`r`n" +
                      "Tempo siderale (Meeus cap.7-8)." + "`r`n" +
                      "Input: luogo, data, ora."
        Color       = [System.Drawing.Color]::FromArgb(255,  40, 140, 200)
    },
    [PSCustomObject]@{
        Label       = "  [5]  Congiunzioni (Horizons)"
        File        = "Horizons-Congiunzioni.ps1"
        Title       = "Congiunzioni Planetarie JPL"
        Description = "Ricerca congiunzioni via JPL" + "`r`n" +
                      "Horizons API (richiede Internet)." + "`r`n" +
                      "Pianeti visibili ad occhio nudo."
        Color       = [System.Drawing.Color]::FromArgb(255, 180,  60, 180)
    },
    [PSCustomObject]@{
        Label       = "  [6]  Eclissi Solari & Lunari"
        File        = "EclissiCalc.ps1"
        Title       = "Eclissi Solari e Lunari"
        Description = "Eclissi solari: elementi Besseliani" + "`r`n" +
                      "da CSV (Espenak & Meeus, 5MKLE)." + "`r`n" +
                      "Eclissi lunari: Meeus AFC cap.33."
        Color       = [System.Drawing.Color]::FromArgb(255, 220, 200,  30)
    }
)

# ============================================================
#  COLORI
# ============================================================
$BG       = [System.Drawing.Color]::FromArgb(255,  18,  18,  30)
$BG2      = [System.Drawing.Color]::FromArgb(255,  28,  28,  45)
$BG3      = [System.Drawing.Color]::FromArgb(255,  12,  12,  22)
$Gold     = [System.Drawing.Color]::FromArgb(255, 220, 180,  60)
$Silver   = [System.Drawing.Color]::FromArgb(255, 180, 180, 200)
$DimText  = [System.Drawing.Color]::FromArgb(255, 100, 100, 130)
$SepColor = [System.Drawing.Color]::FromArgb(255,  60,  60,  90)
$BtnBG    = [System.Drawing.Color]::FromArgb(255,  38,  38,  58)
$BtnHov   = [System.Drawing.Color]::FromArgb(255,  50,  50,  75)
$ExitBG   = [System.Drawing.Color]::FromArgb(255,  50,  20,  20)
$ExitBdr  = [System.Drawing.Color]::FromArgb(255, 160,  40,  40)
$MoonClr  = [System.Drawing.Color]::FromArgb(255, 200, 200, 160)
$StatusOK = [System.Drawing.Color]::FromArgb(255,  60, 200, 100)
$StatusER = [System.Drawing.Color]::FromArgb(255, 220,  60,  60)

# ============================================================
#  FONT
# ============================================================
$FontMonoB = New-Object System.Drawing.Font("Courier New",  9, [System.Drawing.FontStyle]::Bold)
$FontSub   = New-Object System.Drawing.Font("Courier New",  8, [System.Drawing.FontStyle]::Regular)
$FontBtn   = New-Object System.Drawing.Font("Courier New",  9, [System.Drawing.FontStyle]::Bold)
$FontDesc  = New-Object System.Drawing.Font("Courier New",  8, [System.Drawing.FontStyle]::Regular)
$FontDTit  = New-Object System.Drawing.Font("Courier New", 10, [System.Drawing.FontStyle]::Bold)
$FontBanr  = New-Object System.Drawing.Font("Courier New",  8, [System.Drawing.FontStyle]::Bold)
$FontMoon  = New-Object System.Drawing.Font("Courier New",  8, [System.Drawing.FontStyle]::Regular)

# ============================================================
#  METRICHE - TUTTE PRE-CALCOLATE (regola PS 5.1)
#
#  6 bottoni: H = 44, Gap = 10 -> passo = 54
#  BtnTopStart = 130
#  Fine bottoni = 130 + 6*54 = 454  -> AfterBtns = 464
#  Pannello:     PanTop=104, PanH=340 -> AfterPan  = 454
#  SepY = max(464, 454) = 464
#  FormH = SepY + 80 = 544  -> arrotondato a 560
# ============================================================
$FormW       = 640
$FormH       = 560
$Margin      = 18
$BtnW        = 244
$BtnH        = 42
$BtnGap      = 8
$BtnLeft     = 18           # = $Margin
$PanW        = 336
$PanH        = 340
$PanLeft     = 274          # $Margin + $BtnW + 12 = 18+244+12
$PanTop      = 104
$BannerX     = 18
$BannerY     = 14
$BannerW     = 604          # $FormW - $Margin*2
$BannerH     = 72
$Sep1Y       = 92
$Sep1W       = 604
$Col1LblY    = 100
$BtnTopStart = 122
$PanInnerW   = 316          # $PanW - 20

# Calcolo verticale (6 bottoni):
# Fine = 122 + 6*(42+8) = 122 + 300 = 422 -> AfterBtns = 432
# AfterPan = 104 + 340 + 10 = 454   <- dominante
$SepY        = 454
$StatusY     = 462          # $SepY + 8
$FooterY     = 484          # $SepY + 30
$ExitX       = 502          # $FormW - 120 - $Margin

# ============================================================
#  FORM
# ============================================================
$Form = New-Object System.Windows.Forms.Form
$Form.Text            = "AstroLauncher  -  Jean Meeus Suite"
$Form.Size            = New-Object System.Drawing.Size($FormW, $FormH)
$Form.StartPosition   = "CenterScreen"
$Form.BackColor       = $BG
$Form.ForeColor       = $Silver
$Form.FormBorderStyle = "FixedSingle"
$Form.MaximizeBox     = $false
$Form.Font            = $FontSub

# ============================================================
#  BANNER
# ============================================================
$BannerText = "+------------------------------------------------------------+" + "`r`n" +
              "|  *    .       ASTRO LAUNCHER       .         *    .        |" + "`r`n" +
              "|    .    Jean Meeus Astronomical Suite  v1.0      .    *    |" + "`r`n" +
              "+------------------------------------------------------------+"

$lblBanner = New-Object System.Windows.Forms.Label
$lblBanner.Text      = $BannerText
$lblBanner.Font      = $FontBanr
$lblBanner.ForeColor = $Gold
$lblBanner.BackColor = [System.Drawing.Color]::Transparent
$lblBanner.Location  = New-Object System.Drawing.Point($BannerX, $BannerY)
$lblBanner.Size      = New-Object System.Drawing.Size($BannerW, $BannerH)
$lblBanner.TextAlign = "MiddleCenter"
$Form.Controls.Add($lblBanner)

# ============================================================
#  SEPARATORE SUPERIORE
# ============================================================
$sep1 = New-Object System.Windows.Forms.Label
$sep1.Text        = ""
$sep1.BorderStyle = "Fixed3D"
$sep1.Location    = New-Object System.Drawing.Point($Margin, $Sep1Y)
$sep1.Size        = New-Object System.Drawing.Size($Sep1W, 2)
$sep1.BackColor   = $SepColor
$Form.Controls.Add($sep1)

# ============================================================
#  LABEL COLONNA SINISTRA
# ============================================================
$lblCol1 = New-Object System.Windows.Forms.Label
$lblCol1.Text      = "[ STRUMENTI ]"
$lblCol1.Font      = $FontMonoB
$lblCol1.ForeColor = $Silver
$lblCol1.Location  = New-Object System.Drawing.Point($BtnLeft, $Col1LblY)
$lblCol1.Size      = New-Object System.Drawing.Size($BtnW, 18)
$Form.Controls.Add($lblCol1)

# ============================================================
#  PANNELLO DESCRIZIONE (colonna destra)
# ============================================================
$PanDesc = New-Object System.Windows.Forms.Panel
$PanDesc.Location    = New-Object System.Drawing.Point($PanLeft, $PanTop)
$PanDesc.Size        = New-Object System.Drawing.Size($PanW, $PanH)
$PanDesc.BackColor   = $BG2
$PanDesc.BorderStyle = "FixedSingle"
$Form.Controls.Add($PanDesc)

$lblDescHdr = New-Object System.Windows.Forms.Label
$lblDescHdr.Text      = "[ INFO STRUMENTO ]"
$lblDescHdr.Font      = $FontMonoB
$lblDescHdr.ForeColor = $Gold
$lblDescHdr.Location  = New-Object System.Drawing.Point(10, 10)
$lblDescHdr.Size      = New-Object System.Drawing.Size($PanInnerW, 20)
$PanDesc.Controls.Add($lblDescHdr)

$lblDescSep2 = New-Object System.Windows.Forms.Label
$lblDescSep2.Text      = "+-----------------------------+"
$lblDescSep2.Font      = $FontSub
$lblDescSep2.ForeColor = $DimText
$lblDescSep2.Location  = New-Object System.Drawing.Point(10, 32)
$lblDescSep2.Size      = New-Object System.Drawing.Size($PanInnerW, 16)
$PanDesc.Controls.Add($lblDescSep2)

$lblDescTitle = New-Object System.Windows.Forms.Label
$lblDescTitle.Text      = "Seleziona uno strumento"
$lblDescTitle.Font      = $FontDTit
$lblDescTitle.ForeColor = [System.Drawing.Color]::White
$lblDescTitle.Location  = New-Object System.Drawing.Point(10, 52)
$lblDescTitle.Size      = New-Object System.Drawing.Size($PanInnerW, 22)
$PanDesc.Controls.Add($lblDescTitle)

$lblDescBody = New-Object System.Windows.Forms.Label
$lblDescBody.Text      = "Passa il mouse su un bottone" + "`r`n" +
                         "per vedere la descrizione" + "`r`n" +
                         "dello strumento."
$lblDescBody.Font      = $FontDesc
$lblDescBody.ForeColor = $Silver
$lblDescBody.Location  = New-Object System.Drawing.Point(10, 82)
$lblDescBody.Size      = New-Object System.Drawing.Size($PanInnerW, 80)
$PanDesc.Controls.Add($lblDescBody)

$lblDescFile = New-Object System.Windows.Forms.Label
$lblDescFile.Text      = ""
$lblDescFile.Font      = $FontSub
$lblDescFile.ForeColor = $DimText
$lblDescFile.Location  = New-Object System.Drawing.Point(10, 166)
$lblDescFile.Size      = New-Object System.Drawing.Size($PanInnerW, 16)
$PanDesc.Controls.Add($lblDescFile)

# --- ASCII art luna/eclisse nel pannello ---
$MoonText = "        _..._        " + "`r`n" +
            "      .'  .-'`''.    " + "`r`n" +
            "     /  /      \    " + "`r`n" +
            "    |   |  ) (  |   " + "`r`n" +
            "     \  '-.   .'    " + "`r`n" +
            "      '._   ``'      " + "`r`n" +
            "         ''--'      " + "`r`n" +
            "                    " + "`r`n" +
            "   *   .  Meeus .   " + "`r`n" +
            " .   Astronomical . " + "`r`n" +
            "   .  Algorithms  * "

$lblMoon = New-Object System.Windows.Forms.Label
$lblMoon.Text      = $MoonText
$lblMoon.Font      = $FontMoon
$lblMoon.ForeColor = $MoonClr
$lblMoon.Location  = New-Object System.Drawing.Point(10, 192)
$lblMoon.Size      = New-Object System.Drawing.Size($PanInnerW, 140)
$PanDesc.Controls.Add($lblMoon)

# ============================================================
#  BOTTONI (loop su 6 script)
# ============================================================
$BtnTop = $BtnTopStart

foreach ($s in $Scripts) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text      = $s.Label
    $btn.Font      = $FontBtn
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.BackColor = $BtnBG
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderColor        = $s.Color
    $btn.FlatAppearance.BorderSize         = 2
    $btn.FlatAppearance.MouseOverBackColor = $BtnHov
    $btn.Location  = New-Object System.Drawing.Point($BtnLeft, $BtnTop)
    $btn.Size      = New-Object System.Drawing.Size($BtnW, $BtnH)
    $btn.TextAlign = "MiddleLeft"
    $btn.Cursor    = [System.Windows.Forms.Cursors]::Hand

    $capturedScript = $s

    $btn.Add_MouseEnter({
        $lblDescTitle.Text      = $capturedScript.Title
        $lblDescTitle.ForeColor = $capturedScript.Color
        $lblDescBody.Text       = $capturedScript.Description
        $lblDescFile.Text       = "File: " + $capturedScript.File
    }.GetNewClosure())

    $btn.Add_MouseLeave({
        $lblDescTitle.Text      = "Seleziona uno strumento"
        $lblDescTitle.ForeColor = [System.Drawing.Color]::White
        $lblDescBody.Text       = "Passa il mouse su un bottone" + "`r`n" +
                                  "per vedere la descrizione" + "`r`n" +
                                  "dello strumento."
        $lblDescFile.Text       = ""
    }.GetNewClosure())

    $btn.Add_Click({
        $path = Join-Path $ScriptDir $capturedScript.File
        if (Test-Path $path) {
            $lblStatus.Text      = "  >> Avvio: " + $capturedScript.File + " ..."
            $lblStatus.ForeColor = $StatusOK
            # Runspace STA dedicato: nessun nuovo processo, nessuna console.
            # WinForms richiede STA (Single Thread Apartment).
            $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
            $rs  = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
            $rs.ApartmentState = [System.Threading.ApartmentState]::STA
            $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::UseNewThread
            $rs.Open()
            $ps = [System.Management.Automation.PowerShell]::Create()
            $ps.Runspace = $rs
            [void]$ps.AddScript(". '" + $path + "'")
            $ps.BeginInvoke() | Out-Null
        } else {
            $lblStatus.Text      = "  >> ERRORE: file non trovato -> " + $capturedScript.File
            $lblStatus.ForeColor = $StatusER
        }
    }.GetNewClosure())

    $Form.Controls.Add($btn)
    $BtnTop = $BtnTop + $BtnH + $BtnGap
}

# ============================================================
#  SEPARATORE INFERIORE
# ============================================================
$sep2 = New-Object System.Windows.Forms.Label
$sep2.Text        = ""
$sep2.BorderStyle = "Fixed3D"
$sep2.Location    = New-Object System.Drawing.Point($Margin, $SepY)
$sep2.Size        = New-Object System.Drawing.Size($BannerW, 2)
$sep2.BackColor   = $SepColor
$Form.Controls.Add($sep2)

# ============================================================
#  STATUS BAR
# ============================================================
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text      = "  >> Pronto.  Seleziona uno strumento e clicca per avviarlo."
$lblStatus.Font      = $FontSub
$lblStatus.ForeColor = $DimText
$lblStatus.BackColor = $BG3
$lblStatus.Location  = New-Object System.Drawing.Point(0, $StatusY)
$lblStatus.Size      = New-Object System.Drawing.Size($FormW, 20)
$lblStatus.TextAlign = "MiddleLeft"
$Form.Controls.Add($lblStatus)

# ============================================================
#  FOOTER
# ============================================================
$lblFooter = New-Object System.Windows.Forms.Label
$lblFooter.Text      = "Jean Meeus - Astronomical Algorithms  |  PowerShell 5.1  |  2025"
$lblFooter.Font      = $FontSub
$lblFooter.ForeColor = $DimText
$lblFooter.Location  = New-Object System.Drawing.Point(0, $FooterY)
$lblFooter.Size      = New-Object System.Drawing.Size($FormW, 16)
$lblFooter.TextAlign = "MiddleCenter"
$Form.Controls.Add($lblFooter)

# ============================================================
#  BOTTONE ESCI
# ============================================================
$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Text      = "[ ESCI ]"
$btnExit.Font      = $FontMonoB
$btnExit.ForeColor = $Silver
$btnExit.BackColor = $ExitBG
$btnExit.FlatStyle = "Flat"
$btnExit.FlatAppearance.BorderColor = $ExitBdr
$btnExit.FlatAppearance.BorderSize  = 1
$btnExit.Location  = New-Object System.Drawing.Point($ExitX, $StatusY)
$btnExit.Size      = New-Object System.Drawing.Size(120, 20)
$btnExit.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnExit.Add_Click({ $Form.Close() })
$Form.Controls.Add($btnExit)

# ============================================================
#  AVVIO
# ============================================================
[System.Windows.Forms.Application]::Run($Form)

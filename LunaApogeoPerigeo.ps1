# ============================================================
#  Apogeo e Perigeo Lunare - PowerShell 5.1 Windows Forms
#
#  Algoritmo JDE  : J. Meeus, "Astronomical Algorithms", Cap. 50
#  Distanza (km)  : J. Meeus, "Astronomical Algorithms", Cap. 47
#
#  Accuratezza data/ora : +/- 1-2 ore
#  Accuratezza distanza : +/- 500-1500 km
# ============================================================
# REGOLA PS 5.1: mai espressioni aritmetiche dentro
#   New-Object System.Drawing.Point/Size -> pre-calcolare!
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ================================================================
#  Funzioni astronomiche
# ================================================================

function ConvertFrom-JDE {
    param([double]$JDE)
    $jd = $JDE + 0.5
    $Z  = [Math]::Floor($jd)
    $F  = $jd - $Z
    if ($Z -lt 2299161) { $A = $Z }
    else {
        $al = [Math]::Floor(($Z - 1867216.25) / 36524.25)
        $A  = $Z + 1 + $al - [Math]::Floor($al / 4.0)
    }
    $B = $A + 1524
    $C = [Math]::Floor(($B - 122.1) / 365.25)
    $D = [Math]::Floor(365.25 * $C)
    $E = [Math]::Floor(($B - $D) / 30.6001)
    $df   = $B - $D - [Math]::Floor(30.6001 * $E) + $F
    $day  = [Math]::Floor($df)
    $frac = $df - $day
    $month = if ($E -lt 14) { $E - 1 } else { $E - 13 }
    $year  = if ($month -gt 2) { $C - 4716 } else { $C - 4715 }
    $h = [Math]::Floor($frac * 24)
    $m = [Math]::Round(($frac * 24 - $h) * 60)
    if ($m -eq 60) { $h++; $m = 0 }
    if ($h -eq 24) { $h = 0; $day++ }
    return [PSCustomObject]@{
        Year=[int]$year; Month=[int]$month; Day=[int]$day
        Hour=[int]$h;    Minute=[int]$m
    }
}

function Normalize-Angle {
    param([double]$deg)
    $d = $deg % 360.0
    if ($d -lt 0) { $d += 360.0 }
    return $d
}

# ================================================================
#  Calcolo apogei/perigei con distanza
#  Meeus Cap. 50 (JDE) + Cap. 47 (distanza km)
# ================================================================
function Get-LunarApogeoPerigeo {
    param([int]$Anno)

    $PI = [Math]::PI
    $R  = $PI / 180.0

    $mesiNomi = @('',
        'Gennaio','Febbraio','Marzo','Aprile',
        'Maggio','Giugno','Luglio','Agosto',
        'Settembre','Ottobre','Novembre','Dicembre')

    $risultati = @()
    $kApprox   = ($Anno - 1999.97) * 13.2555
    $kInizio   = [int][Math]::Floor($kApprox) - 2

    for ($k0 = $kInizio; $k0 -le ($kInizio + 17); $k0++) {
        foreach ($apogeo in @($false, $true)) {

            $kk   = if ($apogeo) { $k0 + 0.5 } else { [double]$k0 }
            $JDE0 = 2451534.6408 + 27.55454989 * $kk
            $T  = $kk / 1325.55
            $T2 = $T*$T; $T3 = $T2*$T; $T4 = $T3*$T

            $D = Normalize-Angle (171.9179 + 335.9106046*$kk - 0.0100383*$T2 - 0.00001156*$T3 + 0.000000055*$T4)
            $M = Normalize-Angle (347.3477 +  27.1577721*$kk - 0.0008130*$T2 - 0.0000010*$T3)
            $F = Normalize-Angle (316.6109 + 364.5287911*$kk - 0.0125053*$T2 - 0.0000148*$T3)
            $Dr = $D*$R; $Mr = $M*$R; $Fr = $F*$R

            if (-not $apogeo) {
                # Tabella 50.a - correzione JDE per perigeo
                $cJDE  = -1.6769 * [Math]::Sin(2*$Dr)
                $cJDE += +0.4589 * [Math]::Sin(4*$Dr)
                $cJDE += -0.1856 * [Math]::Sin(6*$Dr)
                $cJDE += +0.1092 * [Math]::Sin(8*$Dr)
                $cJDE += -0.0337 * [Math]::Sin(2*$Dr - $Mr)
                $cJDE += +0.0175 * [Math]::Sin(2*$Dr + $Mr)
                $cJDE += -0.0143 * [Math]::Sin($Mr)
                $cJDE += +0.0138 * [Math]::Sin(2*$Dr + 2*$Fr)
                $cJDE += -0.0127 * [Math]::Sin(2*$Dr - 2*$Fr)
                $cJDE += +0.0092 * [Math]::Sin(4*$Dr - $Mr)
                $cJDE += +0.0051 * [Math]::Sin(2*$Dr + 2*$Mr)
                $tipoStr = "PERIGEO"
            } else {
                # Tabella 50.b - correzione JDE per apogeo
                $cJDE  = +0.4392 * [Math]::Sin(2*$Dr)
                $cJDE += -0.0965 * [Math]::Sin(4*$Dr)
                $cJDE += +0.0541 * [Math]::Sin(6*$Dr)
                $cJDE += -0.0137 * [Math]::Sin(8*$Dr)
                $cJDE += +0.0024 * [Math]::Sin(10*$Dr)
                $cJDE += -0.0351 * [Math]::Sin($Mr)
                $cJDE += +0.0077 * [Math]::Sin(2*$Dr - $Mr)
                $cJDE += -0.0115 * [Math]::Sin(2*$Dr + $Mr)
                $cJDE += +0.0100 * [Math]::Sin(2*$Dr + 2*$Fr)
                $cJDE += -0.0095 * [Math]::Sin(2*$Dr - 2*$Fr)
                $cJDE += +0.0097 * [Math]::Sin(4*$Dr - $Mr)
                $cJDE += -0.0018 * [Math]::Sin(2*$Dr + 2*$Mr)
                $tipoStr = "APOGEO"
            }

            $JDE = $JDE0 + $cJDE

            # ---- Distanza (km) - Meeus Cap. 47, 8 termini principali ----
            # M' luna: 0 rad a perigeo, pi rad ad apogeo (approssimazione)
            # Questa ipotesi cattura la variazione principale (fase lunare)
            $T47    = ($JDE - 2451545.0) / 36525.0
            $E47    = 1.0 - 0.002516*$T47 - 0.0000074*$T47*$T47
            $MsR    = Normalize-Angle(357.5291 + 35999.0503*$T47) * $R  # anomalia Sole
            $Mm     = if (-not $apogeo) { 0.0 } else { $PI }             # anomalia Luna

            $rKm  = 385000.56
            $rKm -= 20905.355 * [Math]::Cos($Mm)
            $rKm -=  3699.111 * [Math]::Cos(2*$Dr - $Mm)
            $rKm -=  2955.968 * [Math]::Cos(2*$Dr)
            $rKm -=   569.925 * [Math]::Cos(2*$Mm)
            $rKm -=   246.158 * [Math]::Cos(2*$Dr + $Mm)
            $rKm +=   205.436 * $E47 * [Math]::Cos($MsR)
            $rKm +=   171.733 * [Math]::Cos(4*$Dr - $Mm)
            $rKm +=   152.138 * $E47 * [Math]::Cos(2*$Dr - $MsR)
            $distKm = [Math]::Round([Math]::Abs($rKm))

            $dt = ConvertFrom-JDE -JDE $JDE

            if ($dt.Year -eq $Anno) {
                $risultati += [PSCustomObject]@{
                    Tipo    = $tipoStr
                    Mese    = $dt.Month
                    Giorno  = $dt.Day
                    OraMin  = $dt.Hour * 60 + $dt.Minute
                    DataStr = "$($dt.Day.ToString('00')) $($mesiNomi[$dt.Month])"
                    OraStr  = "$($dt.Hour.ToString('00')):$($dt.Minute.ToString('00'))"
                    DistKm  = $distKm
                }
            }
        }
    }
    return $risultati | Sort-Object Mese, Giorno, OraMin
}

# Barra ASCII proporzionale nel range 352.000-408.000 km
# Perigei -> sinistra   Apogei -> destra
function Get-DistBar {
    param([int]$km, [int]$width = 24)
    $ratio = ($km - 352000.0) / 56000.0
    if ($ratio -lt 0) { $ratio = 0.0 }
    if ($ratio -gt 1) { $ratio = 1.0 }
    $fill  = [int]($ratio * $width)
    return "[" + ("=" * $fill) + ("." * ($width - $fill)) + "]"
}

# km con punto migliaia italiano  (362.158 km)
function Format-Km {
    param([int]$km)
    return ($km.ToString("#,##0").Replace(",", ".")) + " km"
}

# ================================================================
#  INTERFACCIA GRAFICA
#  Soluzione header: usare gli HEADER NATIVI del ListView
#  -> allineamento automatico e perfetto con le colonne dati
# ================================================================

# -- Dimensioni principali --
$fW = 860
$fH = 630

# -- Coordinate (tutte pre-calcolate, mai aritmetica in New-Object) --
$xM   = 8;   $wM   = 844
$yTop = 0;   $hTop = 58
$yIn  = 66;  $hIn  = 44
$yL1  = 118; $hL1  = 16
$yL2  = 134; $hL2  = 16

# Il ListView inizia subito dopo la legenda; il suo header nativo
# fa parte del controllo stesso -> niente pannello separato
$yLv  = 152; $hLv  = 430

$form = New-Object System.Windows.Forms.Form
$form.Text            = "Luna - Apogeo e Perigeo"
$form.ClientSize      = New-Object System.Drawing.Size($fW, $fH)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox     = $false
$form.BackColor       = [System.Drawing.Color]::FromArgb(236, 239, 250)

$fontMono  = New-Object System.Drawing.Font("Courier New", 9)
$fontBold  = New-Object System.Drawing.Font("Courier New", 9,  [System.Drawing.FontStyle]::Bold)
$fontTitle = New-Object System.Drawing.Font("Courier New", 11, [System.Drawing.FontStyle]::Bold)
$fontSmall = New-Object System.Drawing.Font("Courier New", 8)

# ---- Pannello Titolo ----
$pTop = New-Object System.Windows.Forms.Panel
$pTop.BackColor = [System.Drawing.Color]::FromArgb(22, 52, 110)
$pTop.Location  = New-Object System.Drawing.Point($yTop, $yTop)
$pTop.Size      = New-Object System.Drawing.Size($fW, $hTop)
$form.Controls.Add($pTop)

$lTitle = New-Object System.Windows.Forms.Label
$lTitle.Text      = "*  LUNA - APOGEO E PERIGEO  *"
$lTitle.Font      = $fontTitle
$lTitle.ForeColor = [System.Drawing.Color]::White
$lTitle.Location  = New-Object System.Drawing.Point(12, 7)
$lTitle.AutoSize  = $true
$pTop.Controls.Add($lTitle)

$lSub = New-Object System.Windows.Forms.Label
$lSub.Text      = "J. Meeus - Astronomical Algorithms  |  Cap. 50 (date)  Cap. 47 (distanza)  |  Orari in UT"
$lSub.Font      = $fontSmall
$lSub.ForeColor = [System.Drawing.Color]::LightSteelBlue
$lSub.Location  = New-Object System.Drawing.Point(12, 33)
$lSub.AutoSize  = $true
$pTop.Controls.Add($lSub)

# ---- Pannello Input ----
$pIn = New-Object System.Windows.Forms.Panel
$pIn.BackColor   = [System.Drawing.Color]::FromArgb(215, 221, 242)
$pIn.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pIn.Location    = New-Object System.Drawing.Point($xM, $yIn)
$pIn.Size        = New-Object System.Drawing.Size($wM, $hIn)
$form.Controls.Add($pIn)

$lAnno = New-Object System.Windows.Forms.Label
$lAnno.Text     = "Anno:"
$lAnno.Font     = $fontBold
$lAnno.Location = New-Object System.Drawing.Point(10, 12)
$lAnno.AutoSize = $true
$pIn.Controls.Add($lAnno)

$txtAnno = New-Object System.Windows.Forms.TextBox
$txtAnno.Text      = (Get-Date).Year.ToString()
$txtAnno.Font      = $fontBold
$txtAnno.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
$txtAnno.Location  = New-Object System.Drawing.Point(60, 9)
$txtAnno.Size      = New-Object System.Drawing.Size(72, 22)
$pIn.Controls.Add($txtAnno)

$btnCalc = New-Object System.Windows.Forms.Button
$btnCalc.Text      = ">> CALCOLA <<"
$btnCalc.Font      = $fontBold
$btnCalc.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnCalc.BackColor = [System.Drawing.Color]::FromArgb(22, 52, 110)
$btnCalc.ForeColor = [System.Drawing.Color]::White
$btnCalc.Location  = New-Object System.Drawing.Point(144, 7)
$btnCalc.Size      = New-Object System.Drawing.Size(130, 28)
$pIn.Controls.Add($btnCalc)

$lHint = New-Object System.Windows.Forms.Label
$lHint.Text      = "Inserire l'anno e premere CALCOLA  (oppure Invio)"
$lHint.Font      = $fontSmall
$lHint.ForeColor = [System.Drawing.Color]::DimGray
$lHint.Location  = New-Object System.Drawing.Point(286, 13)
$lHint.AutoSize  = $true
$pIn.Controls.Add($lHint)

# ---- Legenda (due righe) ----
$lLeg1 = New-Object System.Windows.Forms.Label
$lLeg1.Text      = "  [P] PERIGEO  =  Luna piu VICINA alla Terra  (356.000 - 370.000 km)"
$lLeg1.Font      = $fontSmall
$lLeg1.ForeColor = [System.Drawing.Color]::FromArgb(150, 10, 10)
$lLeg1.Location  = New-Object System.Drawing.Point($xM, $yL1)
$lLeg1.AutoSize  = $true
$form.Controls.Add($lLeg1)

$lLeg2 = New-Object System.Windows.Forms.Label
$lLeg2.Text      = "  [A] APOGEO   =  Luna piu LONTANA dalla Terra  (404.000 - 407.000 km)"
$lLeg2.Font      = $fontSmall
$lLeg2.ForeColor = [System.Drawing.Color]::FromArgb(10, 55, 160)
$lLeg2.Location  = New-Object System.Drawing.Point($xM, $yL2)
$lLeg2.AutoSize  = $true
$form.Controls.Add($lLeg2)

# ================================================================
#  ListView con HEADER NATIVO
#  La larghezza delle colonne e il testo dell'header sono definiti
#  nello stesso Columns.Add() -> allineamento sempre perfetto.
#
#  Larghezze calibrate su Courier New 9pt (~7.8 px/car):
#    "  [P] PERIGEO"  = 14 car  ->  140 px
#    "  01 Gennaio"   = 13 car  ->  148 px
#    "  20:29 UT"     = 10 car  ->   92 px
#    "  362.158 km"   = 12 car  ->  126 px
#    barra 24 car + prefisso  ->  338 px
# ================================================================
$lv = New-Object System.Windows.Forms.ListView
$lv.View          = [System.Windows.Forms.View]::Details
$lv.FullRowSelect = $true
$lv.GridLines     = $true
$lv.Font          = $fontMono
$lv.BackColor     = [System.Drawing.Color]::FromArgb(252, 253, 255)
$lv.BorderStyle   = [System.Windows.Forms.BorderStyle]::FixedSingle

# Header nativo con stile piatto (NonFlat = testo semplice, niente pulsanti)
$lv.HeaderStyle   = [System.Windows.Forms.ColumnHeaderStyle]::Nonflat

$lv.Location = New-Object System.Drawing.Point($xM, $yLv)
$lv.Size     = New-Object System.Drawing.Size($wM, $hLv)

# -- Colonne: il testo nell'header e la larghezza sono sincronizzati --
$lv.Columns.Add("  Tipo",           140) | Out-Null
$lv.Columns.Add("  Data",           148) | Out-Null
$lv.Columns.Add("  Ora UT",          92) | Out-Null
$lv.Columns.Add("  Distanza (km)",  126) | Out-Null
$lv.Columns.Add("  Posizione  [352k" + ("." * 6) + "=====" + ("." * 6) + "408k]", 338) | Out-Null

$form.Controls.Add($lv)

# ---- StatusBar ----
$sb    = New-Object System.Windows.Forms.StatusStrip
$sbLbl = New-Object System.Windows.Forms.ToolStripStatusLabel
$sbLbl.Text = "Pronto  -  Inserire l'anno e premere CALCOLA"
$sbLbl.Font = $fontSmall
$sb.Items.Add($sbLbl) | Out-Null
$form.Controls.Add($sb)

# ================================================================
#  Evento CALCOLA
# ================================================================
$btnCalc.Add_Click({
    $anno = 0
    if (-not [int]::TryParse($txtAnno.Text.Trim(), [ref]$anno) `
        -or $anno -lt 1 -or $anno -gt 9999) {
        [System.Windows.Forms.MessageBox]::Show(
            "Anno non valido. Inserire un numero intero (es. 2025).",
            "Errore",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $lv.Items.Clear()
    $sbLbl.Text = "Calcolo in corso per l'anno $anno ..."
    $form.Refresh()

    try {
        $eventi = Get-LunarApogeoPerigeo -Anno $anno

        if ($eventi.Count -eq 0) {
            $it = New-Object System.Windows.Forms.ListViewItem("  Nessun evento per $anno")
            $lv.Items.Add($it) | Out-Null
        } else {
            $mesePrecente = 0

            foreach ($ev in $eventi) {

                # Riga separatrice al cambio mese (sottile, non invadente)
                if ($ev.Mese -ne $mesePrecente) {
                    if ($mesePrecente -ne 0) {
                        $sep = New-Object System.Windows.Forms.ListViewItem("")
                        $sep.SubItems.Add("") | Out-Null
                        $sep.SubItems.Add("") | Out-Null
                        $sep.SubItems.Add("") | Out-Null
                        $sep.SubItems.Add("") | Out-Null
                        $sep.BackColor = [System.Drawing.Color]::FromArgb(222, 227, 248)
                        $lv.Items.Add($sep) | Out-Null
                    }
                    $mesePrecente = $ev.Mese
                }

                $sim  = if ($ev.Tipo -eq "PERIGEO") { "[P]" } else { "[A]" }
                $bar  = Get-DistBar -km $ev.DistKm -width 24
                $dStr = Format-Km -km $ev.DistKm

                $it = New-Object System.Windows.Forms.ListViewItem("  $sim $($ev.Tipo)")
                $it.SubItems.Add("  $($ev.DataStr)") | Out-Null
                $it.SubItems.Add("  $($ev.OraStr) UT") | Out-Null
                $it.SubItems.Add("  $dStr")             | Out-Null
                $it.SubItems.Add("  $bar")              | Out-Null

                if ($ev.Tipo -eq "PERIGEO") {
                    $it.ForeColor = [System.Drawing.Color]::FromArgb(150, 10, 10)
                    $it.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 245)
                } else {
                    $it.ForeColor = [System.Drawing.Color]::FromArgb(10, 55, 160)
                    $it.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 255)
                }

                $lv.Items.Add($it) | Out-Null
            }
        }

        # Statistiche nella barra di stato
        $eP = @($eventi | Where-Object { $_.Tipo -eq "PERIGEO" })
        $eA = @($eventi | Where-Object { $_.Tipo -eq "APOGEO"  })
        $nP = $eP.Count;  $nA = $eA.Count

        $statP = if ($nP -gt 0) {
            $mn = ($eP | Measure-Object -Property DistKm -Minimum).Minimum
            $mx = ($eP | Measure-Object -Property DistKm -Maximum).Maximum
            "min $(Format-Km $mn), max $(Format-Km $mx)"
        } else { "n/d" }

        $statA = if ($nA -gt 0) {
            $mn = ($eA | Measure-Object -Property DistKm -Minimum).Minimum
            $mx = ($eA | Measure-Object -Property DistKm -Maximum).Maximum
            "min $(Format-Km $mn), max $(Format-Km $mx)"
        } else { "n/d" }

        $sbLbl.Text = "Anno $anno  |  Perigei: $nP ($statP)  |  Apogei: $nA ($statA)"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Errore durante il calcolo:`n$_", "Errore",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        $sbLbl.Text = "Errore."
    }
})

$txtAnno.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $btnCalc.PerformClick() }
})

$btnCalc.PerformClick()
[void]$form.ShowDialog()

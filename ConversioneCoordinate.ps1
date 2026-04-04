#requires -version 5.1
# ============================================================
#  ConversioneCoordinate.ps1
#  Conversione coordinate equatoriali <-> altazimutali
#  Formule: Jean Meeus - Astronomical Formulae for Calculators
#           cap. 7 (Tempo Siderale) e cap. 8 (Trasformazione)
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ============================================================
#  FUNZIONI ASTRONOMICHE
# ============================================================
function ToRad($deg) { $deg * [Math]::PI / 180.0 }
function ToDeg($rad) { $rad * 180.0 / [Math]::PI }

function Norm360($a) { $r = $a % 360; if ($r -lt 0) { $r += 360 }; return $r }
function Norm24($h)  { $r = $h % 24;  if ($r -lt 0) { $r += 24  }; return $r }

function Get-JD($anno, $mese, $giorno, $oreUT) {
    $d = $giorno + $oreUT / 24.0
    if ($mese -le 2) { $anno--; $mese += 12 }
    $A  = [Math]::Floor($anno / 100.0)
    $B  = 2 - $A + [Math]::Floor($A / 4.0)
    return ([Math]::Floor(365.25*($anno+4716)) + [Math]::Floor(30.6001*($mese+1)) + $d + $B - 1524.5)
}

function Get-GMST0h($JD0h) {
    $T = ($JD0h - 2451545.0) / 36525.0
    return Norm360(100.4606184 + 36000.7700536*$T + 0.000387933*$T*$T - $T*$T*$T/38710000.0)
}

function Get-LST_deg($anno, $mese, $giorno, $oreUT, $lonEst) {
    $JD = Get-JD $anno $mese $giorno 0.0
    return Norm360((Get-GMST0h $JD) + $oreUT*15.0*1.00273790935 + $lonEst)
}

# BUGFIX: formato secondi corretto con separazione esplicita
function OreToHMS($h) {
    $h  = Norm24 $h
    $hh = [Math]::Floor($h)
    $rm = ($h - $hh) * 60.0
    $mm = [Math]::Floor($rm)
    $ss = ($rm - $mm) * 60.0
    $ssStr = ("{0:00.00}" -f $ss)
    return ("{0:D2}h {1:D2}m {2}s" -f [int]$hh, [int]$mm, $ssStr)
}

function GradToDMS($d) {
    $segno = if ($d -lt 0) { "-" } else { "+" }
    $d  = [Math]::Abs($d)
    $dd = [Math]::Floor($d)
    $rm = ($d - $dd) * 60.0
    $mm = [Math]::Floor($rm)
    $ss = ($rm - $mm) * 60.0
    $ssStr = ("{0:00.0}" -f $ss)
    return ("{0}{1:D2}d {2:D2}' {3}`"" -f $segno, [int]$dd, [int]$mm, $ssStr)
}

function Parse-RA($hh, $mm, $ss) {
    return (Norm24([double]$hh + [double]$mm/60.0 + [double]$ss/3600.0)) * 15.0
}

function Parse-Dec($dd_str, $mm, $ss) {
    $neg = $dd_str.Trim().StartsWith("-")
    $dec = [Math]::Abs([double]$dd_str) + [double]$mm/60.0 + [double]$ss/3600.0
    if ($neg) { $dec = -$dec }
    return $dec
}

function Convert-EqToHor($RA_deg, $Dec_deg, $Lat_deg, $LST_deg) {
    $H_r = ToRad(Norm360($LST_deg - $RA_deg))
    $d_r = ToRad $Dec_deg
    $p_r = ToRad $Lat_deg
    $sinh   = [Math]::Sin($p_r)*[Math]::Sin($d_r) + [Math]::Cos($p_r)*[Math]::Cos($d_r)*[Math]::Cos($H_r)
    $alt    = ToDeg([Math]::Asin($sinh))
    $Az_sud = ToDeg([Math]::Atan2([Math]::Sin($H_r), [Math]::Cos($H_r)*[Math]::Sin($p_r) - [Math]::Tan($d_r)*[Math]::Cos($p_r)))
    $Az     = Norm360(180.0 + $Az_sud)
    return [PSCustomObject]@{ Alt=$alt; Az=$Az; H_ore=Norm24((Norm360($LST_deg-$RA_deg))/15.0) }
}

function Convert-HorToEq($Az_nord, $Alt_deg, $Lat_deg, $LST_deg) {
    $A_r = ToRad($Az_nord - 180.0)
    $h_r = ToRad $Alt_deg
    $p_r = ToRad $Lat_deg
    $sind = [Math]::Sin($h_r)*[Math]::Sin($p_r) + [Math]::Cos($h_r)*[Math]::Cos($p_r)*[Math]::Cos($A_r)
    $Dec  = ToDeg([Math]::Asin($sind))
    $H_raw= ToDeg([Math]::Atan2([Math]::Sin($A_r), [Math]::Cos($A_r)*[Math]::Sin($p_r) + [Math]::Tan($h_r)*[Math]::Cos($p_r)))
    $H    = Norm360($H_raw)
    $RA  = Norm360($LST_deg - $H)
    return [PSCustomObject]@{ RA_deg=$RA; Dec_deg=$Dec; H_ore=Norm24($H/15.0) }
}

# ============================================================
#  PALETTE COLORI
# ============================================================
$cBg       = [System.Drawing.Color]::FromArgb(22,  22,  26 )   # sfondo form
$cPanel    = [System.Drawing.Color]::FromArgb(38,  38,  44 )   # sfondo groupbox
$cHeader   = [System.Drawing.Color]::FromArgb(0,   100, 160)   # striscia header
$cInput    = [System.Drawing.Color]::FromArgb(18,  18,  22 )   # sfondo textbox
$cAccent   = [System.Drawing.Color]::FromArgb(0,   150, 215)   # blu accento
$cResult   = [System.Drawing.Color]::FromArgb(52,  52,  60 )   # sfondo area risultati
$cBorder   = [System.Drawing.Color]::FromArgb(0,   120, 190)   # bordo area risultati
$cLbl      = [System.Drawing.Color]::FromArgb(170, 175, 185)   # testo label
$cGold     = [System.Drawing.Color]::FromArgb(255, 200, 50 )   # valori risultato
$cGreen    = [System.Drawing.Color]::FromArgb(70,  210, 115)   # LST / H
$cWarn     = [System.Drawing.Color]::FromArgb(220, 75,  65 )   # warning
$cDim      = [System.Drawing.Color]::FromArgb(100, 100, 110)   # status bar
$cWhite    = [System.Drawing.Color]::White

$FS   = 11   # font size base
$FSr  = 12   # font size risultati
$LH   = 30   # line height

# ============================================================
#  HELPER COSTRUTTORI
# ============================================================
function fnt($size=11, $bold=$false) {
    $style = if ($bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    return New-Object System.Drawing.Font("Consolas", $size, $style)
}

function pt($x,$y)  { New-Object System.Drawing.Point($x,$y) }
function sz($w,$h)  { New-Object System.Drawing.Size($w,$h) }

function New-Lbl($txt,$x,$y,$w=200,$h=28) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text=$txt; $l.Location=pt $x $y; $l.Size=sz $w $h
    $l.ForeColor=$cLbl; $l.TextAlign="MiddleLeft"; $l.Font=fnt $FS
    $l.BackColor=[System.Drawing.Color]::Transparent
    return $l
}

function New-TB($x,$y,$w=80,$val="0") {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location=pt $x $y; $t.Size=sz $w 26
    $t.Text="$val"; $t.BackColor=$cInput; $t.ForeColor=$cWhite
    $t.BorderStyle="FixedSingle"; $t.TextAlign="Center"; $t.Font=fnt $FS
    return $t
}

function New-GB($title,$x,$y,$w,$h) {
    $gb = New-Object System.Windows.Forms.GroupBox
    $gb.Text="  $title  "; $gb.Location=pt $x $y; $gb.Size=sz $w $h
    $gb.ForeColor=$cAccent; $gb.BackColor=$cPanel; $gb.Font=fnt $FS $true
    return $gb
}

function New-Btn($txt,$x,$y,$w=200,$h=44) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text=$txt; $b.Location=pt $x $y; $b.Size=sz $w $h
    $b.BackColor=$cAccent; $b.ForeColor=$cWhite
    $b.FlatStyle="Flat"; $b.FlatAppearance.BorderSize=0
    $b.Font=fnt $FS $true
    return $b
}

# Pannello risultato con sfondo distinto e bordo sinistro colorato
function New-ResultPanel($x,$y,$w,$h) {
    $p = New-Object System.Windows.Forms.Panel
    $p.Location=pt $x $y; $p.Size=sz $w $h
    $p.BackColor=$cResult
    $p.BorderStyle="None"
    return $p
}

function New-ResLbl($txt,$x,$y,$w=550,$color=$null) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text=$txt; $l.Location=pt $x $y; $l.Size=sz $w $LH
    $l.ForeColor=if($color){$color}else{$cGold}
    $l.TextAlign="MiddleLeft"; $l.Font=fnt $FSr $true
    $l.BackColor=[System.Drawing.Color]::Transparent
    return $l
}

function New-ResKey($txt,$x,$y,$w=175) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text=$txt; $l.Location=pt $x $y; $l.Size=sz $w $LH
    $l.ForeColor=[System.Drawing.Color]::FromArgb(140,145,160)
    $l.TextAlign="MiddleLeft"; $l.Font=fnt ($FS-1) $false
    $l.BackColor=[System.Drawing.Color]::Transparent
    return $l
}

# Linea separatrice orizzontale
function New-Sep($parent,$x,$y,$w) {
    $s = New-Object System.Windows.Forms.Label
    $s.Location=pt $x $y; $s.Size=sz $w 1
    $s.BackColor=[System.Drawing.Color]::FromArgb(65,65,80)
    $parent.Controls.Add($s)
}

# Decorazione bordo sinistro per area risultati
function New-LeftBar($parent,$x,$y,$h) {
    $b = New-Object System.Windows.Forms.Label
    $b.Location=pt $x $y; $b.Size=sz 3 $h
    $b.BackColor=$cAccent
    $parent.Controls.Add($b)
}

# ============================================================
#  FORM PRINCIPALE
# ============================================================
$PW = 880

$form = New-Object System.Windows.Forms.Form
$form.Text            = "Conversione Coordinate  -  Equatoriali  <->  Altazimutali"
$form.Size            = sz 920 920
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox     = $false
$form.Font            = fnt $FS
$form.BackColor       = $cBg

# ============================================================
#  HEADER DECORATIVO
# ============================================================
$hdr = New-Object System.Windows.Forms.Panel
$hdr.Location = pt 0 0; $hdr.Size = sz 920 38; $hdr.BackColor = $cHeader
$form.Controls.Add($hdr)

$hdrLbl = New-Object System.Windows.Forms.Label
$hdrLbl.Text = "  CONVERSIONE COORDINATE ASTRONOMICHE   |   Meeus cap. 7-8"
$hdrLbl.Location = pt 0 0; $hdrLbl.Size = sz 920 38
$hdrLbl.ForeColor = [System.Drawing.Color]::FromArgb(210,235,255)
$hdrLbl.TextAlign = "MiddleLeft"
$hdrLbl.Font = fnt $FS $true
$hdrLbl.BackColor = [System.Drawing.Color]::Transparent
$hdr.Controls.Add($hdrLbl)

# ============================================================
#  PANNELLO OSSERVATORE
# ============================================================
$gbObs = New-GB "Osservatore  e  Data / Ora UTC" 10 48 $PW 150
$form.Controls.Add($gbObs)

# riga 1: Lat / Anno Mese Giorno
$r1 = 30; $r2 = $r1 + $LH + 8

$gbObs.Controls.Add((New-Lbl "Latitudine  (deg) :" 12 $r1 185))
$txtLat = New-TB 200 ($r1+2) 90 "45.464"
$gbObs.Controls.Add($txtLat)
$gbObs.Controls.Add((New-Lbl "+N / -S" 298 $r1 80))

$gbObs.Controls.Add((New-Lbl "Longitudine (deg) :" 12 $r2 185))
$txtLon = New-TB 200 ($r2+2) 90 "9.189"
$gbObs.Controls.Add($txtLon)
$gbObs.Controls.Add((New-Lbl "+E / -W" 298 $r2 80))

# separatore verticale
$vSep = New-Object System.Windows.Forms.Label
$vSep.Location=pt 395 22; $vSep.Size=sz 1 100; $vSep.BackColor=[System.Drawing.Color]::FromArgb(65,65,80)
$gbObs.Controls.Add($vSep)

$gbObs.Controls.Add((New-Lbl "Anno  :" 410 $r1 75))
$txtAnno = New-TB 490 ($r1+2) 72 (Get-Date).Year
$gbObs.Controls.Add($txtAnno)
$gbObs.Controls.Add((New-Lbl "Mese :" 572 $r1 65))
$txtMese = New-TB 640 ($r1+2) 48 (Get-Date).Month
$gbObs.Controls.Add($txtMese)
$gbObs.Controls.Add((New-Lbl "Giorno:" 698 $r1 70))
$txtGiorno = New-TB 772 ($r1+2) 50 (Get-Date).Day
$gbObs.Controls.Add($txtGiorno)

$gbObs.Controls.Add((New-Lbl "Ora UT   hh :" 410 $r2 125))
$txtOraH = New-TB 538 ($r2+2) 50 ([int](Get-Date).ToUniversalTime().Hour)
$gbObs.Controls.Add($txtOraH)
$gbObs.Controls.Add((New-Lbl "mm :" 596 $r2 45))
$txtOraM = New-TB 644 ($r2+2) 50 ([int](Get-Date).ToUniversalTime().Minute)
$gbObs.Controls.Add($txtOraM)
$gbObs.Controls.Add((New-Lbl "ss :" 702 $r2 45))
$txtOraS = New-TB 750 ($r2+2) 50 "0"
$gbObs.Controls.Add($txtOraS)

# Riga LST
New-Sep $gbObs 10 100 ($PW-30)

$pnlLST = New-Object System.Windows.Forms.Panel
$pnlLST.Location=pt 10 106; $pnlLST.Size=sz ($PW-30) 32; $pnlLST.BackColor=[System.Drawing.Color]::FromArgb(28,50,38)
$gbObs.Controls.Add($pnlLST)

$dotLST = New-Object System.Windows.Forms.Label
$dotLST.Location=pt 0 0; $dotLST.Size=sz 4 32; $dotLST.BackColor=$cGreen
$pnlLST.Controls.Add($dotLST)

$lblLST = New-Object System.Windows.Forms.Label
$lblLST.Location=pt 12 0; $lblLST.Size=sz ($PW-50) 32
$lblLST.ForeColor=$cGreen; $lblLST.TextAlign="MiddleLeft"
$lblLST.Font=fnt $FS $true; $lblLST.Text="LST : --"
$lblLST.BackColor=[System.Drawing.Color]::Transparent
$pnlLST.Controls.Add($lblLST)

# ============================================================
#  PANNELLO EQ -> HOR
# ============================================================
$gbEQ = New-GB "EQUATORIALI  -->  ALTAZIMUTALI" 10 208 $PW 280
$form.Controls.Add($gbEQ)

# Input section
$i1=32; $i2=$i1+$LH+6

$gbEQ.Controls.Add((New-Lbl "AR  ( hh    mm    ss ) :" 12 $i1 220))
$txtRA_h = New-TB 236 ($i1+2) 58 "10"
$gbEQ.Controls.Add($txtRA_h)
$gbEQ.Controls.Add((New-Lbl "h" 298 $i1 20))
$txtRA_m = New-TB 320 ($i1+2) 58 "57"
$gbEQ.Controls.Add($txtRA_m)
$gbEQ.Controls.Add((New-Lbl "m" 382 $i1 20))
$txtRA_s = New-TB 404 ($i1+2) 72 "35.8"
$gbEQ.Controls.Add($txtRA_s)
$gbEQ.Controls.Add((New-Lbl "s" 480 $i1 20))

$gbEQ.Controls.Add((New-Lbl "Dec  (+N/-S   dd   mm   ss ) :" 12 $i2 268))
$txtDec_d = New-TB 284 ($i2+2) 60 "+8"
$gbEQ.Controls.Add($txtDec_d)
$gbEQ.Controls.Add((New-Lbl "d" 348 $i2 20))
$txtDec_m = New-TB 370 ($i2+2) 58 "25"
$gbEQ.Controls.Add($txtDec_m)
$gbEQ.Controls.Add((New-Lbl "m" 432 $i2 20))
$txtDec_s = New-TB 454 ($i2+2) 72 "58.1"
$gbEQ.Controls.Add($txtDec_s)
$gbEQ.Controls.Add((New-Lbl "s" 530 $i2 20))

# Bottone sulla destra - centrato verticalmente tra i2 righe input
$btnEQ = New-Btn "  >>>  Converti  >>>  " 660 ($i1-2) 195 ($LH*2+14)
$gbEQ.Controls.Add($btnEQ)

# Area risultati
New-Sep $gbEQ 10 106 ($PW-30)

$pnlResEQ = New-ResultPanel 10 112 ($PW-30) 138
$gbEQ.Controls.Add($pnlResEQ)
New-LeftBar $gbEQ 10 112 138

$pnlResEQ.Controls.Add((New-ResKey "Azimut (da Nord) :" 14 4 170))
$lblAz = New-ResLbl "--" 192 4
$pnlResEQ.Controls.Add($lblAz)

$pnlResEQ.Controls.Add((New-ResKey "Altitudine        :" 14 36 170))
$lblAlt = New-ResLbl "--" 192 36
$pnlResEQ.Controls.Add($lblAlt)

$pnlResEQ.Controls.Add((New-ResKey "Angolo Orario (H) :" 14 68 170))
$lblH_eq = New-ResLbl "--" 192 68 550 $cGreen
$pnlResEQ.Controls.Add($lblH_eq)

$pnlResEQ.Controls.Add((New-ResKey "Info              :" 14 100 170))
$lblInfo_eq = New-ResLbl "" 192 100 550 $cDim
$lblInfo_eq.Font = fnt ($FS-1) $false
$pnlResEQ.Controls.Add($lblInfo_eq)

# ============================================================
#  PANNELLO HOR -> EQ
# ============================================================
$gbHOR = New-GB "ALTAZIMUTALI  -->  EQUATORIALI" 10 498 $PW 280
$form.Controls.Add($gbHOR)

$h1=32; $h2=$h1+$LH+6

$gbHOR.Controls.Add((New-Lbl "Azimut da Nord  (deg) :" 12 $h1 228))
$txtAz = New-TB 244 ($h1+2) 110 "180.0"
$gbHOR.Controls.Add($txtAz)
$hintAz = New-Lbl "[ 0-360 ]  N=0  E=90  S=180  W=270" 364 $h1 285
$hintAz.Font = fnt 9; $gbHOR.Controls.Add($hintAz)

$gbHOR.Controls.Add((New-Lbl "Altitudine      (deg) :" 12 $h2 228))
$txtAltIn = New-TB 244 ($h2+2) 110 "30.0"
$gbHOR.Controls.Add($txtAltIn)
$hintAlt = New-Lbl "[ -90/+90 ]  0=orizz  90=zenit" 364 $h2 285
$hintAlt.Font = fnt 9; $gbHOR.Controls.Add($hintAlt)

# Bottone - sotto i due input, non sovrapposto
$btnHOR = New-Btn "  >>>  Converti  >>>  " 660 ($h1-2) 195 ($LH*2+14)
$gbHOR.Controls.Add($btnHOR)

# Area risultati
New-Sep $gbHOR 10 106 ($PW-30)

$pnlResHOR = New-ResultPanel 10 112 ($PW-30) 138
$gbHOR.Controls.Add($pnlResHOR)
New-LeftBar $gbHOR 10 112 138

$pnlResHOR.Controls.Add((New-ResKey "Asc. Retta (AR)   :" 14 4 170))
$lblRA_out = New-ResLbl "--" 192 4
$pnlResHOR.Controls.Add($lblRA_out)

$pnlResHOR.Controls.Add((New-ResKey "Declinazione       :" 14 36 170))
$lblDec_out = New-ResLbl "--" 192 36
$pnlResHOR.Controls.Add($lblDec_out)

$pnlResHOR.Controls.Add((New-ResKey "Angolo Orario (H) :" 14 68 170))
$lblH_hor = New-ResLbl "--" 192 68 550 $cGreen
$pnlResHOR.Controls.Add($lblH_hor)

$pnlResHOR.Controls.Add((New-ResKey "Info              :" 14 100 170))
$lblInfo_hor = New-ResLbl "" 192 100 550 $cDim
$lblInfo_hor.Font = fnt ($FS-1) $false
$pnlResHOR.Controls.Add($lblInfo_hor)

# ============================================================
#  FOOTER
# ============================================================
$footer = New-Object System.Windows.Forms.Panel
$footer.Location=pt 0 792; $footer.Size=sz 920 32; $footer.BackColor=[System.Drawing.Color]::FromArgb(32,32,38)
$form.Controls.Add($footer)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location=pt 14 0; $lblStatus.Size=sz 900 32
$lblStatus.ForeColor=$cDim; $lblStatus.Font=fnt 10; $lblStatus.TextAlign="MiddleLeft"
$lblStatus.Text="Formule: J. Meeus - Astronomical Formulae for Calculators, cap. 7-8"
$lblStatus.BackColor=[System.Drawing.Color]::Transparent
$footer.Controls.Add($lblStatus)

# ============================================================
#  GET-PARAMS
# ============================================================
function Get-Params {
    try {
        $lat   = [double]$txtLat.Text
        $lon   = [double]$txtLon.Text
        $anno  = [int]$txtAnno.Text
        $mese  = [int]$txtMese.Text
        $gg    = [int]$txtGiorno.Text
        $oreUT = [int]$txtOraH.Text + [int]$txtOraM.Text/60.0 + [double]$txtOraS.Text/3600.0

        $LST  = Get-LST_deg $anno $mese $gg $oreUT $lon
        $LSTh = Norm24($LST / 15.0)
        $lblLST.Text = "LST :  {0}   =   {1:F5} deg" -f (OreToHMS $LSTh), $LST
        return [PSCustomObject]@{ Lat=$lat; Lon=$lon; LST=$LST; OK=$true }
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Errore nei dati osservatore / data:`n$($_.Exception.Message)",
            "Errore input", "OK", "Error") | Out-Null
        return [PSCustomObject]@{ OK=$false }
    }
}

# ============================================================
#  EVENTO EQ -> HOR
# ============================================================
$btnEQ.Add_Click({
    $p = Get-Params; if (-not $p.OK) { return }
    try {
        $RA  = Parse-RA  $txtRA_h.Text $txtRA_m.Text $txtRA_s.Text
        $Dec = Parse-Dec $txtDec_d.Text $txtDec_m.Text $txtDec_s.Text
        if ([Math]::Abs($Dec) -gt 90) { throw "Declinazione fuori range [-90, +90]." }

        $r = Convert-EqToHor $RA $Dec $p.Lat $p.LST

        $lblAz.Text  = "{0,9:F4} deg    ( {1} )" -f $r.Az,  (GradToDMS $r.Az)
        $lblAlt.Text = "{0,9:F4} deg    ( {1} )" -f $r.Alt, (GradToDMS $r.Alt)
        $lblH_eq.Text= "H = {0}   ( {1:F4} deg )" -f (OreToHMS $r.H_ore), ($r.H_ore*15)

        if ($r.Alt -lt 0) {
            $lblInfo_eq.ForeColor = $cWarn
            $lblInfo_eq.Text = "  [!]  Oggetto SOTTO l'orizzonte  (alt < 0)"
            $lblStatus.ForeColor = $cWarn
            $lblStatus.Text = "  ATTENZIONE: altitudine negativa - oggetto non osservabile"
        } else {
            $lblInfo_eq.ForeColor = $cGreen
            $lblInfo_eq.Text = "  [ok] Oggetto visibile sopra l'orizzonte"
            $lblStatus.ForeColor = $cDim
            $lblStatus.Text = "  Conversione EQ --> HOR completata."
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Errore:`n$($_.Exception.Message)", "Errore", "OK", "Error") | Out-Null
    }
})

# ============================================================
#  EVENTO HOR -> EQ
# ============================================================
$btnHOR.Add_Click({
    $p = Get-Params; if (-not $p.OK) { return }
    try {
        $Az  = [double]$txtAz.Text
        $Alt = [double]$txtAltIn.Text
        if ([Math]::Abs($Alt) -gt 90) { throw "Altitudine fuori range [-90, +90]." }

        $r = Convert-HorToEq $Az $Alt $p.Lat $p.LST

        $lblRA_out.Text  = "{0}   ( {1:F4} deg )" -f (OreToHMS($r.RA_deg/15.0)), $r.RA_deg
        $lblDec_out.Text = "{0,9:F4} deg    ( {1} )"  -f $r.Dec_deg, (GradToDMS $r.Dec_deg)
        $lblH_hor.Text   = "H = {0}   ( {1:F4} deg )" -f (OreToHMS $r.H_ore), ($r.H_ore*15)
        $lblInfo_hor.ForeColor = $cGreen
        $lblInfo_hor.Text = "  [ok] Conversione HOR --> EQ completata."

        $lblStatus.ForeColor = $cDim
        $lblStatus.Text = "  Conversione HOR --> EQ completata."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Errore:`n$($_.Exception.Message)", "Errore", "OK", "Error") | Out-Null
    }
})

$form.Add_Shown({ $null = Get-Params })
[System.Windows.Forms.Application]::Run($form)

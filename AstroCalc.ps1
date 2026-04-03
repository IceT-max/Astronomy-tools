# AstroCalc.ps1 - Calcolatore astronomico per PowerShell 5.1
# Algoritmi da Jean Meeus - Astronomical Algorithms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

#region === MATEMATICA DI BASE ===
function Rad  { param([double]$d); $d * [Math]::PI / 180.0 }
function Deg  { param([double]$r); $r * 180.0 / [Math]::PI }
function N360 { param([double]$x); $x - 360.0 * [Math]::Floor($x / 360.0) }

function SolveKepler {
    param([double]$M_deg, [double]$ecc)
    $M_r = Rad $M_deg
    $E   = $M_r
    for ($i = 0; $i -lt 50; $i++) {
        $sinE = [Math]::Sin($E)
        $cosE = [Math]::Cos($E)
        $dE   = ($M_r - $E + $ecc * $sinE) / (1.0 - $ecc * $cosE)
        $E   += $dE
        if ([Math]::Abs($dE) -lt 1e-10) { break }
    }
    return $E
}

function TrueAnomaly {
    param([double]$E_rad, [double]$ecc)
    $cosE   = [Math]::Cos($E_rad)
    $cv     = ($cosE - $ecc) / (1.0 - $ecc * $cosE)
    $sv     = [Math]::Sqrt(1.0 - $ecc * $ecc) * [Math]::Sin($E_rad) / (1.0 - $ecc * $cosE)
    return (Deg ([Math]::Atan2($sv, $cv)))
}
#endregion

#region === FUNZIONI ASTRONOMICHE ===
function JulianDate {
    param([int]$Y,[int]$M,[int]$D,[double]$UT)
    if ($M -le 2) { $Y -= 1; $M += 12 }
    $A = [Math]::Floor($Y / 100.0)
    $B = 2 - $A + [Math]::Floor($A / 4.0)
    return ([Math]::Floor(365.25 * ($Y + 4716)) + [Math]::Floor(30.6001 * ($M + 1)) + $D + $B - 1524.5 + $UT / 24.0)
}

function GST {
    param([double]$JD)
    $T = ($JD - 2451545.0) / 36525.0
    $T2 = $T * $T
    $T3 = $T2 * $T
    $g = 280.46061837 + 360.98564736629 * ($JD - 2451545.0) + 0.000387933 * $T2 - $T3 / 38710000.0
    return (N360 $g)
}

function Obliq {
    param([double]$T)
    return (23.439291111 - 0.013004167 * $T - 0.0000001639 * $T * $T + 0.0000005036 * $T * $T * $T)
}

function Ecl2Equ {
    param([double]$lam_deg, [double]$bet_deg, [double]$eps_deg)
    $lr = Rad $lam_deg
    $br = Rad $bet_deg
    $er = Rad $eps_deg
    $sinLr = [Math]::Sin($lr)
    $cosLr = [Math]::Cos($lr)
    $sinBr = [Math]::Sin($br)
    $cosBr = [Math]::Cos($br)
    $sinEr = [Math]::Sin($er)
    $cosEr = [Math]::Cos($er)
    $tanBr = [Math]::Tan($br)
    $ra  = Deg ([Math]::Atan2($sinLr * $cosEr - $tanBr * $sinEr, $cosLr))
    $dec = Deg ([Math]::Asin($sinBr * $cosEr + $cosBr * $sinEr * $sinLr))
    return @{ RA = (N360 $ra); Dec = $dec }
}

function AltAz {
    param([double]$RA_deg,[double]$Dec_deg,[double]$LST_deg,[double]$Lat_deg)
    $H_deg    = N360 ($LST_deg - $RA_deg)
    $Hr       = Rad $H_deg
    $Dr       = Rad $Dec_deg
    $Lr       = Rad $Lat_deg
    $sinH     = [Math]::Sin($Hr)
    $cosH     = [Math]::Cos($Hr)
    $sinD     = [Math]::Sin($Dr)
    $cosD     = [Math]::Cos($Dr)
    $sinL     = [Math]::Sin($Lr)
    $cosL     = [Math]::Cos($Lr)
    $sinAlt   = $sinD * $sinL + $cosD * $cosL * $cosH
    $alt      = Deg ([Math]::Asin($sinAlt))
    $az       = Deg ([Math]::Atan2($sinH, $cosH * $sinL - [Math]::Tan($Dr) * $cosL))
    return @{ Alt = $alt; Az = (N360 ($az + 180.0)) }
}

function FmtRA {
    param([double]$deg)
    $h   = $deg / 15.0
    $hh  = [int][Math]::Floor($h)
    $mm  = [int][Math]::Floor(($h - $hh) * 60.0)
    $ssd = (($h - $hh) * 60.0 - $mm) * 60.0
    $ssi = [int][Math]::Floor($ssd)
    $ssf = [int][Math]::Round(($ssd - $ssi) * 10.0)
    if ($ssf -ge 10) { $ssi += 1; $ssf = 0 }
    return ("{0:00}h {1:00}m {2:00}.{3}s" -f $hh, $mm, $ssi, $ssf)
}

function FmtDec {
    param([double]$deg)
    $DG  = [char]176   # simbolo gradi
    $sgn = if ($deg -lt 0) { "-" } else { "+" }
    $a   = [Math]::Abs($deg)
    $dd  = [int][Math]::Floor($a)
    $mm  = [int][Math]::Floor(($a - $dd) * 60.0)
    $ss  = [int][Math]::Round((($a - $dd) * 60.0 - $mm) * 60.0)
    if ($ss -ge 60) { $mm += 1; $ss = 0 }
    if ($mm -ge 60) { $dd += 1; $mm = 0 }
    return ("{0}{1:00}{2} {3:00}' {4:00}`"" -f $sgn, $dd, $DG, $mm, $ss)
}

function FmtAng {
    param([double]$d)
    $DG = [char]176
    return ("{0:0.00}{1}" -f $d, $DG)
}
#endregion

#region === POSIZIONE SOLE ===
function SunPos {
    param([double]$JD)
    $T  = ($JD - 2451545.0) / 36525.0
    $T2 = $T * $T
    $L0 = N360 (280.46646   + 36000.76983  * $T + 0.0003032 * $T2)
    $M  = N360 (357.52911   + 35999.05029  * $T - 0.0001537 * $T2)
    $Mr = Rad $M
    $e  = 0.016708634 - 0.000042037 * $T - 0.0000001267 * $T2
    $C  = (1.914602 - 0.004817 * $T - 0.000014 * $T2) * [Math]::Sin($Mr) +
          (0.019993 - 0.000101 * $T) * [Math]::Sin(2.0 * $Mr) +
           0.000289 * [Math]::Sin(3.0 * $Mr)
    $theta = N360 ($L0 + $C)
    $nu    = N360 ($M  + $C)
    $nur   = Rad $nu
    $R     = 1.000001018 * (1.0 - $e * $e) / (1.0 + $e * [Math]::Cos($nur))
    $om    = 125.04 - 1934.136 * $T
    $omr   = Rad $om
    $lam   = N360 ($theta - 0.00569 - 0.00478 * [Math]::Sin($omr))
    $eps   = (Obliq $T) + 0.00256 * [Math]::Cos($omr)
    $eq    = Ecl2Equ $lam 0.0 $eps
    return @{ RA = $eq.RA; Dec = $eq.Dec; R = $R; EL = $lam; Mag = -26.74 }
}
#endregion

#region === POSIZIONE LUNA ===
function MoonPos {
    param([double]$JD, $sun)
    $T  = ($JD - 2451545.0) / 36525.0
    $T2 = $T * $T; $T3 = $T2 * $T; $T4 = $T3 * $T
    $Lp = N360 (218.3164477 + 481267.88123421 * $T - 0.0015786 * $T2 + $T3 / 538841.0 - $T4 / 65194000.0)
    $D  = N360 (297.8501921 + 445267.1114034  * $T - 0.0018819 * $T2 + $T3 / 545868.0 - $T4 / 113065000.0)
    $M  = N360 (357.5291092 + 35999.0502909   * $T - 0.0001536 * $T2 + $T3 / 24490000.0)
    $Mp = N360 (134.9633964 + 477198.8675055  * $T + 0.0087414 * $T2 + $T3 / 69699.0   - $T4 / 14712000.0)
    $F  = N360 (93.2720950  + 483202.0175233  * $T - 0.0036539 * $T2 - $T3 / 3526000.0 + $T4 / 863310000.0)
    $E  = 1.0 - 0.002516 * $T - 0.0000074 * $T2
    $Dr=Rad $D; $Mr=Rad $M; $Mpr=Rad $Mp; $Fr=Rad $F
    $dL = 6288774.0*[Math]::Sin($Mpr) + 1274027.0*[Math]::Sin(2*$Dr-$Mpr) +
          658314.0*[Math]::Sin(2*$Dr) + 213618.0*[Math]::Sin(2*$Mpr) +
          (-185116.0*$E)*[Math]::Sin($Mr) + (-114332.0)*[Math]::Sin(2*$Fr) +
          58793.0*[Math]::Sin(2*$Dr-2*$Mpr) +
          (57066.0*$E)*[Math]::Sin(2*$Dr-$Mr-$Mpr) +
          53322.0*[Math]::Sin(2*$Dr+$Mpr) + (45758.0*$E)*[Math]::Sin(2*$Dr-$Mr) +
          (-40923.0*$E)*[Math]::Sin($Mr-$Mpr) + (-34720.0)*[Math]::Sin($Dr) +
          (-30383.0*$E)*[Math]::Sin($Mr+$Mpr) + 15327.0*[Math]::Sin(2*$Dr-2*$Fr) +
          (-12528.0)*[Math]::Sin($Mpr+2*$Fr) + 10980.0*[Math]::Sin($Mpr-2*$Fr) +
          10675.0*[Math]::Sin(4*$Dr-$Mpr) + 10034.0*[Math]::Sin(3*$Mpr) +
          8548.0*[Math]::Sin(4*$Dr-2*$Mpr) + (-7888.0*$E)*[Math]::Sin(2*$Dr+$Mr-$Mpr)
    $dR = -20905355.0*[Math]::Cos($Mpr) - 3699111.0*[Math]::Cos(2*$Dr-$Mpr) -
          2955968.0*[Math]::Cos(2*$Dr) - 569925.0*[Math]::Cos(2*$Mpr) +
          48888.0*$E*[Math]::Cos($Mr) - 3149.0*[Math]::Cos(2*$Fr) +
          246158.0*[Math]::Cos(2*$Dr-2*$Mpr) + (-152138.0*$E)*[Math]::Cos(2*$Dr-$Mr-$Mpr) +
          (-170733.0)*[Math]::Cos(2*$Dr+$Mpr) + (-204586.0*$E)*[Math]::Cos(2*$Dr-$Mr) +
          (-129620.0*$E)*[Math]::Cos($Mr-$Mpr) + 108743.0*[Math]::Cos($Dr) +
          104755.0*$E*[Math]::Cos($Mr+$Mpr) + 10321.0*[Math]::Cos(2*$Dr-2*$Fr)
    $dB = 5128122.0*[Math]::Sin($Fr) + 280602.0*[Math]::Sin($Mpr+$Fr) +
          277693.0*[Math]::Sin($Mpr-$Fr) + 173237.0*[Math]::Sin(2*$Dr-$Fr) +
          55413.0*[Math]::Sin(2*$Dr-$Mpr+$Fr) + 46271.0*[Math]::Sin(2*$Dr-$Mpr-$Fr) +
          32573.0*[Math]::Sin(2*$Dr+$Fr) + 17198.0*[Math]::Sin(2*$Mpr+$Fr) +
          9266.0*[Math]::Sin(2*$Dr+$Mpr-$Fr) + 8822.0*[Math]::Sin(2*$Mpr-$Fr) +
          (8216.0*$E)*[Math]::Sin(2*$Dr-$Mr-$Fr) + 4324.0*[Math]::Sin(2*$Dr-2*$Mpr-$Fr) +
          4200.0*[Math]::Sin(2*$Dr+$Mpr+$Fr) + (-3359.0*$E)*[Math]::Sin(2*$Dr+$Mr-$Fr)
    $moonLon = N360 ($Lp + $dL / 1000000.0)
    $moonLat = $dB / 1000000.0
    $moonDist = 385000.56 + $dR / 1000.0
    $eps    = Obliq $T
    $eq     = Ecl2Equ $moonLon $moonLat $eps
    $psi    = N360 ($moonLon - $sun.EL)
    $psi    = [Math]::Abs($psi)
    if ($psi -gt 180.0) { $psi = 360.0 - $psi }
    $mag    = -12.74 + 0.026 * $psi + 4e-9 * $psi * $psi * $psi * $psi
    return @{ RA = $eq.RA; Dec = $eq.Dec; Dist = $moonDist; Mag = [Math]::Round($mag,1) }
}
#endregion

#region === POSIZIONE PIANETI (Meeus AFC Tavola 23.A) ===
function PlanetPos {
    param([string]$name, [double]$JD, $sun, [double]$eps)
    $T1 = ($JD - 2415020.0) / 36525.0   # T dal 1900 Jan 0.5
    switch ($name) {
        "Mercurio" {
            $L = N360 (178.179078  + 149474.07078  * $T1 + 0.0003011  * $T1*$T1)
            $a = 0.3870986; $e = 0.20561421 + 0.00002046*$T1
            $i = N360 (7.002881   + 0.0018608   * $T1 - 0.0000183  * $T1*$T1)
            $Om= N360 (47.145944  + 1.1852083   * $T1 + 0.0001739  * $T1*$T1)
            $w = N360 (28.753753  + 0.3702806   * $T1 + 0.0001208  * $T1*$T1)
        }
        "Venere" {
            $L = N360 (342.767053 + 58519.21191  * $T1 + 0.0003097  * $T1*$T1)
            $a = 0.7233316; $e = 0.00682069 - 0.00004774*$T1
            $i = N360 (3.393631   + 0.0010058   * $T1 - 0.0000010  * $T1*$T1)
            $Om= N360 (75.779647  + 0.8998500   * $T1 + 0.0004100  * $T1*$T1)
            $w = N360 (54.384186  + 0.5081861   * $T1 - 0.0013864  * $T1*$T1)
        }
        "Marte" {
            $L = N360 (293.737334 + 19141.69551  * $T1 + 0.0003107  * $T1*$T1)
            $a = 1.5236883; $e = 0.09331290 + 0.000092064*$T1
            $i = N360 (1.850333   - 0.0006750   * $T1 + 0.0000126  * $T1*$T1)
            $Om= N360 (48.786442  + 0.7709917   * $T1 - 0.0000014  * $T1*$T1)
            $w = N360 (285.431761 + 1.0697667   * $T1 + 0.0001313  * $T1*$T1)
        }
        "Giove" {
            $L = N360 (238.049257 + 3036.301986  * $T1 + 0.0003347  * $T1*$T1 - 0.00000165*$T1*$T1*$T1)
            $a = 5.202561;  $e = 0.04833475 + 0.000164180*$T1 - 0.0000004676*$T1*$T1
            $i = N360 (1.308736   - 0.0056961   * $T1 + 0.0000039  * $T1*$T1)
            $Om= N360 (99.443414  + 1.0105300   * $T1 + 0.0003512  * $T1*$T1)
            $w = N360 (273.277558 + 0.5994317   * $T1 + 0.0007051  * $T1*$T1)
        }
        "Saturno" {
            $L = N360 (266.564377 + 1223.509884  * $T1 + 0.0003245  * $T1*$T1 - 0.0000058*$T1*$T1*$T1)
            $a = 9.554747;  $e = 0.05589232 - 0.000345700*$T1 - 0.0000007280*$T1*$T1
            $i = N360 (2.492519   - 0.0039189   * $T1 - 0.0000015  * $T1*$T1)
            $Om= N360 (112.790414 + 0.8731951   * $T1 - 0.0001521  * $T1*$T1)
            $w = N360 (338.307800 + 1.0852207   * $T1 + 0.0009454  * $T1*$T1)
        }
    }
    $M_pl    = N360 ($L - $w)
    $E_rad   = SolveKepler $M_pl $e
    $V       = TrueAnomaly $E_rad $e
    $r       = $a * (1.0 - $e * $e) / (1.0 + $e * [Math]::Cos($E_rad))
    $Omr     = Rad $Om
    $ir      = Rad $i
    $u       = Rad (N360 ($V + $w - $Om))
    $sinOm   = [Math]::Sin($Omr); $cosOm = [Math]::Cos($Omr)
    $sinU    = [Math]::Sin($u);   $cosU  = [Math]::Cos($u)
    $sinI    = [Math]::Sin($ir);  $cosI  = [Math]::Cos($ir)
    $xh = $r * ($cosOm * $cosU - $sinOm * $sinU * $cosI)
    $yh = $r * ($sinOm * $cosU + $cosOm * $sinU * $cosI)
    $zh = $r * $sinU * $sinI

    # Terra eliocentrica (da Sole geocentrico)
    $T2k = ($JD - 2451545.0) / 36525.0
    $Ls   = Rad (N360 (280.46646 + 36000.76983*$T2k))
    $Ms   = Rad (N360 (357.52911 + 35999.05029*$T2k - 0.0001537*$T2k*$T2k))
    $Cs   = (1.914602 - 0.004817*$T2k)*[Math]::Sin($Ms) + 0.019993*[Math]::Sin(2*$Ms)
    $thetaS = N360 ((Deg $Ls) + $Cs)
    $nusR   = Rad (N360 ((Deg $Ms) + $Cs))
    $es     = 0.016708634 - 0.000042037*$T2k
    $Rs     = 1.000001018 * (1.0 - $es*$es) / (1.0 + $es*[Math]::Cos($nusR))
    $thetaSr = Rad $thetaS
    $xe = -$Rs * [Math]::Cos($thetaSr)
    $ye = -$Rs * [Math]::Sin($thetaSr)
    $ze = 0.0

    $dx = $xh - $xe; $dy = $yh - $ye; $dz = $zh - $ze
    $Delta = [Math]::Sqrt($dx*$dx + $dy*$dy + $dz*$dz)
    $lamP  = Deg ([Math]::Atan2($dy, $dx))
    $betP  = Deg ([Math]::Atan2($dz, [Math]::Sqrt($dx*$dx + $dy*$dy)))
    $eq    = Ecl2Equ (N360 $lamP) $betP $eps

    # Angolo di fase e magnitudine
    $cosPhase = ($r*$r + $Delta*$Delta - $Rs*$Rs) / (2.0*$r*$Delta)
    if ($cosPhase -lt -1.0) { $cosPhase = -1.0 } elseif ($cosPhase -gt 1.0) { $cosPhase = 1.0 }
    $phi = Deg ([Math]::Acos($cosPhase))
    $mag = switch ($name) {
        "Mercurio" { -0.42 + 5*[Math]::Log10($r*$Delta) + 0.0380*$phi - 0.000273*$phi*$phi + 0.000002*$phi*$phi*$phi }
        "Venere"   { -4.40 + 5*[Math]::Log10($r*$Delta) + 0.0009*$phi + 0.000239*$phi*$phi - 0.00000065*$phi*$phi*$phi }
        "Marte"    { -1.52 + 5*[Math]::Log10($r*$Delta) + 0.016*$phi }
        "Giove"    { -9.40 + 5*[Math]::Log10($r*$Delta) + 0.005*$phi }
        "Saturno"  { -8.88 + 5*[Math]::Log10($r*$Delta) }
    }
    return @{ RA = $eq.RA; Dec = $eq.Dec; Dist = $Delta; Mag = [Math]::Round($mag,1) }
}
#endregion

#region === CALCOLO PRINCIPALE ===
function CalcolaAstri {
    param([double]$lat,[double]$lon,[int]$Y,[int]$Mo,[int]$D,[double]$UT)
    $JD  = JulianDate $Y $Mo $D $UT
    $T2k = ($JD - 2451545.0) / 36525.0
    $GST_deg = GST $JD
    $LST_deg = N360 ($GST_deg + $lon)
    $eps = Obliq $T2k
    $sun  = SunPos $JD
    $moon = MoonPos $JD $sun
    $results = [System.Collections.Generic.List[hashtable]]::new()
    # Sole
    $aas = AltAz $sun.RA $sun.Dec $LST_deg $lat
    $results.Add(@{ Nome="Sole"; RA=$sun.RA; Dec=$sun.Dec; Mag=$sun.Mag; Az=$aas.Az; Alt=$aas.Alt; Dist=("{0:0.000000} UA" -f $sun.R) })
    # Luna
    $aam = AltAz $moon.RA $moon.Dec $LST_deg $lat
    $results.Add(@{ Nome="Luna"; RA=$moon.RA; Dec=$moon.Dec; Mag=$moon.Mag; Az=$aam.Az; Alt=$aam.Alt; Dist=("{0:0} km" -f $moon.Dist) })
    # Pianeti
    foreach ($pname in @("Mercurio","Venere","Marte","Giove","Saturno")) {
        $pl  = PlanetPos $pname $JD $sun $eps
        $aap = AltAz $pl.RA $pl.Dec $LST_deg $lat
        $results.Add(@{ Nome=$pname; RA=$pl.RA; Dec=$pl.Dec; Mag=$pl.Mag; Az=$aap.Az; Alt=$aap.Alt; Dist=("{0:0.000} UA" -f $pl.Dist) })
    }
    return @{ Rows=$results; JD=$JD; T=$T2k; GST=$GST_deg; LST=$LST_deg }
}
#endregion

#region === COLORI ===
$BG      = [Drawing.Color]::FromArgb(13,13,30)
$PANEL   = [Drawing.Color]::FromArgb(20,20,45)
$BORDER  = [Drawing.Color]::FromArgb(60,80,140)
$FG      = [Drawing.Color]::White
$ACCENT  = [Drawing.Color]::FromArgb(100,160,255)
$GRID_BG = [Drawing.Color]::FromArgb(18,18,40)
$GRID_FG = [Drawing.Color]::White
$GRID_HD = [Drawing.Color]::FromArgb(30,30,60)
$ABOVE   = [Drawing.Color]::FromArgb(30,180,80)
$BELOW   = [Drawing.Color]::FromArgb(90,90,100)

$BODY_COLORS = @{
    "Sole"     = [Drawing.Color]::FromArgb(255,215,0)
    "Luna"     = [Drawing.Color]::FromArgb(210,210,210)
    "Mercurio" = [Drawing.Color]::FromArgb(200,180,140)
    "Venere"   = [Drawing.Color]::FromArgb(240,230,180)
    "Marte"    = [Drawing.Color]::FromArgb(230,100,60)
    "Giove"    = [Drawing.Color]::FromArgb(230,200,160)
    "Saturno"  = [Drawing.Color]::FromArgb(210,195,130)
}
#endregion

#region === FORM ===
$form = New-Object System.Windows.Forms.Form
$form.Text            = "AstroCalc - Calcolatore Astronomico"
$form.Size            = New-Object Drawing.Size(980, 600)
$form.StartPosition   = "CenterScreen"
$form.BackColor       = $BG
$form.ForeColor       = $FG
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox     = $false
$form.Font            = New-Object Drawing.Font("Consolas", 9)

# Titolo
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text      = "*** ASTRO CALC - Effemeridi Planetarie ***"
$lblTitle.Font      = New-Object Drawing.Font("Consolas", 12, [Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = $ACCENT
$lblTitle.Location  = New-Object Drawing.Point(10, 10)
$lblTitle.Size      = New-Object Drawing.Size(950, 25)
$lblTitle.TextAlign = "MiddleCenter"
$form.Controls.Add($lblTitle)

# Pannello input
$pnlInput = New-Object System.Windows.Forms.Panel
$pnlInput.Location  = New-Object Drawing.Point(10, 42)
$pnlInput.Size      = New-Object Drawing.Size(955, 60)
$pnlInput.BackColor = $PANEL
$form.Controls.Add($pnlInput)

function MkLabel { param($txt,$x,$y,$w)
    $l = New-Object System.Windows.Forms.Label
    $l.Text=$txt; $l.ForeColor=$ACCENT
    $l.Location=New-Object Drawing.Point($x,$y)
    $l.Size=New-Object Drawing.Size($w,20)
    return $l
}
function MkBox { param($txt,$x,$y,$w)
    $b = New-Object System.Windows.Forms.TextBox
    $b.Text=$txt; $b.BackColor=[Drawing.Color]::FromArgb(30,30,55)
    $b.ForeColor=$FG; $b.BorderStyle="FixedSingle"
    $b.Location=New-Object Drawing.Point($x,$y)
    $b.Size=New-Object Drawing.Size($w,22)
    return $b
}

$pnlInput.Controls.Add((MkLabel "Lat N:"     8  8  48))
$txtLat  = MkBox "45.4654"  58  8  80; $pnlInput.Controls.Add($txtLat)
$pnlInput.Controls.Add((MkLabel "Lon E:"    150  8  48))
$txtLon  = MkBox "9.1859"  200  8  80; $pnlInput.Controls.Add($txtLon)
$pnlInput.Controls.Add((MkLabel "Data UT:"  295  8  65))
$txtDate = MkBox (Get-Date -Format "dd/MM/yyyy")  365  8  90; $pnlInput.Controls.Add($txtDate)
$pnlInput.Controls.Add((MkLabel "Ora UT:"   465  8  60))
$txtTime = MkBox (Get-Date -UFormat "%H:%M:%S")   530  8  80; $pnlInput.Controls.Add($txtTime)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text      = "Default: Milano  Lat=45.4654 N  Lon=9.1859 E  |  Data e Ora in UTC"
$lblHint.ForeColor = [Drawing.Color]::FromArgb(90,110,160)
$lblHint.Location  = New-Object Drawing.Point(8, 36)
$lblHint.Size      = New-Object Drawing.Size(620, 18)
$lblHint.Font      = New-Object Drawing.Font("Consolas", 8)
$pnlInput.Controls.Add($lblHint)

$btnCalc = New-Object System.Windows.Forms.Button
$btnCalc.Text      = "[ CALCOLA ]"
$btnCalc.Font      = New-Object Drawing.Font("Consolas",10,[Drawing.FontStyle]::Bold)
$btnCalc.BackColor = [Drawing.Color]::FromArgb(30,60,120)
$btnCalc.ForeColor = $FG
$btnCalc.FlatStyle = "Flat"
$btnCalc.Location  = New-Object Drawing.Point(650, 10)
$btnCalc.Size      = New-Object Drawing.Size(140, 40)
$pnlInput.Controls.Add($btnCalc)

$btnNow = New-Object System.Windows.Forms.Button
$btnNow.Text      = "[ ADESSO ]"
$btnNow.Font      = New-Object Drawing.Font("Consolas",9)
$btnNow.BackColor = [Drawing.Color]::FromArgb(20,50,40)
$btnNow.ForeColor = $FG
$btnNow.FlatStyle = "Flat"
$btnNow.Location  = New-Object Drawing.Point(800, 10)
$btnNow.Size      = New-Object Drawing.Size(110, 40)
$pnlInput.Controls.Add($btnNow)

# Barra stato / info JD
$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text      = "In attesa di calcolo..."
$lblInfo.ForeColor = [Drawing.Color]::FromArgb(150,180,220)
$lblInfo.BackColor = [Drawing.Color]::FromArgb(10,10,25)
$lblInfo.Location  = New-Object Drawing.Point(10, 108)
$lblInfo.Size      = New-Object Drawing.Size(955, 20)
$lblInfo.Font      = New-Object Drawing.Font("Consolas",8)
$form.Controls.Add($lblInfo)

# DataGridView
$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Location  = New-Object Drawing.Point(10, 133)
$dgv.Size      = New-Object Drawing.Size(955, 415)
$dgv.BackgroundColor = $GRID_BG
$dgv.DefaultCellStyle.BackColor       = $GRID_BG
$dgv.DefaultCellStyle.ForeColor       = $GRID_FG
$dgv.DefaultCellStyle.Font            = New-Object Drawing.Font("Consolas", 10)
$dgv.DefaultCellStyle.SelectionBackColor = [Drawing.Color]::FromArgb(40,40,70)
$dgv.DefaultCellStyle.SelectionForeColor = $GRID_FG
$dgv.AlternatingRowsDefaultCellStyle.BackColor = [Drawing.Color]::FromArgb(22,22,48)
$dgv.AlternatingRowsDefaultCellStyle.SelectionBackColor = [Drawing.Color]::FromArgb(40,40,70)
$dgv.AlternatingRowsDefaultCellStyle.SelectionForeColor = $GRID_FG
$dgv.ColumnHeadersDefaultCellStyle.BackColor = $GRID_HD
$dgv.ColumnHeadersDefaultCellStyle.ForeColor = $ACCENT
$dgv.ColumnHeadersDefaultCellStyle.Font      = New-Object Drawing.Font("Consolas",10,[Drawing.FontStyle]::Bold)
$dgv.ColumnHeadersDefaultCellStyle.SelectionBackColor = $GRID_HD
$dgv.ColumnHeadersDefaultCellStyle.SelectionForeColor = $ACCENT
$dgv.GridColor            = $BORDER
$dgv.RowHeadersVisible    = $false
$dgv.AllowUserToAddRows   = $false
$dgv.ReadOnly             = $true
$dgv.SelectionMode        = "FullRowSelect"
$dgv.EnableHeadersVisualStyles = $false
$dgv.ColumnHeadersHeightSizeMode = "DisableResizing"
$dgv.ColumnHeadersHeight  = 30
$dgv.RowTemplate.Height   = 52
$dgv.BorderStyle          = "None"
$form.Controls.Add($dgv)

# Colonne
$MC = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$cols = @(
    @{N="Corpo";       W=110; A=$MC},
    @{N="Asc. Retta";  W=160; A=$MC},
    @{N="Declinaz.";   W=160; A=$MC},
    @{N="Magnit.";     W=80;  A=$MC},
    @{N="Azimut";      W=100; A=$MC},
    @{N="Altitudine";  W=110; A=$MC},
    @{N="Distanza";    W=130; A=$MC}
)
foreach ($c in $cols) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.HeaderText = $c.N
    $col.Width      = $c.W
    $col.DefaultCellStyle.Alignment = $c.A
    $col.SortMode   = "NotSortable"
    $dgv.Columns.Add($col) | Out-Null
}

# Footer
$lblFoot = New-Object System.Windows.Forms.Label
$lblFoot.Text      = "Algoritmi: Jean Meeus - Astronomical Algorithms / Astronomical Formulae for Calculators"
$lblFoot.ForeColor = [Drawing.Color]::FromArgb(80,100,140)
$lblFoot.Location  = New-Object Drawing.Point(10, 553)
$lblFoot.Size      = New-Object Drawing.Size(955, 16)
$lblFoot.Font      = New-Object Drawing.Font("Consolas",7)
$lblFoot.TextAlign = "MiddleCenter"
$form.Controls.Add($lblFoot)
#endregion

#region === LOGICA BOTTONI ===
function ParseInput {
    try {
        $lat = [double]::Parse($txtLat.Text)
        $lon = [double]::Parse($txtLon.Text)
        $dp  = $txtDate.Text.Split("/")
        $tp  = $txtTime.Text.Split(":")
        $D=[int]$dp[0]; $Mo=[int]$dp[1]; $Y=[int]$dp[2]
        $h=[int]$tp[0]; $m=[int]$tp[1]; $s=0
        if ($tp.Count -ge 3) { $s=[int]$tp[2] }
        $UT = $h + $m/60.0 + $s/3600.0
        return @{ OK=$true; Lat=$lat; Lon=$lon; Y=$Y; Mo=$Mo; D=$D; UT=$UT }
    } catch {
        return @{ OK=$false }
    }
}

function AggiornaGriglia {
    $p = ParseInput
    if (-not $p.OK) {
        $lblInfo.Text = "ERRORE: Verifica i valori inseriti (Data: gg/mm/aaaa  Ora: HH:mm:ss)"
        $lblInfo.ForeColor = [Drawing.Color]::Salmon
        return
    }
    $lblInfo.ForeColor = [Drawing.Color]::FromArgb(150,180,220)
    $lblInfo.Text = "Calcolo in corso..."
    $form.Refresh()
    try {
        $res = CalcolaAstri $p.Lat $p.Lon $p.Y $p.Mo $p.D $p.UT
        $DG    = [char]176
        $gst   = $res.GST / 15.0
        $gstH  = [int][Math]::Floor($gst)
        $gstM  = [int][Math]::Floor(($gst - $gstH) * 60.0)
        $lst   = $res.LST / 15.0
        $lstH  = [int][Math]::Floor($lst)
        $lstM  = [int][Math]::Floor(($lst - $lstH) * 60.0)
        $epsV  = [Math]::Round((Obliq $res.T), 4)
        $Tval  = [Math]::Round($res.T, 6)
        $JDval = [Math]::Round($res.JD, 4)
        $lblInfo.Text = ("JD={0}  T={1}  GST={2:00}h{3:00}m  LST={4:00}h{5:00}m  eps={6}{7}" `
            -f $JDval, $Tval, $gstH, $gstM, $lstH, $lstM, $epsV, $DG)
        $dgv.Rows.Clear()
        foreach ($row in $res.Rows) {
            $ri = $dgv.Rows.Add(
                $row.Nome,
                (FmtRA  $row.RA),
                (FmtDec $row.Dec),
                ("{0:+0.0;-0.0;0.0}" -f $row.Mag),
                (FmtAng $row.Az),
                (FmtAng $row.Alt),
                $row.Dist
            )
            $r  = $dgv.Rows[$ri]
            $bc = $BODY_COLORS[$row.Nome]
            $r.Cells[0].Style.ForeColor = $bc
            $r.Cells[0].Style.Font = New-Object Drawing.Font("Consolas",10,[Drawing.FontStyle]::Bold)
            if ($row.Alt -gt 0) {
                $r.Cells[5].Style.ForeColor  = $ABOVE
                $r.Cells[5].Style.Font       = New-Object Drawing.Font("Consolas",10,[Drawing.FontStyle]::Bold)
                $r.Cells[5].Style.BackColor  = [Drawing.Color]::FromArgb(15,50,25)
            } else {
                $r.Cells[5].Style.ForeColor = $BELOW
                $r.Cells[5].Style.BackColor = [Drawing.Color]::FromArgb(25,25,35)
            }
        }
    } catch {
        $lblInfo.Text = "ERRORE nel calcolo: $_"
        $lblInfo.ForeColor = [Drawing.Color]::Salmon
    }
}

$btnCalc.Add_Click({ AggiornaGriglia })
$btnNow.Add_Click({
    $now = [System.DateTime]::UtcNow
    $txtDate.Text = $now.ToString("dd/MM/yyyy")
    $txtTime.Text = $now.ToString("HH:mm:ss")
    AggiornaGriglia
})
$form.Add_Shown({ AggiornaGriglia })
#endregion

[System.Windows.Forms.Application]::Run($form)

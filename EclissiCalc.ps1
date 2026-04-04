#Requires -Version 5.1
# ============================================================
#  ECLISSI SOLARI & LUNARI  v4  -  PowerShell 5.1
#
#  SOLARI : Elementi Besseliani dal CSV (Espenak & Meeus,
#           Five Millennium Canon) -> circostanze locali esatte
#  LUNARI  : Meeus AFC Cap.33 + campionamento altitudine Luna
#
#  Algoritmo locale solare (cap. 9 Supplemento Astronomico):
#    x(t),y(t) = posizione asse ombra nel piano fondamentale
#    d(t),mu(t)= declinazione e GHA dell'asse del cono
#    xi,eta,zeta = coordinate geocentriche osservatore
#    Delta = distanza osservatore dall'asse del cono
#    Se Delta <= L1(t) : eclisse visibile
#    Magnitudine locale = (L1 - Delta) / (L1 + |L2|)
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ================================================================
#  REPORT HTML PER ESPORTAZIONE PDF
# ================================================================
function Build-HtmlReport {
    param([array]$Dati,[string]$Lat,[string]$Lon,[string]$TZ,[string]$AnnoI,[string]$AnnoF,[string]$GenOn)
    $nSol = @($Dati | Where-Object { $_.Tipo -eq '(*)'}).Count
    $nLun = @($Dati | Where-Object { $_.Tipo -eq '(O)'}).Count
    $nVis = @($Dati | Where-Object { $_.VisStr -notlike 'Non*' -and $_.VisStr -ne 'Fuori percorso'}).Count
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<!DOCTYPE html><html lang="it"><head><meta charset="UTF-8"><title>Eclissi</title><style>')
    [void]$sb.Append('@page{size:A4 landscape;margin:10mm 12mm 12mm 12mm}')
    [void]$sb.Append('*{box-sizing:border-box;margin:0;padding:0}')
    [void]$sb.Append('body{font-family:"Courier New",monospace;font-size:8pt;color:#1a1a2e;background:#fff}')
    [void]$sb.Append('.hdr{background:linear-gradient(135deg,#1a1a5e 0%,#2d2d9a 60%,#4040b0 100%);color:#fff;padding:8px 12px 6px;border-radius:5px;margin-bottom:8px}')
    [void]$sb.Append('.hdr h1{font-size:18pt;letter-spacing:5px;color:#ffd84a;margin-bottom:2px}')
    [void]$sb.Append('.hdr .sub{font-size:7pt;color:#c0b8ee;margin-bottom:3px}')
    [void]$sb.Append('.hdr .meta{font-size:7.5pt;color:#e0d8ff;border-top:1px solid rgba(255,255,255,0.2);padding-top:4px;margin-top:3px;display:flex;gap:12px;flex-wrap:wrap}')
    [void]$sb.Append('.hdr .meta span{background:rgba(255,255,255,0.1);padding:1px 7px;border-radius:3px}')
    [void]$sb.Append('table{width:100%;border-collapse:collapse;font-size:7.5pt}')
    [void]$sb.Append('thead{display:table-header-group}')
    [void]$sb.Append('tr{page-break-inside:avoid}')
    [void]$sb.Append('th{background:#2d2d9a;color:#fff;font-weight:bold;padding:4px 6px;text-align:center;border-bottom:2px solid #1a1a5e;white-space:nowrap;font-size:7pt}')
    [void]$sb.Append('td{padding:3px 6px;border-bottom:1px solid #ddd;white-space:nowrap;vertical-align:middle}')
    [void]$sb.Append('tr.alt td{background:#f0f0fa}')
    [void]$sb.Append('.sol{color:#8a5500;font-weight:bold}.lun{color:#1a1a8a;font-weight:bold}')
    [void]$sb.Append('.tot{color:#cc0000;font-weight:bold}.ann{color:#cc5500;font-weight:bold}')
    [void]$sb.Append('.par{color:#444}.pen{color:#888}.umb{color:#2255aa;font-weight:bold}')
    [void]$sb.Append('.vis{color:#006600;font-weight:bold}.nonvis{color:#bbb}')
    [void]$sb.Append('.c{text-align:center}.r{text-align:right}')
    [void]$sb.Append('.ftr{margin-top:6px;padding-top:4px;border-top:1px solid #ccc;font-size:6.5pt;color:#888;display:flex;justify-content:space-between}')
    [void]$sb.Append('</style></head><body>')
    [void]$sb.Append('<div class="hdr"><h1>ECLISSI &nbsp; SOLARI &nbsp;&amp;&nbsp; LUNARI</h1>')
    [void]$sb.Append('<div class="sub">Calcolo astronomico &mdash; Espenak &amp; Meeus, <em>Five Millennium Canon</em> &mdash; Meeus AFC Cap.33</div>')
    [void]$sb.Append('<div class="meta">')
    [void]$sb.Append("<span>&#128197; Periodo: <b>$AnnoI &ndash; $AnnoF</b></span>")
    [void]$sb.Append("<span>&#127759; Coord: <b>${Lat}&deg;N &nbsp;${Lon}&deg;E</b></span>")
    [void]$sb.Append("<span>&#128336; Fuso: <b>UTC+$TZ</b></span>")
    [void]$sb.Append("<span>Totale: <b>$($Dati.Count)</b> eclissi ($nSol sol., $nLun lun.)</span>")
    [void]$sb.Append("<span>Visibili: <b>$nVis</b></span>")
    [void]$sb.Append("</div></div>")
    [void]$sb.Append('<table><thead><tr>')
    foreach ($h in @('Data','T','Sottotipo','Magn.','1&deg;Cont.UT','Massimo UT','Ult.Cont.UT','Max Loc.','Inizio Vis.','Fine Vis.','Visib.%','Alt.(&deg;)','T.Parz(m)','T.Tot(m)','Nodo','Note')) {
        [void]$sb.Append("<th>$h</th>")
    }
    [void]$sb.Append('</tr></thead><tbody>')
    $alt = $false
    foreach ($r in $Dati) {
        $isSol = $r.Tipo -eq '(*)'
        $trCls = if ($isSol) {'sol'} else {'lun'}
        $subCls = if     ($r.SottoTipo -eq 'Totale')     {'tot'}
                  elseif ($r.SottoTipo -eq 'Anulare')    {'ann'}
                  elseif ($r.SottoTipo -eq 'Ibrida')     {'ann'}
                  elseif ($r.SottoTipo -eq 'Umbrale')    {'umb'}
                  elseif ($r.SottoTipo -eq 'Penombrale') {'pen'}
                  else                                    {'par'}
        $visCls = if ($r.VisStr -like 'Non*' -or $r.VisStr -eq 'Fuori percorso') {'nonvis'} else {'vis'}
        $altCls = if ($alt) {' class="alt"'} else {''}
        $alt = -not $alt
        [void]$sb.Append("<tr$altCls>")
        [void]$sb.Append("<td><b>$($r.DataStr)</b></td>")
        [void]$sb.Append("<td class='$trCls c'>$($r.Tipo)</td>")
        [void]$sb.Append("<td class='$subCls'>$($r.SottoTipo)</td>")
        [void]$sb.Append("<td class='r'>$($r.Magn)</td>")
        [void]$sb.Append("<td class='c'>$($r.UtP1)</td>")
        [void]$sb.Append("<td class='c'><b>$($r.UtMax)</b></td>")
        [void]$sb.Append("<td class='c'>$($r.UtP4)</td>")
        [void]$sb.Append("<td class='c'>$($r.LocMax)</td>")
        [void]$sb.Append("<td class='c $visCls'>$($r.InizioVis)</td>")
        [void]$sb.Append("<td class='c $visCls'>$($r.FineVis)</td>")
        [void]$sb.Append("<td class='c $visCls'>$($r.VisStr)</td>")
        [void]$sb.Append("<td class='r'>$($r.AltMax)</td>")
        [void]$sb.Append("<td class='c'>$($r.DurParz)</td>")
        [void]$sb.Append("<td class='c'>$($r.DurTot)</td>")
        [void]$sb.Append("<td class='c'>$($r.NodoStr)</td>")
        [void]$sb.Append("<td style='color:#999;font-size:6.5pt'>$($r.Flag)</td>")
        [void]$sb.Append("</tr>")
    }
    [void]$sb.Append('</tbody></table>')
    [void]$sb.Append('<div class="ftr">')
    [void]$sb.Append('<span>(*) Solare: Besseliani Espenak &amp; Meeus | Magn.=locale piano fondamentale | (O) Lunare: Meeus AFC Cap.33 | Visib.%=fraz. eclisse con corpo&gt;3&deg; | Alt.=gradi al massimo | T.Parz/Tot=min durata fase</span>')
    [void]$sb.Append("<span>Generato: $GenOn</span>")
    [void]$sb.Append('</div></body></html>')
    return $sb.ToString()
}

# ---- Export PDF via Edge/Chrome headless (come FasiLunari.ps1) ----
function Export-EclissiPdf {
    param([array]$Dati,[string]$OutPath,[string]$Lat,[string]$Lon,[string]$TZ,[string]$AnnoI,[string]$AnnoF)

    # 1. Cerca Edge o Chrome
    $browsers = @(
        'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
        'C:\Program Files\Microsoft\Edge\Application\msedge.exe'
        "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe"
        'C:\Program Files\Google\Chrome\Application\chrome.exe'
        'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    )
    $browser = $browsers | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $browser) {
        throw "Microsoft Edge o Google Chrome non trovati.`nInstallare uno dei due browser per esportare in PDF."
    }

    # 2. Genera HTML in file temporaneo
    $genOn  = [DateTime]::Now.ToString('dd/MM/yyyy HH:mm')
    $html   = Build-HtmlReport $Dati $Lat $Lon $TZ $AnnoI $AnnoF $genOn
    $tmpHtml = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(),
               "EclissiCalc_$([System.Guid]::NewGuid().ToString('N')).html")
    [System.IO.File]::WriteAllText($tmpHtml, $html, [System.Text.Encoding]::UTF8)

    # 3. Percorsi con slash (richiesti da Chromium)
    $absOut  = [System.IO.Path]::GetFullPath($OutPath).Replace('\','/')
    $fileUrl = "file:///$($tmpHtml.Replace('\','/'))"

    # 4. Avvia browser headless
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName    = $browser
    $psi.Arguments   = "--headless=new --disable-gpu --run-all-compositor-stages-before-draw --no-sandbox --print-to-pdf=`"$absOut`" --print-to-pdf-no-header --no-pdf-header-footer `"$fileUrl`""
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true

    $proc = [System.Diagnostics.Process]::Start($psi)
    $ok   = $proc.WaitForExit(30000)
    if (-not $ok) { $proc.Kill(); throw "Timeout: il browser non ha risposto entro 30 secondi." }
    Start-Sleep -Milliseconds 800

    # 5. Cleanup HTML temporaneo
    Remove-Item $tmpHtml -ErrorAction SilentlyContinue

    # 6. Verifica file PDF
    if (-not (Test-Path $OutPath)) {
        throw "Il file PDF non e' stato creato (ExitCode=$($proc.ExitCode))."
    }
    $sz = (Get-Item $OutPath).Length
    if ($sz -lt 1000) { throw "PDF vuoto o corrotto ($sz byte)." }
}

# ================================================================
#  CALCOLI ASTRONOMICI IN C# (compilato a runtime)
#  Elimina l'overhead delle chiamate a funzione PS (~5-15ms ciascuna)
#  per le operazioni critiche: Besseliani e altitudine Sole/Luna
# ================================================================
Add-Type @"
using System;
public static class AstroCalc {
    const double D2R = Math.PI / 180.0;
    const double R2D = 180.0 / Math.PI;

    static double NA(double x) { x = x % 360.0; return x < 0 ? x + 360.0 : x; }

    // ------------------------------------------------------------------
    //  Besselian element evaluation
    //  Returns [Delta, L1, L2]
    // ------------------------------------------------------------------
    public static double[] BessEval(
            double x0, double x1, double x2, double x3,
            double y0, double y1, double y2, double y3,
            double d0, double d1, double d2,
            double mu0, double mu1, double mu2,
            double l10, double l11, double l12,
            double l20, double l21, double l22,
            double tanF1, double tanF2,
            double t, double sinPhi, double cosPhi, double Lon) {
        double t2 = t*t, t3 = t2*t;
        double bx  = x0  + x1*t  + x2*t2  + x3*t3;
        double by  = y0  + y1*t  + y2*t2  + y3*t3;
        double bd  = (d0  + d1*t  + d2*t2) * D2R;
        double bmu = mu0 + mu1*t + mu2*t2;
        double bl1 = l10 + l11*t + l12*t2;
        double bl2 = l20 + l21*t + l22*t2;
        double tR  = (bmu - Lon) * D2R;
        double sinT = Math.Sin(tR), cosT = Math.Cos(tR);
        double sinD = Math.Sin(bd), cosD = Math.Cos(bd);
        double xi   = cosPhi * sinT;
        double eta  = sinPhi*cosD - cosPhi*cosT*sinD;
        double zeta = sinPhi*sinD + cosPhi*cosT*cosD;
        double u = bx - xi, v = by - eta;
        double L1 = bl1 - zeta*tanF1;
        double L2 = bl2 - zeta*tanF2;
        return new double[] { Math.Sqrt(u*u + v*v), L1, L2 };
    }

    // ------------------------------------------------------------------
    //  Greenwich Sidereal Time (degrees)
    // ------------------------------------------------------------------
    public static double GST(double JD) {
        double T = (JD - 2451545.0) / 36525.0;
        double g = 280.46061837 + 360.98564736629*(JD-2451545.0)
                   + 0.000387933*T*T - T*T*T/38710000.0;
        return NA(g);
    }

    // ------------------------------------------------------------------
    //  Posizione Sole (RA, Dec in gradi)
    // ------------------------------------------------------------------
    public static double[] SunPos(double JD) {
        double T  = (JD-2451545.0)/36525.0, T2=T*T;
        double L0 = NA(280.46646+36000.76983*T+0.0003032*T2);
        double M  = NA(357.52911+35999.05029*T-0.0001537*T2);
        double Mr = M*D2R;
        double C  = (1.914602-0.004817*T-0.000014*T2)*Math.Sin(Mr)
                   +(0.019993-0.000101*T)*Math.Sin(2*Mr)
                   +0.000289*Math.Sin(3*Mr);
        double theta = NA(L0+C);
        double om = 125.04-1934.136*T, omR = om*D2R;
        double lam = NA(theta - 0.00569 - 0.00478*Math.Sin(omR));
        double eps = (23.439291111-0.013004167*T+0.00256*Math.Cos(omR))*D2R;
        double lamR = lam*D2R;
        double RA  = NA(Math.Atan2(Math.Sin(lamR)*Math.Cos(eps), Math.Cos(lamR))*R2D);
        double Dec = Math.Asin(Math.Sin(eps)*Math.Sin(lamR))*R2D;
        return new double[] { RA, Dec };
    }

    // ------------------------------------------------------------------
    //  Posizione Luna (RA, Dec in gradi) - termini principali
    // ------------------------------------------------------------------
    public static double[] MoonPos(double JD) {
        double T=((JD-2451545.0)/36525.0), T2=T*T, T3=T2*T, T4=T3*T;
        double Lp=NA(218.3164477+481267.88123421*T-0.0015786*T2+T3/538841.0-T4/65194000.0);
        double D =NA(297.8501921+445267.1114034*T-0.0018819*T2+T3/545868.0-T4/113065000.0);
        double M =NA(357.5291092+35999.0502909*T-0.0001536*T2);
        double Mp=NA(134.9633964+477198.8675055*T+0.0087414*T2+T3/69699.0-T4/14712000.0);
        double F =NA(93.2720950+483202.0175233*T-0.0036539*T2-T3/3526000.0+T4/863310000.0);
        double E =1.0-0.002516*T-0.0000074*T2;
        double Dr=D*D2R, Mr2=M*D2R, Mpr=Mp*D2R, Fr=F*D2R;
        double dL=6288774.0*Math.Sin(Mpr)+1274027.0*Math.Sin(2*Dr-Mpr)
                 +658314.0*Math.Sin(2*Dr)+213618.0*Math.Sin(2*Mpr)
                 -185116.0*E*Math.Sin(Mr2)-114332.0*Math.Sin(2*Fr)
                 +58793.0*Math.Sin(2*Dr-2*Mpr)+57066.0*E*Math.Sin(2*Dr-Mr2-Mpr)
                 +53322.0*Math.Sin(2*Dr+Mpr)+45758.0*E*Math.Sin(2*Dr-Mr2);
        double dB=5128122.0*Math.Sin(Fr)+280602.0*Math.Sin(Mpr+Fr)
                 +277693.0*Math.Sin(Mpr-Fr)+173237.0*Math.Sin(2*Dr-Fr)
                 +55413.0*Math.Sin(2*Dr-Mpr+Fr)+46271.0*Math.Sin(2*Dr-Mpr-Fr);
        double mlon=NA(Lp+dL/1000000.0), mlat=dB/1000000.0;
        double eps=(23.439291111-0.013004167*T)*D2R;
        double lonR=mlon*D2R, latR=mlat*D2R;
        double RA =NA(Math.Atan2(Math.Sin(lonR)*Math.Cos(eps)-Math.Tan(latR)*Math.Sin(eps),
                                  Math.Cos(lonR))*R2D);
        double Dec=Math.Asin(Math.Sin(latR)*Math.Cos(eps)+Math.Cos(latR)*Math.Sin(eps)*Math.Sin(lonR))*R2D;
        return new double[] { RA, Dec };
    }

    // ------------------------------------------------------------------
    //  Altitudine corpo celeste (gradi) - isSolar: 1=Sole, 0=Luna
    // ------------------------------------------------------------------
    public static double Altitude(double JDE, double Lat, double Lon, int isSolar) {
        double lst = NA(GST(JDE) + Lon);
        double[] pos = isSolar == 1 ? SunPos(JDE) : MoonPos(JDE);
        double ha = NA(lst - pos[0]);
        double haR = ha*D2R, decR = pos[1]*D2R, latR = Lat*D2R;
        double sinA = Math.Sin(decR)*Math.Sin(latR)
                     +Math.Cos(decR)*Math.Cos(latR)*Math.Cos(haR);
        sinA = Math.Max(-1.0, Math.Min(1.0, sinA));
        return Math.Asin(sinA)*R2D;
    }
}
"@

# Cultura invariante per parsing numeri dal CSV (evita problemi separatore decimale)
$script:IC = [System.Globalization.CultureInfo]::InvariantCulture
function PD { param([string]$s); return [double]::Parse($s.Trim(), $script:IC) }

# Stopwatch globale per timing e log live
$script:SW = [System.Diagnostics.Stopwatch]::new()

# Log immediato (non throttled) - usare solo per messaggi importanti
function LogF {
    param([string]$msg)
    $ms = $script:SW.ElapsedMilliseconds
    $ts = '{0:D2}:{1:D2}:{2:D2}.{3:D3}' -f `
          [int][Math]::Floor($ms/3600000), `
          [int][Math]::Floor(($ms%3600000)/60000), `
          [int][Math]::Floor(($ms%60000)/1000), ($ms%1000)
    $txt = "[$ts]  $msg"
    $script:LogLines.Add($txt)
    $txtLog.AppendText($txt + "`r`n")
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
    $script:LastUIms = $ms
}

$script:LogLines    = [System.Collections.Generic.List[string]]::new()
$script:LastUIms    = 0    # timestamp ultimo aggiornamento UI

function Log {
    param([string]$msg)
    $ms = $script:SW.ElapsedMilliseconds
    $ts = '{0:D2}:{1:D2}:{2:D2}.{3:D3}' -f `
          [int][Math]::Floor($ms/3600000), `
          [int][Math]::Floor(($ms%3600000)/60000), `
          [int][Math]::Floor(($ms%60000)/1000), ($ms%1000)
    $txt = "[$ts]  $msg"
    $script:LogLines.Add($txt)

    # AppendText e' 10x piu' veloce di Lines= (non riscrive tutto il testo)
    # DoEvents solo ogni 80ms: evita freeze da repaint in loop stretti
    $msSince = $ms - $script:LastUIms
    if ($msSince -ge 80) {
        $txtLog.AppendText($txt + "`r`n")
        $txtLog.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
        $script:LastUIms = $ms
    }
}

# ================================================================
#  MATEMATICA DI BASE
# ================================================================
function Rad  { param([double]$d); return $d * [Math]::PI / 180.0 }
function Deg  { param([double]$r); return $r * 180.0 / [Math]::PI }
function NA   { param([double]$x); $x = $x % 360.0; if ($x -lt 0) { $x += 360.0 }; return $x }

# ================================================================
#  JDE -> data/ora UT
# ================================================================
function JDE-ToUTC {
    param([double]$JD)
    $tmp = $JD + 0.5; $Z = [Math]::Floor($tmp); $F = $tmp - $Z
    if ($Z -lt 2299161) { $A = $Z }
    else {
        $al = [Math]::Floor(($Z - 1867216.25) / 36524.25)
        $A  = $Z + 1 + $al - [Math]::Floor($al / 4.0)
    }
    $B = $A + 1524; $C = [Math]::Floor(($B - 122.1) / 365.25)
    $Dd = [Math]::Floor(365.25 * $C); $E = [Math]::Floor(($B - $Dd) / 30.6001)
    $df = $B - $Dd - [Math]::Floor(30.6001 * $E) + $F
    $day = [Math]::Floor($df)
    $hf = ($df - $day) * 24.0; $h = [Math]::Floor($hf)
    $m  = [Math]::Round(($hf - $h) * 60.0)
    if ($m -eq 60) { $h++; $m = 0 }
    if ($h -eq 24) { $h = 0; $day++ }
    if ($E -lt 14) { $month = [int]($E - 1) } else { $month = [int]($E - 13) }
    if ($month -gt 2) { $year = [int]($C - 4716) } else { $year = [int]($C - 4715) }
    return [PSCustomObject]@{ Year=[int]$year; Month=[int]$month; Day=[int]$day; Hour=[int]$h; Minute=[int]$m }
}

function JDE-ToHHMM {
    param([double]$JDE, [double]$TZ)
    $dt = JDE-ToUTC $JDE
    $tm = $dt.Hour * 60 + $dt.Minute + [int]($TZ * 60)
    $tm = (($tm % 1440) + 1440) % 1440
    $hh = [int][Math]::Floor($tm / 60); $mm = $tm % 60
    return ('{0:D2}:{1:D2}' -f $hh, $mm)
}

# ================================================================
#  GST, ALTITUDINE, SOLE, LUNA
# ================================================================
function GST {
    param([double]$JD)
    $T  = ($JD - 2451545.0) / 36525.0; $T2 = $T*$T; $T3 = $T2*$T
    $g  = 280.46061837 + 360.98564736629*($JD-2451545.0) + 0.000387933*$T2 - $T3/38710000.0
    return (NA $g)
}

function Altitude {
    param([double]$RA, [double]$Dec, [double]$LST, [double]$Lat)
    $ha   = NA($LST - $RA)
    $decR = Rad $Dec; $latR = Rad $Lat; $haR = Rad $ha
    $sinA = [Math]::Sin($decR)*[Math]::Sin($latR) + [Math]::Cos($decR)*[Math]::Cos($latR)*[Math]::Cos($haR)
    $sinA = [Math]::Max(-1.0, [Math]::Min(1.0, $sinA))
    return (Deg([Math]::Asin($sinA)))
}

function SunPos {
    param([double]$JD)
    $T  = ($JD-2451545.0)/36525.0; $T2=$T*$T
    $L0 = NA(280.46646+36000.76983*$T+0.0003032*$T2)
    $M  = NA(357.52911+35999.05029*$T-0.0001537*$T2)
    $Mr = Rad $M; $e = 0.016708634-0.000042037*$T
    $C  = (1.914602-0.004817*$T-0.000014*$T2)*[Math]::Sin($Mr) + (0.019993-0.000101*$T)*[Math]::Sin(2*$Mr) + 0.000289*[Math]::Sin(3*$Mr)
    $theta = NA($L0+$C); $om = 125.04-1934.136*$T; $omr = Rad $om
    $lam = NA($theta-0.00569-0.00478*[Math]::Sin($omr))
    $eps = 23.439291111-0.013004167*$T+0.00256*[Math]::Cos($omr)
    $lamr = Rad $lam; $epsr = Rad $eps
    $RA  = NA(Deg([Math]::Atan2([Math]::Sin($lamr)*[Math]::Cos($epsr),[Math]::Cos($lamr))))
    $Dec = Deg([Math]::Asin([Math]::Sin($epsr)*[Math]::Sin($lamr)))
    return @{ RA=$RA; Dec=$Dec }
}

function MoonPos {
    param([double]$JD)
    $T=$($JD-2451545.0)/36525.0; $T2=$T*$T; $T3=$T2*$T; $T4=$T3*$T
    $Lp=NA(218.3164477+481267.88123421*$T-0.0015786*$T2+$T3/538841.0-$T4/65194000.0)
    $D =NA(297.8501921+445267.1114034*$T-0.0018819*$T2+$T3/545868.0-$T4/113065000.0)
    $M =NA(357.5291092+35999.0502909*$T-0.0001536*$T2+$T3/24490000.0)
    $Mp=NA(134.9633964+477198.8675055*$T+0.0087414*$T2+$T3/69699.0-$T4/14712000.0)
    $F =NA(93.2720950+483202.0175233*$T-0.0036539*$T2-$T3/3526000.0+$T4/863310000.0)
    $E =1.0-0.002516*$T-0.0000074*$T2
    $Dr=Rad $D; $Mr=Rad $M; $Mpr=Rad $Mp; $Fr=Rad $F
    $dL=6288774.0*[Math]::Sin($Mpr)+1274027.0*[Math]::Sin(2*$Dr-$Mpr)+658314.0*[Math]::Sin(2*$Dr)+213618.0*[Math]::Sin(2*$Mpr)+(-185116.0*$E)*[Math]::Sin($Mr)+(-114332.0)*[Math]::Sin(2*$Fr)+58793.0*[Math]::Sin(2*$Dr-2*$Mpr)+(57066.0*$E)*[Math]::Sin(2*$Dr-$Mr-$Mpr)+53322.0*[Math]::Sin(2*$Dr+$Mpr)+(45758.0*$E)*[Math]::Sin(2*$Dr-$Mr)
    $dB=5128122.0*[Math]::Sin($Fr)+280602.0*[Math]::Sin($Mpr+$Fr)+277693.0*[Math]::Sin($Mpr-$Fr)+173237.0*[Math]::Sin(2*$Dr-$Fr)+55413.0*[Math]::Sin(2*$Dr-$Mpr+$Fr)+46271.0*[Math]::Sin(2*$Dr-$Mpr-$Fr)
    $moonLon=NA($Lp+$dL/1000000.0); $moonLat=$dB/1000000.0
    $eps=23.439291111-0.013004167*$T; $epsr=Rad $eps; $lonr=Rad $moonLon; $latr=Rad $moonLat
    $RA=NA(Deg([Math]::Atan2([Math]::Sin($lonr)*[Math]::Cos($epsr)-[Math]::Tan($latr)*[Math]::Sin($epsr),[Math]::Cos($lonr))))
    $Dec=Deg([Math]::Asin([Math]::Sin($latr)*[Math]::Cos($epsr)+[Math]::Cos($latr)*[Math]::Sin($epsr)*[Math]::Sin($lonr)))
    return @{ RA=$RA; Dec=$Dec }
}

function AltAt {
    param([double]$JDE, [double]$Lat, [double]$Lon, [bool]$isSolar)
    $iSol = if ($isSolar) { 1 } else { 0 }
    return [AstroCalc]::Altitude($JDE, $Lat, $Lon, $iSol)
}

# ================================================================
#  FINESTRA VISIBILITA' LUNARE
#  jde_P1/P4 = contatti reali, % = visDur_clamped / (P4-P1)
# ================================================================
function Get-VisWindow {
    param([double]$jde_P1, [double]$jde_P4,
          [double]$Lat, [double]$Lon, [bool]$isSolar,
          [int]$NPTS = 48, [double]$minAlt = 3.0)
    $eclDur = $jde_P4 - $jde_P1
    if ($eclDur -le 0) { return @{ VisPerc=0.0; JDE1=-1.0; JDE4=-1.0 } }
    Log "  VisWindow: dur=$([Math]::Round($eclDur*24,2))h campionamento $NPTS punti..."
    $margin = 15.0/1440.0
    $t_start = $jde_P1 - $margin; $t_end = $jde_P4 + $margin
    $step = ($t_end - $t_start) / [double]($NPTS - 1)
    # Usa C# [AstroCalc]::Altitude al posto di AltAt PS: nessun overhead di chiamata
    $iSol = if ($isSolar) { 1 } else { 0 }
    $rawVis1 = -1.0; $rawVis4 = -1.0; $nVis = 0
    for ($i = 0; $i -lt $NPTS; $i++) {
        $t = $t_start + $i * $step
        if ([AstroCalc]::Altitude($t,$Lat,$Lon,$iSol) -ge $minAlt) {
            $nVis++; if ($rawVis1 -lt 0) { $rawVis1 = $t }; $rawVis4 = $t
        }
    }
    if ($nVis -eq 0) { return @{ VisPerc=0.0; JDE1=-1.0; JDE4=-1.0 } }
    if ($rawVis1 -gt $t_start) {
        $lo = $t_start; $hi = $rawVis1
        for ($k = 0; $k -lt 7; $k++) {
            $mid = ($lo+$hi)*0.5
            if ([AstroCalc]::Altitude($mid,$Lat,$Lon,$iSol) -ge $minAlt) { $hi=$mid } else { $lo=$mid }
        }
        $rawVis1 = ($lo+$hi)*0.5
    }
    if ($rawVis4 -lt $t_end) {
        $lo = $rawVis4; $hi = $t_end
        for ($k = 0; $k -lt 7; $k++) {
            $mid = ($lo+$hi)*0.5
            if ([AstroCalc]::Altitude($mid,$Lat,$Lon,$iSol) -ge $minAlt) { $lo=$mid } else { $hi=$mid }
        }
        $rawVis4 = ($lo+$hi)*0.5
    }
    $effStart = $rawVis1; if ($effStart -lt $jde_P1) { $effStart = $jde_P1 }
    $effEnd   = $rawVis4; if ($effEnd   -gt $jde_P4) { $effEnd   = $jde_P4 }
    if ($effStart -ge $effEnd) { return @{ VisPerc=0.0; JDE1=-1.0; JDE4=-1.0 } }
    $visDur  = $effEnd - $effStart
    $visPerc = [Math]::Round(($visDur / $eclDur) * 100.0, 1)
    if ($visPerc -gt 100.0) { $visPerc = 100.0 }
    return @{ VisPerc=$visPerc; JDE1=$effStart; JDE4=$effEnd }
}

# ================================================================
#  CARICAMENTO DB BESSELIANI (CSV)
#  Indicizza per JDE -> ricerca binaria rapida
# ================================================================
# $script:BessIndex e $script:BessLines gestiti da Load-BessDB

# ================================================================
#  DATABASE ELEMENTI BESSELIANI  -  caricamento 2-FASI
#
#  Fase 1 (al caricamento): legge SOLO julian_date per ogni riga
#          -> 11.898 parse invece di ~300.000 -> ~1-2 secondi
#  Fase 2 (per ogni eclisse trovata da Meeus): parse completo
#          di UNA riga -> triviale
#
#  $script:BessLines  = array di tutte le righe CSV grezze
#  $script:BessIndex  = array ordinato per JDE di {JDE, rigaIdx}
#  $script:BessHdrIdx = mappa nome-colonna -> indice
# ================================================================
$script:BessLines  = $null
$script:BessIndex  = $null
$script:BessHdrIdx = $null
$script:BessIC     = [System.Globalization.CultureInfo]::InvariantCulture

function Load-BessDB {
    param([string]$csvPath = "")

    # ---- Auto-ricerca del CSV (cartella script, cartella corrente, Downloads) ----
    if ([string]::IsNullOrWhiteSpace($csvPath) -or -not (Test-Path $csvPath)) {
        $csvName = "eclipse_besselian_from_mysqldump2.csv"
        $candidates = @(
            (Join-Path $PSScriptRoot $csvName),
            (Join-Path (Split-Path $MyInvocation.ScriptName -Parent) $csvName),
            (Join-Path ([Environment]::GetFolderPath("UserProfile")) "Downloads\$csvName"),
            (Join-Path (Get-Location).Path $csvName)
        )
        $csvPath = ""
        foreach ($c in $candidates) {
            if ($c -and (Test-Path $c)) { $csvPath = $c; break }
        }
    }

    # ---- Download automatico se il file non esiste ----
    if ([string]::IsNullOrWhiteSpace($csvPath) -or -not (Test-Path $csvPath)) {
        $saveDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
        $csvPath = Join-Path $saveDir "eclipse_besselian_from_mysqldump2.csv"
        $url = 'https://eclipse.gsfc.nasa.gov/eclipse_besselian_from_mysqldump2.csv'
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "CSV Besseliani non trovato nella cartella dello script o in Downloads.`n`nScaricare automaticamente da NASA (~7 MB)?`n$url",
            'CSV Besseliani mancante',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($ans -eq 'Yes') {
            try {
                $lblStatus.Text = "  Download CSV da NASA...  (attendere)"; $form.Refresh()
                [System.Windows.Forms.Application]::DoEvents()
                $wc = New-Object System.Net.WebClient
                $wc.DownloadFile($url, $csvPath)
                $wc.Dispose()
            } catch {
                return @{ OK=$false; Err="Download fallito: $($_.Exception.Message)" }
            }
        } else {
            return @{ OK=$false; Err='File non presente. Scaricare manualmente.' }
        }
    }
    if (-not (Test-Path $csvPath)) { return @{ OK=$false; Err="File non trovato: $csvPath" } }

    try {
        $ic = $script:BessIC

        # ---- FASE 1: leggi tutte le righe come stringhe (fast) ----
        Log "ReadAllLines: inizio lettura CSV"
        $allLines = [System.IO.File]::ReadAllLines($csvPath, [System.Text.Encoding]::UTF8)
        Log "ReadAllLines: lette $($allLines.Count) righe"

        # Mappa header -> indice colonna
        $hdr = ($allLines[0] -replace '"','').Split(',')
        $hdrIdx = @{}
        for ($hi = 0; $hi -lt $hdr.Count; $hi++) { $hdrIdx[$hdr[$hi]] = $hi }
        $jdeCol = $hdrIdx['julian_date']
        Log "Header OK - colonne trovate: $($hdr.Count) | julian_date=col.$jdeCol"

        # FASE 1: costruisce indice JDE con UNA sola parse per riga
        $nTot = $allLines.Count - 1
        Log "Fase 1: indicizzazione $nTot righe (solo julian_date)..."
        $idxList = New-Object 'System.Collections.Generic.List[psobject]'
        for ($ri = 1; $ri -lt $allLines.Count; $ri++) {
            $line = $allLines[$ri]
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $cols = $line.Split(',')
            $jdeStr = $cols[$jdeCol] -replace '"',''
            if ([string]::IsNullOrWhiteSpace($jdeStr)) { continue }
            $jde = [double]::Parse($jdeStr, $ic)
            if ($jde -gt 0) { $idxList.Add([PSCustomObject]@{ JDE=$jde; Ri=[int]$ri }) }
            if ($ri % 2000 -eq 0) { Log "  ...riga $ri / $nTot" }
        }
        Log "Indicizzazione completata: $($idxList.Count) eclissi trovate"

        Log "Ordinamento indice JDE..."
        $script:BessLines  = $allLines
        $script:BessIndex  = @($idxList | Sort-Object JDE)
        $script:BessHdrIdx = $hdrIdx
        Log "CSV pronto. $($script:BessIndex.Count) eclissi indicizzate."
        return @{ OK=$true; Err="" }
    } catch {
        return @{ OK=$false; Err=$_.Exception.Message }
    }
}

# Trova l'eclisse piu\' vicina al JDE dato (ricerca binaria sull'indice)
# e restituisce l'hashtable con TUTTI i double pre-parsati (Fase 2: una riga)
function Find-BessElem {
    param([double]$jde_approx)
    Log "  FB: inizio jde=$([Math]::Round($jde_approx,3))"
    $arr = $script:BessIndex
    Log "  FB: arr tipo=$($arr.GetType().Name) count=$(if($arr){$arr.Count}else{'NULL'})"
    if ($arr -eq $null -or $arr.Count -eq 0) { Log "  FB: arr nullo/vuoto"; return $null }

    $lo = 0; $hi = $arr.Count - 1
    Log "  FB: binary search lo=$lo hi=$hi"
    $iter = 0
    while ($lo -lt $hi) {
        $iter++; if ($iter -gt 50) { Log "  FB: LOOP INFINITO!"; break }
        $mid = [int](($lo + $hi) / 2)
        if ($arr[$mid].JDE -lt $jde_approx) { $lo = $mid + 1 } else { $hi = $mid }
    }
    Log "  FB: binary search finito in $iter iter, lo=$lo"
    $best = $arr[$lo]
    Log "  FB: best.JDE=$([Math]::Round($best.JDE,3)) best.Ri=$($best.Ri)"
    if ($lo -gt 0 -and ([Math]::Abs($arr[$lo-1].JDE - $jde_approx) -lt [Math]::Abs($best.JDE - $jde_approx))) {
        $best = $arr[$lo-1]; Log "  FB: aggiornato best con lo-1"
    }
    if ([Math]::Abs($best.JDE - $jde_approx) -gt 0.5) { Log "  FB: troppo lontano, null"; return $null }

    Log "  FB: parse riga $($best.Ri)..."
    $ic   = $script:BessIC
    $hIdx = $script:BessHdrIdx
    if ($hIdx -eq $null) { Log "  FB: hIdx e' NULL!"; return $null }
    Log "  FB: hIdx ok, split riga..."
    $line = $script:BessLines[$best.Ri]
    Log "  FB: riga lunga $($line.Length) chars"
    $cols = $line.Split(',')
    Log "  FB: $($cols.Count) colonne. td_ge_idx=$($hIdx['td_ge'])"

    $tdp = ($cols[$hIdx['td_ge']] -replace '"','').Split(':')
    Log "  FB: td_ge='$($cols[$hIdx['td_ge']])' t0='$($cols[$hIdx['t0']])'"
    $tdH = [double]::Parse($tdp[0],$ic) + [double]::Parse($tdp[1],$ic)/60.0 + [double]::Parse($tdp[2],$ic)/3600.0
    $t0H = [double]::Parse($cols[$hIdx['t0']], $ic)
    $tGE = $tdH - $t0H
    if ($tGE -gt  12) { $tGE -= 24.0 }
    if ($tGE -lt -12) { $tGE += 24.0 }
    Log "  FB: tGE=$([Math]::Round($tGE,4)) JDE_T0=$([Math]::Round($best.JDE-$tGE/24.0,4))"
    Log "  FB: parse coefficienti..."
    $jde = $best.JDE
    $result = @{
        JDE    = $jde
        JDE_T0 = $jde - $tGE / 24.0
        tmin   = [double]::Parse($cols[$hIdx['tmin']], $ic)
        tmax   = [double]::Parse($cols[$hIdx['tmax']], $ic)
        tanF1  = [double]::Parse($cols[$hIdx['tan_f1']], $ic)
        tanF2  = [double]::Parse($cols[$hIdx['tan_f2']], $ic)
        x0=[double]::Parse($cols[$hIdx['x0']],$ic); x1=[double]::Parse($cols[$hIdx['x1']],$ic)
        x2=[double]::Parse($cols[$hIdx['x2']],$ic); x3=[double]::Parse($cols[$hIdx['x3']],$ic)
        y0=[double]::Parse($cols[$hIdx['y0']],$ic); y1=[double]::Parse($cols[$hIdx['y1']],$ic)
        y2=[double]::Parse($cols[$hIdx['y2']],$ic); y3=[double]::Parse($cols[$hIdx['y3']],$ic)
        d0=[double]::Parse($cols[$hIdx['d0']],$ic);  d1=[double]::Parse($cols[$hIdx['d1']],$ic);  d2=[double]::Parse($cols[$hIdx['d2']],$ic)
        mu0=[double]::Parse($cols[$hIdx['mu0']],$ic); mu1=[double]::Parse($cols[$hIdx['mu1']],$ic); mu2=[double]::Parse($cols[$hIdx['mu2']],$ic)
        l10=[double]::Parse($cols[$hIdx['l10']],$ic); l11=[double]::Parse($cols[$hIdx['l11']],$ic); l12=[double]::Parse($cols[$hIdx['l12']],$ic)
        l20=[double]::Parse($cols[$hIdx['l20']],$ic); l21=[double]::Parse($cols[$hIdx['l21']],$ic); l22=[double]::Parse($cols[$hIdx['l22']],$ic)
    }
    LogF "  FB: COMPLETATO tmin=$($result.tmin) tmax=$($result.tmax)"
    return $result
}


# ================================================================
#  VALUTAZIONE ELEMENTI BESSELIANI AL TEMPO t
#  t = ore da T0 (ora di riferimento nel CSV)
#  Returns @{ Delta, L1, L2, InPenombra, InOmbra, SunAlt }
# ================================================================
# Bess-Eval: riceve coefficienti pre-convertiti in double (hashtable C)
# per evitare 20 parse-stringa a ogni chiamata.
# Non calcola SunAlt: il controllo altitudine e' fatto separatamente.
function Bess-Eval {
    param([hashtable]$C, [double]$t,
          [double]$sinPhi, [double]$cosPhi, [double]$Lon,
          [double]$JDE_T0)
    $t2 = $t*$t; $t3 = $t2*$t
    $PI180 = [Math]::PI / 180.0
    $bx  = $C.x0  + $C.x1*$t  + $C.x2*$t2  + $C.x3*$t3
    $by  = $C.y0  + $C.y1*$t  + $C.y2*$t2  + $C.y3*$t3
    $bd  = ($C.d0  + $C.d1*$t  + $C.d2*$t2) * $PI180
    $bmu = $C.mu0 + $C.mu1*$t + $C.mu2*$t2
    $bl1 = $C.l10 + $C.l11*$t + $C.l12*$t2
    $bl2 = $C.l20 + $C.l21*$t + $C.l22*$t2
    $thetaR = ($bmu - $Lon) * $PI180
    $sinT = [Math]::Sin($thetaR); $cosT = [Math]::Cos($thetaR)
    $sinD = [Math]::Sin($bd);     $cosD = [Math]::Cos($bd)
    $xi   = $cosPhi * $sinT
    $eta  = $sinPhi*$cosD - $cosPhi*$cosT*$sinD
    $zeta = $sinPhi*$sinD + $cosPhi*$cosT*$cosD
    $u    = $bx - $xi;  $v = $by - $eta
    $L1   = $bl1 - $zeta*$C.tanF1
    $L2   = $bl2 - $zeta*$C.tanF2
    $Delta = [Math]::Sqrt($u*$u + $v*$v)
    return @{ Delta=$Delta; L1=$L1; L2=$L2;
              InPenombra=($Delta -le $L1); InOmbra=($Delta -le [Math]::Abs($L2));
              JDE=($JDE_T0 + $t/24.0) }
}

# ================================================================
#  CIRCOSTANZE LOCALI SOLARE (con elementi Besseliani)
#  Restituisce: VisPerc, MagnLoc, SottoTipoLoc,
#               JDE_C1, JDE_C4, JDE_LocalMax,
#               UtC1, UtC4, LocC1, LocC4
# ================================================================
function Get-SolarLocalCirc {
    # C e' una hashtable con tutti i double gia' pre-parsati da Load-BessDB:
    # nessuna chiamata a PD() qui -> calcolo veloce
    param([hashtable]$C, [double]$Lat, [double]$Lon, [double]$TZ)

    $tmin   = $C.tmin;   $tmax  = $C.tmax
    $JDE_T0 = $C.JDE_T0
    if ($tmin -ge $tmax) { return @{ Visible=$false; VisPerc=0; MagnLoc=0; SottoTipoLoc="---" } }

    # Geocentrica osservatore
    $latR   = Rad $Lat
    $uGeo   = [Math]::Atan(0.99664719 * [Math]::Tan($latR))
    $sinPhi = 0.99664719 * [Math]::Sin($uGeo)
    $cosPhi = [Math]::Cos($uGeo)

    # Precompila parametri array per chiamata C# (evita hashtable lookup in loop)
    $ax = [double[]]($C.x0,$C.x1,$C.x2,$C.x3)
    $ay = [double[]]($C.y0,$C.y1,$C.y2,$C.y3)
    $ad = [double[]]($C.d0,$C.d1,$C.d2)
    $amu= [double[]]($C.mu0,$C.mu1,$C.mu2)
    $al1= [double[]]($C.l10,$C.l11,$C.l12)
    $al2= [double[]]($C.l20,$C.l21,$C.l22)
    $tf1= $C.tanF1; $tf2= $C.tanF2

    # NOTA: NON usare function interna (PS la crea nello scope globale
    #       e non vede $ax,$ay,... del chiamante). Chiamare BessEval direttamente.

    # Verifica che il tipo C# sia disponibile
    if (-not ([System.Management.Automation.PSTypeName]'AstroCalc').Type) {
        Log "  ERRORE: tipo AstroCalc non trovato - Add-Type fallito?"
        return @{ Visible=$false; VisPerc=0; MagnLoc=0; SottoTipoLoc="Errore-C#" }
    }
    Log "  Bess(C#): campionamento tmin=$([Math]::Round($tmin,2)) tmax=$([Math]::Round($tmax,2))"
    $nPts   = 40
    $dtSamp = ($tmax - $tmin) / ($nPts - 1)
    $tInPen = [System.Collections.Generic.List[double]]::new()
    for ($i = 0; $i -lt $nPts; $i++) {
        $tt = $tmin + $i * $dtSamp
        $rr = [AstroCalc]::BessEval(
            $ax[0],$ax[1],$ax[2],$ax[3], $ay[0],$ay[1],$ay[2],$ay[3],
            $ad[0],$ad[1],$ad[2], $amu[0],$amu[1],$amu[2],
            $al1[0],$al1[1],$al1[2], $al2[0],$al2[1],$al2[2],
            $tf1,$tf2, $tt, $sinPhi,$cosPhi,$Lon)
        if ($rr[0] -le $rr[1]) { $tInPen.Add($tt) }
    }
    Log "  Bess(C#): $($tInPen.Count) punti in penombra"
    if ($tInPen.Count -eq 0) {
        return @{ Visible=$false; VisPerc=0; MagnLoc=0; SottoTipoLoc="Fuori percorso" }
    }

    # Binary search C1
    $t_c1 = $tInPen[0]
    if ($t_c1 -gt $tmin) {
        $lo=$tmin; $hi=$t_c1
        for ($k=0; $k-lt 7; $k++) {
            $mid=($lo+$hi)*0.5
            $rr=[AstroCalc]::BessEval(
                $ax[0],$ax[1],$ax[2],$ax[3],$ay[0],$ay[1],$ay[2],$ay[3],
                $ad[0],$ad[1],$ad[2],$amu[0],$amu[1],$amu[2],
                $al1[0],$al1[1],$al1[2],$al2[0],$al2[1],$al2[2],
                $tf1,$tf2,$mid,$sinPhi,$cosPhi,$Lon)
            if ($rr[0] -le $rr[1]) { $hi=$mid } else { $lo=$mid }
        }
        $t_c1=($lo+$hi)*0.5
    }
    # Binary search C4
    $t_c4 = $tInPen[$tInPen.Count-1]
    if ($t_c4 -lt $tmax) {
        $lo=$t_c4; $hi=$tmax
        for ($k=0; $k-lt 7; $k++) {
            $mid=($lo+$hi)*0.5
            $rr=[AstroCalc]::BessEval(
                $ax[0],$ax[1],$ax[2],$ax[3],$ay[0],$ay[1],$ay[2],$ay[3],
                $ad[0],$ad[1],$ad[2],$amu[0],$amu[1],$amu[2],
                $al1[0],$al1[1],$al1[2],$al2[0],$al2[1],$al2[2],
                $tf1,$tf2,$mid,$sinPhi,$cosPhi,$Lon)
            if ($rr[0] -le $rr[1]) { $lo=$mid } else { $hi=$mid }
        }
        $t_c4=($lo+$hi)*0.5
    }

    # Ternary search: minimo Delta in [t_c1, t_c4]
    $lo=$t_c1; $hi=$t_c4
    for ($k=0; $k-lt 25; $k++) {
        $sp=($hi-$lo)/3.0; $m1=$lo+$sp; $m2=$hi-$sp
        $r1=[AstroCalc]::BessEval(
            $ax[0],$ax[1],$ax[2],$ax[3],$ay[0],$ay[1],$ay[2],$ay[3],
            $ad[0],$ad[1],$ad[2],$amu[0],$amu[1],$amu[2],
            $al1[0],$al1[1],$al1[2],$al2[0],$al2[1],$al2[2],
            $tf1,$tf2,$m1,$sinPhi,$cosPhi,$Lon)
        $r2=[AstroCalc]::BessEval(
            $ax[0],$ax[1],$ax[2],$ax[3],$ay[0],$ay[1],$ay[2],$ay[3],
            $ad[0],$ad[1],$ad[2],$amu[0],$amu[1],$amu[2],
            $al1[0],$al1[1],$al1[2],$al2[0],$al2[1],$al2[2],
            $tf1,$tf2,$m2,$sinPhi,$cosPhi,$Lon)
        if ($r1[0] -lt $r2[0]) { $hi=$m2 } else { $lo=$m1 }
    }
    $t_locMax=($lo+$hi)*0.5
    $rMax=[AstroCalc]::BessEval(
        $ax[0],$ax[1],$ax[2],$ax[3],$ay[0],$ay[1],$ay[2],$ay[3],
        $ad[0],$ad[1],$ad[2],$amu[0],$amu[1],$amu[2],
        $al1[0],$al1[1],$al1[2],$al2[0],$al2[1],$al2[2],
        $tf1,$tf2,$t_locMax,$sinPhi,$cosPhi,$Lon)
    $evMax = @{ Delta=$rMax[0]; L1=$rMax[1]; L2=$rMax[2] }
    Log "  Bess(C#): locMax=$([Math]::Round($t_locMax,4)) Delta=$([Math]::Round($rMax[0],4))"

    # Magnitudine e sottotipo locale
    $magnLoc      = 0.0
    $sottoTipoLoc = "Parziale"
    $absL2        = [Math]::Abs($evMax.L2)
    $denom        = $evMax.L1 + $absL2
    if ($denom -gt 0) { $magnLoc = ($evMax.L1 - $evMax.Delta) / $denom }
    if ($magnLoc -lt 0) { $magnLoc = 0.0 }

    if ($evMax.InOmbra) {
        if ($evMax.L2 -lt 0) { $sottoTipoLoc = "Totale" } else { $sottoTipoLoc = "Anulare" }
        if ($evMax.L2 -lt 0 -and $absL2 -gt 0) {
            $magnLoc = 1.0 + ($absL2 - $evMax.Delta) / $absL2
        }
    }

    # JDE contatti e massimo locale
    $jde_C1    = $JDE_T0 + $t_c1/24.0
    $jde_C4    = $JDE_T0 + $t_c4/24.0
    $jde_LMax  = $evMax.JDE

    # Visibilita': % del periodo C1-C4 in cui il Sole e' sopra l'orizzonte
    $wnd = Get-VisWindow $jde_C1 $jde_C4 $Lat $Lon $true 48 3.0

    $visPerc = $wnd.VisPerc
    $jdeV1   = $wnd.JDE1; $jdeV4 = $wnd.JDE4

    return @{
        Visible      = $true
        VisPerc      = $visPerc
        JDE_C1       = $jde_C1
        JDE_C4       = $jde_C4
        JDE_LMax     = $jde_LMax
        JDE_Vis1     = $jdeV1
        JDE_Vis4     = $jdeV4
        MagnLoc      = [Math]::Round($magnLoc, 3)
        SottoTipoLoc = $sottoTipoLoc
        AltMax       = [Math]::Round($evMax.SunAlt, 1)
    }
}

# ================================================================
#  CALCOLO ECLISSI (Meeus AFC Cap.32-33 per trovare le eclissi,
#  poi Besseliani per le circostanze locali solari)
# ================================================================
function Get-Eclipses {
    param([int]$YearStart, [int]$YearEnd,
          [double]$Lat, [double]$Lon, [double]$TZ)
    $D2R    = [Math]::PI / 180.0
    $result = [System.Collections.Generic.List[PSObject]]::new()
    $mesi   = @('','Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic')
    $k0     = [int][Math]::Floor(($YearStart-1900.0)*12.3685) - 3
    $k1     = [int][Math]::Ceiling(($YearEnd+1-1900.0)*12.3685) + 3

    for ($kb = $k0; $kb -le $k1; $kb++) {
        foreach ($hk in @(0.0, 0.5)) {
            $k = [double]$kb + $hk; $isSolar = ($hk -eq 0.0)
            $T = $k/1236.85; $T2=$T*$T; $T3=$T2*$T

            $JDE = 2415020.75933 + 29.53058868*$k + 0.0001178*$T2 - 0.000000155*$T3 `
                   + 0.00033*[Math]::Sin($D2R*(166.56+132.87*$T-0.009173*$T2))

            $M    = NA(359.2242  + 29.10535608*$k  - 0.0000333*$T2  - 0.00000347*$T3)
            $Mp   = NA(306.0253  + 385.81691806*$k + 0.0107306*$T2  + 0.00001236*$T3)
            $Fang = NA(21.2964   + 390.67050646*$k - 0.0016528*$T2  - 0.00000239*$T3)
            if ([Math]::Abs([Math]::Sin($Fang*$D2R)) -gt 0.36) { continue }

            $Mr = $M*$D2R; $Mpr = $Mp*$D2R; $Fr = $Fang*$D2R
            $dJDE  = (0.1734-0.000393*$T)*[Math]::Sin($Mr)+0.0021*[Math]::Sin(2*$Mr)
            $dJDE -= 0.4068*[Math]::Sin($Mpr); $dJDE += 0.0161*[Math]::Sin(2*$Mpr)
            $dJDE -= 0.0051*[Math]::Sin($Mr+$Mpr); $dJDE -= 0.0074*[Math]::Sin($Mr-$Mpr)
            $dJDE -= 0.0104*[Math]::Sin(2*$Fr)
            $JDE_max = $JDE + $dJDE

            $S  = 5.19595-0.0048*[Math]::Cos($Mr)+0.0020*[Math]::Cos(2*$Mr)-0.3283*[Math]::Cos($Mpr)-0.0060*[Math]::Cos($Mr+$Mpr)+0.0041*[Math]::Cos($Mr-$Mpr)
            $C  = 0.2070*[Math]::Sin($Mr)+0.0024*[Math]::Sin(2*$Mr)-0.0390*[Math]::Sin($Mpr)+0.0115*[Math]::Sin(2*$Mpr)-0.0073*[Math]::Sin($Mr+$Mpr)-0.0067*[Math]::Sin($Mr-$Mpr)+0.0117*[Math]::Sin(2*$Fr)
            $y  = $S*[Math]::Sin($Fr)+$C*[Math]::Cos($Fr); $absY = [Math]::Abs($y)
            $u  = 0.0059+0.0046*[Math]::Cos($Mr)-0.0182*[Math]::Cos($Mpr)+0.0004*[Math]::Cos(2*$Mpr)-0.0005*[Math]::Cos($Mr+$Mpr)
            $n  = 0.5458+0.0400*[Math]::Cos($Mpr); if ($n -lt 0.01) { $n = 0.01 }

            $tipoEcl=""; $sottoTipo=""; $magnGlob=0.0; $valida=$true
            $jde_P1=$JDE_max; $jde_P4=$JDE_max

            if ($isSolar) {
                $r_pen = $u+0.5460; $limEst = 1.5432+$u
                if ($absY -gt $limEst) { $valida=$false }
                else {
                    $tipoEcl = "Solare"
                    if ($absY -le 0.9972) {
                        if ($u -lt 0) { $sottoTipo="Totale"; $magnGlob=1.0+[Math]::Abs($u) }
                        elseif ($u -gt 0.0047) { $sottoTipo="Anulare"; $magnGlob=($r_pen-$absY)/(0.546+2*$u) }
                        else {
                            $cFW=[Math]::Sqrt([Math]::Max(0.0,1.0-[Math]::Sin($Fr)*[Math]::Sin($Fr)))*0.00464
                            if ($u -lt $cFW) { $sottoTipo="Ibrida" } else { $sottoTipo="Anulare" }
                            $magnGlob=($r_pen-$absY)/(0.546+2*$u)
                        }
                    } else {
                        $sottoTipo="Parziale"; $magnGlob=($r_pen-$absY)/(0.546+2*$u)
                        if ($magnGlob -le 0) { $valida=$false }
                    }
                    if ($valida) {
                        $argP = $r_pen*$r_pen-$y*$y
                        if ($argP -gt 0) { $sp=$([Math]::Sqrt($argP)/$n); $jde_P1=$JDE_max-$sp/24.0; $jde_P4=$JDE_max+$sp/24.0 }
                    }
                }
            } else {
                $tipoEcl = "Lunare"; $p_pen=1.2847+$u; $p_umb=1.0129-$u
                $magU=($p_umb-$absY)/0.5450; $magP=($p_pen-$absY)/0.5450
                if ($magU -gt 0) {
                    if ($magU -ge 1.0) { $sottoTipo="Totale" } else { $sottoTipo="Parziale" }
                    $magnGlob=$magU
                } elseif ($magP -gt 0.01) { $sottoTipo="Penombrale"; $magnGlob=$magP }
                else { $valida=$false }
                if ($valida) {
                    $argU=$p_umb*$p_umb-$y*$y
                    if ($argU -gt 0) { $sU=[Math]::Sqrt($argU)/$n; $jde_P1=$JDE_max-$sU/24.0; $jde_P4=$JDE_max+$sU/24.0 }
                    else {
                        $argPN=$p_pen*$p_pen-$y*$y
                        if ($argPN -gt 0) { $sPn=[Math]::Sqrt($argPN)/$n; $jde_P1=$JDE_max-$sPn/24.0; $jde_P4=$JDE_max+$sPn/24.0 }
                    }
                }
            }
            if (-not $valida) { continue }

            $dt = JDE-ToUTC $JDE_max
            if ($dt.Year -lt $YearStart -or $dt.Year -gt $YearEnd) { continue }

            # -------- Durata fasi lunari --------
            $durParzStr=""; $durTotStr=""
            if (-not $isSolar) {
                $pU2=1.0129-$u; $aU=$pU2*$pU2-$y*$y
                if ($aU -gt 0) { $durParzStr='{0:F0}' -f (2.0*[Math]::Sqrt($aU)/$n*60.0) }
                if ($sottoTipo -eq "Totale") {
                    $pT2=0.4679-$u; $aT=$pT2*$pT2-$y*$y
                    if ($aT -gt 0) { $durTotStr='{0:F0}' -f (2.0*[Math]::Sqrt($aT)/$n*60.0) }
                }
            }

            # -------- Visibilita' --------
            $visPerc=0.0; $jdeV1=-1.0; $jdeV4=-1.0; $visStr=""; $flag=""
            $sottoTipoLoc=$sottoTipo; $magnLoc=0.0
            $altMaxStr="---"; $utC1Str="---"; $utC4Str="---"; $locMaxStr="---"

            if ($isSolar) {
                $dtTmp = JDE-ToUTC $JDE_max
                LogF "Eclisse solare $($mesi[$dtTmp.Month]) $($dtTmp.Year) - ricerca Besseliani..."
                $B = Find-BessElem $JDE_max
                LogF "  Find-BessElem completato: $(if($B -ne $null){'trovato JDE='+[Math]::Round($B.JDE,3)}else{'NULL - CSV non caricato?'})"
                if ($B -ne $null) {
                    LogF "  Chiamo Get-SolarLocalCirc..."
                    $circ = Get-SolarLocalCirc $B $Lat $Lon $TZ
                    LogF "  Circostanze locali OK: $($circ.SottoTipoLoc) Magn=$([Math]::Round($circ.MagnLoc,3)) Vis=$($circ.VisPerc)%"
                    $sottoTipoLoc = $circ.SottoTipoLoc
                    $magnLoc      = $circ.MagnLoc
                    $altMaxStr    = '{0:F1}' -f $circ.AltMax
                    if ($circ.Visible) {
                        $jdeV1   = $circ.JDE_Vis1; $jdeV4 = $circ.JDE_Vis4
                        $visPerc = $circ.VisPerc
                        $jde_P1  = $circ.JDE_C1;   $jde_P4 = $circ.JDE_C4
                        $dtC1    = JDE-ToUTC $circ.JDE_C1
                        $dtC4    = JDE-ToUTC $circ.JDE_C4
                        $utC1Str = '{0:D2}:{1:D2}' -f $dtC1.Hour,$dtC1.Minute
                        $utC4Str = '{0:D2}:{1:D2}' -f $dtC4.Hour,$dtC4.Minute
                        $locMaxStr = JDE-ToHHMM $circ.JDE_LMax $TZ
                        if ($jdeV1 -lt 0) {
                            $visStr = "Non visibile"; $flag = "(Sole sotto orizzonte)"
                        } else {
                            if ($visPerc -ge 99.0) { $visStr = "100% (compl.)" }
                            else                   { $visStr = ('{0:F1}' -f $visPerc) + "%" }
                            if ($circ.AltMax -lt 3.0) { $flag = "! massimo non vis." }
                        }
                    } else {
                        $visStr = $circ.SottoTipoLoc   # "Fuori percorso"
                        $flag   = "(fuori penombra)"
                    }
                } else {
                    # CSV non disponibile: approssimazione altitudine
                    $altV = AltAt $JDE_max $Lat $Lon $true
                    $altMaxStr = '{0:F1}' -f $altV
                    if ($altV -lt 3.0) { $visStr="Non visibile"; $flag="(CSV non caricato)" }
                    else { $visStr="~(CSV N/D)"; $flag="CSV non trovato" }
                    $locMaxStr = JDE-ToHHMM $JDE_max $TZ
                }
            } else {
                # LUNARE
                LogF "Eclisse lunare $($mesi[$dt.Month]) $($dt.Year) - calcolo finestra visibilita'..."
                $wnd    = Get-VisWindow $jde_P1 $jde_P4 $Lat $Lon $false 48 3.0
                LogF "  Lunare OK: $sottoTipo Magn=$([Math]::Round($magnGlob,3)) Vis=$($wnd.VisPerc)%"
                $jdeV1  = $wnd.JDE1; $jdeV4=$wnd.JDE4; $visPerc=$wnd.VisPerc
                $altV   = AltAt $JDE_max $Lat $Lon $false
                $altMaxStr = '{0:F1}' -f $altV
                $locMaxStr = JDE-ToHHMM $JDE_max $TZ
                $dtC1 = JDE-ToUTC $jde_P1; $dtC4 = JDE-ToUTC $jde_P4
                $utC1Str = '{0:D2}:{1:D2}' -f $dtC1.Hour,$dtC1.Minute
                $utC4Str = '{0:D2}:{1:D2}' -f $dtC4.Hour,$dtC4.Minute
                if ($jdeV1 -lt 0) {
                    $visStr="Non visibile"; $flag="(Luna sotto orizzonte)"
                } else {
                    if ($visPerc -ge 99.0) { $visStr="100% (compl.)" }
                    else                   { $visStr=('{0:F1}' -f $visPerc)+"%"}
                    if ($altV -lt 3.0) { $flag="! massimo non vis." }
                }
                # Magnitudine: min(magn, 1.0)*100 per parziale/totale
                $magnLoc = [Math]::Min($magnGlob, 1.0)
            }

            # -------- Stringhe --------
            $dataStr  = '{0:D2} {1} {2}' -f $dt.Day,$mesi[$dt.Month],$dt.Year
            $utMaxStr = '{0:D2}:{1:D2} UT' -f $dt.Hour,$dt.Minute
            $magnStr  = if ($isSolar -and $circ -ne $null -and $circ.Visible) { '{0:F3}' -f $magnLoc }
                        elseif ($isSolar) { ('{0:F3}g' -f $magnGlob) }
                        else { '{0:F3}' -f $magnLoc }
            $locV1Str="---"; $locV4Str="---"
            if ($jdeV1 -ge 0) { $locV1Str=JDE-ToHHMM $jdeV1 $TZ; $locV4Str=JDE-ToHHMM $jdeV4 $TZ }
            $tipoAbbr = if ($isSolar) { "(*)" } else { "(O)" }
            $nodoStr  = "Disc."; if ([Math]::Abs($Fang)-lt 90 -or [Math]::Abs($Fang-360)-lt 90) { $nodoStr="Asc." }

            $result.Add([PSCustomObject]@{
                DataStr=$dataStr; Anno=$dt.Year; Mese=$dt.Month; Giorno=$dt.Day
                Tipo=$tipoAbbr; SottoTipo=$sottoTipoLoc; Magn=$magnStr
                UtP1=$utC1Str; UtMax=$utMaxStr; UtP4=$utC4Str
                LocMax=$locMaxStr; InizioVis=$locV1Str; FineVis=$locV4Str
                VisStr=$visStr; VisPerc=$visPerc; AltMax=$altMaxStr
                DurParz=$durParzStr; DurTot=$durTotStr; NodoStr=$nodoStr
                IsSolar=$isSolar; Flag=$flag
            })
        }
    }
    return ($result | Sort-Object Anno, Mese, Giorno)
}

# ================================================================
#  COLORI
# ================================================================
$BG=$([Drawing.Color]::FromArgb(13,13,30)); $PANEL=$([Drawing.Color]::FromArgb(20,20,50))
$BORDER=$([Drawing.Color]::FromArgb(60,80,140)); $FG=[Drawing.Color]::White
$ACCENT=$([Drawing.Color]::FromArgb(100,160,255)); $GOLD=$([Drawing.Color]::FromArgb(255,210,0))
$SILVER=$([Drawing.Color]::FromArgb(200,200,225)); $GREEN=$([Drawing.Color]::FromArgb(60,210,110))
$RED=$([Drawing.Color]::FromArgb(220,80,80)); $GRAY=$([Drawing.Color]::FromArgb(100,100,120))
$ORANGE=$([Drawing.Color]::FromArgb(255,160,40)); $CYAN=$([Drawing.Color]::FromArgb(80,220,220))
$YELLOW=$([Drawing.Color]::FromArgb(200,180,50))

# ================================================================
#  FORM
# ================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text="EclissiCalc v4  -  Eclissi Solari (Besseliani) & Lunari  [Espenak & Meeus]"
$fSz=New-Object Drawing.Size(1185,765); $form.Size=$fSz; $form.StartPosition="CenterScreen"
$form.BackColor=$BG; $form.ForeColor=$FG; $form.FormBorderStyle="FixedSingle"; $form.MaximizeBox=$false
$form.Font=New-Object Drawing.Font("Consolas",9)

$lblBanner=New-Object System.Windows.Forms.Label
$lblBanner.Text=" +====================================================================================================================================+ "
$lblBanner.Font=New-Object Drawing.Font("Consolas",7); $lblBanner.ForeColor=$BORDER; $lblBanner.BackColor=$PANEL
$pB1=New-Object Drawing.Point(5,5); $sB1=New-Object Drawing.Size(1165,12)
$lblBanner.Location=$pB1; $lblBanner.Size=$sB1; $form.Controls.Add($lblBanner)

$lblTitle=New-Object System.Windows.Forms.Label
$lblTitle.Text="  ***  ECLISSI SOLARI (*) con Elementi Besseliani  +  LUNARI (O)  ***  Espenak & Meeus / Meeus AFC Cap.33  ***"
$lblTitle.Font=New-Object Drawing.Font("Consolas",10,[Drawing.FontStyle]::Bold)
$lblTitle.ForeColor=$GOLD; $lblTitle.BackColor=$PANEL
$pTit=New-Object Drawing.Point(5,16); $sTit=New-Object Drawing.Size(1165,24)
$lblTitle.Location=$pTit; $lblTitle.Size=$sTit; $lblTitle.TextAlign="MiddleCenter"; $form.Controls.Add($lblTitle)

$lblBanner2=New-Object System.Windows.Forms.Label
$lblBanner2.Text=" +====================================================================================================================================+ "
$lblBanner2.Font=New-Object Drawing.Font("Consolas",7); $lblBanner2.ForeColor=$BORDER; $lblBanner2.BackColor=$PANEL
$pB2=New-Object Drawing.Point(5,39); $sB2=New-Object Drawing.Size(1165,12)
$lblBanner2.Location=$pB2; $lblBanner2.Size=$sB2; $form.Controls.Add($lblBanner2)

# ---- Pannello input (senza campo CSV: auto-rilevamento/download) ----
$pInput=New-Object System.Windows.Forms.Panel; $pInput.BackColor=$PANEL
$pIP=New-Object Drawing.Point(5,54); $pIS=New-Object Drawing.Size(1165,88)
$pInput.Location=$pIP; $pInput.Size=$pIS; $form.Controls.Add($pInput)

function Add-Field {
    param($parent,[string]$lbl,[int]$x,[int]$y,[int]$w,[string]$def,[string]$hint)
    $l=New-Object System.Windows.Forms.Label; $l.Text=$lbl; $l.ForeColor=$ACCENT; $l.BackColor=$PANEL
    $l.Font=New-Object Drawing.Font("Consolas",8); $lp=New-Object Drawing.Point($x,$y); $ls=New-Object Drawing.Size($w,14)
    $l.Location=$lp; $l.Size=$ls; $parent.Controls.Add($l)
    $t=New-Object System.Windows.Forms.TextBox; $t.Text=$def
    $t.BackColor=[Drawing.Color]::FromArgb(28,28,60); $t.ForeColor=$FG
    $t.Font=New-Object Drawing.Font("Consolas",10,[Drawing.FontStyle]::Bold); $t.BorderStyle="FixedSingle"
    $ty=$y+15; $tp=New-Object Drawing.Point($x,$ty); $ts=New-Object Drawing.Size($w,22)
    $t.Location=$tp; $t.Size=$ts; $parent.Controls.Add($t)
    if ($hint) {
        $h=New-Object System.Windows.Forms.Label; $h.Text=$hint; $h.ForeColor=$GRAY; $h.BackColor=$PANEL
        $h.Font=New-Object Drawing.Font("Consolas",7)
        $hy=$y+39; $hp=New-Object Drawing.Point($x,$hy); $hs=New-Object Drawing.Size($w,13)
        $h.Location=$hp; $h.Size=$hs; $parent.Controls.Add($h)
    }
    return $t
}

# ---- Campi input ----
$txtAnnoI=Add-Field $pInput "Anno Inizio"        10  5  95 "2026" ""
$txtAnnoF=Add-Field $pInput "Anno Fine"          115  5  95 "2026" ""
$txtLat  =Add-Field $pInput "Latitudine +N/-S"   220  5 130 "45.4654" "(es. Milano: 45.47)"
$txtLon  =Add-Field $pInput "Longitudine +E/-O"  360  5 130 "9.1859"  "(es. Milano:  9.19)"
$txtTZ   =Add-Field $pInput "Fuso Orario (h)"    500  5  90 "1"    "(UTC+1=CET)"

# ---- Checkboxes filtro ----
$chkSol=New-Object System.Windows.Forms.CheckBox
$chkSol.Text="(*) Solari"; $chkSol.Checked=$true; $chkSol.ForeColor=$GOLD
$chkSol.BackColor=$PANEL; $chkSol.Font=New-Object Drawing.Font("Consolas",9,[Drawing.FontStyle]::Bold)
$csp=New-Object Drawing.Point(605,6); $css=New-Object Drawing.Size(125,20)
$chkSol.Location=$csp; $chkSol.Size=$css; $pInput.Controls.Add($chkSol)

$chkLun=New-Object System.Windows.Forms.CheckBox
$chkLun.Text="(O) Lunari"; $chkLun.Checked=$true; $chkLun.ForeColor=$SILVER
$chkLun.BackColor=$PANEL; $chkLun.Font=New-Object Drawing.Font("Consolas",9,[Drawing.FontStyle]::Bold)
$clp=New-Object Drawing.Point(605,28); $cls=New-Object Drawing.Size(125,20)
$chkLun.Location=$clp; $chkLun.Size=$cls; $pInput.Controls.Add($chkLun)

$chkVisSolo=New-Object System.Windows.Forms.CheckBox
$chkVisSolo.Text="Solo visibili"; $chkVisSolo.Checked=$false; $chkVisSolo.ForeColor=$GREEN
$chkVisSolo.BackColor=$PANEL; $chkVisSolo.Font=New-Object Drawing.Font("Consolas",9)
$cvp=New-Object Drawing.Point(605,50); $cvs=New-Object Drawing.Size(125,20)
$chkVisSolo.Location=$cvp; $chkVisSolo.Size=$cvs; $pInput.Controls.Add($chkVisSolo)

# ---- Bottoni azione ----
$btnCalc=New-Object System.Windows.Forms.Button
$btnCalc.Text="CALCOLA"; $btnCalc.Font=New-Object Drawing.Font("Consolas",10,[Drawing.FontStyle]::Bold)
$btnCalc.ForeColor=$BG; $btnCalc.BackColor=[Drawing.Color]::FromArgb(80,140,255)
$btnCalc.FlatStyle="Flat"; $btnCalc.FlatAppearance.BorderSize=1; $btnCalc.FlatAppearance.BorderColor=$ACCENT
$bcp=New-Object Drawing.Point(745,5); $bcs=New-Object Drawing.Size(140,78)
$btnCalc.Location=$bcp; $btnCalc.Size=$bcs; $pInput.Controls.Add($btnCalc)

$btnTest=New-Object System.Windows.Forms.Button
$btnTest.Text="DIAG"; $btnTest.Font=New-Object Drawing.Font("Consolas",9,[Drawing.FontStyle]::Bold)
$btnTest.ForeColor=$BG; $btnTest.BackColor=[Drawing.Color]::FromArgb(200,150,0)
$btnTest.FlatStyle="Flat"
$btp=New-Object Drawing.Point(895,5); $bts=New-Object Drawing.Size(70,36)
$btnTest.Location=$btp; $btnTest.Size=$bts; $pInput.Controls.Add($btnTest)

$btnClr=New-Object System.Windows.Forms.Button
$btnClr.Text="CLR"; $btnClr.Font=New-Object Drawing.Font("Consolas",9)
$btnClr.ForeColor=$FG; $btnClr.BackColor=[Drawing.Color]::FromArgb(70,30,30)
$btnClr.FlatStyle="Flat"
$brp=New-Object Drawing.Point(895,47); $brs=New-Object Drawing.Size(70,36)
$btnClr.Location=$brp; $btnClr.Size=$brs; $pInput.Controls.Add($btnClr)

$btnCsv=New-Object System.Windows.Forms.Button
$btnCsv.Text="CSV"; $btnCsv.Font=New-Object Drawing.Font("Consolas",9,[Drawing.FontStyle]::Bold)
$btnCsv.ForeColor=$BG; $btnCsv.BackColor=[Drawing.Color]::FromArgb(50,160,80)
$btnCsv.FlatStyle="Flat"
$bCsvP=New-Object Drawing.Point(975,5); $bCsvS=New-Object Drawing.Size(90,36)
$btnCsv.Location=$bCsvP; $btnCsv.Size=$bCsvS; $pInput.Controls.Add($btnCsv)

$btnPdf=New-Object System.Windows.Forms.Button
$btnPdf.Text="PDF"; $btnPdf.Font=New-Object Drawing.Font("Consolas",9,[Drawing.FontStyle]::Bold)
$btnPdf.ForeColor=$BG; $btnPdf.BackColor=[Drawing.Color]::FromArgb(180,40,40)
$btnPdf.FlatStyle="Flat"
$bPdfP=New-Object Drawing.Point(975,47); $bPdfS=New-Object Drawing.Size(90,36)
$btnPdf.Location=$bPdfP; $btnPdf.Size=$bPdfS; $pInput.Controls.Add($btnPdf)

$lblSep=New-Object System.Windows.Forms.Label
$lblSep.Text=("-"*200); $lblSep.ForeColor=$BORDER; $lblSep.BackColor=$BG
$lblSep.Font=New-Object Drawing.Font("Consolas",7)
$lsp=New-Object Drawing.Point(5,145); $lss=New-Object Drawing.Size(1165,12)
$lblSep.Location=$lsp; $lblSep.Size=$lss; $form.Controls.Add($lblSep)

$dgv=New-Object System.Windows.Forms.DataGridView
$dgvP=New-Object Drawing.Point(5,157); $dgvS=New-Object Drawing.Size(1165,520)
$dgv.Location=$dgvP; $dgv.Size=$dgvS
$dgv.BackgroundColor=[Drawing.Color]::FromArgb(16,16,40); $dgv.ForeColor=$FG
$dgv.GridColor=[Drawing.Color]::FromArgb(35,35,75); $dgv.Font=New-Object Drawing.Font("Consolas",9)
$dgv.BorderStyle="None"; $dgv.RowHeadersVisible=$false; $dgv.AllowUserToAddRows=$false
$dgv.AllowUserToDeleteRows=$false; $dgv.ReadOnly=$true; $dgv.SelectionMode="FullRowSelect"
$dgv.MultiSelect=$false; $dgv.AutoSizeColumnsMode="Fill"; $dgv.ScrollBars="Both"
$hdrSt=New-Object System.Windows.Forms.DataGridViewCellStyle
$hdrSt.BackColor=[Drawing.Color]::FromArgb(22,22,65); $hdrSt.ForeColor=$ACCENT
$hdrSt.Font=New-Object Drawing.Font("Consolas",8,[Drawing.FontStyle]::Bold); $hdrSt.Alignment="MiddleCenter"
$dgv.ColumnHeadersDefaultCellStyle=$hdrSt; $dgv.EnableHeadersVisualStyles=$false; $dgv.ColumnHeadersHeight=22
$cellSt=New-Object System.Windows.Forms.DataGridViewCellStyle
$cellSt.BackColor=[Drawing.Color]::FromArgb(16,16,40); $cellSt.ForeColor=$FG
$cellSt.SelectionBackColor=[Drawing.Color]::FromArgb(50,80,170); $cellSt.SelectionForeColor=$FG
$dgv.DefaultCellStyle=$cellSt
$altSt=New-Object System.Windows.Forms.DataGridViewCellStyle; $altSt.BackColor=[Drawing.Color]::FromArgb(20,20,52)
$dgv.AlternatingRowsDefaultCellStyle=$altSt; $form.Controls.Add($dgv)

$colDefs=@(
    @{N="DataStr";  H="Data         ";W=100;A="MiddleLeft"}
    @{N="Tipo";     H="T ";           W=38; A="MiddleCenter"}
    @{N="SottoTipo";H="Sottotipo    ";W=110;A="MiddleCenter"}
    @{N="Magn";     H="Magn.";        W=62; A="MiddleRight"}
    @{N="UtP1";     H="C1 (UT)";     W=70; A="MiddleCenter"}
    @{N="UtMax";    H="Massimo UT   ";W=90; A="MiddleCenter"}
    @{N="UtP4";     H="C4 (UT)";     W=70; A="MiddleCenter"}
    @{N="LocMax";   H="Max.Loc.";     W=68; A="MiddleCenter"}
    @{N="InizioVis";H="Inizio Vis."; W=78; A="MiddleCenter"}
    @{N="FineVis";  H="Fine Vis.  "; W=73; A="MiddleCenter"}
    @{N="VisStr";   H="Visib.%    "; W=95; A="MiddleCenter"}
    @{N="AltMax";   H="Alt.(*)";      W=55; A="MiddleRight"}
    @{N="DurParz";  H="T.Parz(m)";   W=72; A="MiddleCenter"}
    @{N="DurTot";   H="T.Tot(m)";    W=65; A="MiddleCenter"}
    @{N="NodoStr";  H="Nodo";         W=45; A="MiddleCenter"}
    @{N="Flag";     H="Note          ";W=120;A="MiddleLeft"}
)
foreach ($cd in $colDefs) {
    $col=New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name=$cd.N; $col.HeaderText=$cd.H; $col.DataPropertyName=$cd.N
    $col.FillWeight=$cd.W; $col.DefaultCellStyle.Alignment=$cd.A
    [void]$dgv.Columns.Add($col)
}

$lblStatus=New-Object System.Windows.Forms.Label
$lblStatus.Text="  Caricare il CSV Besseliani e premere [ CALCOLA ]"
$lblStatus.ForeColor=$GRAY; $lblStatus.BackColor=$PANEL; $lblStatus.Font=New-Object Drawing.Font("Consolas",8)
$lsStp=New-Object Drawing.Point(5,692); $lsSts=New-Object Drawing.Size(1165,18)
$lblStatus.Location=$lsStp; $lblStatus.Size=$lsSts; $form.Controls.Add($lblStatus)

$lblLeg=New-Object System.Windows.Forms.Label
$lblLeg.Text="  (*) Solare: Besseliani -> C1/C4=contatti locali, Magn.=magnitudine locale esatta | (O) Lunare: Meeus+altitudine | Visib%=fraz.eclisse con corpo>3* | Alt(*)=gradi al massimo"
$lblLeg.ForeColor=[Drawing.Color]::FromArgb(70,70,110); $lblLeg.BackColor=$BG
$lblLeg.Font=New-Object Drawing.Font("Consolas",7)
$lLp=New-Object Drawing.Point(5,711); $lLs=New-Object Drawing.Size(1165,14)
$lblLeg.Location=$lLp; $lblLeg.Size=$lLs; $form.Controls.Add($lblLeg)

# CellFormatting
$dgv.Add_CellFormatting({
    param($s,$e)
    if ($e.RowIndex -lt 0) { return }
    $rows=$s.Rows; if ($e.RowIndex -ge $rows.Count) { return }
    $r=$rows[$e.RowIndex]
    $tipo=""; $sottoT=""; $visS=""
    try { $tipo=$r.Cells["Tipo"].Value.ToString(); $sottoT=$r.Cells["SottoTipo"].Value.ToString(); $visS=$r.Cells["VisStr"].Value.ToString() } catch { return }
    $baseClr=$SILVER; if ($tipo -eq "(*)") { $baseClr=$GOLD }
    $subClr=$baseClr
    if ($sottoT -eq "Totale") { if ($tipo -eq "(*)") { $subClr=[Drawing.Color]::FromArgb(255,80,80) } else { $subClr=$CYAN } }
    elseif ($sottoT -eq "Anulare") { $subClr=$ORANGE }
    elseif ($sottoT -eq "Ibrida")  { $subClr=[Drawing.Color]::FromArgb(255,140,60) }
    elseif ($sottoT -eq "Penombrale" -or $sottoT -eq "---" -or $sottoT -eq "Fuori percorso") { $subClr=$GRAY }
    $visClr=$FG
    if ($visS -like "Non*" -or $visS -eq "Fuori percorso") { $visClr=$GRAY }
    elseif ($visS -like "100*") { $visClr=[Drawing.Color]::FromArgb(0,255,120) }
    elseif ($visS -like "*%*") { $visClr=$GREEN }
    $altClr=$GREEN; if ($visS -like "Non*" -or $visS -eq "Fuori percorso") { $altClr=$GRAY }
    $winClr=$CYAN;  if ($visS -like "Non*" -or $visS -eq "Fuori percorso") { $winClr=$GRAY }
    $colN=$dgv.Columns[$e.ColumnIndex].Name
    if ($colN -eq "Tipo") { $e.CellStyle.ForeColor=$baseClr; $e.CellStyle.Font=New-Object Drawing.Font("Consolas",9,[Drawing.FontStyle]::Bold) }
    elseif ($colN -eq "SottoTipo") { $e.CellStyle.ForeColor=$subClr; $e.CellStyle.Font=New-Object Drawing.Font("Consolas",9,[Drawing.FontStyle]::Bold) }
    elseif ($colN -eq "Magn") { $e.CellStyle.ForeColor=$baseClr }
    elseif ($colN -eq "VisStr") { $e.CellStyle.ForeColor=$visClr; $e.CellStyle.Font=New-Object Drawing.Font("Consolas",9,[Drawing.FontStyle]::Bold) }
    elseif ($colN -eq "AltMax") { $e.CellStyle.ForeColor=$altClr }
    elseif ($colN -eq "InizioVis" -or $colN -eq "FineVis") {
        $e.CellStyle.ForeColor=$winClr
        if ($visS -notlike "Non*" -and $visS -ne "Fuori percorso") { $e.CellStyle.Font=New-Object Drawing.Font("Consolas",9,[Drawing.FontStyle]::Bold) }
    }
    elseif ($colN -eq "LocMax") { $e.CellStyle.ForeColor=$YELLOW }
    elseif ($colN -eq "DurParz" -or $colN -eq "DurTot") { $e.CellStyle.ForeColor=$CYAN }
    elseif ($colN -eq "NodoStr") { $e.CellStyle.ForeColor=[Drawing.Color]::FromArgb(130,130,170) }
    elseif ($colN -eq "Flag") { $e.CellStyle.ForeColor=[Drawing.Color]::FromArgb(80,80,130) }
    else { $e.CellStyle.ForeColor=$FG }
    if ($visS -like "Non*" -or $visS -eq "Fuori percorso") { $e.CellStyle.ForeColor=[Drawing.Color]::FromArgb(65,65,88) }
})

# ================================================================
#  EVENTI
# ================================================================
$script:eclissiGlobali=@()

$btnCalc.Add_Click({
    $script:SW.Restart()
    $script:LogLines.Clear()
    $txtLog.Clear()
    $btnCalc.Enabled = $false
    $lblStatus.ForeColor=$ACCENT; $lblStatus.Text="  Calcolo in corso..."; $form.Refresh()
    $dgv.Rows.Clear()
    LogF "=== CALCOLA premuto ==="
    # Validazione
    $annI=0;$annF=0;$lat=0.0;$lon=0.0;$tz=0.0
    try {
        $annI=[int]$txtAnnoI.Text; $annF=[int]$txtAnnoF.Text
        $lat=[double]($txtLat.Text -replace ',','.'); $lon=[double]($txtLon.Text -replace ',','.')
        $tz =[double]($txtTZ.Text  -replace ',','.')
        if ($annF -lt $annI) { throw "Anno Fine < Anno Inizio!" }
        if (($annF-$annI) -gt 200) { throw "Max 200 anni." }
        if ([Math]::Abs($lat) -gt 90) { throw "Lat non valida." }
        if ([Math]::Abs($lon) -gt 180) { throw "Lon non valida." }
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,"Errore",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        $lblStatus.Text="  Errore parametri."; $lblStatus.ForeColor=$RED; return
    }
    # Carica CSV se indicato e non gia' caricato
    # CSV auto-rilevamento (nessun campo manuale)
    if ($script:BessIndex -eq $null -or $script:BessIndex.Count -eq 0) {
        $res = Load-BessDB ""   # percorso vuoto = auto-find/download
        if (-not $res.OK) {
            [System.Windows.Forms.MessageBox]::Show("Impossibile caricare il CSV Besseliani:`n$($res.Err)","Attenzione",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
            $lblStatus.Text="  CSV non disponibile."; $lblStatus.ForeColor=$RED; $btnCalc.Enabled=$true; return
        } else {
            LogF "CSV caricato: $($script:BessIndex.Count) eclissi."
        }
    }
    $lblStatus.Text="  Calcolo eclissi e circostanze locali..."; $form.Refresh()
    try {
        $all=Get-Eclipses -YearStart $annI -YearEnd $annF -Lat $lat -Lon $lon -TZ $tz
    } catch {
        $errMsg = $_.Exception.Message
        Log "!!! ERRORE: $errMsg"
        [System.Windows.Forms.MessageBox]::Show("Errore durante il calcolo:`n$errMsg","Errore",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        $lblStatus.Text="  ERRORE - vedere log"; $lblStatus.ForeColor=$RED; return
    }
    if (-not $chkSol.Checked) { $all=@($all|Where-Object{$_.IsSolar -ne $true}) }
    if (-not $chkLun.Checked) { $all=@($all|Where-Object{$_.IsSolar -ne $false}) }
    if ($chkVisSolo.Checked)  { $all=@($all|Where-Object{$_.VisStr -notlike "Non*" -and $_.VisStr -ne "Fuori percorso"}) }
    $script:eclissiGlobali=$all
    foreach ($ec in $all) {
        $ri=$dgv.Rows.Add(); $row=$dgv.Rows[$ri]
        $row.Cells["DataStr"].Value=$ec.DataStr;  $row.Cells["Tipo"].Value=$ec.Tipo
        $row.Cells["SottoTipo"].Value=$ec.SottoTipo; $row.Cells["Magn"].Value=$ec.Magn
        $row.Cells["UtP1"].Value=$ec.UtP1;        $row.Cells["UtMax"].Value=$ec.UtMax
        $row.Cells["UtP4"].Value=$ec.UtP4;        $row.Cells["LocMax"].Value=$ec.LocMax
        $row.Cells["InizioVis"].Value=$ec.InizioVis; $row.Cells["FineVis"].Value=$ec.FineVis
        $row.Cells["VisStr"].Value=$ec.VisStr;    $row.Cells["AltMax"].Value=$ec.AltMax
        $row.Cells["DurParz"].Value=$ec.DurParz;  $row.Cells["DurTot"].Value=$ec.DurTot
        $row.Cells["NodoStr"].Value=$ec.NodoStr;  $row.Cells["Flag"].Value=$ec.Flag
    }
    $nSol=@($all|Where-Object{$_.IsSolar}).Count; $nLun=@($all|Where-Object{-not $_.IsSolar}).Count
    $nVis=@($all|Where-Object{$_.VisStr -notlike "Non*" -and $_.VisStr -ne "Fuori percorso"}).Count
    $nTot=@($all|Where-Object{$_.SottoTipo -eq "Totale"}).Count
    $latS=if($lat -ge 0){"N"}else{"S"}; $lonS=if($lon -ge 0){"E"}else{"O"}
    $bessInfo=if($script:BessIndex -ne $null){"[Besseliani: SI - $($script:BessIndex.Count) ecl.]"}else{"[Besseliani: NO]"}
    LogF "=== Calcolo completato in $($script:SW.ElapsedMilliseconds) ms ==="
    $btnCalc.Enabled = $true
    $lblStatus.ForeColor=$GREEN
    $lblStatus.Text=("  {0} eclissi ({1} sol. {2} lun.) | Visibili: {3} | Totali: {4} | {5:F2}*{6} {7:F2}*{8} TZ+{9} | {10}") -f $all.Count,$nSol,$nLun,$nVis,$nTot,[Math]::Abs($lat),$latS,[Math]::Abs($lon),$lonS,$tz,$bessInfo
})



$btnClr.Add_Click({
    $dgv.Rows.Clear(); $script:eclissiGlobali=@()
    $script:LogLines.Clear(); $txtLog.Clear(); $script:LastUIms=0
    $lblStatus.Text="  Svuotato."; $lblStatus.ForeColor=$GRAY
})

$btnTest.Add_Click({
    $script:SW.Restart(); $script:LogLines.Clear(); $txtLog.Clear()
    Log "=== DIAGNOSTICA COMPONENTI ==="

    Log "1) AstroCalc (C# compilato)..."
    try {
        $v = [AstroCalc]::Altitude(2461041.0, 45.47, 9.19, 1)
        Log "   OK - SunAlt(Milano)=$([Math]::Round($v,2)) gradi"
    } catch { Log "   ERRORE: $($_.Exception.Message)" }

    Log "2) BessDB in memoria..."
    $msg2 = "BessIndex=$(if($script:BessIndex){"$($script:BessIndex.Count) el."}else{"NULL"})"
    $msg2 += " | BessLines=$(if($script:BessLines){"$($script:BessLines.Count) righe"}else{"NULL"})"
    $msg2 += " | BessHdrIdx=$(if($script:BessHdrIdx){"$($script:BessHdrIdx.Count) col"}else{"NULL"})"
    $msg2 += " | BessIC=$(if($script:BessIC){$script:BessIC.Name}else{"NULL"})"
    Log "   $msg2"

    Log "3) Percorso CSV auto: $(if($script:BessIndex){"gia' caricato ($($script:BessIndex.Count))"}else{"non caricato (clicca CALCOLA prima)"})"
    Log "   Find-BessElem (JDE~2461041 = Apr 2026)..."
    try {
        $b = Find-BessElem 2461041.0
        if ($b) { Log "   OK - JDE=$([Math]::Round($b.JDE,3)) tmin=$($b.tmin) tmax=$($b.tmax)" }
        else     { Log "   risultato NULL" }
    } catch { Log "   ERRORE: $($_.Exception.Message)" }

    Log "4) JDE-ToUTC(2451545)=2000-01-01..."
    try { $dt=JDE-ToUTC 2451545.0; Log "   $($dt.Year)-$($dt.Month)-$($dt.Day)" }
    catch { Log "   ERRORE: $($_.Exception.Message)" }

    Log "=== FINE - controllare risultati sopra ==="
    # Forza flush finale del log
    $txtLog.AppendText("")
    [System.Windows.Forms.Application]::DoEvents()
    $lblStatus.ForeColor=$CYAN; $lblStatus.Text="  Diagnostica ok - vedere log"
})

# ---- Handler CSV ----
$btnCsv.Add_Click({
    if ($script:eclissiGlobali.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Nessun dato. Eseguire prima il calcolo.","Attenzione",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $dlg=New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title="Salva eclissi come CSV"; $dlg.Filter="CSV (*.csv)|*.csv|Tutti (*.*)|*.*"
    $annI=$txtAnnoI.Text; $annF=$txtAnnoF.Text
    $dlg.FileName="eclissi_${annI}_${annF}.csv"
    if ($dlg.ShowDialog() -eq "OK") {
        $lines=[System.Collections.Generic.List[string]]::new()
        $lines.Add("Data;Tipo;Sottotipo;Magnitudine;1ContattoUT;MassimoUT;UltContattoUT;MassimoLocale;InizioVisione;FineVisione;Visibilita_perc;AltMax_gradi;DurataParzialeMin;DurataTotaleMin;Nodo;Note")
        foreach ($ec in $script:eclissiGlobali) {
            $lines.Add(('{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11};{12};{13};{14};{15}' -f `
                $ec.DataStr,$ec.Tipo,$ec.SottoTipo,$ec.Magn,$ec.UtP1,$ec.UtMax,$ec.UtP4,
                $ec.LocMax,$ec.InizioVis,$ec.FineVis,$ec.VisStr,$ec.AltMax,
                $ec.DurParz,$ec.DurTot,$ec.NodoStr,$ec.Flag))
        }
        [System.IO.File]::WriteAllLines($dlg.FileName,$lines,[System.Text.Encoding]::UTF8)
        [System.Windows.Forms.MessageBox]::Show("CSV salvato:`n"+$dlg.FileName,"OK",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# ---- Handler PDF (headless Edge/Chrome, identico a FasiLunari.ps1) ----
$btnPdf.Add_Click({
    if ($script:eclissiGlobali.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Nessun dato. Eseguire prima il calcolo.","Attenzione",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $annI=$txtAnnoI.Text; $annF=$txtAnnoF.Text
    $dlg=New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title="Salva report eclissi come PDF"
    $dlg.Filter="PDF (*.pdf)|*.pdf|Tutti i file (*.*)|*.*"
    $dlg.FileName="Eclissi_${annI}_${annF}.pdf"
    if ($dlg.ShowDialog() -ne "OK") { return }

    $btnPdf.Enabled=$false
    $lblStatus.ForeColor=$ACCENT; $lblStatus.Text="  Generazione PDF tramite browser headless..."; $form.Refresh()
    try {
        Export-EclissiPdf $script:eclissiGlobali $dlg.FileName $txtLat.Text $txtLon.Text $txtTZ.Text $annI $annF
        $lblStatus.ForeColor=$GREEN; $lblStatus.Text="  PDF salvato: $($dlg.FileName)"
        [System.Windows.Forms.MessageBox]::Show(
            "PDF generato con successo:`n$($dlg.FileName)",
            "PDF Creato",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        $lblStatus.ForeColor=$RED; $lblStatus.Text="  ERRORE PDF: $_"
        [System.Windows.Forms.MessageBox]::Show("Errore generazione PDF:`n$_","Errore",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $btnPdf.Enabled=$true
    }
})

# ---- Area log live ----
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline     = $true
$txtLog.ScrollBars    = "Vertical"
$txtLog.ReadOnly      = $true
$txtLog.BackColor     = [Drawing.Color]::FromArgb(8,8,20)
$txtLog.ForeColor     = [Drawing.Color]::FromArgb(80,220,120)
$txtLog.Font          = New-Object Drawing.Font("Consolas",8)
$txtLog.BorderStyle   = "None"
$txtLog.WordWrap      = $false
$tLp = New-Object Drawing.Point(5,730)
$tLs = New-Object Drawing.Size(1165,120)
$txtLog.Location      = $tLp
$txtLog.Size          = $tLs
$form.Controls.Add($txtLog)
# Allarga il form per il log
$fSz2 = New-Object Drawing.Size(1185,905)
$form.Size = $fSz2

[System.Windows.Forms.Application]::Run($form)

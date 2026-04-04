#Requires -Version 5.1
# ============================================================
#  FASI LUNARI  -  Jean Meeus, Cap. 32
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ============================================================
#  FUNZIONI ASTRONOMICHE
# ============================================================
function Normalize-Angle {
    param([double]$a)
    $a = $a % 360.0
    if ($a -lt 0) { $a += 360.0 }
    return $a
}

function ConvertFrom-JulianDay {
    param([double]$JD)
    $tmp = $JD + 0.5; $Z = [Math]::Floor($tmp); $F = $tmp - $Z
    if ($Z -lt 2299161) { $A = $Z }
    else {
        $alpha = [Math]::Floor(($Z - 1867216.25) / 36524.25)
        $A = $Z + 1 + $alpha - [Math]::Floor($alpha / 4.0)
    }
    $B=$A+1524; $C=[Math]::Floor(($B-122.1)/365.25); $D=[Math]::Floor(365.25*$C)
    $E=[Math]::Floor(($B-$D)/30.6001)
    $dayFrac=$B-$D-[Math]::Floor(30.6001*$E)+$F; $day=[Math]::Floor($dayFrac)
    $hFrac=($dayFrac-$day)*24.0; $hour=[Math]::Floor($hFrac)
    $minute=[Math]::Floor(($hFrac-$hour)*60.0)
    if ($E -lt 14){$month=$E-1}else{$month=$E-13}
    if ($month -gt 2){$year=$C-4716}else{$year=$C-4715}
    return [PSCustomObject]@{Year=[int]$year;Month=[int]$month;Day=[int]$day;Hour=[int]$hour;Minute=[int]$minute}
}

function Get-LMT {
    param([int]$Hour,[int]$Minute,[double]$Lon)
    $tm=$Hour*60+$Minute+[Math]::Round($Lon/15.0*60.0)
    $tm=(($tm%1440)+1440)%1440
    return ('{0:D2}:{1:D2}' -f [int][Math]::Floor($tm/60),[int]($tm%60))
}

function Get-LunarPhases {
    param([int]$YearStart,[int]$YearEnd,[double]$Lon)
    $D2R=[Math]::PI/180.0
    $pN=@('Luna Nuova','Primo Quarto','Luna Piena','Ultimo Quarto')
    $pS=@('( * )','( ) )','( O )','( ( )')
    $pO=@(0.00,0.25,0.50,0.75)
    $res=[System.Collections.Generic.List[PSObject]]::new()
    $k0=[Math]::Floor(($YearStart-1900.0)*12.3685)-2
    $k1=[Math]::Ceiling(($YearEnd+1-1900.0)*12.3685)+2
    for($kb=[int]$k0;$kb -le [int]$k1;$kb++){
        for($pi=0;$pi -le 3;$pi++){
            $k=[double]$kb+$pO[$pi]; $T=$k/1236.85
            $JD=2415020.75933+29.53058868*$k+0.0001178*$T*$T-0.000000155*$T*$T*$T `
               +0.00033*[Math]::Sin($D2R*(166.56+132.87*$T-0.009173*$T*$T))
            $M =Normalize-Angle(359.2242 +29.10535608 *$k-0.0000333 *$T*$T-0.00000347*$T*$T*$T)
            $Mp=Normalize-Angle(306.0253 +385.81691806*$k+0.0107306 *$T*$T+0.00001236*$T*$T*$T)
            $F =Normalize-Angle(21.2964  +390.67050646*$k-0.0016528 *$T*$T-0.00000239*$T*$T*$T)
            $Mr=$M*$D2R;$Mpr=$Mp*$D2R;$Fr=$F*$D2R;$c=0.0
            if($pi -eq 0 -or $pi -eq 2){
                $c+=(0.1734-0.000393*$T)*[Math]::Sin($Mr)
                $c+=0.0021*[Math]::Sin(2*$Mr);  $c-=0.4068*[Math]::Sin($Mpr)
                $c+=0.0161*[Math]::Sin(2*$Mpr); $c-=0.0004*[Math]::Sin(3*$Mpr)
                $c+=0.0104*[Math]::Sin(2*$Fr);  $c-=0.0051*[Math]::Sin($Mr+$Mpr)
                $c-=0.0074*[Math]::Sin($Mr-$Mpr);$c+=0.0004*[Math]::Sin(2*$Fr+$Mr)
                $c-=0.0004*[Math]::Sin(2*$Fr-$Mr);$c-=0.0006*[Math]::Sin(2*$Fr+$Mpr)
                $c+=0.0010*[Math]::Sin(2*$Fr-$Mpr);$c+=0.0005*[Math]::Sin($Mr+2*$Mpr)
            }else{
                $c+=(0.1721-0.0004*$T)*[Math]::Sin($Mr)
                $c+=0.0021*[Math]::Sin(2*$Mr);  $c-=0.6280*[Math]::Sin($Mpr)
                $c+=0.0089*[Math]::Sin(2*$Mpr); $c-=0.0004*[Math]::Sin(3*$Mpr)
                $c+=0.0079*[Math]::Sin(2*$Fr);  $c-=0.0119*[Math]::Sin($Mr+$Mpr)
                $c-=0.0047*[Math]::Sin($Mr-$Mpr);$c+=0.0003*[Math]::Sin(2*$Fr+$Mr)
                $c-=0.0004*[Math]::Sin(2*$Fr-$Mr);$c-=0.0006*[Math]::Sin(2*$Fr+$Mpr)
                $c+=0.0021*[Math]::Sin(2*$Fr-$Mpr);$c+=0.0003*[Math]::Sin($Mr+2*$Mpr)
                $c+=0.0004*[Math]::Sin($Mr-2*$Mpr);$c-=0.0003*[Math]::Sin(2*$Mr+$Mpr)
                if($pi -eq 1){$c+=0.0028-0.0004*[Math]::Cos($Mr)+0.0003*[Math]::Cos($Mpr)}
                else         {$c+=-0.0028+0.0004*[Math]::Cos($Mr)-0.0003*[Math]::Cos($Mpr)}
            }
            $JD+=$c; $dt=ConvertFrom-JulianDay $JD
            if($dt.Year -ge $YearStart -and $dt.Year -le $YearEnd){
                $ut=('{0:D2}:{1:D2}' -f $dt.Hour,$dt.Minute)
                $lmt=Get-LMT -Hour $dt.Hour -Minute $dt.Minute -Lon $Lon
                $d=('{0:D4}/{1:D2}/{2:D2}' -f $dt.Year,$dt.Month,$dt.Day)
                $res.Add([PSCustomObject]@{Data=$d;UT=$ut;LMT=$lmt;Fase=$pN[$pi];Simb=$pS[$pi];JD=[Math]::Round($JD,4);SortK=$JD})
            }
        }
    }
    return ($res | Sort-Object SortK)
}

# ============================================================
#  GENERAZIONE HTML  (usata sia per l'anteprima che per il PDF)
# ============================================================
function Build-HtmlReport {
    param($Rows, $Periodo, $Coordinate, $LmtInfo, $GeneratedOn)

    $rowsHtml = [System.Text.StringBuilder]::new()
    $ri = 0
    foreach ($r in $Rows) {
        $css = switch ($r.Fase) {
            'Luna Nuova'  { 'nm' }
            'Luna Piena'  { 'fm' }
            default       { 'qu' }
        }
        $even = if ($ri % 2 -eq 0) { '' } else { ' alt' }
        [void]$rowsHtml.AppendLine("<tr class='$even'>")
        [void]$rowsHtml.AppendLine("  <td class='date'>$($r.Data)</td>")
        [void]$rowsHtml.AppendLine("  <td class='center'>$($r.UT)</td>")
        [void]$rowsHtml.AppendLine("  <td class='center lmt'>$($r.LMT)</td>")
        [void]$rowsHtml.AppendLine("  <td class='$css bold'>$($r.Fase)</td>")
        [void]$rowsHtml.AppendLine("  <td class='center $css bold'>$($r.Simb)</td>")
        [void]$rowsHtml.AppendLine("  <td class='right dim'>$($r.JD)</td>")
        [void]$rowsHtml.AppendLine("</tr>")
        $ri++
    }

    return @"
<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8">
<style>
  @page { size: A4 landscape; margin: 12mm 14mm 14mm 14mm; }

  * { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: 'Courier New', Courier, monospace;
    font-size: 8.5pt;
    color: #1a1a2e;
    background: #fff;
  }

  /* ---------- INTESTAZIONE ---------- */
  .header {
    background: linear-gradient(135deg, #2d1b6b 0%, #4a2fa0 60%, #5b3dbf 100%);
    color: #fff;
    padding: 10px 14px 8px 14px;
    border-radius: 6px;
    margin-bottom: 10px;
    page-break-inside: avoid;
  }
  .header h1 {
    font-size: 20pt;
    letter-spacing: 6px;
    color: #ffd84a;
    margin-bottom: 3px;
    font-family: 'Courier New', monospace;
  }
  .header .sub {
    font-size: 8pt;
    color: #c8bfee;
    margin-bottom: 4px;
  }
  .header .meta {
    font-size: 8pt;
    color: #e8e0ff;
    border-top: 1px solid rgba(255,255,255,0.25);
    padding-top: 5px;
    margin-top: 4px;
    display: flex;
    gap: 18px;
    flex-wrap: wrap;
  }
  .header .meta span {
    background: rgba(255,255,255,0.12);
    padding: 2px 8px;
    border-radius: 3px;
  }
  .legend {
    font-size: 8pt;
    color: #ffd84a;
    margin-top: 5px;
  }

  /* ---------- TABELLA ---------- */
  table {
    width: 100%;
    border-collapse: collapse;
    font-size: 8.5pt;
    page-break-inside: auto;
  }
  thead { display: table-header-group; }
  tr    { page-break-inside: avoid; }

  th {
    background: #4a2fa0;
    color: #fff;
    font-weight: bold;
    padding: 6px 8px;
    text-align: left;
    border-bottom: 2px solid #2d1b6b;
    font-size: 8.5pt;
    white-space: nowrap;
  }
  th.center { text-align: center; }
  th.right  { text-align: right; }

  td {
    padding: 4px 8px;
    border-bottom: 1px solid #ddd;
    white-space: nowrap;
    vertical-align: middle;
  }
  tr.alt td { background: #f3f0fb; }

  /* Allineamento celle */
  .date   { font-weight: bold; color: #1a1a2e; }
  .center { text-align: center; }
  .right  { text-align: right; }
  .bold   { font-weight: bold; }

  /* Colori fasi */
  .nm  { color: #5020a8; }   /* Luna Nuova  - viola */
  .fm  { color: #8a5500; }   /* Luna Piena  - oro scuro */
  .qu  { color: #1a5090; }   /* Quarti      - blu */
  .lmt { color: #1a7a46; font-weight: bold; }
  .dim { color: #888; font-size: 8pt; }

  /* ---------- FOOTER ---------- */
  .footer {
    margin-top: 8px;
    padding-top: 5px;
    border-top: 1px solid #ccc;
    font-size: 7pt;
    color: #888;
    display: flex;
    justify-content: space-between;
    page-break-inside: avoid;
  }
</style>
</head>
<body>

<div class="header">
  <h1>FASI &nbsp; LUNARI</h1>
  <div class="sub">Calcolo Astronomico &mdash; Jean Meeus, <em>Astronomical Formulae for Calculators</em>, Cap. 32</div>
  <div class="meta">
    <span>&#128197; Periodo: <strong>$Periodo</strong></span>
    <span>&#127759; $Coordinate</span>
    <span>&#128336; $LmtInfo</span>
    <span>Fasi totali: <strong>$($Rows.Count)</strong></span>
  </div>
  <div class="legend">
    ( * ) Luna Nuova &nbsp;&nbsp;&nbsp;
    ( ) ) Primo Quarto &nbsp;&nbsp;&nbsp;
    ( O ) Luna Piena &nbsp;&nbsp;&nbsp;
    ( ( ) Ultimo Quarto &nbsp;&nbsp;&nbsp;
    &mdash;&nbsp; Ora LMT = UT + Lon/15h (Ora Locale Media)
  </div>
</div>

<table>
  <thead>
    <tr>
      <th>Data</th>
      <th class="center">Ora UT</th>
      <th class="center">Ora LMT</th>
      <th>Fase Lunare</th>
      <th class="center">Simbolo</th>
      <th class="right">Giorno Giuliano</th>
    </tr>
  </thead>
  <tbody>
$($rowsHtml.ToString())
  </tbody>
</table>

<div class="footer">
  <span>Generato il $GeneratedOn</span>
  <span>Fonte: Jean Meeus, &ldquo;Astronomical Formulae for Calculators&rdquo;, Cap. 32 &mdash; Precisione: ~2 minuti</span>
</div>

</body>
</html>
"@
}

# ============================================================
#  ESPORTAZIONE PDF  via Edge o Chrome headless
# ============================================================
function Export-LunarPdf {
    param($Rows, $OutPath, $Periodo, $Coordinate, $LmtInfo)

    # 1. Cerca browser headless
    $candidates = @(
        'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
        'C:\Program Files\Microsoft\Edge\Application\msedge.exe'
        "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe"
        'C:\Program Files\Google\Chrome\Application\chrome.exe'
        'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    )
    $browser = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $browser) {
        throw "Microsoft Edge o Google Chrome non trovati.`nInstallare uno dei due browser per esportare in PDF."
    }

    # 2. Genera HTML temporaneo
    $genOn  = Get-Date -Format 'dd/MM/yyyy HH:mm'
    $html   = Build-HtmlReport -Rows $Rows -Periodo $Periodo -Coordinate $Coordinate -LmtInfo $LmtInfo -GeneratedOn $genOn
    $tmpDir = [System.IO.Path]::GetTempPath()
    $tmpHtml= [System.IO.Path]::Combine($tmpDir, "FasiLunari_$([System.Guid]::NewGuid().ToString('N')).html")
    [System.IO.File]::WriteAllText($tmpHtml, $html, [System.Text.Encoding]::UTF8)

    # 3. Path assoluti con slash
    $absOut  = [System.IO.Path]::GetFullPath($OutPath).Replace('\','/')
    $absHtml = $tmpHtml.Replace('\','/')
    $fileUrl = "file:///$absHtml"

    # 4. Argomenti headless (Edge e Chrome usano la stessa sintassi Chromium)
    $args = @(
        '--headless=new'
        '--disable-gpu'
        '--run-all-compositor-stages-before-draw'
        '--no-sandbox'
        "--print-to-pdf=`"$absOut`""
        '--print-to-pdf-no-header'
        '--no-pdf-header-footer'
        "`"$fileUrl`""
    )

    # 5. Avvio e attesa (timeout 30s)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = $browser
    $psi.Arguments              = $args -join ' '
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true

    $proc = [System.Diagnostics.Process]::Start($psi)
    $timedOut = -not $proc.WaitForExit(30000)   # 30 secondi
    if ($timedOut) { $proc.Kill(); throw "Timeout: il browser non ha risposto entro 30 secondi." }

    # Piccola attesa scrittura disco
    Start-Sleep -Milliseconds 800

    # 6. Cleanup HTML temp
    Remove-Item $tmpHtml -ErrorAction SilentlyContinue

    # 7. Verifica output
    if (-not (Test-Path $OutPath)) {
        throw "Il file PDF non e' stato creato (codice uscita browser: $($proc.ExitCode))."
    }
    $size = (Get-Item $OutPath).Length
    if ($size -lt 1000) {
        throw "Il file PDF sembra vuoto o corrotto ($size byte)."
    }
}

# ============================================================
#  PALETTE COLORI  UI
# ============================================================
$C_BG     = [System.Drawing.Color]::FromArgb( 14, 14, 26)
$C_PANEL  = [System.Drawing.Color]::FromArgb( 24, 24, 44)
$C_HEADER = [System.Drawing.Color]::FromArgb( 30, 20, 70)
$C_ACCENT = [System.Drawing.Color]::FromArgb(110, 70,200)
$C_BLUE   = [System.Drawing.Color]::FromArgb( 55,140,210)
$C_RED    = [System.Drawing.Color]::FromArgb(170, 50, 65)
$C_GOLD   = [System.Drawing.Color]::FromArgb(255,210, 60)
$C_FG     = [System.Drawing.Color]::FromArgb(220,220,240)
$C_DIM    = [System.Drawing.Color]::FromArgb(130,130,160)
$C_GREEN  = [System.Drawing.Color]::FromArgb( 90,200,130)
$C_ROW2   = [System.Drawing.Color]::FromArgb( 20, 20, 40)
$C_SEL    = [System.Drawing.Color]::FromArgb( 60, 45,120)
$C_BORDER = [System.Drawing.Color]::FromArgb( 60, 50,110)
$C_NM     = [System.Drawing.Color]::FromArgb(190,160,255)
$C_FQ     = [System.Drawing.Color]::FromArgb(140,210,255)
$C_DARK   = [System.Drawing.Color]::FromArgb( 50, 50, 80)

# ============================================================
#  FONT
# ============================================================
$F_TITLE  = New-Object System.Drawing.Font('Courier New',20,[System.Drawing.FontStyle]::Bold)
$F_SUB    = New-Object System.Drawing.Font('Courier New',12,[System.Drawing.FontStyle]::Regular)
$F_LEGEND = New-Object System.Drawing.Font('Courier New',11,[System.Drawing.FontStyle]::Regular)
$F_LABEL  = New-Object System.Drawing.Font('Courier New',12,[System.Drawing.FontStyle]::Bold)
$F_INPUT  = New-Object System.Drawing.Font('Courier New',13,[System.Drawing.FontStyle]::Regular)
$F_BTN    = New-Object System.Drawing.Font('Courier New',11,[System.Drawing.FontStyle]::Bold)
$F_GRID   = New-Object System.Drawing.Font('Courier New',11,[System.Drawing.FontStyle]::Regular)
$F_GHDR   = New-Object System.Drawing.Font('Courier New',11,[System.Drawing.FontStyle]::Bold)
$F_SIMB   = New-Object System.Drawing.Font('Courier New',11,[System.Drawing.FontStyle]::Bold)
$F_STATUS = New-Object System.Drawing.Font('Courier New',10,[System.Drawing.FontStyle]::Regular)

# ============================================================
#  FORM
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text            = 'Fasi Lunari  --  J. Meeus, Cap. 32'
$form.Size            = New-Object System.Drawing.Size(1400,820)
$form.MinimumSize     = New-Object System.Drawing.Size(1100,620)
$form.StartPosition   = 'CenterScreen'
$form.BackColor       = $C_BG
$form.ForeColor       = $C_FG
$form.FormBorderStyle = 'Sizable'

# ---- Intestazione ----
$pnlHead           = New-Object System.Windows.Forms.Panel
$pnlHead.Location  = New-Object System.Drawing.Point(0,0)
$pnlHead.Size      = New-Object System.Drawing.Size(1400,140)
$pnlHead.BackColor = $C_HEADER
$pnlHead.Anchor    = 'Top,Left,Right'
$form.Controls.Add($pnlHead)

$lblTitle           = New-Object System.Windows.Forms.Label
$lblTitle.Text      = 'FASI  LUNARI'
$lblTitle.Font      = $F_TITLE
$lblTitle.ForeColor = $C_GOLD
$lblTitle.BackColor = [System.Drawing.Color]::Transparent
$lblTitle.Location  = New-Object System.Drawing.Point(24,14)
$lblTitle.AutoSize  = $true
$pnlHead.Controls.Add($lblTitle)

$lblSub             = New-Object System.Windows.Forms.Label
$lblSub.Text        = 'Calcolo Astronomico  --  Jean Meeus, Astronomical Formulae for Calculators, Cap. 32'
$lblSub.Font        = $F_SUB
$lblSub.ForeColor   = $C_DIM
$lblSub.BackColor   = [System.Drawing.Color]::Transparent
$lblSub.Location    = New-Object System.Drawing.Point(26,58)
$lblSub.AutoSize    = $true
$pnlHead.Controls.Add($lblSub)

$lblSym             = New-Object System.Windows.Forms.Label
$lblSym.Text        = '( * ) Luna Nuova       ( ) ) Primo Quarto       ( O ) Luna Piena       ( ( ) Ultimo Quarto'
$lblSym.Font        = $F_LEGEND
$lblSym.ForeColor   = $C_GOLD
$lblSym.BackColor   = [System.Drawing.Color]::Transparent
$lblSym.Location    = New-Object System.Drawing.Point(26,96)
$lblSym.AutoSize    = $true
$pnlHead.Controls.Add($lblSym)

$pnlHL = New-Object System.Windows.Forms.Panel
$pnlHL.Location = New-Object System.Drawing.Point(0,138)
$pnlHL.Size     = New-Object System.Drawing.Size(1400,3)
$pnlHL.BackColor= $C_ACCENT
$pnlHL.Anchor   = 'Top,Left,Right'
$form.Controls.Add($pnlHL)

# ---- Parametri ----
$Y_PARAM = 148
$pnlParam           = New-Object System.Windows.Forms.Panel
$pnlParam.Location  = New-Object System.Drawing.Point(0,$Y_PARAM)
$pnlParam.Size      = New-Object System.Drawing.Size(1400,108)
$pnlParam.BackColor = $C_PANEL
$pnlParam.Anchor    = 'Top,Left,Right'
$form.Controls.Add($pnlParam)

function New-Lbl{param($t,$x,$y,$f,$c,$p)
    $l=New-Object System.Windows.Forms.Label; $l.Text=$t; $l.Font=$f; $l.ForeColor=$c
    $l.BackColor=[System.Drawing.Color]::Transparent
    $l.Location=New-Object System.Drawing.Point($x,$y); $l.AutoSize=$true
    $p.Controls.Add($l); return $l}

function New-Txt{param($d,$x,$y,$w,$m,$p)
    $t=New-Object System.Windows.Forms.TextBox; $t.Text=$d; $t.Font=$F_INPUT
    $t.BackColor=[System.Drawing.Color]::FromArgb(10,10,22); $t.ForeColor=$C_GOLD
    $t.BorderStyle='FixedSingle'
    $t.Location=New-Object System.Drawing.Point($x,$y); $t.Size=New-Object System.Drawing.Size($w,30); $t.MaxLength=$m
    $p.Controls.Add($t); return $t}

function New-Btn{param($t,$x,$y,$w,$h,$bg,$p)
    $b=New-Object System.Windows.Forms.Button; $b.Text=$t; $b.Font=$F_BTN
    $b.ForeColor=[System.Drawing.Color]::White; $b.BackColor=$bg; $b.FlatStyle='Flat'
    $b.FlatAppearance.BorderSize=0
    $b.Location=New-Object System.Drawing.Point($x,$y); $b.Size=New-Object System.Drawing.Size($w,$h)
    $p.Controls.Add($b); return $b}

New-Lbl 'Anno Inizio:'   16  16 $F_LABEL $C_FG $pnlParam | Out-Null
$txtStart = New-Txt (Get-Date).Year  152  12  90  4  $pnlParam
New-Lbl 'Anno Fine:'    262  16 $F_LABEL $C_FG $pnlParam | Out-Null
$txtEnd   = New-Txt (Get-Date).Year  378  12  90  4  $pnlParam
$btnCalc   = New-Btn 'CALCOLA'      492  10  200  34 $C_ACCENT $pnlParam
$btnExport = New-Btn 'ESPORTA CSV'  706  10  200  34 $C_BLUE   $pnlParam
$btnPdf    = New-Btn 'ESPORTA PDF'  920  10  200  34 $C_RED    $pnlParam
$btnPreview= New-Btn 'ANTEPRIMA'   1134  10  220  34 $C_DARK   $pnlParam
$btnExport.Enabled=$false; $btnPdf.Enabled=$false; $btnPreview.Enabled=$false

$sep=New-Object System.Windows.Forms.Label; $sep.BackColor=$C_BORDER
$sep.Location=New-Object System.Drawing.Point(16,52); $sep.Size=New-Object System.Drawing.Size(1360,2)
$pnlParam.Controls.Add($sep)

New-Lbl 'Latitudine:'    16  62 $F_LABEL $C_FG $pnlParam | Out-Null
$txtLat = New-Txt '45.4654'  140  58  110  12  $pnlParam
New-Lbl '(+ Nord / - Sud)'  262  65 $F_STATUS $C_DIM $pnlParam | Out-Null
New-Lbl 'Longitudine:'   490  62 $F_LABEL $C_FG $pnlParam | Out-Null
$txtLon = New-Txt '9.1859'   634  58  110  12  $pnlParam
New-Lbl '(+ Est / - Ovest)    Es: Milano  Lat 45.4654  Lon 9.1859'  756  65 $F_STATUS $C_DIM $pnlParam | Out-Null

$pnlPL=New-Object System.Windows.Forms.Panel; $pnlPL.Location=New-Object System.Drawing.Point(0,($Y_PARAM+107))
$pnlPL.Size=New-Object System.Drawing.Size(1400,2); $pnlPL.BackColor=$C_ACCENT; $pnlPL.Anchor='Top,Left,Right'
$form.Controls.Add($pnlPL)

# ---- Barra info ----
$Y_INFO=($Y_PARAM+110)
$pnlInfo=New-Object System.Windows.Forms.Panel; $pnlInfo.Location=New-Object System.Drawing.Point(0,$Y_INFO)
$pnlInfo.Size=New-Object System.Drawing.Size(1400,32); $pnlInfo.BackColor=[System.Drawing.Color]::FromArgb(18,18,36)
$pnlInfo.Anchor='Top,Left,Right'; $form.Controls.Add($pnlInfo)
$lblInfo=New-Object System.Windows.Forms.Label
$lblInfo.Text='  UT = Tempo Universale   |   LMT = Ora Locale Media (UT + Longitudine/15h)   |   Precisione algoritmo: ~2 min'
$lblInfo.Font=New-Object System.Drawing.Font('Courier New',10,[System.Drawing.FontStyle]::Regular)
$lblInfo.ForeColor=$C_DIM; $lblInfo.BackColor=[System.Drawing.Color]::Transparent
$lblInfo.Location=New-Object System.Drawing.Point(4,7); $lblInfo.AutoSize=$true
$pnlInfo.Controls.Add($lblInfo)

# ---- Griglia ----
$Y_GRID=$Y_INFO+32
$dgv=New-Object System.Windows.Forms.DataGridView
$dgv.Location=New-Object System.Drawing.Point(0,$Y_GRID); $dgv.Anchor='Top,Bottom,Left,Right'
$dgv.BackgroundColor=$C_BG; $dgv.GridColor=$C_BORDER; $dgv.BorderStyle='None'
$dgv.RowHeadersVisible=$false; $dgv.AllowUserToAddRows=$false; $dgv.AllowUserToDeleteRows=$false
$dgv.ReadOnly=$true; $dgv.SelectionMode='FullRowSelect'; $dgv.MultiSelect=$false
$dgv.CellBorderStyle='SingleHorizontal'; $dgv.EnableHeadersVisualStyles=$false
$dgv.AutoSizeColumnsMode='Fill'; $dgv.RowTemplate.Height=32
$dgv.DefaultCellStyle.BackColor=$C_BG; $dgv.DefaultCellStyle.ForeColor=$C_FG
$dgv.DefaultCellStyle.SelectionBackColor=$C_SEL; $dgv.DefaultCellStyle.SelectionForeColor=[System.Drawing.Color]::White
$dgv.DefaultCellStyle.Font=$F_GRID; $dgv.DefaultCellStyle.Padding=New-Object System.Windows.Forms.Padding(6,0,6,0)
$dgv.AlternatingRowsDefaultCellStyle.BackColor=$C_ROW2
$dgv.ColumnHeadersDefaultCellStyle.BackColor=$C_ACCENT
$dgv.ColumnHeadersDefaultCellStyle.ForeColor=[System.Drawing.Color]::White
$dgv.ColumnHeadersDefaultCellStyle.Font=$F_GHDR
$dgv.ColumnHeadersDefaultCellStyle.Alignment='MiddleCenter'
$dgv.ColumnHeadersDefaultCellStyle.Padding=New-Object System.Windows.Forms.Padding(6,0,6,0)
$dgv.ColumnHeadersHeight=36; $dgv.ColumnHeadersHeightSizeMode='DisableResizing'
$form.Controls.Add($dgv)

$colDefs=@(
    @{N='Data';H='  Data';          W=20;A='MiddleLeft';  Clr=$null;   Fnt=$null}
    @{N='UT';  H='Ora UT';          W=11;A='MiddleCenter';Clr=$null;   Fnt=$null}
    @{N='LMT'; H='Ora LMT';         W=11;A='MiddleCenter';Clr=$C_GREEN;Fnt=$null}
    @{N='Fase';H='  Fase Lunare';   W=26;A='MiddleLeft';  Clr=$null;   Fnt=$null}
    @{N='Simb';H='Simbolo';         W=12;A='MiddleCenter';Clr=$C_GOLD; Fnt=$F_SIMB}
    @{N='JD';  H='Giorno Giuliano'; W=21;A='MiddleRight'; Clr=$C_DIM;  Fnt=$null}
)
foreach($cd in $colDefs){
    $col=New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name=$cd.N;$col.HeaderText=$cd.H;$col.FillWeight=$cd.W;$col.DefaultCellStyle.Alignment=$cd.A
    if($cd.Clr){$col.DefaultCellStyle.ForeColor=$cd.Clr}
    if($cd.Fnt){$col.DefaultCellStyle.Font=$cd.Fnt}
    [void]$dgv.Columns.Add($col)
}
$phaseColors=@{'Luna Nuova'=$C_NM;'Primo Quarto'=$C_FQ;'Luna Piena'=$C_GOLD;'Ultimo Quarto'=$C_FQ}
$dgv.Add_CellFormatting({param($s,$e)
    if($e.RowIndex -ge 0 -and $e.ColumnIndex -eq 3){
        $f=$dgv.Rows[$e.RowIndex].Cells['Fase'].Value
        if($phaseColors.ContainsKey($f)){$e.CellStyle.ForeColor=$phaseColors[$f]}}})

# ---- Status bar ----
$status=New-Object System.Windows.Forms.Label; $status.Dock='Bottom'; $status.Height=30
$status.BackColor=$C_PANEL; $status.ForeColor=$C_DIM; $status.Font=$F_STATUS
$status.TextAlign='MiddleLeft'; $status.Text='  Inserisci i parametri e premi CALCOLA'
$form.Controls.Add($status)

# ---- Resize ----
$form.Add_Resize({
    $w=$form.ClientSize.Width
    foreach($c in @($pnlHead,$pnlHL,$pnlParam,$pnlPL,$pnlInfo)){$c.Width=$w}
    $sep.Width=$w-32
    $dgv.Size=New-Object System.Drawing.Size($w,($form.ClientSize.Height-$Y_GRID-$status.Height))
})

# ============================================================
#  STATO GLOBALE
# ============================================================
$script:allResults=$null; $script:periodoStr=''; $script:coordStr=''; $script:lmtStr=''

# ============================================================
#  BOTTONE CALCOLA
# ============================================================
$btnCalc.Add_Click({
    $errMsg='';$ys=0;$ye=0;$lat=0.0;$lon=0.0
    $inv=[System.Globalization.CultureInfo]::InvariantCulture
    $nst=[System.Globalization.NumberStyles]::Any
    if   (-not [int]::TryParse($txtStart.Text.Trim(),[ref]$ys))                                    {$errMsg='Anno Inizio non valido.'}
    elseif(-not [int]::TryParse($txtEnd.Text.Trim(),  [ref]$ye))                                   {$errMsg='Anno Fine non valido.'}
    elseif($ys -lt 1 -or $ys -gt 9999)                                                             {$errMsg='Anno Inizio fuori range (1-9999).'}
    elseif($ye -lt 1 -or $ye -gt 9999)                                                             {$errMsg='Anno Fine fuori range (1-9999).'}
    elseif($ye -lt $ys)                                                                             {$errMsg='Anno Fine deve essere >= Anno Inizio.'}
    elseif(($ye-$ys) -gt 99)                                                                        {$errMsg='Intervallo massimo: 100 anni.'}
    elseif(-not [double]::TryParse($txtLat.Text.Trim().Replace(',','.'),$nst,$inv,[ref]$lat))       {$errMsg='Latitudine non valida.'}
    elseif($lat -lt -90 -or $lat -gt 90)                                                            {$errMsg='Latitudine fuori range (-90/+90).'}
    elseif(-not [double]::TryParse($txtLon.Text.Trim().Replace(',','.'),$nst,$inv,[ref]$lon))       {$errMsg='Longitudine non valida.'}
    elseif($lon -lt -180 -or $lon -gt 180)                                                          {$errMsg='Longitudine fuori range (-180/+180).'}
    if($errMsg){
        [System.Windows.Forms.MessageBox]::Show($errMsg,'Errore Input',
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)|Out-Null
        $status.Text="  ERRORE: $errMsg"; return}

    $dgv.Rows.Clear(); $status.Text='  Calcolo in corso...'
    $btnCalc.Enabled=$false; $form.Cursor=[System.Windows.Forms.Cursors]::WaitCursor; $form.Refresh()
    try{
        $res=Get-LunarPhases -YearStart $ys -YearEnd $ye -Lon $lon
        $script:allResults=$res
        foreach($r in $res){[void]$dgv.Rows.Add($r.Data,$r.UT,$r.LMT,$r.Fase,$r.Simb,$r.JD)}
        $ns2=if($lat -ge 0){'N'}else{'S'}; $ew=if($lon -ge 0){'E'}else{'W'}
        $anni=if($ys -eq $ye){"$ys"}else{"$ys - $ye"}
        $lmtH=[int][Math]::Floor([Math]::Abs($lon)/15)
        $lmtM=[int](([Math]::Abs($lon)/15-$lmtH)*60)
        $sgn=if($lon -ge 0){'+'}else{'-'}
        $script:periodoStr=$anni
        $script:coordStr=('Lat {0:F4} {1} / Lon {2:F4} {3}' -f [Math]::Abs($lat),$ns2,[Math]::Abs($lon),$ew)
        $script:lmtStr=('LMT = UT {0}{1:D2}h {2:D2}m' -f $sgn,$lmtH,$lmtM)
        $status.Text=("  {0} fasi  |  Periodo: {1}  |  {2}  |  {3}" -f $res.Count,$anni,$script:coordStr,$script:lmtStr)
        $ok=($res.Count -gt 0)
        $btnExport.Enabled=$ok; $btnPdf.Enabled=$ok; $btnPreview.Enabled=$ok
    }catch{
        $status.Text="  ERRORE: $_"
        [System.Windows.Forms.MessageBox]::Show("Errore: $_",'Errore',
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null
    }finally{$btnCalc.Enabled=$true;$form.Cursor=[System.Windows.Forms.Cursors]::Default}
})

# ============================================================
#  BOTTONE ESPORTA CSV
# ============================================================
$btnExport.Add_Click({
    if(-not $script:allResults){return}
    $dlg=New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter='CSV (*.csv)|*.csv|Tutti i file (*.*)|*.*'
    $dlg.FileName=('FasiLunari_{0}-{1}.csv' -f $txtStart.Text,$txtEnd.Text)
    if($dlg.ShowDialog() -eq 'OK'){
        try{
            $script:allResults|Select-Object Data,UT,LMT,Fase,Simb,JD|
                Export-Csv -Path $dlg.FileName -Delimiter ';' -Encoding UTF8 -NoTypeInformation
            $status.Text="  CSV salvato: $($dlg.FileName)"
            [System.Windows.Forms.MessageBox]::Show("CSV salvato:`n$($dlg.FileName)",'OK',
                [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)|Out-Null
        }catch{
            [System.Windows.Forms.MessageBox]::Show("Errore: $_",'Errore',
                [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null
        }
    }
})

# ============================================================
#  BOTTONE ESPORTA PDF
# ============================================================
$btnPdf.Add_Click({
    if(-not $script:allResults){return}
    $dlg=New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter='PDF (*.pdf)|*.pdf|Tutti i file (*.*)|*.*'
    $dlg.FileName=('FasiLunari_{0}-{1}.pdf' -f $txtStart.Text,$txtEnd.Text)
    $dlg.Title='Esporta fasi lunari come PDF'
    if($dlg.ShowDialog() -ne 'OK'){return}

    $btnPdf.Enabled=$false; $status.Text='  Generazione PDF tramite browser headless...'
    $form.Cursor=[System.Windows.Forms.Cursors]::WaitCursor; $form.Refresh()
    try{
        Export-LunarPdf -Rows $script:allResults -OutPath $dlg.FileName `
            -Periodo $script:periodoStr -Coordinate $script:coordStr -LmtInfo $script:lmtStr
        $status.Text="  PDF salvato: $($dlg.FileName)"
        [System.Windows.Forms.MessageBox]::Show(
            "PDF generato con successo:`n$($dlg.FileName)",
            'PDF Creato',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)|Out-Null
    }catch{
        $status.Text="  ERRORE PDF: $_"
        [System.Windows.Forms.MessageBox]::Show(
            "Errore nella generazione del PDF:`n`n$_`n`nVerifica che Microsoft Edge o Chrome siano installati.",
            'Errore PDF',[System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null
    }finally{
        $btnPdf.Enabled=($null -ne $script:allResults -and $script:allResults.Count -gt 0)
        $form.Cursor=[System.Windows.Forms.Cursors]::Default
    }
})

# ============================================================
#  BOTTONE ANTEPRIMA (apre HTML nel browser)
# ============================================================
$btnPreview.Add_Click({
    if(-not $script:allResults){return}
    try{
        $genOn=Get-Date -Format 'dd/MM/yyyy HH:mm'
        $html=Build-HtmlReport -Rows $script:allResults `
            -Periodo $script:periodoStr -Coordinate $script:coordStr `
            -LmtInfo $script:lmtStr -GeneratedOn $genOn
        $tmp=[System.IO.Path]::Combine([System.IO.Path]::GetTempPath(),"FasiLunari_preview.html")
        [System.IO.File]::WriteAllText($tmp,$html,[System.Text.Encoding]::UTF8)
        Start-Process $tmp
    }catch{
        [System.Windows.Forms.MessageBox]::Show("Errore anteprima: $_",'Errore',
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null
    }
})

# ============================================================
#  TASTI RAPIDI
# ============================================================
$form.Add_KeyDown({param($s,$e)
    if($e.KeyCode -eq 'Return'){$btnCalc.PerformClick()}
    if($e.KeyCode -eq 'Escape'){$form.Close()}
})
$form.KeyPreview=$true
$dgv.Size=New-Object System.Drawing.Size(1400,492)
[System.Windows.Forms.Application]::Run($form)

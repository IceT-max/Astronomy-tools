#Requires -Version 5.1
<#
.SYNOPSIS
    JPL Horizons - Ricerca Congiunzioni Planetarie Visibili
    API: https://ssd.jpl.nasa.gov/api/horizons.api
    Richiede connessione Internet.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
Set-StrictMode -Off

#==============================================================================
# STATO GLOBALE
#==============================================================================
$script:HORIZONS_URL     = 'https://ssd.jpl.nasa.gov/api/horizons.api'
$script:SUN_CODE         = '10'
$script:Cancelled        = $false
$script:RisultatiGlobali = @()

$script:CORPI = [ordered]@{
    'Mercurio' = '199'
    'Venere'   = '299'
    'Luna'     = '301'
    'Marte'    = '499'
    'Giove'    = '599'
    'Saturno'  = '699'
    'Urano'    = '799'
    'Nettuno'  = '899'
}

#==============================================================================
# FUNZIONI
#==============================================================================

function Invoke-HorizonsAPI {
    param([string]$Codice, [string]$Site, [string]$T0, [string]$T1)

    $url = $script:HORIZONS_URL +
        '?format=text'            +
        "&COMMAND='$Codice'"      +
        "&MAKE_EPHEM='YES'"       +
        "&EPHEM_TYPE='OBSERVER'"  +
        "&CENTER='coord@399'"     +
        "&COORD_TYPE='GEODETIC'"  +
        "&SITE_COORD='$Site'"     +
        "&START_TIME='$T0'"       +
        "&STOP_TIME='$T1'"        +
        "&STEP_SIZE='1%20h'"      +
        "&QUANTITIES='1,4,9'"     +
        "&OBJ_DATA='NO'"          +
        "&ANG_FORMAT='DEG'"       +
        "&CSV_FORMAT='NO'"        +
        "&APPARENT='AIRLESS'"

    $wc = New-Object System.Net.WebClient
    $wc.Headers.Add('User-Agent', 'PS-HorizonsConjunction/2.0')
    $wc.Encoding = [System.Text.Encoding]::UTF8
    return $wc.DownloadString($url)
}

function Parse-HorizonsText {
    param([string]$Testo)

    $IC    = [System.Globalization.CultureInfo]::InvariantCulture
    $lista = [System.Collections.Generic.List[PSObject]]::new()
    $dentro = $false

    foreach ($riga in ($Testo -split "`r?`n")) {
        if ($riga -match '^\$\$SOE') { $dentro = $true;  continue }
        if ($riga -match '^\$\$EOE') { $dentro = $false; break    }
        if (-not $dentro) { continue }

        # La data e' sempre nel formato "2026-Apr-03 00:00"
        if ($riga -notmatch '(\d{4}-[A-Za-z]{3}-\d{2})\s+(\d{2}:\d{2})') { continue }
        $dateTxt = $Matches[1] + ' ' + $Matches[2]

        # Prendi tutto dopo la parte oraria, rimuovi flag (*, m, C, ecc.)
        $idx   = $riga.IndexOf($Matches[2]) + $Matches[2].Length
        if ($idx -ge $riga.Length) { continue }
        $resto = $riga.Substring($idx).Trim()
        $resto = [System.Text.RegularExpressions.Regex]::Replace($resto, '^[^\d\-\+]+', '')

        $col = @($resto -split '\s+' | Where-Object { $_ -ne '' })
        if ($col.Count -lt 4) { continue }

        $ra=[double]0; $dec=[double]0; $az=[double]0; $el=[double]0; $mag=[double]99
        $ok = $true
        try {
            $ra  = [double]::Parse($col[0], $IC)
            $dec = [double]::Parse($col[1], $IC)
            $az  = [double]::Parse($col[2], $IC)
            $el  = [double]::Parse($col[3], $IC)
        } catch { $ok = $false }
        if (-not $ok) { continue }
        if ($col.Count -ge 5) { try { $mag = [double]::Parse($col[4], $IC) } catch { $mag = 99 } }

        $dt = $null
        try { $dt = [DateTime]::ParseExact($dateTxt,'yyyy-MMM-dd HH:mm',$IC) } catch { continue }

        $lista.Add([PSCustomObject]@{ Data=$dt; AR=$ra; Dec=$dec; Az=$az; El=$el; Mag=$mag })
    }
    return $lista
}

function Get-AngSep {
    param([double]$ar1,[double]$de1,[double]$ar2,[double]$de2)
    $k = [Math]::PI / 180.0
    $c = [Math]::Sin($de1*$k)*[Math]::Sin($de2*$k) +
         [Math]::Cos($de1*$k)*[Math]::Cos($de2*$k)*[Math]::Cos(($ar1-$ar2)*$k)
    return [Math]::Acos([Math]::Max(-1.0,[Math]::Min(1.0,$c))) / $k
}

#==============================================================================
# FORM
#==============================================================================
$frm = New-Object System.Windows.Forms.Form
$frm.Text            = 'JPL Horizons  --  Congiunzioni Planetarie Visibili'
$frm.ClientSize      = New-Object System.Drawing.Size(1180,700)
$frm.StartPosition   = 'CenterScreen'
$frm.FormBorderStyle = 'Sizable'
$frm.MaximizeBox     = $false
$frm.MinimumSize     = New-Object System.Drawing.Size(900, 700)

# Blocca il ridimensionamento verticale: ripristina sempre l'altezza originale
$frm.add_Resize({
    if ($frm.WindowState -eq 'Normal' -and $frm.Height -ne 700) {
        $frm.Height = 700
    }
})
$frm.Font            = New-Object System.Drawing.Font('Consolas',9)

function mkLbl($t,$x,$y,$w=140,$h=20) {
    $l=New-Object System.Windows.Forms.Label
    $l.Text=$t; $l.Location=New-Object System.Drawing.Point($x,$y)
    $l.Size=New-Object System.Drawing.Size($w,$h); $l.TextAlign='MiddleLeft'; $l
}
function mkTxt($x,$y,$w=100,$d='') {
    $t=New-Object System.Windows.Forms.TextBox
    $t.Location=New-Object System.Drawing.Point($x,$y)
    $t.Size=New-Object System.Drawing.Size($w,22); $t.Text=$d; $t
}
function mkGB($t,$x,$y,$w,$h) {
    $g=New-Object System.Windows.Forms.GroupBox
    $g.Text=$t; $g.Location=New-Object System.Drawing.Point($x,$y)
    $g.Size=New-Object System.Drawing.Size($w,$h); $g
}
function mkBtn($t,$x,$y,$w,$h) {
    $b=New-Object System.Windows.Forms.Button
    $b.Text=$t; $b.Location=New-Object System.Drawing.Point($x,$y)
    $b.Size=New-Object System.Drawing.Size($w,$h); $b.FlatStyle='Flat'; $b
}

# ---------- GroupBox Coordinate ----------
$gbC = mkGB 'Coordinate Osservatore' 10 10 450 105
$frm.Controls.Add($gbC)
$gbC.Controls.Add((mkLbl 'Longitudine (E+):' 10 28 140))
$txtLon = mkTxt 155 26 100 '9.1895'
$gbC.Controls.Add($txtLon)
$gbC.Controls.Add((mkLbl 'deg' 260 28 30))
$gbC.Controls.Add((mkLbl 'Altitudine:' 310 28 80))
$txtAlt = mkTxt 393 26 65 '0.122'
$gbC.Controls.Add($txtAlt)
$gbC.Controls.Add((mkLbl 'km' 462 28 25))
$gbC.Controls.Add((mkLbl 'Latitudine  (N+):' 10 62 140))
$txtLat = mkTxt 155 60 100 '45.4654'
$gbC.Controls.Add($txtLat)
$gbC.Controls.Add((mkLbl 'deg' 260 62 30))
$lPre = mkLbl '[Default: Milano, IT]' 310 62 155 20
$lPre.ForeColor=[System.Drawing.Color]::Gray; $gbC.Controls.Add($lPre)

# ---------- GroupBox Periodo ----------
$gbP = mkGB 'Periodo di Ricerca' 10 125 450 80
$frm.Controls.Add($gbP)
$gbP.Controls.Add((mkLbl 'Data Inizio:' 10 28 110))
$dtpI = New-Object System.Windows.Forms.DateTimePicker
$dtpI.Location=New-Object System.Drawing.Point(125,25); $dtpI.Size=New-Object System.Drawing.Size(165,22)
$dtpI.Format='Short'; $dtpI.Value=[DateTime]::Today; $gbP.Controls.Add($dtpI)
$gbP.Controls.Add((mkLbl 'Data Fine:' 10 58 110))
$dtpF = New-Object System.Windows.Forms.DateTimePicker
$dtpF.Location=New-Object System.Drawing.Point(125,55); $dtpF.Size=New-Object System.Drawing.Size(165,22)
$dtpF.Format='Short'; $dtpF.Value=[DateTime]::Today.AddMonths(6); $gbP.Controls.Add($dtpF)
$lGG = mkLbl '' 310 40 130 20; $lGG.ForeColor=[System.Drawing.Color]::DarkBlue; $gbP.Controls.Add($lGG)
$updGG = { $d=($dtpF.Value-$dtpI.Value).Days; $lGG.Text=if($d-gt 0){"$d giorni"}else{'! intervallo'} }
$dtpI.add_ValueChanged($updGG); $dtpF.add_ValueChanged($updGG); & $updGG

# ---------- GroupBox Parametri ----------
$gbPar = mkGB 'Parametri Ricerca' 10 215 450 95
$frm.Controls.Add($gbPar)
$gbPar.Controls.Add((mkLbl 'Separazione max:' 10 28 145))
$txtSep = mkTxt 160 26 60 '5.0'; $gbPar.Controls.Add($txtSep)
$gbPar.Controls.Add((mkLbl 'gradi' 225 28 45))
$gbPar.Controls.Add((mkLbl 'Elevazione min:' 10 60 145))
$txtEl = mkTxt 160 58 60 '5.0'; $gbPar.Controls.Add($txtEl)
$gbPar.Controls.Add((mkLbl 'gradi' 225 60 45))
$chkN = New-Object System.Windows.Forms.CheckBox
$chkN.Text='Solo notte astronomica (sole < -12 deg)'; $chkN.Location=New-Object System.Drawing.Point(290,44)
$chkN.Size=New-Object System.Drawing.Size(310,22); $chkN.Checked=$true; $gbPar.Controls.Add($chkN)

# ---------- GroupBox Pianeti ----------
$gbPl = mkGB 'Pianeti / Corpi Celesti' 470 10 690 305
$gbPl.Anchor=[System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$frm.Controls.Add($gbPl)
$script:ChkCorpi = @{}
$nomi = @($script:CORPI.Keys)
$colW=195; $row=0; $col=0
foreach ($n in $nomi) {
    $cb=New-Object System.Windows.Forms.CheckBox
    $cb.Text=$n; $cb.Location=New-Object System.Drawing.Point((15+$col*$colW),(30+$row*30))
    $cb.Size=New-Object System.Drawing.Size($colW,24); $cb.Checked=$true
    $gbPl.Controls.Add($cb); $script:ChkCorpi[$n]=$cb
    $row++; if($row -ge 4){$row=0;$col++}
}
$bAll=mkBtn 'Seleziona Tutti' 15 165 130 26; $bAll.add_Click({foreach($c in $script:ChkCorpi.Values){$c.Checked=$true}}); $gbPl.Controls.Add($bAll)
$bNon=mkBtn 'Nessuno' 155 165 90 26; $bNon.add_Click({foreach($c in $script:ChkCorpi.Values){$c.Checked=$false}}); $gbPl.Controls.Add($bNon)
$lApi=mkLbl 'API: ssd.jpl.nasa.gov/api/horizons.api' 15 200 390 20
$lApi.ForeColor=[System.Drawing.Color]::DarkBlue; $lApi.Font=New-Object System.Drawing.Font('Consolas',8); $gbPl.Controls.Add($lApi)
$lQ=mkLbl 'OBSERVER | QUANTITIES=1,4,9 | ANG_FORMAT=DEG' 15 220 390 20
$lQ.ForeColor=[System.Drawing.Color]::DarkBlue; $lQ.Font=New-Object System.Drawing.Font('Consolas',8); $gbPl.Controls.Add($lQ)

# ---------- Pulsanti (y=325) ----------
$BLU   = [System.Drawing.Color]::FromArgb(0,120,212)
$ROSSO = [System.Drawing.Color]::FromArgb(196,43,28)
$VERDE = [System.Drawing.Color]::FromArgb(16,124,16)
$VIOLA = [System.Drawing.Color]::FromArgb(100,80,180)
$BIANCO= [System.Drawing.Color]::White

$ANCORA_BL=[System.Windows.Forms.AnchorStyles]::Left
$btnCerca = mkBtn 'CERCA CONGIUNZIONI' 10 325 205 36
$btnCerca.BackColor=$BLU; $btnCerca.ForeColor=$BIANCO
$btnCerca.Font=New-Object System.Drawing.Font('Consolas',10,[System.Drawing.FontStyle]::Bold)
$btnCerca.Anchor=$ANCORA_BL
$frm.Controls.Add($btnCerca)

$btnAnn = mkBtn 'Annulla' 222 325 95 36
$btnAnn.BackColor=$ROSSO; $btnAnn.ForeColor=$BIANCO; $btnAnn.Enabled=$false
$btnAnn.Anchor=$ANCORA_BL
$btnAnn.add_Click({ $script:Cancelled=$true }); $frm.Controls.Add($btnAnn)

$btnExp = mkBtn 'Esporta CSV' 324 325 110 36
$btnExp.BackColor=$VERDE; $btnExp.ForeColor=$BIANCO; $btnExp.Enabled=$false
$btnExp.Anchor=$ANCORA_BL
$frm.Controls.Add($btnExp)

$btnPDF = mkBtn 'Esporta PDF' 441 325 110 36
$btnPDF.BackColor=[System.Drawing.Color]::FromArgb(160,40,40)
$btnPDF.ForeColor=$BIANCO; $btnPDF.Enabled=$false
$btnPDF.Anchor=$ANCORA_BL
$frm.Controls.Add($btnPDF)

$btnPul = mkBtn 'Pulisci' 558 325 80 36
$btnPul.Anchor=$ANCORA_BL
$btnPul.add_Click({
    $dgv.Rows.Clear(); $script:RisultatiGlobali=@()
    $btnExp.Enabled=$false; $btnPDF.Enabled=$false
    $lSt.Text='Risultati cancellati.'; $lSt.ForeColor=[System.Drawing.Color]::DarkGray
}); $frm.Controls.Add($btnPul)

$btnTest = mkBtn 'Testa API' 645 325 95 36
$btnTest.BackColor=$VIOLA; $btnTest.ForeColor=$BIANCO
$btnTest.Anchor=$ANCORA_BL
$frm.Controls.Add($btnTest)

$lCnt = mkLbl '' 748 325 422 36
$lCnt.TextAlign='MiddleRight'
$lCnt.Anchor=[System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$frm.Controls.Add($lCnt)

# ---------- Progress bar (y=370) ----------
$pbar=New-Object System.Windows.Forms.ProgressBar
$pbar.Location=New-Object System.Drawing.Point(10,370); $pbar.Size=New-Object System.Drawing.Size(1160,18)
$pbar.Anchor=[System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$pbar.Minimum=0; $pbar.Maximum=100; $pbar.Value=0; $frm.Controls.Add($pbar)

# ---------- Status label (y=393) ----------
$lSt=New-Object System.Windows.Forms.Label
$lSt.Location=New-Object System.Drawing.Point(10,393); $lSt.Size=New-Object System.Drawing.Size(1160,20)
$lSt.Anchor=[System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$lSt.Text='Pronto.  Verde = ottima visibilita notturna  |  Giallo = parziale  |  Arancio = scarsa. Premere CERCA o TESTA API.'
$lSt.ForeColor=[System.Drawing.Color]::DarkGray; $lSt.Font=New-Object System.Drawing.Font('Consolas',8)
$frm.Controls.Add($lSt)

# ---------- DataGridView (y=418) ----------
$dgv=New-Object System.Windows.Forms.DataGridView
$dgv.Location=New-Object System.Drawing.Point(10,418); $dgv.Size=New-Object System.Drawing.Size(1160,262)
$dgv.Anchor=[System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$dgv.AllowUserToAddRows=$false; $dgv.AllowUserToDeleteRows=$false; $dgv.ReadOnly=$true
$dgv.SelectionMode='FullRowSelect'; $dgv.MultiSelect=$false; $dgv.RowHeadersVisible=$false
$dgv.AutoSizeColumnsMode='Fill'; $dgv.BackgroundColor=[System.Drawing.Color]::White
$dgv.Font=New-Object System.Drawing.Font('Consolas',8.5)
$dgv.ColumnHeadersDefaultCellStyle.Font=New-Object System.Drawing.Font('Consolas',8.5,[System.Drawing.FontStyle]::Bold)
$dgv.AlternatingRowsDefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(240,247,255)
$frm.Controls.Add($dgv)

function addCol($n,$h,$w,$a='Left') {
    $c=New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.Name=$n; $c.HeaderText=$h; $c.FillWeight=$w; $c.DefaultCellStyle.Alignment="Middle$a"; $c.SortMode='Automatic'
    $dgv.Columns.Add($c) | Out-Null
}
addCol 'Data'  'Data e Ora (UTC)'          82 'Center'
addCol 'C1'    'Corpo 1'                   58 'Center'
addCol 'C2'    'Corpo 2'                   58 'Center'
addCol 'Sep'   'Separazione (gradi)'       72 'Right'
addCol 'El1'   'Elevazione Corpo 1'        65 'Right'
addCol 'El2'   'Elevazione Corpo 2'        65 'Right'
addCol 'Mag1'  'Magnitudine C1'            55 'Right'
addCol 'Mag2'  'Magnitudine C2'            55 'Right'
addCol 'Sole'  'Elevazione Sole'           65 'Right'
addCol 'Stato' 'Visibilita'' e Condizioni' 200 'Left'

#==============================================================================
# HANDLER: Testa API
#==============================================================================
$btnTest.add_Click({
    $IC = [System.Globalization.CultureInfo]::InvariantCulture
    $lon=[double]0; $lat=[double]0; $alt=[double]0
    try{$lon=[double]::Parse($txtLon.Text.Replace(',','.'), $IC)}catch{}
    try{$lat=[double]::Parse($txtLat.Text.Replace(',','.'), $IC)}catch{}
    try{$alt=[double]::Parse($txtAlt.Text.Replace(',','.'), $IC)}catch{}
    $site=([string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0:0.0000},{1:0.0000},{2:0.000}', $lon, $lat, $alt))
    $t0=[DateTime]::Today.ToString('yyyy-MM-dd')
    $t1=[DateTime]::Today.AddDays(7).ToString('yyyy-MM-dd')

    $lSt.Text="Test API in corso... (Venere, 7 giorni, site=$site)"
    $lSt.ForeColor=[System.Drawing.Color]::DarkBlue
    [System.Windows.Forms.Application]::DoEvents()

    $rawT=''
    try {
        $rawT   = Invoke-HorizonsAPI -Codice '299' -Site $site -T0 $t0 -T1 $t1
        $parsed = Parse-HorizonsText -Testo $rawT
        $nP     = $parsed.Count
        $linee  = $rawT -split "`r?`n"

        # Trova $$SOE
        $soeN=-1
        for($ii=0;$ii -lt $linee.Count;$ii++){ if($linee[$ii] -match '^\$\$SOE'){$soeN=$ii;break} }

        if($soeN -ge 0) {
            $ctx="--- Righe intorno a `$`$SOE (riga $soeN) ---`n"
            $da=[Math]::Max(0,$soeN-3)
            for($ii=$da;$ii -le [Math]::Min($linee.Count-1,$soeN+8);$ii++){$ctx+="[$ii] $($linee[$ii])`n"}
        } else {
            $ctx="!!! `$`$SOE NON TROVATO !!!`n" + (($linee | Select-Object -First 25) -join "`n")
        }

        $prima=if($nP-gt 0){"Prima riga: $($parsed[0].Data.ToString('yyyy-MM-dd'))  RA=$($parsed[0].AR)  Dec=$($parsed[0].Dec)  Az=$($parsed[0].Az)  El=$($parsed[0].El)"}else{"NESSUNA RIGA PARSATA"}

        $msg = "Site: $site`nPeriodo: $t0 -> $t1`nRighe totali: $($linee.Count)`nRighe parsate: $nP`n$prima`n`n$ctx"
        [System.Windows.Forms.MessageBox]::Show($msg,'Diagnostica API - Venere',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null

        $lSt.Text="Test OK: $nP righe parsate."
        $lSt.ForeColor=if($nP-gt 0){[System.Drawing.Color]::DarkGreen}else{[System.Drawing.Color]::Red}
    } catch {
        $snippet=if($rawT -ne ''){($rawT -split "`r?`n" | Select-Object -First 15)-join "`n"}else{'(nessuna risposta)'}
        [System.Windows.Forms.MessageBox]::Show("ERRORE: $($_.Exception.Message)`n`nPrime 15 righe risposta:`n$snippet",
            'Errore Test API',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        $lSt.Text='Test fallito: '+$_.Exception.Message; $lSt.ForeColor=[System.Drawing.Color]::Red
    }
})

#==============================================================================
# HANDLER: Cerca Congiunzioni
#==============================================================================
$btnCerca.add_Click({
    $IC=[System.Globalization.CultureInfo]::InvariantCulture
    $errV=@()
    $lon=[double]0;$lat=[double]0;$alt=[double]0;$sepM=[double]5;$elM=[double]5

    try{$lon=[double]::Parse($txtLon.Text.Replace(',','.'), $IC)}catch{$errV+='Longitudine non valida.'}
    try{$lat=[double]::Parse($txtLat.Text.Replace(',','.'), $IC)}catch{$errV+='Latitudine non valida.' }
    try{$alt=[double]::Parse($txtAlt.Text.Replace(',','.'), $IC)}catch{$errV+='Altitudine non valida.' }
    try{$sepM=[double]::Parse($txtSep.Text.Replace(',','.'), $IC)}catch{$errV+='Separazione non valida.'}
    try{$elM=[double]::Parse($txtEl.Text.Replace(',','.'),  $IC)}catch{$errV+='Elevazione non valida.' }

    if($lon-lt -180 -or $lon-gt 180){$errV+='Longitudine fuori range (-180..180).'}
    if($lat-lt -90  -or $lat-gt 90 ){$errV+='Latitudine fuori range (-90..90).'}
    if($sepM-le 0   -or $sepM-gt 180){$errV+='Separazione deve essere 0..180 gradi.'}
    if($dtpF.Value -le $dtpI.Value)  {$errV+='Data fine deve essere successiva alla data inizio.'}

    $sel=@($script:CORPI.Keys | Where-Object { $script:ChkCorpi[$_].Checked })
    if($sel.Count -lt 2){$errV+='Selezionare almeno 2 corpi celesti.'}

    if($errV.Count -gt 0){
        [System.Windows.Forms.MessageBox]::Show(($errV -join "`n"),'Errori di input',
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    $dgv.Rows.Clear(); $script:RisultatiGlobali=@()
    $btnExp.Enabled=$false; $btnCerca.Enabled=$false; $btnAnn.Enabled=$true
    $script:Cancelled=$false; $pbar.Value=0; $lCnt.Text=''

    $t0=$dtpI.Value.ToString('yyyy-MM-dd')
    $t1=$dtpF.Value.ToString('yyyy-MM-dd')
    $site=([string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0:0.0000},{1:0.0000},{2:0.000}', $lon, $lat, $alt))
    $totQ=$sel.Count+1; $done=0; $eph=@{}

    # --- Sole ---
    $done++
    $lSt.Text="[$done/$totQ] Query Sole..."; $lSt.ForeColor=[System.Drawing.Color]::DarkBlue
    [System.Windows.Forms.Application]::DoEvents()
    $rawS=''
    try {
        $rawS=$rawS=(Invoke-HorizonsAPI -Codice $script:SUN_CODE -Site $site -T0 $t0 -T1 $t1)
        $eph['__sole']=Parse-HorizonsText -Testo $rawS
        $lSt.Text="[$done/$totQ] Sole: $($eph['__sole'].Count) date."
        $pbar.Value=[int](($done/$totQ)*100); [System.Windows.Forms.Application]::DoEvents()
    } catch {
        $lSt.Text='Errore Sole: '+$_.Exception.Message; $lSt.ForeColor=[System.Drawing.Color]::Red
        $btnCerca.Enabled=$true; $btnAnn.Enabled=$false; return
    }

    # --- Pianeti ---
    foreach($nome in $sel) {
        if($script:Cancelled){break}
        $done++
        $cod=$script:CORPI[$nome]
        $lSt.Text="[$done/$totQ] Query $($nome) (cod $cod)..."; $lSt.ForeColor=[System.Drawing.Color]::DarkBlue
        [System.Windows.Forms.Application]::DoEvents()
        $rawP=''
        try {
            $rawP=(Invoke-HorizonsAPI -Codice $cod -Site $site -T0 $t0 -T1 $t1)
            $eph[$nome]=Parse-HorizonsText -Testo $rawP
            $lSt.Text="[$done/$totQ] $($nome): $($eph[$nome].Count) date parsate."
            $pbar.Value=[int](($done/$totQ)*100); [System.Windows.Forms.Application]::DoEvents()
        } catch {
            $lSt.Text="Errore $($nome): "+$_.Exception.Message; $lSt.ForeColor=[System.Drawing.Color]::Red
            $btnCerca.Enabled=$true; $btnAnn.Enabled=$false; return
        }
    }

    if($script:Cancelled){
        $lSt.Text='Annullato.'; $lSt.ForeColor=[System.Drawing.Color]::DarkRed
        $btnCerca.Enabled=$true; $btnAnn.Enabled=$false; $pbar.Value=0; return
    }

    $lSt.Text='Analisi coppie (dati orari)...'; [System.Windows.Forms.Application]::DoEvents()

    # Lookup orario: chiave = DateTime (data + ora)
    $soleLU=@{}
    foreach($r in $eph['__sole']){ $soleLU[$r.Data]=$r }

    $lu=@{}
    foreach($nome in $sel){
        $lu[$nome]=@{}
        if($eph.ContainsKey($nome)){foreach($r in $eph[$nome]){$lu[$nome][$r.Data]=$r}}
    }

    # Coppie
    $coppie=@()
    for($i=0;$i-lt $sel.Count;$i++){for($j=$i+1;$j-lt $sel.Count;$j++){$coppie+=,@($sel[$i],$sel[$j])}}

    # Giorni univoci presenti nei dati
    $giorni=@($eph['__sole'] | ForEach-Object { $_.Data.Date } | Sort-Object -Unique)

    $conj=[System.Collections.Generic.List[PSObject]]::new()

    foreach($giorno in $giorni) {

        foreach($cp in $coppie) {
            $n1=$cp[0]; $n2=$cp[1]

            # Per ogni ora del giorno trova la migliore finestra di visibilita'
            $bestSep=[double]999; $bestR1=$null; $bestR2=$null; $bestElS=0; $bestOra=$null

            for($h=0; $h-le 23; $h++) {
                $ora = $giorno.AddHours($h)

                if(-not $soleLU.ContainsKey($ora)){continue}
                if(-not $lu[$n1].ContainsKey($ora)){continue}
                if(-not $lu[$n2].ContainsKey($ora)){continue}

                $rS=$soleLU[$ora]; $r1=$lu[$n1][$ora]; $r2=$lu[$n2][$ora]

                # Filtro notte astronomica (opzionale)
                if($chkN.Checked -and $rS.El -gt -12){continue}

                # Almeno un corpo sopra la soglia elevazione
                if($r1.El -lt $elM -and $r2.El -lt $elM){continue}

                $sep=Get-AngSep $r1.AR $r1.Dec $r2.AR $r2.Dec
                if($sep -lt $bestSep){
                    $bestSep=$sep; $bestR1=$r1; $bestR2=$r2; $bestElS=$rS.El; $bestOra=$ora
                }
            }

            # Solo se la separazione minima dell'ora migliore e' entro soglia
            if($bestSep -gt $sepM -or $bestR1 -eq $null){continue}

            $stArr=@()
            if($bestR1.El -ge $elM){$stArr+="$n1 vis."}
            if($bestR2.El -ge $elM){$stArr+="$n2 vis."}
            if    ($bestElS -lt -18){$stArr+='Notte ast.'}
            elseif($bestElS -lt -12){$stArr+='Notte naut.'}
            elseif($bestElS -lt  -6){$stArr+='Crepuscolo'}
            else                    {$stArr+='GIORNO'}

            $bg=if($bestElS-lt -12 -and $bestR1.El-ge $elM -and $bestR2.El-ge $elM){[System.Drawing.Color]::FromArgb(210,255,210)}
               elseif($bestR1.El-ge $elM -or $bestR2.El-ge $elM)                   {[System.Drawing.Color]::FromArgb(255,255,200)}
               else                                                                  {[System.Drawing.Color]::FromArgb(255,230,220)}

            $m1=if($bestR1.Mag-ge 99){'n/a'}else{('{0:0.0}'-f $bestR1.Mag)}
            $m2=if($bestR2.Mag-ge 99){'n/a'}else{('{0:0.0}'-f $bestR2.Mag)}

            $conj.Add([PSCustomObject]@{
                # Mostra data + ora migliore UTC
                Data=$bestOra.ToString('yyyy-MM-dd HH:mm'); C1=$n1; C2=$n2
                Sep=[Math]::Round($bestSep,3)
                El1=[Math]::Round($bestR1.El,1); El2=[Math]::Round($bestR2.El,1)
                Mag1=$m1; Mag2=$m2; Sole=[Math]::Round($bestElS,1)
                Stato=($stArr -join ' | '); BG=$bg
            })
        }
    }

    $sorted=$conj | Sort-Object Data, Sep
    foreach($c in $sorted) {
        $idx=$dgv.Rows.Add(); $row=$dgv.Rows[$idx]
        $row.Cells['Data'].Value  = $c.Data
        $row.Cells['C1'].Value    = $c.C1
        $row.Cells['C2'].Value    = $c.C2
        $row.Cells['Sep'].Value   = "$($c.Sep) deg"
        $row.Cells['El1'].Value   = "$($c.El1) deg"
        $row.Cells['El2'].Value   = "$($c.El2) deg"
        $row.Cells['Mag1'].Value  = $c.Mag1
        $row.Cells['Mag2'].Value  = $c.Mag2
        $row.Cells['Sole'].Value  = "$($c.Sole) deg"
        $row.Cells['Stato'].Value = $c.Stato
        $row.DefaultCellStyle.BackColor=$c.BG
    }

    $script:RisultatiGlobali=$sorted
    $n=$sorted.Count
    $lSt.Text="Completato: $n congiunzione/i nel periodo $t0 / $t1."
    $lSt.ForeColor=if($n-gt 0){[System.Drawing.Color]::DarkGreen}else{[System.Drawing.Color]::DarkOrange}
    $lCnt.Text="$n risultati"; $pbar.Value=100
    $btnExp.Enabled=($n-gt 0); $btnPDF.Enabled=($n-gt 0); $btnCerca.Enabled=$true; $btnAnn.Enabled=$false
})

#==============================================================================
# HANDLER: Esporta CSV
#==============================================================================
$btnExp.add_Click({
    $sfd=New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Title='Salva congiunzioni CSV'; $sfd.Filter='File CSV (*.csv)|*.csv|Tutti (*.*)|*.*'
    $sfd.FileName='congiunzioni_'+[DateTime]::Today.ToString('yyyyMMdd')+'.csv'
    $sfd.InitialDirectory=[Environment]::GetFolderPath('Desktop')
    if($sfd.ShowDialog() -eq 'OK') {
        try {
            $ll=@('Data,Corpo1,Corpo2,"Sep(deg)","El.C1","El.C2",Mag1,Mag2,"El.Sole",Stato')
            foreach($c in $script:RisultatiGlobali){
                $ll+=('{0},{1},{2},{3},{4},{5},{6},{7},{8},"{9}"' -f $c.Data,$c.C1,$c.C2,$c.Sep,$c.El1,$c.El2,$c.Mag1,$c.Mag2,$c.Sole,$c.Stato)
            }
            [System.IO.File]::WriteAllLines($sfd.FileName,$ll,[System.Text.Encoding]::UTF8)
            [System.Windows.Forms.MessageBox]::Show("Salvato:`n$($sfd.FileName)",'OK',
                [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Errore: $_",'Errore',
                [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    }
})

#==============================================================================
# HANDLER: Esporta PDF
#==============================================================================
$btnPDF.add_Click({
    if ($script:RisultatiGlobali.Count -eq 0) { return }

    # Cerca stampante PDF installata (nome varia in base alla lingua di Windows)
    $nomePDFStampante = $null
    foreach ($p in [System.Drawing.Printing.PrinterSettings]::InstalledPrinters) {
        if ($p -match 'PDF') { $nomePDFStampante = $p; break }
    }
    if (-not $nomePDFStampante) {
        [System.Windows.Forms.MessageBox]::Show(
            "Nessuna stampante PDF trovata sul sistema.`nInstallare 'Microsoft Stampa su PDF' dalle impostazioni di Windows.",
            'Stampante PDF non trovata',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Title  = 'Salva congiunzioni come PDF'
    $sfd.Filter = 'File PDF (*.pdf)|*.pdf|Tutti (*.*)|*.*'
    $sfd.FileName = 'congiunzioni_' + [DateTime]::Today.ToString('yyyyMMdd') + '.pdf'
    $sfd.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($sfd.ShowDialog() -ne 'OK') { return }
    $pdfPath = $sfd.FileName

    # Colonne e chiavi per il report
    $pCols = @(
        [PSCustomObject]@{Titolo='Data e Ora (UTC)';     Campo='Data'; W=90}
        [PSCustomObject]@{Titolo='Corpo 1';              Campo='C1';   W=55}
        [PSCustomObject]@{Titolo='Corpo 2';              Campo='C2';   W=55}
        [PSCustomObject]@{Titolo='Separazione';          Campo='Sep';  W=70}
        [PSCustomObject]@{Titolo='Elev. Corpo 1';        Campo='El1';  W=60}
        [PSCustomObject]@{Titolo='Elev. Corpo 2';        Campo='El2';  W=60}
        [PSCustomObject]@{Titolo='Mag. C1';              Campo='Mag1'; W=50}
        [PSCustomObject]@{Titolo='Mag. C2';              Campo='Mag2'; W=50}
        [PSCustomObject]@{Titolo='Elev. Sole';           Campo='Sole'; W=60}
        [PSCustomObject]@{Titolo='Visibilita'' e Condizioni'; Campo='Stato'; W=160}
    )
    $totW = ($pCols | Measure-Object W -Sum).Sum

    $script:PdfPageIdx = 0
    $script:PdfDati    = $script:RisultatiGlobali

    $pd = New-Object System.Drawing.Printing.PrintDocument
    $pd.PrinterSettings.PrinterName = $nomePDFStampante
    $pd.PrinterSettings.PrintToFile = $true
    $pd.PrinterSettings.PrintFileName = $pdfPath
    $pd.DefaultPageSettings.Landscape = $true

    $pd.add_PrintPage({
        param($s, $e)
        $g     = $e.Graphics
        $marg  = $e.MarginBounds
        $fTit  = New-Object System.Drawing.Font('Arial', 13, [System.Drawing.FontStyle]::Bold)
        $fSub  = New-Object System.Drawing.Font('Arial',  8)
        $fHdr  = New-Object System.Drawing.Font('Arial',  8, [System.Drawing.FontStyle]::Bold)
        $fDat  = New-Object System.Drawing.Font('Arial',  7)
        $sfmt  = New-Object System.Drawing.StringFormat
        $sfmt.Alignment    = 'Near'
        $sfmt.LineAlignment= 'Center'
        $sfmt.Trimming     = 'EllipsisCharacter'
        $sfmtR = New-Object System.Drawing.StringFormat
        $sfmtR.Alignment    = 'Far'
        $sfmtR.LineAlignment= 'Center'
        $sfmtR.Trimming     = 'EllipsisCharacter'

        $x = $marg.Left; $y = $marg.Top; $pageW = $marg.Width

        # Calcola larghezze colonne proporzionali alla pagina
        $cW = $pCols | ForEach-Object { [int]($_.W / $totW * $pageW) }

        # Intestazione pagina (solo prima pagina)
        if ($script:PdfPageIdx -eq 0) {
            $g.DrawString('Congiunzioni Planetarie Visibili  --  JPL Horizons', $fTit,
                [System.Drawing.Brushes]::Black, $x, $y)
            $y += 22
            $t0v = $dtpI.Value.ToString('dd/MM/yyyy'); $t1v = $dtpF.Value.ToString('dd/MM/yyyy')
            $g.DrawString("Posizione: Lon $($txtLon.Text)  Lat $($txtLat.Text)  Alt $($txtAlt.Text) km   |   Periodo: $t0v - $t1v   |   $($script:PdfDati.Count) congiunzioni trovate", $fSub,
                [System.Drawing.Brushes]::DimGray, $x, $y)
            $y += 18
        }

        # Riga di intestazione colonne
        $hdrH = 18
        $hdrBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(30,90,180))
        $g.FillRectangle($hdrBrush, $x, $y, $pageW, $hdrH)
        $cx = $x
        for ($i=0; $i -lt $pCols.Count; $i++) {
            $r = New-Object System.Drawing.RectangleF(($cx+2), $y, ($cW[$i]-2), $hdrH)
            $g.DrawString($pCols[$i].Titolo, $fHdr, [System.Drawing.Brushes]::White, $r, $sfmt)
            $cx += $cW[$i]
        }
        $y += $hdrH

        # Righe dati
        $rigaH  = 14
        $altBr  = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(240,247,255))
        $nRiga  = 0
        $penGr  = New-Object System.Drawing.Pen([System.Drawing.Color]::LightGray, 0.5)

        while ($script:PdfPageIdx -lt $script:PdfDati.Count) {
            if (($y + $rigaH) -gt $marg.Bottom) { $e.HasMorePages = $true; break }

            $c = $script:PdfDati[$script:PdfPageIdx]

            # Sfondo alternato o colorato dalla visibilita'
            $rowBr = New-Object System.Drawing.SolidBrush($c.BG)
            $g.FillRectangle($rowBr, $x, $y, $pageW, $rigaH)
            $rowBr.Dispose()

            $vals = @($c.Data, $c.C1, $c.C2, $c.Sep, $c.El1, $c.El2, $c.Mag1, $c.Mag2, $c.Sole, $c.Stato)
            $cx = $x
            for ($i=0; $i -lt $vals.Count; $i++) {
                $r = New-Object System.Drawing.RectangleF(($cx+2), $y, ($cW[$i]-3), $rigaH)
                $g.DrawString([string]$vals[$i], $fDat, [System.Drawing.Brushes]::Black, $r, $sfmt)
                $cx += $cW[$i]
            }
            $g.DrawLine($penGr, $x, ($y+$rigaH), ($x+$pageW), ($y+$rigaH))

            $y += $rigaH; $nRiga++; $script:PdfPageIdx++
        }

        # Bordo tabella
        $g.DrawRectangle([System.Drawing.Pens]::Gray, $x, ($marg.Top + $(if($script:PdfPageIdx -le $nRiga){40}else{0})), $pageW, ($y - $marg.Top - $(if($script:PdfPageIdx -le $nRiga){40}else{0})))

        # Footer pagina
        $g.DrawString("Generato il $([DateTime]::Now.ToString('dd/MM/yyyy HH:mm'))  --  JPL Horizons API",
            $fSub, [System.Drawing.Brushes]::Gray,
            (New-Object System.Drawing.RectangleF($x, ($marg.Bottom-14), $pageW, 14)), $sfmtR)

        foreach ($obj in @($fTit,$fSub,$fHdr,$fDat,$sfmt,$sfmtR,$hdrBrush,$altBr,$penGr)) {
            try { $obj.Dispose() } catch {}
        }
    })

    try {
        $lSt.Text = "Generazione PDF in corso..."
        $lSt.ForeColor = [System.Drawing.Color]::DarkBlue
        [System.Windows.Forms.Application]::DoEvents()
        $pd.Print()
        $lSt.Text = "PDF salvato: $pdfPath"
        $lSt.ForeColor = [System.Drawing.Color]::DarkGreen
        [System.Windows.Forms.MessageBox]::Show("PDF salvato con successo:`n$pdfPath", 'Esportazione PDF',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Errore generazione PDF:`n$($_.Exception.Message)", 'Errore PDF',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        $lSt.Text = 'Errore PDF: ' + $_.Exception.Message
        $lSt.ForeColor = [System.Drawing.Color]::Red
    } finally {
        $pd.Dispose()
    }
})


[System.Windows.Forms.Application]::Run($frm)

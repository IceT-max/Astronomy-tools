#Requires -Version 5.1
<#
.SYNOPSIS
    JPL Horizons  --  GUI Universale per Ricerche e Lookup
    API Effemeridi : https://ssd.jpl.nasa.gov/api/horizons.api
    API Lookup     : https://ssd-api.jpl.nasa.gov/sbdb.api  +  Horizons disambiguazione
    Richiede connessione Internet.
.NOTES
    Caratteri ASCII standard. Solo Consolas. PowerShell 5.1.
    Tab: [1] Lookup  [2] Osservatore  [3] Vettori  [4] Elementi  [5] Avvicinamenti  [6] Risultati
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
Set-StrictMode -Off

#==============================================================================
# COSTANTI GLOBALI
#==============================================================================
$script:URL_EPHEM  = 'https://ssd.jpl.nasa.gov/api/horizons.api'
$script:URL_SBDB   = 'https://ssd-api.jpl.nasa.gov/sbdb.api'  # corpi minori
$script:SelectedCmd = ''

$script:OGGETTI_NOTI = [ordered]@{
    '--- PIANETI ---'       = ''
    'Sole'                  = '10'
    'Mercurio'              = '199'
    'Venere'                = '299'
    'Terra (baricentro)'    = '399'
    'Luna'                  = '301'
    'Marte'                 = '499'
    'Giove'                 = '599'
    'Saturno'               = '699'
    'Urano'                 = '799'
    'Nettuno'               = '899'
    'Plutone'               = '999'
    '--- PIANETI NANI ---'  = ''
    'Cerere'                = '2000001'
    'Eris'                  = '2136199'
    'Makemake'              = '2136472'
    'Haumea'                = '2136108'
    '--- ASTEROIDI ---'     = ''
    'Pallade (2)'           = '2000002'
    'Giunone (3)'           = '2000003'
    'Vesta (4)'             = '2000004'
    'Iride (7)'             = '2000007'
    'Flora (8)'             = '2000008'
    'Igea (10)'             = '2000010'
    'Chirone'               = '2002060'
    'Bennu (OSIRIS-REx)'    = '2101955'
    'Itokawa (Hayabusa)'    = '2025143'
    'Ryugu (Hayabusa2)'     = '2162173'
    '--- COMETE ---'        = ''
    '1P/Halley'             = 'DES=1P;'
    '2P/Encke'              = 'DES=2P;'
    '9P/Tempel 1'           = 'DES=9P;'
    '67P/C-G (Rosetta)'     = 'DES=67P;'
    '81P/Wild 2 (Stardust)' = 'DES=81P;'
    'C/1995 O1 Hale-Bopp'   = 'DES=C/1995 O1;'
    'C/2020 F3 NEOWISE'     = 'DES=C/2020 F3;'
    'C/2023 A3 Tsuchinshan' = 'DES=C/2023 A3;'
    '--- SONDE ---'         = ''
    'ISS (staz. spaziale)'  = '-125544'
    'Hubble (HST)'          = '-48'
    'Voyager 1'             = '-31'
    'Voyager 2'             = '-32'
    'New Horizons'          = '-98'
    'Cassini'               = '-82'
    'Juno'                  = '-61'
    'Parker Solar Probe'    = '-96'
    'BepiColombo'           = '-121'
    'JWST'                  = '-170'
    'GAIA'                  = '-137350'
    'Rosetta'               = '-226'
}

$script:QTY_LIST = @(
    '1  - Asc. Retta e Declinazione (J2000 astrometrica)'
    '2  - Asc. Retta e Dec. (apparente, correzioni FK5)'
    '3  - Rif. al piano virtuale (lat/lon virtuale)'
    '4  - Azimut e Altezza (topocentrici)'
    '6  - Dist. dal Sole (r, rdot, dRA/dt, dDec/dt)'
    '7  - Dist. dal Sole e vel. angolare'
    '8  - Dist. geocentrica (Delta, dDelta/dt)'
    '9  - Magnitudine visuale V e brillanza superficie'
    '10 - Dimensione angolare del disco (arcsec)'
    '11 - Angolo di posizione (PA) asse di rotazione'
    '13 - Angolo di fase del Sole'
    '14 - Elongazione dal Sole e sign. E/W'
    '15 - Visibilita'' notturna consigliata'
    '17 - Angolo eliocentrico (long. eliocentrica)'
    '19 - Variazione distanza Delta (km/s)'
    '20 - Velocita'' radiale osservatore-bersaglio'
    '23 - Distanza dal piano dell''eclittica (AU)'
    '24 - Distanza dal piano galattico (kpc)'
    '25 - Traiettoria apparente (d RA, d Dec)'
    '27 - Asc. Retta e Dec. galattiche'
    '29 - Coord. eclittiche apparenti (lon, lat)'
    '31 - Distanza baricentro SS (AU, dAU/dt)'
    '33 - Coord. galattiche (l, b)'
    '36 - Body-frame angles target-sun'
    '39 - Range e rdot (baricentrico)'
    '41 - Coord. eclitt. eliocentrliche (lon, lat)'
    '42 - Coord. equat. eliocentriche (RA, Dec)'
    '43 - Coord. cartesiane geocentriche (X, Y, Z)'
    '45 - Coord. equatoriali geocentriche (RA, Dec)'
    '47 - Coord. eclittiche geocentriche (lon, lat)'
    '48 - Rifrazi. in AR e Dec (arcsec)'
    '50 - Fase lunare e illuminazione'
)

#==============================================================================
# HELPERS FORM
#==============================================================================
function mkLbl([string]$t, [int]$x, [int]$y, [int]$w=130, [int]$h=20) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $t; $l.Location = New-Object System.Drawing.Point($x,$y)
    $l.Size = New-Object System.Drawing.Size($w,$h); $l.TextAlign = 'MiddleLeft'
    return $l
}
function mkTxt([int]$x, [int]$y, [int]$w=120, [string]$d='') {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point($x,$y)
    $t.Size = New-Object System.Drawing.Size($w,22); $t.Text = $d
    return $t
}
function mkBtn([string]$t, [int]$x, [int]$y, [int]$w=110, [int]$h=26) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $t; $b.Location = New-Object System.Drawing.Point($x,$y)
    $b.Size = New-Object System.Drawing.Size($w,$h); $b.FlatStyle = 'Flat'
    return $b
}
function mkGB([string]$t, [int]$x, [int]$y, [int]$w, [int]$h) {
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text = $t; $g.Location = New-Object System.Drawing.Point($x,$y)
    $g.Size = New-Object System.Drawing.Size($w,$h)
    return $g
}
function mkCmb([int]$x, [int]$y, [int]$w=120, [string[]]$items=@(), [int]$sel=0) {
    $c = New-Object System.Windows.Forms.ComboBox
    $c.Location = New-Object System.Drawing.Point($x,$y)
    $c.Size = New-Object System.Drawing.Size($w,22); $c.DropDownStyle = 'DropDownList'
    foreach ($i in $items) { $c.Items.Add($i) | Out-Null }
    if ($c.Items.Count -gt 0 -and $sel -lt $c.Items.Count) { $c.SelectedIndex = $sel }
    return $c
}
function mkChk([string]$t, [int]$x, [int]$y, [int]$w=200, [bool]$chk=$false) {
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text = $t; $c.Location = New-Object System.Drawing.Point($x,$y)
    $c.Size = New-Object System.Drawing.Size($w,20); $c.Checked = $chk
    return $c
}
function mkDTP([int]$x, [int]$y, [datetime]$date) {
    $d = New-Object System.Windows.Forms.DateTimePicker
    $d.Location = New-Object System.Drawing.Point($x,$y)
    $d.Size = New-Object System.Drawing.Size(155,22); $d.Format = 'Short'; $d.Value = $date
    return $d
}
function mkRtb([int]$x, [int]$y, [int]$w, [int]$h, [bool]$ro=$true, [string]$bg='White') {
    $r = New-Object System.Windows.Forms.RichTextBox
    $r.Location = New-Object System.Drawing.Point($x,$y); $r.Size = New-Object System.Drawing.Size($w,$h)
    $r.Font = New-Object System.Drawing.Font('Consolas',8)
    $r.ReadOnly = $ro; $r.ScrollBars = 'Both'; $r.WordWrap = $false
    $r.BackColor = [System.Drawing.Color]::FromName($bg)
    return $r
}

#==============================================================================
# FUNZIONI HELPER
#==============================================================================
function Enc([string]$s) { return [System.Uri]::EscapeDataString($s) }

function Get-Web([string]$url) {
    $wc = New-Object System.Net.WebClient
    $wc.Headers.Add('User-Agent', 'PS-HorizonsGUI/3.0 (PowerShell 5.1; Windows)')
    $wc.Encoding = [System.Text.Encoding]::UTF8
    return $wc.DownloadString($url)
}

function SetStatus([string]$msg, [string]$col = 'Black') {
    $lStatus.Text = $msg
    $lStatus.ForeColor = [System.Drawing.Color]::FromName($col)
    [System.Windows.Forms.Application]::DoEvents()
}

function SetCmd([string]$id) {
    if ($id -eq '') { return }
    $script:SelectedCmd = $id
    $txtCmd.Text   = $id
    $txtCmdV.Text  = $id
    $txtCmdE.Text  = $id
    $txtCmdCA.Text = $id
    $lSelObj.Text  = "Oggetto attivo: $id"
    $lSelObj.ForeColor = [System.Drawing.Color]::DarkBlue
}

function ShowResult([string]$url, [string]$testo) {
    $rtbUrl.Text = $url
    $rtbOut.Text = $testo
    $tabControl.SelectedIndex = 5
}

#==============================================================================
# FUNZIONI API
#==============================================================================

function Do-Lookup {
    # Due strategie:
    #  1) sbdb.api   -> corpi minori (asteroidi, comete)  JSON pulito
    #  2) horizons.api?format=json  -> major bodies, con parsing risposta disambiguazione
    $q = $txtSearch.Text.Trim()
    if ($q -eq '') { return }
    SetStatus "Ricerca: '$q' ..." 'DarkBlue'
    $lvResults.Items.Clear(); $rtbLookupDetail.Clear()

    $lista = [System.Collections.Generic.List[PSObject]]::new()
    $dbg   = [System.Text.StringBuilder]::new()

    # -----------------------------------------------------------------------
    # STRATEGIA 1: Small Body Database  (asteroidi, comete, pianeti nani)
    #   https://ssd-api.jpl.nasa.gov/sbdb.api?sstr=TERM&full-prec=false
    # -----------------------------------------------------------------------
    try {
        $urlSB = "https://ssd-api.jpl.nasa.gov/sbdb.api?sstr=$(Enc $q)&full-prec=false"
        [void]$dbg.AppendLine("[SBDB] $urlSB")
        $jsonSB = Get-Web $urlSB
        $dSB    = $jsonSB | ConvertFrom-Json

        if ($dSB.list -and $dSB.list.Count -gt 0) {
            # Risposta con lista multipla: {"count":N,"list":[...]}
            foreach ($r in $dSB.list) {
                $lista.Add([PSCustomObject]@{
                    pdes  = if ($r.pdes) { [string]$r.pdes } elseif ($r.des) { [string]$r.des } else { '' }
                    name  = if ($r.name) { [string]$r.name } else { '' }
                    spkid = if ($r.spkid){ [string]$r.spkid} else { '' }
                    kind  = 'SB'
                    alias = ''
                    raw   = $r
                })
            }
            [void]$dbg.AppendLine("[SBDB] lista: $($dSB.list.Count) oggetti")
        } elseif ($dSB.object) {
            # Risposta singola: {"object":{...}}
            # Campi nome: name, shortname, fullname, des (in ordine di preferenza)
            $o = $dSB.object
            $nSB = ''; $pSB = ''
            foreach ($campo in @('name','shortname','fullname','des')) {
                if ($o.$campo -and ([string]$o.$campo).Trim() -ne '') { $nSB = [string]$o.$campo; break }
            }
            foreach ($campo in @('des','pdes','id','spkid')) {
                if ($o.$campo -and ([string]$o.$campo).Trim() -ne '') { $pSB = [string]$o.$campo; break }
            }
            $spkSB = if ($o.spkid -and [string]$o.spkid -ne '') { [string]$o.spkid } else { $pSB }
            $aliSB = ''
            if ($o.shortname -and [string]$o.shortname -ne $nSB) { $aliSB = [string]$o.shortname }
            $lista.Add([PSCustomObject]@{
                pdes = $pSB; name = $nSB; spkid = $spkSB
                kind = if ($o.kind) { [string]$o.kind } else { 'SB' }
                alias = $aliSB; raw = $o
            })
            $campi = ($o | Get-Member -MemberType NoteProperty).Name -join ', '
            [void]$dbg.AppendLine("[SBDB] singolo: nome='$nSB'  des='$pSB'  campi=[$campi]")
        } else {
            [void]$dbg.AppendLine("[SBDB] nessun risultato")
        }
    } catch {
        [void]$dbg.AppendLine("[SBDB] errore: $($_.Exception.Message)")
    }

    # -----------------------------------------------------------------------
    # STRATEGIA 2: Horizons API con format=json
    #   COMMAND='TERM'  MAKE_EPHEM=NO  OBJ_DATA=YES
    #   Se il nome e' ambiguo, la risposta contiene la tabella di disambiguazione.
    #   Se e' univoco, restituisce i dati dell'oggetto.
    # -----------------------------------------------------------------------
    try {
        $urlH = "$($script:URL_EPHEM)?format=json&COMMAND=$(Enc "'$q'")&MAKE_EPHEM=NO&OBJ_DATA=YES"
        [void]$dbg.AppendLine("[Horizons] $urlH")
        $jsonH = Get-Web $urlH
        $dH    = $jsonH | ConvertFrom-Json
        $txt   = if ($dH.result) { [string]$dH.result } else { '' }
        $preview = if ($txt.Length -gt 300) { $txt.Substring(0,300) } else { $txt }
        [void]$dbg.AppendLine("[Horizons] result len=$($txt.Length) chars")
        [void]$dbg.AppendLine("[Horizons] risposta (primi 300 char):`r`n$preview")

        # ------------------------------------------------------------------
        # PARSING RISPOSTA HORIZONS
        # Il testo in $txt puo' essere:
        #   A) Oggetto univoco  -> contiene "Target body name: NOME (ID)"
        #   B) Disambiguazione  -> contiene tabella "ID#  Nome  ..."
        #   C) Errore/no match  -> contiene messaggio di errore
        # ------------------------------------------------------------------

        # Caso A: match univoco con OBJ_DATA=YES
        # Due formati possibili da Horizons:
        #
        # A1) Output effemeride:  "Target body name: Sun (10)"
        #
        # A2) Output OBJ_DATA:   " Revised: July 31, 2013                  Sun                                 10"
        #     La riga ha 3+ spazi come separatori: [Revised: DATE] [NOME] [ID o ID/ID]
        $nFoundA = 0

        # A1: "Target body name: NOME (ID)"
        $reTarget = [regex]'Target body name\s*:\s*(.+?)\s+\((-?\d+)\)'
        foreach ($m in $reTarget.Matches($txt)) {
            $mbName = $m.Groups[1].Value.Trim()
            $mbId   = $m.Groups[2].Value.Trim()
            $gia = $lista | Where-Object { $_.spkid -eq $mbId -or $_.pdes -eq $mbId }
            if (-not $gia) {
                $lista.Add([PSCustomObject]@{
                    pdes = $mbId; name = $mbName; spkid = $mbId; kind = 'MB'; alias = ''; raw = $null
                })
                $nFoundA++
            }
        }

        # A2: riga "Revised: DATE   NOME   ID"  (formato OBJ_DATA senza effemeride)
        foreach ($riga in ($txt -split "`n")) {
            if ($riga -notmatch 'Revised:') { continue }
            # Dividi per 3+ spazi consecutivi: parte[0]=Revised, parte[1]=Nome, parte[2]=ID
            $parti = [regex]::Split($riga.Trim(), '\s{3,}')
            if ($parti.Count -lt 3) { continue }
            $mbName = $parti[1].Trim()
            # L'ID puo' essere "10", "499 / 499" o "-72" -> prendiamo il primo numero
            $idRaw  = $parti[2].Trim()
            if ($idRaw -notmatch '(-?\d+)') { continue }
            $mbId = $Matches[1]
            # Scarta se il "nome" e' vuoto, troppo corto, o sembra una data
            if ($mbName.Length -lt 2 -or $mbName -match '^\d') { continue }
            $gia = $lista | Where-Object { $_.spkid -eq $mbId -or $_.pdes -eq $mbId }
            if (-not $gia) {
                $lista.Add([PSCustomObject]@{
                    pdes = $mbId; name = $mbName; spkid = $mbId; kind = 'MB'; alias = ''; raw = $null
                })
                $nFoundA++
            }
        }

        [void]$dbg.AppendLine("[Horizons] CasoA (Target body / Revised): $nFoundA trovati")

        # Caso B: tabella disambiguazione major bodies e sonde
        # Formato righe:  "   -72      ARTEMIS-P1 (THEMIS-B)"
        #                 "   499      Mars Barycenter"
        # Regole:
        #   - Spazi iniziali (qualsiasi numero)
        #   - ID intero (positivo o negativo)
        #   - Almeno 2 spazi
        #   - Nome: qualsiasi carattere fino a fine riga
        # Escludi: righe header con "---", "ID#", "Name", righe vuote
        $nFoundB = 0
        $intestazioni = @('-------','ID#','Name','Designation','IAU','aliases')
        foreach ($riga in ($txt -split "`n")) {
            $r = $riga.TrimEnd()
            if ($r.Trim() -eq '') { continue }
            # Salta righe header/separatori
            $isHeader = $false
            foreach ($h in $intestazioni) { if ($r -match [regex]::Escape($h)) { $isHeader = $true; break } }
            if ($isHeader) { continue }

            # Regex principale: spazi + ID + spazi + nome
            if ($r -match '^\s+(-?\d+)\s{2,}(\S.*)$') {
                $mbId   = $Matches[1].Trim()
                $mbName = $Matches[2].Trim()
                if ($mbName.Length -lt 2) { continue }
                $gia = $lista | Where-Object { $_.spkid -eq $mbId -or $_.pdes -eq $mbId }
                if (-not $gia) {
                    # Rimuovi eventuale colonna "Designation" e "IAU/aliases" in coda
                    # (sono separate da 3+ spazi nella tabella Horizons)
                    $nomeClean = ($mbName -split '\s{3,}')[0].Trim()
                    if ($nomeClean.Length -lt 2) { $nomeClean = $mbName }
                    $lista.Add([PSCustomObject]@{
                        pdes = $mbId; name = $nomeClean; spkid = $mbId; kind = 'MB'; alias = ''; raw = $null
                    })
                    $nFoundB++
                }
            }
        }
        [void]$dbg.AppendLine("[Horizons] CasoB (tabella MB): $nFoundB trovati")

        # Caso C: small body disambiguation da Horizons (formato DES=...)
        $nFoundC = 0
        foreach ($riga in ($txt -split "`n")) {
            if ($riga -match 'DES=([^;]+);') {
                $des = $Matches[1].Trim()
                $gia = $lista | Where-Object { $_.pdes -eq $des }
                if (-not $gia) {
                    $lista.Add([PSCustomObject]@{
                        pdes = $des; name = $riga.Trim(); spkid = ''; kind = 'SB-DES'; alias = ''; raw = $null
                    })
                    $nFoundC++
                }
            }
        }
        [void]$dbg.AppendLine("[Horizons] CasoC (SB-DES): $nFoundC trovati")

    } catch {
        [void]$dbg.AppendLine("[Horizons] errore: $($_.Exception.Message)")
    }

    # -----------------------------------------------------------------------
    # Popola ListView
    # -----------------------------------------------------------------------
    [void]$dbg.AppendLine("[UI] lista.Count = $($lista.Count)")
    try {
        if ($lista.Count -gt 0) {
            foreach ($r in $lista) {
                # Garantisce che tutti i campi siano stringhe non-null
                $col0 = if ($r.pdes)  { [string]$r.pdes  } else { '' }
                $col1 = if ($r.name)  { [string]$r.name  } else { $col0 }
                $col2 = if ($r.spkid) { [string]$r.spkid } else { $col0 }
                $col3 = if ($r.kind)  { [string]$r.kind  } else { '' }
                $col4 = if ($r.alias) { [string]$r.alias } else { '' }
                # Usa il nome come label principale se la designazione e' solo un numero
                $label = if ($col1 -and $col1 -ne $col0) { "$col0  ($col1)" } else { $col0 }
                $lvi = New-Object System.Windows.Forms.ListViewItem($label)
                $lvi.SubItems.Add($col1) | Out-Null
                $lvi.SubItems.Add($col2) | Out-Null
                $lvi.SubItems.Add($col3) | Out-Null
                $lvi.SubItems.Add($col4) | Out-Null
                $lvi.Tag = $r
                $lvResults.Items.Add($lvi) | Out-Null
                [void]$dbg.AppendLine("[UI] aggiunto: $label")
            }
            SetStatus "$($lista.Count) oggetti trovati per '$q'." 'DarkGreen'
        } else {
            SetStatus "Nessun risultato per '$q'. Prova un ID diretto (es: 499, DES=1P;, -48)." 'DarkOrange'
        }
    } catch {
        [void]$dbg.AppendLine("[UI] ERRORE ListView: $($_.Exception.Message)")
        SetStatus "Errore visualizzazione: $($_.Exception.Message)" 'Red'
    }
    # Mostra log debug nel pannello dettaglio
    $rtbLookupDetail.Text = $dbg.ToString()
}
function Do-ObjInfo([string]$cmd) {
    if ($cmd -eq '') { [System.Windows.Forms.MessageBox]::Show('Nessun COMMAND selezionato.','Attenzione') | Out-Null; return }
    SetStatus "Recupero dati oggetto: $cmd ..." 'DarkBlue'
    try {
        $p   = "format=text&COMMAND=$(Enc "'$cmd'")&MAKE_EPHEM=NO&OBJ_DATA=YES"
        $url = "$($script:URL_EPHEM)?$p"
        $r   = Get-Web $url
        ShowResult $url $r
        SetStatus "Dati oggetto '$cmd' ricevuti." 'DarkGreen'
    } catch {
        ShowResult '' "ERRORE:`n$($_.Exception.Message)"
        SetStatus "Errore: $($_.Exception.Message)" 'Red'
    }
}

function Do-Observer {
    $cmd = $txtCmd.Text.Trim()
    if ($cmd -eq '') { [System.Windows.Forms.MessageBox]::Show('Inserire un COMMAND.','Attenzione') | Out-Null; return }
    SetStatus "Effemeridi osservatore per: $cmd ..." 'DarkBlue'
    try {
        $site = "$($txtLon.Text),$($txtLat.Text),$($txtAlt.Text)"
        $t0   = $dtpStart.Value.ToString('yyyy-MM-dd')
        $t1   = $dtpStop.Value.ToString('yyyy-MM-dd')
        $step = "$($txtStep.Text)$($cmbStepUnit.SelectedItem)"
        $qlist = @()
        foreach ($item in $clbQty.CheckedItems) { $qlist += ($item -split '\s+')[0] }
        if ($qlist.Count -eq 0) { $qlist = @('1','4','9') }
        $qty      = $qlist -join ','
        $apparent = if ($chkAirless.Checked)  { 'AIRLESS' } else { 'REFRACTED' }
        $angfmt   = if ($chkDeg.Checked)      { 'DEG'     } else { 'HMS' }
        $csv      = if ($chkCSVobs.Checked)   { 'YES'     } else { 'NO' }
        $objdata  = if ($chkObjData.Checked)  { 'YES'     } else { 'NO' }
        $calfmt   = if ($chkCalBoth.Checked)  { 'BOTH'    } else { 'CAL' }
        $extra = ''
        if ($chkExtraPrec.Checked) { $extra += '&EXTRA_PREC=YES' }

        $params = @(
            "format=text"
            "COMMAND=$(Enc "'$cmd'")"
            "MAKE_EPHEM=YES"
            "EPHEM_TYPE=OBSERVER"
            "CENTER=coord@399"
            "COORD_TYPE=GEODETIC"
            "SITE_COORD=$(Enc "'$site'")"
            "START_TIME=$(Enc "'$t0'")"
            "STOP_TIME=$(Enc "'$t1'")"
            "STEP_SIZE=$(Enc "'$step'")"
            "QUANTITIES=$(Enc "'$qty'")"
            "ANG_FORMAT=$angfmt"
            "APPARENT=$apparent"
            "CSV_FORMAT=$csv"
            "OBJ_DATA=$objdata"
            "CAL_FORMAT=$calfmt"
            "TIME_DIGITS=MINUTES"
        )
        $url = $script:URL_EPHEM + '?' + ($params -join '&') + $extra
        $r   = Get-Web $url
        ShowResult $url $r
        $nRighe = ($r -split "`n").Count
        SetStatus "Effemeridi ricevute ($nRighe righe). COMMAND=$cmd  step=$step  qty=$qty" 'DarkGreen'
    } catch {
        ShowResult '' "ERRORE:`n$($_.Exception.Message)"
        SetStatus "Errore: $($_.Exception.Message)" 'Red'
    }
}

function Do-Vectors {
    $cmd = $txtCmdV.Text.Trim()
    if ($cmd -eq '') { [System.Windows.Forms.MessageBox]::Show('Inserire un COMMAND.','Attenzione') | Out-Null; return }
    SetStatus "Vettori stato per: $cmd ..." 'DarkBlue'
    try {
        $t0     = $dtpStartV.Value.ToString('yyyy-MM-dd')
        $t1     = $dtpStopV.Value.ToString('yyyy-MM-dd')
        $step   = "$($txtStepV.Text)$($cmbStepUnitV.SelectedItem)"
        $center = $txtCenterV.Text.Trim()
        $rp     = $cmbRefPlane.SelectedItem
        $vt     = ($cmbVecTable.SelectedItem -split '\s+')[0]
        $corr   = $cmbVecCorr.SelectedItem
        $csv    = if ($chkCSVvec.Checked) { 'YES' } else { 'NO' }
        $od     = if ($chkODvec.Checked)  { 'YES' } else { 'NO' }

        $params = @(
            "format=text"
            "COMMAND=$(Enc "'$cmd'")"
            "MAKE_EPHEM=YES"
            "EPHEM_TYPE=VECTORS"
            "CENTER=$(Enc "'$center'")"
            "START_TIME=$(Enc "'$t0'")"
            "STOP_TIME=$(Enc "'$t1'")"
            "STEP_SIZE=$(Enc "'$step'")"
            "REF_PLANE=$rp"
            "VEC_TABLE=$vt"
            "VEC_CORR=$corr"
            "OBJ_DATA=$od"
            "CSV_FORMAT=$csv"
            "OUT_UNITS=AU-D"
            "REF_SYSTEM=ICRF"
        )
        $url = $script:URL_EPHEM + '?' + ($params -join '&')
        $r   = Get-Web $url
        ShowResult $url $r
        SetStatus "Vettori stato ricevuti. COMMAND=$cmd  center=$center  step=$step" 'DarkGreen'
    } catch {
        ShowResult '' "ERRORE:`n$($_.Exception.Message)"
        SetStatus "Errore: $($_.Exception.Message)" 'Red'
    }
}

function Do-Elements {
    $cmd = $txtCmdE.Text.Trim()
    if ($cmd -eq '') { [System.Windows.Forms.MessageBox]::Show('Inserire un COMMAND.','Attenzione') | Out-Null; return }
    SetStatus "Elementi orbitali per: $cmd ..." 'DarkBlue'
    try {
        $t0     = $dtpStartE.Value.ToString('yyyy-MM-dd')
        $t1     = $dtpStopE.Value.ToString('yyyy-MM-dd')
        $step   = "$($txtStepE.Text)$($cmbStepUnitE.SelectedItem)"
        $center = $txtCenterE.Text.Trim()
        $rp     = $cmbRefPlaneE.SelectedItem
        $od     = if ($chkODelem.Checked)  { 'YES' } else { 'NO' }
        $csv    = if ($chkCSVelem.Checked) { 'YES' } else { 'NO' }

        $params = @(
            "format=text"
            "COMMAND=$(Enc "'$cmd'")"
            "MAKE_EPHEM=YES"
            "EPHEM_TYPE=ELEMENTS"
            "CENTER=$(Enc "'$center'")"
            "START_TIME=$(Enc "'$t0'")"
            "STOP_TIME=$(Enc "'$t1'")"
            "STEP_SIZE=$(Enc "'$step'")"
            "REF_PLANE=$rp"
            "REF_SYSTEM=ICRF"
            "TP_TYPE=ABSOLUTE"
            "OBJ_DATA=$od"
            "CSV_FORMAT=$csv"
        )
        $url = $script:URL_EPHEM + '?' + ($params -join '&')
        $r   = Get-Web $url
        ShowResult $url $r
        SetStatus "Elementi orbitali ricevuti. COMMAND=$cmd  center=$center  step=$step" 'DarkGreen'
    } catch {
        ShowResult '' "ERRORE:`n$($_.Exception.Message)"
        SetStatus "Errore: $($_.Exception.Message)" 'Red'
    }
}

function Do-Approach {
    $cmd = $txtCmdCA.Text.Trim()
    if ($cmd -eq '') { [System.Windows.Forms.MessageBox]::Show('Inserire un COMMAND.','Attenzione') | Out-Null; return }
    SetStatus "Close Approach per: $cmd ..." 'DarkBlue'
    try {
        $t0 = $dtpStartCA.Value.ToString('yyyy-MM-dd')
        $t1 = $dtpStopCA.Value.ToString('yyyy-MM-dd')

        $params = @(
            "format=text"
            "COMMAND=$(Enc "'$cmd'")"
            "MAKE_EPHEM=YES"
            "EPHEM_TYPE=APPROACH"
            "START_TIME=$(Enc "'$t0'")"
            "STOP_TIME=$(Enc "'$t1'")"
            "OBJ_DATA=YES"
        )
        $url = $script:URL_EPHEM + '?' + ($params -join '&')
        $r   = Get-Web $url
        ShowResult $url $r
        SetStatus "Avvicinamenti ricevuti. COMMAND=$cmd  periodo: $t0 -> $t1" 'DarkGreen'
    } catch {
        ShowResult '' "ERRORE:`n$($_.Exception.Message)"
        SetStatus "Errore: $($_.Exception.Message)" 'Red'
    }
}

#==============================================================================
# FORM PRINCIPALE
#==============================================================================
$BLU    = [System.Drawing.Color]::FromArgb(0,120,212)
$VERDE  = [System.Drawing.Color]::FromArgb(16,124,16)
$ROSSO  = [System.Drawing.Color]::FromArgb(180,40,30)
$VIOLA  = [System.Drawing.Color]::FromArgb(100,60,180)
$BIANCO = [System.Drawing.Color]::White
$GIALLO = [System.Drawing.Color]::FromArgb(255,255,200)
$NERO   = [System.Drawing.Color]::Black
$GRIGIO = [System.Drawing.Color]::FromArgb(245,245,245)

$frm = New-Object System.Windows.Forms.Form
$frm.Text            = 'JPL Horizons  --  GUI Universale  v3.1  (PS 5.1)'
$frm.ClientSize      = New-Object System.Drawing.Size(1220,890)
$frm.StartPosition   = 'CenterScreen'
$frm.FormBorderStyle = 'Sizable'
$frm.MinimumSize     = New-Object System.Drawing.Size(1000,800)
$frm.Font            = New-Object System.Drawing.Font('Consolas',9)

# TabControl principale
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(5,5)
$tabControl.Size     = New-Object System.Drawing.Size(1205,840)
$tabControl.Anchor   = 'Top,Bottom,Left,Right'
$frm.Controls.Add($tabControl)

# Barra di stato globale (sotto il tabcontrol)
$lStatus = New-Object System.Windows.Forms.Label
$lStatus.Location  = New-Object System.Drawing.Point(5,852)
$lStatus.Size      = New-Object System.Drawing.Size(1205,22)
$lStatus.Anchor    = 'Bottom,Left,Right'
$lStatus.Text      = 'Pronto. Usa [1] Lookup per cercare oggetti celesti.'
$lStatus.Font      = New-Object System.Drawing.Font('Consolas',8)
$lStatus.ForeColor = [System.Drawing.Color]::DarkGray
$frm.Controls.Add($lStatus)

# Oggetto selezionato (sempre visibile, sopra la status bar)
$lSelObj = New-Object System.Windows.Forms.Label
$lSelObj.Location  = New-Object System.Drawing.Point(5,838)
$lSelObj.Size      = New-Object System.Drawing.Size(1205,16)
$lSelObj.Anchor    = 'Bottom,Left,Right'
$lSelObj.Text      = 'Oggetto attivo: (nessuno selezionato)'
$lSelObj.Font      = New-Object System.Drawing.Font('Consolas',8,[System.Drawing.FontStyle]::Bold)
$lSelObj.ForeColor = [System.Drawing.Color]::DarkBlue
$frm.Controls.Add($lSelObj)

# Creazione TabPage
$tabLookup   = New-Object System.Windows.Forms.TabPage; $tabLookup.Text   = ' [1] Lookup '
$tabObserver = New-Object System.Windows.Forms.TabPage; $tabObserver.Text = ' [2] Oss. Effemeridi '
$tabVectors  = New-Object System.Windows.Forms.TabPage; $tabVectors.Text  = ' [3] Vettori Stato '
$tabElements = New-Object System.Windows.Forms.TabPage; $tabElements.Text = ' [4] Elementi Orbitali '
$tabApproach = New-Object System.Windows.Forms.TabPage; $tabApproach.Text = ' [5] Avvicinamenti '
$tabOut      = New-Object System.Windows.Forms.TabPage; $tabOut.Text      = ' [6] Risultati '
foreach ($tp in @($tabLookup,$tabObserver,$tabVectors,$tabElements,$tabApproach,$tabOut)) {
    $tabControl.TabPages.Add($tp)
}

#==============================================================================
# [TAB 1] LOOKUP
#==============================================================================
$p1 = $tabLookup

# --- Riga ricerca ---
$p1.Controls.Add((mkLbl 'Cerca oggetto:' 8 14 105 22))
$txtSearch = mkTxt 116 12 280 ''
$txtSearch.Font = New-Object System.Drawing.Font('Consolas',10)
$p1.Controls.Add($txtSearch)
$chkFuzzy = mkChk 'Fuzzy match' 402 14 120 $true
$p1.Controls.Add($chkFuzzy)
$btnSearch = mkBtn 'CERCA' 530 12 90 26
$btnSearch.BackColor = $BLU; $btnSearch.ForeColor = $BIANCO
$btnSearch.Font = New-Object System.Drawing.Font('Consolas',9,[System.Drawing.FontStyle]::Bold)
$p1.Controls.Add($btnSearch)
$p1.Controls.Add((mkLbl 'Es: mars, ceres, halley, 1999 RQ36, C/2023 A3' 630 14 480 20))

# --- Riga oggetti noti ---
$p1.Controls.Add((mkLbl 'Oggetti noti:' 8 44 95 22))
$cmbKnown = New-Object System.Windows.Forms.ComboBox
$cmbKnown.Location = New-Object System.Drawing.Point(108,42)
$cmbKnown.Size     = New-Object System.Drawing.Size(220,22)
$cmbKnown.DropDownStyle = 'DropDownList'
$cmbKnown.DropDownWidth = 320
foreach ($k in $script:OGGETTI_NOTI.Keys) {
    $cmbKnown.Items.Add($k) | Out-Null
}
$cmbKnown.SelectedIndex = 0
$p1.Controls.Add($cmbKnown)
$btnSelKnown = mkBtn 'Seleziona' 336 42 90 26
$p1.Controls.Add($btnSelKnown)

# --- Riga ID diretto ---
$p1.Controls.Add((mkLbl 'ID diretto:' 8 74 88 22))
$txtDirectID = mkTxt 100 72 180 ''
$p1.Controls.Add($txtDirectID)
$btnSelDirect = mkBtn 'Usa ID' 288 72 80 26
$p1.Controls.Add($btnSelDirect)
$p1.Controls.Add((mkLbl '  (199, 499, DES=67P;, -48, -125544...)' 376 74 400 20))

# Pulsante Info oggetto
$btnInfoObj = mkBtn 'Info Oggetto' 800 12 120 26
$btnInfoObj.BackColor = $VIOLA; $btnInfoObj.ForeColor = $BIANCO
$p1.Controls.Add($btnInfoObj)
$btnUseLV = mkBtn 'USA OGGETTO SELEZIONATO >>>' 800 44 320 26
$btnUseLV.BackColor = $VERDE; $btnUseLV.ForeColor = $BIANCO
$btnUseLV.Font = New-Object System.Drawing.Font('Consolas',9,[System.Drawing.FontStyle]::Bold)
$p1.Controls.Add($btnUseLV)

$p1.Controls.Add((mkLbl '(doppio clic o pulsante per selezionare)' 800 76 360 20))

# --- Separatore ---
$sep1 = New-Object System.Windows.Forms.Label
$sep1.Location = New-Object System.Drawing.Point(0,103)
$sep1.Size = New-Object System.Drawing.Size(1200,2)
$sep1.BorderStyle = 'Fixed3D'
$p1.Controls.Add($sep1)

# --- ListView risultati ---
$gbLV = mkGB 'Risultati Lookup  (sbdb.api = corpi minori  |  Horizons disambiguazione = pianeti/sonde)' 8 110 1175 450
$p1.Controls.Add($gbLV)

$lvResults = New-Object System.Windows.Forms.ListView
$lvResults.Location    = New-Object System.Drawing.Point(8,22)
$lvResults.Size        = New-Object System.Drawing.Size(1155,255)
$lvResults.View        = 'Details'
$lvResults.FullRowSelect = $true
$lvResults.GridLines   = $true
$lvResults.MultiSelect = $false
$lvResults.Font        = New-Object System.Drawing.Font('Consolas',8)
$lvResults.BackColor   = [System.Drawing.Color]::FromArgb(250,250,255)
foreach ($colDef in @(
    @{T='Designazione / PDES'; W=140}
    @{T='Nome'; W=200}
    @{T='SPK-ID'; W=90}
    @{T='Tipo'; W=70}
    @{T='Alias / Designazioni alternate'; W=550}
)) {
    $col = New-Object System.Windows.Forms.ColumnHeader
    $col.Text = $colDef.T; $col.Width = $colDef.W
    $lvResults.Columns.Add($col) | Out-Null
}
$gbLV.Controls.Add($lvResults)

$gbLV.Controls.Add((mkLbl 'Dettaglio oggetto selezionato:' 8 284 280 18))
$rtbLookupDetail = mkRtb 8 304 1155 132 $true 'White'
$rtbLookupDetail.BackColor = [System.Drawing.Color]::FromArgb(248,248,255)
$gbLV.Controls.Add($rtbLookupDetail)

# --- Help panel ---
$gbHelp = mkGB 'Guida rapida formati COMMAND' 8 568 1175 230
$p1.Controls.Add($gbHelp)

$rtbHelp = mkRtb 8 20 1155 200 $true 'White'
$rtbHelp.BackColor = $GIALLO
$rtbHelp.Font = New-Object System.Drawing.Font('Consolas',8)
$rtbHelp.Text = @"
FORMATI COMMAND ACCETTATI DALL'API HORIZONS   (ssd.jpl.nasa.gov/api/horizons.api)
=====================================================================================================
Corpo                Formato COMMAND         Esempio           Note
---------------------+--------------------+-----------------+---------------------------------------------
Sole                 | numero intero       | 10               | 10 = Sole (baric. sistema solare = 0)
Pianeti (baric.)     | NNN                 | 499              | 499=Marte 599=Giove 699=Saturno
Pianeti (nucleo)     | NNN (corpo fisico)  | 499              | I pianeti principali usano lo stesso ID
Lune                 | NNN                 | 301=Luna 402=Deimos 501=Io 502=Europa
Asteroidi            | 2NNNNNNN            | 2000001=Cerere   | 7 cifre: 2 + n. catalogo (0-padded)
                     | NNN;                | 1;               | Ricerca per numero di catalogo
Asteroidi (des.)     | DES=YYYY XNN[N];    | DES=1999 RQ36;   | Designazione provvisoria
Comete (periodic.)   | DES=NP;             | DES=1P; DES=2P;  | N=numero catalogo, P=periodica
Comete (non-per.)    | DES=C/YYYY XX;      | DES=C/2020 F3;   | C=non-periodica, A=da det.
Navicelle spaziali   | -N  (negativo)      | -48  -82  -31    | NAIF SPICE ID (negativo)
Per nome             | stringa libera      | 'Mars'  'Ceres'  | L'API cerca per corrispondenza
=====================================================================================================
CENTRI DI OSSERVAZIONE comuni:  500=Geocentro  @10=Elio-centrico  @0=Baricentro SS  coord@399=Geodetico
"@
$gbHelp.Controls.Add($rtbHelp)

#==============================================================================
# [TAB 2] EFFEMERIDI OSSERVATORE
#==============================================================================
$p2 = $tabObserver

# Header
$p2.Controls.Add((mkLbl 'COMMAND:' 8 14 75 22))
$txtCmd = mkTxt 86 12 180 '499'
$p2.Controls.Add($txtCmd)
$p2.Controls.Add((mkLbl '(ID oggetto, es: 499=Marte  301=Luna  -48=HST)' 272 14 400 22))
$btnGoObs = mkBtn 'CALCOLA >>>' 900 12 130 30
$btnGoObs.BackColor = $BLU; $btnGoObs.ForeColor = $BIANCO
$btnGoObs.Font = New-Object System.Drawing.Font('Consolas',10,[System.Drawing.FontStyle]::Bold)
$p2.Controls.Add($btnGoObs)

# --- Colonna sinistra: Posizione + Periodo + Passo + Opzioni ---
$gbLoc = mkGB 'Posizione Osservatore (Geodetiche)' 8 50 430 115
$p2.Controls.Add($gbLoc)
$gbLoc.Controls.Add((mkLbl 'Longitudine E+ (deg):' 8 26 155))
$txtLon = mkTxt 168 24 95 '9.1895'; $gbLoc.Controls.Add($txtLon)
$gbLoc.Controls.Add((mkLbl 'Latitudine  N+ (deg):' 8 56 155))
$txtLat = mkTxt 168 54 95 '45.4654'; $gbLoc.Controls.Add($txtLat)
$gbLoc.Controls.Add((mkLbl 'Altitudine       (km):' 8 84 155))
$txtAlt = mkTxt 168 82 80 '0.122'; $gbLoc.Controls.Add($txtAlt)
$lLocHint = mkLbl '[Milano, IT]' 270 56 130 20
$lLocHint.ForeColor = [System.Drawing.Color]::Gray
$gbLoc.Controls.Add($lLocHint)

$gbPeriodo = mkGB 'Intervallo Temporale' 8 172 430 85
$p2.Controls.Add($gbPeriodo)
$gbPeriodo.Controls.Add((mkLbl 'Data Inizio:' 8 28 88))
$dtpStart = mkDTP 100 26 ([DateTime]::Today); $gbPeriodo.Controls.Add($dtpStart)
$gbPeriodo.Controls.Add((mkLbl 'Data Fine:' 8 58 88))
$dtpStop  = mkDTP 100 56 ([DateTime]::Today.AddMonths(3)); $gbPeriodo.Controls.Add($dtpStop)

$gbPasso = mkGB 'Passo Temporale' 8 264 430 58
$p2.Controls.Add($gbPasso)
$gbPasso.Controls.Add((mkLbl 'Passo:' 8 24 55))
$txtStep = mkTxt 66 22 60 '1'; $gbPasso.Controls.Add($txtStep)
$cmbStepUnit = mkCmb 132 22 60 @('d','h','m','mo','y') 0; $gbPasso.Controls.Add($cmbStepUnit)
$gbPasso.Controls.Add((mkLbl '  d=giorni  h=ore  m=minuti  mo=mesi  y=anni' 200 24 230 20))

$gbOpt = mkGB 'Opzioni Output' 8 330 430 165
$p2.Controls.Add($gbOpt)
$chkAirless   = mkChk 'Senza rifrazione atmosferica (AIRLESS)' 8 22 330 $true
$chkDeg       = mkChk 'Angoli AR/Dec in gradi (DEG, default=HMS)' 8 46 330 $true
$chkCSVobs    = mkChk 'Formato CSV (valori separati da virgola)' 8 70 330 $false
$chkObjData   = mkChk 'Includi dati fisici oggetto (OBJ_DATA=YES)' 8 94 330 $false
$chkCalBoth   = mkChk 'Data calend. + JD (CAL_FORMAT=BOTH)' 8 118 280 $true
$chkExtraPrec = mkChk 'Precisione extra (EXTRA_PREC=YES)' 8 140 280 $false
foreach ($c in @($chkAirless,$chkDeg,$chkCSVobs,$chkObjData,$chkCalBoth,$chkExtraPrec)) { $gbOpt.Controls.Add($c) }

# --- Colonna destra: Quantita' ---
$gbQty = mkGB 'Quantita'' da calcolare (QUANTITIES)' 448 50 730 445
$p2.Controls.Add($gbQty)

$gbQty.Controls.Add((mkLbl 'Ctrl+clic per multi-selezione. Default: 1, 4, 9' 8 22 420 18))
$btnQtyAll  = mkBtn 'Tutte' 440 20 70 22; $gbQty.Controls.Add($btnQtyAll)
$btnQtyNone = mkBtn 'Nessuna' 516 20 70 22; $gbQty.Controls.Add($btnQtyNone)
$btnQtyStd  = mkBtn 'Standard' 592 20 80 22; $gbQty.Controls.Add($btnQtyStd)

$clbQty = New-Object System.Windows.Forms.CheckedListBox
$clbQty.Location = New-Object System.Drawing.Point(8, 46)
$clbQty.Size     = New-Object System.Drawing.Size(708, 390)
$clbQty.Font     = New-Object System.Drawing.Font('Consolas',8)
$clbQty.BackColor = [System.Drawing.Color]::FromArgb(250,250,255)
foreach ($q in $script:QTY_LIST) { $clbQty.Items.Add($q, $false) | Out-Null }
foreach ($idx in @(0,3,8)) { $clbQty.SetItemChecked($idx,$true) }  # 1, 4, 9
$gbQty.Controls.Add($clbQty)

# Legenda quantita' (inline)
$gbQtyLegend = mkGB 'Quantita'' piu'' usate' 448 500 730 290
$p2.Controls.Add($gbQtyLegend)
$rtbQLeg = mkRtb 8 20 715 260 $true 'White'
$rtbQLeg.BackColor = $GIALLO
$rtbQLeg.Font = New-Object System.Drawing.Font('Consolas',8)
$rtbQLeg.Text = @"
QUANTITA'' PIU'' COMUNI PER OSSERVAZIONE ASTRONOMICA
====================================================================
 1 = Asc.Retta + Declinazione J2000 (astrometrica, RA Dec)
 2 = AR + Dec apparente (con refrazione, aberrazioni FK5)
 4 = Azimut + Altezza topocentrici --- FONDAMENTALE ---
 6 = Distanza dal Sole r (AU), rdot, dRA/dt, dDec/dt
 8 = Distanza geocentrica Delta (AU) + vel. avvic./allontanamento
 9 = Magnitudine visuale V + brillanza superficiale (S-brt)
10 = Dimensione angolare del disco (arcsec)
13 = Angolo di fase solare (deg)
14 = Elongazione dal Sole + sigla E/W
17 = Longitudine eliocentrica apparente del bersaglio
19 = Variazione della distanza Delta (km/s)
20 = Velocita'' radiale rel. all''osservatore (km/s)
29 = Coord. eclittiche apparenti (longitudine + latitudine)
33 = Coordinate galattiche (l, b)
39 = Range geocentrico + rdot (baricentrico)

COMBINAZIONI TIPICHE:
  Osservazione visiva    : QUANTITIES='1,4,9,10,14'
  Fotografia pianeti     : QUANTITIES='1,4,9,10,13,14'
  Asteroidi e comete     : QUANTITIES='1,4,9,6,8,19'
  Radio-telescopio       : QUANTITIES='1,4,6,8,19,20'
  Calcolo rotta/orbita   : QUANTITIES='1,6,8,29,31,39'
"@
$gbQtyLegend.Controls.Add($rtbQLeg)

#==============================================================================
# [TAB 3] VETTORI STATO
#==============================================================================
$p3 = $tabVectors

$p3.Controls.Add((mkLbl 'COMMAND:' 8 14 75 22))
$txtCmdV = mkTxt 86 12 180 '499'
$p3.Controls.Add($txtCmdV)
$p3.Controls.Add((mkLbl '(ID oggetto, es: 499  301  -48)' 272 14 380 22))
$btnGoVec = mkBtn 'CALCOLA >>>' 900 12 130 30
$btnGoVec.BackColor = $BLU; $btnGoVec.ForeColor = $BIANCO
$btnGoVec.Font = New-Object System.Drawing.Font('Consolas',10,[System.Drawing.FontStyle]::Bold)
$p3.Controls.Add($btnGoVec)

$gbVecLeft = mkGB 'Parametri Vettori Stato' 8 50 550 380
$p3.Controls.Add($gbVecLeft)

$gbVecLeft.Controls.Add((mkLbl 'Centro di riferimento:' 8 28 155))
$txtCenterV = mkTxt 166 26 130 '@sun'
$gbVecLeft.Controls.Add($txtCenterV)
$gbVecLeft.Controls.Add((mkLbl '@sun=@10  @0=baricentroSS  500@399=geocentro' 305 28 240 20))

$gbVecLeft.Controls.Add((mkLbl 'Data Inizio:' 8 60 105))
$dtpStartV = mkDTP 116 58 ([DateTime]::Today); $gbVecLeft.Controls.Add($dtpStartV)
$gbVecLeft.Controls.Add((mkLbl 'Data Fine:' 8 90 105))
$dtpStopV  = mkDTP 116 88 ([DateTime]::Today.AddMonths(1)); $gbVecLeft.Controls.Add($dtpStopV)

$gbVecLeft.Controls.Add((mkLbl 'Passo:' 8 122 55))
$txtStepV = mkTxt 66 120 60 '1'; $gbVecLeft.Controls.Add($txtStepV)
$cmbStepUnitV = mkCmb 132 120 60 @('d','h','m','mo','y') 0; $gbVecLeft.Controls.Add($cmbStepUnitV)

$gbVecLeft.Controls.Add((mkLbl 'Piano di riferimento:' 8 155 150))
$cmbRefPlane = mkCmb 160 153 140 @('ECLIPTIC','FRAME','BODY EQUATOR') 0
$gbVecLeft.Controls.Add($cmbRefPlane)

$gbVecLeft.Controls.Add((mkLbl 'Tabella vettori:' 8 188 130))
$cmbVecTable = mkCmb 140 186 230 @(
    '1  (X,Y,Z + VX,VY,VZ)',
    '2  (tipo 1 + derivate acc.)',
    '3  (solo pos. X,Y,Z)',
    '4  (solo vel. VX,VY,VZ)',
    '5  (LT, RG, RR)',
    '6  (tutti: pos.+vel.+LT+RG+RR)'
) 0
$gbVecLeft.Controls.Add($cmbVecTable)

$gbVecLeft.Controls.Add((mkLbl 'Correzione luce:' 8 222 130))
$cmbVecCorr = mkCmb 140 220 120 @('NONE','LT','LT+S') 0
$gbVecLeft.Controls.Add($cmbVecCorr)

$chkCSVvec = mkChk 'Formato CSV' 8 258 150 $false; $gbVecLeft.Controls.Add($chkCSVvec)
$chkODvec  = mkChk 'Dati oggetto' 168 258 150 $true; $gbVecLeft.Controls.Add($chkODvec)

# Info pannello destra
$gbVecInfo = mkGB 'Guida Vettori Stato' 568 50 610 380
$p3.Controls.Add($gbVecInfo)
$rtbVecInfo = mkRtb 8 20 592 350 $true 'White'
$rtbVecInfo.BackColor = $GIALLO
$rtbVecInfo.Text = @"
VETTORI DI STATO  (EPHEM_TYPE = VECTORS)
===========================================
Restituisce posizione e velocita'' del corpo
in un sistema di riferimento cartesiano.

UNITA'' DI OUTPUT (OUT_UNITS=AU-D):
  X, Y, Z     posizione in AU
  VX, VY, VZ  velocita'' in AU/giorno
  LT          tempo di luce (min)
  RG          distanza dal centro (AU)
  RR          vel. radiale (AU/d)

PIANI DI RIFERIMENTO:
  ECLIPTIC     = eclittica media J2000.0
                 (piu'' usato per pianeti)
  FRAME        = ICRF/J2000 equatoriale
                 (piu'' usato per stelle,sonde)
  BODY EQUATOR = equatore del corpo target

CENTRI COMUNI:
  @10  o  @sun   = Sole (baricentro)
  @0             = Baricentro SS
  500  o  @399   = Geocentro
  500@10          = Geocentro rif. Sole
  @599           = Giovicentro

CORREZIONE LUCE:
  NONE   = geometrica (no ritardo)
  LT     = aberrazione di luce (retarded)
  LT+S   = LT + deflessione stellare

TABELLE:
  1  = posizione + velocita'' (piu'' usata)
  2  = +derivate di secondo ordine
  3  = solo X,Y,Z
  4  = solo VX,VY,VZ
  5  = LT,RG,RR (distanza e velocita'')
  6  = tutto (pos+vel+LT+RG+RR)

SISTEMA DI RIF.: ICRF (International
Celestial Reference Frame), J2000.0
"@
$gbVecInfo.Controls.Add($rtbVecInfo)

# Esempi vettori
$gbVecEx = mkGB 'Esempi di configurazione tipica' 8 440 1170 340
$p3.Controls.Add($gbVecEx)
$rtbVecEx = mkRtb 8 20 1150 310 $true 'White'
$rtbVecEx.BackColor = $GRIGIO
$rtbVecEx.Font = New-Object System.Drawing.Font('Consolas',8)
$rtbVecEx.Text = @"
ESEMPI DI QUERY VETTORI STATO TIPICHE
======================================================================================================
Scopo                           | COMMAND | Centro  | Tabella | Piano     | Note
--------------------------------+---------+---------+---------+-----------+----------------------------
Orbita Marte intorno al Sole    | 499     | @sun    | 1       | ECLIPTIC  | Classica meccanica celeste
Orbita Luna intorno alla Terra  | 301     | 500     | 1       | FRAME     | Riferimento geocentrico
Posizione ISS rispetto Terra    | -125544 | 500     | 1       | FRAME     | Coord. cartesiane geocentr.
Orbita Giove (baricentrica)     | 599     | @0      | 6       | ECLIPTIC  | Tutti i campi
Cassini (flyby Saturno)         | -82     | @699    | 1       | ECLIPTIC  | Centro = Saturno
New Horizons (interstellare)    | -98     | @sun    | 2       | FRAME     | +accelerazioni
Cometa di Halley                | DES=1P; | @sun    | 1       | ECLIPTIC  | Orbita eliocentrica
Bennu (OSIRIS-REx target)       | 2101955 | @sun    | 1       | ECLIPTIC  | NEA

NOTA: Per le sonde spaziali usare il NAIF SPICE ID (numero negativo). Per trovare l''ID usare il
      tab [1] Lookup. I pianeti e i corpi minori usano ID positivi (pianeti baricentrici = NNN).
"@
$gbVecEx.Controls.Add($rtbVecEx)

#==============================================================================
# [TAB 4] ELEMENTI ORBITALI
#==============================================================================
$p4 = $tabElements

$p4.Controls.Add((mkLbl 'COMMAND:' 8 14 75 22))
$txtCmdE = mkTxt 86 12 180 '499'
$p4.Controls.Add($txtCmdE)
$p4.Controls.Add((mkLbl '(ID oggetto, es: 499  DES=1P;  2101955)' 272 14 380 22))
$btnGoElem = mkBtn 'CALCOLA >>>' 900 12 130 30
$btnGoElem.BackColor = $BLU; $btnGoElem.ForeColor = $BIANCO
$btnGoElem.Font = New-Object System.Drawing.Font('Consolas',10,[System.Drawing.FontStyle]::Bold)
$p4.Controls.Add($btnGoElem)

$gbElemLeft = mkGB 'Parametri Elementi Orbitali' 8 50 550 320
$p4.Controls.Add($gbElemLeft)

$gbElemLeft.Controls.Add((mkLbl 'Centro di riferimento:' 8 28 155))
$txtCenterE = mkTxt 166 26 130 '@sun'
$gbElemLeft.Controls.Add($txtCenterE)
$gbElemLeft.Controls.Add((mkLbl '@sun=@10  @0=baricentroSS  @599=Giove' 305 28 280 20))

$gbElemLeft.Controls.Add((mkLbl 'Data Inizio:' 8 60 105))
$dtpStartE = mkDTP 116 58 ([DateTime]::Today); $gbElemLeft.Controls.Add($dtpStartE)
$gbElemLeft.Controls.Add((mkLbl 'Data Fine:' 8 90 105))
$dtpStopE  = mkDTP 116 88 ([DateTime]::Today.AddMonths(1)); $gbElemLeft.Controls.Add($dtpStopE)

$gbElemLeft.Controls.Add((mkLbl 'Passo:' 8 122 55))
$txtStepE = mkTxt 66 120 60 '1'; $gbElemLeft.Controls.Add($txtStepE)
$cmbStepUnitE = mkCmb 132 120 60 @('d','h','m','mo','y') 0; $gbElemLeft.Controls.Add($cmbStepUnitE)

$gbElemLeft.Controls.Add((mkLbl 'Piano di riferimento:' 8 155 150))
$cmbRefPlaneE = mkCmb 160 153 140 @('ECLIPTIC','FRAME','BODY EQUATOR') 0
$gbElemLeft.Controls.Add($cmbRefPlaneE)

$chkODelem  = mkChk 'Dati fisici oggetto' 8 192 200 $true; $gbElemLeft.Controls.Add($chkODelem)
$chkCSVelem = mkChk 'Formato CSV'         8 216 200 $false; $gbElemLeft.Controls.Add($chkCSVelem)

# Info pannello
$gbElemInfo = mkGB 'Elementi Kepleriani - Descrizione campi' 568 50 610 320
$p4.Controls.Add($gbElemInfo)
$rtbElemInfo = mkRtb 8 20 592 290 $true 'White'
$rtbElemInfo.BackColor = $GIALLO
$rtbElemInfo.Text = @"
ELEMENTI ORBITALI KEPLERIANI  (EPHEM_TYPE=ELEMENTS)
=====================================================
JDTDB  = Epoca (Julian Date, TDB)
Tp     = Epoca al perielio (JD)
EC     = Eccentricita'' e              [adim.]
QR     = Distanza al perielio q        [AU]
IN     = Inclinazione i               [deg]
OM     = Long. nodo ascendente Omega  [deg]
W      = Arg. del perielio omega      [deg]
MA     = Anomalia media M             [deg]
TA     = Anomalia vera nu             [deg]
A      = Semi-asse maggiore a         [AU]
AD     = Distanza all''afelio Q       [AU]
PR     = Periodo orbitale             [giorni]
N      = Moto medio n                 [deg/d]

TIPO DI ORBITA per eccentricita'':
  e < 1.0   orbita ellittica (chiusa)
  e = 1.0   orbita parabolica
  e > 1.0   orbita iperbolica (aperta)

Per orbite iperboliche (comete non-period.):
  A < 0  (semi-asse "negativo" per convenzione)
  AD non definito
  PR non definito

SISTEMI DI RIFERIMENTO:
  ECLIPTIC  = eclittica media J2000.0  (default)
  FRAME     = ICRF equatoriale J2000.0
  REF_SYSTEM= ICRF (fisso, sempre usato)
  TP_TYPE   = ABSOLUTE (epoca perielio assoluta)
"@
$gbElemInfo.Controls.Add($rtbElemInfo)

# Tabella esempi
$gbElemEx = mkGB 'Esempi di corpi e loro orbite tipiche' 8 380 1170 400
$p4.Controls.Add($gbElemEx)
$rtbElemEx = mkRtb 8 20 1150 370 $true 'White'
$rtbElemEx.BackColor = $GRIGIO
$rtbElemEx.Text = @"
ESEMPI ELEMENTI ORBITALI
===========================================================================================================
Corpo              | COMMAND      | Centro | Piano    | e (EC)  | a (A, AU) | P (anni)  | Note
-------------------+--------------+--------+----------+---------+-----------+-----------+----------------
Mercurio           | 199          | @sun   | ECLIPTIC | 0.2056  |  0.387    |  0.24     |
Venere             | 299          | @sun   | ECLIPTIC | 0.0067  |  0.723    |  0.62     |
Marte              | 499          | @sun   | ECLIPTIC | 0.0934  |  1.524    |  1.88     |
Giove              | 599          | @sun   | ECLIPTIC | 0.0489  |  5.204    | 11.86     |
Saturno            | 699          | @sun   | ECLIPTIC | 0.0565  |  9.537    | 29.46     |
Luna               | 301          | 500    | ECLIPTIC | 0.0549  |  0.00257  |  0.0748   | geocentrica!
Cerere             | 2000001      | @sun   | ECLIPTIC | 0.0758  |  2.769    |  4.60     | pianeta nano
Vesta              | 2000004      | @sun   | ECLIPTIC | 0.0887  |  2.362    |  3.63     | asteroide
Bennu              | 2101955      | @sun   | ECLIPTIC | 0.2037  |  1.126    |  1.20     | NEA, Apollo
Chirone            | 2002060      | @sun   | ECLIPTIC | 0.3789  |  13.60    | 50.18     | centauro
Cometa di Halley   | DES=1P;      | @sun   | ECLIPTIC | 0.9671  |  17.83    | 75.32     | ellittica
Cometa Hale-Bopp   | DES=C/1995 O1; | @sun | ECLIPTIC| 0.9951  |  186.5    | ~2530     | quasi-parabolica
Voyager 1          | -31          | @sun   | FRAME    | >1.0    | iperb.    | --        | fuga dal SS

===========================================================================================================
NOTA: Gli elementi orbitali variano nel tempo (perturbazioni). I valori riportati sopra sono approssimativi.
      Per la data corrente usare l''API per avere i valori aggiornati. Usare sempre un passo temporale
      abbastanza piccolo per corpi con orbite fortemente perturbate (asteroidi vicini ai pianeti).
"@
$gbElemEx.Controls.Add($rtbElemEx)

#==============================================================================
# [TAB 5] CLOSE APPROACH (AVVICINAMENTI)
#==============================================================================
$p5 = $tabApproach

$p5.Controls.Add((mkLbl 'COMMAND:' 8 14 75 22))
$txtCmdCA = mkTxt 86 12 200 '2101955'
$p5.Controls.Add($txtCmdCA)
$p5.Controls.Add((mkLbl '(asteroide / cometa / NEA, es: 2101955=Bennu)' 295 14 440 22))
$btnGoApp = mkBtn 'CALCOLA >>>' 900 12 130 30
$btnGoApp.BackColor = $BLU; $btnGoApp.ForeColor = $BIANCO
$btnGoApp.Font = New-Object System.Drawing.Font('Consolas',10,[System.Drawing.FontStyle]::Bold)
$p5.Controls.Add($btnGoApp)

$gbCALeft = mkGB 'Parametri Close Approach' 8 50 550 220
$p5.Controls.Add($gbCALeft)

$gbCALeft.Controls.Add((mkLbl 'Data Inizio:' 8 30 105))
$dtpStartCA = mkDTP 116 28 ([DateTime]::Today); $gbCALeft.Controls.Add($dtpStartCA)
$gbCALeft.Controls.Add((mkLbl 'Data Fine:' 8 60 105))
$dtpStopCA  = mkDTP 116 58 ([DateTime]::Today.AddYears(10)); $gbCALeft.Controls.Add($dtpStopCA)

$lCANote = New-Object System.Windows.Forms.Label
$lCANote.Text = "NOTA: EPHEM_TYPE=APPROACH e'' disponibile SOLO per corpi minori (asteroidi, comete, NEA/NEO).`nNon funziona con pianeti principali (199, 299, ...) o con le navicelle spaziali."
$lCANote.Location = New-Object System.Drawing.Point(8,100)
$lCANote.Size = New-Object System.Drawing.Size(530,55)
$lCANote.ForeColor = $ROSSO
$gbCALeft.Controls.Add($lCANote)

$gbCALeft.Controls.Add((mkLbl 'MAKE_EPHEM=YES  EPHEM_TYPE=APPROACH' 8 165 530 20))
$gbCALeft.Controls.Add((mkLbl 'START_TIME + STOP_TIME + OBJ_DATA=YES' 8 182 530 20))

# Info
$gbCAInfo = mkGB 'Close Approach - Output e Campi' 568 50 610 350
$p5.Controls.Add($gbCAInfo)
$rtbCAInfo = mkRtb 8 20 592 320 $true 'White'
$rtbCAInfo.BackColor = $GIALLO
$rtbCAInfo.Text = @"
CLOSE APPROACH  (EPHEM_TYPE = APPROACH)
==========================================
Restituisce le date in cui un corpo del
sistema solare si avvicina ai pianeti.

CAMPI DI OUTPUT TIPICI:
  Date_________YYYY-Mon-DD = data (UT)
  Body         = pianeta avvicinato
  CA_TDB       = tempo calc. minimo (TDB)
  CA_UT        = tempo calc. minimo (UT)
  CA_dT        = incertezza T (min)
  CA_dist      = dist. minima (AU)
  CA_dist_min  = dist. minima -3sigma
  CA_dist_max  = dist. minima +3sigma
  v_rel        = vel. relativa (km/s)
  v_inf        = vel. infinita (km/s)
  h            = distanza (LD=dist.lunare)
  Tisserand(J) = param. Tisserand/Giove

OGGETTI SUPPORTATI:
  Asteroidi vicini alla Terra (NEA)
  Oggetti vicini alla Terra (NEO)
  Comete periodiche e non
  Corpi minori in generale

NON SUPPORTATI:
  Pianeti principali (199, 299, ...)
  Navicelle spaziali (-N)
  Lune (301, 401, ...)
  Sole (10)
"@
$gbCAInfo.Controls.Add($rtbCAInfo)

# ListView esempi NEA/comete
$gbCAEx = mkGB 'Oggetti di esempio - doppio clic per selezionare' 8 280 1170 500
$p5.Controls.Add($gbCAEx)
$lvExamples = New-Object System.Windows.Forms.ListView
$lvExamples.Location = New-Object System.Drawing.Point(8,20)
$lvExamples.Size     = New-Object System.Drawing.Size(1150,470)
$lvExamples.View     = 'Details'; $lvExamples.FullRowSelect = $true
$lvExamples.GridLines = $true; $lvExamples.MultiSelect = $false
$lvExamples.Font = New-Object System.Drawing.Font('Consolas',8)
$lvExamples.BackColor = [System.Drawing.Color]::FromArgb(250,250,255)
foreach ($ch in @('COMMAND','Nome','Tipo','Gruppo','Note')) {
    $col = New-Object System.Windows.Forms.ColumnHeader; $col.Text = $ch
    $col.Width = switch ($ch) { 'COMMAND' {130} 'Nome' {180} 'Tipo' {70} 'Gruppo' {80} default {580} }
    $lvExamples.Columns.Add($col) | Out-Null
}
$exData = @(
    ,@('2101955','Bennu','Asteroide','Apollo','Target missione OSIRIS-REx (NASA) - campione riportato a Terra nel 2023')
    ,@('2025143','Itokawa','Asteroide','Apollo','Target missione Hayabusa (JAXA) - primo campione da asteroide 2010')
    ,@('2162173','Ryugu','Asteroide','Apollo','Target missione Hayabusa2 (JAXA) - campioni riportati nel 2020')
    ,@('2004769','Castalia','Asteroide','Apollo','Prima immagine radar di un asteroide (Arecibo 1989)')
    ,@('2029075','1950 DA','Asteroide','Apollo','In lista di rischio impatto futuro (lontano)')
    ,@('2000001','Cerere','Pianeta nano','Fascia princip.','Ex asteroide, ora pianeta nano. Target Dawn (NASA) 2015')
    ,@('2000004','Vesta','Asteroide','Fascia princip.','Target missione Dawn (NASA) 2011-2012')
    ,@('2000433','Eros','Asteroide','Amor','Primo asteroide in orbita ravvicinata alla Terra - target NEAR 2001')
    ,@('2099942','Apophis','Asteroide','Aten','Passaggio ravvicinato Terra 2029 (~31800 km!). Missione ESA Hera.')
    ,@('DES=1P;','1P/Halley','Cometa','HTC','Cometa di Halley, periodo ~75 anni, ultima apparizione 1986')
    ,@('DES=2P;','2P/Encke','Cometa','JFC','Cometa Encke, periodo piu'' breve nota 3.3 anni')
    ,@('DES=9P;','9P/Tempel 1','Cometa','JFC','Target missione Deep Impact (NASA) - impatto nel 2005')
    ,@('DES=67P;','67P/C-G','Cometa','JFC','Target missione Rosetta/Philae (ESA). Incontro 2014-2016')
    ,@('DES=81P;','81P/Wild 2','Cometa','JFC','Target missione Stardust (NASA) - campione coda 2004')
    ,@('DES=C/2020 F3;','C/2020 F3 NEOWISE','Cometa','LPC','Grande cometa visibile a occhio nudo luglio 2020')
    ,@('DES=C/2023 A3;','C/2023 A3 Tsuchinshan','Cometa','LPC','Grande cometa, perielio ottobre 2024, visibilita'' eccezionale')
    ,@('DES=C/1995 O1;','C/1995 O1 Hale-Bopp','Cometa','LPC','Grande cometa del 1997, visibile ad occhio nudo per mesi')
)
foreach ($ex in $exData) {
    $lvi = New-Object System.Windows.Forms.ListViewItem([string]$ex[0])
    $lvi.SubItems.Add([string]$ex[1]) | Out-Null
    $lvi.SubItems.Add([string]$ex[2]) | Out-Null
    $lvi.SubItems.Add([string]$ex[3]) | Out-Null
    $lvi.SubItems.Add([string]$ex[4]) | Out-Null
    $lvi.Tag = [string]$ex[0]
    $lvExamples.Items.Add($lvi) | Out-Null
}
$gbCAEx.Controls.Add($lvExamples)

#==============================================================================
# [TAB 6] RISULTATI
#==============================================================================
$p6 = $tabOut

# URL bar
$gbUrl = mkGB 'URL Richiesta inviata all''API' 5 5 1185 52
$p6.Controls.Add($gbUrl)
$rtbUrl = New-Object System.Windows.Forms.TextBox
$rtbUrl.Location  = New-Object System.Drawing.Point(8,22)
$rtbUrl.Size      = New-Object System.Drawing.Size(1165,22)
$rtbUrl.ReadOnly  = $true
$rtbUrl.Font      = New-Object System.Drawing.Font('Consolas',7)
$rtbUrl.BackColor = [System.Drawing.Color]::FromArgb(235,235,255)
$rtbUrl.Anchor    = 'Left,Right,Top'
$gbUrl.Controls.Add($rtbUrl)

# Pulsanti output
$btnSaveTxt  = mkBtn 'Salva TXT'    5 64 110 28
$btnCopyAll  = mkBtn 'Copia tutto' 122 64 110 28
$btnCopyURL  = mkBtn 'Copia URL'   239 64 110 28
$btnClearOut = mkBtn 'Pulisci'     356 64 90 28
$btnOpenURL  = mkBtn 'Apri URL nel browser' 455 64 190 28
$btnOpenURL.BackColor = $VIOLA; $btnOpenURL.ForeColor = $BIANCO
foreach ($b in @($btnSaveTxt,$btnCopyAll,$btnCopyURL,$btnClearOut,$btnOpenURL)) { $p6.Controls.Add($b) }

$lOutCount = mkLbl '' 656 68 500 22
$lOutCount.ForeColor = [System.Drawing.Color]::DarkGreen
$p6.Controls.Add($lOutCount)

# Output RichTextBox
$gbOut = mkGB 'Output JPL Horizons' 5 100 1185 685
$p6.Controls.Add($gbOut)
$rtbOut = New-Object System.Windows.Forms.RichTextBox
$rtbOut.Location   = New-Object System.Drawing.Point(5,18)
$rtbOut.Size       = New-Object System.Drawing.Size(1173,660)
$rtbOut.Font       = New-Object System.Drawing.Font('Consolas',8)
$rtbOut.ReadOnly   = $true
$rtbOut.BackColor  = [System.Drawing.Color]::FromArgb(10,10,30)
$rtbOut.ForeColor  = [System.Drawing.Color]::FromArgb(0,255,160)
$rtbOut.ScrollBars = 'Both'
$rtbOut.WordWrap   = $false
$rtbOut.Anchor     = 'Top,Bottom,Left,Right'
$gbOut.Controls.Add($rtbOut)

#==============================================================================
# EVENT HANDLERS
#==============================================================================

# --- TAB 1: Lookup ---
$btnSearch.add_Click({ Do-Lookup })
$txtSearch.add_KeyDown({ if ($_.KeyCode -eq 'Return') { Do-Lookup } })

$btnSelKnown.add_Click({
    $sel = $cmbKnown.SelectedItem
    if ($sel -and $script:OGGETTI_NOTI.ContainsKey($sel) -and $script:OGGETTI_NOTI[$sel] -ne '') {
        SetCmd $script:OGGETTI_NOTI[$sel]
        SetStatus "Selezionato: $sel  -->  ID: $($script:OGGETTI_NOTI[$sel])" 'DarkGreen'
    }
})
$cmbKnown.add_DoubleClick({
    $sel = $cmbKnown.SelectedItem
    if ($sel -and $script:OGGETTI_NOTI.ContainsKey($sel) -and $script:OGGETTI_NOTI[$sel] -ne '') {
        SetCmd $script:OGGETTI_NOTI[$sel]
    }
})

$btnSelDirect.add_Click({
    $id = $txtDirectID.Text.Trim()
    if ($id -ne '') { SetCmd $id; SetStatus "ID diretto impostato: $id" 'DarkGreen' }
})
$txtDirectID.add_KeyDown({
    if ($_.KeyCode -eq 'Return') {
        $id = $txtDirectID.Text.Trim()
        if ($id -ne '') { SetCmd $id; SetStatus "ID diretto impostato: $id" 'DarkGreen' }
    }
})

$lvResults.add_SelectedIndexChanged({
    if ($lvResults.SelectedItems.Count -gt 0) {
        $r = $lvResults.SelectedItems[0].Tag
        if ($r) {
            $sb = [System.Text.StringBuilder]::new()
            [void]$sb.AppendLine("Designazione (PDES) : $(if ($r.pdes)  { $r.pdes  } else { 'N/A' })")
            [void]$sb.AppendLine("Nome                : $(if ($r.name)  { $r.name  } else { 'N/A' })")
            [void]$sb.AppendLine("SPK-ID              : $(if ($r.spkid) { $r.spkid } else { 'N/A' })")
            [void]$sb.AppendLine("Tipo                : $(if ($r.kind)  { $r.kind  } else { 'N/A' })")
            if ($r.alias -and $r.alias.Count -gt 0) {
                [void]$sb.AppendLine("Alias / Desig. alt. : $($r.alias -join '  |  ')")
            }
            if ($r.orbit_id) { [void]$sb.AppendLine("Orbit-ID            : $($r.orbit_id)") }
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("Per usare questo oggetto: clicca 'USA OGGETTO SELEZIONATO' oppure fai doppio clic sulla riga.")
            $rtbLookupDetail.Text = $sb.ToString()
        }
    }
})

$btnUseLV.add_Click({
    if ($lvResults.SelectedItems.Count -gt 0) {
        $r = $lvResults.SelectedItems[0].Tag
        if ($r) {
            $id = if ($r.spkid) { [string]$r.spkid } elseif ($r.pdes) { [string]$r.pdes } else { '' }
            if ($id -ne '') {
                SetCmd $id
                SetStatus "Oggetto da Lookup: $($r.name)  ID=$id" 'DarkGreen'
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show('Seleziona prima un oggetto dalla lista.','Attenzione',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
})

$lvResults.add_DoubleClick({
    if ($lvResults.SelectedItems.Count -gt 0) {
        $r = $lvResults.SelectedItems[0].Tag
        if ($r) {
            $id = if ($r.spkid) { [string]$r.spkid } elseif ($r.pdes) { [string]$r.pdes } else { '' }
            if ($id -ne '') { SetCmd $id; SetStatus "Oggetto: $($r.name)  ID=$id" 'DarkGreen' }
        }
    }
})

$btnInfoObj.add_Click({
    $id = $script:SelectedCmd
    if ($id -eq '') {
        $id = $txtDirectID.Text.Trim()
        if ($id -eq '') { [System.Windows.Forms.MessageBox]::Show('Seleziona prima un oggetto.','Attenzione') | Out-Null; return }
    }
    Do-ObjInfo $id
})

# --- TAB 2: Observer ---
$btnGoObs.add_Click({ Do-Observer })

$btnQtyAll.add_Click({  for ($i=0; $i -lt $clbQty.Items.Count; $i++) { $clbQty.SetItemChecked($i,$true)  } })
$btnQtyNone.add_Click({ for ($i=0; $i -lt $clbQty.Items.Count; $i++) { $clbQty.SetItemChecked($i,$false) } })
$btnQtyStd.add_Click({
    for ($i=0; $i -lt $clbQty.Items.Count; $i++) { $clbQty.SetItemChecked($i,$false) }
    foreach ($idx in @(0,3,8)) { $clbQty.SetItemChecked($idx,$true) }   # 1, 4, 9
})

# --- TAB 3: Vectors ---
$btnGoVec.add_Click({ Do-Vectors })

# --- TAB 4: Elements ---
$btnGoElem.add_Click({ Do-Elements })

# --- TAB 5: Approach ---
$btnGoApp.add_Click({ Do-Approach })

$lvExamples.add_DoubleClick({
    if ($lvExamples.SelectedItems.Count -gt 0) {
        $id = [string]$lvExamples.SelectedItems[0].Tag
        if ($id -ne '') {
            SetCmd $id
            $txtCmdCA.Text = $id
            SetStatus "Esempio selezionato: $id" 'DarkGreen'
        }
    }
})
$lvExamples.add_SelectedIndexChanged({
    if ($lvExamples.SelectedItems.Count -gt 0) {
        $id = [string]$lvExamples.SelectedItems[0].Tag
        if ($id -ne '') { $txtCmdCA.Text = $id }
    }
})

# --- TAB 6: Risultati ---
$btnSaveTxt.add_Click({
    if ($rtbOut.Text.Trim() -eq '') { SetStatus 'Nessun risultato da salvare.' 'DarkOrange'; return }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Title = 'Salva output Horizons'
    $sfd.Filter = 'File di testo (*.txt)|*.txt|Tutti i file (*.*)|*.*'
    $sfd.FileName = 'horizons_' + [DateTime]::Now.ToString('yyyyMMdd_HHmm') + '.txt'
    $sfd.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($sfd.ShowDialog() -eq 'OK') {
        [System.IO.File]::WriteAllText($sfd.FileName, $rtbOut.Text, [System.Text.Encoding]::UTF8)
        SetStatus "Salvato: $($sfd.FileName)" 'DarkGreen'
    }
})

$btnCopyAll.add_Click({
    if ($rtbOut.Text.Trim() -eq '') { return }
    [System.Windows.Forms.Clipboard]::SetText($rtbOut.Text)
    SetStatus 'Output copiato negli appunti.' 'DarkBlue'
})

$btnCopyURL.add_Click({
    if ($rtbUrl.Text.Trim() -eq '') { return }
    [System.Windows.Forms.Clipboard]::SetText($rtbUrl.Text)
    SetStatus 'URL copiato negli appunti.' 'DarkBlue'
})

$btnClearOut.add_Click({
    $rtbOut.Clear(); $rtbUrl.Clear(); $lOutCount.Text = ''
    SetStatus 'Output pulito.' 'Gray'
})

$btnOpenURL.add_Click({
    $u = $rtbUrl.Text.Trim()
    if ($u -ne '') {
        try { [System.Diagnostics.Process]::Start($u) | Out-Null }
        catch { SetStatus "Impossibile aprire il browser: $($_.Exception.Message)" 'Red' }
    }
})

# Aggiorna contatore righe quando cambia l'output
$rtbOut.add_TextChanged({
    $n = ($rtbOut.Text -split "`n").Count
    $lOutCount.Text = "Righe: $n   Caratteri: $($rtbOut.Text.Length)"
})

# Chiudi con Escape
$frm.KeyPreview = $true
$frm.add_KeyDown({
    if ($_.KeyCode -eq 'Escape') {
        $frm.Close()
    }
})

#==============================================================================
# AVVIO
#==============================================================================
SetStatus "Pronto. [1]=Lookup oggetti  [2]=Effemeridi  [3]=Vettori  [4]=Elementi  [5]=Avvicinamenti  [6]=Risultati  |  Esc=Esci" 'DarkBlue'
[System.Windows.Forms.Application]::Run($frm)

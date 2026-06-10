<#
.SYNOPSIS
    PowerShell-GUI für das 3-Töpfe-Finanzprinzip (Dr. Andreas Beck).
.DESCRIPTION
    Dieses Skript modelliert die Vermögensentwicklung basierend auf einer 3-Töpfe-Strategie 
    in einer grafischen Benutzeroberfläche.
.VERSION
    4.8
.CHANGELOG v4.8
    - FIX: Sparplan-Parsing korrigiert – Werte ≥ 1.000 € (mit Tausender-Punkt, z.B. "1.250,00")
      wurden nach Neustart als 0 gespeichert, weil die Komma/Punkt-Ersetzung "1.250.00" erzeugte,
      was TryParse nicht parsen konnte. Neues Parsing: Tausender-Punkt entfernen → Komma → Punkt.
    - FIX: Sparplan-Summenspalte verwendet jetzt explizit de-DE-Kultur statt Thread-Default.
    - FIX: Dividenden-Kalender – Robustere DividendenMonate-Prüfung mit zentraler $hasDM-Variable.
      Where-Object-Ergebnisse werden mit Select-Object -First 1 stabilisiert (verhindert
      Array-statt-Einzelobjekt-Problem bei PowerShell-Pipeline-Rückgaben).
    - FIX: Dividenden-Kalender zeigt keine ETFs mit 0,00 € mehr an ($secAmt -gt 0 Guard).
    - FIX: Portfolio-Positionen mit Stückzahl ≤ 0 werden in Dividendenberechnung übersprungen.
    - FIX: Load-PortfolioData verwendet -Encoding UTF8 und @()-Wrapping für DividendenMonate
      (schützt gegen PowerShell Single-Element-Array-Unwrapping nach JSON-Deserialisierung).
    - FIX: Sparplan-Monate-Array beim Laden ebenfalls mit @() gewrappt.
.CHANGELOG v4.7
    - ÄNDERUNG: Option "jährlich" aus dem AutoBackup-Zeitraum-Dropdown und der Logik entfernt.
.CHANGELOG v4.6
    - FEATURE: Neuer Button "Automatisches Backup..." im Einstellungen-Popup.
      Öffnet Show-AutoBackupPopup mit zwei GroupBoxen (Basis-Konfiguration + Zeit-Einstellungen).
    - FEATURE: In-Process-Scheduler (Invoke-AutoBackupIfDue) prüft bei Programm-Start
      und periodisch (alle 15 Min.) ob ein Backup fällig ist.
    - FEATURE: Backup-Format wählbar: ZIP, Ordner-Kopie oder beides.
    - FEATURE: Registry-Export wird mit gesichert (DataPath für Restore auf anderem Gerät).
    - FEATURE: Intervalle: täglich, wöchentlich (Wochentags-Auswahl), monatlich (Monatstag),
      jährlich, bei Programm-Start.
    - FEATURE: Backup-Bereinigung: Ältere Backups automatisch löschen (konfigurierbar).
    - FEATURE: "Jetzt sofort Backup ausführen"-Button im Konfigurations-Dialog.
    - FEATURE: Backup-Konfiguration + LastBackup-Timestamp im bestehenden Config-JSON persistiert.
    - SCHUTZ: Zielverzeichnis darf nicht innerhalb des Datenverzeichnisses liegen
      (verhindert rekursive Backup-von-Backup-Spiralen).
.CHANGELOG v4.5
    - FIX: Sparplan-Daten werden nun zuverlässig gespeichert und beim Neustart geladen.
      Neue zentrale Sync-Funktion (Sync-SparplanFromGrid) mit EndEdit()-Aufruf vor jedem Save.
      Portfolio-Save synchronisiert Sparplan-Grid automatisch mit, um Datenverlust zu verhindern.
    - UI: Vergangene Monate in der Sparplan-Tabelle werden leicht grau hinterlegt (basierend auf aktuellem Monat).
    - UI: Kürzel-Eingabefeld im Wertpapier-Editor bündig zu Name/ISIN/WKN ausgerichtet (X=80).
    - UI: Kürzel-Hinweistext als Tooltip statt Label implementiert.
.CHANGELOG v4.4
    - FEATURE: Neues Feld "Kürzel (opt.)" im Wertpapier-Editor.
      Wird als Anzeigename in Sparraten-Planung und Dividende-Tabs verwendet (wenn gesetzt).
    - REFACTOR: Sparplan-Grid verwendet versteckte SecId-Spalte für zuverlässige Zuordnung
      (unabhängig von Anzeigenamen-Änderungen).
    - VERIFY: Dividendenberechnungen geprüft – alle Turnus-Varianten korrekt addiert und angezeigt.
.CHANGELOG v4.3
    - FEATURE: Sparraten-Planung Detail-Prognose: Zeigt pro ETF mit Sparplan die voraussichtlich neue
      Gesamtposition (Stückzahl + Gesamtwert €) basierend auf aktuellem Tageskurs.
    - FEATURE: Für ausschüttende ETFs zusätzlich: aktuelle Dividende p.a., zusätzliche Dividende
      durch Sparplan-Investition, neue Gesamtdividende p.a.
    - UI: Rechte Sparplan-Tabelle verkleinert, Detail-Panel im freigewordenen Bereich platziert.
.CHANGELOG v4.2
    - FEATURE: Neuer Dividenden-Turnus "monatlich" mit individuellen Bruttobeträgen pro Monat
      (12 feste Monate Jan–Dez, 2-Spalten-Layout im Wertpapier-Editor).
      Durchgängig implementiert: Save/Load, Dividendenfunktion, Kalender, VAP-Popup, Securities-Grid.
.CHANGELOG v4.1
    - FEATURE: Halbjährliche Dividenden: Individuelle Ausschüttungsmonate und Beträge
      pro Payout-Monat (2 frei wählbare Monate mit eigenen €/Anteil-Beträgen).
    - FEATURE: Neuer "Hilfe"-Button in der Button-Zeile – öffnet Hilfetext als Popup.
    - UI: Tab "Hilfe" umbenannt in "Release-Notes" (geleert, für zukünftige Nutzung).
    - UI: Dividende-Label zeigt nun "Dividende gesamt p.a. (Stand 31.12.<Vorjahr>)".
    - REFACTOR: Zentraler App-Name via $script:AppName – alle "3-Toepfe-Tracker"-Referenzen
      durch Variable ersetzt (Registry, Ordner, Backups, UI-Texte, Deinstallation).
.CHANGELOG v4.0
    - FEATURE: Neuer Dividenden-Turnus "keine Ausschüttung" für thesaurierende ETFs.
      Deaktiviert automatisch das Betrag/Anteil-Feld und zeigt entsprechenden Hinweis.
    - FEATURE: ETF-Vorabpauschale Ergebnis wird nach Berechnung direkt in der jeweiligen
      GroupBox im Sub-Tab angezeigt (Steuerjahr, Vorabpauschale, Steuer gesamt).
    - FIX: ConvertTo-Hashtable auf robustere Template-Version aktualisiert
      (unterstützt nun auch Hashtable-Eingaben und DictionaryEntry).
.CHANGELOG v3.9
    - FEATURE: Vierteljährliche Dividenden: Individuelle Ausschüttungsmonate und Beträge pro 
      Payout-Monat (4 frei wählbare Monate mit eigenen €/Anteil-Beträgen).
    - FEATURE: Neuer Sub-Tab "ETF-Vorabpauschale (Prognose)" im Analyse-Tab.
      Zeigt pro Wertpapier eine GroupBox mit Stammdaten und Button "Berechnung".
    - FEATURE: Vorabpauschale-Berechnungspopup mit vollständiger Berechnung nach InvStG §18:
      Basisertrag, Wertsteigerung, Teilfreistellung (Aktien-/Misch-/Immobilienfonds),
      Abgeltungsteuer, Solidaritätszuschlag und optionale Kirchensteuer.
      Basiszinssätze: 2023=2,55%, 2024=2,29%, 2025=2,53%, 2026=3,20%.
    - UI: Formularhöhe vergrößert für 8 GroupBoxes ohne Scrollen.
.CHANGELOG v3.8
    - FEATURE: Neuer Button "Sicherung wiederherstellen (ZIP)..." im Einstellungen-Popup.
      Ermöglicht das vollständige Wiederherstellen eines Deinstallations-Backups (UserData + Registry).
      Unterstützt Clean-Restore (Daten ersetzen) und Merge-Modus (Daten ergänzen).
#>

# ---------- 1. Globale Konfiguration & Assemblies ----------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization # Für DataGridView

# Zentraler App-Name (für Registry, Ordner, UI-Texte, Backups)
$script:AppName = "Finance-Planner"
$script:AppVersion = "v.4.8"

# Registry-Integration für flexiblen Datenpfad
$regPath = "HKCU:\Software\PSC\$($script:AppName)"
if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
$storedPath = (Get-ItemProperty -Path $regPath -Name DataPath -ErrorAction SilentlyContinue).DataPath

if ($storedPath -and (Test-Path $storedPath)) {
    $dataDir = $storedPath
} else {
    $dataDir = Join-Path -Path ([Environment]::GetFolderPath("UserProfile")) -ChildPath $script:AppName
    if (-not (Test-Path $dataDir)) { New-Item -Path $dataDir -ItemType Directory -Force | Out-Null }
    Set-ItemProperty -Path $regPath -Name DataPath -Value $dataDir
}

# Dateipfade
$userConfigFile    = Join-Path -Path $dataDir -ChildPath "finanzplan-config.json"
$portfolioDataFile = Join-Path -Path $dataDir -ChildPath "portfolio-data.json"

# Globale Script-Variablen
$script:data = @{}
$script:simulationsErgebnis = $null
$script:portfolioData = @{ Securities = [System.Collections.ArrayList]::new(); Portfolio = [System.Collections.ArrayList]::new() }

# ---- Portfolio-Datenfunktionen ----
function Load-PortfolioData {
    $default = @{
        Securities = [System.Collections.ArrayList]::new()
        Portfolio  = [System.Collections.ArrayList]::new()
        Sparplan   = [System.Collections.ArrayList]::new()
    }
    if (Test-Path $portfolioDataFile) {
        try {
            $loaded = Get-Content $portfolioDataFile -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($sec in $loaded.Securities) {
                $secH = @{
                    Id         = [string]$sec.Id
                    Name       = [string]$sec.Name
                    Kuerzel    = if ($sec.Kuerzel) { [string]$sec.Kuerzel } else { "" }
                    ISIN       = [string]$sec.ISIN
                    WKN        = [string]$sec.WKN
                    Regionen   = [System.Collections.ArrayList]::new()
                    Sektoren   = [System.Collections.ArrayList]::new()
                    Industrien = [System.Collections.ArrayList]::new()
                    Dividende  = @{ BetragProAnteil = [double]$sec.Dividende.BetragProAnteil; Turnus = [string]$sec.Dividende.Turnus; DividendenMonate = [System.Collections.ArrayList]::new() }
                }
                # Migration: DividendenMonate laden (vierteljährlich/halbjährlich/monatlich)
                if ($null -ne $sec.Dividende.DividendenMonate) {
                    # @() erzwingt Array-Kontext (schützt gegen PS-Single-Element-Unwrapping)
                    foreach ($dm in @($sec.Dividende.DividendenMonate)) {
                        if ($null -ne $dm -and $null -ne $dm.Monat) {
                            $secH.Dividende.DividendenMonate.Add(@{ Monat = [int]$dm.Monat; Betrag = [double]$dm.Betrag }) | Out-Null
                        }
                    }
                }
                foreach ($r in $sec.Regionen)   { $secH.Regionen.Add(@{Name=[string]$r.Name;Prozent=[double]$r.Prozent})   | Out-Null }
                foreach ($s in $sec.Sektoren)   { $secH.Sektoren.Add(@{Name=[string]$s.Name;Prozent=[double]$s.Prozent})   | Out-Null }
                foreach ($i in $sec.Industrien) { $secH.Industrien.Add(@{Name=[string]$i.Name;Prozent=[double]$i.Prozent}) | Out-Null }
                $default.Securities.Add($secH) | Out-Null
            }
            foreach ($pos in $loaded.Portfolio) {
                $default.Portfolio.Add(@{
                    SecurityId = [string]$pos.SecurityId
                    Tageskurs  = [double]$pos.Tageskurs
                    Stueckzahl = [double]$pos.Stueckzahl
                }) | Out-Null
            }
            if ($loaded.Sparplan) {
                foreach ($sp in $loaded.Sparplan) {
                    $monate = [System.Collections.ArrayList]::new()
                    if ($sp.Monate) {
                        foreach ($m in @($sp.Monate)) { $monate.Add([double]$m) | Out-Null }
                    } else {
                        # Migration altes Format Von/Bis/Sparrate
                        for ($mi = 1; $mi -le 12; $mi++) {
                            $val = if ($mi -ge [int]$sp.MonatVon -and $mi -le [int]$sp.MonatBis) { [double]$sp.Sparrate } else { 0.0 }
                            $monate.Add($val) | Out-Null
                        }
                    }
                    while ($monate.Count -lt 12) { $monate.Add(0.0) | Out-Null }
                    $default.Sparplan.Add(@{ SecurityId = [string]$sp.SecurityId; Monate = $monate }) | Out-Null
                }
            }
        } catch { Write-Warning "Fehler beim Laden der Portfolio-Daten: $($_.Exception.Message)" }
    }
    $script:portfolioData = $default
}

function Save-PortfolioData {
    try { $script:portfolioData | ConvertTo-Json -Depth 10 | Set-Content $portfolioDataFile -Encoding UTF8 }
    catch { [System.Windows.Forms.MessageBox]::Show("Fehler beim Speichern: $($_.Exception.Message)", "Fehler", "OK", "Error") }
}

function Get-PortfolioWeightedDistribution {
    param([string]$Category)
    $result = @{}
    $totalValue = 0.0
    $posValues = [System.Collections.ArrayList]::new()
    foreach ($pos in $script:portfolioData.Portfolio) {
        $sec = $script:portfolioData.Securities | Where-Object { $_.Id -eq $pos.SecurityId }
        if ($null -ne $sec) {
            $val = $pos.Tageskurs * $pos.Stueckzahl
            $totalValue += $val
            $posValues.Add(@{Sec=$sec;Val=$val}) | Out-Null
        }
    }
    if ($totalValue -le 0) { return @{} }
    foreach ($pv in $posValues) {
        $weight = $pv.Val / $totalValue
        foreach ($cat in $pv.Sec[$Category]) {
            $k = $cat.Name
            if (-not $result.ContainsKey($k)) { $result[$k] = 0.0 }
            $result[$k] += $cat.Prozent * $weight
        }
    }
    return $result
}

function Get-PortfolioDividendsByMonth {
    # Returns hashtable Month(1-12) -> total dividend for current/displayed year
    $md = @{}
    for ($m = 1; $m -le 12; $m++) { $md[$m] = 0.0 }
    foreach ($pos in $script:portfolioData.Portfolio) {
        if ([double]$pos.Stueckzahl -le 0) { continue }
        $sec = $script:portfolioData.Securities | Where-Object { $_.Id -eq $pos.SecurityId }
        if ($null -eq $sec) { continue }
        $hasDM = ($null -ne $sec.Dividende.DividendenMonate -and $sec.Dividende.DividendenMonate.Count -gt 0)
        if (-not ($sec.Dividende.BetragProAnteil -gt 0 -or $hasDM)) { continue }
        switch ($sec.Dividende.Turnus) {
            "jährlich"        { $md[12] += $sec.Dividende.BetragProAnteil * $pos.Stueckzahl }
            "halbjährlich"    {
                if ($hasDM) {
                    foreach ($dm in $sec.Dividende.DividendenMonate) {
                        $md[[int]$dm.Monat] += [double]$dm.Betrag * $pos.Stueckzahl
                    }
                } else {
                    # Fallback: gleicher Betrag auf Jun/Dez
                    $amt = $sec.Dividende.BetragProAnteil * $pos.Stueckzahl; $md[6] += $amt; $md[12] += $amt
                }
            }
            "vierteljährlich" {
                if ($hasDM) {
                    foreach ($dm in $sec.Dividende.DividendenMonate) {
                        $md[[int]$dm.Monat] += [double]$dm.Betrag * $pos.Stueckzahl
                    }
                } else {
                    # Fallback: gleicher Betrag auf Mär/Jun/Sep/Dez
                    $amt = $sec.Dividende.BetragProAnteil * $pos.Stueckzahl
                    $md[3] += $amt; $md[6] += $amt; $md[9] += $amt; $md[12] += $amt
                }
            }
            "monatlich" {
                if ($hasDM) {
                    foreach ($dm in $sec.Dividende.DividendenMonate) {
                        $md[[int]$dm.Monat] += [double]$dm.Betrag * $pos.Stueckzahl
                    }
                }
            }
        }
    }
    return $md
}

# ---- Helper: Display-Name (Kuerzel wenn gesetzt, sonst Name) ----
function Get-SecDisplayName {
    param($sec)
    if ($null -ne $sec -and -not [string]::IsNullOrWhiteSpace($sec.Kuerzel)) { return $sec.Kuerzel }
    if ($null -ne $sec) { return $sec.Name }
    return ""
}

# ---- Helper: Category-DataGridView (Regionen/Sektoren/Industrien) ----
function New-CategoryDgv {
    param($parent, $y = 10, $height = 300)
    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Location = New-Object System.Drawing.Point(10, $y)
    $dgv.Size = New-Object System.Drawing.Size(400, $height)
    $dgv.AllowUserToAddRows = $true
    $dgv.AllowUserToDeleteRows = $true
    $dgv.AutoSizeColumnsMode = "Fill"
    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.Name = "Name"; $colName.HeaderText = "Bezeichnung"; $colName.FillWeight = 70
    $colPct  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPct.Name  = "Prozent"; $colPct.HeaderText = "Anteil %"; $colPct.FillWeight = 30
    $dgv.Columns.AddRange($colName, $colPct) | Out-Null
    if ($null -ne $parent) { $parent.Controls.Add($dgv) }
    return $dgv
}

# ---- Security Editor Dialog ----
function Show-SecurityEditorDialog {
    param($existingSecurity = $null)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = if ($null -ne $existingSecurity) { "Wertpapier bearbeiten" } else { "Neues Wertpapier" }
    $dlg.Size = New-Object System.Drawing.Size(660, 670)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false

    $tc = New-Object System.Windows.Forms.TabControl
    $tc.Location = New-Object System.Drawing.Point(8, 8); $tc.Size = New-Object System.Drawing.Size(630, 580)

    # -- Stammdaten --
    $tStamm = New-Object System.Windows.Forms.TabPage; $tStamm.Text = "Stammdaten"
    function Add-LF($lbl, $y, $panel) {
        $l = New-Object System.Windows.Forms.Label; $l.Text = $lbl
        $l.Location = New-Object System.Drawing.Point(10, ($y+4)); $l.Size = New-Object System.Drawing.Size(60, 22)
        $t = New-Object System.Windows.Forms.TextBox
        $t.Location = New-Object System.Drawing.Point(80, $y); $t.Size = New-Object System.Drawing.Size(300, 25)
        $panel.Controls.AddRange(@($l,$t)); return $t
    }
    $txtEName = Add-LF "Name:"  20 $tStamm
    # Kürzel-Feld (optional) – Label breiter, TextBox bündig zu den anderen bei X=80
    $lKuerzel = New-Object System.Windows.Forms.Label; $lKuerzel.Text = "Kürzel:"; $lKuerzel.Location = New-Object System.Drawing.Point(10, 59); $lKuerzel.Size = New-Object System.Drawing.Size(60, 22)
    $txtEKuerzel = New-Object System.Windows.Forms.TextBox; $txtEKuerzel.Location = New-Object System.Drawing.Point(80, 55); $txtEKuerzel.Size = New-Object System.Drawing.Size(300, 25)
    $toolTipKuerzel = New-Object System.Windows.Forms.ToolTip
    $toolTipKuerzel.SetToolTip($txtEKuerzel, "Optional: Wird in Sparplan- und Dividende-Tabs als Anzeigename verwendet")
    $tStamm.Controls.AddRange(@($lKuerzel, $txtEKuerzel))
    $txtEISIN = Add-LF "ISIN:"  90 $tStamm
    $txtEWKN  = Add-LF "WKN:"  125 $tStamm

    # Dividende in Stammdaten-Tab
    $grpDiv = New-Object System.Windows.Forms.GroupBox
    $grpDiv.Text = "Dividende"; $grpDiv.Location = New-Object System.Drawing.Point(10, 165); $grpDiv.Size = New-Object System.Drawing.Size(600, 370)
    
    $lBPA = New-Object System.Windows.Forms.Label; $lBPA.Text = "Betrag/Anteil (€):"; $lBPA.Location = New-Object System.Drawing.Point(10,25); $lBPA.Size = New-Object System.Drawing.Size(130,22)
    $txtBPA = New-Object System.Windows.Forms.TextBox; $txtBPA.Text = "0,00"; $txtBPA.Location = New-Object System.Drawing.Point(145,22); $txtBPA.Size = New-Object System.Drawing.Size(100,25)
    
    $lTurnus = New-Object System.Windows.Forms.Label; $lTurnus.Text = "Turnus:"; $lTurnus.Location = New-Object System.Drawing.Point(10,60); $lTurnus.Size = New-Object System.Drawing.Size(130,22)
    $cmbTurnus = New-Object System.Windows.Forms.ComboBox
    $cmbTurnus.DropDownStyle = "DropDownList"
    $cmbTurnus.Items.AddRange(@("keine Ausschüttung","jährlich","halbjährlich","vierteljährlich","monatlich"))
    $cmbTurnus.SelectedIndex = 0
    $cmbTurnus.Location = New-Object System.Drawing.Point(145,57); $cmbTurnus.Size = New-Object System.Drawing.Size(150,25)
    
    $lblDivInfo = New-Object System.Windows.Forms.Label; $lblDivInfo.Text = ""
    $lblDivInfo.Location = New-Object System.Drawing.Point(10,90); $lblDivInfo.Size = New-Object System.Drawing.Size(580,22)
    $lblDivInfo.ForeColor = [System.Drawing.Color]::DarkBlue

    # --- Panel für vierteljährlich individuelle Monate ---
    $pnlQuarterly = New-Object System.Windows.Forms.GroupBox
    $pnlQuarterly.Text = "Vierteljährlich: Individuelle Ausschüttungsmonate"
    $pnlQuarterly.Location = New-Object System.Drawing.Point(10, 118)
    $pnlQuarterly.Size = New-Object System.Drawing.Size(580, 240)
    $pnlQuarterly.Visible = $false

    $monthNamesQ = @("Januar","Februar","März","April","Mai","Juni","Juli","August","September","Oktober","November","Dezember")
    $script:qCmbMonths = @()
    $script:qTxtAmounts = @()

    for ($qi = 0; $qi -lt 4; $qi++) {
        $yOff = 25 + ($qi * 50)
        $lNum = New-Object System.Windows.Forms.Label; $lNum.Text = "Ausschüttung $($qi+1):"
        $lNum.Location = New-Object System.Drawing.Point(10, ($yOff+4)); $lNum.Size = New-Object System.Drawing.Size(100, 22)
        
        $lMon = New-Object System.Windows.Forms.Label; $lMon.Text = "Monat:"
        $lMon.Location = New-Object System.Drawing.Point(115, ($yOff+4)); $lMon.Size = New-Object System.Drawing.Size(45, 22)
        
        $cmbM = New-Object System.Windows.Forms.ComboBox; $cmbM.DropDownStyle = "DropDownList"
        $cmbM.Items.AddRange($monthNamesQ)
        # Default: Mär, Jun, Sep, Dez
        $defaultMonths = @(2, 5, 8, 11)  # 0-based index
        $cmbM.SelectedIndex = $defaultMonths[$qi]
        $cmbM.Location = New-Object System.Drawing.Point(165, $yOff); $cmbM.Size = New-Object System.Drawing.Size(130, 25)
        
        $lAmt = New-Object System.Windows.Forms.Label; $lAmt.Text = "€/Anteil:"
        $lAmt.Location = New-Object System.Drawing.Point(310, ($yOff+4)); $lAmt.Size = New-Object System.Drawing.Size(60, 22)
        
        $txtAmt = New-Object System.Windows.Forms.TextBox; $txtAmt.Text = "0,0000"
        $txtAmt.Location = New-Object System.Drawing.Point(375, $yOff); $txtAmt.Size = New-Object System.Drawing.Size(100, 25)
        
        $pnlQuarterly.Controls.AddRange(@($lNum, $lMon, $cmbM, $lAmt, $txtAmt))
        $script:qCmbMonths += $cmbM
        $script:qTxtAmounts += $txtAmt
    }

    $lblQInfo = New-Object System.Windows.Forms.Label
    $lblQInfo.Location = New-Object System.Drawing.Point(10, 215); $lblQInfo.Size = New-Object System.Drawing.Size(560, 20)
    $lblQInfo.ForeColor = [System.Drawing.Color]::DarkBlue
    $lblQInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $pnlQuarterly.Controls.Add($lblQInfo)

    # --- Panel für halbjährlich individuelle Monate ---
    $pnlHalfYearly = New-Object System.Windows.Forms.GroupBox
    $pnlHalfYearly.Text = "Halbjährlich: Individuelle Ausschüttungsmonate"
    $pnlHalfYearly.Location = New-Object System.Drawing.Point(10, 118)
    $pnlHalfYearly.Size = New-Object System.Drawing.Size(580, 140)
    $pnlHalfYearly.Visible = $false

    $monthNamesH = @("Januar","Februar","März","April","Mai","Juni","Juli","August","September","Oktober","November","Dezember")
    $script:hCmbMonths = @()
    $script:hTxtAmounts = @()

    for ($hi = 0; $hi -lt 2; $hi++) {
        $yOff = 25 + ($hi * 50)
        $lNum = New-Object System.Windows.Forms.Label; $lNum.Text = "Ausschüttung $($hi+1):"
        $lNum.Location = New-Object System.Drawing.Point(10, ($yOff+4)); $lNum.Size = New-Object System.Drawing.Size(100, 22)
        
        $lMon = New-Object System.Windows.Forms.Label; $lMon.Text = "Monat:"
        $lMon.Location = New-Object System.Drawing.Point(115, ($yOff+4)); $lMon.Size = New-Object System.Drawing.Size(45, 22)
        
        $cmbM = New-Object System.Windows.Forms.ComboBox; $cmbM.DropDownStyle = "DropDownList"
        $cmbM.Items.AddRange($monthNamesH)
        # Default: Jun, Dez
        $defaultMonthsH = @(5, 11)  # 0-based index
        $cmbM.SelectedIndex = $defaultMonthsH[$hi]
        $cmbM.Location = New-Object System.Drawing.Point(165, $yOff); $cmbM.Size = New-Object System.Drawing.Size(130, 25)
        
        $lAmt = New-Object System.Windows.Forms.Label; $lAmt.Text = "€/Anteil:"
        $lAmt.Location = New-Object System.Drawing.Point(310, ($yOff+4)); $lAmt.Size = New-Object System.Drawing.Size(60, 22)
        
        $txtAmt = New-Object System.Windows.Forms.TextBox; $txtAmt.Text = "0,0000"
        $txtAmt.Location = New-Object System.Drawing.Point(375, $yOff); $txtAmt.Size = New-Object System.Drawing.Size(100, 25)
        
        $pnlHalfYearly.Controls.AddRange(@($lNum, $lMon, $cmbM, $lAmt, $txtAmt))
        $script:hCmbMonths += $cmbM
        $script:hTxtAmounts += $txtAmt
    }

    $lblHInfo = New-Object System.Windows.Forms.Label
    $lblHInfo.Location = New-Object System.Drawing.Point(10, 118); $lblHInfo.Size = New-Object System.Drawing.Size(560, 20)
    $lblHInfo.ForeColor = [System.Drawing.Color]::DarkBlue
    $lblHInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $pnlHalfYearly.Controls.Add($lblHInfo)

    # --- Panel für monatlich individuelle Beträge (12 Monate, 2-Spalten) ---
    $pnlMonthly = New-Object System.Windows.Forms.GroupBox
    $pnlMonthly.Text = "Monatlich: Bruttobetrag pro Monat (€/Anteil)"
    $pnlMonthly.Location = New-Object System.Drawing.Point(10, 118)
    $pnlMonthly.Size = New-Object System.Drawing.Size(580, 230)
    $pnlMonthly.Visible = $false

    $monthNamesM = @("Januar","Februar","März","April","Mai","Juni","Juli","August","September","Oktober","November","Dezember")
    $script:mTxtAmounts = @()

    for ($mi = 0; $mi -lt 12; $mi++) {
        $col = [Math]::Floor($mi / 6)   # 0=links, 1=rechts
        $row = $mi % 6
        $xBase = 10 + ($col * 285)
        $yOff  = 22 + ($row * 30)

        $lMon = New-Object System.Windows.Forms.Label
        $lMon.Text = "$($monthNamesM[$mi]):"
        $lMon.Location = New-Object System.Drawing.Point($xBase, ($yOff+3))
        $lMon.Size = New-Object System.Drawing.Size(80, 22)

        $txtAmt = New-Object System.Windows.Forms.TextBox
        $txtAmt.Text = "0,0000"
        $txtAmt.Location = New-Object System.Drawing.Point(($xBase + 85), $yOff)
        $txtAmt.Size = New-Object System.Drawing.Size(90, 25)

        $pnlMonthly.Controls.AddRange(@($lMon, $txtAmt))
        $script:mTxtAmounts += $txtAmt
    }

    $lblMInfo = New-Object System.Windows.Forms.Label
    $lblMInfo.Location = New-Object System.Drawing.Point(10, 205); $lblMInfo.Size = New-Object System.Drawing.Size(560, 20)
    $lblMInfo.ForeColor = [System.Drawing.Color]::DarkBlue
    $lblMInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $pnlMonthly.Controls.Add($lblMInfo)

    $grpDiv.Controls.AddRange(@($lBPA,$txtBPA,$lTurnus,$cmbTurnus,$lblDivInfo,$pnlQuarterly,$pnlHalfYearly,$pnlMonthly))
    $tStamm.Controls.Add($grpDiv)

    # Update quarterly info label
    $updateQInfo = {
        $totalQ = 0.0
        for ($qi = 0; $qi -lt 4; $qi++) {
            $v = 0.0
            [double]::TryParse($script:qTxtAmounts[$qi].Text.Replace(",","."), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$v) | Out-Null
            $totalQ += $v
        }
        $lblQInfo.Text = "Summe p.a.: {0:N4} € pro Anteil" -f $totalQ
    }
    foreach ($tA in $script:qTxtAmounts) { $tA.Add_TextChanged($updateQInfo) }

    # Update half-yearly info label
    $updateHInfo = {
        $totalH = 0.0
        for ($hi = 0; $hi -lt 2; $hi++) {
            $v = 0.0
            [double]::TryParse($script:hTxtAmounts[$hi].Text.Replace(",","."), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$v) | Out-Null
            $totalH += $v
        }
        $lblHInfo.Text = "Summe p.a.: {0:N4} € pro Anteil" -f $totalH
    }
    foreach ($tA in $script:hTxtAmounts) { $tA.Add_TextChanged($updateHInfo) }

    # Update monthly info label
    $updateMInfo = {
        $totalM = 0.0
        for ($mi = 0; $mi -lt 12; $mi++) {
            $v = 0.0
            [double]::TryParse($script:mTxtAmounts[$mi].Text.Replace(",","."), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$v) | Out-Null
            $totalM += $v
        }
        $lblMInfo.Text = "Summe p.a.: {0:N4} € pro Anteil | Ø/Monat: {1:N4} €" -f $totalM, ($totalM / 12.0)
    }
    foreach ($tA in $script:mTxtAmounts) { $tA.Add_TextChanged($updateMInfo) }

    $updateDivInfo = {
        $bpa = 0.0
        [double]::TryParse($txtBPA.Text.Replace(",","."), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$bpa) | Out-Null
        $isQuarterly = ($cmbTurnus.SelectedItem -eq "vierteljährlich")
        $isHalfYearly = ($cmbTurnus.SelectedItem -eq "halbjährlich")
        $isMonthly = ($cmbTurnus.SelectedItem -eq "monatlich")
        $isNone = ($cmbTurnus.SelectedItem -eq "keine Ausschüttung")
        $pnlQuarterly.Visible = $isQuarterly
        $pnlHalfYearly.Visible = $isHalfYearly
        $pnlMonthly.Visible = $isMonthly
        if ($isNone) {
            $lblDivInfo.Text = "Keine Ausschüttung (z.B. thesaurierender ETF)."
            $txtBPA.Enabled = $false
            $txtBPA.Text = "0,0000"
            $lBPA.ForeColor = [System.Drawing.Color]::Gray
        } elseif ($isQuarterly) {
            $lblDivInfo.Text = "Individuelle Beträge pro Ausschüttungsmonat unten eingeben:"
            $txtBPA.Enabled = $false
            $lBPA.ForeColor = [System.Drawing.Color]::Gray
        } elseif ($isHalfYearly) {
            $lblDivInfo.Text = "Individuelle Beträge pro Ausschüttungsmonat unten eingeben:"
            $txtBPA.Enabled = $false
            $lBPA.ForeColor = [System.Drawing.Color]::Gray
        } elseif ($isMonthly) {
            $lblDivInfo.Text = "Individuelle Bruttobeträge pro Monat unten eingeben:"
            $txtBPA.Enabled = $false
            $lBPA.ForeColor = [System.Drawing.Color]::Gray
        } else {
            $txtBPA.Enabled = $true
            $lBPA.ForeColor = [System.Drawing.SystemColors]::ControlText
            switch ($cmbTurnus.SelectedItem) {
                "jährlich"        { $lblDivInfo.Text = "= {0:N4} € p.a. (1× jährlich im Dez.)" -f $bpa }
            }
        }
    }
    $txtBPA.Add_TextChanged($updateDivInfo)
    $cmbTurnus.Add_SelectedIndexChanged({ & $updateDivInfo; & $updateQInfo; & $updateHInfo; & $updateMInfo })

    # -- Regionen --
    $tReg = New-Object System.Windows.Forms.TabPage; $tReg.Text = "Regionen"
    $dgvEReg = New-CategoryDgv $tReg 10 420

    # -- Sektoren --
    $tSek = New-Object System.Windows.Forms.TabPage; $tSek.Text = "Sektoren"
    $dgvESek = New-CategoryDgv $tSek 10 420

    # -- Industrien --
    $tInd = New-Object System.Windows.Forms.TabPage; $tInd.Text = "Industrien"
    $dgvEInd = New-CategoryDgv $tInd 10 420

    $tc.TabPages.AddRange(@($tStamm,$tReg,$tSek,$tInd))
    $dlg.Controls.Add($tc)

    $btnDlgOK = New-Object System.Windows.Forms.Button; $btnDlgOK.Text = "Speichern"
    $btnDlgOK.Location = New-Object System.Drawing.Point(455,600); $btnDlgOK.Size = New-Object System.Drawing.Size(90,30)
    $btnDlgOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnDlgCancel = New-Object System.Windows.Forms.Button; $btnDlgCancel.Text = "Abbrechen"
    $btnDlgCancel.Location = New-Object System.Drawing.Point(555,600); $btnDlgCancel.Size = New-Object System.Drawing.Size(90,30)
    $btnDlgCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.AddRange(@($btnDlgOK,$btnDlgCancel))
    $dlg.AcceptButton = $btnDlgOK; $dlg.CancelButton = $btnDlgCancel

    # Pre-fill
    if ($null -ne $existingSecurity) {
        $txtEName.Text = $existingSecurity.Name
        $txtEKuerzel.Text = if ($existingSecurity.Kuerzel) { $existingSecurity.Kuerzel } else { "" }
        $txtEISIN.Text = $existingSecurity.ISIN
        $txtEWKN.Text  = $existingSecurity.WKN
        $bpa = $existingSecurity.Dividende.BetragProAnteil
        $txtBPA.Text = $bpa.ToString("N4", [System.Globalization.CultureInfo]::new("de-DE"))
        $cmbTurnus.SelectedItem = $existingSecurity.Dividende.Turnus
        # Pre-fill quarterly individual months/amounts
        if ($existingSecurity.Dividende.Turnus -eq "vierteljährlich" -and $existingSecurity.Dividende.DividendenMonate -and $existingSecurity.Dividende.DividendenMonate.Count -gt 0) {
            for ($qi = 0; $qi -lt [Math]::Min(4, $existingSecurity.Dividende.DividendenMonate.Count); $qi++) {
                $dm = $existingSecurity.Dividende.DividendenMonate[$qi]
                $script:qCmbMonths[$qi].SelectedIndex = [int]$dm.Monat - 1  # 1-based -> 0-based
                $script:qTxtAmounts[$qi].Text = ([double]$dm.Betrag).ToString("N4", [System.Globalization.CultureInfo]::new("de-DE"))
            }
        }
        # Pre-fill half-yearly individual months/amounts
        if ($existingSecurity.Dividende.Turnus -eq "halbjährlich" -and $existingSecurity.Dividende.DividendenMonate -and $existingSecurity.Dividende.DividendenMonate.Count -gt 0) {
            for ($hi = 0; $hi -lt [Math]::Min(2, $existingSecurity.Dividende.DividendenMonate.Count); $hi++) {
                $dm = $existingSecurity.Dividende.DividendenMonate[$hi]
                $script:hCmbMonths[$hi].SelectedIndex = [int]$dm.Monat - 1
                $script:hTxtAmounts[$hi].Text = ([double]$dm.Betrag).ToString("N4", [System.Globalization.CultureInfo]::new("de-DE"))
            }
        }
        # Pre-fill monthly individual amounts
        if ($existingSecurity.Dividende.Turnus -eq "monatlich" -and $existingSecurity.Dividende.DividendenMonate -and $existingSecurity.Dividende.DividendenMonate.Count -gt 0) {
            foreach ($dm in $existingSecurity.Dividende.DividendenMonate) {
                $mIdx = [int]$dm.Monat - 1  # 1-based -> 0-based
                if ($mIdx -ge 0 -and $mIdx -lt 12) {
                    $script:mTxtAmounts[$mIdx].Text = ([double]$dm.Betrag).ToString("N4", [System.Globalization.CultureInfo]::new("de-DE"))
                }
            }
        }
        foreach ($r in $existingSecurity.Regionen)   { $dgvEReg.Rows.Add($r.Name, $r.Prozent.ToString("N2",[System.Globalization.CultureInfo]::new("de-DE"))) | Out-Null }
        foreach ($s in $existingSecurity.Sektoren)   { $dgvESek.Rows.Add($s.Name, $s.Prozent.ToString("N2",[System.Globalization.CultureInfo]::new("de-DE"))) | Out-Null }
        foreach ($i in $existingSecurity.Industrien) { $dgvEInd.Rows.Add($i.Name, $i.Prozent.ToString("N2",[System.Globalization.CultureInfo]::new("de-DE"))) | Out-Null }
        & $updateDivInfo
        & $updateQInfo
        & $updateHInfo
        & $updateMInfo
    }

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        if ([string]::IsNullOrWhiteSpace($txtEName.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Bitte einen Namen eingeben.", "Hinweis", "OK", "Warning")
            return $null
        }
        $bpaVal = 0.0
        [double]::TryParse($txtBPA.Text.Replace(",","."), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$bpaVal) | Out-Null
        $divMonate = [System.Collections.ArrayList]::new()
        if ($cmbTurnus.SelectedItem -eq "vierteljährlich") {
            $bpaVal = 0.0  # Bei vierteljährlich wird BetragProAnteil nicht verwendet
            for ($qi = 0; $qi -lt 4; $qi++) {
                $monat = $script:qCmbMonths[$qi].SelectedIndex + 1  # 0-based -> 1-based
                $betrag = 0.0
                [double]::TryParse($script:qTxtAmounts[$qi].Text.Replace(",","."), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$betrag) | Out-Null
                if ($betrag -gt 0) {
                    $divMonate.Add(@{ Monat = $monat; Betrag = $betrag }) | Out-Null
                }
            }
        } elseif ($cmbTurnus.SelectedItem -eq "halbjährlich") {
            $bpaVal = 0.0  # Bei halbjährlich wird BetragProAnteil nicht verwendet
            for ($hi = 0; $hi -lt 2; $hi++) {
                $monat = $script:hCmbMonths[$hi].SelectedIndex + 1
                $betrag = 0.0
                [double]::TryParse($script:hTxtAmounts[$hi].Text.Replace(",","."), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$betrag) | Out-Null
                if ($betrag -gt 0) {
                    $divMonate.Add(@{ Monat = $monat; Betrag = $betrag }) | Out-Null
                }
            }
        } elseif ($cmbTurnus.SelectedItem -eq "monatlich") {
            $bpaVal = 0.0  # Bei monatlich wird BetragProAnteil nicht verwendet
            for ($mi = 0; $mi -lt 12; $mi++) {
                $betrag = 0.0
                [double]::TryParse($script:mTxtAmounts[$mi].Text.Replace(",","."), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$betrag) | Out-Null
                if ($betrag -gt 0) {
                    $divMonate.Add(@{ Monat = ($mi + 1); Betrag = $betrag }) | Out-Null
                }
            }
        }
        $sec = @{
            Id         = if ($null -ne $existingSecurity) { $existingSecurity.Id } else { [Guid]::NewGuid().ToString() }
            Name       = $txtEName.Text.Trim()
            Kuerzel    = $txtEKuerzel.Text.Trim()
            ISIN       = $txtEISIN.Text.Trim()
            WKN        = $txtEWKN.Text.Trim()
            Regionen   = [System.Collections.ArrayList]::new()
            Sektoren   = [System.Collections.ArrayList]::new()
            Industrien = [System.Collections.ArrayList]::new()
            Dividende  = @{ BetragProAnteil = $bpaVal; Turnus = $cmbTurnus.SelectedItem.ToString(); DividendenMonate = $divMonate }
        }
        foreach ($row in $dgvEReg.Rows) {
            if (-not $row.IsNewRow -and -not [string]::IsNullOrWhiteSpace($row.Cells["Name"].Value)) {
                $pct = 0.0; [double]::TryParse(([string]$row.Cells["Prozent"].Value -replace ",","."), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$pct) | Out-Null
                $sec.Regionen.Add(@{Name=[string]$row.Cells["Name"].Value; Prozent=$pct}) | Out-Null
            }
        }
        foreach ($row in $dgvESek.Rows) {
            if (-not $row.IsNewRow -and -not [string]::IsNullOrWhiteSpace($row.Cells["Name"].Value)) {
                $pct = 0.0; [double]::TryParse(([string]$row.Cells["Prozent"].Value -replace ",","."), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$pct) | Out-Null
                $sec.Sektoren.Add(@{Name=[string]$row.Cells["Name"].Value; Prozent=$pct}) | Out-Null
            }
        }
        foreach ($row in $dgvEInd.Rows) {
            if (-not $row.IsNewRow -and -not [string]::IsNullOrWhiteSpace($row.Cells["Name"].Value)) {
                $pct = 0.0; [double]::TryParse(([string]$row.Cells["Prozent"].Value -replace ",","."), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$pct) | Out-Null
                $sec.Industrien.Add(@{Name=[string]$row.Cells["Name"].Value; Prozent=$pct}) | Out-Null
            }
        }
        return $sec
    }
    return $null
}

# ---------- 2. Hilfsfunktionen (aus Template übernommen) ----------

function ConvertTo-Hashtable {
    param ([Parameter(ValueFromPipeline)] $InputObject)
    $hash = @{}
    $properties = $null
    if ($InputObject -is [System.Management.Automation.PSCustomObject] -or $InputObject.PSObject -ne $null) {
        $properties = $InputObject.PSObject.Properties
    } elseif ($InputObject -is [hashtable]) {
        $properties = $InputObject.GetEnumerator()
    } else {
        return $InputObject
    }
    foreach ($prop in $properties) {
        $name  = if ($prop -is [System.Collections.DictionaryEntry]) { $prop.Key } else { $prop.Name }
        $value = if ($prop -is [System.Collections.DictionaryEntry]) { $prop.Value } else { $prop.Value }

        if ($value -is [System.Management.Automation.PSCustomObject] -or $value -is [hashtable]) {
            $hash[$name] = ConvertTo-Hashtable -InputObject $value
        } elseif ($value -is [Array] -or $value -is [System.Collections.ArrayList] -or
                  ($null -ne $value -and $value.GetType().IsArray)) {
            $arrayList = New-Object System.Collections.ArrayList
            foreach ($item in $value) {
                if ($item -is [System.Management.Automation.PSCustomObject] -or $item -is [hashtable]) {
                    $arrayList.Add((ConvertTo-Hashtable -InputObject $item)) | Out-Null
                } else {
                    $arrayList.Add($item) | Out-Null
                }
            }
            $hash[$name] = $arrayList
        } else {
            $hash[$name] = $value
        }
    }
    return $hash
}

# Parse-Number Funktion (V2.6)
function Parse-Number {
    param([string]$text)
    if ([string]::IsNullOrWhiteSpace($text)) { return 0.0 }
    
    $t = $text.Trim()
    
    # 1. Prozentzeichen einfach entfernen. Die Division passiert im Event-Handler.
    $t = $t.Replace("%", "").Trim()

    # 2. Tausender-Trennzeichen (Punkte in DE) entfernen
    $t = $t -replace "\.", "" 

    # 3. Dezimal-Trennzeichen (Komma in DE) durch invarianten Punkt ersetzen
    $t = $t -replace ",", "."

    $out = 0.0
    if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$out)) { 
        return [double]$out # Z.B. "2" -> 2.0; "800000.00" -> 800000.0
    }
    throw "Ungültige Zahl: '$text'"
}


function Merge-Hashtables {
    param ([hashtable]$destination, [hashtable]$source)
    foreach ($key in $source.Keys) {
        if ($destination.ContainsKey($key) -and $destination[$key] -is [hashtable] -and $source[$key] -is [hashtable]) {
            Merge-Hashtables -destination $destination[$key] -source $source[$key]
        } else {
            $destination[$key] = $source[$key]
        }
    }
}

function Load-UserData {
    $defaultConfig = @{
        Finanzplan = @{
            Gesamtvermoegen = 800000.0
            JaehrlicherFinanzbedarf = 30000.0
            Planungshorizont = 3
            RenditeTopf1 = 2.0 # Standardwerte sind jetzt 2 (für 2%)
            RenditeTopf2 = 4.0
            RenditeTopf3 = 8.0
            MonatlichesDividendeneinkommen = 500.0
            Simulationsjahre = 25
        }
        AutoBackup = @{
            Enabled        = $false
            TargetPath     = ""
            FormatZip      = $true
            FormatFolder   = $false
            Interval       = "wöchentlich"
            TimeOfDay      = "03:00"
            Weekdays       = @("Mo")
            MonthDay       = "am Ersten eines Monats"
            IncludeRegistry = $true
            CleanupEnabled = $false
            CleanupKeep    = "letzte drei Backups behalten"
            LastBackup     = ""
        }
    }
    if (Test-Path $userConfigFile) {
        try {
            $loadedData = Get-Content $userConfigFile | ConvertFrom-Json | ConvertTo-Hashtable
            Merge-Hashtables -destination $defaultConfig -source $loadedData
        } catch {
            Write-Warning "Fehler beim Laden der Konfig-Datei $userConfigFile. Verwende Standardwerte."
        }
    }
    $script:data = $defaultConfig
}

function Save-UserData { 
    param ($data)
    $data | ConvertTo-Json -Depth 5 | Set-Content $userConfigFile -Encoding UTF8
}

# ---------- 3. Kernlogik: 3-Töpfe-Simulation ----------
# (Diese Funktionen erwarten die Rendite bereits als Dezimalzahl, z.B. 0.02)

function Start-Finanzsimulation {
    [CmdletBinding()]
    param (
        [double]$StartKapital,
        [double]$Bedarf,
        [int]$Horizont,
        [double]$R1, # Erwartet 0.02
        [double]$R2, # Erwartet 0.02
        [double]$R3, # Erwartet 0.06
        [double]$DividendenMonatlich,
        [int]$Jahre
    )
    
    if ($Horizont -lt 2) { throw "Der Planungshorizont muss mindestens 2 Jahre betragen." }
    
    $Topf1_Ziel = $Bedarf
    $Topf2_Ziel = $Bedarf * ($Horizont - 1)
    
    $Topf1 = $Topf1_Ziel
    $Topf2 = $Topf2_Ziel
    $Topf3 = $StartKapital - $Topf1 - $Topf2
    
    if ($Topf3 -le 0) { throw "Das Startkapital reicht nicht, um Topf 1 & 2 zu füllen. Topf 3 ist 0 oder negativ." }
    
    $Jahresdividenden = $DividendenMonatlich * 12
    $Restbedarf = $Bedarf - $Jahresdividenden
    
    $verlauf = New-Object System.Collections.ArrayList
    
    for ($jahr = 1; $jahr -le $Jahre; $jahr++) {
        
        $T1_Start = $Topf1; $T2_Start = $Topf2; $T3_Start = $Topf3
        $Gesamt_Start = $T1_Start + $T2_Start + $T3_Start
        
        $Topf1 = $Topf1 * (1 + $R1); $Topf2 = $Topf2 * (1 + $R2); $Topf3 = $Topf3 * (1 + $R3)
        
        $EntnahmeDividenden = $Jahresdividenden
        $Topf3 = $Topf3 - $EntnahmeDividenden
        
        $EntnahmeRestbedarf = 0.0
        if ($Restbedarf -gt 0) {
            $EntnahmeRestbedarf = $Restbedarf
            $Topf3 = $Topf3 - $EntnahmeRestbedarf
        }
        $TotalEntnahme = $EntnahmeDividenden + $EntnahmeRestbedarf
        
        $RebalancingInfo = ""
        
        $Delta1 = $Topf1_Ziel - $Topf1
        if ($Delta1 -gt 0) { 
            $Topf1 += $Delta1; $Topf3 -= $Delta1; $RebalancingInfo += "T3->T1: {0:N2}; " -f $Delta1
        } elseif ($Delta1 -lt 0) { 
            $Topf1 += $Delta1; $Topf3 -= $Delta1; $RebalancingInfo += "T1->T3: {0:N2}; " -f (-$Delta1)
        }
        
        $Delta2 = $Topf2_Ziel - $Topf2
        if ($Delta2 -gt 0) { 
            $Topf2 += $Delta2; $Topf3 -= $Delta2; $RebalancingInfo += "T3->T2: {0:N2}" -f $Delta2
        } elseif ($Delta2 -lt 0) { 
            $Topf2 += $Delta2; $Topf3 -= $Delta2; $RebalancingInfo += "T2->T3: {0:N2}" -f (-$Delta2)
        }
        
        # (V3.2) Eigenschaftsnamen ohne Leerzeichen
        $record = [PSCustomObject]@{
            Jahr            = $jahr
            Gesamt_Start    = $Gesamt_Start
            T1_Start        = $T1_Start
            T2_Start        = $T2_Start
            T3_Start        = $T3_Start
            Dividenden      = $EntnahmeDividenden
            Rest_Entnahme   = $EntnahmeRestbedarf
            Total_Entnahme  = $TotalEntnahme
            T1_Ende         = $Topf1
            T2_Ende         = $Topf2
            T3_Ende         = $Topf3
            Gesamt_Ende     = $Topf1 + $Topf2 + $Topf3
            Rebalancing     = $RebalancingInfo.Trim()
        }
        $verlauf.Add($record) | Out-Null
        
        if ($Topf3 -le 0 -and ($Delta1 -gt 0 -or $Delta2 -gt 0)) {
            throw "SIMULATION ABGEBROCHEN in Jahr ${jahr}: Topf 3 ist leer und kann T1/T2 nicht mehr auffüllen."
        }
    }
    return $verlauf
}

function Get-BreakEvenAnalyse {
    [CmdletBinding()]
    param (
        [double]$Bedarf, [int]$Horizont, [double]$R1, [double]$R2, [double]$R3, [double]$DividendenMonatlich
    )
    
    $output = New-Object System.Collections.ArrayList
    $output.Add("--- Break-Even-Analyse ---") | Out-Null
    
    $Topf1_Ziel = $Bedarf
    $Topf2_Ziel = $Bedarf * ($Horizont - 1)
    $Jahresdividenden = $DividendenMonatlich * 12
    $NettoBedarf = $Bedarf - $Jahresdividenden
    
    $ZinsgewinnT1_T2 = ($Topf1_Ziel * $R1) + ($Topf2_Ziel * $R2)
    $NettoLastFuerTopf3 = $NettoBedarf - $ZinsgewinnT1_T2
    
    $output.Add("Nettobedarf (Bedarf - Dividenden): {0:C}" -f $NettoBedarf) | Out-Null
    $output.Add("Zins-Überschuss aus T1 & T2 (fließt zu T3): {0:C}" -f $ZinsgewinnT1_T2) | Out-Null
    $output.Add("Netto-Last (die T3-Rendite decken muss): {0:C}" -f $NettoLastFuerTopf3) | Out-Null
    
    if ($NettoLastFuerTopf3 -le 0) {
        $output.Add("Break-Even erreicht: Divi + Zinsen T1/T2 decken Bedarf.") | Out-Null
    } elseif ($R3 -le 0) {
        $output.Add("Break-Even nicht berechenbar: T3-Rendite ist <= 0.") | Out-Null
    } else {
        $BenoetigtesKapitalTopf3 = $NettoLastFuerTopf3 / $R3
        $output.Add("Benötigtes Kapital in Topf 3 für Break-Even: {0:C}" -f $BenoetigtesKapitalTopf3) | Out-Null
    }
    
    return $output -join [Environment]::NewLine
}

# ---------- 4. GUI-Struktur (Basiert auf Template) ----------

# (V3.4) Korrigierte Format-DataGridView Funktion mit Tooltips
function Format-DataGridView {
    param($grid)
    if ($grid.Columns.Count -eq 0) { return }
    
    $grid.AutoSizeColumnsMode = "AllCells"
    $currencyCellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $currencyCellStyle.Format = "C2" # C = Currency
    $currencyCellStyle.Alignment = "MiddleRight"
    
    # 1. Spaltennamen (Properties) finden
    $currencyCols = @("Gesamt_Start", "T1_Start", "T2_Start", "T3_Start", 
                      "Dividenden", "Rest_Entnahme", "Total_Entnahme",
                      "T1_Ende", "T2_Ende", "T3_Ende", "Gesamt_Ende")
                      
    foreach ($colName in $currencyCols) {
        if ($grid.Columns[$colName]) {
            $grid.Columns[$colName].DefaultCellStyle = $currencyCellStyle
        }
    }
    
    # 2. Spalten-Header (Anzeigetext) und Tooltips "schön" machen
    if ($grid.Columns["Jahr"]) {
        $grid.Columns["Jahr"].ToolTipText = "Das simulierte Jahr."
    }
    if ($grid.Columns["Gesamt_Start"]) {
        $grid.Columns["Gesamt_Start"].HeaderText = "Gesamt Start"
        $grid.Columns["Gesamt_Start"].ToolTipText = "Gesamtkapital zu Beginn des Jahres (vor Rendite)."
    }
    if ($grid.Columns["T1_Start"]) {
        $grid.Columns["T1_Start"].HeaderText = "Topf 1 Start"
        $grid.Columns["T1_Start"].ToolTipText = "Kapital in Topf 1 (Liquidität) zu Jahresbeginn."
    }
    if ($grid.Columns["T2_Start"]) {
        $grid.Columns["T2_Start"].HeaderText = "Topf 2 Start"
        $grid.Columns["T2_Start"].ToolTipText = "Kapital in Topf 2 (Sicherheit) zu Jahresbeginn."
    }
    if ($grid.Columns["T3_Start"]) {
        $grid.Columns["T3_Start"].HeaderText = "Topf 3 Start"
        $grid.Columns["T3_Start"].ToolTipText = "Kapital in Topf 3 (Wachstum) zu Jahresbeginn."
    }
    if ($grid.Columns["Dividenden"]) {
        $grid.Columns["Dividenden"].ToolTipText = "Geplante Entnahme aus Dividenden (aus Topf 3)."
    }
    if ($grid.Columns["Rest_Entnahme"]) {
        $grid.Columns["Rest_Entnahme"].HeaderText = "Rest-Entnahme (T3)"
        $grid.Columns["Rest_Entnahme"].ToolTipText = "Restbedarf, der durch Kapitalverkauf aus Topf 3 gedeckt wird."
    }
    if ($grid.Columns["Total_Entnahme"]) {
        $grid.Columns["Total_Entnahme"].HeaderText = "Total Entnahme"
        $grid.Columns["Total_Entnahme"].ToolTipText = "Summe der Entnahmen (Dividenden + Rest-Entnahme)."
    }
    if ($grid.Columns["T1_Ende"]) {
        $grid.Columns["T1_Ende"].HeaderText = "Topf 1 Ende"
        $grid.Columns["T1_Ende"].ToolTipText = "Kapital in Topf 1 am Jahresende (nach Rendite, vor Rebalancing)."
    }
    if ($grid.Columns["T2_Ende"]) {
        $grid.Columns["T2_Ende"].HeaderText = "Topf 2 Ende"
        $grid.Columns["T2_Ende"].ToolTipText = "Kapital in Topf 2 am Jahresende (nach Rendite, vor Rebalancing)."
    }
    if ($grid.Columns["T3_Ende"]) {
        $grid.Columns["T3_Ende"].HeaderText = "Topf 3 Ende"
        $grid.Columns["T3_Ende"].ToolTipText = "Kapital in Topf 3 am Jahresende (nach Rendite, vor Rebalancing)."
    }
    if ($grid.Columns["Gesamt_Ende"]) {
        $grid.Columns["Gesamt_Ende"].HeaderText = "Gesamt Ende"
        $grid.Columns["Gesamt_Ende"].ToolTipText = "Gesamtkapital am Jahresende (nach Rendite, vor Rebalancing)."
    }
    if ($grid.Columns["Rebalancing"]) {
        $grid.Columns["Rebalancing"].ToolTipText = "Umschichtungen am Jahresende, um Töpfe 1 & 2 aufzufüllen (z.B. T3->T1)."
    }

    $grid.Columns["Jahr"].DefaultCellStyle.Alignment = "MiddleCenter"
    $grid.Columns["Rebalancing"].AutoSizeMode = "Fill"
}


# ---------- AutoBackup-System (v4.6) ----------
#
# Architektur:
#   Invoke-AutoBackupIfDue  - Zentraler Scheduler-Check, bei Programm-Start und periodisch
#   Invoke-AutoBackupNow    - Führt das Backup tatsächlich aus (ZIP und/oder Ordner-Kopie)
#   Show-AutoBackupPopup    - Konfigurations-Dialog
#
# Registry wird mit gesichert (Konsistent mit Uninstall-Backup), da der DataPath-Wert
# essentiell für Restore auf anderem Gerät ist.

function Test-BackupPathSafe {
    param([string]$TargetPath, [string]$DataDir)
    if ([string]::IsNullOrWhiteSpace($TargetPath)) { return "Zielverzeichnis darf nicht leer sein." }
    try {
        $t = [System.IO.Path]::GetFullPath($TargetPath.TrimEnd('\'))
        $d = [System.IO.Path]::GetFullPath($DataDir.TrimEnd('\'))
        if ($t -ieq $d) { return "Zielverzeichnis darf nicht das Datenverzeichnis selbst sein." }
        if ($t.StartsWith($d + [System.IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
            return "Zielverzeichnis darf nicht unterhalb des Datenverzeichnisses liegen (rekursive Backups)."
        }
    } catch {
        return "Ungültiger Pfad: $($_.Exception.Message)"
    }
    return $null
}

function Invoke-BackupCleanup {
    param(
        [Parameter(Mandatory=$true)][string]$TargetPath,
        [Parameter(Mandatory=$true)][int]$Keep
    )
    try {
        if (-not (Test-Path $TargetPath)) { return }
        $prefix = "$($script:AppName)_AutoBackup_"
        $zipBackups = @(Get-ChildItem -Path $TargetPath -Filter "$($prefix)*.zip" -File -ErrorAction SilentlyContinue |
                        Sort-Object -Property CreationTime -Descending)
        $folderBackups = @(Get-ChildItem -Path $TargetPath -Filter "$($prefix)*" -Directory -ErrorAction SilentlyContinue |
                           Sort-Object -Property CreationTime -Descending)
        $effectiveKeep = [Math]::Max($Keep, 1)
        if ($zipBackups.Count -gt $effectiveKeep) {
            $toDelete = $zipBackups | Select-Object -Skip $effectiveKeep
            foreach ($f in $toDelete) {
                try { Remove-Item -Path $f.FullName -Force -ErrorAction Stop }
                catch { Write-Warning "Cleanup-Fehler bei $($f.Name): $($_.Exception.Message)" }
            }
        }
        if ($folderBackups.Count -gt $effectiveKeep) {
            $toDelete = $folderBackups | Select-Object -Skip $effectiveKeep
            foreach ($d in $toDelete) {
                try { Remove-Item -Path $d.FullName -Recurse -Force -ErrorAction Stop }
                catch { Write-Warning "Cleanup-Fehler bei $($d.Name): $($_.Exception.Message)" }
            }
        }
    } catch {
        Write-Warning "Invoke-BackupCleanup-Fehler: $($_.Exception.Message)"
    }
}

function Invoke-AutoBackupNow {
    try {
        $bk = $script:data.AutoBackup
        if (-not $bk -or -not $bk.Enabled) { return $false }
        if (-not $bk.FormatZip -and -not $bk.FormatFolder) { return $false }
        $err = Test-BackupPathSafe -TargetPath $bk.TargetPath -DataDir $dataDir
        if ($err) { Write-Warning "AutoBackup: $err"; return $false }
        if (-not (Test-Path $bk.TargetPath)) {
            New-Item -ItemType Directory -Path $bk.TargetPath -Force | Out-Null
        }
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $prefix = "$($script:AppName)_AutoBackup_"

        # Registry exportieren (optional, aber empfohlen)
        $tempRegFile = $null
        if ($bk.IncludeRegistry) {
            try {
                $tempRegFile = Join-Path -Path $env:TEMP -ChildPath "$($script:AppName)_reg_$timestamp.reg"
                $regKey = "HKCU\Software\PSC\$($script:AppName)"
                $null = & reg.exe export $regKey $tempRegFile /y 2>&1
                if (-not (Test-Path $tempRegFile)) { $tempRegFile = $null }
            } catch {
                Write-Warning "Registry-Export fehlgeschlagen: $($_.Exception.Message)"
                $tempRegFile = $null
            }
        }

        # ZIP-Format
        if ($bk.FormatZip) {
            $zipPath = Join-Path -Path $bk.TargetPath -ChildPath "$($prefix)$timestamp.zip"
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
            $tempStagingDir = Join-Path -Path $env:TEMP -ChildPath "$($script:AppName)_stage_$timestamp"
            New-Item -ItemType Directory -Path $tempStagingDir -Force | Out-Null
            try {
                $stageData = Join-Path -Path $tempStagingDir -ChildPath "UserData"
                Copy-Item -Path $dataDir -Destination $stageData -Recurse -Force
                if ($tempRegFile) {
                    Copy-Item -Path $tempRegFile -Destination (Join-Path $tempStagingDir "Registry.reg") -Force
                }
                [System.IO.Compression.ZipFile]::CreateFromDirectory($tempStagingDir, $zipPath)
            } finally {
                if (Test-Path $tempStagingDir) { Remove-Item $tempStagingDir -Recurse -Force -ErrorAction SilentlyContinue }
            }
        }

        # Ordner-Kopie-Format
        if ($bk.FormatFolder) {
            $folderBackup = Join-Path -Path $bk.TargetPath -ChildPath "$($prefix)$timestamp"
            $folderUserData = Join-Path -Path $folderBackup -ChildPath "UserData"
            New-Item -ItemType Directory -Path $folderUserData -Force | Out-Null
            $robocopyArgs = @($dataDir, $folderUserData, "/E", "/R:1", "/W:1", "/NP", "/NFL", "/NDL", "/NJH", "/NJS")
            $null = & robocopy.exe @robocopyArgs 2>&1
            if ($LASTEXITCODE -ge 8) {
                Write-Warning "Robocopy Exit-Code $LASTEXITCODE - Kopie evtl. unvollständig"
            }
            if ($tempRegFile) {
                Copy-Item -Path $tempRegFile -Destination (Join-Path $folderBackup "Registry.reg") -Force
            }
        }

        # Temp-Registry-Datei aufräumen
        if ($tempRegFile -and (Test-Path $tempRegFile)) {
            Remove-Item $tempRegFile -Force -ErrorAction SilentlyContinue
        }

        # Last-Backup-Timestamp aktualisieren
        $script:data.AutoBackup.LastBackup = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
        Save-UserData -data $script:data

        # Backup-Cleanup NACH erfolgreicher Erstellung
        if ($bk.CleanupEnabled) {
            $keep = switch ([string]$bk.CleanupKeep) {
                "letztes Backup behalten"           { 1 }
                "letzte zwei Backups behalten"      { 2 }
                "letzte drei Backups behalten"      { 3 }
                "alle löschen"                      { 0 }
                default                             { 3 }
            }
            Invoke-BackupCleanup -TargetPath $bk.TargetPath -Keep $keep
        }

        return $true
    } catch {
        Write-Warning "AutoBackup-Fehler: $($_.Exception.Message)"
        return $false
    }
}

function Invoke-AutoBackupIfDue {
    try {
        $bk = $script:data.AutoBackup
        if (-not $bk -or -not $bk.Enabled) { return }
        if (-not $bk.FormatZip -and -not $bk.FormatFolder) { return }
        if ([string]::IsNullOrWhiteSpace($bk.TargetPath)) { return }

        $now = Get-Date
        $lastBackup = $null
        if (-not [string]::IsNullOrWhiteSpace($bk.LastBackup)) {
            try { $lastBackup = [datetime]::Parse($bk.LastBackup) } catch { $lastBackup = $null }
        }

        $due = $false
        switch ($bk.Interval) {
            "bei Programm-Start" {
                if (-not $script:autoBackupRanThisSession) { $due = $true }
            }
            "täglich"     { if (-not $lastBackup -or $lastBackup.Date -lt $now.Date) { $due = $true } }
            "wöchentlich" {
                $selectedDays = @($bk.Weekdays)
                if ($selectedDays.Count -eq 0) {
                    if (-not $lastBackup -or ($now - $lastBackup).TotalDays -ge 7) { $due = $true }
                } else {
                    $dayMap = @{
                        "Monday"="Mo"; "Tuesday"="Di"; "Wednesday"="Mi"; "Thursday"="Do";
                        "Friday"="Fr"; "Saturday"="Sa"; "Sunday"="So"
                    }
                    $todayShort = $dayMap[[string]$now.DayOfWeek]
                    if ($selectedDays -contains $todayShort) {
                        if (-not $lastBackup -or $lastBackup.Date -lt $now.Date) { $due = $true }
                    }
                }
            }
            "monatlich"   {
                $dayOfMonth = $now.Day
                $daysInMonth = [datetime]::DaysInMonth($now.Year, $now.Month)
                $targetDay = 1
                switch ([string]$bk.MonthDay) {
                    "am Ersten eines Monats"         { $targetDay = 1 }
                    "am zweiten Tag"                 { $targetDay = 2 }
                    "am dritten Tag"                 { $targetDay = 3 }
                    "am vierten Tag"                 { $targetDay = 4 }
                    "am fünften Tag"                 { $targetDay = 5 }
                    "zur Monatsmitte (15.)"          { $targetDay = 15 }
                    "am Letzten Tag eines Monats"    { $targetDay = $daysInMonth }
                    default                          { $targetDay = 1 }
                }
                if ($dayOfMonth -ge $targetDay) {
                    if (-not $lastBackup -or
                        $lastBackup.Year -lt $now.Year -or
                        ($lastBackup.Year -eq $now.Year -and $lastBackup.Month -lt $now.Month)) {
                        $due = $true
                    }
                }
            }
        }

        if (-not $due) { return }

        # Bei intervallbasierten Backups: Uhrzeit-Check
        if ($bk.Interval -ne "bei Programm-Start") {
            try {
                $targetTime = [datetime]::ParseExact($bk.TimeOfDay, "HH:mm", $null)
                $todayTarget = $now.Date.Add($targetTime.TimeOfDay)
                if ($now -lt $todayTarget) { return }
            } catch {
                Write-Warning "AutoBackup: Ungültige Uhrzeit '$($bk.TimeOfDay)', Backup wird trotzdem ausgeführt."
            }
        }

        $success = Invoke-AutoBackupNow
        if ($success) {
            $script:autoBackupRanThisSession = $true
        }
    } catch {
        Write-Warning "AutoBackupIfDue-Fehler: $($_.Exception.Message)"
    }
}

function Show-AutoBackupPopup {
    $bk = $script:data.AutoBackup

    $popup = New-Object System.Windows.Forms.Form
    $popup.Size = New-Object System.Drawing.Size(600, 825)
    $popup.Text = "Automatisches Backup konfigurieren"
    $popup.StartPosition = "CenterParent"
    $popup.FormBorderStyle = "FixedDialog"
    $popup.MaximizeBox = $false; $popup.MinimizeBox = $false

    # --- GroupBox 1: Basis-Konfiguration ---
    $groupBasis = New-Object System.Windows.Forms.GroupBox
    $groupBasis.Text = "Basis-Konfiguration"
    $groupBasis.Location = New-Object System.Drawing.Point(15, 15)
    $groupBasis.Size = New-Object System.Drawing.Size(555, 340)
    $popup.Controls.Add($groupBasis)

    $chkEnabled = New-Object System.Windows.Forms.CheckBox
    $chkEnabled.Text = "Automatisches Backup aktivieren"
    $chkEnabled.Location = New-Object System.Drawing.Point(15, 25)
    $chkEnabled.Size = New-Object System.Drawing.Size(300, 22)
    $chkEnabled.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $chkEnabled.Checked = [bool]$bk.Enabled
    $groupBasis.Controls.Add($chkEnabled)

    $lblTarget = New-Object System.Windows.Forms.Label
    $lblTarget.Text = "Zielverzeichnis:"
    $lblTarget.Location = New-Object System.Drawing.Point(15, 60)
    $lblTarget.Size = New-Object System.Drawing.Size(100, 20)
    $groupBasis.Controls.Add($lblTarget)

    $txtTarget = New-Object System.Windows.Forms.TextBox
    $txtTarget.Location = New-Object System.Drawing.Point(120, 57)
    $txtTarget.Size = New-Object System.Drawing.Size(320, 22)
    $txtTarget.Text = [string]$bk.TargetPath
    $groupBasis.Controls.Add($txtTarget)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "Durchsuchen..."
    $btnBrowse.Location = New-Object System.Drawing.Point(445, 55)
    $btnBrowse.Size = New-Object System.Drawing.Size(100, 26)
    $groupBasis.Controls.Add($btnBrowse)
    $btnBrowse.Add_Click({
        $folderDlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDlg.Description = "Zielverzeichnis für automatische Backups wählen"
        if ($txtTarget.Text -and (Test-Path $txtTarget.Text)) { $folderDlg.SelectedPath = $txtTarget.Text }
        if ($folderDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtTarget.Text = $folderDlg.SelectedPath
        }
    })

    $lblFormat = New-Object System.Windows.Forms.Label
    $lblFormat.Text = "Format:"
    $lblFormat.Location = New-Object System.Drawing.Point(15, 100)
    $lblFormat.Size = New-Object System.Drawing.Size(100, 20)
    $groupBasis.Controls.Add($lblFormat)

    $chkZip = New-Object System.Windows.Forms.CheckBox
    $chkZip.Text = "ZIP (Export)"
    $chkZip.Location = New-Object System.Drawing.Point(120, 100)
    $chkZip.Size = New-Object System.Drawing.Size(140, 22)
    $groupBasis.Controls.Add($chkZip)

    $chkFolder = New-Object System.Windows.Forms.CheckBox
    $chkFolder.Text = "Ordner und Dateien"
    $chkFolder.Location = New-Object System.Drawing.Point(270, 100)
    $chkFolder.Size = New-Object System.Drawing.Size(170, 22)
    $groupBasis.Controls.Add($chkFolder)

    $chkBoth = New-Object System.Windows.Forms.CheckBox
    $chkBoth.Text = "Beides"
    $chkBoth.Location = New-Object System.Drawing.Point(450, 100)
    $chkBoth.Size = New-Object System.Drawing.Size(90, 22)
    $groupBasis.Controls.Add($chkBoth)

    $chkZip.Checked    = [bool]$bk.FormatZip
    $chkFolder.Checked = [bool]$bk.FormatFolder
    $chkBoth.Checked   = ([bool]$bk.FormatZip -and [bool]$bk.FormatFolder)

    $chkRegistry = New-Object System.Windows.Forms.CheckBox
    $chkRegistry.Text = "Registry-Einstellungen mitsichern (empfohlen)"
    $chkRegistry.Location = New-Object System.Drawing.Point(15, 140)
    $chkRegistry.Size = New-Object System.Drawing.Size(500, 22)
    $chkRegistry.Checked = [bool]$bk.IncludeRegistry
    $groupBasis.Controls.Add($chkRegistry)

    $lblRegInfo = New-Object System.Windows.Forms.Label
    $lblRegInfo.Text = "Sichert den DataPath-Konfigurationseintrag. Wichtig für Restore auf anderem Gerät."
    $lblRegInfo.Location = New-Object System.Drawing.Point(35, 162)
    $lblRegInfo.Size = New-Object System.Drawing.Size(500, 20)
    $lblRegInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblRegInfo.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $groupBasis.Controls.Add($lblRegInfo)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(15, 195)
    $lblStatus.Size = New-Object System.Drawing.Size(530, 20)
    $lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    if ($bk.LastBackup) {
        try {
            $d = [datetime]::Parse($bk.LastBackup)
            $lblStatus.Text = "Letztes erfolgreiches Backup: $($d.ToString('dd.MM.yyyy HH:mm'))"
            $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
        } catch {
            $lblStatus.Text = "Letztes Backup: unbekannt"
        }
    } else {
        $lblStatus.Text = "Noch kein Backup durchgeführt."
        $lblStatus.ForeColor = [System.Drawing.Color]::DarkSlateGray
    }
    $groupBasis.Controls.Add($lblStatus)

    # Sub-GroupBox "Backup-Bereinigung"
    $groupCleanup = New-Object System.Windows.Forms.GroupBox
    $groupCleanup.Text = "Backup-Bereinigung"
    $groupCleanup.Location = New-Object System.Drawing.Point(15, 222)
    $groupCleanup.Size = New-Object System.Drawing.Size(525, 105)
    $groupBasis.Controls.Add($groupCleanup)

    $chkCleanup = New-Object System.Windows.Forms.CheckBox
    $chkCleanup.Text = "Ältere Backups löschen"
    $chkCleanup.Location = New-Object System.Drawing.Point(15, 25)
    $chkCleanup.Size = New-Object System.Drawing.Size(200, 22)
    $chkCleanup.Checked = [bool]$bk.CleanupEnabled
    $groupCleanup.Controls.Add($chkCleanup)

    $lblCleanupRule = New-Object System.Windows.Forms.Label
    $lblCleanupRule.Text = "Regel:"
    $lblCleanupRule.Location = New-Object System.Drawing.Point(230, 28)
    $lblCleanupRule.Size = New-Object System.Drawing.Size(50, 20)
    $groupCleanup.Controls.Add($lblCleanupRule)

    $cmbCleanupKeep = New-Object System.Windows.Forms.ComboBox
    $cmbCleanupKeep.Location = New-Object System.Drawing.Point(285, 25)
    $cmbCleanupKeep.Size = New-Object System.Drawing.Size(225, 22)
    $cmbCleanupKeep.DropDownStyle = "DropDownList"
    $cmbCleanupKeep.Items.AddRange(@(
        "letztes Backup behalten",
        "letzte zwei Backups behalten",
        "letzte drei Backups behalten",
        "alle löschen"
    ))
    $cmbCleanupKeep.SelectedItem = [string]$bk.CleanupKeep
    if (-not $cmbCleanupKeep.SelectedItem) { $cmbCleanupKeep.SelectedItem = "letzte drei Backups behalten" }
    $groupCleanup.Controls.Add($cmbCleanupKeep)

    $lblCleanupHint = New-Object System.Windows.Forms.Label
    $lblCleanupHint.Text = "Bereinigung erfolgt NACH erfolgreichem Backup. Betrifft nur Dateien mit`nPrefix '$($script:AppName)_AutoBackup_'. Fremde Dateien bleiben unberührt."
    $lblCleanupHint.Location = New-Object System.Drawing.Point(15, 55)
    $lblCleanupHint.Size = New-Object System.Drawing.Size(500, 40)
    $lblCleanupHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblCleanupHint.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $groupCleanup.Controls.Add($lblCleanupHint)

    $updateCleanupVisibility = {
        $on = $chkCleanup.Checked
        $lblCleanupRule.Visible = $on
        $cmbCleanupKeep.Visible = $on
    }
    $chkCleanup.Add_CheckedChanged($updateCleanupVisibility)
    & $updateCleanupVisibility

    # Checkbox-Intelligenz (SuspendEvent-Pattern)
    $script:abSuspend = $false
    $syncFromIndividual = {
        if ($script:abSuspend) { return }
        $script:abSuspend = $true
        try {
            $bothChecked = $chkZip.Checked -and $chkFolder.Checked
            $chkBoth.Checked = $bothChecked
            $chkZip.Enabled    = -not $bothChecked
            $chkFolder.Enabled = -not $bothChecked
        } finally { $script:abSuspend = $false }
    }
    $syncFromBoth = {
        if ($script:abSuspend) { return }
        $script:abSuspend = $true
        try {
            if ($chkBoth.Checked) {
                $chkZip.Checked = $true
                $chkFolder.Checked = $true
                $chkZip.Enabled = $false
                $chkFolder.Enabled = $false
            } else {
                $chkZip.Checked = $false
                $chkFolder.Checked = $false
                $chkZip.Enabled = $true
                $chkFolder.Enabled = $true
            }
        } finally { $script:abSuspend = $false }
    }
    $chkZip.Add_CheckedChanged($syncFromIndividual)
    $chkFolder.Add_CheckedChanged($syncFromIndividual)
    $chkBoth.Add_CheckedChanged($syncFromBoth)
    & $syncFromIndividual

    # --- GroupBox 2: Zeit-Einstellungen ---
    $groupTime = New-Object System.Windows.Forms.GroupBox
    $groupTime.Text = "Zeit-Einstellungen"
    $groupTime.Location = New-Object System.Drawing.Point(15, 365)
    $groupTime.Size = New-Object System.Drawing.Size(555, 345)
    $popup.Controls.Add($groupTime)

    $lblInterval = New-Object System.Windows.Forms.Label
    $lblInterval.Text = "Zeitraum:"
    $lblInterval.Location = New-Object System.Drawing.Point(15, 35)
    $lblInterval.Size = New-Object System.Drawing.Size(100, 20)
    $groupTime.Controls.Add($lblInterval)

    $cmbInterval = New-Object System.Windows.Forms.ComboBox
    $cmbInterval.Location = New-Object System.Drawing.Point(120, 32)
    $cmbInterval.Size = New-Object System.Drawing.Size(220, 22)
    $cmbInterval.DropDownStyle = "DropDownList"
    $cmbInterval.Items.AddRange(@("täglich","wöchentlich","monatlich","bei Programm-Start"))
    $cmbInterval.SelectedItem = [string]$bk.Interval
    if (-not $cmbInterval.SelectedItem) { $cmbInterval.SelectedItem = "wöchentlich" }
    $groupTime.Controls.Add($cmbInterval)

    $lblTime = New-Object System.Windows.Forms.Label
    $lblTime.Text = "Uhrzeit:"
    $lblTime.Location = New-Object System.Drawing.Point(15, 75)
    $lblTime.Size = New-Object System.Drawing.Size(100, 20)
    $groupTime.Controls.Add($lblTime)

    $dtTime = New-Object System.Windows.Forms.DateTimePicker
    $dtTime.Format = [System.Windows.Forms.DateTimePickerFormat]::Time
    $dtTime.ShowUpDown = $true
    $dtTime.Location = New-Object System.Drawing.Point(120, 72)
    $dtTime.Size = New-Object System.Drawing.Size(120, 22)
    try {
        $parsedTime = [datetime]::ParseExact([string]$bk.TimeOfDay, "HH:mm", $null)
        $dtTime.Value = (Get-Date).Date.Add($parsedTime.TimeOfDay)
    } catch { $dtTime.Value = (Get-Date).Date.AddHours(3) }
    $groupTime.Controls.Add($dtTime)

    $lblTimeHint = New-Object System.Windows.Forms.Label
    $lblTimeHint.Text = ""
    $lblTimeHint.Location = New-Object System.Drawing.Point(250, 76)
    $lblTimeHint.Size = New-Object System.Drawing.Size(290, 20)
    $lblTimeHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblTimeHint.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $groupTime.Controls.Add($lblTimeHint)

    # Sub-GroupBox "Wochentage" - nur sichtbar bei "wöchentlich"
    $groupWeekdays = New-Object System.Windows.Forms.GroupBox
    $groupWeekdays.Text = "Wochentage"
    $groupWeekdays.Location = New-Object System.Drawing.Point(15, 105)
    $groupWeekdays.Size = New-Object System.Drawing.Size(525, 170)
    $groupTime.Controls.Add($groupWeekdays)

    $weekdayKeys = @("Mo","Di","Mi","Do","Fr","Sa","So")
    $weekdayLabels = @{
        "Mo"="Montag"; "Di"="Dienstag"; "Mi"="Mittwoch"; "Do"="Donnerstag";
        "Fr"="Freitag"; "Sa"="Samstag"; "So"="Sonntag"
    }
    $savedDays = @()
    if ($bk.Weekdays) { $savedDays = @($bk.Weekdays) }
    $wdCheckboxes = @{}

    $wdSpec = @(
        @{ Key="Mo"; X=15;  Y=25; W=72 }
        @{ Key="Di"; X=87;  Y=25; W=72 }
        @{ Key="Mi"; X=159; Y=25; W=72 }
        @{ Key="Do"; X=231; Y=25; W=90 }
        @{ Key="Fr"; X=321; Y=25; W=72 }
        @{ Key="Sa"; X=15;  Y=55; W=72 }
        @{ Key="So"; X=87;  Y=55; W=72 }
    )
    foreach ($spec in $wdSpec) {
        $k = $spec.Key
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = $weekdayLabels[$k]
        $cb.Location = New-Object System.Drawing.Point($spec.X, $spec.Y)
        $cb.Size = New-Object System.Drawing.Size($spec.W, 22)
        $cb.Checked = ($savedDays -contains $k)
        $cb.Tag = $k
        $groupWeekdays.Controls.Add($cb)
        $wdCheckboxes[$k] = $cb
    }

    $btnWdAll = New-Object System.Windows.Forms.Button
    $btnWdAll.Text = "Alle auswählen"
    $btnWdAll.Location = New-Object System.Drawing.Point(15, 85)
    $btnWdAll.Size = New-Object System.Drawing.Size(110, 24)
    $groupWeekdays.Controls.Add($btnWdAll)
    $btnWdAll.Add_Click({
        foreach ($k in $weekdayKeys) { $wdCheckboxes[$k].Checked = $true }
    }.GetNewClosure())

    $btnWdNone = New-Object System.Windows.Forms.Button
    $btnWdNone.Text = "Alle abwählen"
    $btnWdNone.Location = New-Object System.Drawing.Point(135, 85)
    $btnWdNone.Size = New-Object System.Drawing.Size(110, 24)
    $groupWeekdays.Controls.Add($btnWdNone)
    $btnWdNone.Add_Click({
        foreach ($k in $weekdayKeys) { $wdCheckboxes[$k].Checked = $false }
    }.GetNewClosure())

    $lblWdHint = New-Object System.Windows.Forms.Label
    $lblWdHint.Text = "Backup erfolgt an den ausgewählten Tagen, sofern die Uhrzeit`nerreicht ist und heute noch kein Backup durchgeführt wurde."
    $lblWdHint.Location = New-Object System.Drawing.Point(15, 135)
    $lblWdHint.Size = New-Object System.Drawing.Size(500, 26)
    $lblWdHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblWdHint.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $groupWeekdays.Controls.Add($lblWdHint)

    # Sub-GroupBox "Monatstag" - nur sichtbar bei "monatlich"
    $groupMonthDay = New-Object System.Windows.Forms.GroupBox
    $groupMonthDay.Text = "Monatstag"
    $groupMonthDay.Location = New-Object System.Drawing.Point(15, 105)
    $groupMonthDay.Size = New-Object System.Drawing.Size(525, 170)
    $groupTime.Controls.Add($groupMonthDay)

    $lblMonthDay = New-Object System.Windows.Forms.Label
    $lblMonthDay.Text = "Durchführung am:"
    $lblMonthDay.Location = New-Object System.Drawing.Point(15, 28)
    $lblMonthDay.Size = New-Object System.Drawing.Size(130, 20)
    $groupMonthDay.Controls.Add($lblMonthDay)

    $cmbMonthDay = New-Object System.Windows.Forms.ComboBox
    $cmbMonthDay.Location = New-Object System.Drawing.Point(150, 25)
    $cmbMonthDay.Size = New-Object System.Drawing.Size(260, 22)
    $cmbMonthDay.DropDownStyle = "DropDownList"
    $cmbMonthDay.Items.AddRange(@(
        "am Ersten eines Monats",
        "am zweiten Tag",
        "am dritten Tag",
        "am vierten Tag",
        "am fünften Tag",
        "zur Monatsmitte (15.)",
        "am Letzten Tag eines Monats"
    ))
    $cmbMonthDay.SelectedItem = [string]$bk.MonthDay
    if (-not $cmbMonthDay.SelectedItem) { $cmbMonthDay.SelectedItem = "am Ersten eines Monats" }
    $groupMonthDay.Controls.Add($cmbMonthDay)

    $lblMdHint = New-Object System.Windows.Forms.Label
    $lblMdHint.Text = "Backup erfolgt am gewählten Tag des Monats, sofern Uhrzeit erreicht ist."
    $lblMdHint.Location = New-Object System.Drawing.Point(15, 55)
    $lblMdHint.Size = New-Object System.Drawing.Size(500, 20)
    $lblMdHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblMdHint.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $groupMonthDay.Controls.Add($lblMdHint)

    # Toggle: Sichtbarkeit basierend auf Interval
    $updateTimeEnable = {
        $isStart   = ($cmbInterval.SelectedItem -eq "bei Programm-Start")
        $isWeekly  = ($cmbInterval.SelectedItem -eq "wöchentlich")
        $isMonthly = ($cmbInterval.SelectedItem -eq "monatlich")
        $dtTime.Enabled = -not $isStart
        $lblTime.Enabled = -not $isStart
        if ($isStart) {
            $lblTimeHint.Text = "(nicht relevant: Backup beim Programm-Start)"
        } else {
            $lblTimeHint.Text = "Backup erfolgt nicht vor dieser Uhrzeit."
        }
        $groupWeekdays.Visible = $isWeekly
        $groupMonthDay.Visible = $isMonthly
    }
    $cmbInterval.Add_SelectedIndexChanged($updateTimeEnable)
    & $updateTimeEnable

    $lblScheduleInfo = New-Object System.Windows.Forms.Label
    $lblScheduleInfo.Text = "Hinweis: Es wird KEIN Windows-Aufgabenplanungs-Task erstellt. Backups erfolgen nur während das Programm läuft (Start-Check + Scheduler alle 15 Min.). Verpasste Intervalle werden beim nächsten Programmstart nachgeholt."
    $lblScheduleInfo.Location = New-Object System.Drawing.Point(15, 285)
    $lblScheduleInfo.Size = New-Object System.Drawing.Size(530, 55)
    $lblScheduleInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblScheduleInfo.ForeColor = [System.Drawing.Color]::DarkRed
    $groupTime.Controls.Add($lblScheduleInfo)

    # --- Jetzt-Backup-Button ---
    $btnBackupNow = New-Object System.Windows.Forms.Button
    $btnBackupNow.Text = "Jetzt sofort Backup ausführen"
    $btnBackupNow.Location = New-Object System.Drawing.Point(15, 745)
    $btnBackupNow.Size = New-Object System.Drawing.Size(220, 32)
    $popup.Controls.Add($btnBackupNow)
    $btnBackupNow.Add_Click({
        $err = Test-BackupPathSafe -TargetPath $txtTarget.Text -DataDir $dataDir
        if ($err) {
            [System.Windows.Forms.MessageBox]::Show($err, "Ungültiges Zielverzeichnis", "OK", "Warning")
            return
        }
        if (-not $chkZip.Checked -and -not $chkFolder.Checked) {
            [System.Windows.Forms.MessageBox]::Show("Bitte mindestens ein Format (ZIP oder Ordner) auswählen.", "Format fehlt", "OK", "Warning")
            return
        }
        $prev = @{
            Enabled       = $script:data.AutoBackup.Enabled
            TargetPath    = $script:data.AutoBackup.TargetPath
            FormatZip     = $script:data.AutoBackup.FormatZip
            FormatFolder  = $script:data.AutoBackup.FormatFolder
            IncludeRegistry = $script:data.AutoBackup.IncludeRegistry
        }
        $script:data.AutoBackup.Enabled         = $true
        $script:data.AutoBackup.TargetPath      = $txtTarget.Text
        $script:data.AutoBackup.FormatZip       = $chkZip.Checked
        $script:data.AutoBackup.FormatFolder    = $chkFolder.Checked
        $script:data.AutoBackup.IncludeRegistry = $chkRegistry.Checked
        $ok = Invoke-AutoBackupNow
        if (-not $prev.Enabled) { $script:data.AutoBackup.Enabled = $prev.Enabled }
        if ($ok) {
            [System.Windows.Forms.MessageBox]::Show("Backup erfolgreich erstellt in:`n$($txtTarget.Text)", "Backup", "OK", "Information")
            $now = Get-Date
            $lblStatus.Text = "Letztes erfolgreiches Backup: $($now.ToString('dd.MM.yyyy HH:mm'))"
            $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
        } else {
            [System.Windows.Forms.MessageBox]::Show("Backup fehlgeschlagen. Details siehe PowerShell-Warnings.", "Backup-Fehler", "OK", "Error")
        }
    })

    # --- Abbrechen / Speichern ---
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abbrechen"
    $btnCancel.Location = New-Object System.Drawing.Point(355, 745)
    $btnCancel.Size = New-Object System.Drawing.Size(105, 32)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $popup.Controls.Add($btnCancel)
    $popup.CancelButton = $btnCancel

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Speichern"
    $btnSave.Location = New-Object System.Drawing.Point(465, 745)
    $btnSave.Size = New-Object System.Drawing.Size(105, 32)
    $popup.Controls.Add($btnSave)
    $popup.AcceptButton = $btnSave

    $btnSave.Add_Click({
        $selectedWeekdays = @()
        foreach ($k in $weekdayKeys) {
            if ($wdCheckboxes[$k].Checked) { $selectedWeekdays += $k }
        }
        if ($chkEnabled.Checked) {
            $err = Test-BackupPathSafe -TargetPath $txtTarget.Text -DataDir $dataDir
            if ($err) {
                [System.Windows.Forms.MessageBox]::Show($err, "Ungültiges Zielverzeichnis", "OK", "Warning")
                return
            }
            if (-not $chkZip.Checked -and -not $chkFolder.Checked) {
                [System.Windows.Forms.MessageBox]::Show("Bitte mindestens ein Format (ZIP oder Ordner) auswählen.", "Format fehlt", "OK", "Warning")
                return
            }
            if ($cmbInterval.SelectedItem -eq "wöchentlich" -and $selectedWeekdays.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Bitte mindestens einen Wochentag auswählen, an dem das wöchentliche Backup erfolgen soll.", "Wochentag fehlt", "OK", "Warning")
                return
            }
        }
        $script:data.AutoBackup.Enabled         = $chkEnabled.Checked
        $script:data.AutoBackup.TargetPath      = $txtTarget.Text
        $script:data.AutoBackup.FormatZip       = $chkZip.Checked
        $script:data.AutoBackup.FormatFolder    = $chkFolder.Checked
        $script:data.AutoBackup.IncludeRegistry = $chkRegistry.Checked
        $script:data.AutoBackup.Interval        = [string]$cmbInterval.SelectedItem
        $script:data.AutoBackup.TimeOfDay       = $dtTime.Value.ToString("HH:mm")
        $script:data.AutoBackup.Weekdays        = $selectedWeekdays
        $script:data.AutoBackup.MonthDay        = [string]$cmbMonthDay.SelectedItem
        $script:data.AutoBackup.CleanupEnabled  = $chkCleanup.Checked
        $script:data.AutoBackup.CleanupKeep     = [string]$cmbCleanupKeep.SelectedItem
        Save-UserData -data $script:data
        [System.Windows.Forms.MessageBox]::Show("Backup-Einstellungen gespeichert.", "Gespeichert", "OK", "Information")
        $popup.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $popup.Close()
    })

    $popup.ShowDialog() | Out-Null
}

# (Funktion Show-SettingsPopup V4.6 – mit AutoBackup-Button)
function Show-SettingsPopup {
    param($mainForm)
    $popupForm = New-Object System.Windows.Forms.Form
    $popupForm.Size = New-Object System.Drawing.Size(450, 460)
    $popupForm.Text = "Einstellungen & Datenverwaltung"
    $popupForm.StartPosition = "CenterParent"
    $popupForm.FormBorderStyle = "FixedDialog"
    $popupForm.MaximizeBox = $false
    $popupForm.MinimizeBox = $false
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Verwalten Sie hier Ihre Konfigurationsdateien und Exporte."
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(400, 20)
    $popupForm.Controls.Add($label)
    $openConfigFileButton = New-Object System.Windows.Forms.Button
    $openConfigFileButton.Text = "Konfigurationsdatei öffnen (.json)"
    $openConfigFileButton.Location = New-Object System.Drawing.Point(20, 60)
    $openConfigFileButton.Size = New-Object System.Drawing.Size(400, 35)
    $popupForm.Controls.Add($openConfigFileButton)
    $openDataPathButton = New-Object System.Windows.Forms.Button
    $openDataPathButton.Text = "Daten-Verzeichnis im Explorer öffnen"
    $openDataPathButton.Location = New-Object System.Drawing.Point(20, 105)
    $openDataPathButton.Size = New-Object System.Drawing.Size(400, 35)
    $popupForm.Controls.Add($openDataPathButton)
    $exportDataButton = New-Object System.Windows.Forms.Button
    $exportDataButton.Text = "Benutzerdaten als ZIP-Archiv exportieren..."
    $exportDataButton.Location = New-Object System.Drawing.Point(20, 150)
    $exportDataButton.Size = New-Object System.Drawing.Size(400, 35)
    $popupForm.Controls.Add($exportDataButton)
    $moveDataButton = New-Object System.Windows.Forms.Button
    $moveDataButton.Text = "Datenpfad ändern..."
    $moveDataButton.Location = New-Object System.Drawing.Point(20, 195)
    $moveDataButton.Size = New-Object System.Drawing.Size(400, 35)
    $popupForm.Controls.Add($moveDataButton)
    $autoBackupButton = New-Object System.Windows.Forms.Button
    $autoBackupButton.Text = "Automatisches Backup..."
    $autoBackupButton.Location = New-Object System.Drawing.Point(20, 240)
    $autoBackupButton.Size = New-Object System.Drawing.Size(400, 35)
    $autoBackupButton.ForeColor = [System.Drawing.Color]::DarkBlue
    $popupForm.Controls.Add($autoBackupButton)
    $restoreBackupButton = New-Object System.Windows.Forms.Button
    $restoreBackupButton.Text = "Sicherung wiederherstellen (ZIP)..."
    $restoreBackupButton.Location = New-Object System.Drawing.Point(20, 285)
    $restoreBackupButton.Size = New-Object System.Drawing.Size(400, 35)
    $restoreBackupButton.ForeColor = [System.Drawing.Color]::DarkGreen
    $popupForm.Controls.Add($restoreBackupButton)
    $uninstallButton = New-Object System.Windows.Forms.Button
    $uninstallButton.Text = "Deinstallieren"
    $uninstallButton.Location = New-Object System.Drawing.Point(20, 330)
    $uninstallButton.Size = New-Object System.Drawing.Size(400, 35)
    $uninstallButton.ForeColor = [System.Drawing.Color]::DarkRed
    $popupForm.Controls.Add($uninstallButton)
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Schließen"
    $closeButton.Location = New-Object System.Drawing.Point(170, 383)
    $closeButton.Size = New-Object System.Drawing.Size(100, 30)
    $closeButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $popupForm.Controls.Add($closeButton)
    $popupForm.AcceptButton = $closeButton
    
    $openConfigFileButton.Add_Click({
        if (Test-Path $userConfigFile) {
            try { Invoke-Item $userConfigFile } catch { [System.Windows.Forms.MessageBox]::Show("Fehler: $($_.Exception.Message)") }
        } else { [System.Windows.Forms.MessageBox]::Show("Konfigurationsdatei nicht gefunden.") }
    })
    $openDataPathButton.Add_Click({
        if (Test-Path $dataDir) {
            try { Invoke-Item $dataDir } catch { [System.Windows.Forms.MessageBox]::Show("Fehler: $($_.Exception.Message)") }
        } else { [System.Windows.Forms.MessageBox]::Show("Datenverzeichnis nicht gefunden.") }
    })
    $exportDataButton.Add_Click({
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "ZIP-Archiv (*.zip)|*.zip"
        $saveFileDialog.Title = "Benutzerdaten exportieren"
        $saveFileDialog.FileName = "$($script:AppName)_Export_$(Get-Date -Format 'yyyy-MM-dd').zip"
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $exportPath = $saveFileDialog.FileName
            try {
                Compress-Archive -Path "$($dataDir)\*" -DestinationPath $exportPath -Force
                [System.Windows.Forms.MessageBox]::Show("Daten erfolgreich nach `"$exportPath`" exportiert.", "Export abgeschlossen", "OK", "Information")
            } catch { [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren: $($_.Exception.Message)") }
        }
    })
    $moveDataButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Wählen Sie ein neues, leeres Verzeichnis für die Anwendungsdaten."
        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $newPath = $folderBrowser.SelectedPath
            if ($newPath -eq $dataDir) {
                [System.Windows.Forms.MessageBox]::Show("Der ausgewählte Pfad ist bereits der aktuelle Speicherort.", "Information", "OK", "Information")
                return
            }
            if (Get-ChildItem -Path $newPath) {
                [System.Windows.Forms.MessageBox]::Show("Das ausgewählte Verzeichnis ist nicht leer. Bitte wählen Sie einen leeren Ordner aus, um Datenkonflikte zu vermeiden.", "Fehler", "OK", "Error")
                return
            }
            try {
                Move-Item -Path "$($dataDir)\*" -Destination $newPath -Force
                Remove-Item -Path $dataDir -Recurse -Force
                Set-ItemProperty -Path $regPath -Name "DataPath" -Value $newPath
                [System.Windows.Forms.MessageBox]::Show("Der Datenpfad wurde erfolgreich nach `"$newPath`" geändert. Das Programm wird nun neu gestartet.", "Erfolg", "OK", "Information")
                $scriptPath = $MyInvocation.MyCommand.Path
                Start-Process powershell -ArgumentList "-File `"$scriptPath`""
                $popupForm.Close()
                $mainForm.Close()
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Verschieben der Daten: $($_.Exception.Message)", "Fehler", "OK", "Error")
            }
        }
    })
    $autoBackupButton.Add_Click({ Show-AutoBackupPopup })
    $restoreBackupButton.Add_Click({
        # Schritt 1: ZIP-Datei auswählen
        $openDlg = New-Object System.Windows.Forms.OpenFileDialog
        $openDlg.Filter = "ZIP-Archiv (*.zip)|*.zip"
        $openDlg.Title = "Backup-ZIP auswählen..."
        if ($openDlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $zipPath = $openDlg.FileName

        try {
            # Schritt 2: ZIP in temporäres Verzeichnis entpacken und validieren
            $tempDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "PSC_$($script:AppName)_Restore_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
            New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
            Expand-Archive -Path $zipPath -DestinationPath $tempDir -Force

            # Prüfe ob UserData-Ordner vorhanden ist (Pflicht-Inhalt)
            $userDataSource = Join-Path -Path $tempDir -ChildPath "UserData"
            $hasUserData = Test-Path $userDataSource

            if (-not $hasUserData) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Die gewählte ZIP-Datei enthält keinen 'UserData'-Ordner und ist kein gültiges $($script:AppName) Backup.",
                    "Ungültiges Backup", "OK", "Error")
                Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                return
            }

            # Schritt 3: Bestehende Daten prüfen und Merge/Überschreiben fragen
            $existingDataCount = 0
            if (Test-Path $dataDir) {
                $existingDataCount = @(Get-ChildItem -Path $dataDir -Recurse -File -ErrorAction SilentlyContinue).Count
            }

            $restoreMode = "overwrite"
            if ($existingDataCount -gt 0) {
                $choice = [System.Windows.Forms.MessageBox]::Show(
                    "Es sind bereits $existingDataCount Dateien im aktuellen Datenverzeichnis vorhanden.`n`nMöchten Sie die bestehenden Daten vorher löschen und komplett durch das Backup ersetzen?`n`n[Ja] = Alles löschen und Backup wiederherstellen`n[Nein] = Backup-Daten ergänzen (bestehende Dateien werden überschrieben)`n[Abbrechen] = Nichts tun",
                    "Bestehende Daten gefunden", "YesNoCancel", "Question")
                if ($choice -eq "Cancel") {
                    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                    return
                }
                if ($choice -eq "Yes") { $restoreMode = "clean" }
                else { $restoreMode = "merge" }
            }

            # Schritt 4: Daten wiederherstellen
            # 4a: Bei clean-Modus bestehende Daten löschen
            if ($restoreMode -eq "clean" -and (Test-Path $dataDir)) {
                Remove-Item -Path "$($dataDir)\*" -Recurse -Force
            }

            # 4b: Datenverzeichnis sicherstellen
            if (-not (Test-Path $dataDir)) {
                New-Item -Path $dataDir -ItemType Directory -Force | Out-Null
            }

            # 4c: UserData-Inhalt kopieren
            # Prüfe ob der UserData-Ordner einen Unterordner mit App-Name enthält (Backup-Struktur: UserData\[AppName]\...)
            $innerAppDir = Join-Path -Path $userDataSource -ChildPath $script:AppName
            if (Test-Path $innerAppDir) {
                Copy-Item -Path "$innerAppDir\*" -Destination $dataDir -Recurse -Force
            } else {
                Copy-Item -Path "$userDataSource\*" -Destination $dataDir -Recurse -Force
            }

            # 4d: Registry wiederherstellen (falls vorhanden)
            $regRestored = $false
            $regFile = Join-Path -Path $tempDir -ChildPath "Registry_Backup.reg"
            if (Test-Path $regFile) {
                $regImportResult = & reg import $regFile 2>&1
                $regRestored = $true
            }

            # Registry: DataPath auf aktuellen Pfad setzen (verhindert Pfadkonflikte aus altem Backup)
            if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
            Set-ItemProperty -Path $regPath -Name "DataPath" -Value $dataDir

            # Schritt 5: Aufräumen
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue

            # Schritt 6: Zusammenfassung und Neustart
            $summary = "Wiederherstellung erfolgreich abgeschlossen!`n`n"
            $summary += "- Benutzerdaten: Wiederhergestellt ($restoreMode)`n"
            $summary += "- Registry: $(if ($regRestored) { 'Importiert' } else { 'Übersprungen (wird beim Start initialisiert)' })`n"
            $summary += "`nDas Programm wird nun neu gestartet."

            [System.Windows.Forms.MessageBox]::Show($summary, "Wiederherstellung abgeschlossen", "OK", "Information")

            # Neustart
            $scriptPath = $PSCommandPath
            if (-not $scriptPath) { $scriptPath = $MyInvocation.ScriptName }
            if ($scriptPath -and (Test-Path $scriptPath)) {
                Start-Process powershell -ArgumentList "-File `"$scriptPath`""
            }
            $popupForm.Close()
            $mainForm.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler bei der Wiederherstellung: $($_.Exception.Message)",
                "Fehler", "OK", "Error")
            if ($tempDir -and (Test-Path $tempDir)) {
                Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    })
    $uninstallButton.Add_Click({
        # Schritt 1: Bestätigung
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Möchten Sie $($script:AppName) wirklich vollständig deinstallieren?`n`nFolgende Daten werden unwiderruflich entfernt:`n- Alle gespeicherten Finanzplan-Daten`n- Registry-Einträge`n- Das Script selbst`n`nDieser Vorgang kann nicht rückgängig gemacht werden!",
            "Deinstallation bestätigen", "YesNo", "Warning")
        if ($confirm -ne "Yes") { return }

        # Schritt 2: Sicherung anbieten
        $backup = [System.Windows.Forms.MessageBox]::Show(
            "Möchten Sie vorher eine vollständige Sicherung aller Daten erstellen?`n`n(Enthält: Benutzerdaten, Registry-Einträge und das Script)",
            "Sicherung erstellen?", "YesNoCancel", "Question")
        if ($backup -eq "Cancel") { return }

        $scriptPath = $PSCommandPath
        if (-not $scriptPath) { $scriptPath = $MyInvocation.ScriptName }
        if (-not $scriptPath) { $scriptPath = & { $MyInvocation.ScriptName } }

        if ($backup -eq "Yes") {
            # Schritt 3: Zielpfad auswählen und ZIP erstellen
            $saveDlg = New-Object System.Windows.Forms.SaveFileDialog
            $saveDlg.Filter = "ZIP-Archiv (*.zip)|*.zip"
            $saveDlg.Title = "Sicherung speichern unter..."
            $saveDlg.FileName = "$($script:AppName)_Backup_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').zip"
            if ($saveDlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
            $backupZip = $saveDlg.FileName

            try {
                # Temporäres Staging-Verzeichnis
                $stagingDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "PSC_$($script:AppName)_Uninstall_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
                New-Item -Path $stagingDir -ItemType Directory -Force | Out-Null

                # 3a: Benutzerdaten kopieren
                if (Test-Path $dataDir) {
                    $dataBackupDir = Join-Path -Path $stagingDir -ChildPath "UserData"
                    Copy-Item -Path $dataDir -Destination $dataBackupDir -Recurse -Force
                }

                # 3b: Registry-Schlüssel exportieren
                $regExportFile = Join-Path -Path $stagingDir -ChildPath "Registry_Backup.reg"
                $regKeyFull = "HKCU\Software\PSC\$($script:AppName)"
                $regExportResult = & reg export $regKeyFull $regExportFile /y 2>&1
                if (-not (Test-Path $regExportFile)) {
                    # Fallback: Manueller Export als Text
                    $regExportFile = Join-Path -Path $stagingDir -ChildPath "Registry_Backup.txt"
                    $regValues = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
                    if ($regValues) { $regValues | Out-String | Set-Content -Path $regExportFile -Encoding UTF8 }
                }

                # 3c: Script-Datei kopieren
                if ($scriptPath -and (Test-Path $scriptPath)) {
                    Copy-Item -Path $scriptPath -Destination $stagingDir -Force
                }

                # 3d: ZIP erstellen
                if (Test-Path $backupZip) { Remove-Item -Path $backupZip -Force }
                Compress-Archive -Path "$stagingDir\*" -DestinationPath $backupZip -Force

                # Staging aufräumen
                Remove-Item -Path $stagingDir -Recurse -Force -ErrorAction SilentlyContinue

                [System.Windows.Forms.MessageBox]::Show(
                    "Sicherung erfolgreich erstellt:`n$backupZip`n`nDie Deinstallation wird nun fortgesetzt.",
                    "Sicherung abgeschlossen", "OK", "Information")
            } catch {
                $errMsg = $_.Exception.Message
                $abortChoice = [System.Windows.Forms.MessageBox]::Show(
                    "Fehler bei der Sicherung: $errMsg`n`nMöchten Sie die Deinstallation trotzdem fortsetzen?",
                    "Sicherungsfehler", "YesNo", "Error")
                if ($abortChoice -ne "Yes") { return }
            }
        }

        # Schritt 4: Daten entfernen
        try {
            # 4a: Datenverzeichnis löschen
            if (Test-Path $dataDir) {
                Remove-Item -Path $dataDir -Recurse -Force
            }

            # 4b: Registry-Schlüssel entfernen (inkl. übergeordnetem PSC-Schlüssel wenn leer)
            if (Test-Path $regPath) {
                Remove-Item -Path $regPath -Recurse -Force
            }
            $parentRegPath = "HKCU:\Software\PSC"
            if ((Test-Path $parentRegPath) -and @(Get-ChildItem -Path $parentRegPath -ErrorAction SilentlyContinue).Count -eq 0) {
                Remove-Item -Path $parentRegPath -Force -ErrorAction SilentlyContinue
            }

            # 4c: Script-Datei löschen (verzögert per Hintergrundjob)
            if ($scriptPath -and (Test-Path $scriptPath)) {
                $deleteCmd = "Start-Sleep -Seconds 2; Remove-Item -Path '$($scriptPath -replace "'","''")' -Force -ErrorAction SilentlyContinue"
                Start-Process powershell -ArgumentList "-WindowStyle Hidden -Command $deleteCmd" -WindowStyle Hidden
            }

            [System.Windows.Forms.MessageBox]::Show(
                "$($script:AppName) wurde erfolgreich deinstalliert.`nAlle Daten und Einstellungen wurden entfernt.`n`nDas Programm wird jetzt beendet.",
                "Deinstallation abgeschlossen", "OK", "Information")

            # Programm beenden
            $popupForm.Close()
            $mainForm.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler bei der Deinstallation: $($_.Exception.Message)`n`nEinige Daten konnten möglicherweise nicht entfernt werden.",
                "Fehler", "OK", "Error")
        }
    })
    $popupForm.ShowDialog()
}


# Hauptfunktion zum Erstellen und Anzeigen der GUI
function Show-MainForm {
    
    # Globale GUI-Elemente
    $Global:txtGesamtvermoegen = $null; $Global:txtBedarf = $null
    $Global:numHorizont = $null; $Global:txtDividenden = $null
    $Global:numSimulationsjahre = $null
    $Global:txtRendite1 = $null; $Global:txtRendite2 = $null; $Global:txtRendite3 = $null
    $Global:lblStatus = $null; $Global:gridErgebnis = $null
    $Global:btnExportCsv = $null; $Global:tblEingaben = $null
    $Global:btnStartSimulation = $null
    $Global:btnSaveConfig = $null
    $Global:btnEinstellungen = $null

    $script:divYearUpdating = $false  # Guard gegen rekursive SelectedIndexChanged-Aufrufe

    # Formular erstellen
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$($script:AppName)"
    $form.Size = New-Object System.Drawing.Size(1400, 960)
    $form.StartPosition = "CenterScreen"
    $form.MinimumSize = New-Object System.Drawing.Size(1200, 750)

    # TabControl erstellen
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = "Fill"
    
    # --- Tab 1: Finance-Planner ---
    $simulationTab = New-Object System.Windows.Forms.TabPage
    $simulationTab.Text = "Finance-Planner"
    $simulationTab.Padding = New-Object System.Windows.Forms.Padding(10)
    
    # Vertikales Haupt-Layout
    $mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $mainLayout.Dock = "Fill"
    $mainLayout.ColumnCount = 1
    $mainLayout.RowCount = 2 # 1. Top-Panel, 2. Grid
    
    # RowStyles
    $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle "AutoSize"))        # 0: Top-Panel
    $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle "Percent", 100))   # 1: Grid (füllt)

    # --- Zone 0: Top-Panel (Inputs + Actions/Status) ---
    $topPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $topPanel.Dock = "Fill"
    $topPanel.AutoSize = $true
    $topPanel.ColumnCount = 2
    $topPanel.RowCount = 1
    $topPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle "AutoSize"))
    $topPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle "Percent", 100)) # Rechte Spalte füllt

    # --- Zone 0a: Eingabefelder (links im Top-Panel) ---
    $tblEingaben = New-Object System.Windows.Forms.TableLayoutPanel
    $Global:tblEingaben = $tblEingaben
    $tblEingaben.AutoSize = $true
    $tblEingaben.Anchor = "Top, Left" # Linksbündig
    $tblEingaben.ColumnCount = 2
    $tblEingaben.RowCount = 8 # 8 Felder
    
    $tblEingaben.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle "Absolute", 180))
    $tblEingaben.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle "Absolute", 100))
    
    # Helper-Funktion
    $row = 0
    function Add-InputRow {
        param($label, $control)
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $label
        $lbl.Anchor = "Left"; $lbl.TextAlign = "MiddleLeft"; $lbl.AutoSize = $true
        $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        
        $control.Dock = "Fill"
        $control.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        
        $script:tblEingaben.Controls.Add($lbl, 0, $script:row)
        $script:tblEingaben.Controls.Add($control, 1, $script:row)
        $script:row++
    }
    
    # Eingabefelder erstellen
    $Global:txtGesamtvermoegen = New-Object System.Windows.Forms.TextBox
    Add-InputRow -label "Gesamtvermögen (€):" -control $txtGesamtvermoegen
    
    $Global:txtBedarf = New-Object System.Windows.Forms.TextBox
    Add-InputRow -label "Jährl. Finanzbedarf (€):" -control $txtBedarf
    
    $Global:numHorizont = New-Object System.Windows.Forms.NumericUpDown
    $numHorizont.Minimum = 2; $numHorizont.Maximum = 10
    Add-InputRow -label "Planungshorizont (J):" -control $numHorizont
    
    $Global:txtDividenden = New-Object System.Windows.Forms.TextBox
    Add-InputRow -label "Monatl. Dividenden (€):" -control $txtDividenden
    
    $Global:numSimulationsjahre = New-Object System.Windows.Forms.NumericUpDown
    $numSimulationsjahre.Minimum = 1; $numSimulationsjahre.Maximum = 100
    Add-InputRow -label "Simulationsdauer (J):" -control $numSimulationsjahre
    
    $Global:txtRendite1 = New-Object System.Windows.Forms.TextBox
    Add-InputRow -label "Rendite Topf 1 (%):" -control $txtRendite1
    
    $Global:txtRendite2 = New-Object System.Windows.Forms.TextBox
    Add-InputRow -label "Rendite Topf 2 (%):" -control $txtRendite2
    
    $Global:txtRendite3 = New-Object System.Windows.Forms.TextBox
    Add-InputRow -label "Rendite Topf 3 (%):" -control $txtRendite3

    # Eingabe-Tabelle zum Top-Panel (links) hinzufügen
    $topPanel.Controls.Add($tblEingaben, 0, 0)
    $script:row = $row

    # --- Zone 0b: Rechte Spalte (Buttons + Status) ---
    $rightColumnLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $rightColumnLayout.Dock = "Fill"
    $rightColumnLayout.ColumnCount = 1
    $rightColumnLayout.RowCount = 3 # 1. Button-Zeile, 2. Status, 3. Spacer
    $rightColumnLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle "AutoSize")) # Button-Zeile
    $rightColumnLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle "AutoSize")) # Status
    $rightColumnLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle "Percent", 100)) # Spacer
    
    # === ÄNDERUNG: Panel für die Button-Zeile (Start + Export + Einstellungen) ===
    $buttonRowPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $buttonRowPanel.AutoSize = $true
    $buttonRowPanel.ColumnCount = 5
    $buttonRowPanel.RowCount = 1
    $buttonRowPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle "Absolute", 40))
    $buttonRowPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle "AutoSize"))
    $buttonRowPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle "AutoSize"))
    $buttonRowPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle "AutoSize"))
    $buttonRowPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle "AutoSize"))
    $buttonRowPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle "AutoSize"))
    
    # Button "Speichern" (links neben Simulation starten)
    $Global:btnSaveConfig = New-Object System.Windows.Forms.Button
    $btnSaveConfig.Text = "Speichern"
    $btnSaveConfig.Size = New-Object System.Drawing.Size(150, 34)
    $btnSaveConfig.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnSaveConfig.Margin = New-Object System.Windows.Forms.Padding(0, 3, 0, 3)
    
    # Button "Simulation starten"
    $Global:btnStartSimulation = New-Object System.Windows.Forms.Button
    $btnStartSimulation.Text = "Simulation starten"
    $btnStartSimulation.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnStartSimulation.BackColor = [System.Drawing.Color]::PaleGreen
    $btnStartSimulation.Size = New-Object System.Drawing.Size(200, 34)
    $btnStartSimulation.Margin = New-Object System.Windows.Forms.Padding(10, 3, 0, 3)
    
    # Button "Ergebnisse als CSV exportieren"
    $Global:btnExportCsv = New-Object System.Windows.Forms.Button
    $btnExportCsv.Text = "Ergebnisse als CSV exportieren"
    $btnExportCsv.Enabled = $false
    $btnExportCsv.Size = New-Object System.Drawing.Size(230, 34)
    $btnExportCsv.Margin = New-Object System.Windows.Forms.Padding(10, 3, 0, 3)

    # Button "Einstellungen"
    $Global:btnEinstellungen = New-Object System.Windows.Forms.Button
    $btnEinstellungen.Text = "Einstellungen"
    $btnEinstellungen.Size = New-Object System.Drawing.Size(120, 34)
    $btnEinstellungen.Margin = New-Object System.Windows.Forms.Padding(10, 3, 0, 3)

    # Button "Hilfe"
    $Global:btnHilfe = New-Object System.Windows.Forms.Button
    $btnHilfe.Text = "Hilfe"
    $btnHilfe.Size = New-Object System.Drawing.Size(80, 34)
    $btnHilfe.Margin = New-Object System.Windows.Forms.Padding(10, 3, 0, 3)

    # "Beenden" als TabPage (sieht wie ein nativer Tab aus)
    $beendenTab = New-Object System.Windows.Forms.TabPage
    $beendenTab.Text = "✕  Beenden"
    # Buttons zur Button-Zeile hinzufügen
    $buttonRowPanel.Controls.Add($btnSaveConfig, 0, 0)
    $buttonRowPanel.Controls.Add($btnStartSimulation, 1, 0)
    $buttonRowPanel.Controls.Add($btnExportCsv, 2, 0)
    $buttonRowPanel.Controls.Add($btnEinstellungen, 3, 0)
    $buttonRowPanel.Controls.Add($btnHilfe, 4, 0)
    
    # Statusanzeige (lblStatus)
    $Global:lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = "Status: Bereit."
    $lblStatus.Font = New-Object System.Drawing.Font("Consolas", 9)
    $lblStatus.AutoSize = $true
    $lblStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#EEEEEE")
    $lblStatus.Padding = New-Object System.Windows.Forms.Padding(5)
    $lblStatus.TextAlign = "TopLeft"
    $lblStatus.Anchor = "Top, Left" # Linksbündig in seiner Zelle
    $lblStatus.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 0) # Abstand nach oben
    
    # Steuerelemente zur rechten Spalte hinzufügen
    $rightColumnLayout.Controls.Add($buttonRowPanel, 0, 0) # Button-Zeile
    $rightColumnLayout.Controls.Add($lblStatus, 0, 1)      # Status
    
    # Rechte Spalte zum Top-Panel hinzufügen
    $topPanel.Controls.Add($rightColumnLayout, 1, 0) # Rechts
    
    # Top-Panel (mit Inputs und Status) zum Haupt-Layout hinzufügen
    $mainLayout.Controls.Add($topPanel, 0, 0) # In Zeile 0

    # --- Zone 1: Ergebnis-Grid ---
    $Global:gridErgebnis = New-Object System.Windows.Forms.DataGridView
    $gridErgebnis.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 0)
    $gridErgebnis.Dock = "Fill"
    $gridErgebnis.ReadOnly = $true
    $gridErgebnis.AllowUserToAddRows = $false
    $gridErgebnis.AllowUserToDeleteRows = $false
    $gridErgebnis.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $gridErgebnis.AllowUserToResizeRows = $false
    $gridErgebnis.RowHeadersVisible = $false
    $gridErgebnis.ShowCellToolTips = $true # Tooltips aktivieren
    
    # Grid zum Haupt-Layout hinzufügen (Zeile 1)
    $mainLayout.Controls.Add($gridErgebnis, 0, 1)

    # Haupt-Layout zum Tab hinzufügen
    $simulationTab.Controls.Add($mainLayout)
    
    # --- Hilfetext als Variable (für Hilfe-Popup) ---
    $script:HilfeText = @"
Willkommen beim $($script:AppName)!

Dieses Tool hilft Ihnen, die 3-Töpfe-Finanzstrategie (bekannt durch Dr. Andreas Beck) zu simulieren. Sie können sehen, wie sich Ihr Vermögen über Jahre entwickelt, basierend auf Ihren Annahmen.

---
1. DAS PRINZIP (KURZ ERKLÄRT)
---
Die Strategie teilt Ihr Vermögen in drei Töpfe auf, um Sicherheit und Wachstum zu balancieren.

* **Topf 1 (Liquidität):** Ihr "Girokonto". Enthält den Finanzbedarf für 1 Jahr. (Niedrige Rendite, z.B. Geldmarkt).
* **Topf 2 (Sicherheit):** Ihr "Puffer". Enthält den Bedarf für weitere Jahre (siehe 'Planungshorizont'). (Mittlere Rendite, z.B. Anleihen).
* **Topf 3 (Wachstum):** Der gesamte Rest Ihres Vermögens. Dies ist der Motor für Ihr Wachstum. (Hohe Rendite, z.B. Aktien/ETFs).

**Der Ablauf (pro Jahr):**
1.  Alle Töpfe erhalten ihre angenommene Rendite.
2.  Ihr Jahresbedarf wird *ausschließlich* aus Topf 3 entnommen (zuerst Dividenden, dann Kapital).
3.  Am Jahresende werden Topf 1 und Topf 2 aus Topf 3 wieder aufgefüllt ("Rebalancing").

---
2. DIE EINGABEFELDER (LINKS)
---
Hier legen Sie die Regeln für die Simulation fest:

* **Gesamtvermögen (€):** Ihr komplettes Startkapital.
* **Jährl. Finanzbedarf (€):** Ihre jährlichen Lebenshaltungskosten, die Sie entnehmen müssen.
* **Planungshorizont (J):** Ihr *gesamter* Sicherheitspuffer. (Beispiel: 3 = 1 Jahr in Topf 1 + 2 Jahre in Topf 2).
* **Monatl. Dividenden (€):** Der monatliche Cashflow (z.B. Dividenden), den Topf 3 generiert. Dieser wird *zuerst* zur Deckung des Bedarfs genutzt, bevor Kapital verkauft wird.
* **Simulationsdauer (J):** Wie viele Jahre in die Zukunft soll gerechnet werden?
* **Rendite Topf 1/2/3 (%):** Die erwartete Jahresrendite für jeden Topf. Geben Sie die Zahl als Prozent ein (z.B. "6" für 6%).
* **Speichern:** Speichert die aktuellen Eingaben als Standard für den nächsten Start des Programms.

---
3. STEUERUNG & ANALYSE (RECHTS)
---
* **Simulation starten:** Führt die Berechnung durch und füllt die Tabelle (das Grid) unten.
* **Ergebnisse als CSV exportieren:** (Wird nach der Simulation klickbar) Speichert die Ergebnistabelle als CSV-Datei, die Sie z.B. in Excel öffnen können.
* **Einstellungen:** Öffnet ein Menü zur Verwaltung Ihrer Daten (Speicherort ändern, ZIP-Export, etc.).
* **Status-Box (Break-Even-Analyse):** Dieses Feld zeigt Ihnen nach dem Start der Simulation eine wichtige Analyse:
    * **Nettobedarf:** Der Teil Ihres Jahresbedarfs, der *nicht* durch Dividenden gedeckt ist.
    * **Zins-Überschuss:** Die Zinsen, die Ihre Töpfe 1 und 2 (der Puffer) voraussichtlich erwirtschaften.
    * **Netto-Last:** (Nettobedarf - Zins-Überschuss). Dies ist der Betrag, den Topf 3 *mindestens* an Rendite erwirtschaften muss, damit Ihr Kapital nicht schrumpft.
    * **Benötigtes Kapital...:** Zeigt, wie groß Topf 3 sein muss, um diese Netto-Last allein durch seine Rendite zu decken (der "Break-Even").

---
4. DIE ERGEBNISTABELLE (UNTEN)
---
Dies ist das Herzstück. Jede Zeile ist ein Jahr. Fahren Sie mit der Maus über die Spaltenüberschriften für eine Kurzerklärung (Tooltip).

* **... Start:** Kapital zu Jahresbeginn.
* **Dividenden / Rest-Entnahme:** Zeigt die Entnahmen aus Topf 3 zur Deckung Ihres Bedarfs.
* **... Ende:** Kapital am Jahresende (nach Rendite und Entnahme).
* **Rebalancing:** Zeigt die Umschichtungen, um Töpfe 1 & 2 aufzufüllen (z.B. "T3->T1: ...").

---
5. EINSTELLUNGEN (UNTEN)
---
* **Beenden:** Schließt das Programm.

---
tl;dr
---
Ihr Vermögen wird in 3 Töpfe aufgeteilt: Topf 1 = 1 Jahr Lebenshaltung (Tagesgeld), Topf 2 = Puffer für weitere Jahre (Anleihen), Topf 3 = alles andere (Aktien/ETFs). Jedes Jahr wird Ihr Bedarf komplett aus Topf 3 entnommen (Dividenden zuerst, dann Kapitalverkauf), danach werden Topf 1 & 2 per Rebalancing wieder aufgefüllt. Die Simulation zeigt Jahr für Jahr, ob Topf 3 groß genug bleibt, um das System dauerhaft zu tragen. Entscheidend ist die Break-Even-Analyse: Erwirtschaftet Topf 3 genug Rendite, um die Entnahmen zu kompensieren, wächst Ihr Vermögen — andernfalls schrumpft es. Eingaben links anpassen, "Simulation starten" klicken, Ergebnis in der Tabelle ablesen.
"@

    # --- Tab 2: Release-Notes (vormals Hilfe) ---
    $releaseNotesTab = New-Object System.Windows.Forms.TabPage
    $releaseNotesTab.Text = "Release-Notes"
    $releaseNotesTab.Padding = New-Object System.Windows.Forms.Padding(10)
    
    $releaseNotesTextBox = New-Object System.Windows.Forms.RichTextBox
    $releaseNotesTextBox.Dock = "Fill"
    $releaseNotesTextBox.ReadOnly = $true
    $releaseNotesTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $releaseNotesTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $releaseNotesTextBox.Text = ""
    
    $releaseNotesTab.Controls.Add($releaseNotesTextBox)

    # =================== ANALYSE TAB ===================
    $analyseTab = New-Object System.Windows.Forms.TabPage
    $analyseTab.Text = "Analyse"
    $analyseTab.Padding = New-Object System.Windows.Forms.Padding(5)

    $analyseInnerTabs = New-Object System.Windows.Forms.TabControl
    $analyseInnerTabs.Dock = "Fill"

    # ---- Sub-Tab: Daten ----
    $datTab = New-Object System.Windows.Forms.TabPage; $datTab.Text = "Daten"

    # GroupBox: Wertpapiere definieren
    $grpSec = New-Object System.Windows.Forms.GroupBox
    $grpSec.Text = "Wertpapiere definieren"; $grpSec.Location = New-Object System.Drawing.Point(5,5)
    $grpSec.Size = New-Object System.Drawing.Size(1350,280)

    $dgvSec = New-Object System.Windows.Forms.DataGridView
    $dgvSec.Location = New-Object System.Drawing.Point(10,22); $dgvSec.Size = New-Object System.Drawing.Size(980,240)
    $dgvSec.SelectionMode = "FullRowSelect"; $dgvSec.MultiSelect = $false
    $dgvSec.ReadOnly = $true; $dgvSec.AllowUserToAddRows = $false; $dgvSec.RowHeadersVisible = $false
    $dgvSec.AutoSizeColumnsMode = "Fill"
    @("Name","ISIN","WKN","Regionen","Sektoren","Industrien","Div. €/Anteil","Turnus") | ForEach-Object {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $_; $col.HeaderText = $_; $dgvSec.Columns.Add($col) | Out-Null
    }
    $grpSec.Controls.Add($dgvSec)

    $btnSecNew  = New-Object System.Windows.Forms.Button; $btnSecNew.Text = "Neu";       $btnSecNew.Location  = New-Object System.Drawing.Point(1000,22);  $btnSecNew.Size  = New-Object System.Drawing.Size(120,35)
    $btnSecEdit = New-Object System.Windows.Forms.Button; $btnSecEdit.Text = "Bearbeiten"; $btnSecEdit.Location = New-Object System.Drawing.Point(1000,67);  $btnSecEdit.Size = New-Object System.Drawing.Size(120,35)
    $btnSecDel  = New-Object System.Windows.Forms.Button; $btnSecDel.Text  = "Löschen";   $btnSecDel.Location  = New-Object System.Drawing.Point(1000,112); $btnSecDel.Size  = New-Object System.Drawing.Size(120,35)
    $btnSecDel.ForeColor = [System.Drawing.Color]::DarkRed
    $grpSec.Controls.AddRange(@($btnSecNew,$btnSecEdit,$btnSecDel))
    $datTab.Controls.Add($grpSec)

    # GroupBox: Portfolio Zusammenstellung
    $grpPort = New-Object System.Windows.Forms.GroupBox
    $grpPort.Text = "Portfolio Zusammenstellung"; $grpPort.Location = New-Object System.Drawing.Point(5,295)
    $grpPort.Size = New-Object System.Drawing.Size(1350,370)

    $dgvPort = New-Object System.Windows.Forms.DataGridView
    $dgvPort.Location = New-Object System.Drawing.Point(10,22); $dgvPort.Size = New-Object System.Drawing.Size(1320,280)
    $dgvPort.AllowUserToAddRows = $true; $dgvPort.AllowUserToDeleteRows = $true
    $dgvPort.AutoSizeColumnsMode = "Fill"; $dgvPort.RowHeadersVisible = $false

    $colPName = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $colPName.Name = "Wertpapier"; $colPName.HeaderText = "Wertpapier"; $colPName.FillWeight = 40
    $colPKurs = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPKurs.Name = "Tageskurs"; $colPKurs.HeaderText = "Tageskurs (€)"; $colPKurs.FillWeight = 20
    $colPStk  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPStk.Name = "Stueckzahl"; $colPStk.HeaderText = "Stückzahl"; $colPStk.FillWeight = 20
    $colPGes  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPGes.Name = "Gesamtwert"; $colPGes.HeaderText = "Gesamtwert (€)"; $colPGes.ReadOnly = $true; $colPGes.FillWeight = 20

    $dgvPort.Columns.AddRange($colPName,$colPKurs,$colPStk,$colPGes) | Out-Null
    $grpPort.Controls.Add($dgvPort)
    $dgvPort.Add_DataError({ param($s,$e) $e.ThrowException = $false })  # Unterdrückt ComboBox-Validierungsfehler bei Refresh

    $btnPortSave = New-Object System.Windows.Forms.Button; $btnPortSave.Text = "Speichern"
    $btnPortSave.Location = New-Object System.Drawing.Point(10,312); $btnPortSave.Size = New-Object System.Drawing.Size(160,35)
    $btnPortSave.BackColor = [System.Drawing.Color]::FromArgb(70,130,180); $btnPortSave.ForeColor = [System.Drawing.Color]::White
    $grpPort.Controls.Add($btnPortSave)
    $datTab.Controls.Add($grpPort)

    # ---- Helper: Create Distribution Sub-Tab ----
    function New-DistributionTab {
        param([string]$Title, [string]$CategoryKey)
        $tab = New-Object System.Windows.Forms.TabPage; $tab.Text = $Title

        $splitDist = New-Object System.Windows.Forms.SplitContainer
        $splitDist.Dock = "Fill"; $splitDist.Orientation = "Vertical"
        $splitDist.SplitterDistance = 50; $splitDist.IsSplitterFixed = $false

        # Left: DataGridView (replaces ListView - no ghost column)
        $lvDist = New-Object System.Windows.Forms.DataGridView
        $lvDist.Dock = "Fill"
        $lvDist.ReadOnly = $true
        $lvDist.AllowUserToAddRows = $false
        $lvDist.AllowUserToDeleteRows = $false
        $lvDist.RowHeadersVisible = $false
        $lvDist.SelectionMode = "FullRowSelect"
        $lvDist.AutoSizeColumnsMode = "Fill"
        $lvDist.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $colKat = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $colKat.Name = "Kategorie"; $colKat.HeaderText = "Kategorie"; $colKat.FillWeight = 65
        $colPct = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $colPct.Name = "Anteil"; $colPct.HeaderText = "Anteil %"; $colPct.FillWeight = 35
        $colPct.DefaultCellStyle.Alignment = "MiddleRight"
        $lvDist.Columns.AddRange($colKat, $colPct) | Out-Null
        $splitDist.Panel1.Controls.Add($lvDist)

        # Right: Pie Chart
        $chartDist = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
        $chartDist.Dock = "Fill"
        $chartDist.Titles.Add($Title) | Out-Null
        $chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
        $chartArea.Name = "main"
        $chartDist.ChartAreas.Add($chartArea)
        $serPie = New-Object System.Windows.Forms.DataVisualization.Charting.Series
        $serPie.Name = "dist"; $serPie.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
        $serPie["PieLabelStyle"] = "Outside"
        $serPie["PieLineColor"] = "Black"
        $chartDist.Series.Add($serPie)
        $splitDist.Panel2.Controls.Add($chartDist)

        $tab.Controls.Add($splitDist)

        # Refresh function – returns scriptblock with closure
        $refreshFn = [scriptblock]::Create(@"
            `$dist = Get-PortfolioWeightedDistribution -Category '$CategoryKey'
            `$lvDist.Rows.Clear()
            `$serPie.Points.Clear()
            `$sorted = `$dist.GetEnumerator() | Sort-Object Value -Descending
            foreach (`$kv in `$sorted) {
                `$lvDist.Rows.Add(`$kv.Key, ("{0:N2} %" -f `$kv.Value)) | Out-Null
                `$pt = `$serPie.Points.AddXY(`$kv.Key, `$kv.Value)
                `$serPie.Points[`$pt]["ExplodedPie"] = "false"
            }
"@)
        return @{ Tab = $tab; Lv = $lvDist; Chart = $chartDist; Series = $serPie; Refresh = $refreshFn }
    }

    $distReg = New-DistributionTab "Regionen"   "Regionen"
    $distSek = New-DistributionTab "Sektoren"   "Sektoren"
    $distInd = New-DistributionTab "Industrien" "Industrien"

    # ---- Sub-Tab: Dividende ----
    $divTab = New-Object System.Windows.Forms.TabPage; $divTab.Text = "Dividende"

    # Filter row
    $pnlDivFilter = New-Object System.Windows.Forms.Panel
    $pnlDivFilter.Location = New-Object System.Drawing.Point(0,0); $pnlDivFilter.Size = New-Object System.Drawing.Size(1360,45); $pnlDivFilter.Dock = "Top"

    $lblDivYear = New-Object System.Windows.Forms.Label; $lblDivYear.Text = "Jahr:"; $lblDivYear.Location = New-Object System.Drawing.Point(10,12); $lblDivYear.Size = New-Object System.Drawing.Size(40,22)
    $cmbDivYear = New-Object System.Windows.Forms.ComboBox; $cmbDivYear.DropDownStyle = "DropDownList"
    $cmbDivYear.Location = New-Object System.Drawing.Point(55,8); $cmbDivYear.Size = New-Object System.Drawing.Size(120,28)
    $curYear = (Get-Date).Year
    $cmbDivYear.Items.Add("YTD") | Out-Null
    $cmbDivYear.SelectedIndex = 0

    $lblDivMonAvg = New-Object System.Windows.Forms.Label; $lblDivMonAvg.Text = "Ø/Monat: 0,00 €"
    $lblDivMonAvg.Location = New-Object System.Drawing.Point(200,12); $lblDivMonAvg.Size = New-Object System.Drawing.Size(220,22)
    $lblDivMonAvg.Font = New-Object System.Drawing.Font($form.Font, [System.Drawing.FontStyle]::Bold)

    $vorjahr = (Get-Date).Year - 1
    $lblDivTotal = New-Object System.Windows.Forms.Label; $lblDivTotal.Text = "Dividende gesamt p.a. (Stand 31.12.$vorjahr): 0,00 €"
    $lblDivTotal.Location = New-Object System.Drawing.Point(440,12); $lblDivTotal.Size = New-Object System.Drawing.Size(400,22)
    $lblDivTotal.Font = New-Object System.Drawing.Font($form.Font, [System.Drawing.FontStyle]::Bold)

    $pnlDivFilter.Controls.AddRange(@($lblDivYear,$cmbDivYear,$lblDivMonAvg,$lblDivTotal))
    $divTab.Controls.Add($pnlDivFilter)

    # Bar chart: dividende per month
    $chartDiv = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $chartDiv.Location = New-Object System.Drawing.Point(0,45); $chartDiv.Size = New-Object System.Drawing.Size(1350,250)
    $caDiv = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea; $caDiv.Name = "bar"
    $caDiv.AxisX.Title = "Monat"; $caDiv.AxisY.Title = "Dividende (€)"
    $caDiv.AxisX.Interval = 1
    $chartDiv.ChartAreas.Add($caDiv)
    $serBar = New-Object System.Windows.Forms.DataVisualization.Charting.Series
    $serBar.Name = "divbar"; $serBar.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
    $serBar.Color = [System.Drawing.Color]::FromArgb(70,130,180)
    $serBar["PointWidth"] = "0.6"
    $chartDiv.Series.Add($serBar)
    $divTab.Controls.Add($chartDiv)

    # Dividend calendar (tiles like Haushalts-Tracker)
    $grpDivCal = New-Object System.Windows.Forms.GroupBox
    $grpDivCal.Text = "Dividenden-Kalender"; $grpDivCal.Location = New-Object System.Drawing.Point(0,305)
    $grpDivCal.Size = New-Object System.Drawing.Size(1350,355)

    $divCalTilePanel = New-Object System.Windows.Forms.TableLayoutPanel
    $divCalTilePanel.Dock = "Fill"; $divCalTilePanel.ColumnCount = 4; $divCalTilePanel.RowCount = 3
    for ($i = 0; $i -lt 4; $i++) { $divCalTilePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null }
    for ($i = 0; $i -lt 3; $i++) { $divCalTilePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) | Out-Null }

    $monthNamesDE = @("Januar","Februar","März","April","Mai","Juni","Juli","August","September","Oktober","November","Dezember")
    $script:divCalTiles    = [System.Collections.ArrayList]::new()
    $script:divCalContent  = [System.Collections.ArrayList]::new()
    $script:divCalAmounts  = [System.Collections.ArrayList]::new()

    for ($mi = 0; $mi -lt 12; $mi++) {
        $tile = New-Object System.Windows.Forms.Panel
        $tile.Dock = "Fill"; $tile.BorderStyle = "FixedSingle"; $tile.Margin = New-Object System.Windows.Forms.Padding(3)

        $mLbl = New-Object System.Windows.Forms.Label
        $mLbl.Text = $monthNamesDE[$mi]
        $mLbl.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
        $mLbl.Dock = "Top"; $mLbl.TextAlign = "MiddleCenter"; $mLbl.Height = 25

        $cLbl = New-Object System.Windows.Forms.Label
        $cLbl.Location = New-Object System.Drawing.Point(5,30); $cLbl.Size = New-Object System.Drawing.Size(180,65)
        $cLbl.ForeColor = [System.Drawing.Color]::FromArgb(40,40,40)

        $aLbl = New-Object System.Windows.Forms.Label
        $aLbl.Location = New-Object System.Drawing.Point(185,30); $aLbl.Size = New-Object System.Drawing.Size(120,65)
        $aLbl.TextAlign = "TopRight"
        $aLbl.Font = New-Object System.Drawing.Font($form.Font, [System.Drawing.FontStyle]::Bold)

        $tile.Controls.AddRange(@($mLbl,$cLbl,$aLbl))
        $divCalTilePanel.Controls.Add($tile)
        $script:divCalTiles.Add($tile)    | Out-Null
        $script:divCalContent.Add($cLbl)  | Out-Null
        $script:divCalAmounts.Add($aLbl)  | Out-Null
    }
    $grpDivCal.Controls.Add($divCalTilePanel)
    $divTab.Controls.Add($grpDivCal)


    # =================== Sub-Tab: Sparraten-Planung ===================
    $sparTab = New-Object System.Windows.Forms.TabPage; $sparTab.Text = "Sparraten-Planung"

    $sparSplit = New-Object System.Windows.Forms.SplitContainer
    $sparSplit.Dock = "Fill"; $sparSplit.Orientation = "Vertical"
    $sparSplit.SplitterDistance = 50

    # ---- LINKS: GroupBox Zusammenfassung ----
    $grpSpar = New-Object System.Windows.Forms.GroupBox
    $grpSpar.Text = "Jahresübersicht & Prognose"; $grpSpar.Dock = "Fill"

    $sparSummaryPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $sparSummaryPanel.Dock = "Fill"; $sparSummaryPanel.ColumnCount = 1; $sparSummaryPanel.RowCount = 3
    $sparSummaryPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle "AutoSize")) | Out-Null
    $sparSummaryPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle "Percent", 50)) | Out-Null
    $sparSummaryPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle "Percent", 50)) | Out-Null

    # Gesamtsparrate Label
    $lblSparGesamt = New-Object System.Windows.Forms.Label
    $lblSparGesamt.Text = "Gesamtsparrate p.a.: 0,00 €"
    $lblSparGesamt.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblSparGesamt.Margin = New-Object System.Windows.Forms.Padding(8,8,8,4)
    $lblSparGesamt.AutoSize = $true
    $sparSummaryPanel.Controls.Add($lblSparGesamt, 0, 0)

    # Per-ETF DGV
    $dgvSparSum = New-Object System.Windows.Forms.DataGridView
    $dgvSparSum.Dock = "Fill"; $dgvSparSum.ReadOnly = $true
    $dgvSparSum.AllowUserToAddRows = $false; $dgvSparSum.AllowUserToDeleteRows = $false
    $dgvSparSum.RowHeadersVisible = $false; $dgvSparSum.AutoSizeColumnsMode = "Fill"
    $dgvSparSum.Margin = New-Object System.Windows.Forms.Padding(8,0,8,4)
    $dgvSparSum.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $colSSName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSSName.HeaderText = "ETF / Wertpapier"; $colSSName.FillWeight = 55
    $colSSAmt  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSSAmt.HeaderText  = "Summe p.a. (€)";   $colSSAmt.FillWeight  = 45
    $colSSAmt.DefaultCellStyle.Alignment = "MiddleRight"
    $dgvSparSum.Columns.AddRange($colSSName, $colSSAmt) | Out-Null
    $sparSummaryPanel.Controls.Add($dgvSparSum, 0, 1)

    # Prognose Regionalverteilung DGV
    $grpSparReg = New-Object System.Windows.Forms.GroupBox
    $grpSparReg.Text = "Prognose Regionalverteilung (basierend auf Sparplan)"
    $grpSparReg.Dock = "Fill"; $grpSparReg.Margin = New-Object System.Windows.Forms.Padding(8,4,8,8)
    $dgvSparReg = New-Object System.Windows.Forms.DataGridView
    $dgvSparReg.Dock = "Fill"; $dgvSparReg.ReadOnly = $true
    $dgvSparReg.AllowUserToAddRows = $false; $dgvSparReg.AllowUserToDeleteRows = $false
    $dgvSparReg.RowHeadersVisible = $false; $dgvSparReg.AutoSizeColumnsMode = "Fill"
    $dgvSparReg.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $colSRReg = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSRReg.HeaderText = "Region";     $colSRReg.FillWeight = 65
    $colSRPct = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSRPct.HeaderText = "Anteil %";   $colSRPct.FillWeight = 35
    $colSRPct.DefaultCellStyle.Alignment = "MiddleRight"
    $dgvSparReg.Columns.AddRange($colSRReg, $colSRPct) | Out-Null
    $grpSparReg.Controls.Add($dgvSparReg)
    $sparSummaryPanel.Controls.Add($grpSparReg, 0, 2)

    $grpSpar.Controls.Add($sparSummaryPanel)
    $sparSplit.Panel1.Controls.Add($grpSpar)

    # ---- RECHTS: Sparplan-Tabelle (ETF x Monat) + Detailinfos ----
    $sparRightPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $sparRightPanel.Dock = "Fill"; $sparRightPanel.ColumnCount = 1; $sparRightPanel.RowCount = 3
    $sparRightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle "Percent", 50)) | Out-Null
    $sparRightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle "Percent", 50)) | Out-Null
    $sparRightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle "Absolute", 42)) | Out-Null

    $monthNames = @("Januar","Februar","März","April","Mai","Juni","Juli","August","September","Oktober","November","Dezember")
    $monthShortSpar = @("Jan","Feb","Mrz","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez")

    $dgvSpar = New-Object System.Windows.Forms.DataGridView
    $dgvSpar.Dock = "Fill"
    $dgvSpar.AllowUserToAddRows = $false; $dgvSpar.AllowUserToDeleteRows = $false
    $dgvSpar.RowHeadersVisible = $false; $dgvSpar.AutoSizeColumnsMode = "Fill"
    $dgvSpar.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $dgvSpar.Add_DataError({ param($s,$e) $e.ThrowException = $false })

    # Spalte: Wertpapier (ReadOnly)
    $colSparSec = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSparSec.Name = "Wertpapier"; $colSparSec.HeaderText = "ETF / Wertpapier"
    $colSparSec.ReadOnly = $true; $colSparSec.FillWeight = 22
    $colSparSec.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
    $dgvSpar.Columns.Add($colSparSec) | Out-Null

    # Hidden column: SecurityId für zuverlässige Zuordnung
    $colSparSecId = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSparSecId.Name = "SecId"; $colSparSecId.Visible = $false
    $dgvSpar.Columns.Add($colSparSecId) | Out-Null

    # 12 Monatsspalten
    for ($mi = 0; $mi -lt 12; $mi++) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = "M$($mi+1)"; $col.HeaderText = $monthShortSpar[$mi]; $col.FillWeight = 6.5
        $col.DefaultCellStyle.Alignment = "MiddleRight"
        $dgvSpar.Columns.Add($col) | Out-Null
    }

    # Summenspalte (ReadOnly)
    $colSparSum = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSparSum.Name = "Summe"; $colSparSum.HeaderText = "Summe p.a."; $colSparSum.ReadOnly = $true; $colSparSum.FillWeight = 9
    $colSparSum.DefaultCellStyle.Alignment = "MiddleRight"
    $colSparSum.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235,245,255)
    $colSparSum.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $dgvSpar.Columns.Add($colSparSum) | Out-Null

    $sparRightPanel.Controls.Add($dgvSpar, 0, 0)

    # ---- Detail-Info-Panel: Prognose neue Positionsgröße & Dividenden ----
    $grpSparDetail = New-Object System.Windows.Forms.GroupBox
    $grpSparDetail.Text = "Prognose: Neue Positionsgrößen & Dividenden durch Sparplan"
    $grpSparDetail.Dock = "Fill"
    $grpSparDetail.Margin = New-Object System.Windows.Forms.Padding(0,4,0,0)

    $dgvSparDetail = New-Object System.Windows.Forms.DataGridView
    $dgvSparDetail.Dock = "Fill"; $dgvSparDetail.ReadOnly = $true
    $dgvSparDetail.AllowUserToAddRows = $false; $dgvSparDetail.AllowUserToDeleteRows = $false
    $dgvSparDetail.RowHeadersVisible = $false; $dgvSparDetail.AutoSizeColumnsMode = "Fill"
    $dgvSparDetail.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $dgvSparDetail.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $dgvSparDetail.Add_DataError({ param($s,$e) $e.ThrowException = $false })

    $colSD_Name   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSD_Name.HeaderText   = "ETF / Wertpapier";       $colSD_Name.FillWeight   = 22
    $colSD_Name.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
    $colSD_AktStk = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSD_AktStk.HeaderText = "Akt. Stück";             $colSD_AktStk.FillWeight = 10
    $colSD_AktStk.DefaultCellStyle.Alignment = "MiddleRight"
    $colSD_NeuStk = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSD_NeuStk.HeaderText = "Neue Stück (ca.)";       $colSD_NeuStk.FillWeight = 11
    $colSD_NeuStk.DefaultCellStyle.Alignment = "MiddleRight"
    $colSD_NeuStk.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkBlue
    $colSD_NeuWrt = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSD_NeuWrt.HeaderText = "Neuer Gesamtwert (€)";   $colSD_NeuWrt.FillWeight = 15
    $colSD_NeuWrt.DefaultCellStyle.Alignment = "MiddleRight"
    $colSD_NeuWrt.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkBlue
    $colSD_AktDiv = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSD_AktDiv.HeaderText = "Akt. Dividende p.a.";    $colSD_AktDiv.FillWeight = 14
    $colSD_AktDiv.DefaultCellStyle.Alignment = "MiddleRight"
    $colSD_ZusDiv = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSD_ZusDiv.HeaderText = "Zusätzl. Div. p.a.";     $colSD_ZusDiv.FillWeight = 13
    $colSD_ZusDiv.DefaultCellStyle.Alignment = "MiddleRight"
    $colSD_ZusDiv.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGreen
    $colSD_NeuDiv = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSD_NeuDiv.HeaderText = "Neue Dividende p.a.";    $colSD_NeuDiv.FillWeight = 15
    $colSD_NeuDiv.DefaultCellStyle.Alignment = "MiddleRight"
    $colSD_NeuDiv.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $colSD_NeuDiv.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGreen

    $dgvSparDetail.Columns.AddRange($colSD_Name, $colSD_AktStk, $colSD_NeuStk, $colSD_NeuWrt, $colSD_AktDiv, $colSD_ZusDiv, $colSD_NeuDiv) | Out-Null
    $grpSparDetail.Controls.Add($dgvSparDetail)
    $sparRightPanel.Controls.Add($grpSparDetail, 0, 1)

    $btnSparSave = New-Object System.Windows.Forms.Button
    $btnSparSave.Text = "Sparplan speichern"; $btnSparSave.Dock = "Fill"
    $btnSparSave.BackColor = [System.Drawing.Color]::FromArgb(70,130,180)
    $btnSparSave.ForeColor = [System.Drawing.Color]::White
    $btnSparSave.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnSparSave.Margin = New-Object System.Windows.Forms.Padding(0,4,0,0)
    $sparRightPanel.Controls.Add($btnSparSave, 0, 2)

    $sparSplit.Panel2.Controls.Add($sparRightPanel)
    $sparTab.Controls.Add($sparSplit)

    # =================== Sub-Tab: ETF-Vorabpauschale (Prognose) ===================
    $vapTab = New-Object System.Windows.Forms.TabPage; $vapTab.Text = "ETF-Vorabpauschale (Prognose)"
    $vapTab.AutoScroll = $true

    $script:vapPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $script:vapPanel.Dock = "Fill"
    $script:vapPanel.FlowDirection = "TopDown"
    $script:vapPanel.WrapContents = $false
    $script:vapPanel.AutoScroll = $true
    $script:vapPanel.Padding = New-Object System.Windows.Forms.Padding(5)

    $vapTab.Controls.Add($script:vapPanel)

    function Refresh-VorabpauschaleTab {
        $script:vapPanel.Controls.Clear()
        $kultur = [System.Globalization.CultureInfo]::new("de-DE")

        if ($script:portfolioData.Securities.Count -eq 0) {
            $lblEmpty = New-Object System.Windows.Forms.Label
            $lblEmpty.Text = "Keine Wertpapiere definiert. Bitte zuerst im Tab 'Daten' Wertpapiere anlegen."
            $lblEmpty.AutoSize = $true; $lblEmpty.Padding = New-Object System.Windows.Forms.Padding(10)
            $script:vapPanel.Controls.Add($lblEmpty)
            return
        }

        # Ziel: 8 GroupBoxes sichtbar → Panel-Höhe ca. 860px, jede Box ~100px
        foreach ($sec in $script:portfolioData.Securities) {
            $pos = $script:portfolioData.Portfolio | Where-Object { $_.SecurityId -eq $sec.Id }
            $gesamtwert = 0.0
            if ($null -ne $pos) { $gesamtwert = [double]$pos.Tageskurs * [double]$pos.Stueckzahl }

            $grpVap = New-Object System.Windows.Forms.GroupBox
            $grpVap.Text = $sec.Name
            $grpVap.Size = New-Object System.Drawing.Size(1310, 98)
            $grpVap.Margin = New-Object System.Windows.Forms.Padding(3,3,3,2)

            $lblVapISIN = New-Object System.Windows.Forms.Label
            $lblVapISIN.Text = "ISIN: $($sec.ISIN)  |  WKN: $($sec.WKN)"
            $lblVapISIN.Location = New-Object System.Drawing.Point(12, 22); $lblVapISIN.Size = New-Object System.Drawing.Size(400, 20)
            $lblVapISIN.ForeColor = [System.Drawing.Color]::FromArgb(80,80,80)

            $lblVapWert = New-Object System.Windows.Forms.Label
            $lblVapWert.Text = "Aktueller Wert: {0:N2} €" -f $gesamtwert
            $lblVapWert.Location = New-Object System.Drawing.Point(12, 44); $lblVapWert.Size = New-Object System.Drawing.Size(300, 20)
            $lblVapWert.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

            $lblVapTurnus = New-Object System.Windows.Forms.Label
            $turnusText = $sec.Dividende.Turnus
            if ($turnusText -eq "keine Ausschüttung") {
                $turnusText = "keine Ausschüttung (thesaurierend)"
            } elseif ($turnusText -eq "vierteljährlich" -and $sec.Dividende.DividendenMonate -and $sec.Dividende.DividendenMonate.Count -gt 0) {
                $divSumme = ($sec.Dividende.DividendenMonate | ForEach-Object { [double]$_.Betrag } | Measure-Object -Sum).Sum
                $turnusText = "vierteljährlich ({0:N4} € p.a./Anteil)" -f $divSumme
            } elseif ($turnusText -eq "halbjährlich" -and $sec.Dividende.DividendenMonate -and $sec.Dividende.DividendenMonate.Count -gt 0) {
                $divSumme = ($sec.Dividende.DividendenMonate | ForEach-Object { [double]$_.Betrag } | Measure-Object -Sum).Sum
                $turnusText = "halbjährlich ({0:N4} € p.a./Anteil)" -f $divSumme
            } elseif ($turnusText -eq "monatlich" -and $sec.Dividende.DividendenMonate -and $sec.Dividende.DividendenMonate.Count -gt 0) {
                $divSumme = ($sec.Dividende.DividendenMonate | ForEach-Object { [double]$_.Betrag } | Measure-Object -Sum).Sum
                $turnusText = "monatlich ({0:N4} € p.a./Anteil)" -f $divSumme
            } elseif ($sec.Dividende.BetragProAnteil -gt 0) {
                $turnusText = "{0} ({1:N4} €/Anteil)" -f $sec.Dividende.Turnus, $sec.Dividende.BetragProAnteil
            }
            $lblVapTurnus.Text = "Dividende: $turnusText"
            $lblVapTurnus.Location = New-Object System.Drawing.Point(12, 66); $lblVapTurnus.Size = New-Object System.Drawing.Size(500, 20)

            $stk = if ($null -ne $pos) { [double]$pos.Stueckzahl } else { 0.0 }
            $lblVapStk = New-Object System.Windows.Forms.Label
            $lblVapStk.Text = "Stückzahl: {0:N4}" -f $stk
            $lblVapStk.Location = New-Object System.Drawing.Point(420, 22); $lblVapStk.Size = New-Object System.Drawing.Size(200, 20)

            # Button "Berechnung"
            $btnVapCalc = New-Object System.Windows.Forms.Button
            $btnVapCalc.Text = "Berechnung"
            $btnVapCalc.Location = New-Object System.Drawing.Point(1160, 20)
            $btnVapCalc.Size = New-Object System.Drawing.Size(130, 65)
            $btnVapCalc.BackColor = [System.Drawing.Color]::FromArgb(70,130,180)
            $btnVapCalc.ForeColor = [System.Drawing.Color]::White
            $btnVapCalc.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
            $btnVapCalc.FlatStyle = "Flat"
            $btnVapCalc.Tag = $sec.Id

            # Ergebnis-Label für kompakte Anzeige in der GroupBox
            $lblVapResult = New-Object System.Windows.Forms.Label
            $lblVapResult.Text = ""
            $lblVapResult.Location = New-Object System.Drawing.Point(640, 22); $lblVapResult.Size = New-Object System.Drawing.Size(500, 65)
            $lblVapResult.Font = New-Object System.Drawing.Font("Consolas", 8.5)
            $lblVapResult.ForeColor = [System.Drawing.Color]::FromArgb(0,100,0)

            $btnVapCalc.Add_Click({
                param($sender, $e)
                $secId = $sender.Tag
                $secObj = $script:portfolioData.Securities | Where-Object { $_.Id -eq $secId }
                $posObj = $script:portfolioData.Portfolio | Where-Object { $_.SecurityId -eq $secId }
                # Finde das zugehörige Ergebnis-Label im gleichen GroupBox
                $parentGrp = $sender.Parent
                $resultLabel = $parentGrp.Controls | Where-Object { $_ -is [System.Windows.Forms.Label] -and $_.Font.Name -eq "Consolas" }
                Show-VorabpauschalePopup -security $secObj -position $posObj -resultLabel $resultLabel
            })

            $grpVap.Controls.AddRange(@($lblVapISIN, $lblVapWert, $lblVapTurnus, $lblVapStk, $btnVapCalc, $lblVapResult))
            $script:vapPanel.Controls.Add($grpVap)
        }
    }

    # ---- Vorabpauschale Berechnungs-Popup ----
    function Show-VorabpauschalePopup {
        param($security, $position, $resultLabel = $null)
        $kultur = [System.Globalization.CultureInfo]::new("de-DE")
        $curYear = (Get-Date).Year

        $popup = New-Object System.Windows.Forms.Form
        $popup.Text = "ETF-Vorabpauschale berechnen – $($security.Name)"
        $popup.Size = New-Object System.Drawing.Size(620, 680)
        $popup.StartPosition = "CenterParent"
        $popup.FormBorderStyle = "FixedDialog"
        $popup.MaximizeBox = $false; $popup.MinimizeBox = $false

        $gesamtwert = 0.0
        if ($null -ne $position) { $gesamtwert = [double]$position.Tageskurs * [double]$position.Stueckzahl }

        # --- Eingaben GroupBox ---
        $grpInput = New-Object System.Windows.Forms.GroupBox
        $grpInput.Text = "Eingaben für die Berechnung"
        $grpInput.Location = New-Object System.Drawing.Point(10, 10); $grpInput.Size = New-Object System.Drawing.Size(585, 340)

        $yPos = 25
        function Add-VapField($label, $default, $yRef, $parent, $suffix = "") {
            $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = $label
            $lbl.Location = New-Object System.Drawing.Point(15, ($yRef + 4)); $lbl.Size = New-Object System.Drawing.Size(240, 22)
            $txt = New-Object System.Windows.Forms.TextBox; $txt.Text = $default
            $txt.Location = New-Object System.Drawing.Point(260, $yRef); $txt.Size = New-Object System.Drawing.Size(140, 25)
            $parent.Controls.AddRange(@($lbl, $txt))
            if ($suffix -ne "") {
                $lblS = New-Object System.Windows.Forms.Label; $lblS.Text = $suffix
                $lblS.Location = New-Object System.Drawing.Point(405, ($yRef + 4)); $lblS.Size = New-Object System.Drawing.Size(50, 22)
                $parent.Controls.Add($lblS)
            }
            return $txt
        }

        # Steuerjahr
        $lblSJ = New-Object System.Windows.Forms.Label; $lblSJ.Text = "Steuerjahr:"
        $lblSJ.Location = New-Object System.Drawing.Point(15, ($yPos+4)); $lblSJ.Size = New-Object System.Drawing.Size(240, 22)
        $cmbSteuerJahr = New-Object System.Windows.Forms.ComboBox; $cmbSteuerJahr.DropDownStyle = "DropDownList"
        for ($y = $curYear; $y -ge 2023; $y--) { $cmbSteuerJahr.Items.Add($y.ToString()) | Out-Null }
        $cmbSteuerJahr.SelectedIndex = 0
        $cmbSteuerJahr.Location = New-Object System.Drawing.Point(260, $yPos); $cmbSteuerJahr.Size = New-Object System.Drawing.Size(140, 25)
        $grpInput.Controls.AddRange(@($lblSJ, $cmbSteuerJahr))
        $yPos += 35

        # Basiszins (automatisch befüllt)
        $basiszinsMap = @{ "2023" = "2,55"; "2024" = "2,29"; "2025" = "2,53"; "2026" = "3,20" }
        $txtBasiszins = Add-VapField "Basiszins (von Bundesbank):" ($basiszinsMap["$curYear"]) $yPos $grpInput "%"
        $cmbSteuerJahr.Add_SelectedIndexChanged({
            $selYear = $cmbSteuerJahr.SelectedItem
            if ($basiszinsMap.ContainsKey($selYear)) { $txtBasiszins.Text = $basiszinsMap[$selYear] }
            else { $txtBasiszins.Text = "0,00" }
        })
        $yPos += 35

        # Fondswert Jahresanfang
        $txtWertAnfang = Add-VapField "Fondswert am 01.01 (€):" ($gesamtwert.ToString("N2", $kultur)) $yPos $grpInput "€"
        $yPos += 35

        # Fondswert Jahresende
        $txtWertEnde = Add-VapField "Fondswert am 31.12 (€):" ($gesamtwert.ToString("N2", $kultur)) $yPos $grpInput "€"
        $yPos += 35

        # Ausschüttungen
        $divJahr = 0.0
        if ($null -ne $position -and $null -ne $security) {
            $stk = [double]$position.Stueckzahl
            switch ($security.Dividende.Turnus) {
                "keine Ausschüttung" { $divJahr = 0.0 }
                "jährlich"        { $divJahr = $security.Dividende.BetragProAnteil * $stk }
                "halbjährlich"    {
                    if ($security.Dividende.DividendenMonate -and $security.Dividende.DividendenMonate.Count -gt 0) {
                        $divJahr = ($security.Dividende.DividendenMonate | ForEach-Object { [double]$_.Betrag } | Measure-Object -Sum).Sum * $stk
                    } else { $divJahr = $security.Dividende.BetragProAnteil * 2 * $stk }
                }
                "vierteljährlich" {
                    if ($security.Dividende.DividendenMonate -and $security.Dividende.DividendenMonate.Count -gt 0) {
                        $divJahr = ($security.Dividende.DividendenMonate | ForEach-Object { [double]$_.Betrag } | Measure-Object -Sum).Sum * $stk
                    } else { $divJahr = $security.Dividende.BetragProAnteil * 4 * $stk }
                }
                "monatlich" {
                    if ($security.Dividende.DividendenMonate -and $security.Dividende.DividendenMonate.Count -gt 0) {
                        $divJahr = ($security.Dividende.DividendenMonate | ForEach-Object { [double]$_.Betrag } | Measure-Object -Sum).Sum * $stk
                    }
                }
            }
        }
        $txtAusschuettungen = Add-VapField "Erhaltene Ausschüttungen (€):" ($divJahr.ToString("N2", $kultur)) $yPos $grpInput "€"
        $yPos += 40

        # Fondstyp (Teilfreistellung)
        $lblFondstyp = New-Object System.Windows.Forms.Label; $lblFondstyp.Text = "Fondstyp (Teilfreistellung):"
        $lblFondstyp.Location = New-Object System.Drawing.Point(15, ($yPos+4)); $lblFondstyp.Size = New-Object System.Drawing.Size(240, 22)
        $cmbFondstyp = New-Object System.Windows.Forms.ComboBox; $cmbFondstyp.DropDownStyle = "DropDownList"
        $cmbFondstyp.Items.AddRange(@(
            "Aktienfonds (30% TFS)",
            "Mischfonds (15% TFS)",
            "Immobilienfonds Inland (60% TFS)",
            "Immobilienfonds Ausland (80% TFS)",
            "Sonstiger Fonds (0% TFS)"
        ))
        $cmbFondstyp.SelectedIndex = 0
        $cmbFondstyp.Location = New-Object System.Drawing.Point(260, $yPos); $cmbFondstyp.Size = New-Object System.Drawing.Size(300, 25)
        $grpInput.Controls.AddRange(@($lblFondstyp, $cmbFondstyp))
        $yPos += 40

        # Kirchensteuer
        $chkKirchensteuer = New-Object System.Windows.Forms.CheckBox; $chkKirchensteuer.Text = "Kirchensteuer"
        $chkKirchensteuer.Location = New-Object System.Drawing.Point(15, $yPos); $chkKirchensteuer.Size = New-Object System.Drawing.Size(150, 25)
        $lblKiStSatz = New-Object System.Windows.Forms.Label; $lblKiStSatz.Text = "Satz:"
        $lblKiStSatz.Location = New-Object System.Drawing.Point(175, ($yPos+4)); $lblKiStSatz.Size = New-Object System.Drawing.Size(40, 22)
        $cmbKiStSatz = New-Object System.Windows.Forms.ComboBox; $cmbKiStSatz.DropDownStyle = "DropDownList"
        $cmbKiStSatz.Items.AddRange(@("8 % (Bayern/BaWü)", "9 % (übrige Länder)")); $cmbKiStSatz.SelectedIndex = 1
        $cmbKiStSatz.Location = New-Object System.Drawing.Point(220, $yPos); $cmbKiStSatz.Size = New-Object System.Drawing.Size(200, 25)
        $cmbKiStSatz.Enabled = $false
        $chkKirchensteuer.Add_CheckedChanged({ $cmbKiStSatz.Enabled = $chkKirchensteuer.Checked })
        $grpInput.Controls.AddRange(@($chkKirchensteuer, $lblKiStSatz, $cmbKiStSatz))

        $popup.Controls.Add($grpInput)

        # --- Ergebnis GroupBox ---
        $grpResult = New-Object System.Windows.Forms.GroupBox
        $grpResult.Text = "Berechnungsergebnis"
        $grpResult.Location = New-Object System.Drawing.Point(10, 360); $grpResult.Size = New-Object System.Drawing.Size(585, 230)

        $txtResult = New-Object System.Windows.Forms.TextBox
        $txtResult.Multiline = $true; $txtResult.ReadOnly = $true; $txtResult.ScrollBars = "Vertical"
        $txtResult.Location = New-Object System.Drawing.Point(10, 22); $txtResult.Size = New-Object System.Drawing.Size(565, 195)
        $txtResult.Font = New-Object System.Drawing.Font("Consolas", 9.5)
        $txtResult.BackColor = [System.Drawing.Color]::FromArgb(248,248,248)
        $grpResult.Controls.Add($txtResult)
        $popup.Controls.Add($grpResult)

        # --- Buttons ---
        $btnCalc = New-Object System.Windows.Forms.Button; $btnCalc.Text = "Berechnen"
        $btnCalc.Location = New-Object System.Drawing.Point(350, 600); $btnCalc.Size = New-Object System.Drawing.Size(120, 35)
        $btnCalc.BackColor = [System.Drawing.Color]::FromArgb(34,139,34); $btnCalc.ForeColor = [System.Drawing.Color]::White
        $btnCalc.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

        $btnClose = New-Object System.Windows.Forms.Button; $btnClose.Text = "Schließen"
        $btnClose.Location = New-Object System.Drawing.Point(480, 600); $btnClose.Size = New-Object System.Drawing.Size(110, 35)
        $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $popup.CancelButton = $btnClose

        $btnCalc.Add_Click({
            try {
                $ki = [System.Globalization.CultureInfo]::InvariantCulture
                # Parse inputs
                $basiszinsRaw = $txtBasiszins.Text.Replace(",",".")
                $basiszins = 0.0; [double]::TryParse($basiszinsRaw, [System.Globalization.NumberStyles]::Any, $ki, [ref]$basiszins) | Out-Null
                $basiszins = $basiszins / 100.0

                $wertAnfangRaw = $txtWertAnfang.Text -replace "\.", ""; $wertAnfangRaw = $wertAnfangRaw.Replace(",",".")
                $wertAnfang = 0.0; [double]::TryParse($wertAnfangRaw, [System.Globalization.NumberStyles]::Any, $ki, [ref]$wertAnfang) | Out-Null

                $wertEndeRaw = $txtWertEnde.Text -replace "\.", ""; $wertEndeRaw = $wertEndeRaw.Replace(",",".")
                $wertEnde = 0.0; [double]::TryParse($wertEndeRaw, [System.Globalization.NumberStyles]::Any, $ki, [ref]$wertEnde) | Out-Null

                $ausschRaw = $txtAusschuettungen.Text -replace "\.", ""; $ausschRaw = $ausschRaw.Replace(",",".")
                $ausschuettungen = 0.0; [double]::TryParse($ausschRaw, [System.Globalization.NumberStyles]::Any, $ki, [ref]$ausschuettungen) | Out-Null

                # Teilfreistellung
                $tfsMap = @(0.30, 0.15, 0.60, 0.80, 0.00)
                $tfs = $tfsMap[$cmbFondstyp.SelectedIndex]

                # Kirchensteuer
                $kiStSatz = 0.0
                if ($chkKirchensteuer.Checked) {
                    $kiStSatz = if ($cmbKiStSatz.SelectedIndex -eq 0) { 0.08 } else { 0.09 }
                }

                # --- Berechnung gemäß InvStG §18 ---
                # Schritt 1: Basisertrag
                $basisertrag = $wertAnfang * $basiszins * 0.7

                # Schritt 2: Wertsteigerung
                $wertsteigerung = [Math]::Max(0, $wertEnde - $wertAnfang)

                # Schritt 3: Vorabpauschale = min(Basisertrag, Wertsteigerung) - Ausschüttungen (min 0)
                $bemessungsgrundlage = [Math]::Min($basisertrag, $wertsteigerung)
                $vorabpauschale = [Math]::Max(0, $bemessungsgrundlage - $ausschuettungen)

                # Schritt 4: Teilfreistellung anwenden
                $steuerpflichtigeVAP = $vorabpauschale * (1 - $tfs)

                # Schritt 5: Steuern berechnen
                $abgeltungsteuer = 0.25
                $soli = 0.055  # 5,5% auf AbgSt
                # Bei Kirchensteuer: Abgeltungsteuer wird auf 25/(4+KiSt%) reduziert
                $effAbgSt = if ($kiStSatz -gt 0) { 0.25 / (1 + $kiStSatz) } else { 0.25 }
                $steuerGesamt = $steuerpflichtigeVAP * $effAbgSt
                $soliAbzug = $steuerGesamt * $soli
                $kiStAbzug = $steuerGesamt * $kiStSatz
                $steuerTotal = $steuerGesamt + $soliAbzug + $kiStAbzug

                # Ergebnis formatieren
                $k = [System.Globalization.CultureInfo]::new("de-DE")
                $steuerJahr = $cmbSteuerJahr.SelectedItem
                $folgejahr = ([int]$steuerJahr + 1)
                $result = @(
                    "═══ VORABPAUSCHALE-BERECHNUNG $steuerJahr ═══"
                    "Fällig im Januar $folgejahr"
                    ""
                    "Fondswert 01.01.:         {0,15:N2} €" -f $wertAnfang
                    "Fondswert 31.12.:         {0,15:N2} €" -f $wertEnde
                    "Wertsteigerung:           {0,15:N2} €" -f $wertsteigerung
                    ""
                    "Basiszins:                {0,14:N2} %"  -f ($basiszins * 100)
                    "Basisertrag (×0,7):       {0,15:N2} €" -f $basisertrag
                    ""
                    "Bemessungsgrundlage:      {0,15:N2} €" -f $bemessungsgrundlage
                    "- Ausschüttungen:         {0,15:N2} €" -f $ausschuettungen
                    "= Vorabpauschale:         {0,15:N2} €" -f $vorabpauschale
                    ""
                    "Teilfreistellung ({0:N0}%):  -{1,14:N2} €" -f ($tfs*100), ($vorabpauschale * $tfs)
                    "Steuerpflichtige VAP:     {0,15:N2} €" -f $steuerpflichtigeVAP
                    ""
                    "─── Steuerbelastung ───"
                    "Abgeltungsteuer (25%):    {0,15:N2} €" -f $steuerGesamt
                    "Solidaritätszuschlag:     {0,15:N2} €" -f $soliAbzug
                )
                if ($kiStSatz -gt 0) {
                    $result += "Kirchensteuer ({0:N0}%):     {1,15:N2} €" -f ($kiStSatz*100), $kiStAbzug
                }
                $result += ""
                $result += "══════════════════════════════════════"
                $result += "STEUER GESAMT:            {0,15:N2} €" -f $steuerTotal
                $result += "══════════════════════════════════════"

                if ($wertsteigerung -le 0) {
                    $result += ""
                    $result += "HINWEIS: Keine Wertsteigerung → keine"
                    $result += "Vorabpauschale fällig."
                }

                $txtResult.Text = $result -join "`r`n"

                # Ergebnis kompakt in die GroupBox des Haupt-Tabs schreiben
                if ($null -ne $resultLabel) {
                    $compactResult = "Steuerjahr $steuerJahr (fällig Jan $folgejahr)`r`n"
                    $compactResult += "Vorabpauschale: {0:N2} € | TFS ({1:N0}%): -{2:N2} €`r`n" -f $vorabpauschale, ($tfs*100), ($vorabpauschale * $tfs)
                    $compactResult += "► Steuer gesamt: {0:N2} €" -f $steuerTotal
                    $resultLabel.Text = $compactResult
                    if ($steuerTotal -le 0) { $resultLabel.ForeColor = [System.Drawing.Color]::FromArgb(100,100,100) }
                    else { $resultLabel.ForeColor = [System.Drawing.Color]::FromArgb(0,100,0) }
                }
            } catch {
                $txtResult.Text = "Fehler bei der Berechnung: $($_.Exception.Message)"
            }
        })

        $popup.Controls.AddRange(@($btnCalc, $btnClose))
        $popup.ShowDialog() | Out-Null
    }

    $analyseInnerTabs.TabPages.AddRange(@($datTab, $distReg.Tab, $distSek.Tab, $distInd.Tab, $divTab, $sparTab, $vapTab))
    $analyseTab.Controls.Add($analyseInnerTabs)

    # =================== Helper Refresh Functions ===================
    function Refresh-SecuritiesGrid {
        $dgvSec.Rows.Clear()
        $colPName.Items.Clear()
        foreach ($sec in $script:portfolioData.Securities) {
            $regStr  = ($sec.Regionen   | ForEach-Object { "{0}:{1:N0}%" -f $_.Name,$_.Prozent }) -join ", "
            $sekStr  = ($sec.Sektoren   | ForEach-Object { "{0}:{1:N0}%" -f $_.Name,$_.Prozent }) -join ", "
            $indStr  = ($sec.Industrien | ForEach-Object { "{0}:{1:N0}%" -f $_.Name,$_.Prozent }) -join ", "
            # Dividend info: Show individual amounts for quarterly/halbjährlich if available
            $divBpa = $sec.Dividende.BetragProAnteil.ToString("N4", [System.Globalization.CultureInfo]::new("de-DE"))
            if (($sec.Dividende.Turnus -eq "vierteljährlich" -or $sec.Dividende.Turnus -eq "halbjährlich" -or $sec.Dividende.Turnus -eq "monatlich") -and $sec.Dividende.DividendenMonate -and $sec.Dividende.DividendenMonate.Count -gt 0) {
                $divSumme = ($sec.Dividende.DividendenMonate | ForEach-Object { [double]$_.Betrag } | Measure-Object -Sum).Sum
                $divBpa = $divSumme.ToString("N4", [System.Globalization.CultureInfo]::new("de-DE")) + " (p.a.)"
            }
            $dgvSec.Rows.Add($sec.Name, $sec.ISIN, $sec.WKN, $regStr, $sekStr, $indStr, $divBpa, $sec.Dividende.Turnus) | Out-Null
            $colPName.Items.Add($sec.Name) | Out-Null
        }
    }

    function Refresh-PortfolioGrid {
        $dgvPort.Rows.Clear()
        foreach ($pos in $script:portfolioData.Portfolio) {
            $sec = $script:portfolioData.Securities | Where-Object { $_.Id -eq $pos.SecurityId }
            if ($null -ne $sec) {
                $ges = $pos.Tageskurs * $pos.Stueckzahl
                $row = $dgvPort.Rows.Add($sec.Name,
                    $pos.Tageskurs.ToString("N2",[System.Globalization.CultureInfo]::new("de-DE")),
                    $pos.Stueckzahl.ToString("N4",[System.Globalization.CultureInfo]::new("de-DE")),
                    $ges.ToString("N2",[System.Globalization.CultureInfo]::new("de-DE")))
            }
        }
    }

    function Update-DistributionTabs {
        # Refresh all three distribution sub-tabs
        foreach ($distTab in @($distReg, $distSek, $distInd)) {
            $distTab.Lv.Rows.Clear()
            $distTab.Series.Points.Clear()
            $catKey = switch ($distTab.Tab.Text) {
                "Regionen"   {"Regionen"}
                "Sektoren"   {"Sektoren"}
                "Industrien" {"Industrien"}
            }
            $dist = Get-PortfolioWeightedDistribution -Category $catKey
            $sorted = $dist.GetEnumerator() | Sort-Object Value -Descending
            foreach ($kv in $sorted) {
                $distTab.Lv.Rows.Add($kv.Key, ("{0:N2} %" -f $kv.Value)) | Out-Null
                $distTab.Series.Points.AddXY($kv.Key, [Math]::Round($kv.Value,2)) | Out-Null
            }
        }
    }

    function Update-DividendeTab {
        $kultur = [System.Globalization.CultureInfo]::new("de-DE")
        $md = Get-PortfolioDividendsByMonth
        $serBar.Points.Clear()
        $monthShort = @("Jan","Feb","Mrz","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez")
        $totalDiv = 0.0
        for ($m = 1; $m -le 12; $m++) {
            $val = $md[$m]
            $serBar.Points.AddXY($monthShort[$m-1], [Math]::Round($val,2)) | Out-Null
            $totalDiv += $val
        }
        $avgDiv = $totalDiv / 12.0
        $lblDivTotal.Text  = "Dividende gesamt p.a. (Stand 31.12.$((Get-Date).Year - 1)): {0:N2} €" -f $totalDiv
        $lblDivMonAvg.Text = "Ø/Monat: {0:N2} €" -f $avgDiv

        # Dropdown: nur Jahre mit Dividenden-Daten anzeigen (Guard verhindert Rekursion)
        $script:divYearUpdating = $true
        $prevSel = $cmbDivYear.SelectedItem
        $cmbDivYear.Items.Clear()
        $cmbDivYear.Items.Add("YTD") | Out-Null
        if ($totalDiv -gt 0) {
            # Aktuelles Jahr hinzufügen; weitere Jahre wenn Portfolio-Einstieg älter
            $startYear = $curYear
            $latestYear = $curYear
            for ($y = $latestYear; $y -ge $startYear; $y--) { $cmbDivYear.Items.Add([string]$y) | Out-Null }
        }
        if ($cmbDivYear.Items.Contains($prevSel)) { $cmbDivYear.SelectedItem = $prevSel }
        else { $cmbDivYear.SelectedIndex = 0 }
        $script:divYearUpdating = $false

        # Dividend calendar tiles
        for ($mi = 0; $mi -lt 12; $mi++) {
            $m = $mi + 1
            $divAmt = $md[$m]
            $tile = $script:divCalTiles[$mi]

            # Highlight months with dividend
            if ($divAmt -gt 0) {
                $tile.BackColor = [System.Drawing.Color]::FromArgb(235,248,235)
                # Build content: which securities pay in this month
                $lines = [System.Collections.ArrayList]::new()
                foreach ($pos in $script:portfolioData.Portfolio) {
                    if ([double]$pos.Stueckzahl -le 0) { continue }
                    $sec = $script:portfolioData.Securities | Where-Object { $_.Id -eq $pos.SecurityId }
                    if ($null -eq $sec) { continue }
                    $hasDM = ($null -ne $sec.Dividende.DividendenMonate -and $sec.Dividende.DividendenMonate.Count -gt 0)
                    if (-not ($sec.Dividende.BetragProAnteil -gt 0 -or $hasDM)) { continue }
                    $pays = $false
                    $secAmt = 0.0
                    switch ($sec.Dividende.Turnus) {
                        "jährlich"        { if ($m -eq 12) { $pays = $true; $secAmt = $sec.Dividende.BetragProAnteil * $pos.Stueckzahl } }
                        "halbjährlich"    {
                            if ($hasDM) {
                                $dmMatch = $sec.Dividende.DividendenMonate | Where-Object { [int]$_.Monat -eq $m } | Select-Object -First 1
                                if ($null -ne $dmMatch) { $pays = $true; $secAmt = [double]$dmMatch.Betrag * $pos.Stueckzahl }
                            } else {
                                if ($m -eq 6 -or $m -eq 12) { $pays = $true; $secAmt = $sec.Dividende.BetragProAnteil * $pos.Stueckzahl }
                            }
                        }
                        "vierteljährlich" {
                            if ($hasDM) {
                                $dmMatch = $sec.Dividende.DividendenMonate | Where-Object { [int]$_.Monat -eq $m } | Select-Object -First 1
                                if ($null -ne $dmMatch) { $pays = $true; $secAmt = [double]$dmMatch.Betrag * $pos.Stueckzahl }
                            } else {
                                if ($m -in @(3,6,9,12)) { $pays = $true; $secAmt = $sec.Dividende.BetragProAnteil * $pos.Stueckzahl }
                            }
                        }
                        "monatlich" {
                            if ($hasDM) {
                                $dmMatch = $sec.Dividende.DividendenMonate | Where-Object { [int]$_.Monat -eq $m } | Select-Object -First 1
                                if ($null -ne $dmMatch) { $pays = $true; $secAmt = [double]$dmMatch.Betrag * $pos.Stueckzahl }
                            }
                        }
                    }
                    if ($pays -and $secAmt -gt 0) {
                        $displayN = Get-SecDisplayName $sec
                        $lines.Add(("{0}: {1:N2} €" -f $displayN, $secAmt)) | Out-Null
                    }
                }
                $script:divCalContent[$mi].Text = ($lines -join "`n")
                $script:divCalAmounts[$mi].Text = $divAmt.ToString("N2", $kultur) + " €"
                $script:divCalAmounts[$mi].ForeColor = [System.Drawing.Color]::DarkGreen
            } else {
                $tile.BackColor = [System.Drawing.SystemColors]::Control
                $script:divCalContent[$mi].Text = "-"
                $script:divCalAmounts[$mi].Text = ""
            }
        }
    }


    # Sparplan-Grid-Werte → In-Memory-Datenmodell synchronisieren
    function Sync-SparplanFromGrid {
        $dgvSpar.EndEdit()
        $newSparplan = [System.Collections.ArrayList]::new()
        foreach ($row in $dgvSpar.Rows) {
            $secId = [string]$row.Cells["SecId"].Value
            if ([string]::IsNullOrWhiteSpace($secId)) { continue }
            $sec = $script:portfolioData.Securities | Where-Object { $_.Id -eq $secId }
            if ($null -eq $sec) { continue }
            $monate = [System.Collections.ArrayList]::new()
            for ($mi = 1; $mi -le 12; $mi++) {
                $cellVal = [string]$row.Cells["M$mi"].Value
                $v = 0.0
                # DE-Format: Tausender-Punkt entfernen, Dezimalkomma → Punkt
                $cleaned = (($cellVal -replace '€','').Trim() -replace '\.','') -replace ',','.'
                [double]::TryParse($cleaned, [System.Globalization.NumberStyles]::Any,
                    [System.Globalization.CultureInfo]::InvariantCulture, [ref]$v) | Out-Null
                $monate.Add($v) | Out-Null
            }
            $newSparplan.Add(@{ SecurityId = $sec.Id; Monate = $monate }) | Out-Null
        }
        $script:portfolioData.Sparplan = $newSparplan
    }

    function Refresh-SparplanGrid {
        # Eine Zeile pro Wertpapier; Monatswerte aus gespeichertem Sparplan laden
        $dgvSpar.Rows.Clear()
        $kultur = [System.Globalization.CultureInfo]::new("de-DE")
        $currentMonth = (Get-Date).Month  # 1-12
        $pastMonthColor = [System.Drawing.Color]::FromArgb(232,232,232)  # leichtes Grau
        foreach ($sec in $script:portfolioData.Securities) {
            $sp = $script:portfolioData.Sparplan | Where-Object { $_.SecurityId -eq $sec.Id }
            $displayName = Get-SecDisplayName $sec
            $rowVals = @($displayName, $sec.Id)
            $jahresSumme = 0.0
            for ($mi = 0; $mi -lt 12; $mi++) {
                $val = if ($null -ne $sp -and $sp.Monate -and $mi -lt $sp.Monate.Count) { [double]$sp.Monate[$mi] } else { 0.0 }
                $jahresSumme += $val
                $rowVals += if ($val -gt 0) { $val.ToString("N2", $kultur) } else { "" }
            }
            $rowVals += ($jahresSumme.ToString("N2", $kultur) + " €")
            $rowIdx = $dgvSpar.Rows.Add($rowVals)
            # Vergangene Monate grau einfärben
            for ($mi = 0; $mi -lt 12; $mi++) {
                if (($mi + 1) -lt $currentMonth) {
                    $dgvSpar.Rows[$rowIdx].Cells["M$($mi+1)"].Style.BackColor = $pastMonthColor
                }
            }
        }
    }

    function Update-SparplanTab {
        $kultur = [System.Globalization.CultureInfo]::new("de-DE")

        # Per-ETF Jahressummen aus Grid lesen (keyed by SecId)
        $etfTotals = @{}
        $totalSparPA = 0.0
        foreach ($row in $dgvSpar.Rows) {
            $secId = [string]$row.Cells["SecId"].Value
            if ([string]::IsNullOrWhiteSpace($secId)) { continue }
            $jahresSumme = 0.0
            for ($mi = 1; $mi -le 12; $mi++) {
                $cellVal = [string]$row.Cells["M$mi"].Value
                $v = 0.0
                # DE-Format: Tausender-Punkt entfernen, Dezimalkomma → Punkt
                $cleaned = (($cellVal -replace '€','').Trim() -replace '\.','') -replace ',','.'
                [double]::TryParse($cleaned, [System.Globalization.NumberStyles]::Any,
                    [System.Globalization.CultureInfo]::InvariantCulture, [ref]$v) | Out-Null
                $jahresSumme += $v
            }
            # Summenspalte live aktualisieren
            $row.Cells["Summe"].Value = if ($jahresSumme -gt 0) { $jahresSumme.ToString("N2", $kultur) + " €" } else { "" }
            $etfTotals[$secId] = $jahresSumme
            $totalSparPA += $jahresSumme
        }

        # Linke Seite: ETF-Zusammenfassung (Anzeige mit Display-Name)
        $dgvSparSum.Rows.Clear()
        foreach ($kv in ($etfTotals.GetEnumerator() | Where-Object { $_.Value -gt 0 })) {
            $sec = $script:portfolioData.Securities | Where-Object { $_.Id -eq $kv.Key }
            $displayName = Get-SecDisplayName $sec
            $dgvSparSum.Rows.Add($displayName, ("{0:N2} €" -f $kv.Value)) | Out-Null
        }
        $lblSparGesamt.Text = "Gesamtsparrate p.a.: {0:N2} €" -f $totalSparPA

        # ── Detail-Prognose: Neue Positionsgrößen & Dividenden ──
        $dgvSparDetail.Rows.Clear()
        foreach ($row in $dgvSpar.Rows) {
            $secId = [string]$row.Cells["SecId"].Value
            if ([string]::IsNullOrWhiteSpace($secId)) { continue }
            $sec = $script:portfolioData.Securities | Where-Object { $_.Id -eq $secId }
            if ($null -eq $sec) { continue }

            # Sparplan-Jahressumme für dieses WP
            $sparSumme = 0.0
            if ($etfTotals.ContainsKey($secId)) { $sparSumme = $etfTotals[$secId] }
            if ($sparSumme -le 0) { continue }

            # Aktuelle Position finden
            $pos = $script:portfolioData.Portfolio | Where-Object { $_.SecurityId -eq $sec.Id }
            $aktKurs  = if ($null -ne $pos) { [double]$pos.Tageskurs  } else { 0.0 }
            $aktStueck = if ($null -ne $pos) { [double]$pos.Stueckzahl } else { 0.0 }

            if ($aktKurs -le 0) { continue }

            # Neue Stückzahl & Wert
            $zusatzStueck = $sparSumme / $aktKurs
            $neueStueck   = $aktStueck + $zusatzStueck
            $neuerWert    = $neueStueck * $aktKurs

            # Dividenden-Berechnung (nur für ausschüttende ETFs)
            $istAusschuettend = ($sec.Dividende.Turnus -ne "keine Ausschüttung" -and $sec.Dividende.Turnus -ne "")
            $aktDivPA  = ""
            $zusDivPA  = ""
            $neuDivPA  = ""

            if ($istAusschuettend) {
                # Jahres-Dividende pro Anteil berechnen
                $divProAnteilPA = 0.0
                switch ($sec.Dividende.Turnus) {
                    "jährlich" {
                        $divProAnteilPA = [double]$sec.Dividende.BetragProAnteil
                    }
                    "halbjährlich" {
                        if ($sec.Dividende.DividendenMonate -and $sec.Dividende.DividendenMonate.Count -gt 0) {
                            foreach ($dm in $sec.Dividende.DividendenMonate) { $divProAnteilPA += [double]$dm.Betrag }
                        } else {
                            $divProAnteilPA = [double]$sec.Dividende.BetragProAnteil * 2
                        }
                    }
                    "vierteljährlich" {
                        if ($sec.Dividende.DividendenMonate -and $sec.Dividende.DividendenMonate.Count -gt 0) {
                            foreach ($dm in $sec.Dividende.DividendenMonate) { $divProAnteilPA += [double]$dm.Betrag }
                        } else {
                            $divProAnteilPA = [double]$sec.Dividende.BetragProAnteil * 4
                        }
                    }
                    "monatlich" {
                        if ($sec.Dividende.DividendenMonate -and $sec.Dividende.DividendenMonate.Count -gt 0) {
                            foreach ($dm in $sec.Dividende.DividendenMonate) { $divProAnteilPA += [double]$dm.Betrag }
                        }
                    }
                }

                $aktDivWert = $divProAnteilPA * $aktStueck
                $zusDivWert = $divProAnteilPA * $zusatzStueck
                $neuDivWert = $divProAnteilPA * $neueStueck

                $aktDivPA = "{0:N2} €" -f $aktDivWert
                $zusDivPA = "+{0:N2} €" -f $zusDivWert
                $neuDivPA = "{0:N2} €" -f $neuDivWert
            }

            $displayName = Get-SecDisplayName $sec
            $dgvSparDetail.Rows.Add(
                $displayName,
                $aktStueck.ToString("N4", $kultur),
                $neueStueck.ToString("N4", $kultur),
                ("{0:N2} €" -f $neuerWert),
                $aktDivPA,
                $zusDivPA,
                $neuDivPA
            ) | Out-Null
        }

        # ── Prognose Regionalverteilung: Gesamtportfolio + Sparplan ──
        # 1. Aktueller Portfoliowert pro Wertpapier (keyed by SecId)
        $secValues = @{}
        foreach ($pos in $script:portfolioData.Portfolio) {
            $sec = $script:portfolioData.Securities | Where-Object { $_.Id -eq $pos.SecurityId }
            if ($null -ne $sec) {
                if (-not $secValues.ContainsKey($sec.Id)) { $secValues[$sec.Id] = 0.0 }
                $secValues[$sec.Id] += $pos.Tageskurs * $pos.Stueckzahl
            }
        }
        # 2. Sparplan-Jahressummen addieren
        foreach ($kv in $etfTotals.GetEnumerator()) {
            if ($kv.Value -gt 0) {
                if (-not $secValues.ContainsKey($kv.Key)) { $secValues[$kv.Key] = 0.0 }
                $secValues[$kv.Key] += $kv.Value
            }
        }
        # 3. Gewichtete Regionalverteilung berechnen
        $totalKombiniert = ($secValues.Values | Measure-Object -Sum).Sum
        $dgvSparReg.Rows.Clear()
        if ($totalKombiniert -gt 0) {
            $regForecast = @{}
            foreach ($kv in $secValues.GetEnumerator()) {
                $sec = $script:portfolioData.Securities | Where-Object { $_.Id -eq $kv.Key }
                if ($null -ne $sec) {
                    $weight = $kv.Value / $totalKombiniert
                    foreach ($r in $sec.Regionen) {
                        if (-not $regForecast.ContainsKey($r.Name)) { $regForecast[$r.Name] = 0.0 }
                        $regForecast[$r.Name] += $r.Prozent * $weight
                    }
                }
            }
            foreach ($kv in ($regForecast.GetEnumerator() | Sort-Object Value -Descending)) {
                $dgvSparReg.Rows.Add($kv.Key, ("{0:N2} %" -f $kv.Value)) | Out-Null
            }
        }
    }

    # =================== Event Handlers: Analyse Tab ===================
    $btnSecNew.Add_Click({
        $newSec = Show-SecurityEditorDialog
        if ($null -ne $newSec) {
            $script:portfolioData.Securities.Add($newSec) | Out-Null
            Refresh-SecuritiesGrid
            Refresh-SparplanGrid
            Update-DistributionTabs
            Update-DividendeTab
        }
    })

    $btnSecEdit.Add_Click({
        if ($dgvSec.SelectedRows.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Bitte zuerst ein Wertpapier auswählen.", "Hinweis", "OK", "Warning"); return }
        $idx = $dgvSec.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:portfolioData.Securities.Count) { return }
        $existing = $script:portfolioData.Securities[$idx]
        $updated = Show-SecurityEditorDialog -existingSecurity $existing
        if ($null -ne $updated) {
            $script:portfolioData.Securities[$idx] = $updated
            Refresh-SecuritiesGrid
            Refresh-SparplanGrid
            Update-DistributionTabs
            Update-DividendeTab
        }
    })

    $btnSecDel.Add_Click({
        if ($dgvSec.SelectedRows.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Bitte zuerst ein Wertpapier auswählen.", "Hinweis", "OK", "Warning"); return }
        $idx = $dgvSec.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:portfolioData.Securities.Count) { return }
        $secId = $script:portfolioData.Securities[$idx].Id
        $result = [System.Windows.Forms.MessageBox]::Show("Wertpapier und zugehörige Portfolio-Positionen löschen?", "Löschen", "YesNo", "Warning")
        if ($result -eq "Yes") {
            $script:portfolioData.Securities.RemoveAt($idx)
            # Remove matching portfolio entries
            $toRemove = $script:portfolioData.Portfolio | Where-Object { $_.SecurityId -eq $secId }
            foreach ($r in @($toRemove)) { $script:portfolioData.Portfolio.Remove($r) | Out-Null }
            Refresh-PortfolioGrid
            Refresh-SecuritiesGrid
            Refresh-SparplanGrid
            Update-DistributionTabs
            Update-DividendeTab
        }
    })

    # Auto-calculate Gesamtwert in portfolio grid
    $dgvPort.Add_CellValueChanged({
        param($s,$e)
        if ($e.RowIndex -lt 0) { return }
        $row = $dgvPort.Rows[$e.RowIndex]
        $kursRaw = [string]$row.Cells["Tageskurs"].Value
        $stkRaw  = [string]$row.Cells["Stueckzahl"].Value
        $kurs = 0.0; $stk = 0.0
        [double]::TryParse(($kursRaw -replace ",","."), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$kurs) | Out-Null
        [double]::TryParse(($stkRaw  -replace ",","."), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$stk) | Out-Null
        $ges = $kurs * $stk
        $row.Cells["Gesamtwert"].Value = $ges.ToString("N2",[System.Globalization.CultureInfo]::new("de-DE"))
    })

    $btnPortSave.Add_Click({
        try {
            # Offene Zellbearbeitungen abschließen
            $dgvPort.EndEdit()
            # Sparplan-Grid ebenfalls synchronisieren, damit Save-PortfolioData keine Sparplan-Daten überschreibt
            Sync-SparplanFromGrid
            $newPortfolio = [System.Collections.ArrayList]::new()
            foreach ($row in $dgvPort.Rows) {
                if ($row.IsNewRow) { continue }
                $secName = [string]$row.Cells["Wertpapier"].Value
                if ([string]::IsNullOrWhiteSpace($secName)) { continue }
                $sec = $script:portfolioData.Securities | Where-Object { $_.Name -eq $secName }
                if ($null -eq $sec) { continue }
                $kurs = 0.0; $stk = 0.0
                [double]::TryParse(([string]$row.Cells["Tageskurs"].Value -replace ",","."), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$kurs) | Out-Null
                [double]::TryParse(([string]$row.Cells["Stueckzahl"].Value -replace ",","."), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$stk) | Out-Null
                $newPortfolio.Add(@{ SecurityId=$sec.Id; Tageskurs=$kurs; Stueckzahl=$stk }) | Out-Null
            }
            $script:portfolioData.Portfolio = $newPortfolio
            Save-PortfolioData
            Update-DistributionTabs
            Update-DividendeTab
            [System.Windows.Forms.MessageBox]::Show("Portfolio gespeichert!", "Gespeichert", "OK", "Information")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler: $($_.Exception.Message)", "Fehler", "OK", "Error")
        }
    })


    # Sparplan speichern
    $btnSparSave.Add_Click({
        try {
            Sync-SparplanFromGrid
            Save-PortfolioData
            [System.Windows.Forms.MessageBox]::Show("Sparplan gespeichert!", "Gespeichert", "OK", "Information")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler: $($_.Exception.Message)", "Fehler", "OK", "Error")
        }
    })

    # Live-Vorschau bei Zelländerung (ColumnIndex >= 2: überspringe Wertpapier + SecId)
    $dgvSpar.Add_CellValueChanged({ param($s,$e) if ($e.RowIndex -ge 0 -and $e.ColumnIndex -ge 2) { Update-SparplanTab } })

    $cmbDivYear.Add_SelectedIndexChanged({
        if (-not $script:divYearUpdating) { Update-DividendeTab }
    })

    $analyseInnerTabs.Add_SelectedIndexChanged({
        $idx = $analyseInnerTabs.SelectedIndex
        if ($idx -eq 1 -or $idx -eq 2 -or $idx -eq 3) { Update-DistributionTabs }
        elseif ($idx -eq 4) { Update-DividendeTab }
        elseif ($idx -eq 5) { Update-SparplanTab }
        elseif ($idx -eq 6) { Refresh-VorabpauschaleTab }
    })
    # =================== END ANALYSE TAB ===================

    # Tabs zum TabControl hinzufügen
    $tabControl.TabPages.Add($simulationTab)
    $tabControl.TabPages.Add($analyseTab)
    $tabControl.TabPages.Add($releaseNotesTab)
    
    $form.Controls.Add($tabControl)

    # Zweites mini-TabControl ganz rechts (Beenden)
    $tabControlClose = New-Object System.Windows.Forms.TabControl
    $tabControlClose.Size = New-Object System.Drawing.Size(105, 26)
    $tabControlClose.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 105), 0)
    $tabControlClose.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $tabControlClose.TabPages.Add($beendenTab)
    $form.Controls.Add($tabControlClose)
    $tabControlClose.BringToFront()
    
    # ---------- 5. Event-Handler (Logik mit GUI verbinden) ----------
    
    # Formular-Lade-Ereignis
    $form.Add_Load({
        Load-UserData
        Load-PortfolioData
        Refresh-SecuritiesGrid
        Refresh-PortfolioGrid
        Refresh-SparplanGrid
        Update-SparplanTab
        Update-DividendeTab
        
        $kultur = [System.Globalization.CultureInfo]::new("de-DE")
        
        $txtGesamtvermoegen.Text = $script:data.Finanzplan.Gesamtvermoegen.ToString("N2", $kultur)
        $txtBedarf.Text = $script:data.Finanzplan.JaehrlicherFinanzbedarf.ToString("N2", $kultur)
        $numHorizont.Value = [decimal]$script:data.Finanzplan.Planungshorizont
        $txtDividenden.Text = $script:data.Finanzplan.MonatlichesDividendeneinkommen.ToString("N2", $kultur)
        $numSimulationsjahre.Value = [decimal]$script:data.Finanzplan.Simulationsjahre
        
        $txtRendite1.Text = $script:data.Finanzplan.RenditeTopf1.ToString("N2", $kultur)
        $txtRendite2.Text = $script:data.Finanzplan.RenditeTopf2.ToString("N2", $kultur)
        $txtRendite3.Text = $script:data.Finanzplan.RenditeTopf3.ToString("N2", $kultur)
        
        $lblStatus.Text = "Status: Konfiguration geladen. Bereit zur Simulation."
    })
    
    # Tab-Klick "Beenden" -> Form schließen
    $tabControlClose.Add_MouseClick({ $form.Close() })


    
    # Klick-Handler für "Einstellungen"
    $btnEinstellungen.Add_Click({
        Show-SettingsPopup -mainForm $form
    })
    
    # Klick-Handler für "Hilfe" (öffnet Popup mit Hilfetext)
    $btnHilfe.Add_Click({
        $hilfeDlg = New-Object System.Windows.Forms.Form
        $hilfeDlg.Text = "$($script:AppName) – Hilfe"
        $hilfeDlg.Size = New-Object System.Drawing.Size(750, 650)
        $hilfeDlg.StartPosition = "CenterParent"
        $hilfeDlg.FormBorderStyle = "Sizable"
        $hilfeDlg.MinimizeBox = $false
        
        $hilfeRtb = New-Object System.Windows.Forms.RichTextBox
        $hilfeRtb.Dock = "Fill"
        $hilfeRtb.ReadOnly = $true
        $hilfeRtb.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $hilfeRtb.BorderStyle = [System.Windows.Forms.BorderStyle]::None
        $hilfeRtb.Text = $script:HilfeText
        
        $hilfeDlg.Controls.Add($hilfeRtb)
        $hilfeDlg.ShowDialog() | Out-Null
    })
    
    # Klick-Handler für "Speichern"
    $btnSaveConfig.Add_Click({
        try {
            $script:data.Finanzplan.Gesamtvermoegen = Parse-Number $txtGesamtvermoegen.Text
            $script:data.Finanzplan.JaehrlicherFinanzbedarf = Parse-Number $txtBedarf.Text
            $script:data.Finanzplan.Planungshorizont = [int]$numHorizont.Value
            $script:data.Finanzplan.MonatlichesDividendeneinkommen = Parse-Number $txtDividenden.Text
            $script:data.Finanzplan.Simulationsjahre = [int]$numSimulationsjahre.Value
            
            $script:data.Finanzplan.RenditeTopf1 = Parse-Number $txtRendite1.Text
            $script:data.Finanzplan.RenditeTopf2 = Parse-Number $txtRendite2.Text
            $script:data.Finanzplan.RenditeTopf3 = Parse-Number $txtRendite3.Text
            
            Save-UserData -data $script:data
            
            $lblStatus.Text = "Status: Konfiguration erfolgreich gespeichert."
            [System.Windows.Forms.MessageBox]::Show("Eingaben wurden als Standard gespeichert.", "Gespeichert", "OK", "Information")
            
        } catch {
            $lblStatus.Text = "FEHLER: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Speichern: $($_.Exception.Message)", "Fehler", "OK", "Error")
        }
    })
    
    # Klick-Handler für "Simulation starten"
    $btnStartSimulation.Add_Click({
        try {
            $p_StartKapital = Parse-Number $txtGesamtvermoegen.Text
            $p_Bedarf = Parse-Number $txtBedarf.Text
            $p_Horizont = [int]$numHorizont.Value
            $p_Dividenden = Parse-Number $txtDividenden.Text
            $p_Jahre = [int]$numSimulationsjahre.Value
            
            $p_R1 = (Parse-Number $txtRendite1.Text) / 100.0
            $p_R2 = (Parse-Number $txtRendite2.Text) / 100.0
            $p_R3 = (Parse-Number $txtRendite3.Text) / 100.0
            
            $statusText = Get-BreakEvenAnalyse -Bedarf $p_Bedarf -Horizont $p_Horizont -R1 $p_R1 -R2 $p_R2 -R3 $p_R3 -DividendenMonatlich $p_Dividenden
            $lblStatus.Text = $statusText
            
            $ergebnis = Start-Finanzsimulation -StartKapital $p_StartKapital `
                                              -Bedarf $p_Bedarf `
                                              -Horizont $p_Horizont `
                                              -R1 $p_R1 `
                                              -R2 $p_R2 `
                                              -R3 $p_R3 `
                                              -DividendenMonatlich $p_Dividenden `
                                              -Jahre $p_Jahre
                                              
            $gridErgebnis.DataSource = $null
            if ($null -ne $ergebnis -and $ergebnis.Count -gt 0) {
                $bindingList = [System.ComponentModel.BindingList[object]]::new($ergebnis)
                $gridErgebnis.DataSource = $bindingList
                
                $script:simulationsErgebnis = $ergebnis
                $btnExportCsv.Enabled = $true
                
                Format-DataGridView $gridErgebnis
            } else {
                $script:simulationsErgebnis = $null
                $btnExportCsv.Enabled = $false
            }
            
        } catch {
            $lblStatus.Text = "FEHLER: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Fehler bei der Simulation: $($_.Exception.Message)", "Fehler", "OK", "Error")
            $gridErgebnis.DataSource = $null
            $btnExportCsv.Enabled = $false
        }
    })
    
    # Klick-Handler für "CSV Export"
    $btnExportCsv.Add_Click({
        if ($null -eq $script:simulationsErgebnis) { 
            [System.Windows.Forms.MessageBox]::Show("Keine Simulationsdaten zum Exportieren vorhanden.", "Hinweis", "OK", "Warning")
            return 
        }
        
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV-Datei (*.csv)|*.csv"
        $saveDialog.FileName = "finanzsimulation_3_toepfe.csv"
        
        if ($saveDialog.ShowDialog() -eq "OK") {
            try {
                $Kultur = [System.Globalization.CultureInfo]::new("de-DE")
                $script:simulationsErgebnis | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Culture $Kultur
                [System.Windows.Forms.MessageBox]::Show("Export erfolgreich gespeichert: $($saveDialog.FileName)", "Erfolg", "OK", "Information")
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Export: $($_.Exception.Message)", "Fehler", "OK", "Error")
            }
        }
    })

    # AutoBackup: Start-Check + periodischer Scheduler
    $script:autoBackupRanThisSession = $false
    Invoke-AutoBackupIfDue

    # Timer alle 15 Minuten prüfen, ob Backup fällig ist
    $script:autoBackupTimer = New-Object System.Windows.Forms.Timer
    $script:autoBackupTimer.Interval = 15 * 60 * 1000  # 15 Minuten
    $script:autoBackupTimer.Add_Tick({ Invoke-AutoBackupIfDue })
    $script:autoBackupTimer.Start()

    $form.Add_FormClosing({
        if ($script:autoBackupTimer) {
            try { $script:autoBackupTimer.Stop(); $script:autoBackupTimer.Dispose() } catch { }
        }
    })

    # Formular anzeigen (letzter Schritt)
    $form.ShowDialog()
}

# ---------- 6. Skript-Start ----------
Show-MainForm

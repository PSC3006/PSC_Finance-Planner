# CLAUDE.md

Anweisungen für Claude Code in diesem Repository (PSC_Finance-Planner).

## Git- & PR-Workflow

- **Niemals direkt auf `main` arbeiten.** Jede Änderung erfolgt auf einem neuen,
  thematisch benannten Branch.
- Für jede Änderung wird ein **Pull Request als Draft** erstellt.
- **Niemals automatisch mergen.** Das Mergen erfolgt ausschließlich durch den
  Repo-Owner nach Review/Approval.
- **Commit-Messages werden auf Deutsch verfasst**, kurz und prägnant, mit Fokus
  auf das "Warum" der Änderung.

## Projekt-Hintergrund

PSC Finance-Planner ist ein einzelnes PowerShell-Skript mit einer Windows-Forms-GUI
zur Simulation der Vermögensentwicklung nach dem 3-Töpfe-Finanzprinzip
(Dr. Andreas Beck) (Hauptdatei: `Finance-Planner_*.ps1`).

- **Zielumgebung:** Windows (PowerShell, Windows Forms, Registry `HKCU`).
- **Funktionsumfang:** Tabs für Stammdaten (Wertpapiere, Regionen, Sektoren,
  Industrien), Sparraten-Planung mit Detail-Prognose pro ETF, Dividenden-Kalender
  und -Berechnung (alle Turnus-Varianten: monatlich, vierteljährlich, halbjährlich,
  jährlich, keine Ausschüttung), ETF-Vorabpauschale-Berechnung (nach InvStG §18,
  inkl. Abgeltungsteuer/Soli/Kirchensteuer), Finanzsimulation (3-Töpfe-Strategie)
  sowie Analyse-Tab mit Break-Even-Analyse und Portfolio-Verteilung.
- **Daten & Sicherheit:** Konfiguration und Portfoliodaten werden als JSON lokal
  gespeichert. Automatisches Backup-System (Scheduler) mit konfigurierbaren
  Intervallen (täglich/wöchentlich/monatlich/bei Programmstart), Format
  (ZIP, Ordner-Kopie oder beides), Bereinigung alter Backups sowie
  Registry-Export zur Wiederherstellung auf anderen Geräten.

## Code-Konventionen (PowerShell)

- **Funktionsnamen:** Verb-Noun in PascalCase (z.B. `Save-PortfolioData`,
  `Get-PortfolioDividendsByMonth`, `Start-Finanzsimulation`).
- **Einrückung:** 4 Leerzeichen, keine Tabs.
- **Variablen:** Script-weite Variablen mit `$script:`-Prefix und camelCase
  (z.B. `$script:AppName`, `$script:AppVersion`).
- **Sprache:** UI-Texte, Meldungen und Kommentare sind auf Deutsch.
- **Strukturierung:** Größere Abschnitte werden mit Kommentar-Headern
  abgegrenzt (`# ---------- Abschnitt ----------`).
- **Dokumentation:** Wichtige Funktionen erhalten Comment-Based-Help
  (`.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`).
- **Changelog:** Änderungen werden im Changelog-Kommentarblock am Dateianfang
  dokumentiert (Versionsnummer, Datum, Beschreibung).

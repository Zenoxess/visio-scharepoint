# PPSI Visio–SharePoint Mirror

Der Spiegel durchsucht einen konfigurierten Windows-Ordner rekursiv, konvertiert `.vsd`, `.vsdx` und `.vsdm` mit lokalem Microsoft Visio in PDF und spiegelt ausschließlich diese PDFs samt benötigter Ordnerstruktur in einen dedizierten SharePoint-Server-2019-Zielordner. Der weitere Entwicklungsschwerpunkt liegt auf der [Python-Fassung](python/README.md). Die PowerShell-Fassung in [`production`](production/) bleibt als Alternative erhalten.

## PowerShell-Betriebspaket

```text
production/
├─ Invoke-VisioSharePointMirror.ps1
├─ mirror.json
└─ README.md
```

Nach dem Kopieren dieses Ordners wird der Spiegel mit genau einem Betriebsablauf gestartet:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-VisioSharePointMirror.ps1 `
  -ConfigurationPath .\mirror.json
```

Konfiguration, Voraussetzungen, Löschschutz und Aufgabenplanung sind in [`production/README.md`](production/README.md) beschrieben.
Die ausgelieferte `production/mirror.json` enthält ausschließlich deutlich markierte Platzhalter und muss vor dem Einsatz vollständig ausgefüllt werden.

## Python-Version (Entwicklungsschwerpunkt)

[`python/README.md`](python/README.md) erklärt Einrichtung und Start der Python-Fassung. Sie verwendet dieselben sechs Konfigurationsschlüssel sowie `pywin32`, `requests` und Windows-SSPI. Windows und lokales Visio bleiben erforderlich. Die Python-Demo funktioniert unabhängig davon ohne Zusatzbibliotheken:

```powershell
py -3 python\demo_mirror.py
```

Beide Varianten verwenden für identisch konfigurierte Ziele dieselbe Windows-Sperre gegen parallele Läufe auf einem Host. Für den Regelbetrieb wird eine der beiden Varianten eingeplant.

## Ablauf in der Kommandozeile ansehen

Die unabhängige [Konsolen-Demo](demo/README.md) zeigt den Ablauf mit festen Beispieldaten. Sie benötigt weder Konfiguration noch Visio oder SharePoint und führt ausschließlich Konsolenausgaben aus:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\demo\Show-VisioSharePointMirror.ps1
```

Mit `-Scenario UploadError` wird ein Uploadfehler ohne anschließende Bereinigung dargestellt, mit `-Scenario EmptySource` eine leere Quelle. Alle Ausgaben sind als `[DEMO]` markiert. Die Demo prüft keine echte Umgebung und gehört nicht zum Betriebspaket.

## Repository-Bereiche

- `production/` enthält das eigenständige PowerShell-Betriebspaket.
- `python/` enthält die Python-Fassung als Entwicklungsschwerpunkt mit Konfiguration, Abhängigkeiten, Anleitung und separater Konsolen-Demo.
- `src/` enthält das frühere mehrstufige Pipeline-Gerüst einschließlich ausdrücklich markierter Legacy-Platzhalter zur Nachvollziehbarkeit und wird vom aktuellen Skript weder geladen noch benötigt.
- `demo/` enthält die eigenständige Konsolen-Demo ohne echte Verarbeitung.
- `tests/Test-Mirror.ps1` und `tests/test_python_mirror.py` prüfen die beiden Spiegel-Varianten mit simulierten SharePoint-Antworten. Die übrigen Tests betreffen das frühere Gerüst. Tests gehören nicht zum Betriebspaket.
- `docs/` enthält das frühere Entscheidungsregister und Projektnotizen.
- `config/` enthält Beispielkonfigurationen des früheren Gerüsts.

Das Betriebsskript ist eigenständig und lädt keine Dateien aus den anderen Repository-Bereichen. Es besitzt keine Validate-, Preview-, Simulate-, DryRun- oder Testmodi und führt keine inkrementelle State-Datei.

Die Regressionstests werden separat unter Windows PowerShell 5.1 ausgeführt. Sie laden nur Funktionsdefinitionen und ersetzen Netzwerkaufrufe durch Testantworten; ein echter Visio-/SharePoint-Integrationslauf wird damit nicht ersetzt:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Mirror.ps1
```

## Festgelegtes Verhalten

- Quelle: absoluter Windows-Ordner, beispielsweise `P:\Quelle`
- Zielabbildung: `P:\Quelle\A\B\Datei.vsdx` → `<SharePoint-Ziel>\Quelle\A\B\Datei.pdf`
- SharePoint-Anbindung: REST mit integrierter Windows-Authentifizierung
- Synchronisation: vollständiger Upload bei jedem Lauf, danach exakter Zielabgleich
- Löschung: ausschließlich per SharePoint-Papierkorb und niemals für den konfigurierten Zielordner selbst
- Sicherheit: Eine unvollständige Quelle, fehlerhafte Konvertierung, falsche Ziel-ID oder ein Uploadfehler verhindert die Papierkorbphase

Microsoft empfiehlt die unbeaufsichtigte Automation von Office-Desktopanwendungen nicht als serverseitige Architektur. Der geplante Betrieb benötigt deshalb einen dedizierten, lizenzierten und unter dem Ausführungskonto initialisierten Windows-/Visio-Host.

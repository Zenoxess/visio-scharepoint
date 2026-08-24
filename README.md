# PPSI Visio–SharePoint Mirror

Das aktuelle Betriebspaket liegt vollständig in [`production`](production/). Es durchsucht einen konfigurierten Windows-Ordner rekursiv, konvertiert `.vsd`, `.vsdx` und `.vsdm` mit lokalem Microsoft Visio in PDF und spiegelt ausschließlich diese PDFs samt benötigter Ordnerstruktur in einen dedizierten SharePoint-Server-2019-Zielordner.

## Aktuelles Betriebspaket

```text
production/
├─ Invoke-VisioSharePointMirror.ps1
├─ mirror.json
└─ README.md
```

Nach dem Kopieren dieses Ordners wird der Spiegel mit genau einem Betriebsablauf gestartet:

```powershell
pwsh -NoProfile -File .\Invoke-VisioSharePointMirror.ps1 `
  -ConfigurationPath .\mirror.json
```

Konfiguration, Voraussetzungen, Löschschutz und Aufgabenplanung sind in [`production/README.md`](production/README.md) beschrieben.

## Repository-Bereiche

- `production/` ist das einzige aktuelle und auszuliefernde Betriebspaket.
- `src/` enthält das frühere mehrstufige Pipeline-Gerüst zur Nachvollziehbarkeit und wird vom aktuellen Skript weder geladen noch benötigt.
- `tests/` enthält die Tests des früheren Gerüsts und gehört nicht zum Betriebspaket.
- `docs/` enthält das frühere Entscheidungsregister und Projektnotizen.
- `config/` enthält Beispielkonfigurationen des früheren Gerüsts.

Das neue Betriebsskript dot-sourct oder importiert keine Dateien aus diesen historischen Bereichen. Es besitzt keine Validate-, Preview-, Simulate-, DryRun- oder Testmodi und führt keine inkrementelle State-Datei.

## Festgelegtes Verhalten

- Quelle: absoluter Windows-Ordner, beispielsweise `P:\Quelle`
- Zielabbildung: `P:\Quelle\A\B\Datei.vsdx` → `<SharePoint-Ziel>\Quelle\A\B\Datei.pdf`
- SharePoint-Anbindung: REST mit integrierter Windows-Authentifizierung
- Synchronisation: vollständiger Upload bei jedem Lauf, danach exakter Zielabgleich
- Löschung: ausschließlich per SharePoint-Papierkorb und niemals für den konfigurierten Zielordner selbst
- Sicherheit: Eine unvollständige Quelle, fehlerhafte Konvertierung, falsche Ziel-ID oder ein Uploadfehler verhindert die Papierkorbphase

Microsoft empfiehlt die unbeaufsichtigte Automation von Office-Desktopanwendungen nicht als serverseitige Architektur. Der geplante Betrieb benötigt deshalb einen dedizierten, lizenzierten und unter dem Ausführungskonto initialisierten Windows-/Visio-Host.

# Zuordnung der PowerShell-Skripte

## `production` – früheres Pipeline-Gerüst

Dieser Ordner bleibt ausschließlich zur Nachvollziehbarkeit des früheren mehrstufigen Pipeline-Gerüsts erhalten. Er ist nicht mehr das aktuelle Betriebspaket und wird vom neuen Spiegel-Skript nicht geladen:

- `Invoke-VisioSharePointSync.Production.ps1` war der produktive Einstieg dieses früheren Gerüsts und führt intern ausschließlich `Execute` aus.
- `VisioSharePointSync.Core.ps1` validiert die fachliche Sync-Konfiguration.
- `VisioSharePointSync.Pipeline.Runtime.ps1` enthält Orchestrierung, Sicherheitsgates, State-, Datei- und Transportlogik.

Das aktuelle, eigenständige Betriebspaket liegt im Top-Level-Ordner `production/` des Repositories und besteht nur aus Skript, Konfiguration und Kurzanleitung.

## `development`

- `Invoke-VisioSharePointSync.Development.ps1` erlaubt ausschließlich `Validate` und `Simulate`.
- `VisioSharePointSync.Pipeline.Simulation.ps1` enthält die deterministischen In-Memory-Adapter.

Dieser Ordner darf nicht Bestandteil eines Produktivdeployments sein.

## `legacy`

- `Invoke-VisioSharePointSync.Legacy.ps1` ist der sichtbar gekennzeichnete Validate/DryRun-Erststand.
- `VisioSharePointSync.AdapterStubs.Legacy.ps1` enthält alte, absichtlich werfende Adapter-Stubs.

Die Legacy-Dateien bleiben nur für Nachvollziehbarkeit und Regressionstests erhalten. Neue Automatisierungen dürfen sie nicht aufrufen.

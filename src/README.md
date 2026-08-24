# Zuordnung der PowerShell-Skripte

## `production`

Dieser Ordner ist das vollständige spätere Betriebspaket und muss als Einheit kopiert werden:

- `Invoke-VisioSharePointSync.Production.ps1` ist der einzige produktive Einstieg und führt intern ausschließlich `Execute` aus.
- `VisioSharePointSync.Core.ps1` validiert die fachliche Sync-Konfiguration.
- `VisioSharePointSync.Pipeline.Runtime.ps1` enthält Orchestrierung, Sicherheitsgates, State-, Datei- und Transportlogik.

Keine Datei in diesem Ordner lädt Test-Fixtures oder die Simulationsadapter.

## `development`

- `Invoke-VisioSharePointSync.Development.ps1` erlaubt ausschließlich `Validate` und `Simulate`.
- `VisioSharePointSync.Pipeline.Simulation.ps1` enthält die deterministischen In-Memory-Adapter.

Dieser Ordner darf nicht Bestandteil eines Produktivdeployments sein.

## `legacy`

- `Invoke-VisioSharePointSync.Legacy.ps1` ist der sichtbar gekennzeichnete Validate/DryRun-Erststand.
- `VisioSharePointSync.AdapterStubs.Legacy.ps1` enthält alte, absichtlich werfende Adapter-Stubs.

Die Legacy-Dateien bleiben nur für Nachvollziehbarkeit und Regressionstests erhalten. Neue Automatisierungen dürfen sie nicht aufrufen.

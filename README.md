# PPSI Visio–SharePoint Sync

Dieses Repository enthält den ersten, bewusst sicheren Projektmeilenstein für einen späteren Visio-zu-PDF-Abgleich zwischen einem Netzlaufwerk und SharePoint. `Validate` bleibt rein lesend; `Simulate` durchläuft die vollständige Pipeline ausschließlich mit In-Memory-Adaptern. Reale UNC-, COM- und SharePoint-Aufrufe bleiben durch die Execute-Gates und sichtbare Implementierungsplatzhalter gesperrt.

## Enthaltene Artefakte

- `docs/Offene-Fragen.xlsx` ist das führende Entscheidungsregister für die technische Abstimmung.
- `docs/OFFENE-FRAGEN.md` ist der versionierbare Snapshot des initialen Fragenbestands vom 21.08.2026.
- `config/sync.example.json` zeigt das geplante Konfigurationsschema mit ausdrücklich offenen Werten.
- `config/runtime.example.json` beschreibt getrennt davon sichere Laufzeitvorgaben und ist standardmäßig nicht zur Ausführung freigegeben.
- `src/Invoke-VisioSharePointSync.ps1` validiert eine Konfiguration oder zeigt im Dry-Run ausschließlich die später vorgesehenen Phasen.
- `src/Invoke-VisioSharePointSync.Pipeline.ps1` ist ein separater Einstieg für die abgesicherte Pipeline-Vorbereitung mit `Validate`, `Simulate` und `Execute`.
- `src/Invoke-VisioSharePointSync.Production.ps1` ist der eigenständige Einstieg für den späteren Produktivbetrieb. Er besitzt keinen Modus-Schalter und startet ausschließlich den abgesicherten `Execute`-Pfad.
- `src/VisioSharePointSync.Pipeline.Runtime.ps1` enthält die gemeinsame Validierungs-, Sicherheits- und Ablauflogik ohne Fake-Adapter.
- `src/VisioSharePointSync.Pipeline.Simulation.ps1` enthält ausschließlich die In-Memory-Adapter und wird nur vom Pipeline-Einstieg geladen, niemals vom Produktiv-Einstieg.
- `tests/Test-Scaffold.ps1` prüft das Gerüst ohne externe Testmodule.
- `tests/Test-PipelineScaffold.ps1` prüft die Pipeline-Grenzen und insbesondere, dass keine unbeabsichtigten externen Zugriffe stattfinden.

## Sicherer Aufruf

Die Skripte sind für Windows PowerShell 5.1-kompatible Syntax ausgelegt und benötigen in diesem Meilenstein keine externen Module.

```powershell
# Konfiguration ausschließlich validieren (Standardmodus)
powershell.exe -NoProfile -File .\src\Invoke-VisioSharePointSync.ps1 `
  -ConfigurationPath .\config\sync.example.json

# Geplante Phasen anzeigen, weiterhin ohne operative Zugriffe
pwsh -NoProfile -File .\src\Invoke-VisioSharePointSync.ps1 `
  -ConfigurationPath .\config\sync.example.json `
  -Mode DryRun
```

Erwartete Exitcodes:

- `0`: Konfiguration ist vollständig und gültig.
- `2`: Pflichtentscheidungen sind offen oder die Konfiguration ist ungültig.
- `1`: Unerwarteter technischer Fehler beim Laden oder Validieren.

Die Beispielkonfiguration ist absichtlich unvollständig und liefert daher Exitcode `2` sowie die zugehörigen Entscheidungs-IDs.

Für eine programmgesteuerte Auswertung kann `src/VisioSharePointSync.Core.ps1` eingebunden und `Invoke-VisioSharePointSyncAssessment` aufgerufen werden. Das Ergebnisobjekt enthält ausschließlich:

- `IsValid`
- `Errors`
- `Warnings`
- `UnresolvedDecisionIds`
- `PlannedStages`

Im Modus `DryRun` sind die späteren Stufen `InventorySource`, `StageStableSource`, `ConvertVisioToPdf`, `EnsureSharePointFolders`, `UploadOrUpdatePdf`, `ReconcileLocalAndRemoteState`, `PersistSyncState` und `SummarizeRun` nur beschreibende Einträge; keine dieser Stufen wird ausgeführt.

## Separater Pipeline-Einstieg

Der bisherige Einstieg `src/Invoke-VisioSharePointSync.ps1` und sein Verhalten bleiben unverändert: Er kennt weiterhin ausschließlich `Validate` und `DryRun`, lädt keine ausführbaren Adapter und besitzt keinen `Execute`-Modus. Der neue Einstieg `src/Invoke-VisioSharePointSync.Pipeline.ps1` ergänzt dieses sichere Gerüst, ersetzt es aber nicht.

Die Pipeline erhält sowohl die fachliche Sync-Konfiguration als auch eine getrennte Laufzeitkonfiguration:

```powershell
# Beide Konfigurationen validieren; dies ist zugleich der Standardmodus
pwsh -NoProfile -File .\src\Invoke-VisioSharePointSync.Pipeline.ps1 `
  -ConfigurationPath .\tests\fixtures\complete.config.json `
  -RuntimeConfigurationPath .\config\runtime.example.json `
  -Mode Validate

# Plan und Bereitschaft ohne externe Seiteneffekte simulieren
pwsh -NoProfile -File .\src\Invoke-VisioSharePointSync.Pipeline.ps1 `
  -ConfigurationPath .\tests\fixtures\complete.config.json `
  -RuntimeConfigurationPath .\config\runtime.example.json `
  -Mode Simulate

# Zeigt mit der Beispielkonfiguration das sichere Blockieren eines Execute-Versuchs
pwsh -NoProfile -File .\src\Invoke-VisioSharePointSync.Pipeline.ps1 `
  -ConfigurationPath .\tests\fixtures\complete.config.json `
  -RuntimeConfigurationPath .\config\runtime.example.json `
  -Mode Execute `
  -AllowExternalSideEffects
```

Die Modi haben klar getrennte Aufgaben:

- `Validate` prüft beide Konfigurationen, ohne Pipeline-Adapter aufzurufen.
- `Simulate` durchläuft denselben Orchestrator wie der spätere Echtlauf mit deterministischem Fake-Inventar, Fake-PDF, Fake-Remote-ID/eTag, In-Memory-State und In-Memory-Log. Außer dem Lesen der zwei JSON-Dateien gibt es keine Datei-, COM- oder Netzwerkoperation.
- `Execute` ist für einen späteren operativen Lauf reserviert. Eine Freigabe erfordert gleichzeitig den CLI-Schalter `-AllowExternalSideEffects` und `Execution.Enabled: true` in der Runtime-Konfiguration. Zusätzlich müssen sämtliche für die gewählte Route benötigten Fähigkeiten den Status `Ready` besitzen.

Diese doppelte Execute-Freigabe und die Bereitschaftsprüfung arbeiten fail-closed: Fehlt eine Bedingung, endet die Pipeline vor externen Seiteneffekten mit Exitcode `3`. Der Schalter allein aktiviert daher nichts.

`config/runtime.example.json` verwendet sichere Standardwerte. `Execution.Enabled` ist `false`; `Conversion.AdapterKind`, `Conversion.AdapterVersion` und `Operations.QuarantinePath` enthalten den sichtbaren Wert `__PLACEHOLDER_REQUIRED__`, und `Conversion.ExternalExecutablePath` ist `null`. Solche Platzhalter werden weder automatisch ersetzt noch als Implementierung gewertet. Für lokale, nicht eingecheckte Laufzeitwerte ist `config/runtime.local.json` vorgesehen.

Der aktuelle Stand enthält weiterhin keinen produktiven Visio-Konverter, keine reale Authentifizierung und keinen ausführbaren Graph- oder SharePoint-Upload. Das Ersetzen von Platzhalterwerten macht eine Route allein nicht `Ready`; ein `Execute`-Versuch bleibt gesperrt, solange die benötigten Fähigkeiten nicht implementiert und bereit sind.

Exitcodes des separaten Pipeline-Einstiegs:

- `0`: Der angeforderte Modus wurde erfolgreich abgeschlossen.
- `1`: Ein unerwarteter interner Fehler ist aufgetreten.
- `2`: Eine Konfiguration ist ungültig, unvollständig oder nicht lesbar.
- `3`: Die Ausführung wurde sicher blockiert oder die gewählte Route ist nicht bereit.
- `4`: Ein erwarteter operativer Fehler oder Teilerfolg ist in einem später freigegebenen Lauf aufgetreten.

## Eigener Produktiv-Einstieg

Für den späteren unbeaufsichtigten Betrieb gibt es einen bewusst kleinen, separaten Einstieg ohne `Validate`-, `DryRun`- oder `Simulate`-Auswahl. Er lädt nur den Konfigurations-Core und die gemeinsame Runtime, nicht die getrennte Simulationsdatei. Den Modus setzt er intern fest auf `Execute`; dadurch bleibt die produktive Bedienoberfläche testfrei, ohne eine zweite, abweichende Kopie der Ablauf- und Sicherheitslogik zu erzeugen.

```powershell
pwsh -NoProfile -File .\src\Invoke-VisioSharePointSync.Production.ps1 `
  -ConfigurationPath .\config\sync.local.json `
  -RuntimeConfigurationPath .\config\runtime.local.json `
  -AllowExternalSideEffects
```

Die doppelte Freigabe bleibt erhalten: `-AllowExternalSideEffects` und `Runtime.Execution.Enabled: true` müssen gleichzeitig gesetzt sein. Danach müssen weiterhin sämtliche benötigten Adapterfähigkeiten und die Go-live-Freigabe den Status `Ready` besitzen. Bis die weiter unten genannten Integrationsplatzhalter implementiert und fachlich freigegeben sind, endet auch dieser Einstieg sicher mit Exitcode `3`, bevor UNC-, COM- oder SharePoint-Zugriffe stattfinden. Konfigurationen unter `tests/fixtures` sind ausschließlich Testdaten und dürfen für diesen Aufruf nicht verwendet werden.

## Konfiguration

Das Schema besteht aus `Source`, `Conversion`, `SharePoint`, `Sync` und `Operations`. Werte wie `Undecided`, `null` und leere Listen bleiben sichtbar, bis die verknüpfte Frage im Entscheidungsregister beantwortet wurde. `ResolvedDecisionIds` bestätigt zusätzlich die 18 als Blocker markierten Entscheidungen. Diese Liste ersetzt nicht die Dokumentation in Excel: Ohne jede Blocker-ID ist Exitcode `0` ausgeschlossen; unbekannte oder doppelte IDs werden abgewiesen.

Konfigurationsdateien müssen als gültiges, BOM-loses UTF-8 vorliegen. Die erlaubten Schlüssel sind je Abschnitt geschlossen definiert; doppelte JSON-Schlüssel werden vor der Deserialisierung abgewiesen. Passwörter, Client-Secrets, Tokens, private Schlüssel, vorautorisierte Upload-URLs und ähnlich geheime Werte sind unzulässig. Entsprechende, doppelte oder unbekannte Schlüssel werden mit generischer, redigierter Diagnose abgelehnt; benutzerkontrollierte Schlüssel und Werte erscheinen nicht in der Ausgabe.

Die Ziel- und Authentifizierungsfelder sind konditional:

- `SharePointOnline` mit `MicrosoftGraphV1` verwendet `SiteId`, `DriveId` und `TargetFolderId`.
- `SharePointRest` verwendet eine HTTPS-`SiteUrl`, `LibraryName` und den relativen `TargetFolderPath`; damit ist auch `SharePointServer` mit `WindowsIntegrated` ohne Entra-Dummywerte darstellbar.
- `CertificateAppOnly` verlangt Tenant-/Client-ID, Zertifikat-Thumbprint und Zertifikatspeicherort.
- `ManagedIdentity` verlangt keine fiktiven Zertifikatswerte; `WorkloadIdentity` verlangt Tenant-/Client-ID und einen sicheren lokalen `WorkloadIdentityFilePath`. `Delegated` wird im ausdrücklich unbeaufsichtigten Gerüst mit Verweis auf `SEC-001` abgewiesen.
- Graph- und REST-Zielfelder schließen sich gegenseitig aus. SharePoint Server ist in diesem Gerüst nur mit REST und `WindowsIntegrated` zulässig; SharePoint Online lehnt `WindowsIntegrated` ab.

Graph-Site-, Drive- und Ordner-IDs werden als undurchsichtige Einzelsegmentwerte validiert; der spätere Adapter muss sie dennoch URI-escapen oder als SDK-Parameter übergeben. Der Workload-Identity-Referenzpfad wird gemeinsam mit State-, Staging- und Logpfad auf Überschneidungen geprüft.

`Source.RootPath` wird ausschließlich syntaktisch als klassischer UNC-Pfad geprüft. Die Operations-Pfade müssen absolute, lokale Nicht-Root-Pfade sein und dürfen weder identisch sein noch ineinander liegen. Dabei werden keine Pfade aufgelöst oder auf Erreichbarkeit geprüft. Der Dateinamens-Regex besitzt über `Source.RegexTimeoutMilliseconds` einen verpflichtenden Laufzeitgrenzwert von 50 bis 5.000 Millisekunden.

## Entscheidungs- und Änderungsprozess

1. Technische Verantwortliche bearbeiten zuerst `docs/Offene-Fragen.xlsx`.
2. Eine Entscheidung erhält Status, Entscheidung, Begründung, Verantwortliche und Termin.
3. Nach einem abgestimmten Meilenstein wird `docs/OFFENE-FRAGEN.md` anhand der führenden Excel-Datei aktualisiert und mit einem neuen Standdatum versehen.
4. Die Beispielkonfiguration wird in eine nicht eingecheckte `config/sync.local.json` kopiert und mit den freigegebenen, nicht geheimen Werten befüllt.
5. Erst wenn alle Blocker geklärt und die Konfiguration erfolgreich validiert sind, werden produktive Scanner-, Konverter- und SharePoint-Adapter geplant und implementiert.

## Ausdrücklich noch nicht produktiv freigegeben

- tatsächlicher UNC-Zugriff über den öffentlichen `Execute`-Pfad; die rekursive Inventarisierungsfunktion ist implementiert, aber hinter der vollständigen Readiness-Prüfung gesperrt
- Automatisierung einer Visio-Installation oder PDF-Konvertierung
- Authentifizierung an Microsoft Entra ID oder SharePoint
- Anlage, Aktualisierung, Verschiebung oder Löschung von SharePoint-Inhalten
- Aufgabenplanung, Dauerbetrieb oder Alarmierung; lokale State-/Staging-Funktionen und ihre In-Memory-Abnahme sind vorhanden, werden produktiv jedoch erst nach vollständiger Freigabe verwendet

Die späteren Adaptergrenzen sind in `src/VisioSharePointSync.Adapters.ps1` dokumentiert, werden vom öffentlichen Einstiegspunkt nicht geladen und sind absichtlich nicht ausführbar.

## Tests

```powershell
pwsh -NoProfile -File .\tests\Test-Scaffold.ps1
powershell.exe -NoProfile -File .\tests\Test-Scaffold.ps1
pwsh -NoProfile -File .\tests\Test-PipelineScaffold.ps1
powershell.exe -NoProfile -File .\tests\Test-PipelineScaffold.ps1
```

Die Testläufe nutzen nur PowerShell-Bordmittel. `Test-Scaffold.ps1` prüft unter anderem vollständige, unvollständige und unzulässige Konfigurationen, das exakte 18-Blocker-Gate, BOM-loses UTF-8 mit Umlauten, konditionale Graph-/REST- und Authentifizierungsvarianten, geschlossene und redigierte Schlüsselvalidierung einschließlich doppelter JSON-Schlüssel, Graph-ID- und Pfadsyntax, den Regex-Timeout, stabile Exitcodes sowie die garantierte Nebenwirkungsfreiheit des Dry-Runs. `Test-PipelineScaffold.ps1` prüft zusätzlich die Runtime-Konfiguration, die physische Trennung des Produktiv-Einstiegs von den Simulationsadaptern, vollständige Bootstrap-Ergebnisse, Modusgrenzen, Execute-Sperren, die vollständige Fake-Adapter-Stagefolge, Commit-vor-State, unveränderte Dateien, Teilerfolg, Inventur-/State-Blockaden, Zielkollisionen, Remote-ID-/Provenienz-/eTag-Sicherheit, Retry mit Fake-Uhr, atomaren State-Replace und die Log-Positivliste.

## Technische Referenzen

- [Microsoft: Unbeaufsichtigte Automation von Office](https://learn.microsoft.com/en-us/office/client-developer/integration/considerations-unattended-automation-office-microsoft-365-for-unattended-rpa)
- [Microsoft Graph: Unterstützte Formatkonvertierungen](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format?view=graph-rest-1.0)
- [Microsoft: Resource Specific Consent und Sites.Selected](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins-modernize/understanding-rsc-for-msgraph-and-sharepoint-online)
- [Microsoft Graph: Kleine Uploads](https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0)
- [Microsoft Graph: Fortsetzbare Uploads](https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0)

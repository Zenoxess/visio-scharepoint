# PPSI Visio–SharePoint Sync

Dieses Repository enthält den ersten, bewusst sicheren Projektmeilenstein für einen späteren Visio-zu-PDF-Abgleich zwischen einem Netzlaufwerk und SharePoint. `Validate` bleibt rein lesend; `Simulate` durchläuft die vollständige Pipeline ausschließlich mit In-Memory-Adaptern. Reale UNC-, COM- und SharePoint-Aufrufe bleiben durch die Execute-Gates und sichtbare Implementierungsplatzhalter gesperrt.

## Verzeichnisstruktur

Die Skripte sind nach ihrem Einsatzzweck getrennt. Nur `src/production` gehört in ein späteres Betriebspaket:

```text
src/
├─ production/   Produktiveinstieg sowie seine beiden Laufzeitabhängigkeiten
├─ development/  Validierung, Simulation und In-Memory-Adapter
└─ legacy/       Veralteter Erststand und nicht ausführbare alte Adapter-Stubs
tests/           Automatisierte Tests und ausschließlich synthetische Fixtures
```

- `src/production` ist als Einheit kopierbar und enthält keine Simulationsdatei.
- `src/development` bietet ausschließlich `Validate` und `Simulate`; ein `Execute`-Parameter existiert dort nicht.
- `src/legacy` dient nur der Nachvollziehbarkeit und Kompatibilität. Für neue Abläufe dürfen diese Dateien nicht verwendet werden.
- `docs/Offene-Fragen.xlsx` ist das führende Entscheidungsregister; `docs/OFFENE-FRAGEN.md` ist sein versionierbarer Snapshot.
- `config/sync.example.json` und `config/runtime.example.json` sind sichere, bewusst nicht produktionsfähige Beispiele.

Eine kurze Zuordnung jeder Quelldatei steht zusätzlich in `src/README.md`. Alle Skripte bleiben mit Windows PowerShell 5.1 kompatibel und benötigen derzeit keine externen Module.

## Legacy-Einstieg – nur Altbestand

`src/legacy/Invoke-VisioSharePointSync.Legacy.ps1` ist sichtbar als Legacy gekennzeichnet und bleibt ausschließlich für das erste Validate/DryRun-Gerüst erhalten:

```powershell
powershell.exe -NoProfile -File .\src\legacy\Invoke-VisioSharePointSync.Legacy.ps1 `
  -ConfigurationPath .\config\sync.example.json

pwsh -NoProfile -File .\src\legacy\Invoke-VisioSharePointSync.Legacy.ps1 `
  -ConfigurationPath .\config\sync.example.json `
  -Mode DryRun
```

Die Beispielkonfiguration ist absichtlich unvollständig und liefert Exitcode `2`. Für programmgesteuerte Prüfungen kann `src/production/VisioSharePointSync.Core.ps1` eingebunden und `Invoke-VisioSharePointSyncAssessment` aufgerufen werden. Im Legacy-Modus `DryRun` werden die späteren Stufen ausschließlich beschrieben und nicht ausgeführt.

## Entwicklungs- und Abnahme-Einstieg

Der Development-Einstieg lädt Core und Runtime aus `src/production` sowie die getrennten In-Memory-Adapter aus `src/development`. Er erlaubt nur `Validate` und `Simulate`:

```powershell
# Beide Konfigurationen validieren
pwsh -NoProfile -File .\src\development\Invoke-VisioSharePointSync.Development.ps1 `
  -ConfigurationPath .\tests\fixtures\complete.config.json `
  -RuntimeConfigurationPath .\config\runtime.example.json `
  -Mode Validate

# Den vollständigen Orchestrator ohne externe Seiteneffekte simulieren
pwsh -NoProfile -File .\src\development\Invoke-VisioSharePointSync.Development.ps1 `
  -ConfigurationPath .\tests\fixtures\complete.config.json `
  -RuntimeConfigurationPath .\config\runtime.example.json `
  -Mode Simulate
```

`Validate` ruft keine Adapter auf. `Simulate` verwendet Fake-Inventar, Fake-PDF, Fake-Remote-ID/eTag, In-Memory-State und In-Memory-Log. Außer dem Lesen der beiden JSON-Dateien gibt es keine Datei-, COM- oder Netzwerkoperation. Konfigurationen unter `tests/fixtures` sind ausschließlich Testdaten.

## Produktiv-Einstieg

Für den späteren unbeaufsichtigten Betrieb existiert ausschließlich `src/production/Invoke-VisioSharePointSync.Production.ps1`. Er besitzt keinen Modus-Schalter, lädt nur die beiden Dateien aus demselben Ordner und setzt intern fest `Execute`:

```powershell
pwsh -NoProfile -File .\src\production\Invoke-VisioSharePointSync.Production.ps1 `
  -ConfigurationPath .\config\sync.local.json `
  -RuntimeConfigurationPath .\config\runtime.local.json `
  -AllowExternalSideEffects
```

Die doppelte Freigabe bleibt erhalten: `-AllowExternalSideEffects` und `Runtime.Execution.Enabled: true` müssen gleichzeitig gesetzt sein. Danach müssen sämtliche benötigten Adapterfähigkeiten und die Go-live-Freigabe den Status `Ready` besitzen. Fehlt eine Bedingung, endet der Aufruf fail-closed mit Exitcode `3`, bevor Lock, UNC-, COM-, lokale Schreib- oder SharePoint-Zugriffe stattfinden.

`config/runtime.example.json` verwendet sichere Standardwerte: `Execution.Enabled` ist `false`, und offene Laufzeitwerte tragen `__PLACEHOLDER_REQUIRED__`. Der aktuelle Stand enthält weiterhin keinen produktiven Visio-Konverter, keine reale Authentifizierung und keinen freigegebenen Graph- oder SharePoint-Upload. Das Ersetzen von Konfigurationswerten allein macht die Adapter nicht `Ready`.

Exitcodes des Produktiv-Einstiegs:

- `0`: Der spätere operative Lauf wurde erfolgreich abgeschlossen.
- `1`: Ein unerwarteter interner Fehler ist aufgetreten.
- `2`: Eine Konfiguration ist ungültig, unvollständig oder nicht lesbar.
- `3`: Die Ausführung wurde vor Nebenwirkungen sicher blockiert.
- `4`: Ein erwarteter operativer Fehler oder Teilerfolg ist aufgetreten.

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

Die alten Adapter-Stubs liegen sichtbar als Legacy unter `src/legacy/VisioSharePointSync.AdapterStubs.Legacy.ps1`. Sie werden von keinem aktuellen Einstieg geladen und sind absichtlich nicht ausführbar.

## Tests

```powershell
pwsh -NoProfile -File .\tests\Test-LegacyScaffold.ps1
powershell.exe -NoProfile -File .\tests\Test-LegacyScaffold.ps1
pwsh -NoProfile -File .\tests\Test-DevelopmentPipeline.ps1
powershell.exe -NoProfile -File .\tests\Test-DevelopmentPipeline.ps1
```

Die Testläufe nutzen nur PowerShell-Bordmittel. `Test-LegacyScaffold.ps1` prüft den ausdrücklich veralteten Erststand sowie die reine Konfigurationslogik. `Test-DevelopmentPipeline.ps1` prüft zusätzlich die Runtime-Konfiguration, die physische Trennung des Produktiv-Einstiegs von den Simulationsadaptern, das eigenständig kopierbare Production-Bundle, vollständige Bootstrap-Ergebnisse, Modusgrenzen, Execute-Sperren, die vollständige Fake-Adapter-Stagefolge, Commit-vor-State, unveränderte Dateien, Teilerfolg, Inventur-/State-Blockaden, Zielkollisionen, Remote-ID-/Provenienz-/eTag-Sicherheit, Retry mit Fake-Uhr, atomaren State-Replace und die Log-Positivliste.

## Technische Referenzen

- [Microsoft: Unbeaufsichtigte Automation von Office](https://learn.microsoft.com/en-us/office/client-developer/integration/considerations-unattended-automation-office-microsoft-365-for-unattended-rpa)
- [Microsoft Graph: Unterstützte Formatkonvertierungen](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format?view=graph-rest-1.0)
- [Microsoft: Resource Specific Consent und Sites.Selected](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins-modernize/understanding-rsc-for-msgraph-and-sharepoint-online)
- [Microsoft Graph: Kleine Uploads](https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0)
- [Microsoft Graph: Fortsetzbare Uploads](https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0)

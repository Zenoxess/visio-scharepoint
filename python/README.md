# Python-Version des Visio-SharePoint-Spiegels

Die Python-Version trennt Konfiguration, Dateiplanung und externe Zugriffe und erleichtert damit Strukturierung und Tests. Die PowerShell-Version unter `production/` bleibt eine Alternative. Beide Varianten benötigen für den echten Betrieb weiterhin Windows, lokal installiertes und aktiviertes Microsoft Visio sowie SharePoint Server 2019 mit Windows-Authentifizierung. Python ersetzt diese Voraussetzungen nicht.

## Vorbereitung und Start

Benötigt werden Python **3.11 oder neuer**, ein für das Ausführungskonto eingerichtetes Visio-Profil sowie folgende Rechte: Lesen der gesamten Quelle, Schreiben von Logdatei und temporären Dateien und Anlegen, Überschreiben und Recyceln im SharePoint-Ziel. Der dedizierte Zielordner muss bereits existieren; erzwungenes Auschecken in der Bibliothek muss deaktiviert sein (`ForceCheckout=false`). Für Netzwerkquellen empfiehlt sich ein UNC-Pfad statt eines zugeordneten Laufwerks, besonders in geplanten Aufgaben.

Alle folgenden Befehle werden aus der Projektwurzel ausgeführt:

```powershell
py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install -r python\requirements.txt
```

Die Pakete `requests`, `requests-negotiate-sspi` und `pywin32` ermöglichen SharePoint-Zugriffe mit Windows-Anmeldung und Visio-COM. Die sechs Platzhalter in `python/mirror.json` müssen vor dem ersten echten Lauf ersetzt werden:

| Schlüssel | Inhalt |
|---|---|
| `SourcePath` | Absoluter Quellordner, bevorzugt UNC bei Netzwerkquellen |
| `SharePointSiteUrl` | HTTPS-URL der SharePoint-Site |
| `LibraryName` | Anzeigename der Dokumentbibliothek |
| `TargetFolderPath` | Relativer Zielordner innerhalb der Bibliothek |
| `TargetFolderUniqueId` | Tatsächliche SharePoint-`UniqueId` des Zielordners als GUID |
| `LogPath` | Absoluter Pfad der Logdatei |

Pfad und Ziel-GUID müssen zusammenpassen. Die SharePoint-Administration kann die GUID ermitteln. Nach dem Befüllen starten:

```powershell
.venv\Scripts\python.exe python\visio_sharepoint_mirror.py --config python\mirror.json
```

Die Anmeldung verwendet das aktuell ausführende Windows-Konto ohne hinterlegte Kennwörter. Die TLS-Zertifikatsprüfung bleibt aktiv. Bei einer internen Unternehmens-CA muss gegebenenfalls `REQUESTS_CA_BUNDLE` auf eine passende PEM-Zertifikatsdatei zeigen, da Requests standardmäßig sein CA-Bundle statt des Windows-Zertifikatsspeichers verwendet. HTTP-Umleitungen werden abgelehnt; die konfigurierte Site-URL muss direkt erreichbar sein. Siehe [Requests: Zertifikatsprüfung](https://requests.readthedocs.io/en/latest/user/advanced/#ssl-cert-verification).

Die Abhängigkeiten sind in `requirements.txt` auf konkrete Versionen festgelegt. Das SSPI-Paket ist älter; die Windows-Anmeldung muss deshalb mit der tatsächlichen Python-Version und SharePoint-Umgebung im Testlauf bestätigt werden. Die Demo und die Unit-Tests benötigen diese Anmeldung nicht.

## Verhalten der echten Spiegelung

`.vsd`, `.vsdx` und `.vsdm` werden bei jedem Lauf erneut als PDF exportiert und hochgeladen. Beispiel: `Quelle\A\Datei.vsdx` wird zu `<SharePoint-Ziel>/Quelle/A/Datei.pdf`. Andere Quelldateien und leere Ordner werden ignoriert.

**Vollspiegelung:** Überzählige Dateien und Unterordner im Ziel werden nach erfolgreicher Verarbeitung in den SharePoint-Papierkorb verschoben, auch manuell abgelegte Inhalte. Bei leerer Quelle werden sämtliche Zielinhalte recycelt. Der konfigurierte Zielordner selbst bleibt erhalten. Er muss ausschließlich für diesen Spiegel reserviert sein.

Scan-, Konvertierungs-, Authentifizierungs- oder Uploadfehler verhindern die Bereinigung. Eine bereits begonnene Bereinigung lässt sich bei einem späteren Fehler nicht automatisch zurückrollen. Es gibt keinen DryRun-Modus. Exitcode `0` bedeutet Erfolg, `1` einen Fehler; Einzelheiten stehen in Konsole und Logdatei.

Python und PowerShell teilen für identisch konfigurierte Ziele dieselbe Sperre gegen parallele Läufe auf demselben Host. Regelmäßige Starts und eine maximale Laufzeit werden über die Windows-Aufgabenplanung eingerichtet. Die Quellprüfung vergleicht Pfade, Größen und Änderungszeiten; Änderungen mit unveränderter Größe und Zeit können dadurch nicht erkannt werden. Eine PDF wird für den SSPI-Upload vollständig in den Arbeitsspeicher gelesen.

## Konsolen-Demo und lokale Tests

Die Demo benötigt nur Python, keine Paketinstallation oder Konfiguration. Sie verwendet feste Beispieldaten und ausschließlich Konsolenausgaben; keine Dateien werden verarbeitet und keine Verbindungen aufgebaut. Jede Ablaufzeile trägt `[DEMO]`. Sie ist unabhängig vom Betriebsskript und kein Funktionstest oder DryRun gegen ein echtes Ziel.

```powershell
py -3 python\demo_mirror.py
py -3 python\demo_mirror.py --scenario upload-error
py -3 python\demo_mirror.py --scenario empty-source
```

Der Standardfall zeigt einen erfolgreichen Lauf (Exitcode `0`). `upload-error` zeigt den Abbruch beim zweiten Upload ohne Bereinigung (Exitcode `1`); `empty-source` zeigt die Bereinigung sämtlicher Zielinhalte (Exitcode `0`). Den Exitcode direkt danach in PowerShell mit `$LASTEXITCODE`, in der Eingabeaufforderung mit `echo %ERRORLEVEL%` anzeigen.

Die automatisierten lokalen Tests benötigen keine Fremdbibliotheken:

```powershell
py -3 -m unittest discover -s tests -p test_python_mirror.py -v
```

Ein echter Integrationslauf mit Visio und SharePoint wurde noch nicht durchgeführt. Vor dem produktiven Einsatz ist ein vollständiger Probelauf mit mehrseitigen Zeichnungen, Überschreiben und Bereinigung in einem separaten Testziel erforderlich.

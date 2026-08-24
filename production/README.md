# PPSI Visio-PDF-Spiegel

Dieser Ordner ist das vollständige Betriebspaket. Benötigt werden genau diese drei Dateien:

- `Invoke-VisioSharePointMirror.ps1`
- `mirror.json`
- `README.md`

Das Skript besitzt nur einen Betriebsablauf: Es spiegelt alle Visio-Zeichnungen aus einem Quellordner als PDF in einen dedizierten SharePoint-Server-2019-Zielordner.

## Voraussetzungen

- Windows PowerShell 5.1 oder PowerShell 7
- lokal installiertes und für das Ausführungskonto initialisiertes Microsoft Visio
- Windows-Authentifizierung an der verwendeten SharePoint-Site
- Lesezugriff auf den Quellordner sowie Anlage-, Überschreib- und Papierkorbrechte im SharePoint-Ziel
- eine Dokumentbibliothek ohne erzwungenes Auschecken (`ForceCheckout=false`)
- ein bereits vorhandener, ausschließlich für diesen Spiegel reservierter Zielordner

Bei einer geplanten Aufgabe müssen Visio-Profil, Netzlaufwerk und SharePoint-Rechte für genau das Aufgabenkonto eingerichtet sein. Eine Laufwerkszuordnung aus einer anderen Benutzersitzung ist dort normalerweise nicht sichtbar.

## Konfiguration

`mirror.json` enthält genau sechs Werte:

| Feld | Bedeutung |
|---|---|
| `SourcePath` | Absoluter Quellordner, beispielsweise `P:\Quelle` |
| `SharePointSiteUrl` | HTTPS-URL der SharePoint-Site |
| `LibraryName` | Anzeigename der Dokumentbibliothek |
| `TargetFolderPath` | Relativer, bereits vorhandener Zielordner innerhalb der Bibliothek |
| `TargetFolderUniqueId` | SharePoint-`UniqueId` dieses Zielordners als GUID |
| `LogPath` | Absoluter Pfad der fortlaufenden Logdatei |

Die Ziel-ID kann von der SharePoint-Administration per REST oder über die SharePoint-Verwaltungswerkzeuge ermittelt werden. Pfad und ID müssen zusammenpassen; andernfalls beendet sich das Skript ohne SharePoint-Änderung.

Die mitgelieferte Konfiguration enthält absichtlich eine `.invalid`-URL und Beispielwerte. Sie muss vor dem ersten Lauf angepasst werden.

## Aufruf

```powershell
pwsh -NoProfile -File .\Invoke-VisioSharePointMirror.ps1 `
  -ConfigurationPath .\mirror.json
```

Alternativ mit Windows PowerShell 5.1:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\Invoke-VisioSharePointMirror.ps1 `
  -ConfigurationPath .\mirror.json
```

Das Skript startet sich bei Bedarf automatisch in einem STA-Prozess neu, den Visio-COM benötigt. Es gibt keinen Preview-, Validate-, Simulate- oder DryRun-Modus.

## Abbildung und Spiegelverhalten

```text
P:\Quelle\A\B\Datei.vsdx
    -> <SharePoint-Ziel>\Quelle\A\B\Datei.pdf
```

- Verarbeitet werden `.vsd`, `.vsdx` und `.vsdm`; andere Dateien und leere Ordner werden ignoriert.
- Jede Visio-Datei wird bei jedem Lauf neu konvertiert und mit `overwrite=true` hochgeladen.
- Erst nach vollständigem Scan, erfolgreicher Konvertierung aller Dateien, erfolgreichem Upload und erneuter Zielprüfung werden überzählige Inhalte recycelt.
- Ist die Quelle leer, werden alle Inhalte unterhalb des Zielordners in den SharePoint-Papierkorb verschoben. Der Zielordner selbst bleibt bestehen.
- Manuell abgelegte Fremdinhalte unterhalb des dedizierten Zielordners werden ebenfalls recycelt.

Nicht eindeutig adressierbare SharePoint-Namen, Pfade über 260 Zeichen, Reparse Points und PDF-Zielkollisionen brechen den Lauf vor Remote-Änderungen ab. Blockiert im Ziel eine Datei einen benötigten Ordner oder ein Ordner eine erwartete PDF, muss dieser Typkonflikt manuell beseitigt werden.

## Ergebnis und Betrieb

- Exitcode `0`: Spiegel-Lauf vollständig erfolgreich
- Exitcode `1`: Lauf fehlgeschlagen; die konkrete Ursache steht in Konsole und Log

Ein Scan-, Konvertierungs-, Authentifizierungs- oder Uploadfehler verhindert die Papierkorbphase. Schlägt eine Papierkorbaktion selbst fehl, können vorherige Papierkorbaktionen nicht zurückgerollt werden; der Lauf endet dann mit Exitcode `1`.

Die Zeitbegrenzung für eine geplante Aufgabe wird in der Windows-Aufgabenplanung festgelegt. Das Skript enthält bewusst keinen Scheduler, keinen lokalen Synchronisations-State und keine internen Testmodi.

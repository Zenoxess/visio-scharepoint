#requires -Version 5.1
<#
Reine Konsolen-Demo mit festen Beispieldaten. Kein Visio, SharePoint,
Dateizugriff, Netzwerkzugriff oder Import des Betriebsskripts.
#>
[CmdletBinding()]
param(
    [ValidateSet('Success', 'UploadError', 'EmptySource')]
    [string]$Scenario = 'Success'
)

function Write-Demo {
    param([string]$Message, [string]$Level = 'INFO')
    Write-Host ("[DEMO] [{0}] {1}" -f $Level, $Message)
}

$files = @(
    @{ Source = 'Prozesslandkarte.vsdx'; Pdf = 'Prozesslandkarte.pdf' }
    @{ Source = 'Vertrieb\Angebot.vsdx'; Pdf = 'Vertrieb/Angebot.pdf' }
    @{ Source = 'IT\Systeme.vsd'; Pdf = 'IT/Systeme.pdf' }
)
if ($Scenario -eq 'EmptySource') { $files = @() }

Write-Demo 'NUR ANZEIGE: Alle Daten und Aktionen sind erfunden.'
Write-Demo ("Spiegel-Lauf gestartet. Szenario: {0}" -f $Scenario)
Write-Demo 'Quelle: C:\Demo\Quelle (simuliert)'
Write-Demo 'Ziel: /sites/PPSI/Dokumente/Visio/Quelle (simuliert)'
Write-Demo ("{0} Visio-Datei(en) gefunden." -f $files.Count)
Write-Demo 'SharePoint-Ziel und Windows-Anmeldung bestaetigt (simuliert).'

foreach ($file in $files) {
    Write-Demo ("Konvertiert: {0} -> {1} (simuliert)" -f $file.Source, $file.Pdf)
}
Write-Demo ("{0} PDF-Datei(en) erzeugt (simuliert)." -f $files.Count)

$uploaded = 0
foreach ($file in $files) {
    if ($Scenario -eq 'UploadError' -and $uploaded -eq 1) {
        Write-Demo ("Upload fehlgeschlagen: {0} - HTTP 503 (simuliert)." -f $file.Pdf) 'ERROR'
        Write-Demo 'Lauf abgebrochen. Bereinigung wird wegen des Fehlers nicht ausgefuehrt.' 'ERROR'
        Write-Demo 'Demo beendet. Exitcode: 1'
        exit 1
    }
    Write-Demo ("Hochgeladen: {0} (simuliert)" -f $file.Pdf)
    $uploaded++
}
Write-Demo ("{0} PDF-Datei(en) nach SharePoint hochgeladen (simuliert)." -f $uploaded)

if ($Scenario -eq 'EmptySource') {
    Write-Demo 'Quelle ist leer: Alle Inhalte unterhalb des Zielordners wuerden recycelt.' 'WARN'
    Write-Demo 'Spiegelabgleich abgeschlossen: 4 Datei(en) und 3 Ordner recycelt (simuliert).'
} else {
    Write-Demo 'Bereinigung: Veraltete Datei Alt/Prozess.pdf und Ordner Alt recycelt (simuliert).'
    Write-Demo 'Spiegelabgleich abgeschlossen: 1 Datei(en) und 1 Ordner recycelt (simuliert).'
}
Write-Demo 'Der konfigurierte Zielordner bleibt bestehen (simuliert).'
Write-Demo 'Demo beendet. Exitcode: 0'
exit 0

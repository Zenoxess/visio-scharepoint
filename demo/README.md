# Konsolen-Demo

`Show-VisioSharePointMirror.ps1` zeigt den Ablauf mit festen Beispieldaten. Jede Ausgabe ist mit `[DEMO]` gekennzeichnet. Das Skript liest keine Konfiguration oder Quelldateien, erzeugt keine PDFs oder Logdateien und verbindet sich weder mit Visio noch mit SharePoint. Alle dargestellten Aktionen sind erfunden. Windows PowerShell 5.1 genügt.

Die Demo ist unabhängig vom Betriebsskript. Sie zeigt dessen typische Phasen und das Verhalten bei Fehlern; sie ist kein Funktionstest der echten Spiegelung und kein DryRun für ein echtes Ziel. Ausgabedetails und Dateinamen dienen ausschließlich der Veranschaulichung.

Die folgenden Befehle aus der Projektwurzel in PowerShell oder der Windows-Eingabeaufforderung ausführen. Sie starten einen eigenen PowerShell-Prozess, damit der Exitcode die aufrufende Konsole nicht beendet.

Erfolgreicher Lauf mit drei Visio-Dateien und anschließender simulierter Bereinigung (Exitcode `0`):

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\demo\Show-VisioSharePointMirror.ps1
```

Fehler beim zweiten Upload, Abbruch ohne Bereinigung (Exitcode `1`):

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\demo\Show-VisioSharePointMirror.ps1 -Scenario UploadError
```

Leere Quelle mit simulierter Bereinigung aller Zielinhalte; der Zielordner bleibt erhalten (Exitcode `0`):

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\demo\Show-VisioSharePointMirror.ps1 -Scenario EmptySource
```

Den tatsächlichen Exitcode direkt nach dem jeweiligen Aufruf mit `$LASTEXITCODE` in PowerShell oder `echo %ERRORLEVEL%` in der Eingabeaufforderung anzeigen.

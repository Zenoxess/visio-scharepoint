"""Reine Konsolen-Demo: feste Beispieldaten, keinerlei externe Aktionen."""

import argparse


def show(message: str, level: str = "INFO") -> None:
    print(f"[DEMO] [{level}] {message}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--scenario",
        choices=("success", "upload-error", "empty-source"),
        default="success",
        help="Angezeigter Ablauf (Standard: success).",
    )
    scenario = parser.parse_args().scenario
    files = [
        ("Prozesslandkarte.vsdx", "Prozesslandkarte.pdf"),
        (r"Vertrieb\Angebot.vsdx", "Vertrieb/Angebot.pdf"),
        (r"IT\Systeme.vsd", "IT/Systeme.pdf"),
    ]
    if scenario == "empty-source":
        files = []

    show("NUR ANZEIGE: Alle Daten und Aktionen sind erfunden.")
    show(f"Spiegel-Lauf gestartet. Szenario: {scenario}")
    show(r"Quelle: C:\Demo\Quelle (simuliert)")
    show("Ziel: /sites/PPSI/Dokumente/Visio/Quelle (simuliert)")
    show(f"{len(files)} Visio-Datei(en) gefunden.")
    show("SharePoint-Ziel und Windows-Anmeldung bestaetigt (simuliert).")

    for source, pdf in files:
        show(f"Konvertiert: {source} -> {pdf} (simuliert)")
    show(f"{len(files)} PDF-Datei(en) erzeugt (simuliert).")

    for index, (_, pdf) in enumerate(files):
        if scenario == "upload-error" and index == 1:
            show(f"Upload fehlgeschlagen: {pdf} - HTTP 503 (simuliert).", "ERROR")
            show("Lauf abgebrochen. Bereinigung wird wegen des Fehlers nicht ausgefuehrt.", "ERROR")
            show("Demo beendet. Exitcode: 1")
            return 1
        show(f"Hochgeladen: {pdf} (simuliert)")
    show(f"{len(files)} PDF-Datei(en) nach SharePoint hochgeladen (simuliert).")

    if scenario == "empty-source":
        show("Quelle ist leer: Alle Inhalte unterhalb des Zielordners wuerden recycelt.", "WARN")
        show("Spiegelabgleich abgeschlossen: 4 Datei(en) und 3 Ordner recycelt (simuliert).")
    else:
        show("Bereinigung: Veraltete Datei Alt/Prozess.pdf und Ordner Alt recycelt (simuliert).")
        show("Spiegelabgleich abgeschlossen: 1 Datei(en) und 1 Ordner recycelt (simuliert).")
    show("Der konfigurierte Zielordner bleibt bestehen (simuliert).")
    show("Demo beendet. Exitcode: 0")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

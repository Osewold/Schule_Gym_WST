#!/usr/bin/env python3
"""
Excel-Auswertungstool für das Gymnasium WST
============================================
Dieses Skript liest eine Excel-Datei ein und erstellt eine
umfassende Auswertung der enthaltenen Daten.

Verwendung:
    python3 excel_auswertung.py <excel_datei.xlsx>

Optionale Argumente:
    --sheet <Name>    Nur ein bestimmtes Tabellenblatt auswerten
    --output <Datei>  Ergebnis in eine Textdatei schreiben
    --csv             Zusammenfassung als CSV exportieren
"""

import sys
import os
import argparse
from pathlib import Path


# Unterstützte Dateiendungen
SUPPORTED_EXTENSIONS = (".xlsx", ".xls", ".xlsm", ".xlsb", ".ods")

# Datentypen, die als Textspalten behandelt werden
TEXT_DTYPES = ["object", "string", "category"]

# Erkennungsbegriffe für schulspezifische Daten (Kleinbuchstaben)
STUNDENPLAN_BEGRIFFE = frozenset({
    "montag", "dienstag", "mittwoch", "donnerstag", "freitag",
    "mo", "di", "mi", "do", "fr", "stunde", "periode",
})
LEHRER_BEGRIFFE = frozenset({
    "lehrer", "lehrkraft", "kuk", "lehrerin", "name", "kürzel",
    "deputat", "stunden", "fach", "fächer",
})
SCHUELER_BEGRIFFE = frozenset({
    "schüler", "klasse", "jahrgang", "jg", "geburtsdatum",
    "vorname", "nachname", "geburtstag",
})
DEPUTAT_BEGRIFFE = frozenset({"stunden", "deputat", "std"})
KLASSEN_BEGRIFFE = frozenset({"klasse", "jg", "jahrgang"})


def check_dependencies():
    """Überprüft, ob alle benötigten Bibliotheken installiert sind."""
    missing = []
    try:
        import pandas  # noqa: F401
    except ImportError:
        missing.append("pandas")
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        missing.append("openpyxl")

    if missing:
        print("Fehlende Bibliotheken. Bitte installieren mit:")
        print(f"  pip install {' '.join(missing)}")
        sys.exit(1)


check_dependencies()

import pandas as pd  # noqa: E402


def lade_excel(dateipfad: str, sheet_name=None):
    """
    Lädt eine Excel-Datei und gibt ein Dictionary mit DataFrames zurück.

    Args:
        dateipfad: Pfad zur Excel-Datei
        sheet_name: Name des Tabellenblatts (None = alle Blätter)

    Returns:
        dict mit Blattnamen als Keys und DataFrames als Values
    """
    pfad = Path(dateipfad)
    if not pfad.exists():
        print(f"Fehler: Datei '{dateipfad}' nicht gefunden.")
        sys.exit(1)

    if pfad.suffix.lower() not in SUPPORTED_EXTENSIONS:
        print(f"Warnung: Unbekannte Dateiendung '{pfad.suffix}'. Versuche dennoch zu lesen...")

    try:
        if sheet_name:
            daten = {sheet_name: pd.read_excel(dateipfad, sheet_name=sheet_name)}
        else:
            daten = pd.read_excel(dateipfad, sheet_name=None)
        return daten
    except Exception as e:
        print(f"Fehler beim Lesen der Datei: {e}")
        sys.exit(1)


def blatt_info(name: str, df: pd.DataFrame) -> str:
    """Gibt grundlegende Informationen über ein Tabellenblatt zurück."""
    zeilen = []
    zeilen.append(f"\n{'='*60}")
    zeilen.append(f"Tabellenblatt: '{name}'")
    zeilen.append(f"{'='*60}")
    zeilen.append(f"Zeilen: {len(df)}")
    zeilen.append(f"Spalten: {len(df.columns)}")
    zeilen.append(f"Spaltennamen: {list(df.columns)}")

    # Leere Zellen
    leere = df.isnull().sum().sum()
    gesamt = df.size
    if gesamt > 0:
        zeilen.append(f"Leere Zellen: {leere} von {gesamt} ({100*leere/gesamt:.1f}%)")

    return "\n".join(zeilen)


def numerische_statistik(df: pd.DataFrame) -> str:
    """Erstellt Statistiken für numerische Spalten."""
    num_df = df.select_dtypes(include="number")
    if num_df.empty:
        return "\n[Keine numerischen Spalten gefunden]"

    zeilen = ["\n--- Numerische Auswertung ---"]
    stats = num_df.describe().round(2)
    zeilen.append(stats.to_string())
    return "\n".join(zeilen)


def text_auswertung(df: pd.DataFrame) -> str:
    """Erstellt Häufigkeitsanalysen für Textspalten."""
    text_df = df.select_dtypes(include=TEXT_DTYPES)
    if text_df.empty:
        return ""

    zeilen = ["\n--- Textspalten: Häufigste Werte ---"]
    for spalte in text_df.columns:
        nicht_leer = df[spalte].dropna()
        if len(nicht_leer) == 0:
            continue
        eindeutig = nicht_leer.nunique()
        zeilen.append(f"\n  Spalte '{spalte}': {eindeutig} eindeutige Werte")
        top = nicht_leer.value_counts().head(10)
        for wert, anzahl in top.items():
            prozent = 100 * anzahl / len(nicht_leer)
            zeilen.append(f"    {str(wert)[:40]:<40} {anzahl:>5}x  ({prozent:5.1f}%)")

    return "\n".join(zeilen)


def duplikate_pruefen(df: pd.DataFrame) -> str:
    """Überprüft auf doppelte Zeilen."""
    duplikate = df.duplicated().sum()
    if duplikate == 0:
        return "\n[Keine doppelten Zeilen gefunden]"
    return f"\nHinweis: {duplikate} doppelte Zeilen gefunden!"


def daten_vorschau(df: pd.DataFrame, n: int = 5) -> str:
    """Zeigt die ersten n Zeilen des DataFrames."""
    zeilen = [f"\n--- Erste {min(n, len(df))} Zeilen ---"]
    zeilen.append(df.head(n).to_string(index=False))
    return "\n".join(zeilen)


def schulbezogene_auswertung(name: str, df: pd.DataFrame) -> str:
    """
    Versucht schulspezifische Auswertungen zu erkennen und durchzuführen.
    Erkennt u.a. Stundenpläne, Lehrerlisten und Klassenlisten.
    """
    zeilen = []
    spalten_lower = [str(s).lower() for s in df.columns]

    # Stundenplan-Erkennung
    if any(b in spalten_lower for b in STUNDENPLAN_BEGRIFFE):
        zeilen.append("\n--- Stundenplan erkannt ---")
        zeilen.append(f"  Spalten: {list(df.columns)}")

    # Lehrer-Erkennung
    if any(b in spalten_lower for b in LEHRER_BEGRIFFE):
        zeilen.append("\n--- Lehrerdaten erkannt ---")
        # Suche nach Stunden-/Deputat-Spalten
        stunden_spalten = [s for s, sl in zip(df.columns, spalten_lower)
                           if any(b in sl for b in DEPUTAT_BEGRIFFE)]
        for sp in stunden_spalten:
            werte = pd.to_numeric(df[sp], errors="coerce").dropna()
            if len(werte) > 0:
                zeilen.append(f"  Summe '{sp}': {werte.sum():.1f}")
                zeilen.append(f"  Durchschnitt '{sp}': {werte.mean():.1f}")

    # Schüler-Erkennung
    if any(b in spalten_lower for b in SCHUELER_BEGRIFFE):
        zeilen.append("\n--- Schülerdaten erkannt ---")
        klassen_spalten = [s for s, sl in zip(df.columns, spalten_lower)
                           if any(b in sl for b in KLASSEN_BEGRIFFE)]
        for sp in klassen_spalten:
            nicht_leer = df[sp].dropna()
            if len(nicht_leer) > 0:
                zeilen.append(f"  Klassen/Jahrgänge in '{sp}':")
                for k, n in nicht_leer.value_counts().head(20).items():
                    zeilen.append(f"    {k}: {n} Schüler")

    return "\n".join(zeilen)


def exportiere_csv(daten: dict, ausgabe_pfad: str):
    """Exportiert die Auswertungsdaten als CSV-Dateien."""
    basis = Path(ausgabe_pfad).stem
    verzeichnis = Path(ausgabe_pfad).parent
    exportiert = []

    for name, df in daten.items():
        # Sicheren Dateinamen erstellen
        sicherer_name = "".join(c if c.isalnum() or c in "-_" else "_" for c in str(name))
        csv_pfad = verzeichnis / f"{basis}_{sicherer_name}.csv"
        df.to_csv(csv_pfad, index=False, encoding="utf-8-sig")
        exportiert.append(str(csv_pfad))

    return exportiert


def auswertung_durchfuehren(dateipfad: str, sheet_name=None, ausgabe=None, csv_export=False):
    """
    Hauptfunktion: Führt die vollständige Auswertung durch.

    Args:
        dateipfad: Pfad zur Excel-Datei
        sheet_name: Optionaler Name eines bestimmten Tabellenblatts
        ausgabe: Optionaler Pfad für Textausgabe
        csv_export: Falls True, werden Daten als CSV exportiert
    """
    print(f"\nLade Excel-Datei: {dateipfad}")
    daten = lade_excel(dateipfad, sheet_name)

    ergebnis_zeilen = []
    ergebnis_zeilen.append(f"Excel-Auswertung: {Path(dateipfad).name}")
    ergebnis_zeilen.append(f"Anzahl Tabellenblätter: {len(daten)}")
    ergebnis_zeilen.append(f"Blätter: {list(daten.keys())}")

    for name, df in daten.items():
        ergebnis_zeilen.append(blatt_info(name, df))
        ergebnis_zeilen.append(daten_vorschau(df))
        ergebnis_zeilen.append(numerische_statistik(df))
        ergebnis_zeilen.append(text_auswertung(df))
        ergebnis_zeilen.append(duplikate_pruefen(df))
        ergebnis_zeilen.append(schulbezogene_auswertung(name, df))

    ergebnis = "\n".join(ergebnis_zeilen)
    print(ergebnis)

    if ausgabe:
        with open(ausgabe, "w", encoding="utf-8") as f:
            f.write(ergebnis)
        print(f"\nAuswertung gespeichert in: {ausgabe}")

    if csv_export:
        # Ausgabebasis aus dem Eingabepfad ableiten, unabhängig von der Dateiendung
        basis = ausgabe if ausgabe else str(Path(dateipfad).with_suffix("")) + "_auswertung.txt"
        exportierte = exportiere_csv(daten, basis)
        print(f"\nCSV-Dateien exportiert:")
        for pfad in exportierte:
            print(f"  {pfad}")

    return daten


def main():
    ext_liste = ", ".join(SUPPORTED_EXTENSIONS)
    parser = argparse.ArgumentParser(
        description="Excel-Auswertungstool für das Gymnasium WST",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("datei", help=f"Pfad zur Excel-Datei ({ext_liste})")
    parser.add_argument(
        "--sheet", "-s",
        help="Name eines bestimmten Tabellenblatts (Standard: alle Blätter)",
        default=None,
    )
    parser.add_argument(
        "--output", "-o",
        help="Ergebnis zusätzlich in Textdatei speichern",
        default=None,
    )
    parser.add_argument(
        "--csv",
        action="store_true",
        help="Tabellenblätter als CSV-Dateien exportieren",
    )

    args = parser.parse_args()
    auswertung_durchfuehren(args.datei, args.sheet, args.output, args.csv)


if __name__ == "__main__":
    main()

"""
analyze_excel.py – Wertet eine Excel-Tabelle aus und gibt eine Zusammenfassung aus.

Verwendung:
    python analyze_excel.py <excel_datei> [<tabellenblatt>]

Argumente:
    excel_datei     Pfad zur Excel-Datei (.xlsx oder .xls)
    tabellenblatt   Optionaler Name des Tabellenblatts (Standard: erstes Blatt)
"""

import sys
import pandas as pd


def analyze(filepath: str, sheet_name: str | None = None) -> None:
    """Lädt eine Excel-Datei und gibt eine strukturierte Auswertung aus."""
    try:
        xl = pd.ExcelFile(filepath)
    except FileNotFoundError:
        print(f"Fehler: Datei '{filepath}' nicht gefunden.")
        sys.exit(1)
    except Exception as exc:
        print(f"Fehler beim Öffnen der Datei '{filepath}': {exc}")
        sys.exit(1)
    sheet_names = xl.sheet_names

    print(f"Datei: {filepath}")
    print(f"Tabellenblätter ({len(sheet_names)}): {', '.join(sheet_names)}")
    print()

    target_sheet = sheet_name if sheet_name else sheet_names[0]
    if target_sheet not in sheet_names:
        print(f"Fehler: Tabellenblatt '{target_sheet}' nicht gefunden.")
        sys.exit(1)

    df = xl.parse(target_sheet)

    print(f"=== Tabellenblatt: {target_sheet} ===")
    print(f"Zeilen: {len(df)}  |  Spalten: {len(df.columns)}")
    print()

    print("--- Spaltenübersicht ---")
    for col in df.columns:
        non_empty = df[col].notna().sum()
        print(f"  {col!r:40s}  Typ: {df[col].dtype}  |  Befüllt: {non_empty}/{len(df)}")
    print()

    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    if numeric_cols:
        print("--- Statistische Zusammenfassung (numerische Spalten) ---")
        print(df[numeric_cols].describe().to_string())
        print()

    object_cols = df.select_dtypes(include=["object", "str"]).columns.tolist()
    for col in object_cols:
        unique_vals = df[col].dropna().unique()
        if 1 < len(unique_vals) <= 20:
            print(f"--- Eindeutige Werte in '{col}' ---")
            for val in sorted(unique_vals, key=str):
                count = (df[col] == val).sum()
                print(f"  {str(val):40s}  {count}x")
            print()

    missing = df.isnull().sum()
    missing = missing[missing > 0]
    if not missing.empty:
        print("--- Fehlende Werte ---")
        for col, count in missing.items():
            print(f"  {col!r:40s}  {count} fehlend")
        print()

    print("--- Erste 5 Zeilen ---")
    print(df.head().to_string())
    print()


def main() -> None:
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    filepath = sys.argv[1]
    sheet_name = sys.argv[2] if len(sys.argv) >= 3 else None
    analyze(filepath, sheet_name)


if __name__ == "__main__":
    main()

# Schule_Gym_WST

Verwaltungs-Repository für das Gymnasium WST.

## Excel-Auswertung

Das Skript `excel_auswertung.py` ermöglicht die Auswertung von Excel-Tabellen (z.B. Stundenpläne, Lehrerlisten, Schülerdaten).

### Voraussetzungen

```bash
pip install -r requirements.txt
```

### Verwendung

```bash
# Alle Tabellenblätter auswerten
python3 excel_auswertung.py meine_tabelle.xlsx

# Nur ein bestimmtes Tabellenblatt auswerten
python3 excel_auswertung.py meine_tabelle.xlsx --sheet Lehrerliste

# Ergebnis in Textdatei speichern
python3 excel_auswertung.py meine_tabelle.xlsx --output auswertung.txt

# Daten als CSV exportieren
python3 excel_auswertung.py meine_tabelle.xlsx --csv
```

### Funktionen

- **Übersicht**: Anzahl Zeilen, Spalten, leere Zellen pro Tabellenblatt
- **Datenvorschau**: Erste Zeilen der Tabelle
- **Numerische Statistik**: Min, Max, Mittelwert, Standardabweichung
- **Häufigkeitsanalyse**: Häufigste Werte in Textspalten
- **Duplikaterkennung**: Hinweis auf doppelte Zeilen
- **Schulspezifische Auswertung**: Automatische Erkennung von Stundenplänen, Lehrerlisten und Schülerdaten

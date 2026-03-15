# Schule_Gym_WST

Dieses Repository dient der Verwaltung und Planung schulischer Aufgaben am Gymnasium WST.

## Excel-Tabellen auswerten

Das Skript `analyze_excel.py` ermöglicht es, eine umfangreiche Excel-Datei schnell auszuwerten.

### Voraussetzungen

```bash
pip install -r requirements.txt
```

### Verwendung

```bash
python analyze_excel.py <excel_datei> [<tabellenblatt>]
```

**Beispiel:**

```bash
python analyze_excel.py stundenplan.xlsx
python analyze_excel.py stundenplan.xlsx "Klasse 10a"
```

Das Skript gibt aus:

- Alle enthaltenen Tabellenblätter
- Spaltenübersicht mit Datentypen und Befüllungsgrad
- Statistische Zusammenfassung numerischer Spalten
- Häufigkeitsverteilung kategorialer Spalten (bis 20 eindeutige Werte)
- Übersicht fehlender Werte
- Vorschau der ersten 5 Zeilen
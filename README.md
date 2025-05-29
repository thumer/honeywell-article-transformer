# Honeywell Article Transformer

Dieses Konsolenprogramm verarbeitet Artikelbeschreibungen aus einer CSV- oder XLSX-Datei mit Semikolon als Trennzeichen. Die Anwendung erwartet den Pfad zur Eingabedatei als ersten Parameter und erzeugt eine Ausgabedatei im gleichen Format.

## Verwendung

```bash
Honeywell.ArticleTransformer.exe eingabe.csv
```

oder

```bash
Honeywell.ArticleTransformer.exe eingabe.xlsx
```

Bei CSV wird `output.csv` erzeugt, bei XLSX `output.xlsx`. Alle Zeilenumbrüche im Ergebnis nutzen das Windows-Format (UTF-8).

## Transformationsschritte

1. Alle ursprünglichen Zeilenumbrüche (`_x000D_`, `\n`, `\r`) entfernen.
2. Mehrfache Leerzeichen auf ein Leerzeichen reduzieren.
3. Vor den Begriffen "Anschlüsse:", "Technische Daten:", "Daten:" und "Betriebsspannung:" einen Zeilenumbruch einfügen. Steht "Daten:" alleine, auch danach einen Zeilenumbruch. "Technische Daten" darf nicht getrennt werden. Nach jedem Doppelpunkt einen Zeilenumbruch einfügen. Beginnt danach eine Aufzählung, jede Zeile mit `-` beginnen und zuvor einen Zeilenumbruch einfügen.
4. Technische Daten logisch und übersichtlich strukturieren. Keine Formatierungen einsetzen, leere Felder bleiben leer und vorhandene Schreibweise wird nicht korrigiert.

# VBA_Excel_BI

## Folgende Aufgaben waren Teil des Bachlor Moduls "Business intelligence"

Die erste Aufgabe (Aufgabenblatt 5) war es "Plausibilitätsprüfungen" auf einen Excel-Datenbestand anzuwenden und Fehler zu erkennen.
Die Zweite Aufgabe (Aufgabenblatt 6) war es dann aus diesen korrigierten Daten eine dem Unterrichtsstoff nach besprochenen Art "Star Schema" ein kleines Data Warehouse in Excel aufzubauen. In den folgenden Aufgaben wurde dann dieses DW zuerst in eine MS-Acces Datenbank integriert, anschließend in das DW-Tool SAP BW on HANA rein geladen.

## Praktikum -Aufgabenblatt 5
Aufgabe 1 (6 Punkte, Abgabe 5. bzw. 6. Juni 2019):
Sie arbeiten als Junior-Consultant in einem kleinen Bera-tungsunternehmen. Sie sollen ein Data Warehouse für einen Autohändler konzipieren. Ihnen liegen die Rohdaten (rohda-ten.xlsx) für den An- und Verkauf der Autos vor. Im ersten Schritt sollen Sie die Daten harmonisieren und die Mängel beseitigen. (Diese Aufgabe wird fortgeführt.)

a) Markieren Sie manuell die Mängel in den beiden Tabel-len. (1 Punkt)

b) Entwickeln Sie ein VBA-Programm, das leere Datenfelder und potenzielle Ausreißer in den Rohdaten erkennt. Die-se Mängel sollen in einem Protokoll (ID, (potenziell) fehlerhaftes Datenfeld) ausgegeben werden. (3 Punkte)

c) Führen Sie eine Harmonisierung und ggfs. eine Fehlerbe-hebung für Datum und ID durch. Als Ergebnis soll eine Tabelle mit den Datenfeldern ID, Tag Ankauf, Monat An-kauf, Jahr Ankauf, Tag Verkauf, Monat Verkauf, Jahr Verkauf entstehen. Wenn der Verkauf noch nicht erfolgt ist, dann soll der 31.12.9999 als Verkaufsdatum ange-nommen werden. (2 Punkte)


## Praktikum -Aufgabenblatt 6
Aufgabe 1 (6 Punkte, Abgabe 5. bzw. 6. Juni 2019):

Fortführung Data Warehouse für Autohaus Schäfer
Entwickeln Sie ein „kleines“ Data Warehouse auf Basis eines Star-Schemas. Erstellen Sie hierzu eine Faktentabelle mit-tels VBA-Programm. Die Faktentabelle soll alle Auswertungen zum Lagerbestand abdecken, die für das „Berichtswesen Auto-haus Schäfer“ erforderlich sind.
Achten Sie bei der Erstellung des VBA-Programmes darauf, dass dieses nicht geändert werden muss, wenn zukünftig neue Ausprägungen der Merkmale (z.B. eine neue Farbe)in den Quelldaten vorkommen.

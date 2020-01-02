# Änderungen
Die neue Vorlage finden Sie in `corrector.xlsx`, in bereits existierenden corrector-Dateien müssen 
folgende Änderungen vorgenommen werden:

- Der Dateiname für eingeschickte Dateien muss in B2 gesetzt werden, grundlegend ist der Name 
hier egal, sollte aber einem Muster von `ModulNummer_laufendeNummer` folgen.
- Die Abgabefrist wird nun als normales Datum in B3 eingefügt
- B4 enthält nun die Anzahl an Versuchen
- "Dummies" bzw. a1-a100 ist nun in B9-I9 und auf 8 Werte begrenzt, ist aber auf Wunsch erweiterbar
- Die Matrikelnummer ist nun in B10
- Die Werte/Aufgaben sind nun ab Reihe 13
- Spalte E enthält nun eine absolute Wertetoleranz als Alternative zur relativen Toleranz in D

Dateien namens `corrector.xls` werden automatisch zu `.xlsx` konvertiert.

# Bedeutung der neuen/geänderten Felder
##### Titel
Wird in der Grußformel der Mails an Studenten genutzt

##### Dateiname
Von Studenten eingesandte Dateien werden anhand des Namens (ohne .xlsx!) dem corrector zugeordnet, 
dabei müssen diese Dateien auf .xlsx enden. Im Feld des correctors sollte lediglich der Dateiname 
selbst ohne Dateiendung (`203131_1` statt `203131_1.xlsx`) eingetragen werden.

##### Variable/Wert in A8:H9
Hier können Werte eingetragen werden, die neben der Matrikelnummer in die Berechnung der Studenten-
spezifischen Aufgabe/Lösung einfließen sollen.

#### Toleranz
Neben der relativen Toleranz (Spalte D) kann nun auch eine absolute Toleranz (Spalte E) festgelegt 
werden. Dabei gelten folgende Regeln:

- Ist **keine Toleranz** festgelegt, wird kein Fehler toleriert.
- Ist **eine Toleranz** eingetragen, wird nur diese beachtet.
- Sind in **beiden** Spalten Werte vorhanden, gilt die (Teil-)Aufgabe als gelöst, wenn der Wert 
innerhalb einem der Toleranzbereichen liegt.

#### Aufgaben
Die Nummerierung der Aufgaben muss numerisch sein und darf keine Buchstaben enthalten, ebenso 
müssen alle Aufgaben durchlaufend nummeriert sein. Soll eine Aufgabe von PyCor nicht korrigiert 
werden, muss sie trotzdem eingetragen werden.
(Teil-)Aufgaben werden nicht als Fehler angerechnet, sofern keine Lösung im Corrector eingetragen ist.

# Fehlermeldungen/Ignorieren von corrector-Dateien
Correctors werden von PyCor ignoriert, wenn eine Datei namens `PYCOR_IGNORE.txt` im selben 
Ordner vorliegt. Diese Datei wird automatisch beim Feststellen von Fehlern angelegt.

Wird ein Fehler erkannt, erstellt PyCor eine Datei namens `PYCOR_ERROR.txt` im Ordner des correctors, die eine 
Fehlermeldung enthält.
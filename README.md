![Addendum.png](assets/AddendumLogo2020.png)

#### Extends the possibilities through surface analysis and manipulation

### Version 1.34 vom 08.08.2020
lauffähig ab *Autohotkey_H V1.1.32.00 Unicode 64bit* , geschrieben für Albis ab Version 18.40 (Windows XP,8,10)

<img src="Docs\TrennerExtraBreit.png" style="zoom: 67%;" />

### DIES IST KEINE OFFIZIELLE ERWEITERUNG FÜR ALBIS!
##### Der Hersteller der Praxissoftware Albis*, die Compugroup AG, hat mit dieser Software nichts zu tun, geschweige denn, wurde diese offiziell durch die CompuGroup AG legitimiert!
###### * Arztinformationssystem (AIS), Arztsoftware, Praxissoftware, Praxisverwaltungssoftware, Praxisverwaltungssystem (PVS), Praxismanagementsoftware oder Ordinationsmanagementsoftware
<img src="Docs/TrennerExtraBreit.png" style="zoom:67%;" />



# ![Funktionen.png](Docs/Icons/Funktionen.png)  FUNKTIONSÜBERSICHT
## ![Addendum.png](Docs/Icons/Addendum.png) Addendum

<h3> &#9733 PopupBlocker</h3>

- Schließt diverse eher störende, wenig wichtige Fenster in Albis und zugehörigen Programmen

<h3> &#9733 Fensterhandler</h3>

- positioniert Fenster zur besseren Übersicht auf dem Monitor (bestimmte Bereiche im Albisfenster werden nicht mehr verdeckt)
- automatische Bestätigung von Dialognachfragen
- erweitert automatisch Anzeige-Elemente in Albisdialogen für mehr Übersicht 

<h3> &#9733 Auto-Login </h3>

- kann auf Wunsch das Login in ihr Albisprogramm vornehmen

<h3> &#9733 Patientennamendatenbank </h3>

- für eine schnelle und fehlertolerante Suche nach Patienten

<h3> &#9733 alternatives Tagesprotokoll </h3>

- nach Computern getrenntes Tagesprotokoll. Dies könnte für einzelne Analysen interessant sein

<h3> &#9733 Infofenster </h3>

- zeigt den Inhalt ihres Bildvorlagen-Dateiordners an (PDF, sowie Bilddateien)
- Befund-/Bilddateien können von dort ohne Umweg in die Patientenakte importiert werden
- Auflistung aller Karteikarten des Tages (letzte Patienten) 

<h3> &#9733 Unterstützung der PDF-Signierung </h3>

- durch Druck auf einen Hotkey wird die aktuell im FoxitReader geöffnete PDF Datei signiert. (eine Signatur müssen Sie vorher im FoxitReader erstellt haben). Dadurch ist sogar eine ***gesetzeskonforme und rechtssichere Signierung*** der Dateien möglich. **Achtung:** es gibt keine kostenlose Software zur digitalen Signierung. Den [FoxitReader](https://www.foxitsoftware.com/de/pdf-reader/) müssen Sie bei professioneller Nutzung liszensieren lassen! Ebenso die genutzten Command-Line-Tools - [xpdf-Tools](http://www.xpdfreader.com/) und pdftk.

<h3> &#9733 Menusuche </h3>

- Finden und Aufrufen von Menupunkten im Albismenu

<h3> &#9733 Vereinfachung der Albisbedienung </h3>

- erweiterte Tastenkombinationen (Hotkeys) für zusätzliche Funktionalität
  - **Verschieben von Einträgen** im Dauermedikamenten- und Dauerdiagnosenfenster 
  - **Kopieren**, **Ausschneiden** und **Einfügen** mit der von Windows gewohnten Tastenkombination
  - **Schließen** einer Krankenakte oder **Anzeigen** der nächsten geöffneten
  - Einstellen des aktuellen **Tagesdatums**
  - **Addition** von einer **Woche** oder einem **Monat** in **Datumsfeldern**, anstatt nur einem Tag wie bisher 
  - **neuer Shift + F3 Kalender** zeigt mehrere Monate an 

<h3> &#9733 erweitertes Kontextmenu </h3>

- mehr Funktionen im Rechtsklick Menu in der Karteikarte. Bearbeiten (Anzeigen), Drucken, Exportieren, Versand als Fax, per Mail oder per Telegram 

<h3> &#9733 Rezepthelfer </h3>

- Rezeptvorlagen, z.B. mehrzeilige Hilfsmittelrezepte oder Verschreibung mehrerer Medikamente nach Auswahl einer Vorlage 

<h3> &#9733 Anzeige von Beginn und Ende der Lohnfortzahlung  </h3>

- die berechneten Stichtage werden im oberen Teil der Arbeitsunfähigkeitsbescheinigung eingeblendet 

![](Docs/AUFristen.png)

<h3> &#9733 Kontextsensitive Texterweiterungen </h3>

- Erkennung des Kontext in der Karteikarte anhand des Albiskürzel (z.B. lko, dia, bef, info) mit Bereitstellung von Texterweiterungen 

<h3> &#9733 Unterstützung des Laborabrufes </h3>

- Übernimmt automatisch nach manuellem Start des Laborabrufes die weiteren Vorgänge bis hin zum Übertragen der Befunde ins Laborblatt  (funktioniert nur teilweise)

<h3> &#9733 Automatisierung DICOM-Daten Umwandlung </h3>

- [MicroDicom](https://www.microdicom.com), der freie DICOM-Viewer für Windows, wird automatisiert für eine schnelle Umwandlung der Daten in Bild- oder Videodateien 

<h3> &#9733 Funktionsbibliothek für eigene Skriptentwicklung </h3>

- **107 Funktionen** zur Steuerung von Albis sind für die Entwicklung eigener Skripte vorhanden 



![](Docs/TrennerExtraBreit.png)

## ![](../../../../../Eigene%20Dateien/Eigene%20Dokumente/AutoIt%20Scripte/GitHub/Addendum-fuer-Albis-on-Windows/Docs/Icons/Achtung.png) *WICHTIG*

- Ich empfehle die Skripte **nicht** zu **kompilieren!** 

- **Entpacken** Sie die Dateien am besten **auf ein Netzwerklaufwerk** auf das sämtliche Computer in Ihrem Netzwerk Zugriff haben. Alle Skripte greifen auf eine gemeinsame Einstellungsdatei zurück (Addendum.ini), damit z.B. bei Neuinstallation eines Computer sämtliche Einstellungen noch vorhanden sind und nicht extra gesichert werden müssen.

- Lassen Sie die Skripte am besten in den Ordnern in denen diese nach Entpacken sind, da die Programmbibliotheken per relativem Pfadbezug hinzugeladen werden

- es lohnt sich nicht ein einzelnes Skript herunter zu laden, da Bezüge und Aufrufe untereinander bestehen, manche Skripte kommunizieren auch untereinander. 

- **DENKEN** sie immer an den **BACKUP** ihrer wichtigen Daten!

  

## **Haben Sie keine Sorge vor dem Verlust von Praxisdaten durch die Verwendung der Skripte!**

DIE SKRIPTE SCHREIBEN NICHTS IN IHRE DATENBANK! NOCH LESEN SIE ETWAS AUS IHRER DATENBANK AUS! SIE MACHEN DAS, WAS SIE AUCH TÄGLICH MACHEN: SIE ARBEITEN *!AUSSCHLIESSLICH!* MIT DER PROGRAMMOBERFLÄCHE! ES GIBT AUCH KEINERLEI FUNKTIONEN DIE IRGENDETWAS IN DEN AKTEN ODER IN IHREN DATEN LÖSCHEN WÜRDEN OHNE MINDESTENS ZUVOR EIN BACKUP DESSEN ANZULEGEN! ES WIRD KEINE INTERNETVERBINDUNG BENÖTIGT! DIE SKRIPTE VERSENDEN KEINE DATEN ÜBER DAS INTERNET AN IRGENDWEN!

Da alles in Autohotkey geschrieben ist, läßt sich sämtlicher Code in einem normalen Texteditor lesen (Einschränkung: 2 Funktionen mit Maschinencode (Assembler) - einsehbar im Autohotkey-Forum). 

#### *RECHTLICHE HINWEISE UND DIE LIZENSIERUNGSBEDINGUNGEN FINDEN SIE AM ENDE DES DOKUMENTES!*

<br>

![](Docs/TrennerExtraBreit.png)

<br>

## ![](Docs/Icons/Addendum.png) <u>Dauermedikamentenfenster</u>

![](Docs/Dauermedikamente.png) 

Auswahlmöglichkeit aus voreingestellten Kategorien (in Patienten verständlichem Deutsch). Erreichbar nach Drücken (#) der ![](Docs/Icons/Raute.png) Taste in einer geöffneten Zeile. Die Kategorien lassen sich im Quelltext des Addendum.ahk-Skriptes jederzeit ändern. (an einer komfortableren Lösung wird gearbeitet). Mit der Tastenkombination (![](Docs\Icons\Key-White_Strg-Links.png)+ ![](Docs\Icons\hoch.png)oder![](Docs\Icons\runter.png)) lassen sich alle Einträge innerhalb der Ansicht verschieben.

<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

## ![Addendum](Docs/Icons/Addendum.png) <u>neuer Shift+F3 Kalender</u>

![Menu-Suche](Docs/NeuerShiftF3Kalender.png)



<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

## ![Addendum](Docs/Icons/Addendum.png) <u>Menu-Suche</u>

![Menu-Suche](Docs/Menu_Suche.png)

Albis On Windows hat mehr als **740** Menupünkte. Seltene genutzte Formulare zu finden dauert meist ziemlich lange. Drücke ![](Docs/Icons/Alt.png) + ![](Docs/Icons/Key-White-M.png) für einen Suchdialog und öffne den Menupunkt von hier aus - Danke an *Lexikos* dem Author von Autohotkey für dieses wunderbare Skript. 



<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

## ![Hotkey](Docs/Icons/Hotkey.png) <u>sinnvollere Albiskürzel</u>

- **Kopieren**, **Ausschneiden** und **Einfügen** ist mit den üblichen Kürzeln überall in Albis möglich <br>
    - **Kopieren:**                                      ![Strg](Docs/Icons/Key-White_Strg-Links.png)+![c](Docs/Icons/Key-White-c.png)
    - **Ausschneiden:**                            ![Strg](Docs/Icons/Key-White_Strg-Links.png)+![x](Docs/Icons/X.png)
    - **Einfügen:**                                       ![Strg](Docs/Icons/Key-White_Strg-Links.png)+![v](Docs/Icons/Key-White-v.png)
  - **Schließen einer Karteikarte:**   ![Alt](Docs/Icons/Alt.png)+![Runter](Docs/Icons/runter.png) 
  - **zur nächsten Karteikarte:**        ![Alt](Docs/Icons/Alt.png)+![Hoch](Docs/Icons/hoch.png) 
  - **Laborblatt zeigen:**                      ![Alt](Docs/Icons/Alt.png)+![Rechts](Docs/Icons/Rechts.png) 
  - **Karteikarte zeigen:**                    ![Alt](Docs/Icons/Alt.png)+![Links](Docs/Icons/Links.png) 
  - **Einstellen des aktuellen Tagesdatums:**   ![Alt](Docs/Icons/Alt.png)+![F5](Docs/Icons/F5.png) 
  
  

<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

## ![Hausbesuche.png](Docs/Icons/Hausbesuch.png) <u>Formularhelfer Hausbesuche</u>

**Ausdrucken von Rezepten/Über- und Einweisungen ohne gedrucktes Datum**

![Formularhelfer](Docs/Formularhelfer.png)

* ein Fenster mit 5 Formularen (Kassenrezept, Privatrezept, Krankenhauseinweisung, Krankenbeförderung, Überweisung)
* Auswahl der Formularanzahl 0 für keine und maximal 9 (auch über die Zifferntastatur), nach Eingabe einer Ziffer rückt der Eingabefocus zum nächsten Feld weiter. Ist alles eingegeben, dann nur noch Enter drücken. Den Rest übernimmt das Skript. Es ruft die jeweiligen Formulare auf, entfernt wenn notwendig den Haken am Datumsfeld, setzt automatisch die Anzahl und drückt den Knopf für Drucken. 
* im Addendumskript ist ein Hotstring hinterlegt (*FHelfer*). Diesen in irgendeinem Eingabefeld in Albis eingegeben und das Skript startet. 
* 1 Mausklick, 7 Buchstabentasten, max. 5 Ziffern und 1xEnter müssen gedrückt werden. **Das wars!** Für einen Hausbesuch sind die Unterlagen vorbereitet. Optional kann für jeden Patienten noch sein Patientenstammblatt ausgedruckt werden. 



<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

##  ![Rezepthelfer.png](Docs/Icons/Rezepthelfer.png) <u>Schnellrezepte</u>

![](Docs/Schnellrezepte.gif)



<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

### ![KontextMenu](Docs/Icons/KontextMenu.png) <u>Erweitertes Kontextmenu</u>

**mehr Funktionen im Rechtsklick Menu in der Karteikarte**

![erweitertes Kontextmenu](Docs/erweitertes_Kontextmenu.png)

Je nach Karteikartenkürzel werden verschiedene Funktionen angeboten. Unter anderem Öffnen eines Formulares zum Bearbeiten oder der direkte Ausdruck. Wenn Sie PDF-Dateien oder Bild-Dateien direkt in die Karteikarte ablegen, können Sie diese Dateien ebenso Anzeigen oder Ausdrucken lassen, aber die Dateien lassen sich auch in einen nach dem Patienten benannten Dateiordner exportieren (z.B. zur Erleichterung beim Arztwechsel) und da der Faxversand eigentlich auch nur ein Druckvorgang ist, geht auch das inklusive Abfrage der Faxnummer (wenn gewünscht)



<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

## ![](Docs/Icons/Monet.png) <u>MoNet</u> 

**Monitor for your Network**

![Beispielbild des Monet Modules](Docs/Modul%20MoNet%20Screenshot.png)

Statusanzeige der Computer in der Praxis (im Moment nur an oder aus). Bestimmte Computer können für den Zugriff gesperrt werden um versehentliches Herunterfahren zu vermeiden. Entsperren geht nur per Passwort.

Soll ein Clientkommunikationsmodul werden. angedachte Punkte:

- **Herunterfahren** eines entfernten Clients von jedem Rechner aus
- **Hochfahren** per WakeOnLan
- **Kurznachrichtenchat** - die Albis interne Nachrichtenfunktion ist mal wieder ausgefallen
- **Remote-Start** von Skripten auf anderen Clienten
- **Überwachungsfunktionen** - melden z.B. eines Serverausfall, abgestürzter oder nicht funktionierender Software z.B. über eine Nachricht per Telegram App (Funktion)<br>

**Beispielbild des Netzwerkmonitors**. 
Die IP-Adressen und Namen der Rechner müssen in der Addendum.ini hinterlegt sein



<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

## ![InfoIcon](Docs/Icons/Infofenster.png) <u>Infofenster</u>

**Befundeingang, Tagesprotokoll, Praxisinfos**

![Infofenster](Docs/Infofenster.png)

<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

## ![](Docs/Icons/Abrechnungshelfer.png) <u>Abrechnungshelfer - optimieren Sie Ihre Abrechnung</u>

<b>mehr Statistiken, mehr Möglichkeiten der Automatisierung</b>

<img src="Docs/Screenshot-Abrechnungshelfer.png" alt="Screenshot Abrechnungshelfer" style="zoom: 100%;" />

Erstellen Sie ein Tagesprotokoll und nutzen Sie diese Modul um von der Compugroup nicht abgedeckte Regeln zu möglichen Abrechnungsziffern zu entwerfen oder nutzen Sie das Skript um Patienten bestimmten Gruppen zuzuordnen oder um eine erweiterte Statistik durchführen zu können
<br>fertige Statistiken:

- **freie Statistik** - mittels *RegEx* in Tagesprotokollen suchen (**!nur die Gui ist fertig!**)
- **Patienten für die Vorsorgeliste suchen** - findet Patienten bei denen eine Vorsorgeuntersuchung (GVU und/oder Hautkrebsscreening durchgeführt werden kann)
- **nach fehlenden GB Ziffern suchen** - erstellt eine Liste von Patienten bei denen der geriatrische Basiskomplex noch nicht abgerechnet wurde
- **fehlende Chronikerziffern** - erstellt eine Liste von Patienten bei denen die Ziffern 03220 oder 03221 nicht abgerechnet wurden, obwohl dies in den Vorquartalen erfolgt war 
- **neue Chroniker finden** - findet Patienten bei denen man die Ziffern 03220/03221 ansetzen kann. Das Skript nutzt dazu eine Liste von ICD-Schlüsselnummern des Bewertungsausschusses nach § 87 Absatz 1 SGB V, die nach Einschätzung der AG medizinische Grouperanpassung chronische Krankheiten kodieren.
- **Dauerdiagnosenstatistik** - listet und zählt alle Dauerdiagnosen aus dem gewählten Tagesprotokoll, jede Diagnose beeinhaltet auch eine Liste der entsprechenden Patienten



<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

## ![Laborabruf](Docs/Icons/LaborAbruf.png) Laborabruf (teilweise in Addendum integriert)

**Die Automatisierung für den Abruf der Laborwerte**

- der Abruf der Laborwerte ist jetzt teilautomatisiert und ins Addendum-Skript integriert
- Skript erkennt einzelne Abschnitte des Abrufs von Labordaten und übernimmt dann die immer wieder kehrenden Eingabetätigkeiten

- erkannt werden:

  - Öffnen des Labordatenimport-Fensters z.B. nach Aufruf über den Menupunkt Extern/Labor/Daten importieren
    - es wird alles eingetragen was benötigt wird und der Vorgang wird gestartet
    - im Anschluss wird sofort das Laborbuch geöffent
  - im Laborblatt werden nach Aufruf der Funktion ..alle ins Laborblatt.. , sämtliche sich dann öffnenden Dialoge automatisch bearbeitet.Es ist kein weiterer Eingriff durch den Nutzer notwendig.   



<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

## ![GVU](Docs/Icons/GVU.png) Gesundheitsvorsorgeliste

**automatisierte Formularerstellung**

*dieses Skript nutzt die zuvor erstellte **GVU-Formular-Liste** untersuchter Patienten*. 

- **quartalsweise** und **datumsgenaue** Erstellung der Gesundheitsvorsorge- und der Hautkrebsscreeningformulare
- die **Leistungskomplexe** werden automatisch erstellt und in die Akte geschrieben
- nutzt einen Eintrag im **CaveVon**-Fenster um Ihnen den schnellen Überblick zu ermöglichen, wann die letzte Untersuchung war und wann die nächste fällig wäre

***Wichtig***: *es erfolgt kein Zugriff auf die Albis-Datenbanken. Die Erstellung erfolgt vereinfacht ausgedrückt durch die Simulation von Maus und Tastatureingaben^1^.  In der vorherigen Version brauchte das Skript für circa 80 Formulare ungefähr 3 Stunden (wohl gemerkt! Es macht nichts anderes als ein Mensch, ist aber wesentlich schneller!). Wie lange brauchen Sie für diese Formularanzahl? Ich denke darüber haben Sie sich noch nie Gedanken gemacht.*

^1^ mit ausschließlich simulierten Tasten- und Mauseingaben ist ein fehlerfreier Ablauf unmöglich. Das Skript verwendet nicht einen einzigen Befehl zum Senden von Tasten. Die Formularfelder werden ausschließlich direkt manipuliert. Dies macht das ganze aber auch deutlich aufwendiger. Mit einem reinen Maus- und Tastaturrekorder erreichen Sie niemals diesen hohen Grad an Zuverlässigkeit.

*bitte lesen Sie die Anleitung zur Benutzung des Skriptes, sie finden diese in den ersten Zeilen des Skriptes selbst*



<img src="Docs/TrennerExtraBreit.png" style="zoom: 67%;" />

# ![OutOfTheBox](Docs/Icons/OutOfTheBox.png)  <u>KEIN "OUT OF THE BOX"</u>

Dies ist **"keine out of the box"** Lösung! Es müssen entsprechende Anpassungen an Ihre Praxisumgebung manuell in den Skripten vorgenommen werden. Entsprechende Kenntnisse in Programmier- oder Skriptsprachen sind hilfreich. Die Skripte basieren auf einer gemeinsamen speziellen Funktionsbibliothek (Praxomat_Functions.ahk) und weiteren Bibliotheken für AHK. Ich nutze keine Klassenbibliotheken, zum einen da ich in der Objekt orientieren Programmierung kaum wissen habe und zum anderen weil ich einfach davon ausgehe das OOP das es den meisten Kollegen auch so geht.
<br>
<img src="Docs/TrennerExtraBreit.png" style="zoom: 67%;" />

# ![Paragraphen](Docs/Icons/Paragraphen.png) <u>RECHTLICHE HINWEISE</u>

**FOLGENDE ABSCHNITTE GELTEN FÜR ALLE TEILE UND DIE GESAMTE SAMMLUNG DIE UNTER DEM NAMEN** **"Addendum für Albis On Windows"** (nachfolgend Skriptsammlung genannt) herausgegeben wurde

DIE SKRIPTSAMMLUNG IST EIN HILFSANGEBOT AN NIEDERGELASSENE ÄRZTE. 

KOMMERZIELLEN UNTERNEHMEN, DIE SICH MIT DER HERSTELLUNG, DEM VERTRIEB ODER WARTUNG VON SOFT- ODER HARDWARE BESCHÄFTIGEN IST DIE NUTZUNG ALLER INHALTE ODER AUCH NUR TEILE DES INHALTES NUR NACH SCHRIFTLICHER ANFRAGE MIT ANGABEN DER NUTZUNGSGRÜNDE UND MEINER SCHRIFTLICHEN FREIGABE GESTATTET! UNBERÜHRT DAVON SIND DIE VON MIR BENUTZTEN FREMDBIBLIOTHEKEN!

DIESES REPOSITORY DARF AUF EIGENEN SEITEN VERLINKT WERDEN.

DIESE SOFTWARE IST **FREIE SOFTWARE**! (*!freie Software ist nicht dasselbe wie Open Source!*) 

DIE SAMMLUNG ENTHÄLT SKRIPTE/BIBLIOTHEKEN AUS ANDEREN QUELLEN. DAS COPYRIGHT SIEHT IN JEDEM FALL EINE FREIE VERWENDUNG FÜR NICHT KOMMERZIELLE ZWECKE VOR. AUS DIESEM GRUND KÖNNEN DIESE SAMMLUNG ODER AUCH NUR TEILE DAVON ZU KEINEM ZEITPUNKT VERKÄUFLICH SEIN! ANFRAGEN JEGLICHER JURISTISCHER ODER NATÜRLICHER PERSONEN HINSICHTLICH KOMMERZIELLER ANSÄTZE WERDEN IGNORIERT!
<br>
<br>

<img src="Docs/TrennerExtraBreit.png" style="zoom: 67%;" />

# ![Haftungsauschluss](Docs/Icons/Haftungsausschluss.png) <u>AGB’s / HAFTUNGSAUSSCHLUSS</u>

**I.a.** Der Download und die Nutzung der Skripte unterliegen der GNU Lizenz welche von Lexikos dem Gründer der Autohotkey Foundation erstellt wurden. 

Die Inhalte und Skripte dürfen AUSSCHLIESSLICH NUR von Praxisinhabern, die zum Zeitpunkt des Downloads oder bei Nutzung in EIGENER PRAXIS als selbstständig tätige approbierte Ärzte sind, installiert, authorisiert, ausgeführt und aber auch frei verändert werden. Ein Auftrag zum Download, Änderungsvorgaben, Nutzungsänderungen an Angestellte (dies gilt auch für angestellte (approbierte) Ärzte oder eine Auftragsvergabe an ein Unternehmen sind **UNTERSAGT**! Nach jeder noch so kleinen Änderung an einem der Skripte oder den grafischen Beigaben ist zusätzlich zu meinem Pseudonym im jeweiligen Skript ihr eigenes Pseudonym oder ihr eigener Namen hinzu zufügen, falls Sie das veränderte Skript weitergeben möchten. Eine Weitergabe veränderter Dateien ausschließlich unter meinem Pseudonym **UNTERSAGE** ich.

**I.b.** ***Nicht gemeinnützigen Unternehmen*** (wie Softwarehäusern, Vertrieblern, Krankenhäusern, Servicefirmen und all Ihren Mitarbeitern) ist ***weder*** die Erwähnung (insbesondere zu Werbezwecken) noch das Weiterverwenden/Einbinden, das Ändern der Skripte oder die Nutzung meiner Ideen in jeglicher Form, selbst in ähnlicher Weise zum Verändern eigener oder anderer Software z.B. um hiermit Geld zu generieren **UNTERSAGT!** In diesen Fällen behalte ich mir vor einen Urheberanspruch an das kommerziell arbeitende Unternehmen zu stellen, den ich im entsprechenden Fall gesetzlich durchsetzen lasse werde.

**I.c.** Universitäten, interessierten Studenten gestatte ich eine Nutzung nur nach schriftlicher Genehmigung und persönlichem Nachweis einer gemeinnützigen Sache, meiner Haftungsfreistellung und einer Erklärung in keinster Weise mit irgendeinem Unternehmen oder einer juristischen Person wie unter I.b. zusammenzuarbeiten. Die eventuellen Auslagen und Kosten trägt der Anfrage/Antragsteller!<BR>

**I.d.** Gemeinnützige Einrichtungen aus nicht EU-Ländern und nicht-Gxx-Ländern ist die Nutzung wie den Personen unter I.a. unter allen Haftungsausschlüssen für mich jederzeit und ohne Nachfrage gestattet! Hier erkläre ich mich gerne ebenso wie bei I.a. bereit im Falle von Fragen zu helfen.<BR>

#### Download

**II.** mit dem Herunterladen oder Speichern der Dateien in jeglicher Form (entpackt, gepackt, kompiliert, unkompiliert, in Teilen, als Ganzes) übernehmen Sie die volle Verantwortung für deren Inhalt. Ich brauche keinen Nachweis wer die Skripte nutzt, nutzen will oder wer Veränderungen daran vorgenommen hat oder vornehmen will. Desweiteren akzeptieren sie vollständig alle unter Punkt III stehenden Handlungs- und Haftungsregeln.<BR>

#### Haftungsausschluß / Haftungsfreistellung

**III.** die Nutzung des gesamten Inhaltes (einschließlich Programme, Skripte, Bilder, Bibliotheken, Fremdprogramme) erfolgt auf eigene Gefahr! Das Angebot wurde besten Gewissens auf mögliche Urheberrechtsverletzungen untersucht. Es dürften keine Verletzungen enthalten sein. Quellenangaben sind soweit es nachvollziehbar war in den jeweiligen Dateien enthalten.<BR>

**III.a.** ich übernehme keine Haftung durch vermeintliche, unwissentliche oder wissentliche Fehler in den Skripten<BR>

**III.b.** ebenso übernehme ich keine Haftung für Urheberrechtsverletzungen durch Dritte deren Software oder SourceCode hier verwendet wurde/wird und werden wird.<BR>

**III.c.** Sie akzeptieren das die Skripte nur in einer Alpha-Version vorliegen, sie sehen den Inhalt als ein Beispielprojekt! Die dargebotenen Skripte sollen keinen Abrechnungsbetrug ermöglichen, sondern im Gegenteil nur zeigen, welche Möglichkeiten im Einsatz von Software liegen.<BR>

**III.c.** Sie akzeptieren das die Skripte nur in einer Alpha-Version vorliegen. Sie sehen den Inhalt als  Beispielprojekt! Die angebotenen Skripte sollen keinen Abrechnungsbetrug ermöglichen, sondern im Gegenteil nur zeigen, welche Möglichkeiten der Einsatz von RPA-Software in der Praxis bringen würde.<BR>

**III.d.** Eine Richtigkeit der in der in der Software hinterlegten Leistungsziffern kann nicht garantiert werden,es liegt in Ihrer Verantwortung alles zu kontrollieren<BR>

**III.e.** sämtliche Automatisierungsskripte können gegen Ihr lokales KV-Recht oder der KBV verstoßen - das Haftungsrisiko übernehmen sie! Die Skripte sind als Beispiele für mögliche Anwendungen durch Nutzung der freien Skriptsprache Autohotkey zu verstehen (siehe auch IIIc).<BR>

**III.f.** Sie haben kein Recht Updates, Fehlerkorrekturen oder eine Behebung von Folgeproblemen (vermeintlich oder beweisbar) einzufordern. Ebenso haben Sie kein Recht Schadensersatz für (vermeintliche oder beweisbare) Fehler zu fordern.<BR>

**III.g.** Insbesondere distanziere ich mich von jeglichen Versuchen meine Skripte für die Überwachung von Mitarbeitern oder anderen Personen einzusetzen.<BR>

**III.h.** Für Fehler in Rechtschreibung, Grammatik bin ich ebenso nicht verantwortlich zu machen. Dazu wenden Sie sich bitte an meine Deutschlehrer.<BR>

**III.i.** ich übernehme keinerlei HAFTUNG aufgrund hier fehlender rechtlicher Hinweise/Aus- oder Einschlüße. Ihnen sollte nach dem lesen bekannt und bewußt sein, daß ich keinerlei kommerzielle Zwecke verfolge und die Zusammenstellung der Dateien nicht zum Zwecke eigener Bekanntheit erfolgt und ich deshalb niemals wissentlich oder absichtlich fremdes geistiges Eigentums entwendet habe. Die angebotene Sammlung verfolgt ausschließlich unten genannten gemeinnützigen Zweck.<BR>

**III.j** ich erkläre hiermit, das bewußt keine Informatiker, IT-Experten oder Angestellten einer Hotline für dieses Projekt genervt oder gequält wurden. Irgendwelche Ähnlichkeiten zu lebenden oder verstorbenen Personen sind rein zufällig.*

###### * (! für Juristen ! - Artikel ‘III.j’ ist Satire!)


<br>
<br>

<img src="Docs/TrennerExtraBreit.png" style="zoom: 67%;" />

# ![Abschluss](Docs/Icons/Abschluss.png) <u>ABSCHLUSS</u>

## ![Aufruf](Docs/Icons/Aufruf.png) <u>Aufruf zur Gemeinnützigkeit</u>

**DER FINANZIELLE DRUCK DER AUF UNS NIEDERGELASSENEN ÄRZTEN LASSTET, HINSICHTLICH EINER SCHNELLEN HILFREICHEN VERWALTUNGSSOFTWARE, EINER SOFTWARE DIE STÄNDIG NEUEN VORGABEN GEWACHSEN SEIN SOLL, DEM ZUNEHMENDEM DRUCK DER KOMMERZIALISIERUNG WICHTIGER INFORMATIONEN (Studien, Diagnostische Hilfen, Therapien, Formulare etc.), DEM VERMEINTLICHEN ZERTIFIZIERUNGSWAHNSINN, DEN EINGRIFFEN DES GESETZGEBER IN UNSERE ARBEITSZEIT - MUSS DRINGEND EINHALT GEBOTEN WERDEN! DIES HIER IST EIN AUFRUF AN ALLE KOLLEGEN SICH DARÜBER GEDANKEN ZU MACHEN!**

LIEBE KOLLEGEN, DENKEN SIE AUCH AN EINE ENTLASTUNG IHRER ANGESTELLTEN!  UM SO MEHR SIE DIESE VON IMMER WIEDERKEHRENDEN TÄTIGKEITEN AM PC BEFREIEN, UM SO MEHR KÖNNEN ***AUCH SIE SICH*** VON UNNÖTIGER ARBEIT ENTLEDIGEN!



## ![Projekt](Docs/Icons/Projekt.png) <u>Projektmitarbeit</u>

Die Mitarbeit von Kollegen ist ausdrücklich erwünscht! Der Umfang meiner Skripte behindert leider immer wieder die Fehlerkorrekturen, da ich meist an einem anderen Skript arbeite, während ein anderes plötzlich kritische Fehler zeigt. 

Da dieses Projekt unter einem gemeinnützigen, nicht kommerziellen Aspekt steht (Hilfe zur Selbsthilfe, !Mehr Zeit für Patienten!) stehen keine Gelder zur Verfügung! Dieses Projekt soll unsere tägliche Arbeit unterstützen und nicht zu einer weiteren Bereicherung der Medizinindustrie führen. Die Mitarbeit am Projekt endet sofort, wenn ein hinreichender Verdacht auf kommerzielle Interessen besteht! 

Da dieses Projekt unter einem gemeinnützigen, nicht kommerziellen Aspekt steht (Hilfe zur Selbsthilfe, !Mehr Zeit für Patienten!) stehen keine Gelder zur Verfügung! Dieses Projekt soll unsere tägliche Arbeit unterstützen. 



<center> - IXIKO 2020 - </center>

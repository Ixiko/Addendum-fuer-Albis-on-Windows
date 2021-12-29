![Addendum.svg](assets/AddendumLogo2021-2.svg)
#### Extends the possibilities through surface analysis and manipulation

### Release Version 1.68 vom 05.12.2021
lauffähig ausschließlich nur mit *Autohotkey_H ab V1.1.33+ Unicode 64bit*,
geschrieben für Albis ab Version 18.40 (Windows XP,8,10)

![Trenner](Docs/TrennerExtraBreit.png)

### DIES IST KEINE OFFIZIELLE ERWEITERUNG FÜR ALBIS!
##### Der Hersteller der Praxissoftware Albis*, die Compugroup AG, hat mit dieser Software nichts zu tun, geschweige denn, wurde diese offiziell durch die CompuGroup AG legitimiert!
###### * Arztinformationssystem (AIS), Arztsoftware, Praxissoftware, Praxisverwaltungssoftware, Praxisverwaltungssystem (PVS), Praxismanagementsoftware oder Ordinationsmanagementsoftware

![Trenner](Docs/TrennerExtraBreit.png)    

## &#9733; NEUE FUNKTIONEN SEIT DEM LETZTEM RELEASE

![Trenner](Docs/TrennerExtraBreit.png)    

### &#10042; wichtigste Neuerung: *Addendum hat Zugriff auf Albis Datenbanken erhalten*

​		<u>**Module mit Datenbankzugriff**</u><br>

- **Laborjournal**
- **Laborabruf**
- **Export von Patientendaten in einem Durchgang** (Dokumente, Laborblatt, Karteikarte)
- **Abrechnungshelfer** (ähnlich der Albis internen Funktionen) im Infofenster
- **Quicksearch** - den Inhalt aller dBASE Datenbanken anzeigen und gezielt durchsuchen
- **Abrechnungsassistent**  - neue Version 

### &#10042; neue Erweiterungen für das Infofenster

- automatische **Texterkennung** mit Tesseract (Hotfolder)      
- **Autonaming** Assistent
- **Vereinfachung des manuellen Umbennens** einer Datei durch Dateivorschau
- **EMail Versand von Labordaten** an Patienten über die in Albis hinterlegte E-Mailadresse
- **Impfstatistik** für das RKI

### &#10042; weitere Funktionen

- medizinische Berechnungen
- Faxanhänge aus Outlook extrahieren 
- Albis reanimieren
- **PushToTelegram**



![](Docs/TrennerExtraBreit.png)

## ![Addendum.png](Docs/Icons/Addendum48x48.png)ddendum


### &#9733; PopupBlocker

- schließt automatisch diverse störende Dialoge von Albis und anderen Programmen

### &#9733; Fensterhandler

- positioniert und erweitert Albisdialoge oder einzelne Elemente für eine bessere Übersicht 
- keine Anzeige von Werbung mehr beim Kassenrezept
- Albis *kann* individuell an jedem Arbeitsplatz in Größe und/oder Position fixiert werden 

### &#9733; Auto-Login 

- kann auf Wunsch das Login in ihr Albisprogramm vornehmen

### &#9733; eigene Patientendatenbank 

- für eine schnelle und fehlertolerante Suche (Fuzzy Stringvergleich) nach Patienten

### &#9733; Infofenster 

​        zentral ins Albisfenster integriertes Tool für Dokument-Eingänge und Verwaltung ihres Praxisnetzwerk

- **Patient**: zeigt den Posteingang des aktuellen Patienten und Informationen und Tipps zur Abrechnung an.
- **Journal**: der Posteingang. Funktionen für Texterkennung und automatische Namenserkennung. Organisation von Dokumentimport, -anzeige und benennung
- **Protokoll**: Tagesprotokoll aller geöffneten Karteikarten.
- **Netzwerk:** direkter Start einer Remotedesktopsitzung per Klick 
- **Extras**: häufiger benötigte Programme/Skripte lassen sich im Infofenster anzeigen und von dort aus auch starten
- **Impfen**: Statistik für COVID19-Impfungen per Knopfdruck ins KBV/RKI Formular eintragen lassen 

### &#9733; Texterkennung

- **Texterkennung** mit Tesseract
- **Hotfolder**: neu in einen Ordner hinzugefügte PDF Dateien werden automatisch in PDF-Text Dateien umgewandelt

### &#9733; Autonaming 

- Erkennung von Patientennamen und Erstellungsdatum des Dokumentes bzw. des Zeitraums eines Krankenhausaufenthaltes. Aus diesen Informationen wird ein neuer Dateiname generiert.
- Autonaming wird im Anschluß an eine abgeschlossene Texterkennungsvorgang ausgeführt

###  &#9733; Unterstützung der PDF-Signierung 

- per Hotkey wird die aktuell im FoxitReader geöffnete PDF Datei signiert. (eine Signatur müssen Sie vorher im FoxitReader erstellt haben).<br>**Achtung:** es gibt keine kostenlose Software zur digitalen Signierung. Den [FoxitReader](https://www.foxitsoftware.com/de/pdf-reader/) müssen Sie bei professioneller Nutzung liszensieren lassen! Ebenso die genutzten Command-Line-Tools - [xpdf-Tools](http://www.xpdfreader.com/) und pdftk.

###  &#9733; Menusuche 

- Finden und Aufrufen von Menupunkten im Albismenu

###  &#9733; Vereinfachung der Albisbedienung 

- erweiterte Tastenkombinationen (Hotkeys) für zusätzliche Funktionalität
  - **Verschieben von Einträgen** im Dauermedikamenten- und Dauerdiagnosenfenster und im cave! Dialog 
  - **Kopieren**, **Ausschneiden** und **Einfügen** mit der von Windows gewohnten Tastenkombination
  - **Schließen** einer Krankenakte oder **Anzeigen** der nächsten geöffneten
  - Einstellen des aktuellen **Tagesdatums**
  - **Addition** von einer **Woche** oder einem **Monat** in **Datumsfeldern**, anstatt nur einem Tag wie bisher 
  - **neuer Shift + F3 Kalender** zeigt mehrere Monate an 
  - **Hotstrings** für die schnelle Eingabe von Abrechnungsziffern (EBM und GOÄ, Diagnosen und anderes)

###  &#9733; erweitertes Kontextmenu 

- mehr Funktionen im Rechtsklick Menu in der Karteikarte. Bearbeiten (Anzeigen), Drucken, Exportieren, Versand als Fax, per Mail  (Automatisierung für FoxitReader und Sumatra PDF sind integriert)

###  &#9733; Rezepthelfer 

- Rezeptvorlagen, z.B. mehrzeilige Hilfsmittelrezepte oder Verschreibung mehrerer Medikamente nach Auswahl einer Vorlage 

###  &#9733; sinnvollerer Verordnungsplan  

* Der Bundeseinheitlichen Medikationsplan ist gut, es fehlte nur ein größeres Feld für bekannte Arzneimittelallergien oder unerwünschte Wirkungen welche beim Patienten auftraten 

###  &#9733; Anzeige von Beginn und Ende der Lohnfortzahlung  

- die berechneten Stichtage werden im oberen Teil der Arbeitsunfähigkeitsbescheinigung eingeblendet 

![](Docs/AUFristen.png)

###  &#9733; Kontextsensitive Texterweiterungen 

- Erkennung des Eingabekontext in der Karteikarte anhand des Albiskürzel (z.B. lko, dia, bef, info, lp, lbg) 
- Anlegen eigener Textkürzel die sich automatisch zu Leistungskomplexen, Diagnosen, Befundtexten oder anderes erweitern lassen 

###  &#9733; DICOM-Daten Umwandlung 

- [MicroDicom](https://www.microdicom.com), der freie DICOM-Viewer für Windows, wird automatisiert für eine schnelle Umwandlung der Daten in Bild- oder Videodateien (im Moment MRT/CT Aufnahmenumwandlung ins wmi Format (video) (im Moment fehlerhaft)

###  &#9733; RPA-Funktionsbibliothek

- **>120 Funktionen** zur Steuerung von Albis für die Entwicklung eigener Skripte  

###  &#9733; native DBASE-Klasse		

- für die Analyse der von Albis verwendeten DBASE-Datei Strukturen, Portierung von oder Suche nach Daten

###  &#9733; medizinische Berechnungen	

- Funktionsbibliothek für Berechnungen von BMI, eGFR nach der CKD-EPI Formel, KOF und mehr

###  &#9733; Karteikartenexport		

- Export von Karteikarte, Laborblatt und aller Befunde (Bilder-, Worddokumente und PDF-Dateien) ohne größeren Umweg
- Dokumente lassen sich aus dem Karteikartenexport drucken (z.B. um Kopien für Behörden anzufertigen)
- automatische Berechnung der entstandenen Sachkosten (Kopiegebühren), Sachkostenberechnung läßt sich nach Albis übertragen 

###  &#9733; Laborjournal

- tägliche Übersicht medizinisch relevanter Laborwerte 
- konfigurierbare Ausnahmebehandlung einzelner Laborwerte

###  &#9733; Laborabruf

- zeitgesteuerter Abruf von Labordaten mit anschliessendem Import in das Laborblatt der Patienten. Die Uhrzeiten des Abrufs können selbst festgelegt werden. 

### &#9733; Laborimport
- fast vollständig automatisierter Importvorgang neuer Laborwerte ins Laborblatt der Patienten

###  &#9733; Outlook Anhänge extrahieren

- Faxbefunde erreichen mich von der Fritzbox per EMail. Ein neues Skript extrahiert die als PDF angehängten Fax in den Befundordner. Durch die neue Hotfolder-Funktionalität wird dann sofort eine Texterkennung durchgeführt und anschliessend werden die Dateien kategorisiert.

###  &#9733; Albis Reanimator

- Albis hängt sich häufig auf? Mehrmals am Tag sämtliche geladene von Albis geladene Prozesse mit dem Taskmanager löschen? Damit dann endlich Albis wieder startet. Das übernimmt der Albis Reanimator absofort für Sie! 

###  &#9733; Listview-Inhalte kopieren

- Sie brauchen den Inhalt des Wartezimmers, die Dauerdiagnosen, die Dauermedikamenten oder die Liste der Resistenztestung aus dem Laborblatt. Steuern Sie ihre Maus über das entsprechende Element und drücken Sie ![Strg](Docs/Icons/Key-White_Strg-Links.png)+![c](Docs/Icons/shift_Left.png)+![c](Docs/Icons/Key-White-c.png) und Addendum kopiert nicht nur den Inhalt des Elementes sondern bereitet ihn auf damit sie ihn z.B. in ein Dokument per Paste einfügen können. 

###  &#9733; QuickSearch
visuelle Analyse und für die Suche von Daten in allen Albis (dBase) Datenbankdateien 

- Anzeige von Daten und Suche in allen Datenbanken
- Auslesen spezieller Informationen der Strukturen der Datenbanken
- Inhalte der Datenbank können im .csv Format exportiert werden 



[alle Änderungen in Addendum](Docs/Changes_Addendum_main.md)   |   [alle Änderungen in den Funktionsbibliotheken](Docs/Changes_Addendum_includes.md)

![](Docs/TrennerExtraBreit.png)

## ![](Docs/Icons/Achtung.png) *WICHTIG*

- Ich empfehle die Skripte **nicht** zu **kompilieren!** 

- **Entpacken** Sie die Dateien am besten **auf ein Netzwerklaufwerk** auf das sämtliche Computer in Ihrem Netzwerk Zugriff haben. Alle Skripte greifen auf eine gemeinsame Einstellungsdatei zurück (Addendum.ini), damit z.B. bei Neuinstallation eines Computer sämtliche Einstellungen noch vorhanden sind und nicht extra gesichert werden müssen.

- Lassen Sie die Skripte am besten in den Ordnern in denen diese nach Entpacken sind, da die Programmbibliotheken per relativem Pfadbezug hinzugeladen werden

- es lohnt sich nicht ein einzelnes Skript herunter zu laden, da Bezüge und Aufrufe untereinander bestehen, manche Skripte kommunizieren auch untereinander. 

- **DENKEN** sie immer an den **BACKUP** ihrer wichtigen Daten!

  

Da alles in Autohotkey geschrieben ist, läßt sich sämtlicher Code in einem normalen Texteditor lesen (Einschränkung: 2 Funktionen mit Maschinencode (Assembler) - einsehbar im Autohotkey-Forum). 

#### *RECHTLICHE HINWEISE UND DIE LIZENSIERUNGSBEDINGUNGEN FINDEN SIE AM ENDE DES DOKUMENTES!*

![](Docs/TrennerExtraBreit.png)

<br>

## ![](Docs/Icons/Laborjournal.png) Laborjournal

![](Docs/Screenshot-Laborjournal.png) 

- ***Responsives Webinterface*** (Basis: Internet Explorers). 

- ***Gruppierung der Laborparameter*** nach klinischer Bedeutung, erkennbar durch die dickere Schrift und die unterschiedliche farbliche Hervorhebung. 

   ![#f03c15](https://via.placeholder.com/15/f03c15/000000?text=+) ’‘**immer**‘’ (Anzeige: immer (nur) wenn pathologisch) und ![#9400D3](https://via.placeholder.com/15/9400D3/000000?text=+) ‘’**exklusiv**‘’ (Anzeige: auch bei Normwert).

- ***durchschnittliche Über- oder Unterschreitung:*** 

  - Berechnet die durchschnittliche prozentuale Über- bzw. Unterschreitung der Normwertgrenzen je Laborparameter.
  - Für eventuelle Anpassungen wird die maximale Über- oder Unterschreitung als Einzelwert gespeichert.
  - Durch Nutzung eines Faktors (Prozentwert) erscheinen mir, die durch Annährung erreichten "Warngrenzen", auch bei unterschiedlichen			Einheiten und altersabhängigen Normwertgrenzen klinisch bedeutsame Laborwertveränderungen sicher herauszufiltern.

  

<img src="Docs/TrennerSchmal.png" style="zoom:50%;"/>

## ![](Docs/Icons/Addendum48x48.png) Dauermedikamentenfenster

![](Docs/Screenshot-Dauermedikamente.png) 

Auswahlmöglichkeit aus voreingestellten Kategorien (in Patienten verständlichem Deutsch). Erreichbar nach Drücken (#) der ![](Docs/Icons/Raute.png) Taste in einer geöffneten Zeile. Die Kategorien lassen sich im Quelltext des Addendum.ahk-Skriptes jederzeit ändern. (an einer komfortableren Lösung wird gearbeitet). Mit der Tastenkombination (![](Docs\Icons\Key-White_Strg-Links.png)+ ![](Docs\Icons\hoch.png)oder![](Docs\Icons\runter.png)) lassen sich alle Einträge innerhalb der Ansicht verschieben.



<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

## ![Addendum](Docs/Icons/Addendum48x48.png) neuer Shift+F3 Kalender

![Menu-Suche](Docs/Screenshot-ShiftF3Kalender.png)



<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

## ![Addendum](Docs/Icons/KontextMenu.png) Menu-Suche

![Menu-Suche](Docs/Screenshot-Menu_Suche.png)

Albis On Windows hat mehr als **740** Menupünkte. Seltene genutzte Formulare zu finden dauert meist ziemlich lange. Drücke ![](Docs/Icons/Alt.png) + ![](Docs/Icons/Key-White-M.png) für einen Suchdialog und öffne den Menupunkt von hier aus - Danke an *Lexikos* dem Author von Autohotkey für dieses wunderbare Skript. 



<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

## ![Hotkey](Docs/Icons/Hotkey.png) sinnvollere Albiskürzel

- **Kopieren**, **Ausschneiden** und **Einfügen** ist mit den üblichen Kürzeln überall in Albis möglich <br>
    - **Kopieren:**                                      ![Strg](Docs/Icons/Key-White_Strg-Links.png)+![c](Docs/Icons/Key-White-c.png)
    - **Ausschneiden:**                            ![Strg](Docs/Icons/Key-White_Strg-Links.png)+![x](Docs/Icons/X.png)
    - **Einfügen:**                                       ![Strg](Docs/Icons/Key-White_Strg-Links.png)+![v](Docs/Icons/Key-White-v.png)
  
- **weitere Hotkey-Aktionen**

    - **Schließen einer Karteikarte:**   ![Alt](Docs/Icons/Alt.png)+![Runter](Docs/Icons/runter.png) 
    - **zur nächsten Karteikarte:**        ![Alt](Docs/Icons/Alt.png)+![Hoch](Docs/Icons/hoch.png) 
    - **Laborblatt zeigen:**                      ![Alt](Docs/Icons/Alt.png)+![Rechts](Docs/Icons/Rechts.png) 
    - **Karteikarte zeigen:**                    ![Alt](Docs/Icons/Alt.png)+![Links](Docs/Icons/Links.png) 
    - **Einstellen des aktuellen Tagesdatums:**   ![Alt](Docs/Icons/Alt.png)+![F5](Docs/Icons/F5.png) 

- **Hotstrings** (Beispiele)

    - Hotstring: **Kopie** - automatisiert die Berechnung von Gebühren für Kopien nach Eingabe der Seitenzahl<br>**Kopie** bei ***lp*** als aktives Kürzel oder in der Privatabrechnung eingeben. Im folgenden Dialogfenster die Anzahl der Kopien eintragen. Es wird ein zulässiger Abrechnungstext erstellt und in die Karteikarte geschrieben (z.B. ergeben 38 Seiten:   **lp   |** ***(sach:Kopien 38x a 50 cent:19.00)***

        | Hotstring                              | Erweiterung                                             |
        | -------------------------------------- | ------------------------------------------------------- |
        | **JVEG**/**sozialgericht**             | (sach:Anfrage Sozialgericht gem. JVEG:21.00)            |
        | **lageso**                             | (sach:Landesamt für Gesundheit und Soziales:21.00)      |
        | **lagesokurz**                         | (sach:Landesamt für Gesundheit und Soziales:5.00)       |
        | **Rentenversich**/**RLV** oder **DRV** | (sach:Anfrage Rentenversicherung:28.20)                 |
        | **Bundesa** oder **Agentur**           | (sach:Anfrage Bundesagentur für Arbeit gem. JVEG:32.50) |
        | **porto1**/**Standard**                | (sach:Porto Standard:0.80)	; bis 20g                 |
        | **porto2**/**Kompakt**                 | (sach:Porto Kompakt:0.95) 	; bis 50g                 |
        | **porto3**/**Groß**                    | (sach:Porto Groß:1.55)      	; bis 500g              |
        | **porto4**/**Maxi**                    | (sach:Porto Maxi:2.70)       	; bis 1000g            |

    
    
    <u>Einblendung von Tooltips nach partieller Eingabe des auslösenden Hotstrings:</u>
    
    <img src="Docs/Screenshot-JVEG.png" alt="JVEG SOzial" style="zoom:67%;" />
    
    <img src="Docs/Screenshot-Porto.png" alt="Porto" style="zoom: 67%;" />

<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

## ![Hausbesuche.png](Docs/Icons/Hausbesuch.png) Formularhelfer Hausbesuche

**Ausdrucken von Rezepten/Über- und Einweisungen ohne gedrucktes Datum**

![Formularhelfer](Docs/Screenshot-Formularhelfer.png)

* ein Fenster mit 5 Formularen (Kassenrezept, Privatrezept, Krankenhauseinweisung, Krankenbeförderung, Überweisung)
* Auswahl der Formularanzahl 0 für keine und maximal 9 (auch über die Zifferntastatur), nach Eingabe einer Ziffer rückt der Eingabefocus zum nächsten Feld weiter. Ist alles eingegeben, dann nur noch Enter drücken. Den Rest übernimmt das Skript. Es ruft die jeweiligen Formulare auf, entfernt wenn notwendig den Haken am Datumsfeld, setzt automatisch die Anzahl und drückt den Knopf für Drucken. 
* im Addendumskript ist ein Hotstring hinterlegt (*FHelfer*). Diesen in irgendeinem Eingabefeld in Albis eingegeben und das Skript startet. 
* 1 Mausklick, 7 Buchstabentasten, max. 5 Ziffern und 1xEnter müssen gedrückt werden. **Das wars!** Für einen Hausbesuch sind die Unterlagen vorbereitet. Optional kann für jeden Patienten noch sein Patientenstammblatt ausgedruckt werden. 



<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

##  ![Rezepthelfer.png](Docs/Icons/Schnellrezept.png) Schnellrezepte

![](Docs/Schnellrezepte.gif)



<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

### ![KontextMenu](Docs/Icons/KontextMenu.png) Erweitertes Kontextmenu

**mehr Funktionen im Rechtsklick Menu in der Karteikarte**

![erweitertes Kontextmenu](Docs/Screenshot-erweitertes_Kontextmenu.png)

Je nach Karteikartenkürzel werden verschiedene Funktionen angeboten. Unter anderem Öffnen eines Formulares zum Bearbeiten oder direkter Druck. Wenn Sie PDF-Dateien oder Bild-Dateien direkt in die Karteikarte ablegen, können Sie diese Dateien ebenso Anzeigen oder Ausdrucken lassen. Und man kann diese in einen nach dem Patienten benannten Dateiordner exportieren (z.B. zur Erleichterung beim Arztwechsel). Versehentlich importierte Befunde lassen sich wieder in den Befundordner unter anderem Namen exportieren. Und da der Faxversand eigentlich auch nur ein Druckvorgang ist, geht auch das inklusive Abfrage der Faxnummer (wenn gewünscht)



<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

## ![InfoIcon](Docs/Icons/Infofenster.png) Infofenster

**Befundeingang, Tagesprotokoll, Netzwerkübersicht, Praxisinfos**

![Infofenster](Docs/Infofenster.gif)

## ![InfoIcon](Docs/Icons/Infofenster.png) Impfen

**schnellere Erstellung als mit ALBIS. Ausgabe der Daten in einer Tabelle. Optional können die zusammengestellten Daten ins KBV/RKI Formular eingetragen werden** 

![Impfstatistikvergleich](Docs/Impfstatistikvergleich2.gif)

<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

## ![Abrechnungshelfer](Docs/Icons/Abrechnungsassistent.png) Abrechnungsassistent

<img src="Docs/Abrechnungsassistent.png" alt="Screenshot Abrechnungshelfer" style="zoom: 100%;" />

- bietet Vorschläge zu bestimmten Abrechnungspositionen zu Patienten an

  

  

  <img src="Docs/TrennerSchmal.png" style="zoom:50%;" />


## ![Export](Docs/Icons/DocPrinz.png) DocPrinz

***Skript ist ohne Addendum.ahk ausführbar***

![Addendum_Export.ahk](Docs/Addendum_Exporter.gif)

- **Eingabe** von **Nachname, Vorname, Geburtsdatum oder Patientennummer** 
- alle mit den Suchkriterien übereinstimmenden Patienten werden angezeigt 
- ein **Klick** auf einen Patienten und alle Dokumente des Patienten werden angezeigt 
- **Häkchen setzen** für gezielten Export oder ***‘Alle Dokumente auswählen’*** für eine Komplettauswahl
- ***‘Auswahl exportieren’*** kopiert die Dokumente in einen automatisch erzeugten Unterpfad des Basispfades
-  Laborblatt und Karteikarte und Dokumente lassen sich 

<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

## ![Addendum](Docs/Icons/Labor.png) Labor Anzeigegruppen

- automatische Erweiterung der Fenstergröße und der Steuerelemente für mehr Übersicht
- weitere Fenster welche sich an die automatisch an die Bildschirmgröße anpassen: Rentenversicherung Befundbericht V015, S0051

![Labor Anzeigegruppen](Docs/LBAZGGNANI.svg)

<img src="Docs/TrennerSchmal.png" style="zoom:50%;" />

## ![Laborabruf](Docs/Icons/LaborAbruf.png) Laborabruf

**Die Automatisierung für den Abruf der Laborwerte**

- der Abruf der Laborwerte ist jetzt nahezu vollständig automatisiert 
- Skript erkennt einzelne Abschnitte des Abrufs von Labordaten und übernimmt dann die immer wieder kehrenden Eingabetätigkeiten

- erkannt werden:

  - Öffnen des Labordatenimport-Fensters z.B. nach Aufruf über den Menupunkt Extern/Labor/Daten importieren
    - es wird alles eingetragen was benötigt wird und der Vorgang wird gestartet
    - im Anschluss wird sofort das Laborbuch geöffent
  - im Laborblatt werden nach Aufruf der Funktion ..alle ins Laborblatt.. , sämtliche sich dann öffnenden Dialoge automatisch bearbeitet.Es ist kein weiterer Eingriff durch den Nutzer notwendig.   



<img src="Docs/TrennerExtraBreit.png" style="zoom: 67%;" />

# ![OutOfTheBox](Docs/Icons/OutOfTheBox.png)  "OUT OF THE BOX"

Dies ist noch immer **"keine out of the box"** Lösung! Ein wenig Einarbeitung in Autohotkey wird notwendig sein. Ich habe so gut wie keine grafischen Eingabemöglichkeiten für die Änderungen von Einstellungen bereitgestellt. Die meisten Einstellungen müssen noch händisch per Texteditor in der Addendum.ini vorgenommen werden. Wichtige Einstellungen erreicht man allerdings über das einen rechten Mausklick auf das Addendum Tray Menu.
<br>
<img src="Docs/TrennerExtraBreit.png" style="zoom: 67%;" />

# ![Paragraphen](Docs/Icons/Paragraphen.png) RECHTLICHE HINWEISE

**FOLGENDE ABSCHNITTE GELTEN FÜR ALLE TEILE UND DIE GESAMTE SAMMLUNG DIE UNTER DEM NAMEN** **"Addendum für Albis On Windows"** (nachfolgend Skriptsammlung genannt) herausgegeben wurde

DIE SKRIPTSAMMLUNG IST EIN HILFSANGEBOT AN NIEDERGELASSENE ÄRZTE. 

KOMMERZIELLEN UNTERNEHMEN, DIE SICH MIT DER HERSTELLUNG, DEM VERTRIEB ODER WARTUNG VON SOFT- ODER HARDWARE BESCHÄFTIGEN IST DIE NUTZUNG ALLER INHALTE ODER AUCH NUR TEILE DES INHALTES NUR NACH SCHRIFTLICHER ANFRAGE MIT ANGABEN DER NUTZUNGSGRÜNDE UND MEINER SCHRIFTLICHEN FREIGABE GESTATTET! UNBERÜHRT DAVON SIND DIE VON MIR BENUTZTEN FREMDBIBLIOTHEKEN!

DIESES REPOSITORY DARF AUF EIGENEN SEITEN VERLINKT WERDEN.

DIESE SOFTWARE IST **FREIE SOFTWARE**!

DIE SAMMLUNG ENTHÄLT SKRIPTE/BIBLIOTHEKEN AUS ANDEREN QUELLEN. DAS COPYRIGHT SIEHT IN JEDEM FALL EINE FREIE VERWENDUNG FÜR NICHT KOMMERZIELLE ZWECKE VOR. AUS DIESEM GRUND KÖNNEN DIESE SAMMLUNG ODER AUCH NUR TEILE DAVON ZU KEINEM ZEITPUNKT VERKÄUFLICH SEIN! ANFRAGEN JURISTISCHER ODER NATÜRLICHER PERSONEN HINSICHTLICH KOMMERZIELLER ANSÄTZE WERDEN IGNORIERT!
<br>
<br>

<img src="Docs/TrennerExtraBreit.png" style="zoom: 67%;" />

# ![Haftungsauschluss](Docs/Icons/Haftungsausschluss.png) AGB’s / HAFTUNGSAUSSCHLUSS

**I.a.** Der Download und die Nutzung der Skripte unterliegen der GNU Lizenz welche von Lexikos dem Gründer der Autohotkey Foundation erstellt wurden. 

**I.b.** Die Inhalte und Skripte dürfen frei verändert werden. Jegliche Änderung ist vor Weitergabe zukennzeichnen.

#### Download

**II.** mit dem Herunterladen oder Speichern der Dateien in jeglicher Form (entpackt, gepackt, kompiliert, unkompiliert, in Teilen, als Ganzes) übernehmen Sie die volle Verantwortung für deren Inhalt. Ich verlange keinen Nachweis wer die Skripte herunterlädt, plant diese herunter zuladen oder diese vielleicht nie nutzen möchte, oder sogar nutzen will oder wer Veränderungen daran vorgenommen hat oder vornehmen will. Desweiteren akzeptieren sie vollständig alle unter Punkt III stehenden Handlungs- und Haftungsregeln.<BR>

#### Haftungsausschluß / Haftungsfreistellung

**III.** die Nutzung des gesamten Inhaltes (einschließlich Programme, Skripte, Bilder, Bibliotheken, Fremdprogramme) erfolgt auf eigene Gefahr! Das Angebot wurde besten Gewissens auf mögliche Urheberrechtsverletzungen untersucht. Es dürften keine Verletzungen enthalten sein. Quellenangaben sind soweit es nachvollziehbar war in den jeweiligen Dateien enthalten.<BR>

**III.a.** ich übernehme keine Haftung durch vermeintliche, unwissentliche oder eventuell unterstellte, wissentliche Fehler in den Skripten. Programmieren ist Hobby, Leidenschaft und Arbeitserleichterung. <BR>

**III.b.** ebenso übernehme ich keine Haftung für Urheberrechtsverletzungen durch Dritte deren Software oder SourceCode hier verwendet wurde/wird und werden wird. <BR>

**III.c.** Sie akzeptieren das die Skripte nur in einer Alpha-Version vorliegen. Sie sehen den Inhalt als  Beispielprojekt! Die angebotenen Skripte sollen keinen Abrechnungsbetrug ermöglichen, sondern im Gegenteil nur zeigen, welche Möglichkeiten der Einsatz von RPA-Software in der Praxis bringen würde.<BR>

**III.d.** Eine Richtigkeit der in der in der Software hinterlegten Leistungsziffern kann nicht garantiert werden,es liegt in Ihrer Verantwortung alles zu kontrollieren<BR>

**III.e.** sämtliche Automatisierungsskripte können gegen Ihr lokales KV-Recht oder der KBV verstoßen - das Haftungsrisiko übernehmen sie! Die Skripte sind als Beispiele für mögliche Anwendungen durch Nutzung der freien Skriptsprache Autohotkey zu verstehen (siehe auch IIIc).<BR>

**III.f.** Sie haben kein Recht Updates, Fehlerkorrekturen oder eine Behebung von Folgeproblemen (vermeintlich oder beweisbar) einzufordern. Ebenso haben Sie kein Recht Schadensersatz für (vermeintliche oder beweisbare) Fehler zu fordern.<BR>

**III.g.** Insbesondere distanziere ich mich von jeglichen Versuchen meine Skripte für die Überwachung von Mitarbeitern oder anderen Personen einzusetzen.<BR>

**III.h.** Für Fehler in Rechtschreibung, Grammatik bin ich ebenso nicht verantwortlich zu machen. Dazu wenden Sie sich bitte an meine Deutschlehrer.<BR>

**III.i.** ich übernehme keinerlei HAFTUNG aufgrund hier fehlender rechtlicher Hinweise/Aus- oder Einschlüße. Ihnen sollte nach dem Lesen bekannt sein, daß ich keinerlei kommerzielle Zwecke verfolge und die Zusammenstellung der Dateien nicht zum Zwecke eigener Bekanntheit erfolgt und ich deshalb niemals wissentlich oder absichtlich fremdes geistiges Eigentums entwendet habe. Die angebotene Sammlung verfolgt ausschließlich einen gemein-nützigen Zweck.<BR>

<br>
[GNU Licence for Addendum für Albis](Docs/GNU Licence for Addendum für Albis.pdf)

<img src="Docs/TrennerExtraBreit.png" style="zoom: 67%;" />

<center> - IXIKO 2021 - </center>


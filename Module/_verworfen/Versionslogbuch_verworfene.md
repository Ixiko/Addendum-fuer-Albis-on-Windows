#   <u>Versionslogbuch</u> - verworfene Skripte


![](../../Docs/Trenner_klein.png)

## ![Praxomat.png](../../Docs/Icons/Praxomat.png) Praxomat

| Datum          | Teil     | Beschreibung                                                 |
| -------------- | -------- | ------------------------------------------------------------ |
| **08.04.2019** | **F~**   | nur Code übersichtlicher gemacht                             |
| **13.12.2018** | **F+**   | SkriptIcon ist in das Skript integriert                      |
| **04.10.2018** |          | **Praxomat Version 0.77**                                    |
| **04.10.2018** | **F+**   | HelpViewGui verbessert, Buttons anstatt Text zum direkten Ausführen von Befehlen gedacht |
|                | **F+**   | Praxomat zeigt in der Gui jetzt mehr Informationen           |
|                | **F+**   | das PraxomatGui kann sich jetzt automatisch an die richtige Stelle positionieren, es verschiebt sich auch mit dem Albisfenster |
| **30.06.2018** | **F~**   | diverse Fehlerbereinigungen, insbesondere bei der Erstellung der GVU Liste traten Formatfehler auf so daß die GVU-Liste durch das Modul GVU_Queue_ausfüllen nicht abgearbeitet werden konnte, verloren gegangene Funktionen für dieses Modul der Praxomat-Functions.ahk wieder hinzugefügt |
| **29.05.2018** | **P+**   | **Praxomat V0.70** - die GUI zeigt jetzt einen weichen Farbverlauf des Timers und nicht mehr der Hintergrundfarbe. Der Farbverlauf ist vorberechnet aus einem vorher erstellten Bild mit einem horizontalen Gradienten über mehrere Farben. Mit Hilfe eines Skriptes wurde dann eine Liste mit RGB Farbwerten (3000 Stück) erzeugt. Bei max. 25min eingestellter Gesprächszeit wird jeweils alle 500ms ein neuer Farbwert im Hintergrund des Fensters gesetzt. |
| **24.05.2018** | **P+~**  | **Praxomat V0.67** /                                         |
| **18.05.2018** | **P~**   | **diverse Fehlerkorrekturen** - das Pausieren des Timer funktioniert jetzt zuverlässig, Wartezeitdateien werden wieder eingelesen, GVU Liste wird wieder eingelesen - war ein ein Problem in der *GetQuartal*-Funktion, Parsing der Wartezeitanzeige 01:05 nicht 1:5 wie vorher |
| **17.05.2018** | **P~**   | **TenMinutes Gui:** die Ziffernauswahl arbeitet erstmals absolut zuverlässig und fehlerfrei |
| **14.05.2018** | **P~**   | **GVU Liste**: zugehörige Funktion GetQuartal() verbessert, *Codebereinigung* - unötiger Code entfernt, PraxomatIcon aus der Taskbar entfernt |
| **04.05.2018** | **P~**   | **TenMinutes Gui**: Fehler das manchmal die Gesprächszeit nicht berechnet wurde beseitigt |
| **04.04.2018** | **P~**   | **Praxomat V0.62**: das Ausführen der Funktion zum Erstellen der GVU Liste nutzte bisher das aktuelle Datum, dies führte allerdings dazu, das die Liste mit einem neuen Quartal fortgeführt wurde, auch wenn man gerne noch ein paar Formulare im vergangen Quartal unterbringen wollte. Jetzt wird das eingestellte Programmdatum dazu genutzt. Kleinere Fehlerkorrekturen in der GVU Listenfunktion durchgeführt |
| **02.03.2018** | **P+**   | die PraxomatGui ist nicht immer AlwaysOnTop, je nach Mouseposition wird dieser Wert- ein oder ausgeschaltet |
| **28.01.2018** | **P~**   | **diverse Fehlerkorrekturen** seit dem 23.01. , Problem gelöst das manchmal die Nachfrage ob ein Patient aus der Wartezimmerliste gelöscht werden zweimal erschien |
| **20.12.2017** | **P+**   | die Timer Gui ist jetzt vollständig eine GDI Gui, somit kein Flackern der Ausgabe mehr - damit läßt sich aber das Fenster nicht als child an Albis binden |
|                | **P#**   | **Praxomat Version 0.55**                                    |
| **19.12.2017** | **P~**   | das Unterbringen von Ziffern mit Faktor z.B. 03230(x:1) - wird nicht von Albis akzeptiert (bei der Vorbereitung zur Abrechnung hat er diese Eintragungen nicht akzeptiert), deshalb entfällt jetzt das (x:1) |
| **15.12.2017** | **P~**   | Praxomat: Steuerung + F11 und ENTER funktionierten trotzt diverser vermeindlicher Fehlerbehebungen gar nicht mehr, durch das Überprüfen welches Fenster vorn liegt um die Timer-Gui je nachdem sichtbar oder unsichtbar zu machen, kam es zu einem Flackern der Anzeige, diese Funktion ist wieder entfernt |
| **12.12.2017** | **P~**   | das GDI-Overlay Fenster wird versteckt wenn das Albisfenster nicht mehr vorne liegt, leider läßt sich das Overlay Fenster nicht als ChildWindow anbinden , es wird dann gar nicht mehr angezeigt |
| **07.12.2017** | **P+**   | gleich am Anfang wir die GVU-Liste eingelesen, damit die Anzahl der aufgenommenen Patienten im TopToolTip angezeigt werden kann, siehe 06.12. |
|                | **P+**   | Anzeige eines ToolTip oberhalb der Timer-Gui da das Aufnehmen der Behandlungspatienten immer noch nicht reibungslos funktionierte über die Praxomatinterne Funktion TopToolTip() |
| **06.12.2017** | **P~**   | die zigfachste Fehlerkorrektur bei der Aufnahme des Behandlungspatienten. goto Routinen gesetzt, gosub gelassen wo es hingehört. |
|                | **PF+**  | die Routine für die Aufnahme in die GVUListe überprüft auf schon vorhandene Eintragungen, im Anschluß wird im Timer-Gui die Gesamtzahl der Untersuchungen angezeigt |
| **05.12.2017** | **P+**   | die Hintergrundfarbe der Timer-Gui ändert sich nur wenn ein Patient in Behandlung gesetzt wurde |
| **03.12.2017** | **P+**   | Information zum Programmablauf werden jetzt in der letzten Zeile des Timer-Gui angezeigt (Infopraxo:), dort sollen dann auch Fehler oder andere wichtige Hinweise erscheinen, |
|                | **P#**   | Praxomat/Addendum erreicht **"Level"** 0.5                   |
| **02.12.2017** | **P+~**  | der obere Teil der Timer-Gui ist jetzt mit der GDI+ Library gestaltet, leider verschwindet die GDI Gui wenn man sie als child binden möchte, lt. Autohotkey-Forum ist die Transparenz eines |
| **01.12.2017** | **P-~**  | das Überprüfen ob die richtigen Einträge von z.B. lko 03230(x:2) in den Feldern (Edit3, RichEdit..) funktioniert nicht, AHK kann die Felder nicht auslesen, diese Funktion war zur Fehlerüberprüfung gedacht, produzierte aber nur Fehler (Praxomat_Functions.ahk). Die bisherigen Programmteile habe ich als Snippets gesichert und aus dem Funktionscode entfernt. |
|                | **P-**   | die ehemalige Errinnerungsfunktion für das Timer-Gui damit ich das Eintragen der Gesprächszeit nicht vergesse ist jetzt überflüssig geworden, die Funktion ist jetzt für eine manuelle |
|                | **P~**   | das Timer-Gui ist jetzt als child-Fenster an das Albis Fenster gebunden |
|                | **P~**   | bei Privatpatienten wird wieder die Gesprächszeit eingetragen, nachdem das Eintragen beendet ist wird die Akte geschlossen und nachgefragt ob der Patient aus der Wartezimmerliste |
| **30.11.2017** | **P+-**  | farbliche Kennzeichnung Praxomat Timeranzeige nach Zeitüberschreitung türkis, ---------> rot , rudimentär programmiert mit nicht so schönen Farben |
|                | **P+**   | nach drücken auf die *PAUSE*-Taste wird der Timer angehalten , das Hinweisfenster sieht sehr hübsch aus im Gegensatz zur farblichen Kennzeichnung der Timer-Gui (siehe oben) |
| **28.11.2017** | **P~**   | Fehlerkorrektur des Patientenaufruf, aufgerufener Patient wird registriert, dies kann mehr als 10sec dauern, es hat wohl etwas damit zu tun das die Albis Fenster Modus - Overlapped	eingestellt haben |
|                | **PF+*** | ein GUI ( TENMINUTESGUI() ) zur Auswahl der 3 möglichen Gesprächsziffern erstellt, anhand der Gesprächsdauer wird auch gleich der Faktor vorgeschlagen , dieser läßt sich ändern ebenso welche Ziffer gewählt wird |
|                | **P+**   | per Strg+ Win + linke Maus zu erreichende WindowInfo Funktion integriert (gibt viele Daten aus - aber funktioniert nicht so wie ich es gern hätte) - ist im Moment für mich für die weitere Programmierung notwendig |
|                | **PF+-** | Beginn der Integration eines DebugFensters auf Basis eines ListView |
| **26.11.2017** | **P+***  | beim Öffnen eines Patienten aus dem Wartezimmer heraus - wird der Timer neu gestartet - dieser Timer wird an den Patientennamen gebunden									schaltet man zurück auf das Wartezimmer pausiert der Timer bis man zurück zu ist (*im Moment noch in jedem Patientenfenster: das Auswählen eines neuen Patienten aus dem Wartezimmer	 bei noch laufenden Timer führt zur Nachfrage ob man die Timerdaten abspeichern möchte*) |
|                | **F-**   | *das Speichern der Timerdaten - soll bei Überschreitung jeder 10.min zum Einschreiben der Gesprächsziffer lko 03230 oder 35110 mit jeweiligem Faktor auffordern*) |
| **24.11.2017** | **P+**   | Patientenaufruf  aus dem Wartezimmer (nur Arzt) setzt nach Zustimmung den Patientenstatus auf -in Behandlung- -> dies funktioniert nur mit der ENTER-Taste |
| **23.11.2017** | **P+**   | Timer zeigt die Zahl der Wartenden Patienten und die maximale Wartezeit an |
| **20.11.2017** | **P#**   | Praxomat V0.4 - sag ich mal                                  |
| **16.11.2017** | **P+***  | Albisstart jetzt mit Autologin für die jeweiligen Rechner - Daten werden in der ini Datei hinterlegt |
| **14.11.2017** | **P+-**  | animiertes ToolTip ( --Bezier_MouseMove2() und TToolTip()-- ) funktioniert wenn ein Pat. in die GVU Liste eingetragen wurde - damit man weiß das er auch in der Liste gelandet ist 							(*Liste muss auf doppelte Einträge überprüft werden*) - erledigt am 06.12.2017 |
| **11.11.2017** | **P+-**  | wenn Albis abstürzt - wird es neu gestartet - da er Albis mehrfach startet - lässt sich nicht verhindern - wird mittels Messagebox zuvor nachgefragt |
|                | **M**+/- | es existiert ein Modul welches es ermöglicht, Quartalsprotokolle (aus Tagesprotokollen) zu erstellen, im Anschluß	werden alle unwichtigen Daten entfernt, dies ist der Grundbaustein für 								weitere Auswertungen, Erstellung von Statistiken. Hauptsächlich werde ich es wohl zur automatischen Erkennungen von fehlenden Ziffern verwenden. Listenvergleich AOK,BARMER |
|                | **M**-*  | Ziffern ein und legt ein neues oder ergänzt ein Errinnerungscode im cave Fenster Zeile 9 (z.B. GVU/HKS 03^17) - dies steht für Monat und Jahr geplant ist das diese Liste später automatisch abgearbeitet werden kann und das dann ohne jeden Eingriff, das Modul öffnet die Patientenakte, ändert das Datum in der Akte auf das in der GVU Liste hinterlegte Datum, füllt erst das GVU und dann das KVU Formular aus, schreibt die Ziffern in die Akte, ändert den Code in cave und schließt die Akte wieder |
|                | **M**+   | es besteht innerhalb des Praxomat ein Programm welches das GVU/KVU Formular fast ohne 	Eingriff - per Standard ausfüllt (das spart clicks und Zeit), das Modul schreibt die entsprechenden |
|                | **P**+   | Patienten lassen sich per Hotkey in eine externe Textdatei ablegen - GVU/KVU Liste |
|                | **P+**   | StrC + StrX + StrV funktionieren jetzt innerhalb von Albis - sonst wurde immer die Einleseroutine für die Kartenlesegeräte aufgrufen |

<br>

![](../../Docs/Trenner_klein.png)

## ![Dicom2Albis.png](../../Docs/Icons/Dicom2Albis.png) Dicom2Albis

|     Datum      | Teil    | Beschreibung                                                 |
| :------------: | ------- | ------------------------------------------------------------ |
| **13.12.2018** | **F+**  | SkriptIcon ist in das Skript integriert                      |
| **14.03.2018** | **M~**  | **DICOM2ALBIS V0.98**: Fehlerkorrekturen beim Aufruf des externen Programmes für die Umwandlung von Serienaufnahmen (MRT,CT) nach AVI oder WMI , bei mDicom.exe die beiden Fenster (Listview und das Logo) werden vor Start des externen Programmes beendet, sie stören sonst nur |
|                | **MF+** | **DICOM2ALBIS V1.00**: mit Hilfe der MicroDicom.exe lassen sich jetzt Serien von Bildern automatisiert in WMV umwandeln, der Filename wird aus den Metadaten der DICOMDIR Datei generiert, Dicom2Albis hat jetzt alle Funktionen die ich brauche (eine fehlt nur noch - gemixte DICOM CD's Röntgen und Schichtaufnahmen werden nicht eingelesen diese CD's sind selten - dazu muss ich noch eine Lösung finden |
| **06.03.2018** | **M+**  | bei nicht zu konvertierendem DICOM Inhalt ruft das Skript das kostenlose MicroDicom Tool auf (die Umwandlung in wmv oder avi muss noch manuell erfolgen), nach Fertigstellung der Umwandlung wird die CD automatisch ausgeworfen |
| **20.02.2018** | **M~**  | erstmals durchgehende Funktionsweise, alle Pfade nachdem Autostart des Skriptes (CD einlegen , wird als DICOM CD erkannt und das Dicom2Albis Skript wird gestartet) Fehlerbereinigung, Fehlerüberprüfung falls es Probleme beim Umwandeln gab und im Befundordner keine umgewandelte Datei erscheint, vorher hatte das Skript im 100ms Abstand auf den Ok Button des Befundübernahmefenster geklickt, jetzt wird das erscheinende Albisfenster mit Hinweis auf die nicht vorhandenen Dateien erkannt und es wird die nächste Datei ausprobiert |
| **16.02.2018** | **M~**  | die Parameterübergabe an das Skript korrigiert, notwendig für den automatischen Aufruf durch das Praxomat_st Skript nach dem dieses ein eingelegtes DICOM Medium erkannt hat |
| **01.02.2018** | **MF+** | die Größe des VerlaufsGui läßt sich jetzt verändern, die letzte Position und Größe wird in der Praxomat.ini gespeichert |
| **31.01.2018** | **MF+** | Skriptpause Button und Skriptabbruch Button der GUI hinzugefügt, bei vorzeitigem Abbruch fragt das Script ob bereits konvertierte Bilder gelöscht werden sollen |
| **28.01.2018** | **M+~** | Dicom2Albis, erste Versuche Dicom2Albis vollständig zu automatisieren haben noch nicht zum vollen Erfolg geführt, eine neu eingelegte CD wird erkannt, die Prüfung auf eine DICOM CD funktioniert, aber der Dicom2Albis Start schlägt fehl, zudem funktioniert Dicom2Albis aufeinmal nicht mehr, obwohl ich innerhalb dieses Skriptes keine Veränderungen vorgenommen habe |
| **14.01.2018** | **M~**  | Dicom2Albis - die Schrift wurde auf kleinen Bildern zu groß angezeigt, es wird jetzt automatisch die Schriftgröße anhand der Bildgröße angepasst, es besteht weiterhin der Fehler das beim ersten Start des 	Skriptes sogleich die Patientenakte aufgerufen wird ohne das versucht wird Bilder von der CD zu lesen |
| **18.12.2017** | **M~**  | Fehlerkorrekturen - Patientennamen enthalten oft folgendes Zeichen (^), manchmal sind diese Zeichen auch am Ende des Namens zu finden, bisher habe ich mit StringReplace dieses Zeichen zwischen Nachnamen und Vornamen ersetzt als Beispiel: Müller^Marie -> Müller, Marie. Letzteres ist die Schreibweise um sie Albis zu übergeben. Mit der Änderung werden vorgestellte und nachgestellte ^-Zeichen getrimmt. Zusätzlich habe ich die Animation Bezeichner-GUI verbessert (dieses Feature musste sein!). |
| **15.12.2017** | **M~**  | Dicom2Albis Fehlerkorrektur nach dem ersten richtigen Einsatz von Dicom2Albis, das Programm hat teilweise die Daten der zuvor erstellen DICOM.txt Datei benutzt, die Datei wurde scheinbar nicht überschrieben und wird jetzt jedes mal gelöscht |
| **13.12.2017** | **M+**  | komplette Funktionsfähigkeit!, Qualität der jpeg ist jetzt auf 90 eingestellt, das Umwandeln und Einsortieren muss nacheinander passieren, ein späteres Einsortieren geht leider nur manuell (im Anschluss hat der Vorführeffekt alles zunichte gemacht) |
| **10.12.2017** | **M+*** | **Dicom2Albis** - ein Modul um die Bilder einer Röntgen-CD automatisiert einzulesen - ist so gut wie fertig, das Modul wandelt mit Hilfe einer freien Software die sonst nicht lesbare DICOMDIR Datei in eine xml-Datei um dort wird dann nach den wichtigen Informationen gesucht<br />im Anschluss werden die Dateien mit Hilfe einer weiteren freien Software *dicom2.exe* in eine PNG-Datei umgewandelt. <br />Dicom2Albis schreibt Text in das Bild (DICOM Bilder - haben keinen Text in den Bilddaten und benötigen deshalb ein spezielles Anzeigeprogramm) <br />da mir PNG Dateien zuviel Platz wegnehmen werden diese Dateien noch in JPEG mit einer Qualität von 75 umgewandelt.<br />Dann erfolgt das automatisierte Einsortieren in die richtige Patientenakte. |
|                | **M-**  | -es fehlen noch Einstellungsmöglichkeiten für die Grafikausgabe (im Moment nur meine Vorstellungen), das Ausgabedateiformat ist fest vorgegeben	(JPEG)<br />-eine Installations- oder Einstellungsroutine für den automatischen Start nach Einlegen einer CD mit DICOM Inhalt muss noch erstellt werden |

<br>

![](../../Docs/Trenner_klein.png)

## ![GVU](../../Docs/Icons/GVU.png) Gesundheitsvorsorgeliste

| Datum          | Teil   | Beschreibung                                                 |
| -------------- | ------ | ------------------------------------------------------------ |
| **03.10.2019** | **F+** | der seit 01.07.2019 geltenden Abrechnungsbedingung für die 01732 angepasst |
| **25.06.2019** | **F~** | Code verkürzt, Code optimiert für mehr Geschwindigkeit, vermehrter Einsatz von RegEx-Befehle für flexiblere Stringvergleiche V1.88 |
| **24.06.2019** | **F+** | Hautkrebsscreening-Formular - alle Karzinom-Arten werden auf 'nein gesetzt |
| **29.03.2019** | **F+** | Fenster das der Patient heute Geburtstag hat wird automatisch geschlossen V1.86 |
| **28.03.2019** | **F~** | das Hautkrebsscreening-Formular wurde zu diesem Quartal verändert, Skript wurde entsprechend angepasst, Code wurde optimiert und fehlerfreies RPA! 100 Formulare ohne Unterbrechung! V1.85 |
| **07.01.2019** | **F~** | Zuverlässigkeit erhöht, Aufruf der Formulare erfolgt jetzt per Menubefehl nicht mehr per internem Albismakro, dafür erscheinen zwei Fenster weniger und das Erstellen der Dokumentation ist deutlich schneller |
|                | **F+** | das Skript nutzt jetzt auch einen Hook um störende Fenster zu entfernen und um erkennen zu können |
| **13.12.2018** | **F+** | SkriptIcon ist in das Skript integriert                      |
| **13.10.2018** | **F~** | durch Verbesserung der AlbisGetCaveZeile Funktion, enormen Geschwindig-keitszuwachs der Abarbeitung erreicht von >5s runter auf 0.9s, Abrechnungs-bedingung 'Mindestalter 35 Jahre' wird überprüft |
| **12.10.2018** | **F~** | Skript reorganisiert , Code übersichtlicher formatiert, Kommentierungen verbessert und/oder noch einige hinzugefügt, allgemeine Skriptbeschreibung verständlicher gestaltet |
| **08.10.2018** | **F~** | kleinere Fehlerbehebungen                                    |
|                | **F+** | wenn man seine Abrechnung noch nicht zum neuen Quartalsanfang nicht geschafft hat, dann gibt es manchmal ein Problem mit dem GNR Vorschlag, es muss dann das Abrechnungsquartal zuvor ausgewählt werden |
| **30.06.2018** | **F+** | merkt sich jetzt die sich aktuell in Abarbeitung befindende GVU-Liste, der INI-Eintrag wird erst nach vollständiger Abarbeitung wieder entfernt |
| **07.04.2018** | **M~** | das Dokumentieren des Untersuchungsdatums im Cave!von Fenster hatte nicht immer funktioniert, die Berechnung für den Abstand zwischen 2 Vorsorgeuntersuchungen war nicht immer korrekt |
| **06.04.2018** | **M~** | das Prüfen auf einen gültigen Abrechnungsschein zum Abrechnen der Vorsorgeformulare funktioniert jetzt |
| **07.01.2018** | **M~** | *GVU.txt Datei - Änderung im Prinzip als .csv Datei - Delimiter ist jetzt ein ";" , die Leerzeichen als Trennzeichen hatten dazu geführt das zuviele zu bearbeitende Variablen entstanden sind Änderungen wurden im Praxomat Modul_GVU_Queue Ausfüllhilfe und in Praxomat (Label GVUListe) durchgeführt |
| **06.01.2018** | **M~** | Praxomat Modul_GVU_Queue Ausfüllhilfe, Optimierung des Ablaufes<br /> Sicherung der geänderten "Cave! von" Fenster-Zeile 9 in die *GVU.txt Listendatei<br />Modul füllt die beiden Formulare Gesundheitsvorsorge und Hautkrebsscreening durch die verbesserten oder neuen Funktionen zuverlässiger aus , ohne wie zuvor manchmal nicht mehr weiterzukommen |
| **04.01.2018** | **M+** | Praxomat Modul_GVU_Queue Ausfüllhilfe, das zuvor die mit Praxomat im Quartal erstellten GVU-Listen abarbeitet hat seinen ersten Probelauf durch |

<br>

![](../../Docs/Trenner_klein.png)

## ![SonoCapture.png](../../Docs/Icons/SonoCapture.png) SonoCapture

| Datum | Teil | Beschreibung |
| ----- | ---- | ------------ |
| **13.12.2018** | **F+** | SkriptIcon ist in das Skript integriert     |
| **02.12.2018** | **F~**  | Unicode DllCall der avicap32.dll - warum die ANSI DllCalls auch mal funktionierten, ich weiß es nicht. So ist es allerdings stabiler lauffähig. |
| **02.08.2018** | **F~**  | RegRead über externe Funktion                                |
| **14.04.2018** | **M~**  | das Skript kann jetzt mit der Windows dpi scale Funktion umgehen. Die Fenster sollten somit bei allen Auflösungen und Skalierungseinstellungen korrekt aussehen |
| **13.04.2018** | **MF**+ | **Sono Capture V06.0**: Erinnerungsfenster integriert, das Ultraschallgerät muss unbedingt vor dem Zugriff auf den WinCap-Treiber eingeschaltet sein, sonst kommt es zu einem nicht behebbaren Treiberkonflikt, so daß sich das Skript nicht beenden läßt. Meist ist ein Neustart des Computer dann auch noch notwendig. |
| **14.03.2018** | **MF+** | **Sono Capture V0.53:** nicht konstant erkanntes Irfanview Fenster (schlecht gelöste Abfrage) korrigiert, ein Nutzer kann jetzt in der Praxomat.ini sein bevorzugtes Preview-Programm einstellen oder er schreibt "No" ein, wenn er bei Albis die Previewfunktion ausgestellt hat,	Auflösung in der GUI Preview erhöht, PAL-Auflösung (720x576 Pixel) |
| **23.02.2018** | **M+**  | **SonoCapture hinzugefügt**, eigenes Modul für die Aufnahme der Sonobilder mit sofortiger automatischer Eintragung mit korrektem Bezeichner in die Patientenakte. **ERSATZ**: für meine *unzuverlässige Praxisarchivierungs-software (CGM Praxisarchiv)*  - der Start dauert mir immer zu lange, läßt sich seit Monaten eh nicht starten, kann mir die Daten nicht mal mehr ansehen jetzt, der Videoinput über das Capturemodul der Software funktionierte mal und mal nicht. Zudem muss das Praxisarchiv jedesmal aktiviert werden wenn ein Rechner erneuert oder ausgetauscht wurde. |

<br>

![](../../Docs/Trenner_klein.png)

## ![ScanPool.png](../../Docs/Icons/ScanPool.png) ScanPool

|     Datum      |  Teil  | Beschreibung                                                 |
| :------------: | :----: | :----------------------------------------------------------- |
| **28.05.2019** | **F~** | Verbesserung beim Abfangen und Bearbeiten der Foxitdialoge, diese werden jetzt zu 100% richtig erkannt und abgehandelt |
| **18.05.2019** | **F+** | **ExtractNamesFromFileName()** - entfernt Dr. und Prof.      |
| **15.05.2019** | **F+** | Funktionen zur Behandlung des Speichern unter Systemdialoges in Windows10, **FoxitReader_ConfirmSaveAs()**, **FoxitReader_CloseSaveAs()**, **FoxitReader_ExceptionDialog()** |
|                | **F~** | Albis Pdf's	- automatisches Schließen der von Albis nach import geöffneten Foxitreaderfenster (V0.9.89) |
| **14.05.2019** | **F~** | Fuzzy Matching von Patientenfenster Albis und Pdf Dateinnamen funktioniert jetzt super! |
| **10.05.2019** | **F+** | PdfAddMetaDataInfo() -  (V0.9.88)                            |
| **02.05.2019** | **F~** | PdfReader_Close(), FoxitReader_SignaturSetzen(), diverse Fehler behoben, mehrere Dokumente aufeinmal lassen sich wieder übertragen und signieren |
| **12.04.2019** | **F~** | teilweise Reprogrammierung des Signierprozesse für eine deutlich verbesserte Zuverlässigkeit des Ablaufs |
|                | **F~** | Vorschau 		- Listview wird zu schmal dargestellt, erst ein manuelles Resize der Gui korrigiert es |
|                | **F~** | Foxitdialoge 	- Speichern unter - riesiges Problem - er hat dauernd die falschen Ordner offen oder nimmt sich irgendeinen anderen Dateinamen |
|                | **F~** | Rechtsklick nach einem Befund übertragen Vorgang funktioniert manchmal nicht mehr im Listviewfenster sondern nur in anderem Teil der Gui |
|                | **F+** | Befundordner ist eine Datenbank (V0.9.86)                    |
| **05.04.2019** | **F+** | Teile des Signiervorganges werden jetzt durch einen **Fensterhook** gesteuert. Dieser Methode ist deutlich zuverlässiger und wesentlich schneller! |
|                | **F+** | Der **Foxitreader** muss nicht mehr extra heruntergeladen und auch nicht installiert werden. Im \include befindet sich eine partable Version die sich von allen Clients aus starten läßt! (0.987) |
| **02.04.2019** | **F~** | Befund desselben Patienten sind jetzt farblich gleich hinterlegt in der Auflistung (0.986) |
| **01.04.2019** | **F~** | Fehlerbehebung in den Funktionen Einsortieren der Befunde, Öffnen einer Patientenakte, Anzeigen aller Befunde (0.985) |
| **22.01.2019** | **F~** | ContextMenu geht wieder sicher, Pdf Vorschau verschwindet wenn kein Befund ausgewählt wurde, Statuszeile optimiert |
| **15.01.2019** | **F~** | viele Fehler behoben                                         |
| **13.12.2018** | **F+** | SkriptIcon ist in das Skript integriert                      |
| **30.11.2018** | **F+** | das Vorschaubild läßt sich jetzt in 90° Schritten drehen     |
| **28.11.2018** | **F~** | sehr viele Fehlerkorrekturen                                 |
| **23.11.2018** | **F+** | *Fuzzy search* im Hauptfenster bei der Suche nach einer Datei |
| **22.11.2018** | **F~** | Korrektur der Fensterpositionierung vorgenommen, besseres berechnen, Hook zur Erkennung einer sich öffnenden Praxomat-Gui erstellt |
| **14.11.2018** | **F+** | fokusfreies Scrollen der Befund Gui ermöglicht               |
| **31.10.2018** | **F~** | Toleranz-Inkrement pro Suchvorgang (ein Loop)  zur FindText-Funktion hinzugefügt, so daß eine 100%-ige Übereinstimmung nicht immer notwendig ist |
| **28.09.2018** | **F+** | Pdf blättern mit dem Mausrad funktioniert super              |
| **03.09.2018** | **F+** | Annäherungsfunktion zur Namenssuche , Sift3 Funktion gleiche öffnen +0.5 - das war Teil1 |
| **25.08.2018** | **F+** | Erkennen das eine Datei signiert wurde und nicht mehr verändert werden kann |
| **25.08.2018** | **F~** | beim Signieren zur 1.Seite schalten                          |
| **23.08.2018** | **F~** | PdfIndex.txt nach jedem Löschvorgang neu erstellen           |
| **23.08.2018** | **F+** | das Passwort wird verschlüsselt in einer Variable gehalten   |
| **23.08.2018** | **F+** | PDF Preview!!                                                |
| **21.08.2018** | **F+** | ContextMenu direktes Einsortieren auch ohne signieren        |
| **12.08.2018** | **F+** | beim Öffnen einer Akte nimmt er immer an das diese noch nicht geöffnet wurde |
| **13.08.2018** | **F+** | Schließen des FoxitReader nachdem der Befund in die Akte einsortiert wurde und schließen der Patientenakte |
| **11.08.2018** | **F~** | Aktualisieren der Listview korrigiert                        |
| **09.08.2018** | **F+** | während des Signiervorganges ein Progress Fenster anzeigen und Interaktionen mit dem ScanPool Hauptfenster verhindern (Sperren des UserInputs, Abbruch per Esc - Hinweis geben!) |
| **02.08.2018** | **F+** | Edit1 löschen wenn alle Befunde Button gedrückt wurde        |
| **12.08.2018** | **F+** | öffnen aller Befunde eines Patienten                         |
| **04.08.2018** | **F+** | MsgBox Infobox oder Fehlerangabe sollte wie ein Childwindow den Zugriff auf die Gui sperren |
| **04.08.2018** | **F~** | vor dem Start von FoxitReader prüfen ob ein BEfund überhaupt noch existiert - sonst muss man ein leeres FoxitReader fenster schließen |
| **02.08.2018** | **F~** | indizierte und in die Akte einsortierte Dateien müssen auch aus dem files Array elöscht werden |
| **11.08.2018** | **F~** | Signiervorgang beschleunigt                                  |
| **09.08.2018** | **F+** | Statustext nach Einsortierung in die Akte auffrischen - Anzahl Dokumente und Seiten |
| **02.08.2018** | **F+** | Incrementelle Suchfunktion +0.5                              |
<br>

![](../../Docs/Trenner_klein.png)

## ![Monet](../../Docs/Icons/Monet.png) MoNet - (Mo)nitor for your (Net)work

| Datum          | Teil   | Beschreibung                                                 |
| -------------- | ------ | ------------------------------------------------------------ |
| **13.12.2018** | **F+** | SkriptIcon ist in das Skript integriert                      |
| **09.03.2018** | **M+** | **MoNet**: Monitor for your Network (nicht nur für Ärzte) - Anzeige des Status der Computer in der Praxis, auch als dezentrales Tool gedacht um von jedem Punkt im LAN aus alle Computer herunterzufahren und auch über WakeOnLan am nächsten morgen hochzufahren. Möglicherweise später auch Zeitgesteuertes hochfahren, weitere Funktionen wie LAN-Minichat, 	versenden von Einzel-Kommandos an die Computer oder aber auch ferngesteuertes Script oder Programm starten sind angedacht. |





#### Versionslogbuch - eine Beschreibung

- **F** 	Funktion

- **M**	Modul 

  

#### verwendete Begriffe

- **Hook** ([englisch](https://de.wikipedia.org/wiki/Englische_Sprache) für *Haken*, auch **Einschubmethode** genannt) bezeichnet in der [Programmierung](https://de.wikipedia.org/wiki/Programmierung) eine [Schnittstelle](https://de.wikipedia.org/wiki/Schnittstelle), mit der fremder [Programmcode](https://de.wikipedia.org/wiki/Programmcode) in eine bestehende Anwendung integriert werden kann, um diese zu erweitern, deren Ablauf zu verändern oder um bestimmte Ereignisse abzufangen. Dies kann entweder im [Quelltext](https://de.wikipedia.org/wiki/Quelltext) geschehen, der entsprechend modifiziert wird, über [Konfigurationsdateien](https://de.wikipedia.org/wiki/Konfiguration_(Computer)), die den Ablauf eines fertigen [Programms](https://de.wikipedia.org/wiki/Computerprogramm) verändern, oder über Aufruf von Funktionen, denen der auszuführende Programmcode in irgendeiner Form mitgegeben wird. In der Regel ist das Standardverhalten von Einschubmethoden, gar nichts zu tun.[[1/]](https://de.wikipedia.org/wiki/Hook_(Informatik)#cite_note-1)

  Hooks können auch vom Betriebssystem zum Abfangen von Nachrichten zur Verfügung gestellt werden. Damit lassen sich z. B. Fenster abfangen worauf hin z.B. Programmcode zum Schließen des Fensters (**Popup Blocker**) ausgeführt wird. (*Quelle: Wikipedia)

<br>
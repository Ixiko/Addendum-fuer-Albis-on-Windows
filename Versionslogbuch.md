#   <u>Versionslogbuch</u>

#### Versionslogbuch - eine Beschreibung

- **F+** 	neue Funktion die fehlerfrei läuft und der keine Subfunktionen fehlen

  **F+***	neue wichtige Funktion
  **F/f-**	wie oben , allerdings z.T. fehlerbehaftet (F) oder es fehlt noch eine Teilfunktion(f)
  **P/F-**	Programmcode oder Funktion entfernt, da entweder zu fehlerhaft / nicht mit AHK 					     durchführbar oder nicht mehr benötigt

- **P#**	neue Programmversion oder wichtiger Hinweis

  **P+**	neue wichtige Programmfunktion
  **M+**	neues Modul *für wichtiges Modul

- **MF+** neue Funktion innerhalb eines vorhandenen Moduls
  **F/M/P~**	Fehlerkorrekturen

#### verwendete Begriffe

- **Hook** ([englisch](https://de.wikipedia.org/wiki/Englische_Sprache) für *Haken*, auch **Einschubmethode** genannt) bezeichnet in der [Programmierung](https://de.wikipedia.org/wiki/Programmierung) eine [Schnittstelle](https://de.wikipedia.org/wiki/Schnittstelle), mit der fremder [Programmcode](https://de.wikipedia.org/wiki/Programmcode) in eine bestehende Anwendung integriert werden kann, um diese zu erweitern, deren Ablauf zu verändern oder um bestimmte Ereignisse abzufangen. Dies kann entweder im [Quelltext](https://de.wikipedia.org/wiki/Quelltext) geschehen, der entsprechend modifiziert wird, über [Konfigurationsdateien](https://de.wikipedia.org/wiki/Konfiguration_(Computer)), die den Ablauf eines fertigen [Programms](https://de.wikipedia.org/wiki/Computerprogramm) verändern, oder über Aufruf von Funktionen, denen der auszuführende Programmcode in irgendeiner Form mitgegeben wird. In der Regel ist das Standardverhalten von Einschubmethoden, gar nichts zu tun.[[1/]](https://de.wikipedia.org/wiki/Hook_(Informatik)#cite_note-1)

  Hooks können auch vom Betriebssystem zum Abfangen von Nachrichten zur Verfügung gestellt werden. Damit lassen sich z. B. Fenster abfangen worauf hin z.B. Programmcode zum Schließen des Fensters (**Popup Blocker**) ausgeführt wird. (*Quelle: Wikipedia)



<br>

![](Docs\Trenner.png)

## ![Addendum.png](Docs\Icons\Addendum.png) Addendum

| Datum          | Teil     | Beschreibung                                                 |
| -------------- | -------- | ------------------------------------------------------------ |
| **11.11.2017** | **MF+**  | **Popup Blocker für TeamViewer**                             |
| **26.11.2017** | **MF+**  | **Popup Blocker** für das Fenster: "**fehlende KVK**"  - > in das Praxomat-Überwachungsprogramm integriert |
| **10.12.2017** | **M+**   | **Hashtag-Übermittlung** (interne Patientennummer) mit einem kleinen Infotext um Patienten die bei **Telegram** angemeldet sind eindeutig identifizieren zu können<br />diese Funktion wird nur aufgerufen, wenn man in der Telegram-App im Chat das Wort Hash und anschliessend Enter, Space oder Tab drückt |
| **15.12.2017** | **F+**   | **Praxomat_st - Telegram-Desktop-App** - funktioneller-Hotkey - ein einfaches Enter erzeugt eine neue Zeile, bei einem schnellen Doppel-Enter wird Strg + Enter gesendet um die Nachricht abzusenden, diese Tastenkombination hat mich gestört, aber ich wollte dennoch die Formatierbarkeit einer Nachricht erhalten<br>**Praxomat_sT V0.90** |
| **22.12.2017** | **F+**   | neuer **Hotstring für Mail** oder sonstiges: 'Bestellungen': **Info`s über Möglichkeiten zu Rezeptbestellungen** |
|                | **F~**   | Fehlerbehebung innerhalb der HotKey's und Hotstrings         |
| **21.04.2018** | **MF+**  | **Menusuche** - Integration von Lexiko's hervorragendem Menusearch Skript - funktioniert auch mit anderen Programmen!<br>**Praxomat_sT V0.92** |
| **23.04.2018** | **F+**   | **TrayMenu geändert**, Praxomat_st hat ein **Infofenster** erhalten |
|                | **F+**   | **Errinnerungsfenster** - errinnern die Schwestern humorvoll an ihre einmal pro Woche durchzuführenden Tätigkeiten - **Sprechzimmer **mit Formularen/Arbeitsmaterialien bestücken, den **Hausbesuchskoffer** zu bestücken, beide Fenster erstellen aus einer Text-Liste, welche einfach zu editieren ist, eine formatierte Ausgabe auf dem Bildschirm. Die Liste enthält die Dinge mit denen Sprechzimmer oder Koffer zu bestücken sind. Der Aufruf der Funktion ist Client-abhängig gestaltet, je nachdem an welchem Arbeitsplatz die dazu beauftragte Schwester hauptsächlich arbeitet.<br>**Praxomat_sT V0.93** |
| **24.04.2018** | **F-**   | **AlbisHotKeyHilfe()** - Hotkeyhilfe in der Statusbar von Albis nach Praxomat-Functions.ahk verschoben |
| **18.05.2018** | **F-**   | **ToolTipAutomatic-Label** - entfernt, wird nicht mehr gebraucht |
| **20.05.2018** | **F+/~** | Registry auslesen klappt jetzt auch auf 64-bit System, nutze eine neu implementierte Funktion aus der Praxomat-Functions Bibliothek, an der TrayTip Ausgabe gearbeitet. Das Infofenster zeigt ein paar Informationen zum laufenden Skript an.<br>**Praxomat_sT V0.94** |
| **25.05.2018** | **F~**   | Hotkeys verändert, Ausführung verbessert                     |
| **31.05.2018** | **F+**   | **Laufzeitüberwachung für das Laborprogramm (AIS Connector)** - hinzugefügt <br>**Praxomat_sT V0.95** |
| **14.11.2018** | **F~**   | Shell-Hook für Parent-Fenster hinzugefügt (27.10.2018), die WinEventHooks erfassen Control- und Childwindow Events besser. Die beiden grundlegenden Hook-Funktionen sind jetzt clientunabhängig. Eventhook Routinen sind optimiert auf Geschwindigkeit und Erfassungsgenauigkeit (14.11.2018). |
| **08.12.2018** | **F+**   | IPC - Inter Process Communication - zwischen dem Addendum und Praxomat Skript steht, das Addendumskript kann somit weitere Funktionen übernehmen (V0.98b) |
|                | **F~**   | **Hooks** komplett überarbeitet, im Prinzip das komplette Skripte rund erneuert, Shell-Hooks werden nicht mehr benötigt, alles läuft über *WinEventHooks* |
| **13.12.2018** | **F+**   | SkriptIcon ist in das Skript integriert                      |
| **29.12.2018** | **F+**   | bei über Nacht laufenden Computern stellt Addendum das Albis Programmdatum einen Tag weiter und das Addendumskript wird neu geladen, dies verhindert Fehler im Skript durch Änderungen des Tagesdatums |
|                | **~**    | Versionierung geändert aus 0.98b wurde V0.98.2, kleinste Programmänderungen erhalten auch einen Zähler an der 4.Stelle - V0.98.21 |
|                | **F~**   | Funktion **NoTrayIcons()** war unter Windows10 wirkungslos, neueste Version ist jetzt integriert so daß bei Absturz oder Neustart eines Skript das Icon für ein nicht mehr laufendes Skript in der Traybar entfernt wird. |
| **05.01.2019** | **F+**   | **PopUp-Bearbeitung** für das Fenster "GNR der Anford.Ident" dem integrierten PopUpBlocker hinzugefügt (V0.98.22) |
| **06.01.2019** | **F+**   | neue Hotkey's zur Vereinfachung des Verschiebens von Einträgen im Dauermedikament und Dauerdiagnosen Fenster, drücke Steuerung und eine der beiden Pfeiltasten für hoch und runter zum Verschieben von Einträgen |
|                | **F+**   | **AutoDoc** - integriertes neues Modul zum schnelleren Erstellen von Einträgen im Dokumentenfenster. Das Makro-Modul in Albis hat sicherlich einen guten Ansatz, aber es benötigt mir zuviele Eingaben und Klicks bis es fertig ist. Mit AutoDoc setzt man eine neue Zeile z.B. durch Drücken der Einfügetaste, drückt danach + und es erscheint eine Auswahlbox für Deine Makro's (nutzt nicht die Albismakros!). |
| **08.01.2019** | **F+**   | weitere Fenster des Vorganges Labor abrufen sind jetzt per Hook automatisiert , *"Labor auswählen"* und *"Labordaten"* werden jetzt ohne Nutzereingriff automatisch ausgefüllt und bestätigt |
| **13.01.2019** | **F~**   | **Hooks** überarbeitet, deutlich schnellere und zuverlässigere Erfassung von Fenstern |
| **22.01.2019** | **M~**   | **Addendum.ahk** - ist ab heute nur noch mit Autohotkey_H zu starten, die anderen Skripte werden je nach Bedarf von Autohotkey_L nach Autohotkey_H umgeschrieben |
|                | **F+**   | Hook auf das Laborabrufprogramm (MCSVianova WebClient) gesetzt |
| **28.01.2019** | **F+**   | den Ablauf des Laborabrufes fast vollständig vom Skript Laborabrufen ins Addendumskript integriert |
| **08.02.2019** | **F~**   | den Shift+F3 Kalenderersatz vervollständigt                  |
| **17.02.2019** | **M+**   | **Schulzettel Modul**                                        |
| **20.02.2019** | **F+**   | **Start aller Module über das TrayIcon des Addendum-Skriptes möglich**, über Einstellungen in der Addendum.ini kann man festlegen welche Skripte auf welchem Client ausgeführt werden dürfen V 0.98.33 |
| **20.02.2019** | **F~**   | Fehler in der neuen Datumsfunktion +1 Woche +1 Monat         |
| **21.03.2019** | **F~**   | einige Fehler beim Shift+F3 Kalender behoben, hat auch in den Diagnosenfeldern den Kalender geöffnet, jetzt auch Anzeige im Datumsfeld der Patientenakte |
| **13.04.2019** | **F+**   | **Hotstringaufrufe** können jetzt an *bestimmte Bedingungen* speziell für die Handhabung in Albis angepasst werden. Die Funktion **AlbisGetActiveControl()** kann unter anderem Erkennen in welchem Fensterbereich sich der Texteingabe-Cursor befindet. So ist es jetzt möglich Hotstrings nur für die Ersetzung z.B. beim Kürzel 'lko' in der Karteikarte zu erstellen. Diese Kürzel sind dann z.B. nicht nach Eingabe von 'dia' erreichbar und umgekehrt. V 0.98.34 |
| **15.04.2019** | **F~**   | die **Modulstarter Funktion** (erreichbar über das Traymenu) kontrolliert ob ein Modul schon läuft und gibt einen Hinweis dazu aus. Gleichzeitig läßt sich über dieses Menu *ein laufendes Modul beenden*, insbesondere dann wenn es abgestürzt ist. |
| **01.05.2019** | **F~**   | Fehlerprotokollierung hatte selbst einige Fehler so daß nichts gespeichert werden konnte, eine zusätzliche Protokollierungsfunktion für OnExit - Gründe hinzugefügt |
| **01.05.2019** | **F~**   | das RestartScript in AddendumReload.ahk umbenannt            |
| **01.05.2019** | **F+**   | Addendum Starts erfolgen jetzt über das AddendumStarter.ahk Skript welches gleichzeitig bei jedem Start überprüft dqs alle wichtigen Dateien und Einstellungen vorhanden sind.  (V0.98.35) |
| **11.05.2019** | **F+**   | HotStrings hinzugefügt                                       |
| **13.05.2019** | **F~**   | HotString Fehler in der Erkennung des Eingabefocus bei AutoDoc() beseitigt |
| **01.06.2019** | **M+**   | der Laborabruf ist nach Aufruf des Menu "Labor abrufen" vom Skript zunehmend vollständiger automatisiert (V0.98.45) |
| **25.06.2019** |          | **V0.98.55** - Verbesserungen in den Addendum Funktionen (AddendumFunctions.ahk) und den anderen Skripten |
| **27.06.2019** | **M+**   | Abrechnungshelfer.ahk - kann deutlich mehr jetzt (V0.98.55)  |
| **14.08.2019** | **F+**   | **Hotstring-Funktion *AutoDiagnose()*** erweitert. <br>- _Ausgabe von Diagnosenketten_: z.B. Diabetes mellitus mit multiplen Komplikationen, G. {E14.72G}; Diabetische Polyneuropathie {\*G63.2G}; diabetische Nephropathie {\*N08.3G}; <br>- _Abfrage von und Handling der Seitenangaben_ in Diagnosen (funktioniert auch  mit den Diagnoseketten) |
| **23.08.2019** | **F~**   | Verbesserungen und Fehlerbeseitigung der Funktion AutoDiagnose |
| **23.08.2019** | **F+**   | Drucken des letzten aktuellen Laborparameter mit einer Hotkey-Kombination integriert  (V0.98.58) |
| **27.08.2019** | **M~**   | Auslagerung der Auto-ICD-Diagnosen in eine eigene Datei für bessere Editierbarkeit und angedachte Modularität |
| **01.09.2019** | **F~**   | **Auto-ICD-Diagnose** - verbessert, Fehler beseitigt, Diagnoseliste besteht derzeit aus circa 170 häufigen allgemeinärztlichen Diagnose und circa 250 Abkürzungen/Begriffen. <br>**Funktionsbibliotheken** - sind jetzt klar getrennt nach Addendum eigenen (\include) und allgemeinen (\lib) Bibliotheken<br>**Abrechnungshelfer** - neues Makro für die Auswertung aller Dauerdiagnosen in einem Tagesprotokoll (war hilfreich für die Auto-ICD-Diagnose Funktion) |
| **05.09.2019** | **M+**   | Vereinfachung der Übernahme von neuen ICD-Diagnosen in die Hotstringliste mit Gui für optionale Angaben, Synonyme, Ausgabe als formatierter Code in Scite |
| **10.09.2018** | **F~+**  | **Shift-F3 Kalender** - mehr Datumfelder werden erkannt, der Kalender wird nur noch bei Datumsfeldern angezeigt. Zuvor bestand ein Problem in der zuverlässigen Erkennung |
| **11.09.2018** | **F+**   | **Rezepthelfer** - erstelle Hilfsmittelrezepte ohne eine Liste öffnen zu müssen, tippe in eine Zeile im Rezept z.B. ATStrumpf oder Strumpf oder Strümpfe oder ... und es wird Dir das Rezept mit allen Zeilen befüllt. (**V0.98.61**) |
| **18.09.2019** | **F+**   | **Tray-Menu erweitert:** - in der Addendum.ini können jetzt Programme eingetragen, welche über das Tray-Icon in der Taskbar startbar gemacht werden |
| **25.09.2019** | **F+**   | **Auto-ICD-Diagnosen Hotstrings** - ein Fenster zum Anlegen eigener Hotstrings für die ICD Textexpandierung hinzugefügt (siehe Screenshots). (V0.98.62) |
| **27.09.2019** | **F~**   | Optimierung und Fehlerbeseitigung in den WinEventhooks, **AutoICD()** Funktion leicht verbessert |
| **02.10.2019** | **M+~**  | **Abrechnungshelfer-Modul**: ein großer Schritt zu mehr Funktionalität und umfangreicherer automatisierten Übertragung der Ergebnisse nach Albis (V0.98.65) |
| **23.11.2019** | **F*~**  | **Rezepthelfer** - Aufruf gespeicherter Rezeptvorlagen per Textkürzel wieder entfernt (kann mir diese nicht merken), dafür wird jetzt ein Auswahlfeld in jedes Rezept eingeblendet |
| **02.01.2020** | **F+** | das "Dauerdiagnosen von" Fenster wird neu gezeichnet, die Dauermedikamente wird verbreitert und das Werbeelement wird entfernt |
| **07.01.2020** | **F~** | es wird verhindert das Addendum doppelt gestartet werden kann |
|  | **F~** | die kontextsensitive Hotstring Erkennung in der Patientenakte funktioniert wieder  |
|  | **F~** | **Addendum_Protocol()** - das senden einer Nachricht mit Telegram funktioniert, die fehlerauslösende Zeile wird auch gesendet und im Protokoll auf Festplatte gespeichert |
| **14.01.2020** | **F~** | **AlbisFocusEventProc()** - Fehler in der Steuerelementerkennung beseitigt |
| **19.01.2020** | **M+** | **Addendum Toolbar** - Addendum integriert eine zusätzliche Toolbar in Albis. Über diese lassen sich Skripte starten und Einstellungen vornehmen. Welche Skripte gestartet werden können hängt von den vorgenommenen Einstellungen in der Addendum.ini ab. |
| **24.01.2020** | **F+** | **AlbisResizeDauerdiagnosen()** - Addendum-Albis.ahk: kleine Funktion ausschließlich zum Verändern der Größe des Dauerdialogfensters |
| **30.01.2020** | **M+** | **Befund und Bildimport** - sind jetzt direkt über ein in Albis integrierten Dialog (wird bei den Stammdaten/DauerDiagnosen/Dauermedikamente erstellt) möglich. Dazu wird der Scanordner (alias "Befundordner") auf pdf und jpg Dateien untersucht. Der Imprt erfolgt direkt in die Patientenakte zum Erstellungsdatum der Befund/Bilddatei. Die Funktion ist nicht in Verbindung mit dem Praxisarchiv gedacht. V1.25 |
| **31.01.2020** | **F+** | einzelne Skriptfunktionen lassen sich jetzt über das Traymenu wahlweise zu- oder abschalten, derzeit: Albis und Microsoft Word automatisch auf dem Monitor positionieren (nur für 4k Monitore), Automatisierung für die Abrechnung der Gesundheitsvorsorgen (Überprüfung auf Abrechenbarkeit, Formulare, Ziffern...), automatisches Signieren einer PDF mit Hilfe des FoxitReader V1.26|
| **11.02.2020** | **F~** | Tray-Menu Funktion zum starten von Addendum-Modulen und starten anderer Skripte funktioniert nach blödsinnigem Programmierfehler wieder |
| **11.02.2020** | **F+** | mehr zu- und abschaltbare Einstellungen, Vorbereitungen für einen zeitgesteuerten Laborabruf welcher pathologische Blutwerte per Telegram-Messenger verschicken kann |
| **12.02.2020** | **F~** | massive Nachbesserung an der PDF- und Bildimportfunktion, Dateien die nicht zugeordnet werden können, zieht man jetzt einfach in das Posteingangsfenster und prompt erscheinen diese dort. Danach lassen diese sich schnell in die Akte importieren.  |
| **12.02.2020** | **F+** | nutzt man den FoxitReader lässt sich mit einer einzigen Taste der komplette Vorgang zum Signieren eines Dokumentes starten. Was vorher nicht zuverlässig mit dem ScanPool-Skript funktionierte, läuft jetzt so schnell, das ich den Vorgang optisch ausgebremst halte. Du drückst Deine Wunschtaste im aktiviertem Foxit-Fenster, der Mauszeiger bewegt sich zum Dokumentbereich, zieht dort den Bereich für Dein voreingestelltes PDF-Sign auf, alle folgenden Dialoge samt Checkboxen, Buttons werden gedrückt, die Datei wird mit allen Sicherheitsnachfragen gespeichert und die signierte Datei wird danach sofort geschlossen. Existiert noch ein geöffnetes FoxitReader Fenster wird dieses nach vorne geholt und Du kannst gleich weiterlesen! Zeit für den gesamten Vorgang max. 5s (ausgebremst) - notwendige Eingriffe ein Druck auf eine Taste! That's RPA! ..... und falls Du das nicht brauchst schaltest Du es einfach über das Tray-Menu oder die Addendum.ini aus!  (V1.28) |
| **12.02.2020** | **F~** | Signieren beschleunigt, RPA Abbrüche beseitigt  |
| **31.03.2020** | **F~** | Labelfehler bei Ausführung auf einem unbekanntem Client beseitigt, Addendum_Reload/Addendum/Addendum_Monitor/Addendum_Starter - für ausschließlichen Ausführung des Addendum.ahk Skriptes mit AutohotkeyH eingestellt  |
| **01.04.2020** | **F+** | Addendum_Protocol: ein hinterneinander auftretender Fehlercode wird nicht gesendet und auch nicht protokolliert, damit wie es bei mir vorgekommen ist, nicht über Nacht alle 3-5Sekunden derselbe Fehler gemeldet wird |
| **04.04.2020** | **F~** | Importvorgang für PDF Datei deutlich (um das circa 5fache) beschleunigt, Erkennung der Dialoge verbessert |
| **04.04.2020** | **F+** | signierter FoxitTab wird je nach Einstellung automatisch geschlossen |
| **05.04.2020** | **F+** | der Dateipfad signierter PDF-Dateien wenn diese nach Albis importiert wurden wird zusammen mit der PatientenID in eine Textdatei gespeichert. Dies ist für ein schnelleres Wiederauffinden/Zusammenstellen alle Patientenbefunde gedacht. |
| **05.04.2020** | **F~** | Addendum_Starter.ahk verbessert: AddendumMonitor.ahk wird nicht mehr doppelt gestartet, wenn es im Hintergrund läuft, falls das Addendum-Skript noch läuft wird gefragt ob es neu gestartet werden soll  |
| **07.04.2020** | **F~** | öffnen einer Patientenakte aus dem Windows Dateiexplorer heraus, mit Linksklick eine mit einem Patientennamen versehene PDF Datei auswählen und per F6 wird die zugehörige Akte geöffnet, dies erleichtert den Import von Befunden, nicht erkannte Personennamen werden über den Aufruf der ersten 2 Buchstaben des Zu- und Vornamen über den Patient öffnen Dialog in Albis gesucht und bei passenden Namen zum Öffnen angeboten  V1.29 |

![](D:/Eigene Dateien/Eigene Dokumente/AutoIt Scripte/GitHub/Addendum-fuer-Albis-on-Windows/Docs/Trenner_klein.png)

## ![AddendumFunctions](Docs\Icons\AddendumFunctions.png) Addendum Funktionsbibliotheken (..\Include)

##### .                                                       *Addendum_**.ahk*

Addendum_Albis.ahk

|     Datum      |  Teil   | Beschreibung                                                 |
| :------------: | :-----: | :----------------------------------------------------------- |
| **30.11.2017** | **F+**  | Funktionen für das direkte Einschreiben von Ziffern in die Akte hinterlegt <br >**AlbisPrepareInput(Name)**<br>**AlbisSendInputLT(kk, inhalt, kk_Ausnahme, kk_voll)**<br>**Versicherungsstatus(PatData)** |
| **11.12.2017** | **F~**  | Fehler in der Funktion **GetCurrentAlbisPatID()** - PatName:= SubStr(WT, PosA + 1, PosB - [->PosA<-] - 1) - das PosA in den [] Klammern fehlte, deshalb konnte er beim Vergleichen des aktuellen Patienten den Speichervorgang nicht auslösen |
| **19.12.2017** | **F~**  | die **Helper_Gui** aus dem Praxomat Skript ist jetzt als Funktion ausgelagert, dies war keine wirkliche Fehlerbereinigung |
| **22.12.2017** | **F~**  | das Auswählen des 'Abbrechen' Buttons führte dazu das die TenMinutes-Gui Routine in der Warteschleife verharrte , hat immer auf einen lko Wert <> 0 gewartet, gelöst indem lko bei Abbruch auf 1 gesetzt wurde und im Praxomat Programm zusätzlich abgefragt wird das lko nicht 1 ist, zudem wird jedem Patienten der unter die 10 Minuten kommt ein lko 	apk (Arzt-Patienten-Kontakt) in die Akte eingeschrieben |
|                | **F~**  | bei Privatpatienten wurde die Gesprächszeit nicht in die Akte geschrieben, korrigiert |
|                | **F+**  | Funktion für Startkontrolle und Loginautomation für Albis in die Functions aufgenommen, dafür wurde diese Funktion aus Dicom2Albis entfernt. Jetzt genügt ein 3zeiliger Code für die Überprüfung ob Albis läuft und um Albis zu starten und sich einzuloggen |
| **06.01.2018** | **F+**  | zwei neue Funktionen für bessere Interaktionen mit Controls erstellt:     **SureControlClick(CName, WinTitle, WinText="")** <br />**SureControlCheck(CName, WinTitle, WinText="")**<br>die Funktionen prüfen ob der Befehl an das Control wirklich ausgeführt wurde. |
|                | **F~+** | **WaitAndActivate (WinTitle, Debug=1, DbgWhwnd=0)** verbessert und eine extra Funktion ausschließlich für Albis hinzugefügt **AlbisWaitAndActivate(WinTitle, Debug=1, DbgWhwnd=0)** das Debugging ist eigentlich keines, sondern nur eine Listview Ausgabe |
|                | **F+**  | **AlbisGetCaveZeile(nr)** - auslesen einer Zeile aus einem Cave! von Fenster - mit Versuch der Erkennung falls der Zugriff fehl schlägt |
|                | **F+**  | **AlbisSucheInAkte (inhalt, Richtung="backward", Vorkommen="")** - brauche ich erstmal um herauszufinden wann die letzte GVU wirklich abgerechnet wurde, die Funktion ist offen gestaltet, so daß sie für andere Zwecke ebenso funktionsfähig ist - (Funktion ist sehr unzuverlässig aufgrund der vielen überlagerten zumeist verstecken Fenster im Bereich der Akte) |
|                | **F+**  | **AlbisPatientenKarteikarte()**, stellt die Karteikarte per Strg+Shift+F9 bereit |
| **23.01.2018** | **F+~** | **AlbisInBehandlungSetzen()** ; was vorher nach If then Abfrage kam ist jetzt in eine Funktion ausgelagert. Dabei einen Fehler beseitigt der das Aufrufen verhindern konnte und wartet jetzt so lange bis der gewählte Patient verfügbar ist |
| **28.01.2018** | **F+**  | **WinForms_GetClassNN(WinID, fromElement, ElementName**) - ermittelt das ClassNN eines Elementes (z.B. Button) in einem Fenster wenn es zum Microsoft WinForms Framework gehört |
| **15.03.2018** | **F+**  | **AlbisInBehandlungSetzen()**	- es wird die Raumnummer des jeweiligen Sprechzimmer in dem die Funktion aufgerufen wird eingetragen, es wird die Zimmer Nr. hinzugefügt oder bei Wechsel eines Raumes ersetzt - im Moment nur Zahlenersetzung |
| **04.04.2018** | **F+**  | **AlbisLeseProgrammDatum()** - Auslesen des eingestellten Programmdatums. Praxomat_Functions enthält jetzt **31 Funktionen** mit direktem Zugriff auf Albis |
| **06.04.2018** | **F+**  | **FindWindow(WinTitle,WinClass:="", WinText:="", DetectHiddenWins = off, DectectHiddenTexts = off)** - neue Funktion für die Suche nach einem ganz bestimmten Fenster - gibt die ID zurück |
|                | **F+-** | **AlbisAbrechnungsscheinVorhanden(Quartal)** - im Albisfenster findet sich eine ComboBox , diese gehört zur Patientenfenster Toolbar, dort läßt sich per ControlGet - die Information auslesen, ob ein gültiger Abrechnungsschein angelegt wurde. Dies verhindert das Anlegen der Vorsorgeformulare falls ein Patient versehentlich in die GVU Liste aufgenommen wurde. |
| **11.04.2018** | **F+**  | **SprechzimmerAuffuellen(nr,AddendumDir)** eine GUI <br>die Schwestern vergessen manchmal die Sprechzimmer mit Formularen und Arbeitsmaterialien zu befüllen, diese Funktion erstellt ein GUI als Errinnerungsfunktion am Anmeldungsrechner, gesteuert wird dies über das Praxomat_st Skript , es wird an einem Tag der Woche angezeigt. <br>In der Textdatei (Die Liste.txt), welche im Praxomat_st Skriptverzeichnis liegen sollte, kann man die benötigten Materialien eintragen. <br>Die Funktion kann je nach ausgewählter Schriftart eine mehrspaltige Liste anzeigen. |
| **13.04.2018** | **F~**  | **DriveData(Drv)**: Hook zur automatischen Erkennung einer neu angelegten DICOM CD hat sich auch bei neu eingebunden Festplatten (z.B. Netzwerklaufwerke) aktiviert. Startet jetzt nur noch bei folgenden Filesystemen: CDFS und UDF |
| **14.04.2018** | **F+**  | **LDT2CSV(file, raw=0)**: LDT Parser - splittet die Daten auf und erzeugt eine csv Datei z.B. für Excelauswertungen oder anderes<br>raw = 0 - es werden die Feldkennungen in Klartext übersetzt<br>raw = 1 - er trennt die Zeilen nur auf und läßt die Zifferncodes erhalten<br>*bekanntes Problem*:<br>LDT-Dateien werden in **ISO 8859-15 Latin 9** kodiert, bisher habe ich noch keine Möglichkeit für eine Konvertierung nach UTF-8 gefunden |
| **14.04.2018** | **F~**  | **Zimmerauffuellen()** und **Kofferauffuellen()** - die Errinnerungsfenster für die Schwestern |
| **15.04.2018** | **F+**  | den Funktionen **SureControlClick()** und **SureControlCheck()** kann jetzt auch nur die ID des Fenster übergeben werden |
| **21.04.2018** | **F+**  | **MenuGetAll()** und **MenuGetAll_sub()** aus Lexiko's hervorragendem Menu-Search.ahk Skript |
| **22.04.2018** | **F~**  | **Quell-Code der Library übersichtlicher gestaltet**, faltbare nummerierte Unterteilungen, Funktionen sind jeweils einmal aufgeführt damit man nicht lange suchen muss |
| **23.04.2018** | **MF+** | **AlbisCaveVonToolTip()** - Overlay-Gui für das Cave! von Fenster - Errinnerung daran das noch keine Impfungen eingetragen wurden |
| **23.04.2018** | **F+**  | TrayMenu geändert, Addendum hat ein Infofenster erhalten     |
| **24.04.2018** | **F+**  | **Zimmerauffuellen()** und **Kofferauffuellen()** - die Errinnerungsfenster für die Schwestern |
|                | **F+**  | **AlbisHotKeyHilfe()** - Hotkeyanzeige für die Statusbar des Albisfenster, ursprünglich Bestandteil des Praxomat_st Skriptes, jetzt eine ausgelagerte Funktion, verbesserte Formatierung |
| **28.04.2018** | **F~**  | **ExtractFromString-function** hat eine bessere Beschreibung bekommen |
| **03.05.2018** | **F~+** | **AlbisGetCaveZeile()**, **AlbisAkteOeffnen()** Steuerung von Albis verbessert , es wird kein Tastenkürzel mehr an Albis gesendet, Umstellung auf SendMessage-Befehle, beide Funktionen überprüfen jetzt ob die Ausführung ihrer gesendeten Befehle an Albis erfolgreich war |
| **06.05.2018** | **F+**  | **TabCtrl_GetCurSel(), TabCtrl_GetItemText()** - die erste Funktion gibt den 1-basierten Index der aktuell ausgewählten Registerkarte zurück, die zweite gibt den Registerkarten-Namen zurücke |
| **09.05.2018** | **F+**  | **AlbisOptPatientenfenster(nr)** öffnet das Optionsfenster - Patientenfenster - für das Labormodul Laborabruf eingestellt, nicht nur zum Aufrufen dieses Fensters, die Funktion gibt das Handle |
|                | **F+**  | **StarteAlbis(....., Auto=0)** zweiten Startmodus hinzugefügt, *(Auto=0)* Start mit Rückfrage - wie bisher, [Auto=1] Start ohne Rückfrage und ohne Anzeige von ToolTips oder Hinweisen |
| **10.05.2018** | **F+**  | **FindChildWindow(ParentObject, ChildWinTitle, DetectHiddenWindow="On"),[EnumChildWindow()]** - diese Funktion findet alle ChildWindows eines ParentWindow, 										/ |
|                | **F+**  | **WinGetMinMaxState(hwnd)** - Funktion gibt den Zustand eines Fensters zurück, z steht für "zoomed" oder maximiert, i steht für iconic oder minimiert, die Funktion kann auch mit MDI_ChildWindows umgehen, vorausgesetzt man hat das entsprechende Handle des Fenster, dieses kann mit **FindChildWindow()** erhalten werden |
|                | **F+~** | **StarteAlbis(CompName, User, Pass, CallingProcess, AddendumDir, Auto=0)** - Funktion verbessert. <br>1. **AutoStartfunktion** integriert - **ALBIS** läßt sich mit der Einstellung (**Auto=1**) ohne weitere Nachfrage starten<br>2. **volle Automation** - Funktion kann gestartet werden, selbst wenn ALBIS schon gestartet ist, erkennt selbstständig z.B. das noch kein Loginvorgang stattgefunden hat und trägt selbstständig den Nutzer und das Passwort (wie in der Praxomat.ini hinterlegt) ein und öffnet dann das *Wartezimmer*-Fenster. Muss nichts mehr durchgeführt werden, findet auch keine Interaktion mit Albis statt. <br>3. die **Maximierung** des *Wartezimmer*-Fenster funktioniert jetzt durch Ermitteln der Fensterhandles, dazu wurde die **FindChildWindow()** integriert. |
| **11.05.2018** | **F~**  | **StarteAlbis()** - Korrektur eines Fehlers bei der Erkennung das das Wartezimmer-Fenster noch nicht geöffnet ist, Implementation zweier Funktioen (siehe nächste) zur Erkennung blockierender 'child' Fenster die ein Arbeiten mit einem schon geöffneten Albis behindern würden |
|                | **F+**  | **AlbisIsBlocked()** - stellt fest ob das Albis Fenster deaktiviert ist (durch ein child Fenster blockiert) - man kann dieser Funktion auch andere Fensterhandle übergeben |
|                | **F+**  | **AlbisCloseLastActivePopups()** - schließt alle blockierenden PopUp-Fenster (Formulare, Menufenster) die sonst interagierende Befehle eines Skript blockieren würden. |
| **13.05.2018** | **F+**  | **TimeCode(MD)** - kleine Funktion die bei MD=1 Datum und Zeit(inkl. Millisekunden) als String zurückgibt, für Protokollfunktionen z.B. wann ein Fehler aufgetreten ist |
|                | **F+~** | **StarteAlbis()** - Run Albis,, **+UseErrorLevel**, + AlbisPID* zum Ermitteln das der Startvorgang erfolgreich lief. <br>*Fehlerprotokoll* spendiert, absofort werden alle Fehlerausgaben in den Ordner logs'n'data/Errologs in der Datei Errorbox.txt im Ordner protokolliert und im speziellen Fall dieser Funktion wird zusätzlich ein Screenshot dort abgelegt. |
|                | **F+**  | AlbisGetCaveZeile(), AlbisAkteOeffnen() Steuerung von Albis verbessert , es wird kein Tastenkürzel mehr an Albis gesendet, Umstellung auf SendMessage-Befehle beide Funktionen überprüfen jetzt ob die Ausführung ihrer gesendeten Befehle an Albis erfolgreich waren |
| **14.05.2018** | **F~**  | **GetQuartal()** - unerklärlicherweise hat diese Funktion, obwohl völlig fehlerhaft Monatelang die richtigen Ergebnisse ausgespuckt -> heute erstmals nicht, der gleich einen neue Funktionsübergabe spendiert "heute", damit muss nicht bei jedem Aufruf ein Datum generiert werden, wenn es nur um das aktuelle Quartal geht |
| **15.05.2018** | **F-**  | **Modulstarter()** - unbenutzter und auch nicht funktionierter Code entfernt |
| **16.05.2018** | **F+**  | **Json2Obj()** - neue interne Funktion für die neue Funktion **AlbisMenu()** |
| **18.05.2018** | **F~**  | **PIC-GUI** - kann jetzt ein Fenstername übergeben werden, die Funktion selbst gibt nur noch ein Handle zurück, im Prinzip kann diese Funktion für die Anzeige von Bildern genutzt werden aufgrund der Gestaltung kann bei optionaler Wahl eines anderen GUI-Titel die Funktion für eine vielfache und gleichzeitige Anzeige von GUI-Fenstern gentutzt werden |
|                | **F-**  | **ToolTipAutomatic-Label** - entfernt, wird nicht mehr gebraucht |
| **20.05.2018** | **F+**  | an der **TrayTip Ausgabe** gearbeitet. Das Infofenster zeigt ein paar mehr Informationen zum laufenden Skript an. *Addendum V0.95* |
|                | **F+**  | **RegReadUniCode64()** - Registrierungsaufrufe von 32-Bit-Anwendungen, die auf 64-Bit-Maschinen ausgeführt werden, werden normalerweise vom System abgefangen und von HKLM / SOFTWARE an HKLM / SOFTWARE / Wow6432Node umgeleitet. Um diese Umleitung zu umgehen und um das in der Registry stehende Hauptverzeichnis von Addendum für Albis on Windows lesen zu können, musste ich diese Funktionen integrieren. |
| **21.05.2018** | **F+**  | **MonitorScreenShot()** - schießt einen kompletten Monitorscreenshot, diese Funktion ist hilfreich bei der Fehlersuche in Skripten die unbeaufsichtigt laufen sollen |
| **24.05.2018** | **F-**  | **GDI_GUI(), HilfeMenu(), Help_HotKeyGui()** - entfernt und dem Praxomat-Skript hinzugefügt, die Funktionen werden auch nur dort benötigt, somit schlankere Bibliothek |
| **25.05.2018** | **F~**  | Hotkeys verändert                                            |
| **31.05.2018** | **F+**  | **ErrorBox()** - eine Funktion um Daten ins Fehlerlogbuch (Verzeichnis: /logs'n'data/ErrorLogs/Errorbox.txt) zu schreiben. Die Funktion ermöglicht auch einen Screenshot eines ausgewählten Monitors oder aller Monitore zu erstellen. |
|                | **F+**  | **CheckAISConnector()** - sieht nach ob der AIS Connector (Laborverbindungsprogramm) läuft und startet es bei Bedarf neu. Jeder Neustart der AIS Connector.exe wird mit der Funktion *Errorbox()* protokolliert. |
|                | **#A#** | **Addendum Version V0.72**                                   |
| **09.06.2018** | **F~**  | Bibliothek aufgeräumt, Layout verbessert, SaveHBitMapToFile für andere Screenshotoperationen hinzugefügt |
| **26.06.2018** | **F+**  | **listDirectory()** als spezifische Funktion für die Einsortierung vorab eingescannter Befunde erstellt |
|                | **F~**  | **AlbisIsBlocked()** - optionaler Parameter hinzugefügt, die Funktion kann nur ein Statement zurückgeben oder aber mit oder ohne Rückfrage ein oder mehrere blockierende Fenster schließen |
| **29.06.2018** | **F-**  | ChipKartenNachfrage() - Funktion nach Praxomat_st verlegt, wird nur von diesem Skript benötigt |
| **11.07.2018** | **F+**  | **IndexedDir()** - schreibt eine Liste mit allen pdf Befunden als Text Datei in den logs`n`data Ordner, dies dient der schnelleren Anzeige aller noch ungelesen Befunde |
| **15.07.2018** | **F~**  | Hotkeybereich verbessert, Code verschlankt                   |
| **28.07.2018** | **F~**  | **AlbisPrepareInput()** - verbessert für mehr Flexibilität   |
|                | **F+**  | **StrDiff()** - SIFT3 : Super Fast and Accurate string distance algorithm, Nutze ich um Rechtschreibfehler auszugleichen |
|                | **F+**  | **WaitForNewPopUpWindow()** - Funktion wartet auf ein neues PopUpWindow für ein als Handle übergebenes Parent Window und gibt Titel, Klasse, Text und Hwnd zurück, diese Funktion ersetzt in bestimmten Fällen das interne Kommande WinWait, insbesondere dann wenn in Abhängigkeit einer Eingabe in ein Dialogfeld mehrere Möglichkeiten bestehen, welches Fenster sich öffnen könnte und der Name des erscheinenden Fensters nicht vorher gesehen werden kann |
| **29.07.2018** | **F+**  | **AlbisAkteGeoeffnet()** - zum Prüfen ob die Akte eines bestimmten Patienten geöffnet ist, muss auch im Vordergrund liegen |
|                | **F~**  | **AlbisUebertrageGrafischenBefund()** - hat sich manchmal in einem Loop aufgehangen |
| **30.07.2018** | **F+**  | neuer Hotkey Strg+F10 startet bei aktivem Albisfenster das Modul ScanPool - Suchen, Anzeigen, Signieren und automatisches Einsortieren neuer Patienten Befunde im PDF-Format in die Patientenakte |
| **06.08.2018** | **F~**  | **AlbisAkteOeffnen** - Verbesserung der Funktion hinsichtlich Zuverlässigkeit. Jetzt umfassende Erkennung sämtlich möglicher sich öffnender Dialogfenster |
| **27.08.2018** | **F+**  | **PraxTT** - Tooltip Ersatz im Addendum/Nutzer Design mit off-Timer Feature |
|                | **F+**  | **ExceptionHelper()** - zur Anzeige von runtime Fehlern in eigenen Funktionen |
| **31.08.2018** | **M+**  | große Veränderung zumindestens namentlich. **Praxomat_st** gibt es nicht mehr, es heißt jetzt **Addendum** wie die	Sammlung und ist das Hauptskript das die Steuerungs- und Überwachungsaufgaben hat |
| **04.09.2018** | **M~**  | die Fensterüberwachungsroutinen basieren jetzt zum größten Teil auf WinEventHooks , das spart bis zu 8% CPU Zeit |
| **07.09.2018** | **F+**  | **kleine Funktion die per Alt+q** erreicht werden kann, fertigt einen Screenshot einer auswählbaren Region an,	der Screenshot wird in die Zwischenablage gelegt |
| **23.09.2018** | **M~**  | Laufzeitverbesserung und weniger Verbrauch von CPU-Zeit durch Nutzung der **WinEventHook-Methode**, keine Vielzahl an If-Abfragen, der Hook startet jeweils ein anderes Label für die Fensterbehandlung |
| **06.10.2018** | **F+**  | alle bisherigen Albisfenster sind jetzt per WinEventHook automatisiert, der CPU-Verbrauch ist runter auf 0,1%!!! |
| **12.10.2018** | **F~**  | etliche Funktionen hoffentlich verbessert, neue allgemeine Funktionen hinzugefügt, andere spezielle in die Skripte integriert |
| **13.10.2018** | **F~**  | Verbesserung der Funktionen **AlbisGetCaveZeile()** und **AlbisSetCaveZeile()** Funktion, enormer Geschwindigkeitszuwachs bei der Abarbeitung erreicht, von >5s runter auf 0.9s, und **AlbisSetCaveZeile()** überprüft jetzt ob die Zeile tatsächlich geschrieben wurde |
|                | **F+**  | **AlbisPatGeburtsdatum()** - ermittelt das Geburtsdatum aus dem aktuellen Fenstertitel des Albisfenster |
|                | **F+**  | **Datediff()** - neue Funktion z.B. zum Ermitteln des Patientenalters hinzugefügt. |
| **15.10.2018** | **M+**  | Beginn des **AutoComplete**-Modules! neue Albisfenster werden abgefangen, arbeite an der Verbesserung der Hook-Funktionen |
| **24.10.2018** | **F+**  | **AlbisGetAllOpenMDITabs** - Patientenakten, Wartezimmer, Scheinrückseite, Labor werden in sogenannten MDI-Fenstern (Multidokumentfenster) angezeigt, diese sind sogenannte Child-Fenster eines Child-Fenster und über die internen Autohotkeybefehle wie WinGet nicht so einfach zu finden. Diese Funktion soll das umschalten zwischen den einzelnen Dokumentfenstern erleichtern |
| **25.10.2018** | **F+**  | **AlbisAutoLogin()** und **AlbisDefaultLogin**, erste Funktion ist für automatisiertes Einloggen und automatisches Maximieren des Wartezimmerfensters, die zweite Funktion liest aus der Praxomat.ini die Passwörter für diesen Vorgang aus |
| **31.10.2018** | **F~**  | die Namen lokaler Variablen geändert, wenn diese gleich globalen Variablen aus Funktionsbibliotheken waren |
| **11.11.2018** | **F~**  | etliche Funktionen komplett überarbeitet, insbesondere das Ermitteln von Patientendaten erfolgt aus dem Titel des Albisfenster per RegEx, alle Funktionen für Albis beginnen jetzt mit Albis.... |
| **04.12.2018** | **F+**  | **AlbisGetWartezimmerID()** - ermittelt das Handle des Wartezimmerfenster innerhalb des MDI Fenster |
|                | **F+**  | **AlbisGetAllMDIWin()** - erstellt ein globales Objekt, welches die Namen, die Klassen und die Handles aller geöffneten MDI Fenster enthält |
|                | **F+**  | **AlbisGetMDIMaxStatus()** - stellt fest ob das gewählte MDI Fenster maximiert und im Vordergrund ist |
|                | **F+**  | **AlbisGetActiveMDIChild()** - ermittelt das Handle des aktuellen MDI-Childfensters |
|                | **F+**  | **AlbisGetActiveWindowType()** - ermittelt den aktuell bearbeiteten Dokumenttyp - Wartezimmer, Terminplaner, Akte |
| **21.12.2018** | **F+**  | **AlbisGetAllMDITabNames()** und **AlbisGetSpecificHMDI()** - zwei neue Funktionen für das Handling von MDI Fenstern, mit der ersten Funktion kann man die Namen aller Tabs eines SysTabControls321 ermitteln, die zweite Funktion ermöglicht eine spezielle Fenster-ID eines MDI-Fenster zu ermitteln indem man einfach einen Namen übergibt z.B. AlbisGetSpecificHMDI("Wartezimmer") |
|                | **F+**  | **ObjFindValue()** - sucht nach dem übergebenen Wert in mehreren key:value Objekten, welche in einem indizierten Array enthalten sind, gibt den Index key zurück, gehört zu *AlbisGetAllMDITabNames()* |
| **23.12.2018** | **F+**  | **AlbisActivateMDITab()** - aktiviert ein MDI-Tab, so kann man im MDI-Fenster "Wartezimmer" z.B. zwischen "Überblick" und "Arzt" oder auch weiter nach "Abgemeldet" umschalten |
| **04.01.2019** | **F+**  | **ControlSetTextEx()** - erweiterte ControlSetText Funktion - kontrolliert ob der Text tatsächlich eingefügt wurde und man kann noch eine Verzögerung übergeben |
|                | **F~**  | **AlbisAutoLogin()** - nach drücken auf Nein in der MsgBox wurde die Hookfunktion im Addendum-Skript erneut getriggert, die MsgBox erschien unmittelbar erneut so daß man manuell kein Passwort eingeben konnte, eine erneute Nachfrage wird jetzt für 30 Sekunden unterdrückt |
| **05.01.2019** | **F+**  | **AlbisRestrainLabWindow()** - diese Funktion bearbeitet die Fenster die beim Laborabruf oder beim Übertragen ins Laborblatt erscheinen, den Anfang macht das Fenster "GNR der Anford.Ident", für dieses Fenster speichert die Funktion Daten aus dem Fenster in eine Datei damit diese später in die Akte eingefügt werden können und sie wählt bei 2 vorliegenden Quartalen automatisch das richtige Quartal aus, AlbisRestrainLabWindow() soll das Laborabruf Modul zuverlässiger machen, da es automatisch durch die Hooks im  Addendumskript aufgerufen wird. |
| **08.01.2019** | **F+**  | **AlbisLaborbuch()** - Funktion ruft das Laborbuch auf       |
| **19.01.2019** | **F~**  | **AlbisOeffneWartezimmer()** - Code verbessert, gibt jetzt 0 oder 1 für "nicht erfolgreich" oder "erfolgreich geöffnet" zurück |
|                | **F+**  | **AlbisGet_StammPos()** - liest aus der Local.ini die Positionen der Stammdatenfenster aus |
| **17.02.2019** | **F+**  | **AlbisStatus()** - ermittelt ob Albis bereit ist Befehle zu empfangen |
| **21.03.2019** | **F~**  | **AlbisSetzeProgrammDatum()** - das eingetragene Datum wurde manchmal von Albis nicht akzeptiert und erzeugte ein Fehlerhinweis |
| **25.03.2019** | **F+**  | **GetFocusedControlClassNN()** - ClassNN des aktuell fokusierten Controls im aktiven Fenster ermitteln |
| **25.03.2019** | **F~**  | **WMIEnumProcessExist()** - Codelänge gekürzt, dürfte minimal schneller sein |
| **25.03.2019** | **F~**  | **AlbisSetCaveZeile()** - Code optimiert und versucht die Erkennung des Fensters deutlich zu verbessern, absolut fehlerfrei zu bekommen |
| **31.03.2019** | **F+**  | **AlbisMenu()** - eigentlich nur ein Wrapper für den *Send- oder Postmessage* Befehl kann aber hier das zu erwartende Fenster übergeben werden um. Die Funktion wartet selbstständig bis das Fenster erscheint. |
| **09.04.2019** | **F+**  | **AlbisOeffneLaborBlatt(), AlbisOeffneKarteikarte()** - wechseln zum Laborblatt oder zur Patientenkarteikarte per Funktionsaufruf |
| **13.04.2019** | **F+**  | **AlbisGetActiveWindowType()** -  erkennt welcher Inhalt im Patientenfenster geöffnet ist - Patientenakte/Laborblatt/Biometriedaten/Rechnungen/Abrechnung |
|                | **F+**  | **AlbisKarteikartenFensterAktivieren()** -  Funktion sollte vor Funktionen die aus der Karteikarte etwas lesen oder etwas schreiben wollen aufgerufen werden. |
| **04.05.2019** | **F+**  | **ControlCmd()** - smarte Funktion für die Vereinfachung der Robotic Process Automation (noch im Aufbau), welche nur noch einen Befehlsnamen für unterschiedliche Controls kennt, die Funktion entscheidet anhand des ClassNN Namens welche Befehle verwendet werden |
| **11.05.2019** | **F+**  | **AlbisOeffnePatient()** - zum Aufrufen des Dialogfensters 'Patient oeffnen' |
|                | **F~**  | Library entrümpelt!, unbenutzte Funktionen entfernt, weniger Code bei anderen Funktionen, kleinere Optimierungen |
| **15.05.2019** | **F+**  | **ControlCmd()** - Funktionsumfang erweitert, Fehlerbereinigung |
| **18.05.2019** | **F+**  | **FormatedFileCreationTime()** - Auslesen und Formatieren des Erstellungsdatums einer Datei ins deutsche Format |
| **26.05.2019** | **F+**  | **AlbisSchreibeLkMitFaktor** - Vereinfachung zum Senden eines Leistungskomplex Hotstrings |
| **01.06.2019** | **F~**  | **MCSVianovaWebClient** - Funktion bekam das Fenster nicht geschlossen, Bearbeitung dieses Fensters startet selbstständig alle weiteren notwendigen Schritte bis selbst die Labordaten in die Patientenakten eingefügt sind ; Codeoptimierungen in anderen Funktionen |
| **07.06.2019** | **F~**  | **AlbisAutoLogin()** und **AlbisDefaultLogin()** - beide Funktionen waren durch fehlerhafte Übergaben praktisch funktionslos, in einigen anderen Funktionen konnte Quellcode verkürzt werden durch ersetzen mit neuen Funktionen |
| **19.06.2019** | **F+**  | **AlbisCloseMDITab()** - schließt ein per Title identifiziertes MDIClient Fenster (z.B. eine Tagesprotokollanzeige) |
|                | **F+**  | **AlbisErstelleTagesprotokoll()** - erstellt Tagesprotokolle durch Aufruf des Albismenu, der Funktion können Daten in Form von Quartalsangaben oder Datums-Zeiträumen übergeben werden. Tagesprotokolle können automatisch gespeichert werden. Der Dokument-Tab des Tagesprotokoll kann auf Wunsch nach Abschluss aller Vorgänge automatisch geschlossen werden. |
| **23.06.2019** | **F~**  | **AlbisAkteOeffnen()** - Codeausführung ist jetzt deutlich schneller und zuverlässiger |
|                | **F~**  | **WaitForNewPopUpWindow()** - der Funktion die Möglichkeit hinzugefügt nicht nur Popup-Fenster zu erkennen, sondern auch auf eine Änderung des Fenstertitels des Parent-Fenster zu reagieren, dies beschleunigt die Ausführung in den überwiegenden Fällen |
| **25.06.2019** | **F~**  | Verbesserungen in der Fenster Erkennung mehrerer Funktionen  |
|                | **F+**  | **AlbisErstelleTagesprotokoll()** - erstellt Tagesprotokolle durch Aufruf des Albismenu, der Funktion können Daten in Form von Quartalsangaben oder Datums-Zeiträumen übergeben werden. Tagesprotokolle können automatisch gespeichert werden. Der Dokument-Tab des Tagesprotokoll kann auf Wunsch nach Abschluss aller Vorgänge automatisch geschlossen werden. |
| **27.06.2019** | **F+**  | **AlbisDateiAnzeigen(file)** - neue Funktion um in Albis z.B. ein mit dem Abrechnungshelfer erstelltes Protokoll anzeigen zu lassen |
| **07.08.2019** | **F+**  | **AlbisIsElevated()** stellt fest ob Albis mit UAC Virtualisierung gestartet wurde. Damit kann man die Notwendigkeit feststellen das aufgerufene Skripte ebenso mit UAC-Virtualisierung gestartet werden damit diese Albis steuern können. |
| **09.08.2019** | **F~**  | **AlbisLaborAuswaehlen()** - bearbeitet jetzt auch das Fenster welches erscheint wenn keine Labordaten im Laborordner vorhanden sind. WinTitle: "ALBIS", WinText: "Keine Datei(en) im Pfad....." |
| **23.08.2019** | **F+**  | **AlbisDruckeLaborBlatt()** - Ausdruck der Laborwerte über eine Aufruf durch ein Hotkey oder aus einem Skript heraus. Funktion hat einen automatischen Einstellungsdialog. Ein Ausdruck auch durch Übergabe des Druckernamens möglich. Der Funktion können auch Parameter wie ein Dateiname zur Erzeugung einer PDF-Datei übergeben werden. |
| **05.09.2019** | **F~**  | unbenutzte Funktionen entfernt,  ältere Funktionen an die aktuellen Konventionen von Autohotkey_L 1.1.30 angepasst |
| **06.09.2019** | **F+**  | **AlbisHeilMittelKatologPosition()** - verschiebt den CGM Heilmittelkatalog in die Mitte des aktuellen Monitors und löst damit das ab und zu auftretende Problem das der Druck von F3 in der Heilmittelverordnung die Ausfüllhilfe nicht angezeigt hat. Die Ausfüllhilfe wird nur ausserhalb des sichtbaren Monitorbereiches positioniert und war deshalb nur nicht sichtbar |
| **12.09.2019** | **F~**  | **AlbisAutoLogin()** - der Klick auf den OK Button wird jetzt per Mausklick-Simulation durchgeführt, der Click mittels einer anderen Technik hat das Login-Fenster nicht geschlossen und den Hook für die Erkennung des Login-Fenster erneut ausgelöst |
|  | **F+**  | **MoveWinToCenterScreen()** - verschiebt ein außerhalb des sichtbaren Monitorbereich liegendes Fenster in die Mitte des aktuellen Monitor, funktioniert gerade auch im Multi-Monitor-Einsatz |
| **13.09.2019** | **F~**  | **AlbisAutoLogin()** - korrigiert so das das Dialogfenster auch sicher geschlossen wird |
| **16.09.2019** | **F~**  | **FindChildWindow()** - Funktion kompatibel zur Autohotkey-Version V1.1.30.03 gemacht<br>**AlbisGetActiveMDIChild()** - Code bereinigt |
| **25.09.2019** | **F+**  | 3 neue Funktionen für die Automatisierung von Albis hinzugefügt. **AlbisDruckeBlankoFormular()**, **AlbisDruckePatientenausweis()**, **AlbisRezeptHelfer()** |
| **28.09.2019** | **F+**  | **AlbisFristenRechner()** und **AlbisFristenGui()**  - Beginn und Ende der Krankengeldzahlung werden in jedem Arbeitsunfähigkeitsformular angezeigt |
| **01.10.2019** | **F~**  | **AlbisErstelleTagesprotokoll()** - mehr Statusanzeigen<br>**AlbisDateiAnzeigen()** - verbesserte Zuverlässigkeit |
| **02.10.2019** | **F+**  | **AlbisKarteikartenFocusSetzen()** - ähnlich der Funktion *AlbisKarteikarteAktivieren()* hat diese Funktion endlich die Zuverlässigkeit die ich mir gewünscht habe, vermutlich werde ich die ältere Funktion irgendwann komplett entfernen |
|                | **F+**  | **AlbisSchreibeInKarteikarte()** - ist ursprünglich zusammen mit *AlbisKarteikartenFocusSetzen()* für mein Modul Abrechnungshelfer entwickelt worden. Jetzt wird es möglich sein, jeden Abend ohne Beisein die während der Sprechstunde eingegeben Ziffern/Diagnosen (oder was auch immer man will), automatisiert korrigieren oder ergänzen zu lassen. |
| **23.11.2019** | **F+**  | **AlbisRezeptHelferGui()** - erstellt ein Auswahlfeld im Rezept zur Auswahl von Rezeptvorlagen, ermöglicht auch das Erstellen neuer Vorlogen<br>**AlbisRezeptFelderLoeschen(),<br>AlbisRezeptFelderAuslesen(), <br>AlbisRezeptSchalterLoeschen()<br>** - Rezepthelferfunktionen - Löschen aller Eingaben in einem Rezept                                   (Medikamenten- sowie Zusatzfelder)<br>- Auslesen aller Felder eines Rezeptes (Medikamente, Zusätze, Schalter) |
| **21.12.2019** | **F~**  | **AlbisRezeptHelferGui()** <br>**1.)** das Rezeptformular wird automatisch unterhalb der Dauermedikamente angezeigt, damit diese immer vollständig zu sehen sind.<br>**2.)** das Rezeptfenster zeigt keine Werbung mehr an! |
|  | **F+**  | **AlbisRezeptAutIdem()** - setzt automatisch ein Aut-Idem Häkchen, wenn man in den Dauermedikamenten in der Nachtmedikation ein großes A notiert hat<br>die Funktion installiert einen Hook auf das Rezeptformular (Muster16) von Albis, dies funktioniert leider noch nicht gut |
| **22.12.2019** | **F~**  | **AlbisGetMDIClientHandle()** - ermittelt das handle des MDI Client Bereiches jetzt zuverlässig |
| **23.12.2019** | **F+**  | **AlbisRezept_DauermedikamenteAuslesen()** - im Rezeptfenster werden in einer Listbox alle Dauermedikamente des Patienten angezeigt, dieses Feld wird ausgelesen. Die Daten können dann für verschiedene weitere Steuerungen verwendet werden. Zum Beispiel könnte es automatisch ein Hinweis erfolgen, wenn man ein Medikament verordnet hat, gegen das der Patient allergisch ist oder das bei ihm kontraindiziert ist. |
| **27.12.2019** | **F~** | **AlbisFristenGui()** - Fehlerkorrektur und behandelt auch die Ausnahme das der Shift+F3 Kalender automatisch geschlossen wird falls das AU-Formular geschlossen wurde |
|  | **F~** | **AlbisDateiAnzeigen()** - Fehler beim Schliessen des Dialogfensters beseitigt |
|  | **F~** | **AlbisZeigeKarteikarte()** - schließt jetzt zwei Dialogfenster die auftreten können, wenn eine Tastaturfokus gesetzt ist und eine Eingabe erfolgt war |
| **28.12.2019** | **F~** | **AlbisGetMDIMaxStatus()** - Aufruf der Funktion ist jetzt ohne vorherigen Aufruf von AlbisGetAllMDIWin() möglich, der Funktion kann ein Handle oder der Name des Fensters übergeben werden |
|  | **F+** | **AlbisDateiSpeichern()** - speichert ein von Albis erstelltes Protokoll/Statistik |
|  | **F+** | **AlbisErstelleZiffernStatistik()** - erstellt eine Ziffernstastik (Menu Statistik\Leistungsstatistik\EBM2000Plus/2009Ziffernstatistik), als Parameter kann ein Quartal, ein Zeitraum oder ein Tag übergeben werden, der Dateipfad und der Dateiname in welches die Statistik gespeichert werden soll und ob die Statistikausgabe in Windows geschlossen werden soll. |
| **29.12.2019** |  | **AlbisFormular()** - Aufrufen von Formularen mittels Übergabe eines String z.B. eHautkrebsScreening Nicht-Dermatologe = eHKS_ND, Gesundheitsvorsorge = GVU |
|  | **F+** | **AlbisHautkrebsScreening()** - befüllt das eHautkrebsScreening (nicht Dermatologen) Formular *(Das Hautkrebsscreeningformular für Nichtdermatologen wurde mehrfach geändert in den letzten Quartalen. die ClassNN der Buttons hat sich dabei jedesmal geändert, damit wurden nicht die richtigen Buttons angesprochen und somit war das Formular nicht korrekt ausgefüllt oder es ließ sich nicht speichern da sich auch für den 'Speichern' Button die ClassNN geändert hatte. Letzteres versuche ich nun mittels Übergabe des Namen/Text des Speichern-Buttons treffsicherer zu machen)* |
| **07.01.2020** | **F~** | Addendum_Albis/**AlbisGetActiveControl()** - AfxFrameOrView und AfxMDIFrame wurden in intern in Albis umbenannt, aus AfxFrameOrView90 und AfxMidiFrame90 wurde AfxFrameOrView140 und AfxMidiframe140, die Erkennung wurde deshalb aufgrund eventuell in Zukunft folgender Änderung der Klassennamen verallgemeinert |
| **08.01.2020** | **F~** | Addendum_Albis/**AlbisGetActiveControl()** -Anpassung an die Klassennamensänderung |
| **16.01.2020** | **F~** | **AlbisRezeptHelferGui()** - Neuzeichnen und versetztes Anzeigen der zusätzlichen Auswahl korrigiert, breiteres Steuerelement für mehr Informationen |
| **30.01.2020** | **F+** | **AlbisImportiereBild()** - importiert eine jpg Datei in eine Patientenakte |
| **30.01.2020** | **F~** | **AlbisImportierePdf()** - Performance verbessert, **AlbisUebertrageGrafischenBefund()** und **AlbisKarteikartenFocus()** - Fehlerrückgabe angepasst, Performance und Zuverlässigkeit deutlich verbessert |
| **03.02.2020** | **F+** | **IfapVerordnung()** - Medikamente lassen sich per Kurzform über einen Funktionsaufruf mit Hilfe von ifap verordnen, dazu muss eine PZN des entsprechenden Medikamentes hinterlegt sein. die Funktion druckt automatisch (auf Nutzernachfrage) zum Medikament hinterlegte Word-Dokumente ohne das Microsoft Word sichtbar wird. Damit erhalten Sie die Möglichkeit selbst entworfene Dosierungsangaben und Patientenhinweise automatisch mit der Medikamentenverordnung auszudrucken oder sie es wird ein Dokument ausgedruckt, welches der Patient z.B. als Einverständniserklärung unterschreiben muss. Sie können auch mehrere Dokumente mit einer Verordnung ausdrucken. Dazu müssen diese aber alle in einem Verzeichnis auf ihrer Festplatte liegen. |

![Praxomat.png](Docs/Icons/Abrechnungshelfer.png) Abrechnungshelfer
| Datum          | Teil     | Beschreibung                                                 |
| -------------- | -------- | ------------------------------------------------------------ |
|  **16.12.2019**  |  auswertbare Tagesprotokolle werden als auswählbare Baumstruktur (Treeview) angezeigt und können einzeln oder insgesamt ausgewählt werden. |
|  **17.12.2019**  |  Das Makro: "freie Statistik" hat eine Gui bekommen und wird in Zukunft interessante Daten ermittlen können |
|  **03.01.2020**  |  Vollständigiger Parser der Tagesprotokollausgaben erreicht. Die Daten des Protokolls werden in ein Autohotkey-Objekt umgewandelt und sind damit Analysen schneller zugänglich. |


![](Docs/Trenner_klein.png)

## ![Praxomat.png](Docs/Icons/Praxomat.png) Praxomat

| Datum          | Teil     | Beschreibung                                                 |
| -------------- | -------- | ------------------------------------------------------------ |
|                | **P+**   | StrC + StrX + StrV funktionieren jetzt innerhalb von Albis - sonst wurde immer die Einleseroutine für die Kartenlesegeräte aufgrufen |
|                | **P**+   | Patienten lassen sich per Hotkey in eine externe Textdatei ablegen - GVU/KVU Liste |
|                | **M**+   | es besteht innerhalb des Praxomat ein Programm welches das GVU/KVU Formular fast ohne 	Eingriff - per Standard ausfüllt (das spart clicks und Zeit), das Modul schreibt die entsprechenden |
|                | **M**-*  | Ziffern ein und legt ein neues oder ergänzt ein Errinnerungscode im cave Fenster Zeile 9 (z.B. GVU/HKS 03^17) - dies steht für Monat und Jahr geplant ist das diese Liste später automatisch abgearbeitet werden kann und das dann ohne jeden Eingriff, das Modul öffnet die Patientenakte, ändert das Datum in der Akte auf das in der GVU Liste hinterlegte Datum, füllt erst das GVU und dann das KVU Formular aus, schreibt die Ziffern in die Akte, ändert den Code in cave und schließt die Akte wieder |
|                | **M**+/- | es existiert ein Modul welches es ermöglicht, Quartalsprotokolle (aus Tagesprotokollen) zu erstellen, im Anschluß	werden alle unwichtigen Daten entfernt, dies ist der Grundbaustein für 								weitere Auswertungen, Erstellung von Statistiken. Hauptsächlich werde ich es wohl zur automatischen Erkennungen von fehlenden Ziffern verwenden. Listenvergleich AOK,BARMER |
| **11.11.2017** | **P+-**  | wenn Albis abstürzt - wird es neu gestartet - da er Albis mehrfach startet - lässt sich nicht verhindern - wird mittels Messagebox zuvor nachgefragt |
| **14.11.2017** | **P+-**  | animiertes ToolTip ( --Bezier_MouseMove2() und TToolTip()-- ) funktioniert wenn ein Pat. in die GVU Liste eingetragen wurde - damit man weiß das er auch in der Liste gelandet ist 							(*Liste muss auf doppelte Einträge überprüft werden*) - erledigt am 06.12.2017 |
| **16.11.2017** | **P+***  | Albisstart jetzt mit Autologin für die jeweiligen Rechner - Daten werden in der ini Datei hinterlegt |
| **20.11.2017** | **P#**   | Praxomat V0.4 - sag ich mal                                  |
| **23.11.2017** | **P+**   | Timer zeigt die Zahl der Wartenden Patienten und die maximale Wartezeit an |
| **24.11.2017** | **P+**   | Patientenaufruf  aus dem Wartezimmer (nur Arzt) setzt nach Zustimmung den Patientenstatus auf -in Behandlung- -> dies funktioniert nur mit der ENTER-Taste |
| **26.11.2017** | **P+***  | beim Öffnen eines Patienten aus dem Wartezimmer heraus - wird der Timer neu gestartet - dieser Timer wird an den Patientennamen gebunden									schaltet man zurück auf das Wartezimmer pausiert der Timer bis man zurück zu ist (*im Moment noch in jedem Patientenfenster: das Auswählen eines neuen Patienten aus dem Wartezimmer	 bei noch laufenden Timer führt zur Nachfrage ob man die Timerdaten abspeichern möchte*) |
|                | **F-**   | *das Speichern der Timerdaten - soll bei Überschreitung jeder 10.min zum Einschreiben der Gesprächsziffer lko 03230 oder 35110 mit jeweiligem Faktor auffordern*) |
| **28.11.2017** | **P~**   | Fehlerkorrektur des Patientenaufruf, aufgerufener Patient wird registriert, dies kann mehr als 10sec dauern, es hat wohl etwas damit zu tun das die Albis Fenster Modus - Overlapped	eingestellt haben |
|                | **PF+-** | Beginn der Integration eines DebugFensters auf Basis eines ListView |
|                | **P+**   | per Strg+ Win + linke Maus zu erreichende WindowInfo Funktion integriert (gibt viele Daten aus - aber funktioniert nicht so wie ich es gern hätte) - ist im Moment für mich für die weitere Programmierung notwendig |
|                | **PF+*** | ein GUI ( TENMINUTESGUI() ) zur Auswahl der 3 möglichen Gesprächsziffern erstellt, anhand der Gesprächsdauer wird auch gleich der Faktor vorgeschlagen , dieser läßt sich ändern ebenso welche Ziffer gewählt wird |
| **30.11.2017** | **P+-**  | farbliche Kennzeichnung Praxomat Timeranzeige nach Zeitüberschreitung türkis, ---------> rot , rudimentär programmiert mit nicht so schönen Farben |
|                | **P+**   | nach drücken auf die *PAUSE*-Taste wird der Timer angehalten , das Hinweisfenster sieht sehr hübsch aus im Gegensatz zur farblichen Kennzeichnung der Timer-Gui (siehe oben) |
| **01.12.2017** | **P-~**  | das Überprüfen ob die richtigen Einträge von z.B. lko 03230(x:2) in den Feldern (Edit3, RichEdit..) funktioniert nicht, AHK kann die Felder nicht auslesen, diese Funktion war zur Fehlerüberprüfung gedacht, produzierte aber nur Fehler (Praxomat_Functions.ahk). Die bisherigen Programmteile habe ich als Snippets gesichert und aus dem Funktionscode entfernt. |
|                | **P~**   | bei Privatpatienten wird wieder die Gesprächszeit eingetragen, nachdem das Eintragen beendet ist wird die Akte geschlossen und nachgefragt ob der Patient aus der Wartezimmerliste |
|                | **P~**   | das Timer-Gui ist jetzt als child-Fenster an das Albis Fenster gebunden |
|                | **P-**   | die ehemalige Errinnerungsfunktion für das Timer-Gui damit ich das Eintragen der Gesprächszeit nicht vergesse ist jetzt überflüssig geworden, die Funktion ist jetzt für eine manuelle |
| **02.12.2017** | **P+~**  | der obere Teil der Timer-Gui ist jetzt mit der GDI+ Library gestaltet, leider verschwindet die GDI Gui wenn man sie als child binden möchte, lt. Autohotkey-Forum ist die Transparenz eines |
| **03.12.2017** | **P+**   | Information zum Programmablauf werden jetzt in der letzten Zeile des Timer-Gui angezeigt (Infopraxo:), dort sollen dann auch Fehler oder andere wichtige Hinweise erscheinen, |
|                | **P#**   | Praxomat/Addendum erreicht **"Level"** 0.5                   |
| **05.12.2017** | **P+**   | die Hintergrundfarbe der Timer-Gui ändert sich nur wenn ein Patient in Behandlung gesetzt wurde |
| **06.12.2017** | **P~**   | die zigfachste Fehlerkorrektur bei der Aufnahme des Behandlungspatienten. goto Routinen gesetzt, gosub gelassen wo es hingehört. |
|                | **PF+**  | die Routine für die Aufnahme in die GVUListe überprüft auf schon vorhandene Eintragungen, im Anschluß wird im Timer-Gui die Gesamtzahl der Untersuchungen angezeigt |
| **07.12.2017** | **P+**   | gleich am Anfang wir die GVU-Liste eingelesen, damit die Anzahl der aufgenommenen Patienten im TopToolTip angezeigt werden kann, siehe 06.12. |
|                | **P+**   | Anzeige eines ToolTip oberhalb der Timer-Gui da das Aufnehmen der Behandlungspatienten immer noch nicht reibungslos funktionierte über die Praxomatinterne Funktion TopToolTip() |
| **12.12.2017** | **P~**   | das GDI-Overlay Fenster wird versteckt wenn das Albisfenster nicht mehr vorne liegt, leider läßt sich das Overlay Fenster nicht als ChildWindow anbinden , es wird dann gar nicht mehr angezeigt |
| **15.12.2017** | **P~**   | Praxomat: Steuerung + F11 und ENTER funktionierten trotzt diverser vermeindlicher Fehlerbehebungen gar nicht mehr, durch das Überprüfen welches Fenster vorn liegt um die Timer-Gui je nachdem sichtbar oder unsichtbar zu machen, kam es zu einem Flackern der Anzeige, diese Funktion ist wieder entfernt |
| **19.12.2017** | **P~**   | das Unterbringen von Ziffern mit Faktor z.B. 03230(x:1) - wird nicht von Albis akzeptiert (bei der Vorbereitung zur Abrechnung hat er diese Eintragungen nicht akzeptiert), deshalb entfällt jetzt das (x:1) |
| **20.12.2017** | **P+**   | die Timer Gui ist jetzt vollständig eine GDI Gui, somit kein Flackern der Ausgabe mehr - damit läßt sich aber das Fenster nicht als child an Albis binden |
|                | **P#**   | **Praxomat Version 0.55**                                    |
| **28.01.2018** | **P~**   | **diverse Fehlerkorrekturen** seit dem 23.01. , Problem gelöst das manchmal die Nachfrage ob ein Patient aus der Wartezimmerliste gelöscht werden zweimal erschien |
| **02.03.2018** | **P+**   | die PraxomatGui ist nicht immer AlwaysOnTop, je nach Mouseposition wird dieser Wert- ein oder ausgeschaltet |
| **04.04.2018** | **P~**   | **Praxomat V0.62**: das Ausführen der Funktion zum Erstellen der GVU Liste nutzte bisher das aktuelle Datum, dies führte allerdings dazu, das die Liste mit einem neuen Quartal fortgeführt wurde, auch wenn man gerne noch ein paar Formulare im vergangen Quartal unterbringen wollte. Jetzt wird das eingestellte Programmdatum dazu genutzt. Kleinere Fehlerkorrekturen in der GVU Listenfunktion durchgeführt |
| **04.05.2018** | **P~**   | **TenMinutes Gui**: Fehler das manchmal die Gesprächszeit nicht berechnet wurde beseitigt |
| **14.05.2018** | **P~**   | **GVU Liste**: zugehörige Funktion GetQuartal() verbessert, *Codebereinigung* - unötiger Code entfernt, PraxomatIcon aus der Taskbar entfernt |
| **17.05.2018** | **P~**   | **TenMinutes Gui:** die Ziffernauswahl arbeitet erstmals absolut zuverlässig und fehlerfrei |
| **18.05.2018** | **P~**   | **diverse Fehlerkorrekturen** - das Pausieren des Timer funktioniert jetzt zuverlässig, Wartezeitdateien werden wieder eingelesen, GVU Liste wird wieder eingelesen - war ein ein Problem in der *GetQuartal*-Funktion, Parsing der Wartezeitanzeige 01:05 nicht 1:5 wie vorher |
| **24.05.2018** | **P+~**  | **Praxomat V0.67** /                                         |
| **29.05.2018** | **P+**   | **Praxomat V0.70** - die GUI zeigt jetzt einen weichen Farbverlauf des Timers und nicht mehr der Hintergrundfarbe. Der Farbverlauf ist vorberechnet aus einem vorher erstellten Bild mit einem horizontalen Gradienten über mehrere Farben. Mit Hilfe eines Skriptes wurde dann eine Liste mit RGB Farbwerten (3000 Stück) erzeugt. Bei max. 25min eingestellter Gesprächszeit wird jeweils alle 500ms ein neuer Farbwert im Hintergrund des Fensters gesetzt. |
| **30.06.2018** | **F~**   | diverse Fehlerbereinigungen, insbesondere bei der Erstellung der GVU Liste traten Formatfehler auf so daß die GVU-Liste durch das Modul GVU_Queue_ausfüllen nicht abgearbeitet werden konnte, verloren gegangene Funktionen für dieses Modul der Praxomat-Functions.ahk wieder hinzugefügt |
| **04.10.2018** | **F+**   | HelpViewGui verbessert, Buttons anstatt Text zum direkten Ausführen von Befehlen gedacht |
|                | **F+**   | das PraxomatGui kann sich jetzt automatisch an die richtige Stelle positionieren, es verschiebt sich auch mit dem Albisfenster |
|                | **F+**   | Praxomat zeigt in der Gui jetzt mehr Informationen           |
| **04.10.2018** |          | **Praxomat Version 0.77**                                    |
| **13.12.2018** | **F+**   | SkriptIcon ist in das Skript integriert                      |
| **08.04.2019** | **F~**   | nur Code übersichtlicher gemacht                             |

<br>

![](Docs/Trenner_klein.png)

## ![Dicom2Albis.png](Docs/Icons/Dicom2Albis.png) Dicom2Albis

|     Datum      | Teil    | Beschreibung                                                 |
| :------------: | ------- | ------------------------------------------------------------ |
| **10.12.2017** | **M+*** | **Dicom2Albis** - ein Modul um die Bilder einer Röntgen-CD automatisiert einzulesen - ist so gut wie fertig, das Modul wandelt mit Hilfe einer freien Software die sonst nicht lesbare DICOMDIR Datei in eine xml-Datei um dort wird dann nach den wichtigen Informationen gesucht<br />im Anschluss werden die Dateien mit Hilfe einer weiteren freien Software *dicom2.exe* in eine PNG-Datei umgewandelt. <br />Dicom2Albis schreibt Text in das Bild (DICOM Bilder - haben keinen Text in den Bilddaten und benötigen deshalb ein spezielles Anzeigeprogramm) <br />da mir PNG Dateien zuviel Platz wegnehmen werden diese Dateien noch in JPEG mit einer Qualität von 75 umgewandelt.<br />Dann erfolgt das automatisierte Einsortieren in die richtige Patientenakte. |
|                | **M-**  | -es fehlen noch Einstellungsmöglichkeiten für die Grafikausgabe (im Moment nur meine Vorstellungen), das Ausgabedateiformat ist fest vorgegeben	(JPEG)<br />-eine Installations- oder Einstellungsroutine für den automatischen Start nach Einlegen einer CD mit DICOM Inhalt muss noch erstellt werden |
| **13.12.2017** | **M+**  | komplette Funktionsfähigkeit!, Qualität der jpeg ist jetzt auf 90 eingestellt, das Umwandeln und Einsortieren muss nacheinander passieren, ein späteres Einsortieren geht leider nur manuell (im Anschluss hat der Vorführeffekt alles zunichte gemacht) |
| **15.12.2017** | **M~**  | Dicom2Albis Fehlerkorrektur nach dem ersten richtigen Einsatz von Dicom2Albis, das Programm hat teilweise die Daten der zuvor erstellen DICOM.txt Datei benutzt, die Datei wurde scheinbar nicht überschrieben und wird jetzt jedes mal gelöscht |
| **18.12.2017** | **M~**  | Fehlerkorrekturen - Patientennamen enthalten oft folgendes Zeichen (^), manchmal sind diese Zeichen auch am Ende des Namens zu finden, bisher habe ich mit StringReplace dieses Zeichen zwischen Nachnamen und Vornamen ersetzt als Beispiel: Müller^Marie -> Müller, Marie. Letzteres ist die Schreibweise um sie Albis zu übergeben. Mit der Änderung werden vorgestellte und nachgestellte ^-Zeichen getrimmt. Zusätzlich habe ich die Animation Bezeichner-GUI verbessert (dieses Feature musste sein!). |
| **14.01.2018** | **M~**  | Dicom2Albis - die Schrift wurde auf kleinen Bildern zu groß angezeigt, es wird jetzt automatisch die Schriftgröße anhand der Bildgröße angepasst, es besteht weiterhin der Fehler das beim ersten Start des 	Skriptes sogleich die Patientenakte aufgerufen wird ohne das versucht wird Bilder von der CD zu lesen |
| **28.01.2018** | **M+~** | Dicom2Albis, erste Versuche Dicom2Albis vollständig zu automatisieren haben noch nicht zum vollen Erfolg geführt, eine neu eingelegte CD wird erkannt, die Prüfung auf eine DICOM CD funktioniert, aber der Dicom2Albis Start schlägt fehl, zudem funktioniert Dicom2Albis aufeinmal nicht mehr, obwohl ich innerhalb dieses Skriptes keine Veränderungen vorgenommen habe |
| **31.01.2018** | **MF+** | Skriptpause Button und Skriptabbruch Button der GUI hinzugefügt, bei vorzeitigem Abbruch fragt das Script ob bereits konvertierte Bilder gelöscht werden sollen |
| **01.02.2018** | **MF+** | die Größe des VerlaufsGui läßt sich jetzt verändern, die letzte Position und Größe wird in der Praxomat.ini gespeichert |
| **16.02.2018** | **M~**  | die Parameterübergabe an das Skript korrigiert, notwendig für den automatischen Aufruf durch das Praxomat_st Skript nach dem dieses ein eingelegtes DICOM Medium erkannt hat |
| **20.02.2018** | **M~**  | erstmals durchgehende Funktionsweise, alle Pfade nachdem Autostart des Skriptes (CD einlegen , wird als DICOM CD erkannt und das Dicom2Albis Skript wird gestartet) Fehlerbereinigung, Fehlerüberprüfung falls es Probleme beim Umwandeln gab und im Befundordner keine umgewandelte Datei erscheint, vorher hatte das Skript im 100ms Abstand auf den Ok Button des Befundübernahmefenster geklickt, jetzt wird das erscheinende Albisfenster mit Hinweis auf die nicht vorhandenen Dateien erkannt und es wird die nächste Datei ausprobiert |
| **06.03.2018** | **M+**  | bei nicht zu konvertierendem DICOM Inhalt ruft das Skript das kostenlose MicroDicom Tool auf (die Umwandlung in wmv oder avi muss noch manuell erfolgen), nach Fertigstellung der Umwandlung wird die CD automatisch ausgeworfen |
| **14.03.2018** | **M~**  | **DICOM2ALBIS V0.98**: Fehlerkorrekturen beim Aufruf des externen Programmes für die Umwandlung von Serienaufnahmen (MRT,CT) nach AVI oder WMI , bei mDicom.exe die beiden Fenster (Listview und das Logo) werden vor Start des externen Programmes beendet, sie stören sonst nur |
|                | **MF+** | **DICOM2ALBIS V1.00**: mit Hilfe der MicroDicom.exe lassen sich jetzt Serien von Bildern automatisiert in WMV umwandeln, der Filename wird aus den Metadaten der DICOMDIR Datei generiert, Dicom2Albis hat jetzt alle Funktionen die ich brauche (eine fehlt nur noch - gemixte DICOM CD's Röntgen und Schichtaufnahmen werden nicht eingelesen diese CD's sind selten - dazu muss ich noch eine Lösung finden |
| **13.12.2018** | **F+**  | SkriptIcon ist in das Skript integriert                      |

<br>

![](Docs/Trenner_klein.png)

## ![GVU](Docs/Icons/GVU.png) Gesundheitsvorsorgeliste

| Datum          | Teil   | Beschreibung                                                 |
| -------------- | ------ | ------------------------------------------------------------ |
| **04.01.2018** | **M+** | Praxomat Modul_GVU_Queue Ausfüllhilfe, das zuvor die mit Praxomat im Quartal erstellten GVU-Listen abarbeitet hat seinen ersten Probelauf durch |
| **06.01.2018** | **M~** | Praxomat Modul_GVU_Queue Ausfüllhilfe, Optimierung des Ablaufes<br /> Sicherung der geänderten "Cave! von" Fenster-Zeile 9 in die *GVU.txt Listendatei<br />Modul füllt die beiden Formulare Gesundheitsvorsorge und Hautkrebsscreening durch die verbesserten oder neuen Funktionen zuverlässiger aus , ohne wie zuvor manchmal nicht mehr weiterzukommen |
| **07.01.2018** | **M~** | *GVU.txt Datei - Änderung im Prinzip als .csv Datei - Delimiter ist jetzt ein ";" , die Leerzeichen als Trennzeichen hatten dazu geführt das zuviele zu bearbeitende Variablen entstanden sind Änderungen wurden im Praxomat Modul_GVU_Queue Ausfüllhilfe und in Praxomat (Label GVUListe) durchgeführt |
| **06.04.2018** | **M~** | das Prüfen auf einen gültigen Abrechnungsschein zum Abrechnen der Vorsorgeformulare funktioniert jetzt |
| **07.04.2018** | **M~** | das Dokumentieren des Untersuchungsdatums im Cave!von Fenster hatte nicht immer funktioniert, die Berechnung für den Abstand zwischen 2 Vorsorgeuntersuchungen war nicht immer korrekt |
| **30.06.2018** | **F+** | merkt sich jetzt die sich aktuell in Abarbeitung befindende GVU-Liste, der INI-Eintrag wird erst nach vollständiger Abarbeitung wieder entfernt |
| **08.10.2018** | **F~** | kleinere Fehlerbehebungen                                    |
|                | **F+** | wenn man seine Abrechnung noch nicht zum neuen Quartalsanfang nicht geschafft hat, dann gibt es manchmal ein Problem mit dem GNR Vorschlag, es muss dann das Abrechnungsquartal zuvor ausgewählt werden |
| **12.10.2018** | **F~** | Skript reorganisiert , Code übersichtlicher formatiert, Kommentierungen verbessert und/oder noch einige hinzugefügt, allgemeine Skriptbeschreibung verständlicher gestaltet |
| **13.10.2018** | **F~** | durch Verbesserung der AlbisGetCaveZeile Funktion, enormen Geschwindig-keitszuwachs der Abarbeitung erreicht von >5s runter auf 0.9s, Abrechnungs-bedingung 'Mindestalter 35 Jahre' wird überprüft |
| **13.12.2018** | **F+** | SkriptIcon ist in das Skript integriert                      |
| **07.01.2019** | **F~** | Zuverlässigkeit erhöht, Aufruf der Formulare erfolgt jetzt per Menubefehl nicht mehr per internem Albismakro, dafür erscheinen zwei Fenster weniger und das Erstellen der Dokumentation ist deutlich schneller |
|                | **F+** | das Skript nutzt jetzt auch einen Hook um störende Fenster zu entfernen und um erkennen zu können |
| **28.03.2019** | **F~** | das Hautkrebsscreening-Formular wurde zu diesem Quartal verändert, Skript wurde entsprechend angepasst, Code wurde optimiert und fehlerfreies RPA! 100 Formulare ohne Unterbrechung! V1.85 |
| **29.03.2019** | **F+** | Fenster das der Patient heute Geburtstag hat wird automatisch geschlossen V1.86 |
| **24.06.2019** | **F+** | Hautkrebsscreening-Formular - alle Karzinom-Arten werden auf 'nein gesetzt |
| **25.06.2019** | **F~** | Code verkürzt, Code optimiert für mehr Geschwindigkeit, vermehrter Einsatz von RegEx-Befehle für flexiblere Stringvergleiche V1.88 |

<br>

![](Docs/Trenner_klein.png)

## ![SonoCapture.png](Docs/Icons/SonoCapture.png) SonoCapture

| Datum          | Teil    | Beschreibung                                                 |
| -------------- | ------- | ------------------------------------------------------------ |
| **23.02.2018** | **M+**  | **SonoCapture hinzugefügt**, eigenes Modul für die Aufnahme der Sonobilder mit sofortiger automatischer Eintragung mit korrektem Bezeichner in die Patientenakte. **ERSATZ**: für meine *unzuverlässige Praxisarchivierungs-software (CGM Praxisarchiv)*  - der Start dauert mir immer zu lange, läßt sich seit Monaten eh nicht starten, kann mir die Daten nicht mal mehr ansehen jetzt, der Videoinput über das Capturemodul der Software funktionierte mal und mal nicht. Zudem muss das Praxisarchiv jedesmal aktiviert werden wenn ein Rechner erneuert oder ausgetauscht wurde. |
| **14.03.2018** | **MF+** | **Sono Capture V0.53:** nicht konstant erkanntes Irfanview Fenster (schlecht gelöste Abfrage) korrigiert, ein Nutzer kann jetzt in der Praxomat.ini sein bevorzugtes Preview-Programm einstellen oder er schreibt "No" ein, wenn er bei Albis die Previewfunktion ausgestellt hat,	Auflösung in der GUI Preview erhöht, PAL-Auflösung (720x576 Pixel) |
| **13.04.2018** | **MF**+ | **Sono Capture V06.0**: Erinnerungsfenster integriert, das Ultraschallgerät muss unbedingt vor dem Zugriff auf den WinCap-Treiber eingeschaltet sein, sonst kommt es zu einem nicht behebbaren Treiberkonflikt, so daß sich das Skript nicht beenden läßt. Meist ist ein Neustart des Computer dann auch noch notwendig. |
| **14.04.2018** | **M~**  | das Skript kann jetzt mit der Windows dpi scale Funktion umgehen. Die Fenster sollten somit bei allen Auflösungen und Skalierungseinstellungen korrekt aussehen |
| **02.08.2018** | **F~**  | RegRead über externe Funktion                                |
| **02.12.2018** | **F~**  | Unicode DllCall der avicap32.dll - warum die ANSI DllCalls auch mal funktionierten, ich weiß es nicht. So ist es allerdings stabiler lauffähig. |
| **13.12.2018** | **F+** | SkriptIcon ist in das Skript integriert     


<br>

![](Docs/Trenner_klein.png)

## ![ScanPool.png](Docs/Icons/ScanPool.png) ScanPool

|     Datum      |  Teil  | Beschreibung                                                 |
| :------------: | :----: | :----------------------------------------------------------- |
| **02.08.2018** | **F+** | Incrementelle Suchfunktion +0.5                              |
| **09.08.2018** | **F+** | Statustext nach Einsortierung in die Akte auffrischen - Anzahl Dokumente und Seiten |
| **11.08.2018** | **F~** | Signiervorgang beschleunigt                                  |
| **02.08.2018** | **F~** | indizierte und in die Akte einsortierte Dateien müssen auch aus dem files Array elöscht werden |
| **04.08.2018** | **F~** | vor dem Start von FoxitReader prüfen ob ein BEfund überhaupt noch existiert - sonst muss man ein leeres FoxitReader fenster schließen |
| **04.08.2018** | **F+** | MsgBox Infobox oder Fehlerangabe sollte wie ein Childwindow den Zugriff auf die Gui sperren |
| **12.08.2018** | **F+** | öffnen aller Befunde eines Patienten                         |
| **02.08.2018** | **F+** | Edit1 löschen wenn alle Befunde Button gedrückt wurde        |
| **09.08.2018** | **F+** | während des Signiervorganges ein Progress Fenster anzeigen und Interaktionen mit dem ScanPool Hauptfenster verhindern (Sperren des UserInputs, Abbruch per Esc - Hinweis geben!) |
| **11.08.2018** | **F~** | Aktualisieren der Listview korrigiert                        |
| **13.08.2018** | **F+** | Schließen des FoxitReader nachdem der Befund in die Akte einsortiert wurde und schließen der Patientenakte |
| **12.08.2018** | **F+** | beim Öffnen einer Akte nimmt er immer an das diese noch nicht geöffnet wurde |
| **21.08.2018** | **F+** | ContextMenu direktes Einsortieren auch ohne signieren        |
| **23.08.2018** | **F+** | PDF Preview!!                                                |
| **23.08.2018** | **F+** | das Passwort wird verschlüsselt in einer Variable gehalten   |
| **23.08.2018** | **F~** | PdfIndex.txt nach jedem Löschvorgang neu erstellen           |
| **25.08.2018** | **F~** | beim Signieren zur 1.Seite schalten                          |
| **25.08.2018** | **F+** | Erkennen das eine Datei signiert wurde und nicht mehr verändert werden kann |
| **03.09.2018** | **F+** | Annäherungsfunktion zur Namenssuche , Sift3 Funktion gleiche öffnen +0.5 - das war Teil1 |
| **28.09.2018** | **F+** | Pdf blättern mit dem Mausrad funktioniert super              |
| **31.10.2018** | **F~** | Toleranz-Inkrement pro Suchvorgang (ein Loop)  zur FindText-Funktion hinzugefügt, so daß eine 100%-ige Übereinstimmung nicht immer notwendig ist |
| **14.11.2018** | **F+** | fokusfreies Scrollen der Befund Gui ermöglicht               |
| **22.11.2018** | **F~** | Korrektur der Fensterpositionierung vorgenommen, besseres berechnen, Hook zur Erkennung einer sich öffnenden Praxomat-Gui erstellt |
| **23.11.2018** | **F+** | *Fuzzy search* im Hauptfenster bei der Suche nach einer Datei |
| **28.11.2018** | **F~** | sehr viele Fehlerkorrekturen                                 |
| **30.11.2018** | **F+** | das Vorschaubild läßt sich jetzt in 90° Schritten drehen     |
| **13.12.2018** | **F+** | SkriptIcon ist in das Skript integriert                      |
| **15.01.2019** | **F~** | viele Fehler behoben                                         |
| **22.01.2019** | **F~** | ContextMenu geht wieder sicher, Pdf Vorschau verschwindet wenn kein Befund ausgewählt wurde, Statuszeile optimiert |
| **01.04.2019** | **F~** | Fehlerbehebung in den Funktionen Einsortieren der Befunde, Öffnen einer Patientenakte, Anzeigen aller Befunde (0.985) |
| **02.04.2019** | **F~** | Befund desselben Patienten sind jetzt farblich gleich hinterlegt in der Auflistung (0.986) |
| **05.04.2019** | **F+** | Teile des Signiervorganges werden jetzt durch einen **Fensterhook** gesteuert. Dieser Methode ist deutlich zuverlässiger und wesentlich schneller! |
|                | **F+** | Der **Foxitreader** muss nicht mehr extra heruntergeladen und auch nicht installiert werden. Im \include befindet sich eine partable Version die sich von allen Clients aus starten läßt! (0.987) |
| **12.04.2019** | **F~** | teilweise Reprogrammierung des Signierprozesse für eine deutlich verbesserte Zuverlässigkeit des Ablaufs |
|                | **F+** | Befundordner ist eine Datenbank (V0.9.86)                    |
|                | **F~** | Rechtsklick nach einem Befund übertragen Vorgang funktioniert manchmal nicht mehr im Listviewfenster sondern nur in anderem Teil der Gui |
|                | **F~** | Foxitdialoge 	- Speichern unter - riesiges Problem - er hat dauernd die falschen Ordner offen oder nimmt sich irgendeinen anderen Dateinamen |
|                | **F~** | Vorschau 		- Listview wird zu schmal dargestellt, erst ein manuelles Resize der Gui korrigiert es |
| **02.05.2019** | **F~** | PdfReader_Close(), FoxitReader_SignaturSetzen(), diverse Fehler behoben, mehrere Dokumente aufeinmal lassen sich wieder übertragen und signieren |
| **10.05.2019** | **F+** | PdfAddMetaDataInfo() -  (V0.9.88)                            |
| **14.05.2019** | **F~** | Fuzzy Matching von Patientenfenster Albis und Pdf Dateinnamen funktioniert jetzt super! |
| **15.05.2019** | **F+** | Funktionen zur Behandlung des Speichern unter Systemdialoges in Windows10, **FoxitReader_ConfirmSaveAs()**, **FoxitReader_CloseSaveAs()**, **FoxitReader_ExceptionDialog()** |
|                | **F~** | Albis Pdf's	- automatisches Schließen der von Albis nach import geöffneten Foxitreaderfenster (V0.9.89) |
| **18.05.2019** | **F+** | **ExtractNamesFromFileName()** - entfernt Dr. und Prof.      |
| **28.05.2019** | **F~** | Verbesserung beim Abfangen und Bearbeiten der Foxitdialoge, diese werden jetzt zu 100% richtig erkannt und abgehandelt |
<br>

![](Docs/Trenner_klein.png)

## ![Monet](Docs/Icons/Monet.png) MoNet - (Mo)nitor for your (Net)work

| Datum          | Teil   | Beschreibung                                                 |
| -------------- | ------ | ------------------------------------------------------------ |
| **09.03.2018** | **M+** | **MoNet**: Monitor for your Network (nicht nur für Ärzte) - Anzeige des Status der Computer in der Praxis, auch als dezentrales Tool gedacht um von jedem Punkt im LAN aus alle Computer herunterzufahren und auch über WakeOnLan am nächsten morgen hochzufahren. Möglicherweise später auch Zeitgesteuertes hochfahren, weitere Funktionen wie LAN-Minichat, 	versenden von Einzel-Kommandos an die Computer oder aber auch ferngesteuertes Script oder Programm starten sind angedacht. |
| **13.12.2018** | **F+** | SkriptIcon ist in das Skript integriert                      |



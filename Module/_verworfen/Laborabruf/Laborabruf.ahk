;-------------------------------------------------------------------------------------------------------- Modul Laborabruf V0.75 -------------------------------------------------------------------------------------------------------
;------------------------------------------------------------------------------------------------ Addendum für ALBIS on WINDOWS -------------------------------------------------------------------------------------------------
;---------------------------------------------------------------------------------------- written by Ixiko -this version is from 12.12.2018 -------------------------------------------------------------------------------------------
;------------------------------------------------------------------------------ please report errors and suggestions to me: Ixiko@mailbox.org -----------------------------------------------------------------------------------
;--------------------------------------------------------------------------- use subject: "Addendum" so that you don't end up in the spam folder --------------------------------------------------------------------------------
;------------------------------------------------------------------------------------ GNU Lizenz - can be found in main directory  - 2017 -----------------------------------------------------------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

; Automatisierung				: automatisierter Datendownload Ihrer Labordaten auf Ihren Rechner, Übertragung ins Laborbuch und selbstständiges Einsortieren
;										  die Zeitpunkte des Abrufes sind frei einstellbar (per Eintrag in der Addendum.ini und Start über das Addendum.ahk Skript)
; erweiterte Informationen	: Notiz in jede Patientenakte, ob End- oder Teilbefund, es notiert auch die Probennummer, Sie ersparen sich das Suchen im Laborbuch!
; Smartphoneprotokoll		: über Einstellungen in der Addendum.ini können Sie sich den Status des Abrufes auf Ihr Handy schicken lassen
; ParameterAlarm				: mit einer weiteren Einstellung ist es möglich sich hochgradig auffällige pathologische Werte inkl. der Telefonnummer des Patienten zuschicken zu lassen
;									  	  als default ist +/- 30% als Überschreitung des oberen oder unteren Grenzwertes eines Parameter eingestellt

/* dieses Skript ist für einen komplett automatisierten/unabhängigen Ablauf entwickelt worden ### Bitte unbedingt LESEN! #### ---->

	ich habe versucht das bei unerwarteten Ereignissen während des Skriptablaufes (z.B. Skript wartet auf ein Fenster das nicht kommen kann, weil ein anderes Fenster sich öffnete)
	das Skript beendet wird um ein weiterarbeiten mit Albis z.B. am nächsten Sprechstundentag nicht zu behindern. Zur Unterstützung der Fehlersuche wird ein Screenshot
	vom Bildschirm erstellt und im Skriptordner gespeichert. Da ich die App Telegram benutze, ist deshalb eine Benachrichtigungsfunktion integriert, damit ich/Sie erkennen können,
	das ein Abruch aufgrund von Skriptfehlern/Problemen erfolgt ist. Nach einem Abbruch sollte das Skript nach einem vorgegebenen Intervall erneut gestartet werden

 Der Aufruf benötigt ein laufendes AlbisProgramm. Falls Albis nicht gestartet ist, sollte es Albis ohne Nachfrage starten können!
 Dazu muss in der Addendum.ini für den aufrufenden Client ein Username und Passwort eingetragen sein.

 #################################### ! BEACHTEN SIE FOLGENDES ! ################################################
 Autohotkey ist eine Skriptsprache. Aufgrund des nicht verschlüsselbaren Programmcodes nutzt es nichts die Passwörter verschlüsselt abzulegen, diese wären dennoch auslesbar.
 Sorgen Sie dafür das kein Unbefugter Zugriff auf das Addendumverzeichnis bekommt!

	1. nutzen sie eine Firewall! 	oder
	2. nutzen Sie am besten einen Proxyserver in Ihrer Praxis für eingehende und ausgehende Internetverbindungen! Teilen Sie Ihr Praxisnetzwerk physisch in 2 unterschiedliche Bereiche!
	3. Schützen Sie ihre Computer auch vor dem Zugriff Unbefugter, die möglicherweise als Patienten gerade in ihrem Sprechzimmer sitzen!
	4. Protokollieren Sie auch den Zugriff auf Ihre Resourcen. Microsoft Windows hält dafür Einstellungen und/oder Tools bereit!
	5. Ein Kompilieren der Skripte ist sicherlich eine Möglichkeit. Halbwegs versierte kommen allerdings trotzdem an die Daten


 ------------------------------------------------------------------------------- ZUM SKRIPTABLAUF ------------------------------------------------------------------------------------------------------------
																					  Abbruch des Skriptes mit Strg+Shift+ä
 --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 Dies ist das komplexeste Skript das ich bisher für die Sammlung geschrieben habe. Ich habe mir angewöhnt die Skriptteile zu nummerieren und zu bezeichnen.
 Auf diese Weise halte ich es übersichtlicher und durch Codefolding des Editors (bei mir Scite) erhalte ich schneller Zugriff auf Teile des Skriptes.
 Da keines der Skripte ohne Ihren Eingriff laufen würde, sollten Sie sich ruhig durch die Skriptsammlung wühlen. Sie müssen nicht unbedingt Autohotkey beherrschen,
 aber Sie sollten zumindestens den Syntax verstehen.

 I. das Skript läßt sich manuell starten und ist in der Lage zwischen einem Aufruf durch ein anderes Skript oder auch dem Windows Taskplaner zu unterscheiden
    hierzu muss dem Skriptaufruf nur ein Text übergeben werden per cmdline oder AHK Skript z.B. als
			Run, PfadzumSkript\Laborabruf.ahk

a) Start mit Öffnen des WebClienten - dies ist das Programm meines Labors zum downloaden der Daten über das Internet
		-falls Ihre Praxis einen anderen WebClienten nutzt - müssen Sie an dieser Stelle den Aufruf und Ablauf als Skript erstellen
		 alle folgenden Labels sind Albisfunktionen und müssen nicht geändert werden
b) Datenholen: ruft das Menu 'Extern/Labor/Daten holen' auf,

																																																																					 */

/*																								------------------------ TODO LISTE ------------------------																									|
																																																																					|
	~~~~ dieses Skript muss absolut fehlerfrei laufen! Ich brauche ein Handling für unerwartet auftauchende Fenster oder dann Benachrichtigung per Telegram! ~~~~							|
																																																																					|
	1. Problem beim Überprüfen des Albisstatus - blockiert?																																																|
	2. Telegram - Versand des Abrufstatus,	insbesondere wenn Fehler aufgetreten sind																					+V0.05														|
	3. LDT Parser Integration - Versand der Laborwerte																																	+V0.2															|
	4. Hinweis Gui - starten des Skriptes ankündigen																																	+V0.05														|
	5. Block UserInput - solange das Skript läuft, allerdings muss irgendwie doch ein manueller Eingriff möglich sein										+V0.01														|
	6. Oeffnen Vorbereiten - Checkboxhaken werden nicht richtig gesetzt mal wieder , wie überprüfe ich das?																													|
	*erledigt	7. ----/Laborbuch Label - anstatt FindWindow() die FindChildWindow() Funktion nutzen!/ 																															|
	21.05.2018	F+	optische und funktionelle Verbesserung der Verlaufs GUI, Layoutverbesserungen und Kürzen von Quell-Code in 5. Laborabruf, 														|
																																																																					 */

;{1. Scripteinstellungen

	#NoEnv
	#SingleInstance
	SetBatchlines, -1
	;SetWinDelay, -1
	SetTitleMatchMode, 2
	DetectHiddenWindows, On
	DetectHiddenText, On
	FileEncoding, UTF-8
	OnExit, JetztAberRaus

	global code:= []
	debug = 0				;true = 1

	hIBitmap:= Create_LaborAbruf_ico(true)
	Menu Tray, Icon, hIcon:  %hIBitmap%

/*
	If (debug) {
		FileRead, scriptfile, %A_ScriptFullPath%
		Loop, Parse, scriptfile, `n
		{
			code[A_Index]:= A_LoopField
		}
	}
*/

	;{--Minimieren des SciteWindow falls das Skript zu Testzwecken aus Scite gestartet wird
		IfWinExist, ahk_class SciTEWindow
		{
					WinActivate
					 WinMinimize
		}
	;}

;}

;{2. Includes

	#include %A_ScriptDir%\..\..\..\include\AddendumFunctions.ahk
	#include %A_ScriptDir%\..\..\..\include\ini.ahk
	#include %A_ScriptDir%\..\..\..\include\GDIP_All.ahk
	#include %A_ScriptDir%\..\..\..\include\ACC.ahk
	#include %A_ScriptDir%\..\..\..\include\COM.ahk
	#include %A_ScriptDir%\..\..\..\include\Telegram.ahk
	#include %A_ScriptDir%\..\..\..\include\JSON.ahk

;}

;{3. Variablen und Konstanten

   ;{ 3.a.	Variablen

	;-----------------------------------------------------------------------------------------------------------------------------------------------------------
	global AddendumDir		:= FileOpen("C:\albiswin\AddendumDir","r").Read()
	;-----------------------------------------------------------------------------------------------------------------------------------------------------------

	PraxoINI = %AddendumDir%\Addendum.ini

	Autostart = %1%																							;--für Unterscheidung Autostart oder manueller Start
	Heute:= A_DD . "`." . A_MM . "`." . A_YYYY
	Script:= A_ScriptName
	CompName:= A_ComputerName
	StringReplace, CompName, CompName, -,, All

	KeineDaten:= 0 												;flag
	Protokoll(PraxoIni . "`, " . Autostart . "`, " . Heute . "`, " . Script . "`, " . Compname, 2)

	;}

	;{ 3.b. 	hinterlegen Sie den Namen des Client-Rechners der dieses Skript ausführen darf, ich denke das es pro Praxis/Arzt nur einen Arbeitsplatz für den Abruf der Labordaten gibt

	IniRead, OnlyRunOnClient, %PraxoINI%, Labor_Abrufen, OnlyRunOnClient, %CompName%
	If (OnlyRunOnClient <> CompName) {
			MsgBox, 1, Laborabruf, Dieses Skript ist nicht für diesen Computer vorgesehen`nEine Ausführung könnte zu einem Absturz von Albis führen.`nDas Skript wird beendet!, 30
			ExitApp
	}

;}

	;{ 3.c. 	Einstellungen für dieses Skript werden hier eingelesen


		;das Standardverzeichnis indem ihr Labor die .ldt-Datei für die Weiterverarbeitung durch Albis ablegt
	IniRead, LDTDir, %PraxoINI%, Labor_abrufen, LDTDirectory, C:\Labor

	;die Größe und Position der VerlaufsGui werden beim Labor VerlaufGui ausgelesen und definiert

		;der Name Ihres Labors wie Sie ihn im Menu Extern\Labor\Daten holen nach dem Aufruf entnehmen können
	IniRead, LaborName, %PraxoINI%, Labor_Abrufen, LaborName

		;hier in der Addendum.ini Datei eintragen unter welchem Kürzel eine Notiz zum Datenempfang beim Patienten eingetragen werden soll
	IniRead, Aktenkuerzel, %PraxoINI%, Labor_Abrufen, Aktenkuerzel, Labor

		;--------------------------------------------------------- Benachrichtigung per Telegram App über stark veränderte Laborwerte ------------------------------------------------------------
		; 1. für diese Einstellung müssen Sie die Telegram App (Iphone, Android, Windows, Linux...) heruntergeladen (Telegram.org) haben und sich per Telefonnummer (nur Mobiltelefon) angemeldet sein
		; 2. Sie müssen einen über die App einen Bot erstellt haben (es gibt Anleitungen dazu im Internet)
		; 3. das Bot-Token und ihre Chat-ID (an wen die Nachricht gehen soll) müssen dann noch in der Bot.ini hinterlegt werden
		;  	3.a. Vorsicht mit den hinterlegen des Bot-Token! - stellen Sie sicher das die Addendum.ini nicht von irgendwem ausgelesen werden kann, der erstellte Bot könnte dann durch fremde benutzt werden!
		;		 	ich habe noch keine gute Idee den Bot-Token zu schützen , das Senden per Telegram auf Ihr Handy oder den PC zu Hause, sollte allerhöchstens ein Computer in ihrem Netzwerk durchführen
		;			dieser sollte sehr gut vor dem Zugriff von Außen geschützt sein!
		;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		;Telegram Optionen: Ja-TelNr, Ja+TelNr oder Nein , per Default ist die Option Nein eingestellt
		;(+/-TelNr - Telefonnumer des Pat. wird wenn vorhanden mitgesendet)(+/-Crypt - per End-zu-End Verschlüsselung oder nicht - ACHTUNG: End-zu-End sendet nur auf ein eingestelltes Gerät!)

	IniRead, TGramOpt, %PraxoINI%, Labor_Abrufen, TelegramMsg, Nein
	If Instr(TGramOpt, "Ja") {
			IniRead, BotName, %PraxoINI%, Telegram, Bot1
			IniRead, BotToken, %PraxoINI%, Telegram, %BotName%_Token
			IniRead, BotChatID, %PraxoINI%, Telegram, %BotName%_ChatID
		If Instr(TGramOpt, "+") 			;flag für Telefonnummernübertragung ja oder nein (wenn ja dann verschlüsselter Chat)
						TNr:= 1
		else
						TNr:=0
	}

		;es werden im Moment nur Blutwerte erfasst die PathGrenze (in Prozent) ober- oder unterhalb der Normgrenzen liegen - dies gilt im Moment für alle Werte,
		;die Sinnhaftigkeit einer Festlegung auf verschiedene Alarmgrenzen für unterschiedliche Laborwerte überlege ich mir noch
	IniRead, PathGrenze, %PraxoINI%, Labor_Abrufen, PathGrenze, 30

	;diesen Eintrag in der Addendum.ini nicht ändern!!!
	IniRead, LastDayUsed, %PraxoINI%, Labor_Abrufen, LastDayUsed
		If (LastDayUsed<>Heute) {
			IniWrite, %Heute%, %PraxoINI%, Labor_Abrufen, LastDayUsed
			FileAppend, `[%Heute%`]`n, %A_ScriptDir%\LaborAbrufProtokolle.txt
		}

	Protokoll(LDTDir . "`, " . LaborName . "`, " . Aktenkuerzel . "`, " . TGramOpt . "`, " . PathGrenze, 2)

;}

	;{ 3.d.	WM_Command Konstanten für Albis

	Laborbuch:= 34162
	AlleUebertragen:= 34157	;geht nur bei geöffnetem Laborbuch
	Datenholen:= 32965
	LaborAnrufen:= 6016
	global Patientenfenster:=	32886

	;}

;}

;{4. AutoExecute-Bereich

	MsgBox, 4096, Addendum für Albis on Windows, Der Start des Abrufes der Laborwerte erfolgt in 10 Sekunden.

		;Fenster für die Anzeige des Fortschrittes einblenden
	gosub VerlaufGui

		Protokoll("Abbruch des Skriptes mit Strg+Shift+#", 1)
		Protokoll("starte Fehler-Überwachungsroutine des Laborabruf Skriptes `(Verzögerung 10min`)", 1)

		;Abbruchroutine wird nach dieser Zeit gestartet. Dies ist notwendig falls unerwartete Fehler auftauchen.
	SetTimer, RunTimeError, -600000			;10min - Albis braucht manchmal lange um zu starten
		Protokoll("Die Fehler-Überwachung ist gestartet", 1)

		;erhalten des Albis Windowhandles
	global AlbisWinID:= GetAlbisWinID()
																								;If debug, WeiterOderStop(A_LineNumber)

		;startet Albis oder überprüft das Albis nicht blockiert ist, loggt sich wenn nötig in Albis ein um den Laborabruf später vorzunehmen zu können
	;	Protokoll("Prüfe ob Albis läuft und keine Fenster Eingaben blockieren", 1)
	;StarteAlbis(CompName, Login1, Login2, "Laborabruf", 1)
																								;If debug, WeiterOderStop(A_LineNumber)
	;	Protokoll("Die Überprüfung wurde mit 'ok' abgeschlossen.", 1)

		;Programmdatum auf das heutige Datum setzen, damit die Einträge in die Akte am richtigen Tag eingetragen werden
		Protokoll("Stelle Albis auf das aktuelle Programmdatum ein.", 1)
	If !(AlbisIsBlocked(AlbisWinID))
			AlbisSetzeProgrammDatum(Heute)

		Protokoll("Das aktuelle Programmdatum ist eingestellt.", 1)

		Protokoll("Starte die Laufzeitüberwachung von Albis.", 1)
		;diese Routine schaut ob sich Albis während des Vorganges aufhängt, stoppt das Skript bis zu 2min
		;und erkennt auch wenn Albis wieder funktioniert und macht an der angehaltenen Stelle weiter
	SetTimer, CheckAlbisRuntime, 100
		Protokoll("Die Laufzeitüberwachung von Albis wurde erfolgreich gestartet.", 1)
		;~~GOTO BEFEHL für Debug-Sequencen
	;goto OeffnenVorbereiten

;}

;{5. Laborabruf

		Protokoll("<Start eines neuen Laborabrufes>", 2)

WebClient:
;{ a.) MCS vianova infoBox - dies ist der Anfang mit meinem Labor, hier müssen Sie evtl. eine eigene Automatisierungsroutine erstellen

			;** evtl. diese Routine als Funktion bereitstellen , damit Kollegen die eine andere Laborsoftware haben diese hier leichter einfügen können

			;Abbruchroutine wird nach dieser Zeit gestartet. Dies ist notwendig falls unerwartete Fehler auftauchen. Das Minus steht für den einmaligen Aufruf
		SetTimer, RunTimeError, -120000			;2min

		Protokoll("Starte das Programm zum Übertragen der Labordaten")

		;Aufruf der Laborsoftware - WebClient Fensters wird per WM-Command gestartet
	If !(MCSinfoBoxID:=WinExist("MCS vianova infoBox-webClient")) {
			WinActivate		, ahk_class OptoAppClass
			WinWaitActive	, ahk_class OptoAppClass
			PostMessage,0x111, %Laboranrufen%,,, ahk_id %AlbisWinID%
			WinWaitActive	, MCS vianova infoBox-webClient, ,6
	}

																													;If debug, WeiterOderStop(A_LineNumber)

		;WebClient ist geöffnet, dann Laborabruf beginnen
	If MCSinfoBoxID {
			WinActivate, ahk_id %MCSinfoBoxID%
			ButtonClass:=WinForms_GetClassNN(MCSinfoBoxID, "Button", "Abruf")
			Protokoll("ClassNN Abruf-Button: " . ButtonClass , 1)
			ControlClick, %ButtonClass%, ahk_Id %MCSinfoBoxID%
	}

		Protokoll("Warte auf das Ende der Labordatenübertragung")


	i=1

	; wartet darauf das neue Daten eingelesen wurden
	Loop {

		oAcc := Acc_Get("Object", "4.1.4.5.4.1.4.4.4", 0, "ahk_id " MCSinfoBoxID)

				If Instr(oAcc.accValue(1),"Ende Dateien abrufen") {

						If Instr(oAcc.accValue(2), "Keine Daten zum Abruf") {
									Protokoll("Es sind keine neuen Labordaten vorhanden.")
								gosub CloseWebClient
								KeineDaten:=1
									sleep 4000
								goto Datenholen
						}

						break
				}

		sleep -1

	}

		Protokoll("Die Labordatenübertragung ist abgeschlossen.")

																													;If debug, WeiterOderStop(A_LineNumber)

		gosub CloseWebClient
			sleep, 4000

	;Auch wenn dieses Fenster nicht zu schließen geht, kann das Abrufen der Laborwerte weiter gehen
	If WinExist("MCS vianova infoBox-webClient") {
		Protokoll("Das MCS vianova Fenster konnte nicht geschlossen werden.")
	}

;}

LDTSichern:
;{ b.) sichern der LDT Datei(en) für spätere Zwecke

			;Abbruchroutine wird nach dieser Zeit gestartet. Dies ist notwendig falls unerwartete Fehler auftauchen.
		SetTimer, RunTimeError, -120000			;2min

			Protokoll("Erstelle eine Sicherung der übertragenen Daten.")

		;auslesen aller Dateien im LDT-Verzeichnis über den Loop und jeweils kopieren in das Labordatenverzeichnis für spätere Verwendung, die neue
		;Datei erhält ein prefix mit dem Datum des Abrufes
	Loop, %LDTDir%\*.LDT
	{
			Protokoll("Datei gefunden: " . A_LoopFileName)
			FileCopy, %A_LoopFileFullPath%, %A_ScriptDir%\Labordaten\%Heute%-%A_LoopFileName%
	}

	If A_LastError
			Protokoll("das Kopieren der .ldt Dateien erzeugte folgenden Fehler: " . A_LastError)
	else
			Protokoll("Beim Kopiervorgang der .ldt Dateien wurde kein Fehler festgestellt.")

			;kurzes Päuschen
		sleep, 1000

;}

Datenholen:
;{ c.) auf das Menu: Extern/Labor/Daten holen zugreifen und die dann folgenden Fenster bearbeiten

	;Abbruchroutine wird nach dieser Zeit gestartet. Dies ist notwendig falls unerwartete Fehler auftauchen.
	SetTimer, RunTimeError, -120000			;2min

	;Menubefehl 'Extern/Labor/Daten holen' - siehe Hauptverzeichnis\support\_Infos zur Automatisierung - Ablaufbeschreibung.ahk (DIES ISTKEIN SKRIPT!)
	If !FindWindow("Labor auswählen", "#32770")				;falls das Fenster aus irgendwelchen Gründen schon geöffnet ist, eine kleine Abfrage
	{
			PostMessage, 0x111, %Datenholen%,,, ahk_id %AlbisWinID%
			WinWaitActive, Labor auswählen ahk_class #32770,, 3
	}
	;Es könnten zwei Fenster "Keine Datei(en) im Pfad" oder "Labor auswählen" erscheinen - die FindWindow() Funktion ist in AddendumFunctions.ahk enthalten
	Loop {

				;Fenstermöglichkeit 1 untersuchen
			If (foundWinID:= FindWindow("ALBIS", "#32770", "Keine Datei`(en`) im Pfad")) {

					Protokoll("Es waren keine LDT Dateien im Verzeichnis: " . LDTDIR . " vorhanden.")
					Protokoll("Möglicherweise wurden die Dateien manuell abgerufen.", 1)
					Protokoll("Das Programm wird im Laborbuch nach noch nicht einsortierten Befunden suchen!")
					SureControlClick("Button1", "", "", foundWinID)
					WinWaitClose, ahk_id %foundWinID%
						sleep, 2000
					goto LaborBuch

			}

				;Fenstermöglichkeit 2 untersuchen
			If (foundWinID:= FindWindow("Labor auswählen", "#32770")) {

					WinActivate, ahk_id %foundWinID%
						WinWaitActive, ahk_id %foundWinID%
							ControlSetText, Edit1, %LaborName%, ahk_id %foundWinID%
								sleep, 100

								;falls mehrere Labore zur Auswahl stehen muss hier eine Unterscheidung stattfinden
							ControlGetText, CText, Edit1, ahk_id %foundWinID%
							if (CText<>LaborName) {
									Protokoll("Im Fenster: 'Labor auswählen' konnte der Name des Labors: `'" . LaborName . "`' nicht gesetzt werden.")
									Protokoll("Das Skript wird jetzt beendet",1)
									SureControlClick("Button3", "", "", foundWinID)
										sleep, 10000
									ExitApp
							}
							else {
								SureControlClick("Button1", "", "", foundWinID)
								break
							}

			}

			sleep, 500

	}


	;Fenster Labordaten erscheint nach einem weiteren Fenster
	Protokoll("Warte auf das Fenster: 'Labordaten'.")
	Protokoll("Bitte warten Sie die Übernahme der Daten zunächst ab.",1)

	;suchen und warten auf die ID des Labordaten Fenster
	Loop {

		If (foundWinID:= FindWindow("Labordaten")) {
					WinActivate, ahk_id %foundWinID%
						WinWaitActive, ahk_id %foundWinID%, 2
							SendInput, {ALT Down}o{ALT Up}				;so lange bis ich weiß welchen Namen der OK Button des Fensters hat
							SendInput, {ENTER}
						SureControlClick("Button1", "", "", foundWinID)
							WinWaitClose, ahk_id %foundWinID%, 3
									break
		}

		If (foundWinID:= FindWindow("ALBIS", "#32770", "Keine Datei")) {
				Protokoll("Es waren keine LDT Dateien im Verzeichnis: " . LDTDIR . " vorhanden.")
				Protokoll("Möglicherweise wurden die Dateien manuell abgerufen.", 1)
				Protokoll("Starte den nächsten Schritt und rufe das Laborbuch auf.")
				SureControlClick("Button1", "", "", foundWinID)
					WinWaitClose, ahk_id %foundWinID%, 3
				goto Laborbuch
		}

			sleep, 100
	}

;}

OeffnenVorbereiten:
;{ d.) Auslesen der Einstellungen unter Optionen/Patientenfenster/NachÖffnen - Fenster automatisch geöffnet werden sollten -

	;Abbruchroutine wird nach dieser Zeit gestartet. Dies ist notwendig falls unerwartete Fehler auftauchen.
	SetTimer, RunTimeError, -120000			;2min

	;Öffnen des Einstellungsfensters Optionen/Patientenfenster
	foundWinID:= AlbisOptPatientenfenster(7)			;Tab7 = Nach Öffnen

	;für Laborabruf: Status folgender Button zwischenspeichern 				2,3,4,6,7,8,9,11,13,14,16,17,18
	;für Laborabruf: Status "On" für folgende Button zwischenspeichern 	13,17 der Rest wird ausgeschaltet

	;über alle 18 Buttons Information auslesen
	Loop 18 {

			var:= A_Index
			;wenn der Button in der Leseliste ist - die Information zu seinem Status auslesen
			If (var in 2,3,4,6,7,8,9,11,13,14,16,17,18) {

					;Speichern des Status in der Variable TButton
					ControlGet, TButton%var%, checked,, Button%var%, ahk_id %foundWinID%

					;mit dieser Liste wird der benötigte Zustand gesetzt
					If (var in 13,17) {
							Control, check,, Button%var%, ahk_id %foundWinID%
							checker.= var . ": checked`n"
					}
					else {
							Control, uncheck,, Button%var%, ahk_id %foundWinID%
							checker.= var . ": unchecked`n"
					}

			}
	}

	MsgBox, %checker%

	;Schliessen des Option/Patientenfenster
	ControlClick, Button19, ahk_id %foundWinID%
		sleep, 100

;}

Laborbuch:
;{ e.) Einsortieren der eingelesenen Befunde in die Patientenakte

		;Abbruchroutine wird nach dieser Zeit gestartet. Dies ist notwendig falls unerwartete Fehler auftauchen.
		SetTimer, RunTimeError, -120000			;2min

	;sicher gehen das das vorherige Fenster nicht noch geöffnet ist
	If FindWindow("ALBIS", "#32770", "Keine Datei", "", "", "on", "on") {
				SureControlClick("Button1", "ALBIS", "Keine Datei")
	}

	;öffnen des Laborbuches per WM_Command Befehl
	PostMessage, 0x111, %Laborbuch%,,, ahk_id %AlbisWinID%

	Protokoll("Warte auf das Öffnen des Laborbuches.",1)
	Protokoll("Falls sich das Laborbuch in 10 sec nicht öffnet,",1)
	Protokoll("rufen Sie es bitte manuell auf!",1)

	;suchen der ID des Laborbuch Fenster
	Loop {

			If (foundWinID:= FindChildWindow(AlbisWinID, "Laborbuch", "On")) {
						WinActivate, ahk_id %foundWinID%
							WinWaitActive, ahk_id %foundWinID%
						LaborbuchID:= foundWinID
						Protokoll("Das Laborbuch wurde/ist geöffnet.")
						Protokoll("Starte als nächstes das Einsortieren der Befunde in die Akte.", 1)
						break
			} else {
				PostMessage, 0x111, %Laborbuch%,,, ahk_id %AlbisWinID%
			}

		sleep, 500

	}


	sleep, 2000
;}

AlleUebertragen:
;{ f.) Alle Befunde werden in die Akte übertragen

	;Abbruchroutine wird nach dieser Zeit gestartet. Dies ist notwendig falls unerwartete Fehler auftauchen.
	SetTimer, RunTimeError, -120000			;2min

	;WM_Command für Alle übertragen - in Albis per Button oder F8 ausführbar
	If LaborbuchID
		PostMessage, 0x111, %AlleUebertragen%,,, ahk_id %LaborbuchID%
	else
		PostMessage, 0x111, %AlleUebertragen%,,, ahk_id %AlbisWinID%

		sleep, 5000

	Protokoll("Warte auf eine Nachfrage des Programmes - Übernahme bestätigen.")
	WinWait, ALBIS, Möchten Sie die zugeordneten Anforderungen, 8
		If (ErrorLevel=0) {
			ControlClick, Button1, ALBIS, Möchten Sie die zugeordneten Anforderungen
			Protokoll("Übernahme aller Daten bestätigt")
				sleep, 3000
		} else {
			Protokoll("Es sind keine Daten zur Übernahme vorhanden.")
			Protokoll("Das Skript wird nun beendet.", 1)
				sleep, 8000
			ExitApp
		}

	;manchmal ist dies Fenster da und bleibt offen nachdem das Skript endet
	foundWinID:= FindWindow("ALBIS", "", "Möchten Sie die zugeordneten Anforderungen")
	ControlClick, Button1, ahk_Id %foundWinID%

;}

AlleGNRuebernehmen:
;{ g.) Fenster alle GNR übernehmen

	;Abbruchroutine wird nach dieser Zeit gestartet. Dies ist notwendig falls unerwartete Fehler auftauchen.
	SetTimer, RunTimeError, -120000			;2min

	;Afz- Zähler der Anforderungen - Anforderung - Objekt zum Zwischenspeichern ausgelesener Daten
	Afz:= 0, Anforderung:= "", ticker1:=0

	;alle Gebührennummern übertragen und Daten aus dem Fenster 'GNR der Anford.-Ident' übernehmen
	;das Eingangs-Datum brauche ich für die Zeile in der die Daten eingeschrieben werden
	;PatientenID, Endbefund oder Teilbefund und die Anforderungs-Nr- das möchte ich zusätzlich noch in der Akte stehen haben
	Loop {

			foundWinID:= FindWindow("GNR", "#32770")
			If foundWinID {
						Protokoll("WinID: " foundWinID "`, WinTitel: " wintitle,1)
						WinGetTitle, wintitle, ahk_id %foundWinID%
						Afz ++																															;eins weiter zählen
						foundWinID=0
						ticker1:= A_TickCount
						ControlGetText, PName, Static6, GNR ahk_class #32770 											;Patientenname
						ControlGetText, AfNr, Static2, GNR ahk_class #32770 												;AnforderungsID
						ControlGetText, ADatum, Static4, GNR ahk_class #32770 											;Anforderungsdatum
						ControlGetText, BStatus, Static8, GNR ahk_class #32770 											;End- oder Teilbefund diese Daten möchte ich
																																						;später noch in die Akte übertragen

						foundPos:=RegExMatch(PName, "O)([0-9])\w+", PatID )
						If foundPos {
								PName:= SubStr(PName, 1, foundPos-2)
								Anforderung%Afz%:= PatID[0] "`;" PName "`;" ADatum "`;" Bstatus "`;" AfNr
								FileAppend, % Heute "`;" Anforderung%Afz% "`n", AnforderungenDaten.txt
								Protokoll(PName "`, " ADatum "`, " AfNr "`, " BStatus, 1)
						}

						errorL:=SureControlClick("Button8", "GNR der Anford ahk_class #32770")								;alle Gebührennummern übernehmen
							if (errorL) {
								Protokoll("Der Button 'Alle GNR' konnte nicht ausgelöst werden",1)
								Protokoll("Versuche es jetzt mit dem Tastenkürzel zu senden.",1)
								WinActivate, GNR der Anford ahk_class #32770
								SendInput, {ALT Down}ü{ALT Up}
								sleep, 500
							}

							WinWaitClose, GNR ahk_class #32770,, 4
						;Routine muss hier so lange warten bis ein neues GNR Fenster geöffnet ist, sonst wird das schon geöffnete dauernd eingelesen
				}

				;könnte doppelt vorhanden sein, sucht nach dem Fenster das zum Laborbuch gehört
				foundWinID:= FindWindow("ALBIS", "#32770", "", "Laborbuch", "Afx:00400000:b", "on")
				If foundWinID {
						 ControlClick, Button1, ahk_id %foundWinID%																;OK Button klicken
				}

				ticker2:= A_TickCount

			;wenn sich 15s kein neues Fenster öffnet sind wahrscheinlich alle Daten in die Karteikarte übertragen
			If  (((Ticker2-Ticker1)/1000) > 15) {

						If (Afz=0) {
								Protokoll("Es gab keine Daten zur Übernahme!")
								Protokoll("Das Skript wird nun beendet!", 1)
						}	else {
								Protokoll("Der Laborabruf und die Übernahme der Daten", 1)
								Protokoll("in die Albisdatenbank sind abgeschlossen.", 1)
								Protokoll("Übertrage noch ein paar Daten in die Akte.", 1)
						}

				sleep, 4000
				break
			}

		sleep, 1500
	}


;}

StatusEintragen:
;{ h.) Übertragen des Status [Endbefund/Teilbefund] in die Akten der Patienten

		;Abbruchroutine wird nach dieser Zeit gestartet. Dies ist notwendig falls unerwartete Fehler auftauchen.
		SetTimer, RunTimeError, -300000			;5min

	; 1: PatID[0] ;2: PName 3: ADatum 4: Bstatus 5: AfNr

	WinActivate, ahk_id %AlbisWinID%
		WinWaitActive, ahk_id %AlbisWinID%, 2

	Loop %Afz% {

		LaborZeile:= Anforderung%Afz%
		StringSplit, part, LaborZeile, `;
		Aktennotiz:= "Abnahmedatum: " . part3 . "`, Status: " . part4 . "`, AnforderungsNr.: " . part5
		Protokoll("Öffne Akte von Pat.: " . part2)
		Protokoll(AktenNotiz, 1)

		error:= AlbisAkteOeffnen(part1)
			if error {
					MsgBox, 4, Achtung, Scheinbar konnte die angeforderte Akte nicht geöffnet werden.`nSie haben jetzt die Möglichkeit die Akte manuell zu öffnen., 60
					if ErrorLevel {
							Protokoll("Die Zeit für den manuellen Vorgang zum öffnen der Patientenakte mit der ID " . part1 . " ist abgelaufen.")
							continue
					}
					IfMsgBox, Cancel
						{
							Protokoll("Der Nutzer hat den Vorgang für den Patienten mit der ID " . part1 . " abgebrochen.")
							continue
						}
			}

		error:= AlbisPrepareInput("KKk")
			if !error {

					SendInput, labor													;Kürzel Labor
						sleep, 100
					SendInput, {TAB}													;nächstes Feld
						sleep, 100
					SendInput, %Aktennotiz%										;Übergabe der Notiz
						sleep, 100
					SendInput, {TAB}{ESCAPE}									;Beenden der Eingabe
						sleep, 100
					SendInput, {LControl Down}{F4}{LControl Up}	;schließen der Patientenakte

					WinWaitClose, %part2% ahk_class OptoAppClass, 10
					If ErrorLevel
							SendMessage, 0x111, 57602,, ahk_id %AlbisWinID%					;57602 Akte schließen per WM_Command

			}

	}

If Instr(TGramOpt, "Ja") {

			If (AutoStart="") {
					message:= "manueller "
			} else {
					message:= "automatischer "
			}

			message:= "Laborabruf vom: " Heute " `n"
			message.= "Uhrzeit : " A_Hour ":" A_Min "`n"

			If (KeineDaten = 1) {
				message.= "Es lagen keine neuen Laborwerte vor.`n"
			} else if (KeineDaten = 0) {
				message.= "Es wurden Werte von " Afz " Patienten abgerufen.`n"
			} else if (KeineDaten = 2)

			Telegram_SendText(BotToken, message, BotChatID)

}

;}

EinstellungenZurueckSetzen:
;{ i.) vorgenommene Fenstereinstellungen zurücksetzen

		;Abbruchroutine wird nach dieser Zeit gestartet. Dies ist notwendig falls unerwartete Fehler auftauchen.
		SetTimer, RunTimeError, -120000			;2min

	foundWinID:= AlbisOptPatientenfenster(7)

	;alle 18 Buttons zurücksetzen
	Loop 18 {

			;wenn der Button in der Leseliste ist - die Information zu seinem Status auslesen
			If A_Index in 2,3,4,6,7,8,9,11,13,14,16,17,18
			{

					if TButton%A_Index% = 1
							Control, check,, Button%A_Index%, ahk_id %foundWinID%
					else
							Control, uncheck,, Button%A_Index%, ahk_id %foundWinID%

			}

	}
;}

Benachrichtigung:
;{ j.) Benachrichtigung per Telegram(dies ändern wenn Sie es per Mail versenden möchten - nicht implementiert bisher)




;}


ExitApp


^+ä::ExitApp

return

;}

;{6. AlbisRuntime Kontrolle
CheckAlbisRuntime:

	ListLines, off
	AlbisWinID:= GetAlbisWinID()

	if DllCall("IsHungAppWindow", "UInt", AlbisWinID) {

			Protokoll("Der Albis-Prozeß scheint sich aufgehängt zu haben.",1)
			Protokoll("Wenn Albis in 2 Minuten weiterhin nicht funktioniert wird dieses Skript beendet",1)
			SetTimer, CheckAlbisRuntime, off
			TimeNow:= A_TickCount
			Protokoll("Der Albis-Prozeß hat sich aufgehängt.",2)

			;hier wartet das Skript darauf das Albis wieder funktioniert
			Loop
			{
							sleep, 200
						ZeitVorbei:= Round(A_TickCount - TimeNow)						;Millisekunden seitdem der Timer gestartet wurde
						Sekunden := (ZeitVorbei//1000)												;Millisekunden in Sekunden

						;entweder Albis geht wieder dann versucht er es nochmal
						if !DllCall("IsHungAppWindow", "UInt", AlbisWinID) {
								Protokoll("Der Albis-Prozeß arbeitet nach " . Sekunden . " wieder.")
								Protokoll("Das Skript versucht seine Arbeit wieder aufzunehmen.")
								SetTimer, CheckAlbisRuntime, On									;Timer wieder einschalten
								goto weitermachen
						}

						;oder das Skript muss hier abgebrochen werden
						if (Sekunden > 120) {
								Protokoll("Der Albis-Prozeß arbeitet nach 2 Minuten immer noch nicht.")
								Protokoll("Das Skript wird nun beendet.", 1)
									sleep, 10000
								ExitApp
						}

			} ;Ende Loop

		}

weitermachen:
listlines, on
return
;}

;{7. Verlaufs Gui

VerlaufGui:

	;{ GUI Variablen, IniRead, Festlegung, Berechnung

		Sysget, Mon, Monitor, 1

			;die Größe und Position der VerlaufsGui als X,Y, Breite, Höhe
			GSizeDef:= "1000|100|650|876"
		IniRead, GSize, %PraxoINI%, Labor_Abrufen, Gui, %GSizeDef%

		If (GSize = "Error") or (GSize = "|||") {							;Zuordnen der Fensterposition
			ACw:= 650																;Breite des Verlauf Gui
			ACx:= Mon1Right - ACw - 80									;x Position
			ACy:= 100																;y Position
			ACh:= Mon1Bottom - 150											;Höhe der Gui
		} else If Instr(GSize,"|") {
			StringSplit, GSize, GSize, |
				ACx:=GSize1, ACy:=GSize2, ACw:=GSize3, ACh:=GSize4
		}

		;Festlegen der minimalen und maximalen Breite und Höhe
		minw:= 400, minh:= 500, maxw:= Floor(Mon1Right/3), maxh:= Mon1Bottom - 30

		;Größe des ListView-Controls
		LVw:= ACw-30, LVh:= ACh-50

	;}

	Gui, AC1:NEW
	Gui, AC1:Font, S9 CDefault, Futura Bk Bt
	Gui, AC1:Margin,5,5
	Gui, AC1:+LastFound +AlwaysOnTop +Resize +HwndhAC1	MinSize%minw%x%minh% MaxSize%maxw%x%maxh%
	Gui, AC1:Add, Listview, xm ym w%LVw% h%LVh% r40 vDasLV BackgroundAAAAFF Grid NoSort , Laborabruf Fortschritt... 					;
	Gui, AC1:Add, Button, x0 y0 gAbfrage vSPause , Skript pausieren
	Gui, AC1:Add, Button, x180 y300 gAbfrage vSAbbruch, Skript abbrechen
		GuiControlGet, Abb, Pos, SAbbruch

	Gui, AC1:Show, x%ACx% y%ACy% w%ACw% h%ACh% , Labor abrufen - Addendum für Albis on Windows
	Gui, AC1:Default

	LV_ModifyCol(0,0)
	LV_ModifyCol(1,650)


return

;{					Abfrage

Abfrage:

	ctrl:= A_GuiControl
	if ctrl = SPause
			Pause, Toggle
	if ctrl = SAbbruch
			ExitApp

return

;}

;{					GuiSize

AC1GuiSize:

	;The window has been minimized.  No action needed.
	if ErrorLevel = 1
		return

	;Otherwise, the window has been resized or maximized. Resize the Edit control to match.
	Critical

	WinGetPos, wax, way, waw, wah, ahk_id %hAC1%
		position:= wax . "|" . way . "|" . waw . "|" . wah

	hNew:= A_GuiHeight - 55
	wNew:= A_GuiWidth - 10

	xAbb:= A_GuiWidth - AbbW - 20
	yAbb:= hNew + 20

	GuiControl, Move, DasLV, w%wNew% h%hNew%
	GuiControl, Move, SPause, x20 y%yAbb%
	GuiControl, Move, SAbbruch, x%xAbb% y%yAbb%
	;GuiControl,  BackgroundAAAAFF Grid, MeineListView
	WinSet, ReDraw,, ahk_id %hAC1%

	Critical, Off

return

;}

;{					GuiClose

AC1GuiClose:
CloseVerlauf:

	;Verlaufs Gui beenden
	WinGetPos, wax, way, waw, wah, ahk_id %hAC1%
	position:= wax . "|" . way . "|" . waw . "|" . wah
	if ((wax<>"") AND (way<>"") AND (waw<>"") AND (wah<>"")) {
					If (wax >-1 and way >-1)
							IniWrite, %position%, %PraxoINI%, Labor_Abrufen, Gui
		}

;}

;}

;{8. Hinweis Gui - zeigt an das der Laborabruf startet, damit niemand während dessen mit Albis arbeitet
HinweisGui:

	;Gui Hw1: New


return

;}

;{9. Wichtige Labels beim Beenden des Skriptes
RunTimeError:

	Protokoll("Das Skript wird beendet da ein Funktionsaufruf zu lange gedauert hat.")
	Protokoll("Den Screenshot wird finden sie im SkriptOrdner.",1)

	;Screenshoterstellung machen, da Skript ja unbeaufsichtigt laufen soll um eine Fehlersuche zu vereinfachen
	pToken := Gdip_Startup()
	raster := 0x40000000 + 0x00CC0020 ; get layered windows

	;Monitorgröße bestimmen
	Sysget, MonitorInfo, Monitor, 1
	sX := MonitorInfoLeft, sY := MonitorInfoTop
	sW := MonitorInfoRight - MonitorInfoLeft
	sH := MonitorInfoBottom - MonitorInfoTop
	screen:= sX . "|" . sY . "|" . sW . "|" . sH

	Zeitstempel:=  A_DD . A_MM .  A_YYYY . A_Hour . A_Min . A_Sec

	outfile = %A_ScriptDir%\Screenshot_%ZeitStempel%.jpg

	pBitmap := Gdip_BitmapFromScreen(screen,raster)

	Gdip_SetBitmapToClipboard(pBitmap)
	Gdip_SaveBitmapToFile(pBitmap, outfile)
	Gdip_DisposeImage(pBitmap)
	Gdip_Shutdown(pToken)


JetztAberRaus:

	Protokoll("<Ende des Laborabrufes>")

		sleep, 6000

	IfWinExist, ahk_class SciTEWindow
	{
				WinActivate, ahk_class SciTEWindow
					WinMaximize,  ahk_class SciTEWindow
	}

ExitApp
;}

;{10. Funktionen
CloseWebClient: ;{
	WinActivate, ahk_id %MCSinfoBoxID%
		ButtonClass:=WinForms_GetClassNN(MCSinfoBoxID, "Button", "S&chließen")
	ControlClick, %ButtonClass%, ahk_Id %MCSinfoBoxID%
return
;}

;kleine Funktion für kurze Protokollnotizen- eine Funktion fürs Schreiben in eine Datei und in das Verlaufs-GUI
Protokoll(note="empty note", t=0) {

	;t=0 schreiben in beide Ausgaben, t=1 nur Verlaufs-Gui, t=2 nur Protokollausgabe

	Heute:= A_DD . "`." . A_MM . "`." . A_YYYY
	Uhrzeit = %A_Hour%`:%A_Min%`:%A_Sec%
	Schreibsrein = %Heute%`,%Uhrzeit%`: %note%`n
	destfile = %A_ScriptDir%\LaborAbrufProtokolle.txt

	If (t=2 or t=0)
			FileAppend, %Schreibsrein%, %destfile%
	If (t=1 or t=0)
			LV_Add("", note)

}

WeiterOderStop(line) {

	PF:=0				; flag

	;a. sucht bis zu 30 Zeilen rückwärts nach dem letztem Kommentar (LLi-LastLines index)
	Loop, 30 {
		LLi:= line - A_Index
		If Instr(code[LLi], ";")
					break
		;kein Kommentar gefunden dann auch hier Abbruch, aber dann auch überspringen der Suche nach der Abschnittsbezeichnung
		If Instr(code[LLi], ";{") {
					PF:=1
					break
		}
	}

	lastCLines:="Der zuletzt abgearbeitete Codeabschnitt:`n`n" . LLi . ": " . Trim(code[LLi]) . "`n"
	LAi:= LLi			;für Teil c

	;b. anfügen der nächsten 5 Zeilen Code um besser erkennen zu können wo das Programm gerade ist
	Loop {
		LLi++
		If !code[LLi]=""
			lastCLines.= LLi . ": " . Trim(code[LLi]) . "`n"
	}

	parts=0			;flag für die Suche

	If (PF=0) {

		;c. sucht jetzt rückwärts vom letzten Kommentar nach der letzten Bezeichnung des Teilabschnittes
		Loop {

			LAi --
			If (pos1:=Instr(code[LAi], ";{")) {
						If (pos2:=Instr(code[LAi], "`.`)")) {		; ;{ x.) bezeichnet einen Unterabschnitt
						}
			}

		}
	}


	MsgBox, 4, Weiter oder Stop?, %lastlines%`n`nDeine Entscheidung!
	IfMsgBox, Yes
		return

	ExitApp
}
;}

Create_LaborAbruf_ico(NewHandle := False) {
Static hBitmap := Create_LaborAbruf_ico()
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "AAABAAEAMDAAAAEAGACoHAAAFgAAACgAAAAwAAAAYAAAAAEAGAAAAAAAAAAAAGAAAABgAAAAAAAAAAAAAAAAAAAAAAD/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/7nZH/oo//oo//oo//oo//oo//oo//oo8AAAAAAAAAAAD/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo8AAAD/oo//oo//oo//oo//oo//oo9jPSxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdILBr/oo/5nov6n4z/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/6n4z/oo//oo9ILBpCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9jPSxCKBdCKBdCKBdCKBdCKBdCKBdCKBdLLh35nov/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9LLh1CKBdCKBdCKBdCKBdCKBdCKBdCKBdkPi3/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf0m4j/oo//oo//oo//oo/+oY77n4z6n4z6n4z6n4z6n4z6n4z6n4z6n4z6n4z6n4z6n4z6n4z6n4z6n4z6n4z7n4z+oY7/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo/+oY7/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo/9oY7/oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo/+oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/6nJD/oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/9oY7/oo//oo/5noxCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo/zqABCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo/8oI1CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo/zqABCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo/0m4hCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo/+oY5CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo/zqABCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf7oI3/oo/+oY5GKxlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf9oY7/oo//oo/zqABCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo9DKRhCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf7oI3/oo//oo9DKRhCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdRMiD/oo//oo/zqABCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo/8oI1CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo/zqABFKhlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/+oY7/oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdEKhn/oo//oo/zqABCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo/7n4xCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo/zqABHKxpCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/9oY7+oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf4nov/oojzqABCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo/9oY5CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo/7pFpJLBxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/7oI39oY7/oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf6n4z7pFrzqABCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/8oI3zqABOLx9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdLLh34nov/oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf4nov7pFrzqABCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo/5noxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf3nYr/ooj/oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdEKhn/oo/+oY76n4xCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/7pFrzqABCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo/5notCKBdCKBdCKBdCKBdCKBdCKBdCKBf6n4z7pFr/oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf5nov8oI36n4xCKBdCKBdCKBdCKBdCKBf/oo/7pFr/oohCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf4nYr/oo/5notCKBdCKBdCKBdCKBdCKBf4nYr+o3v/oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf9oY7+oY7/oo9CKBdCKBdCKBf/oo/6n4z5notCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf9oY7/oo//oo9CKBdCKBdCKBf/oo/6n4z9oY5CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf6n4z/oo//oo9CKBdCKBdCKBf/oo/6n4z9oY5CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf6n4z/oo//oo9CKBdCKBdCKBf/oo/6n4z+oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf6n4z/oo//oo9CKBdCKBdCKBf/oo/6n4z9oY5CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf6n4z/oo//oo9CKBdCKBdCKBf/oo/6n4z9oY5CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo/7nY//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf9oY7/oo//oo9CKBdCKBdCKBf/oo/6n4z9oY5CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/7nZD/oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf9oY7/oo//oo9CKBdCKBdCKBf/oo/6n4z5notCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf9oY7+oo//oo//oo+naVf/oo//oo//oo/5notCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9iPCtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBddOin9oY7/oo//oo//oo//oo//oo//oo//oo/4nYtcOShCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdjPSz/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf8oI35nov2nIn2nIn2nIn2nIn2nIn2nIn2nIn2nIn2nInul4T2nIlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo/9oY7/oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/9oY7/oo//oo//oo//oo//oo//oo//oo//oo//oo9iPCtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdjPSz/oo//oo//oo//oo//oo//oo8AAAD/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo8AAAAAAAAAAAD/oo//oo//oo//oo//oo//oo/0k5X/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo8AAAAAAADAAAAAAAMAAIAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAAAAQAAwAAAAAADAAA="
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
VarSetCapacity(Dec, DecLen, 0)
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
DllCall("Gdiplus.dll\GdipCreateHICONFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
Return hBitmap
}
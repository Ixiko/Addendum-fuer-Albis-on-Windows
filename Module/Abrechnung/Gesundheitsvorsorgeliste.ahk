
;{ -----------------  Addendum für Albis on Windows - Modul Gesundheitsvorsorgeliste.ahk V1.89 vom 04.07.2020 -  Ixiko

; 	nur lauffähig ab der Albis Version 12.071.005 (CG hatte mit dieser Albisversion eine Änderung des "Cave! von" Fenster eingeführt
;   und nur lauffähig mit der jeweils aktuellen Autohotkey-Version zum Versionsdatum
;	-----------------------------------------------------------------------------------------------------------------------------------------------------------------
;	################# ZU BEACHTEN !!! ###################
; 	-----------------------------------------------------------------------------------------------------------------------------------------------------------------
;   dieses Skriptmodul ist nicht ohne Aufbereitung der eigenen Patientenakten für eine durch die KV nicht zu beanstandende Abrechnung lauffähig.
;	Ich nutze die Zeile 9 im "Cave! von" Fenster, für Kurznotizen zum jeweiligen Patienten zB. könnte folgendes dort stehen: "01740, C16 PP K19, GVU/HKS 01/16, GB 18.1A".
;	So weiß ich mit einem Blick, die Vorsorge-Coloskopie Ziffer habe ich schon abgerechnet, der Pat. hatte 2016 eine Coloskopie "C16" dort wurden Polypen entfernt "PP"
;   eine Kontrolle soll im Jahr 2019 "K19" erfolgen, als nächstes folgt das Quartal der letzten Vorsorgeuntersuchung. GB steht für geriatrisches Basisassesment.
;	in diesem Falle 2018 im 1.Quartal habe ich die Ziffer 03360 "GB18.1A" abgerechnet, wenn ich im folgenden Quartal bin, dann wäre die Ziffer B (03362) einzutragen.
;   Diese Eintragungen setzte ich bisher manuell.
;	Da Albis keine andere Überprüfung zuläßt, holt sich das Skript das Vorsorgequartal aus dieser Zeile.
;	Da wohl niemand Kürzelnotizen mit diesem Syntax verwendet (Komma separierte Eintragungen), würde dieses Skript zwar durchaus funktionieren,
;	aber es kann nicht wissen ob 2 Jahre seit der letzten Vorsorge vergangen sind. Das sollte zum Anfang der Skriptbenutzung manuell noch kontrolliert werden.
;	Sonst erhalten Sie lange Fehlerlisten von Ihrer KV.
;	Die GVU Liste erstelle ich mit dem Addendum Hauptskript (Hotkey's Strg + F7 oder Alt + F7 - > siehe dort).
;

;## TODO
;# 1.eine Start-Gui mit der Möglichkeit zum Ansehen einer GVU Liste um von dort direkt bei Fehlern zu den entsprechenden Patienten springen zu können
;#      und dann um die Formularerstellung zu starten
;# 4. Fehler mit folgendem Hinweis: das Fenster Leistungskette hat sich nicht geöffnet mit öffnen Sie es manuell...
;
;

;}

;{1. Skriptablaufeinstellungen
debug=0
#NoEnv

CoordMode, ToolTip, Screen
CoordMode, Pixel, Screen
CoordMode, Mouse, Screen
CoordMode, Caret, Screen
CoordMode, Menu, Screen

;#ErrorStdOut
SetBatchLines, -1
SetKeyDelay, -1
SetWinDelay, -1
SetControlDelay, -1
SetTitleMatchMode, 2
DetectHiddenWindows, On
DetectHiddenText, On
FileEncoding, UTF-8


hIBitmap:= Create_GVU_ico(true)
Menu Tray, Icon, hIcon:  %hIBitmap%


;ermittle ob eventuell dieses Modul schon gestartet wurde und breche den 2.Prozeß ab (funktioniert nicht, muss noch gemacht werden)
;If DllCall("GetCurrentProcessId") = 0 {
;			MsgBox, Dieses Skript sollte in keiner zweiten Instanz laufen!, 5
;			ExitApp
;	}

OnExit, ExitModul

;}

;{2. Hotkey
Hotkey, ^!p		, GVUPause
Hotkey, ^!l		, GVUListVars
Hotkey, End  	, ExitModul

;}

;{3. Minimieren des SciteWindow

	SciteWasActive:= 0

	IfWinActive, ahk_class SciTEWindow
	{
				SciteWasActive:= 1
				WinMinimize, ahk_class SciTEWindow
	}
;}

;{4. Variablendeklarationen

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  diese Variablen sind global zu halten
	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		global AlbisWinID, AddendumDir, lvar, data, line
		global hHook
		global WorkOff:= []
	;}

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  weitere Variablen
	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		sleeptime    				:= 55						    			; Wartezeiten die variabel gehalten sind für verschiedene Fensteroperationen
		sleeptime1				:= 10
		sleeptime2				:= 30
		boxtime1    				:= 2								    		;2 Sekunden manuelle Eingriffsmöglichkeit am Anfang und Ende jedes Durchlaufes
		boxtime2     				:= 2
		data            				:= []									    	;vorher data:= Object()
		lvar            				:= []							    			;speichert die jeweils aktuelle Zeile der aktuell bearbeiteten GVU Liste
		line            				:= []								    		;speichert die Zeilen der aktuellen GVU Liste
		TC            				:= A_DD "." A_MM "." A_YYYY  	;TimeCode
		ControlsFirst				:= Object()                            	;enthält später Daten zu gesetzten Checkboxen im GVU Formular
		ControlsYet				:= Object()                              	;das ist dann das Objekt für den Vergleich
		minGVU_Abstand 		:= 36						     			;größer gleich 36 Monate gültig seit 01.01.2020
		minAlterEinmalig   	:= 18
		minAlterMehrmalig	:= 35
		AbrBedNegativOld1	:= 0
		AbrBedNegativOld2	:= 0
		AbrBedNegativOld3	:= 0
		AbrBedNegativ1		:= 0
		AbrBedNegativ2		:= 0
		AbrBedNegativ3		:= 0
		Counter					:= 0						     				;Counter ist die Backup Variable für den folgenden loop
		ohneGVU					:= 0									    	;wieviele Patienten sind noch ohne Abrechnung aus der Liste
	;}

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  Auslesen des Applikationsverzeichnisses (AddendumDir))
	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		;AddendumDir := RegReadUniCode64("HKEY_LOCAL_MACHINE", "SOFTWARE\Addendum für AlbisOnWindows", "ApplicationDir")
		global AddendumDir		:= FileOpen("C:\albiswin.loc\AddendumDir","r").Read()
	;}

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  Funktion zum Auswählen *gvu.txt (Quartalsdatei)
	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		IniRead, GVUFile, % AddendumDir "\Addendum.ini", GVU_Queue, aktuelles_GVUFile
		If InStr(GVUFile, "ERROR") or (GVUFile = "") {
			FileSelectFile, FSF_GVUFile, 3, %AddendumDir%\Tagesprotokolle , GVU Quartals-Datei öffnen, GVU-Dokumente (*GVU.txt)
			If (GVUFile = "") {
				MsgBox, 0x40000, AfAoW - Modul GVU Liste ausfüllen, Es scheint so als hätten Sie keine Datei ausgewählt`nDas Modul wird jetzt beendet.
				ExitApp
			}
			SplitPath, FSF_GVUFile, GVUFile
			IniWrite, % GVUFile, % AddendumDir "\Addendum.ini", GVU_Queue, aktuelles_GVUFile
		}

		RegExMatch(GVUFile, "^\d\d", Quartal)
		RegExMatch(GVUFile, "(?<=^\d\d)\d\d", Jahr)
		aktQuartal:= Quartal "/" Jahr

	;}

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  Liste ungewollter Fenster die einfach geschlossen werden können
	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		unwanted:= Object()
		unwanted.push:= "GNR-Vorschlag"
		unwanted.push:= "ALBIS - [Prüfung EBM/KRW]"
		unwanted.push:= "Überprüfung der Versichertendaten"
		unwanted.push:= "Herzlichen Glückwunsch ahk_class #32770"
	;}

;}

;{5. MsgBox-/Hinweistexte

MsgTxt001 =
(
Die Abrechnungsbedingungen für die Vorsorgeuntersuchungen
                                      sind erfüllt.

---------------------------------------------------------------------------
Ich starte in %boxtime1%s die Erstellung der Formulare.
---------------------------------------------------------------------------
---------------------------------------------------------------------------
"OK" 			: Vorgang sofort Ausführen dann drücken Sie.
"Abbrechen"	: Modul jetzt beenden
---------------------------------------------------------------------------
Strg und + - pausiert das Script jederzeit
mit der erneuten Kombination läuft es weiter.
Esc  - beendet das Programm jederzeit
)

MsgTxt002 =
(
Die Abrechnungsbedingungen für die Vorsorgeuntersuchungen
                               sind NICHT erfüllt!

---------------------------------------------------------------------------
In %boxtime1%s mache ich mit der Erstellung der Formulare
für den nächsten Patienten weiter.
---------------------------------------------------------------------------
---------------------------------------------------------------------------
"OK" 			: Vorgang sofort Ausführen dann drücken Sie.
"Abbrechen"	: Modul jetzt beenden
---------------------------------------------------------------------------
Hinweis
Strg und + - pausiert das Script jederzeit mit der erneuten Kombination läuft es weiter.
Esc  - beendet das Programm jederzeit
)

MsgTxt003 =
(
Sie können jetzt noch die gemachten Eintragungen zu kontrollieren.
Drücke auf 'Ja' wenn Du mit der der nächsten Akte gleich weiter machen möchtest?
Drücke auf 'Nein' und das Programm wird beendet. Die Änderungen werden gespeichert.
Dieses Fenster schließt sich nach %boxtime2%s.
)

MsgTxt004 =
(
Scheinbar liegt bei diesem Patienten kein gültiger Abrechnungsschein vor.
Falls dem nicht so ist, drücken Sie 'JA' um die Formulare doch anzulegen.
Drücke auf 'Nein' um mit dem nächsten Patienten fortzufahren.
)

MsgTxt005 =
(
Das GVU-Formular konnte nicht aufgerufen werden.
Bitte öffnen Sie es jetzt manuell oder drücken Sie auf
'Abbrechen' zum Beenden des Skriptes.
)

MsgTxt006 =
(
Das HKS-Formular konnte nicht aufgerufen werden.
Bitte öffnen Sie es jetzt manuell oder drücken Sie auf
'Abbrechen' zum Beenden des Skriptes.
)

MsgTxt007 =
(
Die 2.Seite des GVU-Formular konnte nicht aufgerufen werden.
Bitte drücken Sie den Button 'Weiter" im Formular oder
drücken Sie hier auf 'Abbrechen'. Das Skripte wird dann beendet!
)


;}

;{6. prüfe ob Albis läuft - starte die Hook Prozesse

	AlbisWinID:= AlbisWinID()
	If !AlbisWinID
		MsgBox, Starte bitte zunächst Albis!

	gosub InitializeWinEventHooks

;}

;{7. weitere Vorbereitungen (Einlesen der Vorsorgeliste, Gui)

	; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  Einlesen des Abrechnungsrelevanten Quartals, Festellung der Anzahl der Dateifelder, ich möchte später pro erledigtem Patienten ein Textfeld in die entsprechende Zeile einfügen
	; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  hier wird jede Zeile in ein Array eingelesen
	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
			FileRead, filedata, % AddendumDir "\TagesProtokolle\" GVUFile
			Loop, Parse, filedata, `n, `r
			{

					If StrLen(A_LoopField) < 3
							continue

					Counter++
					line[Counter]:= A_LoopField

					If InStr(A_LoopField, "Mindestalter 35 Jahre")
					{
						AbrBedNegativOld1++
						continue
					}
					else If InStr(A_LoopField, "kein Abrechnungsschein")
					{
						AbrBedNegativOld2++
						continue
					}
					else If InStr(A_LoopField, "eine GVU kann erst in")
					{
						AbrBedNegativOld3++
						continue
					}

					If !InStr(A_LoopField, "Neu:")
						ohneGVU++
			}

	;}

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  Loop schaut welcher Backup-Name noch frei wäre und fertigt eine Kopie der aktuellen GVU-Liste an
	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
			Loop {
				backupFile:= AddendumDir "\TagesProtokolle\" SubStr(GVUFile, 1, StrLen(GVUFile) - 4) . A_Index . ".txt"
			} until !FileExist(backupFile)

		;-: Anlegen einer Sicherheitskopie der *gvu.txt Datei
			FileCopy, % AddendumDir "\TagesProtokolle\" GVUFile, % backupFile
			MsgBox, 0x40004, % "AfAoW - Modul GVU Liste ausfüllen", % "Es sind " . Counter . " Patienten in der Quartalliste: " . aktQuartal . " vorhanden.`nSoll ich mit dem Ausfüllen beginnen?`n'Nein' beendet das Programm."
				IfMsgBox, No
					ExitApp
	;}

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  Arbeitsbereich des Monitors bestimmen
	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
			SM_CMONITORS := 80
			SysGet, monCount, % SM_CMONITORS
			Loop, % monCount
					SysGet, Mon%A_Index%, Monitor, %A_Index%
			SysGet, Mon1, MonitorWorkArea, 1
	;}

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  die einzige GUI wird erstellt - benötigt um den Programmablauf kontrollieren zu können
	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
			Gui, AC1: NEW
			Gui, AC1: Font, S10 CDefault, Futura Bk Bt
			Gui, AC1: +LastFound +AlwaysOnTop +HwndhAC
			Gui, AC1: Margin, 5, 5
			Gui, AC1: Add, Listview, r8 xm y0 w550 h700 gAAListView BackgroundAAAAFF HwndhAC_LV Grid NoSort, % "GVU-Quartal: " lvar[2] ", Pat.gesamt: " Counter ", unbearbeitet: " ohneGVU
			;Gui, AC1: Add, Button, x20 y705 vPause gAAListView, Pause
			Gui, Add, StatusBar, vACSBar HwndhACSBar, Beenden (Ende), Pause (Strg + Alt + P)		                                                	; Programmschritt weiter Strg + Pfeil rechts
			Gui, AC1: Show, x1330 y100 w560 h750 Hide, Addendum für AlbisOnWindows - Gesundheitsvorsorgeliste    ;w350 h750
			Gui, AC1: Default
		;-: exakte Größe der Gui bestimmen
			AC:= GetWindowSpot(hAC)
		;-: Gui auf dem Monitor automatisch platzieren
			SetWindowPos(hAC, (Mon1Right - AC.W - 10), 43, AC.W, (Mon1Bottom - 60))					; y = 43, dort endet der Fenstertitelbereich und Menubereich des Albisfenster
			AC:= GetWindowSpot(hAC)
			ControlGetPos, x, y, w, ACSbarH,, ahk_id %hACSBar%
			ControlMove,,,,, ( AC.CH - (2 * AC.BH) - ACSBarH ), ahk_id %hAC_LV%
			Gui, AC1: Show
		;}

			SetTimer, EventHook_WinHandler, 200
;}

;{8. Prozeduren: gehe Zeile für Zeile durch die GVUListe und erstelle Formulare, Einträge und Ziffern für noch nicht abgearbeitete Patienten

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  a. Zeilenweises bearbeiten der ausgewählten Quartalsliste
	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		LineNum:= FormularZaehler := 0

		Loop, % line.Count() {

			; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			; Listview vorbereiten, Variablen zurücksetzen, aktuelle Patientendaten anzeigen
			; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
				c := 1
				lvar[4]	    	:= ""																; immer leeren sonst enthält diese eventuell noch den vorherigen Eintrag
				thisround  	:= A_Index													; Formularzähler
				currentline 	:= Trim(line[thisround])               					; temporäre Variable
				lvar           	:= StrSplit(currentline, "`;")								; Aufteilen der aktuellen Zeile in vier Variablen als Beispiel, lvar[3] = Patienten-ID
				GVUMonat  	:= SubStr(lvar[1], 4, 2)									; Ermitteln des Untersuchungsmonats
				GVUJahr      	:= SubStr(lvar[1], StrLen(lvar[1]) - 1, 2)			; Ermitteln des Untersuchungsjahres
				FormularZaehler ++

			; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			;  nach Strsplit entstehen 4 Variablen - lvar[1]: Untersuchungsdatum, lvar[2]: Untersuchungsquartal, lvar[3]: PatientenID, lvar[4]: enthält evtl. die Kopie der Zeile 9 des ""Cave! von"" Fenster
			;  Beispielzeilen: 17.01.2018;01/18;17716;Alt:01740, GVU/HKS 12#14,Neu:GVU/HKS, GVU/HKS 01^18
			;	                     17.01.2018;01/18;11027;Alt:01740, GVU/HKS 10#15,Neu:01740, GVU/HKS 01^18
			;                        17.01.2018;01/18;20273;kein Abrechnungsschein angelegt gewesen.
			;
			;  wenn lvar[4] eine Kopie enthält dann ist diese schon Zeile abgearbeitet, dies ist sozusagen gleich ein Backup der Zeile 9 im Cave! von Fenster oder es wird dort die Fehlermeldung
			;  eingeschrieben, warum bei dem jeweiligen Patienten keine Formulare angelegt werden konnten
			; ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

			If (lvar[4] = "")
			{
					; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
					; Listview vorbereiten, Variablen zurücksetzen, aktuelle Patientendaten anzeigen
					; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
						GUI, AC1:Default
						LV_Delete()                                                                                                                                        	;löschen des Inhalts meines ListView Fenster
						LV_Add("", SubStr("000" Counter, -2) " Untersuchungen gesamt. Davon bearbeitet: " Counter - ohneGVU)
						LV_Add("", Substr("00" c, -1) ": Beginne mit Formular Nr.: " SubStr("000" thisround, -2) ", Patienten-ID: " SubStr("00000" lvar[3], -4) ", Untersuchungsdatum: " lvar[1])
						FileAppend	, % "`nFormular Nr.: " SubStr("000" thisround, -2) ", Patienten-ID: " SubStr("00000" lvar[3], -4) ", Untersuchungsdatum: " lvar[1] " | "
											, % AddendumDir "\Tagesprotokolle\GesundheitsvorsorgelisteLog.txt", UTF-8
					;}

					; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
					; Patientenakte aus der Liste öffnen
					; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
						If !RegExMatch(lvar[3], "\d+")				;prüft auf fehlerhafte Patienten-ID
						{
								MsgBox, 0x40000, Gesundheitsvorsorgeliste, % "In der Liste ist eine Zeile ohne Patienten-ID vorhanden.`nFahre mit dem nächsten Patienten fort!", 3
								FileAppend, % "`nIn der Liste ist eine Zeile ohne Patienten-ID ( " SubStr("00000" lvar[3], -4) " ) vorhanden. Fahre mit dem nächsten Patienten fort!"
													, % AddendumDir "\Tagesprotokolle\GesundheitsvorsorgelisteLog.txt", UTF-8
								continue
						}

						If !AlbisAkteGeoeffnet("", "", "", lvar[3])
						{
								If !AlbisAkteOeffnen(lvar[3])																								;kein WaitAndActivate notwendig (hat diese Funktion integriert)
										MsgBox, % "Akte: " lvar[3] " konnte nicht geöffnet werden.`nUnternimm was!"
						}

						Sleep, % sleeptime
					;}

					; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
					; prüft das Alter des Patienten
					; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
						age	:= Floor(DateDiff( "YY", AlbisPatientGeburtsdatum(), lvar[1]))
						If (age >= minAlterEinmalig)
						{
								LV_Add("", Substr("00" c++, -1) ": 'Mindestalter " minAlterEinmalig " Jahre `(" age "`) ist erfüllt!" )
								FileAppend, % "Mindestalter " minAlterEinmalig " Jahre `(" age "`) ist erfüllt! | " , % AddendumDir "\Tagesprotokolle\GesundheitsvorsorgelisteLog.txt", UTF-8
						}
						else
						{
								LV_Add("", Substr("00" c++, -1) ": Abrechnungsbedingung - 'Mindestalter " minAlterEinmalig " Jahre' `(" age "`) ist nicht erfüllt!" )
								AbrBedNegativ1++
								MsgBox, 0x40000, Gesundheitsvorsorgeliste, %  "Die Abrechnungsbedingung älter als " minAlterEinmalig " Jahre `(" age "`) ist nicht erfüllt!`nFahre mit dem nächsten Patienten fort!", 5
								line[thisround]:= currentline ";!!!Mindestalter " minAlterEinmalig " Jahre ist nicht erfüllt!"
								FileAppend, % line[thisround] , % AddendumDir "\Tagesprotokolle\GesundheitsvorsorgelisteLog.txt", UTF-8
								AlbisAkteSchliessen()
								continue
						}

					;}

					; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
					; prüft ob Chipkarte für das Arbeitsquartal eingelesen wurde , falls ein Patient ohne Abrechnungsschein in der GVU Liste eingetragen ist
					; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
						LV_Add("", Substr("00" c++, -1) ": Prüfe auf Vorliegen eines gültigen Abrechnungsscheines")

						abr:= AlbisAbrechnungsscheinVorhanden(aktQuartal)
						LV_Add("", Substr("00" c++, -1) ": Ergebnis der Abrechnungsscheinprüfung: " abr)
						FileAppend, % "Ergebnis der Abrechnungsscheinprüfung: " abr " | " , % AddendumDir "\Tagesprotokolle\GesundheitsvorsorgelisteLog.txt", UTF-8
					;}

					; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
					; kein Abrechnungsschein angelegt - dann führt er folgende Zeilen aus
					; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
						If (abr = 0)
						{
								LV_Add("", "bei dem geöffneten Patienten wurde kein Abrechnungsschein angelegt")
								LV_Add("", "eine Abrechnung der Untersuchung ist in dem Quartal nicht möglich.")
								FileAppend, % "es war kein Abrechnungsschein für das Quartal " aktQuartal " angelegt!" , % AddendumDir "\Tagesprotokolle\GesundheitsvorsorgelisteLog.txt", UTF-8
								Sleep, % sleeptime
								AbrBedNegativ2++

							;falls ein Programmierfehler vorliegt oder einfach nur der Schein nicht erkannt wurde kann der User hier nochmal eingreifen
								MsgBox, 0x40001, Gesundheitsvorsorgeliste, % MsgTxt004
								IfMsgBox, Ok
									gosub, Abrechnung_Positiv
								IfMsgBox, Cancel
								{
									line[thisround]:= currentline . ";!! es ist kein Abrechnungsschein angelegt gewesen !!"
									FileAppend, % line[thisround], % AddendumDir "\Tagesprotokolle\GesundheitsvorsorgelisteLog.txt", UTF-8
									AlbisAkteSchliessen()
								}
						}

					;}

					; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
					; Abrechnungsschein vorhanden - dann das hier
					; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
						If (abr > 0)
						{
								gosub, Abrechnung_Positiv
								Loop, % WorkOff.Length()
									WorkOff.RemoveAt(A_Index)                                                 	;Inhalte löschen
						}

					;}

					ohneGVU --

			}

		}

	;}

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  b. es sind alle Formulare erstellt. Und Tschüssikowski! (EXITMODUL)
	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{

	ExitModul:

		IniWrite, % GVUFile, % AddendumDir "\Addendum.ini", GVU_Queue, aktuelles_GVUFile

	FileSpeichern:

		LV_Delete()

		FileDelete, % AddendumDir "\TagesProtokolle\" GVUFile
		Loop, % line.Count()
		{
				If StrLen(line[A_Index]) > 3
				{
						If ( A_Index < line.Count() )
							FileAppend, % line[A_Index] "`n" , % AddendumDir "\TagesProtokolle\" GVUFile, UTF-8
						else
							FileAppend, % line[A_Index] 		, % AddendumDir "\TagesProtokolle\" GVUFile, UTF-8
				}
		}

		if hWinEventHook
			UnhookWinEvent(hWinEventHook, HookProcAdr)

		;If SciteWasActive
		;	WinActivate, ahk_class SciTEWindow

		maxAbrBedNeg					:= AbrBedNegativ1 + AbrBedNegativ2 + AbrBedNegativ3 + AbrBedNegativOld1 + AbrBedNegativOld2 + AbrBedNegativOld3
		maxAbrBedAktuellNeg		:= AbrBedNegativ1 + AbrBedNegativ2 + AbrBedNegativ3
		ABGneg1							:= AbrBedNegativ1 + AbrBedNegativOld1
		ABGneg2							:= AbrBedNegativ2 + AbrBedNegativOld2
		ABGneg3							:= AbrBedNegativ3 + AbrBedNegativOld3
		AbrGesPositiv					:= Counter   	- maxAbrBedNeg
		AbrGesAktuell					:= ohneGVU    	- maxAbrBedNeg

		Endtext=
		(LTrim


		-----------------------------------------------------------------------
		Ergebnisse aller Durchläufe:
		-----------------------------------------------------------------------
		%Counter% U. sind gemacht worden
		%AbrGesPositiv% U. sind davon abgerechnet
		%maxAbrBedNeg% U. konnten nicht abgerechnet werden

		nicht erfüllte Bedingungen:
		%ABGneg1% x war das Mindestalter von 35 nicht erreicht
		%ABGneg2% x lag kein Abrechnungsschein vor
		%ABGneg3% x stimmte der minimale Abstand (%minGVU_Abstand%) nicht
        -----------------------------------------------------------------------
		Ergebnisse des aktuellen Durchlaufs:
		-----------------------------------------------------------------------
		%FormularZaehler%  U. wurden bearbeitet
		%ohneGVU% U. sind noch nicht abgerechnet.
		%maxAbrBedAktuellNeg% x U. konnten nicht abgerechnet werden

		nicht erfüllte Bedingungen:
		%AbrBedNegativ1% x war das Mindestalter von 35 nicht erreicht
		%AbrBedNegativ2% x lag kein Abrechnungsschein vor
		%AbrBedNegativ3% x stimmte der minimale Abstand (%minGVU_Abstand%) nicht
		##################################


		)

		FileAppend, % Endtext  , % AddendumDir "\Tagesprotokolle\GesundheitsvorsorgelisteLog.txt", UTF-8
		MsgBox, % Endtext

		OnExit
;}

;}

ExitApp

;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;		PROZEDUREN		PROZEDUREN		PROZEDUREN		PROZEDUREN		PROZEDUREN		PROZEDUREN		PROZEDUREN		PROZEDUREN		PROZEDUREN
;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

;{9. Label Abrechnung_Positiv und Ortho1

Abrechnung_Positiv: 						;{ überprüft ob die Voraussetzungen zur Abrechnung einer Gesundheitsvorsorge erfüllt sind

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  a) sucht nach einer Zeile welche den String "GVU" enthält, wenn nichts notiert ist, sichert er die Zeile 9
	;  -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
			CZeile:= Trim(AlbisGetCaveZeile("", "GVU"))
			If (StrLen(CZeile) = 0)
				CZeile:= Trim(AlbisGetCaveZeile("9", ""))
		; Sichern der Kürzelzeile 9
			FileAppend, % TC ";" lvar[3] ";" AlbisCurrentPatient() ";" CZeile "`n", % AddendumDir . "\Tagesprotokolle\cave9.txt", UTF-8
	;}
	;Pause
	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  b) Parsen:        	 der 9.Zeile aus dem "Cave! von" Fenster
	;  -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
			;ist die Zeile 9 noch leer, wird gvufound auf 0 gesetzt
			If RegExMatch(CZeile, "i)G[VU]+.H[KS]+.[KVU]*\s*(\d+).(\d+)", thisGVU)  ; Erkennung auch bei Buchstabenverdrehern (GUV
			{
					gvufound:= true
					CaveMonat:= thisGVU1
					CaveJahr:= thisGVU2
					;RegExMatch(CZeile, "(?<=G[VU][VU].H[KS][KS]\s)\d\d", CaveMonat)
					;RegExMatch(CZeile, "(?<=G[VU][VU].H[KS][KS]\s\d\d.)\d\d", CaveJahr)
			}
			else
			{
					gvufound:= CaveMonat:=  CaveJahr:= 0
			}

			LV_Add("",  Substr("00" c++, -1) ": " thisGVU " gefunden - letzte Untersuchung: " SubStr("0" CaveMonat, -1) "/" SubStr("0" CaveJahr, -1) )
	;}

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  c) Berechnung:	Abstand zwischen den Vorsorgeuntersuchungen
	;  -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
			if !gvufound
			{
					LV_Add("",  Substr("00" c++, -1) ": Es ist noch nie eine Vorsorgeuntersuchung abgerechnet worden.")
					FileAppend, % "Es ist noch nie eine Vorsorgeuntersuchung abgerechnet worden. | " , % AddendumDir "\Tagesprotokolle\GesundheitsvorsorgelisteLog.txt", UTF-8
					CaveJahr := CaveMonat :=0
				; 2 Jahre + 1 Monat (damit es nachher funktioniert)
					deltaGVU := minGVU_Abstand + 1
			}
			else
			{
					If (age < minAlterMehrmalig)
					{
							AbrBedNegativ3++
							line[thisround]:= currentline ";Alt:" CZeile ";Eine GVU darf zwischen dem " minAlterEinmalig ". und dem " minAlterMehrmalig ". Lebensjahr nur einmal abgerechnet werden!"
							LV_Add("",  Substr("00" c++, -1) ": Eine GVU darf zwischen dem " minAlterEinmalig ". - " minAlterMehrmalig " Lebensjahr nur einmal abgerechnet werden." )
							FileAppend, % line[thisround], % AddendumDir "\Tagesprotokolle\GesundheitsvorsorgelisteLog.txt", UTF-8
							MsgBox, 0x40001, Gesundheitsvorsorgeliste, % MsgTxt002, % boxtime1
					    	IfMsgBox, Cancel
									ExitApp
					}
					else
					{
						; Berechnung funktioniert nur ab dem Jahr 2000
							deltaGVU:=	(GVUJahr * 12 + GVUMonat) - (CaveJahr * 12 + CaveMonat)
							LV_Add("",  Substr("00" c++, -1) ": Abstand: " deltaGVU " Monate.")
							FileAppend, % "GVU-Abstand: " SubStr("00" deltaGVU, -1) " Monate. | " , % AddendumDir "\Tagesprotokolle\GesundheitsvorsorgelisteLog.txt", UTF-8
					}
			}
	;}

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  d) Entscheidung:	was passiert wenn der Abstand zwischen Vorsorgeuntersuchungen stimmt (Label Ortho1 ausführen) oder nicht
	;  -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
			If (deltaGVU >= minGVU_Abstand) 												                                      ;- also größer gleich 24 Monate Abstand (bis Ende März 2019)
			{
						if (GVUJahr = 0)
								GVUJahr:="", GVUMonat:=""

						MsgBox, 0x40001, Gesundheitsvorsorgeliste, % MsgTxt001 , % boxtime1
						IfMsgBox Timeout
								gosub, Ortho1
						IfMsgBox Ok
								gosub, Ortho1
						IfMsgBox Cancel
						{
							line[thisround]:= currentline ";Alt:" CZeile ",Neu:"
							ExitApp
						}
			}
			else
			{
						AbrBedNegativ3++
						line[thisround]:= currentline ";Alt:" CZeile ";!! eine GVU kann erst in " (minGVU_Abstand - deltaGVU) " Monaten abgerechnet werden !!"
						FileAppend, % line[thisround], % AddendumDir "\Tagesprotokolle\GesundheitsvorsorgelisteLog.txt", UTF-8
						MsgBox, 0x40001, Gesundheitsvorsorgeliste, % MsgTxt002, % boxtime1
						IfMsgBox, Cancel
									ExitApp
			}
	;}

return
;}

Ortho1: 										;{	GVU und HKS-Formular werden automatisiert aufgerufen und bearbeitet

	;  Ändern des Programmdatums auf das Untersuchungsdatum, sonst würden alle Untersuchungen auf dem eingestellten Datum abgerechnet
		LV_Add("",  Substr("00" c++, -1) ": Setze das Programmdatum auf das Untersuchungsdatum")
		AlbisActivate(1)
		AlbisSetzeProgrammDatum(lvar[1])																	;kein WaitAndActivate notwendig (hat diese Funktion integriert)
		LV_Add("",  Substr("00" c++, -1) ": Datum geändert")

	; HKS Formular wird aufgerufen wenn Patient >= 35 Jahre
		If (age >= minAlterEinmalig) {

				LV_Add("",  Substr("00" c++, -1) ": warte auf HKS-Formular")
				If !hHKSWin:= AlbisFormular("eHKS_ND") {
					while !(hHKSWin:=WinExist("Hautkrebsscreening - Nichtdermatologe ahk_class #32770") ) 	{
							MsgBox,, Addendum für AlbisOnWindows, % MsgTxt006	;...Das HKS-Formular konnte nicht aufgerufen werden.... manuell öffnen
							If MsgBox, No
								ExitApp
					}
				}

			; Befüllen des Formulars
				AlbisHautkrebsScreening("opb", true, true )
				LV_Add("",  Substr("00" c++, -1) ": HKS Formular ausgefüllt")

			; Daten/Leistungen in die Patientenakte eintragen
				AlbisActivate(1)
				AlbisSchreibeInKarteikarte("", lvar[1], "lko", "01732-01746")
				LV_Add("",  Substr("00" c++, -1) ": Ziffern sind in die Akte eingetragen.")

			; Cave! von Text
				GVUText := "GVU/HKS " GVUMonat "^" GVUJahr ", "

		}
		else {	; jünger als 35 Jahre dann nur den GVU Komplex eintragen

				AlbisActivate(1)
				AlbisSchreibeInKarteikarte("", lvar[1], "lko", "01732")
				LV_Add("",  Substr("00" c++, -1) ": Ziffer ist in die Akte eingetragen.")

			; Cave! von Text
				GVUText := "GVU " GVUMonat "^" GVUJahr ", "

		}

	; cave Fenster - die Kürzel auffrischen   #########FEHLER
		LV_Add("", "        " GVUText)
		if (CZeile = "")											;CZeile ist die ursprüngliche cave 9 Zeile - war sie leer?
				neueZeile	:= GVUText
		 else {
				CZeile		:= StrReplace(CZeile, thisGVU, "###")
				neueZeile	:= RegExReplace(CZeile, "###\s*\,*", GVUText)
		}
		neueZeile:= RTrim(neueZeile, ", ")

	; der hook fängt dieses Fenster manchmal nicht ab, das Cave! Fenster läßt sich aber bei dieser Warnung nicht öffnen
		if WinExist("ALBIS - [Prüfung EBM/KRW]") {
				WinActivate   	, % "ALBIS - [Prüfung EBM/KRW]"
				WinWaitActive	, % "ALBIS - [Prüfung EBM/KRW]",,2
				SendInput, {ESC}
				WinWaitClose	, % "ALBIS - [Prüfung EBM/KRW]",,2
				LV_Add("", "Fenster: Prüfung EBM/KRW geschlossen")
		}

	; ersetzt oder schreibt in die Zeile 9 im CaveVon! Fenster - das Abrechnungsquartal der Komplexe
		AlbisActivate(1)
		ControlGetFocus, cFocus, % "ahk_id " AlbisWinID()
		If InStr(cFocus, "Edit")
			SendInput, {Escape}
		AlbisSetCaveZeile(9, neueZeile)

	; Zeile 9 wird im Array gespeichert (wird später in der GVU-Liste gespeichert)
		line[thisround]:= currentline ";Alt:" CZeile ",Neu:" neueZeile

		LV_Add("",  Substr("00" c++	, -1) ": Zeile 9 geändert. Cave! von geschlossen.")
		LV_Add("",  Substr("00" c			, -1) ": Das Erstellen der Formulare und Leistungen ist für diesen Vorsorgefall abgeschlossen!")
		LV_Add("",  Substr("00" c			, -1) ": Beginne gleich mit dem nächsten Fall......")

	; Schließen der aktuellen Akte, Mitteilung an den Nutzer das mit dem nächsten Patienten fortgefahren wird
		MsgBox, 262148, Addendum für AlbisOnWindows - Gesundheitsvorsorgeliste, % MsgTxt003,  % boxtime1
			IfMsgBox, No
				ExitApp

	; die aktuelle Patientenakte wird geschlossen
		AlbisAkteSchliessen(lvar[3])
		;ohneGVU --

	; Löschen des Inhalts des ListView Fenster
		Gui, AC1:Default
		LV_Delete()

return
;}

;}

;{10. Hotkeybereich
GVUPause:
	Pause, On
return
GVUListVars:
	ListVars
return


;}

;{11. SUBROUTINEN/Gui's/Fensteranordung

WinMoveMsgBox1: ;{

	SetTimer, WinMoveMsgBox1, OFF
	ID:=WinExist(WinName)
	WinMove, ahk_id %ID%, , 850, 850

return
;}

StartGui: ;{

return
;}

AAListView: ;{

		if A_GuiEvent = Normal
		{
			LV_GetText(Reihentext, A_EventInfo)  ; Ermittelt den Text aus dem ersten Feld der Reihe.
			ToolTip Text: "%Reihentext%"
		}
		;MsgBox, %A_GuiControl%
		If InStr(A_GuiControl, "Pause")
			Pause, Toggle

return
;}


;}

;{12. Funktionen

WinWasCreated(WinTitle, WinClass:="", WinText:="") {

	For index, val in WorkOff
		If InStr(val, WinTitle) and InStr(val, WinClass) and InStr(val, WinText)
				return 1

return 0
}

AlbisWaitivate(WinTitle, Debug=1, DbgWhwnd=0) {                                   	;--Fehlermeldungsfenster (GNR-Vorschlag, Prüfung EBM/KRW, Überprüfung der Versichertendaten) automatisch schliessen

	; Debug= 0 dann erfolgt keine Ausgabe in ein Listviewfenster, doch noch immer läßt sich der Errorlevel auslesen
		If (Debug =1) {
            	Gui, %DbgWhwnd%:Default
            	LV_Add("", "      Warte auf Fenster: " . WinTitle)
		}

		WinActivate, %WinTitle%,,1
            WinWaitActive, %WinTitle%,,5

		while !WinExist(WinTitle) {

            		phwnd:= DLLCall("GetLastActivePopup", "uint", AlbisWinID)
            		WinGetTitle, WT, ahk_id %phwnd%

            		If InStr(WT, "GNR-Vorschlag") {
                        	WinActivate, %WT%
                        		VerifiedClick("Button1", "GNR-Vorschlag")
                                    Sleep, % sleeptime
                                    If (Debug =1) {
                                    	LV_Add("", "Fenster: GNR Vorschlag - geschlossen")
                                    }
            		}

            		If InStr(WT, "ALBIS - [Prüfung EBM/KRW]") {
                        	WinActivate, %WT%
                        		SendInput, {ESC}
                                    Sleep, 100
                                    If (Debug =1) {
                                    	LV_Add("", "Fenster: Prüfung EBM/KRW geschlossen")
                                    }
                        }

            		If Instr(WT, "Überprüfung der Versichertendaten") {
                        		VerifiedClick("Button1", "Überprüfung der Versichertendaten")
                                    Sleep, 100
                                    If (Debug =1) {
                                    	LV_Add("", "Fenster: Überprüfung der Versichertendaten wurde ignoriert.")
                                    }
                        }

            		If WinWasCreated(WinTitle)
					 			 return "isClosed"

					If ((A_Index>20) and InStr(WinTitle, "Leistungskette")) {
							 return "isClosed"
					} else if (A_Index>20) {
							MsgBox, 0x40000, Achtung, Das Fenster %WinTitle%	hat sich nicht geöffnet.`nöffne es bitte manuell und drücke dann ok.
					}

            		sleep, 100
		}

		IfWinNotActive, %WinTitle%
            	WinActivate, %WinTitle%
                        WinWaitActive, %WinTitle%, , 3

return "ok"
}

Create_GVU_ico(NewHandle := False) {                                                        ;--Task-Icon
Static hBitmap := Create_GVU_ico()
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "AAABAAEAMDAAAAEAGACoHAAAFgAAACgAAAAwAAAAYAAAAAEAGAAAAAAAAAAAAGAAAABgAAAAAAAAAAAAAAAAAAAAAAD/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo8AAAAAAAAAAAD/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/1lJX/oo//oo//oo//oo8AAAD/oo//oo//oo//oo//oo//oo9jPSxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdkPi3/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/9oY7/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo/2lpX/oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9jPSxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdkPi3/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdFKhlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdLLh3MgW7XiHZGKxlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdVNCPsloP/oo//oo/nkn+gZVNCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/6nJD/oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdXNSTxmYb/oo//oo//oo//oo/ul4RWNSRCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdXNSTxmYb/oo//oo//oo//oo//oo//oo/zmoeoalhCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdYNiXlkX7/oo//oo//oo//oo//oo//oo//oo//oo/2nIlmPy5CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdaNybymof/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/5notuRTRCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdYNiXxmYb/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/7oIy0cV9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdaNybzmof/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/+oY61cmBCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdWNSTxmYb/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/9oI65dWNCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdUMyLij33/oo//oo//oo//oo//oo//oo//oo/WiHVCKBf3nYr/oo//oo//oo//oo//oo//oo//oo//oo/9oY26dmRCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdQMSDtl4T/oo//oo//oo//oo//oo//oo//oo+dY1FCKBdCKBdCKBfwmIb/oo//oo//oo//oo//oo//oo//oo//oo/9oY65dWNCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdJLBzok4H/oo//oo//oo//oo//oo//oo//oo+lZ1ZCKBdCKBdCKBdCKBd/Tz7/oo//oo//oo//oo//oo//oo//oo//oo//oo/9oI62c2BCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdEKhnfjXr/oo//oo//oo//oo//oo//oo//oo/Pg3FDKRhCKBdCKBdCKBdCKBdCKBe9d2X/oo//oo//oo//oo//oo//oo//oo//oo//oo/8oI21c2BCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfWh3X/oo//oo//oo//oo//oo//oo//oo/sloNCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfslYL/oo//oo//oo//oo//oo//oo//oo//oo//oo/6n4ypa1lCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBfCemj/oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf+oY7/oo//oo//oo//oo//oo//oo//oo//oo//oo/ul4RKLRxCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdpQjD8oI3/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBe2c2FCKBdDKRhCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/ShXJCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdGKxnvl4X/oo//oo//oo//oo//oo//oo//oo/ul4RPMB9CKBdJLRuzcV//oo//oo//oo9CKBdCKBdCKBdCKBdMLh1CKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo/+oY7QhHFCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBfIf2z/oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBfejXr/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBf8oI3/oo//oo//oo//oo//oo//oo//oo//oo//oo/0mohCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdHKxr4nov/oo//oo//oo//oo//oo//oo//oo/ainhDKRhCKBf3nYr/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo/7oIxNMB5CKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBe6dmT/oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo+4dGJCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBfkkX7/oo//oo//oo//oo//oo//oo//oo//oo9CKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdGKxlCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo/wmYZCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBfpk4D/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBfvl4X/oo//oo//oo//oo//oo//oo//oo/ymYdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBfij33/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBf5noz/oo//oo//oo//oo//oo//oo/ymYZCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBe0cV//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/7n41CKBdDKRhCKBf+oY7/oo//oo//oo//oo//oo/tloRCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdEKhn5nov/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBf/oo//oo//oo//oo//oo9gOylCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBfJgG3/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/4nYtCKBf2nYr/oo//oo//oo/nkn9CKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdLLh30moj/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/+oY5uRDNCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdyRzX9oY7/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/RhHJCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBeEU0H9oI7/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/6nou6dmTDfGn6n4z/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo+2c2BDKRhCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdsQzHtl4T/oo//oo//oo//oo//oo//oo//oo//oo/+oY7gjnxgOylCKBdCKBdXNiTDfGn/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo+IVURCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo/7nY//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdEKhl+Tj3cjHnymof+oY7/oo/3nYrzmofKgG1gPCpCKBdCKBdCKBdCKBdCKBdCKBdmPi6rbFrbi3j5noz/oo//oo//oo//oo//oo9JLBxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/7nZD/oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdHKxpCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdEKhlILBpCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/0k5T/oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo/+oY5iPCtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdjPSz/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9iPCtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdjPSz/oo//oo//oo//oo//oo//oo8AAAD/oo//oo//oo//oo/2lZX/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo8AAAAAAAAAAAD/oo//oo//oo//oo//oo//oo//oo/7nY//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo8AAAAAAADAAAAAAAMAAIAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAAAAQAAwAAAAAADAAA="
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

MyPauseToggle(debug=0) {

	if debug=0
		return

	;ListLines
	MsgBox, 4, Weiter?, Drücke Ja für weiter,`nNein für Abbruch
		IfMsgBox, No
				ExitApp

}


;}

;{13. WinEventHook all functions and labels

InitializeWinEventHooks:        	;{                                                                                          	;falls ich mal speziell nur einen Hook auf Albis setzten will sollte ich eine Funktion oder ein Label bereitstellen

	EVENT_SKIPOWNTHREAD				:= 0x0001
	EVENT_SKIPOWNPROCESS			:= 0x0002
	EVENT_OUTOFCONTEXT				:= 0x0000

	AlbisPID := AlbisPID()
	HookProcAdr 	:= RegisterCallback("WinEventProc", "F")
	hWinEventHook	:= SetWinEventHook( 0x0003, 0x0003, 0, HookProcAdr, AlbisPID, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
	hWinEventHook	:= SetWinEventHook( 0x0010, 0x0010, 0, HookProcAdr, AlbisPID, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS  )
	hWinEventHook	:= SetWinEventHook( 0x8000, 0x8000, 0, HookProcAdr, AlbisPID, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS  )
	hWinEventHook	:= SetWinEventHook( 0x8002, 0x8002, 0, HookProcAdr, AlbisPID, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS  )

return
;}

WinEventProc(hEventHook, event, hwnd, idObject, idChild, eventThread, eventTime) {						; abfangen von Albis Popup Fenstern
	static hOldHook
	Critical
	If (hwnd = 0)
		return 0
	If (hOldHook = hHook:= GetHex(hwnd))
		return 0
	hOldHook:= hHook
	SetTimer, EventHook_WinHandler,  -0
return 0
}

EventHook_WinHandler:         	;{                                                                                            	; Eventhookhandler für popup (child) Fenster in Albis

		hpop	:= GetLastActivePopup(AlbisWinID())

		If (hHook = hpop)
				hook	:= WinGetTitle(hHook) "|" WinGetClass(hHook) "|" WinGetText(hHook)
		else
				hook:= WinGetTitle(hpop) "|" WinGetClass(hpop) "|" WinGetText(hpop) "|" WinGetTitle(hHook) "|" WinGetClass(hHook) "|" WinGetText(hHook)

		WorkOff.push((hook))
		;SciTEOutput(hook, 0, 1, 0)

		if InStr(hook, "Patient hat in diesem Quartal") {
				VerifiedClick("Button1", "ALBIS ahk_class #32770", "Patient hat in diesem Quartal")
				LV_Add("", "Fenster: Patient hat in diesem Quartal... - geschlossen")
		}
		else if InStr(hook, "Wollen Sie die Leistungen")
		{
				VerifiedClick("Button2", "ALBIS ahk_class #32770", "Wollen Sie die Leistungen")
				LV_Add("", "Fenster: Wollen Sie die Leistungen wirklich diesem Schein zuordnen... - geschlossen")
		}
		else if InStr(hook, "Die aktuelle Zeile wurde")
		{
				VerifiedClick("Button2", "ALBIS ahk_class #32770", "Die aktuelle Zeile wurde")
				LV_Add("", "Fenster: Die aktuelle Zeile wurde nicht gespeichert... - geschlossen")
		}
		else if InStr(hook, "Herzlichen Glückwunsch zum Geburtstag")
		{
				VerifiedClick("Button1", "Herzlichen Glückwunsch zum Geburtstag ahk_class #32770")
				LV_Add("", "Fenster: Herzlichen Glückwunsch zum Geburtstag - geschlossen")
		}
		else If InStr(hook, "GNR-Vorschlag")
		{
           		VerifiedClick("Button1", "GNR-Vorschlag")
				LV_Add("", "Fenster: GNR Vorschlag - geschlossen")
   		}
		else if InStr(hook, "ALBIS - [Prüfung EBM/KRW]")
		{
               	WinActivate   	, % "ALBIS - [Prüfung EBM/KRW]"
               	WinWaitActive	, % "ALBIS - [Prüfung EBM/KRW]",, 2
           		SendInput, {ESC}
                Sleep, 200
               	LV_Add("", "Fenster: Prüfung EBM/KRW geschlossen")
         }
		 else If Instr(hook, "Überprüfung der Versichertendaten")
		{
              	VerifiedClick("Button1", "Überprüfung der Versichertendaten")
				LV_Add("", "Fenster: Überprüfung der Versichertendaten wurde ignoriert.")
         }

Return
;}



;}

;{14. Includebereich

	#include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
	#include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
	#include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
	#include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk
	#include %A_ScriptDir%\..\..\include\Addendum_Misc.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_PdfHelper.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
	#include %A_ScriptDir%\..\..\include\Addendum_Screen.ahk
	#include %A_ScriptDir%\..\..\include\Addendum_Window.ahk
	#include %A_ScriptDir%\..\..\lib\ACC.ahk
	#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
	#include %A_ScriptDir%\..\..\lib\ini.ahk
	#Include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
	#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
	#Include %A_ScriptDir%\..\..\lib\Sift.ahk
	#include %A_ScriptDir%\..\..\include\Gui\PraxTT.ahk

;}



	/*   GVU abgeschaltet - muss nicht mehr angelegt werden
	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  b) GVU Formular:     	Aufrufen
	;  -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
			Sleep, % sleeptime
			LV_Add("",  Substr("00" c++, -1) ": Warte auf GVU Seite 1")

			hGvuWin:= CallMenuWaitForWindow(33212, AlbisWinID, "Muster 30")     	; "Gesundheitsvorsorge (30)": "33212"
			If !hGvuWin
			{
					while !( hGvuWin:=WinExist("Muster 30 ahk_class #32770") )
					{
							MsgBox,, Addendum für AlbisOnWindows, % MsgTxt005    ;Das GVU-Formular konnte nicht aufgerufen werden..... manuell öffnen....
							If MsgBox, No
								ExitApp
					}
			}

			LV_Add("",  Substr("00" c++, -1) ": GVU Formular ist aufgerufen")
	;}

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  c) GVU Formular:     	den Status aller Checkboxen einlesen dann Button 'alte Formulare' drücken
	;  -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		; Auslesen aller Controls und der Einstellungen
			ControlsFirst:= AlbisMuster30ControlList()

		; Button 63 (Alte Daten) wird gedrückt
			If !VerifiedClick("Button63", "Muster 30 ahk_class #32770")
				MsgBox, % "Der Button -Alte Daten- konnte nicht angeklickt werden.`nBitte wenn möglich manuell auslösen."

			LV_Add("","GVU Seite 1 - Alte Daten Button gedrückt")
			LV_Add("", Substr("00" c++, -1) ": Warte auf Fenster ''alte Formulardaten übernehmen oder keine alten Daten vorhanden''.")
			Sleep, 500
			WinWhat:= 0
	;}

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  d) GVU Formular:     	liest den Bestand an gesetzten Formularfeldern ein Objekt, vergleicht diese mit dem vorherigen Bestand
	;  -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		; Loop der auf die Fenster ALTE FORMULARDATEN ÜBERNEHMEN, KEINE ALTE DATEN VORHANDEN oder Datenübernahme ohne Popup Fenster wartet und dann entsprechendes ausführt
			Loop
			{
					ControlsYet:= AlbisMuster30ControlList()
					For key, val in ControlsFirst
					{
							sl:= key
							cf:= ControlsFirst[sl]
							cy:= ControlsYet[sl]
							If (ControlsFirst[sl] <> ControlsYet[sl])
									WinWhat:=3
					}

					If WinExist("Alte Formulardaten")
										WinWhat:= 1
					else if WinExist("ALBIS", "Keine alten Daten")
										WinWhat:= 2

					sleep, 100

			} until (WinWhat > 0)

		;dieser Teil gibt die Objekte wieder frei, RAM sparen!
			For key, val in ControlsFirst
					ControlsFirst.Delete(key)

			For key, val in ControlsYet
					ControlsYet.Delete(key)
	;}

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  e) GVU Formular:     	Befüllen des Formulars entweder mit alten Daten ansonsten werden Häkchen bei 'Beweg.apparat' und bei 'Blutdruck' gesetzt
	;  -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
			If (WinWhat = 1)
			{
							WinActivate, Alte Formulardaten übernehmen
							WinWaitActive, Alte Formulardaten übernehmen, ,3
							Send, {SPACE}
							Sleep, % sleeptime
							Send, {ENTER}
							LV_Add("",  Substr("00" c, -1) ": Alte Formulardaten übernommen")
							Sleep, % sleeptime
			}
			else if (WinWhat = 2)
			{
							VerifiedClick("Button1", "ALBIS","Keine alten Daten")
							LV_Add("",  Substr("00" c, -1) ": Fenster: ''keine alten Daten'' - geschlossen")
							Sleep, % sleeptime
							VerifiedClick("Button52", "Muster 30 ahk_class #32770")
							LV_Add("",  Substr("00" c++, -1) ": Häkchen bei ''Beweg.apparat'' gesetzt")
							Sleep, % sleeptime
							VerifiedClick("Button57", "Muster 30 ahk_class #32770")
							LV_Add("",  Substr("00" c++, -1) ": Häkchen bei ''Blutdruck'' gesetzt")
			}

			;bei WinWhat = 3 - also ohne Fenster "Alte Formulardaten" muss nur auf "Weiter" (Button61) gedrückt werden, das erfolgt als nächster Befehl bei f)

	;}

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  f) GVU Formular:     	Button WEITER wird gedrückt und es wird geprüft ob die 2.Seite des GVU Formular erscheint
	;  -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
			LV_Add("",  Substr("00" c++, -1) ": warte auf GVU Seite 2")
			while !WinExist("Gesundheitsuntersuchung (Seite 2)")
			{
					VerifiedClick("Button61", "Muster 30 ahk_class #32770")
					If ( A_Index > 10 )
					{
						MsgBox,, Addendum für AlbisOnWindows, % MsgTxt007    ;Das GVU-Formular konnte nicht aufgerufen werden..... manuell öffnen....
							IfMsgBox, No
								ExitApp
					}
					Sleep, % sleeptime
			}
	;}

	;  -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  g) GVU Formular:     	Hier geht es weiter mit dem Ausfüllen der 2.Seite des Formular
	;  -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
			WaitAndActivate( "Gesundheitsuntersuchung (Seite 2)", 1, "AC1", 3)

			VerifiedCheck("Button37", "Gesundheitsuntersuchung (Seite 2)")															; Button37 - "Orthopädischer Behandlungsbedarf"
			VerifiedCheck("Button52", "Gesundheitsuntersuchung (Seite 2)")
			VerifiedCheck("Button54", "Gesundheitsuntersuchung (Seite 2)")
			  VerifiedClick("Button61", "Gesundheitsuntersuchung (Seite 2)")	                            							; Button61 - "Speichern" und weiter geht es

			While WinExist("Gesundheitsuntersuchung (Seite 2)")
			{
					WinActivate, Gesundheitsuntersuchung (Seite 2)
					Sleep, % sleeptime
					Send, {SPACE}
					Sleep, % sleeptime
					Send, {ENTER}
					Sleep, % sleeptime
			}

			LV_Add("",  Substr("00" c++, -1) ": ''Gesundheitsuntersuchung (Seite 2)'' ausgefüllt")
			LV_Add("",  Substr("00" c, -1) ": warte auf Fenster: ''GNR-Vorschlag''")
	;}
	*/

			/*
		;~ ; Eingabebereich der Patientenakte um dieses Fenster zu aktivieren, sonst können dort keine Daten eingetragen werden!
			;~ AlbisPrepareInput("lko")
			;~ hFControl	:= GetFocusedControl()
			;~ hParent		:= GetAncestor(hFControl, 1)

		;~ ; Eintragen der Daten und Abrechnungsziffern
			;~ VerifiedSetText("Edit2"            	, lvar[1]             	, "ahk_id " hParent, 100)			;. " ahk_exe " . PAexe
			;~ VerifiedSetText("Edit3"            	, "lko"                	, "ahk_id " hParent, 100)			;. " ahk_exe " . PAexe
			;~ VerifiedSetText("RichEdit20A1"	, "01732-01746 "	, "ahk_id " hParent, 100)			;. " ahk_exe " . PAexe

		;~ ; Fokus auf das letzte Feld der Zeile setzen
			;~ ControlFocus, RichEdit20A1, % "ahk_id " hParent
			;~ Sleep, % sleeptime

		;~ ; abschließen der Eingaben durch Senden der 'Tab'-Taste
			;~ SendInput, {Tab}
			;~ Sleep, % sleeptime

					; Escape senden falls der Focus in einem Edit oder RichEditControl noch sein sollte, dann würde die Akte nicht ordentlich zu schliessen sein
			;If InStr(GetFocusedControlClassNN(), "RichEdit20A1") || InStr(GetFocusedControlClassNN(), "Edit")
			;	SendInput, {Escape}
		 */









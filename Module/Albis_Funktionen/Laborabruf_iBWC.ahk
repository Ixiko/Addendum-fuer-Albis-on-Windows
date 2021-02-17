; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                          	** elektronischer Laborabruf ** Automatisierung für infoBoxWebClient.exe
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;
;		Beschreibung:
;
; 		Inhalt:
;       Abhängigkeiten:
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Laborabruf_iBWC.ahk last change:    	15.02.2021
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  ; -------------------------------------------------------------------------------------------------------------------------------------------------------
  ; Einstellungen
  ; -------------------------------------------------------------------------------------------------------------------------------------------------------;{
	#NoEnv
	#Persistent
	#SingleInstance               	, Force
	#MaxThreads                  	, 100
	#MaxThreadsBuffer       	, On
	#KeyHistory						, Off
	;#MaxMem 						4096
	SetBatchLines					, -1
	;ListLines    						, Off
	;}

  ; -------------------------------------------------------------------------------------------------------------------------------------------------------
  ; Variablen
  ; -------------------------------------------------------------------------------------------------------------------------------------------------------;{
	global ThreadControl := true
	global qq := Chr(0x20)
	global amsg
	global adm := Object()

	adm.Labor           	:= Object()
	adm.scriptname	:= RegExReplace(A_ScriptName, "\.ahk$")
	adm.ID             	:= GetAddendumID()
	;}

  ; -------------------------------------------------------------------------------------------------------------------------------------------------------
  ; Pfade und Einstellungen
  ; -------------------------------------------------------------------------------------------------------------------------------------------------------;{

		admLaborAbrufGui()

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Pfad zu Addendum und den Einstellungen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)
		adm.ini	:= AddendumDir "\Addendum.ini"
		If !FileExist(adm.ini) {
			MsgBox, 1024, % adm.scriptname, % (amsg .= datestamp(2) "Die Einstellungsdatei ist nicht vorhanden!`n`t[" adm.ini "]")
			ExitApp
		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Backup Pfad für empfangene Labordateien einlesen und bei Bedarf anlegen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
		adm.AlbisPath   	:= AlbisPath
		adm.AlbisDBPath 	:= AlbisPath "\db"

		workini	:= IniReadExt(AddendumDir "\Addendum.ini")
		If (StrLen(workini) = 0) {

		}
		adm.DBPath := IniReadExt("Addendum", "AddendumDBPath")           	; Datenbankverzeichnis
		If InStr(adm.DBPath, "ERROR") {
			MsgBox, 1024, % adm.scriptname, % "In der Addendum.ini ist kein Pfad für die Addendum Datendateien hinterlegt!`n"
			ExitApp
		}
		If !FilePathCreate(adm.DBPath "\Labordaten") {
			MsgBox, 1024, % adm.scriptname, % "Der Pfad für das Backup der empfangenen`n`tLabordateien konnte nicht angelegt werden.`n`t[" adm.DBPath "\Labor" "]`n"
			ExitApp
		}

		adm.LogFilePath := adm.DBPath "\Labordaten\LaborabrufLog.txt"

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Skripteinstellungen laden
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		; Einstellungen laden
			adm.Labor.ExecuteFile      	:= IniReadExt("LaborAbruf"	, "Laborabruf_Extern"	, ""            	)  	; externes Programm das für den Abruf ausgeführt werden muss
			adm.Labor.LDTDirectory   	:= IniReadExt("LaborAbruf"	, "LDTDirectory"     	, "C:\Labor"	)
			adm.Labor.LaborName    	:= IniReadExt("LaborAbruf"	, "LaborName"        	, ""             	)
			adm.Labor.Laborkuerzel   	:= IniReadExt("LaborAbruf"	, "Aktenkuerzel"      	, "labor"    	)

		; Laborname ist unbekannt. Skript holt sich den Namen aus einer Albisdatenbank.
			If (StrLen(adm.Labor.LaborName) = 0 || InStr(adm.Labor.LaborName, "ERROR")) {
				labore := GetDBFData(adm.AlbisDBPath "\LABSTAMM.DBF" , {"ID":"rx:\d+"}, ["ID", "NAME", "PFAD"], 0)
				If (labore.Count() > 1) {
					For idx, labor in labore
						laborliste .= (idx > 1 ? "|" : "") labor.NAME
					LNr := ListBoxAbfrage(Laborliste, "Laborabruf", labore.Count() " Labore gefunden. Wähle ein Labor aus!")
					If (StrLen(LNr) = 0) {
						MsgBox, 1024, % adm.scriptname, % "Sie haben kein Labor ausgewählt!`n", 6
						ExitApp
					}
				}
				else
					LNr := 1

				adm.Labor.LaborName 	:= labore[LNr].NAME
				adm.Labor.LDTDirectory	:= labore[LNr].PFAD
				IniWrite, % adm.Labor.LaborName	, % adm.ini, Laborabruf, LaborName
				IniWrite, % adm.Labor.LDTDirectory	, % adm.ini, Laborabruf, LDTDirectory
			}

		; Programm für den Labodatendownload bekannt/vorhanden?
			If (StrLen(adm.Labor.ExecuteFile) = 0 || InStr(adm.Labor.ExecuteFile, "ERROR")) {
				MsgBox, 1024, % adm.scriptname, % " Das Programm für den Download`n`tder Labordaten ist nicht vorhanden!`n`t[" adm.Labor.ExecuteFile "]`n",12
				FileSelectFile, exeFile, 3, % "C:\", % "Wählen Sie die InfoBoxWebClient.exe aus", ausführbare Dateien (*.exe)
				if ErrorLevel {
					MsgBox, 1024, % adm.scriptname, % " Sie haben nichts ausgewählt!`n"
					ExitApp
				}
				If FileExist(exeFile) {
					adm.Labor.ExecuteFile := exeFile
					IniWrite, % exeFile, % adm.ini, Laborabruf, Laborabruf_Extern
				}
			}
			If !FileExist(adm.Labor.ExecuteFile) && !debug {
				FileAppend, % datestamp(2) " | infoBoxWebClient.exe`t- ist nicht vorhanden. " adm.Labor.ExecuteFile "`n", % adm.LogFilePath
				MsgBox, 1024, % adm.scriptname, % "[infoBoxWebClient]`nDas auszuführende externe Programm für den`nelektronischen Datenabruf ist nicht vorhanden`n[" adm.Labor.ExecuteFile "]", 12
				ExitApp
			}

	;}

  ; -------------------------------------------------------------------------------------------------------------------------------------------------------
  ; Laborabruf - Automatisierung (infoBoxWebClient und Labordatenimportieren)
  ; -------------------------------------------------------------------------------------------------------------------------------------------------------;{
		; Albis ist nicht gestartet: Abbruch des Laborabrufes
			If !WinExist("ahk_class OptoAppClass") {
				FileAppend	, % (amsg .= datestamp(2) "| " A_ThisFunc "()`t- Albis ist nicht gestartet`n")
									, % adm.LogFilePath
				ExitApp
			}

		; Automatisierungsfunktionen von Addendum.ahk ausstellen per IPC!
		; Addendum.ahk automatisiert den Laborabruf zum Teil und würde wenn die Funktionen
		; nicht angehalten werden in fehlerhafter RPA-Ausführung enden
			If !(writeStr := admLaborAbrufToggle("off")) {
				MsgBox, 1024, % adm.scriptname, % (amsg := writeStr), 6
				FuncExitApp()
			}

		; 	neues Tagesdatum in die Logdatei schreiben
			FileAppend	, % "---------------------------------------- " A_DD "." A_MM "." A_YYYY "," A_Hour ":" A_Min " ----------------------------------------`n"
								, % adm.LogFilePath

		; 1. externes Programm für Download der LDT Datei starten
			If !infoBoxWebClient()
				FuncExitApp()

		; 2. LDT Datei importieren
			If !AlbisLaborImport(adm.Labor.LaborName)
				FuncExitApp()

		; 3. Laborbuch - alle ins Laborblatt übertragen
			AlbisOeffneLaborbuch()

		; Automatisierungsfunktionen im Hauptskript wieder anschalten
			RestoreLaborAutomationSettings:
			FuncExitApp()

	;}


ExitApp

; -------------------------------------------------------------------------------------------------------------------------------------------------------
; RPA Funktionen
; -------------------------------------------------------------------------------------------------------------------------------------------------------;{
infoBoxWebClient() {                                                               	;-- Programm für elektronischen Labordatenempfang

	; die Funktion startet die Client.exe, wartet auf die Beendigung des Abrufes und legt ein Dateibackup
	; der empfangenen .LDT Datei an und schließt das externe Programmfenster im Anschluß

	; externes Programm starten
		vianovaClass := "vianova infoBox-webClient" ;" ahk_class WindowsForms10.Window.8"
		amsg .= datestamp(2) "|" A_ThisFunc "() `t- starte externes Programm für den Laborabruf`n"
		Run, % adm.Labor.ExecuteFile
		WinWait, % vianovaClass,, 10
		If !(hVianova := WinExist(vianovaClass)) {
			FileAppend	, % (amsg .= datestamp(2)  "|" A_ThisFunc "() `t- infoBoxWebClient: Fenster nicht gefunden`n")
								, % adm.LogFilePath
			MsgBox, 1024, % adm.scriptname, % "Fenster: vianova infoBoWebClient `nwurde nicht gefunden!`n"
			return 0
		}

	; Abruf Button drücken
		ctrlClass := Controls("", "ControlFind, Abruf, Button, return classNN", hVianova)
		VerifiedClick(ctrlClass, hVianova)

	; wartet bis keine Daten mehr hinzukommen
		amsg .= datestamp(2) "|" A_ThisFunc "() `t- rufe Labordaten ab`n"
		adm.Labor.Daten := false
		Loop {
			oAcc := Acc_Get("Object", "4.1.4.5.4.1.4.4.4" , 0, "ahk_id " hVianova)
			t := ""
			Loop % oAcc.accChildCount()	{

				 oAcc := Acc_Get("Object", "4.1.4.5.4.1.4.4.4." A_Index, 0, "ahk_id " hVianova)
				 t .= oAcc.accValue(2) "`n"

				 If RegExMatch(t, "(Dateien\serhalten)|(Sichere.*ins\sArchiv)")
					adm.Labor.Daten := true

				If RegExMatch(t, "Ende\sDateien\sabrufen")
					break

			}

			If (A_Index > 600) || adm.Labor.Daten || RegExMatch(t, "Ende\sDateien\sabrufen")
				break
			sleep 50

		}

	; Dateieingang prüfen
		FileEncoding, CP28605   ; spezielles LDT Encoding (kopierte Datei ist sonst nicht lesbar)
		Loop, Files, % adm.Labor.LDTDirectory "\*.LDT"
		{
				FileGetTime, LDTTime, % A_LoopFileFullPath, M
				LDTTime := SubStr(LDTTime, 1 , 8)
				ANow := A_YYYY . A_MM . A_DD

			; eine neue Datei ist da
				If (ANow = LDTTime) {
					datestamp := (A_YYYY "-" A_MM "-" A_DD " " A_Hour "_" A_Min "_" A_Min)
					FileCopy, % A_LoopFileFullPath, % adm.DBPath "\Labordaten\" datestamp . A_LoopFileName
					writeText .= A_LoopFileName ", "
				}

		}

	; Logdaten schreiben
		IniWrite	, % datestamp(1) , % adm.ini, % "LaborAbruf", % "Letzter_Abruf"
		If (StrLen(writeTExt) > 0) {
			IniWrite			, % datestamp(1) "|0", % adm.ini, % "LaborAbruf", % "Letzter_Abruf_mit_Daten"
			FileAppend	, % (amsg .= datestamp(2)  "| " A_ThisFunc "()  `t- # Daten vorhanden [" RTrim(writeText, ", ") "]`n")
								, % adm.LogFilePath, UTF-8
			LabDataReceived := true
		}	else {
			IniWrite			, % datestamp(1) , % adm.ini, % "LaborAbruf", % "Letzter_Abruf_ohne_Daten"
			FileAppend	, % (amsg .= datestamp(2) "| " A_ThisFunc "()  `t- # Keine Daten vorhanden`n")
								, % adm.LogFilePath, UTF-8
			LabDataReceived := false
		}

	; externes Programmfenster schließen
		amsg .= datestamp(2) "|" A_ThisFunc "()  - WebClient Fenster wird geschlossen`n"
		sleep 500
		Controls("","reset","")
		hVianova := WinExist(vianovaClass)
		ctrlClass := Controls("", "ControlFind, Schließen, Button, return classNN", hVianova)
		WinActivate, % vianovaClass
		WinWaitActive, % vianovaClass,, 2
		If !VerifiedClick(ctrlClass, hVianova) {
			sleep 500
			VerifiedClick(ctrlClass, hVianova)
		}



return LabDataReceived
}

AlbisLaborImport(LaborName) {												;-- Labordaten importieren

	; Aufruf des Dialoges Labor wählen über das Menu Extern\Labor\Daten importieren ....
		hLaborwaehlen := AlbisLaborwaehlen()
		If !VerifiedChoose("ListBox1", hLaborwaehlen, LaborName) {
			FileAppend	, % (amsg .= datestamp(2)  "| " A_ThisFunc "() `t- Labor wählen: Listboxeintrag " LaborName
								. 	 " konnte nicht ausgewählt werden.`n")
								, % adm.LogFilePath
			return 0
		}

	; OK drücken, 2 Sekunden auf das Schliessen des Fenster warten
	; der Hinweisdialog 'Keine Datei(en) im Pfad' wird abgefangen und geschlossen, das Skript bricht dann hier ab
		err    	:= VerifiedClick("Button1", hLaborwaehlen,,, 2)
		amsg 	.= datestamp(2) "| " A_ThisFunc "() `t- Warte auf den nächsten Laborabruf-Dialog"
		WinWait, ALBIS, Keine Datei(en) im Pfad, 2
		If (ErrorLevel = 0) 	{
			FileAppend	, % 	(amsg .= datestamp(2) "| " A_ThisFunc "() `t- Labor wählen: keine Daten im Ordner "
								.   	adm.Labor.LDTDirectory "`n")
								, % 	adm.LogFilePath
			If !(err := VerifiedClick("Button1", "ALBIS", "Keine Datei(en) im Pfad"))
				FileAppend	, % (amsg .= datestamp(2) "| " A_ThisFunc "() `t- Fenster ''Keine Datei(en) im Pfad'': "
									. 	  "konnte nicht geschlossen werden`n")
									, % adm.LogFilePath
			return err
		}

	; ein anderes Fenster [Labordaten Sammelimport] öffnet sich jetzt
		while !WinExist("Labordaten ahk_class #32770", "Sammelimport") {
			If (A_Index >= 100)
				break
			Sleep 50
			Loop 10
				Sleep 5
		}
		If !WinExist("Labordaten ahk_class #32770", "Sammelimport") {
			FileAppend	, % (amsg .= datestamp(2) "| " A_ThisFunc "() `t- Labordaten Sammelimport: hat sich nicht geöffnet`n")
								, % adm.LogFilePath
			return 0
		}

	; kurze Pause damit sich das Fenster noch komplett aufbauen kann
		Sleep 1500

	; Checkbox: [für Sammelimport vormerken] anhaken
		If !VerifiedCheck("Button5", "Labordaten ahk_class #32770",,, 1) {
			FileAppend	, % (amsg .= datestamp(2) "| " A_ThisFunc "() `t- Labordaten Sammelimport: Checkbox"
								. 	 " ''für Sammelimport vormerken'' konnte nicht gesetzt werden`n")
								, % adm.LogFilePath
			return 0
		}
	; Button: Sammelimport drücken
		If !VerifiedClick("Button1"	, "Labordaten ahk_class #32770",,, 3) {
			FileAppend	, % (amsg .= datestamp(2)  "| " A_ThisFunc "() `t- Labordaten Sammelimport: "
								. 	 "konnte nicht geschlossen werden`n")
								, % adm.LogFilePath
			return 0

		}

	; Import: 	Fortschrittsanzeige detektieren und solange warten bis der Import abgeschlossen ist
	; 				wartet auf das erste Popupfenster und wartet bis dieses geschlossen wurde
		PopupWin := false
		writeStr := datestamp(2) "| " A_ThisFunc "() `t- Labordaten Sammelimport: letztes Fenster (Importfortschritt) konnte nicht abgefangen werden"
		Loop {

			If (A_Index >= 80) {
				FileAppend, % (amsg .= datestamp(2) writeStr " [1]`n"), % adm.LogFilePath
				return 0
			}

			hPopupWin := GetHex(DllCall("GetLastActivePopup", "uint", AlbisWinID()))
			If (AlbisWinID() <> hPopupWin) {

				WinGetTitle	, PopTitle	, % "ahk_id " hPopupWin
				WinGetClass	, PopClass, % "ahk_id " hPopupWin
				while WinExist(PopTitle " ahk_class " PopClass) {
					If (A_Index >= 80) {
						FileAppend, % (amsg .= datestamp(2) writeStr " [2]`n"), % adm.LogFilePath
						break
					}
					SleepLoop(50, 10, 5)
				}

				hPopupWin := GetHex(DllCall("GetLastActivePopup", "uint", AlbisWinID()))
				If (AlbisWinID() <> hPopupWin)
					return 0
				else
					return 1
			}

			SleepLoop(50, 10, 5)

		}

}

;}

; -------------------------------------------------------------------------------------------------------------------------------------------------------
; Helfer
; -------------------------------------------------------------------------------------------------------------------------------------------------------;{
SleepLoop(a, b, c) {

	Sleep % a
	Loop % b
		Sleep % c

}

ListBoxAbfrage(liste, GTitle:="", GText:="") {

	global

	inputDone := false

	Gui, LBW: Default
	Gui, +AlwaysOnTop +LastFound -DPIScale +HWNDhLBW
	Gui, Color, cFFCCAA, c777744
	Gui, Font, s11 bold	, Calibri
	Gui, Add, Text   	, xm ym w270, % GText
	Gui, Font, s11 cWhite Normal, Calibri
	Gui, Add, ListBox	, y+10 w270 r5 0x100 vLBR AltSubmit, % liste
	Gui, Font, s11 cBlack Normal, Calibri
	Gui, Add, Button	, xm y+10 vLBWOK 	gBtnHandler, OK
	Gui, Add, Button	, x+20 vLBWCancel 	gLBWGuiEscape, Abbruch
	Gui, Show, AutoSize, Laborabruf

	while !inputdone || WinExist(GTitle " ahk_class AutoHotkeyGUI") {
		Sleep 10
		Loop 5
			Sleep 2
	}

	Gui, LBW: Submit, Hide
	Gui, LBW: Destroy

return LBR

LBWGuiClose:
LBWGuiEscape:
	inputDone := true
	LBR := ""
	Gui, LBW: Destroy
return

BtnHandler:

	Gui, LBW: Submit, Hide
	Gui, LBW: Destroy
	inputDone := true

return LBR
}


; - - - - -
admLaborAbrufGui(Title:="") {

		global

		GuiTitle := (StrLen(Title) = 0 ? adm.scriptname : Title) " - IPC Gui"

	; IPC Gui erstellen
		AlbisWinID:= WinExist("ahk_class OptoAppClass")
		AlbisPos	:= GetWindowSpot(AlbisWinID)
		LCw      	:= 480, LCh := 300

		Gui, LC: New, +hwndhLabCall +ToolWindow -Caption +AlwaysOnTop
		Gui, LC: Margin, 0, 0
		Gui, LC: Color, c172842, c172842
		Gui, LC: Font, s8 q5 cWhite, Consolas
		amsg := "----------------------------------------------------------`n"
		amsg .= datestamp(1) "  " adm.scriptname " gestartet`n"
		amsg .= "----------------------------------------------------------`n`n"
		Gui, LC: Add, Edit     	, % "xm ym w" LCw " h" LCh " vLCdbg hwndhLCdbg", % amsg
		Gui, LC: Add, Progress	, % "xm y+0 w" LCw " h20 Background172842 cWhite vLCPrg hwndhLCPrg", 100

		Gui, LC: Show, % "x" (AlbisPos.X+7) " y" (AlbisPos.Y+AlbisPos.H-LCh-30) " w" LCw " h" LCh+20 " NA", % GuiTitle ;AlbisPos.W-LCw-
		WinSet, Trans, 220, % "ahk_id " hLabCall
		OnMessage(0x4a, "Receive_WM_COPYDATA")
		blockSendWMCOPYDATA := false

	; keine Nachrichten senden, wenn Addendum.ahk nicht läuft
		if !adm.ID {
			blockSendWMCOPYDATA := true
			return 0
		}
		else if blockSendWMCOPYDATA
			return 0

	; Autoupdate Debug Gui
		func_Autoupdate := Func("admAutoupdateLCdbg")
		SetTimer, % func_Autoupdate, 50

}

admAutoupdateLCdbg() {

	global LC, LCdbg, hLCdbg
	static vOutput, RoundZ

	If (StrLen(amsg) = 0)
		return

	vOutput .= RTrim(RegExReplace(amsg, "\s+\t", "  "), "`n") "`n"
	amsg := ""
	GuiControl, LC:, %hLCdbg%, % vOutput

}

admLaborAbrufToggle(Switch="Off", Title:="") {

	; sendet eine Nachricht an Addendum.ahk, um die Laborabrufautomatisierung an- oder auszuschalten
	; beim ersten Aufruf wird eine versteckte Gui erstellt (die ist notwendig für die Interprocesskommunikation)
	;
	; Switch:                        	off   	- 	Befehl zum Ausschalten senden
	;				on/restore	[destroy]	- 	Befehl zum Anschalten oder was dasselbe ist, restore um das Laborabruf flag wiederherzustellen
	;				      								optional oder auch allein kann destroy übergeben werden um die IPC-MessageGui zu schliessen
	; 													und OnMessage zu beenden

		global


	; sendet eine Nachricht an Addendum.ahk
		if RegExMatch(Switch, "i)(on|restore|off)") {

			; ThreadControl wird hier immer true gesetzt
				ThreadControl := true, eLvl := 1

			; Nachricht an Addendum.ahk senden
				If InStr(Switch, "off") {
					writeStr := " |" A_ThisFunc "()`t- Addendum hat nicht geantwortet!`nDie Automatisierungsfunktionen konnten nicht angehalten werden."
					eL := Send_WM_COPYDATA("AutoLaborAbruf|aus|" hLabCall, adm.ID)
				}
				else if RegExMatch(Switch, "i)(on|restore)") {
					writeStr := " |" A_ThisFunc "()`t- Addendum hat nicht geantwortet!`nDie Automatisierungsfunktionen konnten nicht wiederhergestellt werden."
					eL := Send_WM_COPYDATA("AutoLaborAbruf|an|" hLabCall, adm.ID)
				}

			; auf Antwort von Addendum.ahk warten (max. 5s wartet das Skript hier auf eine Änderung der ThreadControl Variable)
			; die Messageworker() Funktion ändert die Variable auf false sobald eine Antwort von Addendum.ahk eingetroffen ist
				while, (ThreadControl && (A_Index <= 100)) {
					ToolTip, % "Laborabruf wartet auf Addendum: " ThreadControl "`nAddendumID: " adm.ID ", " hLabCall "`nEL: " eL "`nA_Index: " A_Index "/" 100
					Sleep 50
				}
				ToolTip

			; keine Antwort, dann notieren und Fehler zurückmelden
				If ThreadControl {
					FileAppend, % (amsg .= datestamp(2) " |" StrReplace(writeStr, "`n", " ")) "`n", % adm.LogFilePath
					eLvl := 0
				}

		}

	; IPC-MessageGui schliessen, OnMessage beenden
		If RegExMatch(Switch, "i)destroy")
			return 1


return eLvl
}

; - - - -
FuncExitApp() {

	global func_Autoupdate, LC, hLCPrg

	admLaborAbrufToggle("restore destroy")
	OnMessage(0x4a, "")
	SetTimer, % func_Autoupdate, Delete
	wTime := 20000
	Interrupts := 100
	Loop, % Floor(wTime/Interrupts) {
		ToolTip % 100 - Floor((100*(A_Index*interrupts))/wTime)
		GuiControl, LC:, %hLCPrg%, % 100 - Floor((100*(A_Index*interrupts))/wTime)
		sleep % Interrupts
	}
	ExitApp

}

MessageWorker(msg) {                                                       		;--  verarbeitet die eingegangen Nachrichten

		recv := {	"txtmsg"		: (StrSplit(msg, "|").1)
					, 	"opt"     	: (StrSplit(msg, "|").2)
					, 	"fromID"	: (StrSplit(msg, "|").3)}

		;SciTEOutput("`nLabor Incoming: `nmsg: " recv.txtmsg "`nfrom: " recv.fromID "`nThreadControl: " ThreadControl "`n")
	; main script receive's data from this thread and calls to continue the OCR process
		If InStr(recv.txtmsg, "Laborabruf angehalten")            	; Thread waits until ThreadControl is false
			ThreadControl := false
		else If InStr(recv.txtmsg, "Laborabruf fortgesetzt")                ; to stop on emergency
			ThreadControl := false

		PraxTT(recv.txtmsg, "6 1")

return
}


;}

; -------------------------------------------------------------------------------------------------------------------------------------------------------
; Includes
; -------------------------------------------------------------------------------------------------------------------------------------------------------;{
#Include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Datum.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PDFHelper.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PraxTT.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Window.ahk

#Include %A_ScriptDir%\..\..\lib\acc.ahk
#Include %A_ScriptDir%\..\..\lib\class_JSON.ahk
#Include %A_ScriptDir%\..\..\lib\class_Neutron.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\lib\Sift.ahk
;}





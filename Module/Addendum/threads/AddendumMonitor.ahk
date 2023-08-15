;-----------------------------------------------------------------------------------------------------------------------------------
;------------------------------------------------ ADDENDUM MONITOR ----------------------------------------------------
;-----------------------------------------------------------------------------------------------------------------------------------
													Version:= "1.29" , vom:= "23.12.2022"
;------------------------------------------------------ Runtime-Skript ----------------------------------------------------------
;-------------------- startet Addendum bei einem Absturz oder (un-)absichtlichen Beenden neu ----------------------
;-------------------------------------------- Addendum für AlbisOnWindows -----------------------------------------------
;-----------------------------------------------------------------------------------------------------------------------------------

D1:= "AddendumMonitor V" Version " - The monitoring script - "
D2=
(LTrim
Addendum für Albis on Windows
written by Ixiko - this version %Version% is from %vom%
---------------------------------------------------------------------------------
please report errors and suggestions to me: Ixiko@mailbox.org
use subject: 'Addendum' so that you don't end up in the spam folder
GNU Lizenz can be found in Docs directory  - 2017
)

;{ Skriptablaufeinstellungen

	#NoEnv
	#Persistent
	#SingleInstance			, Force
	#MaxThreadsPerHotkey , 2
	#KeyHistory              	, 0
	SetTitleMatchMode    	, 2        	; Fast is default
	SetTitleMatchMode    	, Fast    	; Fast is default
	DetectHiddenWindows	, On      	; Off is default
	CoordMode, ToolTip 	, Screen
	SetBatchLines            	, -1
	SetWinDelay               	, -1
	SetControlDelay        	, -1
	SendMode                	, Input
	AutoTrim                   	, On
	FileEncoding             	, UTF-8

	OnError("FehlerProtokoll")

;}

  ; Variablen   ;{

	global RestartProcess
	global ObservationRuns	:= false
	global DoRestart          	:= true
	global q                       	:= Chr(0x22)
	global Addendum      	:= Object()
	global overwatch         	:= Array()
	global sinks

  ; Addendum object
	RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)
	Addendum.compname    	:= StrReplace(A_ComputerName, "-")
	Addendum.ScriptName    	:= "Addendum.ahk"
	Addendum.Dir               	:= AddendumDir
	Addendum.Ini                	:= AddendumDir 	"\Addendum.ini"
	Addendum.ScriptPath      	:= AddendumDir 	"\Module\Addendum\" Addendum.ScriptName
	Addendum.Restarts        	:= 0

	If FileExist(A_AppData "\AutoHotkeyH\AutoHotkeyH_U64.exe")
		Addendum.AHKH_exe	:= A_AppData "\AutoHotkeyH\AutoHotkeyH_U64.exe"
	else if FileExist(AddendumDir "\include\AHK_H\x64w\AutoHotkeyH_U64.exe")
		Addendum.AHKH_exe	:= Addendum.Dir "\include\AHK_H\x64w\AutoHotkeyH_U64.exe"
	else
		throw A_ScriptName ": AutohotkeyH_U64.exe ist nicht vorhanden.`nDas Skript kann nicht ausgeführt werden!"
  ;}

  ; TrayMenu ;{
	Menu, Tray, NoStandard
	If (StrLen(hIcon := AddendumMonitor_ico()) > 0)
		Menu, Tray, Icon, % "HICON: " hIcon
	else If FileExist(IconDir := Addendum.Dir "\assets\ModulIcons\AddendumMonitor.ico")
		Menu, Tray, Icon, % IconDir

	;func_AddendumCheck := Func("RestartAddendum").Bind("override")
	Menu, Tray, Add, % StrReplace(A_ScriptName, ".ahk")	, % "admMonInfo"
	Menu, Tray, Add
	Menu, Tray, Add, % "Addendum jetzt prüfen"            	, % "admMonCheck"
	Menu, Tray, Add, % "Addendum (neu)starten"           	, % "admRestartNow"
	Menu, Tray, Add, % "Überwachungsmanager öffnen"	, % "MonManager"
	Menu, Tray, Add
	Menu, Tray, Add, % "Reload"		                            	, % "admMonReload"
	Menu, Tray, Add, % "Exit"			                             	, % "admMonExitApp"
  ;}

  ; get data from Addendum.ini
	IniReadExt(Addendum.Ini)

  ; Addendum DBPath und Logfile Pfad ;{
	Addendum.DBPath	:= IniReadExt("Addendum", "AddendumDBPath")
	If InStr(Addendum.DBPath, "Error") || !RegExMatch(Addendum.DBPath, "i)[a-z]\:\\") || !InStr(FileExist(Addendum.DBPath), "D") {
		Addendum.DBPath := ""
		MsgBox, Das Datenverzeichnis von Addendum für Albis on Windows ist konnte nicht ermittelt werden. Bitte hinterlegen Sie dieses in der Addendum.ini
		ExitApp
	}

	Addendum.LogPath	:= IniReadExt("Addendum", "AddendumLogPath")
	If InStr(Addendum.LogPath, "Error") || !RegExMatch(Addendum.LogPath, "i)[a-z]\:\\") || !InStr(FileExist(Addendum.LogPath), "D") {
		MsgBox, Das Datenverzeichnis von Addendum für Albis on Windows ist konnte nicht ermittelt werden. Bitte hinterlegen Sie dieses in der Addendum.ini
		ExitApp
	}
	Addendum.MonLogPath := Addendum.LogPath "\AddendumMonitorLog.txt"

	;~ If Addendum.DBPath && !InStr(FileExist(Addendum.DBPath "\logs"), "D") {
		;~ FileCreateDir, % Addendum.DBPath "\logs"
		;~ If ErrorLevel
			;~ Addendum.MonLogPath := Addendum.DBPath
	;~ }
;}

  ; Interskript communication gui ;{
	;~ while (hwnd := WinExist("Addendum Monitor Gui", "ahk_class AutoHotkeyGUI"))

	Gui, AMsg: New	, +HWNDhMsgGui +ToolWindow
	Gui, AMsg: Add	, Edit, xm ym w100 h100 HWNDhEditMsgGui
	Gui, AMsg: Show	, AutoSize NA Hide, % "Addendum Monitor Gui"

	Addendum.hMsgGui     	:= hMsgGui
	Addendum.hEditMsgGui	:= hEditMsgGui

	; starts receiving messages
	OnMessage(0x4A	, "Receive_WM_COPYDATA")
	OnMessage(0x404, "AHK_NOTIFYICON")

	; }

;}

  ; loading Telegram BotToken and BotChatID ;{
	Addendum.Telegram := Object()
	TelegramBots()
	BotName	:= IniReadExt("Telegram", "Bot1")
  ;}

  ; loads restart properties (script detects high cpu usage of Addendum.ahk and will close-restart Addendum.ahk) ;{
	Addendum.MaxAverageCPU 	:= IniReadExt("AddendumMonitor", "HighCPU_MaxAvarageCPU" 	, 12)
	Addendum.RestartAfter       	:= IniReadExt("AddendumMonitor", "HighCPU_RestartAfter"         	, 360) 	; seconds
	Addendum.MinIdleTime     	:= IniReadExt("AddendumMonitor", "HighCPU_MinIdleTime"         	, 180)	; seconds
	Addendum.TimedCheck     	:= IniReadExt("AddendumMonitor", "HighCPU_Timer"                    	, 180)	; seconds
	Addendum.CrashCheck     	:= IniReadExt("AddendumMonitor", "HighCPU_CrashCheck"          	, 0)		; seconds  (leider werden Abstürze übersehen, lasse den Monitor dazwischen per Timer prüfen)
	TrTip1 := 	"CPU-Last Grenze:    "      	Addendum.MaxAverageCPU 	"%  `n"
			  ;~ .		"Neustart nach:        " 		Addendum.RestartAfter        	"s`n"
			  .		"Überprüfung alle:    "      	Addendum.TimedCheck      	"s"

	TrTip2 := 	"Addendum-Neustarts: "   	Addendum.Restarts               	"s`n"
			 .		"CPU Überlastung:     nicht geprüft`n"
			 . 	  	"nächster Check: "         	Addendum.TimedCheck      	"s"
	Menu, Tray, Tip, % TrTip1 "`n" TrTip2
	Addendum.TimerCall := A_TickCount
  ;}

  ; loads names of other scripts/processes to be monitored, prepares a part of the WQL-Statement
	WQL := OverwatchLoad(Addendum.Ini, "AddendumMonitor")

  ; Sink Objekte erstellen
	sinks := new SinkObjects(WQL)
	sinks.Create()

  ; Timerfunktion für 5 minütige Skriptausführungskontrolle (Skript hängt oder ist abgestürzt)
	If Addendum.TimedCheck {
		fnTimedCheck := Func("RestartAddendum").Bind(true)
		SetTimer, % fnTimedCheck, % Addendum.TimedCheck * 1000
		Addendum.TimerCall := A_TickCount
	}

	If Addendum.ChrashCheck {
		fnCrashCheck := Func("RestartAddendum").Bind(false)
		SetTimer, % fnCrashCheck, % Addendum.CrashCheck * 1000
	}

  ; erster Check von Addendum.ahk
	RestartAddendum()

  ; alle 2h Neustart bis ich den Verursacher für das Speicherbelegungsproblem gefunden habe
	SetTimer, admRestartMonitor, % -2*3600*1000

	TrayTip, AddendumMonitor, Überwachung gestartet, 1

return

;{ Hotkeys

^+!ö::   	;{ ExitApp
	gosub admMonExitApp
return ;}

^!ö::    	;{ Skriptneustart
	gosub admRestartMonitor
return ;}

#IfWinExist ahk_class AutoHotkey2, Neustart
~Esc::  	;{ Tastenkombination um Skriptneustart während der Anzeige des Countdownfenster abzubrechen
	If DoRestart {
		DoRestart := false
		FileAppend, % "Abbruch des Addendumneustart durch Nutzer, Zeit: " TimeStamp() ", Client: " A_ComputerName "`n", % Addendum.LogPath "\OnExit-" A_YYYY "-Protokoll.txt"
	}
return
#IfWinExist


;}

;}

;{ TrayLabels + class SinkObjects

admMonReload:    	;{
sinks.Release()
Reload  ;}

admMonExitApp:    	;{
sinks.Release()
ExitApp ;}

admMonTimerInfo: 	;{
	nextcall 		:= Addendum.TimedCheck-((A_TickCount - Addendum.TimerCall)/1000)
	nextcall_min	:= Floor(nextcall/60)
	nextcall_sec	:= Floor(nextcall - (nextcall_min*60))
	TrTip2       	:=	"Addendum-Neustarts:  "     	Addendum.Restarts               	"x`n"
						.		"CPU Überlastung: "          	(ObservationRuns ? "läuft":"pausiert") "`n"
						. 	 	"Check in:     "                    	nextcall_min "m " SubStr("00" nextcall_sec, -1) "s"
	Menu, Tray, Tip, % TrTip1 "`n" TrTip2
return ;}

admMonInfo:         	;{
return ;}

admMonCheck:    	;{
	RestartAddendum("override")
  ; AddendumMonitor neu laden,wenn Fenster mit dieser Bezeichnung vorhanden sind
  ; unklar warum es bis zu 15 Threads mit sichtbaren (style layered) Gui's überhaupt gibt, solange ich das nicht verhindern kann -> NEU LADEN
	WinGet, wlist, List, % "AddendumMonitor.ahk", % "ahk_class AutoHotkeyGUI"
	If (StrSplit(wlist, "`n").Count() > 1)
		goto admRestartMonitor_NoTrayTip
return ;}

admRestartNow:    	;{

	TrayMsg1 := "Addendum.ahk wird in 2s beendet und neu gestartet!"
	TrayMsg2 := "Addendum.ahk konnte nicht beendet werden!"
	TrayMsg3 := "Addendum.ahk ProcessID= "
	TrayMsg4 := "Addendum.ahk konnte nicht gestartet werden!"

	TrayTip, AddendumMonitor, % TrayMsg1, 2
	Sleep 2000

  ; ein laufendes Addendum beenden
	If (Addendum.PID := AHKProcessExist(Addendum.ScriptName))
		If !ProcessClose(Addendum.PID) {
			TrayTip, AddendumMonitor, % TrayMsg2 , 2
			return
		}

  ; neu starten
	admPID := RestartAddendum("override")
	TrayTip, AddendumMonitor, % (admPID ? TrayMsg3 . admPID : TrayMsg4), 2

return ;}

admRestartMonitor: 	;{
	TrayTip, AddendumMonitor, Skript wird neu gestartet, 1
admRestartMonitor_NoTrayTip:
	sinks.Release()
	Sleep 3000
	Reload
return ;}

MonManager:       	;{

  ;-: Gui            	;{

	If MonManagerExist {
		Gui, MN: Show
		return
	}

	Gui, MN:  New  	, +hwndhMn -DPIScale ;-Theme
	Gui, MN: Margin	, 5, 5
	Gui, MN: Font    	, s9 q5

	Gui, MN: Add, ListView, xm ym w1025 r5 vMNLV hwndMNhLV AltSubmit gMNGuiHandler, Programm|Pfad|cmdline Optionen|Ausführen auf|Delay|Status
	LV_ModifyCol(1, 130)
	LV_ModifyCol(2, 450)
	LV_ModifyCol(3, 150)
	LV_ModifyCol(4, 200)
	LV_ModifyCol(5, 40)
	LV_ModifyCol(6, 45)
	For each, ow in overwatch
		LV_Add("", ow.name, ow.path, ow.opts, ow.runon, ow.delay, (ow.state?"An":"Aus"))

	Gui, MN: Font, s8 q5
	Gui, MN: Add, Text, xm y+5  BackgroundTrans vMNText1                                    	, Dateipfad des Autohotkey Skript/ausführbare Datei

	Gui, MN: Font, s10 q5
	Gui, MN: Add, Edit    	, xm y+2 			w500 	vMNPath       	ReadOnly
	Gui, MN: Add, Text   	, x+2                                                                         		, \
	Gui, MN: Add, Edit    	, x+2   				w150 	vMNName    	ReadOnly
	Gui, MN: Add, Edit    	, x+5    			w75  	vMNDelay     	gMNGuiHandler
	Gui, MN: Add, UpDown, x+5    	        			vMNDelayUD

	Gui, MN: Font, s9 q5
	Gui, MN: Add, Slider 	, x+0 yp-4   		w55 		vMNStateSL  Center NoTicks Thick30 Range0-1 AltSubmit gMNGuiHandler, 0
	Gui, MN: Font, s12
	Gui, MN: Add, Text    	, x+0 yp+6   	w25  	vMNState     	ReadOnly, Aus

	Gui, MN: Font, s8 q5
	Gui, MN: Add, Text   	, xm y+5  BackgroundTrans vMNText2                                    	, cmdline Optionen

	Gui, MN: Font, s10 q5
	Gui, MN: Add, Edit    	, xm y+2 			w500 	vMNOptions
	Gui, MN: Add, Edit    	, x+2    			w500 	vMNRunOn

	Gui, MN: Font, s12 q5
	Gui, MN: Add, Text		, xm y+20 	                                            	, Überwachung
	Gui, MN: Font, s10 q5
	Gui, MN: Add, Button	, x+10 yp-2 	vMNNew  	gMNGuiHandler 	, Hinzufügen
	Gui, MN: Add, Button	, x+10      	vMNDelete   	gMNGuiHandler 	, Entfernen
	Gui, MN: Add, Button	, x+10      	vMNGo     	gMNGuiHandler 	, Änderung übernehmen
	Gui, MN: Add, Button	, x+50      	vMNHide  	gMNGuiHandler 	, Fenster schließen

	Gui, MN: Font, s8 q5
	cp 	:= GuiControlGet("MN", "Pos", "MNText1")
	dp 	:= GuiControlGet("MN", "Pos", "MNName")
	Gui, MN: Add, Text, % "x" dp.X " y" cp.Y  "  BackgroundTrans" 	, Skript/Exe
	dp 	:= GuiControlGet("MN", "Pos", "MNDelay")
	Gui, MN: Add, Text, % "x" dp.X " y" cp.Y  "  BackgroundTrans" 	, Verzögerung
	dp 	:= GuiControlGet("MN", "Pos", "MNStateSL")
	Gui, MN: Add, Text, % "x" dp.X+10 " y" cp.Y  "  BackgroundTrans" 	, Überwachung ist

	cp 	:= GuiControlGet("MN", "Pos", "MNText2")
	dp 	:= GuiControlGet("MN", "Pos", "MNRunOn")
	Gui, MN: Add, Text, % "x" dp.X " y" cp.Y  "  BackgroundTrans" 	, Ausführen auf (* auf jedem Client oder kommagetrennte Liste mit Clientnamen)

	cp 	:= GuiControlGet("MN", "Pos", "MNHide")
	dp 	:= GuiControlGet("MN", "Pos", "MNLV")
	GuiControl, MN: Move, MNHide, % "x" dp.X+dp.W-cp.W

	MNGui_EnableControls(false)

	Gui, MN: Show,, % "[AddendumMonitor] Überwachungsmanager"

	MonManagerExist := hMn

return ;}

MNGuiHandler: 		;{

	Critical
	Gui, MN: Submit, NoHide

	If (A_GuiControl = "MNLV")                    	{

		If (A_GuiEvent = "Normal") {
			row := A_EventInfo
			If overwatch[row].name {
				GuiControl, MN:, MNPath     	, % overwatch[row].path
				GuiControl, MN:, MNName   	, % overwatch[row].name
				GuiControl, MN:, MNOptions	, % overwatch[row].opts
				GuiControl, MN:, MNRunOn 	, % overwatch[row].runon
				GuiControl, MN:, MNDelay   	, % overwatch[row].delay
				GuiControl, MN:, MNStateSL    	, % overwatch[row].state
				GuiControl, MN:, MNState    	, % overwatch[row].state ? "An" : "Aus"

				GuiControl, MN: Enable1, MNDelete
				MNGui_EnableControls(true)
			}
			else {

				GuiControl, MN: Enable0, MNDelete
				MNGui_EnableControls(false)

			}
		}

	}
	else If (A_GuiControl = "MNGo")           	{

	; schaut ob das Programm bereits überwacht wird und speichert nur Änderungen
		foundProc := false
		For owNr, ow in overwatch
			If (ow.name = MNName && ow.path = MNPath) {
				ow.opts 	:= MNOptions
				ow.opts 	:= MNRunOn
				ow.delay 	:= MNDelay
				ow.state 	:= MNStateSL
				foundProc := true
				LV_Modify(owNr,, MNName, MNPath, MNOptions, MNRunOn, MNDelay, (MNStateSL ? "An":"Aus"))
				break
			}

	; Programm ist noch nicht in der Überwachungsliste
		If !foundProc {
			LV_Add("", MNName, MNPath, MNOptions, MNRunOn, MNDelay, (MNStateSL ? "An":"Aus"))
			overwatch.Push({"name":MNName, "path":MNPath, "opts":MNOptions, "runon":MNRunOn, "delay":MNDelay, "state":MNStateSL})
			owNr := LV_GetCount()
		}

	; Einstellungen sichern
		IniWrite	, % MNName "|" MNPath "|" MNOptions "|" MNRunOn "|" MNDelay "|" (MNStateSL ? "An":"Aus")
					, % Addendum.Ini
					, % "AddendumMonitor"
					, % "overwatch" owNr

	; Eingabefelder leeren und Interaktionsmöglichkeiten ausstellen
		MNGui_EnableControls(false)

	; WQL Query vorbereiten
		WQL := OverwatchWQL()

	; WMI Events neu starten
		sinks := new SinkObjects(WQL)
		sinks.Release()
		sinks.Create(WQL)
		TrayTip, AddendumMonitor, WMI Überwachung wurde neu gestartet, 1

	}
	else If (A_GuiControl = "MNNew")         	{

		Thread, NoTimers
		FileSelectFile, overwatchfile,, % A_ProgramFiles, Skript oder ausführbare Datei auswählen, Skript-Exe (*.ahk; *.exe)
		Thread, NoTimers, false

		If !overwatchfile
			return

		SplitPath, overwatchfile, ProcName, ProcPath

		For row, ow in overwatch
			If (ow.name = ProcName && ow.path = ProcPath) {
				MsgBox, 0x1000, % A_ScriptName, % "Diese Datei wird schon überwacht!`nsiehe " ow.name " in Zeile " row "`nKlicken Sie auf die Zeile um die Einstellungen zu ändern!", 10
				return
			}

		GuiControl, MN:, MNPath     	, % ProcPath
		GuiControl, MN:, MNName  	, % ProcName
		GuiControl, MN:, MNOptions	, % ""
		GuiControl, MN:, MNRunOn 	, % "*"
		GuiControl, MN:, MNDelay   	, % "0"
		GuiControl, MN:, MNStateSL    	, % 1
		GuiControl, MN:, MNState    	, % "An"

		MNGui_EnableControls(true)

		ProcName := ProcPath := overwatchfile := ""

	}
	else if (A_GuiControl = "MNStateSL")      	{
		GuiControl, MN:, MNState, % (MNStateSL=1?"An":"Aus")
	}
	else If (A_GuiControl = "MNDelete")         	{

	  ; sucht nach der zu löschenden Zeile
		foundProc := false
		For owNr, ow in overwatch
			If (ow.name = MNName && ow.path = MNPath) {
				MsgBox, 0x1004, % A_Scriptname, % "Überwachung von " MNName " beenden?"
				IfMsgBox, No
					return
				foundProc := true
				break
			}

	  ; LV Zeile und overwatch item entfernen
	  ; inikeys sind durchgehend nummeriert. Lücken schließen.
		If foundProc {

			LV_Delete(owNr)
			overwatch.RemoveAt(owNr)

			; ini Reorganisieren
			For iniIndex, ow in overwatch
				If (iniIndex >= owNr)
					IniWrite	, % ow.name "|" ow.path "|" ow.opts "|" ow.runon "|" ow.delay "|" (ow.state ? "An":"Aus")
					, % Addendum.Ini
					, % "AddendumMonitor"
					, % "overwatch" iniIndex

			; letzten IniKey löschen
				IniDelete, % Addendum.Ini, % "AddendumMonitor", % "overwatch" iniIndex+1


		}

	}
	else If (A_GuiControl = "MNHide")         	{
		Gui, MN: Hide
	}

return ;}

MNGuiClose:        	;{
MNGuiEscape:
	Gui, MN: Hide
return ;}

newline:                  	;{
	SciTEOutput(" ")
return ;}

BTTOff:                  	;{
BTT()
return ;}

MNGui_EnableControls(state) {

	;~ state := state=true ? 1 : 0
	GuiControl, % "MN: Enable" state, MNGo
	GuiControl, % "MN: Enable" state, MNOptions
	GuiControl, % "MN: Enable" state, MNRunOn
	GuiControl, % "MN: Enable" state, MNDelay
	GuiControl, % "MN: Enable" state, MNStateSL

	If !state {

		GuiControl, MN: Enable0, MNDelete

		GuiControl, MN:, MNPath     	, % ""
		GuiControl, MN:, MNName   	, % ""
		GuiControl, MN:, MNOptions	, % ""
		GuiControl, MN:, MNRunOn 	, % ""
		GuiControl, MN:, MNDelay   	, % ""
		GuiControl, MN:, MNState    	, % ""

	}

}
;}

class SinkObjects {

	__New(WQL:="") {		; it does not create sink objects, call next method

		this.WQL := WQL

	}

	Create() {						; this creates the SinkObjects

	  ; Get WMI service object.
		this.winmgmts  	:= ComObjGet("winmgmts:")

	  ; Create sink objects for receiving event noficiations.
		ComObjConnect(this.CreateSink	:= ComObjCreate("WbemScripting.SWbemSink"), "ProcessCreate_")
		ComObjConnect(this.DeleteSink	:= ComObjCreate("WbemScripting.SWbemSink"), "ProcessDelete_")

	  ; Register for process deletion notifications:
		this.winmgmts.ExecNotificationQueryAsync(this.DeleteSink
			, "SELECT * FROM __InstanceDeletionEvent"
			. " within " 1                                                                  	; Set event polling interval, in seconds.
			. " WHERE TargetInstance ISA 'Win32_Process'"
			. " and TargetInstance.Name LIKE 'Autohotkey%'"
			. this.WQL)

	}

	Release() {

		ObjRelease(this.winmgmts)
		ObjRelease(this.CreateSink)
		ObjRelease(this.DeleteSink)

	}

}

;}


;{ Funktionen
RestartAddendum(TimedCheck:=false, dbg:=false) {                                             	;-- startet(prüft) Addendum.ahk

	; Variablen ;{
		global 	ObservationRuns
		static 	msgEnd 	:= "`n---------------------------------------------------------------`n"

		Addendum.TimerCall	:= A_TickCount
		PTSum 		              		:= 0

	;}

	; während Autonaming nichts machen
		If WinExist("Addendum (Datei umbenennen) ahk_class AutoHotkeyGUI")
			return

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Addendum.ahk Prozeß ist nicht vorhanden
		If !(Addendum.PID := AHKProcessExist(Addendum.ScriptName)) {

			If (A_TimeIdlePhysical > 0) {
				idle       	:= GetTimestrings(A_TimeIdlePhysical)
				IdleTime 	:= idle.hour ":" idle.min ":" idle.sec "::" idle.msec
			}

			FileAppend	, %  	TimeStamp()                                                  	", "
								.     	RegExReplace(A_ScriptName, "\.ahk$")            	", "
								.     	A_ComputerName                                           	", "
								.     	(TimedCheck ? "TIMER" : "EVENT")
								.    	(StrLen(idleTime) > 0 ? ", idle: " idleTime : "")  	"`n"
								, % 	Addendum.LogPath "\OnExit-" A_YYYY "-Protokoll.txt", UTF-8

			Run   	    	, % Addendum.AHKH_exe " /f " q Addendum.ScriptPath q

			; wartet max. 20s auf den Autohotkeyprozeß
				while (!(PID := AHKProcessExist(Addendum.ScriptName)) && A_Index <= 200)
					Sleep 100

		return PID
		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Addendum.ahk Prozeß (PID) gefunden
 	; dann wird über die Länge von 10 Sek. die CPU-Auslastung von Addendum messen
		else {


		  ; eine Überwachung läuft noch, dann nichts machen
			If (TimedCheck != "override") && ObservationRuns
				return

			If (TimedCheck = "override")
				TrayTip, Überprüfung von Addendum.ahk, % "Zeitdauer bis Abschluß: circa 10s", 15, 16

		  ; CPU Status (CPU Auslastung überschritten) oder Addendum.ahk antwortet nicht
			ObservationRuns := true
			cpuStatus		:= GetCPUStatus(Addendum.PID, 10, Addendum.MaxAverageCPU)
			ObservationRuns := true
			;~ admScript 	:= GetAddendumStatus()  ; # ausgestellt, ist Addendum beschäftigt antwortet es nicht
			admScript := {"status":true}
			ObservationRuns := false

		  ; admScript.status =  noanswers oder answers, noanswers = var noanswers, oberserveTime = recTime
			If (TimedCheck = "override") {
				healthstate := cpuStatus.HighLoad	? 	(!admScript.status 	? "1.":"")           	"CPU Lastüberschreitung festgestellt (" cpuStatus.Average ")`n"
																	: 	""
				healthstate .= !admScript.status    	? 	(healthstate 	? "           2.":"") "Addendum hat sich möglicherweise aufgehängt (keine Antwort)"
																	: 	(!healthstate && admScript.status ? "Es wurden keine Probleme festgestellt." : "")
				TrayTip	, % "Überprüfung von Addendum.ahk", % "abgeschlossen: " TimeStamp() "`nBewertung: " healthstate, 12, 16
			}

			If (dbg = true || !admScript.status || cpuStatus.HighLoad) {
				SciTEOutput(" " TimeStamp())
				scitemsg := "  "  A_ThisFunc ": Addendum hat innerhalb von " . admScript.oberserveTime . "s " . (!admScript.status ? admScript.noanswers "x nicht" : "") . " geantwortet "
				SciTEOutput(scitemsg)
				SciTEOutput((admScript.status ? "  " A_ThisFunc ": Innerhalb der Beobachtungszeit wurde " (!admScript.allnoanswers ? "immer geantwortet." : admScript.allnoanswers "x nicht geantwortet.") : ""))
				SciTEOutput((cpuStatus.HighLoad ? "  " A_ThisFunc ": cpuStatus.Average: " cpuStatus.Average "%" : ""))
			}

		  ; cpuStatus > 0 - Kontrollwert überschritten, admScript = false = Addendum.ahk hat nicht geantwortet
			If cpuStatus.HighLoad || !admScript.status {


			  ; Protokoll schreiben
				If Addendum.MonLogPath {
					msg :=	(cpuStatus.HighLoad ? "hohe CPU Auslastung registriert (" cpuStatus.Average ")`n                                `t" : "")
					msg .=	"Addendum: " (!admScript ? "keine Antwort" : "Antwort brauchte " admScript " s")
				}

			  ; Addendum sendet keinen Status zurück => Neustart!
				msg .= (msg ? "`n" : "")
				If !admScript.status {
					Addendum.Restarts += 1
					msg .= (!ProcessClose(Addendum.PID) 	?	"Addendum.ahk konnte nicht mittels Process, Close beendet werden."
																				:	"Addendum.ahk wurde mittels Process, Close beendet.")
				}
			  ; HighLoad wurde registriert, Messung der CPU-Last über mehrere Sekunden erfolgt jetzt
				else if cpuStatus.HighLoad {
					Addendum.Restarts += 1
					If (procDelMsg := ProcessDelete_OnHighCPULoad(Addendum.ScriptName
																							,	Addendum.PID
																							,	Addendum.RestartAfter
																							, 	Addendum.IdleTime
																							, 	Addendum.MaxAverageCPU))
							msg .= (msg ? "`n" : "")  procDelMsg
				}

			  ; Tray Symbole beendeter Autohotkey Programme entfernen
				If (!admScript.status || cpuStatus.HighLoad) {
					RefreshTray()
				}

			  ; schreibt das Protokoll
				If (Addendum.MonLogPath && msg)
					FileAppend, % TimeStamp() "|`t" msg . msgEnd, % Addendum.MonLogPath, UTF-8

			}

		}

}

AHK_NOTIFYICON(wParam, lParam) {                                                	;-- OnMessage Callback
	if (lParam = 0x200) { ; WM_LBUTTONUP
		SetTimer, admMonTimerInfo, -1
	return 0
    }
}

AHKProcessExist(ProcName, cmd="") {                                                	;-- is only searching for Autohotkey processes

	; use cmd = "ProcessID" to receive the PID of an Autohotkey process

	StrQuery := "Select * from Win32_Process Where Name Like 'Autohotkey%'"
	For process in ComObjGet("winmgmts:\\.\root\CIMV2").ExecQuery(StrQuery)	{
		RegExMatch(process.CommandLine, "[\w\-\_]+\.ahk", name)
		If InStr(name, ProcName)
			return StrLen(cmd) > 0 ? process[cmd] : process.ProcessID
	}

return false
}

GetCPUStatus(PID,ObserveTime=4, MaxLoad=5, checkInterval:=100) {	;-- Messung der CPU Auslastung

	; Rückgabeparameter ist 0 wenn der Grenzwert nicht überschritten wurden
	; andernfalls wird die ermittelte CPU Last zurückgegeben
	; MaxLoad - Grenzwert der CPU Last, Überschreitung generiert eine HighLoad Meldung

	PTStum  	:= 0
	starttime	:= A_TickCount

	;Loop % (loops := Floor((WatchTime*1000)/checkInterval)) {
	Loop {
		checks    	:=  	A_Index
		CPULoad 	:= 	GetProcessCpu(PID)
		PTSum     += 	CPULoad ? CPULoad : 0
		loopTime 	:= 	Round((A_TickCount - starttime)/1000, 1)
		If (loopTime >= ObserveTime)
			break
		Sleep % checkInterval
	}

	PTAverage	:= Round(PTSum/checks, 2)

return {"Average": PTAverage, "HighLoad": (PTAverage>MaxLoad?1:0)}
}

GetAddendumID() {                                                                          	;-- für Interskriptkommunikation
	PDetectHiddenWindows	:= A_DetectHiddenWindows
    PTitleMatchMode         	:= A_TitleMatchMode
    DetectHiddenWindows On
    SetTitleMatchMode 2
	AddendumID := WinExist("Addendum Message Gui ahk_class AutoHotkeyGUI")
	DetectHiddenWindows 	, % PDetectHiddenWindows  	; Restore original setting for the caller.
    SetTitleMatchMode     	, % PTitleMatchMode           	; Same.
return AddendumID
}

GetAddendumStatus() {                                                                     	;-- testet ob Addendum.ahk antwortet

	; letzte Änderung: 22.10.2022

		global MsgRec

	; die ID des Addendum Fensters läßt sich nicht ermitteln
		If !(AddendumID := GetAddendumID())
			return 0

	; versucht Addendum zur einer Antwort zu überreden (wartet max 4s auf eine Antwort)
		MsgRec := ""
		noanswers := 0
		maxnoanswers := 40
		ObserveTime := 12
		ObserveTimeMax := 40
		QPC(true)
		starttime := A_TickCount

	; wenn Addendum 10x hintereinander nicht antwortet innerhalb von mindestens 4 Sekunden und max. 10 Sekunden Beobachtungsdauer
		while ((recTime:=QPC(false)) <= ObserveTimeMax ) {
			MsgRec := ""
			Send_WM_COPYDATA("Status||" Addendum.hMsgGui, AddendumID)		; Statusabfrage senden
			Loop 10 {                                                                                            	; 500 ms warten
				Loops := A_Index
				If (MsgRec = "okay")
					break                                                                                        	; Abbruch wenn MsgRec Okay enthält
				Sleep 50
			}
			Sleep % (10-Loops)*50                                                                      	; den Rest bis 500ms ausschlafen

			noanswers 	:=  	lastanswer = "okay" 	? 0 : noanswers                         	; Zähler wird zurück gesetzt, wenn zwischendurch eine Antwort eintraf
			noanswers 	+= 	MsgRec != "okay"  	? 1 : 0
			allnoanswers += 	MsgRec != "okay"  	? 1 : 0
			lastanswer 	:= 	MsgRec
			If (noanswers >= maxnoanswers)                                                                         	; Abbruch nach 10 ausbleibenden Antworten entspricht 1s
				break

		  ;~ ; wenn die minimale Oberservierungszeit abglaufen ist, aber aktuell keine Antworten von Addendum kamen, wird die Oberservierungszeit verlängert
			;~ If (rectTime >= ObserveTime && noanswers > 0)

		}

return {"status": noanswers>=maxnoanswers ? false : true, "noanswers": noanswers, "allnoanswers":allnoanswers, "oberserveTime": Format("{1:.1f}", recTime)}
}

GetAppImagePath(appname) {                                                             	;-- Installationspfad eines Programmes

	headers:= {	"DISPLAYNAME"                  	: 1
					,	"VERSION"                         	: 2
					, 	"PUBLISHER"             	         	: 3
					, 	"PRODUCTID"                    	: 4
					, 	"REGISTEREDOWNER"        	: 5
					, 	"REGISTEREDCOMPANY"    	: 6
					, 	"LANGUAGE"                     	: 7
					, 	"SUPPORTURL"                    	: 8
					, 	"SUPPORTTELEPHONE"       	: 9
					, 	"HELPLINK"                        	: 10
					, 	"INSTALLLOCATION"          	: 11
					, 	"INSTALLSOURCE"             	: 12
					, 	"INSTALLDATE"                  	: 13
					, 	"CONTACT"                        	: 14
					, 	"COMMENTS"                    	: 15
					, 	"IMAGE"                            	: 16
					, 	"UPDATEINFOURL"            	: 17}

   appImages := GetAppsInfo({mask: "IMAGE", offset: A_PtrSize*(headers["IMAGE"] - 1) })
   Loop, Parse, appImages, "`n"
	If Instr(A_loopField, appname)
		return A_loopField

return ""
}

GetAppsInfo(infoType) {                                                                     	;-- Informationen über ein installiertes Programm erhalten

	static CLSID_EnumInstalledApps := "{0B124F8F-91F0-11D1-B8B5-006008059382}"
        , IID_IEnumInstalledApps     	:= "{1BC752E1-9046-11D1-B8B3-006008059382}"

        , DISPLAYNAME            	:= 0x00000001
        , VERSION                    	:= 0x00000002
        , PUBLISHER                  	:= 0x00000004
        , PRODUCTID                	:= 0x00000008
        , REGISTEREDOWNER    	:= 0x00000010
        , REGISTEREDCOMPANY	:= 0x00000020
        , LANGUAGE                	:= 0x00000040
        , SUPPORTURL               	:= 0x00000080
        , SUPPORTTELEPHONE  	:= 0x00000100
        , HELPLINK                     	:= 0x00000200
        , INSTALLLOCATION     	:= 0x00000400
        , INSTALLSOURCE         	:= 0x00000800
        , INSTALLDATE              	:= 0x00001000
        , CONTACT                  	:= 0x00004000
        , COMMENTS               	:= 0x00008000
        , IMAGE                        	:= 0x00020000
        , READMEURL                	:= 0x00040000
        , UPDATEINFOURL        	:= 0x00080000

   pEIA := ComObjCreate(CLSID_EnumInstalledApps, IID_IEnumInstalledApps)

   while DllCall(NumGet(NumGet(pEIA+0) + A_PtrSize*3), Ptr, pEIA, PtrP, pINA) = 0  {
      VarSetCapacity(APPINFODATA, size := 4*2 + A_PtrSize*18, 0)
      NumPut(size, APPINFODATA)
      mask := infoType.mask
      NumPut(%mask%, APPINFODATA, 4)

      DllCall(NumGet(NumGet(pINA+0) + A_PtrSize*3), Ptr, pINA, Ptr, &APPINFODATA)
      ObjRelease(pINA)
      if !(pData := NumGet(APPINFODATA, 8 + infoType.offset))
         continue
      res .= StrGet(pData, "UTF-16") . "`n"
      DllCall("Ole32\CoTaskMemFree", Ptr, pData)  ; not sure, whether it's needed
   }
   Return res
}

GetProcessCpu(PID) {                                                                         	;-- CPU load (calculation with cpu cores)

	; https://www.autohotkey.com/boards/viewtopic.php?t=71922
	static time, coreNumber, prevProcTime, prevTickCount

	if !coreNumber {
		VarSetCapacity(time, 32)
		VarSetCapacity(SYSTEM_INFO, 24 + A_PtrSize*3)
		DllCall("GetSystemInfo", "Ptr", &SYSTEM_INFO)
		coreNumber := NumGet(SYSTEM_INFO, 8 + A_PtrSize*3, "UInt")
	}
	GetProcessTime(PID, time)
	kernelTime	:= NumGet(time, 16, "UInt64") / 10000
	userTime   	:= NumGet(time, 24, "UInt64") / 10000
	procTime  	:= kernelTime + userTime
	if prevProcTime
	  cpu := Round((procTime - prevProcTime)/coreNumber/(A_TickCount - prevTickCount) * 100, 2)+0
	prevProcTime	:= procTime
	prevTickCount := A_TickCount

return cpu
}

GetProcessTime(PID, ByRef time) {                                                        	;-- ProcessTime for one process?
   hProc := DllCall("OpenProcess", "UInt", PROCESS_QUERY_INFORMATION := 0x400, "UInt", false, "UInt", PID, "Ptr")
   DllCall("GetProcessTimes", "Ptr", hProc, "Ptr", &time, "Ptr", &time + 8, "Ptr", &time + 16, "Ptr", &time + 24)
   DllCall("CloseHandle", "Ptr", hProc)
}

GetProcessTimes(PID=0)    {                                                                	;-- ProcessTimes = average CPU Usage?
   Static oldKrnlTime, oldUserTime
   Static newKrnlTime, newUserTime
   oldKrnlTime := newKrnlTime, oldUserTime := newUserTime
   hProc := DllCall("OpenProcess", "Uint", 0x400, "Uint", 0, "Uint", pid, "Ptr")
   DllCall("GetProcessTimes", "Ptr", hProc, "int64*", CreationTime, "int64*", ExitTime, "int64*", newKrnlTime, "int64*", newUserTime)
   DllCall("CloseHandle", "Ptr", hProc)
Return (newKrnlTime-oldKrnlTime + newUserTime-oldUserTime)/10000000 * 100
}

GetSystemTimes() {                                                                           	;-- Total CPU Load

   Static oldIdleTime, oldKrnlTime, oldUserTime
   Static newIdleTime, newKrnlTime, newUserTime

   oldIdleTime := newIdleTime, oldKrnlTime := newKrnlTime, oldUserTime := newUserTime
   DllCall("GetSystemTimes", "int64*", newIdleTime, "int64*", newKrnlTime, "int64*", newUserTime)

Return (1 - (newIdleTime-oldIdleTime)/(newKrnlTime-oldKrnlTime + newUserTime-oldUserTime)) * 100
}

GetState(PID,TimeOut := 200) {                                                           	;-- Prozessstatus

	/*   * SendMessageTimeout values

		#define SMTO_NORMAL                           	0x0000
		#define SMTO_BLOCK                              	0x0001
		#define SMTO_ABORTIFHUNG                 	0x0002
		#if(WINVER >= 0x0500)
		#define SMTO_NOTIMEOUTIFNOTHUNG 	0x0008
		#endif /* WINVER >= 0x0500 */
		#endif /* !NONCMESSAGES */

		SendMessageTimeout(
			__in HWND hWnd,
			__in UINT Msg,
			__in WPARAM wParam,
			__in LPARAM lParam,
			__in UINT fuFlags,
			__in UINT uTimeout,
			__out_opt PDWORD_PTR lpdwResult);

	 */

	; WM_NULL =0x0000
	; SMTO_ABORTIFHUNG =0x0002
	; TimeOut: milliseconds to wait before deciding it is not responding - 100 ms seems reliable under 100% usage
		NR_temp 	:= 0     	; init
		WinGet, wid, ID, % "ahk_pid " PID
		Responding := DllCall("SendMessageTimeout", "UInt", wid, "UInt", 0x0000, "Int", 0, "Int", 0, "UInt", 0x0002, "UInt", TimeOut, "UInt *", NR_temp)

return wid = "" ? 99 : Responding ? 1 : 0 ; 99 Background, 1=responding, 0= Not Responding
}

GetTimestrings(ms, maxTime:="Auto") {                                               	;-- Stunden, Minuten, Sekunden aus Millisekunden berechnen

	; maxTime = "Auto"

	If (maxTime >= 2 || maxTime = "Auto")  {
		hour	:= ms // 360000
		ms   	-= hour * 360000
	}
	If (maxTime >= 1 || maxTime = "Auto") {
		min	:= ms // 60000
		ms   	-= min * 60000
	}
	sec  	:= ms // 1000
	ms   	-= sec * 1000

return {"hour":SubStr("0" hour, -1), "min":SubStr("0" min, -1), "sec":SubStr("0" sec, -1), "msec":ms}
}

GuiControlGet(guiname, cmd, vcontrol) {                                             	;-- GuiControlGet wrapper
	GuiControlGet, cp, % guiname ": " cmd, % vcontrol
	If (cmd = "Pos")
		return {"X":cpX, "Y":cpY, "W":cpW, "H":cpH}
return cp
}

MessageWorker(InComing) {                                                             	;-- handles received messages

	global MsgRec

	recv := {	"txtmsg"		: (StrSplit(InComing, "|").1)
				, 	"opt"     	: (StrSplit(InComing, "|").2)
				, 	"fromID"	: (StrSplit(InComing, "|").3)}

	If (recv.txtmsg = "Status")
		MsgRec := recv.opt

}

IniReadExt(SectionOrFullFilePath, Key="", DefaultVal="", convert=true) { 	;-- eigene IniRead funktion für Addendum

	/* Beschreibung

		-	beim ersten Aufruf der Funktion nur den Pfad und Namen zur ini Datei übergeben!
				workini := IniReadExt("C:\temp\Addendum.ini")

		-	die Funktion behandelt einen geschriebenen Wert welcher ein "ja" oder "nein" ist, als logische Wahrheitswerte (wahr und falsch).
		-	UTF-16-Zeichen (verwendet in .ini Dateien) werden per Default in UTF-8 Zeichen umgewandelt

		letzte Änderung: 14.02.2022

	 */

		static admDir
		static WorkIni

	; Arbeitsini Datei wird erkannt wenn der übergebene Parameter einen Pfad ist
		If RegExMatch(SectionOrFullFilePath, "^([A-Z]:\\|\\\\[A-Z]{2,}).*")	{
			If !FileExist(SectionOrFullFilePath)	{
				MsgBox, 0x1024, % "Addendum für AlbisOnWindows"
						            	 , % "Die .ini Datei existiert nicht!`n`n" WorkIni "`n`nDas Skript wird jetzt beendet.", 10
				ExitApp
			}
			WorkIni	:= SectionOrFullFilePath
			admDir	:= RegExMatch(WorkIni, "^([A-Z]:\\|\\\\[A-Z]{2,}).*?AlbisOnWindows", rxDir) ? rxDir : Addendum.Dir
			return WorkIni
		}

	; Workini ist nicht definiert worden, dann muss das komplette Skript abgebrochen werden
		If !WorkIni {
			MsgBox,, Addendum für AlbisOnWindows, %	"Bei Aufruf von IniReadExt muss als erstes`n"
																			. 	"der Pfad zur ini Datei übergeben werden.`n"
																			.	"Das Skript wird jetzt beendet.", 10
			ExitApp
		}

	; Section, Key einlesen, ini Encoding in UTF.8 umwandeln
		IniRead, OutPutVar, % WorkIni, % SectionOrFullFilePath, % Key
		If convert
			OutPutVar := StrUtf8BytesToText(OutPutVar)

	; Bearbeiten des Wertes vor Rückgabe
		If (InStr(OutPutVar, "ERROR") || StrLen(OutPutVar) = 0)
			If DefaultVal  { ; Defaultwert vorhanden, dann diesen Schreiben und Zurückgeben
				OutPutVar := DefaultVal
				IniWrite, % DefaultVal, % WorkIni, % SectionOrFullFilePath, % key
				If ErrorLevel
					TrayTip, % A_ScriptName, % "Der Defaultwert <" DefaultVal "> konnte geschrieben werden.`n`n[" WorkIni "]", 2
			}
			else return "ERROR"

		If InStr(OutPutVar, "%AddendumDir%")
			OutPutVar :=  StrReplace(OutPutVar, "%AddendumDir%", admDir)
		else if RegExMatch(OutPutVar, "%\.exe$") && !RegExMatch(OutPutVar, "i)[A-Z]\:\\")
			return GetAppImagePath(OutPutVar)
		else if RegExMatch(OutPutVar, "i)^\s*(ja|nein|An|Aus)\s*$", bool)
			return (bool1= "ja" || bool1 = "An") ? true : false

return Trim(OutPutVar)
}

StrUtf8BytesToText(vUtf8) {                                                                    	;-- Umwandeln von Text aus .ini Dateien nach UTF-8
	if A_IsUnicode 	{
		VarSetCapacity(vUtf8X, StrPut(vUtf8, "CP0"))
		StrPut(vUtf8, &vUtf8X, "CP0")
		return StrGet(&vUtf8X, "UTF-8")
	} else
		return StrGet(&vUtf8, "UTF-8")
}

OverwatchLoad(iniPath, iniSection) {                                                     	;--

	IniRead, tmp, % iniPath, % iniSection          ; loads whole section
	tmp := StrUtf8BytesToText(tmp)
	If (!InStr(tmp, "ERROR") && StrLen(tmp)>0)
		For each, keyval in StrSplit(tmp, "`n") {
			RegExMatch(keyval, "^(?<Key>.*?)\d*=(?<Val>.*)$", Ini)
			If (IniKey = "overwatch")
				overwatch.Push({	"name"	: StrSplit(IniVal, "|").1
										, 	"path"	: StrSplit(IniVal, "|").2
										, 	"opts"	: StrSplit(IniVal, "|").3
										, 	"runon"	: StrSplit(IniVal, "|").4
										, 	"delay"	: StrSplit(IniVal, "|").5
										, 	"state"	: (StrSplit(IniVal, "|").6 = "An" ? 1 : 0) })
		}

return OverwatchWQL()
}

OverwatchWQL() {                                                                              	;--

	For each, ow in overwatch {
		runon := StrReplace(ow.runon, ",", "|")
		If RegExMatch(ow.Name, "\.exe$") && (RegExMatch(Addendum.compname, "(" runon ")") || runon = "*")
			WQL .= " or TargetInstance.Name = '" ow.Name "'"
	}

return WQL
}

Process(Subcommand, PIDOrName:="", Value:="") {                           	;-- wrapper
	Process, % SubCommand, % PIDOrName, % Value
return ErrorLevel
}

ProcessClose(PID) {                                                                            	;-- force close a process, returns last PID
	Process, Close, % PID
return ErrorLevel
}

ProcessDelete_OnObjectReady(obj) {                                                 	;-- is called when a process terminates

	; this is prepared for multi restart purposes, but I did'nt finished it yet
	; this can only restart one process per interval

	static RestartProcess
	static rProcStack := Array()

    Process := obj.TargetInstance

	;verhindert die doppelte Ausführung
		If (StrLen(RestartProcess) > 0)
			return

		If RegExMatch(Process.Name, "i)AutoHotkey")
			RegExMatch(Process.CommandLine, "(?<Name>[\w\-]+)\.(?<Ext>\w+)(?=\" q "\s*$)", proc)
		else
			RegExMatch(Process.Name, "(?<Name>[\w\-]+)\.(?<Ext>\w+)\s*$", proc)
		RestartProcess := procName "." procExt

	; prüft hier nur um Addendum.ahk neu zu starten
		If InStr(RestartProcess, "Addendum.ahk") 	{
			If WinExist("Addendum.ahk ahk_class #32770") 	{
				; falls eine Autohotkey Fehlermeldung angezeigt wird, wird Addendum.ahk nicht sofort neu gestartet
					WinGetText, wText, % "Addendum.ahk ahk_class #32770"
					If RegExMatch(wText, "Error:\s.*Line#.*--->") {
						RegExMatch(Process.CommandLine, "[A-Z]:[\w\\_\säöüÄÖÜß.\-]+\.ahk", SkriptPath)
						fehler :={"Message":wText, "File":SkriptPath}
						FehlerProtokoll(fehler, 1)
					}
			}
			If !AHKProcessExist("Addendum.ahk")
				If ShowRestartProgress("Addendum.ahk", 5)
					RestartAddendum()
			DoRestart := true, RestartProcess :=  ""
			return
		}

	; weitere zu überwachende Prozesse / Skripte
		else {
			If overwatch.Count()
				For each, ow in overwatch
					If InStr(RestartProcess, ow.name)
						If ((procExt = "exe") && !Process("Exist", RestartProcess)) || ((procExt = "ahk") && !AHKProcessExist(RestartProcess)) {
							TrayTip, AddendumMonitor, % "Neustart von " RestartProcess "`n in " ow.delay " Sekunden"
							Sleep ow.delay * 1000
							If (procExt = "exe")
								Run, % q . ow.path "\" RestartProcess q " " ow.opts
							else if (procExt = "ahk")
								Run, % Addendum.AHKH_exe " /f " q . ow.path "\" RestartProcess q
						}
		}

		DoRestart := true, RestartProcess :=  ""

}

ProcessDelete_OnHighCPULoad(ProcName, PID, WTime                      	;-- monitors a process over a given time to detect process hung not detected by windows
	, MinIdleTime, AvCPU=10, WSleep=100
	, DropAmount=50, DropTime=600) {

	/*   ProcessDelete_OnHighCPULoad()

		- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Description
		- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		The function is to estimate the probability of an Autohotkey script that is in an endless loop. If all conditions are met,
		the function closes the loop process. It is also possible to close processes that have hung for other reasons.
		But only in this case that MS Windows noticed a hang.

		- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Parameters
		- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		ProcName		    		-	only the name to view in Tooltip
		PID				    		-	Process ID of the process to be observed
		WTime			    		-	how long the process is observed (in seconds) before the overall evaluation takes place
		IdleTime	    			-	if 0 process will be closed if in idle state or not,
											higher than it must be a minimum idle in seconds to close process
		WSleep			    		-	get informations every x ms
		AvCPU			    		-	the average CPU usage above which the process will be closed
		Drop[Amount/Time]	-	The amount in percent and length in ms of the actual CPU load drop compared to the average cpu load
											that leads to the monitoring being aborted (the probability of an infinite loop is low)
		- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Others
		- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		InfoAfterLoop		     	an information window is displayed after a maximum of one tenth of the total monitoring loops
		ObservationIsRunning	to prevent another call if observation is still running

	*/

		global ObservationRuns

		ObservationRuns    	:= true
		InactiveObservation	:= false
		InactiveWait          	:= false
		ObservationLoops  	:= Floor(WTime*1000/WSleep)              	; WTime[s] * 1000 = WTime[ms]/WSleep[ms]
		InfoAfterLoop    		:= Round(ObservationLoops/10)
		CountLoadDrops   	:= false
		DropObservation		:= 0
		DropAmount          	:= Round(DropAmount/100, 1)					; to be a percent value

	; observation of process
		Loop % ObservationLoops {

			; get process informations
				ProcessTime	:= 	GetProcessTimes(PID)
				SystemTime	:= 	GetSystemTimes()
				WTimeSec  	:= 	Round(A_Index*WSleep/1000)
				IdleTime    	:= 	Round(A_TimeIdlePhysical/1000)
				PTSum         +=	ProcessTime
				AvPT				:= 	PTSum/A_Index

			; detect CPU load drops
				If (!CountLoadDrops && (AvPT*DropAmount >= ProcessTime))
					CountLoadDrops := true, DTSum := DropsCount := 0, DropCountStartTime := A_TickCount,	DropObservation ++

			; count CPU load drops
				If CountLoadDrops {

					DropCount 	++
					DTSum			+=	ProcessTime
					AvDT			:=	DTSum/DropCount

				; time for the load drop monitoring has expired
					If (A_TickCount-DropCountStartTime >= DropTime) {
						; a significant CPU load drop was found over the period. The process seems to be alive.
							If (AvDT <= AvPT*DropAmount) {
								BTT()
								return 0
							}
						; the CPU load drop was too short to be evaluated
							else
								CountLoadDrops := false, DTSum := DropCount := 0
					}

				}

			; checks 10 times in the observation period whether the process is inactive or waits here
				If (!CountLoadDrops && Mod(A_Index, ObservationLoops/10) = 0) || InactiveObservation {
					If !InactiveWait
						SetTimer, CountInactive, -10
				} else
					Sleep % WSleep

			; Information is only displayed if the client is physically inactive for 30 seconds
				If (A_Index >= 10)  {																										;&& (IdleTime >= MinIdleTime) (A_Index >= InfoAfterLoop)
					msg :=	"Name:                    " 	ProcName 	     		                                          	"`n"
						.    	"PID:                        " 	PID     			                                                    	"`n"
						. 		"IdleTime:                " 	IdleTime "s"	                                                       	"`n"
						.		"Round:                   " 	SubStr("0000" A_Index, -3) "/" ObservationLoops 	"`n"
						.		"Process state:          " 	ProcessState	                                                    	"`n"
						.		"Time:                      "	WTimeSec "s"                                                     	"`n"
						.		"CPU total:              " 	SubStr("0" Round(SystemTime), -1)                        	"`n"
						. 		"CPU usage:            " 	SubStr("0" Round(ProcessTime), -1)	                 		"`n"
						.		"av. CPU usage:       " 	Round(AvPT, 1)                                                  	"`n"
						.		"drop observations: "  	DropObservation                                                	"`n"
						.		"Inactive times:        " 	InactiveObservation
					BTT(msg, A_ScreenWidth-400, A_ScreenHeight-400,, "Style4")
				}

			; process is 5 times in advance not responding this loop breaks to close this process
				If (InactiveObservation > 4)
					break

		}

	; If a condition is met, the observed process is now terminated
		BTT()

		If (InactiveObservation >= 5) || (AvPT >= AvCPU && IdleTime >= MinIdleTime) {
			Process, Close, % PID
			ObservationRuns := false
			return msg
		}

		ObservationRuns := false

return false

; count in advance how many time the process is not responding
CountInactive: ;{

	InactiveWait := true
	ProcessState := GetState(PID, 200)
	InactiveObservation := !ProcessState ? InactiveObservation+1 : 0
	InactiveWait := false

return ;}
}

RefreshTray() {                                                                                    	;-- by SWIN - http://www.autohotkey.com/community/viewtopic.php?t=8086
   WM_MOUSEMOVE := 0x200
   ControlGetPos, xTray,, wTray,, ToolbarWindow321, ahk_class Shell_TrayWnd
   endX := xTray + wTray
   x := 5
   y := 12
   Loop {
      if (x > endX)
         break
      point := (y << 16) + x
      PostMessage, %WM_MOUSEMOVE%, 0, %point%, ToolbarWindow321, ahk_class Shell_TrayWnd
      x += 18
   }
}

QPC(R := 0) {                                                                                      	;-- genaueste Zeitmessung, Rückgabe in Sekunden (floating point)
     static P := 0, F := 0, Q := DllCall("QueryPerformanceFrequency", "Int64P", F)
    return ! DllCall("QueryPerformanceCounter", "Int64P", Q) + (R ? (P := Q) / F : (Q - P) / F)
}

Receive_WM_COPYDATA(wParam, lParam) {                                     	;-- receives messages from other processes

    StringAddress	:= NumGet(lParam + 2*A_PtrSize)
	fn_MsgWorker	:= Func("MessageWorker").Bind(StrGet(StringAddress))
	SetTimer, % fn_MsgWorker, -10

return
}

Send_WM_COPYDATA(ByRef StringToSend, ScriptID) {                         	;-- send messages to other processes

    static TimeOutTime            	:= 4000

    Prev_DetectHiddenWindows	:= A_DetectHiddenWindows
    Prev_TitleMatchMode         	:= A_TitleMatchMode
    DetectHiddenWindows On
    SetTitleMatchMode 2

    VarSetCapacity(CopyDataStruct, 3*A_PtrSize, 0)
    SizeInBytes := (StrLen(StringToSend) + 1) * (A_IsUnicode ? 2 : 1)
    NumPut(SizeInBytes, CopyDataStruct, A_PtrSize)
    NumPut(&StringToSend, CopyDataStruct, 2*A_PtrSize)

    SendMessage, 0x4a, 0, &CopyDataStruct,, % "ahk_id " ScriptID,,,, % TimeOutTime ; 0x4a is WM_COPYDATA.
	eL := ErrorLevel

    DetectHiddenWindows, % Prev_DetectHiddenWindows 	; Restore original setting for the caller.
    SetTitleMatchMode	 ,  % Prev_TitleMatchMode         	; Same.

return eL  ; Return SendMessage's reply back to our caller.
}

ShowRestartProgress(scriptname, duration) {                                       	;-- zeigt die Zeit bis Addendum erneut gestartet wird

	; duration = Dauer in Sekunden
	global DoRestart

	DoRestart 	:= true
	slTime   	:= 50
	slFactor 	:= Floor(100 / slTime)
	Loops   	:= Floor((duration * 1000) / 100)

	Progress,	% "B2 cW202842 cBFFFFFF cTFFFFFF FM9 FS8 zY2 zH10 w250 WM700 WS200"
				,	% "...Neustart in 10s"
				, 	% "Neustart: " scriptname " Abbruch mit Escape"
				, 	% "----------"
				, 	Futura Bk Bt
	prghwnd := WinExist("ahk_class AutoHotkey2", "...Neustart")
	WinSet  	, Transparent, 200, % "ahk_id " prghwnd
	WinGetPos,,, ww, wh	, % "ahk_id " prghwnd
	WinGetPos,,,,  th    	, % "ahk_class Shell_TrayWnd"
	WinMove	, % "ahk_id " prghwnd,, % A_ScreenWidth - ww - 5, % A_ScreenHeight - th - wh - 5

	Loop % Loops {

		If !DoRestart
			break

		Progress % (rest := Round(100 - (A_Index * slFactor)))
		ControlSetText, Static2, % "...Neustart in " (Round(rest/duration, 1)) "s", % "ahk_id " prghwnd
		Sleep % slTime

		If !DoRestart
			break

	}
	Progress, Off

return DoRestart  	; Abbruch per Hotkey-Methode gibt ein false zurück
}

TimeStamp() {                                                                                      	;-- Formattime wrapper
	FormatTime, time, % A_Now, dd.MM.yyyy HH:mm:ss
return time
}

TelegramBots() {                                                                                	;-- Telegram Bot Einstellungen

	Addendum.Telegram.Bots  	:= Object()
	Addendum.Telegram.Chats 	:= Object()

	Loop {
		BotName	:= IniReadExt("Telegram", "Bot" A_Index)
		If (!BotName || InStr(BotName, "ERROR"))
			break
		else
			Addendum.Telegram.Bots.Push(ReadBots(BotName))
	}

  ; Chatnamen und ID's
	ReadChats()

}

ReadBots(BotName) {                                                                         	;-- Telegram-Bot Daten

	; Daten der Bots sind in der .ini zu verwalten

	Bot := {	"BotName"        	: BotName
				, 	"Token"             	: IniReadExt("Telegram", BotName "_Token")
				, 	"ID"                    	: IniReadExt("Telegram", BotName "_ID", 0)
				, 	"Active"              	: IniReadExt("Telegram", BotName "_Active", "nein")
				, 	"BotLastMsg"      	: IniReadExt("Telegram", BotName "_LastMsg", "--")
				, 	"BotLastMsgTime"	: IniReadExt("Telegram", BotName "_LastMsgTime", "000000")}

return Bot
}

ReadChats() {

   ; Telegram Gruppenchat Namen und ID's auslesen
	chats := Object()
	groupbreak := false
	Loop {

		idx := A_Index
		chat := {"Name"	: IniReadExt("Telegram", "Chat" A_Index "_Name")
		  		, 	 "ID"   		: IniReadExt("Telegram", "Chat" A_Index "_ID")}
		For key, val in chat
			If (!val || InStr(val, "ERROR")) {
				groupbreak := true
				break
			}
		If groupbreak
			break

		Addendum.Telegram.Chats[chat.Name] := chat.ID
		;~ SciTEOutput("Loops: " idx " " chat.Name)

	}

}

AddendumMonitor_ico(NewHandle := False) {
Static hBitmap := AddendumMonitor_ico()
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "AAABAAEAMDAAAAEAGACoHAAAFgAAACgAAAAwAAAAYAAAAAEAGAAAAAAAAAAAAGQAAABkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB0iWNYjS5IkA9BkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBGkAtTjiRtileFhoUAAAAAAAAAAAAAAAAAAACDh39WjStAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBQjiB/h3kAAAAAAAAAAACCh35MjxhAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBKjxKBhnsAAAAAAABVjihAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBUjiYAAABxiV5AkQBAkQBtqzyZxHeZxHeZxHeZxHd5skxAkQBAkQBAkQBAkQBAkQBAkQBAkQBElAaZxHeZxHeZxHeZxHeZxHdJlgxAkQBAkQBAkQBipS6ZxHeZxHeZxHeZxHeJu2JAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCTwW+ZxHeZxHeZxHeZxHdZoCJAkQByiWBVjihAkQBAkQBwrUD////////////////l8NxAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDX6Mr///////////////99tFFAkQBAkQBAkQC01Jv///////////////+rz49AkQBAkQBAkQBAkQBAkQBAkQBAkQBboST////////////////5/PdLlw9AkQBXjS1FjwxAkQBAkQBBkgLq8+P///////////////9foypAkQBAkQBAkQBAkQBAkQBAkQBAkQCVwnH///////////////+41qBAkQBAkQBBkgLt9ef///////////////9oqDVAkQBAkQBAkQBAkQBAkQBAkQBAkQCXw3T////////////////A26tAkQBAkQBJkBFAkQFAkQBAkQBAkQCqzo3///////////////+bxXlAkQBAkQBAkQBAkQBAkQBAkQBAkQBTnBn9/vz////////////x9+xDkwRAkQBpqTf////////////////l8NxBkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQDR5MH///////////////99tFFAkQBAkQBCkQVAkQBAkQBAkQBAkQBnpzT////////////////W58hAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDO477///////////////9urD5AkQCjyoT///////////////+gyIBAkQBAkQBAkQBAkQBAkQBAkQBAkQBNmRL8/fv////////////1+fFGlAhAkQBAkQBCkQVAkQBAkQBAkQBAkQBAkQDj79n////////////+//5UnRtAkQBAkQBAkQBAkQBAkQBAkQBAkQCMvWX///////////////+pzYxAkQDd69H///////////////9foypAkQBAkQBAkQBAkQBAkQBAkQBAkQCFuVz///////////////+11JxAkQBAkQBAkQBCkQVAkQBAkQBAkQBAkQBAkQChyYH///////////////+MvWZAkQBAkQBAkQBAkQBAkQBAkQBAkQBNmBH7/Pn////////////l8NxXnh/////////////////Z6cxAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQC/2qn///////////////9xrUJAkQBAkQBAkQBCkQRAkQBAkQBAkQBAkQBAkQBdoif////////////////I37VAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDF3bH///////////////+y05j///////////////+Ww3NAkQBAkQBAkQBAkQBAkQBAkQBAkQBElAb0+fD////////////s9OVCkgNAkQBAkQBAkQBCkQRAkQBAkQBAkQBAkQBAkQBAkQDa6c3////////////6/PhJlgxAkQBAkQBAkQBAkQBAkQBAkQBAkQCCt1j////////////////9/v3////////////9/v1UnRtAkQBAkQBAkQBAkQBAkQBAkQBAkQBzrkT///////////////+qzo5AkQBAkQBAkQBAkQBCkQRAkQBAkQBAkQBAkQBAkQBAkQCXw3T///////////////9/tVRAkQBAkQBAkQBAkQBAkQBAkQBAkQBIlgv2+vP////////////////////////////O471AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCt0JH///////////////9npzRAkQBAkQBAkQBAkQBCkQRAkQBAkQBAkQBAkQBAkQBAkQBWnh39/v3///////////+516JAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQC82KX///////////////////////////+LvGRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDm8d7////////////j79lAkQBAkQBAkQBAkQBAkQBCkQRAkQBAkQBAkQBAkQBAkQBAkQBAkQDR5MH////////////y9+1DkwRAkQBAkQBAkQBAkQBAkQBAkQBAkQB5skz////////////////////////6/PhNmBFAkQBAkQBAkQBAkQBAkQBAkQBAkQBhpCz///////////////+gyIBAkQBAkQBAkQBAkQBAkQBCkQRAkQBAkQBAkQBAkQBAkQBAkQBAkQCPvmn///////////////9wrUBAkQBAkQBAkQBAkQBAkQBAkQBAkQBEkwXx9+z////////////////////C3K5AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCbxXn///////////////9coSZAkQBAkQBAkQBAkQBAkQBCkQRAkQBAkQBAkQBAkQBAkQBAkQBAkQBPmhT7/fr///////////+rz49AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCz05r///////////////////+AtlVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDU5sb////////////X6MpAkQBAkQBAkQBAkQBAkQBAkQBCkQRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDI4Lb////////////n8d9AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBxrUH////////////////1+fJHlQlAkQBAkQBAkQBAkQBAkQBAkQBAkQBQmhb9/vz///////////+UwXBAkQBAkQBAkQBAkQBAkQBAkQBBkQNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCFuVz///////////////9ipS5AkQBAkQBAkQBAkQBAkQBAkQBAkQBBkgL1+fL////////////I4LZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCJu2H////////////9/vxTnBpAkQBAkQBAkQBAkQBAkQBAkQBBkQNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBJlgz4+/X///////////+ex31AkQBAkQBAkQBAkQBAkQBAkQBAkQBcoSX////////////////y9+1EkwVAkQBAkQBAkQBAkQBAkQBAkQBAkQDD3K/////////////M4rtAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQC/2qn////////////b6s9AkQBAkQBAkQBAkQBAkQBAkQBAkQCbxnr///////////////////9zrkRAkQBAkQBAkQBAkQBAkQBAkQBJlgz4+/b///////////+Ju2FAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB8tFD///////////////9aoCNAkQBAkQBAkQBAkQBAkQBAkQDb6s////////////////////+x0pdAkQBAkQBAkQBAkQBAkQBAkQCDt1n////////////5/PdMmBBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBElAby9+3///////////+YxHZAkQBAkQBAkQBAkQBAkQBaoCP////////////////////////u9ehCkgNAkQBAkQBAkQBAkQBAkQDC3K7////////////B26xAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQC21Z3////////////X6MlAkQBAkQBAkQBAkQBAkQCcxnv///////////////////////////9urD5AkQBAkQBAkQBAkQBKlw34+/b///////////9+tVNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQByrkP////////////+//5Wnh5AkQBAkQBAkQBAkQDc69D///////////////////////////+t0JJAkQBAkQBAkQBAkQCDuFr////////////0+fBGlAhAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkgLs9OX///////////+UwXBAkQBAkQBAkQBcoSb///////////////+31p/////////////r8+RBkQFAkQBAkQBAkQDE3bD///////////+21Z1AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCt0JH////////////S5cNAkQBAkQBAkQCdxnz////////////a6s5Wnh39/v3///////////9qqThAkQBAkQBKlw34+/b///////////9zrkRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBqqTj////////////9/v1TnBpAkQBAkQDd69H///////////+Xw3RAkQDR5cL///////////+ozYtAkQBAkQCFuVz////////////t9edCkgNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDk79v///////////+Pv2pAkQBdoif////////////9/v1VnRxAkQCPvmn////////////n8d9BkQFAkQDF3bH///////////+rz49AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCkyoX////////////O471AkQCex37////////////P479AkQBAkQBPmhT7/fr///////////9mpzNKlw35/Pf///////////9oqDVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBfoyn////////////9/vxQmhXe7NP///////////+MvWVAkQBAkQBAkQDI4Lb///////////+ky4aGuV3////////////l8NxBkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDd7NL///////////+ozYv////////////7/PlOmRNAkQBAkQBAkQCGuV3////////////j79rF3rL///////////+gyIBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCaxXj////////////9/v3////////////F3bFAkQBAkQBAkQBAkQBJlgz4+/X////////////9/vz///////////9doidAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBXnh/+//7///////////////////////+Bt1dAkQBAkQBAkQBAkQBAkQC/2qr////////////////////////Y6ctAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDU5sb////////////////////1+fJHlQpAkQBAkQBAkQBAkQBAkQB9tFH///////////////////////+VwnJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCSwG3///////////////////+51qFAkQBAkQBAkQBAkQBAkQBAkQBElAby9+3////////////////9/v1UnRtAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQFAkQBAkQBAkQBA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"
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

;}

;{ Includes
#Include %A_ScriptDir%\..\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\..\lib\class_Btt.ahk
#Include %A_ScriptDir%\..\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\..\lib\SciTEOutput.ahk
;}


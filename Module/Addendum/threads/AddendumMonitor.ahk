;-----------------------------------------------------------------------------------------------------------------------------------
;------------------------------------------------ ADDENDUM MONITOR ----------------------------------------------------
;-----------------------------------------------------------------------------------------------------------------------------------
													Version:= "1.21" , vom:= "18.10.2021"
;------------------------------------------------------ Runtime-Skript ----------------------------------------------------------
;------------------- startet Addendum bei einem Absturz oder (un-))absichtlichen Schliessen neu --------------------
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

;{ Variablen

		global Restartscript
		global ObservationRuns	:= false
		global DoRestart          	:= true
		global q                       	:= Chr(0x22)
		global winmgmts       	:= ComObjGet("winmgmts:") 	; Get WMI service object.
		global Addendum      	:= Object()

	; Addendum object ;{
		RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)
		Addendum.compname    	:= StrReplace(A_ComputerName, "-")
		Addendum.ScriptName    	:= "Addendum.ahk"
		Addendum.Dir               	:= AddendumDir
		Addendum.Ini                	:= AddendumDir 	"\Addendum.ini"
		Addendum.ScriptPath      	:= AddendumDir 	"\Module\Addendum\" Addendum.ScriptName

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
		Menu, Tray, Add
		Menu, Tray, Add, % "Reload"		                            	, % "admMonReload"
		Menu, Tray, Add, % "Exit"			                             	, % "admMonExitApp"
	;}

	; get data from Addendum.ini
		IniReadExt(Addendum.Ini)

	; Addendum DBPath und Logfile Pfad ;{
		Addendum.DBPath := IniReadExt("Addendum", "AddendumDBPath")
		If InStr(Addendum.DBPath, "Error") || !RegExMatch(Addendum.DBPath, "i)[a-z]\:\\") || !InStr(FileExist(Addendum.DBPath), "D")
			Addendum.DBPath := ""
		else
			Addendum.MonLogPath := Addendum.DBPath 	"\sonstiges\AddendumMonitorLog.txt"

		If Addendum.DBPath && !InStr(FileExist(Addendum.DBPath "\sonstiges"), "D") {
			FileCreateDir, % Addendum.DBPath "\sonstiges"
			If ErrorLevel
				Addendum.MonLogPath := ""
		}
	;}

	; Interskript communication gui ;{
		;If !WinExist("Addendum Monitor Gui ahk_class AutoHotkeyGui") {

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
		BotName	:= IniReadExt("Telegram", "Bot1")
		If !InStr(BotName, "Error") {
				Addendum.Telegram := Object()
				BotToken	:= IniReadExt("Telegram", BotName "_Token")
				BotChatID	:= IniReadExt("Telegram", BotName "_ChatID")
				Addendum.Telegram.Push({"BotName":BotName, "Token": BotToken, "ChatID": BotChatID})
		}
	;}

	; loads restart properties (script detects high cpu usage of Addendum.ahk and will close-restart Addendum.ahk) ;{
		Addendum.MaxAverageCPU 	:= IniReadExt("Addendum", "HighCPU_MaxAvarageCPU"	, 5)
		Addendum.RestartAfter       	:= IniReadExt("Addendum", "HighCPU_RestartAfter"         	, 360) 	; seconds
		Addendum.MinIdleTime     	:= IniReadExt("Addendum", "HighCPU_MinIdleTime"       	, 180)	; seconds
		Addendum.TimedCheck     	:= IniReadExt("Addendum", "HighCPU_Timer"                    	, 180)	; seconds
		TrTip1 := 	"CPU Last Grenze:   "    	Addendum.MaxAverageCPU 	"%`n"
				  .		"Neustart nach:      " 		Addendum.RestartAfter        	"s`n"
				  .		"Überprüfung alle: "   	Addendum.TimedCheck      	"s"
		TrTip2 := 	"CPU Überlastung:  nicht geprüft`n"
				 . 	  	"nächster Check in: "  	Addendum.TimedCheck      	"s"
		Menu, Tray, Tip, % TrTip1 "`n" TrTip2
		Addendum.TimerCall := A_TickCount
	;}

;}

;{ Sink Objekte erstellen Benachrichtigung nach 5 Sekunden\ Kontrolle das Addendum läuft alle 15 Minuten

	; Create sink objects for receiving event noficiations.
		ComObjConnect(CreateSink	:= ComObjCreate("WbemScripting.SWbemSink"), "ProcessCreate_")
		ComObjConnect(DeleteSink	:= ComObjCreate("WbemScripting.SWbemSink"), "ProcessDelete_")

	; Register for process deletion notifications:
		winmgmts.ExecNotificationQueryAsync(DeleteSink
			, "SELECT * FROM __InstanceDeletionEvent"
			. " WITHIN " 5                                                                  	; Set event polling interval, in seconds.
			. " WHERE TargetInstance ISA 'Win32_Process'"
			. " AND TargetInstance.Name LIKE 'Autohotkey%'")

	; Timerfunktion für 15 minütige Skriptausführungskontrolle (Skript hängt oder ist abgestürzt)
		fnADM      	:= Func("RestartAddendum").Bind(true)
		fnADMTime 	:= Addendum.TimedCheck*1000
		SetTimer, % fnADM, % fnADMTime
		Addendum.TimerCall := A_TickCount

	; erster Check von Addendum.ahk
		RestartAddendum()
;}

return

;{ Hotkeys
^+!ö::ExitApp
^!ö::    	;{ Skriptneustart
	TrayTip, AddendumMonitor, Skript wird neu gestartet, 1
	gosub ReleaseObjects
	Sleep 3000
	Reload
return ;}
#IfWinExist ahk_class AutoHotkey2, Neustart
~Esc::  	;{ Tastenkombination um Skriptneustart während der Anzeige des Countdownfenster abzubrechen
	If DoRestart {
		DoRestart := false
		FileAppend, % "Abbruch des Addendumneustart durch Nutzer, Zeit: " TimeStamp() ", Client: " A_ComputerName "`n", % Addendum.Dir "\logs'n'data\OnExit-Protokoll.txt"
	}
return
#IfWinExist

ReleaseObjects: ;{

	ObjRelease(winmgmts)
	ObjRelease(createSink)
	ObjRelease(DeleteSink)
	winmgmts := deleteSink := ""

return ;}

;}
;}

;{ TrayLabels

AHK_NOTIFYICON(wParam, lParam) {
	if (lParam = 0x200) { ; WM_LBUTTONUP
		SetTimer, admMonTimerInfo, -1
	return 0
    }
}

admMonReload:    	;{
gosub ReleaseObjects
Reload  ;}

admMonExitApp:    	;{
gosub ReleaseObjects
ExitApp ;}

admMonTimerInfo: 	;{
	nextcall 		:= Addendum.TimedCheck-((A_TickCount - Addendum.TimerCall)/1000)
	nextcall_min	:= Floor(nextcall/60)
	nextcall_sec	:= Floor(nextcall - (nextcall_min*60))
	TrTip2   		:= "CPU Überlastung:  " 	(ObservationRuns ? "ja":"nein") "`n"
						. 	 "nächster Check in: " 	nextcall_min "m " SubStr("00" nextcall_sec, -1) "s"
	;TrayTip, Addendum Monitor, % TTipMsg, 1, 16
	Menu, Tray, Tip, % TrTip1 "`n" TrTip2
return ;}

admMonInfo:         	;{
return ;}

admMonCheck:    	;{
RestartAddendum("override")
return ;}

admRestartNow:    	;{

	TrayTip, AddendumMonitor, Addendum.ahk wird gleich beendet und neu gestartet!, 2
	Sleep 2000
	If (Addendum.PID := AHKProcessExist(Addendum.ScriptName)) {

		If !ProcessClose(Addendum.PID) {
			TrayTip, AddendumMonitor, Addendum.ahk konnte nicht beendet werden!, 2
			return
		}

		if (PID := RestartAddendum("override"))
			TrayTip, AddendumMonitor, % "Addendum.ahk ProcessID= " PID		, 2
		else
			TrayTip, AddendumMonitor, % "Addendum.ahk konnte nicht gestartet werden!"	, 2

	}

return ;}
;}

;{ Funktionen
ProcessDelete_OnObjectReady(obj) {                                             	;-- Called when a process terminates

	; this is prepared for multi restart purposes, but I did'nt finished it yet
	; this can only restart one process per interval

	static Restartscript

    Process := obj.TargetInstance

	;verhindert die doppelte Ausführung
		If !RegExMatch(Process.CommandLine, "\w+\.ahk(?=\" q ")", Script) || (StrLen(Restartscript) > 0)
			return

	; prüft ob Addendum neu gestartet werden muss
		Restartscript := Script
		If InStr(Restartscript, "Addendum.ahk") 	{
			If WinExist("Addendum.ahk ahk_class #32770") 	{
				; falls eine Autohotkey Fehlermeldung angezeigt wird, wird Addendum.ahk nicht sofort neu gestartet
					WinGetText, wText, % "Addendum.ahk ahk_class #32770"
					If RegExMatch(wText, "Error:\s.*Line#.*--->") {
							RegExMatch(Process.CommandLine, "[A-Z]:[\w\\_\säöüÄÖÜ.]+\.ahk", SkriptPath)
							fehler := Object()
							fehler.Message	:= wText
							fehler.File      	:= SkriptPath
							FehlerProtokoll(fehler, 1)
					}
			}
			If !AHKProcessExist("Addendum.ahk")
				If ShowRestartProgress("Addendum.ahk", 5)
					RestartAddendum()
		}

		DoRestart := true
		Restartscript :=  ""

}

ProcessDelete_OnHighCPULoad(ProcName, PID, WTime                   	;-- monitors a process over a given time to detect process hung not detected by windows
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
				msg :=	"Name:                    " 	ProcName 	     		                                  	"`n"
					.    	"PID:                        " 	PID     			                                            	"`n"
					. 		"IdleTime:                " 	IdleTime "s"	                                               	"`n"
					.		"Round:                   " 	SubStr("0000" A_Index, -3) "/" ObservationLoops	"`n"
					.		"Process state:          " 	ProcessState	                                            	"`n"
					.		"Time:                      "	WTimeSec "s"                                                	"`n"
					.		"CPU total:              " 	SubStr("0" Round(SystemTime), -1)                 	"`n"
					. 		"CPU usage:            " 	SubStr("0" Round(ProcessTime), -1)	          		"`n"
					.		"av. CPU usage:       " 	Round(AvPT, 1)                                          	"`n"
					.		"drop observations: "  	DropObservation                                         	"`n"
					.		"Inactive times:        " 	InactiveObservation
					BTT(msg, A_ScreenWidth-400, A_ScreenHeight-400,, "Style4")
				}

			; process is 5 times in advance not responding this loop breaks to close this process
				If InactiveObservation >= 5
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
CountInactive:

	InactiveWait := true
	ProcessState := GetState(PID, 200)
	If !ProcessState
		InactiveObservation ++
	else
		InactiveObservation := false

	InactiveWait := false

return
}

ShowRestartProgress(scriptname, duration) {                                    	;-- zeigt die Zeit bis Addendum erneut gestartet wird

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

RestartAddendum(TimedCheck:=false) {                                           	;-- startet(prüft) Addendum.ahk

	; Variablen ;{
		static 	msgEnd 	:= "`n---------------------------------------------------------------`n"

		Addendum.TimerCall	:= A_TickCount
		PTSum 		              		:= 0
		Addendum.PID           	:= AHKProcessExist(Addendum.ScriptName)
	;}

	; Addendum.ahk Prozeß ist nicht vorhanden
		If !Addendum.PID {

			If (A_TimeIdlePhysical > 0)
				FormatTime, idleTime, % A_TimeIdlePhysical, HH:mm:ss

			FileAppend	, %   	TimeStamp()                                                  	", "
								.     	RegExReplace(A_ScriptName, "\.ahk$")            	", "
								.     	A_ComputerName                                           	", "
								.     	(TimedCheck ? "TIMER" : "EVENT")
								.    	(StrLen(idleTime) > 0 ? ", idle: " idleTime : "")  	"`n"
								, % 	Addendum.Dir "\logs'n'data\OnExit-Protokoll.txt", UTF-8

			Run   	    	, % Addendum.AHKH_exe " /f " q Addendum.ScriptPath q

			; Ausführung max. 20s
				PID := 0
				while (PID = 0) && (A_Index <= 200) {
				 PID := AHKProcessExist(Addendum.ScriptName)
				 Sleep 100
				}

		return PID
		}
	; über 4 Sekunden die CPU-Auslastung von Addendum messen
		else {

			; eine Überwachung läuft noch dann nichts machen
				If (TimedCheck <> "override") && ObservationRuns
					return

			; CPU Status (CPU Auslastung überschritten) oder Addendum.ahk antwortet nicht
				cpuStatus		:= GetCPUStatus(Addendum.PID, 10, Addendum.MaxAverageCPU)
				admStatus 	:= GetAddendumStatus()

				If (TimedCheck = "override")
					SciTEOutput("  " TimeStamp() " | CPU: " cpuStatus.Average " | Addendum: " (admStatus ? admStatus " ms" : "no answer"))

			; cpuStatus > 0 - Kontrollwert überschritten, admStatus = false - Addendum.ahk kann nicht antworten
				If cpuStatus.HighLoad || !admStatus {

					; Protokoll schreiben
						If Addendum.MonLogPath {
							msg :=	(cpuStatus.HighLoad ? "hohe CPU Auslastung registriert (" cpuStatus.Average ")`n                                `t" : "")
							msg .=	"Addendum: " (!admStatus ? "keine Antwort" : " Antwort brauchte " admStatus " ms")
						}

					; Addendum sendet keinen Status zurück => Neustart!
						If !admStatus
							If !ProcessClose(Addendum.PID)
								msg := (StrLen(Trim(msg)) > 0 ? msg "`n" : "") "Addendum.ahk konnte nicht per Process, Close beendet werden."
							else
								msg := (StrLen(Trim(msg)) > 0 ? msg "`n" : "") "Addendum.ahk wurde durch Process, Close beendet."

					; HighLoad wurde registriert, Messung der CPU-Last über mehrere Sekunden erfolgt jetzt
						else if cpuStatus.HighLoad
							If (procDelMsg := ProcessDelete_OnHighCPULoad(Addendum.ScriptName, Addendum.PID, Addendum.RestartAfter, Addendum.IdleTime, Addendum.MaxAverageCPU))
								msg := (StrLen(Trim(msg)) > 0 ? msg "`n" : "") . procDelMsg

					; schreibt das Protokoll
						If Addendum.MonLogPath & (StrLen(Trim(msg)) > 0)
							FileAppend, % TimeStamp() "|`t" msg . msgEnd, % Addendum.MonLogPath, UTF-8

				}

		}

}

GetCPUStatus(PID,WatchTime=4, MaxLoad=5, checkInterval:=100) {  ;-- Messung der CPU Auslastung

	; Rückgabeparameter ist 0 wenn der Grenzwert nicht überschritten wurden
	; andernfalls wird die ermittelte CPU Last zurückgegeben
	; MaxLoad - Grenzwert der CPU Last, Überschreitung generiert eine HighLoad Meldung

	PTStum  	:= 0
	starttime	:= A_TickCount

	;Loop % (loops := Floor((WatchTime*1000)/checkInterval)) {
	Loop {
		loops     	:= A_Index
		CPULoad 	:= 	GetProcessCpu(PID)
		PTSum  	+= 	CPULoad ? CPULoad : 0
		loopTime 	:= Round((A_TickCount - starttime)/1000, 1)
		If (loopTime >= WatchTime)
			break
		Sleep % checkInterval
	}

	PTAverage := PTSum/loops
	Average := StrSplit(PTAverage, ".").1 "." SubStr(StrSplit(PTAverage, ".").2, 1, 3)

return {"Average": Average, "HighLoad": (PTAverage>MaxLoad?1:0)}
}

GetAddendumStatus() {                                                                     	;-- testet ob Addendum.ahk antwortet

		global MsgRec

	; die ID des Addendum Fensters läßt sich nicht ermitteln
		If !(AddendumID := GetAddendumID())
			return 0

	; versucht Addendum zur einer Antwort zu überreden (wartet 4s auf Antwort)
		MsgRec := ""
		Send_WM_COPYDATA("Status||" Addendum.hMsgGui, GetAddendumID())
		starttime := A_TickCount
		while (MsgRec <> "okay") && ((recTime:=(A_TickCount-starttime)) < 4000)
			Sleep 1

return (MsgRec = "okay" ? recTime : 0)
}

AHKProcessExist(ProcName, cmd="") {                                                	;-- is only searching for Autohotkey processes

	; use cmd = "ProcessID" to receive the PID of an Autohotkey process

	StrQuery := "Select * from Win32_Process Where Name Like 'Autohotkey%'"
	For process in ComObjGet("winmgmts:\\.\root\CIMV2").ExecQuery(StrQuery)	{
		RegExMatch(process.CommandLine, "[\w\-\_]+\.ahk", name)
		If InStr(name, ProcName)
			If (StrLen(cmd) > 0)
				return process[cmd]
			else
				return process.ProcessID
	}

return false
}

ProcessClose(PID) {                                                                            	;-- force close a process, returns last PID
	Process, Close, % PID
return ErrorLevel
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
	prevProcTime := procTime
	prevTickCount := A_TickCount

return cpu
}

GetProcessTime(PID, ByRef time) {
   hProc := DllCall("OpenProcess", "UInt", PROCESS_QUERY_INFORMATION := 0x400, "UInt", false, "UInt", PID, "Ptr")
   DllCall("GetProcessTimes", "Ptr", hProc, "Ptr", &time, "Ptr", &time + 8, "Ptr", &time + 16, "Ptr", &time + 24)
   DllCall("CloseHandle", "Ptr", hProc)
}

GetProcessTimes(PID=0)    {                                                                	;-- ProcessTime = CPU Usage?

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

MessageWorker(InComing) {

	global MsgRec

	recv := {	"txtmsg"		: (StrSplit(InComing, "|").1)
				, 	"opt"     	: (StrSplit(InComing, "|").2)
				, 	"fromID"	: (StrSplit(InComing, "|").3)}

	If (recv.txtmsg = "Status")
		MsgRec := recv.opt

}

IniReadExt(SectionOrFullFilePath, Key:="", DefaultValue:="") {            	;-- eigene IniRead funktion für Addendum

	; beim ersten Aufruf der Funktion !nur! Übergabe des ini Pfades mit dem Parameter SectionOrFullFilePath
	; die Funktion behandelt einen geschriebenen Wert der "ja" oder "nein" ist, als Wahrheitswert, also true oder false
	; letzte Änderung: 08.06.2020

		static WorkIni

	; Arbeitsini Datei wird erkannt wenn der übergebene Parameter einen Pfad ist
		If RegExMatch(SectionOrFullFilePath, "^[A-Z]\:.*\\")	{
			If !FileExist(SectionOrFullFilePath)	{
					MsgBox,, Addendum für AlbisOnWindows, % "Die .ini Datei existiert nicht!`n`n" WorkIni "`n`nDas Skript wird jetzt beendet.", 10
					ExitApp
			}
			WorkIni := SectionOrFullFilePath
			return WorkIni
		}

	; Section, Key einlesen
		IniRead, OutPutVar, % WorkIni, % SectionOrFullFilePath, % Key

	; Bearbeiten des Wertes vor Rückgabe
		If InStr(OutPutVar, "ERROR")
			If (StrLen(DefaultValue) > 0) ; Defaultwert vorhanden, dann diesen Schreiben und Zurückgeben
				IniWrite, % (OutPutVar := DefaultValue), % WorkIni, % SectionOrFullFilePath, % key
			else
				return "ERROR"
		else if InStr(OutPutVar, "%AddendumDir%")
				OutPutVar := StrReplace(OutPutVar, "%AddendumDir%", Addendum.Dir)
		else if RegExMatch(OutPutVar, "i)^\s*(ja|nein)\s*$", bool)
				OutPutVar := (bool1= "ja") ? true : false

return Trim(OutPutVar)
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

TimeStamp() {
	FormatTime, time, % A_Now, dd.MM.yyyy HH:mm:ss
return time
}

Receive_WM_COPYDATA(wParam, lParam) {                                   	;-- empfängt Nachrichten von anderen Skripten

    StringAddress	:= NumGet(lParam + 2*A_PtrSize)
	fn_MsgWorker	:= Func("MessageWorker").Bind(StrGet(StringAddress))
	SetTimer, % fn_MsgWorker, -10

return
}

Send_WM_COPYDATA(ByRef StringToSend, ScriptID) {                       	;-- für die Interskriptkommunikation - keine Netzwerkkommunikation!

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

GetAddendumID() {                                                                        	;-- für Interskriptkommunikation
	PDetectHiddenWindows	:= A_DetectHiddenWindows
    PTitleMatchMode         	:= A_TitleMatchMode
    DetectHiddenWindows On
    SetTitleMatchMode 2
	AddendumID := WinExist("Addendum Message Gui ahk_class AutoHotkeyGUI")
	DetectHiddenWindows 	, % PDetectHiddenWindows  	; Restore original setting for the caller.
    SetTitleMatchMode     	, % PTitleMatchMode           	; Same.
return AddendumID
}

;}

;{ Includes
#Include %A_ScriptDir%\..\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\..\lib\class_Btt.ahk
#Include %A_ScriptDir%\..\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\..\lib\SciTEOutput.ahk
;}


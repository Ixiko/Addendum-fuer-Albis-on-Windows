;-----------------------------------------------------------------------------------------------------------------------------------
;------------------------------------------------ ADDENDUM MONITOR ----------------------------------------------------
;-----------------------------------------------------------------------------------------------------------------------------------
													Version:= "1.0" , vom:= "25.02.2021"
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
		DetectHiddenWindows	, Off     	; Off is default
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

		RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)

		global CompName	:= StrReplace(A_ComputerName, "-")

		global Restartscript
		global ObservationRuns	:= false
		global DoRestart          	:= true
		global q                       	:= Chr(0x22)
		global winmgmts       	:= ComObjGet("winmgmts:") 	; Get WMI service object.
		global Addendum      	:= Object()

	; Addendum object
		Addendum.Dir           	:= AddendumDir
		Addendum.Ini           	:= AddendumDir "\Addendum.ini"
		Addendum.compname	:= StrReplace(A_ComputerName, "-")

		If FileExist(A_AppData "\AutoHotkeyH\AutoHotkeyH_U64.exe")
			Addendum.AHKH_exe	:= A_AppData "\AutoHotkeyH\AutoHotkeyH_U64.exe"
		else if FileExist(AddendumDir "\include\AHK_H\x64w\AutoHotkeyH_U64.exe")
			Addendum.AHKH_exe	:= Addendum.Dir "\include\AHK_H\x64w\AutoHotkeyH_U64.exe"
		else
			throw A_ScriptName ": AutohotkeyH_U64.exe ist nicht vorhanden.`nDas Skript kann nicht ausgeführt werden!"

	; Scripticon
		Menu, Tray, NoStandard
		If (StrLen(hIcon := Base64toHICON()) > 0)
			Menu, Tray, Icon, % "HICON: " hIcon
		else If FileExist(IconDir := Addemdum.Dir "\assets\ModulIcons\AddendumMonitor.ico")
			Menu, Tray, Icon, % IconDir

		Menu, Tray, Add, % StrReplace(A_ScriptName, ".ahk")	, % "admMonInfo"
		Menu, Tray, Add
		Menu, Tray, Add, % "Reload"		                            	, % "admMonReload"
		Menu, Tray, Add, % "Exit"			                             	, % "admMonExitApp"
		Menu, Tray, Add
		Menu, Tray, Add, % "nächster Check in "                   	, % "admMonTimerInfo"

	; get data from Addendum.ini
		IniReadExt(Addendum.Ini)

	; Addendum DBPath
		Addendum.DBPath := IniReadExt("Addendum", "AddendumDBPath")
		If InStr(Addendum.DBPath, "Error") || !RegExMatch(Addendum.DBPath, "i)[a-z]\:\") || !InStr(FileExist(Addendum.DBPath), "D")
			Addendum.DBPath := ""
		If !InStr(FileExist(Addendum.DBPath "\sonstiges"), "D")
			FileCreateDir, % Addendum.DBPath "\sonstiges"

	; loading Telegram BotToken and BotChatID
		BotName	:= IniReadExt("Telegram", "Bot1")
		If !InStr(BotName, "Error") {
				Addendum.Telegram := Object()
				BotToken	:= IniReadExt("Telegram", BotName "_Token")
				BotChatID	:= IniReadExt("Telegram", BotName "_ChatID")
				Addendum.Telegram.Push({"BotName":BotName, "Token": BotToken, "ChatID": BotChatID})
		}

	; loads restart properties (script detects high cpu usage of Addendum.ahk and will close-restart Addendum.ahk)
		Addendum.MaxAverageCPU 	:= IniReadExt("Addendum", "HighCPU_MaxAvarageCPU"	, 5)
		Addendum.RestartAfter       	:= IniReadExt("Addendum", "HighCPU_RestartAfter"         	, 360) 	; seconds
		Addendum.MinIdleTime     	:= IniReadExt("Addendum", "HighCPU_MinIdleTime"       	, 180)	; seconds
		Addendum.TimedCheck     	:= IniReadExt("Addendum", "HighCPU_Timer"                    	, 180)	; seconds
		TrTip1 := 	"CPU Last Grenze:   "    	Addendum.MaxAverageCPU 	"%`n"
				  .		"Neustart nach:      " 		Addendum.RestartAfter        	"s`n"
				  .		"Überprüfung alle: "   	Addendum.TimedCheck      	"s"
		TrTip2 := 	"CPU Überlastung:                     `n"
				 . 	  	"nächster Check in:                      "
		Menu, Tray, Tip, % TrTip1 "`n" TrTip2

	; Interskript communication gui
		Gui, AMsg: New	, +HWNDhMsgGui +ToolWindow
		Gui, AMsg: Add	, Edit, xm ym w100 h100 HWNDhEditMsgGui
		Gui, AMsg: Show	, AutoSize NA Hide, % "Addendum Monitor Gui"
		Addendum.hMsgGui     	:= hMsgGui
		Addendum.hEditMsgGui	:= hEditMsgGui

	; starts receiving messages
		OnMessage(0x4A	, "Receive_WM_COPYDATA")
		OnMessage(0x404, "AHK_NOTIFYICON")

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

;}


return

^+!ö::ExitApp
^!ö::    	;{ Skriptneustart
	TrayTip, AddendumMonitor, Skript wird neu gestartet, 1
	Sleep 1000
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
;}

;{ TrayLabels

AHK_NOTIFYICON(wParam, lParam) {
	if (lParam = 0x200) { ; WM_LBUTTONUP
	SetTimer, admMonTimerInfo, -1
        return 0
    }
}
admMonReload:
Reload
admMonExitApp:
ExitApp
admMonTimerInfo:
	nextcall 		:= Addendum.TimedCheck-((A_TickCount - Addendum.TimerCall)/1000)
	nextcall_min	:= Floor(nextcall/60)
	nextcall_sec	:= Floor(nextcall - (nextcall_min*60))
	TrTip2   		:= "CPU Überlastung:  " 	(ObservationRuns ? "ja":"nein") "`n"
						. 	 "nächster Check in: " 	nextcall_min "m " SubStr("00" nextcall_sec, -1) "s"
	;TrayTip, Addendum Monitor, % TTipMsg, 1, 16
	Menu, Tray, Tip, % TrTip1 "`n" TrTip2
return
admMonInfo:
return
;}

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
				If (A_Index >= InfoAfterLoop) && (IdleTime >= MinIdleTime) {
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

		global 	MsgRec
		static 	ProcName := "Addendum.ahk"

		Addendum.TimerCall := A_TickCount
		PTSum 				:= 0
		AddendumPID	:= AHKProcessExist("Addendum.ahk")

	; Speicherverbrauch reduzieren
		If TimedCheck
			ScriptMem()

	; Addendum.ahk Prozeß ist nicht vorhanden
		If !AddendumPID {

			If (A_TimeIdlePhysical > 0)
				FormatTime, idleTime, % A_TimeIdlePhysical, HH:mm:ss

			FileAppend	, %   	TimeStamp()                                                  	","
								.     	RegExReplace(A_ScriptName, "\.ahk$")            	","
								.     	A_ComputerName                                           	","
								.     	(TimedCheck ? "TIMER" : "EVENT")
								.    	(StrLen(idleTime) > 0 ? ",idle: " idleTime : "")  	"`n"
								, % 	Addendum.Dir "\logs'n'data\OnExit-Protokoll.txt"

			Run   	    	, % Addendum.AHKH_exe " /f " q Addendum.Dir "\Module\Addendum\Addendum.ahk" q

		}
	; über 4 Sekunden die CPU-Auslastung von Addendum messen
		else {

			; eine Überwachung läuft noch dann nichts machen
				If ObservationRuns
					return

			; liegt die CPU-Last über dem maximal Wert oder Addendum.ahk ist inaktiv
			; wird eine längere Überwachung gestartet
				Loop 20 {
					PTSum += GetProcessTimes(AddendumPID)
					Sleep 200
				}

			; versucht Addendum zur einer Antwort zu überreden (wartet 6s auf Antwort)
				Send_WM_COPYDATA("Status||" Addendum.hMsgGui, GetAddendumID())
				starttime := A_TickCount
				while (MsgRec <> "okay") && ((recTime:=(A_TickCount-starttime)) < 4000)
					Sleep 1
				SciTEOutput(" " A_Hour ":" A_Min "| Addendumstatus: " MsgRec " - received after " recTime " ms (CPU Load: " Round(PTSum/20, 1) ")")
				If (StrLen(MsgRec) = 0)
					FileAppend, % TimeStamp() "| Addendum hat nicht geantwortet (CPU Load: " Round(PTSum/20, 1) ")", % Addendum.DBPath "\sonstiges\AddendumMonitorLog.txt", UTF-8
				MsgRec := ""

			; Antwort ja, aber dennoch Überwachungstarten
				If (PTSum/20 >= Addendum.MaxAverageCPU) || !GetState(AddendumPID, 200) {
					If Addendum.DBPath
						FileAppend, % TimeStamp() "| hohe CPU Auslastung registriert (" Round(PTSum/20, 1) ")", % Addendum.DBPath "\sonstiges\AddendumMonitorLog.txt", UTF-8
					If (msg := ProcessDelete_OnHighCPULoad(ProcName, AddendumPID, Addendum.RestartAfter, Addendum.IdleTime, Addendum.MaxAverageCPU))
						If Addendum.DBPath		; schreibt ein Protokoll
							FileAppend, % TimeStamp() "`n" msg "`n---------------------------------------------------------------`n", % Addendum.DBPath "\sonstiges\AddendumMonitorLog.txt", UTF-8
				}

		}

}

AHKProcessExist(ProcName, cmd="") {                                              	;-- is only searching for Autohotkey processes

	; use cmd = "PID" to receive the PID of an Autohotkey process

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

GetProcessTimes(PID=0)    {                                                            	;-- ProcessTime = CPU Usage?

   Static oldKrnlTime, oldUserTime
   Static newKrnlTime, newUserTime

   oldKrnlTime := newKrnlTime, oldUserTime := newUserTime

   hProc := DllCall("OpenProcess", "Uint", 0x400, "Uint", 0, "Uint", pid)
   DllCall("GetProcessTimes", "Uint", hProc, "int64*", CreationTime, "int64*", ExitTime, "int64*", newKrnlTime, "int64*", newUserTime)
   DllCall("CloseHandle", "Uint", hProc)

Return (newKrnlTime-oldKrnlTime + newUserTime-oldUserTime)/10000000 * 100
}

GetSystemTimes() {                                                                       	;-- Total CPU Load

   Static oldIdleTime, oldKrnlTime, oldUserTime
   Static newIdleTime, newKrnlTime, newUserTime

   oldIdleTime := newIdleTime
   oldKrnlTime := newKrnlTime
   oldUserTime := newUserTime

   DllCall("GetSystemTimes", "int64P", newIdleTime, "int64P", newKrnlTime, "int64P", newUserTime)
   Return (1 - (newIdleTime-oldIdleTime)/(newKrnlTime-oldKrnlTime + newUserTime-oldUserTime)) * 100
}

GetState(PID,TimeOut := 200) {                                                       	;-- Prozessstatus

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

ScriptMem() {                                                                                   	;-- gibt belegten Speicher frei

	static PID := 0

	If !PID{
		DHW := A_DetectHiddenWindows, TMM := A_TitleMatchMode
		DetectHiddenWindows	, On
		SetTitleMatchMode		, 2
		WinGet, PID, PID, % "\" A_ScriptName " ahk_class AutoHotkey"
		DetectHiddenWindows	, % DHW
		SetTitleMatchMode	 	, % TMM
	}
	hProc	:= DllCall("OpenProcess", "UInt", 0x001F0FFF, "Int", 0, "Int", PID)
	result	:= DllCall("SetProcessWorkingSetSize", "UInt", hProc, "Int", -1, "Int", -1)
	DllCall("CloseHandle", "Int", hProc)

return result
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

Base64toHICON() {                                                                       	;-- 16x16 PNG image (236 bytes), requires WIN Vista and later
Local B64 := "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"
local Bin, Blen, nBytes:=2304, hICON:=0

  VarSetCapacity( Bin,nBytes,0 ), BLen := StrLen(B64)
  If DllCall( "Crypt32.dll\CryptStringToBinary", "Str",B64, "UInt",BLen, "UInt",0x1
            , "Ptr",&Bin, "UIntP",nBytes, "Int",0, "Int",0 )
     hICON := DllCall( "CreateIconFromResourceEx", "Ptr",&Bin, "UInt",nBytes, "Int",True
                     , "UInt",0x30000, "Int",16, "Int",16, "UInt",0, "UPtr" )
Return hICON
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
	Prev_DetectHiddenWindows := A_DetectHiddenWindows
    Prev_TitleMatchMode := A_TitleMatchMode
    DetectHiddenWindows On
    SetTitleMatchMode 2
	AddendumID := WinExist("Addendum Message Gui ahk_class AutoHotkeyGUI")
	DetectHiddenWindows 	, % Prev_DetectHiddenWindows  	; Restore original setting for the caller.
    SetTitleMatchMode     	, % Prev_TitleMatchMode           	; Same.
return AddendumID
}


#Include %A_ScriptDir%\..\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\..\lib\class_Btt.ahk
#Include %A_ScriptDir%\..\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\..\lib\SciTEOutput.ahk
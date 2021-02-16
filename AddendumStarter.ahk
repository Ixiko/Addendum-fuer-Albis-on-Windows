;	------------------------------------------------------------------------------------------------------------------------------------
;	                                                                  AddendumH-Start-Skript
;	                                                             Addendum für AlbisOnWindows
;	-----------------------------------------------------------------------------------------------------------------------------------
;	-----------------------------------------------------------------------------------------------------------------------------------
;
;	1.       		Skript für den automatischen Aufruf von Addendum.ahk aus dem Autostart-Ordner heraus
;   2.   	          stellt sicher das Addendum mit der Autohotkey_H Unicode 64bit Variante gestartet wird.
;                                AutohotkeyH_U64.exe (multithread unicode 64bit Autohotkey-Version)
;	3.           Installiert notwendige Fonts, legt nach dem ersten Start Datenordner im Addendum-Verzeichnis an
;
;	-----------------------------------------------------------------------------------------------------------------------------------
;	-----------------------------------------------------------------------------------------------------------------------------------
;	                                               written by Ixiko -this version is from 05.04.2020
;	                                    please report errors and suggestions to me: Ixiko@mailbox.org
;	                                use subject: "Addendum" so that you don't end up in the spam folder
;	                                         GNU Lizenz - can be found in main directory  - 2017
;	-----------------------------------------------------------------------------------------------------------------------------------

	#NoEnv
	#NoTrayIcon
	#SingleInstance Force
	SetBatchLines, -1
	FileEncoding, UTF-8
	OnError("FehlerProtokoll")

;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Client Namen feststellen
;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	CompName	:= StrReplace(A_ComputerName, "-")
	len            	:= Floor((20-StrLen(CompName))/2)
	TTipCN     	:= CompName . SubStr("                   ", 1, len)
	Futura       	:= 0
	jkp            	:= 0

; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Speichern des Pfades (AddendumDir) in einer versteckten Datei im albiswin.loc Ordner
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	If !FileExist("C:\albiswin.loc\AddendumDir")
	{
			If !InStr(FileExist("C:\albiswin.loc"), "D")
						FileCreateDir, % "C:\albiswin.loc"
			FileAppend, % A_ScriptDir, % "C:\albiswin.loc\AddendumDir"
			FileSetAttrib, +H, % "C:\albiswin.loc\AddendumDir"
	}
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Laborordner anlegen falls nicht vorhanden
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	If !Instr( FileExist(A_ScriptDir . "\logs'n'data\_DB\Labordaten"),  "D")
				FileCreateDir, % A_ScriptDir "\logs'n'data\_DB\Labordaten"

; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Überprüfen ob die von Addendum benutzten Fonts installiert sind (derzeit Futura und jkpAwesome)
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	Loop 6
		If !FileExist("C:\Windows\Fonts\TT014" . (A_Index - 1) . "M_.TTF")
		{
				FileCopy, % A_ScriptDir "\support\Fonts\Futura_Bk_Fontset\TT014" . (A_Index - 1) . "M_.TTF", "C:\Windows\Fonts\"
				If ErrorLevel
					Futura ++
		}

	If !FileExist("C:\Windows\Fonts\jkpAwesome.ttf")
	{
			FileCopy, % A_ScriptDir "\support\Fonts\jkpAwesome.ttf", "C:\Windows\Fonts\"
			If ErrorLevel
					jkp ++
	}

	;If Futura || jkp
	;		MsgBox, 1, Addendum für Albis on Windows, Addendum kann nicht gestartet werden!`nerforderliche Schriftarten konnten nicht nach C:\Windows\Fonts kopiert werden!`n%Futura%%jkp%

; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; nachschauen und dann auch nachfragen, ob das Skript in den Autostart-Ordner geschrieben werden darf
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	AutoStartOrdner    	:= []
	AutoStartOrdner[1]	:= A_Startup
	AutoStartOrdner[2]	:= A_StartupCommon
	AutoStartOrdner[3] 	:= A_AppData "\Microsoft\Windows\Start Menu\Programs\Startup"	;\Microsoft\Windows\Start Menu\Programs\Startup

	Loop 2	{

			If !FileExist(AutoStartOrdner[A_Index] "\AddendumStarter") 	{

				Exception("Kein Link für Addendum im Autostartordner (" AutoStartOrdner[A_Index] ") vorhanden.")
				IniRead, Autostart, % A_ScriptDir "\Addendum.ini", AutoStartAbfrage, % CompName
				If InStr(Autostart, "ERROR") {

					MsgBox, 4, Addendum für Albis on Windows, % "Möchten Sie Addendum dem Autostart-Ordner hinzufügen? "
					IfMsgBox, Yes
					{
						IniWrite, 1	, % A_ScriptDir "\Addendum.ini", AutoStartAbfrage, % CompName
						FileCreateShortcut, % A_ScriptDir "\AddendumStarter.ahk", % AutoStartOrdner[A_Index] "\AddendumStarter.lnk", % A_ScriptDir,, % "Startet Addendum für AlbisOnWindows", % A_ScriptDir "\support\Addendum.ico"
					}

					IfMsgBox, No
						IniWrite, 0	, % A_ScriptDir "\Addendum.ini", AutoStartAbfrage, % CompName

				}
				else if (Autostart = 1)
					FileCreateShortcut, % A_ScriptDir "\AddendumStarter.ahk", % AutoStartOrdner[A_Index] "\AddendumStarter.lnk", % A_ScriptDir,, % "Startet Addendum für AlbisOnWindows", % A_ScriptDir "\support\Addendum.ico"

			}
	}

; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Die AutohotkeyH exe in das AppData Verzeichnis kopieren - der lokale Start vermeidet Konflikte mit der Kernel dll -> ntdll.dll
; bei einer neuen AutohotkeyH Version wird die Datei im AppData Verzeichnis ersetzt
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------

	If !FileExist(A_ScriptDir "\include\AHK_H\AutoHotkeyH_U64.exe") {

		MsgBox, 1, Addendum für Albis on Windows, %	"Die Datei  *AutoHotkeyH_U64.exe* ist nicht vorhanden.`n"
																			.	"Addendum benötigt diese Datei zum Ausführen der Skripte.`n"
																			.	"Bitte überprüfen Sie das Verzeichnis:`n"
																			.	"\include\AHKH`n"
		ExitApp

	}

	If !Instr( FileExist(A_AppData "\AutohotkeyH"),  "D")                                ; noch kein Verzeichnis vorhanden
		FileCreateDir	, % A_AppData "\AutoHotkeyH"

	If FileExist(A_ScriptDir "\include\AHK_H\AutoHotkeyH_U64.exe") {

		FileGetTime, installDirTime	, % A_ScriptDir "\include\AHK_H\AutoHotkeyH_U64.exe"
		FileGetTime, AppDirTime  	, % A_AppData "\AutoHotkeyH\AutoHotkeyH_U64.exe"

		If !FileExist(A_AppData "\AutohotkeyH\AutoHotkeyH_U64.exe") || (installDirTime <> AppDirTime)            ; noch keine Datei vorhanden oder wurde versehentlich gelöscht
			FileCopy, % A_ScriptDir "\include\AHK_H\AutoHotkeyH_U64.exe", % A_AppData "\AutoHotkeyH\AutoHotkeyH_U64.exe", 1 ; 1 = überschreiben

	}

	If FileExist(A_AppData "\AutohotkeyH\AutoHotkeyH_U64.exe")
		AHKH_FilePath := A_AppData "\AutoHotkeyH\AutoHotkeyH_U64.exe"
	else {
		MsgBox, Kann keine AutohotkeyH.exe finden.
		ExitApp
	}

; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Albis starten wenn es nicht gestartet ist
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	If !AlbisExist() 	{
		IniRead, AlbisDir, % A_ScriptDir "\Addendum.ini", Albis, AlbisWorkDir
		IniRead, AlbisExe, % A_ScriptDir "\Addendum.ini", Albis, AlbisExe
		If !InStr(AlbisDir, "Error") && !InStr(AlbisExe, "Error")
				Run, % AlbisDir "\" AlbisExe, % "c:\albiswin.loc"
	}

; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Pause solange das Addendum-Skript noch läuft, keine zweite Instanz starten! - vermutlich funktionierte deshalb der Skriptneustart nicht
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	If ScriptIsRunning("Addendum") {

		wait 			:= 1 	; in Sekunden
		Sleeping 	:= 20	; Check Intervall in ms

		Progress, % "B1 M P0 T cW202842 cB8088C2 cTFFFFFF zH25 w" 500 " WM400 WS500", % "Addendum wird noch ausgeführt, warte auf die Beendigung ...", % wait " s", % "AddendumStarter wartet ...", Futura Bk Bt

		hwnd := WinExist("AddendumStarter wartet")
		WinGetPos	,,,, wh	, % "ahk_id " hwnd
		WinGetPos	,,,, th	, ahk_class Shell_TrayWnd
		WinMove		, % "ahk_id " hwnd,,, % A_ScreenHeight - wh - th


		while ScriptIsRunning("Addendum")  {

			Progress, % Floor(((A_Index*Sleeping/1000)*100)/wait)
			ControlSetText, Static1, % Round(wait - (A_Index*Sleeping/1000),1) " s", % "ahk_id " hwnd

			If (A_Index*Sleeping <=  wait*1000)
					Sleep % Sleeping
			else
					break

		}

		Progress, Off
	}

; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Addendum mit AutohotkeyH 64bit Unicode Version starten
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
	If (ProcID := ScriptIsRunning("Addendum")) {

			MsgBox, 4, Addendum für Albis on Windows, Addendum.ahk wird noch ausgeführt.`nMöchten Sie es neustarten?, % wait * 2
			IfMsgBox, Yes
			{
					Process, Close, % ProcID
					Run, %AHKH_FilePath% /f "%A_ScriptDir%\Module\Addendum\Addendum.ahk",, UseErrorLevel
					If ErrorLevel
						MsgBox, 1, Addendum für Albis on Windows, Addendum konnte nicht gestartet werden!, 4
			}

	}
	else
			Run, %AHKH_FilePath% /f "%A_ScriptDir%\Module\Addendum\Addendum.ahk",, UseErrorLevel

	If !ScriptIsRunning("AddendumMonitor")
		Run, %AHKH_FilePath% /f "%A_ScriptDir%\Module\Addendum\threads\AddendumMonitor.ahk",, UseErrorLevel
;}

ExitApp

; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Funktionen
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
ScriptIsRunning(scriptname) {

	scriptname := RegExReplace(scriptname, "\.ahk$")

	For Process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process") {
		If RegExMatch(Process.Name, "AHK") || RegExMatch(Process.Name, "i)Autohotkey")  {
            If RegExMatch(Process.CommandLine, "\\" scriptname "\.ahk")
				  return Process.ProcessId
		}
	}

return false
}

AlbisExist() {                                                                                      	;-- schaut nach ob Albis läuft, gibt wahr oder falsch zurück
	;result = 0 wenn kein Albis Prozess läuft , andernfalls die PID(Prozess-ID) des Albis Prozesses
	result:= WMIEnumProcessExist("Albis")
	result+= WMIEnumProcessExist("AlbisCS")
return result
}

WMIEnumProcessExist(ProcessSearched) {                                          		;-- logische Funktion zur Suche ob ein bestimmter Prozeß existiert

    ;Funktion gibt nur 0 - für nicht vorhanden und 1 für vorhanden aus (als schnell Variante gedacht)
	;zudem ist der Funktion die exakte Schreibweise des Prozeß egal (GROSS/klein oder nur ein Teil davon)

	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
		   liste .= process.Name . "`n"

	Sort, liste, U

	Loop, parse, liste, `n
		If Instr(A_LoopField, ProcessSearched, false ,1)
			return true

return false
}

; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Include
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Include %A_ScriptDir%\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\lib\SciTEOutput.ahk


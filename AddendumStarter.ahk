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
;	                                               written by Ixiko -this version is from 21.08.2019
;	                                    please report errors and suggestions to me: Ixiko@mailbox.org
;	                                use subject: "Addendum" so that you don't end up in the spam folder
;	                                         GNU Lizenz - can be found in main directory  - 2017
;	-----------------------------------------------------------------------------------------------------------------------------------

	#NoEnv
	SetBatchLines, -1
	FileEncoding, UTF-8
	OnError("FehlerProtokoll")

;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Client Namen feststellen
;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	CompName:= StrReplace(A_ComputerName, "-")
	len:= Floor((20-StrLen(CompName))/2)
	TTipCN:= CompName . SubStr("                   ", 1, len)

; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Speichern des Pfades (AddendumDir) in einer versteckten Datei im albiswin.loc Ordner
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	If !FileExist("C:\albiswin.loc\AddendumDir")
	{
			if !InStr(FileExist("C:\albiswin.loc"), "D")
						FileCreateDir, % "C:\albiswin.loc"
			FileAppend, % A_ScriptDir, % "C:\albiswin.loc\AddendumDir"
			FileSetAttrib, +H, % "C:\albiswin.loc\AddendumDir"
	}
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Laborordner anlegen falls nicht vorhanden
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	If !Instr( FileExist(A_ScriptDir . "\logs'n'data\_DB\Labordaten"),  "D")
				FileCreateDir, % A_ScriptDir . "\logs'n'data\_DB\Labordaten"

; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Überprüfen ob die von Addendum benutzten Fonts installiert sind (derzeit Futura und jkpAwesome)
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	Loop 6
	If !FileExist("C:\Windows\Fonts\TT014" . (A_Index - 1) . "M_.TTF")
		FileCopy, % A_ScriptDir . "\support\Fonts\Futura_Bk_Fontset\TT014" . (A_Index - 1) . "M_.TTF", "C:\Windows\Fonts\"

	If !FileExist("C:\Windows\Fonts\jkpAwesome.ttf")
		FileCopy, % A_ScriptDir . "\support\Fonts\jkpAwesome.ttf", "C:\Windows\Fonts\"

; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; nachschauen und dann auch nachfragen, ob das Skript in den Autostart-Ordner geschrieben werden darf
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	AutoStartOrdner    	:= A_AppData "\Microsoft\Windows\Start Menu\Programs\Startup"	;\Microsoft\Windows\Start Menu\Programs\Startup
	AutoStartOrdner    	:= []
	AutoStartOrdner[1]	:= A_Startup
	AutoStartOrdner[2]	:= A_StartupCommon

	Loop, 2
	{
		   ;FileCreateShortcut,                           ZIEL						  , 							VERKNÜPFUNG					, ARBEITSVERZEICHNIS	, PARAMETER, BESCHREIBUNG, SYMBOLDATEI, TASTENKÜRZEL, SYMBOLNUMMER
			;FileCreateShortcut, % A_ScriptDir "\AddendumStarter.ahk", % AutoStartOrdner[A_Index] "\AddendumStarter.lnk", % A_ScriptDir				,					, % "Startet Addendum für AlbisOnWindows", % A_ScriptDir "\support\Addendum.ico"
			If !FileExist(AutoStartOrdner[A_Index] "\AddendumStarter")
			{
					Exception("Kein Link für Addendum im Autostartordner (" AutoStartOrdner[A_Index] ") vorhanden.")
					IniRead, Autostart, % A_ScriptDir "\Addendum.ini", AutoStartAbfrage, % CompName
					If InStr(Autostart, "ERROR")
					{
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
; Albis starten wenn es nicht gestartet ist
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	If !AlbisExist()
	{
		IniRead, AlbisDir, % A_ScriptDir "\Addendum.ini", Albis, AlbisDir
		IniRead, AlbisExe, % A_ScriptDir "\Addendum.ini", Albis, AlbisExe
		If !InStr(AlbisDir, "Error") && !InStr(AlbisExe, "Error")
				Run, % AlbisDir "\" AlbisExe
	}

; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Pause solange das Addendum-Skript noch läuft, keine zweite Instanz starten! - vermutlich funktionierte deshalb der Skriptneustart nicht
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	while ScriptExist("Addendum.ahk")
		If A_Index < 40
			Sleep 200
		else
			ExitApp

; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Addendum mit AutohotkeyH 64bit Unicode Version starten
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	Run, %A_ScriptDir%\include\AHK_H\x64w\AutohotkeyH_U64.exe /f "%A_ScriptDir%\Module\Addendum\Addendum.ahk"
	;Run, %A_ScriptDir%\include\AHK_H\x64w\AutohotkeyH_U64.exe /f "%A_ScriptDir%\Module\Addendum\AddendumMonitor.ahk"
	ExitApp

; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Funktionen
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------

ScriptExist(scriptname) {															                     	;-- true oder false wenn ein Skript schon ausgeführt wird

	dtWin:= A_DetectHiddenWindows
	DetectHiddenWindows, On

	WinGet, hwnd, List, ahk_class AutoHotkey

	Loop, % hwnd
	{
			ID := hwnd%A_Index%
			WinGetTitle, Titel, ahk_id %ID%
			WinGet, PID, PID, ahk_id %ID%
			SplitPath, Titel, runningAHK
			If InStr(runningAHK, scriptname)
							return PID
	}

	DetectHiddenWindows, % dtWin

return 0
}

AlbisExist() {																					;-- schaut nach ob Albis läuft, gibt wahr oder falsch zurück
	;result = 0 wenn kein Albis Prozess läuft , andernfalls die PID(Prozess-ID) des Albis Prozesses
	result:= WMIEnumProcessExist("Albis")
	result+= WMIEnumProcessExist("AlbisCS")
return result
}

WMIEnumProcessExist(ProcessSearched) {                                                            		;-- logische Funktion zur Suche ob ein bestimmter Prozeß existiert

    ;Funktion gibt nur 0 - für nicht vorhanden und 1 für vorhanden aus (als schnell Variante gedacht)
	;zudem ist der Funktion die exakte Schreibweise des Prozeß egal (GROSS/klein oder nur ein Teil davon)

	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
		   liste .= process.Name . "`n"

	Sort, liste, U

	Loop, parse, liste, `n
		If Instr(A_LoopField, ProcessSearched, false ,1)
			return 1

	return 0
}

; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Include
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	#Include %A_ScriptDir%\include\Addendum_Protocol.ahk


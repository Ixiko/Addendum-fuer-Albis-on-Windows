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
;	                                               written by Ixiko -this version is from 28.07.2021
;	                                    please report errors and suggestions to me: Ixiko@mailbox.org
;	                                use subject: "Addendum" so that you don't end up in the spam folder
;	                                         GNU Lizenz - can be found in main directory  - 2017
;	-----------------------------------------------------------------------------------------------------------------------------------

	#NoEnv
	#NoTrayIcon
	#SingleInstance Force
	SetBatchLines, -1
	ListLines, Off
	FileEncoding, UTF-8
	OnError("FehlerProtokoll")

; -------------------------------------------------------------------------------------------------------------------------------------
;	Client Namen feststellen
; ------------------------------------------------------------------------------------------------------------------------------------- ;{
	CompName	:= StrReplace(A_ComputerName, "-")
	len            	:= Floor((20-StrLen(CompName))/2)
	TTipCN     	:= CompName . SubStr("                   ", 1, len)
	Futura       	:= 0
	jkp            	:= 0
;}

; -------------------------------------------------------------------------------------------------------------------------------------
;	erste Aufruf von Addendum für Albis on Windows, dann müssen zunächst ein paar Dinge angelegt werden
; -------------------------------------------------------------------------------------------------------------------------------------
	global AddendumDir
	AddendumDir := A_ScriptDir

; -------------------------------------------------------------------------------------------------------------------------------------
; Addendum Datenordner anlegen falls nicht vorhanden
; -------------------------------------------------------------------------------------------------------------------------------------;{
	If !Instr(FileExist(A_ScriptDir "\logs'n'data\_DB"),  "D")
		FileCreateDir, % A_ScriptDir "\logs'n'data\_DB"
	If !Instr(FileExist(A_ScriptDir "\logs'n'data\_DB\Labordaten"),  "D")
		FileCreateDir, % A_ScriptDir "\logs'n'data\_DB\Labordaten"

;}

; -------------------------------------------------------------------------------------------------------------------------------------
; Überprüfen ob die von Addendum benutzten Fonts installiert sind (derzeit Futura und jkpAwesome)
; -------------------------------------------------------------------------------------------------------------------------------------;{
	Loop 6
		If !FileExist("C:\Windows\Fonts\TT014" . (A_Index - 1) . "M_.TTF")		{
			FileCopy, % A_ScriptDir "\support\Fonts\Futura_Bk_Fontset\TT014" . (A_Index - 1) . "M_.TTF", "C:\Windows\Fonts\"
			If ErrorLevel
				Futura ++
		}

	If !FileExist("C:\Windows\Fonts\jkpAwesome.ttf")	{
			FileCopy, % A_ScriptDir "\support\Fonts\jkpAwesome.ttf", "C:\Windows\Fonts\"
			If ErrorLevel
					jkp ++
	}

	;If Futura || jkp
	;		MsgBox, 1, Addendum für Albis on Windows, Addendum kann nicht gestartet werden!`nerforderliche Schriftarten konnten nicht nach C:\Windows\Fonts kopiert werden!`n%Futura%%jkp%
;}

; -------------------------------------------------------------------------------------------------------------------------------------
; nachschauen und dann auch nachfragen, ob das Skript in den Autostart-Ordner geschrieben werden darf
; -------------------------------------------------------------------------------------------------------------------------------------;{
	AutoStartOrdner    	:= []
	AutoStartOrdner[1]	:= A_Startup
	AutoStartOrdner[2]	:= A_StartupCommon
	AutoStartOrdner[3] 	:= A_AppData "\Microsoft\Windows\Start Menu\Programs\Startup"

	Loop 2	{

		If !FileExist(AutoStartOrdner[A_Index] "\AddendumStarter") 	{

			Exception("Kein Link für Addendum im Autostartordner (" AutoStartOrdner[A_Index] ") vorhanden.")
			IniRead, Autostart, % A_ScriptDir "\Addendum.ini", AutoStartAbfrage, % CompName
			If InStr(Autostart, "ERROR") {

				MsgBox, 4, Addendum für Albis on Windows, % "Möchten Sie Addendum dem Autostart-Ordner hinzufügen? "
				IfMsgBox, Yes
				{
					IniWrite, 1	, % A_ScriptDir "\Addendum.ini", AutoStartAbfrage, % CompName
					FileCreateShortcut	, % A_ScriptDir "\AddendumStarter.ahk"
												, % AutoStartOrdner[A_Index] "\AddendumStarter.lnk"
												, % A_ScriptDir,, % "Startet Addendum für AlbisOnWindows"
												, % A_ScriptDir "\support\Addendum.ico"
				}

				IfMsgBox, No
					IniWrite, 0	, % A_ScriptDir "\Addendum.ini", AutoStartAbfrage, % CompName

			}
			else if (Autostart = 1)
				FileCreateShortcut	, % A_ScriptDir "\AddendumStarter.ahk"
											, % AutoStartOrdner[A_Index] "\AddendumStarter.lnk"
											, % A_ScriptDir
											,, % "Startet Addendum für AlbisOnWindows"
											, % A_ScriptDir "\support\Addendum.ico"

		}
	}
;}

; -------------------------------------------------------------------------------------------------------------------------------------
; Die AutohotkeyH exe in das AppData Verzeichnis kopieren - der lokale Start vermeidet Konflikte mit der Kernel dll -> ntdll.dll
; bei einer neuen AutohotkeyH Version wird die Datei im AppData Verzeichnis ersetzt
; -------------------------------------------------------------------------------------------------------------------------------------;{
	If !FileExist(A_ScriptDir "\include\AHK_H\AutoHotkeyH_U64.exe") {

		MsgBox, 1, Addendum für Albis on Windows, %	"Die Datei  *AutoHotkeyH_U64.exe* ist nicht vorhanden.`n"
																			.	"Addendum benötigt diese Datei zum Ausführen der Skripte.`n"
																			.	"Bitte überprüfen Sie das Verzeichnis:`n"
																			.	"\include\AHKH`n"
		ExitApp

	}

	AHKH_Path1 := A_AppData "\AutohotkeyH"
	AHKH_Path2 := A_AppDataCommon "\AutohotkeyH"

  ; noch kein Verzeichnis vorhanden
	If !Instr(FileExist(AHKH_Path1),  "D") && !InStr(FileExist(AHKH_Path2, "D"))  {
		FileCreateDir	, % AHKH_Path1
		If ErrorLevel
			FileCreateDir	, % AHKH_Path2

			MsgBox, % ErrorLevel ": " AHKH_Path2
	}
	AHKH_Path := Instr(FileExist(AHKH_Path1),  "D") ? AHKH_Path1 : Instr(FileExist(AHKH_Path2),  "D") ? AHKH_Path2 : AHKH_Path1

	FileGetTime, installDirTime	, % A_ScriptDir "\include\AHK_H\AutoHotkeyH_U64.exe"
	FileGetTime, AppDirTime    	, % A_AppData "\AutoHotkeyH\AutoHotkeyH_U64.exe"

  ; noch keine Datei vorhanden oder wurde versehentlich gelöscht
	If !FileExist(AHKH_Path "\AutoHotkeyH_U64.exe") || (installDirTime <> AppDirTime)
		FileCopy, % A_ScriptDir "\include\AHK_H\AutoHotkeyH_U64.exe", % AHKH_Path "\AutoHotkeyH_U64.exe", 1

	If FileExist(AHKH_Path "\AutoHotkeyH_U64.exe")
		AHKH_FilePath := AHKH_Path "\AutoHotkeyH_U64.exe"
	else {
		MsgBox, % "AutoHotkeyH konnte nicht ins Verzeichnis:`n" AHKH_Path "`nkopiert werden."
		ExitApp
	}
	;}

; -------------------------------------------------------------------------------------------------------------------------------------
; Albis starten wenn es nicht gestartet ist
; -------------------------------------------------------------------------------------------------------------------------------------
	If !AlbisExist() 	{
		RegRead, AlbisInstallPath	, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
		RegRead, AlbisLocalPath	, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, LocalPath-1
		IniRead, AlbisDir, % A_ScriptDir "\Addendum.ini", Albis, AlbisWorkDir
		IniRead, AlbisExe, % A_ScriptDir "\Addendum.ini", Albis, AlbisExe

		If FileExist(AlbisInstallPath "\" AlbisExe)
			Run, % AlbisInstallPath "\" AlbisExe, % AlbisLocalPath
	}

; -------------------------------------------------------------------------------------------------------------------------------------
; Pause solange das Addendum-Skript noch läuft, keine zweite Instanz starten!
; -------------------------------------------------------------------------------------------------------------------------------------
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

; -------------------------------------------------------------------------------------------------------------------------------------
; Addendum mit AutohotkeyH 64bit Unicode Version starten
; -------------------------------------------------------------------------------------------------------------------------------------;{
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

; -------------------------------------------------------------------------------------------------------------------------------------
; Funktionen
; -------------------------------------------------------------------------------------------------------------------------------------
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

AlbisExist() {                                                     	;-- schaut nach ob Albis läuft, gibt wahr oder falsch zurück
	;result = 0 wenn kein Albis Prozess läuft , andernfalls die PID(Prozess-ID) des Albis Prozesses
	result:= WMIEnumProcessExist("Albis")
	result+= WMIEnumProcessExist("AlbisCS")
return result
}

WMIEnumProcessExist(ProcessSearched) {      		;-- logische Funktion zur Suche ob ein bestimmter Prozeß existiert

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

FilePathCreate(path) {                                      	;-- erstellt einen Dateipfad falls dieser noch nicht existiert

	If !FilePathExist(path) {
		FileCreateDir, % path
		If ErrorLevel
			return 0
		else
			return 1
	}

return 1
}

FilePathExist(path) {                                          	;-- prüft ob ein Dateipfad vorhanden ist
	If !InStr(FileExist(path "\"), "D")
		return 0
return 1
}

IniReadExt(SectionOrFullFilePath, Key:="", DefaultValue:="", convert:=true) {                                          	;-- eigene IniRead funktion für Addendum

	; beim ersten Aufruf der Funktion !nur! Übergabe des ini Pfades mit dem Parameter SectionOrFullFilePath
	; die Funktion behandelt einen geschriebenen Wert der "ja" oder "nein" ist, als Wahrheitswert, also true oder false
	; UTF-16 in UTF-8 Zeichen-Konvertierung
	; Der Pfad in Addendum.Dir wird einer anderen Variable übergeben. Brauche dann nicht immer ein globales Addendum-Objekt
	; letzte Änderung: 31.01.2021

		static admDir
		static WorkIni

	; Arbeitsini Datei wird erkannt wenn der übergebene Parameter einen Pfad ist
		If RegExMatch(SectionOrFullFilePath, "^[A-Z]\:.*\\")	{
			If !FileExist(SectionOrFullFilePath)	{
				MsgBox,, % "Addendum für AlbisOnWindows", % "Die .ini Datei existiert nicht!`n`n" WorkIni "`n`nDas Skript wird jetzt beendet.", 10
				ExitApp
			}
			WorkIni := SectionOrFullFilePath
			If RegExMatch(WorkIni, "[A-Z]\:.*?AlbisOnWindows", rxDir)
				admDir := rxDir
			else
				admDir := Addendum.Dir
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
		If InStr(OutPutVar, "ERROR")
			If (StrLen(DefaultValue) > 0) { ; Defaultwert vorhanden, dann diesen Schreiben und Zurückgeben
				OutPutVar := DefaultValue
				IniWrite, % DefaultValue, % WorkIni, % SectionOrFullFilePath, % key
				If ErrorLevel
					MsgBox, % "Der Defaultwert <" DefaultValue "> konnte geschrieben werden.`n`n[" WorkIni "]"
			}
			else return "ERROR"
		else if InStr(OutPutVar, "%AddendumDir%")
				return StrReplace(OutPutVar, "%AddendumDir%", admDir)
		else if RegExMatch(OutPutVar, "i)^\s*(ja|nein)\s*$", bool)
				return (bool1= "ja") ? true : false

return Trim(OutPutVar)
}

StrUtf8BytesToText(vUtf8) {                                                                                                       	;-- Umwandeln von Text aus .ini Dateien in UTF-8
	if A_IsUnicode 	{
		VarSetCapacity(vUtf8X, StrPut(vUtf8, "CP0"))
		StrPut(vUtf8, &vUtf8X, "CP0")
		return StrGet(&vUtf8X, "UTF-8")
	} else
		return StrGet(&vUtf8, "UTF-8")
}


; -------------------------------------------------------------------------------------------------------------------------------------
; Include
; -------------------------------------------------------------------------------------------------------------------------------------
#Include %A_ScriptDir%\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\lib\SciTEOutput.ahk


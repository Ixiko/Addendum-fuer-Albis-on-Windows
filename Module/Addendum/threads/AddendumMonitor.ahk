;-----------------------------------------------------------------------------------------------------------------------------------
;------------------------------------------------ ADDENDUM MONITOR ----------------------------------------------------
;-----------------------------------------------------------------------------------------------------------------------------------
													Version:= "0.47" , vom:= "24.09.2020"
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
		#SingleInstance, Force
		;#NoTrayIcon
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

		CompName	:= StrReplace(A_ComputerName, "-")

		hIcon := Base64toHICON()
		If (StrLen(hIcon) > 0)
			Menu, Tray, Icon, % "HICON: " hIcon
		else {
			IconDir := A_ScriptDir "\..\..\..\assets\ModulIcons\AddendumMonitor.ico"
			If FileExist(IconDir)
				Menu, Tray, Icon, % IconDir
			else
				TrayTip, AddendumMonitor, Icon kann nicht geladen werden
		 }

		OnError("FehlerProtokoll")

	;}

;{ Variablen

		global t
		global DoRestart          	:= true
		global q                       	:= Chr(0x22)
		global WaitForScript   	:= ""
		global Addendum      	:= Object()
		global Restarts           	:= Object()
		global WatchedScripts	:= {	"Addendum.ahk"	: {"command": "run"
												, 	"run"              		: (AddendumDir "\include\AHK_H\x64w\AutoHotkeyH_U64 " q AddendumDir "\Module\Addendum\Addendum.ahk" q)
												, 	"tip"                    	: "Addendum-Skript"}}
		global winmgmts       	:= ComObjGet("winmgmts:") 	; Get WMI service object.

	; get data from Addendum.ini
		Addendum.AddendumDir 	:= FileOpen("C:\albiswin.loc\AddendumDir","r", "UTF-8").Read()
		Addendum.AddendumIni  	:= AddendumDir "\Addendum.ini"
		Addendum.compname   	:= StrReplace(A_ComputerName, "-")

	; initializing Ini-Path
		IniReadExt(Addendum.AddendumIni)

	; loading Telegram BotToken and BotChatID
		BotName	:= IniReadExt("Telegram", "Bot1")
		If !InStr(BotName, "Error") {
				Addendum.Telegram := Object()
				BotToken	:= IniReadExt("Telegram", BotName "_Token")
				BotChatID	:= IniReadExt("Telegram", BotName "_ChatID")
				Addendum.Telegram.Push({"BotName":BotName, "Token": BotToken, "ChatID": BotChatID})
		}


	;}

;{ Sink Objekte erstellen \ Kontrolle das Addendum läuft alle 30 Minuten

	; Create sink objects for receiving event noficiations.
		ComObjConnect(CreateSink	:= ComObjCreate("WbemScripting.SWbemSink"), "ProcessCreate_")
		ComObjConnect(DeleteSink	:= ComObjCreate("WbemScripting.SWbemSink"), "ProcessDelete_")

	; Register for process deletion notifications:
		winmgmts.ExecNotificationQueryAsync(DeleteSink
			, "SELECT * FROM __InstanceDeletionEvent"
			. " WITHIN " 3                                                                  	; Set event polling interval, in seconds.
			. " WHERE TargetInstance ISA 'Win32_Process'"
			. " AND TargetInstance.Name LIKE 'Autohotkey%'")

	; Timerfunktion für halbstündliche Skriptkontrolle
		fnADM      	:= Func("RestartAddendum").Bind(true)
		fnADMTime 	:= 15*60*1000
		SetTimer, % fnADM, % fnADMTime

;}


return

^!ö::    	;{ Skriptneustart

	TrayTip, AddendumMonitor, Skript wird neu gestartet, 4
	Sleep 4000
	Reload

return ;}
~^Esc:: 	;{ Tastenkombination um Skriptneustart während der Anzeige des Countdownfenster abzubrechen
	DoRestart := false
	FileAppend, % "Abbruch des Addendumneustart durch Nutzer, Zeit: " TimeStamp() ", Client: " A_ComputerName "`n", % Addendum.AddendumDir "\logs'n'data\OnExit-Protokoll.txt"
return ;}
^+!ö::ExitApp

TimeStamp() {

	FormatTime, time, % A_Now, dd.MM.yyyy HH:mm:ss

return time
}

ProcessDelete_OnObjectReady(obj) {                                             	;-- Called when a process terminates

	; this is prepared for multi restart purposes, but I did'nt finished it yet
	; this can only restart one process per interval

	static Procs:= []
	static Restartscript

    Process := obj.TargetInstance
	If !RegExMatch(Process.CommandLine, "\w+\.ahk(?=\" q ")", Script) || (StrLen(Restartscript) > 0)
		return

	Restartscript := Script
	If InStr(Restartscript, "Addendum.ahk") 	{

			If WinExist("Addendum.ahk ahk_class #32770") 	{

				; if matched Autohotkey Error Message gui popped up - no script restart here! but save info's to error protocol
					WinGetText, wText, % "Addendum.ahk ahk_class #32770"
					If RegExMatch(wText, "Error:\s.*Line#.*--->") {
							RegExMatch(Process.CommandLine, "[A-Z]:[\w\\_\säöüÄÖÜ.]+\.ahk", SkriptPath)
							fehler := Object()
							fehler.Message	:= wText
							fehler.File      	:= SkriptPath
							FehlerProtokoll(fehler, 1)
					}

			}

			If !AHKProcessExist("Addendum.ahk") {

					If ShowRestartProgress("Addendum.ahk", 10)
						RestartAddendum()

			}

	}

	DoRestart := true
	Restartscript :=  ""

}

AHKProcessExist(ProcName, cmd="") {                                              	;-- is only searching for Autohotkey processes

	; use cmd = "PID" to receive the PID of an Autohotkey process

	StrQuery := "Select * from Win32_Process Where Name Like 'Autohotkey%'"
	For process in ComObjGet("winmgmts:\\.\root\CIMV2").ExecQuery(StrQuery)
	{
			RegExMatch(process.CommandLine, "[\w\-\_]+\.ahk", name)
			If InStr(name, ProcName)
				If (StrLen(cmd) > 0)
					return process[cmd]
				else
					return true
	}

return false
}

ShowRestartProgress(scriptname, duration) {                                    	;-- zeigt die Zeit bis Addendum erneut gestartet wird

	; duration = Dauer in Sekunden
	global DoRestart

	DoRestart 	:= true
	Loops   	:= Floor((duration * 1000) / 100)

	Progress, B2 cW202842 cBFFFFFF cTFFFFFF FM9 FS8 zY2 zH10 w250 WM700 WS200, % "...Neustart in 10s",  % "Neustart: " scriptname , ----------, Futura Bk Bt
	prghwnd := WinExist("ahk_class AutoHotkey2", "...Neustart")
	WinSet  	, Transparent, 200, % "ahk_id " prghwnd
	WinGetPos,,, ww, wh	, % "ahk_id " prghwnd
	WinGetPos,,,,  th    	, % "ahk_class Shell_TrayWnd"
	WinMove	, % "ahk_id " prghwnd,, % A_ScreenWidth - ww - 5, % A_ScreenHeight - th - wh - 5
	Loop % Loops {

		Progress % (rest := (100 - A_Index))
		ControlSetText, Static2, % "...Neustart in " Floor(rest/10)+1 "s", % "ahk_id " prghwnd
		Sleep, 100
		If !DoRestart
			break

	}
	Progress, Off

return DoRestart  	; Abbruch per Hotkey-Methode gibt ein false zurück
}

RestartAddendum(TimedCheck:=false) {                                           	;-- startet Addendum

	static AHKH_FilePath := A_AppData "\AutoHotkeyH\AutoHotkeyH_U64.exe"

	If !AHKProcessExist("Addendum.ahk") {

		FormatTime	, TimeIdle	, % A_TimeIdle	, HH:mm:ss

		;~ TrayTip     	, AddendumMonitor, % "Neustart von " A_ScriptName "`ncommandline:`n"
							;~ . AHKH_FilePath " /f " q "....`n"
							;~ . SubStr(Addendum.AddendumDir, -20) "\Module\Addendum\Addendum.ahk" q, 4

		Run   	    	, % AHKH_FilePath " /f " q Addendum.AddendumDir "\Module\Addendum\Addendum.ahk" q

		FileAppend, %                                	TimeStamp()
							. ", "                           	A_ScriptName
							. ", "                           	A_ComputerName
							. ", TimeIdle: "            	TimeIdle
							. ", restart Addendum.ahk"
							. " because of "	(TimedCheck ? "timer":"event")
							. "`n", % Addendum.AddendumDir "\logs'n'data\OnExit-Protokoll.txt"
	}

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
				OutPutVar := StrReplace(OutPutVar, "%AddendumDir%", Addendum.AddendumDir)
		else if RegExMatch(OutPutVar, "i)^\s*(ja|nein)\s*$", bool)
				OutPutVar := (bool1= "ja") ? true : false

return Trim(OutPutVar)
}

Create_AddendumMonitor_Icon(NewHandle := true) {
Static hBitmap := 0
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 2788 << !!A_IsUnicode)
B64 := "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"
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
DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
Return hBitmap
}

Base64toHICON() { ; 16x16 PNG image (236 bytes), requires WIN Vista and later
Local B64 := "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"
local Bin, Blen, nBytes:=2304, hICON:=0

  VarSetCapacity( Bin,nBytes,0 ), BLen := StrLen(B64)
  If DllCall( "Crypt32.dll\CryptStringToBinary", "Str",B64, "UInt",BLen, "UInt",0x1
            , "Ptr",&Bin, "UIntP",nBytes, "Int",0, "Int",0 )
     hICON := DllCall( "CreateIconFromResourceEx", "Ptr",&Bin, "UInt",nBytes, "Int",True
                     , "UInt",0x30000, "Int",16, "Int",16, "UInt",0, "UPtr" )
Return hICON
}

#Include %A_ScriptDir%\..\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\..\lib\SciTEOutput.ahk
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                            	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                            	by Ixiko started in September 2017 - last change 14.02.2021 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; FUNKTIONEN  	|                                                 	|                                                 	|                                                 	|                                             	(23)
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; Umwandeln:		|	GetHex                                    	GetDec
; WinEventhook: 	|	SetWinEventHook                     	UnhookWinEvent
; Prozesse:			|	StdOutToVar                            	IsProcessElevated				     		ScriptIsRunning								ScriptMem
; IPC:					|	Receive_WM_COPYDATA	    	Send_WM_COPYDATA			    	GetAddendumID
; INI: 					|	IniReadExt                                	IniAppend                                  	StrUtf8BytesToText
; MsgBox:				|	KurzePause	                          	Weitermachen
; Dateien:				|	FileIsLocked                             	FilePathCreate						    	FilePathExist                               	isFullFilePath
;								GetAppImagePath						GetAppsInfo
; Sonstiges:			|	IsRemoteSession
;______________________________________________________________________________________________________________________________________________
; UMWANDELN (2)
GetHex(hwnd) {                                                                                                                    	;-- Umwandlung Dezimal nach Hexadezimal
return Format("0x{:X}", hwnd)
}

GetDec(hwnd) {                                                                                                                    	;-- Umwandlung Hexadezimal nach Dezimal
return Format("{:u}", hwnd)
}

;______________________________________________________________________________________________________________________________________________
; WINEVENTHOOK (2)
SetWinEventHook(evMin, evMax, hmodWEvProc, lpfnWEvProc, idProcess, idThread, dwFlags) { 	;-- WinEventHook starten
	return DllCall("SetWinEventHook", "uint", evMin, "uint", evMax, "Uint", hmodWEvProc	, "uint", lpfnWEvProc, "uint", idProcess, "uint", idThread, "uint", dwFlags)
}

UnhookWinEvent(hWinEventHook, HookProcAdr) {                                                                    	;-- WinEventHook beenden
	DllCall( "UnhookWinEvent", "Ptr", hWinEventHook )
	DllCall( "GlobalFree", "Ptr", HookProcAdr ) ; free up allocated memory for RegisterCallback
}

;______________________________________________________________________________________________________________________________________________
; PROZESSE (4)
StdoutToVar(sCmd, sEncoding:="UTF-8", sDir:="", ByRef nExitCode:=0) {                               	;-- cmdline Ausgabe in einen String umleiten

    DllCall( "CreatePipe",           PtrP,hStdOutRd, PtrP,hStdOutWr, Ptr,0, UInt,0 )
    DllCall( "SetHandleInformation", Ptr,hStdOutWr, UInt,1, UInt,1                 )

            VarSetCapacity( pi, (A_PtrSize == 4) ? 16 : 24,  0 )
    siSz := VarSetCapacity( si, (A_PtrSize == 4) ? 68 : 104, 0 )
    NumPut( siSz,         si,  0,                                      		"UInt" )
    NumPut( 0x100,     si,  (A_PtrSize == 4) ? 44 : 60, 	"UInt" )
    NumPut( hStdOutWr, si,  (A_PtrSize == 4) ? 60 : 88, 	"Ptr"  )
    NumPut( hStdOutWr, si,  (A_PtrSize == 4) ? 64 : 96, 	"Ptr"  )

    If (!DllCall( "CreateProcess", "Ptr",0, "Ptr",&sCmd, "Ptr",0, "Ptr",0, "Int",True, "UInt",0x08000000, "Ptr",0, "Ptr",sDir?&sDir:0, "Ptr",&si, "Ptr",&pi ))
        Return ""
      , DllCall( "CloseHandle", Ptr,hStdOutWr )
      , DllCall( "CloseHandle", Ptr,hStdOutRd )

    DllCall( "CloseHandle", Ptr,hStdOutWr ) ; The write pipe must be closed before reading the stdout.
    While ( 1 )  { ; Before reading, we check if the pipe has been written to, so we avoid freezings.

        If (!DllCall( "PeekNamedPipe", Ptr,hStdOutRd, Ptr,0, UInt,0, Ptr,0, UIntP,nTot, Ptr,0 ))
            Break

        If ( !nTot ) { ; If the pipe buffer is empty, sleep and continue checking.
            Sleep, 100
            Continue
        } ; Pipe buffer is not empty, so we can read it.

        VarSetCapacity(sTemp, nTot+1)
        DllCall( "ReadFile", Ptr,hStdOutRd, Ptr,&sTemp, UInt,nTot, PtrP,nSize, Ptr,0 )
        sOutput .= StrGet(&sTemp, nSize, sEncoding)

    }

    ; * SKAN has managed the exit code through SetLastError.
    DllCall( "GetExitCodeProcess", Ptr,NumGet(pi,0), UIntP,nExitCode )
    DllCall( "CloseHandle",        Ptr,NumGet(pi,0)                              )
    DllCall( "CloseHandle",        Ptr,NumGet(pi,A_PtrSize)                   )
    DllCall( "CloseHandle",        Ptr,hStdOutRd                                   )

Return sOutput
}

IsProcessElevated(ProcessID) {                                                                                                	;-- ermittelt ob ein Prozess mit UAC-Virtualisierung läuft
    if !(hProcess := DllCall("OpenProcess", "uint", 0x0400, "int", 0, "uint", ProcessID, "ptr"))
        throw Exception("OpenProcess failed", -1)
    if !(DllCall("advapi32\OpenProcessToken", "ptr", hProcess, "uint", 0x0008, "ptr*", hToken))
        throw Exception("OpenProcessToken failed", -1), DllCall("CloseHandle", "ptr", hProcess)
    if !(DllCall("advapi32\GetTokenInformation", "ptr", hToken, "int", 20, "uint*", IsElevated, "uint", 4, "uint*", size))
        throw Exception("GetTokenInformation failed", -1), DllCall("CloseHandle", "ptr", hToken) && DllCall("CloseHandle", "ptr", hProcess)
    return IsElevated, DllCall("CloseHandle", "ptr", hToken) && DllCall("CloseHandle", "ptr", hProcess)
}

ScriptIsRunning(scriptname) {                                                                                                 	;-- ein bestimmtes Skript wird ausgeführt?

	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
        If RegExMatch(process.commandline, scriptname "\.ahk")
			 return process.ProcessID

return 0
}

ScriptMem() {                                                                                                                        	;-- gibt belegten Speicher frei

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

;______________________________________________________________________________________________________________________________________________
; INTERPROZESSCOMMUNICATION (3)
Receive_WM_COPYDATA(wParam, lParam) {                                                                         	;-- empfängt Nachrichten von anderen Skripten

    StringAddress	:= NumGet(lParam + 2*A_PtrSize)
	fn_MsgWorker	:= Func("MessageWorker").Bind(StrGet(StringAddress))
	SetTimer, % fn_MsgWorker, -10

return
}

Send_WM_COPYDATA(ByRef StringToSend, ScriptID) {                                                            	;-- für die Interskriptkommunikation - keine Netzwerkkommunikation!

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

GetAddendumID() {                                                                                                              	;-- für Interskriptkommunikation
	Prev_DetectHiddenWindows := A_DetectHiddenWindows
    Prev_TitleMatchMode := A_TitleMatchMode
    DetectHiddenWindows On
    SetTitleMatchMode 2
	AddendumID := WinExist("Addendum Message Gui ahk_class AutoHotkeyGUI")
	DetectHiddenWindows 	, % Prev_DetectHiddenWindows  	; Restore original setting for the caller.
    SetTitleMatchMode     	, % Prev_TitleMatchMode           	; Same.
return AddendumID
}

;______________________________________________________________________________________________________________________________________________
; INI (3)
IniReadExt(SectionOrFullFilePath, Key:="", DefaultValue:="", convert:=true) {                             	;-- eigene IniRead funktion für Addendum

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
					TrayTip, % A_ScriptName, % "Der Defaultwert <" DefaultValue "> konnte geschrieben werden.`n`n[" WorkIni "]", 2
			}
			else return "ERROR"
		else if InStr(OutPutVar, "%AddendumDir%")
				return StrReplace(OutPutVar, "%AddendumDir%", admDir)
		else if RegExMatch(OutPutVar, "%\.exe$") && !RegExMatch(OutPutVar, "i)[A-Z]\:\\")
				return GetAppImagePath(OutPutVar)
		else if RegExMatch(OutPutVar, "i)^\s*(ja|nein)\s*$", bool)
				return (bool1= "ja") ? true : false

return Trim(OutPutVar)
}

IniAppend(value, filename, section, key) {                                                                                 	;-- vorhandenen Werten weitere Werte hinzufügen

	IniRead, schreib, % filename, % section, % key
	If Instr(schreib, "Error")
		schreib:=""
	IniWrite, % schreib . value, % filename, % section, % key	;, UTF-8

}

StrUtf8BytesToText(vUtf8) {                                                                                                       	;-- Umwandeln von Text aus .ini Dateien in UTF-8
	if A_IsUnicode 	{
		VarSetCapacity(vUtf8X, StrPut(vUtf8, "CP0"))
		StrPut(vUtf8, &vUtf8X, "CP0")
		return StrGet(&vUtf8X, "UTF-8")
	} else
		return StrGet(&vUtf8, "UTF-8")
}

;______________________________________________________________________________________________________________________________________________
; MESSAGEBOX (2)
KurzePause(Pausenzeit:=5) {                                                                                                   	;-- kurze Pause (Debug Hilfe)

	MsgBox, 0x1024, % "Kurze Pause", % "MACH WEITER !!!", % Pausenzeit
	IfMsgBox, No
		return 0

return 1
}

Weitermachen(Message="", Title="Guck!", TimeOut=0) {                                                          	;-- weiter oder abbrechen (Debug Hilfe)

	msg := StrLen(Message) > 0  ? Message "`n" : "Weiter oder Abbrechen?"
	tout	:= TimeOut ? TimeOut : ""
	MsgBox, 0x1021, % Title, % msg, % tout
	IfMsgBox, Cancel
		return 0

return 1
}
:*R:.wm::If !Weitermachen()`nreturn 99

;______________________________________________________________________________________________________________________________________________
; DATEIEN (6)
FileIsLocked(FullFilePath)  {                                                                                                       	;-- ist die Datei gesperrt?

	file	:= FileOpen(FullFilePath, "rw")
	LE		:= A_LastError
	iO 	:= IsObject(file)
	If iO
		file.Close()

return LE = 32 ? true : false
}

FilePathCreate(path) {                                                                                                              	;-- erstellt einen Dateipfad falls dieser noch nicht existiert

	If !FilePathExist(path) {
		FileCreateDir, % path
		If ErrorLevel
			return 0
		else
			return 1
	}

return 1
}

FilePathExist(path) {                                                                                                                 	;-- prüft ob ein Dateipfad vorhanden ist
	If !InStr(FileExist(path "\"), "D")
		return 0
return 1
}

isFullFilePath(path) {                                                                                                               	;-- prüft Pfadstring auf die Angabe eines Laufwerkes
	If RegExMatch(path, "[A-Z]\:\\")
		return 1
return 0
}

GetAppImagePath(appname) {                                                                                                	;-- Installationspfad eines Programmes

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

GetAppsInfo(infoType) {

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

;______________________________________________________________________________________________________________________________________________
; SONSTIGES (1)
IsRemoteSession() {                                                                                                                 	;-- true oder false wenn eine RemoteSession besteht
	SysGet, SessionRS, 4096
	If (SessionRS <> 0)
		return true
return false
}





; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                            	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                            	by Ixiko started in September 2017 - last change 22.07.2022 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; FUNKTIONEN  	|                                                 	|                                                 	|                                                 	|                                             	(36)
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; Umwandeln:		|	GetHex                                    	GetDec
; WinEventhook: 	|	SetWinEventHook                     	UnhookWinEvent
; Prozesse:			|	StdOutToVar                            	IsProcessElevated				     		ScriptIsRunning								ScriptMem
;                         	|	getProcessName                   	GetProcessNameFromID             	GetProcessProperties
; IPC:					|	Receive_WM_COPYDATA	    	Send_WM_COPYDATA			    	GetAddendumID
; INI: 					|	IniReadExt                                	IniAppend                                  	StrUtf8BytesToText
; MsgBox:				|	KurzePause	                          	Weitermachen
; Dateien:				|	FileIsLocked                             	FilePathCreate						    	FilePathExist                               	isFullFilePath
;							|	FileGetDetail                          	FileGetDetails                             	GetDetails
;							|	GetAppImagePath						GetAppsInfo                              	GetFileSize                                  	GetAlbisPaths
; Sonstiges:			|	IsRemoteSession                     	HasVal                                        	GetDrives
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
; PROZESSE (7)
StdoutToVar(sCmd, sEncoding:="UTF-8", sDir:="", ByRef nExitCode:=0) {                               	;-- cmdline Ausgabe in einen String umleiten

    DllCall( "CreatePipe"					, "PtrP"	,hStdOutRd, "PtrP",hStdOutWr, "Ptr",0, "UInt",0 )
    DllCall( "SetHandleInformation"	, "Ptr"	,hStdOutWr, "UInt"	,1, "UInt",1)

               VarSetCapacity( pi, (A_PtrSize == 4) ? 16 : 24,  0 )
    siSz := VarSetCapacity( si, (A_PtrSize == 4) ? 68 : 104, 0 )
    NumPut( siSz,         si,  0,                                      		"UInt" )
    NumPut( 0x100,     si,  (A_PtrSize == 4) ? 44 : 60, 	"UInt" )
    NumPut( hStdOutWr, si,  (A_PtrSize == 4) ? 60 : 88, 	"Ptr"  )
    NumPut( hStdOutWr, si,  (A_PtrSize == 4) ? 64 : 96, 	"Ptr"  )

    If (!DllCall( "CreateProcess", "Ptr",0, "Ptr",&sCmd, "Ptr",0, "Ptr",0, "Int",True, "UInt",0x08000000, "Ptr",0, "Ptr",sDir?&sDir:0, "Ptr",&si, "Ptr",&pi ))
        Return ""
      , DllCall( "CloseHandle", "Ptr",hStdOutWr )
      , DllCall( "CloseHandle", "Ptr",hStdOutRd )

    DllCall( "CloseHandle", "Ptr",hStdOutWr ) ; The write pipe must be closed before reading the stdout.
    While ( 1 )  { ; Before reading, we check if the pipe has been written to, so we avoid freezings.

        If (!DllCall( "PeekNamedPipe", "Ptr",hStdOutRd, "Ptr",0, "UInt",0, "Ptr",0, "UIntP",nTot, "Ptr",0 ))
            Break

        If ( !nTot ) { ; If the pipe buffer is empty, sleep and continue checking.
            Sleep, 100
            Continue
        } ; Pipe buffer is not empty, so we can read it.

        VarSetCapacity(sTemp, nTot+1)
        DllCall( "ReadFile", "Ptr",hStdOutRd, "Ptr",&sTemp, "UInt",nTot, "PtrP",nSize, "Ptr",0 )
        sOutput .= StrGet(&sTemp, nSize, sEncoding)

    }

    ; * SKAN has managed the exit code through SetLastError.
    DllCall( "GetExitCodeProcess"	, "Ptr",NumGet(pi,0), "UIntP",nExitCode	)
    DllCall( "CloseHandle"       	, "Ptr",NumGet(pi,0)                           	)
    DllCall( "CloseHandle"       	, "Ptr",NumGet(pi,A_PtrSize)               	)
    DllCall( "CloseHandle"       	, "Ptr",hStdOutRd                               	)

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
	DllCall("CloseHandle", "UInt", hProc)

return result
}

getProcessName(PID) { 			                                                           		                 					;-- get running processes with search using comma separated list

		s := 100096  ; 100 KB will surely be HEAPS
		array := []

	;Get the handle of this script with PROCESS_QUERY_INFORMATION (0x0400)
		h := DllCall("OpenProcess", "UInt", 0x0400, "Int", false, "UInt", PID, "Ptr")
	;Open an adjustable access token with this process (TOKEN_ADJUST_PRIVILEGES = 32)
		DllCall("Advapi32.dll\OpenProcessToken", "Ptr", h, "UInt", 32, "PtrP", t)
		VarSetCapacity(ti, 16, 0)  ; structure of privileges
		NumPut(1, ti, 0, "UInt")  ; one entry in the privileges array...
	;Retrieves the locally unique identifier of the debug privilege:
		DllCall("Advapi32.dll\LookupPrivilegeValue", "Ptr", 0, "Str", "SeDebugPrivilege", "Int64P", luid)
		NumPut(luid, ti, 4, "Int64")
		NumPut(2, ti, 12, "UInt")  ; enable this privilege: SE_PRIVILEGE_ENABLED = 2
	;Update the privileges of this process with the new access token:
		r := DllCall("Advapi32.dll\AdjustTokenPrivileges", "Ptr", t, "Int", false, "Ptr", &ti, "UInt", 0, "Ptr", 0, "Ptr", 0)
		DllCall("CloseHandle", "Ptr", t)  ; close this access token handle to save memory
		DllCall("CloseHandle", "Ptr", h)  ; close this process handle to save memory
	;Increase performance by preloading the library
		hModule := DllCall("LoadLibrary", "Str", "Psapi.dll")
	;Open process with: PROCESS_VM_READ (0x0010) | PROCESS_QUERY_INFORMATION (0x0400)
	   h := DllCall("OpenProcess", "UInt", 0x0010 | 0x0400, "Int", false, "UInt", PID, "Ptr")
	   if !h
	      return 0
	   VarSetCapacity(n, s, 0)  ; a buffer that receives the base name of the module:
	   e := DllCall("Psapi.dll\GetModuleBaseName", "Ptr", h, "Ptr", 0, "Str", n, "UInt", A_IsUnicode ? s//2 : s)
	   if !e    ; fall-back method for 64-bit processes when in 32-bit mode:
	      if e := DllCall("Psapi.dll\GetProcessImageFileName", "Ptr", h, "Str", n, "UInt", A_IsUnicode ? s//2 : s)
	         SplitPath n, n
	   DllCall("CloseHandle", "Ptr", h)  ; close process handle to save memory
	  DllCall("FreeLibrary", "Ptr", hModule)  ; unload the library to free memory
	return n
}

GetProcessNameFromID(hwnd) {
	Process:= Object()
	Process:= GetProcessProperties(hwnd)
	return Process.Name
}

GetProcessProperties(hwnd) {

	Process:= Object()
	WinGet PID, PID, % "ahk_id " hWnd
    StrQuery := "SELECT * FROM Win32_Process WHERE ProcessId=" . PID
    Enum := ComObjGet("winmgmts:").ExecQuery(StrQuery)._NewEnum
    If (Enum[Process])
        ExePath := Process.ExecutablePath


Return Process
}

;______________________________________________________________________________________________________________________________________________
; INTERPROZESSCOMMUNICATION (4)
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

GetScriptID(scriptname) {                                                                                                       	;-- nur für Skripte mit eigener Message-Gui
	Prev_DetectHiddenWindows := A_DetectHiddenWindows
    Prev_TitleMatchMode := A_TitleMatchMode
    DetectHiddenWindows On
    SetTitleMatchMode 2
	hwnd := WinExist("Addendum " StrReplace(ScriptName, ".ahk") " ahk_class AutoHotkeyGUI")
	DetectHiddenWindows 	, % Prev_DetectHiddenWindows  	; Restore original setting for the caller.
    SetTitleMatchMode     	, % Prev_TitleMatchMode           	; Same.
return hwnd
}

;______________________________________________________________________________________________________________________________________________
; INI (3)
IniReadExt(SectionOrFullFilePath, Key:="", DefaultValue:="", convert:=true) {                             	;-- eigene IniRead funktion für Addendum

	/* Beschreibung

		-	beim ersten Aufruf der Funktion nur den Pfad und Namen zur ini Datei übergeben!
				workini := IniReadExt("C:\temp\Addendum.ini")

		-	die Funktion behandelt einen geschriebenen Wert welcher ein "ja" oder "nein" ist, als logische Wahrheitswerte (wahr und falsch).
		-	UTF-16-Zeichen (verwendet in .ini Dateien) werden per Default in UTF-8 Zeichen umgewandelt

		letzte Änderung: 18.06.2022

	 */

		static admDir
		static WorkIni

	; Arbeitsini Datei wird erkannt wenn der übergebene Parameter einen Pfad ist
		If RegExMatch(SectionOrFullFilePath, "i)^([A-Z]:\\|\\\\\w{2,})")	{
			If !FileExist(SectionOrFullFilePath)	{
				MsgBox, 0x1024	, % "Addendum für AlbisOnWindows"
											, % "Die .ini Datei existiert nicht!`n`n" WorkIni "`n`nDas Skript wird jetzt beendet.", 10
				ExitApp
			}
			WorkIni	:= SectionOrFullFilePath
			admDir	:= RegExMatch(WorkIni, "i)^([A-Z]:\\|\\\\\w{2,}).*?AlbisOnWindows", rxDir) ? rxDir : Addendum.Dir
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
		IniRead, value, % WorkIni, % SectionOrFullFilePath, % Key
		value := Trim(convert ? StrUtf8BytesToText(value) : value)
		value := StrReplace(value, "%AddendumDir%", admDir)

	; Bearbeiten des Wertes vor Rückgabe
		If (InStr(value, "ERROR") || StrLen(value)=0)
			If DefaultValue  { ; Defaultwert vorhanden, dann diesen Schreiben und Zurückgeben
				value := Trim(DefaultValue)
				IniWrite, % DefaultValue, % WorkIni, % SectionOrFullFilePath, % Key
				If ErrorLevel
					TrayTip, % A_ScriptName, % "Der Defaultwert <" DefaultValue "> konnte geschrieben werden.`n`n[" WorkIni "]", 5
			}
			else return "ERROR"

	; bestimmte Werte werden vor der Rückgabe geändert
		If RegExMatch(value, "i)\.exe\s*$") && !RegExMatch(value, "i)^([A-Z]:\\|\\\\\w{2,})")
			return GetAppImagePath(value)
		else if RegExMatch(value, "i)^\s*(ja|nein)\s*$", bool)
			return bool1= "ja" ? true : false


return value
}

IniInsertSection(iniPath, InsertBeforeSection, newSectionName, backupPath:="") {                     	;-- eine Sektion in eine Ini-Datei einfügen

	If !FileExist(iniPath)
		return 0

	SplitPath, iniPath, iniName
	If RegExMatch(backup, "i)[A-Z]\:\") && InStr(FileExist(Backup), "D") {
		FileCopy, % iniPath, % backupPath "\" A_YYYY "-" A_MM "-" A_DD " " A_Hour ":" A_Min ":" A_Sec "_" iniName, 1
		If ErrorLevel
			return 0
		else
			SciTEOutput(A_ThisFunc ": Ini file backup created.")
	}
	ini := FileOpen(iniPath, "r", "UTF-8").Read()
	If !InStr(ini, "[" newSectionName "]") {
		ini := RegExReplace(ini, "i)(\[.*" InsertBeforeSection ".*?\])", "[" newSectionName "]`n" "`n$1")
		FileOpen(iniPath, "w", "UTF-8").Write(ini)
		If !ErrorLevel
			SciTEOutput(A_ThisFunc ": new section for ini file: " iniName " was added.`n" )
		else
			return 0
	}

return 1
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

	MsgBox, 0x1021	, % Title
								, % (StrLen(Message) > 0  ? Message "`n" : "Weiter oder Abbrechen?")
								, % (TimeOut ? TimeOut : "")
	IfMsgBox, Cancel
		return 0

return 1
}
:*R:.wm::If !Weitermachen()`nreturn 99

;______________________________________________________________________________________________________________________________________________
; DATEIEN (13)
FileIsLocked(FullFilePath)  {                                                                                                       	;-- ist die Datei gesperrt?

	If !FileExist(FullFilePath)
		return false

	file	:= FileOpen(FullFilePath, "rw")
	LE		:= A_LastError
	If IsObject(file)
		file.Close()

return LE = 32 ? true : false
}

FilePathCreate(path) {                                                                                                              	;-- erstellt einen Dateipfad falls dieser noch nicht existiert
	If !FilePathExist(path) {
		FileCreateDir, % path
		return ErrorLevel ? 0 : 1
	}
return 1
}

FilePathExist(path) {                                                                                                                 	;-- prüft ob ein Dateipfad vorhanden ist
	; akzeptiert keine leeren Pfade
	If (StrLen(path)>0)
		If (StrLen(pathExist := FileExist(path)) > 0)
			If InStr(pathExist, "D")
				return true
return false
}

isFullFilePath(path) {                                                                                                               	;-- prüft Pfadstring auf die Angabe eines Laufwerkes
	If RegExMatch(path, "[A-Z]\:\\")
		return true
return false
}

FileGetDetail(FilePath, Index) {                                                                                               	;-- Bestimmte Dateieigenschaft per Index abrufen
   Static MaxDetails := 350
   SplitPath, FilePath, FileName , FileDir
	FileDir 			:=!FileDir ? A_WorkingDir : ""
	Shell      	:= ComObjCreate("Shell.Application")
	Folder	 	:= Shell.NameSpace(FileDir)
	Item  		:= Folder.ParseName(FileName)
   Return Folder.GetDetailsOf(Item, Index)
}

FileGetDetails(FilePath) {                                                                                                        	;-- Array der konkreten Dateieigenschaften erstellen
   Static MaxDetails := 350
   Shell := ComObjCreate("Shell.Application")
   Details := []
   SplitPath, FilePath, FileName , FileDir
   If (FileDir = "")
      FileDir := A_WorkingDir
   Folder := Shell.NameSpace(FileDir)
   Item := Folder.ParseName(FileName)
   Loop, %MaxDetails% {
      If (Value := Folder.GetDetailsOf(Item, A_Index - 1))
         Details[A_Index - 1] := [Folder.GetDetailsOf(0, A_Index - 1), Value]
   }
   Return Details
}

GetDetails() {                                                                                                                        	;-- Array der möglichen Dateieigenschaften erstellen
   Static MaxDetails := 350
   Shell := ComObjCreate("Shell.Application")
   Details := []
   Folder := Shell.NameSpace(A_ScriptDir)
   Loop, %MaxDetails% {
      If (Value := Folder.GetDetailsOf(0, A_Index - 1)) {
         Details[A_Index - 1] := Value
         Details[Value] := A_Index - 1
      }
   }
   Return Details
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

   ;~ appImages := GetAppsInfo({mask: "IMAGE", offset: A_PtrSize*(headers["IMAGE"] - 1) })
   appImages := GetAppsInfo({mask: "INSTALLLOCATION", offset: A_PtrSize*(headers["INSTALLLOCATION"] - 1) })
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

GetFileAssoc(extension){
    VarSetCapacity(numChars, 4)
    DllCall("Shlwapi.dll\AssocQueryStringW", "UInt", 0x0, "UInt"
        , 0x2, "WStr", "." . extension, "Ptr", 0, "Ptr", 0, "Ptr", &numChars)
    numChars:= NumGet(&numChars, 0, "UInt")
    VarSetCapacity(progPath, numChars*2)
    DllCall("Shlwapi.dll\AssocQueryStringW", "UInt", 0x0, "UInt"
        , 0x2, "WStr", "." . extension, "Ptr", 0, "Ptr", &progPath, "Ptr", &numChars)
    return StrGet(&progPath,NumGet(&numChars, 0, "UInt"),"UTF-16")
}

GetFileSize(path, units="") {                                                                                                      	;-- FileGetSize wrapper
	FileGetSize, fSize, % path, % units
return fSize
}

GetFileSemVer(path) {
    FileGetVersion, _t1, % path
    RegExMatch(_t1, "^([0-9]+)\.([0-9]+)\.([0-9]+)((\.([0-9]+))+)?$", _match)
    return Format("{}.{}.{}", _match1,_match2,_match3)
}

GetAlbisPaths() {                                                                                                                   	;-- Albisverzeichnisse

	SetRegView	, % (A_PtrSize = 8 ? 64 : 32)
	RegRead   	, MainPath	, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-MainPath
	RegRead    	, LocalPath 	, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-LocalPath
	RegRead   	, Exe         	, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-Exe

return {"MainPath":MainPath, "LocalPath":LocalPath, "Exe":Exe, "Briefe":MainPath "\Briefe", "db":MainPath "\db", "Vorlagen":MainPath "\tvl"}
}

;______________________________________________________________________________________________________________________________________________
; SONSTIGES (3)
IsRemoteSession() {                                                                                                                 	;-- true oder false wenn eine RemoteSession besteht
	SysGet, SessionRS, 4096
return SessionRS
}

HasVal(haystack, needle) {
    for index, value in haystack
        if (value = needle)
            return index
    if !(IsObject(haystack))
        throw Exception("Bad haystack!", -1, haystack)
    return 0
}

DriveTypeExist(DriveType:="CDROM", ByRef DrvType:="") {                                                       	; gibt true zurück wenn ein bestimmter Laufwerkstyp angeschlossen ist

	; oder gibt false zurück wenn nicht

	; ruft die Funktion GetDrives() auf um alle angeschlossenen Laufwerkstypen zu erhalten
	; stellt dann ein Objekt (DrvTyp) zusammen, das nach Laufwerksbuchstaben geordnet alle Laufwerke des gesuchten Typs enthält
	; dieses Objekt wird über ByRef zurückgegeben

	For drv, DriveData in (drives := GetDrives())
		If (DriveData.Type = DriveType) {
			If !IsObject(DrvType)
				DrvType := Object()
			DrvType[drv] := DriveData
		}

return drvType.Count() ? true : false
}

GetDrives() {                                                                                                                           	;-- ermittelt verfügbare Laufwerke und deren Daten

	static DriveGetCmds := ["Type", "Capacity", "Label", "Filesystem", "Serial", "Status", "StatusCD"]

	Drives := Object()
	DriveGet, DriveList    	, List

	For each, drive in StrSplit(DriveList) {
		Drives[drive] := Object()
		For every, cmd in DriveGetCmds {
			DriveGet, value, % cmd, % drive ":"
			If value
				Drives[drive][cmd] := value
		}
	}

return Drives
}



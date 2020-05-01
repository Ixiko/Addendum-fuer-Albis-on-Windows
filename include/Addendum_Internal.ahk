; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                            	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                                            	by Ixiko started in September 2017 - last change 27.04.2020 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; INTERNE FUNKTIONEN                                                                                                                                                                                                                          	(16)
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) TimeCode                                    	(02) PrintArr                                  	(03) GetHex                                  	(04) GetDec                                  	(05) MCodeU
; (06) SetWinEventHook                         	(07) UnhookWinEvent                    	(08) StdOutToVar                          	(09) dirgetparent                           	(10) IsProcessElevated
; (11) RunAsTask                                   	(12) Json2Obj                               	(13) IniReadExt                              	(14) hk                                          	(15) Errorbox
; (15) RegRead
;1
TimeCode(DaT) {                                                                                                                  	;-- used for protokoll functions - Month & Time (DaT) = 1 - it's clear!

	If DaT = 1
		TC:= A_DD "." A_MM "." A_YYYY "`, "

		TC.= A_Hour ":" A_Min ":" A_Sec "`." A_MSec

return TC
}
;2
PrintArr(Arr, Option := "w800 h500, Object name", GuiNum:= 90		                                    	;-- show values of an array in a listview gui for debugging
, Colums:= "Nr|PatID|", statustext:="Gesamtzahl") {
	static initGui:= []
	Option:= StrSplit(Option, ",")

    for index, obj in Arr {
        if (A_Index = 1) {
            for k, v in obj {
                Columns .= k "|"
                cnt++
            }
			If !(init[GuiNum]) {
					Gui, %GuiNum%: Margin, 0, 0
					Gui, %GuiNum%: Add, ListView, % Option[1], % Columns
					Gui, %GuiNum%: Add, Statusbar
					init[GuiNum]:= 1
			}
        }
        RowNum := A_Index
        Gui, %GuiNum%: default
        LV_Add("")
        for k, v in obj {
            LV_GetText(Header, 0, A_Index)
            if (k <> Header) {
                FoundHeader := False
                loop % LV_GetCount("Column") {
                    LV_GetText(Header, 0, A_Index)
                    if (k <> Header)
                        continue
                    else {
                        FoundHeader := A_Index
                        break
                    }
                }
                if !(FoundHeader) {
                    LV_InsertCol(cnt + 1, "", k)
                    cnt++
                    ColNum := "Col" cnt
                } else
                    ColNum := "Col" FoundHeader
            } else
                ColNum := "Col" A_Index
            LV_Modify(RowNum, ColNum, (IsObject(v) ? "Object()" : v))
        }
    }
    loop % LV_GetCount("Column")
        LV_ModifyCol(A_Index, "AutoHdr")
	SB_SetText("   " LV_GetCount() " " statustext)

    Gui, %GuiNum%: Show,, % Option[2]
	return RowNum
}
;3
GetHex(hwnd) {                                                                                                                    	;-- Umwandlung Dezimal nach Hexadezimal
return Format("0x{:x}", hwnd)
}
;4
GetDec(hwnd) {                                                                                                                    	;-- Umwandlung Hexadezimal nach Dezimal
return Format("{:u}", hwnd)
}
;5
MCodeU(mcode) {                                                                                                                	;-- ist der gleiche Code wie MCode - wegen Vorkommen in anderer Bibliothek geändert
static e := {1:4, 2:1}, c := (A_PtrSize=8) ? "x64" : "x86"
if (!regexmatch(mcode, "^([0-9]+),(" c ":|.*?," c ":)([^,]+)", m))
return
if (!DllCall("crypt32\CryptStringToBinary", "str", m3, "uint", 0, "uint", e[m1], "ptr", 0, "uint*", s, "ptr", 0, "ptr", 0))
return
p := DllCall("GlobalAlloc", "uint", 0, "ptr", s, "ptr")
if (c="x64")
DllCall("VirtualProtect", "ptr", p, "ptr", s, "uint", 0x40, "uint*", op)
if (DllCall("crypt32\CryptStringToBinary", "str", m3, "uint", 0, "uint", e[m1], "ptr", p, "uint*", s, "ptr", 0, "ptr", 0))
return p
DllCall("GlobalFree", "ptr", p)
}
;6
SetWinEventHook(eventMin, eventMax, hmodWinEventProc, lpfnWinEventProc, idProcess, idThread, dwFlags) {
	return DllCall("SetWinEventHook", "uint", eventMin, "uint", eventMax, "Uint", hmodWinEventProc	, "uint", lpfnWinEventProc, "uint", idProcess, "uint", idThread, "uint", dwFlags)
}
;7
UnhookWinEvent(hWinEventHook, HookProcAdr) {
	DllCall( "UnhookWinEvent", "Ptr", hWinEventHook )
	DllCall( "GlobalFree", "Ptr", HookProcAdr ) ; free up allocated memory for RegisterCallback
}
;8
StdoutToVar(sCmd, sEncoding:="UTF-8", sDir:="", ByRef nExitCode:=0) {                               	;-- cmdline Ausgabe in einen String umleiten
    DllCall( "CreatePipe",           PtrP,hStdOutRd, PtrP,hStdOutWr, Ptr,0, UInt,0 )
    DllCall( "SetHandleInformation", Ptr,hStdOutWr, UInt,1, UInt,1                 )

            VarSetCapacity( pi, (A_PtrSize == 4) ? 16 : 24,  0 )
    siSz := VarSetCapacity( si, (A_PtrSize == 4) ? 68 : 104, 0 )
    NumPut( siSz,      si,  0,                                       	"UInt" )
    NumPut( 0x100,     si,  (A_PtrSize == 4) ? 44 : 60, "UInt" )
    NumPut( hStdOutWr, si,  (A_PtrSize == 4) ? 60 : 88, "Ptr"  )
    NumPut( hStdOutWr, si,  (A_PtrSize == 4) ? 64 : 96, "Ptr"  )

    If ( !DllCall( "CreateProcess", Ptr,0, Ptr,&sCmd, Ptr,0, Ptr,0, Int,True, UInt,0x08000000, Ptr,0, Ptr,sDir?&sDir:0, Ptr,&si, Ptr,&pi ) )
        Return ""
      , DllCall( "CloseHandle", Ptr,hStdOutWr )
      , DllCall( "CloseHandle", Ptr,hStdOutRd )

    DllCall( "CloseHandle", Ptr,hStdOutWr ) ; The write pipe must be closed before reading the stdout.
    While ( 1 )
    { ; Before reading, we check if the pipe has been written to, so we avoid freezings.
        If ( !DllCall( "PeekNamedPipe", Ptr,hStdOutRd, Ptr,0, UInt,0, Ptr,0, UIntP,nTot, Ptr,0 ) )
            Break
        If ( !nTot )
        { ; If the pipe buffer is empty, sleep and continue checking.
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
;9
dirgetparent(path,parent:=1) {                                                                                              	;-- returns a string containing parent dir
	path:=RTrim(path,"\")
	while parent>=idx:=A_Index
			Loop % path, 1
				if (idx=parent)
						return A_LoopFileDir (A_LoopFileDir?"\":"")
				else if path:=A_LoopFileDir
						break
}
;10
IsProcessElevated(ProcessID) {                                                                                                	;-- ermittelt ob ein Prozess mit UAC-Virtualisierung läuft
    if !(hProcess := DllCall("OpenProcess", "uint", 0x0400, "int", 0, "uint", ProcessID, "ptr"))
        throw Exception("OpenProcess failed", -1)
    if !(DllCall("advapi32\OpenProcessToken", "ptr", hProcess, "uint", 0x0008, "ptr*", hToken))
        throw Exception("OpenProcessToken failed", -1), DllCall("CloseHandle", "ptr", hProcess)
    if !(DllCall("advapi32\GetTokenInformation", "ptr", hToken, "int", 20, "uint*", IsElevated, "uint", 4, "uint*", size))
        throw Exception("GetTokenInformation failed", -1), DllCall("CloseHandle", "ptr", hToken) && DllCall("CloseHandle", "ptr", hProcess)
    return IsElevated, DllCall("CloseHandle", "ptr", hToken) && DllCall("CloseHandle", "ptr", hProcess)
}
;11
RunAsTask() {                                                                                                                        	;-- automatische UAC Virtualisierung für Skripte

 ;  By SKAN,  http://goo.gl/yG6A1F,  CD:19/Aug/2014 | MD:22/Aug/2014

  Local CmdLine, TaskName, TaskExists, XML, TaskSchd, TaskRoot, RunAsTask
  Local TASK_CREATE := 0x2,  TASK_LOGON_INTERACTIVE_TOKEN := 3

  Try TaskSchd  := ComObjCreate( "Schedule.Service" ),    TaskSchd.Connect()
    , TaskRoot  := TaskSchd.GetFolder( "\" )
  Catch
      Return "", ErrorLevel := 1

  CmdLine       := ( A_IsCompiled ? "" : """"  A_AhkPath """" )  A_Space  ( """" A_ScriptFullpath """"  )
  TaskName      := "[RunAsTask] " A_ScriptName " @" SubStr( "000000000"  DllCall( "NTDLL\RtlComputeCrc32"
                   , "Int",0, "WStr",CmdLine, "UInt",StrLen( CmdLine ) * 2, "UInt" ), -9 )

  Try RunAsTask := TaskRoot.GetTask( TaskName )
  TaskExists    := ! A_LastError

  If ( not A_IsAdmin and TaskExists )      {

    RunAsTask.Run( "" )
    ExitApp

  }

  If ( not A_IsAdmin and not TaskExists )  {

    Run *RunAs %CmdLine%, %A_ScriptDir%, UseErrorLevel
    ExitApp

  }

  If ( A_IsAdmin and not TaskExists )      {

    XML := "
    ( LTrim Join
      <?xml version=""1.0"" ?><Task xmlns=""http://schemas.microsoft.com/windows/2004/02/mit/task""><Regi
      strationInfo /><Triggers /><Principals><Principal id=""Author""><LogonType>InteractiveToken</LogonT
      ype><RunLevel>HighestAvailable</RunLevel></Principal></Principals><Settings><MultipleInstancesPolic
      y>Parallel</MultipleInstancesPolicy><DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries><
      StopIfGoingOnBatteries>false</StopIfGoingOnBatteries><AllowHardTerminate>false</AllowHardTerminate>
      <StartWhenAvailable>false</StartWhenAvailable><RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAva
      ilable><IdleSettings><StopOnIdleEnd>true</StopOnIdleEnd><RestartOnIdle>false</RestartOnIdle></IdleS
      ettings><AllowStartOnDemand>true</AllowStartOnDemand><Enabled>true</Enabled><Hidden>false</Hidden><
      RunOnlyIfIdle>false</RunOnlyIfIdle><DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteApp
      Session><UseUnifiedSchedulingEngine>false</UseUnifiedSchedulingEngine><WakeToRun>false</WakeToRun><
      ExecutionTimeLimit>PT0S</ExecutionTimeLimit></Settings><Actions Context=""Author""><Exec>
      <Command>"   (  A_IsCompiled ? A_ScriptFullpath : A_AhkPath )       "</Command>
      <Arguments>" ( !A_IsCompiled ? """" A_ScriptFullpath  """" : "" )   "</Arguments>
      <WorkingDirectory>" A_ScriptDir "</WorkingDirectory></Exec></Actions></Task>
    )"

    TaskRoot.RegisterTask( TaskName, XML, TASK_CREATE, "", "", TASK_LOGON_INTERACTIVE_TOKEN )

  }

Return TaskName, ErrorLevel := 0
}
;12
Json2Obj( str ) {                                                                                                                    	;-- Uses a two-pass iterative approach to deserialize a json string

	; Copyright © 2013 VxE. All rights reserved.

	quot := """" ; firmcoded specifically for readability. Hardcode for (minor) performance gain
	ws := "`t`n`r " Chr(160) ; whitespace plus NBSP. This gets trimmed from the markup
	obj := {} ; dummy object
	objs := [] ; stack
	keys := [] ; stack
	isarrays := [] ; stack
	literals := [] ; queue
	y := nest := 0

; First pass swaps out literal strings so we can parse the markup easily
	StringGetPos, z, str, %quot% ; initial seek
	while !ErrorLevel
	{
		; Look for the non-literal quote that ends this string. Encode literal backslashes as '\u005C' because the
		; '\u..' entities are decoded last and that prevents literal backslashes from borking normal characters
		StringGetPos, x, str, %quot%,, % z + 1
		while !ErrorLevel
		{
			StringMid, key, str, z + 2, x - z - 1
			StringReplace, key, key, \\, \u005C, A
			If SubStr( key, 0 ) != "\"
				Break
			StringGetPos, x, str, %quot%,, % x + 1
		}
	;	StringReplace, str, str, %quot%%t%%quot%, %quot% ; this might corrupt the string
		str := ( z ? SubStr( str, 1, z ) : "" ) quot SubStr( str, x + 2 ) ; this won't

	; Decode entities
		StringReplace, key, key, \%quot%, %quot%, A
		StringReplace, key, key, \b, % Chr(08), A
		StringReplace, key, key, \t, % A_Tab, A
		StringReplace, key, key, \n, `n, A
		StringReplace, key, key, \f, % Chr(12), A
		StringReplace, key, key, \r, `r, A
		StringReplace, key, key, \/, /, A
		while y := InStr( key, "\u", 0, y + 1 )
			if ( A_IsUnicode || Abs( "0x" SubStr( key, y + 2, 4 ) ) < 0x100 )
				key := ( y = 1 ? "" : SubStr( key, 1, y - 1 ) ) Chr( "0x" SubStr( key, y + 2, 4 ) ) SubStr( key, y + 6 )

		literals.insert(key)

		StringGetPos, z, str, %quot%,, % z + 1 ; seek
	}

	; Second pass parses the markup and builds the object iteratively, swapping placeholders as they are encountered
	key := isarray := 1

	; The outer loop splits the blob into paths at markers where nest level decreases
	Loop Parse, str, % "]}"
	{
		StringReplace, str, A_LoopField, [, [], A ; mark any array open-brackets

		; This inner loop splits the path into segments at markers that signal nest level increases
		Loop Parse, str, % "[{"
		{
			; The first segment might contain members that belong to the previous object
			; Otherwise, push the previous object and key to their stacks and start a new object
			if ( A_Index != 1 )
			{
				objs.insert( obj )
				isarrays.insert( isarray )
				keys.insert( key )
				obj := {}
				isarray := key := Asc( A_LoopField ) = 93
			}

			; arrrrays are made by pirates and they have index keys
			if ( isarray )
			{
				Loop Parse, A_LoopField, `,, % ws "]"
					if ( A_LoopField != "" )
						obj[key++] := A_LoopField = quot ? literals.remove(1) : A_LoopField
			}
			; otherwise, parse the segment as key/value pairs
			else
			{
				Loop Parse, A_LoopField, `,
					Loop Parse, A_LoopField, :, % ws
						if ( A_Index = 1 )
                            key := A_LoopField = quot ? literals.remove(1) : A_LoopField
						else if ( A_Index = 2 && A_LoopField != "" )
                            obj[key] := A_LoopField = quot ? literals.remove(1) : A_LoopField
			}
			nest += A_Index > 1
		}

		If !--nest
			Break

		; Insert the newly closed object into the one on top of the stack, then pop the stack
		pbj := obj
		obj := objs.remove()
		obj[key := keys.remove()] := pbj
		If ( isarray := isarrays.remove() )
			key++

	}

	Return obj
} ; json_toobj( str )
;13
IniReadExt(SectionOrFullFilePath, Key:="", DefaultValue:="") {                                                  	;-- eigene IniRead funktion für Addendum

	; beim ersten Aufruf !nur! Übergabe des ini Pfades mit dem Parameter SectionOrFullFilePath
	; Funktion übersetzt eine aus der ini gelesene "1" oder eine "0" als Wahrheitswerte, also "true" oder "false"

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
		else if RegExMatch(OutPutVar, "^\s*(1|0)[^\d]", bool)
				OutPutVar := (bool = 1) ? true : false

return OutPutVar
}
;14
hk(keyboard:=0, mouse:=0, message:="", timeout:=3) {                                                       	;-- Tastatur- und/oder Mauseingriffe abschalten

	;https://www.autohotkey.com/boards/viewtopic.php?f=6&t=33925
	;!F1::hk(1,1,"Keyboard keys and mouse buttons disabled!`nPress Alt+F2 to enable")   ; Disable all keyboard keys and mouse buttons
	;!F2::hk(0,0,"Keyboard keys and mouse buttons restored!")         ; Enable all keyboard keys and mouse buttons
	;!F3::hk(1,0,"Keyboard keys disabled!`nPress Alt+F2 to enable")   ; Disable all keyboard keys (but not mouse buttons)
	;!F4::hk(0,1,"Mouse buttons disabled!`nPress Alt+F2 to enable")   ; Disable all mouse buttons (but not keyboard keys)

   static AllKeys
   static optKeyboard, optMouse

   if !AllKeys {
      s := "||NumpadEnter|Home|End|PgUp|PgDn|Left|Right|Up|Down|Del|Ins|"
      Loop, 254
         k := GetKeyName(Format("VK{:0X}", A_Index))
       , s .= InStr(s, "|" k "|") ? "" : k "|"
      For k,v in {Control:"Ctrl",Escape:"Esc"}
         AllKeys := StrReplace(s, k, v)
      AllKeys := StrSplit(Trim(AllKeys, "|"), "|")
   }
   ;------------------
   For k,v in AllKeys {
      IsMouseButton := Instr(v, "Wheel") || Instr(v, "Button")
      Hotkey, *%v%, Block_Input, % (keyboard && !IsMouseButton) || (mouse && IsMouseButton) ? "On" : "Off"
   }
   if (StrLen(message) > 0) {
      Progress, B1 M FS12 ZH0, %message%
	  optKeyboard	:= keyboard
	  optMouse    	:= mouse
      SetTimer, HkTimeoutTimer, % -1000*timeout
   }
   else
      Progress, Off

Block_Input:
Return

hkTimeoutTimer:

   Progress, Off
   If (optKeyboard=1) || (optMouse=1)
		hk(0, 0, "Die Sperrung des Nutzereingriffes ist aufgrund des übergebenen Zeitintervalles aufgehoben worden!" )

Return
}
;15
ErrorBox(ErrorString, CallingScript:="", Screenshot:=false) {                                                       	;-- eine Funktion um Daten ins Fehlerlogbuch zu schreiben

	;Fehlerlogbuch V1.0 : Verzeichnis - logs'n'data\ErrorLogs\Errorbox.txt
	;Screenshot kann eine Zahl>0 sein für den jeweiligen Monitor oder "All" dann macht er einen Screenshot von allen Monitoren des Clients
	logpath:= AddendumDir "\logs'n'data\ErrorLogs"

	Zeitstempel:=  TimeCode(1) . " | "
	Computer:= A_ComputerName . " | "
	If (CallingScript="")
			CallingScript:= A_ScriptName
	CallingScript.= " | "

	FileAppend, % Zeitstempel Computer Skript SC ErrorString "`n", % logpath "\Errorbox.txt"

}
;16
RegRead64(path, subkey, valuename:="")	{                                                                           	;-- 64bit RegRead wrapper

	SetRegView 64
	RegRead, value, % path, % subkey, % valuename

return value
}


; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                            	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                                            	by Ixiko started in September 2017 - last change 24.01.2020 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; SONSTIGES                                                                                                                                                                                                                                                	(17)
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) ObjFindValue                         	(02) WaitFileExist                                  	(03) FormatSeconds
; (04) DateDiff                                  	(05) FormatedFileCreationTime             	(06) ConvertGerDateToEng                  	(07) LineDelete                         (08) Send_WM_COPYDATA
; (09) GetAddendumID                     	(10) KeyValueObjectFromList                 (11) sleepx                                            (12) NoScriptStopSleep
; (13) ProcessExist                            	(14) ScriptExist
; (15) SendSuspend                        	(16) IsRemoteSession                          	(17) StrFlip                                            (18) Clip                               	(19) ScriptIsRunning
;1
ObjFindValue(arr, value) {                                                                         	;-- Array in Array Suche

	; findet den Wert in mehreren key:value Objekten, welche in einem indizierten Array enthalten sind, gibt den index key zurück

	For idx, obj in arr
	{
				For key, val in arr[idx]
				{
                    	If InStr(val, value)
                            return idx
				}
	}
return 0
}
;2
WaitFileExist(filename) {	                                                                             	;-- wait until a specific file is written (not closed)
   while NOT FileExist(filename)
   {
      Sleep 500
   }
}
;3
FormatSeconds(Sekunden) {

	ListLines, Off

	Return SubStr("0" . Sekunden // 3600, -1) . ":"
        . SubStr("0" . Mod(Sekunden, 3600) // 60, -1) . ":"
        . SubStr("0" . Mod(Sekunden, 60), -1)
}
;4
DateDiff(fnTimeUnits, fnStartDate, fnEndDate) {                                          	;-- berechnet Tagesdifferenzen zwischen zwei Tagen

	; deutsches Datumsformat dd.MM.yyyy wird automatisch umgerechnet
	; returns the difference between two timestamps in the specified units

	; declare local, global, static variables

		; set default return value
		TimeDifference := 0

		; Convert date from german date format to englisch dateformat (only DD.MM.YYYY to YYYY.DD.MM)
		fnStartDate:= ConvertGerDateToEng(fnStartDate)
		fnEndDate:= ConvertGerDateToEng(fnEndDate)

		; validate parameters
		If fnTimeUnits not in YY,MM,DD,HH,MI,SS
			Throw Exception("fnTimeUnits were not valid")

		If fnStartDate is not date
			Throw Exception("fnStartDate was not a date")

		If fnEndDate is not date
			Throw Exception("fnEndDate was not a date")


		; initialise variables
		DatePadding := "00000101000000"
		StartDate := fnStartDate . SubStr(DatePadding, StrLen(fnStartDate)+1) ; normalise start date to 14 digits
		EndDate   := fnEndDate . SubStr(DatePadding, StrLen(fnEndDate  )+1) ; normalise end   date to 14 digits


		; for day or time, use native function
		If fnTimeUnits in DD,HH,MI,SS
		{

			TimeDifference := EndDate
			TimeUnit := fnTimeUnits = "SS" ? "S"
                     :  fnTimeUnits = "MI" ? "M"
                     :  fnTimeUnits = "HH" ? "H"
                     :  fnTimeUnits = "DD" ? "D"
                     :                       ""

			EnvSub, TimeDifference, %StartDate%, %TimeUnit%
		}

		; for year or month
		FormatTime, StartYear , %StartDate%, yyyy
		FormatTime, StartMonth, %StartDate%, MM
		FormatTime, StartDay  , %StartDate%, dd
		FormatTime, StartTime , %StartDate%, HHmmss

		FormatTime, EndYear , %EndDate%, yyyy
		FormatTime, EndMonth, %EndDate%, MM
		FormatTime, EndDay  , %EndDate%, dd
		FormatTime, EndTime , %EndDate%, HHmmss

		If fnTimeUnits in MM
		{

			StartInMonths := (StartYear*12)+StartMonth
			EndInMonths   := (EndYear  *12)+EndMonth

			TimeDifference := EndInMonths-StartInMonths

			If (StartDate < EndDate)
				If (EndDay < StartDay)
					TimeDifference--

			If (StartDate > EndDate)
				If (EndDay > StartDay)
					TimeDifference++
		}

		If fnTimeUnits in YY
		{
			TimeDifference := EndYear-StartYear

			If (StartDate < EndDate)
				If (EndMonth <= StartMonth)
					If (EndDay < StartDay)
						TimeDifference--

			If (StartDate > EndDate)
				If (EndMonth >= StartMonth)
					If (EndDay > StartDay)
						TimeDifference++
		}


	; return
	Return TimeDifference
}
;5
FormatedFileCreationTime(filepath) {                                                         	;-- formatiert ins deutsche Format 18.05.2019
	;bei leerem creationtime String gibt FormatTime automatisch das Datum von heute aus
	FileGetTime, creationtime, % filepath, C
	FormatTime, creationtime, % creationtime, dd.MM.yyyy
return creationtime
}
;6
ConvertGerDateToEng(dateStr) {                                                                	;-- eigene Funktion konvertiert das deutsche Datumsformat in das übliche englische
return SubStr(dateStr, 7, 4) . SubStr(dateStr, 4, 2) . SubStr(dateStr, 1, 2)
}
;7
LineDelete(V, L, R := "", O := "", ByRef M := "") {	                                          	;-- Deletes a specific line or a range of lines from a variable containing one or more lines of text.

	; DESCRIPTION of function LineDelete() - see AHK.Rare on GitHub (Ixiko)

	T := StrSplit(V, "`n").MaxIndex()
	If(L > 0 && L <= T && (O = "" || O = "B")){
		V := StrReplace(V, "`r`n", "`n"), S := "`n" V "`n"
		P := (O = "B") ? InStr(S, "`n",,, L + 1)
		   : InStr(S, "`n",,, L)
		M := (R <> "" && R > 0 && O = "" ) ? SubStr(S, P + 1, InStr(S, "`n",, P, 2 + (R - L)) - P - 1)
		   : (R <> "" && R < 0 && O = "" ) ? SubStr(S, P + 1, InStr(S, "`n",, P, 3 + (R - L + T)) - P - 1)
		   : (R <> "" && R > 0 && O = "B") ? SubStr(S, P + 1, InStr(S, "`n",, P, R - L) - P - 1)
		   : (R <> "" && R < 0 && O = "B") ? SubStr(S, P + 1, InStr(S, "`n",, P, 1 + (R - L + T)) - P - 1)
		   : SubStr(S, P + 1, InStr(S, "`n",, P, 2) - P - 1)
		X := SubStr(S, 1, P - 1) . SubStr(S, P + StrLen(M) + 1), X := SubStr(X, 2, -1)
	}
	Else If(L < 0 && L >= -T && (O = "" || O = "B")){
		V := StrReplace(V, "`r`n", "`n"), S := "`n" V "`n"
		P := (R <> "" && R < 0 && O = "" ) ? InStr(S, "`n",,, R + T + 1)
		   : (R <> "" && R > 0 && O = "" ) ? InStr(S, "`n",,, R)
		   : (R <> "" && R < 0 && O = "B") ? InStr(S, "`n",,, R + T + 2)
		   : (R <> "" && R > 0 && O = "B") ? InStr(S, "`n",,, R + 1)
		   : InStr(S, "`n",,, L + T + 1)
		M := (R <> "" && R < 0 && O = "" ) ? SubStr(S, P + 1, InStr(S, "`n",, P, 2 + (L - R)) - P - 1)
		   : (R <> "" && R > 0 && O = "" ) ? SubStr(S, P + 1, InStr(S, "`n",, P, 3 + (T - R + L)) - P - 1)
		   : (R <> "" && R < 0 && O = "B") ? SubStr(S, P + 1, InStr(S, "`n",, P, (L - R)) - P - 1)
		   : (R <> "" && R > 0 && O = "B") ? SubStr(S, P + 1, InStr(S, "`n",, P, 1 + (T - R + L)) - P - 1)
		   : SubStr(S, P + 1, InStr(S, "`n",, P, 2) - P - 1)
		X := SubStr(S, 1, P - 1) . SubStr(S, P + StrLen(M) + 1), X := SubStr(X, 2, -1)
	}
	Return X
}
;8
Send_WM_COPYDATA(ByRef StringToSend, ScriptID) {                                	;-- für die Interskriptkommunikation - keine Netzwerkkommunikation!

    static TimeOutTime            	:= 4000
    Prev_DetectHiddenWindows	:= A_DetectHiddenWindows
    Prev_TitleMatchMode         	:= A_TitleMatchMode
    DetectHiddenWindows On
    SetTitleMatchMode 2

    VarSetCapacity(CopyDataStruct, 3*A_PtrSize, 0)
    SizeInBytes := (StrLen(StringToSend) + 1) * (A_IsUnicode ? 2 : 1)
    NumPut(SizeInBytes, CopyDataStruct, A_PtrSize)
    NumPut(&StringToSend, CopyDataStruct, 2*A_PtrSize)
    SendMessage, 0x4a, 0, &CopyDataStruct,, ahk_id %ScriptID%,,,, % TimeOutTime ; 0x4a is WM_COPYDATA.

    DetectHiddenWindows, % Prev_DetectHiddenWindows 	; Restore original setting for the caller.
    SetTitleMatchMode	 ,  % Prev_TitleMatchMode         	; Same.

return ErrorLevel  ; Return SendMessage's reply back to our caller.
}
;9
GetAddendumID() {                                                                                   	;-- für Interskriptkommunikation
	Prev_DetectHiddenWindows := A_DetectHiddenWindows
    Prev_TitleMatchMode := A_TitleMatchMode
    DetectHiddenWindows On
    SetTitleMatchMode 2
	AddendumID:= WinExist("AddendumGui") ; ahk_class AutohotkeyGui")
	DetectHiddenWindows %Prev_DetectHiddenWindows%  ; Restore original setting for the caller.
    SetTitleMatchMode %Prev_TitleMatchMode%         ; Same.
return AddendumID
}
;10
KeyValueObjectFromLists(keyList, valueList, delimiter:="`n"
, IncludeKeys:="", KeyREx:="", IncludeValues:="", ValueREx:="") {                	;-- Funktion um z.B. zwei Listen aus WinGet zusammenzuführen

	keyArr:= valueArr:= []
	merged:= Object()
	mustMatches:=0

	If !(IncludeKeys = "")
			mustMatches+=1
	If !(IncludeValues = "")
			mustMatches+=1

	keyArr		:= StrSplit(keyList	 , delimiter)
	valueArr	:= StrSplit(valueList, delimiter)

	Loop % keyArr.MaxIndex()
	{
				If (KeyREx = "")
					mkey:= keyArr[A_Index]
				else
					RegExMatch(keyArr[A_Index], KeyREx, mkey)

				If (ValueREx = "")
					mval:= valueArr[A_Index]
				else
					RegExMatch(valueArr[A_Index], ValueREx, mval)

				matched:=0
				If IncludeKeys != ""
						If mkey in %IncludeKeys%
                            	matched:= 1
				else
						matched:= 1

				If IncludeValues != ""
						If mval in %IncludeValues%
                            	matched += 1
				else
						matched += 1

				If (matched>mustMatches)
						merged[(keyArr[A_Index])]:= valueArr[A_Index]
	}

return merged
}
;11
sleepx(MaxWait) {                                                                                      	;-- mit b oder c unterbrechbare sleep Funktion

		a:=A_TickCount
		while (A_TickCount - a < MaxWait)
		{
				If GetKeyState("b", "P")
						return "b"
				If GetKeyState("c", "P")
						return "c2"
				sleep, 50                                                                                   	; ein sleep <=30 ist nicht unterbrechbar!
		}

return "c1"
}
;12
NoScriptStopSleep(waitingtime) {                                                               	;-- wie sleepx - nur nicht unterbrechbar
		a:=A_TickCount
		while (A_TickCount - a < waitingtime)
					sleep, 50
}
;13
ProcessExist(ProcName, cmd:="") {                                                                	;-- is searching only for Autohotkey Scripts

	; use cmd = "PID" to receive the PID of an Autohotkey process
	q := Chr(0x22)

	StrQuery := "Select * from Win32_Process Where Name Like 'AutoHotkey%'"
	For Process in winmgmts.ExecQuery(StrQuery)
	{
			RegExMatch(Process.CommandLine, "\w+\.ahk(?=\" q ")", name)
			If InStr(name, ProcName)
				If StrLen(cmd) = 0
					return true
				else
					return Process[cmd]
	}

return false
}
;14
ScriptExist(scriptname) {                                                                             	;-- true oder false wenn ein Skript schon ausgeführt wird

	dtWin:= A_DetectHiddenWindows
	DetectHiddenWindows, On

	WinGet, hwnd, List, ahk_class AutoHotkey

	Loop, %hwnd%
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
;15
SendSuspend(scriptname) {                                                                        	;-- Sendet einen Suspend-Befehl zu einem anderen Skript.

	DetectHiddenWindows, On
	WM_COMMAND := 0x111
	ID_FILE_SUSPEND := 65404
	PostMessage, WM_COMMAND, ID_FILE_SUSPEND,,, %scriptname% ahk_class AutoHotkey

}
;16
IsRemoteSession() {                                                                                   	;-- true oder false wenn eine RemoteSession besteht

	SysGet, SessionRS, 4096
	If SessionRS <> 0
		return true

return false
}
;17
StrFlip(string) {                                                                                          	;-- String reverse

	VarSetCapacity(new, n:=StrLen(string))
   Loop % n
      new .= SubStr(string, n--, 1)

return new
}
;18
Clip(Text:="", Reselect:="") {                                                                        	;-- Clip() - Send and Retrieve Text Using the Clipboard

	; by berban - updated February 18, 2019
	; https://www.autohotkey.com/boards/viewtopic.php?f=6&t=62156

	Static BackUpClip, Stored, LastClip

	If (A_ThisLabel = A_ThisFunc)
	{
			If (Clipboard == LastClip)
				Clipboard := BackUpClip
			BackUpClip := LastClip := Stored := ""
	}
	Else
	{
			If !Stored
			{
				Stored := True
				BackUpClip := ClipboardAll                         	; ClipboardAll must be on its own line
			}
			Else
				SetTimer, %A_ThisFunc%, Off

	; LongCopy gauges the amount of time it takes to empty the clipboard which can predict how long the subsequent clipwait will need
		LongCopy := A_TickCount
		Clipboard := ""
		LongCopy -= A_TickCount

		If (Text = "") {
			SendInput, ^c
			ClipWait, LongCopy ? 0.6 : 0.2, True
		} Else {
			Clipboard := LastClip := Text
			ClipWait, 10
			SendInput, ^v
		}

		SetTimer, %A_ThisFunc%, -700
		Sleep 20                                                            	; Short sleep in case Clip() is followed by more keystrokes such as {Enter}

		If (Text = "")
			Return LastClip := Clipboard
		Else If ReSelect and ((ReSelect = True) or (StrLen(Text) < 3000))
			SendInput, % "{Shift Down}{Left " StrLen(StrReplace(Text, "`r")) "}{Shift Up}"
	}

Return

Clip:
Return Clip()
}
;19
ScriptIsRunning(scriptname) {

	for Prozess in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
        If RegExMatch(Prozess.commandline, scriptname "\.ahk")
			 return Prozess.ProcessID

}







GetStdOut(cmd, init:=false) {

	;Get StdOut by Coco - it's working with AHK v1.1.28 on Windows8
	;https://autohotkey.com/boards/viewtopic.php?t=3941

	static base := {__New:"GetStdOut", __Delete:"GetStdOut"}
	static on_exit := new base(true)

	if IsObject(cmd) {
		if init {
			dhw := A_DetectHiddenWindows
			DetectHiddenWindows On
			Run %ComSpec% /k,, Hide UseErrorLevel, pid
			if (ErrorLevel == "ERROR")
				return false
			while !WinExist("ahk_pid " pid)
				Sleep 10
			DllCall("AttachConsole", "UInt", cmd.PID:=pid)
			DetectHiddenWindows %dhw%
		} else {
			DllCall("FreeConsole")
			Process Close, % cmd.PID
		}
		return init ? cmd : true
	}
	shell := ComObjCreate("WScript.Shell")
	exec := shell.Exec(cmd)
	while !exec.StdOut.AtEndOfStream
		stdout := exec.StdOut.ReadAll()
	return stdout
}

StdOutToVar(cmd) {						                                                            								;-- catches the command line stream

	;https://github.com/cocobelgica/AutoHotkey-Util/blob/master/StdOutToVar.ahk
	DllCall("CreatePipe", "PtrP", hReadPipe, "PtrP", hWritePipe, "Ptr", 0, "UInt", 0)
	DllCall("SetHandleInformation", "Ptr", hWritePipe, "UInt", 1, "UInt", 1)

	VarSetCapacity(PROCESS_INFORMATION, (A_PtrSize == 4 ? 16 : 24), 0)    ; http://goo.gl/dymEhJ
	cbSize := VarSetCapacity(STARTUPINFO, (A_PtrSize == 4 ? 68 : 104), 0) ; http://goo.gl/QiHqq9
	NumPut(cbSize, STARTUPINFO, 0, "UInt")                                ; cbSize
	NumPut(0x100, STARTUPINFO, (A_PtrSize == 4 ? 44 : 60), "UInt")        ; dwFlags
	NumPut(hWritePipe, STARTUPINFO, (A_PtrSize == 4 ? 60 : 88), "Ptr")    ; hStdOutput
	NumPut(hWritePipe, STARTUPINFO, (A_PtrSize == 4 ? 64 : 96), "Ptr")    ; hStdError

	if !DllCall(
	(Join Q C
		"CreateProcess",            					; http://goo.gl/9y0gw
		"Ptr",  0,                   							; lpApplicationName
		"Ptr",  &cmd,                						; lpCommandLine
		"Ptr",  0,                   							; lpProcessAttributes
		"Ptr",  0,                   							; lpThreadAttributes
		"UInt", true,                						; bInheritHandles
		"UInt", 0x08000000,     					; dwCreationFlags
		"Ptr",  0,                   							; lpEnvironment
		"Ptr",  0,                   							; lpCurrentDirectory
		"Ptr",  &STARTUPINFO,        				; lpStartupInfo
		"Ptr",  &PROCESS_INFORMATION 		; lpProcessInformation
	)) {
		DllCall("CloseHandle", "Ptr", hWritePipe)
		DllCall("CloseHandle", "Ptr", hReadPipe)
		return ""
	}

	DllCall("CloseHandle", "Ptr", hWritePipe)
	VarSetCapacity(buffer, 4096, 0)
	while DllCall("ReadFile", "Ptr", hReadPipe, "Ptr", &buffer, "UInt", 4096, "UIntP", dwRead, "Ptr", 0)
		sOutput .= StrGet(&buffer, dwRead, "CP0")

	DllCall("CloseHandle", "Ptr", NumGet(PROCESS_INFORMATION, 0))         ; hProcess
	DllCall("CloseHandle", "Ptr", NumGet(PROCESS_INFORMATION, A_PtrSize)) ; hThread
	DllCall("CloseHandle", "Ptr", hReadPipe)
	return sOutput
}

; NoTrayOrphans() - a bunch of functions to remove tray icons of dead processes.
; Initially that function was there: http://www.autohotkey.com/board/topic/80624-notrayorphans/?p=512781
; Thanks to N. Nazzal a.k.a. Chef: http://www.autohotkey.com/board/user/13176-nazzal/
noTrayOrphansWin10() {
	icons := TrayIcon_GetInfo()
	Loop, % icons.MaxIndex()
	if (!icons[A_Index].process && icons[A_Index].hWnd)
		TrayIcon_Remove(icons[A_Index].hwnd, icons[A_Index].uID)
}

noTrayOrphans() {
	tray_icons := tray_icons()
	For index In tray_icons
	{
		If (index == 0)
			Continue
		If (tray_icons[index, "sProcess"] = "")
			tray_iconRemove(tray_icons[index, "hWnd"], tray_icons[index, "uID"], "", tray_icons[index, "hIcon"])
	}

}
tray_icons() {
	arr := []
	arr[0] := ["sProcess", "Tooltip", "nMsg", "uID", "idx", "idn", "Pid", "hWnd", "sClass", "hIcon"]
	Index := 0
	trayWindows := "Shell_TrayWnd|NotifyIconOverflowWindow"
	Loop, Parse, trayWindows, |
	{
		WinGet, taskbar_pid, PID, ahk_class %A_LoopField%
		hProc := DllCall("OpenProcess", "uInt", 0x38, "Int", 0, "uInt", taskbar_pid)
		pProc := DllCall("VirtualAllocEx", "uInt", hProc, "uInt", 0, "uInt", 32, "uInt", 0x1000, "uInt", 0x4)
		idxTB := tray_getTrayBar()
		SendMessage, 0x0418, 0, 0, ToolbarWindow32%idxTB%, ahk_class %A_LoopField%
		Loop, % ErrorLevel		{
			SendMessage, 0x0417, A_Index - 1, pProc, ToolbarWindow32%idxTB%, ahk_class %A_LoopField%
			VarSetCapacity(btn, 32, 0), VarSetCapacity(nfo, 32, 0)
			DllCall("ReadProcessMemory", "uInt", hProc, "uInt", pProc, "uInt", &btn, "uInt", 32, "uInt", 0)
			iBitmap := NumGet(btn, 0), idn := NumGet(btn, 4), Statyle := NumGet(btn, 8)
			If dwData := NumGet(btn, 12, "uInt")
				iString := NumGet(btn, 16)
			Else
				dwData := NumGet(btn, 16, "Int64"), iString := NumGet(btn, 24, "Int64")
			DllCall("ReadProcessMemory", "uInt", hProc, "uInt", dwData, "uInt", &nfo, "uInt", 32, "uInt", 0)
			If NumGet(btn, 12, "uInt")
				hWnd := NumGet(nfo, 0), uID := NumGet(nfo, 4), nMsg := NumGet(nfo, 8), hIcon := NumGet(nfo,20)
			Else
				hWnd := NumGet(nfo, 0, "Int64"), uID := NumGet(nfo, 8, "uInt"), nMsg := NumGet(nfo,12,"uInt")
			WinGet, pid, PID, ahk_id %hWnd%
			WinGet, sProcess, ProcessName, ahk_id %hWnd%
			WinGetClass, sClass, ahk_id %hWnd%
			VarSetCapacity(sTooltip,128), VarSetCapacity(wTooltip,128*2)
			DllCall("ReadProcessMemory", "uInt", hProc, "uInt", iString, "uInt", &wTooltip, "uInt", 128*2, "uInt", 0)
			DllCall("WideCharToMultiByte", "uInt", 0, "uInt", 0, "Str", wTooltip, "Int", -1, "Str", sTooltip, "Int", 128, "uInt", 0, "uInt", 0)
			idx := A_Index - 1
			Tooltip := A_IsUnicode ? wTooltip : sTooltip
			Index++
			For a, b In arr[0]
				arr[Index, b] := %b%
		}
		DllCall("VirtualFreeEx", "uInt", hProc, "uInt", pProc, "uInt", 0, "uInt", 0x8000)
		DllCall("CloseHandle", "uInt", hProc)
	}
	Return arr
}
tray_iconRemove(hWnd, uID, nMsg = 0, hIcon = 0, nRemove = 0x2) {
	VarSetCapacity(nid, size := 936 + 4 * A_PtrSize)
	NumPut(size, nid, 0, "uInt")
	NumPut(hWnd, nid, A_PtrSize)
	NumPut(uID, nid, A_PtrSize * 2, "uInt")
	NumPut(1|2|4, nid, A_PtrSize * 3, "uInt")
	NumPut(nMsg, nid, A_PtrSize * 4, "uInt")
	NumPut(hIcon, nid, A_PtrSize * 5, "uInt")
	Return DllCall("shell32\Shell_NotifyIconA", "uInt", nRemove, "uInt", &nid)
}
tray_getTrayBar() {
	ControlGet, hParent, hWnd,, TrayNotifyWnd1, ahk_class Shell_TrayWnd
	ControlGet, hChild, hWnd,, ToolbarWindow321, ahk_id %hParent%
	Loop	{
		ControlGet, hWnd, hWnd,, ToolbarWindow32%A_Index%, ahk_class Shell_TrayWnd
		If (hWnd = hChild)
			idxTB := A_Index
		If (hWnd = hChild) || !hWnd
			Break
	}
	Return idxTB
}

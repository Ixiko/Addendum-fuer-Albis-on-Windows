#NoEnv
#include %A_ScriptDir%\..\..\include\Praxomat_Functions.ahk

SetTitleMatchMode, 2
DetectHiddenWindows, On
DetectHiddenText, on

If WinExist("SciTE")
	WinMinimize, SciTE

LvGetText(hListView,r=1,c=1) {
	r-=1
	ntemp:=0
	dwProcessId:=0
	DllCall("GetWindowThreadProcessId", UInt,hListView, UIntP,dwProcessId)
	hProcess := DllCall("OpenProcess", UInt,0x001F0FFF, Int,false, UInt,dwProcessId) ; 0x001F0FFF = PROCESS_ALL_ACCESS
	lpProcessBuf1 := DllCall("VirtualAllocEx", UInt,hProcess, UInt,ntemp, UInt,28, UInt,0x1000, UInt,4)
	   ; 28 = size of lvItem, 0x1000 = MEM_COMMIT, 4 = PAGE_READWRITE
	lpProcessBuf2 := DllCall("VirtualAllocEx", UInt,hProcess, UInt,ntemp, UInt,512, UInt,0x1000, UInt,4) ; text buffer
	VarSetCapacity(t, 511, 1)
	VarSetCapacity(lvItem, 28)
	NumPut(1, lvItem, "uint")  ; mask
	NumPut(0, lvItem, A_PtrSize, "int")  ; iItem 0 means the first row; anyhow, when sending the LVM_GETITEMTEXT message, we can specify a different item
	NumPut(c-1, lvItem, A_PtrSize * 2, "int")  ; iSubItem 0 means the first column
	NumPut(lpProcessBuf2, lvItem, A_PtrSize * 5, "ptr")  ; pszText
	NumPut(512, lvItem, A_PtrSize * 6) ; cchTextMax
	DllCall("WriteProcessMemory", UInt,hProcess, UInt,lpProcessBuf1, UInt,&lvItem, UInt,28, UInt,ntemp)
	DllCall("SendMessage", "uint", hListView, "uint", (A_IsUnicode ? 0x1000 + 0x73 : 0x1000 + 0x2D), "uint", r, "ptr", lpProcessBuf1)
	DllCall("ReadProcessMemory", UInt,hProcess, UInt,lpProcessBuf2, UInt,&t, UInt,512, UInt,ntemp)
	DllCall("VirtualFreeEx", UInt,hProcess, UInt,lpProcessBuf1, UInt,0, UInt,0x8000) ; 0x8000 = MEM_RELEASE
	DllCall("VirtualFreeEx", UInt,hProcess, UInt,lpProcessBuf2, UInt,0, UInt,0x8000)
	DllCall("CloseHandle", UInt,hProcess)
	return t
}

WM_Command(W, L, M, H) { 
   ; --------------------------------------------------------------------------- 
   ; M : Msg = WM_COMMAND 
   ; H : Hwnd = HWND der Gui 
   ; L : lParam = HWND des Controls 
   ; Das Folgende gilt für die hier behandelten Notifications: 
   ; W : wParam = Notification (Bytes 1 und 2) 
   ;              Interne ID des Controls (Bytes 3 und 4) 
   ; --------------------------------------------------------------------------- 
   Global 
   Static EN_SETFOCUS := 0x100 
   Static EN_KILLFOCUS := 0x200 
   Static AED 
   Static NED 
   Local EN 
   Critical 
   If (L = ED1ID) { 
      ; Byte 1 und 2 (Notification) abgreifen 
      EN := W >> 16 
      If (EN = EN_SETFOCUS) { 
         GuiControlGet, AED, , ED1 
         GuiControl, , TX1, % "Edit erhält den Fokus!`nInhalt:`n" AED 
         Return 0 
      } 
      If (EN = EN_KILLFOCUS) { 
         GuiControlGet, NED, , ED1 
         GuiControl, , TX1, % "Edit verliert den Fokus!`nInhalt:`n" NED 
         Return 0 
      } 
   } 
} 

WinActivate, ahk_class OptoAppClass
	
;found:= AlbisSucheInAkte("01732", "up", 1)
;abr:= AlbisAbrechnungsscheinVorhanden("04/17")
;Zeile:= AlbisGetCaveZeile(9, 0, 0)



MsgBox, %Heute%

Sleep, 200
;ControlGet, ctxt, List, , SysListView322, ahk_class OptoAppClass

^LButton::gosub Versuche
^ä::gosub DerAusstieg

return

DerAusstieg:

If WinExist("- SciTE")
	WinMaximize, - SciTE

ExitApp

ControlsObjekte:
;{
		ControlsFirst:= Object()
		ControlsYet:= Object()
		ControlsFirst:= AlbisMuster30ControlList()
		raus:=0
		Loop {
			
			ControlsYet:= AlbisMuster30ControlList()
				
				For key, val in ControlsFirst {
							sl:= key
							cf:= ControlsFirst[sl]
							cy:= ControlsYet[sl]
							FileAppend, %sl%: %cf%-%cy%`n, WinWin.txt	
							If ControlsFirst[sl] <> ControlsYet[sl] {
								raus =1
								break
							}	
					
				}
			
			Sleep, 100
		} until raus = 1

		;For key, val in Controls
		;		Linien.= key . ": " . val . "`n"

		;FileAppend, %Linien%, ControlsGVU.txt
		a:= ControlsFirst[key] 
		b:= ControlsYet[key]

		MsgBox, Du hast %key% geändert und deswegen wurde die Schleife beendet.`nInhalt zuerst %a% danach %b%
;}
return

Versuche:
;{
MouseGetPos, mx, my, mWin, mControl
;ControlGet, handle, Hwnd, , ahk_id %mControl%, ahk_id %mWin%
;MsgBox % LvGetText(handle,1,4)  ; reading text of row1 col4
ControlGet, Liste, List, , %mControl%, ahk_id %mWin%
ControlGetText, txt, %mControl%, ahk_id %mWin%
MsgBox Win Handle: %mWin%`nControl Handle: %mControl%`nInhalt:`n%Liste%`nText: %txt%
;}
return
; not pausing InputBox
; modified by IXIKO 2022

#NoEnv
#Persistent
SetBatchLines, -1


InputBoxEx("Instruction", "content", "InputBoxEx demonstration", "Default", "", "")
SetTimer, InputBoxExCheck, 200
return

InputBoxExCheck:
	If IBEx {
		MsgBox, % IBEX
		ExitApp
	}
return

InputBoxEx(Instruction := "", Content := "", Title := "", Default := "", Control := "", Options := "", Width := 430, x := "", y := "", Owner := "", Icon := "", IconIndex := 0) {

    Static py, p1, p2, c, cy, ch, Input, e, ey, eh, f, ww, hwnd, hwndf
	global ExitCode, IBEX

    Gui New, hWndhWnd LabelInputBoxEx -0xA0000
    Gui % hWnd ": " (Owner ?  "+Owner" . Owner : "")
    Gui % hWnd ": Font"
    Gui % hWnd ": Color"	, White
    Gui % hWnd ": Margin"	, 10, 12
    py := 10

    If (Instruction != "") {
        Gui % hWnd ": Font", s12 c003399, Segoe UI
        Gui % hWnd ": Add", Text, vp1 y10                                                                                                             	, % Instruction
        py := 40
    }

    Gui % hWnd ": Font", s9 cDefault, Segoe UI

    If (Content != "")
        Gui % hWnd ": Add", Link, % "vp2 x10 y" . py . " w" . (Width-20)                                                                      	, % Content

    GuicontrolGet c, Pos, % (Content != "") ? "p2" : "p1"
    py := (Instruction != "" || Content !="") ? (cy + ch + 16) : 22
    Gui % hWnd ": Add", % (Control != "") ? Control : "Edit", % "vIBEX x10 y" py " w" Width-20 "h21 " Options	        	, % Default
	DefaultOld := Default

    GuiControlGet e, Pos, IBEX
    py := ey + eh + 20
    Gui % hWnd ": Add", Text, % "hWndf y" py " -Background +Border" ; Footer

    Gui % hWnd ": Add", Button, % "gInputBoxExOK x" . (Width - 176) . " yp+12 w80 h23 Default"                          	, &OK
    Gui % hWnd ": Add", Button, % "gInputBoxExClose xp+86 yp w80 h23"                                                            	, &Cancel

    Gui % hWnd ": Show", % "w" Width " x" (x?x:"Center") " y" (y?y:"Center")                                                              	, % Title
    Gui % hWnd ": +SysMenu"
    If (Icon != "") {
        hIcon := LoadPicture(Icon, "Icon" IconIndex, ErrorLevel)
        SendMessage WM_SETICON := 0x80, 0, hIcon,, % "ahk_id " hWnd
	}

    WinGetPos,,, ww,, % "ahk_id " hWnd
    Guicontrol MoveDraw, %f%, % "x-1 w" ww " h" 48

    If (Owner) {
        WinSet Disable,, ahk_id %Owner%
    }

    GuiControl Focus, Input
    Gui % hWnd ": Font"
    ;~ WinWaitClose % "ahk_id " hWnd

return

InputBoxExESCAPE: ;{
InputBoxExCLOSE:
InputBoxExOK:

	Critical

	If (Owner)
		WinSet Enable,, % "ahk_id " Owner

	Gui % hWnd ": Submit"
    Gui % hWnd ": Destroy"

	ErrorLevel := IBEXExitCode := (A_ThisLabel == "InputBoxExOK" || A_ThisHotkey = "ENTER") && IBEX != "" ? 0 : 1

Return (A_ThisLabel == "InputBoxExOK" || A_ThisHotkey = "ENTER") ? IBEX : ""  ;}
}

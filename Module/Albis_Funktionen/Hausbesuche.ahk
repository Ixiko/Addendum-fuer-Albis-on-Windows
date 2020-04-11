;###############################################################
;------------------------------------------------------------------------------------------------------------------------------------
;--------------------------------------------- Addendum für AlbisOnWindows -----------------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------
;---------------------------------------- Modul: Formularhelfer für Hausbesuche -------------------------------------------
															Version = vom 30.01.2019
;------------------------------------------------------------------------------------------------------------------------------------

SetTitleMatchMode, 2		;Fast is default
SetTitleMatchMode, Fast		;Fast is default
DetectHiddenWindows, Off	;Off is default
CoordMode, Mouse, Client
CoordMode, Pixel, Client
CoordMode, ToolTip, Client
CoordMode, Menu, Client
SetKeyDelay, -1, -1
SetBatchLines, -1
SetWinDelay, -1
SetControlDelay, -1
SendMode, Input
AutoTrim, On
FileEncoding, UTF-8

#MaxHotkeysPerInterval 99000000
#HotkeyInterval 99000000
#KeyHistory 0

ListLines Off

;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Script Prozess ID feststellen
;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	scriptPID:= DllCall("GetCurrentProcessId")

;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; globale Variablen
;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
	CompName:= StrReplace(A_ComputerName, "-")

	OnExit, DasEnde

	global AlbisWinID            	:= AlbisWinID()
	global AlbisPID                	:= AlbisPID()
	global AddendumDir			:= FileOpen("C:\albiswin.loc\AddendumDir","r").Read()
	global AddendumDBPath 	:= AddendumDir . "\logs'n'data\_DB"
	global hPrg2, BgColor
	col:= Object()
	FocusColor:= "BfD2ff"

	If !AlbisWinID
	{
			MsgBox,, Addendum für Albis on Windows, Bitte erst Albis starten!
			ExitApp
	}

	WinActivate, ahk_class OptoAppClass
	Menu, Tray, Icon, % "hIcon: " Create_Hausbesuche_png(true)
	gosub, InitializeWinEventHooks
;}

;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Einstellungen /Daten einlesen
;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{

	;Hintergrundfarbe ------------------------------------------------------------------------------------------------------------------
	IniRead, BgColor, % AddendumDir . "\Addendum.ini", Addendum, DefaultBgColor3
	If InStr(ErrorLevel, "FAIL") {
		BgColor:=  "8fA2ff"
		IniWrite, % BgColor,  % AddendumDir . "\Addendum.ini", Addendum, DefaultBgColor3
	}
	;Formularanzahl --------------------------------------------------------------------------------------------------------------------
	IniRead, Dk, % AddendumDir . "\Addendum.ini", Addendum, Druckkopien, 3|2|1|1|1
	copies:= StrSplit(Dk, "|")

	;Arbeitsbereich des Monitors bestimmen ----------------------------------------------------------------------------------------
	AlbisMonitorPos:= GetMonitorIndexFromWindow(AlbisWinID)
	SysGet, mon, MonitorWorkArea, % AlbisMonitorPos
	If (monRight > 1920)
			IniKey:= "FormularhelferPos_HighDpi"
	else
			IniKey:= "FormularhelferPos"

	;gespeicherte Fensterposition für Client einlesen -------------------------------------------------------------------------------
	IniRead, xyPosition, % AddendumDir . "\Addendum.ini", % CompName, % Inikey, % "xCenter yCenter "
	RegExMatch(xyPosition, "(?<=x)\d+", fPosX)
	RegExMatch(xyPosition, "(?<=y)\d+", fPosY)

	;Hausbesuchspatienten Datenbank

;}

;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Patientenerfassung
;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
	AlbisT		:= AlbisGetActiveWinTitle()
	Patname	:= AlbisCurrentPatient()
	Birth			:= AlbisPatientGeburtsdatum()
	PatID		:= AlbisAktuellePatID()
	If (Patname="")
			PatLabel:= "Warte auf eine Patientenakte..."
	else
			PatLabel:= "ID: " PatID "          Name: " Patname "          Geburtstag: " Birth
;}

; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Formular Gui
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{

	col1X:= 5, 	topY:= 150, 	fields:=5
	Beschreibung:= "In die Felder mit den Zahlen lassen sich nur einstellige Ziffern zwischen 0-9 eingeben. Am schnellsten geht es indem man eine Ziffer drückt. Der Eingabefokus rückt anschliessend automatisch eine Position von links nach rechts weiter. Wenn man bei 4 Formularen ein 5.mal eine Ziffer drückt beginnt die Eingabe wieder im ganz linken Feld. Drücke ENTER um die ausgewählten Formulare ausdrucken zu lassen. Du kannst natürlich auch sofort die ENTER Taste drücken wenn schon alles richtig eingestellt ist. Warte bitte bis dieses Fenster geschlossen wurde!"
	TasturhilfeStart:= "l 123456789 ik es"
	TasturhilfeDazwischen:= "jl 123456789 ik es"
	TasturhilfeEnde:= "j 123456789 ik es"

	Gui, FhHB: New, hwndhFhHb
	Gui, FhHb: Margin, 0,0
	Gui, FhHb: Color, % BgColor

; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Kassenrezept
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	Gui, FhHb: Add, GroupBox	, % "hWndhGB1        	x" col1X " y" topY " w202 h120 Hide"
	Gui, FhHb: Font, s36 q5 bold cWhite	, Futura Md Bk
	Gui, FhHb: Add, Edit    			, % "hWndhE1        		x" (col1x+8) " y" (topY+38) " w30 vE1 -E0x200 Limit1", % copies[1]
	Gui, FhHb: Add, Picture        	, % "hWndhPic1         	x" (col1x+64)  " y" (topY+38) , % "HBITMAP: " Create_Kassenrezept_png()
	Gui, FhHb: Font, s28 q5 bold cBlack 	, Futura Bk Bt
	Gui, FhHb: Add, Text            	, % "                          	x" (col1x+40) " y" (topY+60) " w20 h23 BackgroundTrans cBlack +0x200", X
	Gui, FhHb: Add, UpDown     	, % "                           	x" (col1x+163 ) " y" (topY+38) " w34 h68 Range1-9 vUD1 -16", % copies[1]
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Privatrezept
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	col:= ControlGetPos(hGB1), col2X:= col["X"] + col["W"] - 5
	Gui, FhHb: Add, GroupBox	, % "hWndhGB2	x" col2X " y" (topY+5) " w202 h" (col.H-5) " Hide"
	Gui, FhHb: Font, s36 q5 bold cWhite, Futura Md Bk
	Gui, FhHb: Add, Edit    			, % "hWndhE2 vE2	x" (col2X+8) " y" (topY+38) " w30 -E0x200 Limit1", % copies[2]
	Gui, FhHb: Add, Picture        	, % "hWndhPicPrivatrezept x" (col2x+64)  " y" (topY+38) , % "HBITMAP: " hBitmapPR:= Create_Privatrezept_png()
	Gui, FhHb: Font, s28 q5 bold cBlack, Futura Bk Bt
	Gui, FhHb: Add, Text            	, % "x" (col2x +   40 ) " y" (topY + 60 ) " w20 h23 BackgroundTrans cBlack +0x200", X
	Gui, FhHb: Add, UpDown     	, % "x" (col2x + 163 ) " y" (topY + 38 ) " w34 h68 Range1-9 vUD2 -16", % copies[2]
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Krankenhaus
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	col:= ControlGetPos(hGB2), col3X:= col["X"] + col["W"] - 5
	Gui, FhHb: Add, GroupBox	, % "hWndhGB3	x" col3X " y" (topY+5) " w202 h179 Hide"
	Gui, FhHb: Font, s36 q5 bold cWhite, Futura Md Bk
	Gui, FhHb: Add, Edit    			, % "hWndhE3 vE3 x" (col3X+8) " y" (topY+38) " w30 -E0x200 Limit1", % copies[3]
	Gui, FhHb: Add, Picture        	, % "hWndhPicKH x" (col3x+64)  " y" (topY+38) , % "HBITMAP: " hBitmapKH:= Create_KHBehandlung_png()
	Gui, FhHb: Font, s28 q5 bold cBlack, Futura Bk Bt
	Gui, FhHb: Add, Text            	, % "x" (col3x +   40 ) " y" (topY + 60 ) " w20 h23 BackgroundTrans cBlack +0x200", X
	Gui, FhHb: Add, UpDown     	, % "x" (col3x + 163 ) " y" (topY + 38 ) " w34 h68 Range1-9 vUD3 -16", % copies[3]
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Krankenbeförderung
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	col:= ControlGetPos(hGB3), col4X:= col["X"] + col["W"] - 5
	Gui, FhHb: Add, GroupBox	, % "hWndhGB4	x" col4X " y" (topY+5) " w243 h179 Hide"
	Gui, FhHb: Font, s36 q5 bold cWhite, Futura Md Bk
	Gui, FhHb: Add, Edit    			, % "hWndhE4 vE4 x" (col4X+8) " y" (topY+38) " w30 -E0x200 Limit1", % copies[4]
	Gui, FhHb: Add, Picture        	, % "hWndhPicKB x" (col4x+64)  " y" (topY+38) , % "HBITMAP: " hBitmapKH:= Create_Krankenbefoerderung_png()
	Gui, FhHb: Font, s28 q5 bold cBlack, Futura Bk Bt
	Gui, FhHb: Add, Text            	, % "x" (col4x +   40 ) " y" (topY + 60 ) " w20 h23 BackgroundTrans cBlack +0x200", X
	Gui, FhHb: Add, UpDown     	, % "vUDKB x" (col4x + 203 ) " y" (topY + 38 ) " w34 h68 Range1-9 vUD4 -16", % copies[4]
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Überweisung
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	col:= ControlGetPos(hPicKB), col5X:= col["X"] + col["W"] + 35
	GuiH_Temp:= col["H"] + col["Y"] + 40
	Gui, FhHb: Add, GroupBox	, % "hWndhGB5	x" col5X " y" (topY+5) " w220 h179 Hide"
	Gui, FhHb: Font, s36 q5 bold cWhite, Futura Md Bk
	Gui, FhHb: Add, Edit    			, % "hWndhE5 vE5 x" (col5X+8) " y" (topY+38) " w30 -E0x200 r1 Limit1", % copies[5]
	CTLCOLORS.Attach(hE5, BgColor)
	Gui, FhHb: Add, Picture        	, % "hWndhPicUB x" (col5x+64)  " y" (topY+38) , % "HBITMAP: " hBitmapKH:= Create_Ueberweisung_png()
	Gui, FhHb: Font, s28 q5 bold cBlack, Futura Bk Bt
	Gui, FhHb: Add, Text            	, % "x" (col5x +   40 ) " y" (topY + 60 ) " w20 h23 BackgroundTrans cBlack +0x200", X
	Gui, FhHb: Add, UpDown     	, % "vUDUB x" (col5x + 202 ) " y" (topY + 38 ) " w34 h68 Range1-9 vUD5 -16", % copies[5]
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Überschrift, Beschreibung, anvisierter Patientenname
; --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	col:= ControlGetPos(hPicUB), GuiW:= col["X"] + col["W"] + 35
;-: Überschrift
	Gui, FhHb: Add, Progress, % "hWndhPrg x-2 y0 w" (GuiW+4) " h60 +C0x172046", 100
	Gui, FhHb: Font, s36 q5 bold cWhite, Futura Md Bk
	Gui, FhHb: Add, Text, % "x0 y5 w" GuiW " h46 +0x200 Center BackgroundTrans" , Formularhelfer Hausbesuche
	Gui, FhHb: Add, Progress, % "hWndhPrg2 x-2 y140 w" (GuiW+4) " h30 +C0x172046", 100
	Gui, FhHb: Font, s16 q5 Normal cWhite, Futura Md Bk
	Gui, FhHb: Add, Text, % "vPLabel x-2 y142 w" (GuiW+4) " h28 Center BackgroundTrans", % PatLabel
;-: Beschreibung
	Gui, FhHb: Font, s12 q5 cBlack Normal, Futura Bk Bt
	Gui, FhHb: Add, Text  			, % "x10 y64 w" (GuiW-10) " Wrap", % Beschreibung
;-::ein Progress noch
	Gui, FhHb: Add, Progress, % "hWndhPrg3 x-2 y" (GuiH_Temp - 10 ) " w" (GuiW+4) " h30 +C0x172046", 100
;-: Checkbox für Patientenbox-Ausdruck
	Gui, FhHb: Font, s18 q5 bold cBlack, Futura Bk Bt
	addRowY:= 70
	Gui, FhHb: Add, Progress	, % "x45 y" (col.y + addRowY - 8 ) " w380 h36 +C0x172046", 100
	Gui, FhHb: Add, Checkbox, % "HWNDhCheckPatAusweis x25 y" (col.y + addRowY) " w20 h20 "
	Gui, FhHb: Add, Text        	, % "x45 y" (col.y + addRowY - 5) " w400 h28 BackgroundTrans cWhite +0x200", % " + Patientenstammblatt drucken"
	Gui, FhHb: Add, Button    	, % "gPStammblatt HWNDhBPSb x" (GuiW//2 -500//2 - 20//2) " y" (GuiH_Temp + 50 - 10) " w500 Center h36 BackgroundTrans +0x200", nur das Patientenstammblatt drucken
	col:= ControlGetPos(hBPSb)
	GuiH:= col["H"] + col["Y"] + 10
;-: Fokus aus dem EditControl nehmen - sieht nicht schön aus
	ControlFocus,, ahk_id %hPrg2%
;-: Hintergrundfarbe der Edit Controls setzen
	CTLCOLORS.Attach(hE1, BgColor)
	CTLCOLORS.Attach(hE2, BgColor)
	CTLCOLORS.Attach(hE3, BgColor)
	CTLCOLORS.Attach(hE4, BgColor)
	CTLCOLORS.Attach(hE5, BgColor)

	If ( (fPosX+GuiW > monRight) or (fPosY+GuiH > monBottom) )
				xyPosition:= "xCenter yCenter "

	Gui FhHb: Show, % xyPosition "w" GuiW " h" GuiH , Addendum für Albis on Windows - Formularhelfer Hausbesuche
	WinSet, Redraw,, ahk_id %hFhHb%
;}

; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Vorbereitung für den Focusview
; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
	global FsCtl:=1					;fokussiertes Control!

	SetTimer, InputFocus, 300
	;SetTimer, WaitInput, -50

	Hotkey, IfWinActive, Formularhelfer Hausbesuche
	Hotkey, Right		, goRight
	Hotkey, Left		, goLeft
	Hotkey, Up		, Increase
	Hotkey, Down	, Decrease
	Hotkey, Esc		, DasEnde
	Hotkey, Enter		, Printf
	Hotkey, IfWinActive

	Loop 9
		OnMessage( 255+A_Index, "DetectKeyButtonPress" ) ; 0x100 to 0x108

Return

FhHbGuiEscape:
FhHbGuiClose:
    ExitApp
;}


;{  Hotkey Labels
goRight: ;{
	lfs		:= FsCtl
	If (FsCtl+1 > fields)
		FsCtl:= 1
	else
		FsCtl++
	flag:= 1
	CTLCOLORS.Change(hE%lfs%, BgColor)
	ControlFocus,, ahk_id %hPrg2%
return
;}
goLeft: ;{
	lfs		:= FsCtl
	If (FsCtl-1 < 1)
		FsCtl:= fields
	else
		FsCtl--
	CTLCOLORS.Change(hE%lfs%, BgColor)
	flag:= 1
	ControlFocus,, ahk_id %hPrg2%
return
;}
Increase: ;{
	GuiControlGet, copZiffer, FhHb:, % "E" FsCtl
	If (copZiffer+1 > 9)
		copZiffer:= 0
	else
		copZiffer++
	GuiControl, FhHb:, % "E" FsCtl, % copZiffer
	ControlFocus,, ahk_id %hPrg2%
return
;}
Decrease: ;{
	GuiControlGet, copZiffer,FhHb:, % "E" FsCtl
	If (copZiffer-1 < 0)
		copZiffer:= 9
	else
		copZiffer--
	GuiControl, FhHb:, % "E" FsCtl, % copZiffer
	ControlFocus,, ahk_id %hPrg2%
return
;}
Printf: ;{

  ;Gui Informationen abrufen
	Gui, FhHb: Submit, NoHide
  ;geöffnete PopUp-Fenster schließen, sonst behindern diese den Ablauf
	AlbisCloseLastActivePopups(AlbisWinID)
	Dk:=""
	Loop, % fields
			Dk.= E%A_Index% . "|"
	Dk:= RTrim(Dk, "|")
	IniWrite, % Dk, % AddendumDir . "\Addendum.ini", Addendum, Druckkopien
	MsgBox, 4, Addendum für Albis on Windows, Diese Anzahl an Formularen ausdrucken?
	If MsgBox, No
			return
	;Aufrufen der Formulare
	WinMinimize, Formularhelfer Hausbesuche
	AlbisDruckeBlankoFormular("KR" E1 " PR" E2 " KH" E3 " KB" E4 " UB" E5, 0)			;wähle 0 für Drucken, 1 = ist der Spooler
	ControlGet, isChecked	, Checked	,,, ahk_id %hCheckPatAusweis%
	If isChecked
			AlbisDruckePatientenAusweis()
	;Abfragen ob Patient in die Liste der Hausbesuche aufgenommen werden soll (Menu: 33003)
	;MsgBox, 4, Addendum für Albis on Windows, Soll dieser Patient in die Haus-`nbesuchsliste aufgenommen werden?
	;IfMsgBox, Yes
	;
	ControlFocus,, ahk_id %hPrg2%
	WinSet, Redraw,, ahk_id %hFhHb%

	ExitApp
return
;}


;}

AlbisDruckePatientenAusweis() {

	static BM_Click	:= 0x00F5
	AlbisCloseLastActivePopups(AlbisWinID)
	WinMinimize, Formularhelfer Hausbesuche
	If !WinExist("Patientenausweis")
			PostMessage, 0x111, 33003,,, ahk_id %AlbisWinID%
	Sleep, 2000
	ControlGet, hPrint, HWND,, Button23, Patientenausweis
	PostMessage, % BM_Click,,,, ahk_id %hPrint%
	WinWaitClose, Patientenausweis,, 10
	If WinExist("Patientenausweis")
		MsgBox,, Addendum für Albis on Windows, Bitte das Patientenausweis-Fenster schließen,`nbevor weiter gemacht werden kann!
	WinActivate, Formularhelfer Hausbesuche
	return

}

DetectKeyButtonPress(wParam, lParam) {							;{

	static ReplaceKeys:= {"NumpadIns": "0", "NumpadEnd": "1", "NumpadDown": "2", "NumpadPgDn": "3", "NumpadLeft": "4", "NumpadClear": "5", "NumpadRight": "6", "NumpadHome":"7", "NumpadUp":"8", "NumpadPgUp":"9"}
	static last_key:=0

	If ((A_TickCount-last_key)<200)
		return

	last_key:= A_TickCount

	If !WinActive("Formularhelfer Hausbesuche")
		return

	SetFormat, Integer, Hex
	keyname:= GetKeyName("SC" SubStr((((lParam>>16) & 0xFF)+0xF000),-2))
	SetFormat, Integer, d

	For key, replace in ReplaceKeys
			If (key=keyname)
					keyname:= StrReplace(keyname, key, replace)

	If InStr("0123456789", keyname)
	{
			GuiControl, FhHb:, % "E" FsCtl, % keyname
			lfs:= FsCtl
			If (FsCtl+1 > 5)
				FsCtl:= 1
			else
				FsCtl++
			flag:= 1
			CTLCOLORS.Change(hE%lfs%, BgColor)
			ControlFocus,, ahk_id %hPrg2%
	}


	return
}
;}

InputFocus: ;{

	If flag
		CTLCOLORS.Change(hE%FsCtl%, FocusColor)
	else
		CTLCOLORS.Change(hE%FsCtl%, BgColor)

	flag:= !flag

return
;}

PStammblatt: ;{

	AlbisDruckePatientenAusweis()
	ControlFocus,, ahk_id %hPrg2%
	WinSet, Redraw,, ahk_id %hFhHb%

return
;}

;{ WinEventhook Labels
InitializeWinEventHooks:        	;{                                                                                          	;falls ich mal speziell nur einen Hook auf Albis setzten will sollte ich eine Funktion oder ein Label bereitstellen

	ExcludeScriptMessages = 1 ; 0 to include
	ExcludeGuiEvents = 1 ; 0 to include
	dwFlags := ( ExcludeScriptMessages = 1 ? 0x1 : 0x0 )
	HookProcAdr := RegisterCallback("WinEventProc", "F")
	hWinEventHook := SetWinEventHook( 0x00008004, 0x00008004, 0, HookProcAdr, AlbisPID, 0, dwFlags )

return
;}

WinEventProc(hHook, event, hwnd, idObject, idChild, eventThread, eventTime) {		    				;im Moment werden nur sich öffnende Albisfenster abgefangen
	Critical
	EHookHwnd:= GetHex(hwnd)
	SetTimer, AlbisFensterTitel_Hook,  -0
return 0
}

AlbisFensterTitel_Hook:         	;{                                                                                            	;Eventhookhandler für popup (child) Fenster in Albis

		Patname	:= AlbisCurrentPatient()
		Birth			:= AlbisPatientGeburtsdatum()
		PatID		:= AlbisAktuellePatID()
		If (PatID!=PatID_old)
		{
				PatID_old:= PatID
				If (Patname="") {
					PatLabel:= "Warte auf eine geöffnete Patientenakte..."
					GuiControl, FhHb:, PLabel, % PatLabel
				} else {
					PatLabel:= "ID: " PatID "          Name: " Patname "          Geburtstag: " Birth
					GuiControl, FhHb:, PLabel, % PatLabel
					WinActivate, ahk_id %hFhHb%
				}
		}

Return
;}
;}

;{ Base64 encoded Bilder

Create_Ueberweisung2_png(NewHandle := False) {
Static hBitmap := 0
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 11244 << !!A_IsUnicode)
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

Create_Ueberweisung_png() {
VarSetCapacity(B64, 11244 << !!A_IsUnicode)
B64 := "iVBORw0KGgoAAAANSUhEUgAAAIcAAABgCAYAAAA3pmEwAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAATpAAAE6QB6tWnywAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAACAASURBVHic7Z15dBzHfec/v+rumcEM7psESZAgRVOkIlEWaeq05MSW5HV0mInlKJEUabPJxl69RN44cVZ5G9txlDxtYuc50dvDyXrtXW8cW17bK/lK1pGTyLpF6yBNUuIFXiBuYDD3dHf99o8eDAAClCiegIjvewNUd1dXV1f9uupXv6vkwIEDtr6+XljEIipQVdLptLr19fXS1KigE9OuwiK1XDhQmNHh4nSzf/9+caOrWe6773e47NIeVq/u4tYPbKnkfjuQiFZ+izgRjvWP8rWv/xOZbIF161Zwxy/9HqqKCxEJvGPtcj72W1sBsMWA4FgCiccRI6gqIoIqVXoRouEHkWpaAESIsgmKRqNQlcbknNOb5nK4K0LENaiCv68fb01n9R3mvKfy/2SqOq1J5jye/xCWdjbzsd/ayqEjwzz54x3VKy5EL7R+fTeVLocghHyS4Og4Nj2BGAfxPBRFgyB6eQXTvRTbNwgiuOtW4+/eB2GIBiFOUz12IosWS4jj4G2+DG951zl/9XB4BMId4BpKT/4UPI/M48+RvOM6iLmI52KzBUxNDDwHmyli01lMSz3iuRFRlQIQ0KKPxD2wipbKmNZ6NFOg/PQu4u97J3Yijx2dwKQSOCvaMI455+97sviXH28ntMrg0AS/+MFrcIyQTCbo7u5g8vOIphWU237+qupJAdyeFXh168FaqI4eJjqehBG4ZB0YAwjeimWgGn09xoAqdiKDxGNITc05fPUpTP+Kw4E0GIP/k/1kAwtjWdQP8N65msJjz+OuW4btGwHH4K3vJtjbh7OyDUoBkogRHhvFjmbx1i3HP9CPqYsTu2oD/p5j5B79PO7KTqSlHrIFEh/YTGLTmvPyzm+MaADoOzbKiy/uRsTwgZs3UZtK0NpcyzVXbZjKOTIyok0NGSQ8Wj2p2RLqbMI01J2Hyp9Z2OERJL4DiXvYckj5+d24PUsIB8YwqQTEPJyORuzQBBhDODSGaamHoo/EPML+EZwV7eA45B99Eu/SlThtDWiuhNPVgk3ncBpr0VIZO57DNNeB62BqE5i68/NBnAjBsXGwFtPWwOs/PYhX9DnUP8z1H9iCiVXGCQRiW3j++RcnR445sLAmzpOCxBwS10ZfhrO0BZnGqJoVrQC4y5qr5xTBXdlWSUPd/bdUzk7DkqapdPdMnmM673L8+ZnNK9XcJ85X4eHmuKvC9c3g9Y7vPkUIRzL4r+xDampYemgE77JuWvsm0LIPsdmkMCdxKBBxnwufy5/ewMGhMuJFnW0LI8R6TvxthEM5bLAkauxSCdOWxknF3/BZ4VAeG3QiCDaXxbTnKT/3GrF1ywgH05i2OghC7FgOZ0kjkkpgR3MQWrAWLQeEIxlqrl8PgL8vwCRbovcIB7HpYUxdHBQ0W0TqEmimhKRiaGCx2QJOwiMcTBO/5mJsrh7T/A5AscPjeBuvwLtMsZkcJhGDmEfsfcDg03O+z9yto0rYN4DNZN+wMRYC7Og4sir6tsStw+2ORo7g4A7C/BAEFlOfBEDzRfBcjOeAr7grN0TEkclic89Q7hvFW9kedWbCi+4pBiCKiXsz7rFDw2j4ChQCyjt6kYYU4csHMO0N2IkCkkqg6QL+C/twOpuQ5jhhXxpTm4iWd6qYVBvuiouj+vZm0TDEf6UX096IJDyCl3vBRowynkAphA3LUBHCfAnyZawZjggvbzHGw44NQuBAzEDJR7DYkg/MJnxXROYY5QQ7kSHc24vUxADB23wZpR8+idu9DDueBjHgOJi2ZuyRfkgmENfBjk9g6mrBD9FyGYnHsIUiTlcnsUsvZvoQOjM9owKncX5mWuJxhKE57oHy7iOErx8FHHAd7MgYpqGW5NarZuRTgCDE33uM8rZ9aKGMJGLRO5ZKSFsDic1rEYnNuEkcIOnhLmvBlkpg67AjE3hXrMF//nVMWyN4DtR6mMZ6THMjsZ72yqg9G5rJIokYmi0gqTjSVIupixPsG8Jd0xERge9Tc13UzkEOMA5ayqPEoFQAv4hIClVFs8OYxo45nyci6qrqLO5CAHdtD+6Gd4Bjoi/FcUn8/PsgVNAQVKLViiNo1xLEcaAmAcXSVEGeC4FFwwBTk2BydpxJiydibk7n/LSZupJUBA0yBId3AWALo8Qv74HLV6N+EBE7ijgmku14Qnhod7SELZUwzXGSN1+OWkVDRYxUCheQ6NgO5QkP7YpkPdkspssFP8QsacL4AZovEw64iAhOdwdSE8MFnCVN4Dk4LfUz3sjmhgkO74zqb/O4S1uxY7mIQArl6MN1HZyVreD7UF9L8JM9MLlK8rPY0SHENdjxQUITzQSaGYR4LGK4Bw4g3uwlt6rKBbVaudCgQPHJndixaMQxdTXYbBEdyUJjCuMabKFM8n2XIYnJ9jmZ1co8h6oSjqfRbA5pbYZyGc0X8ZZ0zJ0fIRzOQ6nC87tgGj1QQeIuNleCMMRpSM1YxUyH3zeO01oXSYU9Ux0HbTkA10GMEI7lUD8SFEp9EpPwsEUfRHHiMcKSD8ZgKvdPf5Kg2LADE4s6KiyUMDVxBLCFElJNFyGRwAhosYzGXIwxlXuK0SitYAuDuKs68ctDiBtDixZ3VSe6wmL7C6DgLElgywFOYvbHs2CJA8Dfth0VxVPFDo0Bira1Iq4zK68A+E24qy4HIDj4U4J9O7CHx7G+D+k8fv8ADb9zBydapYXHxin9y0+xmTxuZyM2X8bbsJJg+wE0CHEvWoqkYvg7juK9cyV21xHCo8OY+jpMdzPlkSxh3xA6UcbpbkWzEQNsRzKYxiSJGy9H8wk0CZJqhPE0GrqocbBpF/EFiSUI+8uYthpsLEF4rA+3uxtrw4jnGB2F5V1ofhwdO4a7vBXCNZiupdEUOfATJF6GZe9GUim0XEKHX4D62e+7YIlDREi899oojaAruqpf8pvcWU2Z2iTaEuLGDLY/jVcfo/Tivmj10ZhEXDeSoK5ZgqDEVrcTW9sBgQU0YkpdB29ZIyqCigE/wFnSitoAp7EWU5vEWdmG2hB1HbxLVqATBexIBuedayi/sIfYlneAIxCPQUEQYzCJJAFpxIsjjgeUoVRAi3mwAWJD7Fg/BGUANJ9BvGjFoaUcWi5OvbJaSEdqjurrFwvYchZxXJC5P4YFSxzADGJ4M8JQNGImgyA6DiM1gJbKSKwGHIMWShCEaDZPeddhTFMd0pDEXbMUISKYOevRmEJFyD/6NDqWxSxtRQt53HXdlLf3Yvb3oZki0loH2w+D56ATeUx3O6a5FtUQ7ZuA1Z2RQKqpCS2VK9OTC1bA95GmNrAhZIYhlsJ4ScLCILbsIyqIiaG+D2Ufiddj0xZUUDcOqWZAoXg4YpideETcxoETSCxOzJC6mzD1by+G1E4UI7kEgGcwTSmmTyFVTTKgBb+yjJ/SN70ZZsoltbKoni2rnF2aVmUbdiAzTUQ6TaV9MmmoatABJBVDUvFKmRI9piGOxFzsULbyHME0JpC3JD5/G0Lq4kgyBq6ptP/xIuuprowIY+rMW5cVn4itnYvMpGpxYzrmmPxPE3OVOf3ciWp6YuJY+JLzGVDAPziEHclEUtFEHEUxMYcwXQBHcFrq0WwexOCtXwuJdW9HFdNsBK+D5medvmBGDoHIZiMRQ7NF7EQW05BErSL5Iuq5hANjSH0S/6Ve3HVrMVLD21IDeTzEzDkYzCk+P177t5Ax/eXcjkboaJydacNyACxC4dvPRyJt5uYO3o7QOWyGTyg+x0Jw5DCSqkE8j6p9oFYMfYwDYRiJ1t9AEQ1EYl0BXDcSw0N0n1UwBvX9SBTtvIVBrFxGEcRzpjFjc3ejHRtHek6ukwUlfv16xHNRLRH2vnSSy+NzD5svRk3umEiuI4KcouWZaSshx8nAVFXm7BExEZ8d7O2Nllb5ItLRAvkiNpOLRNGuC2Ufp6sTOz4BxsHpaK0o16ZQ3r2HYNfeiCASMYxx0VIJzRdxlnZEQqBsjvgNV1elfG8Gf89+gl2vR0u9ZA3uxWvwVq6YM69NJhCGT6pcAdym1GTz4CwtvlH284ZwPE/5yZcxtQmcrib8g8NorkTqQ1fPqSc5Vbxl3UqQjrSux3ekDQK0UMSpqz1jlTsTmFzKEo8R9BcQm4jOaxGn1UFzRUxTHXY0W/kSBdOQigyqbWTTIkZQawHBONFIEhZ87IhTNYeURp9wOI2zpClSQiqEg2mkMYmTiEV6vWlQIOgtIrFEdJQoR8KrYhhZkzmKDowjzXUQ85DAInWJaEWaBU32YMdymLiDLQWA4HY0EGZyOPW1CBBOZDF1tYhAmMkhtUmMCFooYV2D43lYVXRiF07qDCxl3Ya5l1rGdWGeEcZ0CCDaiLvynUBkzxHs/Sn+T/ZD3MVOZBHPw9Sm0FwBjRmM8ZC6GNKQwh4ZRa0Sv2EDXmcj5Eq4y96DxDw0CAn7f0T5n3ciyRi27MN4DppT6FgJ0xIncfO7cOsTM2okidW4K7qj+hx4En97L8GBEdyL2nE6mwkPDRMefhWpS2DTeSSVIPH+zYjv4tQ1oZ6PzQtOx1I0PYD6SXQsh7U+hD624IGfBy+BHQlwG1qxE6PYUhyTSmIVEEXTRUjN7ruzuloJjvThLlt6hss8hrOs84zwAs6yViTuguNiB9KY9nqIe1AOQC2aL6P5EqYhhbuqE0KLBhb/0DAaKKZlWmGqxD9wBWQL2Fw5kqcImNoatFAmPDoEflNkUKQWd8VsBaGzog3vsh7sWAbTUIuprcG7dCVh3whuTwdaDpBJA+9YHOPGUDcSt+N6mEQKqwNQzIIYBANhAEEmMgwvF9FiDiQZGYJnx5DaRk4ktziBmaBi/TJSLp9W49tiCT3NMo6HlkpoqTxDKviGdSj7mIqRk81kCA4dBiAcTyO5CYLXj4HnRC4K/WMEh0ahWMA0piKmNxnHjmXANVC02IkMpqUe09xA6B6BmBuJ4h2l/MxuTH0SnShEPJYT6VqkqQ78gGD/IKY2QXhoGOfuTsLBQcBU6llC8yXKrx7GLKkneO0YkowhNXHCI8PYsTTixKDGw12+jODgEcQG2JFxbKGAiGIzhwkGR3DbW0CEYKAft6MlmuKGx5FYHIIyNjuBTSQwLtj82EyPgmmYk+ewmSLlV+LYTCZarRgTLVaCMDLw8YNI7dxQh83mMfW1uKtXVsW207oSf+9BtFSK7kEhHiO27qKT7txJBAcOYYdHIVkDxuB0tOE0z7EsPZ44Rsdwlx2r2HNMKZ6mjJ+mi8/BZgpIMo44prr+mvxvC+Wo2olIejrjfXWmAOD4Ndx0Q2A7lkOaUjPuV50S05/QTi6waKGEqUvOeIByEpJ1jeyz5Ph7dGYbvCnPISJ4a3vQso8k4pEyB6BcjgjF8yqKoAbEGLRUxulom6soJJnEjqYRz0Hqa9F0Bqezfc68bwgRtLU5sptIJpG62op12Zvc5hjgGDClJREqDaU2Gh0qeRWQumQ1PV24rihSM9POUmc07Ey9ik4jCT3urJlcEelMwnyzNK5B6mqo9u70PPrGaZipfD3+2lw4Ic8hdbU4Z8ASzNTVRjalk6g9NabVaW+F9tbTqIngv+ZjGldHw+xoLxIfINh3DEkkInPHQgkRg1qLSdZgc0Uk4RIeG8HpaMVOZIldtQ4dnyB4fSBa5nsu3oo2wokc4npooYTmSkhzCs2VkZgbecZlc9jxHIlr12PexIp9vqDq8XY2pIF2InNGNLs2k51JYKcIU1+LO+mSqWMQz+Gu6UIcg7/jIE53BzqRQxIeJpVEWuvQUhmvozGyBW2rwzSmCNJZnCXNSMLFFsrYXGSl5e84iHfxcqQ2gdPZSLBvELetHhIeapK4S5qjaXmBoDpyBAeHcZc1V+fa6QPrqUA1YvpM/el3qh3PIHWpUyLeGUQ/bRpQQIez+D85EDGbMSF47TCaL+FuXI3mimixTHB4iPh16yl9/yVMewPhkSFiV6/HP7AbLYeYZS2EvYM4XW1osYSzvJVw/zG85S24y1tPqPFcCKgwpBMEr7+M290aOQ6XQ8J+g8ROzyjXZvOY2rkNZN5SOfk8JnlqSrCowww4QjhQhGBSFV8GU0CHs5gljZG9x1g2YkbjHjqei3xM0vnIryUMsROFqjuhHcmgQRBdU9BcCfV93JUdaL6ISS6MqWM2jmNIlYhBq7JsMQd3BUBwWo9xiJ12GVPlhKd89yTcjukMbCL6TXNndOqmXa+kTfuk0M+dwSuYJcetlKpWYoosWMKYCQMVdXasstRbxCIqMBCNF2F6trHHIi5sVCP7mJhXXXerQuml/TgdjZH62lpwHFAbKaQQNJ1DXIMkE9jxPKYpRTg8Eckz4lFZquAk49ggRDw3CgrjGkwytoDn5AsHlZhgQF28Kk5TwI5kCA8NoekCIJiuaDln03mc5npIxbFHR5D6JFoOsCOZyEfCM9ihDNTEoOBTyhfACk5XI1oO0FyJ2CXdeOu7FiexeY6IOASohDYCMCIk37cxGiHypUgoVFuRzL3FLrX5IliLU5ucZmF2vK/XhYr53QZTElKZ3u0ueBsiNXfDdOHyW4fTUCl8DiP9Cx7BLtDSm+c7T6gQh6CF8jRbQoELxbj2vGH+W6hW7JMUZ1lzxSZ0EYuIUJ1WNDjeX+vEzi6LOH1oxYlpPo8dVeKQSvxQ8Rw0LBHu/+H5rNcFAaczCe78JY8qz2GzRUxDFBpRHBene808p+u3AcJBzoR64WxharViNYrQm4oDDuKsZL4zTAsbCnYUdP4SR5UhFWshfsF4Ry7iJBDpVhTCoj+HDegiLmREuhURTMKryrk0DAgO7licVM4mFJz2MnP7HM4PVM0EI0/rij2HY3GX5CKz2oqybRKqpqKCsUQDz+lZjC1i/mKKbh2p7H4Q0UjmC9/BXbYEmyngre1EsyWwSnBoANNYV8nvEB4awrukGzs2gaRqUJT45T04JwiRtIiFg+pSVuLelPGrQOK6n4kCmXoONl/CaW/ApgskLlsRxesWgZKPu3YJNl/CbatDQ4tpSEbBaRex4DHVi+7kFCGAwbvkUma52FR8K52WqfSJl7unNtVsuPQWvv53f8G/+Y3/yDM//uoplbEQoIDYCU7d/PHso2JDqoTDmWgqcCSKYOeu43zIOX7n4w/SsfQq/t39Hwf34je/YYFCUPBfAp2/xBHZkCoY17xlF8UziWKxyKZNm9i8eTOtra10dXVxww03APDRj360eu3WW2/l29/+NgBf/epX2bRp04zf7t27+fSnP109vummm/jUpz5FuVzmyiuv5KmnngLgnnvu4ZZbbgFg+/btbN68meeee25Wed/85jf5yle+wl133QXAD37wA6655hp6enq4//77CYKAPXv2sGnTJtLpNABf/vKXufvuuwG4/fbb+dznPgfA2NgYmzZtYnR0FIDvff8ptlz763Sv2cqdd3+SoaHxc9PYJ4kpYx+RKPKOmR3991zAWsu2bdvI5XJA1JAvv/wyAHv27KG+vp73vve9vPrqq9xxxx309fUxMDDA0aNHeeCBB6rlNDU1cfDgQay1bN26lb6+Pj7zmc9wySWXkM/nefbZZ9m8eTOPPvoopVKJsbExXnzxRcbHxwnDkG3btvHHf/zHuG7UND09Pfzwhz9k165dhGHIXXfdxc0338yHP/xhHnzwQTZu3MjmzZvZtm0bQSXGaX9/P7t37wZgx44d/OhHP+Lee++tlu/7Ptu3b+e2X/gYW2+/njt+8Wf5i89/jfsf+Cxf+9+fOZfN/oaYUrzVxKcUbzYkPLL/nOpWwkIBADt4lPDIfuzIAFhLeGQ/Wipw5cbL+MQ9v8TBI9fxta99jdLhfdjxEdqaGvn4r3xoqiA/h+YybOhZyX+495cB+N7jj9G7/SUuX3cRrz73NC+tW01NPE5jXR0v/OA7vPL0k1x+8Tuwg30AfOzOXyAem9oe4x/SI6hfonRwD5mJCTau7uY3br2ZrmSMzrYm7MCR6B36DhIW0tj0KFouER7ZD4FPJpPhc5/+Q+7/1/dG+Y4d5D9/7i/ZfMX6KjFs2byBXbt7z3IrvzVUicNmi5j2iuuisTjtA+e0IiYfhViSplFMez/SMA5Go7RX5quPfYsnf/Jj9u07yq/ceSNLN4TIExPs7T3Altv+FQA1NXH+5Yn/gtQUeO6pV/nop/89Y2MZjg0O8nM/vxavMc///F/f5/m9T7Jly8UkEjFe3Pc02/e+yo03bkGaRgC47hdvr86wX//qQ0htBnEDapaNcvddN/G7f/QnfO4L/43f+PXbueXOe9i5qzd6h9YBTEsxyu/5mPZ+cEK23n49f/WlL3HXr19byTfIroM7ufqqSwB4dfteiqUSq1YtIQwtzjyxq5lyagpt1QtbxAP3ZziXDKmbiMzl1KxFvE1YDhKL1SDeJjB1LFnazTXXXsuK7iN881v/l9/bFUfMChoaGvnQh+4BwPO8Sv5Wyr6w7aXDvPLKq3zvu49x+aafI1dcwe8/+F956pnDbLnyvdTU1PD8Cy/w6o6DPPgHDyFu5Mi0desvV6eV+qarELMXJIl4m/jCXz/Kbbd/n7/571/koT/9Mo67jA9+8Lao7s6liNeG8gSxeFNUF4lz5y//Jjt29vEXf/WPUfu6G1FSBBVe9I8e+h/8n2/9EwC5sX8kmXzz6AHnAhXXBMHEnKmQAApItEPTuUI8EcPzPNITeZAYY+MZWlpaKvUwvOc9P8tDDz0EwKpVq/j7f3gCz/Nob+/gE7//4HGlGd797uv54he/yObNm3n4P32W9934fjZevplyuczjj3+XX/3V+0gmk/zJnz5MJpPhnVe8i127oo16Pv67nyAen+Y6IQ4g7H5tP/fffz+PPfYYt9z6QX77t3+bp55+lrvvuReA9ESBtvYYY+MTNDe3VNvQceP84R9+kvvuu69SnseGDT/Dj5/6JwC+8XcP8eOnXuW6n/3I2WreU8I0Calz3j3AN27cyCc/+Ul6e3t55JFH6OnpqV575plnePjhhzl27BiHDh3ioosuore3l8HBQR5++OFqvq1bt1bTruvy+c9/nuuvv57HHnuMW2+9lbVr1/Laa6/xrne9i0QiQT6fZ9WqVTQ1TblF/vmf/3l15Ljqqqktvbq6unjxxRe57777uPHGG6tlLl++nJaWFh544AFuuukmvvKVr/ChD03jg4APf/jDPPTQQ+zcGe289Ju/+W/567/+Anff90e8a/N6vvil7+B57klHVDwXcD7xiU98qiZRQkvjmJpYJcykA85SzrWc44orruB73/sejz76KN3d3TzyyCO0trby7LPPsn//fnbu3Mng4CB33XUXH/nIR+jt7WXHjh3s3Lmz+rv66qsZHx+noaGB97znPXR3d9Pf38/OnTu57bbbOHz4MM3Nzfzar/0asViMHTt2cOWVV3LjjTcyNjbGCy+8wO7du6vldXV10dLSQjqd5s477+Syyy7jb//2b3n88ce58sor+bM/+zNqa2u55JJL+MY3vsF3v/tdNm3axGc/+1lSqRRPPPEEN9xwAz09PXR1dbF3717uvfdeVq5cyRUbm/n6N/4fj33nSXpWdfGlv/kDelad2Rhqbx0CzjKOHu2bCvsU7HsFd3klBINVwv48i8Y+ZxOK055E5p2Z4Cwve0UL0wK7GUHDAqajiWD3Edw1nahGInO3NnaiUhdxilCEcKgI/rnfwVoBSeZxG2czwRFDqkTmgZPrN6uUn98ThS+qj2OHJgh2HsK5aCnuzZefy7pfEBAUwhrclVee82crSnjwH2GO2HtTElI/rE4iYoTkL1wzs5B3b8AOTzB7qpkV826O8yeTfqtlvh2eNVf8wfmDKTmH1WiXHwTwwFs7I6PxwCyfDPKyiNOGgoR7QM9snNYziWoIBmlKVY19QBCpZxY1y3yk7wUKUQjnz7J1LkzVrhzMiIu5iEVUI/v4L+2vRBmuoBp5b/F3Nn4L4Tusis/dZS1TwVtsmfDQkyxOImcTitPmLgzr83A0R3U3c2NwujpYJI6zCUXsOPPfTFBBi/60lZWLuGtZJI6ziQVgJigiKiKR1bg7v7nnRZw7iIia6gaAaqNN+RaxCJjaAFA12gOkesEGhIdfX5xUziIUxWn1I1OReYqKryxIQ3JKzmEsTsfozA06FnFmofM/+slMjzcmj2IQm3RqWsRZgSoEO0Dn5/akMD0mWKZAWCjjxiZPLcZCP6tYAKNydXli6pOY1PwwbF3E/ECVOLTsV4lZZVHNcraxENp3SnirdmoWCcqEh//+/NToAoLTWTcPzQSnUBWfuxcvqwrBxHFwutefT9fZtz8UCPtZANEEhbBvHKchVSEQB3GWs8iQnk0o2GHmdTRBkYjTCPedW/fHRcwNrar1J48n/+pxZ+f6zbyKTi/veCZnrnNTEBF1VVVQcJY2RV72i/qVc4jjA4pD6dlt2JGxaGNlETQIwIYQi0EYQmCjyEmOA4UizoouVASdyKDj6WiTaGvRfBFJJSHuQaEExRJm1XLCfQfBgNQkcVZ04a1eMSeNVMXnCJBKVD3eNPQJel9Z5DnOIlTBbS9V7DkEKBD0voRps0h9EvE8tFBEMlmoq0VcQa0B14vMlAslaI2BN4YYB+rLmKaa6AMPAScJrgsoBAYNPMQdx11XDyU/emTQT9A7hsTn7ugqQ2oaaqrEIY7iLV3c8+3cQXE740ChcixEjKrLTJ+BydBcAJP+QyFTNiF2Wj6Yyey6047NtPwFYG5/mepS1n/tGO7x22EuYkEizBTxXz6A09WKzRYARZIJTCpOOJrDNKWwR8eirdbCgNg7liKJ2XsIV0hI+OewiDjzWEW4iJNH3MPp6UQdwTTVYVrrMa11EPNwljSBYzCdDThLm3A6m6KwocDoWJbnXniNSV6oOnK8/PJe3n/jFeflXRZxZuHEHJyuJpQplnd6mmnnmHY+myuyd18fWyr+bFXi2Lmrl7985FusXr2UIQEBmgAAAGFJREFUD7x/y9mr+SLOGeQE6ePPHesf5dFv/DPpiTxr1ixjkmwiL/vGEBuOV284z2E6FnGOoUwZAYoIxu2JvOwLhQLxeAOwGI56ERWU8oRhiKuq+vLLryyOFYuYgZaWFv3/zi4tA3d2NnYAAAAASUVORK5CYII="

Return Create_Image(B64, 11244)
}

Create_Krankenbefoerderung_png() {
VarSetCapacity(B64, 10532 << !!A_IsUnicode)
B64 := "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"

Return Create_Image(B64, 10532)
}

Create_Kassenrezept_png() {
		VarSetCapacity(B64, 3456 << !!A_IsUnicode)
		B64 := "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"
Return Create_Image(B64, 3456)
		}

Create_BtmRezept_png(NewHandle := False) {
		Static hBitmap := 0
		If (NewHandle)
		   hBitmap := 0
		If (hBitmap)
		   Return hBitmap
		VarSetCapacity(B64, 3756 << !!A_IsUnicode)
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

Create_PrivatRezept_png() {
		VarSetCapacity(B64, 3096 << !!A_IsUnicode)
		B64 := "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"
Return Create_Image(B64, 3096)
		}

Create_KHBehandlung_png() {
VarSetCapacity(B64, 5112 << !!A_IsUnicode)
B64 := "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"
Return Create_Image(B64, 5112)
}

Create_Image(EncodedImage, Size) {
	VarSetCapacity(B64, Size << !!A_IsUnicode)
	B64:=EncodedImage
	If !DllCall("Crypt32.dll\CryptStringToBinaryW", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
	   Return False
	VarSetCapacity(Dec, DecLen, 0)
	If !DllCall("Crypt32.dll\CryptStringToBinaryW", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
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

Create_Hausbesuche_png(NewHandle := False) {
Static hBitmap := 0
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 2916 << !!A_IsUnicode)
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
DllCall("Gdiplus.dll\GdipCreateHICONFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
Return hBitmap
}

ControlGetPos(ControlID:="") {                	;-- ControlGetPos als Funktion - handle
	ControlGetPos, X, Y, W, H,, ahk_id %ControlID%
	Return {"X":(X), "Y":(Y), "W":(W), "H":(H)}
}


;}

DasEnde:	;{

	OnExit

	Win:= Object()


	Win:= GetWindowInfo(hFhHb)
	IniWrite, % "x" Win.WindowX " y" Win.WindowY " ", % AddendumDir . "\Addendum.ini", % CompName, % Inikey

	Loop 9
		OnMessage( 255+A_Index, "" ) ; 0x100 to 0x108

	if hWinEventHook
		UnhookWinEvent(hWinEventHook, HookProcAdr)

ExitApp
;}

;{ Include
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Functions.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#Include %A_ScriptDir%\..\..\include\Gui\PraxTT.ahk
#Include %A_ScriptDir%\..\..\lib\class_CtlColors.ahk
;}






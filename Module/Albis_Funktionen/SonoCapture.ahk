;----------------------------------------------- Modul Sono Capture V0.65 -------------------------------------------------
;------------------------------------------- Addendum für AlbisOnWindows ------------------------------------------------
;------------------------------------- written by Ixiko -this version is from 14.12.2018 -----------------------------------
;--------------------------- please report errors and suggestions to me: Ixiko@mailbox.org ---------------------------
;------------------------ use subject: "Addendum" so that you don't end up in the spam folder ------------------------
;--------------------------------- GNU Lizenz - can be found in main directory  - 2017 ---------------------------------
;-----------------------------------------------------------------------------------------------------------------------------------
;{ thx to
; http://ej.bantz.com/video/
; http://ej.bantz.com/video/detail/
; http://www.dotnet4all.com/dotnet-code/2004/12/video-capture.html
;AND ALL USERS FROM FROM AUTOHOTKEY FORUM www.autohotkey.com
;burque505 - for unicode version of Avicap-wrapper - https://autohotkey.com/boards/viewtopic.php?f=6&t=59041&hilit=capture
;}

;# TODO: #done# make it possible to edit the driver and let the user choose his preview app he's using with Albis
;CHANGES DONE:
;| **14.12.2018** | **F~** | übersichtlicheres Skriptlayout


;{1. Scripteinstellungen
	#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
	#SingleInstance force
	SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
	SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
	SetTitleMatchMode, 2
	DetectHiddenWindows, Off
	DetectHiddenText, Off
	SetBatchLines -1            ; Script nicht ausbremsem (Default: "sleep, 10" alle 10 ms)
	SetControlDelay, -1         ; Wartezeit beim Zugriff auf langsame Controls abschalten
	SetWinDelay, -1             ; Verzögerung bei allen Fensteroperationen abschalten
	CoordMode, Mouse, Screen
	CoordMode, Pixel, Screen
	CoordMode, Tooltip, Screen
	FileEncoding, UTF-8

	OnExit, Ende

	ModulShort:= "A6"
;}

;{2. variable declarations

	;------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; client name
	;------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		CompName:= StrReplace(A_ComputerName, "-")
		len:= Floor((20-StrLen(CompName))/2)
		TTipCN:= CompName . SubStr("                   ", 1, len)
	;}

	;------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; menu
	;------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		hIBitmap:= Create_SonoCapture_ico(true)
		Menu Tray, Icon, hIcon:  %hIBitmap%
		Menu, Tray, NoStandard
		Menu, Tray, Tip, % "Addendum für Albis on Windows`nSkript: SonoCapture.ahk`nClient: " TTipCN "`nPID: " GetHex(scriptPID)
		Menu, Tray, Add, Variablen anzeigen, ZeigeVariablen
		Menu, Tray, Add,
		Menu, Tray, Add, Skript Neu Starten, SkriptReload
		Menu, Tray, Add, Beenden, handle_exit
	;}

	;------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; global variables and others
	;------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		global hSonoCap								;main gui hwnd
		global AlbisWinID		:= AlbisWinID()
		global AddendumDir := RegReadUniCode64("HKEY_LOCAL_MACHINE", "SOFTWARE\Addendum für AlbisOnWindows", "ApplicationDir")
		scriptPID := DllCall("GetCurrentProcessId")
		export := []
		df			:= DPIFactor()
		PatID	:= AlbisAktuellePatID()
		Patient	:= AlbisCurrentPatient()
		GebDat:= AlbisPatientGeburtsdatum()

	;}

	;------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; load avicap32 lib and start GDIP_Plus
	;------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		hModule := DllCall("LoadLibrary", "str", "avicap32.dll")
		If !pToken := Gdip_Startup()
		{
				MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
				ExitApp
		}
	;}

	;------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Variables for lock/unlock detection
	;------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		NOTIFY_FOR_ALL_SESSIONS     		:=   1
		WTS_SESSION_LOCK                		:=   0x7
		WTS_SESSION_UNLOCK           		:=   0x8
		WM_WTSSESSION_CHANGE     		:=   0x02B1
		WM_USER                                 		:=   0x0400
		WM_CAP_START                        		:=   WM_USER 									; 0x1024
		WM_CAP_GRAB_FRAME_NOSTOP	:=   WM_USER + 61
		WM_CAP_FILE_SAVEDIB             		:=   WM_CAP_START + 25

		WM_CAP                                      	:=   0x400
		WM_CAP_DRIVER_CONNECT       	:=   WM_CAP + 10
		WM_CAP_DRIVER_DISCONNECT    	:=   WM_CAP + 11
		WM_CAP_EDIT_COPY                    	:=   WM_CAP + 30
		WM_CAP_SET_PREVIEW                	:=   WM_CAP + 50
		WM_CAP_SET_PREVIEWRATE            	:=   WM_CAP + 52
		WM_CAP_SET_SCALE                    	:=   WM_CAP + 53
	;}

	;------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; ini read
	;------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		If FileExist(AddendumDir . "\Addendum.ini") {
				IniRead, PraxisName	    	, % AddendumDir "\Addendum.ini", 	Allgemeines   	, PraxisName
				IniRead, Strasse 		    	, % AddendumDir "\Addendum.ini", 	Allgemeines   	, Strasse
				IniRead, PLZ	                	, % AddendumDir "\Addendum.ini", 	Allgemeines   	, PLZ
				IniRead, Ort	                	, % AddendumDir "\Addendum.ini", 	Allgemeines   	, Ort
				IniRead, Module	    		, % AddendumDir "\Addendum.ini", 	% CompName	, Module
				IniRead, BefundOrdner   	, % AddendumDir "\Addendum.ini", 	ScanPool       	, BefundOrdner
				IniRead, Comp              	, % AddendumDir "\Addendum.ini", 	Computer     	, % CompName
				IniRead, capWidth         	, % AddendumDir "\Addendum.ini", 	SonoCapture		, capWidth        	, 720
				IniRead, capHeight        	, % AddendumDir "\Addendum.ini", 	SonoCapture		, capHeight       	, 576
				IniRead, BgColor           	, % AddendumDir "\Addendum.ini", 	SonoCapture		, BgColor           	, 172842
				IniRead, PicPre              	, % AddendumDir "\Addendum.ini", 	SonoCapture		, Previewer        	, Irfanview
				IniRead, SPicFont           	, % AddendumDir "\Addendum.ini", 	SonoCapture		, Font                	, Futura Bk Bt
				IniRead, SPicOptions       	, % AddendumDir "\Addendum.ini", 	SonoCapture		, TextOptions1     	, x2 y0 cffFFFFFF Left r1
				IniRead, PTextOptions       	, % AddendumDir "\Addendum.ini", 	SonoCapture		, TextOptions2     	, x2 y14 cffFFFFFF
				IniRead, Untersuchungen	, % AddendumDir "\Addendum.ini", 	SonoCapture		, Untersuchungen	, Sono Abdomen|Weichteile|Gefäße|Nieren und Blase|Restharn|Sonstiges
				PraxisText := PraxisName ", " Strasse " in" PLZ " " Ort
		}	else {
				MsgBox, 1, , Es gibt bisher keine Einstellungen in der .ini Datei`nIch nehme daher die Defaultwerte, 1
					capWidth:=720, capHeight:=576, BgColor:=172842					;PAL Auflösung
		}
	;}

	;------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; gui calculations
	;------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		gboxwidth:=capwidth+10,gboxheight:=capheight+80
		buttonys:=capheight + 20, buttonxs:=capwidth - 230 - 10
		picxs:=capwidth - 300 - 10
		FPS:= 15
	;}

	;------------------------------------------------------------------------------------------------------------------------------------------------------------------
;}

;{3. starts main and sub gui, checks the driver, starts capture driver

	;OnMessage(0x201, "WM_LBUTTONDOWN")

	;------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; main gui
	;------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		Gui, Sono: +dpiscale -SysMenu HWNDhSonoCap
		Gui, Sono: Margin		, 5, 5
		;Gui, Sono: Add	    	, Progress, % "x0 y0 w900 h900 vBack c" BgColor, 100

		Gui, Sono: Color		, c%BgColor%
		Gui, Sono: Font			, s18 Futura Bk Bt q5
		Gui, Sono: Add			, Text, Center cFFFFFF Section , Patient: %Patient%

		Gui, Sono: Font			, s12 Futura Bk Bt q5
		Gui, Sono: Add			, ComboBox, xs vUntersuchung , %Untersuchungen%
		Gui, Sono: Add			, Edit, r1 w300 vZusatz HWNDhZusatz, zusätzliche Angaben

		Gui, Sono: Add			, Picture, xs+%picxs% y0, %AddendumDir%\assets\PraxisLogo300x120.png

		Gui, Sono: Font			, s10 Futura Bk Bt q5
		Gui, Sono: Add			, GroupBox, xm y110 w%gboxwidth% h%gboxheight% Section cFFFFFF, Video
		Gui, Sono: Add			, Text	,% "xs+5 y124 w" capwidth " h" capheight " Section vVidPlaceholder HWNDhVidPlaceholder"

		Gui, Sono: Font			, s18 q5
		Gui, Sono: Add			, Button	, % "xs+5 ys+" buttonys " w190 h28 cFFFFFF gCopyToClipBoard", Capture
		Gui, Sono: Add			, Button	, % "xs+" buttonxs " ys+" buttonys " w230 h54 cFFFFFF gEnde", Bilder übertragen`n und Beenden
		Gui, Sono: Font			, s12 cFFFFFF q5
		GuiControl, +0x7, VidPlaceholder ; frame
	  ;-: Statusbar
		Gui, Sono: Add			, StatusBar , -Theme Background%BgColor% cFFFFFF,, Done
		SB_SetIcon("shell32.dll", 222, 1)
	  ;-: Standarduntersuchung auswählen
		Gui, Sono: Show, xCenter yCenter AutoSize , Hausarztpraxis Clemenz - Sono Capture
		GuiControl, Sono: Choose, ComboBox1, 1
		gosub ConnectToDriver
		return
	;}

	;------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; sub gui to remember user to start ultrasound device first
	;------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		BGLorDet:=LorD(BgColor, 10)
		GuiControlGet, Vid, Sono: Pos, VidPlaceholder
	;-:dpi aware
		;Wx:= Round(Wx // df), Wy:= Round(Wy // df)
	;-:vertrical center of RTD Gui inside VidPlaceholder element with border of 10%
		Sw:= Round(VidW - (VidW // 10)), Sh:= Round(VidH - (VidH // 10))			;dont fill the hole VidPlaceholder region
		Wy:= VidY + Round((VidH - Sh)//2), Wx:= VidX + Round((VidW - Sw)//2)

	 ;-:draw sub gui
		Gui, RTD: -Caption -dpiscale HWNDhRTD
		Gui, RTD: Margin	, 0, 0
		Gui, RTD: Color		, c%BGLorDet%
		Gui, RTD: Add    	, Progress, % "x0 y0 w" Sw " h" Sh " vBlink HWNDhblink BackgroundRed"

		;Gui, RTD: Font		, % "s1 q5 " Font
		;Gui, RTD: Add		, GroupBox, % "x0 y0 w" Sw " h" Sh " HWNDhGBBig"

		Gui, RTD: Font		, % "s42 q5 " Font
		Gui, RTD: Add		, Text	, % "x0 y" (0.5*(Sh/4)) " w" sw " Center c000000 +BackgroundTrans HWNDhNote1", NICHT VERGESSEN!

		Gui, RTD: Font		, % "s26 q5 " Font
		Gui, RTD: Add		, Text	, % "x0 y" (1.5*(Sh/4)) " w" Sw " Center +BackgroundTrans cFFFF00 HWNDhNote2", Schalte zuerst`ndas Ultraschallgerät ein,`nbevor Du weitermachst!

		Gui, RTD: Font		, % "s20 q5 " Font
		Gui, RTD: Add		, Button , % "x0 y" (2.8*(Sh/4)) " gRTDclicked vRTDb1 HWNDhRTDb1", Ja, ist eingeschaltet!
		Gui, RTD: Font		, % "s16 q5 " Font
		Gui, RTD: Add		, Button , % "x0 y" (3.2*(Sh/4)) " gRTDclicked vRTDb2 HWNDhRTDb2", Abbrechen

		Gui, RTD: Show		, % "Hide x0 y0 w" sw " h" sh, Remember the Driver

	 ;-:move button1 x position
		GuiControlGet, rtd1, RTD: Pos, RTDb1
		GuiControl, RTD: Move, RTDb1, % "x" Round((Sw - rtd1W)/2)

	 ;-:move button2 x position
		GuiControlGet, rtd2, RTD: Pos, RTDb2
		GuiControl, RTD: Move, RTDb2, % "x" Round((Sw - rtd2W)/2) " y" (rtd1Y + rtd1H + 10)

		Gui, RTD: Show, % "x" Wx " y" Wy , Remember the Driver

	;-:Set this Gui as a child window
		DllCall("SetParent", "uint", hRTD, "uint", hSonoCap)
			sleep, 100
		WinSet, Redraw,, ahk_id %hRTD%
		Gui, RTD: Default

		onoff:= 0
		CoordMode, Pixel, Screen

	 ;-:Animation
		;SetTimer, Blink, 1000

	;}


return

Blink: ;{
	onoff:= !onoff
	GuiControl, RTD: Hide%onoff%, Blink
Return
;}


RTDclicked: ;{
	CName:= A_GuiControl
	If InStr(CName, "RTD1")
	{
			SetTimer, Blink, Off
			Gui, RTD: Destroy
			gosub ConnectToDriver
	}
	else If InStr(CName, "RTD2" )
	{
			ExitApp
	}
return
;}

Ende: ;{
	OnExit
	Gdip_Shutdown(pToken)
	DllCall("FreeLibrary", "str", hModule)
	;OnMessage(0x201, "")
ExitApp
;}

#IfWinExist, Sono Capture
LButton & RButton::gosub, CopytoClipBoard
#IfWinExist

;}

;{4. Labels - ConnectToDriver, SavePicture from Stream, *DisconnectDriver can be found at handle_exit: Label

ConnectToDriver:					;{   4a. Capture

		OnExit, handle_exit
		WinGetPos, wx, wy, ww, wh, % "ahk_id " hSonoCap
		wx:=wx+ww

	  ; select driver
		SelectedDriver:= Cap_GetDriverDescription(0)
		SciTEOutput("GetDriverDescription: " . (ErrorLevel=0 ? "failed" : "ok") . "`nSelected Driver: " . SelectedDriver, 1, 1, 0)

	  ; Connect and preview
		capHwnd := Cap_CreateCaptureWindow(hSonoCap, 10, 124, capwidth, capheight, "SonoCapture.ahk - " . PraxisName)
		SciTEOutput(el:= "CreateCapture: " . (capHwnd=0 ? "failed" : "ok"), 0, 1, 0)

	  ; Connect to driver
		SendMessage, WM_CAP_DRIVER_CONNECT, %SelectedDriver%, 0, , % "ahk_id " capHwnd
		capcon:=ErrorLevel
		Loop
		{							;
			sleep, 10
		} until A_Index>100 or ErrorLevel=1

		SciTEOutput("Driver_Connection: " . (ErrorLevel=0 ? "failed" : "ok") . ", A_LastError :" . A_LastError, 0, 1, 0)
		If (ErrorLevel<>1) {
				MsgBox, Failed to connect to capture driver: %ErrorLevel%
				ExitApp
		}

	 ; Capture start
		SendMessage, WM_CAP_START, %SelectedDriver%, 0, , % "ahk_id " capHwnd
		SciTEOutput( "Cap_Start: " . (ErrorLevel="FAIL" ? "failed" : "ok"), 0, 1, 0)
		If (capstart="FAIL")
				MsgBox, CaptureStart: %ErrorLevel%

	  ; Set the preview scale
		SendMessage, WM_CAP_SET_SCALE, 1, 0, , % "ahk_id " capHwnd
		SciTEOutput("Set_Scale: " . (ErrorLevel=0 ? "failed" : "ok"), 0, 1, 0)

	  ; Set the preview rate in milliseconds
		MSC := round((1/FPS)*1000)
		SendMessage, WM_CAP_SET_PREVIEWRATE, %MSC%, 0, , % "ahk_id " capHwnd
		SciTEOutput("Set_Previewrate to " . FPS . " frames/s: " (ErrorLevel=0 ? "failed" : "ok"), 0, 1, 0)

	  ; Start previewing the image from the camera
		SendMessage, WM_CAP_SET_PREVIEW, 1, 0, , % "ahk_id " capHwnd
		SciTEOutput("Set_Preview: " . (ErrorLevel=0 ? "failed" : "ok"), 0, 1, 0)

Return
;}--------------------------------------------------------------------------------------------

CopyToClipBoard:					;{    4b. Writing text to image

    countf ++																					;Picture counter
	SciTEOutput(countf, 0, 1, 0)
    ControlGetText, Untersuchung,, % "ahk_id " hUntersuchungen
    ControlGetText, Zusatz,, % "ahk_id " hZusatz
    Gui, Sono: Submit, NoHide
    if Instr(Zusatz, "zusätzliche Angaben") {
          Zusatz:=""
    }

    imagefile			:= BefundOrdner . "\" . Patient . "`, " . Trim(Untersuchung) . SubStr("000" . countf, -3) . ".jpg"
	SPicBildtext		:= "Pat.:(" . PatID . ") " Patient . " " . GebDat . ", Datum: " . A_DD . "`." . A_MM . "`." . A_YYYY .  ", " . Trim(Untersuchung) . " " . Trim(Zusatz)
    nofile 				:= 1
	export[countf,1]	:= imagefile
	export[countf,2]	:= SPicBildtext

	sendAgain:
	ClipBoard:= ""
	SendMessage, WM_CAP_EDIT_COPY, 0, 0, , % "ahk_id " capHwnd
   ; sleep, 500
	ClipWait, 2

	ImageTextOverlay("clipboard", imagefile, SPicBildtext, PraxisText, SPicOptions, PTextOptions, SPicFont, BgColor)
/*
  If FileExist(imagefile) {

          nofile:= 0
          WinActivate, Albis ahk_class OptoAppClass

          Aktentext	:= Trim(Untersuchung) . " " . Trim(Zusatz) . "[" . countf . "]"
          image		:= Untersuchung . countf . ".png"

          ;AlbisPrepareInput("bild1")   ;öffnet das Grafische Befund Fenster

          ;AlbisUebertrageGrafischenBefund(image, Aktentext)

		;if the user has not set Albis preview, this part of the program does not need to be resolved (new to V0.53)
		if !(picpre="No")
		{
					WinWait, %picpre%,, 2													;picpre contains name of preview window
					Loop {
							hpicpre:= FindWindow(picpre)

					} until hpicpre

					;this routine uses WM_Command for Menu/Quit for IrfanView (use WinSpy to get this information for the most windows)
					 If Instr(picpre, "IrfanView") {
								SendMessage, 0x111, 1080,,, ahk_id %hpicpre%
					}

					i:=0
					While WinExist("ahk_id " . hpicpre)
					{
							Win32_SendMessage(hpicpre)					;new to version 0.52
							sleep, 100
							i++
							If (i>10)
									break
					}
					i:=0
		}

		WinActivate, Albis ahk_class OptoAppClass
          WinActivate, ahk_id %hSonoCap%
			ToolTip, 1

    }
*/

Return
;}

handle_exit: ;{
	OnExit
	;OnMessage(0x201, "")
	Gdip_Shutdown(pToken)
    SendMessage,% WM_CAP_DRIVER_DISCONNECT, 1, 0, , % "ahk_id " capHwnd
	DllCall("FreeLibrary", "str", hModule)
    ;DllCall( "gdi32.dll\DeleteObject", "uint", hbm_buffer )
    ;DllCall( "gdi32.dll\DeleteDC"     , "uint", hdc_frame )
    ;DllCall( "gdi32.dll\DeleteDC"     , "uint", hdc_buffer )
ExitApp
;}



;}

;{5. Important local functions
ImageTextOverlay(imgInput, imgOutput, Text1, Text2, Options1, Options2, Font, BgColor) {

		if InStr(imgInput, "clipboard")
				pBitmap := Gdip_CreateBitmapFromClipboard()
		else
				pBitmap := Gdip_CreateBitmapFromFile(imgInput)

		iHeight	:= Gdip_GetImageHeight(pBitmap)
		iWidth	:= Gdip_GetImageWidth(pBitmap)

		Fontsize1:= Floor(iHeight/45)
		Options1:= Options1 . " s" . Fontsize1

		Fontsize2:= Floor(iHeight/55)
		Options2:= Options2 . " s" . Fontsize2

		RegExMatch(Options1, "i)R(\d)", Rendering)
		RegExMatch(Options1, "i)S(\d+)(p*)", Size)
		RegExMatch(Options1, "i)X([\-\d\.]+)(p*)", xpos)
		RegExMatch(Options1, "i)Y([\-\d\.]+)(p*)", ypos)

		rwidth:= Size * zl + 10
		rheight:= Rendering * zn + 10
		rxpos:= xpos - 15
		rypos:= ypos - 15

	;-:begin to draw on capture
		G := Gdip_GraphicsFromImage(pBitmap)
	;-:text background
		Gdip_SetCompositingMode(G, 4)
		pBrush := Gdip_BrushCreateSolid("0xff" . BgColor)
		Gdip_FillRectangle(G, pBrush, 0, 0, iWidth, 30)
		Gdip_DeleteBrush(pBrush)
	;-:white line between text
		pBrush := Gdip_BrushCreateSolid("0xffFFFFFF")
		Gdip_FillRectangle(G, pBrush, 0, 16, iWidth-40 , 1)
		Gdip_DeleteBrush(pBrush)
	;-:draw text
		Gdip_SetCompositingMode(G, 0)
		Gdip_SetSmoothingMode(G, 4)
		Gdip_TextToGraphics(G, Text2, Options2, Font)
		Gdip_TextToGraphics(G, Text1, Options1, Font)
		Gdip_SaveBitmapToFile(pBitmap, imgOutput, 100)
		Gdip_DeleteGraphics(G), Gdip_DisposeImage(pBitmap)

}

GetClientRect(hwnd) {
	VarSetCapacity(rc, 16)
	result := DllCall("GetClientRect", "PTR", hwnd, "PTR", &rc, "UINT")
	return {x : NumGet(rc, 0, "int"), y : NumGet(rc, 4, "int"), w : NumGet(rc, 8, "int"), h : NumGet(rc, 12, "int")}
}

Cap_FileSaveAs(capHwnd, filename) {
        DLLCall("avicap32.dll\capFileSaveAs", "Uint", capHwnd, "Int", &filename)
}

Cap_CreateCaptureWindow(hSonoCap, x, y, w, h, lpszWindowName) {

  WS_CHILD := 0x40000000
  WS_VISIBLE := 0x10000000

  lwndC := DLLCall("avicap32.dll\capCreateCaptureWindowW"
                  , "Str", lpszWindowName
                  , "UInt", WS_VISIBLE | WS_CHILD ; dwStyle
                  , "Int", x
                  , "Int", y
                  , "Int", w
                  , "Int", h
                  , "UInt", hSonoCap
                  , "Int", 0)

  Return lwndC
}

Cap_GetDriverDescription(wDriver) {
  VarSetCapacity(lpszName, 100)
  VarSetCapacity(lpszVer, 100)
  res := DLLCall("avicap32.dll\capGetDriverDescriptionW"
                  , "Short", wDriver
                  , "Str", lpszName
                  , "Int", 100
                  , "Str", lpszVer
                  , "Int", 100)
  If res
    capInfo := lpszName ; " | " lpszVer
  Return capInfo
}

WM_LBUTTONDOWN(wParam, lParam) {

	If GetKeystate("LControl", "D")
			PostMessage, 0xA1, 2,,, ahk_id %hSonoCap%

}

LorD(x, n) {		; LIGHTEN or DARKEN x must be Hex , add with decimal f.e. n = 1 or n = -1
						;this can be done better! but it works and looks nice

	Red:= GetDec(SubStr(x, 1, 2)) + n
	Red:= (Red>255) ? 255 : Red
	Red:= (Red<0) ? 0 : Red

	Green:= GetDec(SubStr(x, 3, 2)) + n
	Green:= (Green>255) ? 255 : Green
	Green:= (Green<0) ? 0 : Green

	Blue:= GetDec(SubStr(x, 5, 2)) + n
	Blue:= (Blue>255) ? 255 : Blue
	Blue:= (Blue<0) ? 0 : Blue

	x:= GetHex(Red) . GetHex(Green) . GetHex(Blue)

  Return x
}

Create_SonoCapture_ico(NewHandle := False) {
Static hBitmap := Create_SonoCapture_ico()
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "AAABAAEAMDAAAAEAGACoHAAAFgAAACgAAAAwAAAAYAAAAAEAGAAAAAAAAAAAAGAAAABgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo8AAAAAAAAAAAAAAAAAAAAAAAAAAAD/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo8AAAAAAAAAAAAAAAAAAAD/oo/9oY7/oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/9oY7/oo8AAAAAAAAAAAD/oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo8AAAAAAAD/oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo9CKBdCKBdCKBe2c2Hok4GNWEdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBfqlIL/oo//oo/KgG5CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBe+eGX/oo//oo//oo+kZ1VCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBfLgW7/oo//oo//oo/8oI3ymofymofvmIXwmIXOgnBDKRhCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBfOgnD/oo//oo//oo//oo//oo//oo//oo//oo/ejXpJLBxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdJLBz/oo//oo//oo//oo//oo//oo//oo//oo//oo/0m4ijZ1RCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdGKxn/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/5nourbFpCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdDKRj+oY7/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/8oI3WiHVCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf6n4z/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/9oY7BemhCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf1m4j/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/OgnBDKRhCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBfIfmz/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/aindGKxlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfolIH/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/lkX9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdHKxrxmYb/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/tl4Swb1zShXLzmof7n43Pg3FMLx5CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf3nYr/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/1nIm6dmNCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf4nov/oo//oo//oo//oo//oo//oo/vmIVCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/FfGpCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf8oI3/oo//oo//oo//oo//oo9CKBdCKBdDKRhCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/EfWpCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo+0cmBCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo/+oY5CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdDKRhCKBf/oo//oo//oo//oo//oo9CKBdCKBf1m4j/oo//oo//oo//oo//oo//oo//oo//oo/vmIVCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdILBpCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo+4dGFCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfmkX//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/ymodCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBe0cV//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo+AUD9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfThXP/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo+/eGdCKBdCKBdCKBdCKBdCKBf/oo/bi3hCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf3nYr/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/ymYZCKBdCKBdCKBdCKBdCKBf/oo//oo/DfGlCKBdCKBdCKBf/oo/8oI1CKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf5nov/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/ymodCKBdCKBdCKBdCKBdCKBf/oo//oo/vmIVCKBdCKBdEKhlCKBf/oo//oo9CKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfUhnT/oo//oo//oo//oo//oo//oo//oo//oo//oo/Jf21CKBdCKBdCKBdCKBdCKBf1nIn/oo//oo9CKBdCKBdCKBdCKBf/oo//oo/ok4FCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdPMB/2nIr/oo//oo//oo//oo//oo//oo/ymoeIVURCKBdCKBdCKBdCKBdCKBfnk4D/oo//oo/DfGlCKBdCKBdILBpCKBf/oo//oo99Tj1CKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeDUkD/oo//oo//oo/9oY3jkH29d2VGKhpCKBdCKBdCKBdCKBdCKBflkX7/oo//oo/fjntOLx9CKBdCKBdCKBf/oo//oo/kkH1CKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfDemjBemiycF6eY1JCKBdCKBdCKBdCKBdCKBdCKBdCKBfzm4f/oo//oo/qlIJCKBdCKBdCKBdCKBf7oIz/oo/9oY5jPSxCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo/gjntCKBdCKBdCKBdCKBfslYL/oo//oo/MgW9CKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo/HfmxCKBdCKBdCKBdCKBfhj33/oo//oo/ikH1EKhlCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo/1m4hCKBdCKBdCKBdCKBdCKBfhj33/oo//oo/ymodCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo9CKBdCKBdCKBdCKBdCKBftl4T/oo//oo/ymodCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/7nZD/oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdJLBxCKBf9oI3/oo//oo/xmYZCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdDKRhCKBfnkn//oo//oo//oo/ShXJCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo/9oI3LgW5EKhlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo/UhnNCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo8AAAD/oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf8oI3/oo/ShXJCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo8AAAAAAAAAAAD/oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo8AAAAAAAAAAAAAAAAAAAD/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdjPSz/oo//oo//oo8AAAAAAAAAAAAAAAAAAAAAAAAAAAD/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/7nY//oo//oo//oo//oo8AAAAAAAAAAAAAAAAAAAD4AAAAAB8AAPAAAAAADwAA4AAAAAAHAADAAAAAAAMAAIAAAAAAAQAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAAAAQAAwAAAAAADAADgAAAAAAcAAPAAAAAADwAA+AAAAAAfAAA="
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

SkriptReload: ;{

	Script:= A_ScriptDir "\" A_ScriptName
	scriptPID := DllCall("GetCurrentProcessId")
	run,  Autohotkey.exe /f "%AddendumDir%\include\RestartScript.ahk" "%Script%" "2" "%A_ScriptHwnd%" "%scriptPID%"

return

ZeigeVariablen:

	ListVars

return

;}

RefreshDrivers:     ;{ not used at the moment

	foundDriver:= 0
	thisInfo := Cap_GetDriverDescription(A_Index-1)
    If thisInfo
    {
			foundDriver:= 1
			SelectedDriver:= thisinfo
     }

Return
;}

GDIplusError:
GDIplusStop: ;{
   If (#GDIplus_lastError != "")
      MsgBox 16, GDIplus Test, Error in %#GDIplus_lastError%
   GDIplus_Stop()
    ExitApp
Return
;}

PraxTTOff: ;{
	AnimateWindow(hPraxTT, PraxTTDuration, BLEND)
	Gui, PraxTT: Destroy
return
;}

;}

;{ 6. Includes

	#Include %A_ScriptDir%\..\..\lib\gdip_all.ahk
	#Include %A_ScriptDir%\..\..\lib\GDIPlusHelper.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Functions.ahk
	#Include %A_ScriptDir%\..\..\lib\ini.ahk
	#Include %A_ScriptDir%\..\..\include\Gui\PraxTT.ahk
	;#Include %A_ScriptDir%\..\..\include\SciTEOutput.ahk

;}

;############################################################################################################





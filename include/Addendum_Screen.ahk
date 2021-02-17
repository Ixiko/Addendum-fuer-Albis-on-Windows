;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                            	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                                            	by Ixiko started in September 2017 - last change 27.09.2020 - this file runs under Lexiko's GNU Licence
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

;	MONITOR / SCREEN                                                                                                                                                                                                                             	(06)
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; GetMonitorIndexFromWindow          	GetMonitorAt                            	screenDims                                    	MonitorScreenShot                        	DPIFactor
; DPI
; ____________________________________________________________________________________________________________________________________________________________

GetMonitorIndexFromWindow(windowHandle) {

	; Starts with 1.
	; https://autohotkey.com/board/topic/69464-how-to-determine-a-window-is-in-which-monitor/

	monitorIndex := 1
	VarSetCapacity(monitorInfo, 40)
	NumPut(40, monitorInfo)

	if (monitorHandle := DllCall("MonitorFromWindow", "uint", windowHandle, "uint", 0x2)) && DllCall("GetMonitorInfo", "uint", monitorHandle, "uint", &monitorInfo) {

		monitorLeft  		:= NumGet(monitorInfo,  4, "Int")
		monitorTop    	:= NumGet(monitorInfo,  8, "Int")
		monitorRight  	:= NumGet(monitorInfo, 12, "Int")
		monitorBottom 	:= NumGet(monitorInfo, 16, "Int")
		workLeft      		:= NumGet(monitorInfo, 20, "Int")
		workTop       	:= NumGet(monitorInfo, 24, "Int")
		workRight     		:= NumGet(monitorInfo, 28, "Int")
		workBottom    	:= NumGet(monitorInfo, 32, "Int")
		isPrimary     		:= NumGet(monitorInfo, 36, "Int") & 1

		SysGet, MonCount, MonitorCount
		Loop % MonCount 	{                                    		; Compare location to determine the monitor index.
			SysGet, tempMon, Monitor, %A_Index%
			if ((monitorLeft = tempMonLeft) && (monitorTop = tempMonTop) && (monitorRight = tempMonRight) && (monitorBottom = tempMonBottom))
				return monitorIndex
		}

	}

return monitorIndex
}

GetMonitorAt(Lx, Ly, Ldefault:=1) {                                                        	                                    	;-- Get the index of the monitor containing the specified x and y co-ordinates.

	; https://autohotkey.com/board/topic/19990-windowpad-window-moving-tool/page-2
	; letzte Änderung: 27.09.2020

    SysGet, Lm, MonitorCount
    Loop % Lm {   ; Check if the window is on this monitor.

        SysGet, Mon, Monitor, %A_Index%
        if (Lx >= MonLeft && Lx <= MonRight && Ly >= MonTop && Ly <= MonBottom)
            return A_Index

    }

return Ldefault
}

ScreenDims(MonNr:=1) {	                                                                                        		    			;-- returns a key:value pair of width screen dimensions (only for primary monitor)

	Sysget, MonitorInfo, Monitor, % MonNr
	X	:= MonitorInfoLeft
	Y	:= MonitorInfoTop
	W	:= MonitorInfoRight - MonitorInfoLeft
	H 	:= MonitorInfoBottom - MonitorInfoTop

	DPI    	:= A_ScreenDPI
	Orient	:= (W>H)?"L":"P"
	yEdge	:= DllCall("GetSystemMetrics", "Int", SM_CYEDGE)
	yBorder	:= DllCall("GetSystemMetrics", "Int", SM_CYBORDER)

 return {X:X, Y:Y, W:W, H:H, DPI:DPI, OR:Orient, yEdge:yEdge, yBorder:yBorder}
}

MonitorScreenShot(MonNr, ScriptName:="", Path:="") {                                                                     	;-- erstellt einen Screenshot von einem Monitor

	; use "All" if you like to get all screens
	; letzte Änderung: 27.09.2020

	static raster := 0x40000000 + 0x00CC0020 ; get layered windows

	If !RegExMatch(MonNr, "^\s*(?<Nr>\d+)\s*$", Mon) {
		throw % "Wrong value in parameter 'MonNr'. Only integer is aloud."
		return 0
	}

	MSSpToken := Gdip_Startup()

	;Monitorgröße bestimmen
	If MonNr in 1,2
	{
		Sysget, MonitorInfo, Monitor, % MonNr
		sX 	:= MonitorInfoLeft
		sY 	:= MonitorInfoTop
		sW 	:= MonitorInfoRight    	- MonitorInfoLeft
		sH 	:= MonitorInfoBottom 	- MonitorInfoTop
		screen:= sX . "|" . sY . "|" . sW . "|" . sH
	} else if (MonNr=="All") {

	}

	Zeitstempel:=  TimeCode(1)
	outfile := Path "\Screenshot_" ScriptName "_" ZeitStempel ".jpg"

	pBitmap := Gdip_BitmapFromScreen(screen,raster)
	Gdip_SetBitmapToClipboard(pBitmap)
	Gdip_SaveBitmapToFile(pBitmap, outfile)
	Gdip_DisposeImage(pBitmap)
	Gdip_Shutdown(MSSpToken)

}

DPIFactor() {
RegRead, DPI_value, HKEY_CURRENT_USER, Control Panel\Desktop\WindowMetrics, AppliedDPI
; the reg key was not found - it means default settings
; 96 is the default font size setting
if (errorlevel=1) OR (DPI_value=96 )
	return 1
else
	Return  DPI_Value/96
}

DPI(in="",setdpi=1) 	{

	/*
	Name                  	: DPI
	Purpose            	: Return scaling factor or calculate position/values for AHK controls (font size, position (x y), width, height)
	Version             	: 0.31
	Source               	: https://github.com/hi5/dpi
	AutoHotkey Forum : https://autohotkey.com/boards/viewtopic.php?f=6&t=37913
	License              	: see license.txt (GPL 2.0)
	Documentation   	: See readme.md @ https://github.com/hi5/dpi

	History:

	* v0.31	: refactored "process" code, just one line now
	* v0.3	: - Replaced super global variable ###dpiset with static variable within dpi() to set dpi
				  - Removed r parameter, always use Round()
				  - No longer scales the Rows option and others that should be skipped (h-1, *w0, hwnd etc)
	* v0.2	: public release
	* v0.1	: first draft

	*/

	static dpi:=1, AppliedDPI

	 if (setdpi <> 1)
		dpi := setdpi

	 if !AppliedDPI {
		; If the AppliedDPI key is not found the default settings are used. 96 is the default value.
		RegRead, AppliedDPI, HKEY_CURRENT_USER, Control Panel\Desktop\WindowMetrics, AppliedDPI
	 	if (ErrorLevel=1) || (AppliedDPI=96)
			AppliedDPI := 96
	}

	 if (dpi <> 1)
		AppliedDPI := dpi

	factor := (AppliedDPI / 96)
	;factor := (96 / screenDims(1).DPI)

	If !in
	  return factor

	 Loop, parse, in, %A_Space%%A_Tab%
	 {

		 option := A_LoopField
		 if RegExMatch(option,"i)(w0|h0|h-1|xp|yp|xs|ys|xm|ym)$") || RegExMatch(option,"i)(icon|hwnd)") ; these need to be bypassed
			out .= option A_Space
		 else if RegExMatch(option,"i)^\*{0,1}(x|xp|y|yp|w|h|s)[-+]{0,1}\K(\d+)",number) ; should be processed
			out .= StrReplace(option, number, Round(number*factor)) A_Space
		 else ; the rest can be bypassed as well (variable names etc)
			out .= option A_Space

	}

Return Trim(out)
}
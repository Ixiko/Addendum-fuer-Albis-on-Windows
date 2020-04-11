;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                            	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                                            	by Ixiko started in September 2017 - last change 07.10.2019 - this file runs under Lexiko's GNU Licence
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

;	MONITOR / SCREEN                                                                                                                                                                                                                             	(05)
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; GetMonitorIndexFromWindow          	GetMonitorAt                            	screenDims                                    	MonitorScreenShot                        	DPIFactor
; ____________________________________________________________________________________________________________________________________________________________

GetMonitorIndexFromWindow(windowHandle) {
	; Starts with 1.
	; https://autohotkey.com/board/topic/69464-how-to-determine-a-window-is-in-which-monitor/
	monitorIndex := 1

	VarSetCapacity(monitorInfo, 40)
	NumPut(40, monitorInfo)

	if (monitorHandle := DllCall("MonitorFromWindow", "uint", windowHandle, "uint", 0x2))
		&& DllCall("GetMonitorInfo", "uint", monitorHandle, "uint", &monitorInfo)
	{
		monitorLeft   		:= NumGet(monitorInfo,  4, "Int")
		monitorTop    	:= NumGet(monitorInfo,  8, "Int")
		monitorRight  	:= NumGet(monitorInfo, 12, "Int")
		monitorBottom 	:= NumGet(monitorInfo, 16, "Int")
		workLeft      		:= NumGet(monitorInfo, 20, "Int")
		workTop       	:= NumGet(monitorInfo, 24, "Int")
		workRight     		:= NumGet(monitorInfo, 28, "Int")
		workBottom    	:= NumGet(monitorInfo, 32, "Int")
		isPrimary     		:= NumGet(monitorInfo, 36, "Int") & 1

		SysGet, monitorCount, MonitorCount

		Loop, % monitorCount
		{
			SysGet, tempMon, Monitor, %A_Index%

			; Compare location to determine the monitor index.
			if ((monitorLeft = tempMonLeft) and (monitorTop = tempMonTop)
				and (monitorRight = tempMonRight) and (monitorBottom = tempMonBottom))
			{
				monitorIndex := A_Index
				break
			}
		}
	}

	return % monitorIndex
}

GetMonitorAt(Lx, Ly, Ldefault:=1) {                                                        	                                    	;-- Get the index of the monitor containing the specified x and y co-ordinates.
	; https://autohotkey.com/board/topic/19990-windowpad-window-moving-tool/page-2
    ;Ly := 100
    SysGet, Lm, MonitorCount
    ; Iterate through all monitors.
    Loop, %Lm%
    {   ; Check if the window is on this monitor.
        SysGet, Mon, Monitor, %A_Index%
        if (Lx >= MonLeft && Lx <= MonRight && Ly >= MonTop && Ly <= MonBottom)
            return A_Index
    }

    return Ldefault
}

screenDims(MonNr:=1) {	                                                                                        		    			;-- returns a key:value pair of width screen dimensions (only for primary monitor)

	Sysget, MonitorInfo, Monitor, % MonNr
	X	:= MonitorInfoLeft
	Y	:= MonitorInfoTop
	W	:= MonitorInfoRight - MonitorInfoLeft
	H 	:= MonitorInfoBottom - MonitorInfoTop

	DPI := A_ScreenDPI
	Orient := (W>H)?"L":"P"
	yEdge := DllCall("GetSystemMetrics", "Int", SM_CYEDGE)
	yBorder := DllCall("GetSystemMetrics", "Int", SM_CYBORDER)

 return {X:X, Y:Y, W:W, H:H, DPI:DPI, OR:Orient, yEdge:yEdge, yBorder:yBorder}
}

MonitorScreenShot(MonNr, ScriptName:="", Path:="") {                                                                     	;-- erstellt einen Screenshot von einem Monitor

	;use "All" if you like to get all screens

	MSSpToken := Gdip_Startup()
	raster := 0x40000000 + 0x00CC0020 ; get layered windows

	;Monitorgröße bestimmen
	If (MonNr=1) or (MonNr=2) {
		Sysget, MonitorInfo, Monitor, %MonNr%
		sX := MonitorInfoLeft, sY := MonitorInfoTop
		sW := MonitorInfoRight - MonitorInfoLeft
		sH := MonitorInfoBottom - MonitorInfoTop
		screen:= sX . "|" . sY . "|" . sW . "|" . sH
	} else if (MonNr=="All") {

	}

	Zeitstempel:=  TimeCode(1)

	outfile = %Path%\Screenshot_%ScriptName%_%ZeitStempel%`.jpg

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

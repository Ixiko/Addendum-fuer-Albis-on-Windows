; Addendum Something - alles das nicht mehr gebraucht wird oder vielleicht doch noch gebraucht werden könnte

GoogleTranslate(phrase,LangIn,LangOut) {                                                              	;--für Addendum - Strg+C (Deutsch)

		Critical
		base := "https://translate.google.com.tw/?hl=en&tab=wT#"
		path := base . LangIn . "/" . LangOut . "/" . phrase
		IE := ComObjCreate("InternetExplorer.Application")
		;~ IE.Visible := true
		IE.Navigate(path)

		While IE.readyState!=4 || IE.document.readyState!="complete" || IE.busy
				Sleep 50

		Result := IE.document.all.result_box.innertext
		IE.Quit

return Result
}

TeamViewerGSClose() {                                                                                     		;--schließt das Teamviewer Fenster automatisch

	;schließen des gesponserte Sitzung Fenster von Teamviewer
	If WinExist("Gesponserte Sitzung") or WinExist("Verbindungs Timeout!")
				ControlClick, Button4, ahk_exe TeamViewer.exe ahk_class #32770

return
}

MailAndTelegramWindow() {                                                                                 	;--fixieren der Positionen von ClawsMail und Telegramfenster am AnmeldungsPC

	;{-ClawsMailFenster
	ClawsID:= WinExist("Claws Mail ahk_class gdkWindowToplevel")
	If (ClawsID) {

			WinGet, minmax, minmax, ahk_id %ClawsID%
			if (minmax<>0) {
				WinRestore, ahk_id %ClawsID%
			}

			Claws_RecPointer:= GetWindowPos(ClawsID, ClawsX, ClawsY, ClawsW, ClawsH)
			WinGet, Claws_Style, Style, ahk_id %ClawsID%
			WinGet, Claws_ExStyle, ExStyle, ahk_id %ClawsID%


			if (ClawsX<>1686 OR ClawsY<>12 OR ClawsW<>605 OR ClawsH<>1031) {
						WinMove, ahk_id %ClawsID%,, 1686, 12, 605, 1031
				}
			if ((ClawsStyle<> 0x160F0000) OR (ClawsExStyle<> 0x00000110)) {
						WinSet, Style, 0x160F0000, ahk_id %ClawsID%  ; Entfernt die Titelleiste des aktiven Fensters (WS_CAPTION).
						WinSet, ExStyle, 0x0000011, ahk_id %ClawsID%
			}
	}
	;}

	;{-Telegram Fenster
	TGramID:= WinExist("Telegram ahk_class Qt5QWindowIcon")
	If (TGramID) {

			WinGet, minmax, minmax, ahk_id %TGramID%
			if (minmax<>0) {
				WinRestore, ahk_id %TGramID%
			}

			TGram_RecPointer:= GetWindowPos(TGramID, TGramX, TGramY, TGramW, TGramH)
			;WinGet, TGram_Style, Style, ahk_id %TGramID%
			;WinGet, TGram_ExStyle, ExStyle, ahk_id %TGramID%

			if (TGramX<>2302 OR TGramY<>13 OR TGramW<>654 OR TGramH<>1024) {
						WinMove, ahk_id %TGramID%,, 2302, 13, 654, 1024
				}

	}
	;}

}

StrPutVar(string, ByRef var, encoding) {												    					;-- HTML encode/decode or other encode/decode functions
    ; Ensure capacity.
    SizeInBytes := VarSetCapacity( var, StrPut(string, encoding)
        ; StrPut returns char count, but VarSetCapacity needs bytes.
        * ((encoding="utf-16"||encoding="cp1200") ? 2 : 1) )
    ; Copy or convert the string.
    StrPut(string, &var, encoding)
   Return SizeInBytes
}

UriEncode2(str) {
   b_Format := A_FormatInteger
   data := ""
   SetFormat,Integer,H
   SizeInBytes := StrPutVar(str,var,"utf-8")
   Loop, %SizeInBytes%
   {
   ch := NumGet(var,A_Index-1,"UChar")
   If (ch=0)
      Break
   if ((ch>0x7f) || (ch<0x30) || (ch=0x3d))
      s .= "%" . ((StrLen(c:=SubStr(ch,3))<2) ? "0" . c : c)
   Else
      s .= Chr(ch)
   }
   SetFormat,Integer,%b_format%
   return s
}

TTip(msg, time, nr:= 17) {                                                                                        	;-- timed ToolTip
	static tnr
	tnr:= nr
	MouseGetPos, mx, my
	ToolTip, % msg, % my, % my, % tnr
	SetTimer, ToolTipEnd, % time * 1000
return

ToolTipEnd:
	ToolTip,,,, % tnr
return
}

DebugGui(GTitle, ColNames, x, y, width, height) {
;Gui fr Skriptverlauf oder Debugging oder Infos zum Fortschritt ......
gheight:= height +80
global DbgButton1
global DbgButton2
global DbgLV1
Gui, 77:NEW
Gui, 77:Font, S10 CDefault, Futura Bk Bt
Gui, 77:+LastFound +AlwaysOnTop
Gui, 77:Add, Button, x10 y5 vDbgButton1 gDebugGui, Speichern
Gui, 77:Add, Button, x100 y5 vDbgButton2 gDebugGui, Abbrechen
Gui, 77:Add, Listview, x0 y40 r10 w%width%-20 h%height%-10 vDbgLV1 gDebugGui BackgroundAAAAFF LV0x10 Grid NoSort, %ColNames%
Gui, 77:Show,x%x% y%y% w%width% h%gheight%, %GTitle%
Gui, 77:Default
hwnd:=WinExist()
return hwnd
}

Fade(Type="",WinID=0,TotalTime=500,Trans=250) {

		; Window Fade Funktion - gleichmäßige Transparenzänderung da die Funktion
		; über A_TickCount getaktet wird

		Time1 := A_TickCount
		WinShow, ahk_id %WinID%
		Loop	{
				Trans:= Round(((A_TickCount-Time1)/TotalTime)*255)
						If (Type = "in") {
									Trans1:= Trans
						} else if (Type = "out") {
									Trans1:= 255-Trans
						}

				WinSet, Transparent, %Trans1%, ahk_id %WinID%
				if (Trans >= 255) {
								break
						}
		}

		return
	}

Fade_Out(Gui_ID) {

	DllCall("AnimateWindow","UInt",GUI_ID,"Int",1000,"UInt","0x90000") ; Fade out when clicked
	return
}

CreateGradient(Handle, Colors*) {
   GuiControlGet, LC, Pos, %Handle%
   ColorCnt := Colors.Length()
   Size := ColorCnt * 2 * 4
   VarSetCapacity(Bits, Size, 0)
   Addr := &Bits
   For Each, Color In Colors
      Addr := Numput(Color, NumPut(Color, Addr + 0, "UInt"), "UInt")
    HBMP := DllCall("CreateBitmap", "Int", 2, "Int", ColorCnt, "UInt", 1, "UInt", 32, "Ptr", 0, "Ptr")
    HBMP := DllCall("CopyImage", "Ptr", HBMP, "UInt", 0, "Int", 0, "Int", 0, "UInt", 0x2008, "Ptr")
    DllCall("SetBitmapBits", "Ptr", HBMP, "UInt", Size, "Ptr", &Bits)
    HBMP := DllCall("CopyImage", "Ptr", HBMP, "UInt", 0, "Int", LCW, "Int", LCH, "UInt", 0x2008, "Ptr")
    DllCall("SendMessage", "Ptr", Handle, "UInt", 0x0172, "Ptr", 0, "Ptr", HBMP, "Ptr")
    Return True
}

MsgBoxEx(Text, Title := "", Buttons := "", Icon := "", ByRef CheckText := "", Styles := "", Timeout := "", Owner := "", FontOptions := "", FontName := "", BGColor := "", Callback := "") {
    Static hWnd, y2, p, px, pw, c, cw, cy, ch, f, o, gL, hBtn, lb, DHW, ww, Off, k, v, RetVal
    Static Sound := {2: "*48", 4: "*16", 5: "*64"}

    Gui New, hWndhWnd LabelMsgBoxEx -0xA0000 -DPIScale
    Gui % (Owner) ? "+Owner" . Owner : ""
    Gui Font
    Gui Font, % (FontOptions) ? FontOptions : "s9", % (FontName) ? FontName : "Segoe UI"
    Gui Color, % (BGColor) ? BGColor : "White"
    Gui Margin, 10, 12

    If (IsObject(Icon)) {
        Gui Add, Picture, % "x20 y24 w32 h32 Icon" . Icon[1], % (Icon[2] != "") ? Icon[2] : "shell32.dll"
    } Else If (Icon + 0) {
        Gui Add, Picture, x20 y24 Icon%Icon% w32 h32, user32.dll
        SoundPlay % Sound[Icon]
    }

    Gui Add, Link, % "x" . (Icon ? 65 : 20) . " y" . (InStr(Text, "`n") ? 24 : 32) . " vc", %Text%
    GuicontrolGet c, Pos
    GuiControl Move, c, % "w" . (cw + 30)
    y2 := (cy + ch < 52) ? 90 : cy + ch + 34

    Gui Add, Text, vf -Background ; Footer

    Gui Font
    Gui Font, s9, Segoe UI
    px := 42
    If (CheckText != "") {
        CheckText := StrReplace(CheckText, "*",, ErrorLevel)
        Gui Add, CheckBox, vCheckText x12 y%y2% h26 -Wrap -Background AltSubmit Checked%ErrorLevel%, %CheckText%
        GuicontrolGet p, Pos, CheckText
        px := px + pw + 10
    }

    o := {}
    Loop Parse, Buttons, |, *
    {
        gL := (Callback != "" && InStr(A_LoopField, "...")) ? Callback : "MsgBoxExBUTTON"
        Gui Add, Button, hWndhBtn g%gL% x%px% w90 y%y2% h26 -Wrap, %A_Loopfield%
        lb := A_LoopField
        o[hBtn] := px
        px += 98
    }
    GuiControl +Default, % (RegExMatch(Buttons, "([^\*\|]*)\*", Match)) ? Match1 : StrSplit(Buttons, "|")[1]

    Gui Show, Autosize Center Hide, %Title%
    DHW := A_DetectHiddenWindows
    DetectHiddenWindows On
    WinGetPos,,, ww,, ahk_id %hWnd%
    GuiControlGet p, Pos, %lb% ; Last button
    Off := ww - (px + pw)
    For k, v in o {
        GuiControl Move, %k%, % "x" . (v + Off - 14)
    }
    Guicontrol MoveDraw, f, % "x-1 y" . (y2 - 10) . " w" . ww . " h" . 48

    Gui Show
    Gui +SysMenu %Styles%
    DetectHiddenWindows %DHW%

    If (Timeout) {
        SetTimer MsgBoxExTIMEOUT, % Round(Timeout) * 1000
    }

    If (Owner) {
        WinSet Disable,, ahk_id %Owner%
    }

    GuiControl Focus, f
    Gui Font
    WinwaitClose ahk_id %hWnd%
    Return RetVal

    MsgBoxExESCAPE:
    MsgBoxExCLOSE:
    MsgBoxExTIMEOUT:
    MsgBoxExBUTTON:
        SetTimer MsgBoxExTIMEOUT, Delete

        If (A_ThisLabel == "MsgBoxExBUTTON") {
            RetVal := StrReplace(A_GuiControl, "&")
        } Else {
            RetVal := (A_ThisLabel == "MsgBoxExTIMEOUT") ? "Timeout" : "Cancel"
        }

        If (Owner) {
            WinSet Enable,, ahk_id %Owner%
        }

        Gui Submit
        Gui %hWnd%: Destroy
    Return
}

ResConImg(OriginalFile, NewWidth:="", NewHeight:="", NewName:="", NewExt:="", NewDir:="", PreserveAspectRatio:=true, BitDepth:=24) {										;-- Resize and convert images. png, bmp, jpg, tiff, or gif

    /*  ResConImg
         *    By kon
         *    Updated November 2, 2015
         *    http://ahkscript.org/boards/viewtopic.php?f=6&t=2505
         *
         *  Resize and convert images. png, bmp, jpg, tiff, or gif.
         *
         *  Requires Gdip.ahk in your Lib folder or #Included. Gdip.ahk is available at:
         *      http://www.autohotkey.com/board/topic/29449-gdi-standard-library-145-by-tic/
         *
         *  ResConImg( OriginalFile             ;- Path of the file to convert
         *           , NewWidth                 ;- Pixels (Blank = Original Width)
         *           , NewHeight                ;- Pixels (Blank = Original Height)
         *           , NewName                  ;- New file name (Blank = "Resized_" . OriginalFileName)
         *           , NewExt                   ;- New file extension can be png, bmp, jpg, tiff, or gif (Blank = Original extension)
         *           , NewDir                   ;- New directory (Blank = Original directory)
         *           , PreserveAspectRatio      ;- True/false (Blank = true)
         *           , BitDepth)                ;- 24/32 only applicable to bmp file extension (Blank = 24)
     */

    SplitPath, OriginalFile, SplitFileName, SplitDir, SplitExtension, SplitNameNoExt, SplitDrive
    pBitmapFile := Gdip_CreateBitmapFromFile(OriginalFile)                  ; Get the bitmap of the original file
    Width := Gdip_GetImageWidth(pBitmapFile)                                ; Original width
    Height := Gdip_GetImageHeight(pBitmapFile)                              ; Original height
    NewWidth := NewWidth ? NewWidth : Width
    NewHeight := NewHeight ? NewHeight : Height
    NewExt := NewExt ? NewExt : SplitExtension
    if SubStr(NewExt, 1, 1) != "."                                          ; Add the "." to the extension if required
        NewExt := "." NewExt
    NewPath := ((NewDir != "") ? NewDir : SplitDir)                         ; NewPath := Directory
            . "\" ((NewName != "") ? NewName : "Resized_" SplitNameNoExt)       ; \File name
            . NewExt                                                            ; .Extension
    if (PreserveAspectRatio) {                                              ; Recalcultate NewWidth/NewHeight if required
        if ((r1 := Width / NewWidth) > (r2 := Height / NewHeight))          ; NewWidth/NewHeight will be treated as max width/height
            NewHeight := Height / r1
        else
            NewWidth := Width / r2
    }
    pBitmap := Gdip_CreateBitmap(NewWidth, NewHeight                        ; Create a new bitmap
    , (SubStr(NewExt, -2) = "bmp" && BitDepth = 24) ? 0x21808 : 0x26200A)   ; .bmp files use a bit depth of 24 by default
    G := Gdip_GraphicsFromImage(pBitmap)                                    ; Get a pointer to the graphics of the bitmap
    Gdip_SetSmoothingMode(G, 4)                                             ; Quality settings
    Gdip_SetInterpolationMode(G, 7)
    Gdip_DrawImage(G, pBitmapFile, 0, 0, NewWidth, NewHeight)               ; Draw the original image onto the new bitmap
    Gdip_DisposeImage(pBitmapFile)                                          ; Delete the bitmap of the original image
    Gdip_SaveBitmapToFile(pBitmap, NewPath)                                 ; Save the new bitmap to file
    Gdip_DisposeImage(pBitmap)                                              ; Delete the new bitmap
    Gdip_DeleteGraphics(G)                                                  ; The graphics may now be deleted
}

RotateAroundCenter(G, Angle, Width, Height) {																				;-- GDIP rotate around center

	Gdip_TranslateWorldTransform(G, Width / 2, Height / 2)
	Gdip_RotateWorldTransform(G, Angle)
	Gdip_TranslateWorldTransform(G, - Width / 2, - Height / 2)

}

SaveHBITMAPToFile(hBitmap, sFile) {                                                                                         	;-- hBitMap to jpg, gif, png, bmp

	DllCall("GetObject", "Uint", hBitmap, "int", VarSetCapacity(oi,84,0), "Uint", &oi)
	hFile:=	DllCall("CreateFile", "Uint", &sFile, "Uint", 0x40000000, "Uint", 0, "Uint", 0, "Uint", 2, "Uint", 0, "Uint", 0)
	DllCall("WriteFile", "Uint", hFile, "int64P", 0x4D42|14+40+NumGet(oi,44)<<16, "Uint", 6, "UintP", 0, "Uint", 0)
	DllCall("WriteFile", "Uint", hFile, "int64P", 54<<32, "Uint", 8, "UintP", 0, "Uint", 0)
	DllCall("WriteFile", "Uint", hFile, "Uint", &oi+24, "Uint", 40, "UintP", 0, "Uint", 0)
	DllCall("WriteFile", "Uint", hFile, "Uint", NumGet(oi,20), "Uint", NumGet(oi,44), "UintP", 0, "Uint", 0)
	DllCall("CloseHandle", "Uint", hFile)

}

WinGetTextFast(hwnd) {                                                                                                              	;-- WinGetText ALWAYS uses the "fast" mode - TitleMatchMode only affects

	; WinText/ExcludeText parameters.  In Slow mode, GetWindowText() is used
	; to retrieve the text of each control.
	WinGet controls, ControlListHwnd
	static WINDOW_TEXT_SIZE := 32767 ; Defined in AutoHotkey source.
	VarSetCapacity(buf, WINDOW_TEXT_SIZE * (A_IsUnicode ? 2 : 1))
	text := ""
	Loop Parse, controls, `n
	{
		if !DllCall("IsWindowVisible", "ptr", A_LoopField)
			continue
		if !DllCall("GetWindowText", "ptr", A_LoopField, "str", buf, "int", WINDOW_TEXT_SIZE)
			continue
		text .= buf "`r`n"
	}
	return text
}

GetActiveWindow() {
    Return (DllCall("User32.dll\GetForegroundWindow", "Ptr"))
}

GetForegroundWindow() {
	return DllCall("GetTopWindow", "Ptr")
}

BtnBox(Title := "", Prompt := "", List := "") {                                                                              	;--	show a custom MsgBox with arbitrarily named buttons
;-------------------------------------------------------------------------------
    ; show a custom MsgBox with arbitrarily named buttons
    ; return the text of the button pressed
    ;
    ; Title is the title for the GUI
    ; Prompt is the text to display
    ; List is a pipe delimited list of captions for the buttons
    ; Seconds is the time in seconds to wait before timing out
    ;---------------------------------------------------------------------------

    ; create GUI
    Gui, BtnBox: New, +LastFound, %Title%
    Gui, -MinimizeBox
    Gui, Margin, 30, 18
    Gui, Add, Text,, %Prompt%
    Loop, Parse, List, |
        Gui, Add, Button, % (A_Index = 1 ? "" : "x+10") " gBtn", %A_LoopField%

    ; main loop
    Gui, Show
 Return, Result


    ;-----------------------------------
    ; event handlers
    ;-----------------------------------
    Btn: ; all the buttons come here
        Result := A_GuiControl
        Gui, Destroy
    Return

    BtnBoxGuiClose: ; {Alt+F4} pressed, [X] clicked
        Result := "WinClose"
        Gui, Destroy
    Return

    BtnBoxGuiEscape: ; {Esc} pressed
        Result := "EscapeKey"
        Gui, Destroy
    Return
}

ReadIndex() {						                                                                                	;-- this is a specialized function for ScanPool.ahk

		;Teile der Variablen sind globale Variablen

		PageSum	:=0, FileCount:=0, allfiles:="", tidx:=0

		FileCount	:= ScanPoolArray("Load", PdfIndexFile)					;erstellt den files Array aus der pdfIndex.txt Datei
		If FileCount
				PageSum	:= ScanPoolArray("CountPages")

		PdfDirList		:= ReadDir(BefundOrdner, "pdf")
		RegExReplace(PdfDirList, "m)\n", "", filesInDir)
		Progress, ,  ...lösche nicht mehr vorhandene Dateien aus dem Index..

	;nicht mehr vorhandene Dateien aus dem Index nehmen
		For key, val in ScanPool
		{
				If !InStr(PdfDirList, StrSplit(val, "|").1)
				{
						Progress, , % "...lösche nicht mehr vorhandene Dateien (" tidx++ ") aus dem Index..."
						ScanPool.Delete(key)
				}
		}

	fc_old:= FileCount, modmod:= Round( (filesInDir - FileCount)/50 )
	;nach noch nicht aufgenommenen Dateien suchen
		Loop, Parse, PdfDirList, `n, `r
		{
				If !ScanPoolArray("Find", A_LoopField)
				{
						If Mod(FileCount, modmod) = 0
								Progress, , % FileCount - fc_old " neue Dateien hinzugefügt (" FileCount ")"
						pages:= pdfinfo(BefundOrdner "\" A_LoopField, "Pages")
						FileGetSize, FSize, % BefundOrdner "\" A_LoopField, K
						PageSum += pages
						FileCount ++
						ScanPool.Push(A_LoopField . "|" . FSize . "|" . pages)
						continue
				}

				If (Mod(A_Index, 50) = 0)
				{
						SB_SetText(FileCount, 1)
						SB_SetText("Pdf Dateien mit " . PageSum . " Seiten", 2)
				}

		}

	;Sortieren der eingelesenen und aktualisierten Dateien
		Progress, , % "...sortiere die Dateien..."
		ScanPoolArray("Sort")
		ScanPoolArray("Save")
		Progress, , % "Der ScanPool-Index wurde aktualisiert."

	;Status ausgeben
		SB_SetText(FileCount, 1)
		SB_SetText("Pdf Dateien mit " . PageSum . " Seiten", 2)
		SB_SetText("Der ScanPool Ordner ist eingelesen`.", 4)
		Gui, BO: Show, NA

		;ObjTree()
		FileCount:= CountValidKeys(ScanPool)

return FileCount

HotkeyDebug:

		ListVars
		Pause,Toggle

return
}

NoScriptStopSleep(waitingtime) {                                                               	;-- wie sleepx - nur nicht unterbrechbar
		a := A_TickCount
		while (A_TickCount - a < waitingtime)
			sleep, 50
}

FormatSeconds(Sekunden) {

	ListLines, Off

	Return SubStr("0" . Sekunden // 3600, -1) . ":"
        . SubStr("0" . Mod(Sekunden, 3600) // 60, -1) . ":"
        . SubStr("0" . Mod(Sekunden, 60), -1)
}

WaitFileExist(filename) {	                                                                             	;-- wait until a specific file is written (not closed)
   while !FileExist(filename)
        Sleep 500
}

SleepOnRdPSession(ms, altms) {                                                                               	;-- unterschiedliche sleep Zeiten ms = wenn eine Remotedesktop-Verbindung besteht, altms = wenn nicht

	; manchmal notwendig wegen verzögertem Übertragung einer Fensterdarstellung während einer bestehenden Remotesitzung. Manche Fenster flickern lange und man kann nichts lesen.
	; wenn altms = 0 kehrt die Funktion sofort zurück

	If altms = 0
			return

	If WMIEnumProcessExist("RdClient.Windows.exe")
		sleep % ms
	else
		sleep % altms

return
}

WMIEnumProcessID(ProcessSearched) {

	ProcList:= Object()
	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
		   ProcList[process.Name]:= process.ProcessID

	return ProcList.ProcessSearched
}

WMIEnumProcesses() {                                                                                          	;--returns process list
	;==== ein deutlich besserer Ersatz für Process, Exist nur 32bit ======
	;funktioniert zuverlässiger, in der Original Library waren zudem Fehler bei den ID Bezeichnungen
	; gibt eine Liste existierender Prozesse samt PID zurück, die Liste muss nachfolgend ausgewertet werden

	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
        list .= process.ProcessID . "|" . process.Name . "`n"

    return list
}

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

SendSuspend(scriptname) {                                                                        	;-- Sendet einen Suspend-Befehl zu einem anderen Skript.

	DetectHiddenWindows, On
	WM_COMMAND := 0x111
	ID_FILE_SUSPEND := 65404
	PostMessage, WM_COMMAND, ID_FILE_SUSPEND,,, %scriptname% ahk_class AutoHotkey

}

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


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Addendum_Controls
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
;\/\/ RICHEDIT \/\/
Rich_FindText(hEdit, Text, Mode:="WHOLEWORD") {

	static EM_FINDTEXT:= 1080
	Static FR:= {DOWN: 1, WHOLEWORD: 2, MATCHCASE: 4}
	Flags := 0

	For Each, Value In Mode
         If FR.HasKey(Value)
            Flags |= FR[Value]

	Sel := RE_GetSel(hEdit)
    Min := (Flags & FR.DOWN) ? Sel.E : Sel.S
	Max := (Flags & FR.DOWN) ? -1 : 0

	VarSetCapacity(FT, 16 + A_PtrSize, 0)
	NumPut(Min	  	,   FT, 0, "Int")
	NumPut(Max	  	,   FT, 4, "Int")
	NumPut(&Text	,   FT, 8, "Ptr")

	SendMessage, EM_FINDTEXT, %Flags%, &FT,, % "ahk_id " hEdit
	S := NumGet(FTX, 8 + A_PtrSize, "Int"), E := NumGet(FTX, 12 + A_PtrSize, "Int")
	If (S = -1) && (E = -1)
         Return False

	RE_SetSel(hEdit, S, E)
	RE_ScrollCaret(hEdit)

Return ErrorLevel=4294967295 ? -1 : ErrorLevel
}
RE_FindText(hEdit, sText, cpMin=0, cpMax=-1, flags="") {
	static EM_FINDTEXT=1080,WHOLEWORD=2,MATCHCASE=4		 ;WM_USER + 56
	hFlags := 0
	loop, parse, flags, %A_Tab%%A_Space%,%A_Space%%A_Tab%
		if (A_LoopField != "")
			hFlags |= %A_LOOPFIELD%
	VarSetCapacity(FT, 12)
	NumPut(cpMin,  FT, 0)
	NumPut(cpMax,  FT, 4)
	NumPut(&sText, FT, 8)
	SendMessage, EM_FINDTEXT, hFlags, &FT,, ahk_id %hEdit%
Return ErrorLevel
}

RE_GetSel(hEdit) {                                                                                         	;-- Funktionen von HiEdit.ahk - diese funktionieren mit dem RichEdit-Control in Albis
	static EM_GETSEL=176
	VarSetCapacity(s, 4), VarSetCapacity(e, 4)
	SendMessage, EM_GETSEL, &s, &e,, ahk_id %hEdit%
	s := NumGet(s), e := NumGet(e)
Return {S: s, E: e}
}

RE_GetTextLength(hEdit) {
	static WM_GETTEXTLENGTH=14
	SendMessage, WM_GETTEXTLENGTH, 0, 0,, ahk_id %hEdit%
	Return ErrorLevel
}

RE_ReplaceSel(hEdit, text=""){
	static  EM_REPLACESEL=194
	SendMessage, EM_REPLACESEL, 0, &text,, ahk_id %hEdit%
Return ErrorLevel
}

RE_ScrollCaret(hEdit){
	static EM_SCROLLCARET=183
	SendMessage, EM_SCROLLCARET, 0, 0,, ahk_id %hEdit%
	Return ErrorLevel
}

RE_SetSel(hEdit, nStart=0, nEnd=-1) {
	static EM_SETSEL=0x0B1
	SendMessage, EM_SETSEL, nStart, nEnd,, ahk_id %hEdit%
Return ErrorLevel
}

LV_MouseGetCellPos(ByRef LV_CurrRow, ByRef LV_CurrCol, LV_LView) {

	/*                              	DESCRIPTION

			Link: https://autohotkey.com/board/topic/30486-listview-tooltip-on-mouse-hover/

	*/

	static LVIR_LABEL                           := 0x0002                                                                    	; LVM_GETSUBITEMRECT constant - get label info
	static LVM_GETITEMCOUNT      	:= 4100                                                                       	; gets total number of rows
	static LVM_SCROLL                    	:= 4116                                                                       	; scrolls the listview
	static LVM_GETTOPINDEX          	:= 4135                                                                       	; gets the first displayed row
	static LVM_GETCOUNTPERPAGE 	:= 4136                                                                       	; gets number of displayed rows
	static LVM_GETSUBITEMRECT    	:= 4152                                                                       	; gets cell width,height,x,y

	ControlGetPos	, LV_lx, LV_ly, LV_lw, LV_lh 			, , % "ahk_id" LV_LView                          	; get info on listview
	SendMessage	, LVM_GETITEMCOUNT		, 0, 0, , % "ahk_id" LV_LView
	LV_TotalNumOfRows		:= ErrorLevel                                                                                 	; get total number of rows
	SendMessage	, LVM_GETCOUNTPERPAGE	, 0, 0, , % "ahk_id" LV_LView
	LV_NumOfRows 			:= ErrorLevel                                                                                 	; get number of displayed rows
	SendMessage	, LVM_GETTOPINDEX			, 0, 0, , % "ahk_id" LV_LView
	LV_topIndex               	:= ErrorLevel                                                                                 	; get first displayed row

	mMode := A_CoordModeMouse
	CoordMode, MOUSE, RELATIVE
	MouseGetPos, LV_mx, LV_my
	LV_mx -= LV_lx, LV_my -= LV_ly
	VarSetCapacity(LV_XYstruct, 16, 0)                                                                                     	; create struct

	Loop,% LV_NumOfRows + 1                                                                                              	; gets the current row and cell Y,H
	{	LV_which := LV_topIndex + A_Index - 1                                                                         	; loop through each displayed row
		NumPut(LVIR_LABEL, LV_XYstruct, 0)                                                                               	; get label info constant
		NumPut(A_Index - 1, LV_XYstruct, 4)                                                                               	; subitem index
		SendMessage, LVM_GETSUBITEMRECT, %LV_which%, &LV_XYstruct,, ahk_id %LV_LView% 	; get cell coords
		LV_RowY 				:= NumGet(LV_XYstruct,4)                                                                 	; row upperleft y
		LV_RowY2 			:= NumGet(LV_XYstruct,12)                                                               	; row bottomright y2
		LV_currColHeight 	:= LV_RowY2 - LV_RowY                                                                    	; get cell height
		If(LV_my <= LV_RowY + LV_currColHeight)                                                                   	; if mouse Y pos less than row pos + height
		{	LV_currRow   := LV_which + 1                                                                                   	; 1-based current row
			LV_currRow0 := LV_which                                                                                          	; 0-based current row, if needed
			LV_currCol	:= 0                                                                                                     	; LV_currCol is not needed here, so I didn't do it! It will always be 0.
																																				; See my ListviewInCellEditing function for details on finding LV_currCol if needed.
			return LV_currRow
			Break
		}
	}
	CoordMode, MOUSE, % mMode

return
}

LV_GetItemText(item_index, sub_index, ctrl_id, win_id) {                         				;-- read the text from an item in a TListView

		; https://autohotkey.com/board/topic/18299-reading-listview-of-another-app/    ----  code from Tigerite
		MAX_TEXT:= 260
		item_index -= 1
        VarSetCapacity(szText	, MAX_TEXT, 0)
        VarSetCapacity(szClass	, MAX_TEXT, 0)
        ControlGet, hListView, Hwnd, , % ctrl_id, % "ahk_id " win_id
        DllCall("GetClassName", "UInt",hListView, "Str",szClass, "Int",MAX_TEXT)
        if (DllCall("lstrcmpi", "Str",szClass, "Str","SysListView32") == 0 || DllCall("lstrcmpi", "Str",szClass, "Str","TListView") == 0)
            LV_ItemText(hListView, item_index, sub_index, szText, MAX_TEXT)

return %szText%
}

LVM_GetText(h, r, c=1) {

	;https://autohotkey.com/board/topic/41650-ahk-l-60-listview-handle-library-101/
	r -= 1                                                     	; convert to 0 based index

	VarSetCapacity(t, 511, 1)
	VarSetCapacity(lvItem, A_PtrSize * 7)
	NumPut(1 	, lvItem, "uint")                   	; mask
	NumPut(r   	, lvItem, A_PtrSize, "int")      	; iItem
	NumPut(c-1	, lvItem, A_PtrSize * 2, "int") 	; iSubItem
	NumPut(&t	, lvItem, A_PtrSize * 5, "ptr") 	; pszText
	NumPut(512	, lvItem, A_PtrSize * 6)         	; cchTextMax

	If (A_IsUnicode)
		DllCall("SendMessage", "uint", h, "uint", 4211, "uint", r, "ptr", &lvItem) ; LVM_GETITEMTEXTW
	Else
		DllCall("SendMessage", "uint", h, "uint", 4141, "uint", r, "ptr", &lvItem) ; LVM_GETITEMTEXTA

Return t
}


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Addendum_Windows
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
WinToClient(hwnd, ByRef x, ByRef y) {

    WinGetPos, wx, wy,,, ahk_id %hwnd%
    VarSetCapacity(pt, 8)
    NumPut(x + wx, pt, 0)
    NumPut(y + wy, pt, 4)
    DllCall("ScreenToClient", "uint", hwnd, "uint", &pt)
    x := NumGet(pt, 0, "int")
    y := NumGet(pt, 4, "int")

}

ClientToWin(hwnd, ByRef x, ByRef y) {                                                                 	     	    			;-- Convert client co-ordinates (cx,cy) to window co-ordinates (wx,wy) - Lexikos
	;https://autohotkey.com/board/topic/24813-windows-information-is-it-possible/page-2?&#entry161431
    VarSetCapacity(pt, 8)
    NumPut(x, pt, 0)
    NumPut(y, pt, 4)
    DllCall("ClientToScreen", "uint", hwnd, "uint", &pt)
    WinGetPos, wx, wy,,, ahk_id %hwnd%
    x := NumGet(pt, 0, "int") - wx
    y := NumGet(pt, 4, "int") - wy
}

ScreenToClient(hwnd, ByRef x, ByRef y) {
    VarSetCapacity(pt, 8)
    NumPut(x, pt, 0)
    NumPut(y, pt, 4)
    DllCall("ScreenToClient", "uint", hwnd, "uint", &pt)
    x := NumGet(pt, 0, "int")
    y := NumGet(pt, 4, "int")
}

MouseGetWinTitle() {                                                                                                                 	;-- ermittelt den Fenstertitel über das der Mauszeiger steht
	MouseGetPos,x ,y , WinID, CId
	WinGetTitle, WTitle, ahk_id %WinID%
return WTitle
}

GetWindowPos(hWnd, ByRef X, ByRef Y, ByRef W, ByRef H) {
    VarSetCapacity(RECT, 16, 0)
    DllCall("GetWindowRect", "Ptr", hWnd, "Ptr", &RECT)
    DllCall("MapWindowPoints", "Ptr", 0, "Ptr", GetParent(hWnd), "Ptr", &RECT, "UInt", 2)
    X := NumGet(RECT, 0, "Int")
    Y := NumGet(RECT, 4, "Int")
    w := NumGet(RECT, 8, "Int") - X
    H := NumGet(RECT, 12, "Int") - Y
}


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Addendum_DB.ahk
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; some ways to find patient names in text by comparing against known patient data
FindName(text)                                 	{                          	; PDF renaming - Hilfsfunktion

		static ThousandYear	:= SubStr(A_YYYY, 1, 2)
		static HundredYear	:= SubStr(A_YYYY, 3, 2)
		static RxNames := [	"i)Patient[ien]*[\s;:.]*(Herrn|Herr|Frau)*\s*(?<Name>[A-Z][\pL-]+[\s,.]+[A-Z][\pL-]+)[\s,.]+((geb[.oren]*\s*(am)*)|(\*\*))\s*(?<Geb>\d{1,2}\.\d{1,2}\.\d{2,4})"
									,	"i)(Nach)name[:,;.\s]+(?<name2>[A-Z][\pL-]+).*Vo(rn|m)ame[:,;.\s]+(?<name1>[A-Z][\pL-]+)"
									,	"i)Vo(rn|m)ame[:,;.\s]+(?<name1>[A-Z][\pL-]+).*Name[:,;.\s]+(?<name2>[A-Z][\pL-]+)"]

		PatIDBuf	:= Object()

	; Methode 1:	RegExsStrings nutzen
		For RxIdx, rxString in RxNames {

				spos := 1, Next := false, PatNachname := PatVorname := ""
				while (spos := RegExMatch(text, rxString, Pat, spos)) {

					If (StrLen(PatName1) > 0) && (StrLen(PatName2) > 0) {
						spos  += StrLen(PatName2 PatName1)
					} else if (StrLen(PatName) > 0) {
						spos  += StrLen(PatName)
						PatName1 := StrSplit(PatName, "\s+").1
						PatName2 := StrSplit(PatName, "\s+").2
					}

					PatIDArr := admDB.StringSimilarityID(PatName1, PatName2)
					If IsObject(PatIDArr)
						For idx, PatID in PatIDArr {
							Next := true
							If PatIDBuf.haskey(PatID)
								PatIDBuf[PatID] += 1
							else
								PatIDBuf[PatID] := 1

							SciTEOutput("Methode 1-" RxIdx ": " (PatID) " , hits: " PatIDBuf[PatID] " ~ " PatName1 ", " PatName2)
						}

				}

				If Next
					break
		}

	; Methode 2:	Suche nach Patientennamen über gefundene Geburtstage
		spos := 1
		while (spos := RegExMatch(text, "(?<Tag>\d{1,2})[.,;](?<Monat>\d{1,2})[.,;](?<Jahr>(\d{4}|\d{2}))[\s:,;.]", D, spos)) {

			spos += StrLen(D)
			If (StrLen(DJahr) = 2)                    	; für Datumsformate 01.08.24, erstellt bei größeren Jahren als das aktuelle Jahr (zB. 20), 01.08.1924
				DJahr := (DJahr > HundredYear) ? (SubStr("00" (ThousendYear - 1), -1) . DJahr) : (SubStr("00" (ThousendYear), -1) . DJahr)

			PatIDArr := admDB.MatchID("gd", SubStr("00" DTag, -1) "." SubStr("00" DMonat, -1) "." DJahr)
			SciTEOutput(PatIDArr.1)
			If IsObject(PatIDArr)
				For idx, PatID in PatIDArr {
					If PatIDBuf.haskey(PatID)
						PatIDBuf[PatID] += 1
					else
						PatIDBuf[PatID] := 1
					SciTEOutput("Methode 2: " PatID " , hits: " PatIDBuf[PatID])
				}

		}

	; Methode 3:	Findet zwei aufeinander folgende Worte mit Großbuchstaben am Anfang (Personennamen, Eigennamen ...)
	;               	und vergleicht diese per StringSimiliarity-Algorithmus mit den Namen in der Addendum-Patientendatenbank
		spos := 1
		while (spos := RegExMatch(text, "([A-Z][\pL-]+)[\s,.]+([A-Z][\pL-]+)", name, spos)) {

				;SciTEOutput("spos: " spos)
			; überflüssige Zeichen entfernen
				name1 := RegExReplace(name1, "[\s\n\r]"),
				name2 := RegExReplace(name2, "[\s\n\r]")

			; Bindestrichworte ignorieren
				If RegExMatch(name1, "\-$") {
					spos += StrLen(name1)                    ; ein Wort weiter
					continue
				} else if RegExMatch(name2, "\-$") {
					spos += StrLen(name)                     ; beide Wörter weiter
					continue
				} else if (StrLen(name1) = 0) || (StrLen(name2) = 0) {
					spos += StrLen(name)
					continue
				}

				PatIDArr := admDB.StringSimilarityID(name1, name2)
				If IsObject(PatIDArr) {
					spos += StrLen(name)
					For idx, PatID in PatIDArr {
						If PatIDBuf.haskey(PatID)
							PatIDBuf[PatID] += 1
						else
							PatIDBuf[PatID] := 1
						SciTEOutput("Methode 3: " PatID " , hits: " PatIDBuf[PatID])
					}

				}
				else
					spos += StrLen(name1)

		}

	; Ausgabe in Konsole dabei höchste Trefferzahl ermitteln und diese Patienten ID merken
		maxHits := 0
		For PatID, hits in PatIDBuf {

			SciTEOutput("hits: " hits " - (" PatID ") " oPat[PatID].Nn ", " oPat[PatID].Vn)
			If (maxHits < hits) {
				BestPatID	:= PatID
				maxHits	:= hits
			}

		}

	; schauen ob es gleiche Trefferzahlen gibt, Nutzer muss dann die Entscheidung für den richtigen Namen vornehmen
		For PatID, hits in PatIDBuf {
			If (maxHits = hits) && !(PatID = BestPatID) {

				If IsObject(BestPatID) {
					tmpID    	:= BestPatID
					BestPatID	:= Array()
					BestPatID.Push(tmpID)
				}
				BestPatID.Push(PatID)
				SciTEOutput("Entscheidungsproblem. Gleiche Trefferzahl von " maxHits " bei PatID " BestPatID " und " PatID)

			}
		}

return BestPatID
}

RxNames(Str, RxMatch="GetNames")                      	{                       	;-- Stringfunktion für die Behandlung von PDF-Dateinamen

	; eingeführt mit Addendum_Gui.ahk
	; letzte Änderung: 07.02.2021

	static rxPerson1 	:= "[A-ZÄÖÜ][\pL]+[\s-]*([A-ZÄÖÜ][\pL-]+)*"
	static rxPerson2 	:= "[A-ZÄÖÜ][\pL]+([\-\s][A-ZÄÖÜ][\pL]+)*"

	static RxExtract   	:= "^\s*(?<NN>" rxPerson1 ")[\s,]+(?<VN>" rxPerson1 ")[\s,]+(?<DokTitel>.*)*" ;?\.[a-z]*
	static RxNames1	:= "^\s*(?<NN>" rxPerson1 ")[\s,]+(?<VN>" rxPerson1 ")[\s,]+"
	static RxNames2	:= "^\s*(?<NN>" rxPerson2 ")[\s,]+(?<VN>" rxPerson2 ")[\s,]+"

	if (RxMatch = "ContainsName")
		return RegExMatch(Str, RxNames1)
	else If (RxMatch = "GetNames") {
		RegExMatch(Str, RxExtract, Pat)
		return {"Nn":PatNN, "Vn":PatVN, "DokTitel":PatDokTitel}
	}
	else if (RxMatch ="ReplaceNames")
		return RegExReplace(Str, RxNames2)

}

EscapeStrRegEx(str)                                                 	{                      	;-- formt eine String RegEx fähig um
	For idx, Char in StrSplit("[](){}*\/|-+.?^$")
		str := StrReplace(str, char , "\" char)
return str
}

StrFlip(string)                                                         	{                         	;-- String reverse
	VarSetCapacity(new, n:=StrLen(string))
   Loop % n
      new .= SubStr(string, n--, 1)
return new
}

class string_old                                                               	{                        	;-- wird alle String Funktionen von Addendum enthalten

	; benötigt Addendum_Datum.ak
	; letzte Änderung 17.02.2021

	; RegEx Strings erstellen
		static rx := RegExStrings()

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; diverse Stringfunktionen
	;----------------------------------------------------------------------------------------------------------------------------------------------
		EscapeStrRegEx(str)                           	{      	;-- formt einen String RegEx fähig um
			For idx, Char in StrSplit("[](){}*\/|-+.?^$")
				str := StrReplace(str, char , "\" char)
		return str
		}

		Flip(string)                                       	{        	;-- String reverse
			VarSetCapacity(new, n:=StrLen(string))
			Loop % n
				new .= SubStr(string, n--, 1)
		return new
		}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Funktionen für die PDF Kategorisierung von Addendum-Gui.ahk
	;----------------------------------------------------------------------------------------------------------------------------------------------
		ContainsName(str)                         	{       	;-- enthält Personennamen?
			return RegExMatch(Str, this.rx.Names1)
		}

		GetDocDate(str)                               	{       	;-- Dokumentdatum aus dem Dateinamen erhalten

			RegExMatch(str, "(?<Date>\d{1,2}\.\d{1,2}\.\d{2,4}).?([^\d.\s\-]|$)", Doc)
			If RegExMatch(DocDate, "^\d{1,2}\.\d{1,2}\.\d{2,4}$")
				return FormatDate(DocDate, "DMY", "dd.MM.yyyy")

		return
		}

		GetNames(str)                               	{         	;-- Personennamen zurückgeben
			RegExMatch(Str, this.rx.Extract, Pat)
			return {"Nn":PatNN, "Vn":PatVN, "DokTitel":PatDokTitel}
		}

		isFullNamed(str)                            	{        	;-- Dateiname enthält Nachname,Vorname, Dokumentbezeichung und Datum

			str := RegExReplace(str, this.rx.fileExt)
			If RegExMatch(str, this.rx.CaseName, FOE) {
				Part1	:= RegExReplace(FOEN	, "[\s,]+$")
				Part2	:= RegExReplace(FOEI 	, "[\s,]+$")
				Part3	:= RegExReplace(FOED	, "^[\s,v\.]+")
				If (StrLen(Part1) = 0 || StrLen(Part2) = 0 || StrLen(Part3) = 0)
					return false
				else
					return true
			}

		return false
		}

		class Replace extends string             	{

			Names(str, rplStr:="")                   	{         	;-- Personennamen ersetzen
				return RegExReplace(Str, this.rx.Names2, rplStr)
			}

			FileExt(str, ext:="", rplStr:="")       	{        	;-- Dateiendungen ersetzen
				ext := !ext ? "[a-z]+" : ext
				return RegExReplace(Str, "\." ext "$" , rplStr)
			}

		}

}


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Addendum.ahk
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
global Script        	:= Object()
Script.Gui         	:= Object()
Script.User       	:= {"Interaction" : false, "Interface" : ""}

Wartezeit:                                     	;{

	FileReadLine, WZeit, % WZStatistikFile, 1
	FileReadLine, WPat, % WZStatistikFile, 2
		if (errorlevel = 0) {
			x = %A_ScreenWidth% - 400
			y = %A_ScreenHeight% - 120
			SplashPraxID = PraxomatInfo("Addendum für Albis on Windows", "konnte die Wartezeitstatistikdatei nicht öffnen", % x , % y, 1000)
			;SetTimer, SplashTextAus, %showtime%
		}

return
;}
UserAway:                                      	;{	                                                            				;-- macht verschiedene Dinge wenn niemand den Computer eine zeitlang benutzt

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; pausiert den Timer der nopopup_Labels nach 30sek. wenn niemand am Computer ist
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		If   	 	(A_TimeIdle > 60000) &&  !Addendum.useraway {

			Addendum.useraway := true

				If IsLabel(nopopup_%compname%)
					SetTimer, nopopup_%compname%, Off

				SetTimer, UserAway, 500                                        	;jetzt schneller nachsehen damit es nicht 5min. dauert bis die Überwachungsroutine nopopup_compname wieder anspringt

		}
		else if	(A_TimeIdle < 300)  	&&  Addendum.useraway {

				Addendum.useraway := true

				If IsLabel(nopopup_%compname%)
					SetTimer, nopopup_%compname%, On

				SetTimer, UserAway, 300000                                    	;die Überwachungsroutine wird wieder seltener aufgerufen werden

		}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; wenn ein neuer Tag angebrochen ist, zuerst das Albis Programmdatum aktualisieren und dann das Skript neu starten (verhindert Programmfehler, gleichzeitig
	; werden bei einem Skriptneustart die Addendum eigene Patientendatenbank sortiert und das Tagesprotokoll neu begonnen)
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		If !(A_DD = Addendum.AktuellerTag) {
			AlbisSetzeProgrammDatum(A_DD . "." . A_MM . "." . A_YYYY)
			SkriptReload()
		}

return
;}
DoingStatisticalThings:                     	;{

return
;}

AutoSkriptReload() {                                                                                                        	;-- Timerfunktion, startet Skript um 0 Uhr immer neu
	; Addendum.ahk setzt bei Albis das Tagesdatum um 0Uhr auf den neuen Tag
	If ( TimeDiff("000000", "now", "m")=0 )
		SkriptReload()
}

WinEventHook_Helper_Thread:                                                                                      	;{ Labordaten - Fenster abfangen

	hpop:= GetLastActivePopup(AlbisWinID())
	If (hpop != AlbisWinID()) && (hpop != hpop_old)	{

		EHWT		:= WinGetTitle(hpop)
		EHWText	:= WinGetText(hpop)
		hpop_old	:= hpop

		If InStr(EHWT, "Labordaten")		{

			SetTimer, WinEventHook_Helper_Thread, Off
			VerifiedCheck("Button5"	, hpop,,, 1)
			VerifiedClick("Button1" 	, hpop)
			AlbisLaborbuchOeffnen()

		}

	}

return
;}

WinEvent_HMV(hHook, event, hwnd, idObject, idChild, eventThread, eventTime) {           	; Hookprocedure welche nur bei geöffneten Heilmittelverordnungsformularen aktiv ist
	Critical
	HmvHookHwnd:= GetHex(hwnd)
	HmvEvent:= GetHex(event)
	SetTimer, Heilmittelverordnung_WinHandler,  -0
return 0

Heilmittelverordnung_WinHandler:                                                                                 	;{ Eventhookhandler - kümmert sich nur um Heilmittelverordnungsfenster

return ;}
}

SplashAus: ;{	; ----------------------- zugehörige Funktionen für die WinEventHandler Labels
	SplashTextOff
return ;}
			else if InStr(EHWT   	, "CGM LIFE")	                                                                                    	 {  	; Verbindungsfehler (gesperrt durch Windows Firewall)
				WinClose, % "CGM LIFE ahk_class Afx:00400000:0:00010005:86101803"
			}


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Addendum_Gui.ahk
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
FuzzyKarteikarte_O1(NameStr, OnlyPatID=false)                   	{               	;-- fuzzy name matching function, öffnet eine Karteikarte

	; mit OnlyPatID = true kann man nur einen String mit enthaltenem Patientennamen testen und bekommt bei einem Treffer die Patienten-ID zurück

		static rxPerson := "[A-ZÄÖÜ][\pL]+[\s-]*([A-ZÄÖÜ][\pL-]+)*"

		DiffSafe 	:= Array()
		minDiff		:= 100
		PatFound 	:= true
		umlauts 	:= {"Ä":"Ae", "Ü":"Ue", "Ö":"Oe", "ä":"ae", "ü":"ue", "ö":"oe", "ß":"sz"}


		If !IsObject(Pat := RxNames(NameStr, "GetNames")) {
			PraxTT(	"Die Dateibezeichnung enthält keinen Namen eines Patienten.`nDer Karteikartenaufruf wird abgebrochen!", "3 2")
			return
		}

		NVName 	:= RegExReplace(Pat.Nn Pat.Vn	, "[\s]")         	; NachnameVorname
		VNName 	:= RegExReplace(Pat.Vn Pat.Nn	, "[\s]")         	; VornameNachname

	; Stringdifference Methode mit Suche des Patienten in der Addendum Patientendatenbank     	;{
		For PatID, Pat in oPat 		{

				If InStr(PatID, "MaxPat")
					break

			; Leerzeichen und Minus aus den Namen entfernen
				DbName	:= RegExReplace(Pat.Nn Pat.Vn, "[\s]")         	; NachnameVorname

			; Stringdifferenz Methode anwenden
				DiffA	:= StrDiff(DBName, NVname), DiffB := StrDiff(DbName, VNname)
				Diff  	:= DiffA <= DiffB ? DiffA : DiffB

			 ; Treffergenauigkeit (Stringsimilarity) kleiner gleich 0.11
			 ; -> passender Patient wurde gefunden. Karteikarte des Patienten wird aufgerufen
				If (Diff <= 0.11) {
					SciTEOutput()
					If !OnlyPatID
						admGui_Karteikarte(PatID)
					else
						return {"ID":PatID, "Nn":Pat.Nn, "Vn":Pat.Vn, "Gd":Pat.Gd}
				}

			; sammelt den kleinsten Differenzwert
				If (Diff < minDiff)
					minDiff	:= Diff, bestDiff := PatID, bestNn := Pat.Nn, bestVn := Pat.Vn, bestGd := Pat.Gd

			; enthält eine der beiden Zeichenketten 2 Zeichen mehr als die andere wird die Patienten ID gepuffert
				If ( Diff <= 0.13 && Diff > 0.11 )
					If (Abs(StrLen(DbName) - StrLen(NVname)) > 2)
						DiffSafe.Push(PatID)

		}

		;PatIDs := admDB.StringSimilarityID(Pat.Nn, Pat.Vn)

		If (minDiff < 0.2) {
			If !OnlyPatID
				admGui_Karteikarte(bestDiff)
			else
				return {"ID":bestDiff, "Nn":bestNn, "Vn":bestVn, "Gd":bestGd}
		}

	; kein Treffer dann leeren String zurückgeben
		If OnlyPatID
			return "#" minDiff ", " bestDiff

		;}

	; Suchen über die Albis Patientensuche                                                                                	;{

		; Code nimmt die ersten beiden Buchstaben vom Vor- und Zuname und übergibt sie dem Patient öffnen Dialog
		; stehen mehrere Personen zur Auswahl im Dialog "Patient auswählen" wird per String Differenz der dichteste Treffer gesucht
		; wenn Albis für diese Kombination nur einen Treffer hat, wird sofort die Akte geöffnet (unabhängig vom Skript)
		; --- die umlautveränderten Namen werden hier für die Suche noch nicht genutzt!!
		PatMatches 	:= Object()

		PraxTT("Der Patient ist nicht in der Datenbank!`nVersuche es mit einer Patientensuche über den Albisdialog.", "6 0")
		Sleep, 200

	; öffnet den Dialog "Öffne Patient" und übergibt nur die ersten beiden Buchstaben vom Vor- und Zunamen
		AlbisDialogOeffnePatient("open", SubStr(name[1], 1, 2) ", " SubStr(name[2], 1, 2))

		while !(WinExist("Patient auswählen ahk_class #32770") || WinExist("ALBIS ahk_class #32770", "Patient <") || (A_Index > 40))
			Sleep 50

		; Patient ist nicht vorhanden ;{
			while (hwnd := WinExist("ALBIS ahk_class #32770", "Patient <")) && (A_Index < 10) 	{
				PatFound := false
				VerifiedClick("Button1", hwnd,,, true) ; OK
				WinWait, % "Patient öffnen ahk_class #32770",, 2
				If WinExist("Patient öffnen ahk_class #32770") {
					AlbisDialogOeffnePatient("close")
					PraxTT("Albis hat keine Patienten zur Auswahl gestellt.`nEs kann keine Suche durchgeführt werden.", "6 2")
					return false
				}
				sleep, 200
			}
		;}

		If (hwnd := WinExist("Patient auswählen ahk_class #32770")) {

			PatFound	:= true
			LVPats   	:= AlbisLVContent(hwnd, "SysListView321", "Name|Vorname|Geb.-Datum")
			For row, Patient in LVPats	{
				Patient.2 := StrReplace(Patient.2, "-")
				Patient.3 := StrReplace(Patient.3, "-")
				If (StrDiff(Patient.2 Patient.3, Name.1 Name.2) <= 0.12 ) || (StrDiff(Patient.2 Patient.3, Name.2 Name.1) <= 0.12 )
					PatMatches.Push(Patient)
			}

			If (PatMatches.MaxIndex() = 1) {

				hinweis :=  "Ich denke dieser Patient könnte es sein:`n`n" PatMatches[1][2] "," PatMatches[1][3] ", geb. am " PatMatches[1][4]
				hinweis .= "`nletzter Behandlungstag: " PatMatches[1][6] "`n`nSoll ich diese Akte öffnen?"

				Msgbox, 0x1024, Addendum für Albis on Windows, % hinweis
				IfMsgBox, No
					return

				VerifiedClick("Button2", "Patient auswählen ahk_class #32770",,, true)
				WinWait, % "Patient öffnen ahk_class #32770",, 3
				If ErrorLevel	{
					MsgBox, Der Patient öffnen Dialog fehlt mir jetzt`num weiter zu machen!
					return false
				}
				VerifiedSetText("Edit1", PatMatches[1][5], "Patient öffnen ahk_class #32770", 200)

			; Akte wird jetzt geöffnet durch drücken von OK
				while WinExist("Patient öffnen ahk_class #32770") 	{

					VerifiedClick("Button2", "Patient öffnen ahk_class #32770")	    	; Button OK drücken
					WinWaitClose, Patient öffnen ahk_class #32770,, 1
					If WinExist("Patient öffnen ahk_class #32770") {	                     	; Fenster ist immer noch da? Dann sende ein ENTER.
							WinActivate, Patient öffnen ahk_class #32770
							ControlFocus, Edit1, Patient öffnen ahk_class #32770
							SendInput, {Enter}
					}
					If (A_Index > 10) {
						PatFound := false
						break
					}
					sleep, 200

				}

			}
			else
				PatFound := false

			VerifiedClick("Button2", "Patient auswählen ahk_class #32770",,, true) ; Abbruch
			AlbisDialogOeffnePatient("close")

		}

	;}

return PatFound
}

FuzzyKarteikarte_O2(NameStr)                                             	{               	; fuzzy name matching function, öffnet eine Karteikarte

	; RegEx Variablen
		static rxPerson	:= "[A-ZÄÖÜ][\pL]+[\s-]*([A-ZÄÖÜ][\pL-]+)*"
		static umlauts	:= {"Ä":"Ae", "Ü":"Ue", "Ö":"Oe", "ä":"ae", "ü":"ue", "ö":"oe", "ß":"sz"}

	; prüft ob Namen übergeben wurden
		If !IsObject(Pat := string.GetNames(NameStr)) {
			PraxTT(	"Die Dateibezeichnung enthält keinen Namen eines Patienten.`nDer Karteikartenaufruf wird abgebrochen!", "3 2")
			return
		}

	; Namen heraussuchen
		PatMatches := admDB.StringSimilarityEx(Pat.Nn, Pat.Vn)
		If ((PatMatches.Count() -1) = 1) {
			For PatID, Pat in PatMatches
				If (PatID <> "diff") {
					admGui_Karteikarte(PatID)
					break
				}
		}
		else {
			m := PatMatches.diff
			SciTEOutput(" (matches: " PatMatches.Count() - 1 ") mindiff: " m.minDiff ", [" m.bestID "] " m.bestNn ", " m.bestVn)
			For PatID, Pat in PatMatches
				If (PatID <> "diff")
					SciTEOutput(" [" PatID "] " Pat.Nn ", " Pat.Vn)
			return
		}

	return

	; Suchen über die Albis Patientensuche                                                                                	;{

		; Code nimmt die ersten beiden Buchstaben vom Vor- und Zuname und übergibt sie dem Patient öffnen Dialog
		; stehen mehrere Personen zur Auswahl im Dialog "Patient auswählen" wird per String Differenz der dichteste Treffer gesucht
		; wenn Albis für diese Kombination nur einen Treffer hat, wird sofort die Akte geöffnet (unabhängig vom Skript)
		; --- die umlautveränderten Namen werden hier für die Suche noch nicht genutzt!!
		PatMatches 	:= Object()

		PraxTT("Der Patient ist nicht in der Datenbank!`nVersuche es mit einer Patientensuche über den Albisdialog.", "6 0")
		Sleep, 200

	; öffnet den Dialog "Öffne Patient" und übergibt nur die ersten beiden Buchstaben vom Vor- und Zunamen
		AlbisDialogOeffnePatient("open", SubStr(name[1], 1, 2) ", " SubStr(name[2], 1, 2))

		while !(WinExist("Patient auswählen ahk_class #32770") || WinExist("ALBIS ahk_class #32770", "Patient <") || (A_Index > 40))
			Sleep 50

		; Patient ist nicht vorhanden ;{
			while (hwnd := WinExist("ALBIS ahk_class #32770", "Patient <")) && (A_Index < 10) 	{
				PatFound := false
				VerifiedClick("Button1", hwnd,,, true) ; OK
				WinWait, % "Patient öffnen ahk_class #32770",, 2
				If WinExist("Patient öffnen ahk_class #32770") {
					AlbisDialogOeffnePatient("close")
					PraxTT("Albis hat keine Patienten zur Auswahl gestellt.`nEs kann keine Suche durchgeführt werden.", "6 2")
					return false
				}
				sleep, 200
			}
		;}

		If (hwnd := WinExist("Patient auswählen ahk_class #32770")) {

			PatFound	:= true
			LVPats   	:= AlbisLVContent(hwnd, "SysListView321", "Name|Vorname|Geb.-Datum")
			For row, Patient in LVPats	{
				Patient.2 := StrReplace(Patient.2, "-")
				Patient.3 := StrReplace(Patient.3, "-")
				If (StrDiff(Patient.2 Patient.3, Name.1 Name.2) <= 0.12 ) || (StrDiff(Patient.2 Patient.3, Name.2 Name.1) <= 0.12 )
					PatMatches.Push(Patient)
			}

			If (PatMatches.MaxIndex() = 1) {

				hinweis :=  "Ich denke dieser Patient könnte es sein:`n`n" PatMatches[1][2] "," PatMatches[1][3] ", geb. am " PatMatches[1][4]
				hinweis .= "`nletzter Behandlungstag: " PatMatches[1][6] "`n`nSoll ich diese Akte öffnen?"

				Msgbox, 0x1024, Addendum für Albis on Windows, % hinweis
				IfMsgBox, No
					return

				VerifiedClick("Button2", "Patient auswählen ahk_class #32770",,, true)
				WinWait, % "Patient öffnen ahk_class #32770",, 3
				If ErrorLevel	{
					MsgBox, Der Patient öffnen Dialog fehlt mir jetzt`num weiter zu machen!
					return false
				}
				VerifiedSetText("Edit1", PatMatches[1][5], "Patient öffnen ahk_class #32770", 200)

			; Akte wird jetzt geöffnet durch drücken von OK
				while WinExist("Patient öffnen ahk_class #32770") 	{

					VerifiedClick("Button2", "Patient öffnen ahk_class #32770")	    	; Button OK drücken
					WinWaitClose, Patient öffnen ahk_class #32770,, 1
					If WinExist("Patient öffnen ahk_class #32770") {	                     	; Fenster ist immer noch da? Dann sende ein ENTER.
							WinActivate, Patient öffnen ahk_class #32770
							ControlFocus, Edit1, Patient öffnen ahk_class #32770
							SendInput, {Enter}
					}
					If (A_Index > 10) {
						PatFound := false
						break
					}
					sleep, 200

				}

			}
			else
				PatFound := false

			VerifiedClick("Button2", "Patient auswählen ahk_class #32770",,, true) ; Abbruch
			AlbisDialogOeffnePatient("close")

		}

	;}

return PatFound
}

admGui_Update(wParam, lParam, Msg, hWnd)                	{                	; OnMessage Funktion soll das Neuzeichnen des Fensters verbessern

	;SciTEOutput(GetHex(wParam) "," GetHex(lParam) ", " GetHex(Msg) ", " GetHex(hwnd) ": " WinGetClass(hwnd))
	global hadm, lastOnMsg
	global fn_Redraw

	if (lastOnMsg = Msg)
		return 0
	lastOnMsg := Msg

	;fn_Redraw := Func("RedrawWindow").Bind(hadm)
	;~ fn_Redraw := Func("UpdateGui").Bind(hadm)
	;~ SetTimer, % fn_Redraw, -400

return 0
}

UpdateGui(hGui)                                                            	{                	; versucht

	WinGet, ctrList, ControlListHwnd, % "ahk_id " hGui
	ctrList := StrSplit(ctrList, "`n")
	For cIdx, hwnd in ctrList {
		If !DllCall("IsWindowVisible","Ptr", hWnd) {
			HiddenCtrls ++
			WinSet, Style, +0x10000000, % "ahk_id " hwnd
		}
	}

	WinSet, Redraw,, % "ahk_id " hwnd
	SciTEOutput("HiddenCtrls: " !HiddenCtrls ? 0 : HiddenCtrls "/" ctrList.Count())
}

toptop:                                                                                              	;{ zeichnet das Fenster in Abständen neu

	; Gui schliessen unter bestimmten Bedingungen
		Aktiv := Addendum.AktiveAnzeige := AlbisGetActiveWindowType()
		If !RegExMatch(Aktiv, "i)Karteikarte|Laborblatt|Biometriedaten") || !WinExist("ahk_class OptoAppClass") {
			admGui_Destroy()
			return
		}

	; Albis ist nicht aktiv
		If !WinActive("ahk_class OptoAppClass") {
			SetTimer, toptop, % (Addendum.iWin.aRT := !WinActive("ahk_class OptoAppClass") ? Addendum.iWin.RefreshTime*2 : Addendum.iWin.RefreshTime)
			return
		}

	; Fenster ab und zu - neu zeichnen lassen
		RedrawWindow(hadm)
		try {
			Gui, adm: Show, NA
		}

return ;}

adm_Notes:			                                                                    			;{ Notizen           	- Edit Handler

return ;}

admGui_RenameO(filename, prompt="", newfilename="") {               	; Dialog für Dokument umbenennen oder teilen

	; Variablen
		global hadm, admPV, admGui_HRenIB
		static admRen_Width 	:= 350
		static admRen_Height	:= 206
		static oPV, fileout, fileout1, fileext, newadmFile

	; PDF Vorschau anzeigen
		oPV := admGui_PDFPreview(fileName)
		If !IsObject(oPV)
			return

	; Vorbereitung Dokument zerlegen
		If Instr(prompt, "Aufteilen") {

			title    	:= "Dokument zerlegen"
			fileout	:= ""

		}
	; Vorbereitung Dokument umbenennen
		else {

			title   	:= "Dokument umbenennen"
			prompt	:= "Ändern Sie hier die Dateibezeichnung"
			SplitPath, filename,,, fileext

		; neuen Dateinamen vorbereiten
			newfilename := Trim(newfilename)
			If (StrLen(newfilename ) > 0) {

				If RegExMatch(newfilename, "\.[a-z]+$")
					SplitPath, newfilename,,,, fileout
				else
					fileout := newfilename

				fileout := RegExReplace(fileout, "\s*(Befund)*\s*\.[a-z]+\d*$")
			}
			else {

				;RegExMatch(filename, "\s*(Befund\s*)\.\w*$", fileout)	; => fileout1
				fileout := RegExReplace(filename, "\s*(Befund)*\s*\.[a-z]+\d*$")
			}

		}

	; DPI-Skalierung ausschalten
		Result := DllCall("User32\SetProcessDpiAwarenessContext", "UInt" , 1)

	; Inputfensterposition
		admPos	:= GetWindowSpot(hadm)
		RnWidth	:= admPos.W + (2 * admPos.BW) + 10
		RnWidth	:= RnWidth < admRen_Width ? admRen_Width : RnWidth

	; Inputbox anzeigen
		InputBox, newadmFile	, % title
									    	, % Trim(prompt)                                          	; prompt
											, % ""                                                          	; Hide
											, % RnWidth                                                	; Width
											, % admRen_Height                                      	; Height
											, % admPos.X + admPos.W -	 RnWidth + 5	; X
											, % admPos.Y + admPos.H                        	; Y
											, % "Locale"                                                	; Locale
											, % ""                                                          	; Timeout
											, % fileout                                                    	; Default

		IBErrLvl := ErrorLevel

	; PDF Vorschaufenster automatisch schliessen
		If (oPV.previewer = "SumatraPDF") {

			Process, Close, % oPV.PID
			If WinExist(filename)
				SumatraInvoke("Exit", oPV.ID)

		} else
			Gui, admPV: Destroy

	; Dokument zerlegen
		If Instr(prompt, "Aufteilen") {

		; split pages: 	qpdf.exe --empty --pages in.pdf 1-2 -- split1.pdf
		; 						qpdf.exe --empty --pages in.pdf 3-4 -- split2.pdf
		;	1-2 D1, 3-4 D2
		; (?<Pages>\d+-\d+)\s*D(?<Nr>\d+)

				nwfiles	:= Array()
				SplitPath, filename,,, fileext, fileout

				SciTEOutput(" ------------------------------------------------------------------------------------------------------")
				SciTEOutput("  # teile PDF: " filename)

				For idx, scmd in StrSplit(newadmFile, ",") {

					If RegExMatch(scmd, "\s*(?<Pages>\d*-*\d+)\s*D(?<Nr>\d+", Dok) {

						splitcmd := 	Addendum.PDF.qpdfPath "\qpdf.exe --empty --pages " Addendum.BefundOrdner "\" fileName
									 .  	" " DokPages " -- " Addendum.BefundOrdner "\" fileout DokNr "." fileext
						SciTEOutput("  - teile Seite(n) [" DokPage "] Dokument " fileout DokNr "." fileext " zu`n  [" splitcmd "]")
						stdout := StdoutToVar(splitcmd)
						If (StrLen(Stdout) > 0)
							SciTEOutput(ParseStdOut(stdout))

						haystack :=	"error in numeric range|"
										. 	"number \d+ out of range|"
										.	"No such file or directory|"
										.	"Permission denied|"
										.	"unexpected character"

						If !RegExMatch(stdout, haystack)
							nwfiles.Push(fileout DokNr "." fileext)

					}

				}

				SciTEOutput("  # Dokument aufgeteilt.")

			; Anzeige auffrischen
				admGui_Default("admJournal")
				For fnr, newfile in nwfiles {

					isNewFile := true
					For key, pdf in ScanPool
						If Instr(pdf.name, newfile) {
							isNewFile := false
							break
						}

					thispdf := GetFileData(Addendum.BefundOrdner, newfile)

					If isNewFile
						ScanPool[key] := thispdf
					else {

						ScanPool.Push(thispdf)
						;LV_Add
					}

				}
				admGui_Reload(false)

				return
		}
	; Dokument umbenennen
		else {

			; -------------------------------------------------------------------------------------------------------------------------------------
			; HINWEISE BEI UMBENENNUNGSFEHLERN
			; -------------------------------------------------------------------------------------------------------------------------------------;{
			; Nutzer hatte abgebrochen
				If IBErrLvl {
					PraxTT("Sie haben abgebrochen!", "2 1")
					RedrawWindow(hadm)
					return
				}
			; Nutzer hat OK oder Enter gedrückt, aber nichts geändert
				If ((newadmFile . fileout1 . fileext) = filename) {
					PraxTT("Sie haben den Dateinamen nicht geändert!", "2 1")
					RedrawWindow(hadm)
					return
				}
			; Nutzer hat alles gelöscht
				If (StrLen(newadmFile) = 0) {
					PraxTT("Ihr neuer Dateiname enthält keine Zeichen.", "2 1")
					RedrawWindow(hadm)
					return
				}
			;}

			; -------------------------------------------------------------------------------------------------------------------------------------
			; DOKUMENTE UMBENENNEN
			; -------------------------------------------------------------------------------------------------------------------------------------;{
				newadmFile := newadmFile "." fileext
				FileMove, % Addendum.BefundOrdner "\" filename, % Addendum.BefundOrdner "\" newadmFile, 1
				If ErrorLevel {
					MsgBox, 0x1024, Addendum für Albis on Windows, % 	"Das Umbenennen der Datei`n"
																									.	"  <" filename ">  `n"
																									. 	"wurde von Windows aufgrund des`n"
																									.	"Fehlercode (" A_LastError ") abgelehnt."
					RedrawWindow(hadm)
					return
				}

			; ZUGEHÖRIGES TEXTDOKUMENT UMBENENNEN
				txtfile := Addendum.BefundOrdner "\Text\" RegExReplace(filename, "\.\w+$", ".txt")
				If (fileext = "pdf") && FileExist(txtfile)
					FileMove, % txtfile, % Addendum.BefundOrdner "\Text\" RegExReplace(newadmFile, "\.\w+$", ".txt"), 1

			; BACKUP DATEI UMBENENNEN
				If (fileext = "pdf") && FileExist(Addendum.BefundOrdner "\Backup\" filename)
					FileMove, % Addendum.BefundOrdner "\Backup\" fileName, % Addendum.BefundOrdner "\Backup\" newadmFile, 1

			;}

			; ScanPool auffrischen
				For key, pdf in ScanPool
					If (pdf.name = filename) {
						ScanPool[key].name := newadmFile
						break
					}

			; Journal auffrischen
				admGui_Default("admJournal")
				Loop, % LV_GetCount() {

					LV_GetText(rFilename, row := A_Index, 1)
					If Instr(rFilename, fileName) {
						LV_Modify(row,, newadmFile)
						break
					}

				}

				RedrawWindow(hadm)

		}

return
}

ReadPatientDatabase_O1(PatDBPath) {										               				;-- liest die .csv Datei Patienten.txt als Object() ein

	PatDB	:= Object()

	If (StrLen(PatDBPath) = 0)
		PatDBPath := Addendum.DBPath "\Patienten.txt"

	If FileExist(PatDBPath)	{

		; Einlesen der Datenbank als Textliste, Sortieren aufsteigend nach PatID, Aussortieren doppelter Einträge (später neue Einträge unter den Skripten kommunizieren?)
			DBtmp := FileOpen(PatDBPath, "r").Read()
			Sort, DBtmp, N U
			SciTEOutput(StrLen(DBtmp))
			;FileOpen(PatDBPath, "w", "UTF-8").Write(DBtmp)
			DBtmp := StrSplit(DBtmp, "`n", "`r")

		; Einlesen in ein Objekt

			/*  Aufbau PatDB - Objektes

				1. PatID = key
				2. Nachname    	(Nn)
				3. Vorname     		(Vn)
				4. Geschlecht    	(Gt)
				5. Geburtsdatum 	(Gd)
				6. Krankenkasse 	(Kk)
				7. letzteGVU 		(letzteGVU)

			 */
			For DBtmpIdx, line in DBtmp 	{
				If (StrLen(line) = 0)
					continue
				Str := StrSplit(line, ";", A_Space)
				PatID := Str[1]
				PatDB[PatID] := {"Nn": Str[2], "Vn": Str[3], "Gt": Str[4], "Gd": Str[5], "Kk": Str[6]}

			}

	}

	JSONData.Save(Addendum.DBpath "\Patienten.json", PatDB, true,, 1, "UTF-8")

return PatDB
}

ReadPatientDatabase_O2(PatDBPath) {										               				;-- liest die .csv Datei Patienten.txt als Object() ein

	PatDB	:= Object()

	If (StrLen(PatDBPath) = 0)
		PatDBPath := Addendum.DBPath "\Patienten.txt"

	If FileExist(PatDBPath)	{

		; Einlesen der Datenbank als Textliste, Sortieren aufsteigend nach PatID, Aussortieren doppelter Einträge (später neue Einträge unter den Skripten kommunizieren?)
			DBtmp := FileOpen(PatDBPath, "r").Read()
			Sort, DBtmp, N U
			FileOpen(PatDBPath, "w", "UTF-8").Write(DBtmp)
			DBtmp := StrSplit(DBtmp, "`n", "`r")
			;~ FileDelete  	, % AddendumDBPath "\Patienten.txt"
			;~ FileAppend	, % DBtmp, % AddendumDBPath "\Patienten.txt", UTF-8

		; Einlesen in ein Objekt
			;~ 1.PatID = key; 2.Nachname (Nn); 3.Vorname (Vn); 4.Geschlecht (Gt); 5.Geburtsdatum (Gd); 6.Krankenkasse (Kk); 7.letzteGVU (letzteGVU)
			For DBtmpIdx, line in DBtmp 	{

				If (StrLen(line) = 0)
					continue
				Str := StrSplit(A_LoopField, ";", A_Space)
				PatID := Str[1]
				PatDB[PatID] := {"Nn": Str[2], "Vn": Str[3], "Gt": Str[4], "Gd": Str[5], "Kk": Str[6], "letzteGVU": Str[7]}
			}

			;PatDB["MaxPat"] := maxPat := PatDB.Count()

	}

return PatDB
}

admGui_RemoveImports(Imports)                                                        	; PatDocs - Importe entfernen
admGui_RemoveImports(ImportList)                               	{               	; PdfReport-Array aufräumen

	global PatDocs

	Loop, Parse, % RTrim(ImportList, "`n"), `n
		Loop % PatDocs.MaxIndex()
			If InStr(PatDocs[A_Index], A_LoopField) {
				PatDocs.RemoveAt(A_Index)
				continue
			}

}

;CheckJournal() { ...
For idx, pdf in ScanPool
	If (pdf.name = pdfFile) {
		ScanPool[idx] := GetPDFData(Addendum.BefundOrdner, pdfFile)
		break
	}

;BefundIndex() { ....
ScanPool.Push(GetPDFData(Addendum.BefundOrdner, filename))

	; Originaldatei umbennen
		FileMove, % Addendum.BefundOrdner "\" oldadmfile, % Addendum.BefundOrdner "\" (newadmfile := newadmfile "." FileExt), 1
		If ErrorLevel {
			MsgBox, 0x1024, Addendum für Albis on Windows
									, % 	"Das Umbenennen der Datei`n"
									.	"  <" oldadmFile ">  `n"
									. 	"wurde von Windows aufgrund des`n"
									.	"Fehlercode (" A_LastError ") abgelehnt."
			gosub RNGuiBeenden
			return
		}

	; Backup Datei umbennen
		If (fileext = "pdf") && FileExist(Addendum.BefundOrdner "\Backup\" oldadmfile)
			FileMove, % Addendum.BefundOrdner "\Backup\" oldadmfile, % Addendum.BefundOrdner "\Backup\" newadmfile, 1

	; Textdokument umbennen
		txtfile := Addendum.BefundOrdner "\Text\" RegExReplace(oldadmfile, "\.\w+$", ".txt")
		If (FileExt = "pdf") && FileExist(txtfile)
			FileMove, % txtfile, % Addendum.BefundOrdner "\Text\" RegExReplace(newadmfile, "\.pdf$", ".txt"), 1

			FileMove, % pdfPath	, % docPath "\" 				newadmFile ".pdf"	, 1               	; Originaldatei umbenennen
			FileMove, % bupPath	, % docPath "\Backup\" 	newadmFile ".pdf"	, 1               	; Backupdatei umbenennen
			FileMove, % txtPath	, % docPath "\Text\" 		newadmFile ".txt" 	, 1           		; zugehörige Text Datei umbenennen

		If IsObject(fDates)
			If fDates.Behandlung[1]
				newadmFile .= "v. " fDates.Behandlung[1]	" "
				else
			newadmFile .= "v. " fDates.Dokument[1] 	" "

				pdfPath  	:= docPath "\" filename
				bupPath 	:= docPath "\Backup\" filename
				txtPath   	:= docPath "\Text\" StrReplace(filename, ".pdf", ".txt")


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Addendum_Ini.ahk
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
admAlbisMenuAufrufe()                                      	; Albis Menubezeichnungen und wm_command Befehle
admAlbisMenuAufrufe() {                                                                       	;-- Menubezeichnungen und wm_command Befehle

		Addendum.AlbisMenu := Object()        	; wm_command Daten für Dialogaufrufe in Albis
		Addendum.AlbisMenu.Privatliquidation:= { "Auswahlliste": "33023"
		, "Behandlungsliste": "32891"
		, "Ausgangsbuch": "33125"
		, "Offene_Posten": "32892"
		, "Mahnliste": { "Alle": "33145"
			, "Mahnstufe_1": "33142"
			, "Mahnstufe_2": "33143"
			, "Mahnstufe_3": "33144"
			, "Mahnbescheid": "34192"
			, "Fällige": "33850"}
		, "Quittungsliste": "32893"
		, "Stornierte_Restbeträge": "34864"
		, "Journal": "32894"
		, "Stornierte": "33851"
		, "Buchungsliste": "33841"
		, "Rechnungen_Buchungen": "33842"
		, "Rechnungen_Mahnungen": "33844"
		, "Kostenplan": "34067"
		, "Faktorzuordnungen": "34155"
		, "Quittungsliste_löschen": "32895"
		, "Sachkostenaufstellung": "34080"
		, "KH-Abschlag-_und_Vorteilsausgleich": "34156"}

}


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Addendum_Protocoll.ahk
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
HashThisString(str) {
return CryptHash(&str, StrLen(str), "MD5")
}
;intern
CryptHash(pData, nSize, SID = "CRC32", nInitial = 0) {
	CALG_SHA := CALG_SHA1 := 1 + CALG_MD5 := 0x8003
	If Not CALG_%SID%
	{
		FormatI := A_FormatInteger
		SetFormat, Integer, H
		sHash := DllCall("ntdll\RtlComputeCrc32", "Uint", nInitial, "Uint", pData, "Uint", nSize, "Uint")
		SetFormat, Integer, %FormatI%
		StringUpper,	sHash, sHash
		StringReplace,	sHash, sHash, X, 000000
		Return	SubStr(sHash,-7)
	}

	DllCall("advapi32\CryptAcquireContextA", "UintP", hProv, "Uint", 0, "Uint", 0, "Uint", 1, "Uint", 0xF0000000)
	DllCall("advapi32\CryptCreateHash", "Uint", hProv, "Uint", CALG_%SID%, "Uint", 0, "Uint", 0, "UintP", hHash)
	DllCall("advapi32\CryptHashData", "Uint", hHash, "Uint", pData, "Uint", nSize, "Uint", 0)
	DllCall("advapi32\CryptGetHashParam", "Uint", hHash, "Uint", 2, "Uint", 0, "UintP", nSize, "Uint", 0)
	VarSetCapacity(HashVal, nSize, 0)
	DllCall("advapi32\CryptGetHashParam", "Uint", hHash, "Uint", 2, "Uint", &HashVal, "UintP", nSize, "Uint", 0)
	DllCall("advapi32\CryptDestroyHash", "Uint", hHash)
	DllCall("advapi32\CryptReleaseContext", "Uint", hProv, "Uint", 0)

	FormatI := A_FormatInteger
	SetFormat, Integer, H
	Loop,	%nSize%
		sHash .= SubStr(*(&HashVal + A_Index - 1), -1)
	SetFormat, Integer, %FormatI%
	StringReplace,	sHash, sHash, x, 0, All
	StringUpper,	sHash, sHash
	Return	sHash
}


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Addendum_Internal.ahk
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ProcessExist(ProcName, cmd:="") {                                                                                            	;-- sucht nur nach Autohotkeyprozessen

	; use cmd = "PID" to receive the PID of an Autohotkey process
	q := Chr(0x22)

	StrQuery := "Select * from Win32_Process Where Name Like 'AutoHotkey%'"
	For Process in winmgmts.ExecQuery(StrQuery)	{
		RegExMatch(Process.CommandLine, "\w+\.ahk(?=\" q ")", name)
		If InStr(name, ProcName)
			If StrLen(cmd) = 0
				return true
			else
				return Process[cmd]
	}

return false
}


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Addendum_PDFReader.ahk
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
RGBColorToBGR(vColor) {												; function: convert RGB 6 character hex color code to BGR
	; example:  "8e3e2d" -> "2d3e8e"
  If (StrLen(vColor) = 6)
    Return SubStr(vColor,5,2) SubStr(vColor,3,2) SubStr(vColor,1,2)
}

SetSystemCursor(Cursor := "") {										; function: SetSystemCursor() set custom system cursor or restore default cursor

	; - no parameter     : restore default cursor
	; - parameter "CROSS": change custom system cursor to cross
	; original function by Flipeador , https://www.autohotkey.com/boards/viewtopic.php?p=206703#p206703

	Static Cursors := {APPSTARTING: 32650, ARROW: 32512, CROSS: 32515, HAND: 32649, HELP: 32651, IBEAM: 32513
                    , NO: 32648, SIZEALL: 32646, SIZENESW: 32643, SIZENS: 32645
                    , SIZENWSE: 32642, SIZEWE: 32644, UPARROW: 32516, WAIT: 32514}

  If (Cursor == "")
    ; Restore default cursors
    Return DllCall("User32.dll\SystemParametersInfoW", "UInt", 0x0057, "UInt", 0, "Ptr", 0, "UInt", 0)

  ; Replace default cursors with custom cursor
  Cursor := InStr(Cursor, "3") ? Cursor : Cursors[Cursor]
  For Each, ID in Cursors
  {
    ; 2 = IMAGE_CURSOR | 0x00008000 = LR_SHARED
    hCursor := DllCall("User32.dll\LoadImageW", "Ptr", 0, "Int", Cursor, "UInt", 2, "Int", 0, "Int", 0, "UInt", 0x00008000, "Ptr")
    hCursor := DllCall("User32.dll\CopyIcon", "Ptr", hCursor, "Ptr")
    DllCall("User32.dll\SetSystemCursor", "Ptr", hCursor, "UInt",  ID)
  }
} ; https://msdn.microsoft.com/en-us/library/windows/desktop/ms648395(v=vs.85).aspx







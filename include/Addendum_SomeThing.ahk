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







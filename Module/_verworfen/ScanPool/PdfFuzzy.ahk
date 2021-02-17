;{ 1. Scripteinstellungen / Includes / Tray Menu

	#NoEnv
	#Persistent
	#MaxMem 4095	; INCREASE MAXIMUM MEMORY ALLOWED FOR EACH VARIABLE - NECESSARY FOR THE INDEXING VARIABLE
	#SingleInstance off

	SendMode Input
	;SetWorkingDir %A_ScriptDir%
	SetTitleMatchMode, 2
	DetectHiddenWindows, On
	DetectHiddenText, On
	SetControlDelay -1
	SetWinDelay, -1
	SetBatchLines -1
	CoordMode, Mouse, Screen
	CoordMode, Pixel, Screen
	CoordMode, Menu, Screen
	CoordMode, Caret, Screen
	CoordMode, Tooltip, Screen
	FileEncoding, UTF-8

	#Include %A_ScriptDir%\..\..\..\include\AddendumFunctions.ahk
	;#Include %A_ScriptDir%\..\..\..\include\FindText.ahk
	#include %A_ScriptDir%\..\..\..\include\GDIP_All.ahk
	#include %A_ScriptDir%\..\..\..\include\XA.ahk
	#include %A_ScriptDir%\..\..\..\include\LVA.ahk
	#include %A_ScriptDir%\..\..\..\include\WMCopyData.ahk
	#include %A_ScriptDir%\lib\ScanPool_PdfHelper.ahk


;}

;{ 2. Variblen Setup / Registry auslesen

	;-------------------------------------------------------------------- globale Variablen -----------------------------------------------------------------------
	;------------------------------------------------------------------------------------------------------------------------------------------------------------------
	global AddendumDir, AlbisWinID 					;ACHTUNG: diese sollten in jedem Skript global sein
	global BefundOrdner									;BefundOrdner - der Ordner in dem sich die ungelesenen PdfDateien befinden

	;------------------------------------------------------- Ermitteln der beiden wichtigsten Pfade ------ ------------------------------------------------------
		AddendumDir := RegReadUniCode64("HKEY_LOCAL_MACHINE", "SOFTWARE\Addendum für AlbisOnWindows", "ApplicationDir")
		PdfDataPath:= A_ScriptDir . "`\PdfData\"
		PdfIndexPath:= A_ScriptDir . "`\PdfData\_BefundOrdner.xml"
	;-------------------------------------------------------- diverse Variablen setzen oder definieren ----------------------------------------------------------
	;--------------------------------------------------------------------------------------------------------------------------------------------------------------------

;}

		Name1 = %1%
		Name2 = %2%
		PdfPath = %3%
		ToolTip % SPID "`n" Name1 "`n" Name2 "`n" PdfPath


		PdfName:= StrReplace(PdfPath, "`,", " ")
		PdfName:= StrReplace(PdfName, "`;", " ")
		PdfName:= StrReplace(PdfName, ":", " ")
		PdfName:= StrReplace(PdfName, "`\", " ")
		PdfName:= StrReplace(PdfName, "   ", " ")
		PdfName:= StrReplace(PdfName, "  ", " ")
		PdfName:= StrReplace(PdfName, " ", "_")
		PdfName:= SubStr(PdfName, 1, StrLen(PdfName)-4)

		LastPage:= PdfPageCount(PdfPath)

		q := Chr(0x22)

		XpdfPath := AddendumDir . "\include\pdfk\pdftotext.exe" 			;. q
		CmdString := XpdfPath . " -f 1 -l " . LastPage . " -nopgbrk -nodiag " . q . PdfPath . q . " " q . "-" . q 					;A_ScriptDir . "\pdfData\" . Substr("00000000" . indexnr, -6) . "-Pdf.txt" . q ;" -hide"	; . q
		PdfText:= StdOutToVar(CmdString)

		Condensed:="", wC1:=0, wC2:=0
		Loop, Parse, PdfText, `r`n
					{
							if (A_LoopField="")
										continue

							line:= StrReplace(A_LoopField, "`.", " ")
							line:= StrReplace(line, "`,", " ")
							line:= StrReplace(line, "`;", " ")
							line:= StrReplace(line, "`:", " ")
							line:= StrReplace(line, "`#", " ")
							line:= StrReplace(line, "`+", " ")
							line:= StrReplace(line, "`:", " ")
							line:= StrReplace(line, "^", " ")
							line:= StrReplace(line, "`/", " ")
							line:= StrReplace(line, "`\", " ")
							line:= StrReplace(line, "   ", " ")
							line:= StrReplace(line, "  ", " ")

							condensed.= Trim(line) . " "
					}

					;Condensed:= RegExReplace(Condensed, "\.", ". ")
					;Condensed:= RegExReplace(Condensed, ":", ": ")
					;Condensed:= RegExReplace(Condensed, ";", "; ")
					;Condensed:= RegExReplace(Condensed, "#", " # ")

		FileAppend % Condensed, %A_ScriptDir%\PdfData\%PdfName%.txt

		Loop, Parse, Condensed, %A_Space%
		{
			ToolTip, % "Count: " A_Index ", " Name1 ", " Name2 ", PdfName: " PdfName ", PdfPath: " PdfPath, 2000, 100, 12 		;"`nPages: " LastPage "`nCondensed Text: `n"  Substr(Condensed,1, 50), 2000, 100, 12

			If (StrReplace(PdfName, " ", "") = "")
								continue

			;FileAppend, % A_LoopField "`n", %A_ScriptDir%\pdfer.txt
			a:= StrDiff(A_LoopField, Name1)
			b:= StrDiff(A_LoopField, Name2)

			if (a<0.2) {
					lst1.= A_LoopField . ", "
					;FileAppend, % Name1 ": " a "  `[ " A_LoopField " `]`n", %A_ScriptDir%\PdfData\%PdfName%-sift.txt
					wC1++
			}

			if (b<0.2) {
					lst2.= A_LoopField . ", "
					;FileAppend, % Name2 ": " b "  `[ " A_LoopField " `]`n", %A_ScriptDir%\PdfData\%PdfName%-sift.txt
					wC2++
			}

		}

		If (wC1 AND wC2) {

					MsgBox, % "Folgende relavanten Übereinstimmungen wurden gefunden zu: " Name1 "`n" lst1 "`n`nund zu: " Name2 "`n" lst2
					TargetScriptTitle:= "ScanPool - Befunde ahk_class AutoHotkeyGUI"

					result := Send_WM_COPYDATA(Name1 . "|" . lst1 . "|" . Name2 . "|" . lst2 , TargetScriptTitle)
					if result = FAIL
					  MsgBox SendMessage failed. Does the following WinTitle exist?:`n%TargetScriptTitle%
					else if result = 0
						MsgBox Message sent but the target window responded with 0, which may mean it ignored it.

	}


ExitApp


/*
    ObjRegisterActive(Object, CLSID, Flags:=0)

        Registers an object as the active object for a given class ID.
        Requires AutoHotkey v1.1.17+; may crash earlier versions.

    Object:
            Any AutoHotkey object.
    CLSID:
            A GUID or ProgID of your own making.
            Pass an empty string to revoke (unregister) the object.
    Flags:
            One of the following values:
              0 (ACTIVEOBJECT_STRONG)
              1 (ACTIVEOBJECT_WEAK)
            Defaults to 0.

    Related:
        http://goo.gl/KJS4Dp - RegisterActiveObject
        http://goo.gl/no6XAS - ProgID
        http://goo.gl/obfmDc - CreateGUID()
*/
ObjRegisterActive(Object, CLSID, Flags:=0) {
    static cookieJar := {}
    if (!CLSID) {
        if (cookie := cookieJar.Remove(Object)) != ""
            DllCall("oleaut32\RevokeActiveObject", "uint", cookie, "ptr", 0)
        return
    }
    if cookieJar[Object]
        throw Exception("Object is already registered", -1)
    VarSetCapacity(_clsid, 16, 0)
    if (hr := DllCall("ole32\CLSIDFromString", "wstr", CLSID, "ptr", &_clsid)) < 0
        throw Exception("Invalid CLSID", -1, CLSID)
    hr := DllCall("oleaut32\RegisterActiveObject"
        , "ptr", &Object, "ptr", &_clsid, "uint", Flags, "uint*", cookie
        , "uint")
    if hr < 0
        throw Exception(format("Error 0x{:x}", hr), -1)
    cookieJar[Object] := cookie
}
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;                                    Addendum Patienten
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;      	Funktion: 				 	einfacher Suchfilter für die Suche nach Adress-, Telefon und wenigen anderen Daten
;
;		Abhängigkeiten:		siehe includes
;
;      	begonnen:       	    	03.05.2021
; 		letzte Änderung:	 	18.09.2021
;
;	  	Addendum für Albis on Windows by Ixiko started in September 2017
;      	- this file runs under Lexiko's GNU Licence
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣

  ; Einstellungen
	#NoEnv
	#Persistent
	#KeyHistory, Off

	SetBatchLines    	, -1
	ListLines             	, Off
	SetWinDelay        	, -1
	SetControlDelay   	, -1

	global PatDB            	; Patientendatenbank
	global Addendum

  ; startet Windows Gdip
   	If !(pToken:=Gdip_Startup()) {
		MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
		ExitApp
	}

  ; Tray Icon erstellen
	If (hIconPatienten := Create_Patienten_ico())
    	Menu, Tray, Icon, % "hIcon: " hIconPatienten

  ; Albis Datenbankpfad / Addendum Verzeichnis
	RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
	RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)

	Addendum                 	:= Object()
	Addendum.Dir            	:= AddendumDir
	Addendum.Ini              	:= AddendumDir "\Addendum.ini"
	Addendum.DBPath      	:= AddendumDir "\logs'n'data\_DB"
	Addendum.DBasePath  	:= AddendumDir "\logs'n'data\_DB\DBase"
	Addendum.AlbisDBPath	:= AlbisPath "\DB"
	Addendum.compname	:= StrReplace(A_ComputerName, "-")                                                                 	; der Name des Computer auf dem das Skript läuft
	Addendum.propsPath 	:= A_ScriptDir "\Coronaimpfdokumentation.json"
	SciTEOutput()

  ; Patientendatenbank
	global outfilter 	:= ["NR", "PRIVAT", "GESCHL", "NAME", "VORNAME", "GEBURT", "Alter", "PLZ", "ORT", "STRASSE", "HAUSNUMMER", "TELEFON", "TELEFON2", "TELEFAX", "ARBEIT", "LAST_BEH", "MORTAL"]
	global Table 		:= ["NR"    	, "PR" 	, "Geschl."	, "Name"	, "Vorname"	, "Geb." 	, "Alter" 	, "PLZ"   	, "Ort"	, "Strasse"	, "Nr"    	, "Telefon1"    	, "Telefon2"   	, "Fax", "Beruf", "LBeh.", "Verst."]
	global colType	:= ["Integer"	, "Text"	, "Text"   	, "Text"  	, "Text"      	, "Integer"	, "Integer"	, "Integer"	, "Text"	, "Text"   	, "Integer"	, "Text Logical"	, "Text Logical"	, "Text Logical", "Text", "Text", "Text"]
	global colWidth := [42, 25, 25, 95, 98, 66, 40, 48, 125, 136,	34, 112, 120, 100, 160, 66, 66]
	PatDB 	:= ReadPatientDBF(Addendum.AlbisDBPath, outfilter, outfilter)

  ; Gui aufrufen
	Patienten()

return
^!a::Reload

Patienten() {

		global

		today := A_DD "." A_MM "." A_YYYY
		fields := {}
		For i, col in table
			cols .= (i>1?"|":"") col
		For i, col in outfilter
			fields[col] := i
		For i, width in colWidth
			maxwidth += width

	; Datum und Geschlechterbez. wird in lesbares Format geändert
		For PatID, data in PatDB
			For field, fieldvalue in data {
				If (RegExMatch(field, "(GEBURT|LAST_BEH|DEL_DATE|MORTAL)") && fieldvalue)
					data[field] := ConvertDBASEDate(fieldvalue)
				else if (field = "GESCHL")
					data[field] := fieldvalue = 1 ? "m":"w"
				else if (field = "PRIVAT")
					data[field] := fieldvalue = "t" ? "P":""
			}

		For PatID, data in PatDB
			data.Alter := Age(data.Geburt, today)

		;~ outfilter.InsertAt(7, "Alter")

	; GUI zeichen
		Gui, PAT: new, % "-DPIScale +HWNDhPAT +Resize MinSize" maxwidth+50 "x300"
		Gui, PAT: Color, c20AA45, c50EE95
		For i, width in colWidth
			Gui, PAT: Add, Edit, % (i = 1 ? "xm" : "x+0") " ym w" (i = 1 ? width+2 : width) " r1 AltSubmit vPATSF" i " gSUCHFELD"
		Gui, PAT: Add, Text   	, % "x+2 yp+3 Backgroundtrans vPATNFO", % "0      "
		Gui, PAT: Add, ListView	, % "xm y+4 w" maxwidth+30 " h800 Grid vPATLV gLVHandler", % cols
		Gui, PAT: Show, AutoSize, Patienten

	; Spaltenbreite anpassen
		For i, width in colWidth
			LV_ModifyCol(i, width " " colType[i])

	; Focus auf den Nachnamen
		GuiControl, PAT: Focus, PATSF4

return

SUCHFELD:   	;{

		Gui, PAT: Default
		Gui, PAT: ListView, PATLV

	; Eingaben nach Feldern zusammenstellen
		search := {}
		Gui, PAT: Submit, NoHide
		Loop, % outfilter.Count() {
			str := "PatSF" A_Index
			str := %str%
			If (str)
				search[outfilter[A_Index]] := (str)
		}

	; Listview löschen, neuzeichen anhalten
		GUIControl, PAT: -Redraw, PATLV
		LV_Delete()

	; Daten suchen
		For PatID, data in PatDB {
			LVCol 	:= Object(), LVCol.1 := PatID, found := 0
			For field, fieldvalue in data {
				LVCol[fields[field]] := fieldvalue
				If search.haskey(field) {

					If (field = "LAST_BEH") && RegExMatch(fieldvalue, search[field])
						found ++
					else if InStr(fieldvalue, search[field])
						found ++

				}
			}
			If (found = search.Count())
				LV_Add("", LVCol*)

		}

		GUIControl, PAT: +Redraw, PATLV
		GuiControl, PAT:, PATNFO, % LV_GetCount()

return ;}

LVHandler: ;{

return ;}

PATGuiSize:    	;{

	Critical, Off	; erst Critical Off soll Critical On dann schneller machen oder zuverlässiger machen (hab vergessen woher ich das habe)
	if A_EventInfo = 1
		return
	Critical
	PATw := A_GuiWidth, PATh:= A_GuiHeight
	;GuiControl, PAT: -Redraw 		, PATLV
	GuiControl, PAT: MoveDraw	, PATLV, % "w" PATW-20 " h" PATH-35
	;GuiControl, PAT: +Redraw 	, PATLV
	WinSet, Redraw,, % "ahk_id " hPAT

return ;}

PATGUIClose:	;{
PATGUIEscape:

	wqs  	:= GetWindowSpot(hQS)
	winSize := "x" wqs.X " y" wqs.Y " w" wqs.CW " h" wqs.CH
	IniWrite, % winSize, % Addendum.Ini, % Addendum.compname ,% "Patienten_Fenstergroesse"

ExitApp ;}
}

Create_Patienten_ico() {

VarSetCapacity(B64, 7508 << !!A_IsUnicode)
B64 := "AAABAAEAJDAAAAEAGADoFQAAFgAAACgAAAAkAAAAYAAAAAEAGAAAAAAAAAAAAGQAAABkAAAAAAAAAAAAAAD29/X////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////29/X9/v1EkwVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBGlAj8/fv7/PpHlQlAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBFlAdNmBFNmBFLlw9AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBKlw34+vf5+/hKlw1AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBKlw1ZoCJfoymMvWb////////s9OVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBNmRL2+PT3+fVMmBBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB7s0/4+/X///////////+NvWf////////s9OWIumDy+O7C3K5oqDZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBSmxjy9fD09vNPmhRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBxrUL+//7///////////+NvWf////////s9OWbxnr////////7/Pl1sEdAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBWnh3w8+3y9fBSmxhAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCGuV3v9un///////+NvWf////////u9eiUwXD////////////s9OZEkwVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBZoCLs8erw9O5UnRtAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBFlAdlpzJzrkSMvWb////////1+fJNmBGNvWf7/fr///////9xrUFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBeoijp7ufu8utXnh9AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCMvWb////////9/vxBkQFAkQDa6c3///////+AtlZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBipC3n6+Pr7uhcoSVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBLlw+nzIn///////////9rqjmHul/8/fv///////9lpzJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBlpzLk6OHk6eFlpjFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBDkwSQv2vX6Mr+//7////////////////////////////////h7tdCkgNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBsqjvg5dzf5Ntuqz1AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBHlQrS5cP////////////////////////////////////////3+vRuqz1AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB4sUvZ3dXa3tZ3sUpAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCkyoX////////////////////////////////////8/fvO471jpS9AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCEuFvR1s7U2NCBt1dAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDq8+P////////E3bBuqz2jyoT////////x9+xgpCtGlAhAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCQv2vLzsnO0cuKvGNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkgP///////////9ZoCJAkQCMvWb////////s9OVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCcxnvExcPHycaZxHdAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDx9+z////////D3K9YnyCMvWb////////s9OVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCrzpC9vb28vLyv0JVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCz05n////////////q8+OMvWb////////s9OWQv2v9/vzX6Mp0r0ZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDE2rOvr6+vr6/E2rNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBRmxfm8d7////////8/fuMvWb////////s9OWex37////////9/v2Hul9AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDd6NSgoKCioqLa59BAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBPmhSrz4/q8+Pr8+SMvWb////////s9OWUwXD////////////3+vRJlgxAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDz9vGQkJCUlJTt8+lAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkgKMvWb////////s9OVEkwWPvmn///////////9zrkRAkQBAkQBAkQBAkQBAkQBAkQBAkQBKlw36+/kAAACIiIj4+vdKlw1AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBJlgxHlQpAkQBAkQCHul/y+O7y+O7g7dZAkQBAkQDp8uH///////9/tVRAkQBAkQBAkQBAkQBAkQBAkQBAkQBlpzLm6+MAAAAAAADb39d5skxAkQBAkQBAkQBAkQBAkQBAkQBAkQBGlAi82ab9/v3+//7i7ti616OVwnGDt1lzrkRzrkR2sEijyoT+//7///////9rqjlAkQBAkQBAkQBAkQBAkQBAkQBAkQCXxHTMz8oAAAAAAAC+vr2916hAkQBAkQBAkQBAkQBAkQBAkQBAkQCaxXj////////9/vzv9un////////////////////////////////////l8NxCkgNAkQBAkQBAkQBAkQBAkQBAkQBDkwTl7d+kpKQAAAAAAACMjIzu8ut2sEhAkQBAkQBAkQBAkQBAkQBAkQCex33////////I4LZeoij////////////////////////////////q8+NipS5AkQBAkQBAkQBAkQBAkQBAkQBDkwSz0prY29UAAAAAAAAAAAAAAACioqLx9O6iyIJOmRNAkQBAkQBAkQBAkQBDkwSTwW/E3bDX6MrX6MrZ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"
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return -1
VarSetCapacity(Dec, DecLen, 0)
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return -2
hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
DllCall("gdiplus\GdipCreateHICONFromBitmap", A_PtrSize ? "UPtr" : "UInt", pBitmap, A_PtrSize ? "UPtr*" : "uint*", hIcon)
DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
Return hIcon
}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Includes
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
#Include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DATUM.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PDFHelper.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PraxTT.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Window.ahk

#Include %A_ScriptDir%\..\..\lib\acc.ahk
#Include %A_ScriptDir%\..\..\lib\class_JSON.ahk
#Include %A_ScriptDir%\..\..\lib\class_LV_Colors.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\lib\Sift.ahk
;}





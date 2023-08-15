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
; 		letzte Änderung:	 	31.05.2022
;
;	  	Addendum für Albis on Windows by Ixiko started in September 2017
;      	- this file runs under Lexiko's GNU Licence
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣

  ; Einstellungen
	#NoEnv
	#Persistent
	#KeyHistory, Off

	SetBatchLines            	, -1
	ListLines                    	, Off
	SetWinDelay               	, -1
	SetControlDelay          	, -1
	SetKeyDelay		        	, 20
	SendMode	    	    	, Input
	SetTitleMatchMode    	, 2
	SetTitleMatchMode     	, Fast
	DetectHiddenWindows  	, On

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

  ; Addendum Verzeichnis
	RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)
	workini := IniReadExt(AddendumDir "\Addendum.ini")

	Addendum                 	:= Object()
	Addendum.Dir            	:= AddendumDir
	Addendum.Ini              	:= AddendumDir "\Addendum.ini"
	Addendum.DBPath      	:= AddendumDir "\logs'n'data\_DB"
	Addendum.DBasePath  	:= AddendumDir "\logs'n'data\_DB\DBase"
	Addendum.AlbisPath 	:= GetAlbisPath()                                   	; Albis Basispfad
	Addendum.AlbisDBPath	:= Addendum.AlbisPath "\db"
	Addendum.compname	:= StrReplace(A_ComputerName, "-")                                           	; der Name des Computer auf dem das Skript läuft
	Addendum.propsPath 	:= A_ScriptDir "\Coronaimpfdokumentation.json"
	Addendum.Default     	:= Object()
	Addendum.Default := { "Font"         	: IniReadExt("Addendum"	, "StandardFont"                        	)
	                            	,   "BoldFont"  	: IniReadExt("Addendum"	, "StandardBoldFont"                    	)
	                            	,   "FontSize"   	: IniReadExt("Addendum"	, "StandardFontSize"                    	)
	                            	,   "FntColor"   	: IniReadExt("Addendum"	, "DefaultFntColor"                     	)
	                            	,   "BGColor"   	: IniReadExt("Addendum"	, "DefaultBGColor"                     	)
	                            	,   "BGColor1" 	: IniReadExt("Addendum"	, "DefaultBGColor1"                   	)
	                            	,   "BGColor2" 	: IniReadExt("Addendum"	, "DefaultBGColor2"                   	)
	                            	,   "BGColor3"	: IniReadExt("Addendum"	, "DefaultBGColor3"                   	)
	                            	,   "PRGColor" 	: IniReadExt("Addendum"	, "DefaultProgressColor", "FFFFFF"	)}

  ; Patientendatenbank
	global outfilter 	:= ["NR", "PRIVAT", "GESCHL", "NAME", "VORNAME", "GEBURT", "Alter"
								, "PLZ", "ORT", "STRASSE", "HAUSNUMMER", "TELEFON", "TELEFON2", "TELEFAX", "ARBEIT", "LAST_BEH", "MORTAL"]
	global Table 		:= ["NR"    	, "PR" 	, "Geschl."	, "Name"   	, "Vorname"	, "Geb." 	, "Alter" 	, "PLZ"   	, "Ort"	, "Strasse"
								, "Nr"    	, "Telefon1"       	, "Telefon2"  	, "Fax"        	, "Beruf" 	, "LBeh."	, "Verst."]
	global colType	:= ["Integer"	, "Text"	, "Text"   	, "Text"      	, "Text"      	, "Integer"	, "Integer"	, "Integer"	, "Text"	, "Text"   	, "Integer"
								, "Text Logical"	, "Text Logical"	, "Text Logical", "Text", "Text", "Text"]
	global colWidth := [42, 25, 25, 95, 98, 66, 40, 48, 125, 136,	34, 112, 120, 100, 160, 66, 66]
	PatDB 	:= ReadPatientDBF(Addendum.AlbisDBPath, outfilter, outfilter)

  ; Gui aufrufen
	Patienten()

return

; Reload Hotkey
^!a::Reload

; - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Patienten()                                                                     	{

		global

		static PatVorname, PatNachname, PatNR

		today := A_DD "." A_MM "." A_YYYY

	; fields ist das Tabellen-Objekt
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

	; GUI zeichen
		Gui, PAT: new, % "-DPIScale +HWNDhPAT +Resize MinSize" maxwidth+50 "x300"
		Gui, PAT: Margin	, 5 , 5
		Gui, PAT: Color	, % "c" "355C7D" , % "c" "99B898"

		Gui, PAT: Font	, s12 cWhite
		Gui, PAT: Add	, Text     	, % "xm ym Backgroundtrans vPATSText1"        	, % "Suche:      ⮛ nur in bestimmten Feldern ⮛"
		Gui, PAT: Add	, Text     	, % "x+100 ym Backgroundtrans vPATSText2" 	, % "⮚ in allen Feldern ⮚"

		Gui, PAT: Font	, s9 cBlack
		Gui, PAT: Add	, Edit	     	, % "x+1 ym w600 r1 hwndhPATSAF vPATSAF gSUCHFELD"    ; AltSubmit
		WinSet, ExStyle, 0x0 , % "ahk_id " hPATSAF
		EM_SetMargins(hPATSAF, 2, 2)
		GuiControlGet, cp, Pat: Pos, PatSAF
		GuiControl, Pat: Move, PATSAF, % "h" cpH-3
		y := cpY + cpH - 3 + 2

		For i, width in colWidth {
			Gui, PAT: Add, Edit	     	, % (i = 1 ? "xm y" y : "x+1") " w" (i = 1 ? width : width-1) " r1  hwndhwnd vPATSF" i " gSUCHFELD" ; AltSubmit
			WinSet, ExStyle, 0x0 , % "ahk_id " hwnd
			EM_SetMargins(Hwnd, 2, 2)
			GuiControlGet, cp, Pat: Pos, % "PATSF" i
			GuiControl, Pat: Move, % "PATSF" i, % "h" cpH-3
	   }

		Gui, PAT: Font	, s10 cWhite
		Gui, PAT: Add	, Text      	, % "x+2 yp+3 Backgroundtrans vPATNFO", % "00000 "

		Gui, PAT: Font	, s9 cBlack
		GuiControlGet, dp, Pat: Pos, PatSF1
		Gui, PAT: Add	, ListView	, % "xm y" dpY+dpH+1 " w" maxwidth+50 " h800 Grid hwndPAThLV vPATLV gLVHandler", % cols
		WinSet, ExStyle, 0x0 , % "ahk_id " PAThLV

		LV_Colors.Attach(PAThLV, 1) ; static mode
		LV_Colors.SelectionColors("0x3399FF"	, "0xFFFFFF")
		LV_Colors.AlternateRows("0x79A277" 	, "0x0")
		LV_Colors.OnMessage()

		GuiControlGet, cp, Pat: Pos, PATSText1
		GuiControlGet, dp, Pat: Pos, PatSF1
		GuiControl, Pat: Move, PATSText1, % "y" dpY - 0 - cpH

		GuiControlGet, cp, Pat: Pos, PATSText1
		GuiControlGet, dp, Pat: Pos, PATSText2
		GuiControlGet, ep, Pat: Pos, PatSF10
		GuiControlGet, fp, Pat: Pos, PatSF17
		GuiControl, Pat: Move, PatSAF, % "x" epX " w" fpX+fpW-epX
		GuiControl, Pat: Move, PATSText2, % "x" epX-dpW-2 " y" cpY

		GuiControlGet, dp, Pat: Pos, Edit10
		BorderYH := dpY+dpH+2+5

		Gui, PAT: Show	, AutoSize NoActivate NA, Patienten

	; Spaltenbreite anpassen
		For i, width in colWidth
			LV_ModifyCol(i, width " " colType[i])

	; Focus auf den Nachnamen
		GuiControl, PAT: Focus, PATSF4

	; ahk_class AutoHotkeyGUI
		Hotkey, IfWinActive, Patienten ahk_class AutoHotkeyGUI
		Hotkey, LControl & LButton, Namensenden

		OnClipboardChange("CheckTelNumber")

return

SUCHFELD:   	        	;{

		Gui, PAT: Default
		Gui, PAT: ListView, PATLV

	; 400ms warten vor dem Suchen
		Loop 40
			Sleep 10

	; Eingaben nach Feldern zusammenstellen
		search := {}
		Gui, PAT: Submit, NoHide
		Loop, % outfilter.Count() {
			str := "PatSF" A_Index
			If (str := %str%)
				search[outfilter[A_Index]] := str
		}
		If (search.Count()=0 && !PatSAF)
			return

	; PATSAF ist für die Suche in allen Spalten
		searchAll 	:= ""
		PatSAF  	:= RegExReplace(PatSAF, "i)[^\s\p{L}\d]+")
		PatSAF  	:= RegExReplace(PatSAF, "\s{2,}", " ")
		searchAll 	:= StrSplit(PatSAF, A_Space)

	; Listview löschen, neuzeichen anhalten
		GUIControl, PAT: -Redraw, PATLV
		LV_Delete()

	; Daten suchen
	; benutzt die Daten aus den Feldern Telefon|Fax als Suchmuster im Suchstring (erspart das Untersuchern auf Vorwahlnummern)
		only1   	:= only2 := 0
		For PatID, data in PatDB {

			LVCol        	:= Object()
			LVCol.1    	:= PatID
			foundB      	:= foundA := 0
			LastFoundA 	:= LastFoundB := 0
			onlyRowData	:= false

			For field, fieldvalue in data {

				LVCol[fields[field]] := fieldvalue
				If !onlyRowData {

					If RegExMatch(field, "i)(TELEFON|FAX)")
						onlydigits := true, fieldvalue := RegExReplace(fieldvalue, "[^\d]+")
					else
						onlydigits := false

					If search.haskey(field) {
						If (field = "LAST_BEH") && RegExMatch(fieldvalue, search[field])
							foundB ++
						else if InStr(fieldvalue, search[field])
							foundB ++
					}

					If IsObject(searchAll) {
						If (foundA < searchAll.Count())
							For wpos, word in searchAll {

								If (foundA >= searchAll.Count())
									break

								word 	:= onlydigits ? RegExReplace(word, "[^\d]+") : word
								fvalLen	:= StrLen(fieldvalue)
								wdLen 	:= StrLen(word)
								fAhow 	:= 0

								If (fvalLen >= wdLen && wdLen > 0 && InStr(fieldvalue, word))
									foundA ++,fAhow := 1
								else If (onlydigits && wdLen >= fvalLen && fvalLen > 0 && InStr(word, fieldvalue))
									foundA ++,fAhow := 2

							}

					}

					If (foundB+foundA > 0 && foundB=search.Count() && foundA=searchAll.Count())
						onlyRowData := true

				}

			}

			only1 += onlyRowData ? 1 : 0
			only2 += !onlyRowData ? 1 : 0

			If (foundB+foundA > 0 && foundB=search.Count() && foundA=searchAll.Count())
				LV_Add("", LVCol*)

		}

		GUIControl, PAT: +Redraw, PATLV
		GuiControl, PAT:, PATNFO, % LV_GetCount()

	; ein Treffer
	 If (LV_GetCount() = 1) {
		LV_GetText(PatNR          	, 1, 1)
		LV_GetText(PatNachname	, 1, 4)
		LV_GetText(PatVorname 	, 1, 5)
		LV_GetText(PatGeb        	, 1, 6)
		If Telsearch {
			MouseGetPos, mx, my
			ToolTip, % "Telefonnummer gehört zu:`n[" PatNR "] " PatNachname ", " PatVorname " *" PatGeb, % mx, % my-25, 2
			SetTimer, Toff, -10000
		}
	}
	else {
		PatVorname := PatNachname := PatGeb := PatNR := ""
		If Telsearch {
			MouseGetPos, mx, my
			ToolTip, % "Telefonnummer im Clipboard: ist unbekannt", % mx, % my-25, 2
			SetTimer, Toff, -5000
		}
	}

	Telsearch := false

return ;}

LVHandler:                	;{

	If (A_GuiEvent = "DoubleClick") {
		LV_GetText(PatID, A_EventInfo, 1)
		AlbisAkteOeffnen("", PatID)
	}

return ;}

Namensenden:           	;{

	MouseGetPos,,, hMouseOverWin
	mouseWinTitle 	:= WinGetTitle(hMouseOverWin)
	mouseWinClass	:= WinGetClass(hMouseOverWin)

	If (mouseWinTitle <> "Patienten" &&  mouseWinClass <> "AutoHotkeyGUI") {
		SendInput, {LControl Down}
		SendInput, {LButton Down}
		SendInput, {LButton Up}
		SendInput, {LControl Up}
		SendInput, {LControl Up}
		return
	}

	If !PatVorname || !PatNachname
		return
	Send, % "{Raw}" PatVorname
		Sleep 150
		SendInput, {Tab}
		Sleep 200
	Send, % "{Raw}" PatNachname
		Sleep 150
		SendInput, {Tab}
		Sleep 200
		If !WinActive("Kontakt verwalten ahk_class #32770") {
			SendInput, {Tab}
			Sleep 200
		}
	Send, % "{Raw}[" PatNR "]"

return ;}

PATGuiSize:    	        	;{ 25+12

	Critical, Off	; erst Critical Off soll Critical On dann schneller machen oder zuverlässiger machen (hab vergessen woher ich das habe)
	EInfo := A_EventInfo = 1 ? return : A_EventInfo
	Critical
	PATw := A_GuiWidth, PATh:= A_GuiHeight
	GuiControl, PAT: MoveDraw	, PATLV, % "w" PATw-10 " h" PATh-BorderYH
	WinSet, Redraw,, % "ahk_id " hPAT

return ;}

PATGUIClose:	        	;{
PATGUIEscape:

	wqs  	:= GetWindowSpot(hQS)
	winSize := "x" wqs.X " y" wqs.Y " w" wqs.CW " h" wqs.CH
	IniWrite, % winSize, % Addendum.Ini, % Addendum.compname ,% "Patienten_Fenstergroesse"

ExitApp ;}
}

CheckTelNumber(ctyp) {

	global hPATSAF, TelSearch := false

	CoordMode, Mouse	, Screen
	CoordMode, ToolTip	, Screen

	ClipWait, 1
	clip := Clipboard
	If RegExMatch(clip, "\s*(?<Vorwahl>0\d+[\-\/\s]|\(0\s*\d+\s*\))*(?<Nr>[\d\s]+)\s*$", tel) {
		cp := GetWindowSpot(hPATSAF)
		ToolTip, % "Telefonnummer im Clipboard entdeckt", % cp.X, % cp.Y-25, 2
		SetTimer, Toff, -5000
		GuiControl, PAT:, PatSAF, % telVorwahl . telNr
		TelSearch := true
	}

return
Toff:
 ToolTip,,,, 2
return
}

; - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -


; - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
EM_SetMargins(Hwnd, Left := "", Right := "")                   	{
   ; EM_SETMARGINS = 0x00D3 -> http://msdn.microsoft.com/en-us/library/bb761649(v=vs.85).aspx
   Set := 0 + (Left <> "") + ((Right <> "") * 2)
   Margins := (Left <> "" ? Left & 0xFFFF : 0) + (Right <> "" ? (Right & 0xFFFF) << 16 : 0)
   Return DllCall("User32.dll\SendMessage", "Ptr", HWND, "UInt", 0x00D3, "Ptr", Set, "Ptr", Margins, "Ptr")
}

Class LV_Colors                                                            	{

	; ===================================================
	; Namespace:      LV_Colors
	; Function:       Individual row and cell coloring for AHK ListView controls.
	; Tested with:    AHK 1.1.23.05 (A32/U32/U64)
	; Tested on:      Win 10 (x64)
	; Changelog:
	;     1.1.04.01/2016-05-03/just me - added change to remove the focus rectangle from focused rows
	;     1.1.04.00/2016-05-03/just me - added SelectionColors method
	;     1.1.03.00/2015-04-11/just me - bugfix for StaticMode
	;     1.1.02.00/2015-04-07/just me - bugfixes for StaticMode, NoSort, and NoSizing
	;     1.1.01.00/2015-03-31/just me - removed option OnMessage from __New(), restructured code
	;     1.1.00.00/2015-03-27/just me - added AlternateRows and AlternateCols, revised code.
	;     1.0.00.00/2015-03-23/just me - new version using new AHK 1.1.20+ features
	;     0.5.00.00/2014-08-13/just me - changed 'static mode' handling
	;     0.4.01.00/2013-12-30/just me - minor bug fix
	;     0.4.00.00/2013-12-30/just me - added static mode
	;     0.3.00.00/2013-06-15/just me - added "Critical, 100" to avoid drawing issues
	;     0.2.00.00/2013-01-12/just me - bugfixes and minor changes
	;     0.1.00.00/2012-10-27/just me - initial release
	; ===================================================
	; CLASS LV_Colors
	;
	; The class provides six public methods to set individual colors for rows and/or cells, to clear all colors, to
	; prevent/allow sorting and rezising of columns dynamically, and to deactivate/activate the message handler for
	; WM_NOTIFY messages (see below).
	;
	; The message handler for WM_NOTIFY messages will be activated for the specified ListView whenever a new instance is
	; created. If you want to temporarily disable coloring call MyInstance.OnMessage(False). This must be done also before
	; you try to destroy the instance. To enable it again, call MyInstance.OnMessage().
	;
	; To avoid the loss of Gui events and messages the message handler might need to be set 'critical'. This can be
	; achieved by setting the instance property 'Critical' ti the required value (e.g. MyInstance.Critical := 100).
	; New instances default to 'Critical, Off'. Though sometimes needed, ListViews or the whole Gui may become
	; unresponsive under certain circumstances if Critical is set and the ListView has a g-label.
	; ===================================================
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; META FUNCTIONS ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ===================================================
   ; __New()         Create a new LV_Colors instance for the given ListView
   ; Parameters:     HWND        -  ListView's HWND.
   ;                 Optional ------------------------------------------------------------------------------------------
   ;                 StaticMode  -  Static color assignment, i.e. the colors will be assigned permanently to the row
   ;                                contents rather than to the row number.
   ;                                Values:  True/False
   ;                                Default: False
   ;                 NoSort      -  Prevent sorting by click on a header item.
   ;                                Values:  True/False
   ;                                Default: True
   ;                 NoSizing    -  Prevent resizing of columns.
   ;                                Values:  True/False
   ;                                Default: True
   ; ===================================================
   __New(HWND, StaticMode := False, NoSort := True, NoSizing := True) {
      If (This.Base.Base.__Class) ; do not instantiate instances
         Return False
      If This.Attached[HWND] ; HWND is already attached
         Return False
      If !DllCall("IsWindow", "Ptr", HWND) ; invalid HWND
         Return False
      VarSetCapacity(Class, 512, 0)
      DllCall("GetClassName", "Ptr", HWND, "Str", Class, "Int", 256)
      If (Class <> "SysListView32") ; HWND doesn't belong to a ListView
         Return False
      ; ----------------------------------------------------------------------------------------------------------------
      ; Set LVS_EX_DOUBLEBUFFER (0x010000) style to avoid drawing issues.
      SendMessage, 0x1036, 0x010000, 0x010000, , % "ahk_id " . HWND ; LVM_SETEXTENDEDLISTVIEWSTYLE
      ; Get the default colors
      SendMessage, 0x1025, 0, 0, , % "ahk_id " . HWND ; LVM_GETTEXTBKCOLOR
      This.BkClr := ErrorLevel
      SendMessage, 0x1023, 0, 0, , % "ahk_id " . HWND ; LVM_GETTEXTCOLOR
      This.TxClr := ErrorLevel
      ; Get the header control
      SendMessage, 0x101F, 0, 0, , % "ahk_id " . HWND ; LVM_GETHEADER
      This.Header := ErrorLevel
      ; Set other properties
      This.HWND := HWND
      This.IsStatic := !!StaticMode
      This.AltCols := False
      This.AltRows := False
      This.NoSort(!!NoSort)
      This.NoSizing(!!NoSizing)
      This.OnMessage()
      This.Critical := "Off"
      This.Attached[HWND] := True
   }
   ; ===================================================
   __Delete() {
      This.Attached.Remove(HWND, "")
      This.OnMessage(False)
      WinSet, Redraw, , % "ahk_id " . This.HWND
   }
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; PUBLIC METHODS ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ===================================================
   Clear(AltRows := False, AltCols := False) {
	   ; Clear()         Clears all row and cell colors.
	   ; Parameters:     AltRows     -  Reset alternate row coloring (True / False)
	   ;                                Default: False
	   ;                 AltCols     -  Reset alternate column coloring (True / False)
	   ;                                Default: False
	   ; Return Value:   Always True.
	   ; ===================================================
      If (AltCols)
         This.AltCols := False
      If (AltRows)
         This.AltRows := False
      This.Remove("Rows")
      This.Remove("Cells")
      Return True
   }
   AlternateRows(BkColor := "", TxColor := "") {
	   ; ===================================================
	   ; AlternateRows() Sets background and/or text color for even row numbers.
	   ; Parameters:     BkColor     -  Background color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
	   ;                                Default: Empty -> default background color
	   ;                 TxColor     -  Text color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
	   ;                                Default: Empty -> default text color
	   ; Return Value:   True on success, otherwise false.
	   ; ===================================================
      If !(This.HWND)
         Return False
      This.AltRows := False
      If (BkColor = "") && (TxColor = "")
         Return True
      BkBGR := This.BGR(BkColor)
      TxBGR := This.BGR(TxColor)
      If (BkBGR = "") && (TxBGR = "")
         Return False
      This["ARB"] := (BkBGR <> "") ? BkBGR : This.BkClr
      This["ART"] := (TxBGR <> "") ? TxBGR : This.TxClr
      This.AltRows := True
      Return True
   }
   AlternateCols(BkColor := "", TxColor := "") {
	   ; ===================================================
	   ; AlternateCols() Sets background and/or text color for even column numbers.
	   ; Parameters:     BkColor     -  Background color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
	   ;                                Default: Empty -> default background color
	   ;                 TxColor     -  Text color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
	   ;                                Default: Empty -> default text color
	   ; Return Value:   True on success, otherwise false.
	   ; ===================================================
      If !(This.HWND)
         Return False
      This.AltCols := False
      If (BkColor = "") && (TxColor = "")
         Return True
      BkBGR := This.BGR(BkColor)
      TxBGR := This.BGR(TxColor)
      If (BkBGR = "") && (TxBGR = "")
         Return False
      This["ACB"] := (BkBGR <> "") ? BkBGR : This.BkClr
      This["ACT"] := (TxBGR <> "") ? TxBGR : This.TxClr
      This.AltCols := True
      Return True
   }
   SelectionColors(BkColor := "", TxColor := "") {
	   ; ===================================================
	   ; SelectionColors() Sets background and/or text color for selected rows.
	   ; Parameters:     BkColor     -  Background color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
	   ;                                Default: Empty -> default selected background color
	   ;                 TxColor     -  Text color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
	   ;                                Default: Empty -> default selected text color
	   ; Return Value:   True on success, otherwise false.
	   ; ===================================================
      If !(This.HWND)
         Return False
      This.SelColors := False
      If (BkColor = "") && (TxColor = "")
         Return True
      BkBGR := This.BGR(BkColor)
      TxBGR := This.BGR(TxColor)
      If (BkBGR = "") && (TxBGR = "")
         Return False
      This["SELB"] := BkBGR
      This["SELT"] := TxBGR
      This.SelColors := True
      Return True
   }
   ; ===================================================
   ; Row()           Sets background and/or text color for the specified row.
   ; Parameters:     Row         -  Row number
   ;                 Optional ------------------------------------------------------------------------------------------
   ;                 BkColor     -  Background color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
   ;                                Default: Empty -> default background color
   ;                 TxColor     -  Text color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
   ;                                Default: Empty -> default text color
   ; Return Value:   True on success, otherwise false.
   ; ===================================================
   Row(Row, BkColor := "", TxColor := "") {
      If !(This.HWND)
         Return False
      If This.IsStatic
         Row := This.MapIndexToID(Row)
      This["Rows"].Remove(Row, "")
      If (BkColor = "") && (TxColor = "")
         Return True
      BkBGR := This.BGR(BkColor)
      TxBGR := This.BGR(TxColor)
      If (BkBGR = "") && (TxBGR = "")
         Return False
      This["Rows", Row, "B"] := (BkBGR <> "") ? BkBGR : This.BkClr
      This["Rows", Row, "T"] := (TxBGR <> "") ? TxBGR : This.TxClr
      Return True
   }
   ; ===================================================
   ; Cell()          Sets background and/or text color for the specified cell.
   ; Parameters:     Row         -  Row number
   ;                 Col         -  Column number
   ;                 Optional ------------------------------------------------------------------------------------------
   ;                 BkColor     -  Background color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
   ;                                Default: Empty -> row's background color
   ;                 TxColor     -  Text color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
   ;                                Default: Empty -> row's text color
   ; Return Value:   True on success, otherwise false.
   ; ===================================================
   Cell(Row, Col, BkColor := "", TxColor := "") {
      If !(This.HWND)
         Return False
      If This.IsStatic
         Row := This.MapIndexToID(Row)
      This["Cells", Row].Remove(Col, "")
      If (BkColor = "") && (TxColor = "")
         Return True
      BkBGR := This.BGR(BkColor)
      TxBGR := This.BGR(TxColor)
      If (BkBGR = "") && (TxBGR = "")
         Return False
      If (BkBGR <> "")
         This["Cells", Row, Col, "B"] := BkBGR
      If (TxBGR <> "")
         This["Cells", Row, Col, "T"] := TxBGR
      Return True
   }
   ; ===================================================
   ; NoSort()        Prevents/allows sorting by click on a header item for this ListView.
   ; Parameters:     Apply       -  True/False
   ;                                Default: True
   ; Return Value:   True on success, otherwise false.
   ; ===================================================
   NoSort(Apply := True) {
      If !(This.HWND)
         Return False
      If (Apply)
         This.SortColumns := False
      Else
         This.SortColumns := True
      Return True
   }
   ; ===================================================
   ; NoSizing()      Prevents/allows resizing of columns for this ListView.
   ; Parameters:     Apply       -  True/False
   ;                                Default: True
   ; Return Value:   True on success, otherwise false.
   ; ===================================================
   NoSizing(Apply := True) {
      Static OSVersion := DllCall("GetVersion", "UChar")
      If !(This.Header)
         Return False
      If (Apply) {
         If (OSVersion > 5)
            Control, Style, +0x0800, , % "ahk_id " . This.Header ; HDS_NOSIZING = 0x0800
         This.ResizeColumns := False
      }
      Else {
         If (OSVersion > 5)
            Control, Style, -0x0800, , % "ahk_id " . This.Header ; HDS_NOSIZING
         This.ResizeColumns := True
      }
      Return True
   }
   ; ===================================================
   ; OnMessage()     Adds/removes a message handler for WM_NOTIFY messages for this ListView.
   ; Parameters:     Apply       -  True/False
   ;                                Default: True
   ; Return Value:   Always True
   ; ===================================================
   OnMessage(Apply := True) {
      If (Apply) && !This.HasKey("OnMessageFunc") {
         This.OnMessageFunc := ObjBindMethod(This, "On_WM_Notify")
         OnMessage(0x004E, This.OnMessageFunc) ; add the WM_NOTIFY message handler
      }
      Else If !(Apply) && This.HasKey("OnMessageFunc") {
         OnMessage(0x004E, This.OnMessageFunc, 0) ; remove the WM_NOTIFY message handler
         This.OnMessageFunc := ""
         This.Remove("OnMessageFunc")
      }
      WinSet, Redraw, , % "ahk_id " . This.HWND
      Return True
   }
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; PRIVATE PROPERTIES  ++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   Static Attached := {}
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; PRIVATE METHODS +++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   On_WM_NOTIFY(W, L, M, H) {
      ; Notifications: NM_CUSTOMDRAW = -12, LVN_COLUMNCLICK = -108, HDN_BEGINTRACKA = -306, HDN_BEGINTRACKW = -326
      Critical, % This.Critical
      If ((HCTL := NumGet(L + 0, 0, "UPtr")) = This.HWND) || (HCTL = This.Header) {
         Code := NumGet(L + (A_PtrSize * 2), 0, "Int")
         If (Code = -12)
            Return This.NM_CUSTOMDRAW(This.HWND, L)
         If !This.SortColumns && (Code = -108)
            Return 0
         If !This.ResizeColumns && ((Code = -306) || (Code = -326))
            Return True
      }
   }
   ; -------------------------------------------------------------------------------------------------------------------
   NM_CUSTOMDRAW(H, L) {
      ; Return values: 0x00 (CDRF_DODEFAULT), 0x20 (CDRF_NOTIFYITEMDRAW / CDRF_NOTIFYSUBITEMDRAW)
      Static SizeNMHDR := A_PtrSize * 3                  ; Size of NMHDR structure
      Static SizeNCD := SizeNMHDR + 16 + (A_PtrSize * 5) ; Size of NMCUSTOMDRAW structure
      Static OffItem := SizeNMHDR + 16 + (A_PtrSize * 2) ; Offset of dwItemSpec (NMCUSTOMDRAW)
      Static OffItemState := OffItem + A_PtrSize         ; Offset of uItemState  (NMCUSTOMDRAW)
      Static OffCT :=  SizeNCD                           ; Offset of clrText (NMLVCUSTOMDRAW)
      Static OffCB := OffCT + 4                          ; Offset of clrTextBk (NMLVCUSTOMDRAW)
      Static OffSubItem := OffCB + 4                     ; Offset of iSubItem (NMLVCUSTOMDRAW)
      ; ----------------------------------------------------------------------------------------------------------------
      DrawStage := NumGet(L + SizeNMHDR, 0, "UInt")
      , Row := NumGet(L + OffItem, "UPtr") + 1
      , Col := NumGet(L + OffSubItem, "Int") + 1
      , Item := Row - 1
      If This.IsStatic
         Row := This.MapIndexToID(Row)
      ; CDDS_SUBITEMPREPAINT = 0x030001 --------------------------------------------------------------------------------
      If (DrawStage = 0x030001) {
         UseAltCol := !(Col & 1) && (This.AltCols)
         , ColColors := This["Cells", Row, Col]
         , ColB := (ColColors.B <> "") ? ColColors.B : UseAltCol ? This.ACB : This.RowB
         , ColT := (ColColors.T <> "") ? ColColors.T : UseAltCol ? This.ACT : This.RowT
         , NumPut(ColT, L + OffCT, "UInt"), NumPut(ColB, L + OffCB, "UInt")
         Return (!This.AltCols && !This.HasKey(Row) && (Col > This["Cells", Row].MaxIndex())) ? 0x00 : 0x20
      }
      ; CDDS_ITEMPREPAINT = 0x010001 -----------------------------------------------------------------------------------
      If (DrawStage = 0x010001) {
         ; LVM_GETITEMSTATE = 0x102C, LVIS_SELECTED = 0x0002
         If (This.SelColors) && DllCall("SendMessage", "Ptr", H, "UInt", 0x102C, "Ptr", Item, "Ptr", 0x0002, "UInt") {
            ; Remove the CDIS_SELECTED (0x0001) and CDIS_FOCUS (0x0010) states from uItemState and set the colors.
            NumPut(NumGet(L + OffItemState, "UInt") & ~0x0011, L + OffItemState, "UInt")
            If (This.SELB <> "")
               NumPut(This.SELB, L + OffCB, "UInt")
            If (This.SELT <> "")
               NumPut(This.SELT, L + OffCT, "UInt")
            Return 0x02 ; CDRF_NEWFONT
         }
         UseAltRow := (Item & 1) && (This.AltRows)
         , RowColors := This["Rows", Row]
         , This.RowB := RowColors ? RowColors.B : UseAltRow ? This.ARB : This.BkClr
         , This.RowT := RowColors ? RowColors.T : UseAltRow ? This.ART : This.TxClr
         If (This.AltCols || This["Cells"].HasKey(Row))
            Return 0x20
         NumPut(This.RowT, L + OffCT, "UInt"), NumPut(This.RowB, L + OffCB, "UInt")
         Return 0x00
      }
      ; CDDS_PREPAINT = 0x000001 ---------------------------------------------------------------------------------------
      Return (DrawStage = 0x000001) ? 0x20 : 0x00
   }
   ; -------------------------------------------------------------------------------------------------------------------
   MapIndexToID(Row) { ; provides the unique internal ID of the given row number
      SendMessage, 0x10B4, % (Row - 1), 0, , % "ahk_id " . This.HWND ; LVM_MAPINDEXTOID
      Return ErrorLevel
   }
   ; -------------------------------------------------------------------------------------------------------------------
   BGR(Color, Default := "") { ; converts colors to BGR
      Static Integer := "Integer" ; v2
      ; HTML Colors (BGR)
      Static HTML := {AQUA: 0xFFFF00, BLACK: 0x000000, BLUE: 0xFF0000, FUCHSIA: 0xFF00FF, GRAY: 0x808080, GREEN: 0x008000
                    , LIME: 0x00FF00, MAROON: 0x000080, NAVY: 0x800000, OLIVE: 0x008080, PURPLE: 0x800080, RED: 0x0000FF
                    , SILVER: 0xC0C0C0, TEAL: 0x808000, WHITE: 0xFFFFFF, YELLOW: 0x00FFFF}
      If Color Is Integer
         Return ((Color >> 16) & 0xFF) | (Color & 0x00FF00) | ((Color & 0xFF) << 16)
      Return (HTML.HasKey(Color) ? HTML[Color] : Default)
   }
}

Create_Patienten_ico()                                                   	{

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
; - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -


; - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Includes
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
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
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#include %A_ScriptDir%\..\..\lib\class_loaderbar.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\lib\Sift.ahk
;}





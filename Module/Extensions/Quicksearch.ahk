; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                 	Addendum Quicksearch  - - geschrieben zur visuellen Analyse und Untersuchung von dBase Datenbankdateien (Albis on Windows)
;
;
;
;      Funktionen:           	- Anzeige von Daten und Suche in allen Datenbanken
;									- auslesen spezieller Informationen der Strukturen der Datenbanken
;									- Inhalte der Datenbank können im .csv Format exportiert werden
;
;		Hinweis:				alpha Version, Nutzung auf eigenes Risiko!
;
;		Abhängigkeiten:	siehe includes
;
;	                    			Addendum für Albis on Windows
;                        			by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;									begin: 02.04.2021,	last change 24.06.2021
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
/*  Notizen

	PATEXTRA - POS:  	EMail 				                    			= 93, 7
								Chroniker?                        				= 83
															                    		= 84
								Anmerkungen/weitere Info.				= 0
															                    		= 95
								Geb.Datum   	                    			= 99
								Ausnahmeindikation                     		= 97

*/

; Einstellungen                                                                         	;{
	#NoEnv
	#Persistent
	#SingleInstance, Off
	#KeyHistory, Off

	SetBatchLines                  	, -1
	ListLines                        	, On
	SetWinDelay                    	, -1
	SetControlDelay            	, -1
	AutoTrim                       	, On
	FileEncoding                 	, UTF-8
;}

; Variablen Albis Datenbankpfad / Addendum Verzeichnis      	;{

	global Props, PatDB, adm

  ; Pfad zu Albis und Addendum ermitteln
	RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
	RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)

  ; Addendum Objekt
	adm                 	:= Object()
	adm.Dir            	:= AddendumDir
	adm.Ini              	:= AddendumDir "\Addendum.ini"
	adm.DBPath      	:= AddendumDir "\logs'n'data\_DB"
	adm.AlbisPath    	:= AlbisPath
	adm.AlbisDBPath 	:= adm.AlbisPath "\DB"
	adm.compname	:= StrReplace(A_ComputerName, "-")                                                                 	; der Name des Computer auf dem das Skript läuft

  ; Nutzereinstellungen werden beim Erstellen des Objektes automatisch geladen
	Props := new Quicksearch_Properties(A_ScriptDir "\resources\Quicksearch_UserProperties.json")

  ; Patientendaten laden
	filter   	:= ["NR", "PRIVAT", "GESCHL", "NAME", "VORNAME", "GEBURT",  "PLZ", "ORT"
					, "STRASSE", "HAUSNUMMER", "TELEFON", "TELEFON2", "TELEFAX", "ARBEIT", "LAST_BEH", "MORTAL"]
	PatDB 	:= ReadPatientDBF(adm.AlbisDBPath,,, "allData")
;}

	;~ SciTEOutput(" - - - - - - - - - - - - - - - - - ")
	;~ SciTEOutput("PatDB: " PatDB.Count())
	QuickSearch_Gui()

return

^!+::Reload
#IfWinActive Quicksearch
Esc::ExitApp
#IfWinActive

Quicksearch(dbname, SearchString) {

		; {"KUERZEL": "lko"}, ["DATUM", "PATNR", "INHALT"]
		; (COVID|COVIPC|Sars|Covid|Cov\-)
		global dbfiles, dbfdata
		pattern := Object()

		For idx, line in StrSplit(SearchString, "`n") {
			RegExMatch(Trim(line), "(?<field>.*)?\=(?<rx>.*)$", m)
			If mrx
				pattern[mfield] := "rx:" mrx
			t .= mfield " = " pattern[mfield] "`n"
		}

		;SciTEOutput("`npattern: " pattern.count() "`n" t)

		dbf       	:= new DBASE(adm.AlbisDBPath "\" dbname ".dbf", 2)
		filepos   	:= dbf.OpenDBF()
		dbfdata  	:= dbf.Search(pattern, 0, "QuickSearch_Progress")
		filepos   	:= dbf.CloseDBF()

		Gui, QS: ListView, QSDBR
		LV_Delete()

	; Spalten zusammenstellen
		For rowNr, set in dbfdata {

			; Spalten zusammenstellen
				cols := Object()
				For flabel, value in set {
					If Props.isVisible(fLabel) {
						If (dbname = "BEFUND") && RegExMatch(flabel, "(INHALT|KUERZEL)") && StrLen(value) = 0
							continue
						cols[Props.HeaderVirtualPos(flabel)] := value
					}
					If (flabel = "removed")
						removed := value
				}

				If removed
					cols[1] := "*" cols[1]

			; Zeile hinzufügen
				LV_Add("", cols*)
		}

		ColTyp := Array()
		DataType 	:= {"N":"Integer", "D":"Integer", "C":"Text", "L":"Text", "M":"Text"}
		For idx, flabel in Props.Labels()
			If Props.isVisible(fLabel) {
				ftype := DataType[(Props.HeaderField(flabel, "type"))]
				LV_ModifyCol(idx, fType)
			}

		;~ JSONData.Save(A_Temp "\Quicksearch.json", matches, true,, 2, "UTF-8")
		;~ Run, % A_Temp "\Quicksearch.json"


}

QuickSearch_Gui() {

	;-: VARIABLEN                                              	;{

	  ; GUI
		global 	hQS, QSDB, QSDBFL, QSDBS, QSDBR, QSPGS, QSPGST, QST1, QST2, QST3, QST4
		global	QSMinKB      	, QSMaxKB    		, QSFilter                                       	; Eingabefelder
		global	QSMinKBText	, QSMaxKBText		, QSFilterText                                  	; Textfelder
		global 	QShDB         	, QShDBFL    		, QShDBR                                         ; hwnd
		global	QSFeldWahlT	, QSFeldWahl		, QSFeldAlle			, QSFeldKeine     	, QSFelderT
		global	QSSearch			, QSSaveRes			, QSSaveAs			, QSSaveSearch	, QSQuit			, QSReload                	; Buttons
		global 	QSBlockNR		, QSBlockSize		, QSLBlock			, QSNBlock
		global 	QSEXTRANFO 	, QSEXTRAS		, QSXTRT
		global 	BlockNr			, BlockSize		    	, lastBlockSize		, lastBlockNr			, dbfRecords
		global 	QSmrgX			, QSmrgY				, QSLvX, QSLvY, QSLvW, QSLvH
		global 	ReloadGui := false


	  ; ANDERES
		global dbfiles, DBColW, DBFilesW, dbname

		static	QSBlockNRBez, QSEXTRASx
		static EventInfo, empty := "- - - - - - -"
		static FSize            	:= (A_ScreenHeight > 1080 ? 9 : 9)
		static gcolor          	:= "805D3C"
		;static tcolor           	:= "E4E4C4"
		static tcolor           	:= "FFE1C4"
		static DBFieldW    	 	:= 550
		static DBSearchW   	:= 300
		static GFR              	:= " gQuickSearch_DBFilter "
		static gQSDBFields	:= ""
		static BGT              	:= " BackgroundTrans "
		static ASM              	:= " AltSubmit "
		static ColorTitles     	:= " cWhite"
		static ColorData      	:= " cBlack"
				 DBFilesW   	  	:= 310

		BlockNr := 0
		dbstructsPath := A_ScriptDir "\resources\AlbisDBase_Structs.json"


		If !FileExist(dbstructsPath) {
			dbfiles  	:= DBASEStructs(adm.AlbisDBPath)
			JSONData.Save(dbstructsPath, dbfiles, true,, 1, "UTF-8")
		} else {
			dbfiles := JSONData.Load(dbstructsPath, "", "UTF-8")
		}

		IniRead, QSMinKB 	, % adm.Ini, % "QuickSearch", % "QSMinKB"
		IniRead, QSMaxKB	, % adm.Ini, % "QuickSearch", % "QSMaxKB"
		IniRead, QSFilter   	, % adm.Ini, % "QuickSearch", % "QSFilter"
		IniRead, BlockSize   	, % adm.Ini, % "QuickSearch", % "BlockSize"    	, % "4000"
		IniRead, WinSize    	, % adm.Ini, % "QuickSearch", % "WinSize"      	, % "xCenter yCenter w1200 h800"
		IniRead, lastdbname 	, % adm.Ini, % "QuickSearch", % "lastDB"
		IniRead, lastS        	, % adm.Ini, % "QuickSearch", % "lastSearch"
		IniRead, ExtrasStatus 	, % adm.Ini, % "QuickSearch", % "Extras"


		RegExMatch(WinSize, "i)w(?<W>\d+)\s*h(?<H>\d+)", GClient)
		GClientW -= 5

		lastdbname 	:= InStr(lastdbname, "ERROR") ? "" : lastdbname
		lastBlockSize	:= BlockSize
		lastBlockNr 	:= BlockNr
		lastS          	:= RegExReplace(lastS, "[\|]", "`n")

	;}

		Gui, QS: new    	, HWNDhQS -DPIScale +Resize
		Gui, QS: Color  	, % "c" gcolor , % "c" tcolor
		Gui, QS: Margin	, 5, 5

	;-: SPALTE 1: DATENBANKEN                       	;{

		; - Filter -
		Gui, QS: Font	, % "s" FSize-1 " q5 Normal" ColorTitles, Futura Bk Bt
		Gui, QS: Add	, Text        	, % "xm ym-3	 " BGT "               	vQSMinKBText"                                                              	, % "min KB"
		Gui, QS: Add	, Text        	, % "x+10 	 " BGT "               	vQSMaxKBText"                                                             	, % "max KB"
		Gui, QS: Add	, Text        	, % "x+10 	 " BGT "             	vQSFilterText"                                                                	, % "DB-Filter"

		; - Felder -
		Gui, QS: Add	, GroupBox  	, % "x+10 w50 h10                           	vQSFeldWahlT"                                                    	, % "Felder wählen"

		Gui, QS: Font	, % "s" FSize " q5 Normal" ColorData
		Gui, QS: Add	, Edit           	, % "xm y+4 	w" FSize*8 	" Number	 " BGT "    	" GFR " -E0x200 vQSMinKB "                  	, % QSMinKB
		Gui, QS: Add	, Edit           	, % "x+5    	w" FSize*8 	" Number	 " BGT "    	" GFR " -E0x200 vQSMaxKB "                 	, % QSMaxKB
		Gui, QS: Add	, Edit           	, % "x+5    	w" FSize*20	"           	 " BGT ASM GFR " -E0x200 vQSFilter  hwndhEdit  "     	, % QSFilter
		Edit_SetMargin(hEdit, 4, 1, 4, 4)

		Gui, QS: Add	, Button        	, % "x+10                                     	 " BGT "                                          	vQSFeldAlle" 	, % "Alle"
		Gui, QS: Add	, Button        	, % "x+10                                       	 " BGT "                                          	vQSFeldKeine"	, % "Keine"

		Gui, QS: Font	, % "s" FSize+2 " q5 Italic" ColorTitles
		Gui, QS: Add	, Text        	, % "x+15 ym+10 " BGT "    Center                                                        vQSFelderT  "    	, % "anzuzeigende Felder auswählen"

		; - Filter Elemente versetzen -
		GuiControlGet, cp, QS:Pos, QSMinKBText
		GuiControlGet, dp, QS:Pos, QSMinKB
		GuiControl, QS: Move, QSMinKBText	, % "x" dpX + Floor((dpW/2) - (cpW/2))
		GuiControl, QS: Move, QSMinKB   	, % " h" cpH + 2

		GuiControlGet, cp, QS:Pos, QSMaxKBText
		GuiControlGet, dp, QS:Pos, QSMaxKB
		GuiControl, QS: Move, QSMaxKBText, % "x" Floor(dpX + (dpW//2 - cpW//2))
		GuiControl, QS: Move, QSMaxKB   	, % " h" cpH + 2

		GuiControlGet, cp, QS:Pos, QSFilterText
		GuiControlGet, dp, QS:Pos, QSFilter
		GuiControl, QS: Move, QSFilterText	, % "x" Floor(dpX + (dpW/2 - cpW/2))
		GuiControl, QS: Move, QSFilter     	, % " h" cpH + 2

		GuiControl, QS: Move, QSFeldAlle   	, % " h" cpH + 2
		GuiControl, QS: Move, QSFeldKeine 	, % " h" cpH + 2


		; - Felder wählen versetzen -
		gbY := cpY-1,	gpmrg := 2
		GuiControlGet, cp, QS:Pos, QSFeldAlle
		GuiControlGet, dp, QS:Pos, QSFeldKeine
		GuiControl, QS: Move, QSFeldWahlT, % "x" cpX-2*gpmrg " y" gbY " w" (dpX+dpW-cpX+4*gpmrg) " h" cpW+5*gpmrg

		; - Datenbankauswahl
		LVOptions1 := "-E0x200 vQSDB HWNDQShDB gQS_Handler AltSubmit -Multi Grid"
		GuiControlGet, cp, QS:Pos, QSMinKB
		GuiControlGet, dp, QS:Pos, QSFilter
		mrgX := cpX, DBFilesW := dpX+dpW-mrgX, GCol2X := dpX+dpW+5

		Gui, QS: Font	, % "s" FSize " q5 Normal" ColorData
		Gui, QS: Add	, ListView    	, % "xm				 y+5 	w" DBFilesW 	" r16  " LVOptions1            	, % "Datenbank|records|Größe|letzte Änderung"

		; - Datenbanken anzeigen
		QuickSearch_DBFilter()

		lp := GetWindowSpot(QShDB)
		DBColW   	:= [95,65,75]
		DBColW.4 	:= lp.CW - DBColW.1 - DBColW.2 - DBColW.3
		Gui, QS: ListView, QSDB
		LV_ModifyCol()
		LV_ModifyCol(1, DBColW.1)
		LV_ModifyCol(2, DBColW.2 " Right Integer")
		LV_ModifyCol(3, DBColW.3 " Right Integer")
		LV_ModifyCol(4, DBColW.4 " Left")
		LV_ModifyCol(2, "SortDesc")
	;}

	;-: SPALTE 2: FELDINFOs                           	;{
		GuiControlGet, cp, QS:Pos, QSDB
		LVOptions2	:= " vQSDBFL gQuickSearch_DBFields HWNDQShDBFL -Readonly -E0x200 -LV0x18 AltSubmit -Multi Grid Checked"

		Gui, QS: Font	, % "s" FSize-1 " q5 Normal" ColorData
		Gui, QS: Add	, ListView    	, % "x" GCol2X " y" cpY " w" DBFieldW	" h" cpH " " LVOptions2   	, % "Nr|Feldname|Type|Länge|Nr.DB|Suchparameter"

		; Größenanpassung
		Gui, QS: Show, Hide
		Gui, QS: ListView, QSDBFL
		For colNr, colOption in ["40 Right Integer", "105 Text", "40 Center", "45 Center", "45 Center", "250 Text"]
			LV_ModifyCol(colNr, colOption)

		; letzte Spalte editierbar machen
		QSICEFI := New LV_InCellEdit(QShDBFL)
		QSICEFI.SetColumns(6)

		GuiControlGet, dp, QS:Pos, QSDBFL
		GCol3X 	:= dpX+dpW+5
	;}

	;-: SPALTE 3: DATENBANKSUCHE                 	;{
		GuiControlGet, cp, QS:Pos, QSFelderT
		GuiControlGet, dp, QS:Pos, QSFeldWahlT
		Gui, QS: Font	, % "s" FSize+2 " q5 Italic" ColorTitles, Futura Bk Bt
		Gui, QS: Add	, Text        	, % "x+15 y" cpY " w" DBSearchW " " BGT "    Center                                                           "	, % "RegEx Suche"

		GuiControlGet, cp, QS:Pos, QSDB
		GuiControlGet, dp, QS:Pos, QSDBFL
		Gui, QS: Font	, % "s" FSize " q5 Normal" ColorData
		Gui, QS: Add	, Edit         	, % "x" dpX+dpW+5 " y" dpY 	" w" DBSearchW " h" dpH " -E0x200 vQSDBS gQS_Handler" 	, % lastS
	;}

	;-: SPALTE 4: ZUSÄTZLICHE INFO'S            	;{
		GuiControlGet	, cp              	, QS: Pos	, QSFelderT
		Gui, QS: Font	, % "s" FSize+2 " q5 Italic" ColorTitles, Futura Bk Bt
		Gui, QS: Add	, Checkbox  	, % "x+5  y" cpY " w400 " BGT "                           vQSXTRT gQS_Handler"                        	, % "verknüpfte Daten"

		GuiControlGet	, cp           	, QS: Pos	, QSXTRT
		GuiControl    	, QS: Move	, QSXTRT	, % "y" dpY - cpY - 5   	; Position anpassen
		GuiControl    	, QS:        	, QSXTRT	, % ExtrasStatus           	; Checkbox setzen oder entfernen
		GuiControlGet	, cp           	, QS: Pos	, QSXTRT

		Gui, QS: Font	, % "s" FSize+1 " q5 Normal Bold Underline" ColorData
		Gui, QS: Add	, Edit         	, % "x" cpX " y" dpY " w400 h" cpH " -E0x200 Center vQSEXTRANFO"

		GuiControlGet	, ep           	, QS: Pos	, QSEXTRANFO
		Gui, QS: Font	, % "s" FSize-1 " q5 Normal" ColorData
		Gui, QS: Add	, Edit         	, % "y+1w400 h" dpH-epH-1 " -E0x200 vQSEXTRAS"

		GuiControl, % "QS:" (ExtrasStatus ? "Enable" : "Disable"), % "QSEXTRANFO"
		GuiControl, % "QS:" (ExtrasStatus ? "Enable" : "Disable"), % "QSEXTRAS"

		GuiControlGet, dp            	, QS: Pos	, QSEXTRAS
		QSEXTRASx := dpX
		GuiW := dpX+dpW-5

	;}

	;-: ZEILE 2: BEFEHLE                                    	;{
		Gui, QS: Font	, % "s" FSize+1 " q5" ColorTitles, Futura Md Bt
		Gui, QS: Add	, Text        	, % "xm 	y+2 	 " BGT "   "			                                                    	, % "Position:"
		Gui, QS: Add	, Text        	, % "x+5       	 " BGT "   vQST1"                                                   	, % empty "              "

		Gui, QS: Add	, Text        	, % "x+20       	 " BGT "  "				                                                	, % "Datensätze:"
		Gui, QS: Add	, Text        	, % "x+5       	 " BGT "   vQST2"                                                     	, % empty "   "

		Gui, QS: Add	, Text        	, % "x+20       	 " BGT "  "				                                                	, % "Treffer:"
		Gui, QS: Add	, Text        	, % "x+5       	 " BGT "   vQST3"                                                    	, % empty "         "

        Gui, QS: Add	, Text        	, % "x+40       	 " BGT "  "				                                                	, % "ausgewählter Patient:"
		Gui, QS: Add	, Text        	, % "x+5 w350 	 " BGT "   vQST4"                                                    	, % ""

		; - Progress -
		Gui, QS: Add	, Progress     	, % "xm  	y+2 w" GuiW " vQSPGS"			                                		, % 0
		GuiControlGet, cp, QS: Pos, QSPGS
		Gui, QS: Add	, Text        	, % "x" cpX " y" cpY " w" cpW " h" cpH "  " BGT "   Center vQSPGST"

		; - Buttons -
		Gui, QS: Add	, Button        	, % "      	y+5         	vQSSearch                                          	gQS_Handler "	, % "Suchen"

		Gui, QS: Add	, Button        	, % "x+10		          	vQSSaveRes                                        	gQS_Handler "	, % "Suche speichern als"

		Gui, QS: Font	, % "s" FSize+3 " q5" ColorData, Futura Md Bt
		GuiControlGet, cp, QS: Pos, QSSaveRes
		Gui, QS: Add	, DDL         	, % "x+2 h" cpH "      	vQSSaveAs r5                                      	gQS_Handler "	, % ".csv (Tab)|.csv (;)|.csv (Leerzeichen)"
		GuiControl, QS: Choose, QSSaveAs, 1

		Gui, QS: Font	, % "s" FSize+1 " q5" ColorData, Futura Md Bt
		Gui, QS: Add	, Button        	, % "x+20		          	vQSSaveSearch                                 	gQS_Handler "	, % "Sucheinstellungen speichern"

		Gui, QS: Add	, Button        	, % "x+20		          	vQSLBlock                                         	gQS_Handler "	, % "<< vorheriger Block"

		Gui, QS: Font	, % "s" FSize-2 " q5" ColorTitles, Futura Md Bt
		GuiControlGet, cp, QS: Pos, QSLBlock
		Gui, QS: Add	, Text         	, % "x" cpX+cpW+5 " y" cpY-4 "	vQSBlockNrBez"                                        	, % "Block Nr|Größe"

		Gui, QS: Font	, % "s" FSize-1 " q5" ColorData, Futura Md Bt
		Gui, QS: Add	, Edit         	, % "y+0 w" FSize*4 " vQSBlockNR Number 							gQS_Handler "	, % BlockNr
		Gui, QS: Add	, Edit         	, % "x+0 w" FSize*6 " vQSBlockSize Number 							gQS_Handler "	, % BlockSize

		Gui, QS: Font	, % "s" FSize+1 " q5" ColorData, Futura Md Bt
		Gui, QS: Add	, Button        	, % "x+5 y" cpY "	  	vQSNBlock                                         	gQS_Handler "	, % "nächster Block >>"

		Gui, QS: Add	, Button        	, % "x+150	   	      	vQSReload                                        	gQS_Handler "	, % "Reload"
		Gui, QS: Add	, Button        	, % "x+50	              	vQSQuit                                           	gQS_Handler "	, % "Beenden"

	;}

	;-: ZEILE 3: Ausgabelistview                          	;{
		Gui, QS: Show, %  WinSize " Hide", Quicksearch

		GuiControlGet, cp, QS: Pos, QSSearch
		LVOptions3 	:= "vQSDBR gQS_Handler HWNDQShDBR -E0x200 -LV0x10 AltSubmit -Multi Grid"
		wqs  	    	:= GetWindowSpot(hQS)
		LvH           	:= wqs.CH - cpY - cpH

		Gui, QS: Font	, % "s" FSize+1 " q5 Normal" ColorData, Futura Bk Bt
		Gui, QS: Add	, ListView    	, % "xm y" cpY+cpH+5 " w" GuiW " h" LvH " "  LVOptions3	, % "- -"
		Props.hLV   	:= QShDBR

	;}

	;-: Gui anpassen                                       	;{
		hMon	:= MonitorFromWindow(hQS)
		mon  	:= GetMonitorInfo(hMon)
		wqs  	:= GetWindowSpot(hQS)
		Lv     	:= GetWindowSpot(QShDBR)

		Gui, QS: Show,, Quicksearch

		Lv         	:= GetWindowSpot(QShDBR)
		QSmrgX 	:= wqs.W - Lv.W
		QSmrgY 	:= cpY + cpH + 10

		WinMove, % "ahk_id " hQS,,,, % wqs.W+1, % wqs.H+1
		WinMove, % "ahk_id " hQS,,,, % wqs.W, % wqs.H
	;}

	; lädt die zuletzt benutzte Datenbank
	If lastdbname {
		BlockInput, On
		QuickSearch_ListFields(dbname := lastdbname)
		BlockInput, Off
	}

	; Hotkeys
	fn_QSHK := Func("QuickSearch_Hotkeys")
	Hotkey, IfWinActive, Quicksearch ahk_class AutoHotkeyGUI
	Hotkey, Left   	, 	% fn_QSHK
	Hotkey, Right 	, 	% fn_QSHK
	Hotkey, Enter 	, 	% fn_QSHK
	Hotkey, TAB  	, 	% fn_QSHK
	Hotkey, IfWinActive


return

QS_Handler: 	;{

	Gui, QS: Submit, NoHide

	Switch A_GuiControl	{

		Case "QSReload":                      	;{     Skript neu laden
			ReloadGui :=true
			gosub QSGuiReload
		;}

		Case "QSDB":                             	;{     Datenbank öffnen
			If (A_EventInfo = EventInfo) || (StrLen(A_EventInfo) = 0)
				return
			EventInfo := A_EventInfo
			If Instr(A_GuiEvent	, "DoubleClick") {

					Last_tblRow := 0

				; Datenbankname lesen
					Gui, QS: ListView, QSDB
					LV_GetText(dbname, EventInfo, 1)
					QuickSearch_ListFields(dbname)
			}
		;}

		Case "QSDBR":                             	;{     Datentabelle
			       If (A_GuiEvent = "ColClick"                  	)	{
				colHDR := A_EventInfo
				MouseGetPos, mx, my
				ToolTip, % colHDR, % mx, % my-10, 6
				SetTimer, QSColTipOff, -3000
			}
			else if (A_GuiEvent = "DoubleClick"            	) 	{

				If !WinExist("ahk_class OptoAppClass")
					return

				If (Pcol := Props.PatIDColumn()) {
					Gui, QS: ListView, QSDBR
					LV_GetText(PatID, A_EventInfo, Pcol)
					PatID    	:= StrReplace(PatID, "*")
					PatName 	:= PatDB[PatID].NAME ", " PatDB[PatID].VORNAME
					MsgBox	, 0x2024
								, Quicksearch - Quickclick
								, % "Soll die Karteikarte des Patienten:`n" PatName
								. 	 ", *" ConvertDBASEDate(PatDB[PatID].GEBURT)
								.	 "`nin Albis angezeigt werden?", 10
					IfMsgBox, No
						return
					IfMsgBox, TimeOut
						return
					GuiControl, QS:, QST4, % PatDB[PatID].NAME ", " PatDB[PatID].VORNAME
					AlbisAkteOeffnen(PatName, PatID)
				}

			}
			else if RegExMatch(A_GuiEvent, "(Normal|I)"	)  	{

				tblRow := A_EventInfo
				If (Last_tblRow = tblRow)
					return
				Last_tblRow := tblRow

			  ; PatID finden
				If (Pcol := Props.PatIDColumn()) {
					Gui, QS: ListView, QSDBR
					LV_GetText(PatID, tblRow, Pcol)
					PatID := StrReplace(PatID, "*")
					GuiControl, QS:, QST4, % "[" PatID "] " PatDB[PatID].NAME ", " PatDB[PatID].VORNAME " *" PatDB[PatID].GEBURT
				}

			 ; verlinkte Daten aus anderer Datenbank anzeigen (wenn Häkchen bei verknüpfte Daten gesetzt ist)
				If QSXTRT && Props.LabelExist("TEXTDB")
					QuickSearch_Extras(tblRow)

			}
		;}

		Case "QSLBlock":                        	;{
			QuickSearch_ListFields(dbname, 2)
		;}

		Case "QSNBlock":                     	;{
			QuickSearch_ListFields(dbname, 1)
		;}

		Case "QSSearch": 	                    	;{     Datenbanksuche
			If !dbname
				return
			Gui, QS: Submit, NoHide
			Quicksearch(dbname, QSDBS)
		;}

		Case "QSSaveRes":                    	;{ 	Suchergebnisse speichern
			If !dbname
				return
			Gui, QS: Submit, NoHide
			QuickSearch_SaveResults(QSSaveAs)
		;}

		Case "QSXTRT":                         	;{ 	verknüpfte Daten aktivieren/deaktivieren
			GuiControl, % "QS:" (ExtrasStatus ? "Enable" : "Disable")	, % "QSEXTRANFO"
			GuiControl, % "QS: Enable" (QSXTRT ? "1" : "0")          	, % "QSEXTRAS"
		;}

	}

return ;}

QSColTipOff:  	;{
	ToolTip,,,, 6
return ;}

QSGuiClose:
QSGuiEscape: 	;{
QSGuiReload:

	QuickSearch_GuiSave()
	If ReloadGui
		Reload

ExitApp ;}

QSGuiSize:    	;{

	Critical, Off	; erst Critical Off soll Critical On dann schneller machen oder zuverlässiger machen (hab vergessen woher ich das habe)
	if (A_EventInfo = 1), return
	Critical
	QSW := A_GuiWidth, QSH:= A_GuiHeight
	QSEXTRASw := QSW-QSEXTRASx-5

	GuiControl, QS: MoveDraw	, QSDBR       	, % "w" QSW-10 " h" QSH-QSmrgY
	GuiControl, QS: MoveDraw	, QSPGR       	, % "w" QSW-10
	GuiControl, QS: MoveDraw	, QSEXTRANFO	, % "w" (QSEXTRASw < 300 ? 300 : QSEXTRASw)
	GuiControl, QS: MoveDraw	, QSEXTRAS   	, % "w" (QSEXTRASw < 300 ? 300 : QSEXTRASw)
	GuiControl, QS: +Redraw 	, QSDBR
	WinSet, Redraw,, % "ahk_id " QShDBR

return ;}

}

QuickSearch_GuiSave()                                                                                     	{	; Fenstereinstellung speichern

	global hQS, QSDB, QSDBFL, QSDBS, QSDBR, QSPGS, QST1, QST2, QST3, QSXTRT
	global QSMinKB, QSMaxKB, QSFilter, BlockSize, winSize, dbname

	Gui, QS: Default
	Gui, QS: ListView, Default
	Gui, QS: Submit, NoHide

	wqs  	:= GetWindowSpot(hQS)
	lastS 	:= RegExReplace(QSDBS, "[\r\n]+", "|")

	IniWrite, % QSMinKB 	, % adm.Ini, % "QuickSearch", % "QSMinKB"
	IniWrite, % QSMaxKB	, % adm.Ini, % "QuickSearch", % "QSMaxKB"
	IniWrite, % QSFilter   	, % adm.Ini, % "QuickSearch", % "QSFilter"
	IniWrite, % BlockSize    	, % adm.Ini, % "QuickSearch", % "QSBlockSize"
	IniWrite, % winSize    	, % adm.Ini, % "QuickSearch", % "x" wqs.X " y" wqs.Y " w" wqs.CW " h" wqs.CH
	IniWrite, % dbname    	, % adm.Ini, % "QuickSearch", % "lastDB"
	IniWrite, % lastS        	, % adm.Ini, % "QuickSearch", % "lastSearch"
	IniWrite, % QSXTRT    	, % adm.Ini, % "QuickSearch", % "Extras"

	If !Props.SaveProperties()
		MsgBox, Einstellungen konnten nicht gesichert werden!

return
}

QuickSearch_Hotkeys()                                                                                      	{

	global dbname, QS, QSDBS, QSBlockNr, QSBlockSize

	If !WinActive("Quicksearch ahk_class AutoHotkeyGUI") {
		SendInput, % "{" A_ThisHotkey "}"
		return
	}

	thisHotkey := A_ThisHotkey
	Gui, QS: Default
	Gui, QS: Submit, NoHide
	GuiControlGet, gfocus, QS: FocusV

	Switch thisHotkey {

		Case "TAB":
			If RegExMatch(gfocus, "i)(QSBlockNr|QSBlockSize)")
				QuickSearch_ListFields(dbname, 3)
			else if RegExMatch(gfocus, "i)QSDBS")
				SendInput, {Down}
			else
				SendInput, {Enter}

		Case "Enter":
			If RegExMatch(gfocus, "i)(QSBlockNr|QSBlockSize)")
				QuickSearch_ListFields(dbname, 3)
			else if RegExMatch(gfocus, "i)QSDBS")
				SendInput, {Down}
			else
				SendInput, {Enter}

		Case "Left":
			If RegExMatch(gfocus, "i)(QSDBS|QSBlockNR|QSBlockSize)") {
				SendInput, {Left}
				return
			}
			QuickSearch_ListFields(dbname, 2)

		Case "Right":
			If RegExMatch(gfocus, "i)(QSDBS|QSBlockNR|QSBlockSize)") {
				SendInput, {Right}
				return
			}
			QuickSearch_ListFields(dbname, 1)

	}

}

QuickSearch_Progress(index, maxIndex, len, matchcount)                                    	{	; Suchfortschritt anzeigen

	global	DBSearchW
	prgPos := Floor((index*100)/maxIndex)
	GuiControl, QS:, QSPGS	, % prgPos
	GuiControl, QS:, QST1 	, % index
	GuiControl, QS:, QST3 	, % matchcount

}

QuickSearch_Extras(tblRow)                                                                               	{	; verknüpfte Daten aus anderer Datenbank laden

		If (Props.dbname <> "BEFUND" || !Props.isVisible("TEXTDB")) {
			SciTEOutput("TEXTDB is visible: " Props.isVisible("TEXTDB"))
			return
		}

		Gui, QS: ListView, QSDBR
		LV_GetText(TEXTDB, tblRow, Props.HeaderVirtualPos("TEXTDB"))
		If (TEXTDB = 0)
			return

		dbf       	:= new DBASE(adm.AlbisDBPath "\BEFTEXTE.dbf", 1)
		records		:= dbfRecords := dbf.records
		steps      	:= records/100
		filepos   	:= dbf.OpenDBF()

		found := false
		Loop % records {
			set := dbf.ReadRecord(["PATNR", "DATUM", "LFDNR", "POS", "TEXT"])
			If (TEXTDB = set.LFDNR) {
				found := true
				extra .= set.TEXT
				extraLines := set.POS + 1
			}
			else If (found && TEXTDB <> set.LFDNR) {
				break
			}
			If (Mod(A_Index, 100) = 0)
				QuickSearch_Progress(A_Index, records, StrLen(records), matchcount)
		}

		filepos   	:= dbf.CloseDBF()
		extrahead	:=  extraLines " Zeile" (extraLines = 1 ? "" : "n")
							. " aus BEFTEXT.dbf für TEXTDB Eintrag: " TEXTDB

		GuiControl, QS:, QSEXTRANFO	, % extrahead
		GuiControl, QS:, QSEXTRAS   	, % extra

}

QuickSearch_ListFields(dbname, step:=0)                                                             	{ 	; erstellt die Ausgabetabelle

		global	 QS, QSDBS, QSBlockNr, QSBlockSize, QSPGS, QSPGST
		global 	dbfiles, empty, dbfdata
		global 	BlockNr, BlockSize, lastBlockNr, lastBlockSize, dbfRecords
		static 	lastdbname

		Gui, QS: Submit, NoHide

	; neue Datenbank soll abgerufen werden
		If (dbname <> lastdbname) {
			dbf               	:= new DBASE(adm.AlbisDBPath "\" dbname ".dbf", 0)
			dbfRecords    	:= dbf.records
			dbfConnection	:= IsObject(dbf.connected) ? dbf.connected : ""           	; weitere Daten in anderer Datenbank vorhanden?
			dbf                	:= ""
			lastdbname    	:= dbname
			step               	:= 0
			Props.LVColumnSizes()                                                                       	; Spalten wie bei letzter Anzeige wiederherstellen
		}

		BlockNr 	:= QSBlockNr
		BlockSize 	:= QSBlockSize


		If (StrLen(BlockSize) 	= 0	|| !RegExMatch(BlockSize	, "^\d+$"))
			BlockSize := lastBlockSize
		If !RegExMatch(BlockNr	, "^\d+$")
			BlockNr := lastBlockNr

		If      	(step = 0) {     	; erste Anzeige

			; Einstellungsobjekt anlegen                              	;{
				If !Props.WorkDB(dbname, dbfiles[dbname].dbfields) {
					throw A_ThisFunc ": Einstellungen zur Datenbank: " dbname " konnten nicht angelegt werden!"
					return
				}
			;}

			; Ausgabelistview erstellen                              	;{

				LV_DeleteCols("QS", "QSDBR")                           	; Spalten entfernen

				thisColPos := 1
				For colNr, label in Props.Labels()
					If Props.isVisible(label) {

						Gui, QS: ListView, QSDBR

						dtype 	:= Props.DataType(flabel)
						width 	:= Props.HeaderField(label, "width")
						width	:= !width ? "AutoHDR" : width

						If (thisColPos = 1)
							LV_ModifyCol(thisColPos, width " " dtype, label)
						else
							LV_InsertCol(thisColPos	, width " " dtype, label)

						LV_ModifyCol(thisColPos, width)
						Props.HeaderSetVirtualPos(label, thisColPos)
						thisColPos ++

					}
			;}

			; Datenbankfelderinfo's anzeigen                     	;{
				Gui, QS: ListView, QSDBFL
				LV_Delete()

				For colNr, label in Props.Labels() {
					field := Props.HeaderField(label)
					LV_Add(Props.isVisible(Label)?"Check":"", field.col, field.name, field.Type, field.len, field.pos)
					t .= label "=`n"
				}

			;}

			; Statistik                                                          	;{
				GuiControl, QS:, QSDBS, % t
				GuiControl, QS:, QST1	, % empty
				GuiControl, QS:, QST2	, % dbfiles[dbname].records
				GuiControl, QS:, QST3	, % empty
			;}

		}
		else if	(step = 1)     	; nächster Anzeigeblock
			BlockNr ++
		else if 	(step = 2) 		; vorheriger Anzeigeblock
			BlockNr := BlockNr - 1 < 0 ? 0 : BlockNr - 1
		else if 	(step = 3)    		; manuell eingegebener Anzeigeblock
			BlockNr := (BlockNR*BlockSize >= dbfiles[dbname].records) ? Floor(dbfRecords/BlockSize) : BlockNr

		GuiControl, QS:, QSBlockNR 	, % BlockNr
		GuiControl, QS:, QSBlockSize	, % BlockSize

		lastBlockSize	:= BlockSize
		lastBlockNr	:= BlockNr

	; Daten aus Albis Datenbank lesen                  	;{
		GuiControl, QS:, QSPGST, % "lade Block " BlockNr " ..."
		dbf       	:= new DBASE(adm.AlbisDBPath "\" dbname ".dbf", 2)
		records		:= dbfRecords := dbf.records
		filepos   	:= dbf.OpenDBF()
		dbfdata    	:= dbf.ReadBlock(BlockSize,, BlockNr, "QuickSearch_Progress")
		filepos   	:= dbf.CloseDBF()
		GuiControl, QS:, QSPGST, % ""
		QuickSearch_Progress(BlockNr*BlockSize+1 "-" (BlockNr+1)*BlockSize, records, 0, dbfdata.Count())
	;}

	; Vorschau der Daten im Ergebnisfenster          	;{
		Gui, QS: Default
		Gui, QS: ListView, QSDBR
		LV_Delete()
		GuiControl, QS: -Redraw, QSDBR

		For rowNr, set in dbfdata {

			; Spalten zusammenstellen
				cols := Object()
				For flabel, value in set {

					If Props.isVisible(fLabel) {
						If (flabel = "removed") {
							value := value = 1 ? "*" : ""
						} else if (flabel = "ID") {
							If (value = -1) {
								cols[Props.HeaderVirtualPos("removed")] := "*"
								continue
							}
						}
						cols[Props.HeaderVirtualPos(flabel)] := value
					}

				}

			; Zeile hinzufügen
				LV_Add("", cols*)

		}

		GuiControl, QS: +Redraw, QSDBR

	;}


}

QuickSearch_DBFields(GChwnd:="", GuiEvent:="", EventInfo:="", ErrLevel:="")    	{	; Eventhandler DBFields

		global dbfiles, dbname, dbfdata, QShDBFL, QShDBR
		static DataType 	:= {"N":"Integer", "D":"Integer", "C":"Text", "L":"Text", "M":"Text"}
		static evTime := 0

		/*
			A: Eine Zeile wurde aktiviert, was standardmäßig geschieht, wenn sie doppelt angeklickt wurde. Die Variable A_EventInfo enthält die Nummer der Zeile.
			C: Die ListView hat die Mauserfassung freigegeben.
			E: Der Benutzer hat begonnen, das erste Feld einer Zeile zu editieren (der Benutzer kann das Feld nur editieren, wenn -ReadOnly in den Optionen der ListView vorhanden ist).
				Die Variable A_EventInfo enthält die Nummer der Zeile.
			F:	Die ListView hat den Tastaturfokus erhalten.
			f 	(kleines F): Die ListView hat den Tastaturfokus verloren.
			I: 	Der Zustand einer Zeile hat sich in irgendeiner Form geändert; zum Beispiel weil sie ausgewählt, abgewählt, abgehakt usw. wurde.
				Wenn der Benutzer eine neue Zeile auswählt,	empfängt die ListView mindestens zwei solcher Benachrichtigungen: eine für das Abwählen der vorherigen Zeile und eine für das Auswählen der neuen Zeile.
				[v1.0.44+]:    	Die Variable A_EventInfo enthält die Nummer der Zeile.
				[v1.0.46.10+]: ErrorLevel enthält null oder mehr der folgenden Buchstaben, um mitzuteilen,
				wie das Element geändert wurde: S (ausgewählt) oder s (abgewählt), und/oder F (fokussiert) oder f (defokussiert), und/oder C (Häkchen gesetzt) oder c (Häkchen entfernt).
				SF beispielsweise bedeutet, dass die Zeile ausgewählt und fokussiert wurde. Um festzustellen, ob ein bestimmter Buchstabe vorhanden ist, verwenden Sie eine Parsende Schleife
				oder die GroßKleinSensitiv-Option von InStr(); zum Beispiel: InStr(ErrorLevel, "S", true). Hinweis: Aus Gründen der Kompatibilität mit zukünftigen Versionen sollte ein Skript nicht
				davon ausgehen, das "SsFfCc" die einzigen möglichen Buchstaben sind. Außerdem können Sie Critical in der ersten Zeile von g-Label einfügen, um sicherzustellen,
				dass alle "I"-Benachrichtigungen empfangen werden (andernfalls könnten einige davon verloren gehen, wenn das Skript nicht mit ihnen Schritt halten kann).

			K: Der Benutzer hat eine Taste gedrückt, während die ListView den Fokus hat. A_EventInfo enthält den virtuellen Tastencode der Taste (eine Zahl zwischen 1 und 255). Dieser Code kann
				via GetKeyName() in einen Tastennamen oder in ein Zeichen übersetzt werden. Zum Beispiel Taste := GetKeyName(Format("vk{:x}", A_EventInfo)). Bei den meisten Tastaturlayouts
				können die Tasten A bis Z via Chr(A_EventInfo) in das entsprechende Zeichen übersetzt werden. F2 wird unabhängig von WantF2 erfasst. Enter hingegen wird nicht erfasst; um es
				dennoch zu erfassen, können Sie, wie unten beschrieben, eine Standardschaltfläche nutzen.

			M: Auswahlrechteck. Der Benutzer hat damit begonnen, ein Auswahlrechteck über mehrere Zeilen oder Symbole zu ziehen.

			S: Der Benutzer hat begonnen, in der ListView zu scrollen.

			s (kleines S): Der Benutzer hat aufgehört, in der ListView zu scrollen.
		*/

		Critical
		;~ If (GuiEvent = "K")
			;~ SciTEOutput(GuiEvent "," EventInfo ", " ErrLevel)
		;~ else
		;~ If GuiEvent = "I"
			;~ SciTEOutput((A_TickCount - evTime < 200 ? "" : "`n") "GEvent: " GuiEvent "`t| EInfo: " EventInfo "`t|EL:" ErrLevel)
			;~ evTime := A_TickCount

	; Checkboxstatus hat sich geändert
		If RegExMatch(GuiEvent . EventInfo, "(K32|C0)") {

		  ; Feldinfo Listview auf Default setzen und gLabel Aufrufe ausstellen
			Gui, QS: ListView	, QSDBFL
			GuiControl, QS: -g, QSDBFL

		  ; zeilenweises Auslesen des Checkboxstatus
			Loop % LV_GetCount() {

				dbfieldsRow	:= A_Index
				ischecked  	:= LV_RowIsChecked(QShDBFL, dbfieldsRow)   ; Checkboxstatus
				flabel         	:= Props.HeaderLabel(dbfieldsRow)
				;SciTEOutput(flabel " is " Props.isVisible(fLabel) " <> " ischecked )

			  ; ----------------------------------------------------------------------------
			  ; eine Spalteneinstellung hat sich geändert
			  ; ----------------------------------------------------------------------------
				If (Props.isVisible(fLabel) <> ischecked) {

					; Spalte in virtueller Tabelle ein- oder ausblenden, Klasse nummeriert im Feld-Info Listview
						Gui, QS: ListView, QSDBR     ; Ausgabelistview
						colpos := Props.VirtualView(fLabel, ischecked)

					; Spalte wird ausgeblendet
						If !ischecked {

						  ; Speichern der Pixelbreite vor dem Löschen der Spalte
							Props.HeaderSetVirtualWidth(flabel, LVM_GetColWidth(QShDBR, colpos))

						  ; Spalte löschen
							Gui, QS: ListView, QSDBR     ; Ausgabelistview
							LV_DeleteCol(colpos)

						}
					; Spalte wird eingeblendet
						else {

							dtype 	:= Props.DataType(flabel)
							width 	:= Props.HeaderField(flabel, "width")
							width	:= StrLen(width) = 0 || width = 0 ? "AutoHDR" : width

							Gui, QS: ListView, QSDBR	; Ausgabelistview
							LV_InsertCol(colpos, width " " fType, flabel)
							For row, field in dbfdata
								LV_Modify(row, "Col" colpos, field[flabel])
							If (width = "AutoHDR")
								LV_ModifyCol(colpos, "AutoHDR")
						}

				}
			}
		}

	; gLabel Aufruf wieder ermöglichen
		GuiControl, QS: +gQuickSearch_DBFields, QSDBFL

}

QuickSearch_DBFilter(GChwnd:="", GuiEvent:="", EventInfo:="", ErrLevel:="")     	{	; Eventhandler DBFilter

	global	QSMinKB      	, QSMaxKB    	, QSFilter
	global	dbfiles, QShDB, DBColW, DBFilesW

	Gui, QS: Submit, NoHide

	If (QSMinKB = 0)
		GuiControl, QS:, QSMinKB	, % (QSMinKB := "")
	If (QSMaxKB = 0)
		GuiControl, QS:, QSMaxKB	, % (QSMaxKB := "")

	m1 := QSMinKB 	? true : false
	m2 := QSMaxKB 	? true : false
	m3 := QSFilter  	? true : false

	Gui, QS: ListView, QSDB
	LV_Delete()
	For dbfName, db in dbfiles {
		hits := 0
		hits += !m3 ? 1 : (InStr(dbfName, QSFilter)	? 1 : 0)
		hits += !m2 ? 1 : (db.sizeDBF <= QSMaxKB	? 1 : 0)
		hits += !m1 ? 1 : (db.sizeDBF >= QSMinKB 	? 1 : 0)
		If (hits = 3)
			LV_Add("", dbfName, db.records, db.sizeDBF " KB", db.lastupdateE, QSFormat(db.modified))
	}

	lp := GetWindowSpot(QShDB)
	DBColW.4 := lp.CW - DBColW.1 - DBColW.2 - DBColW.3
	LV_ModifyCol(1, DBColW.1)
	LV_ModifyCol(2, DBColW.2 " Right Integer")
	LV_ModifyCol(3, DBColW.3 " Right Integer")
	LV_ModifyCol(4, DBColW.4 " Left")

}

QuickSearch_SaveResults(format:="csv (Tab)")                                                    	{	; Ausgabetabelle speichern

		global QS, QSDBR, QShDBR

		FileSelectFile, filesavepath,, % A_ScriptDir, Dateiverzeichnis und Dateinamen wählen, *.csv

		filesavepath .=  ".csv"
		DelimiterChar := A_Tab

		table         	:= ""
		tableheader 	:= LV_GetHeader(QShDBR, "QS", "QSDBR")

		For col, label in tableheader.nr
			table .= (col > 1 ? DelimiterChar : "") label
		table .= "`n"
		maxCols := col

		Gui, QS: Default
		Gui, QS: ListView, QSDBR

		Loop, % LV_GetCount() {
			row := A_Index
			Loop % maxCols {
				col := A_Index
				LV_GetText(cell, row, col)
				table .= cell (col < maxCols ? DelimiterChar : "`n")
			}
		}

		FileOpen(filesavepath, "w", "UTF-8").Write(table)
		FileGetSize, fsize, % filesavepath
		If FileExist(filesavepath)
			MsgBox, % "Die Datei wurde gespeichert.`nDateigröße: " FSize " kb"
		else
			MsgBox, Die Datei wurde nicht gespeichert!

}

class Quicksearch_Properties                                                                              	{	; Tabellenmanager

		__New(propertiesPath) {

			SplitPath, propertiesPath,, outFileDir,, outFileNameNoExt
			this.FileDir     	:= outFileDir
			this.FileName 	:= outFileNameNoExt ".json"
			this.FilePath     	:= this.FileDir "\" this.FileName

			If FilePathCreate(this.FileDir) && FileExist(this.FilePath)
				this.qs := this.LoadProperties()
			else
				this.qs := Object()

		}

		WorkDB(dbname, dbfields)                  	{	; Arbeitsdatenbank festlegen und Tabellen-Informationen hinzufügen

			this.dbname := dbname

		; Klasse wurde initialisiert?
			If !IsObject(this.qs)                            	{
				MsgBox, 1,  A_ThisFunc, "Die Klasse muss zuerst initialisiert werden"
				return
			}

		; Objekt für virtuelle Datenbanktabelle anlegen
			If !IsObject(this.qs[this.dbname])        	{

				this.qs[this.dbname] := Object()

				If !IsObject(this.qs[this.dbname].header)	{                                                  	; Einstellungen der Spalten
					this.qs[this.dbname].header       	:= Object()                                        	; Spaltenüberschriften
					this.qs[this.dbname].header.enum := Object()                                        	; in Reihenfolge wie in der DBase Datei
					this.qs[this.dbname].header.text 	:= Object()                                       	; nach Namen sortiert
				}

			  ; aus einem Objekt mach zwei
				this.qs[this.dbname].PatID := ""
				For label, field in dbfields {

					; der Name des Patienten wird bei vorhandenen Labeln wie "PATNR" oder "NR" angezeigt
						If RegExMatch(label, "i)(PATNR|NR)")
							this.qs[this.dbname].PatID := label

					If !this.qs[this.dbname].header.enum.haskey(field.pos)
						this.qs[this.dbname].header.enum[field.pos] := label

				  ; Label anlegen
					If !IsObject(this.qs[this.dbname].header.text[label]) {
						this.qs[this.dbname].header.text[label]             	:= Object()
						this.qs[this.dbname].header.text[label].visible  	:= true                     	; default = sichtbar
						this.qs[this.dbname].header.text[label].name     	:= label                  	; Bezeichnung der Spalte (Feldes)
						this.qs[this.dbname].header.text[label].width     	:= "AutoHDR"
						this.qs[this.dbname].header.text[label].pos       	:= field.pos            	; Spaltennummer in der dBase Datei
						this.qs[this.dbname].header.text[label].col      	:= field.pos            	; Spaltennummer in der Gui Listview
						this.qs[this.dbname].header.text[label].type      	:= field.type            	; Datentype (z.B. Text, Integer, Datum ..)
						this.qs[this.dbname].header.text[label].len        	:= field.len            	; Länge des Datenfeldes (Zeichenanzahl)
					}

				}

				  ; removed field anhängen
					lastfield := this.qs[this.dbname].header.enum.Count() + 1
					this.qs[this.dbname].header.enum[lastfield] := label := "removed"
					this.qs[this.dbname].header.text[label]             	:= Object()
					this.qs[this.dbname].header.text[label].visible  	:= true
					this.qs[this.dbname].header.text[label].name     	:= label
					this.qs[this.dbname].header.text[label].width     	:= "AutoHDR"
					this.qs[this.dbname].header.text[label].pos       	:= lastfield
					this.qs[this.dbname].header.text[label].col      	:= lastfield                	; Spaltennummer in der Gui Listview
					this.qs[this.dbname].header.text[label].type      	:= "C"                          	; "C" = Char
					this.qs[this.dbname].header.text[label].len        	:= 1                           	; ein Zeichen

				; Tabellenobjekt speichern
					this.SaveProperties(false)

			}

		return 1
		}

		DataType(label)                                  	{	; automatische Spalteneinstellungen für die Windows Listview anhand der dBase Feldtypen
			static DataType 	:= {"N":"Integer", "D":"Integer", "C":"Text", "L":"Text", "M":"Text"}
		return DataType[this.qs[this.dbname].header.text[label].type]
		}

		SortType(label)                                  	{

			static SortType := {"PATNR":"Logical"}

		}


	; Teil-Objekte zurückgeben
		HeaderField(label, param:="")              	{	; alle Daten eines Feldes (Spaltenüberschriften)
			If !param
				return this.qs[this.dbname].header.text[label]
			else
				return this.qs[this.dbname].header.text[label][param]
		}

		Header()                                           	{	; Objekt mit Spaltenüberschriften und Daten
		return this.qs[this.dbname].header.text
		}

		Labels()                                             	{	; Array mit Spaltenüberschriften
		return this.qs[this.dbname].header.enum
		}

		Labels2() {
			arr:=[]
			For idx, field in this.header()
				If field.haskey("Label")
					arr[field.pos] := field.Label
		return arr
		}


	; Tabelleninformationen erhalten/ändern, virtuelle Tabelle ändern
		HeaderCount()                                     	{	; Anzahl aller Spalten
		return this.qs[this.dbname].header.enum.Count()
		}

		HeaderLabel(colNr)                            	{	; Name des Feldes an der Position in der dBase Datei
		return this.qs[this.dbname].header.enum[colNr]
		}

		HeaderPos(label)                               	{	; Position des Feldes in der dBase Datei zurückgeben
		return this.qs[this.dbname].header.text[Label].pos
		}

		HeaderVirtualPos(label)                      	{	; Position in der virtuellen Tabelle (Gui-Tabelle)
		return this.qs[this.dbname].header.text[label].col
		}

		HeaderSetVirtualWidth(label, width)    	{	; Pixelbreite eine Spalte des realen Steuerelementes speichern
		this.qs[this.dbname].header.text[label].width := width
		}

		HeaderSetVirtualPos(label, colpos)      	{	; Position eines Feldes in der virtuellen Tabelle setzen
			this.qs[this.dbname].header.text[label].col := colpos
			Gui, QS: ListView, QSDBFL                                                                	; ändert die Positionsanzeige im Field-Info Listview
			LV_Modify(this.HeaderPos(label),, colpos)
		}

		HeaderSetVisible(label, VHDRPos)        	{	; virtuelle Spalte ein- oder ausblenden und alle virtuellen Positionen rechts daneben sichtbarer Spalten verschieben

			this.qs[this.dbname].header.text[label].visible := !VHDRPos ? false : true
			this.HeaderSetVirtualPos(label, VHDRPos)

			hdrPos := this.HeaderPos(label)                                                           	;
			Loop, % this.HeaderCount() - hdrPos {	    		                                	; alle rechts davon sichtbaren virtuellen Tabellenfelder um eine Position nach rechts oder links verschieben
				nlabel := this.HeaderLabel(hdrPos+A_Index-1)	                              	; Bezeichnung finden
				If (nlabel <> label) && this.isVisible(nlabel) {                                  	; ist es eine sichtbare Spalte
					nPos	:= this.HeaderVirtualPos(nlabel) + (VHDRPos ? 1 : -1)			; wird die Position um 1 (nach rechts oder links) versetzt
					this.HeaderSetVirtualPos(nlabel, nPos)	                                   		; und zwar jetzt
				}
			}

		}

		PatIDColumn()                                 		{	; gibt die Spalte der Patientennummer in der virtuellen Listview zurück

			static PatIDCol := false

			PatIDLabel := this.qs[this.dbname].PatID
			If !PatIDLabel && !PatIDCol {
				For colNr, label in this.Labels()
					If RegExMatch(label, "i)(PATNR|NR)") {
						PatIDLabel 	:= this.qs[this.dbname].PatID := label
						PatIDCol   	:= true
						break
					}
			}

		return PatIDLabel ? this.HeaderVirtualPos(PatIDLabel) : 0
		}

		LabelExist(label)                               		{	; true wenn Spaltenüberschrift vorhanden ist

			If IsObject(this.qs[this.dbname].header.text[label])
				return true

		return false
		}


	; virtuelle Tabelle
		VirtualView(label, view)                       	{	; virtuelle Spalte anzeigen oder ausblenden

			; Ausblenden
				If !view {
					VHDRPos := this.HeaderVirtualPos(label)
					this.HeaderSetVisible(label, 0)
					return VHDRPos
				}

			; Einblenden
				colNr := this.HeaderPos(label), VHDRPos := 0
				For hdrPos, flabel in this.Labels() {
					If (this.HeaderVirtualPos(flabel) > 0)
						lastVirtualPos := this.HeaderVirtualPos(flabel) + 1	; letzte sichtbare Spalte
					If (this.isVisible(flabel) && (hdrPos > colNr)) {            	; Spalte ist sichtbar und steht hinter der einzufügenden Spalte in der dBase Datei
						VHDRPos := this.HeaderVirtualPos(flabel)
						break
					}
				}

				If (VHDRPos = 0)
					VHDRPos := lastVirtualPos
				this.HeaderSetVisible(label, VHDRPos)

		return VHDRPos                                                                  	; Einfüge- oder Ausblendeposition zurückgeben
		}

		isVisible(label)                                    	{	; Spalte ist sichtbar
		return this.qs[this.dbname].header.text[label].visible
		}

		LVColumnSizes()                                 	{	; Pixelbreite aller Spalten
			If !this.hLV
				return
			LVHeader := LV_ColumnGetSize(this.hLV, 0)
			For colLabel, colWidth in LVHeader.width
				this.HeaderSetVirtualWidth(colLabel, colWidth)
		}


	; Tabelleneinstellungen laden oder speichern
		LoadProperties()                                  	{	; Objekt laden
			If !this.FilePath || !FileExist(this.FilePath) {
				throw A_ThisFunc ": Einstellungsdatei <" this.FileName "> ist nicht im `n"
						. "Ordner: " this.FileDir " vorhanden."
				return
			}
		return JSONData.Load(this.FilePath, "", "UTF-8")
		}

		SaveProperties(saveColWidth:=true)    	{	; Objekt speichern

			If !this.FilePath
				If !FilePathCreate(this.FileDir) {
					throw A_ThisFunc ": Speicherpfad für das Sichern der Einstellungen konnte nicht angelegt werden.`n"
							. "Ordner: " this.FileDir
					return
				}

			; Spaltenbreite speichern
				If saveColWidth && this.hLV
					this.LVColumnSizes()

			JSONData.Save(this.FilePath, this.qs, true,, 1, "UTF-8")

		return FileExist(this.FilePath) ? 1 : 0
		}


	; +++++++++ PRIVATE FUNKTIONEN ++++++++++++
		CheckLabel(label, FuncName) {
			return
			If (StrLen(label) = 0)
				return 0
			If !IsObject(this.qs[this.dbname].header.text[label]) {
				throw FuncName ": Die Spaltenüberschrift <" label "> gibt es nicht im Tabellenobjekt [" this.dbname "]"
				return 0
			}
		return 1
		}

		CheckBool(bool, FuncName, param) {
			If !(bool=0 || bool=1) {
				throw FuncName ": Parameter [" param "] darf nur 1 oder 0 sein!"
				return 0
			}
		return 1
		}

}

QSFormat(timestamp)                                                                                        	{
	FormatTime, timeformated, % timestamp "000000", yyyy.MM.dd
return timeformated
}

DBFiles_FieldFromPos(pos)                                                                                 	{
	global dbfiles, dbname
	For flabel, field in dbfiles[dbname].dbfields
		If (field.pos = pos)
			break
return flabel
}

DBFiles_PosFromField(field)                                                                               	{
global dbfiles, dbname
return dbfiles[dbname].dbfields[field].pos
}

Edit_SetMargin(hEdit, mLeft:=0, mTop:=0, mRight:=0, mBottom:=0) {
	VarSetCapacity(RECT, 16, 0 )

	; SendMessage, 0xB2,, &RECT,, ahk_id %hEdit% ; EM_GETMARGIN
	DllCall("GetClientRect", "ptr", hEdit, "ptr", &RECT)
	right  := NumGet(RECT, 8, "Int")
	; bottom := NumGet(RECT, 12, "Int")

	static dpi := A_ScreenDPI / 96
	NumPut(0     + Ceil(mLeft*dpi) , RECT, 0, "Int")
	NumPut(0     + Ceil(mTop*dpi)  , RECT, 4, "Int")
	NumPut(right - Ceil(mRight*dpi), RECT, 8, "Int")
	; NumPut(bottom - mBottom, RECT, 12, "Int")
	SendMessage, 0xB3, 0x0, &RECT,, ahk_id %hEdit% ; EM_SETMARGIN
}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Listview Funktionen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
LV_DeleteCols(GuiName, LVName, ColRange:="All" )                                        	{

	Gui, % GuiName ": ListView", % LVName
	resultCols := LV_GetCount("Column")
	If (resultCols > 1)
		If (ColRange = "All") {
			LV_Delete()
			Loop % resultCols-1
				LV_DeleteCol(resultCols-(A_Index-1))
		}

}

LV_GetHeader(hLV, GuiName, LVName)                                                            	{

	LVHeader           	:= Object()
	LVHeader.name	:= Object()
	LVHeader.nr       	:= Array()
	LVHeader.width    	:= Object()

	Gui, % GuiName ": ListView", % LVName

	Loop, % LV_GetCount("Colum") {
		LV_GetText(colName, 0, A_Index)
		width := LVM_GetColWidth(hLV, A_Index)
		LVHeader.name[colName] 	:= A_Index
		LVHeader.width[colName] 	:= width
		LVHeader.nr.Push(colName)
	}

return LVHeader
}

LV_ColumnGetSize(hLV, colNr:=0)                                                                    	{	; Pixelbreite des Listviewsteuerelementes erhalten

	If (colNr > 0)
		return LVM_GetColWidth(hLV, c)
	else
		return LV_GetHeader(hLV, "QS", "QSDBR")

}

LV_RowIsChecked(hLV, row)                                                                               	{
	SendMessage, 0x102C, row-1, 0xF000,, % "ahk_id " hLV  ; 0x102C ist LVM_GETITEMSTATE. 0xF000 ist LVIS_STATEIMAGEMASK.
return (ErrorLevel >> 12) - 1
}

LVM_GetColOrder(h)                                                                                         	{

	; h = ListView handle.

	hdrH := DllCall("SendMessage", "uint", h, "uint", 4127) 						; LVM_GETHEADER
	hdrC := DllCall("SendMessage", "uint", hdrH, "uint", 4608) 					; HDM_GETITEMCOUNT
	VarSetCapacity(o, hdrC * A_PtrSize)
	DllCall("SendMessage", "uint", h, "uint", 4155, "uint", hdrC, "ptr", &o) 	; LVM_GETCOLUMNORDERARRAY
	Loop, % hdrC
		result .= NumGet(&o, (A_Index - 1) * A_PtrSize) + 1 . ","

Return SubStr(result, 1, StrLen(result)-1)
}

LVM_SetColOrder(h, c)                                                                                      	{ ; ##### ÄNDERN!

	; h = ListView handle.
	; c = 1 based indexed comma delimited list of the new column order.
	c := StrSplit(c, ",")
	VarSetCapacity(c, c0 * A_PtrSize)
	Loop, % c0
		NumPut(c%A_Index% - 1, c, (A_Index - 1) * A_PtrSize)

Return DllCall("SendMessage", "uint", h, "uint", 4154, "uint", c0, "ptr", &c) ; LVM_SETCOLUMNORDERARRAY
}

LVM_GetColWidth(hLV, c)                                                                                  	{

	; h = ListView handle.
	; c = 1 based column index to get width of.

Return DllCall("SendMessage", "uint", hLV, "uint", 4125, "uint", c-1, "uint", 0) ; LVM_GETCOLUMNWIDTH
}

LVM_SetColWidth(hLV, cLV, wLV=-1)                                                                 	{
	; hLV = ListView handle.
	; cLV = 1 based column index to get width of.
	; wLV = New width of the column in pixels. Defaults to -1. The following values are supported in report-view mode:
	;	"-1" - Automatically sizes the column.
	;	"-2" - Automatically sizes the column to fit the header text. If you use this value with the last column, its width is set to fill the remaining width of the list-view control.
	Return DllCall("SendMessage", "uint", hLV, "uint", 4126, "uint", cLV-1, "int", wLV) ; LVM_SETCOLUMNWIDTH
}
;}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
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
#Include %A_ScriptDir%\..\..\lib\class_LV_InCellEdit.ahk
#Include %A_ScriptDir%\..\..\lib\class_Neutron.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\lib\Sift.ahk
;}


/*
[Einstellungen]
QSMinKB=
QSMaxKB=
QSFilter=Patien
QSBlockSize=4000








 */

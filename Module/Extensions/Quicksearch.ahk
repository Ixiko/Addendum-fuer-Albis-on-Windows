; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                 	Addendum Quicksearch  - - geschrieben zur visuellen Analyse und Untersuchung von dBase Datenbankdateien
;
;
;
;      Funktionen:           	- Anzeige von Datenbankinhalten und Suche in allen Datenbanken
;                               	- Anzeige vorhandener Feldnamen (Tabellenkopf/Spaltennamen) mit maximaler Zeichenlänge und Zeichentyp
;                                	- Tabellenspalten lassen sich aus- und wieder einblenden
;									- das veränderte Tabellenlayout (inkl. Spaltenbreite) wird nach Skriptstart oder nach einem Tabellenwechsel wiederhergestellt
;									- Patientennamen werden bei einem einfachen Klick auf einen Datensatz angezeigt
;									- Karteikarten werden nach einem Doppelklick in Albis geöffnet
;									- verknüpfte Daten werden nach einem einfachen Klick angezeigt (unterstützt derzeit nur Daten aus BEFTEXTE.dbf)
;									- Suchergebnisse und verbundene Daten können exportiert werden
;									- die eingebenen Suchparameter können mit einer Bezeichnung gespeichert werden, um sie erneut verwenden zu können
;
;		Hinweis:				beta Version, Nutzung auf eigenes Risiko!
;
;		Abhängigkeiten:	siehe includes
;
;	                    			Addendum für Albis on Windows
;                        			by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;									begin: 02.04.2021,	last change 30.07.2021
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
	ListLines                        	, Off
	SetWinDelay                    	, -1
	SetControlDelay            	, -1
	AutoTrim                       	, On
	FileEncoding                 	, UTF-8

	Menu, Tray, Icon, % "hIcon: " Create_QuickSearch_ICO()

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
	;Props := new Quicksearch_Properties(A_ScriptDir "\resources\Quicksearch_UserProperties.json")
	Props := new Quicksearch_Properties(adm.DBPath "\QuickSearch\resources\Quicksearch_UserProperties.json")

  ; Patientendaten laden
	filter   	:= ["NR", "PRIVAT", "GESCHL", "NAME", "VORNAME", "GEBURT",  "PLZ", "ORT"
					, "STRASSE", "HAUSNUMMER", "TELEFON", "TELEFON2", "TELEFAX", "ARBEIT", "LAST_BEH", "MORTAL"]
	PatDB 	:= ReadPatientDBF(adm.AlbisDBPath,,, "allData")
;}

	QuickSearch_Gui()

return

^!+::Reload
#IfWinActive Quicksearch
Esc::ExitApp
#IfWinActive

Quicksearch(dbname, SearchString)                                                                   	{	; Datenbanksuche + Gui Anzeige

		global dbfiles, dbfdata, InCellEdit

	; Gui-Daten
		Gui, QS: Default
		Gui, QS: Submit, NoHide
		GuiControl, QS:, QST4, % ""

	; InCell Edit stoppen
		InCellEdit.OnMessage(false)

		pattern := Object()
		For idx, line in StrSplit(SearchString, "`n") {
			RegExMatch(Trim(line), "(?<field>.*)?\s*\=\s*(?<rx>.*)$", m)
			If mrx {
				;pattern[mfield] := "rx:" mrx
				pattern[mfield] := mrx
				t .= mfield " = " pattern[mfield] "`n"
			}
		}

		dbf       	:= new DBASE(adm.AlbisDBPath "\" dbname ".dbf", 2)
		filepos   	:= dbf.OpenDBF()
		dbfdata  	:= dbf.SearchFast(pattern, "all fields", 0	, 		{ "SaveMatchingSets"    	: false
																							,  "MSetsPath"               	: A_Temp "\QuickSearch_" dbname ".txt"
																							, "ReturnDeleted"          	: QSDBRView }
																		         			, "QuickSearch_Progress")
		filepos   	:= dbf.CloseDBF()

	; Suchergebnis im Props Objekt ablegen (kann auf Festplatte gespeichert werden)
		Gui, QS: ListView, QSDBR
		LV_Delete()

	; Spalten zusammenstellen
		equalIDs := Object()
		For rowNr, set in dbfdata {

			; weiter wenn gelöschte Datensätze nicht angezeigt werden sollen
				If !Props.ShowRemovedSets() && set.removed
					continue

			; Spalten zusammenstellen
				cols := Object()
				For flabel, val in set {
					If Props.isVisible(fLabel) {

						If (dbname = "BEFUND") && RegExMatch(flabel, "i)(INHALT|KUERZEL)") && StrLen(val) = 0
							continue

						If RegExMatch(flabel, "i)" Props.DateLabels) && (dbname <> "CAVE")
							val := RegExMatch(val, "\d{8}") ? ConvertDBASEDate(val) : val
						else if RegExMatch(flabel, "i)" Props.DateLabels) && (dbname = "CAVE")
							val := RegExMatch(val, "\d{8}") ? SubStr(val, 1, 2) "." SubStr(val, 3, 2) "." SubStr(val, 5, 4) : val
						else if RegExMatch(flabel, "i)DATETIME")
							FormatTime, val, % val, dd.MM.yyyy HH:mm
						else if RegExMatch(flabel, "i)VERSART")
							val := Props.VERSART[val]
						else if RegExMatch(flabel, "i)PATNR")
							equalIDs[val] := !equalIDs.haskey(val) ? 1 : equalIDs[val] + 1

						cols[Props.HeaderVirtualPos(flabel)] := val

					}

				}

			; Zeile hinzufügen
				cols.1 := set.removed ? "*" cols.1 : cols.1
				LV_Add("", cols*)

		}

		ColTyp := Array()
		DataType 	:= {"N":"Integer", "D":"Integer", "C":"Text", "L":"Text", "M":"Text"}
		For idx, flabel in Props.Labels()
			If Props.isVisible(fLabel)
				LV_ModifyCol( idx, DataType[Props.HeaderField(flabel, "type")] )

		GuiControl, QS:, QST4, % equalIDs.Count()
		InCellEdit.OnMessage(true)

		;~ JSONData.Save(A_Temp "\test.json", equalIDs, true,, 1, "UTF-8")
		;~ Run, % A_Temp "\test.json"

}

QuickSearch_Gui()                                                                                             	{	; Gui + Handler

	;-: VARIABLEN                                              	;{

	  ; GUI
		global 	hQS, QSDB, QSDBFL, QSDBS, QSDBR, QSPGS, QSPGST
		global 	QST1, QST1a, QST2, QST2a, QST3, QST3a, QST4, QST4a, QST5, QSTGB1, QSTGB2, QSTGB3
		global	QSMinKB      	, QSMaxKB    		, QSFilter                                       	; Eingabefelder
		global	QSMinKBText	, QSMaxKBText		, QSFilterText                                  	; Textfelder
		global 	QShDB         	, QShDBFL    		, QShDBR                                         ; hwnd
		global	QSFeldWahlT	, QSFeldWahl		, QSFeldAlle			, QSFeldKeine     	, QSFelderT
		global	QSSearch			, QSSaveRes			, QSSaveAs			, QSSaveSearch	, QSQuit			, QSReload                	; Buttons
		global 	QSBlockNR		, QSBlockSize		, QSLBlock			, QSNBlock
		global 	QSEXTRANFO 	, QSEXTRAS	    	, QSXTRT          	, QSDBSSave    	, QSDBSParam	, QSDBName
		global	QSDBRView  	, QSDBRPatView
		global 	BlockNr			, BlockSize		    	, lastBlockSize		, lastBlockNr			, dbfRecords  	, lastSavePath
		global 	QSmrgX			, QSmrgY				, QSLvX, QSLvY, QSLvW, QSLvH
		global	InCellEdit
		global 	ReloadGui := false

	  ; ANDERES
		global dbfiles, DBColW, DBFilesW, dbname

		static	QSBlockNRBez, QSEXTRASx, QSReloadw
		static	EventInfo, empty := "- - - - - - -"
		static	FSize            	:= (A_ScreenHeight > 1080 ? 9 : 8)
		static	gcolor          	:= "5f5f50"
		static	tcolor           	:= "DDDDBB"
		static	DBField
		static	DBSearchW   	:= 300
				DBFilesW   	  	:= 310
		static	GFR              	:= " gQuickSearch_DBFilter "
		static	gQSDBFields		:= ""
		static	BGT              	:= " BackgroundTrans "
		static	ASM              	:= " AltSubmit "
		static	ColorTitles     	:= " cWhite"
		static	ColorData      	:= " cBlack"

		BlockNr := 0
		dbstructsPath := A_ScriptDir "\resources\AlbisDBase_Structs.json"

		; Datenbank Tabelle
		DBColW  		:= [80, 65, 80, 75]
		LVWidth1  	:= 20                                  ; Rand der Listview
		For idx, colWidth in DBColW
			LVWidth1 += colWidth

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
		IniRead, lastSavePath	, % adm.Ini, % "QuickSearch", % "lastSavePath"  	, % ""


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

	;-: ZEILE 1: AUSWAHL / SUCHE                 	;{

	;-: FELD 1: DATENBANKEN                       	;{

		; - Filter -
		Gui, QS: Font	, % "s" FSize-1 " q5 Normal" ColorTitles, Futura Bk Bt
		Gui, QS: Add	, Text        	, % "xm ym-3	 " BGT "               	vQSMinKBText"                                                              	, % "min KB"
		Gui, QS: Add	, Text        	, % "x+10 	 " BGT "               	vQSMaxKBText"                                                             	, % "max KB"
		Gui, QS: Add	, Text        	, % "x+10 	 " BGT "             	vQSFilterText"                                                                	, % "DB-Filter"

		; - Felder -
		Gui, QS: Add	, GroupBox  	, % "x+10 w50 h10                                                                               	vQSFeldWahlT"	, % "Felder wählen"

		Gui, QS: Font	, % "s" FSize " q5 Normal" ColorData
		Gui, QS: Add	, Edit           	, % "xm y+4 	w" FSize*8 	" Number  	" BGT "    	" GFR " -E0x200        	vQSMinKB	"  	, % QSMinKB
		Gui, QS: Add	, Edit           	, % "x+5    	w" FSize*8 	" Number  	" BGT "    	" GFR " -E0x200        	vQSMaxKB	"  	, % QSMaxKB

		; - Datenbank Filter/Name -
		GuiControlGet, cp, QS:Pos	, QSMaxKB
		QSFilterW := LVWidth1 + 5 - cpX - cpW - 5
		Gui, QS: Add	, Edit           	, % "x+5    	w" QSFilterW	" hwndhEdit	" BGT ASM GFR " -E0x200         	vQSFilter   	" 	, % QSFilter

		Gui, QS: Add	, Button        	, % "x+9                                        	 " BGT "                                          	vQSFeldAlle	"	, % "Alle"
		Gui, QS: Add	, Button        	, % "x+5                                       	 " BGT "                                          	vQSFeldKeine"	, % "Keine"

		Gui, QS: Font	, % "s" FSize+3 " q5 Italic" ColorTitles
		Gui, QS: Add	, Text        	, % "x+15 ym+10 " BGT "    Center                                                            vQSFelderT 	"  	, % "akt. Datenbank:"

		Gui, QS: Font	, % "s" FSize+6 " q5 Bold " ColorTitles
		Gui, QS: Add	, Text        	, % "x+15  ym+5 w200    " BGT "    Center                                                 vQSDBName" 	, % (lastdbname ? "[" lastdbname ".dbf]")


		; - Filter Elemente versetzen -
		GuiControlGet, cp, QS:Pos	, QSMinKBText
		GuiControlGet, dp, QS:Pos	, QSMinKB
		GuiControl, QS: Move        	, QSMinKBText	, % "x" dpX + Floor((dpW/2) - (cpW/2))
		GuiControl, QS: Move        	, QSMinKB       	, % " h" cpH + 2

		GuiControlGet, cp, QS:Pos	, QSMaxKBText
		GuiControlGet, dp, QS:Pos	, QSMaxKB
		GuiControl, QS: Move        	, QSMaxKBText	, % "x" Floor(dpX + (dpW//2 - cpW//2))
		GuiControl, QS: Move        	, QSMaxKB   	, % " h" cpH + 2

		GuiControlGet, cp, QS:Pos	, QSFilterText
		GuiControlGet, dp, QS:Pos	, QSFilter
		GuiControl, QS: Move        	, QSFilterText	, % "x" Floor(dpX + (dpW/2 - cpW/2))
		GuiControl, QS: Move        	, QSFilter         	, % " h" cpH + 2

		GuiControl, QS: Move        	, QSFeldAlle   	, % " h" cpH + 2
		GuiControl, QS: Move        	, QSFeldKeine 	, % " h" cpH + 2


		; - Felder wählen versetzen -
		gbY := cpY-1,	gpmrg := 2
		GuiControlGet, cp, QS:Pos, QSFeldAlle
		GuiControlGet, dp, QS:Pos, QSFeldKeine
		GuiControl, QS: Move, QSFeldWahlT, % "x" cpX-2*gpmrg " y" gbY " w" (dpX+dpW-cpX+4*gpmrg) " h" cpH+10*gpmrg    ; ist die Groupbox

		; - Datenbankauswahl
		LVOptions1 	:= "-E0x200 vQSDB HWNDQShDB gQS_Handler AltSubmit -Multi Grid"
		GuiControlGet, cp, QS:Pos, QSMinKB
		GuiControlGet, dp, QS:Pos, QSFilter
		mrgX := cpX, DBFilesW := dpX+dpW-mrgX, DBFieldX := dpX+dpW+5

		lp 	        	:= GetWindowSpot(QShDB)
		Gui, QS: Font	, % "s" FSize " q5 Normal" ColorData
		Gui, QS: Add	, ListView    	, % "xm				 y+2 	w" LVWidth1 	" r16  " LVOptions1                                               	, % "Datenbank|Records|Größe|Änderung"

		; - Datenbanken anzeigen
		QuickSearch_DBFilter()

		; - Spaltenbreite einrichten
		;      DBColW.4 	:= lp.CW - DBColW.1 - DBColW.2 - DBColW.3
		Gui, QS: ListView, QSDB
		LV_ModifyCol()
		LV_ModifyCol(1, DBColW.1)
		LV_ModifyCol(2, DBColW.2 " Right Integer")
		LV_ModifyCol(3, DBColW.3 " Right Integer")
		LV_ModifyCol(4, DBColW.4 " Left")
		LV_ModifyCol(2, "SortDesc")

	;}

	;-: FELD 2: DATENBANKFELDER                   	;{
		GuiControlGet, cp, QS:Pos, QSDB
		LVOptions2	    	:= " -Readonly -E0x200 -LV0x18 AltSubmit -Multi Grid Checked vQSDBFL HWNDQShDBFL gQuickSearch_DBFields"
		LVColumnSizes2 	:= ["38 Right Integer", "85 Text", "31 Center", "45 Center", "45 Center", "230 Text"]
		DBFieldW            	:= 0
		For colNr, colOption in LVColumnSizes2 {
			RegExMatch(colOption, "\d+", colWidth)
			DBFieldW += colWidth
		}
		DBFieldW += 20

		Gui, QS: Font	, % "s" FSize " q5 Normal" ColorData
		Gui, QS: Add	, ListView    	, % "x" DBFieldX " y" cpY " w" DBFieldW	" h" cpH " " LVOptions2   	, % "Nr|Feldname|Typ|Länge|Nr.DB|Suchparameter"

		; Größenanpassung
		Gui, QS: Show, Hide
		Gui, QS: ListView, QSDBFL
		For colNr, colOption in LVColumnSizes2
			LV_ModifyCol(colNr, colOption)

		; letzte Spalte editierbar machen
		InCellEdit := New LV_InCellEdit(QShDBFL)
		InCellEdit.SetColumns(6)

		GuiControlGet, dp, QS:Pos, QSDBFL
		GCol3X 	:= dpX+dpW+5
	;}

	;-: FELD 3: DATENBANKSUCHE                 	;{
		GuiControlGet, cp, QS:Pos, QSFelderT
		GuiControlGet, dp, QS:Pos, QSFeldWahlT
		Gui, QS: Font	, % "s" FSize+2 " q5 Italic" ColorTitles, Futura Bk Bt
		Gui, QS: Add	, Text        	, % "x" GCol3X " y" cpY " w" DBSearchW " " BGT                                                       	, % "RegEx Suche"

		GuiControlGet, cp, QS:Pos, QSDB
		GuiControlGet, dp, QS:Pos, QSDBFL
		opt := " -E0x200 vQSDBS HWNDQShDBS gQS_Handler"
		Gui, QS: Font	, % "s" FSize " q5 Normal" ColorData
		Gui, QS: Add	, Edit         	, % "x" dpX+dpW+5 " y" dpY 	" w" DBSearchW " h" dpH . opt                                    	, % ""

		Gui, QS: Font	, % "s" FSize-1 " q5 Normal" ColorData
		Gui, QS: Add	, Button        	, % " y+2 vQSDBSSave  h" FSize-1+2 "	gQS_Handler"                                           	, % "Speichern"
		GuiControlGet, cp, QS:Pos, QSDBSSave
		GuiControlGet, dp, QS:Pos, QSDBS

		Gui, QS: Font	, % "s" FSize+1 " q5 Normal" ColorData
		Gui, QS: Add	, ComboBox 	, % " x+1 y" cpY " w" dpW-cpW-2+100 " h" FSize " r8 vQSDBSParam 	gQS_Handler"  	, % ""

	;}

	;-: FELD 4: ZUSÄTZLICHE INFO'S                	;{
		GuiControlGet	, cp              	, QS: Pos	, QSFelderT
		Gui, QS: Font	, % "s" FSize+2 " q5 Italic" ColorTitles
		Gui, QS: Add	, Checkbox  	, % "x" dpX+dpW+5 "  y" cpY " w400 " BGT " vQSXTRT gQS_Handler" 	, % "verknüpfte Daten"

		GuiControlGet	, cp           	, QS: Pos	, QSXTRT
		GuiControl    	, QS: Move	, QSXTRT	, % "y" dpY - cpY - 5   	; Position anpassen
		GuiControl    	, QS:        	, QSXTRT	, % ExtrasStatus           	; Checkbox setzen oder entfernen
		GuiControlGet	, cp           	, QS: Pos	, QSXTRT

		Gui, QS: Font	, % "s" FSize+1 " q5 Normal Bold Underline" ColorData
		Gui, QS: Add	, Edit         	, % "x" cpX " y" dpY " w400 h" cpH " -E0x200 Center vQSEXTRANFO"

		GuiControlGet	, ep           	, QS: Pos	, QSEXTRANFO
		Gui, QS: Font	, % "s" FSize " q5 Normal" ColorData
		Gui, QS: Add	, Edit         	, % "y+1w400 h" dpH-epH-1 " -E0x200 vQSEXTRAS"

		GuiControl, % "QS:" (ExtrasStatus ? "Enable" : "Disable"), % "QSEXTRANFO"
		GuiControl, % "QS:" (ExtrasStatus ? "Enable" : "Disable"), % "QSEXTRAS"

	;}

	;}

	;-: ZEILE 2: BEFEHLE                                    	;{

		GuiControlGet, dp, QS: Pos , QSEXTRAS
		QSEXTRASx 	:= dpX, GuiW := dpX + dpW-5
		row1H       	:= dpY + dpH
		GuiControlGet, dp, QS: Pos , QSDBSSave
		row2Y         	:= dpY + dpH

		;-: FELD 1: INFO'S                                  	;{
		Gui, QS: Font	, % "s" FSize+1 " q5" ColorTitles, Futura Md Bt

		Gui, QS: Add	, GroupBox	, % "xm     	y" row1H-6  " w100 h50" 	 "  vQSTGB1"

		Gui, QS: Add	, Text        	, % "xm+3  	y" (row1H+3)       BGT " vQST1a"                              	, % "Position:"
		Gui, QS: Add	, Text        	, % "x+2                            	 " BGT " vQST1"                                	, % empty "   "

		Gui, QS: Add	, Text        	, % "xm+3  	y+1                	 " BGT " vQST2a"                              	, % "Datensätze:"
		Gui, QS: Add	, Text        	, % "x+2                           	 " BGT " vQST2"                                	, % empty "   "

		GuiControlGet, cp, QS: Pos , QST2
		GuiControlGet, dp, QS: Pos , QST2a
		col2X := dpW + cpW + 20, InfoH := dpH+cpH+12
		GuiControl, QS: Move, QST1a  	, % "w" 	dpW
		GuiControl, QS: Move, QST1    	, % "x" 	dpX+dpW+2
		GuiControl, QS: Move, QSTGB1	, % "w" 	dpW+cpW+6 " h" InfoH

		; - - - - -

		Gui, QS: Add	, GroupBox	, % "x" col2X      " y" row1H-6  " w100 h" InfoH "  vQSTGB2"

		Gui, QS: Add	, Text        	, % "x" col2X+3 " y" (row1H+3)	   BGT " vQST3a"                   	, % "Treffer gesamt:"
		Gui, QS: Add	, Text        	, % "x+2                               	 " BGT " vQST3"                        	, % empty " "

		Gui, QS: Add	, Text        	, % "x" col2X+3 " y+1             	 " BGT " vQST4a"                      	, % "PatNR:"
		Gui, QS: Add	, Text        	, % "x+2                               	 " BGT " vQST4"                     	, % empty " "

		GuiControlGet, cp, QS: Pos , QST2
		GuiControlGet, dp, QS: Pos , QST3a
		col2X := dpW + cpW + 20
		GuiControl, QS: Move, QST4a  	, % "w" 	dpW
		GuiControl, QS: Move, QST4    	, % "x" 	dpX+dpW+2
		GuiControl, QS: Move, QSTGB2	, % "w" 	LVWidth1-dpX+mrgX+3

		; - - - - -

		Gui, QS: Add	, GroupBox	, % "x" DBFieldX     	" y" row1H-6  " w" DBFieldW " h" InfoH "  vQSTGB3"
        Gui, QS: Add	, Text        	, % "x" DBFieldX+3  	" y" row1H+3        		BGT " "                 	, % "ausgewählter Patient:"
		Gui, QS: Add	, Text        	, % "x" DBFieldX+3 	" y+1 w" DBFieldW-6		BGT " vQST5"       	, % ""


		;}

		;-: FELD 2: PROGRESS                              	;{
		Gui, QS: Add	, Progress     	, % "xm  	y+3 w" GuiW " vQSPGS"			                                		, % 0
		GuiControlGet, cp, QS: Pos, QSPGS
		Gui, QS: Add	, Text        	, % "x" cpX " y" cpY " w" cpW " h" cpH "  " BGT "   Center vQSPGST", % "+*~"
		;}

		;-: FELD 3: BEFEHLE                               	;{
		BtnH := cpH-8
		Gui, QS: Add	, Button        	, % "      	y+2              	vQSSearch                       	gQS_Handler "	, % "Suchen"

		Gui, QS: Add	, Button        	, % "x+10		                	vQSSaveRes                    	gQS_Handler "	, % "Suche speichern als"

		Gui, QS: Font	, % "s" FSize+3 " q5" ColorData
		GuiControlGet	, cp, QS: Pos, QSSaveRes
		Gui, QS: Add	, DDL         	, % "x+2 h" cpH " w115   	vQSSaveAs r5                 	gQS_Handler "	, % ".csv (;)|.csv (Tab)|.csv (,)"
		GuiControl    	, QS: Choose, QSSaveAs, 1

		Gui, QS: Font	, % "s" FSize+1 " q5" ColorData
		Gui, QS: Add	, Button        	, % "x+20		          	vQSLBlock                             	gQS_Handler "	, % "⏪ vorheriger Block"

		Gui, QS: Font	, % "s" FSize-2 " q5" ColorTitles
		GuiControlGet	, cp, QS: Pos, QSLBlock
		Gui, QS: Add	, Text         	, % "x" cpX+cpW+5 " y" cpY-3 "	vQSBlockNrBez " BGT                     	, % "Block Nr|Größe"

		Gui, QS: Font	, % "s" FSize-1 " q5" ColorData
		Gui, QS: Add	, Edit         	, % "y+0 w" FSize*4 " vQSBlockNR Number              	gQS_Handler "	, % BlockNr
		Gui, QS: Add	, Edit         	, % "x+0 w" FSize*6 " vQSBlockSize Number				gQS_Handler "	, % BlockSize

		Gui, QS: Font	, % "s" FSize+1 " q5" ColorData
		Gui, QS: Add	, Button        	, % "x+5 y" cpY "	  	vQSNBlock                             	gQS_Handler "	, % "nächster Block ⏩"

		Gui, QS: Font	, % "s" FSize-1 " q5" ColorTitles
		Gui, QS: Add	, Checkbox  	, % "x+10  y" cpY " " 	BGT "          	vQSDBRView     	gQS_Handler" 	, % "gelöschte Datensätze anzeigen"
		Gui, QS: Add	, Checkbox  	, % "y+3 "             	BGT "          	vQSDBRPatView   	gQS_Handler" 	, % "PATNR mit Namen ersetzen"

		Gui, QS: Font	, % "s" FSize+1 " q5" ColorData
		Gui, QS: Add	, Button        	, % "x+250 y" cpY "              	      	vQSReload          	gQS_Handler "	, % "Reload"
		;}

	;}

	;-: ZEILE 3: AUSGABE                                	;{
		Gui, QS: Show, %  WinSize " Hide", Quicksearch

		GuiControlGet, cp, QS: Pos, QSReload
		QSReloadw	:= cpW
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

		WinMove, % "ahk_id " hQS,,,, % wqs.W+1	, % wqs.H+1
		WinMove, % "ahk_id " hQS,,,, % wqs.W   	, % wqs.H
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
		Hotkey, Left               	, 	% fn_QSHK
		Hotkey, Right             	, 	% fn_QSHK
		Hotkey, Enter             	, 	% fn_QSHK
		Hotkey, NumpadEnter 	, 	% fn_QSHK
		Hotkey, TAB              	, 	% fn_QSHK
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

				If !WinExist("ahk_class OptoAppClass") || !A_EventInfo
					return

				If (Pcol := Props.PatIDColumn()) {
					Gui, QS: ListView, QSDBR
					LV_GetText(PatID, A_EventInfo, Pcol)
					PatID    	:= StrReplace(PatID, "*")
					PatName 	:= PatDB[PatID].NAME ", " PatDB[PatID].VORNAME
					GuiControl, QS:, QST5, % PatDB[PatID].NAME ", " PatDB[PatID].VORNAME
					AlbisAkteOeffnen(PatName, PatID)
				}

			}
			else if RegExMatch(A_GuiEvent, "(Normal|I)"	)  	{

				tblRow := A_EventInfo

				If (!tblRow || Last_tblRow = tblRow)
					return
				Last_tblRow := tblRow

			  ; PatID finden
				If (Pcol := Props.PatIDColumn()) {
					Gui, QS: ListView, QSDBR
					LV_GetText(PatID, tblRow, Pcol)
					PatID := StrReplace(PatID, "*")
					GuiControl, QS:, QST5, % "[" PatID "] " PatDB[PatID].NAME ", " PatDB[PatID].VORNAME
														.	" *" ConvertDBASEDate(PatDB[PatID].GEBURT)
				}

			 ; verlinkte Daten aus anderer Datenbank anzeigen (wenn Häkchen bei verknüpfte Daten gesetzt ist)
				If QSXTRT && Props.LabelExist("TEXTDB")
					QuickSearch_Extras(tblRow)

			}
		;}

		Case "QSLBlock":                        	;{   	vorheriger Datenblock
			QuickSearch_ListFields(dbname, 2)
		;}

		Case "QSNBlock":                     	;{    	nächster Datenblock
			QuickSearch_ListFields(dbname, 1)
		;}

		Case "QSSearch": 	                    	;{     Datenbanksuche
			If !dbname
				return
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

		Case "QSDBSSave":                     	;{ 	Suchparameterliste speichern
			If !dbname
				return
			Quicksearch_SaveSearch(dbname, QSDBSParam, QSDBS)
		;}

		Case "QSDBSParam":                	;{ 	neue Parameterliste wurde ausgewählt

			;SciTEOutput("`n - - - - - - - - - - - label: " QSDBSParam "`n" Props.SParams_GetParams(QSDBSParam))
			Props.SParams_CBSet(QSDBSParam)

		;}

		Case "QSFeldAlle":                     	;{ 	Alle/Keine Spalten der Datenbank anzeigen
		Case "QSFeldKein":

			checkAll := A_GuiControl = "QSFeldAlle" ? true : false
			Gui, QS: ListView, QSDB
			Loop % LV_GetCount()
				LV_Modify(A_Index, (!checkAll ? "-":"") "Check")

		;}

		Case "QSDBRView":                     	;{ 	Props Object Einstellung ändern

			Props.ShowRemovedSets(QSDBRView)

		;}

		Case "QSDBRPatview":                 	;{ 	PATNR durch Patientennamen ersetzen
			Props.ShowPatNames(QSDBRPatView)

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
	QSW := A_GuiWidth, QSH:= A_GuiHeight, QSEXTRASw := QSW-QSEXTRASx-5
	GuiControl, QS: MoveDraw	, QSDBR       	, % "w" QSW-10 " h" QSH-QSmrgY
	GuiControl, QS: MoveDraw	, QSPGS       	, % "w" QSW-10
	GuiControl, QS: MoveDraw	, QSEXTRANFO	, % "w" (QSEXTRASw < 300 ? 300 : QSEXTRASw)
	GuiControl, QS: MoveDraw	, QSEXTRAS   	, % "w" (QSEXTRASw < 300 ? 300 : QSEXTRASw)
	GuiControl, QS: MoveDraw	, QSReload      	, % "x" QSW-QSReloadw-5
	GuiControl, QS: +Redraw 	, QSDBR
	WinSet, Redraw,, % "ahk_id " QShDBR

return ;}

}

QuickSearch_GuiSave()                                                                                     	{	; Fenstereinstellung speichern

	global hQS, QSDB, QSDBFL, QSDBS, QSDBR, QSPGS, QST1, QST2, QST3, QSXTRT, QSDBRView
	global QSMinKB, QSMaxKB, QSFilter, BlockSize, winSize, dbname

	Gui, QS: Default
	Gui, QS: ListView, Default
	Gui, QS: Submit, NoHide

	wqs  	:= GetWindowSpot(hQS)
	winSize :=  "x" wqs.X " y" wqs.Y " w" wqs.CW " h" wqs.CH
	lastS 	:= RegExReplace(QSDBS, "[\r\n]+", "|")

	IniWrite, % QSMinKB 	, % adm.Ini, % "QuickSearch", % "QSMinKB"
	IniWrite, % QSMaxKB	, % adm.Ini, % "QuickSearch", % "QSMaxKB"
	IniWrite, % QSFilter   	, % adm.Ini, % "QuickSearch", % "QSFilter"
	IniWrite, % BlockSize    	, % adm.Ini, % "QuickSearch", % "QSBlockSize"
	IniWrite, % winSize    	, % adm.Ini, % "QuickSearch", % "winSize"
	IniWrite, % dbname    	, % adm.Ini, % "QuickSearch", % "lastDB"
	IniWrite, % lastS         	, % adm.Ini, % "QuickSearch", % "lastSearch"
	IniWrite, % QSXTRT    	, % adm.Ini, % "QuickSearch", % "Extras"

	If !Props.SaveProperties()
		MsgBox, Einstellungen konnten nicht gesichert werden!

return
}

QuickSearch_Hotkeys()                                                                                      	{	; Hotkey Kontexthandler

	global dbname, QS, QSDBS, QSBlockNr, QSBlockSize

	If !WinActive("Quicksearch ahk_class AutoHotkeyGUI") {
		SendInput, % "{" A_ThisHotkey "}"
		return
	}

	thisHotkey := A_ThisHotkey
	Gui, QS: Default
	Gui, QS: Submit, NoHide
	GuiControlGet, gfocus, QS: FocusV
	;SciTEOutput("Focus: " gfocus ", " thisHotkey)

	Switch thisHotkey {

		Case "TAB":
			If RegExMatch(gfocus, "i)(QSBlockNr|QSBlockSize)")
				QuickSearch_ListFields(dbname, 3)
			else if RegExMatch(gfocus, "i)QSDBS")
				SendInput, {Down}
			else
				SendInput, {TAB}

		Case "Enter":
		Case "NumPadEnter":
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

QuickSearch_Extras(tblRow, TEXTDB:=0, TEXTDate:=0)                                         	{	; verknüpfte Daten aus anderer Datenbank laden

		static DBIndex, dbf

		If (!Props.isVisible("TEXTDB")) {
			;SciTEOutput("TEXTDB is visible: " Props.isVisible("TEXTDB") "," Props.HeaderVirtualPos("TEXTDB"))
			return
		}

	; TEXTDB = LFDNR in BEFTEXTE
		Gui, QS: ListView, QSDBR
		If !TEXTDB {
			LV_GetText(TEXTDB, tblRow, Props.HeaderVirtualPos("TEXTDB"))
			If !TEXTDB
				return
		} else
			SaveFeature := true
		If !TEXTDate
			LV_GetText(TEXTDate, tblRow, Props.HeaderVirtualPos("DATUM"))

	; BEFTEXTE.dbf indizieren oder Index laden
		If !IsObject(DBIndex) {

			DBIndex     	:= {"Quartal":{}, "LFDNR":{}}
			DbIndexFile1	:= adm.DBPath "\DBase\DBIndex_BEFTEXTE.json"
			DbIndexFile2	:= adm.DBPath "\DBase\DBIndex_BEFTEXTE_LFDNR-Sort.json"
			dbf           	:= new DBASE(adm.AlbisDBPath "\BEFTEXTE.dbf", 2)

			If FileExist(DbIndexFile1) && FileExist(DbIndexFile2) {
				DBIndex.Quartal	:= JSONData.Load(DbIndexFile1, "", "UTF-8")
				DBIndex.LFDNR 	:= JSONData.Load(DbIndexFile2, "", "UTF-8")
			}

		  ; Indizierung wird auch ausgeführt wenn die letzte Indizierung mehr als 30 Tage her ist
			today := A_YYYY A_MM A_DD
			IndexFromDay := today - 30
			IndexStartRecord := DBIndex.Quartal.LastIndex && (DBIndex.Quartal.LastIndex < IndexFromDay) ? DBIndex.Quartal.LastRecord : 0

		  ; jetzt indizieren
			If !FileExist(DbIndexFile1) || !FileExist(DbIndexFile2) || IndexStartRecord {

				; Hinweisfenster wird eingeblendet ;{
					global hQS, hQSi, QSDBS, QShDBS
					static QSiT, QSi
					op := GetWindowSpot(QShDBS)
					Gui, QSi: new, -Caption +ToolWindow  HWNDhQSi -DPIScale +Owner%hQS% ; +AlwaysOnTop
					Gui, QSi: Color, c000000
					Gui, QSi: Font, s26 q5 bold, Futura Bk Bt
					Gui, QSi: Add, Text, % "x0 y0 w" op.CW " cDarkRed Center vQSiT", % "Indexerstellung`nbitte warten...."
					GuiControlGet, cp, QSi: Pos, QSiT
					GuiControl, QSi: Move, QSiT, % "y" Floor(op.CH/2)-Floor(cpH)
					Gui, QSi: Show, % "x" op.X " y" op.Y " w" op.CW " h" op.CH " NoActivate", QuickSearch...Indexer
					WinSet, Trans, 200, % "ahk_id " hQSi
				;}

				; Indexer erstellt zwei Index Objekte, eines mit Zuordnung zu LFDNR und das andere zum Quartal ;{

						dbf.ShowAt	:= 100                                                                       ; für Fortschrittsanzeige
						records  		:= dbfRecords := dbf.records
						LFDNRsteps  	:= records < 1200 ? Floor(records/100) : 1200
						filepos       	:= dbf.OpenDBF()

					; Index ergänzen wenn letzter Index länger als einen Monat her ist
						If IndexStartRecord
							dbf._SeekToRecord(IndexStartRecord)

					; indizieren
						Loop % (records - IndexStartRecord) {

							;~ If (Mod(A_Index, 100) = 0)
								;~ QuickSearch_Progress(A_Index, records, StrLen(records), matchcount)

							set := dbf.ReadRecord(["PATNR", "DATUM", "LFDNR", "POS", "TEXT"], "QuickSearch_Progress")

						  ; nach Datum/Quartal indizieren
							Quartal := GetQuartalEx(set.DATUM, "YYYYQQ")
							If (LastQuartal <> Quartal) {
								LastQuartal := Quartal
								If !DBIndex.Quartal.haskey(Quartal)
									DBIndex.Quartal[Quartal] := [dbf.recordnr, dbf.filepos]       	; Datensatz Nummer im Objekt als Einsprungpunkt sichern
							}

						  ; nach LFDNR indizieren
							If (Mod(A_Index, LFDNRsteps) = 0) {
								If !DBIndex.LFDNR.haskey(set.LFDNR)
									DBIndex.LFDNR[set.LFDNR] := [dbf.recordnr, dbf.filepos]         	; Datensatz Nummer im Objekt als Einsprungpunkt sichern
							}

						}
				;}

				QuickSearch_Progress(records, records, StrLen(records), DBIndex.Quartal.Count()-3 "+" DBIndex.LFDNR.Count())

				DBIndex.Quartal.LastIndex  	:= dbf.lastupdatedBase
				DBIndex.Quartal.LastRecord 	:= dbf.records
				DBIndex.Quartal.LastFilePos 	:= dbf.filepos

				filepos := dbf.CloseDBF()

				JSONData.Save(DbIndexFile1, DBIndex.Quartal	, true,, 1, "UTF-8")
				JSONData.Save(DbIndexFile2, DBIndex.LFDNR	, true,, 1, "UTF-8")

			}

			Gui, QSi: Destroy
		}

	; Vorbereitung Daten auslesen
		If !IsObject(dbf)
			dbf      	:= new DBASE(adm.AlbisDBPath "\BEFTEXTE.dbf", 2)
		records 	:= dbfRecords := dbf.records
		steps      	:= records/100
		filepos   	:= dbf.OpenDBF()

	; filepointer in Richtung TEXTDB Eintrag versetzen
		IndexStartRecord := 0
		If TEXTDB {
			lastrecData 	:= [1, filepos]
			LFDNRPos 	:= 0
			For LFDNR, recData in DBIndex.LFDNR {
				LFDNRPos ++
				If (LFDNR >= TEXTDB)
					break
				If (LFDNRPos < DBIndex.LFDNR.Count())
					lastrecData := recData
			}
			IndexStartRecord := lastrecData.1
		}
		else If TEXTDate {
			Quartal := GetQuartalEx(TEXTDate, "YYYYQQ")
			IndexStartRecord := DBIndex.Quartal[Quartal]
		}

	; filepointer wird gesetzt
		dbf._SeekToRecord(IndexStartRecord)

	; TEXTDB Eintrag finden und Daten zusammen stellen
		found := false
		Loop % (records-IndexStartRecord) {

			set := dbf.ReadRecord(["PATNR", "DATUM", "LFDNR", "POS", "TEXT"])
			If (TEXTDB = set.LFDNR) {
				found := true
				extra .= set.TEXT
				extraLines := set.POS + 1
			}
			else If (found && TEXTDB <> set.LFDNR) {
				break
			}

			If !SaveFeature && (Mod(A_Index, 100) = 0)
				QuickSearch_Progress(A_Index, records, StrLen(records), matchcount)
		}

		filepos   	:= dbf.CloseDBF()
		If !SaveFeature
			QuickSearch_Progress(records, records, StrLen(records), "---")

	; Daten ausgeben
		If !SaveFeature {
			extrahead	:=  extraLines " Zeile" (extraLines = 1 ? "" : "n")
								. " aus BEFTEXT.dbf für TEXTDB Eintrag: " TEXTDB

			GuiControl, QS:, QSEXTRANFO	, % extrahead
			GuiControl, QS:, QSEXTRAS   	, % extra
		}

return extra
}

QuickSearch_ListFields(dbname, step:=0)                                                             	{ 	; erstellt die Ausgabetabelle

		global	QS, QSDBS, QSBlockNr, QSBlockSize, QSPGS, QSPGST, QSDBSParam, QSDBRView
		global 	dbfiles, empty, dbfdata
		global 	BlockNr, BlockSize, lastBlockNr, lastBlockSize, dbfRecords
		static 	lastdbname

		Gui, QS: Submit, NoHide
		GuiControl, QS:, QST4, % ""
		InCellEdit.OnMessage(false)

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

	; Blockgröße berechnen                                        	;{
		BlockNr 	:= QSBlockNr
		BlockSize 	:= QSBlockSize

		If (StrLen(BlockSize) 	= 0	|| !RegExMatch(BlockSize	, "^\d+$"))
			BlockSize := lastBlockSize
		If !RegExMatch(BlockNr	, "^\d+$")
			BlockNr := lastBlockNr
	;}

	; ----------------------------------------------------------------------------------------------------------------------------------------------------------------
	; erstmalige Anzeige
	; ----------------------------------------------------------------------------------------------------------------------------------------------------------------
		If      	(step = 0) {     	; erste Anzeige

			; Einstellungsobjekt anlegen                              	;{
				If !Props.WorkDB(dbname, dbfiles[dbname].dbfields) {
					throw A_ThisFunc ": Einstellungen zur Datenbank: " dbname " konnten nicht angelegt werden!"
					return
				}
			;}

			; Namen der ausgewählten Datenbank anzeigen
				GuiControl, QS:, QSDBName, % "[" dbname ".dbf]"

			; Datenbanksuche - Parameter/Auswahl
				Props.SParams_CBSet(Props.SParams.lastLabel)

			; Gui Inhalte zurücksetzen
				GuiControl, QS:, QSEXTRASNFO	, % ""                   	; Zusatzinfos leeren
				GuiControl, QS:, QSEXTRAS       	, % ""                  	; Zusatzinfos leeren
				GuiControl, QS:, QST5              	, % ""                  	; ausgewählten Eintrag leeren

			; BlockNr auf Null setzten
				BlockNr := lastBlockNr := 0

			; Ausgabelistview erstellen                              	;{

				LV_DeleteCols("QS", "QSDBR")                           	; Spalten entfernen

				thisColPos := 1
				For colNr, flabel in Props.Labels()
					If Props.isVisible(flabel) && (flabel <> removed) {

						Gui, QS: ListView, QSDBR

						dtype 	:= Props.DataType(flabel)
						width 	:= Props.HeaderField(flabel, "width")
						width	:= RegExReplace(width, "i)[a-z]+")
						;width	:= (!width || width > 600) ? (thisColPos = 1 ? "Auto" :  "AutoHdr") : width

						If (thisColPos = 1) {
							LV_ModifyCol(thisColPos, (!width ? "100":width) " " dtype, flabel)
						}
						else
							LV_InsertCol(thisColPos	 , (!width ? "AutoHdr":width) " " dtype, flabel)

						;LV_ModifyCol(thisColPos, width)
						Props.HeaderSetVirtualPos(flabel, thisColPos)
						thisColPos ++

					}

			;}

			; Datenbankfelderinfo's anzeigen                     	;{
				Gui, QS: ListView, QSDBFL
				LV_Delete()

				For colNr, flabel in Props.Labels() {
					field := Props.HeaderField(flabel)
					LV_Add(Props.isVisible(flabel)?"Check":"", field.col, field.name, field.Type, field.len, field.pos)
				}

			;}

			; Statistik                                                          	;{
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

		lastBlockSize	:= BlockSize
		lastBlockNr	:= BlockNr

	; Checkbox - gelöschte Datensätze anzeigen - einstellen
		If (Props.dbname = "WZIMMER") {
			SRS := Props.ShowRemovedSets()
			Props.ShowRemovedSets(true)
		}

	; 	Controls anpassen
		GuiControl, QS:, QSDBRView	, % Props.ShowRemovedSets()
		GuiControl, QS:, QSBlockNR 	, % BlockNr
		GuiControl, QS:, QSBlockSize	, % BlockSize

	; Daten aus Albis Datenbank lesen                      	;{
		GuiControl, QS:, QSPGST, % "lade Block " BlockNr " ..."
		dbf       	:= new DBASE(adm.AlbisDBPath "\" dbname ".dbf", 2)
		records		:= dbfRecords := dbf.records
		filepos   	:= dbf.OpenDBF()
		dbfdata    	:= dbf.ReadBlock(BlockSize,, BlockNr, "QuickSearch_Progress")
		filepos   	:= dbf.CloseDBF()
		GuiControl, QS:, QSPGST, % ""
		QuickSearch_Progress(BlockNr*BlockSize+1 "-" (BlockNr+1)*BlockSize, records, 0, dbfdata.Count())
	;}

	; Ausgabe der Daten im Ergebnisfenster             	;{
		Gui, QS: Default
		Gui, QS: ListView, QSDBR
		LV_Delete()
		GuiControl, QS: -Redraw, QSDBR
		equalIDs := {}

		For rowNr, set in dbfdata {

			; weiter wenn gelöschte Datensätze nicht angezeigt werden sollen
				If (!Props.ShowRemovedSets() && set.removed)
					continue

			; Spalten zusammenstellen
				cols := Object()
				For flabel, val in set {

					If Props.isVisible(fLabel) {

						If RegExMatch(flabel, "i)" Props.DateLabels)
							If  (dbname <> "CAVE")
								val := RegExMatch(val, "\d{8}") ? ConvertDBASEDate(val) : val
							else
								val := RegExMatch(val, "\d{8}") ? SubStr(val, 1, 2) "." SubStr(val, 3, 2) "." SubStr(val, 5, 4) : val
						else if RegExMatch(flabel, "i)DATETIME")
							FormatTime, val, % val, dd.MM.yyyy HH:mm
						else if RegExMatch(flabel, "i)VERSART")
							val := Props.VERSART[val]
						else if RegExMatch(flabel, "i)PATNR")
							equalIDs[val] := !equalIDs.haskey(val) ? 1 : equalIDs[val] + 1

						cols[Props.HeaderVirtualPos(flabel)] := val
					}

				}

			; Zeile hinzufügen
				cols.1 := set.removed ? "*" cols.1 : cols.1
				LV_Add("", cols*)

		}

		;~ JSONData.Save(A_Temp "\test.json", equalIDs, true,, 1, "UTF-8")
		;~ Run, % A_Temp "\test.json"

		GuiControl, QS: +Redraw, QSDBR
		GuiControl, QS:, QST4, % equalIDs.Count()

		If (Props.dbname = "WZIMMER")
			Props.ShowRemovedSets(SRS)

		InCellEdit.OnMessage(true)


	;}

; 21233 22801 20916
}

QuickSearch_DBFields(GChwnd:="", GuiEvent:="", EventInfo:="", ErrLevel:="")    	{	; Eventhandler DBFields

		global dbfiles, dbname, dbfdata, QShDBFL, QShDBR
		static DataType 	:= {"N":"Integer", "D":"Integer", "C":"Text", "L":"Text", "M":"Text"}
		static evTime := 0
		static lastClickTime, ClicksMeasured

		/*  GuiEvent Bechreibung
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

		;Critical

	; Checkboxstatus hat sich geändert
		If RegExMatch(GuiEvent . EventInfo, "(K32|C0)") {

		  ; Feldinfo Listview auf Default setzen und gLabel Aufrufe ausstellen
			Gui, QS: ListView	, QSDBFL
			GuiControl, QS: -g, QSDBFL

		  ; zeilenweises Auslesen des Checkboxstatus
			FieldCount 	:= 0
			FieldRows  	:= LV_GetCount()
			Loop % FieldRows {

				FieldCount  	++
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

Quicksearch_SaveSearch(dbname, textlabel, SearchString)                                   	{	; Suchparameter sichern

		global dbfiles, dbfdata

	; zurück bei leerer Bezeichnung
		If !textlabel {
			MsgBox, Geben Sie zunächst eine Beschreibung für die Suche ein!
			return
		}

	; Suchparameter-String übergeben und sichern
		If Props.SParams.Haskey(textlabel) {

			If params {
				MsgBox, 0x4, QuickSearch, % "Die Beschreibung:`n<<" textlabel ">>`nist schon vergeben.`nSoll diese überschrieben werden?"
				IfMsgBox, No
					return
			}
			else {
				MsgBox, 0x4, QuickSearch, %	"Es sind keine Suchparameter zum Speichern eingegeben worden`n"
															. 	"und die Beschreibung:`n<<" textlabel ">>`nist schon vergeben.`n"
															. 	"JA      zum Entfernen der Suche aus Quicksearch.`n"
															. 	"NEIN um die bisherige Suche zu behalten!"
				IfMsgBox, No
					return
			}

		}

	; prüft ob eine leere Parameterkette übergeben wurde und erstellt einen neuen String (nur Parameter mit Werten)
		For idx, line in StrSplit(SearchString, "`n") {
			RegExMatch(Trim(line), "(?<field>.*)?\=(?<rx>.*)$", m)
			If mrx
				params .= (idx > 1 ? "##" : "") mfield "=" mrx
		}
		params := LTrim(params, "##")

	; Speichern
		Props.SParams.lastLabel := textlabel
		Props.SParams[textlabel] := params
		Props.SParams_Save()

	; Combobox anpassen
		GuiControl, QS:, QSDBSParam	, % "|" Props.SParams_GetLabels()                        	; Combobox befüllen ("|" am Anfang ersetzt die Liste)
		GuiControl, QS: ChooseString	, QSDBSParam, % textlabel                                	; zuletzt angezeigte Suche vorwählen

}

QuickSearch_SaveResults(format:="csv (Tab)", lastSavePath:="")                            	{	; Ausgabetabelle speichern

		global QS, QSDBR, QShDBR

		Gui, QS: Default
		Gui, QS: ListView, QSDBR
		Gui, QS: Submit, NoHide

		If !(rows := LV_GetCount())
			return

		RegExMatch(format, "(?<Ext>\w+)\s+\((?<Delim>[\w,;\|]+)", File)
		DelimiterChar := FileDelim = "Tab" ? A_Tab : FileDelim

		FileSelectFile, filesavepath,, % (lastSavePath ? lastSavePath : A_ScriptDir), Dateiverzeichnis und Dateinamen wählen, % "*." FileExt
		SplitPath, filesavepath, outname, outdir
		If !filesavepath || (StrLen(outname) = 0)
			return

	; als letzten aufgerufenen Dateipfad sichern
		IniWrite, % outdir, % adm.Ini, % "QuickSearch", % "lastSavePath"

		table         	:= ""
		tableheader 	:= LV_GetHeader(QShDBR, "QS", "QSDBR")
		TEXTDBCol  	:= 0

	; Überschriftenzeile erstellen, "TEXTDB" Spalte registrieren
		For col, label in tableheader.nr {
			table .= (col > 1 ? DelimiterChar : "") label
			If QSXTRT && (label = "TEXTDB")
				TEXTDBCol := col
		}

	; TEXTDB Überschrift hinzufügen wenn vorhanden
		If TEXTDBCol
			table .= DelimiterChar "TEXTDB Inhalt"
		table .= "`n"

	; sichert auch Daten aus der BEFTEXT.dbf wenn es eine Spalte TEXTDB gibt
		maxCols	:= col
		steps     	:= rows >= 100 ? Floor(rows/100) : 1
		Loop, % rows {

			Gui, QS: ListView, QSDBR

			row := A_Index
			If (Mod(A_Index, steps) = 0)
				QuickSearch_Progress(A_Index, rows, StrLen(rows), "---")

			TEXTDBContent := ""
			Loop % maxCols {

				col := A_Index
				LV_GetText(cell, row, col)
				If (col = TEXTDBCol) && (cell > 0)
					TEXTDBContent := QuickSearch_Extras(row, cell)
				table .= cell (col < maxCols ? DelimiterChar : TEXTDBContent ? DelimiterChar TEXTDBContent "`n" : "`n")

			}

		}
		QuickSearch_Progress(rows, rows, StrLen(rows), "---")

	; Tabelle in eine Datei speichern
		FileOpen(filesavepath, "w", "UTF-8").Write(table)
		If FileExist(filesavepath) {
			FileGetSize, FSize, % filesavepath
			MsgBox, % "Die Datei wurde gespeichert.`nDatensätze: " rows "`nDateigröße: " (FSize > 1024 ? Round(FSize/1024, 1) " MB" : FSize " kb")
		} else
			MsgBox, Die Datei wurde nicht gespeichert!

}

class Quicksearch_Properties                                                                              	{	; Tabellenmanager

		__New(propertiesPath) {

			SplitPath, propertiesPath,, outFileDir,, outFileNameNoExt
			this.FileDir     	:= outFileDir
			this.FileName 	:= outFileNameNoExt ".json"
			this.FilePath     	:= this.FileDir "\" this.FileName
			this.DataPath		:= RegExReplace(this.FileDir, "\\.*$", "")

			If FilePathCreate(this.DataPath)
				If FilePathCreate(this.FileDir) && FileExist(this.FilePath)
					this.qs := this.LoadProperties()
				else
					this.qs := Object()

			this.VERSART   	:= ["Mitglied", "2", "Angehöriger", "4", "Rentner"]
			this.DateLabels	:= "(STAT_VON|STAT_BIS|DATUM|EINLESTAG|AUS2DAT|GUELTIG|GUELTVON|GUELTVON2|GUELTBIS|"
										. "GUELTBIS2|GEBURT|VERSBEG|GEBDATBIS|GEBDATVON|LAST_B_FR|LAST_B_TO|PATGRFROM|PATGRTO|"
										. "PATIENTVON|PATIENTBIS|AUFNAME|ENDE|VALIDFROM|AUSSTDAT|DATSPEIC|VALIDUNTIL|GUELT_VON|GUELT_BIS|"
										. "PRDATE|BEZAHLTAM|STORNODATE)|MAHNDATE"

		}

		WorkDB(dbname, dbfields)                  	{	; Arbeitsdatenbank festlegen und Tabellen-Informationen hinzufügen

			this.dbname := dbname

		; Dateipfad der Suchparameterdatei und Suchparameter laden
			this.SParamsPath  	:= this.FileDir "\QuickSearch-Suche_" this.dbname ".json"
			this.SParams        	:= this.SParams_Load()

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
						If RegExMatch(label, "i)^(PATNR|NR)$")
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

		; gelöschte Datensätze anzeigen
			If !this.qs[this.dbname].haskey(ShowRemovedSets)
				this.qs[this.dbname].ShowRemovedSets := false

		; Default labels anlegen für Datenbanksuche anlegen
		; Default Parameter anlegen
			this.SParamsDefaults := ""
			For idx, label in this.Labels()
				this.SParamsDefaults .= label "=`n"
			this.SParamsDefaults := RTrim(this.SParamsDefaults, "`n")

		return 1
		}

		DataType(label)                                  	{	; automatische Spalteneinstellungen für die Windows Listview anhand der dBase Feldtypen
			static DataType 	:= {"N":"Integer", "D":"Integer", "C":"Text", "L":"Text", "M":"Text"}
		return DataType[this.qs[this.dbname].header.text[label].type]
		}

		SortType(label)                                  	{

			static SortType := {"PATNR":"Logical"}

		}

	; ------------------------------------------------------------------------------------
	; Teil-Objekte zurückgeben                                                               	;{
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
	;}

	; ------------------------------------------------------------------------------------
	; Tabelleninformationen erhalten/ändern, virtuelle Tabelle ändern       ;{
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

		HeaderSetVisible(label, VHDRPos)        	{	; virtuelle Spalte ein- oder ausblenden

			this.qs[this.dbname].header.text[label].visible := !VHDRPos ? false : true
			this.HeaderSetVirtualPos(label, VHDRPos)

			hdrPos := this.HeaderPos(label)                                                           	;
			Loop, % this.HeaderCount() - hdrPos {	    		                                	; alle rechts davon sichtbaren virtuellen Tabellenfelder um eine
																														; Position nach rechts oder links verschieben
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
					If RegExMatch(label, "i)^(PATNR|NR)$") {
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
	;}

	; ------------------------------------------------------------------------------------
	; virtuelle Tabelle                                                                              	;{
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
	;}

	; ------------------------------------------------------------------------------------
	; Tabelleneinstellungen laden oder speichern                                     	;{
		LoadProperties()                                  	{	; Objekt laden
			If !this.FilePath || !FileExist(this.FilePath) {
				throw A_ThisFunc ": Einstellungsdatei <" this.FileName "> ist nicht im `n"
						. "Ordner: " this.FileDir " vorhanden."
				return
			}
		return JSONData.Load(this.FilePath, "", "UTF-8")
		}

		SaveProperties(saveColWidth:=true)    	{	; Objekt speichern

			global QSDBRView

			; Dateipfad anlegen
				If !this.FilePath
					If !FilePathCreate(this.FileDir) {
						throw A_ThisFunc ": Speicherpfad für das Sichern der Einstellungen konnte nicht angelegt werden.`n"
								. "Ordner: " this.FileDir
						return
					}

			; Spaltenbreite speichern
				If saveColWidth && this.hLV
					this.LVColumnSizes()

			; Checkbox - gelöschte Datensätze anzeigen
				Gui, QS: Submit, NoHide
				this.qs[this.dbname].ShowRemovedSets := QSDBRView

			; Objekt speichern
				JSONData.Save(this.FilePath, this.qs, true,, 1, "UTF-8")

		return FileExist(this.FilePath) ? 1 : 0
		}

		ShowRemovedSets(show:="")                	{	; Einstellung letzte Datensätze anzeigen
			If (StrLen(show) > 0)
				this.qs[this.dbname].ShowRemovedSets := show
			else
				return this.qs[this.dbname].ShowRemovedSets
		}

		ShowPatNames(show:="")                  	{	; anstatt einer Nummer (PATNR), Patientennamen anzeigen

			global QSDBR, QS

		  ; gibt die Einstellung zurück
			If (StrLen(show) = 0)
				return this.qs[this.dbname].ShowPatNames

		  ; neue Einstellung speichern
			this.qs[this.dbname].ShowPatNames := show

		  ; welche Spaltennummer hat PATNR derzeit
			PatNRCol := this.PatIDColumn()

			Gui, QS: Default
			Gui, QS: ListView, QSDBR

			Loop % LV_GetCount() {

				LV_GetText(PatID, A_Index, PatNRCol)
				if show {
					LV_Modify(A_Index, 	"Col" PatNRCol, PatDB[PatID].name
												.	", " 	PatDB[PatID].VORNAME
												;. 	" *" 	ConvertDBASEDate(PatDB[PatID].GEBURT)
												.	" [" PatID "]")
				} else {
					RegExMatch(PatID, "\[(?<NR>\d+)\]", Pat)
					LV_Modify(A_Index, "Col" PatNRCol, PatNR)
				}

			}

			If Show
				LV_ModifyCol(PatNRCol, "Left Auto")
			else
				LV_ModifyCol(PatNRCol, "Right AutoHDR")

			SciTEOutput(PatNRCol)
		}
	;}

	; ------------------------------------------------------------------------------------
	; Suchparameter-Handler der aktuellen Tabelle                              	;{
		SParams_Load()                                   	{	; gesicherte Suchparameter laden

			; Funktion wird bei jedem WorkDB() Aufruf ausgeführt. Die Suchparamenter müssen deshalb nicht extra geladen werden.

			If !InStr(FileExist(this.FileDir), "D") {
				throw A_ThisFunc ": Speicherordner <" this.FileDir "> ist nicht vorhanden."
				return
			}
			If FileExist(this.SParamsPath)
				SParams := JSONData.Load(this.SParamsPath, "", "UTF-8")
			else
				SParams := Object()

		return SParams
		}

		SParams_Save(delObject:=false)	       	{

			If !InStr(FileExist(this.FileDir), "D")  {
				throw A_ThisFunc ": Speicherordner <" this.FileDir "> ist nicht vorhanden."
				return
			}

			If (this.SParams.Count() > 0)
				JSONData.Save(this.SParamsPath, this.SParams, true,, 1, "UTF-8")

			If delObject
				this.SParams := "", this.SParamsPath := ""

		}

		SParams_GetLabels()                           	{	; Textbezeichnungen für Combobox erstellen

			If (this.SParams.Count() > 0)
				For label, params in this.SParams {
					If (label = "lastLabel")
						continue
					labels .= label "|"
				}
			else
				labels := this.SLabels

		return RTrim(labels, "|")
		}

		SParams_GetParams(TextLabel)           	{	; Parameter für Ausgabe in Edit-Control umwandeln

			If TextLabel && this.SParams.haskey(TextLabel)
				return RegExReplace(this.SParams[TextLabel], "[#]{2}", "`n")
			else
				return this.SParamsDefaults

		}

		SParams_CBSet(TextLabel)              		{	; Combobox befüllen, Suchparameter anzeigen

			global QS, QSBDS, QSDBSParam

			Gui, QS: Default
			Gui, QS: Submit, NoHide

			If (QSDBSParam <> TextLabel) {
				GuiControl, QS:                     	, QSDBSParam	, % "|" this.SParams_GetLabels()              	; Combobox befüllen ("|" am Anfang ersetzt die Liste)
				GuiControl, QS:  ChooseString	, QSDBSParam	, % TextLabel                                            	; zuletzt angezeigte Suche vorwählen
			}

			GuiControl	, QS:                       	, QSDBS        	, % this.SParams_GetParams(TextLabel)  	; und anzeigen

		}
	;}

	; +++++++++ PRIVATE FUNKTIONEN ++++++++++++
		CheckLabel(label, FuncName)            	{
			return
			If (StrLen(label) = 0)
				return 0
			If !IsObject(this.qs[this.dbname].header.text[label]) {
				throw FuncName ": Die Spaltenüberschrift <" label "> gibt es nicht im Tabellenobjekt [" this.dbname "]"
				return 0
			}
		return 1
		}

		CheckBool(bool, FuncName, param)  	{
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


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Listview Funktionen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
LV_DeleteCols(GuiName, LVName, ColRange:="All" )                                         	{

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

LV_ColumnGetSize(hLV, colNr:=0)                                                                     	{	; Pixelbreite des Listviewsteuerelementes erhalten

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

LVM_SetColWidth(hLV, cLV, wLV=-1)                                                                  	{
	; hLV = ListView handle.
	; cLV = 1 based column index to get width of.
	; wLV = New width of the column in pixels. Defaults to -1. The following values are supported in report-view mode:
	;	"-1" - Automatically sizes the column.
	;	"-2" - Automatically sizes the column to fit the header text. If you use this value with the last column, its width is set to fill the remaining width of the list-view control.
	Return DllCall("SendMessage", "uint", hLV, "uint", 4126, "uint", cLV-1, "int", wLV) ; LVM_SETCOLUMNWIDTH
}
;}


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; andere Funktionen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
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

Create_QuickSearch_ICO(NewHandle := False) {
	Static hBitmap := Create_QuickSearch_ICO()
	If (NewHandle)
	   hBitmap := 0
	If (hBitmap)
	   Return hBitmap
	VarSetCapacity(B64, 5096 << !!A_IsUnicode)
	B64 := "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"
	If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
	   Return False
	VarSetCapacity(Dec, DecLen, 0)
	If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
	   Return False
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

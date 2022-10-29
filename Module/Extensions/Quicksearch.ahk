; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                 	Addendum Quicksearch  - - geschrieben zur visuellen Analyse und Untersuchung von dBase Datenbankdateien
;
;
;
;      Funktionen:           	- Anzeige von Datenbankinhalten mit (RegEx) Suchfunktion
;                               	- Auflistung alle Datenbankdateien im Albiswin\db Ordner, dazu Dateigröße, Anzahl der Datensätze und die Zeit des letzten Zugriffes
;                               	- Anzeige von Informationen über die gewählte DBase Datei wie Feldnamen (Tabellenkopf/Spaltennamen), Feldtyp und maximale Zeichenlänge
;
;                                	- einstellbare Ausgabetabelle
;                                	- Tabellenspalten lassen sich aus- und wieder einblenden
;                               	- individuelle Tabellenlayout für jede einzelne Datenbank
;									- das veränderte Tabellenlayout (inkl. Spaltenbreite) wird nach Skriptstart oder nach einem Tabellenwechsel wiederhergestellt
;									- Zeit- und Datumsinformationen werden automatisch ins Deutsche Zeitformat konvertiert
;                                  	- Patientennamen werden bei einem einfachen Klick auf einen Datensatz angezeigt
;									- Karteikarten werden nach einem Doppelklick in Albis geöffnet
;									- verknüpfte Daten werden nach einem einfachen Klick angezeigt (unterstützt derzeit nur Daten aus BEFTEXTE.dbf)
;									- die Suchergebnisse und verlinkte Daten (aus anderen Datenbank) können als .csv-Datei exportiert werden
;									- eingebene Suchparameter lassen sich mit einer freien Bezeichnung/Beschreibung sichern und auch wiederherstellen
;									- UNIX Datums- und Zeitformate werden automatisch ins deutscher Format umgewandelt
;									- anstatt eine Zahl zur Identifikation eines Patientendatensatzes kann zusätzlich der Patientenname angezeigt werden
;									- gelöschte Datensätze lassen sich optional ausblenden
;
;		Hinweis:				beta Version, Nutzung auf eigenes Risiko!
;
;		Abhängigkeiten:	siehe includes
;
;	                    			Addendum für Albis on Windows
;                        			by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;									begin: 02.04.2021,	last modification: 31.05.2022
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
/*  Notizen

	PATEXTRA - POS:  	EMail 				                    			= 93, 7
								                                        				= 81
								keinen Medplan?                 				= 82
								Chroniker?                        				= 83
															                    		= 84
								Anmerkungen/weitere Info.				= 0
															                    		= 95
								Geb.Datum   	                    			= 99
								Ausnahmeindikation                     		= 97
								Geb.Datum (Rechnungsempfänger) 	= 98

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

	global Props, PatDB, adm, q:= Chr(0x22)

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
	adm.compname	:= StrReplace(A_ComputerName, "-")                                                      	; der Name des Computer auf dem das Skript läuft

  ; Nutzereinstellungen werden beim Erstellen des Objektes automatisch geladen
	Props := new Quicksearch_Properties(adm.DBPath "\QuickSearch\resources\Quicksearch_UserProperties.json")

  ; Patientendaten laden
	filter   	:= ["NR", "PRIVAT", "GESCHL", "NAME", "VORNAME", "GEBURT",  "PLZ", "ORT"
					, "STRASSE", "HAUSNUMMER", "TELEFON", "TELEFON2", "TELEFAX", "ARBEIT", "LAST_BEH", "MORTAL"]
	PatDB 	:= ReadPatientDBF(adm.AlbisDBPath, filter, "allData")
	For PatID, Pat in PatDB
		Pat.AGE := Age(Pat.GEBURT, A_YYYY A_MM A_DD)
;}

	QuickSearch_Gui()

return

^!#::Reload
#IfWinActive, Quicksearch ahk_class AutoHotkeyGUI
^Esc::ExitApp
#IfWinActive

Quicksearch(dbname, searchstring)                                                                    	{	; Datenbanksuche + Gui Anzeige

		global dbfiles, dbfdata, InCellEdit, Running

	; Gui-Daten
		Gui, QS: Default
		Gui, QS: Submit, NoHide
		GuiControl, QS:, QST4, % ""

	; InCell Edit stoppen
		InCellEdit.OnMessage(false)

	; RegEx Suchstrings erstellen
		pattern := Object(), RegExSearch := false
		For idx, line in StrSplit(searchstring, "`n") {
			RegExMatch(Trim(line), "^(?<field>.*)?\s*\=\s*(?<RegEx>rx:)*(?<rx>.*)$", m)
			If mrx
				pattern[mfield] := mRegEx . mrx
			if mRegEx
				RegExSearch := true
		}

	; Datenbanksuche starten
		Running 	:= "DBSearch"
		dbf       	:= new DBASE(adm.AlbisDBPath "\" dbname ".dbf", 2)
		filepos   	:= dbf.OpenDBF()
		If !RegExSearch {
			dbfdata	:= dbf.SearchFast(pattern, "all fields", 0	, {"SaveMatchingSets"    	: false          	; Treffer in Datei speichern
																					,  	"LogicalComparison" 	: "and"         	; Treffer nur wenn das gesamte Muster übereinstimmt
																					,	"maxMatches"              	: 12000       	; maximale Übereinstimmungen die gefunden werden sollen
																					,  	"MSetsPath"               	: A_Temp "\QuickSearch_" dbname ".txt"
																					,	"ReturnDeleted"          	: QSDBRView }
																					, 	"QuickSearch_Progress")
		}
		else {
		  ; Search(pattern, startrecord=0, callbackFunc="", opt="")
			dbfdata	:= dbf.Search(pattern, 0, "QuickSearch_Progress"	, {  "SaveMatchingSets"    	: false          	; Treffer in Datei speichern
																										, "LogicalComparison" 	: "and"         	; "and" = Treffer nur wenn das gesamte Muster übereinstimmt
																										, "maxMatches"             	: 12000       	; maximale Items um RAM-Fehler zu vermeiden
																										, "MSetsPath"               	: A_Temp "\QuickSearch_" dbname ".txt"
																										, "debug"                      	: 0x2          	; es können mehrere Debugvarianten gleichzeitig genutzt werden
																										, "OutputDebug"          	: "QuickSearch_OutputDebug"
																										, "ReturnDeleted"          	: QSDBRView })
		}
		filepos   	:= dbf.CloseDBF()
		dbf          	:= ""

	; Balken auf 100%
		GuiControl, QS:, QSPGS	, 100

	; Suchergebnis im Props Objekt ablegen (kann auf Festplatte gespeichert werden)
		Gui, QS: ListView, QSDBR
		LV_Delete()

	; bei einigen Datenbanken ist es besser immer die gelöschten Datensätze anzuzeigen
		ShowRemovedSets	:= Props.ShowRemovedSets(RegExMatch(Props.dbname, "(WZIMMER|LABREAD)") ? 1 : "")
		ShowPatNames 	:= Props.ShowPatNames()
		AutoDateFormat 	:= Props.AutoDateFormat()
		PATNRColumn   	:= Props.PATNRColumn()                                               	; Spaltennummer der PATNR Spalte

	; Spalten zusammenstellen
		equalIDs           	:= Object()
		cutLen               	:= [10, 9, 8, 6, 6, 6]

		For rowNr, set in dbfdata {

		  ; weiter wenn gelöschte Datensätze nicht angezeigt werden sollen
			If (!ShowRemovedSets && set.removed)
				continue

		  ; Spalten zusammenstellen
			cols := Object()
			noRowData := false
			For flabel, val in set {
				If Props.isVisible(fLabel) {

					If (dbname = "BEFUND")
						If (RegExMatch(flabel, "i)(INHALT|KUERZEL)") && StrLen(val) = 0) {      	; Feldnamen (Inhalt,Kuerzel) ohne Wert nicht anzeigen
							noRowData := true
							break
						}

				  ; Formatieren von Datumzahlen und Zeitstrings
					If AutoDateFormat && RegExMatch(flabel, "i)^(" Props.DateLabels "|DATETIME|" Props.TimeLabels ")$") {
						if RegExMatch(flabel, "i)DATETIME")
							FormatTime, val, % val, dd.MM.yyyy HH:mm
						else if RegExMatch(flabel, "(" Props.DateLabels ")")
							val := ConvertDBASEDate(val)
						else if RegExMatch(flabel, "(" Props.TimeLabels ")") {
							FormatTime, time, % today SubStr("0" val, -3) , HH:mm
							val := time = "00:00" ? val : time
						}
						else {
							If (dbname = "CAVE")
								val := RegExMatch(val, "\d{8}") ? SubStr(val, 1, 2) "." SubStr(val, 3, 2) "." SubStr(val, 5, 4) : val
							else if (dbname = "HMVSTAT")
								val := val
							else If  (dbname <> "CAVE")
								val := RegExMatch(val, "\d{8}") ? ConvertDBASEDate(val) : val
						}
					}

				  ; Daten umwandeln
					If RegExMatch(flabel, "i)VERSART")                                                                           	; Versicherungsart lesbar darstellen
						val := Props.VERSART[val]
					else if RegExMatch(flabel, "i)PATNR") {                                                                       	; zählt und vergleicht eingesammelte PATNR
						equalIDs[val] := !equalIDs.haskey(val) ? 1 : equalIDs[val] + 1
						val := ShowPatNames 	? SubStr("         [" val "]" , -1*(cutLen[StrLen(val)]))
															. "  " PatDB[val].NAME ", " PatDB[val].VORNAME ", *" ConvertDBASEDate(PatDB[val].GEBURT)
															. " (" PatDB[val].AGE ")" : val
					}

					cols[Props.HeaderVirtualPos(flabel)] := val                                                                	; ausgelesene Werte einer Spalte zuordnen

				}

			}

		  ; keine anzeigbaren Daten
			If noRowData
				continue

		  ; Zeile hinzufügen
			If ShowRemovedSets
				cols.1 := (set.removed ? "*" : " ") cols.1
			LV_Add("", cols*)

		}

	; Breite der PatientNR Spalte automatisch anpassen
		If PATNRColumn && ShowPatNames
			LV_ModifyCol(PATNRColumn, "AutoHDR")

	; die Anzahl verschiedener Patienten einblenden
		GuiControl, QS:, QST4, % equalIDs.Count()

	; Suchparameter eingaben wieder ermöglichen
		InCellEdit.OnMessage(true)

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
		global 	QSRXSearch
		global	QSSearch			, QSSaveRes			, QSSaveAs			, QSSaveSearch	, QSQuit			, QSReload                	; Buttons
		global 	QSBlockNR		, QSBlockSize		, QSLBlock			, QSNBlock
		global 	QSEXTRANFO 	, QSEXTRAS	    	, QSEXTRA2      	, QSXTRT          	, QSDBSSave   	, QSDBSParam
		global	QSDBName  	, QSDBSNew      	, QSDBRView     	, QSDBRPatView	, QSDBRDFormat
		global 	BlockNr			, BlockSize		   	, lastBlockSize		, lastBlockNr			, dbfRecords  	, lastSavePath
		global 	QSmrgX			, QSmrgY				, QSLvX, QSLvY, QSLvW, QSLvH
		global	InCellEdit
		global 	ReloadGui := false

	  ; ANDERES
		global dbfiles, DBColW, dbname

		static	QSBlockNRBez, QSEXTRASx, QSReloadw
		static	DBField
		static	gQSDBFields		:= ""
		static	EventInfo, empty := "- - - - - - - - - - -"
		static	FSize              	:= (A_ScreenHeight > 1080 ? 9 : 8)
		static	gcolor            	:= "5f5f50"
		static	tcolor             	:= "DDDDBB"
		static	DBSearchW   	:= 300
	  global DBFilesW   	  	:= 310
		static	GFR              	:= " gQuickSearch_DBFilter "
		static	BGT              	:= " BackgroundTrans "
		static	ASM              	:= " AltSubmit "
		static	ColorTitles     	:= " cWhite"
		static	ColorData      	:= " cBlack"

		BlockNr := 0
		dbstructsPath := A_ScriptDir "\resources\AlbisDBase_Structs.json"

		; Datenbank Tabelle
		DBColW	:= [80, 65, 80, 75]
		LVWidth1	:= 20                                  ; Rand der Listview
		For idx, colWidth in DBColW
			LVWidth1 += colWidth

		If !FileExist(dbstructsPath) {
			dbfiles 	:= DBASEStructs(adm.AlbisDBPath)
			JSONData.Save(dbstructsPath, dbfiles, true,, 1, "UTF-8")
		} else {
			dbfiles := JSONData.Load(dbstructsPath, "", "UTF-8")
		}

		IniRead, QSMinKB 	, % adm.Ini, % "QuickSearch", % "QSMinKB"
		IniRead, QSMaxKB	, % adm.Ini, % "QuickSearch", % "QSMaxKB"
		IniRead, QSFilter   	, % adm.Ini, % "QuickSearch", % "QSFilter"
		IniRead, BlockSize   	, % adm.Ini, % "QuickSearch", % "BlockSize"    	, % "4000"
		IniRead, WinSize    	, % adm.Ini, % "QuickSearch", % "WinSize"      	, % "xCenter yCenter w1455 h830"
		IniRead, lastdbname 	, % adm.Ini, % "QuickSearch", % "lastDB"
		IniRead, lastS        	, % adm.Ini, % "QuickSearch", % "lastSearch"
		IniRead, ExtrasStatus 	, % adm.Ini, % "QuickSearch", % "Extras"
		IniRead, lastSavePath	, % adm.Ini, % "QuickSearch", % "lastSavePath"  	, % ""

	; für den Fall fehlerhafter oder einer nicht anzeigbaren Position
		RegExMatch(WinSize, "i)x(?<X>\-*\d+)", GClient)
		RegExMatch(WinSize, "i)y(?<Y>\-*\d+)", GClient)
		RegExMatch(WinSize, "i)w(?<W>\-*\d+)", GClient)
		RegExMatch(WinSize, "i)h(?<H>\-*\d+)", GClient)
		GClientX     	:= !GClientX 	||	GClientX < 10  	? 0    	: GClientX
		GClientY     	:= !GClientY 	||	GClientY < 10  	? 0    	: GClientY
		GClientW   	:= !GClientW	||	GClientW <1455	? 1455	: GClientW - 5
		GClientH   	:= !GClientH	||	GClientH <830 	? 830 	: GClientH

		If !IsInsideVisibleArea(GClientX, GClientY, GClientW, GClientH, CoordInjury)
			GClientX := GClientY := "Center"

		winSize       	:= "x" GClientX " y" GClientY " w" GClientW " h" GClientH

	; Korrektur anderer Daten
		BlockSize    	:= !BlockSize	|| !RegExMatch(BlockSize	, "^\d+$") ? 4000 	: BlockSize
		lastdbname 	:= InStr(lastdbname, "ERROR") ? "" : lastdbname
		lastBlockSize	:= BlockSize
		lastBlockNr 	:= BlockNr
		lastS          	:= RegExReplace(lastS, "[\|]", "`n")

	;}

		Gui, QS: new    	, HWNDhQS -DPIScale +Resize minSize1600x900
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
		Gui, QS: Add	, Text        	, % "x+20  ym+5 w200    " BGT "                                                             vQSDBName" 	, % (lastdbname ? lastdbname ".dbf" : "- - -")


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

		lp := GetWindowSpot(QShDB)
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
		LV_ModifyCol(1, "Sort")

		; - zur Zeile der aktuellen Datenbank scrollen
		If lastdbname
			Loop % LV_GetCount() {
				LV_GetText(rowDB, A_Index, 1)
				If (rowDB = lastdbname) {
					LV_Modify(A_Index, "Select")
					LV_Modify(A_Index+7 > LV_GetCount() ? LV_GetCount() : A_Index+7, "Vis")
					break
				}
			}

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
		Gui, QS: Add	, Checkbox  	, % "x" GCol3X " y" cpY " w" DBSearchW " " BGT " vQSRXSearch gQS_Handler"           	, % "RegEx-Suche"

		GuiControlGet, cp, QS:Pos, QSDB
		GuiControlGet, dp, QS:Pos, QSDBFL
		opt := " -E0x200 vQSDBS HWNDQShDBS gQS_Handler"
		Gui, QS: Font	, % "s" FSize " q5 Normal" ColorData
		Gui, QS: Add	, Edit         	, % "x" dpX+dpW+5 " y" dpY 	" w" DBSearchW " h" dpH . opt                                    	, % ""

		Gui, QS: Font	, % "s" FSize " q5 Normal" ColorData
		Gui, QS: Add	, ComboBox 	, % "y+2 w" DBSearchW+1 " r8 vQSDBSParam 	gQS_Handler"                                  	, % ""

		GuiControlGet, cp, QS:Pos, QSDBSParam
		Gui, QS: Font	, % "s" FSize-1 " q5 Normal" ColorData
		Gui, QS: Add	, Button        	, % "x+5 y" cpY+1 " h" cpH-2 " vQSDBSSave    gQS_Handler"                                  	, % "Speichern"
		Gui, QS: Add	, Button        	, % "x+3 h" cpH-2 " vQSDBSNew    gQS_Handler"                                                     	, % "neue Suche"

		GuiControlGet, cp, QS:Pos, QSDBSSave
		GuiControlGet, dp, QS:Pos, QSDBS

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
		Gui, QS: Add	, Edit         	, % "y+1 w400 h" dpH-epH-cpH-1 " -E0x200 vQSEXTRAS"
		Gui, QS: Add	, Edit         	, % "y+1 w400 h" cpH " -E0x200 hwndhwnd vQSEXTRA2"

	  ; Edit Banner setzen
		banner := "kopierbare Debug-Ausgaben"
		SendMessage, 0x1501, 1, &banner,, % "ahk_id " hwnd
		Edit_SetMargin(hwnd, 1, 1, 1, 1)

		GuiControl, % "QS:" (ExtrasStatus ? "Enable" : "Disable"), % "QSEXTRANFO"
		GuiControl, % "QS:" (ExtrasStatus ? "Enable" : "Disable"), % "QSEXTRAS"
	;}

	;}

	;-: ZEILE 2: BEFEHLE                                    	;{

		GuiControlGet, cp, QS: Pos , QSEXTRA2
		GuiControlGet, dp, QS: Pos , QSDBSSave
		GuiW          	:= cpX + cpW-5
		QSEXTRASx 	:= cpX
		row1H       	:= cpY + cpH
		row2Y         	:= dpY + dpH

		;-: FELD 1: INFO'S                                  	;{
		Gui, QS: Font	, % "s" FSize " " ColorTitles
		Gui, QS: Add	, GroupBox	, % "xm     	y" row1H-6  " w100 h50" 	 "  vQSTGB1"

		Gui, QS: Add	, Text        	, % "xm+3  	y" (row1H+3)       BGT " vQST1a"                              	, % "Position:"
		Gui, QS: Add	, Text        	, % "x+2                            	 " BGT " vQST1"                                	, % empty "     "

		Gui, QS: Add	, Text        	, % "xm+3  	y+1                	 " BGT " vQST2a"                              	, % "Datensätze:"
		Gui, QS: Add	, Text        	, % "x+2                           	 " BGT " vQST2"                                	, % empty "     "

		GuiControlGet, cp, QS: Pos , QST2
		GuiControlGet, dp, QS: Pos , QST2a
		col2X := dpW+cpW+22
		InfoH := dpH+cpH+12

		GuiControl, QS: Move, QST1a  	, % "w" 	dpW
		GuiControl, QS: Move, QST1    	, % "x" 	dpX+dpW+2
		GuiControl, QS: Move, QSTGB1	, % "w" 	dpW+cpW+16 " h" InfoH

		; - - - - -

		Gui, QS: Add	, GroupBox	, % "x" col2X      " y" row1H-6  " w100 h" InfoH "  vQSTGB2"

		Gui, QS: Add	, Text        	, % "x" col2X+3 " y" (row1H+3)	   BGT " vQST3a"                   	, % "Treffer:"
		Gui, QS: Add	, Text        	, % "x+2                               	 " BGT " vQST3"                        	, % empty " "

		Gui, QS: Add	, Text        	, % "x" col2X+3 " y+1             	 " BGT " vQST4a"                      	, % "Patienten:"
		Gui, QS: Add	, Text        	, % "x+2                               	 " BGT " vQST4"                     	, % empty " "

		GuiControlGet, cp, QS: Pos , QST2
		GuiControlGet, dp, QS: Pos , QST3a
		GuiControlGet, ep, QS: Pos , QST4a

		GuiControl, QS: Move, % (dpW >= epW ? "QST4a" 	: "QST3a")	, % "w" (dpW >= epW ? dpW : epW)
		GuiControl, QS: Move, % (dpW >= epW ? "QST4" 	: "QST3")	, % "x"  (dpW >= epW ? dpX+dpW+2 : epX+epW+2)	;dpX+dpW+2
		GuiControl, QS: Move, QSTGB2	, % "w" 	LVWidth1-dpX+mrgX+3
		col2X := dpW + cpW + 20

		; - - - - -

		Gui, QS: Add	, GroupBox	, % "x" DBFieldX     	" y" row1H-6  " w" DBFieldW " h" InfoH "  vQSTGB3"
        Gui, QS: Add	, Text        	, % "x" DBFieldX+3  	" y" row1H+3        		BGT " "                 	, % "ausgewählter Patient:"
		Gui, QS: Add	, Text        	, % "x" DBFieldX+3 	" y+1 w" DBFieldW-6		BGT " vQST5"       	, % ""


		;}

		;-: FELD 2: PROGRESS                              	;{
		Gui, QS: Add	, Progress     	, % "xm  	y+3 w" GuiW " vQSPGS"			                                		, % 0
		GuiControlGet, cp, QS: Pos, QSPGS
		Gui, QS: Font	, % "s" FSize-1 " " ColorTitles
		Gui, QS: Add	, Text        	, % "x" cpX " y" cpY " w" cpW " h" cpH-2 "  " BGT "   Center vQSPGST", % " - - - - - - - - - - - - - - - "
		;}

		;-: FELD 3: BEFEHLE                               	;{
		BtnH := cpH-8
		Gui, QS: Font	, % "s" FSize " " ColorTitles
		Gui, QS: Add	, Button        	, % "      	y+4                          	vQSSearch         	gQS_Handler", % "Suchen"
		Gui, QS: Add	, Button        	, % "x+15		                            	vQSSaveRes       	gQS_Handler", % "Suche speichern als"

		Gui, QS: Font	, % "s" FSize+2 " " ColorData
		GuiControlGet	, cp, QS: Pos, QSSaveRes
		Gui, QS: Add	, DDL         	, % "x+2 h" cpH " w115 r5            	vQSSaveAs       	gQS_Handler", % ".csv (;)|.csv (Tab)|.csv (,)"
		GuiControl    	, QS: Choose, QSSaveAs, 1
		GuiControlGet	, sp, QS: Pos	, QSSaveAs
		GuiControl    	, QS: Move	, QSSearch	, % "h" spH
		GuiControl    	, QS: Move	, QSSaveRes	, % "h" spH

		Gui, QS: Font	, % "s" FSize+1 " " ColorData
		Gui, QS: Add	, Button        	, % "x" DBFieldX " y" cpY " h" spH " 	vQSLBlock           	gQS_Handler", % "⏪ vorheriger Block"

		Gui, QS: Font	, % "s" FSize-2 " " ColorTitles
		GuiControlGet	, cp, QS: Pos, QSLBlock
		cx := cpX+cpW+5
		Gui, QS: Add	, Text         	, % "x" cx " y" cpY-2 " " BGT "     		vQSBlockNrBez"                       	, % "Block Nr|Größe"

		Gui, QS: Font	, % "s" FSize-1 " " ColorData
		opt := " Number HWNDhEdit "

		Gui, QS: Add	, Edit         	, % "y+-1 w" FSize*4 " " opt "      	vQSBlockNR       	gQS_Handler", % BlockNr
		Edit_SetMargin(hEdit, 2, 0, 2, 0)
		GuiControl     	, QS: Move	, QSBlockNr, % "h19"

		Gui, QS: Add	, Edit         	, % "x+0 w" FSize*6 " " opt "	    	vQSBlockSize      	gQS_Handler", % BlockSize
		Edit_SetMargin(hEdit, 2, 0, 2, 0)
		GuiControl    	, QS: Move	, QSBlockSize, % "h19"

		Gui, QS: Font	, % "s" FSize+1
		Gui, QS: Add	, Button        	, % "x+5 y" cpY " h" spH "             	vQSNBlock        	gQS_Handler", % "nächster Block ⏩"

		Gui, QS: Font	, % "s" FSize-1 " " ColorTitles
		Gui, QS: Add	, Checkbox  	, % "x+10  y" cpY   . 	BGT "          	vQSDBRView     	gQS_Handler", % "gelöschte Datensätze anzeigen"
		Gui, QS: Add	, Checkbox  	, % "y+2                              	      	vQSDBRPatView   	gQS_Handler", % "PATNR mit Namen ersetzen"
		GuiControlGet	, cp, QS: Pos, QSDBRView
		GuiControlGet	, dp, QS: Pos, QSDBRPatView
		cx := cpX + (dpW > cpW ? dpW : cpW) + 10
		Gui, QS: Add	, Checkbox  	, % "x" cx " y" cpY " w" cpW . BGT " vQSDBRDFormat  	gQS_Handler", % "Datumspalten lesbar machen"

		Gui, QS: Font	, % "s" FSize+1 " " ColorData
		Gui, QS: Add	, Button        	, % "x+250 y" cpY " h" spH " 	      	vQSReload          	gQS_Handler", % "Reload"

	  ; Höhe der Buttons anpassen

		;}

	;}

	;-: ZEILE 3: AUSGABE                                	;{
		Gui, QS: Show, %  WinSize " Hide", Quicksearch

		GuiControlGet, cp, QS: Pos, QSReload
		QSReloadw	:= cpW
		LVOptions3 	:= "vQSDBR gQS_Handler HWNDQShDBR -E0x200 -LV0x10 AltSubmit -Multi Grid"
		wqs  	    	:= GetWindowSpot(hQS)
		LvH           	:= wqs.CH - cpY - cpH

		Gui, QS: Font	, % "s" FSize " Normal" ColorData, Futura Bk Bt
		Gui, QS: Add	, ListView    	, % "xm y" cpY+cpH+7 " w" GuiW " h" LvH " "  LVOptions3	, % "- -"
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

		wqs  	:= GetWindowSpot(hQS)
		wqs.X 	:= wqs.X	< -10 ? 0 : wqs.X
		wqs.Y 	:= wqs.Y	< -10 ? 0 : wqs.Y

	  ; dies ruft GuSize auf
		WinMove, % "ahk_id " hQS,,,, % wqs.W+1	, % wqs.H+1
		WinMove, % "ahk_id " hQS,, % wqs.X, % wqs.Y, % wqs.W, % wqs.H
	;}

	; lädt die zuletzt benutzte Datenbank
		If lastdbname {
			BlockInput, On
			QuickSearch_ListFields(dbname := lastdbname)
			BlockInput, Off
		}

	; Hotkeys ;{
		fn_QSHK := Func("QuickSearch_Hotkeys")
		Hotkey, IfWinActive, Quicksearch ahk_class AutoHotkeyGUI
		Hotkey, Left               	, 	% fn_QSHK
		Hotkey, Right             	, 	% fn_QSHK
		Hotkey, Enter             	, 	% fn_QSHK
		Hotkey, NumpadEnter 	, 	% fn_QSHK
		Hotkey, TAB              	, 	% fn_QSHK
		Hotkey, Escape             	, 	% fn_QSHK
		Hotkey, IfWinActive
	;}

return

QS_Handler: 	;{

	Critical

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

		Case "QSDBR":                             	;{   	Datentabelle


			;SciTEOutput("A_GuiControl: " A_GuiControl ", A_GuiEvent: " A_GuiEvent ", A_Eventinfo: " A_EventInfo ", PATNRCol: " Pcol := Props.PATNRColumn())

			       If (A_GuiEvent = "ColClick"                  	)	{
				colHDR := A_EventInfo
				MouseGetPos, mx, my
				ToolTip, % colHDR, % mx, % my-10, 6
				SetTimer, QSColTipOff, -3000
			}
		 ; Karteikartenaufruf
			else if (A_GuiEvent ~= "i)^(DoubleClick|A)$") 	{

				If (!WinExist("ahk_class OptoAppClass") || A_EventInfo=0) {
					return
				}

				If (Pcol := Props.PATNRColumn()) {
					Gui, QS: ListView, QSDBR
					LV_GetText(PatCell, A_EventInfo, Pcol)
					RegExMatch(PatCell, "\d+", PATNR)
					GuiControl, QS:, QST5, % "[" PATNR "] " (PatName := PatDB[PATNR].NAME ", " PatDB[PATNR].VORNAME
														. 	" *" ConvertDBASEDate(PatDB[PATNR].GEBURT))
					AlbisAkteOeffnen(PatName, PATNR)
				}

			}
		; Patientenname und bei Bedarf den Inhalt der TEXTDB anzeigen
			else if RegExMatch(A_GuiEvent, "(Normal)"	)  	{	; |I
				tblRow := A_EventInfo
				If (!tblRow || Last_tblRow = tblRow)
					return
				Last_tblRow := tblRow

			  ; PatNR finden
				If (Pcol := Props.PATNRColumn()) {
					Gui, QS: ListView, QSDBR
					LV_GetText(PatCell, tblRow, Pcol)
					RegExMatch(PatCell, "\d+", PATNR)
					GuiControl, QS:, QST5, % "[" PATNR "] " (PatName := PatDB[PATNR].NAME ", " PatDB[PATNR].VORNAME
														. 	" *" ConvertDBASEDate(PatDB[PATNR].GEBURT))
				}

			 ; verlinkte Daten aus anderer Datenbank anzeigen (wenn Häkchen bei verknüpfte Daten gesetzt ist)
				If QSXTRT
					If Props.LabelExist("TEXTDB") 	;|| Props.LabelExist("TEXT"))
						QuickSearch_Extras(tblRow)
					else if (Props.LabelExist("TXTHINW") || Props.LabelExist("TEXTE"))
						QuickSearch_TxtHinw(tblRow, Props.LabelExist("TXTHINW") ? "TXTHINW" : "TEXTE")

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

		Case "QSDBSNew":                     	;{ 	Suchparameterliste leeren
			If !dbname
				return
			Props.SParams_NewSearch()
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

		Case "QSDBRView":                     	;{ 	entfernte Datensätze anzeigen umschalten
			Props.ShowRemovedSets(QSDBRView)
		;}

		Case "QSDBRPatview":                 	;{ 	PATNR durch Patientennamen ersetzen
			Props.ShowPatNames(QSDBRPatView)
		;}

		Case "QSDBRDFormat":               	;{ 	AutoDatumFormatierung
			Props.AutoDateFormat(QSDBRDFormat)
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
	GuiControl, QS: MoveDraw	, QSEXTRA2   	, % "w" (QSEXTRASw < 300 ? 300 : QSEXTRASw)
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

  ; überprüfe nochmal welches Fenster tatsächlich aktiv ist
	MouseGetPos,,, hMouseOverWin
	mouseWinTitle 	:= WinGetTitle(hMouseOverWin)
	mouseWinClass	:= WinGetClass(hMouseOverWin)
		; oder
	hfocus := GetFocusedControlHwnd()
	hAncestor := GetAncestor(hfocus, 1)
	AncestorWinTitle := WinGetTitle(hAncestor)
	AncestorWinClass := WinGetClass(hAncestor)
	;~ SciTEOutput("[" hfocus ", " hAncestor "] " AncestorWinTitle " ahk_class " AncestorWinClass)

	If (mouseWinTitle <> "QuickSearch" && mouseWinClass <> "AutoHotkeyGUI")
	|| (AncestorWinTitle <> "QuickSearch" && AncestorWinClass <> "AutoHotkeyGUI")
	|| !WinActive("Quicksearch ahk_class AutoHotkeyGUI") {
		SendInput, % "{" A_ThisHotkey "}"
		return
	}

  ; Hotkey-Befehl ausführen
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

		Case "Escape":
			nil := ""

	}

}

QuickSearch_Progress(index, maxIndex, len, matchcount)                                    	{	; Suchfortschritt anzeigen

	global	DBSearchW
	prgPos := Floor((index*100)/maxIndex)
	GuiControl, QS:, QSPGS	, % prgPos
	GuiControl, QS:, QST1 	, % index
	GuiControl, QS:, QST3 	, % matchcount

}

QuickSearch_OutputDebug(debugmsg)                                                              	{	; Debugstrings in eigener Ausgabe
	global QS, QSEXTRA2
	GuiControl, QS:, QSEXTRA2, % debugmsg
}

QuickSearch_Extras(tblRow, TEXTDB:=0, TEXTDate:=0)                                         	{	; verknüpfte Daten aus anderer Datenbank laden

		static DBIndex, dbf

		TEXTDBCol := Props.dbname = "CAVE" ? "TEXT" : "TEXTDB"
		If !Props.isVisible(TEXTDBCol)
			return

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; TEXTDB ist der Link zu den zusätzlichen Texten
	; TEXTDB = LFDNR in BEFTEXTE
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Gui, QS: ListView, QSDBR
		If !TEXTDB {
			LV_GetText(TEXTDB, tblRow, Props.HeaderVirtualPos(TEXTDBCol))
			If !TEXTDB
				return
		} else
			SaveFeature := true

		If !TEXTDate && Props.HeaderVirtualPos("DATUM")
			LV_GetText(TEXTDate, tblRow, Props.HeaderVirtualPos("DATUM"))

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; BEFTEXTE.dbf indizieren oder Index laden
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
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
			IndexFromDay := A_YYYY A_MM A_DD
			IndexFromDay += -30, days
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

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Vorbereitung für das Daten auslesen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If !IsObject(dbf)
			dbf      	:= new DBASE(adm.AlbisDBPath "\BEFTEXTE.dbf", 2)
		records 	:= dbfRecords := dbf.records
		steps      	:= records/100
		filepos   	:= dbf.OpenDBF()

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; filepointer in Richtung des TEXTDB Eintrages versetzen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
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
		dbf._SeekToRecord(IndexStartRecord)                   	; filepointer gesetzt

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; TEXTDB Eintrag finden und Daten zusammen stellen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
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
		filepos := dbf.CloseDBF()
		If !SaveFeature
			QuickSearch_Progress(records, records, StrLen(records), "---")

	; Daten ausgeben
		If !SaveFeature {
			;~ ListVars
			;~ Pause, Toggle
			extrahead	:=  extraLines " Zeile" (extraLines = 1 ? "" : "n")
								. " aus BEFTEXT.dbf für TEXTDB Eintrag: " TEXTDB

			GuiControl, QS:, QSEXTRANFO	, % extrahead
			GuiControl, QS:, QSEXTRAS   	, % extra
		}

return extra
}

QuickSearch_TxtHinw(tblRow, HeaderLabel)                                                          	{ 	; verknüpfte Texthinweise in den Labordatenbanken

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; TXTHINW ist der Link zu den zusätzlichen Texten
	; TXTHINW = LABTEXTE.dbf
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Gui, QS: ListView, QSDBR

		TXTNR := Array()
		LV_GetText(TXT, tblRow, Props.HeaderVirtualPos("TXTHINW"))
		If TXT
			TXTNR.Push([TXT, "TXTHINW"])
		LV_GetText(TXT, tblRow, Props.HeaderVirtualPos("TEXTE"))
		If TXT
			TXTNR.Push([TXT, "TEXTE"])
		LV_GetText(TXT, tblRow, Props.HeaderVirtualPos("ERGTXT"))
		If TXT
			TXTNR.Push([TXT, "ERGTXT"])

		If (TXTNR.Count()=0)
			return

		TEXTE := "\b("
		For each, nr in TXTNR
			TEXTE .= (A_Index>1?"|":"") nr.1
		TEXTE .= ")\b"

		LV_GetText(PATID, tblRow, Props.HeaderVirtualPos("PATNR"))
		If !RegExMatch(PatID, "^[\[\s]*(?<ID>\d+)", Pat)
			return

		dbf     	:= new DBASE(adm.AlbisDBPath "\LABTEXTE.dbf", 2)
		records	:= dbfRecords := dbf.records
		steps    	:= records/100
		filepos 	:= dbf.OpenDBF()
		items 	:= dbf.Search({"PATNR":"rx:" PATID, "ID":"rx:" TEXTE}, 0)
		res     	:= dbf.CloseDBF()
		dbf    	:= ""

		labtxt := ""
		For index, item in items {
			txt :=  RegExReplace(item.TEXT, "([\r\n]+)", "$1 ")
			If (item.pos=0)
				For idx, nr in TXTNR
					If (item.ID = nr.1)
						newline := (index>1 ? "`n`n" : "") nr.2 " (" item.ID "):`n "
			labtxt 	.= newline . txt
			newline	:= ""
		}

		labtxt :=  RegExReplace(labtxt, "([a-zäöüß])([A-ZÄÖÜ])", "$1 $2")
		labtxt :=  RegExReplace(labtxt, "\.([A-ZÄÖÜ])", ". $1")

		GuiControl, QS:, QSEXTRANFO	, % "Texthinweise für PATNR " PATID " Eintrag Nr: " TEXTE " aus LABTEXTE.dbf"
		GuiControl, QS:, QSEXTRAS   	, % labtxt ;"`n" JSON.DUMP(items)

}

QuickSearch_ListFields(dbname, step:=0)                                                             	{ 	; erstellt die Ausgabetabelle

		global	QS, QSDBS, QSBlockNr, QSBlockSize, QSPGS, QSPGST, QSDBSParam, QSDBRView, QSDBRDFormat
		global 	dbfiles, empty, dbfdata
		global 	BlockNr, BlockSize, lastBlockNr, lastBlockSize, dbfRecords
		static 	lastdbname, SRS

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
				GuiControl, QS:, QSDBName, % dbname ".dbf"

			; Datenbanksuche - Parameter/Auswahl
				Props.SParams_CBSet("LastLabels")

			; Checkbox Eintrag 'gelöschte Datensätze anzeigen' wiederherstellen
				ShowRemovedSets := Props.ShowRemovedSets()
				Props.ShowRemovedSets(ShowRemovedSets)

			; Gui Inhalte zurücksetzen
				GuiControl, QS:, QSEXTRASNFO	, % ""                   	; Zusatzinfos leeren
				GuiControl, QS:, QSEXTRAS       	, % ""                  	; Zusatzinfos leeren
				GuiControl, QS:, QST5              	, % ""                  	; ausgewählten Eintrag leeren

			; BlockNr auf Null setzten
				BlockNr := lastBlockNr := 0

			; Ausgabelistview: Spalten erstellen                              	;{

				LV_DeleteCols("QS", "QSDBR")                           	; alle Spalten entfernen

				thisColPos := 1
				For colNr, flabel in Props.Labels()
					If Props.isVisible(flabel) && (flabel<>"removed") {

						dtype 	:= Props.DataType(flabel)
						width 	:= Props.HeaderField(flabel, "width")
						width	:= RegExReplace(width, "i)[a-z]+")
						width  	:= width > 1200 ? 200 : width

						Gui, QS: ListView, QSDBR
						If (thisColPos = 1)
							LV_ModifyCol(thisColPos, (!width ? "100" : width)       	" " dtype, flabel)
						else
							LV_InsertCol(thisColPos	 , (!width ? "AutoHdr" : width)	" " dtype, flabel)

						Props.HeaderSetVirtualPos(flabel, thisColPos)
						thisColPos ++

					}

			;}

			; Datenbankfelderinfo's anzeigen                     	;{
				Gui, QS: ListView, QSDBFL
				LV_Delete()

				For colNr, flabel in Props.Labels() {
					If (flabel = "removed")
						continue
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

	; bei einigen Datenbanken ist es besser immer die gelöschten Datensätze anzuzeigen
		If RegExMatch(Props.dbname, "(WZIMMER|LABREAD)")
			Props.ShowRemovedSets(1)

	; 	Controls anpassen
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

	; Defaults setzen                                                   	;{
		Gui, QS: Default
		Gui, QS: ListView, QSDBR
		GuiControl, QS: -Redraw, QSDBR
		LV_Delete()
	;}

	; Checkbox Einstellung 'PATNR mit Namen ersetzen' ;{
		ShowRemovedSets	:= Props.ShowRemovedSets()
		PATNRColumn   	:= Props.PATNRColumn()                                               	; Spaltennummer der PATNR Spalte
		ShowPatNames	 	:= PATNRColumn ? Props.ShowPatNames() : 0                	; Patientennamen  einblenden, letzte Einstellung
		AutoDateFormat 	:= Props.AutoDateFormat()                                             	; Autoformatierung von Datumszahlen

		GuiControl, QS:, QSDBRDFormat	, % AutoDateFormat                              	; Autoformatierung Datumszahlen
		GuiControl, QS:, QSDBRPatView	, % (ShowPatNames ? 1 : 0)                    	; zusätzlich Namen einblenden
		GuiControl, % "QS: Enable" (PATNRColumn ? "1" : "0"), QSDBRPatView       	; Tabelle ohne/mit PATNR - Steuerelement wird in- oder reaktiviert
		If PATNRColumn
			LV_ModifyCol(PATNRColumn, (ShowPatNames ? "Left": "Right"))                  	; Ausrichtung des Spalteninhaltes
	;}

	; Ausgabe der Daten im Ergebnisfenster             	;{
		equalIDs := {}
		cutLen 	:= [10, 9, 8, 6, 6, 6]
		today	:= A_YYYY A_MM A_DD
		Props.tmpcnt := 0
		For rowNr, set in dbfdata {

			; weiter wenn gelöschte Datensätze nicht angezeigt werden sollen
				If (!ShowRemovedSets && set.removed)
					continue

			; Spalten zusammenstellen
				cols := Object(), noRowData := false
				For flabel, val in set {
					If Props.isVisible(fLabel) {

						If (dbname = "BEFUND")
							If (RegExMatch(flabel, "i)(INHALT|KUERZEL)") && StrLen(val) = 0) {      	; Feldnamen (Inhalt,Kuerzel) ohne Wert nicht anzeigen
								noRowData := true
								break
							}

					  ; - - - - - - - - - - - - - - - - - - - - - - - -
					  ; Formatieren von Datumzahlen
					  ; - - - - - - - - - - - - - - - - - - - - - - - -
						If AutoDateFormat && RegExMatch(flabel, "i)^(" Props.DateLabels "|" Props.TimeLabels ")"  "$") {
							if (flabel = "DATETIME")
								FormatTime, val, % val, dd.MM.yyyy HH:mm
							else if RegExMatch(flabel, "i)^(" Props.TimeLabels ")$") {
								FormatTime, time, % today SubStr("0" val, (flabel = "UHRZEIT" ? -5:-3)) , % (flabel = "UHRZEIT" ? "HH:mm:ss" : "HH:mm")
								val := (time=="00:00:00" || time=="00:00" ) ? val : time
							}
							else if RegExMatch(flabel, "i)^(" Props.DateLabels ")$") {
								If RegExMatch(val, "\d{8}")
									val := dbname = "CAVE" ? SubStr(val, 1, 2) "." SubStr(val, 3, 2) "." SubStr(val, 5, 4) : ConvertDBASEDate(val)
							}
						}

					  ; - - - - - - - - - - - - - - - - - - - - - - - -
					  ; Daten umwandeln
					  ; - - - - - - - - - - - - - - - - - - - - - - - -
						If RegExMatch(flabel, "i)VERSART")                                                            	; Versicherungsart lesbar darstellen
							val := Props.VERSART[val]
						else if (RegExMatch(flabel, "i)PATNR") && val > 0) {                                   	; zählt und vergleicht eingesammelte PATNR
							equalIDs[val] := !equalIDs.haskey(val) ? 1 : equalIDs[val]+1
							If ShowPatNames
								val := SubStr("          [" val "]" , -1*(cutLen[StrLen(val)])) "  "
									.  	 (!PatDB[val].Name && !PatDB[val].Vorname ? "✘gelöschter Patient✘" : PatDB[val].Name ", " PatDB[val].Vorname)
						}

						cols[Props.HeaderVirtualPos(flabel)] := val                                             	; ausgelesene Werte einer Spalte zuordnen

					}
				}

			; keine anzeigbaren Daten
				If noRowData
					continue

			; * für gelöschten Datensatz in der ersten Spalte anzeigen
				If ShowSetRemoved
					cols.1 := (set.removed ? "*" : " ") cols.1

			; Zeile hinzufügen
				LV_Add("", cols*)

		}

		If PATNRColumn && ShowPatNames                                                                 	; Ausrichtung des Spalteninhaltes
			LV_ModifyCol(PATNRColumn, "AutoHDR")

		GuiControl, QS: +Redraw, QSDBR                                                                 	; Datenausgabe-Listview jetzt neu zeichnen
		GuiControl, QS:, QST4, % equalIDs.Count()                                                    	; Anzahl identischer PATNr anzeigen

		;~ If RegExMatch(Props.dbname, "(WZIMMER|LABREAD)")
			;~ Props.ShowRemovedSets(SRS)

		InCellEdit.OnMessage(true)

	;}

}

QuickSearch_DBFields(GChwnd:="", GuiEvent:="", EventInfo:="", ErrLevel:="")    	{	; Eventhandler DBFields

		global dbfiles, dbname, dbfdata, QShDBFL, QShDBR
		static DataType 	:= {"N":"Integer", "D":"Integer", "C":"Text", "L":"Text", "M":"Text"}
		static evTime := 0
		static lastClickTime, ClicksMeasured

		/*  GuiEvent Bechreibung
			A: 	Eine Zeile wurde aktiviert, was standardmäßig geschieht, wenn sie doppelt angeklickt wurde.
					Die Variable A_EventInfo enthält die Nummer der Zeile.
			C: 	Die ListView hat die Mauserfassung freigegeben.
			E: 	Der Benutzer hat begonnen, das erste Feld einer Zeile zu editieren (der Benutzer kann das Feld nur editieren,
					wenn -ReadOnly in den Optionen der ListView vorhanden ist). Die Variable A_EventInfo enthält die Nummer der Zeile.
			F:		Die ListView hat den Tastaturfokus erhalten.
			f 		(kleines F): Die ListView hat den Tastaturfokus verloren.
			I: 		Der Zustand einer Zeile hat sich in irgendeiner Form geändert; zum Beispiel weil sie ausgewählt, abgewählt, abgehakt usw. wurde.
					Wenn der Benutzer eine neue Zeile auswählt,	empfängt die ListView mindestens zwei solcher Benachrichtigungen: eine für das
					Abwählen der vorherigen Zeile und eine für das Auswählen der neuen Zeile.
					[v1.0.44+]:    	Die Variable A_EventInfo enthält die Nummer der Zeile.
					[v1.0.46.10+]: ErrorLevel enthält null oder mehr der folgenden Buchstaben, um mitzuteilen,
					wie das Element geändert wurde: S (ausgewählt) oder s (abgewählt), und/oder F (fokussiert) oder f (defokussiert),
					und/oder C (Häkchen gesetzt) oder c (Häkchen entfernt).
					SF beispielsweise bedeutet, dass die Zeile ausgewählt und fokussiert wurde. Um festzustellen, ob ein bestimmter Buchstabe
					vorhanden ist, verwenden Sie eine Parsende Schleife oder die GroßKleinSensitiv-Option von InStr(); zum Beispiel: InStr(ErrorLevel, "S",
					true). Hinweis: Aus Gründen der Kompatibilität mit zukünftigen Versionen sollte ein Skript nicht davon ausgehen, das "SsFfCc" die
					einzigen möglichen Buchstaben sind. Außerdem können Sie Critical in der ersten Zeile von g-Label einfügen, um sicherzustellen,
					dass alle "I"-Benachrichtigungen empfangen werden (andernfalls könnten einige davon verloren gehen, wenn das Skript nicht mit
					ihnen Schritt halten kann).

			K: 	Der Benutzer hat eine Taste gedrückt, während die ListView den Fokus hat. A_EventInfo enthält den virtuellen Tastencode der Taste
					(eine Zahl zwischen 1 und 255). Dieser Code kann via GetKeyName() in einen Tastennamen oder in ein Zeichen übersetzt werden.
					Zum Beispiel Taste := GetKeyName(Format("vk{:x}", A_EventInfo)). Bei den meisten Tastaturlayouts 	können die Tasten A bis Z via
					Chr(A_EventInfo) in das entsprechende Zeichen übersetzt werden. F2 wird unabhängig von WantF2 erfasst. Enter hingegen wird nicht
					erfasst; um es dennoch zu erfassen, können Sie, wie unten beschrieben, eine Standardschaltfläche nutzen.

			M: 	Auswahlrechteck. Der Benutzer hat damit begonnen, ein Auswahlrechteck über mehrere Zeilen oder Symbole zu ziehen.

			S: 	Der Benutzer hat begonnen, in der ListView zu scrollen.

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

Quicksearch_SaveSearch(dbname, textlabel, searchstring)                                   	{	; Suchparameter sichern

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
		For idx, line in StrSplit(searchstring, "`n") {
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
		filesavepath := RegExReplace(filesavepath, "i)\." FileExt "$") "." FileExt
		FileOpen(filesavepath, "w", "UTF-8").Write(table)
		If FileExist(filesavepath) {
			FileGetSize, FSize, % filesavepath
			MsgBox, 0x1000, Quicksearch, % "Die Datei wurde gespeichert.`nDatensätze: " rows "`nDateigröße: " (FSize > 1024 ? Round(FSize/1024, 1) " MB" : FSize " kb")
		} else
			MsgBox, 0x1000, Quicksearch, Die Datei wurde nicht gespeichert!

}

class Quicksearch_Properties                                                                              	{	; Listviewmanager

	 /* Quicksearch_Properties - Listview- and Properties managing class

		Funktionen:       	1. Ermittelt, sichert und stellt die vom Nutzer gemachten Veränderungen an der Oberfläche wieder her
									2. Betreibt eine virtuelle Tabelle damit Spalten aus- und wieder eingeblendet werden können
										dafür werden alle notwendigen Informationen aus dem DBase Header bezogen
									3. Sichert die eingebenen Suchparameter zusammen mit den Datenbankeinstellungen

									Alle Änderungen werden individuell für jede Datenbank vorgenommen und gespeichert.

		Hinweise:         		einige Klassenfunktionen wir ShowRemovedSets() können ohne Parameter aufgerufen werden
									und geben dann nur die Einstellung zurück


	*/

		__New(propertiesPath)                         	{ 	; Basisdatenobjekt anlegen, laden der Einstellungen

		; Dateipfade erstellen
			SplitPath, propertiesPath,, outFileDir,, outFileNameNoExt
			this.PropDir        	:= outFileDir
			this.PropName  	:= outFileNameNoExt
			this.PropFilePath   	:= this.PropDir "\" this.PropName ".json"
			this.DataPath    	:= RegExReplace(this.PropDir, "\\.*$", "")

		; Dateipfade prüfen
			resData := FilePathCreate(this.DataPath)
			resProp	:=  FilePathCreate(this.PropDir)
			this.qs 	:= this.LoadProperties()

		; Klartexte zu Versicherungsart
			this.VERSART   	:= ["Mitglied", "2", "Angehöriger", "4", "Rentner"]

		; Spaltennamen mit Datums- oder Zeitformaten für die automatische Umwandlung
			this.DateLabels	:= "ABVERKAUF|AUS2DAT|ABDAT|ABDATUM|ABGESETZT|ABNAHMDATE|ANTRAG|AUFNAHME|ANDATUM|AUFNAME|AUSSTDAT|AUSDATUM|"
										. "AU_DR_BIS|AU_BIS|ANLDATUM|"
										. "BESCHSEIT|BERICHDATE|BEZAHLTAM|BEZ_AM|BIRTHDAY|BIS|"
										. "DATETIME|DATUMPWD|DATUMV|DATUMVON|DATUMBIS|DATECREATE|DATECLOSE|DATSPEIC|DATE|DRUCKDATUM|DATUM|"
										. "EINAM|EINLESTAG|ENDE|ERINDAT|EDATUM|EINDATUM|FORMDATUM|FREIBIS|"
										. "GUELTIG|GUELTVON|GUELTBIS|GUELT_VON|GUELT_BIS|GUELTBIS2|GUELTVON2|GUELTIGAB|GUELTIGVON|GUELTIGBIS|"
										. "GEBURT|VERSBEG|GEBDATBIS|GEBDATVON|"
										. "KKDATUM|LAST_B_FR|LAST_B_TO|LAST_BEH|LESTAG|LASTCHANGE|MAHNDATE|MAHNAM|MORTAL|"
										. "NACH_AM|"
										. "PATGRFROM|PATGRTO|PATGEB|PRDATE|PATIENTVON|PATIENTBIS|PDATUM|RAUCHSTART|RAUCHEND|"
										. "STAT_VON|STAT_BIS|STORNODATE|SEIT|"
										. "TIMESTAMP|UNFTAG|"
										. "VALIDFROM|VALIDUNTIL|VERSSBEG|VON|"
										. "ZZIEL|ZIELDAT"
			this.noDLDB   	:= "TERMINE"
			this.HEXLabels  	:= "COLOR"
			this.TimeLabels := "ABNAHMTIME|ZIELZEIT|ERINZEIT|TIMESEND|ZEIT|UHRZEIT|UM"

		}

		WorkDB(dbname, dbfields)                  	{	; Arbeitsdatenbank festlegen und Tabellen-Informationen hinzufügen

			this.dbname := dbname

		; Dateipfad der Suchparameterdatei und Suchparameter laden
			this.SParamsPath  	:= this.PropDir "\QuickSearch-Suche_" this.dbname ".json"
			this.SParams        	:= this.SParams_Load()

		; Klasse wurde initialisiert?
			If !IsObject(this.qs)                            	{
				MsgBox, 1,  A_ThisFunc, "Die Klasse mittels __New() zuerst initialisiert werden"
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
			If (StrLen(this.qs[this.dbname].ShowRemovedSets) = 0)
				this.qs[this.dbname].ShowRemovedSets := false

		; Default labels anlegen für Datenbanksuche anlegen
		; Default Parameter anlegen
			this.SParamsDefaults := ""
			For idx, label in this.Labels()
				this.SParamsDefaults .= label "=`n"
			this.SParamsDefaults := RTrim(this.SParamsDefaults, "`n")
			this.PatIDCol          	:= false

		return 1
		}

		DataType(label)                                  	{	; automatische Spalteneinstellungen für die Windows Listview anhand der dBase Feldtypen
			static DataType 	:= {"N":"Integer", "D":"Integer", "C":"Text", "L":"Text", "M":"Text"}   ; L = Bool-Werte?
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

		HeaderSetVirtualWidth(label, width)    	{	; Pixelbreite einer Spalte aus dem realen Listviewelement speichern
			this.qs[this.dbname].header.text[label].width := width
		}

		HeaderSetVirtualPos(label, colpos)      	{	; Position eines Feldes in der virtuellen Tabelle setzen
			this.qs[this.dbname].header.text[label].col := colpos
			Gui, QS: ListView, QSDBFL                                              	; ändert die Positionsanzeige im Field-Info Listview
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

		PATNRColumn()                              		{	; gibt die Spalte mit der Patientennummer in der virtuellen Listview zurück

			static noPATNR := "(BGGOAE|bggoaedm|BGGOAUS|EBM_N|EBM_NDM|EBMAUS|EBMAUS_N|EBMEIN|EBMEIN_N|"
										. "PGSICH|sbartikl|SDPQ)"

			If RegExMatch(this.dbname, "i)\b" noPATNR "\b") {
				this.qs[this.dbname].PatID := "-"
				this.PatIDCol := 0
				return 0
			}

			PatIDLabel := RegExReplace(this.qs[this.dbname].PatID, "^-$")
			If (StrLen(PatIDLabel)=0 && !this.PatIDCol) {                              ; StrLen - muss bleiben!
				labels := this.Labels()
				For colNr, flabel in this.Labels() {
					If RegExMatch(flabel, "\b(PATNR|PATID|NR)\b") {
						PatIDLabel 	:= this.qs[this.dbname].PatID := flabel
						this.PatIDCol	:= this.HeaderVirtualPos(PatIDLabel)
						break
					}
				}
			}

		return PatIDLabel ? this.HeaderVirtualPos(PatIDLabel) : 0
		}

		LabelExist(label)                               		{	; true wenn die Spaltenüberschrift vorhanden ist

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

			this.loadProps := "fromOrigin"

		 ; Pfad prüfen, alternativen temperorären Pfad anlegen
			If !this.PropFilePath || !(PathString := RegExMatch(this.PropFilePath, "i)[A-Z]:\\.*\.json")) || !(PathExist := FilePathExist(this.PropDir)) {

				newpath := A_AppData "\Addendum"
				MsgBox, 0x1000, QuickSearch
									, % (!this.PropFilePath || !PathString ? "Es wurde ein" (!this.PropFilePath ? " leerer " : " inkorrekter") " Dateipfad übergeben.`n")
										. (!PathExist ? "Das übergebene Verzeichnis [" this.PropDir "] ist nicht angelegt." : "")
										. "Die Einstellungen können im Moment nur lokal gespeichert werden:`n"
										. "[" newpath "Quicksearch_UserProperties.json]`n"
										. "Die Speicherung auf einem Netzwerklaufwerk, bevorzugt`n"
										. "im Addendum-Ordner ist für die Arbeit mit dem Tool geeigneter!`n"
										. "Bitte beheben Sie das Verbindungs-/Dateisystemproblem.`n"

				this.PropDir        	:= newPath
				this.PropName   	:= "Quicksearch_UserProperties"
				this.PropFilePath   	:= this.PropDir "\" this.PropName ".json"

			}

		; Dateiexistenz prüfen
			If !FileExist(this.PropFilePath) {
				MsgBox, 0x1000, QuickSearch, % "Einstellungsdatei ist nicht vorhanden. Backupdatei verwenden?"
				IfMsgBox, No
					ExitApp
				this.loadProps := this.ReplacePropFileWithBackup()
			}
		; auf fehlerhafte Datei prüfen
			else {
				tmp := FileOpen(this.PropFilePath, "r", "UTF-8").Read()
				If !RegExMatch(tmp, "i)^[\s\n\r]*\{[:,\d\s\{\}\[\]\n\r_\-\+\pL]+" q) {
					MsgBox, 0x1000, QuickSearch, % "Einstellungsdatei ist leer oder hat ein falsches Format. Backupdatei verwenden?"
					IfMsgBox, No
						ExitApp
					this.loadProps := this.ReplacePropFileWithBackup()
				}
			}

		return RegExMatch(this.loadProps, "i)(fromBackup|fromOrigin)") ? JSONData.Load(this.PropFilePath, "", "UTF-8") : Object()
		}

		ReplacePropFileWithBackup()             	{	; die Originaldatei durch die neueste Backup-Datei ersetzen

		  ; notwendig gewordene Funktion nachdem 2x die Einstellungsdatei wahrscheinlich aufgrund eines Windows-Darstellungsfehlers
		  ; das SysListview-Steuerelement während des Speichervorganges einfror. Sämtliche Einstellungen waren verloren gegangen.
		  ; Zur Sicherheit legt das Skript Backupdateien an (mehrere im Falle das auch hier Schreibfehler vorkommen).

		  ; Backup vorhanden, dann die defekte Datei ersetzen
			backups := this.GetPropertiesBackups()
			If backups.filetime.newestIndex {

				BackupFilePath := this.PropDir "\" this.PropName "_backup" backups.filetime.newestIndex ".json"
				FileCopy, % BackupFilePath, % this.PropFilePath, 1
				If ErrorLevel {
					LastErrorMessage := FormatMessage(LastError := A_LastError)
					throw A_ThisFunc ": Der Versuch die defekte Einstellungsdatei wiederherzustellen war erfolglos.`n"
												. "Bitte beheben Sie zunächst die Ursache des Dateisystemfehlers:`n"
												.  LastErrorMessage " [" LastError "] "
					return "Exit"
				}

			}

		return "fromBackup"
		}

		GetPropertiesBackups()                        	{  ; den Namen der nächsten Backupdatei festlegen

			backups := Object()
			backups.Index := Array()
			backups.filetime := {"newest":0, "newestIndex":0, "oldest":0, "oldestIndex":0}

		  ; prüft jede Backupdatei auf Fehler, ermittelt die jüngste und älteste von maximal 3 Backupdateien
			Loop 3 {

				If FileExist(BackupPath := this.PropDir "\" this.PropName "_backup" A_Index ".json") {
					backups.Index.Push(A_Index)
					tmp := FileOpen(BackupPath, "r", "UTF-8").Read()
					If RegExMatch(tmp, "i)^[\s\n\r]*\{[:,\d\s\{\}\[\]\n\r_\-\+\pL]+" q) {
						FileGetTime, ftime, BackupPath, M
						If (!backups.filetime.newest || backups.filetime.newest < ftime)
							backups.filetime.newest := ftime, backups.filetime.newestIndex := A_Index
						If (!backups.filetime.oldest || backups.filetime.oldest > ftime)
							backups.filetime.oldest := ftime, backups.filetime.oldestIndex := A_Index
					}
				}
			}

		return backups
		}

		GetPropertiesBackupName()               	{	; Backupdateinamen erstellen

		 ; freien Dateinamen finden
			freeIndex 	:= 0
			backups 	:= this.GetPropertiesBackups()
			If (backups.Index.Count() > 0)
				Loop 3 {
					usedIndex := A_Index, indexFound := false
					For each, backupIndex in backups.Index
						If (backupIndex = usedIndex) {
							indexFound := true
							break
						}
					If !indexFound {
						freeIndex := usedIndex
						break
					}
				}

			newestIndex := filetime.newestIndex ? filetime.newestIndex : 0
			useIndex :=  freeIndex ? freeIndex : newestIndex ? newestIndex : 1

		return this.PropDir "\" this.PropName "_backup" useIndex ".json"
		}

		SaveProperties(saveColWidth:=true)    	{	; Objekt speichern

			global hQS, QS, QSDBRView

			; Dateipfad anlegen
				If !this.PropFilePath
					If !FilePathCreate(this.PropDir) {
						throw A_ThisFunc ": Speicherpfad für das Sichern der Einstellungen konnte nicht angelegt werden.`n"
								. "Ordner: " this.PropDir
						return
					}

			; Backup der alten Einstellungsdatei anlegen
				BackupFilePath := this.GetPropertiesBackupName()
				FileCopy, % this.PropFilePath, % BackupFilePath, 1
				If ErrorLevel {
					LastErrorMessage := FormatMessage(LastError := A_LastError)
					MsgBox, 0x1000, Quicksearch, "Es konnte aufgrund folgenden Dateisystemfehlers kein Backup erstellt werden:.`n"
																.  LastErrorMessage " [" LastError "] "
				}

			; Spaltenbreite speichern
				If saveColWidth && this.hLV
					this.LVColumnSizes()

			; Checkbox - gelöschte Datensätze anzeigen
				Gui, QS: Submit, NoHide
				this.qs[this.dbname].ShowRemovedSets := QSDBRView

			; Objekt speichern, nur wenn Gui ansprechbar ist und this.qs nicht leer ist
				timeouts := 0
				Loop 5 {
					timeouts += !CheckWindowStatus(hQS, 100) ? 1 : 0
					If (timeouts = 3), return 0
					Sleep 100
				}
				If (IsObject(this.qs) && this.qs.Count() > 0)
					JSONData.Save(this.PropFilePath, this.qs, true,, 1, "UTF-8")
				else {
					MsgBox, 0x1004, Quicksearch, % "Das interne Einstellungsobjekt scheint defekt zu sein.`n"
																	.  "Die Einstellungen konnten nicht gespeichert werden.`n"
																	.  "Das Skript wird nach bestätigen mit Ja beendet!"
					IfMsgBox, Yes
						ExitApp
					return 0
				}

		return FileExist(this.PropFilePath) ? 1 : 0
		}

		ShowRemovedSets(show:="")                	{	; gelöschte Datensätze anzeigen/nicht anzeigen
			global QS, QSDBRView
			static SRSLast
			If (StrLen(show) > 0) {
				this.qs[this.dbname].ShowRemovedSets := show ? 1 : 0
				If (SRSLast <> show) {
					GuiControl, QS:, QSDBRView, % (show ? 1 : 0)              	; Checkbox Gelöschte Datensätze anzeigen setzen oder entfernen
					SRSLast := show
				}
			}
		return this.qs[this.dbname].ShowRemovedSets
		}

		ShowPatNames(show:="")                  	{	; nicht nur die PATNR, sondern auch den Patientennamen anzeigen

			global QS, QSDBR
			static cutLen 	:= [10, 9, 8, 6, 6, 6]

		 ; Defaulteinstellung bei erstem Aufruf
			If (StrLen(this.qs[this.dbname].ShowPatNames) = 0)
				this.qs[this.dbname].ShowPatNames := show := 0

		  ; gibt die Einstellung zurück
			If (StrLen(show) = 0)
				return this.qs[this.dbname].ShowPatNames

		  ; Defaults einstellen
			Gui, QS: Default
			Gui, QS: ListView, QSDBR

		  ; neue Einstellung speichern
			this.qs[this.dbname].ShowPatNames := (show ? 1 : 0)

		  ; welche Spaltennummer hat PATNR derzeit
			PatNRCol := this.PATNRColumn()

		  ; Wertanzeige in der PATNR Spalte ändern
			Loop % LV_GetCount() {

				LV_GetText(PatCell, A_Index, PatNRCol)
				RegExMatch(PatCell, "(?<SetRemoved>\*)*.*?(?<NR>\d+)", Pat)
				if (show && PatCell<>"0") {
					PatName 	:= !PatDB[PatNR].Name && !PatDB[PatNR].Vorname ? "✘ gelöschter Patient ✘" : PatDB[PatNR].Name ", " PatDB[PatNR].Vorname
					PatNR   	:= (PatSetRemoved ? PatSetRemoved : " ") . SubStr("          [" PatNR "]" , -1*(cutLen[StrLen(PatNR)]))
					LV_Modify(A_Index, "Col" PatNRCol, PatNR "  " PatName)
				} else {
					LV_Modify(A_Index, "Col" PatNRCol, SubStr(" " PatSetRemoved, -1) PatNR)
				}

			}

			 LV_ModifyCol(PatNRCol, (Show ? "Left AutoHDR" : "Right AutoHDR"))

		}

		AutoDateFormat(autostate:="")           	{	; Autoformatierung von Datumszahlen ein und ausschalten

			global QSDBR, QS, QSDBRDFormat

		  ; Defaulteinstellung bei erstem Aufruf setzen
			If (StrLen(this.qs[this.dbname].AutoDateFormat) = 0)
				this.qs[this.dbname].AutoDateFormat := 1

		  ; gibt die Einstellung zurück
			If (StrLen(autostate) = 0)
				return this.qs[this.dbname].AutoDateFormat

		  ; Status hat sich nicht geändert
			If (this.qs[this.dbname].AutoDateFormat = autostate)
				return autostate

		  ; speichert die neue Einstellung
			this.qs[this.dbname].AutoDateFormat := autostate

		  ; Defaults einstellen
			;~ Gui, QS: Default
			;~ Gui, QS: ListView, QSDBR

		}
	;}

	; ------------------------------------------------------------------------------------
	; Suchfeld-Handler der aktuellen Tabelle                                        	;{
		SParams_Load()                                   	{	; gesicherte Suchparameter laden

			; Funktion wird bei jedem WorkDB() Aufruf ausgeführt. Die Suchparamenter müssen deshalb nicht extra geladen werden.

			If !InStr(FileExist(this.PropDir), "D") {
				throw A_ThisFunc ": Speicherordner <" this.PropDir "> ist nicht vorhanden."
				return
			}
			If FileExist(this.SParamsPath)
				SParams := JSONData.Load(this.SParamsPath, "", "UTF-8")
			else
				SParams := Object()

		return SParams
		}

		SParams_Save(delObject:=false)	       	{	; Suchparameter sichern

			If !InStr(FileExist(this.PropDir), "D")  {
				throw A_ThisFunc ": Speicherordner <" this.PropDir "> ist nicht vorhanden."
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

		SParams_GetParams(TextLabel="")        	{	; Parameter für Ausgabe in Edit-Control umwandeln

			If TextLabel && this.SParams.haskey(TextLabel)
				return RegExReplace(this.SParams[TextLabel], "[#]{2}", "`n")
			else
				return this.SParamsDefaults

		}

		SParams_CBSet(TextLabel)              		{	; Combobox befüllen, Suchparameter anzeigen

			global QS, QSBDS, QSDBSParam

			Gui, QS: Default
			Gui, QS: Submit, NoHide

			TextLabel := TextLabel = "LastLabels" ? this.SParams.lastLabel : TextLabel

			If (QSDBSParam <> TextLabel) {
				GuiControl, QS:                     	, QSDBSParam	, % "|" this.SParams_GetLabels()              	; Combobox befüllen ("|" am Anfang ersetzt die Liste)
				GuiControl, QS:  ChooseString	, QSDBSParam	, % TextLabel                                            	; letzte Sucheinstellung wählen
			}

			GuiControl	, QS:                       	, QSDBS        	, % this.SParams_GetParams(TextLabel)  	; und anzeigen

		}

		SParams_NewSearch()                        	{	; alle Suchparameter wiederherstellen + CB leeren

			Gui, QS: Default
			GuiControl, QS: Focus	, QSDBSParam
			Send, {BackSpace}

		}
	;}


	; +++++++++++++++++++++++++++++++++++++
	; +++++++++ PRIVATE FUNKTIONEN ++++++++++++   	;{
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
	;}

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

LVM_SetColOrder(h, c)                                                                                      	{ ; ##### ÄNDERN!

	; h = ListView handle.
	; c = 1 based indexed comma delimited list of the new column order.
	c := StrSplit(c, ",")
	VarSetCapacity(c, c0 * A_PtrSize)
	Loop, % c0
		NumPut(c%A_Index% - 1, c, (A_Index - 1) * A_PtrSize)

Return DllCall("SendMessage", "uint", h, "uint", 4154, "uint", c0, "ptr", &c) ; LVM_SETCOLUMNORDERARRAY
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

	static dpi := A_ScreenDPI / 96

	VarSetCapacity(RECT, 16, 0 )

	; SendMessage, 0xB2,, &RECT,, ahk_id %hEdit% ; EM_GETMARGIN
	DllCall("GetClientRect", "ptr", hEdit, "ptr", &RECT)
	right  := NumGet(RECT, 8, "Int")
	; bottom := NumGet(RECT, 12, "Int")

	NumPut(0     + Ceil(mLeft*dpi) , RECT, 0, "Int")
	NumPut(0     + Ceil(mTop*dpi)  , RECT, 4, "Int")
	NumPut(right - Ceil(mRight*dpi), RECT, 8, "Int")
	; NumPut(bottom - mBottom, RECT, 12, "Int")
	SendMessage, 0xB3, 0x0, &RECT,, % "ahk_id " hEdit ; EM_SETMARGIN
}

FormatMessage(MessageId) {
; ========================================================
; Function......: FormatMessage
; DLL...........: Kernel32.dll
; Library.......: Kernel32.lib
; U/ANSI........: FormatMessageW (Unicode) and FormatMessageA (ANSI)
; Author........: jNizM
; Modified......:
; Links.........: https://msdn.microsoft.com/en-us/library/ms679351.aspx
;                 https://msdn.microsoft.com/en-us/library/windows/desktop/ms679351.aspx
; =======================================================
    static size := 2024, init := VarSetCapacity(buf, size)
    if !(DllCall("kernel32.dll\FormatMessage", "UInt", 0x1000, "Ptr", 0, "UInt", MessageId, "UInt", 0x0800, "Ptr", &buf, "UInt", size, "UInt*", 0))
        return DllCall("kernel32.dll\GetLastError")
    return StrGet(&buf)
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
;~ #Include %A_ScriptDir%\..\..\lib\class_Neutron.ahk
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

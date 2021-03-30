; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                 	Addendum Quicksearch
;
;      Funktion:           	RegExSuche in allen Albisdatenbanken
;
;
;		Hinweis:
;
;		Abhängigkeiten:	siehe includes
;
;	                    			Addendum für Albis on Windows
;                        			by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;									begin: 24.03.2021,	last change 27.03.2021
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

; Einstellungen                                                                         	;{
	#NoEnv
	#Persistent
	#KeyHistory, Off

	SetBatchLines                  	, -1
	ListLines                        	, Off
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
	;adm.AlbisPath    	:= AlbisPath
	adm.AlbisPath    	:= "E:"
	adm.AlbisDBPath 	:= adm.AlbisPath "\DB"
	adm.compname	:= StrReplace(A_ComputerName, "-")                                                                 	; der Name des Computer auf dem das Skript läuft

  ; Nutzereinstellungen werden beim Erstellen des Objektes automatisch geladen
	Props := new Quicksearch_Properties(A_ScriptDir "\resources\Quicksearch_UserProperties.json")

  ; Patientendaten laden
	PatDB 	:= ReadPatientDBF(adm.AlbisDBPath, ["NR", "NAME", "VORNAME", "GEBURT"],, Anzeige)

;}

	SciTEOutput(" - - - - - - - - - - - - - - - - - ")
	QuickGui()

return

Esc::ExitApp
^!+::Reload

class Quicksearch_Properties {

		__New(propertiesPath) {

			this.db := Object()

			SplitPath, propertiesPath,, outFileDir,, outFileNameNoExt
			this.FileDir     	:= outFileDir
			this.FileName 	:= outFileNameNoExt ".json"
			this.FilePath     	:= this.FileDir "\" this.FileName

			If FilePathCreate(this.FileDir)
				If FileExist(this.FilePath)
					this.Props := this.LoadProperties()

		}

		WorkDB(dbname, dbfields)                  	{	; Arbeitsdatenbank festlegen und Tabellen-Informationen hinzufügen

			this.dbname := dbname

			If !IsObject(this.db)                            	{ 	; Objekt mit den Einstellung anlegen
				MsgBox, 1,  A_ThisFunc, "Die Klasse muss zuerst initialisiert werden"
				return
			}
			If !IsObject(this.db[this.dbname])        	{	; spezifische Einstellung dieser Datenbank
				this.db[this.dbname] := Object()
			}
			If !IsObject(this.db[this.dbname].header)	{	; Einstellungen der Spalten
				this.db[this.dbname].header       	:= Object()
				this.db[this.dbname].header.enum := Object()
				this.db[this.dbname].header.text 	:= Object()
			}

		  ; aus einem Objekt mach zwei
			For label, field in dbfields {

				If !this.db[this.dbname].header.enum.haskey(field.pos)
					this.db[this.dbname].header.enum[field.pos] := label

				If !IsObject(this.db[this.dbname].header.text[label]) {		                    	; default = sichtbar
					this.db[this.dbname].header.text[label]             	:= Object()
					this.db[this.dbname].header.text[label].Visible 	:= true
					this.db[this.dbname].header.text[label].name     	:= label
					this.db[this.dbname].header.text[label].pos       	:= field.pos
					this.db[this.dbname].header.text[label].Type      	:= field.Type
					this.db[this.dbname].header.text[label].len        	:= field.len
				}
			}

		return 1
		}

	; Teil-Objekte zurückgeben
		Field(label)                                          	{	; alle Daten eines Feldes (Spaltenüberschriften)
		return this.db[this.dbname].header.text[label]
		}

		Header()                                           	{	; Objekt mit Spaltenüberschriften und Daten
		return this.db[this.dbname].header.text
		}

		Labels()                                             	{	; Array mit Spaltenüberschriften
		return this.db[this.dbname].header.enum
		}

		Labels2() {
			arr:=[]
			For idx, field in this.header()
				If field.haskey("Label")
					arr[field.pos] := field.Label
		return arr
		}

	; Tabelleninformationen erhalten
		FieldsCount()                                     	{	; Anzahl von Spalten
		return this.db[this.dbname].header.enum.Count()
		}

		ColumnPos(Label)                             	{	; Position eines Feldes in der Tabelle
		return this.db[this.dbname].header.text[Label].pos
		}

		isVisible(Label)                                   	{	; Spalte ist sichtbar
		return this.db[this.dbname].header.text[Label].visible
		}

	; Tabelleninformationen setzen
		SetVisible(Label, status)                      	{	; Sichtbarkeitseinstellung einer Spalte ändern
			If !this.CheckBool(status, A_ThisFunc, "status")
				return
			this.db[this.dbname].header.text[Label].visible := status
		}

		LoadProperties()                                  	{	; Objekt laden
			If !this.FilePath || !FileExist(this.FilePath) {
				throw A_ThisFunc ": Einstellungsdatei <" this.FileName "> ist nicht im vorhanden.`n"
						. "Ordner: " this.FileDir
				return
			}
		return JSONData.Load(this.FilePath, "", "UTF-8")
		}

		SaveProperties()                                  	{	; Objekt speichern
			If !this.FilePath
				If !FilePathCreate(this.FileDir) {
					throw A_ThisFunc ": Speicherpfad für das Sichern der Einstellungen konnte nicht angelegt werden.`n"
							. "Ordner: " this.FileDir
					return
				}
			JSONData.Save(this.FilePath, this.Props, true,, 1, "UTF-8")
		return FileExist(this.FilePath) ? 1 : 0
		}

	; +++++++++ PRIVATE FUNKTIONEN ++++++++++++
		CheckLabel(label, FuncName) {
			return
			If (StrLen(label) = 0)
				return 0
			If !IsObject(this.db[this.dbname].header.text[label]) {
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

QuickGui() {

	;-: VARIABLEN                                              	;{

	  ; GUI
		global 	hQS, QSDB, QSDBFL, QSDBS, QSDBR, QSPGS, QST1, QST2, QST3
		global	QSMinKB      	, QSMaxKB    	, QSFilter                                       	; Eingabefelder
		global	QSMinKBText	, QSMaxKBText	, QSFilterText                                  	; Textfelder
		global 	QShDB         	, QShDBFL    	, QShDBR                                         ; hwnd
		global	QSFeldWahlT	, QSFeldWahl	, QSFeldAlle	, QSFeldKeine
		global	QSSearch, QSSaveRes, QSSaveSearch, QSQuit, QSReload                	; Buttons

	; ANDERES
		global dbfiles

		static EventInfo, empty := "- - - - - - -"
		static FSize         	:= (A_ScreenHeight > 1080 ? 9 : 8)
		static gcolor        	:= "BBBB99"
		static tcolor        	:= "e4e4c4"
		static DBFilesW   	:= 320
		static DBFieldW   	:= 250
		static DBSearchW	:= 500
		static gQSDBFilter 	:= "gQuicksearchDBFilter"
		static gQSDBFields	:= "gQuicksearchDBFields"

		dbfiles       	:= DBASEStructs(adm.AlbisDBPath)

		IniRead, QSMinKB 	, % A_ScriptDir "\" A_ScriptName, % "Einstellungen", QSMinKB		, % ""
		IniRead, QSMaxKB	, % A_ScriptDir "\" A_ScriptName, % "Einstellungen", QSMaxKB	, % ""
		IniRead, QSFilter   	, % A_ScriptDir "\" A_ScriptName, % "Einstellungen", QSFilter    	, % ""

	;}

	Gui, QS: new	, HWNDhQS -DPIScale

	;-: SPALTE 1: DATENBANKEN                       	;{

		; - Filter -
		Gui, QS: Color	, % "c" gcolor , % "c" tcolor
		Gui, QS: Font	, % "s" FSize-1 " q5 Normal", Futura Bk Bt
		Gui, QS: Add	, Text        	, % "xm ym-5	BackgroundTrans             	vQSMinKBText"	, % "min KB"
		Gui, QS: Add	, Text        	, % "x+10 	BackgroundTrans             	vQSMaxKBText"	, % "max KB"
		Gui, QS: Add	, Text        	, % "x+10 	BackgroundTrans          	vQSFilterText"   	, % "DB-Filter"
		; - Felder -
		Gui, QS: Add	, GroupBox  	, % "x+10 w50 h10                           	vQSFeldWahlT"	, % "Felder wählen"

		Gui, QS: Font	, % "s" FSize-1 " q5 Normal", Futura Bk Bt
		Gui, QS: Add	, Edit           	, % "xm y+2 	w" FSize*8 	" Number	BackgroundTrans                	" gQSDBFilter " vQSMinKB"	, % QSMinKB
		Gui, QS: Add	, Edit           	, % "x+5    	w" FSize*8 	" Number	BackgroundTrans               	" gQSDBFilter " vQSMaxKB"	, % QSMaxKB
		Gui, QS: Add	, Edit           	, % "x+5    	w" FSize*19	"           	BackgroundTrans AltSubmit	" gQSDBFilter " vQSFilter"   	, % QSFilter
		; - Felder -
		Gui, QS: Add	, Button        	, % "x+10 	BackgroundTrans            	vQSFeldAlle" 	, % "Alle"
		Gui, QS: Add	, Button        	, % "x+5   	BackgroundTrans            	vQSFeldKeine"	, % "Keine"
		Gui, QS: Font	, % "s" FSize+2 " q5 Italic", Futura Bk Bt
		Gui, QS: Add	, Text        	, % "x+15 ym-2 cBlue BackgroundTrans  Center                     "	, % "Rückgabewerte`nauswählen"

		; - Filter Elemente versetzen -
		GuiControlGet, cp, QS:Pos, QSMinKBText
		GuiControlGet, dp, QS:Pos, QSMinKB
		GuiControl, QS: Move, QSMinKBText, % "x" dpX + Floor((dpW/2) - (cpW/2))

		GuiControlGet, cp, QS:Pos, QSMaxKBText
		GuiControlGet, dp, QS:Pos, QSMaxKB
		GuiControl, QS: Move, QSMaxKBText, % "x" Floor(dpX + (dpW//2 - cpW//2))

		GuiControlGet, cp, QS:Pos, QSFilterText
		GuiControlGet, dp, QS:Pos, QSFilter
		GuiControl, QS: Move, QSFilterText, % "x" Floor(dpX + (dpW/2 - cpW/2))

		; - Felder Texte versetzen -
		gbY := cpY-1,	gpmrg := 2
		GuiControlGet, cp, QS:Pos, QSFeldAlle
		GuiControlGet, dp, QS:Pos, QSFeldKeine
		GuiControl, QS: Move, QSFeldWahlT, % "x" cpX-2*gpmrg " y" gbY " w" (dpX+dpW-cpX+4*gpmrg) " h" cpW+5*gpmrg

		; - Datenbankauswahl
		LVOptions1 := "vQSDB HWNDQShDB gQS_Handler AltSubmit -Multi Grid"
		GuiControlGet, cp, QS:Pos, QSMinKB
		GuiControlGet, dp, QS:Pos, QSFilter
		mrgX := cpX, DBFilesW := dpX+dpW-mrgX, GCol2X := dpX+dpW+5
		Gui, QS: Font	, % "s" FSize " q5 Normal"
		Gui, QS: Add	, ListView    	, % "xm				 y+5 	w" DBFilesW 	" r16  " LVOptions1	, % "Datenbank|records|Größe|letzte Änderung"

		QuicksearchDBFilter()

		Gui, QS: ListView, QSDB
		LV_ModifyCol()
		LV_ModifyCol(1, 90)
		LV_ModifyCol(2, "Right Integer")
		LV_ModifyCol(3, "Right Integer")
		LV_ModifyCol(4, "Center")
		LV_ModifyCol(2, "SortDesc")
	;}

	;-: SPALTE 2: FELDINFOs                           	;{
		LVOptions2	:= gQSDBFields " HWNDQShDBFL -LV0x18 AltSubmit -Multi Grid Checked"
		GuiControlGet, cp, QS:Pos, QSDB
		Gui, QS: Font	, % "s" FSize-1 " q5 Normal"
		Gui, QS: Add	, ListView    	, % "x" GCol2X " y" cpY " w" DBFieldW	" h" cpH " vQSDBFL " LVOptions2	, % "Nr|Feldname|Type|Länge"

		LV_ModifyCol(1, "40 Right Integer")
		LV_ModifyCol(2, "85 Text")
		LV_ModifyCol(3, "40 Center")
		LV_ModifyCol(4, "45 Center")

		GuiControlGet, dp, QS:Pos, QSDBFL
		GCol3X 	:= dpX+dpW+5
	;}

	;-: SPALTE 3: DATENBANKSUCHE                 	;{
		GuiControlGet, cp, QS:Pos, QSDB
		GuiControlGet, dp, QS:Pos, QSDBFL
		Gui, QS: Add	, Edit         	, % "x+5	y" dpY " w" DBSearchW	" h" dpH " vQSDBS  	gQS_Handler "
		GuiControlGet, cp, QS:Pos, QSDBS
		GClientW := cpX+cpW-mrgX
	;}

	;-: ZEILE 2: Befehle                                    	;{
		Gui, QS: Font	, % "s" FSize+1 " q5", Futura Md Bt
		Gui, QS: Add	, Text        	, % "xm 	y+5 	BackgroundTrans "			, % "Position:"
		Gui, QS: Font	, % "Normal", Futura Md Bt
		Gui, QS: Add	, Text        	, % "x+5       	BackgroundTrans vQST1", % empty

		Gui, QS: Font	, % "Normal", Futura Md Bt
		Gui, QS: Add	, Text        	, % "x+20       	BackgroundTrans"				, % "Datensätze:"
		Gui, QS: Font	, % "Normal", Futura Md Bt
		Gui, QS: Add	, Text        	, % "x+5       	BackgroundTrans vQST2", % empty

		Gui, QS: Font	, % "Normal", Futura Md Bt
		Gui, QS: Add	, Text        	, % "x+20       	BackgroundTrans"				, % "Treffer:"
		Gui, QS: Font	, % "Normal", Futura Md Bt
		Gui, QS: Add	, Text        	, % "x+5       	BackgroundTrans vQST3", % empty "               "

		; - Progress -
		Gui, QS: Add	, Progress     	, % "xm  	y+0 w" GClientW " vQSPGS"					, % 0

		; - Buttons -
		Gui, QS: Add	, Button        	, % "      	y+5 	vQSSearch      	gQS_Handler ", % "Suchen"
		Gui, QS: Add	, Button        	, % "x+10		  	vQSSaveRes    	gQS_Handler ", % "Ergebnisse speichern"
		Gui, QS: Add	, Button        	, % "x+20		  	vQSSaveSearch 	gQS_Handler ", % "Sucheinstellungen speichern"
		Gui, QS: Add	, Button        	, % "x+50		  	vQSReload           	gQS_Handler ", % "Reload"
		Gui, QS: Add	, Button        	, % "x+200	  	vQSQuit           	gQS_Handler ", % "Beenden"

	;}

	;-: ZEILE 3: Ausgabelistview                           	;{
		LVOptions3 := "vQSDBR gQS_Handler HWNDQShDBR -E0x200 -LV0x10 AltSubmit -Multi Grid"
		Gui, QS: Font	, % "s" FSize-1 " q5 Normal", Futura Bk Bt
		Gui, QS: Add	, ListView    	, % "xm y+5 w" GClientW	" r25 " LVOptions3	, % "- -"
	;}

	Gui, QS: Show, % "x1300 y10 AutoSize Hide", Quicksearch
	hMon	:= MonitorFromWindow(hQS)
	mon  	:= GetMonitorInfo(hMon)
	wqs  	:= GetWindowSpot(hQS)
	Gui, QS: Show, % "x" mon.R-wqs.W " y" Floor(mon.B/2-wqs.H/2) " NA", Quicksearch

return

QS_Handler: ;{

	Switch A_GuiControl	{

		Case "QSReload":      	;{
			Reload
		;}

		Case "QSDB":      	;{
			If (A_EventInfo = EventInfo) || (StrLen(A_EventInfo) = 0)
				return
			EventInfo := A_EventInfo
			If Instr(A_GuiEvent	, "DoubleClick") {
				; Datenbankname lesen
					Gui, QS: ListView, QSDB
					LV_GetText(dbname, EventInfo, 1)
					QuicksearchListFields(dbname)

			}
		return ;}

		Case "QSSearch": 	;{

			Gui, QS: Submit, NoHide
			Quicksearch(dbname, QSDBS)

		return ;}

	}

return ;}

QSGuiClose:
QSGuiEscape: ;{
	Gui, QS: Submit, NoHide
	IniWrite, % QSMinKB 	, % A_ScriptDir "\" A_ScriptName, % "Einstellungen", QSMinKB
	IniWrite, % QSMaxKB	, % A_ScriptDir "\" A_ScriptName, % "Einstellungen", QSMaxKB
	IniWrite, % QSFilter   	, % A_ScriptDir "\" A_ScriptName, % "Einstellungen", QSFilter
	If !Props.SaveProperties()
		MsgBox, Einstellungen konnten nicht gesichert werden!
ExitApp ;}
}

Quicksearch(dbname, srstr) {

		; {"KUERZEL": "lko"}, ["DATUM", "PATNR", "INHALT"]
		; (COVID|COVIPC|Sars|Covid|Cov\-)
		global dbfiles
		pattern := Object()

		For idx, line in StrSplit(srstr, "`n") {
			RegExMatch(Trim(line), "(?<field>.*)?\=(?<rx>.*)$", m)
			If mrx
				pattern[mfield] := "rx:" mrx
		}

		dbf       	:= new DBASE(adm.AlbisDBPath "\" dbname ".dbf", 2)
		filepos   	:= dbf.OpenDBF()
		matches  	:= dbf.Search(pattern, 0, "QuicksearchProgress")
		filepos   	:= dbf.CloseDBF()

		dbfields := dbfiles[dbname].dbfields
		Gui, QS: ListView, QSDBR
		LV_Delete()
		For rowNr, set in matches {
			; Spalten zusammenstellen
				cols := Object()
				For flabel, value in set
					If Props.isVisible(fLabel) && (colpos := Props.ColumnPos(flabel))
						cols[colpos] := value
			; Zeile hinzufügen
				LV_Add("", cols*)
		}
		LV_ModifyCol()



		;~ JSONData.Save(A_Temp "\Quicksearch.json", matches, true,, 2, "UTF-8")
		;~ Run, % A_Temp "\Quicksearch.json"


}

QuicksearchProgress(index, maxIndex, len, matchcount) {

	global	DBSearchW
	prgPos := Floor((index*100)/maxIndex)
	GuiControl, QS:, QSPGS	, % prgPos
	GuiControl, QS:, QST1	, % index
	GuiControl, QS:, QST3	, % matchcount

}

QuicksearchListFields(dbname) {

		global dbfiles, empty, dbfdata

	; Einstellungsobjekt anlegen                              	;{
		If !Props.WorkDB(dbname, dbfiles[dbname].dbfields) {
			throw A_ThisFunc ": Einstellungen zur Datenbank: " dbname " konnten nicht angelegt werden!"
			return
		}
	;}

	; neue Spalten anlegen                                     	;{
		LV_DeleteCols("QS", "QSDBR")                           	; Ausgabelistview - Spalten entfernen
		For colNr, label in Props.Labels()
			If (colNr = 1)
				LV_ModifyCol(1, "AutoHDR", Label)
			else
				LV_InsertCol(colNr, "AutoHDR", Label)

		LV_ModifyCol(1, 50)
	;}

	; Datenbankfelderinfo's anzeigen                     	;{
		Gui, QS: ListView, QSDBFL
		LV_Delete()

		For colNr, label in Props.Labels() {
			field := Props.Field(label)
			LV_Add(Props.isVisible(Label)?"Check":"", colNr, field.name, field.Type, field.len)
			t .= label "=`n"
		}
		LV_ModifyCol()
	;}

	; Statistik                                                          	;{
		GuiControl, QS:, QSDBS, % t
		GuiControl, QS:, QST1	, % empty
		GuiControl, QS:, QST2	, % dbfiles[dbname].records
		GuiControl, QS:, QST3	, % empty
	;}

	; Daten aus Albis Datenbank lesen                  	;{
		dbf       	:= new DBASE(adm.AlbisDBPath "\" dbname ".dbf", 2)
		filepos   	:= dbf.OpenDBF()
		dbfdata    	:= dbf.ReadBlock(1000,, 0)
		filepos   	:= dbf.CloseDBF()
	;}

	; Vorschau der Daten im Ergebnisfenster          	;{
		Gui, QS: ListView, QSDBR
		LV_Delete()
		For rowNr, set in dbfdata {

			; Spalten zusammenstellen
				cols := Object()
				For flabel, value in set
					If Props.isVisible(fLabel) && (colpos := Props.ColumnPos(flabel))
						cols[colpos] := value

			; Zeile hinzufügen
				LV_Add("", cols*)

		}
		LV_ModifyCol()
	;}

}

QuicksearchDBFields(GChwnd:="", GuiEvent:="", EventInfo:="", ErrLevel:="") {

		global dbfiles, dbname, QShDBFL, QShDBR

	; Checkboxstatus hat sich geändert
		If RegExMatch(GuiEvent . EventInfo, "(K32|C0)") {

		  ; aktuelle Spalten der Ausgabelistview ermitteln
			LVHeader := LV_GetHeader(QShDBR, "QS", "QSDBR")

		  ; Feldinfo Listview auf Default setzen
			Gui, QS: ListView, QSDBFL

		  ; zweilenweise Überprüfung des Checkboxstatus
			Loop, % LV_GetCount() {

			  ; Checkboxstatus
				ischecked := LV_RowIsChecked(QShDBFL, A_Index)
				flabel := DBFiles_FieldFromPos(A_Index)

			  ; ----------------------------------------------------------------------------
			  ; eine Spalteneinstellung hat sich geändert
			  ; ----------------------------------------------------------------------------
			   SciTEOutput(A_Index ": " Props.db[dbname].header[fLabel].checked " - "  ischecked)
				If (Props.db[dbname].header[fLabel].checked <> ischecked) {

					; Änderung speichern
						Props.db[dbname].header[fLabel].checked := ischecked

					; Ausgabelistview auf Default setzen
						Gui, QS: ListView, QSDBR

					; -----------------------------------------------------------------------------------
					; Spalte einfügen
					; 	eingefügt wird nach der ersten vorhandenen Spalte der
					; 	Ausgabelistview welche sich vor der Spalte in der Datenbank findet
					; -----------------------------------------------------------------------------------
						If ischecked && !LVHeader.name.haskey(fLabel) {
							colNr := DBFiles_PosFromField(flabel), InsertCol := false
							Loop % colNr {
								colLabel := DBFiles_FieldFromPos(prevCol := colNr-A_Index)
								If  (prevCol = 0) || LVHeader.name.haskey(colLabel) {
									LV_InsertCol(InsertCol := prevCol+1,, flabel)
									break
								}
							}
							If !InsertCol
								LV_InsertCol(1,, flabel)

						  ; fügt die fehlenden Daten hinzu
							Gui, QS: ListView, QSDBR
							For row, field in dbfdata
								LV_Modify(row, "Col" colNr, field[flabel])
						}
					; ----------------------------------------------------------------------------
					; Spalte entfernen
					; ----------------------------------------------------------------------------
						else if !ischecked && LVHeader.name.haskey(fLabel) {
							colNr := LVHeader.name[flabel]
							LV_DeleteCol(colNr)
						}

				}

			}

		}

}

QuicksearchDBFilter(GChwnd:="", GuiEvent:="", EventInfo:="", ErrLevel:="") {

	global	QSMinKB      	, QSMaxKB    	, QSFilter
	global	dbfiles

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


}

QSFormat(timestamp) {
	FormatTime, timeformated, % timestamp "000000", yyyy.MM.dd
return timeformated
}

DBFiles_FieldFromPos(pos) {
	global dbfiles, dbname
	For flabel, field in dbfiles[dbname].dbfields
		If (field.pos = pos)
			break
return flabel
}

DBFiles_PosFromField(field) {
global dbfiles, dbname
return dbfiles[dbname].dbfields[field].pos
}

LV_DeleteCols(GuiName, LVName, ColRange:="All" ) {

	Gui, % GuiName ": ListView", % LVName
	resultCols := LV_GetCount("Column")
	If (resultCols > 1)
		If (ColRange = "All") {
			LV_Delete()
			Loop % resultCols-1
				LV_DeleteCol(resultCols-(A_Index-1))
		}

}

LV_GetHeader(hLV, GuiName, LVName) {

	LVHeader           	:= Object()
	LVHeader.name	:= Object()
	LVHeader.nr       	:= Array()

	Gui, % GuiName ": ListView", % LVName

	Loop, % LV_GetCount("Colum") {
		LV_GetText(colName, 0, A_Index)
		LVHeader.name[colName] := A_Index
		LVHeader.nr.Push(colName)
	}

return LVHeader
}

LV_RowIsChecked(hLV, row) {
	SendMessage, 0x102C, row-1, 0xF000,, % "ahk_id " hLV  ; 0x102C ist LVM_GETITEMSTATE. 0xF000 ist LVIS_STATEIMAGEMASK.
return (ErrorLevel >> 12) - 1
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
QSFilter=



If (colNr = 1)
						LV_InsertCol(1,, flabel)
					else {





 */

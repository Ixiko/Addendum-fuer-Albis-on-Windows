; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                                           ADDENDUM - GESUNDHEITSVORSORGEN -
;
;    	Funktion:           	- 	zeigt alle Patienten an bei denen eine Vorsogeuntersuchung im Abrechnunsquartal abgerechnet werden kann
;
;       Features:      			- 	holt sich alle Daten aus der Albisdatenbank und ist dadurch sehr schnell in der Auswertung
;                       			-	Albis-Datenbank Pfad wird selbstständig aus der Windowsregistrierung gelesen
;                        			- 	Filterung nach KV-Abrechnungsregeln, die auch selbst festgelegt werden können
;                       			 	Ausschlußregeln - Mindestalter oder letzte Vorsorgeuntersuchung ist zu lange her
;                        			- 	vorgeschlagene Patienten werden in einer Datei im Addendum-Verzeichnis im JSON-Format gespeichert
;                        			- 	GUI mit Anzeige gefundener Patienten, Aufruf der Karteikarte per Doppelklick im Listviewfenster
;
;    	Basisskript:        	-	keines
;
;		Abhängigkeiten: 	-	include\Addendum_Albis.ahk, _Controls.ahk, _DB.ahk, _DBASE.ahk, _Datum.ahk, _Misc.ahk ...
;                                   -	lib\SciTEOutput.ahk
;                                   -	lib\class_JSON.ahk
;                                   -	lib\GDIP.ahk
;                                   -	lib\ACC.ahk
;
;	                    	Addendum für Albis on Windows
;                        	by Ixiko started in September 2017 - last change 27.12.2020 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

;-------------------------------------------------------------------------------------------------------------------------------------------
; Skripteinstellungen
;-------------------------------------------------------------------------------------------------------------------------------------------;{
	#NoEnv
	#KeyHistory, Off
	#MaxMem 4096

	SetBatchLines, -1
	ListLines    	, Off
	SetRegView	, 64

	SciTEOutput()

;}

;-------------------------------------------------------------------------------------------------------------------------------------------
; Objekte / Variablen / Daten (PatDB, GVUListe)
;-------------------------------------------------------------------------------------------------------------------------------------------;{

	; wichtige Pfade
		RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)        ; Addendum-Verzeichnis
		RegRead, PathAlbis, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-MainPath

		admDBPath    	:= AddendumDir "\logs'n'data\_DB"
		basedir         	:= PathAlbis "\db"
		propertiesPath	:= A_ScriptDir "\Einstellungen"
		pathGVUListe 	:= AddendumDir "\Tagesprotokolle"

	; Patientendatenpfad anlegen
		If !InStr(FileExist(admDBPath "\PatData"), "D")
			FileCreateDir, % admDBPath "\PatData"

	; Skript-Einstellungen laden/anlegen
		If !InStr(FileExist(propertiesPath), "D")
			FileCreateDir, % propertiesPath
		If !FileExist(propertiesPath "\Gesundheitsvorsorgen.json") {
			props := Object()
			props.minAbstand := 3
			props.minAlter    	:= 35
			FileOpen(propertiesPath "\Gesundheitsvorsorgen.json", "w", "UTF-8").Write(JSON.DUMP(props,,2))
		} else {
			props := JSON.Load(FileOpen(propertiesPath "\Gesundheitsvorsorgen.json", "r", "UTF-8").Read())
		}

	; minimaler GVU Untersuchungsabstand
		minAbstand  	:= props.minAbstand     	; Jahre
		minAlter            := props.minAlter         	; Abrechnung möglich ab dem Alter von

	; das aktuelle Abrechungsquartal
		Abrechnungsquartal 	:= "0420"
		AbrQ      	:= QuartalTage({"Aktuell":Abrechnungsquartal})

	; Patientendaten aus der PATIENT.DBF laden
		infilter 	:= ["NR", "NAME", "VORNAME", "GEBURT", "MORTAL", "LAST_BEH", "RECHART"]
		PatDBF	:= ReadPatientDBF(basedir, infilter,, 1)

	; vorhandene Untersuchungen laden
		GVUListe := ReadGVUListe(pathGVUListe, Abrechnungsquartal)

		;FileOpen(A_ScriptDir "\GVUPatienten.JSON", "w", "UTF-8").Write(JSON.DUMP(PatDBF,,2))

		;gosub JumpListQuartale
		;ExitApp
;}

;-------------------------------------------------------------------------------------------------------------------------------------------
; GUI
;-------------------------------------------------------------------------------------------------------------------------------------------;{
	Gui:

		; VARIABLEN/GUIGRÖSSE
			FSize     	:= (A_ScreenHeight > 1080 ? 10 : 9)
			Font	    	:= Calibri
			tcolor    	:= "DDDDBB"

		;-: Gui  Anfang
			Gui, GV: new	, HWNDhGVU -DPIScale +AlwaysOnTop
			Gui, GV: Color	, cCCCCAA

			Gui, GV: Font	, % "s" FSize+4 " Bold q5", % Font
			Gui, GV: Add	, Text	, xm ym, Abrechnungsregeln

			Gui, GV: Font	, % "s" FSize " Normal q5", % Font
			Gui, GV: Add	, Text	, % "xm y+" FSize+2, Abstand Untersuchungen:
			Gui, GV: Add	, Edit 	, % "x+2 w" FSize*2 " r1 vGVAbstand gGuiGVHandler", % minAbstand
			Gui, GV: Add	, Text	, % "x+2 ", (Jahre)

			Gui, GV: Add	, Text	, % "xm y+" FSize, Mindestalter Patient:
			Gui, GV: Add	, Edit 	, % "x+2 w" FSize*2 " r1 vGVAlter gGuiGVHandler", % minAlter
			Gui, GV: Add	, Text	, % "x+2 ", (Jahre)

			Gui, GV: Font	, % "s" FSize+4 " Bold q5", % Font
			Gui, GV: Add	, Text	, xm y+20, Abrechnungsquartal

			Gui, GV: Font	, % "s" FSize+1 " Normal q5", % Font
			Gui, GV: Add	, Edit 	, % "xm y+" FSize+1 " w" FSize*6 " r1 vGVQuartal gGuiGVHandler", % Abrechnungsquartal

			Gui, GV: Font	, % "s" FSize+1 " Normal q5", % Font
			Gui, GV: Add	, Button	, % "xm y+20 vGVErstellen gGuiGVHandler", % "Patientenliste erstellen"

			GuiControlGet, cl, GV: Pos, GVAbstand
			GuiControl, GV: Move, GVAbstand	, % "y" clY - 2
			GuiControlGet, cl, GV: Pos, GVAlter
			GuiControl, GV: Move, GVAlter     	, % "y" clY - 2

			If FileExist(pathGVUListe "\" Abrechnungsquartal "-GVUVorschlag.JSON") {
				Gui, GV: Add, Button	, % "xm y+10 vGVVS HWNDhGVVS gGuiGVHandler", % "GVU Vorschläge anzeigen (Q " Abrechnungsquartal ")"
			} else
				Gui, GV: Add, Button	, % "xm y+10 vGVVS HWNDhGVVS gGuiGVHandler", % "GVU Vorschläge erstellen (Q " Abrechnungsquartal ")"

			;~ If FileExist(admDBPath "\PatData\PatientExtra.json") {
				;~ Gui, GV: Add, Button	, % "xm y+10 vGVPE HWNDhGVPE gGuiGVHandler", % "Ziffernvorschläge anzeigen (Q " Abrechnungsquartal ")"
			;~ } else
				Gui, GV: Add, Button     	, % "xm y+10         	vGVPE   	HWNDhGVPE 	gGuiGVHandler", % "Ziffernvorschläge erstellen"
				Gui, GV: Add, Button     	, % "xm y+10         	vGVSL   	HWNDhGVJLQ	gGuiGVHandler", % "Sprungliste Quartale anlegen"
				Gui, GV: Add, Button     	, % "xm y+10         	vGVCKF	HWNDhGVCKF	gGuiGVHandler", % "2.Chronikerziffer fehlend"
				Gui, GV: Add, Button     	, % "xm y+10         	vGVGLI 	HWNDhGVGLI 	gGuiGVHandler", % "gelöschte Inhalte anzeigen"
				Gui, GV: Add, Button     	, % "xm y+10         	vGVFS  	HWNDhGVFS 	gGuiGVHandler", % "Formularstatistik"
				Gui, GV: Add, Edit        	, % "xm y+15 w75 	vGVID     	HWNDhGVID"
				Gui, GV: Add, ComboBox	, % "x+5    	 w100 	vGVKZL	HWNDhGVKZL" , % "aem|anam|bef|dia|ekg|info|lufu|lzrr|ther|z"
				GuiControl, GV: ChooseString, GVKZL, bef
				Gui, GV: Add, Button     	, % "x+10               	vGVBT  	HWNDhGVBT 	gGuiGVHandler", % "Befundtexte "

				Gui, GV: Font	, % "s" FSize-4 " Normal q5", % Font
				ControlGetPos, cpX, cpY,, cpH,, % "ahk_id " hGVID
				Gui, GV: Add, Text        	, % "x" cpX " y" (cpY-cpH-FSize-2), % "Patienten Nr"
				ControlGetPos, cpX, cpY,, cpH,, % "ahk_id " hGVKZL
				Gui, GV: Add, Text        	, % "x" cpX " y" (cpY-cpH-FSize-2), % "Karteikartenkürzel"

			Gui, GV: Font	, % "s" FSize-2 " Normal q5", % Font
			Gui, GV: Add	, Text	, xm y+40 vGVTTip1, % 	   "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
			Gui, GV: Add	, Text	, xm       	vGVTTip2, % 	   "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
			Gui, GV: Add	, Text	, xm 			vGVTTip3, % 	   "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"

			Gui, GV: Show, AutoSize Hide, Abrechnungsvorschläge

			gui := GetWindowSpot(hGVU)
			lvCols := "PATID|Name|Vorname|Geb.Tag|Ziffern|letzte GVU|Beh.Tage"

			Gui, GV: Font	, % "s" FSize-1 " Normal q5", % Font
			Gui, GV: Add	, Text	, % "x" gui.CW " ym vGVStats ", % "zusätzliche Patienten: " GVUVorschlag.Count() ", in Liste: " GVUListe.Count() ", Pat. in DB: " PatDBF.Count()
			GuiControlGet, cp, GV: Pos, GVStats

			Gui, GV: Font	, % "s" FSize-1 " Normal q5", % Font
			Gui, GV: Add, Listview	, % "x" gui.cw " y+5 w" StrLen(lvCols)*(FSize-1) " h" gui.h - cpH - 5 " vGVLV gGuiGVHandler", % lvCols

			Gui, GV: Show, x200 y400 AutoSize


		;gosub ZiffernSammler

	return

	GuiGVHandler:   	;{

			If (A_GuiControl = "GVLV") {

				If(A_GuiEvent = "DoubleClick") {

					Gui, GV: ListView, GVLV
					LV_GetText(LVPatID, (crow := A_EventInfo), 1)

					AlbisActivate(2)
					If AlbisAkteOeffnen(LVPatID) {
						Gui, GV: ListView, GVLV
						LV_Delete(crow)
						GVUListe := ReadGVUListe(pathGVUListe, Abrechnungsquartal)
						GuiControl, GV:, GVStats, % "zusätzliche Patienten: " LV_GetCount() ", in Liste: " GVUListe.Count() ", Pat. in DB: " PatDBF.Count()
					}

				}

			}
			else If (A_GuiControl = "GVVS") {
				Gui, GV: Submit, NoHide
				GuiControlGet, btnTxt, GV:, GVVS
				If InStr(btnTxt, "erstellen")
					gosub GVUVorschlaege

			}
			else If (A_GuiControl = "GVPE") {
				Gui, GV: Submit, NoHide
				GuiControlGet, btnTxt, GV:, GVPE
				If InStr(btnTxt, "erstellen")
					gosub ZiffernSammler

			}
			else If (A_GuiControl = "GVSL") {
					gosub JumpListQuartale

			}
			else If (A_GuiControl = "GVCKF") {
					gosub Chronikerziffernfehlt

			}
			else If (A_GuiControl = "GVGLI") {
					gosub GelöschteZiffern

			}
			else If (A_GuiControl = "GVBT") {
					gosub BefundTexte

			}
			else If (A_GuiControl = "GVFS") {
					gosub Formulare_zaehlen

			}

	return ;}

	GVGuiClose:
	GVGuiEscape:
	ExitApp

;}


Impfungen: ;{

	inpattern   	:= {"PATNR":"15319", "KUERZEL":"dia", "INHALT":"rx:Z2[3-7]"}
	outpattern 	:= ["DATUM", "PATNR", "INHALT"]
	ziffern       	:= GetDBFData(basedir "\BEFUND.dbf", inpattern, outpattern,, 1)
	For zidx, m in ziffern
		t .= ConvertDBASEDate(m.Datum) ":`t" m.Inhalt "`n"

return ;}

Formulare_zaehlen: ;{

	inpattern   	:= {"KUERZEL":"rx:[A-Za-z\s]{1,5}", "DATUM":"rx:20[12][0-9]\d{4}"}
	outpattern 	:= ["DATUM", "KUERZEL", "INHALT"]		                                            	; DSVGO = Löschsymbol?
	;GetDBFData(DBASEfilepath, inpattern:="", outpattern:="", seek:=0, options:="", debug:=0)
	dbfData     	:= GetDBFDataEX(basedir "\BEFUND.dbf", inpattern, outpattern, 0 ,, 1)
	GuiControl, GV:, GVTTip1, % "Formulare: " dbfData.Count()

	Formulare := Object()
	Formulare.Summen := Object()
	Formulare.Summen.Gesamt := 0

	For index, m in dbfData {

		Jahr	:= SubStr(m.Datum, 1, 4)
		Q 	:= SubStr("0" Ceil(SubStr(m.Datum, 5, 2)/3), -1)
		key 	:= Jahr "[" Q "]"
		kzl	:= m.KUERZEL

		If (last_kzl = kzl)
			continue

		last_kzl := kzl

		If !IsObject(Formulare[key])
			Formulare[key] := Object()
		If !Formulare[key].haskey(kzl)
			Formulare[key][kzl] := 1
		else
			Formulare[key][kzl] += 1

		If !IsObject(Formulare.Summen[Jahr])
			Formulare.Summen[Jahr] := Object()
		If !Formulare.Summen[Jahr].haskey(kzl)
			Formulare.Summen[Jahr][kzl] := 1
		else
			Formulare.Summen[Jahr][kzl] += 1

		If !Formulare.Summen.haskey(kzl)
			Formulare.Summen[kzl] := 1
		else
			Formulare.Summen[kzl] += 1

		Formulare.Summen.Gesamt += 1

		If Mod(index, 50) = 0
			GuiControl, GV:, GVTTip2, % "Fortschritt: "  index "/" dbfData.Count()

	}

	FileOpen(admDBPath "\sonstiges\FormularStatistik.JSON", "w", "UTF-8").Write(JSON.DUMP(Formulare,,2, "UTF-8"))
	Run, % admDBPath "\sonstiges\FormularStatistik.JSON"

return
;}

BefundTexte: ;{

	Gui, GV: Submit, NoHide

	If RegExMatch(GVID, "\d+")
		inpattern   	:= {"PATNR":GVID, "KUERZEL":GVKZL}
	else
		inpattern   	:= {"KUERZEL":GVKZL}

	outpattern 	:= ["PATNR", "DATUM", "TEXTDB", "INHALT"]
	ref               	:= GetDBFDataEX(basedir "\BEFUND.dbf", inpattern, outpattern, 0,, 1)

	refListe	:= "", txtbuf 	:= ""

	;FiletxtDB := FileOpen(admDBPath "\sonstiges\" GVKZL ".txt", "w", "UTF-8")

	For mIndex, m in ref {
		If (m.TEXTDB <> 0)
			refListe .= (mIndex > 1 ? "`n" : "" ) m.TEXTDB
		else
			txtbuf .= (mIndex > 1 ? "`n" : "" ) m.Inhalt
	}

	Sort, refListe, N
	FileOpen(admDBPath "\sonstiges\" GVKZL "_TEXTDB.txt", "w", "UTF-8").Write(refListe)
	Run, % admDBPath "\sonstiges\" GVKZL "_TEXTDB.txt"

	return
	txtbuf := RegExReplace(txtbuf, "i)\d+,\d+°*\s*C*\s*,*", " ")
	txtbuf := RegExReplace(txtbuf, "i)BS[GR]\s*\:*\s*\d+\/*\d*\s*,*", " ")
	txtbuf := RegExReplace(txtbuf, "i)P\:\s*\d+[\/\\]*\d*\s*(min)*[,;\s]*", " ")
	txtbuf := RegExReplace(txtbuf, "i)BZ(\-Schnelltest)*\s*\:*\s*\d+\.*\d*\s*(mmol|mg)*[\\\/]*(l|mmol|dl)*.*?(\s|\n)", " ")
	txtbuf := RegExReplace(txtbuf, "i)Gewicht.*?(kg|g)[\s,;]*", " ")
	txtbuf := RegExReplace(txtbuf, "i)Größe\s*\:*\s*\d+[,.]*\d*\s*(cm|m)*[\s,;]*", " ")
	txtbuf := RegExReplace(txtbuf, "[,;]\s*\n", "")
	txtbuf := RegExReplace(txtbuf, "[\s]{2,}", " ")
	Sort, txtbuf, U
	FileOpen(admDBPath "\sonstiges\" GVKZL ".txt", "w", "UTF-8").Write(txtbuf)
	Run, % admDBPath "\sonstiges\" GVKZL ".txt"

	;FiletxtDB.CloseDBF()
ExitApp
return
;}

GetTextDBText(AlbisDBPath, pattern) {

	dbfile 	:= new DBASE(AlbisDBPath "\BEFTEXTE.dbf", 1)
	VarSetCapacity(recordbuf, dbfile.lendataset)
	fpos  	:= dbfile.OpenDBF()
	while !(dbfile.dbf.AtEOF) {

		bytes	:= this.dbf.RawRead(recordbuf, dbfile.lendataset)
		set 	:= StrGet(&recordbuf, dbfile.lendataset, dbfile.encoding)

		PatNR := SubStr(set, 1, 9)
		Datum := SubStr(set, 10, 8)

	}

}

Chronikerziffernfehlt: ;{ 03221 fehlt

	#Persistent
	StartTime1 := A_TickCount
	basedir	:= GetAlbisPath() "\db"
	BEFUNDidx := ReadDBASEIndex(admDBPath, "BEFUND")
	StartTime2 := A_TickCount

	sammler := Object()

	QNow	:= GetQuartalEx("31.12.2020", "YYYY-Q")
	RegExMatch(QNow, "(?<Jahr>\d+).(?<Q>\d)", D)
	QLast1	:= (DQ - 1 = 0) ? (DJahr-1 "-" 4 ) : (DJahr "-" DQ - 1)
	QLast2	:= (DQ - 2 = 0) ? (DJahr-1 "-" 4 ) : (DQ - 2 = -1) ? (DJahr-1 "-" 3 ) :(DJahr "-" DQ - 2)

	SciTEOutput("QN: " BEFUNDidx[QNow] "`nQL1: " BEFUNDidx[QLast1] "`nQL2: " BEFUNDidx[QLast2])


	inpattern   	:= {"KUERZEL":"lko", "DATUM":"rx:20201[012]\d\d"}
	outpattern 	:= ["PATNR", "DATUM", "KUERZEL", "INHALT"]		                                            	; DSVGO = Löschsymbol?
	;GetDBFData(DBASEfilepath, inpattern:="", outpattern:="", seek:=0, options:="", debug:=0)
	ziffern       	:= GetDBFDataEX(basedir "\BEFUND.dbf", inpattern, outpattern, 0 ,, 1)

	For ZIndex, m in ziffern {

		PATID := m.PATNR
		;~ If !IsObject(sammler[PATID])
			;~ sammler[PATID] := Object()

		If !sammler.HasKey(PATID)
			sammler[PATID] := 0

		If InStr(m.inhalt, "03220") {
				sammler[PATID] += 1
		}

		If InStr(m.inhalt, "03221") {
				sammler[PATID] += 2
		}

	}

	FileOpen(admDBPath "\sonstiges\chroniker.JSON", "w", "UTF-8").Write(JSON.DUMP(sammler,,1))

	For PATID, val in sammler {

		LV_Add("", PatID, PatDBF[PATID].Name, PatDBF[PATID].VORNAME, PatDBF[PATID].GEBURT, val,, PatDBF[PATID].BEHTage.Count())

	}

return
;}

GetDBFDataEX(DBASEfilepath, inpattern:="", outpattern:="", seek:=0, options:="", debug:=0) {     	;-- holt Daten aus einer beliebigen Datenbank

	dbfile      	:= new DBASE(DBASEfilepath, debug)

	If (seek > 0)
		startrecord := Floor((seek - (dbfile.headerlen + 1)) / dbfile.lendataset)
	else
		startrecord := 0

	SciTEOutput("starts with record nr: " startrecord)

	res        	:= dbfile.OpenDBF()
	;matches 	:= dbfile.Search(inpattern, 0, options)
	matches 	:= dbfile.Search(inpattern, startrecord, options)
	res         	:= dbfile.CloseDBF()
	dbfile    	:= ""

	SciTEOutput("matches: " matches.Count())
	If !IsObject(outpattern)
		return matches

	data := Array()
	For midx, m in matches {

		strObj:= Object()
		For cidx, ckey in outpattern
			strObj[ckey] := m[ckey]

		data.Push(strObj)

	}


return data
}

VorschlaegeAnzeigen: ;{

	If FileExist(pathGVUListe "\" Abrechnungsquartal "-GVUVorschlag.JSON") {

			GVUVorschlag := JSON.Load(FileOpen(pathGVUListe "\" Abrechnungsquartal "-GVUVorschlag.JSON", "r").Read())
			RemoveAt := Array()
			For PatID, g in GVUVorschlag {

				inList := false
				For PatIDListe, liste in GVUListe
					 If (PatIDListe = PatID+1) {
						RemoveAt.Push(PatID)
						inList := true
						break
					 }

				If !inList
					LV_Add("", PatID+1, g.Name, g.VORNAME, g.GEBURT,, g.LastGVU, g.BEHTage.Count())

			}

			LV_ModifyCol()
			GuiControl, GV:, GVStats, % "zusätzliche Patienten: " LV_GetCount() ", in Liste: " GVUListe.Count() ", Pat. in DB: " PatDBF.Count()
			FileOpen(pathGVUListe "\" Abrechnungsquartal "-GVUVorschlag.JSON", "w", "UTF-8").Write(JSON.DUMP(GVUVorschlag,,1))

		}


return ;}

VorschlaegeBearbeiten: ;{


return
;}

JumpListQuartale: ;{    ;         120090317

		GuiControl, GV:, GVTTip1, % "erstelle Index der BEFUND.DBF"
		GuiControl, GV:, GVTTip2, % ""
		GuiControl, GV:, GVTTip3, % ""

		If !InStr(FileExist(admDBPath "\sonstiges"), "D")
			FileCreateDir, % admDBPath "\sonstiges"

		Q := Object()
		JLBefundDBF := Object()
		pt := ""
		prg1 := prg2 := 0

		file := new DBASE(basedir "\BEFUND.dbf", 4)
		VarSetCapacity(recordbuf, file.lendataset, 0x20)
		file.OpenDBF()                                                	     ; ist immer nur ein Lesezugriff!!

		while (!file.dbf.AtEOF) {

			bytes	:= file.dbf.RawRead(recordbuf, file.lendataset)
			set 	:= StrGet(&recordbuf, file.lendataset, "CP1252")
			removedSet := Trim(SubStr(set, -2)) = "*" ? true : false

			Jahr   	:= SubStr(set, 10, 4)
			Q      	:= Ceil(SubStr(set, 14, 2)/3)

			filepos :=
			prg1 := Round(file.dbf.Tell() / file.dbf.Length() * 90)
			If (prg1 > prg2) {
				pt .= "."
				prg2 := prg1
				GuiControl, GV:, GVTTip2, % pt
				GuiControl, GV:, GVTTip3, % Round(file.dbf.Tell() / file.dbf.Length() * 100) "%"
			}

			If (QLast < Q) || (QLast =4 && Q = 1) || (LJahr < Jahr) {
				If !JLBefundDBF.haskey(Jahr "-" Q)
					JLBefundDBF[Jahr "-" Q] := file.dbf.Tell()
				QLast := Q
				LJahr := Jahr
			} else if (QLast = "") && (Q <> "") {
				JLBefundDBF[Jahr "-" Q] := file.dbf.Tell()
				QLast := Q
			}

			;SciTEOutput(set)
			;SciTEOutput(data.PatNR ": " data.Datum " - " data.DateTime "`t" PatID ": "   Jahr "  " Monat "  rSet: " (removedSet ? "true":"false") "`n")
			;SciTEOutput(data.PatNR ": " data.Datum " - " data.DateTime "  rSet: " (removedSet ? "true":"false") "`n")

			;~ If A_Index = 10
				;~ break

		}

		file.dbf.CloseDBF()
		FileOpen(admDBPath "\sonstiges\BEFUNDdbf.idx", "w", "UTF-8").Write(JSON.DUMP(JLBefundDBF))

return
;}

ZiffernSammler: ;{


		PatExtraTmp		:= Object()
		WarDa 				:= Object()
			PJahr := 0


			;~ inpattern   	:= {"KUERZEL":"lko", "DATUM":"rx:20[12][0-9]\d{4}"}
			;~ GuiControl, GV:, GVTTip, % "sammle Leistungskomplexe für Abrechnungstage`n...\PatData\Ziffern.json"
			;~ ziffern       	:= GetDBFData(basedir "\BEFUND.dbf", inpattern, "",, 1)


			;~ GuiControl, GV:, GVTTip, % "Quartalszählung erstellt "
			;~ FileOpen(admDBPath "\PatData\PatWarDa.json", "w", "UTF-8").Write(JSON.DUMP(WarDa,,1))
			;~ Run, % admDBPath "\PatData\PatWarDa.json"

		Gui, GV: Submit, NoHide
		;GVQuartal

		Abr           	:= {	"01740":"1x im Arztfall"
								, 	"01732":"alle 3 Jahre"
								, 	"01746":"alle 3 Jahre"
								, 	"03360":"2x im Behandlungsfall"
								, 	"03362":"1x im Quartal,nicht neben 03360"
								, 	"03220":"Chroniker"
								,	"03221":"wenn 03220 abgerechnet"}

		GuiControl, GV:, GVTTip1, % "sammle Leistungskomplexe.....\PatData\Ziffern.json"

		inpattern   	:= {"KUERZEL":"lko", "DATUM":"rx:20[1-2][0-9]\d{4}"}						; jedes Datum ab 2010 wird untersucht!!
		outpattern 	:= ["DATUM", "PATNR", "INHALT"]		                                            	; DSVGO = Löschsymbol?
		ziffern       	:= GetDBFData(basedir "\BEFUND.dbf", inpattern, outpattern,, 1)


		; sucht nur nach bestimmten Ziffern
		For ZIndex, m in ziffern {

			If !IsObject(PatExtraTmp[(PatID := m.PATNR)])
				PatExtraTmp[PatID] := Object()

			;PatExtraTmp[PatID].Chroniker := PatDBF[PAtID].Chroniker

			If m.removed
				ZZ := "*Z"
			else
				ZZ := "Z"

			For lidx, LZiffer in StrSplit(m.Inhalt, "-")
				If RegExMatch(LZiffer, "^\d+")
					For AbrZiffer, AbrRegel in Abr
						If RegExMatch(LZiffer, "^" AbrZiffer) {
							PatExtraTmp[PatID][ZZ RTrim(LZiffer, "H")] := m.Datum
							break
						}

			If (Mod(ZIndex, 10) = 0)
				GuiControl, GV:, GVTTip2, % "ZIndex: " ZIndex "/" Ziffern.Count()

			}

		PatExtra := Object()
		For PatExID, ZExtra in PatExtraTmp
			If (ZExtra.Count() > 0) {
				PatExtra[PatExID] := ZExtra
			}

		FileOpen(admDBPath "\PatData\PatExtra.json", "w", "UTF-8").Write(JSON.DUMP(PatExtra,,1))

		GuiControl, GV:, GVTTip1, % "Patienten mit Daten:  `t" PatExtra.Count()
		GuiControl, GV:, GVTTip2, % "Patienten ohne Daten:`t" PatExtraTmp.Count() - PatExtra.Count()
		GuiControl, GV:, GVTTip3, % "                          ---- FERTIG ----"

				;zählt ob ein Patient im Quartal da war
		For ZIndex, m in ziffern {
				; wann wurde bei einem Paitenten etwas abgerechnet
				Jahr := SubStr(m.Datum, 1, 4)
				WDJahr := "J" SubStr(m.Datum, 3, 2)
				WDQ	:=  "Q" Ceil(SubStr(m.Datum, 5, 2)/3)
				PatID := m.PatNR
				If !IsObject(WarDa[PatID])
					WarDa[PatID] :=Object()

				 If !IsObject(WarDa[PAtID][WDJahr])
					WarDa[PAtID][WDJahr] := Array()

				 If !WarDa[PAtID][WDJahr].haskey(WDQ)
					WarDa[PAtID][WDJahr][WDQ] := 1
				else
					WarDa[PAtID][WDJahr][WDQ] += 1

				zaehler ++
				PJahr := Jahr > PJahr ? JAhr : PJahr
				IF Mod(zaehler, 1000) = 0
					GuiControl, GV:, GVTTip, % "Quartalszählung:  Jahr: " PJahr " - lko zeile: " zaehler

			}



return ;}

				;~ GuiControl, GV:, GVTTip, % "Patienten: " PatExtra.Count() "     Ziffern durchsucht: " ziffernsumme "`nPatienten ohne Daten: " PatExtra.Count() - PatExtraFinally.Count()

GelöschteZiffern: ;{


		PatExtraTmp		:= Object()
		WarDa 				:= Object()
			PJahr := 0

		Gui, GV: Submit, NoHide
		;GVQuartal

		GuiControl, GV:, GVTTip1, % "sammle gelöschte Leistungskomplexe....."

		inpattern   	:= {"KUERZEL":"lko", "DATUM":"rx:(2016|2018|2019|2020)\d{4}"}						; jedes Datum ab 2010 wird untersucht!!
		outpattern 	:= ["DATUM", "PATNR", "INHALT"]		                                            	; DSVGO = Löschsymbol?
		ziffern       	:= GetDBFData(basedir "\BEFUND.dbf", inpattern, outpattern,, 1)


		; sucht nur nach bestimmten Ziffern
		For ZIndex, m in ziffern {

			;SciTEOutput(ZIndex ": " m.removed)

			If m.removed {
				PatID := m.PATNR
				If !IsObject(PatExtraTmp[PatID])
					PatExtraTmp[PatID] := Object()
				If PatExtraTmp[PatID].haskey(m.Datum)
					PatExtraTmp[PatID][m.Datum] .= " | " m.Inhalt
				else
					PatExtraTmp[PatID][m.Datum] := m.Inhalt
			}

			If (Mod(ZIndex, 10) = 0)
				GuiControl, GV:, GVTTip2, % "ZIndex: " ZIndex "/" Ziffern.Count()

			}

		PatExtra := Object()
		For PatExID, ZExtra in PatExtraTmp
			If (ZExtra.Count() > 0)
				PatExtra[PatExID] := ZExtra


		FileOpen(admDBPath "\PatData\ziffern.json", "w", "UTF-8").Write(JSON.DUMP(ziffern,,1))
		FileOpen(admDBPath "\PatData\ErasedPatData.json", "w", "UTF-8").Write(JSON.DUMP(PatExtra,,1))

		GuiControl, GV:, GVTTip1, % "Patienten mit Daten:  `t" PatExtra.Count()
		GuiControl, GV:, GVTTip2, % "Patienten ohne Daten:`t" PatExtraTmp.Count() - PatExtra.Count()
		GuiControl, GV:, GVTTip3, % "                          ---- FERTIG ----"



return ;}

GVUVorschlaege: ;{

	; Arbeitsobjekte
		GVUVorschlag := Object()

	; vorhandene Untersuchungen laden
		GVUListe := ReadGVUListe(pathGVUListe, Abrechnungsquartal)
		SciTEOutput("GVUListe erfasst mit " GVUListe.Count() " Untersuchungen")

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Patienten und das Datum all ihrer Vorsorgeuntersuchungen heraussuchen
	;-------------------------------------------------------------------------------------------------------------------------------------------;{
		inpattern   	:= {"KUERZEL":"lko", "INHALT":"01732"}
		outpattern 	:= {"DATUM":"rx:(20|19)0[0-9]\d{4}"}                                       	; RegEx - alle Datumsformate vor 2010 sind ausgeschlossen
		GVUPatients	:= GetDBFData(basedir "\BEFUND.dbf", inpattern, outpattern,,1)
		SciTEOutput("Anzahl abgerechneter Vorsorgeuntersuchungen: " GVUPatients.Count())
		outlist := Array()

		For idx, m in GVUPatients
			PatDBF[m.PatNR].LastGVU := m.Datum

		For idx, m in GVUPatients {

			PatID := m.PatNR

			LastBEHM	:= SubStr(PatDBF[PatID].LAST_BEH, 5, 2)                                 	; Monat
			LastBEHJ	:= SubStr(PatDBF[PatID].LAST_BEH, 1, 4)                                 	; Jahr
			LastBEHQ	:= SubStr("0" . Ceil(LastBEHM/3), -1) . SubStr(LastBEHJ, 3,2)      	; Quartal z.B. 0320
			gvuspan 	:= HowLong(PatDBF[PatID].LastGVU, AbrQ.DBaseEnd)
			PatAge 		:= HowLong(PatDBF[PatID].GEBURT, AbrQ.DBaseEnd)

		; Filter - verstorben, kein Behandlung in diesem Quartal oder bereits gelistet
			If (StrLen(PatDBF[PatID].MORTAL) > 0) || (LastBEHQ <> Abrechnungsquartal) || GVUListe.haskey(PatID) || (gvuspan.years < minAbstand) || (PatAge.years < minAlter) {
				outlist.Push(PatID)
				continue
			}

			; Pat aufnehmen
			GVUVorschlag[PatID] := {	"LastGVU" 	: ConvertDBaseDate(PatDBF[PatID].LastGVU)
												,	"LastBEH"   	: SubStr(PatDBF[PatID].LAST_BEH, 7, 2) "." LastBEHM "." LastBEHJ
												,	"BEHTage" 	: []
												, 	"Name"     	: PatDBF[PatID].NAME
												, 	"Vorname" 	: PatDBF[PatID].VORNAME
												, 	"Geburt"    	: PatDBF[PatID].GEBURT
												,	"Alter"			: PatAge.years}

			; debugging

			;SciTEOutput(m.PatNR "; " ConvertDBaseDate(m.Datum) "; " LastBEH  "  (" PatDBF[PatID].NAME ", " PatDBF[PatID].VORNAME ", Alter: " timespan.years  ")")

		}

		SciTEOutput("Anzahl weiterer GVU: " (gvuc := GVUVorschlag.Count()) "`n")

		;FileOpen(A_ScriptDir "\GVUMAtches.JSON", "w", "UTF-8").Write(JSON.DUMP(GVUMatches,,2))
	;}

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; unerfasste Patienten in GVUListe und GVUMatches mit angelegtem Abrechnungsschein in abrechnungsfähigem Alter hinzufügen
	;------------------------------------------------------------------------------------------------------------------------------------------- ;{
		rxDatumPattern	:= AbrQ.Jahr "(" AbrQ.MBeginn "|" SubStr("0" (AbrQ.MBeginn+1), -1) "|" AbrQ.MEnde ")\d{2}" ; RegEx Einschluss Abrechnungsscheine
		inpattern   	:= {"KUERZEL":"lko", "DATUM":"rx:" rxDatumPattern}
		outpattern 	:= ""
		AbrKomplexe:= GetDBFData(basedir "\BEFUND.dbf", inpattern, outpattern,,1)

		ohneGVU := 0
		For idx, lko in AbrKomplexe {

			PatID := lko.PatNR

			if GVUVorschlag.haskey(PatID) {

				foundBEH := false
				For behIDX, lDatum in GVUVorschlag[PatID].BEHTage
					if (lDatum = lko.Datum) {
						foundBEH := true
						break
					}
				If !foundBEH
					GVUVorschlag[PatID].BEHTage.Push(lko.Datum)

			}

		}

		FileOpen(pathGVUListe "\" Abrechnungsquartal "-GVUVorschlag.JSON", "w", "UTF-8").Write(JSON.DUMP(GVUVorschlag,,1))
		SciTEOutput(pathGVUListe "\" Abrechnungsquartal "-GVUVorschlag.JSON - gespeichert" "`n" )

	;}

return ;}


;-------------------------------------------------------------------------------------------------------------------------------------------
; FUNKTIONEN
;-------------------------------------------------------------------------------------------------------------------------------------------;{




;}

;-------------------------------------------------------------------------------------------------------------------------------------------
; INCLUDES
;-------------------------------------------------------------------------------------------------------------------------------------------;{
#Include %A_ScriptDir%\..\..\Include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\Include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\Include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\Include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\Include\Addendum_Datum.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\Include\Addendum_PdfHelper.ahk
#Include %A_ScriptDir%\..\..\Include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\Include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PraxTT.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\Include\Addendum_Window.ahk

#Include %A_ScriptDir%\..\..\lib\ACC.ahk
#Include %A_ScriptDir%\..\..\lib\class_JSON.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#Include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\lib\Sift.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
;}



/*
If !IsObject(PatExtra[PatID]) {
				PatExtra[PatID] := Object()
				;~ PatExtra[PatID].NAME                	:= PatDBF[PatID].NAME
				;~ PatExtra[PatID].VORNAME         	:= PatDBF[PatID].VORNAME
				;~ PatExtra[PatID].GEBURT             	:= PatDBF[PatID].GEBURT
				;~ PatExtra[PatID].LAST_BEH            	:= PatDBF[PatID].LAST_BEH
				PatExtra[PatID][Jahr]                  	:= Object()
				PatExtra[PatID][Jahr][Quartal]     	:= Object()
				PatExtra[PatID][Jahr][Quartal]["lko"]:= Object()
			}
			else If !IsObject(PatExtra[PatID][Jahr]) {
				PatExtra[PatID][Jahr]                  	:= Object()
				PatExtra[PatID][Jahr][Quartal]     	:= Object()
				PatExtra[PatID][Jahr][Quartal]["lko"]	:= Object()
			}
			else If !IsObject(PatExtra[PatID][Jahr][Quartal]) {
				PatExtra[PatID][Jahr][Quartal]     	:= Object()
				PatExtra[PatID][Jahr][Quartal]["lko"]	:= Object()
				PatExtra[PatID][Jahr][Quartal]["lko"][MTag] := Array()
			}
			else If !IsObject(PatExtra[PatID][Jahr][Quartal]["lko"][MTag]) {
				PatExtra[PatID][Jahr][Quartal]["lko"][MTag] := Array()
			}



 */
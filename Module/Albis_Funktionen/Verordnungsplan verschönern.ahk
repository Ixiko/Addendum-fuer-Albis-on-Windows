; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . .
; . . . . . .                                                                  	VERORDNUNGSPLAN VERSCHÖNERN
global                                                         						DatumVom:= "29.01.2021"
; . . . . . .
; . . . . . .                      ROBOTIC PROCESS AUTOMATION FOR THE GERMAN MEDICAL SOFTWARE "ALBIS ON WINDOWS"
; . . . . . .                             BY IXIKO STARTED IN SEPTEMBER 2017 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE
; . . . . . .                          THIS SCRIPT USES THE - MICROSOFT COM OBJECT MODEL - TO INTERACT WITH MS WORD
; . . . . . .                                                   THIS SOFTWARE IS ONLY AVAIBLE IN GERMAN LANGUAGE
; . . . . . .
; . . . . . .      - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; . . . . . .      !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !!
; . . . . . .      - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; . . . . . .                                      SCRIPT ONLY WORKS WITH AUTOHOTKEY OR AUTOHOTKEY_H - V1.1.32.XX
; . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .


/* VBA-SKRIPT DIREKT ÜBER MICROSOFT WORD STARTEN

	so kann man das AHK Skript direkt in Word starten wenn man diesen VB-Code als Makro in die Normal.dot kopiert.
	!Skriptpfad anpassen!

	Sub Medplan()

		Dim cmdln As String
		Dim retVal

		cmdln = "C:\Program Files\AutoHotkey\Autohotkey.exe /f ""D:\....\Verordnungsplan verschönern.ahk"""
		retVal = Shell(cmdln, vbNormalNoFocus)

	End Sub

 */

;----------------------------------------------------------------------------------------------------------------------------------------------
; Skripteinstellungen
;----------------------------------------------------------------------------------------------------------------------------------------------;{
	#NoEnv
	SetBatchLines, -1
;}

;----------------------------------------------------------------------------------------------------------------------------------------------
; Variablen für die Ersetzungen
;----------------------------------------------------------------------------------------------------------------------------------------------;{

	RegRead, AlbisWinDir, HKLM, SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
	RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)

	ifapDBPath	:= AlbisWinDir "\ifapDB\ibonus3\Datenbank"
	userMedDB	:= AddendumDir "\include\Daten\medkategorien.json"

	If FileExist(userMedDB)
		medis := JSON.Load(FileOpen(userMedDB, "r").Read())
	else
		medis:=Object()

	medisCount	 := medis.Count()   ; Bestand der bekannten Medikamentennamen
	rowsToDelete := Array()

	RemoveThatShit := {	"Aaa"                     	: ""
								, 	"\-*1\s*A"              	: ""
								, 	"Ab[Zz]"                 	: ""
								, 	"Acino"                  	: ""
								, 	"Actavis"                  	: ""
								, 	"AL"                       	: ""
								, 	"Ari"                      	: ""
								, 	"Arzneim"              	: ""
								, 	"Aristo"                  	: ""
								, 	"Allomedic"            	: ""
								,	"Augentropfen"      	: ""
								, 	"Aurobindo"           	: ""
								, 	"Arzne"                  	: ""
								, 	"AWD"                   	: ""
								,	"Axico"                  	: ""
								,	"Axicorp\s*P"          	: ""
								, 	"BAYER|Bayer"       	: ""
								, 	"Becat"                  	: ""
								, 	"beta"                    	: ""
                            	,	"Br\W"                   	: ""
								,	"\s*\-\s+Ct(\d+)"     	: " $1"
								,	"\-*\s*CT"              	: ""
								,	"Deutsch"              	: ""
								,	"Dosiergel"            	: ""
								,	"Dosieraeros"        	: ""
								,	"ED"                      	: ""
								,	"Emra.Med"           	: ""
								,	"Eurimpharm"        	: ""
								,	"Fair.Med"             	: ""
								,	"Fertigspritze"        	: ""
								,	"Filmtabl*e*t*t*e*n*"	: ""
								,	"GALEN"                	: ""
								,	"Glenmark"           	: ""
								,	"GmbH*"               	: ""
								,	"Hartka"                	: ""
								,	"Henning"              	: ""
								,	"Hennig"                	: ""
								,	"Heum*a*n*n*"         	: ""
								,	"HEXAL"                 	: ""
								,	"Hkp"                    	: ""
								,	"Huebe"                	: ""
								,	"Inject"                     	: ""
								,	"Inte"                     	: ""
								,	"Isis"                      	: ""
								,	"kohlpharma*"       	: ""
								,	"KwikPen"              	: ""
								,	"Lomapharm"         	: ""
								,	"Lichten(\d+)"           	: " $1"
								,	"Mili"                     	: ""
								,	"Milinda"               	: ""
								,	"Msr"                     	: ""
								,	"Myl"                     	: ""
								,	"Nachfuell"            	: ""
								,	"Net"                     	: ""
								,	"Novolet"               	: ""
								,	"Orifarm"              	: ""
								,	"Pe\s"                    	: ""
								,	"Pen\s"                  	: ""
								,	"Penf*i*l*l*"               	: ""
								,	"Phar*m*a*"          	: ""
								,	"Protect(\d+)"        	: " $1"
								,	"[Rr]atiop*h*a*r*m*"  	: ""
								,	"Retard"                 	: ""
								,	"Retardtable*t*t*e*n*"	: ""
								,	"SANDOZ"               	: ""
								,	"Sta(\d)"                 	: " $1"
								,	"STADA|Stada"      	: ""
								,	"\ssto"                   	: ""
								,	"Tabl\W"                  	: ""
								,	"Tabletten"             	: ""
								,	"T[Aa][Hh]"            	: ""
								,	"TEI"                      	: ""
								,	"TEVA"                   	: ""
								,	"Tro"                     	: ""
								,	"Vital"                    	: ""
								,	"Weichkaps*e*l*n*"	: ""
								,	"Winthrop"	            	: ""
								,	"Zentiva"               	: ""
								,	"4Wochen"            	: "" }

	MedForm	:= {	"AMP"	: "Ampulle"
						, 	"ATR" 	: "Augentropfen"
						, 	"CRE"	: "Creme"
						, 	"DOS" 	: "Dosierspray"
						, 	"EMU"	: "Emulsion"
						, 	"FER" 	: "Fertigspritze"
						, 	"FLU"	: "Flüssig"
						, 	"FTA" 	: "Tablette"
						, 	"GEL"	: "Gel"
						, 	"GRA"	: "Granulat"
						, 	"HKM"	: "Kapsel"
						, 	"HKP"	: "Kapsel"
						, 	"HPI" 	: "Kapsel"
						, 	"HVW"	: "Kapsel"
						, 	"IFB"  	: "Infusionsbeutel"
						, 	"IHP"  	: "Inhalationspack"
						, 	"INF" 	: "Infusionslösung"
						, 	"ILO" 	: "Injektionslösung"
						, 	"ISU" 	: "Suspension"
						, 	"KAP"	: "Kapsel"
						, 	"KMP"	: "Kapsel"
						, 	"KMR"	: "Kapsel"
						, 	"KPG"	: "Kombipackung"
						, 	"LOT"	: "Lotion"
						, 	"LSE" 	: "Lösung"
						, 	"MIL" 	: "Lotion"
						, 	"PEN"	: "Fertigpen"
						, 	"PFL" 	: "Pflaster"
						, 	"PIK"  	: "Infusionslösung"
						, 	"PLE" 	: "Pulverzubereitung"
						, 	"PST" 	: "Paste"
						, 	"PUL"	: "Pulver"
						, 	"REK" 	: "Retardkapsel"
						, 	"RET" 	: "Retardtablette"
						, 	"SHA"	: "Shampoo"
						, 	"SIR"  	: "Sirup"
						, 	"SPR" 	: "Spray"
						, 	"STI"  	: "Stift"
						, 	"TAB"	: "Tablette"
						, 	"TEI"  	: "Tropfen"
						, 	"TMR"	: "Tablette"
						, 	"TRO"	: "Tropfen"
						, 	"UTA"	: "Tablette"
						, 	"VGE"	: "Vaginalcreme"
						, 	"WKA"	: "Kapsel"
						, 	"XNC"	: "Nachtcreme"
						, 	"ZAM"	: "Ampulle"}

	Hinweise	:= {1: 	"• Einnahme der Kapseln mit Flüssigkeit zur selben Tageszeit, unabhängig von den Mahlzeiten "
							.	"• Kapseln nicht zerkauen oder zerstoßen • nicht zusammen mit Grapefruitsaft einnehmen"}

;}

;----------------------------------------------------------------------------------------------------------------------------------------------
; COM Verbindung zu MS WORD herstellen
;----------------------------------------------------------------------------------------------------------------------------------------------;{
	if !IsObject(oWord) {
			try
				oWord := ComObjActive("Word.Application")
			catch {
				;MsgBox, 4, Verordnungsplan, % "Wollen Sie das Trainingscenter starten?"
				;IfMsgBox, Yes
					gosub, Training

				ExitApp
			}
	}

;}

;----------------------------------------------------------------------------------------------------------------------------------------------
; Medikationsplan/Verordnungsplan - geöffnetes Dokument suchen
;----------------------------------------------------------------------------------------------------------------------------------------------;{
	docCount	:= oWord.Documents.Count
	Loop, % docCount {

		Content := oWord.Documents.Item(A_Index).Content.Text
		If RegExMatch(Content, "i)V.e.r.o.r.d.n.u.n.g.s.p.l.a.n|Medikationsplan") {

			docname	:= oWord.Documents.Item(A_Index).Name
			vopln    	:= oWord.Documents.Item(A_Index)

			oWord.Documents(A_Index).Activate
			docNr:= A_Index
			break
		}

	}

;}

;----------------------------------------------------------------------------------------------------------------------------------------------
; Medikamentenallergien oder -nebenwirkungen erkennen und gesondert in der Tablette positionieren
;----------------------------------------------------------------------------------------------------------------------------------------------;{

	; DIE EINTRÄGE ZU ALLERGIEN KÖNNEN SICH ÜBER MEHRERE ZEILEN ERSTRECKEN UND
	; WERDEN JETZT ZU EINER EINZIGEN ZEILE ZUSAMMENGEFASST. ES WIRD EINE ANDERE SCHRIFTART GEWÄHLT
	; UND DIE HINTERGRUNDFARBE DER ZELLE WIRD DUNKLER GEWÄHLT
		startrow:= 1
		table		:= oWord.ActiveDocument.Tables(1)
		rows		:= table.Rows.Count
		cols  	:= table.columns.Count
		cell   	:= table.Cell(2, 2)
		cell.Select

		Text  	:= cell.Range.Text
		Text  	:= SubStr(Text, 1, StrLen(Text) - 2)
		Text  	:= Trim(Text)

		If RegExMatch(Text, "##|Problem|Allergie") {

			cell.Range.Font.Size       	:= 11
			cell.Range.Font.Bold      	:= true
			cell.Range.Font.italic      	:= true
			cell.Range.Font.underline	:= true
			cell.Range.Text              	:= "!!! Bekannte Allergie oder Nebenwirkungen in der Anamnese bei folgenden Medikamenten: "

			Loop % (rows - 2) {                                                    	; Ende der Hinweise finden
				rowAllergieLast := A_Index + 2
				Text := table.Cell(2+A_Index, 2).Range.Text
				If InStr(Text, "##")
					break
			}

		; FELD LINKS UND ALLE FELDER RECHTS DER MEDIKAMENTENBEZEICHNUNG LEEREN, ALLE ZELLEN EINER ZEILE VERBINDEN ;{
			Loop % (rowAllergieLast - 2) {

				row := A_Index + 1
				table.Cell(row, 1).Range.Text := ""
				Loop 6
					table.Cell(row, 4 + A_Index).Range.Text := ""

				table.Rows(row).Cells.Merge
				;table.Cell(row, 2).Merge(table.Cell(row, cols))  ; vorbereitung für anderes Layout
				thisrow := table.Rows(row)
				thisrow.Range.Font.Size := 10

			}
		;}

		; DIE MEDIKAMENTE MÜSSEN NEU NUMMERIERT WERDEN
			Loop % (rows - rowAllergieLast - 2)
				table.Cell(rowAllergieLast + A_Index, 1).Range.Text := A_Index

		; LÖSCHT DIE ZEILE DIE DAS ENDE DER HINWEISE MARKIERTE
			table.Rows(rowAllergieLast).Delete

		; ALLE HINWEISZEILEN WERDEN ZU EINER EINZIGEN ZELLE VERBUNDEN, DIESE ZELLE ERHÄLT EINEN ETWAS DUNKLEREN HINTERGRUND
			table.Cell(2,1).Merge(table.Cell(rowAllergieLast-1,1))
			table.Cell(2,1).Shading.BackgroundPatternColor := -603930625

		; HINWEISBEREICH ALS ERSTE ZEILE DER TABELLE EINSTELLEN (DURCH AUSSCHNEIDEN UND EINFÜGEN DER 1.ZEILE NACH DEN HINWEISEN)
			table.Rows(1).Range.Cut
			table.Rows(2).Range.Paste
			;table.CellRows(1).Range.Cut
			;table.Rows(2).Range.Paste

		; STARTZEILE FÜR DIE ERSETZUNG IST JETZT DIE 3.ZEILE
			startrow := 2

		}


;}

;----------------------------------------------------------------------------------------------------------------------------------------------
; Medikamentbezeichnungen bereinigen von Pharmaherstellern, Packungsgrößen, sonstigen Angaben
;----------------------------------------------------------------------------------------------------------------------------------------------;{
	Loop % ((rows := table.Rows.Count) - startrow) {

		; SPALTE 2 AB ZEILE 2 - MEDIKAMENTBEZEICHNUNG - KÜRZEN UND KLARER DARSTELLEN
			thisrow	:= startrow + A_Index                        	; startrow kann 2 oder 3 sein !
			Cell   	:= table.Cell(thisrow, 2)
			Cell.Select
			Text   	:= Cell.Range.Text
			Text  	:= SubStr(Text, 1, StrLen(Text) - 2)
			Text  	:= Trim(Text)

		; LEERE ZEILE MERKEN - WIRD SPÄTER ENTFERNT
			If (StrLen(Text) = 0) {
				rowsToDelete.InsertAt(1, thisrow)
				continue
			}

		; UNNÖTIGES ENTFERNEN
			newText := RegExReplace(Text . " "	, "\sN[123]\s+"                  	, " ")                                          	; N1 N2 N3 entfernen
			newText := RegExReplace(newText , "i)\s\d+x*\d*\s*(St|Huebe)"	, " ")                                          	; Stückzahl entfern
			newText := RegExReplace(newText , "i)\s[\dx]+\s*ml\s*"           	, " ")                                          	; ml Mengenangabe entfernen
			newText := RegExReplace(newText , "^\!")                                                                                  	; ! - am Anfang entfernen
			newText := RegExReplace(newText , "\(.*\)")                                                                                	; Klammern () und Inhalt
			For shit, removewith in RemoveThatShit                                                                                      	; entfernt Herstellernamen
				newText := RegExReplace(newText , "\s" shit "\s*", (!removewith ? " " : removewith))

			newText .= " "

		; ERMITTELT DIE APPLIKATIONSFORM (AUS RET -> RETARDTABLETTE Z.B.), ENTFERNT DIE KURZBEZEICHNUNGEN (RET, TAB)
			For Abkuerzung, Form in MedForm
				If RegExMatch(newText, "i)\s" Abkuerzung "\s") {
					Loop 2	                                                                                                                              	; sind manchmal 2x vorhanden
						newText := RegExReplace(newText , "i)\s" Abkuerzung "\s", " ")                                           	; z.B. Etoricoxib Heu 90mg Fta FTA N3 100 St
					break
				}

			newText .= " "

		; FÜGT EIN LEERZEICHEN ZWISCHEN DOSIS UND EINHEIT EIN
			If RegExMatch(newText, "\s(?<Dosis>[\d\.]+)\s*(?<Einheit>(mg)|(I\.*E\.*)|(µg))\s", Med )
				newText := RegExReplace(newText, "\s([\d\.]+)\s*((mg)|(I\.*E\.*)|(µg))\s", " $1 $2 ")

			newText := RegExReplace(newText , "\s{2,}", " ")                                                                            	; entfernt überflüssige Leerzeichen
			Cell.Range.Text := newText                                                                                                        	; fertig - Medikamentenbezeichnung auffrischen

		; GRUND DER EINNAHME / MEDIKAMENTE KATEGORISIEREN
			RegExMatch(newText, "^[A-Z][\s-]*[A-Za-z\/\-]+\s*([A-Za-z\/\-]\s)*([A-Z]\d+\s)*\s*[A-Za-z\/\-]*", Medikament)    	; erfasst damit auch Medikamente wie Berlinsulin H Basal z.B.
			Medikament := Trim(Medikament)

			If !medis.haskey(Medikament)
				medis[Medikament] := Object()

			If (!medis[Medikament].use || !medis[Medikament].ingredient) {
				result := AbfrageGui(Medikament, medis[Medikament])
				medis[Medikament].use         	:= result.use
				medis[Medikament].ingredient	:= result.ingredient
			}

		; APPLIKATIONSFORM EINTRAGEN (else ist für eine formatierte Ausgabe)
			If !RegExMatch(Form, "\(")
				table.Cell(thisrow, 3).Range.Text := Form
			else {
				RegExMatch(Form, "(.*)\s(\(.*\))", FM)
				FontSize := table.Cell(thisrow,3).Range.Font.Size
				table.Cell(thisrow, 3).Range.Text := FM1 Chr(10) FM2
				;SciTEOutput(StrLen(FM1) ", " StrLen(FM2))
				;table.Cell(thisrow, 3).Range(StrLen(FM1), StrLen(FM2)).Font.Size	:= FontSize - 4

			}

		; MEDIKAMENT ANWENDUNG UND WIRKSTOFFE EINTRAGEN
			table.Cell(thisrow,   4).Range.Text := medis[Medikament].ingredient
			table.Cell(thisrow, 10).Range.Text := medis[Medikament].use
			Form := ""

	}

;}

;----------------------------------------------------------------------------------------------------------------------------------------------
; Kategorisierung speichern
;----------------------------------------------------------------------------------------------------------------------------------------------;{
	If (medis.Count() > medisCount)                  ; nur speichern wenn neue Daten hinzugekommen sind
		writeJSONFile(userMedDB, medis, 4)
;}

;----------------------------------------------------------------------------------------------------------------------------------------------
; leere Tabellenzeilen entfernen
;----------------------------------------------------------------------------------------------------------------------------------------------;{
	For idx, row in rowsToDelete
		table.Rows(row).Delete
;}

;----------------------------------------------------------------------------------------------------------------------------------------------
; alle Zellen die eine einzelne Null enthalten leeren, Zellen verbinden die Text enthalten
;----------------------------------------------------------------------------------------------------------------------------------------------;{
	Loop % (table.Rows.Count - startrow) {

		row                       	:= A_Index + startrow
		MergeRows           	:= false
		MergeRowsOverride	:= false

		Loop 4 {

			; ZELLENTEXT AUSLESEN
				column := 4+A_Index
				table.Cell(row, column).select
				Text := table.Cell(row, 4+A_Index).Range.Text
				If RegExMatch(Text, "^###") {                                              	; Kommentarzeilen werden nicht verbunden
					MergeRowsOverride := true
					;break
				}

			; SYMBOLE FÜR 1/2, 1/4 oder 3/4
				Text1 := " " Text " "
				Text1 := RegExReplace(Text1, "\s1/2\s", "½")
				Text1 := RegExReplace(Text1, "\s1/4\s", "¼")
				Text1 := RegExReplace(Text1, "\s3/4\s", "¾")
				Text1 := RegExReplace(Text1, "(\d+)(IE|Hub)", "$1 $2")                 ; LEERZEICHEN ZWISCHEN DOSIS UND EINHEIT
				Text1 := Trim(Text1)
				If (Text <> Text1)
					table.Cell(row, 4+A_Index).Range.Text := (Text := Text1)

			; LÖSCHEN VON 0 oder A (steht für aut idem)
				If RegExMatch(Text, "^\s*0|A\s*$")
					table.Cell(row, 4+A_Index).Range.Text := ""
				else if RegExMatch(Text, "[A-Za-zß\.;\:-]") && !RegExMatch(Text, "\d+\s*IE|Hub") || RegExMatch(Text, "i)Bedarf"){
					MergeRows := true
					;break
				}

		}

		If MergeRows && !MergeRowsOverride {

			table.Cell(row, 5).Merge(table.Cell(row, 8))
			Text := RegExReplace(table.Cell(row, 5).Range.Text, "\n|\r", " ")
			Text := Text " "
			Text := RegExReplace(Text, "([A-Z][a-z]+)\s+([a-z]+)", "$1 $2")
			Text := RegExReplace(Text, "\sBed\.\s", " Bedarf ")
			Text := RegExReplace(Text, "Bei\s", "bei ")
			Text := RegExReplace(Text, "Bis\s", "bis ")
			Text := RegExReplace(Text, "T\.\s", "Tabl. ")
			Text := RegExReplace(Text, "z\.\s", "zur ")
			Text := RegExReplace(Text, "Nac\s", "Nacht")
			If RegExMatch(Text " ", "\s([Wochen\s]+)\s", WS) {
				Text	:= StrReplace(Text, WS1, RegExReplace(Trim(WS1), "\s+", ""))
			}
			table.Cell(row, 5).Range.Text := SubStr(Text, 1, StrLen(Text)-1)

		}

	}

	; CARET IN DIE LETZE BEARBEITETE ZELLEN SETZEN
		table.Cell(row,1).select

;}

ExitApp

;----------------------------------------------------------------------------------------------------------------------------------------------
; Interne Funktionen
;----------------------------------------------------------------------------------------------------------------------------------------------;{
GetCellText(tableObj, row, col) {

		try
			Text := table.Cell(row, col).Range.Text
		catch
		{
			SciTEOutput(row ", " A_Index)
			return "###"
		}

		If Select
			Cell.Select

		Text := SubStr(Text, 1, StrLen(Text) - 1)

return Trim(Text)
}

Training: ;{

	data := AbfrageGui("Metoprolol", {"use":"nothing", "ingredient":"Zucker"})
	MsgBox, % data.use

return ;}

Convert_IfapDB_Wirkstoffe(filepath) {

	;base:= "M:\albiswin\ifapDB\ibonus3\Datenbank"
	If !(dbf := FileOpen(filePath, "r", "CP1252")) {
			MsgBox, % "Dbf - file read failed.`n" filePath
			return 0
	}

	If !(save := FileOpen(A_ScriptDir "\Wirkstoffe.txt", "w", "UTF-8")) {
		MsgBox, % "Can't open file to save data.`n" A_ScriptDir "\Wirkstoffe.txt"
		return 0
	}

	Wirkstoffe := ""
	VarSetCapacity(buffin, 300, 0)
	LenDataSet := 156
	dbf.Seek((pos := 5), 0)

	while (!dbf.AtEOF) {

		pos += LenDataSet                                           	; Leseposition + eine Datensatzlänge
		dbf.RawRead(buffin, LenDataSet)                        	; liest einen Datensatz
		string := StrGet(&buffin, LenDataSet, "cp1252")
		string := Trim(StrReplace(string, "`r`n", " "))
		string := StrReplace(string, "`n", " ")

		If !InStr(string, "Ø") && !RegExMatch(string, "\d+[a-z]\%*")
			Wirkstoffe .= string "`n"

	}

	dbf.Close()

	Wirkstoffe := RTrim(Wirkstoffe, "`n")
	Sort, Wirkstoffe, U
	save.Write(Wirkstoffe)
	save.Close()

	Run, % A_ScriptDir "\Wirkstoffe.txt"

}

AbfrageGui(Medikament, data) {

	global

	GuiEnd := false

	Gui, VO: New
	Gui, VO: Font, s13 bold
	Gui, VO: Add, Text	, % "xm w400 	Center"                                	, % "[" Medikament "]"
	Gui, VO: Font, s10 Normal
	Gui, VO: Add, Text	, % "xm w60"                                             	, % "Grund"
	Gui, VO: Add, Edit	, % "x+5 w300	vVOuse"                              	, % data.use
	Gui, VO: Add, Text	, % "xm w60"                                             	, % "Wirkstoff"
	Gui, VO: Add, Edit	, % "x+5 w300	vVOingr"                           	, % data.ingredient
	Gui, VO: Add, Button, % "xm y+15	vVOok 	    	gVOSpeichern"	, % "Sichern"
	Gui, VO: Add, Button, % "x+15      	vVOcancel 	gVOGuiClose" 	, % "Abbruch"

	Gui, VO: Show, , Medikament kategorisieren

	Hotkey, IfWinActive, Medikament kategorisieren
	Hotkey, Enter, VOSpeichern
	Hotkey, IfWinActive

	Loop {
		If GuiEnd || !WinExist("Medikament kategorisieren")
			break
		Sleep 10
	}
	;~ while (!GuiEnd || WinExist("Medikament kategorisieren")) {
		;~ break
		;~ Sleep 10
	;~ }

return {"use":VOuse, "ingredient":VOingr}

VOGuiEscape:
VOGuiClose:
	Gui, VO: Destroy
	GuiEnd := true
return
VOSpeichern:
	Gui, VO: Submit
	GuiEnd := true
return {"use":VOuse, "ingredient":VOingr}

}

writeJSONFile(file, obj, indent) {

	file := FileOpen(file, "w", "UTF-8")
	file.Write(JSON.Dump(obj,,indent, "UTF-8"))
	file.Close()

}

;}

;----------------------------------------------------------------------------------------------------------------------------------------------
; Includes
;----------------------------------------------------------------------------------------------------------------------------------------------;{
#include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\lib\class_JSON.ahk
;}
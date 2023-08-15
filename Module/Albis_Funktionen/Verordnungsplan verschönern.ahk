; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . .
; . . . . . .                                                                  	VERORDNUNGSPLAN VERSCHÖNERN
global                                                         						DatumVom:= "28.12.2022"
; . . . . . .
; . . . . . .                      ROBOTIC PROCESS AUTOMATION FOR THE GERMAN MEDICAL SOFTWARE "ALBIS ON WINDOWS"
; . . . . . .                             BY IXIKO STARTED IN SEPTEMBER 2017 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE
; . . . . . .                          THIS SCRIPT USES THE - MICROSOFT COM OBJECT MODEL - TO INTERACT WITH MS WORD
; . . . . . .                                                   THIS SOFTWARE IS ONLY AVAIBLE IN GERMAN LANGUAGE
; . . . . . .
; . . . . . .      - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; . . . . . .      !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !!
; . . . . . .      - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; . . . . . .                                   THIS SCRIPT ONLY WORKS WITH AUTOHOTKEY OR AUTOHOTKEY_H - V1.1.32.XX
; . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .


/* VBA-SKRIPT: AUTOHOTKEY SKRIPT AUS MICROSOFT WORD STARTEN

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
	Hotkey, ~Esc, ExitScript

;}

;----------------------------------------------------------------------------------------------------------------------------------------------
; Variablen für die Ersetzungen
;----------------------------------------------------------------------------------------------------------------------------------------------;{

	RegRead, AlbisWinDir, HKLM, SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
	RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)

	ifapDBPath	:= AlbisWinDir "\ifapDB\ibonus3\Datenbank"
	userMedDB	:= AddendumDir "\include\Daten\medkategorien.json"

	If FileExist(userMedDB)
		medis := JSONData.Load(userMedDB, "", "UTF-8")
	else
		medis:=Object()

	medisCount   	:= medis.Count()   ; Bestand der bekannten Medikamentennamen
	rowsToDelete   	:= Array()

	ReplaceThatShit	:= {	"Aaa"                                 	: ""
								, 	"\-*1\s*A\s*(Pharm*a*|Pha)" 	: " "
								, 	"Ab[Zz]"                             	: ""
								, 	"Acino"                              	: ""
								, 	"Acis"                                	: ""
								, 	"Accord"                              	: ""
								, 	"Actavis"                              	: ""
								, 	"\bAL\b"                               	: ""
								, 	"Ari"                                  	: ""
								, 	"Arzneim"                          	: ""
								, 	"Aristo"                              	: ""
								, 	"Allomedic"                        	: ""
								,	"Augentropfen"                  	: ""
								, 	"Aurobindo"                       	: ""
								, 	"Arzne"                              	: ""
								, 	"Atid"                                	: ""
								,	"Aventis*"                             	: ""
								, 	"\bAWD\b"                           	: ""
								,	"Axcount"                             	: ""
								,	"Axico"                              	: ""
								,	"Axicorp\s*P"                     	: ""
								, 	"Basics"                             	: ""
								, 	"BAYER|Bayer"                   	: ""
								, 	"Becat"                              	: ""
								, 	"beta"                                	: ""
                            	,	"Br\W"                               	: ""
								,	"\s*\-\s+Ct(\d+)"                 	: " $1"
								,	"\-*\s*CT"                          	: ""
								,	"Deutsch"                          	: ""
								,	"Dosiergel"                        	: ""
								,	"Dosieraeros"                    	: ""
								,	"dura"                                  	: ""
								,	"dura\s*B*"                         	: ""
								,	"ED"                                     	: ""
								,	"Emra.Med"                          	: ""
								,	"Eurimpharm"                       	: ""
								,	"Fair.Med"                            	: ""
								,	"Fertigspritze"                    	: ""
								,	"Filmtabl*e*t*t*e*n*"            	: ""
								,	"Fs"                                   	: ""
								,	"Fta\sFTA"                            	: " FTA "
								,	"Filmt\sFTA"                         	: " FTA "
								,	"GALEN"                            	: ""
								,	"Glenmark"                       	: ""
								,	"GmbH*"                           	: ""
								,	"Hartka"                               	: ""
								,	"\-Hcl"                                	: ""
								,	"Hennin*g"                          	: ""
								,	"Heum*a*n*n*"                    	: ""
								,	"HEUMANN"                     	: ""
								,	"(\-|\s)HEXAL"                      	: ""
								,	"HEXAL"                               	: ""
								,	"Hkp"                                   	: ""
								,	"Huebe"                            	: ""
;								,	"Inject"                                 	: ""
								,	"Inte"                                    	: ""
								,	"Isis"                                  	: ""
								,	"kohlpharma*"                   	: ""
								,	"KwikPen"                             	: ""
								,	"Lomapharm"                    	: ""
								,	"Lichten(\d+)"                       	: " $1"
								,	"Lich"                                  	: ""
								,	"Micro"                                	: ""
								,	"Mili"                                   	: ""
								,	"Milinda"                              	: ""
								,	"Msr"                                 	: ""
								,	"Myl"                                    	: ""
								,	"Nachfuell"                           	: ""
								,	"Net"                                    	: ""
								,	"Novolet"                           	: ""
								,	"Orifarm"                             	: ""
								,	"Pe"                                    	: ""
								,	"Pen"                                 	: ""
								,	"Penf*i*l*l*"                           	: ""
								,	"\bPha*r*m*a*\b"                 	: ""
								,	"\bPh\b"                             	: ""
								,	"Protect(\d+)"                       	: " $1"
								,	"Rat"                                  	: ""
								,	"[Rr]atio*p*h*a*r*m*"         	: ""
								,	"Retard"                             	: ""
								,	"\bSanofi\b"                      	: ""
								,	"\bSANDOZ\b"                     	: ""
								,	"Sta(\d)"                             	: " $1"
								,	"\bS(TADA|tada)\b"              	: ""
								,	"\ssto"                                  	: ""
								;~ ,	"sto"                                 	: ""
								,	"Tab\sTAB"                           	: " TAB "
								,	"Tabl\W"                             	: ""
								,	"Tabletten"                         	: ""
								,	"\bT[Aa][Hh]\b"                   	: ""
								,	"\bTAD\b"                           	: ""
								,	"teilbar"                                	: ""
;								,	"TEI"                                  	: ""
								,	"TEVA"                               	: ""
								,	"\b(TOP|Top)\b"                   	: ""
								,	"Tro"                                 	: ""
;								,	"UTA"                                 	: ""
								,	"Vital"                                	: ""
								,	"Weichkaps*e*l*n*"            	: ""
								,	"\bWinthrop\b"                    	: ""
								,	"Zentiva"                           	: ""
								,	"4Wochen"                        	: ""
								,	"\barma\b"                         	: "" }

	MedForm			:= {	"AMP"    	: "Ampulle"
								, 	"ATR"     	: "Augentropfen"
								, 	"CRE"    	: "Creme"
								, 	"DOS"     	: "Dosierspray"
								, 	"EMU"    	: "Emulsion"
								, 	"FER"     	: "Fertigspritze"
								, 	"FLU"    	: "Flüssig"
								, 	"FLU"    	: "Flüssig"
								, 	"FTA"    	: "Filmtablette"
								, 	"GEL"		: "Gel"
								, 	"GRA"		: "Granulat"
								, 	"HKM"		: "Kapsel"
								, 	"HKP"		: "Kapsel"
								, 	"HPI" 		: "Kapsel"
								, 	"HVW"		: "Kapsel"
								, 	"IFB"  		: "Infusionsbeutel"
								, 	"IHP"  		: "Inhalationspack"
								, 	"INF" 		: "Infusionslösung"
								, 	"ILO" 		: "Injektionslösung"
								, 	"ISU" 		: "Suspension"
								, 	"KAP"		: "Kapsel"
								, 	"KMP"		: "Kapsel"
								, 	"KMR"		: "Kapsel"
								, 	"KPG"		: "Kombipackung"
								, 	"LOT"		: "Lotion"
								, 	"LSE" 		: "Lösung"
								, 	"MIL" 		: "Lotion"
								, 	"PEN"		: "Fertigpen"
								, 	"PFL" 		: "Pflaster"
								, 	"PIK"  		: "Infusionslösung"
								, 	"PLE" 		: "Pulverzubereitung"
								, 	"PST" 		: "Paste"
								, 	"PUL"		: "Pulver"
								, 	"REK" 		: "Retardkapsel"
								, 	"RET" 		: "Retardtablette"
								, 	"SHA"		: "Shampoo"
								, 	"SIR"  		: "Sirup"
								, 	"SPR" 		: "Spray"
								, 	"STI"  		: "Stift"
								, 	"TAB"		: "Tablette"
								, 	"TEI"  		: "Tropfen"
								, 	"TMR"		: "Tablette"
								, 	"TRO"		: "Tropfen"
								, 	"UTA"		: "Tablette"
								, 	"VGE"		: "Vaginalcreme"
								, 	"WKA"		: "Kapsel"
								, 	"XNC"		: "Nachtcreme"
								, 	"ZAM"		: "Ampulle"}

	Hinweise          	:= {1: 	"• Einnahme der Kapseln mit Flüssigkeit zur selben Tageszeit, unabhängig von den Mahlzeiten "
				           			.	"• Kapseln nicht zerkauen oder zerstoßen • nicht zusammen mit Grapefruitsaft einnehmen"}

	JSONData.save(AddendumDir "\include\Daten\MedRegEx1.json", ReplaceThatShit, true,, 1, "UTF-8")
	JSONData.save(AddendumDir "\include\Daten\MedRegEx2.json", MedForm, true,, 1, "UTF-8")

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
		If RegExMatch(Content, "i)(V.e.r.o.r.d.n.u.n.g.s.p.l.a.n|Medikationsplan|Medikament.*?Form.*?Wirkstoff.*?Hinweise.*?Grund)") {

			docname	:= oWord.Documents.Item(A_Index).Name
			VoPlan    	:= oWord.Documents.Item(A_Index).Tables(1)

			oWord.Documents(A_Index).Activate
			docNr := A_Index
			break
		}

	}

	oDoc := oWord.Documents(docname)


;}

;----------------------------------------------------------------------------------------------------------------------------------------------
; Medikamentenallergien oder -nebenwirkungen erkennen und gesondert in der Tablette positionieren
;----------------------------------------------------------------------------------------------------------------------------------------------;{

	; DIE EINTRÄGE ZU ALLERGIEN KÖNNEN SICH ÜBER MEHRERE ZEILEN ERSTRECKEN UND
	; WERDEN JETZT ZU EINER EINZIGEN ZEILE ZUSAMMENGEFASST. ES WIRD EINE ANDERE SCHRIFTART GEWÄHLT
	; UND DIE HINTERGRUNDFARBE DER ZELLE WIRD DUNKLER GEWÄHLT
		startrow:= 1
		VoPlan	:= oWord.ActiveDocument.Tables(1)
		rows		:= VoPlan.Rows.Count
		cols  	:= VoPlan.columns.Count
		cell   	:= VoPlan.Cell(2, 2)
		cell.Select

		Text  	:= cell.Range.Text
		Text  	:= SubStr(Text, 1, StrLen(Text) - 2)
		Text  	:= Trim(Text)

		If RegExMatch(Text, "##|Problem|Allergie") {

			cell.Range.Font.Size       	:= 11
			cell.Range.Font.Bold      	:= true
			cell.Range.Font.italic      	:= true
			cell.Range.Font.underline	:= true
			cell.Range.Text              	:= "!!! Allergien oder Nebenwirkungen gegen folgende Medikamenten: "

			Loop % (rows - 2) {                                                    	; Ende der Hinweise finden
				rowAllergieLast := A_Index + 2
				Text := VoPlan.Cell(2+A_Index, 2).Range.Text
				If InStr(Text, "##")
					break
			}

		; FELD LINKS UND ALLE FELDER RECHTS DER MEDIKAMENTENBEZEICHNUNG LEEREN, ALLE ZELLEN EINER ZEILE VERBINDEN ;{
			Loop % (rowAllergieLast - 2) {

				row := A_Index + 1
				VoPlan.Cell(row, 1).Range.Text := ""
				Loop 6
					VoPlan.Cell(row, 4 + A_Index).Range.Text := ""

				VoPlan.Rows(row).Cells.Merge
				thisrow := VoPlan.Rows(row)
				thisrow.Range.Font.Size := 10

			}
		;}

		; DIE MEDIKAMENTE MÜSSEN NEU NUMMERIERT WERDEN
			Loop % (rows - rowAllergieLast - 2)
				VoPlan.Cell(rowAllergieLast + A_Index, 1).Range.Text := A_Index

		; LÖSCHT DIE ZEILE DIE DAS ENDE DER HINWEISE MARKIERTE
			VoPlan.Rows(rowAllergieLast).Delete

		; ALLE HINWEISZEILEN WERDEN ZU EINER EINZIGEN ZELLE VERBUNDEN, DIESE ZELLE ERHÄLT EINEN ETWAS DUNKLEREN HINTERGRUND
			VoPlan.Cell(2,1).Merge(VoPlan.Cell(rowAllergieLast-1,1))
			VoPlan.Cell(2,1).Shading.BackgroundPatternColor :=  0xDDDDDD ; 0xBBGGRR (BGR = Blue Green Red)

		; HINWEISBEREICH ALS ERSTE ZEILE DER TABELLE EINSTELLEN (DURCH AUSSCHNEIDEN UND EINFÜGEN DER 1.ZEILE NACH DEN HINWEISEN)
			VoPlan.Rows(1).Range.Cut
			VoPlan.Rows(2).Range.Paste

		; STARTZEILE FÜR DIE ERSETZUNG IST JETZT DIE 3.ZEILE
			startrow := 2

		}

;}

;----------------------------------------------------------------------------------------------------------------------------------------------
; Medikamentbezeichnungen bereinigen von Pharmaherstellern, Packungsgrößen, sonstigen Angaben
;----------------------------------------------------------------------------------------------------------------------------------------------;{
	Loop % ((rows := VoPlan.Rows.Count) - startrow) {

		; SPALTE 2 AB ZEILE 2 - MEDIKAMENTBEZEICHNUNG - KÜRZEN UND KLARER DARSTELLEN
			thisrow	:= startrow + A_Index                        	; startrow kann 2 oder 3 sein !
			Cell   	:= VoPlan.Cell(thisrow, 2)
			Cell.Select
			Cell.Range.Font.bold := false
			Text   	:= Cell.Range.Text
			Text  	:= SubStr(Text, 1, StrLen(Text) - 2)
			Text  	:= Trim(Text)

		; LEERE ZEILE MERKEN - WIRD SPÄTER ENTFERNT
			If (StrLen(Text) = 0) {
				rowsToDelete.InsertAt(1, thisrow)
				continue
			}

		; UNNÖTIGES ENTFERNEN
			newText := Text " "
			newText := RegExReplace(newText , "^\!")                                                                                  	; ! - am Anfang entfernen
			newText := RegExReplace(newText , "\(.*\)")                                                                                	; Klammern () und Inhalt
			For shit, replacewith in ReplaceThatShit                                                                                      	; entfernt Herstellernamen
				newText := RegExReplace(newText , "\s" shit "\s*", (!replacewith ? " " : replacewith))

		; ERMITTELT DIE APPLIKATIONSFORM (AUS RET -> RETARDTABLETTE Z.B.), ENTFERNT DIE KURZBEZEICHNUNGEN (RET, TAB)
			For Abkuerzung, Form in MedForm
				If RegExMatch(newText " ", "i)\s" Abkuerzung "\s")
					break

		; FÜGT EIN LEERZEICHEN ZWISCHEN DOSIS UND EINHEIT EIN
			MedZusatz := ""
			If RegExMatch(newText " ", "\s(?<Dosis>\d+(\.\d+|\/\d+)*(?!\sSt))\s*(?<Einheit>mg|µg|g|IE\.*|I\.*E\.*)*\s", Med )
				newText := StrReplace(newText, Med, " " (MedZusatz := MedDosis . MedEinheit) " ")

		; ERSETZT
			newText := RegExReplace(newText , "\s{2,}", " ")                                                                            	; entfernt überflüssige Leerzeichen
			Cell.Range.Text := newText                                                                                                        	; fertig - Medikamentenbezeichnung auffrischen


		; GRUND DER EINNAHME / MEDIKAMENTE KATEGORISIEREN
		  ; erfasst damit auch Medikamente wie Berlinsulin H Basal
		  ; es verbleiben manchmal Applikationsbezeichnungen oder Packungsgrößen
			RegExMatch(newText, "^\p{Lu}[\s-]*[\pL\/\-]+\s*([\pL\/\-]\s)*(\p{Lu}\d+\s)*\s*[\pL\/\-]*", Medikament)
			Medikament := RegExReplace(Medikament , "\sN\d(\s|$)", " ")
			For Abkuerzung, Aform in MedForm
				Medikament := RegExReplace(Medikament , "i)\s" Abkuerzung "(\s|$)", " ")
			Medikament := Trim(Medikament)

			If !medis.haskey(Medikament)
				medis[Medikament] := Object()

			If (!medis[Medikament].use || !medis[Medikament].ingredient) {
				result := AbfrageGui(Medikament, medis[Medikament])
				If (result.use && result.ingredient)
					medis[Medikament] := result
			}

		; SCHRIFTGRÖSSE FÜR APPLIKATIONSFORM UND PACKUNGSGRÖSSE VERKLEINERN (Medikamentenbezeichnungen passen in eine Zeile)
			Cell.Select

			MedPart      	:= MedZusatz ? MedZusatz : Medikament
			StrLenMed		:= InStr(newText, MedPart) + StrLen(Trim(MedPart))
			myRange  	:= Cell.Range
			MedRange	:= [myRange.Start, myRange.Start+StrLenMed]
			TblRange  	:= [myRange.Start+StrLenMed, myRange.End]

			myRange.SetRange(MedRange.1, MedRange.2)
			myRange.Font.Bold   	:= true

			myRange.SetRange(TblRange.1, TblRange.2)
			myRange.Font.Bold	:= false
			myRange.Font.Size 	:= 9

		; APPLIKATIONSFORM EINTRAGEN (else ist für eine formatierte Ausgabe)
			VoPlan.Cell(thisrow, 3).Range.Font.bold := false
			If !RegExMatch(Form, "\(")
				VoPlan.Cell(thisrow, 3).Range.Text := Form
			else {
				RegExMatch(Form, "(.*)\s(\(.*\))", FM)
				VoPlan.Cell(thisrow, 3).Range.Text := FM1 Chr(10) FM2
			}

		; MEDIKAMENT ANWENDUNG UND WIRKSTOFFE EINTRAGEN
			VoPlan.Cell(thisrow, 	4).Range.Font.bold := false
			VoPlan.Cell(thisrow,   4).Range.Text := medis[Medikament].ingredient
			VoPlan.Cell(thisrow, 10).Range.Font.bold := false
			VoPlan.Cell(thisrow, 10).Range.Text := medis[Medikament].use
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
		VoPlan.Rows(row).Delete
;}

;----------------------------------------------------------------------------------------------------------------------------------------------
; alle Zellen die eine einzelne Null oder ein A (Aut idem Kennzeichnung) enthalten leeren, Zellen verbinden die Text enthalten
;----------------------------------------------------------------------------------------------------------------------------------------------;{
	Loop % (VoPlan.Rows.Count - startrow) {

		row                       	:= A_Index + startrow
		MergeRows           	:= false
		MergeRowsOverride	:= false

		Loop 4 {

		  ; ZELLENTEXT AUSLESEN
			column := 4+A_Index
			VoPlan.Cell(row, column).Range.Font.bold := false
			VoPlan.Cell(row, column).select
			Text := VoPlan.Cell(row, column).Range.Text
			If RegExMatch(Text, "^###")                                               	; Kommentarzeilen werden nicht verbunden
				MergeRowsOverride := true

		  ; LÖSCHEN VON 0 oder A (A - steht für aut idem)
		  ;~ SciTEOutput(row "," column " = " Text)
			If RegExMatch(Text, "\s*(0|A)\s*")
				VoPlan.Cell(row, column).Range.Text := ""

		  ; SYMBOLE FÜR 1/2, 1/4 oder 3/4
			Text1 := " " Text " "
			Text1 := RegExReplace(Text1, "\s1/2\s", "½")
			Text1 := RegExReplace(Text1, "\s1/4\s", "¼")
			Text1 := RegExReplace(Text1, "\s3/4\s", "¾")
			Text1 := RegExReplace(Text1, "i)(\d+)(IE|Hub)", "$1 $2")       	; LEERZEICHEN ZWISCHEN DOSIS UND EINHEIT
			Text1 := Trim(Text1)
			If (Text <> Text1)
				VoPlan.Cell(row, column).Range.Text := (Text := Text1)

			If RegExMatch(Text, "[\pL\.;\:\-]{2,}") && !RegExMatch(Text, "i)\d+\s*(IE|Hub)") || RegExMatch(Text, "i)Bedarf")
				MergeRows := true

		}

		If MergeRows && !MergeRowsOverride {

			VoPlan.Cell(row, 5).Range.Font.bold := false
			VoPlan.Cell(row, 5).Merge(VoPlan.Cell(row, 8))

			Text := RegExReplace(VoPlan.Cell(row, 5).Range.Text, "(\n|\r)", " ")
			Text := Text " "
			Text := RegExReplace(Text, "([A-Z][a-z]+)\s+([a-z]+)", "$1 $2")
			Text := RegExReplace(Text, "\sBed\.\s", " Bedarf ")
			Text := RegExReplace(Text, "Bei\s", "bei ")
			Text := RegExReplace(Text, "Bis\s", "bis ")
			Text := RegExReplace(Text, "T\.\s", "Tabl. ")
			Text := RegExReplace(Text, "z\.\s", "zur ")
			Text := RegExReplace(Text, "Nac\s", "Nacht")
			If RegExMatch(Text " ", "\s([Wochen\s]+)\s", WS)
				Text := StrReplace(Text, WS1, RegExReplace(Trim(WS1), "\s+", ""))

			VoPlan.Cell(row, 5).Range.Font.bold := false
			VoPlan.Cell(row, 5).Range.Text := SubStr(Text, 1, StrLen(Text)-1)

		}

	}

	; CARET IN DIE LETZE BEARBEITETE ZELLEN SETZEN
		VoPlan.Cell(row,1).select

;}

ExitApp

ExitScript:
MsgBox, 0x1000, % A_ScriptName, % "Skript wurde per Hotkey abgebrochen.", 2
ExitApp

;----------------------------------------------------------------------------------------------------------------------------------------------
; Interne Funktionen
;----------------------------------------------------------------------------------------------------------------------------------------------;{
GetCellText(tableObj, row, col) {


	try
		Text := tableObj.Cell(row, col).Range.Text
	catch	{
		;SciTEOutput(row ", " A_Index)
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

	global userMedDB
	static VOok, VOcancel, VOShow, VOuse, VOingr


	GuiEnd := false

	Gui, VO: New
	Gui, VO: Font, s13 bold
	Gui, VO: Add, Progress	, % "xm ym w400 	c"                           	, % "[" Medikament "]"
	Gui, VO: Add, Text    	, % "xm ym w400 	Center"                           	, % "[" Medikament "]"

	Gui, VO: Font, s10 Normal
	Gui, VO: Add, Text		, % "xm w60"                                             	, % "Grund"
	Gui, VO: Add, Edit		, % "x+5 w300	vVOuse"                              	, % data.use

	Gui, VO: Add, Text	, % "xm w60"                                             	, % "Wirkstoff"
	Gui, VO: Add, Edit	, % "x+5 w300	vVOingr"                           	, % data.ingredient

	Gui, VO: Add, Button, % "xm y+15	vVOok 	    	gVOSpeichern"	, % "Sichern"
	Gui, VO: Add, Button, % "x+15      	vVOcancel 	gVOGuiClose" 	, % "Abbruch"

	Gui, VO: Font, s8 Normal
	Gui, VO: Add, Button, % "xm y+5   	vVOShow   	gVOAnzeigen"	, % "Medikamentendaten ansehen"

	GuiControlGet, cp, VO:Pos, VOcancel
	GuiControlGet, dp, VO:Pos, VOShow
	GuiCOntrol, VO: Move, VOShow, % "x" Floor((cpX+cpW)/2 - dpW/2)

	Gui, VO: Show, , Medikament kategorisieren

	Hotkey, IfWinActive, Medikament kategorisieren
	Hotkey, Enter, VOSpeichern
	Hotkey, IfWinActive

	Loop {
		If (GuiEnd || !WinExist("Medikament kategorisieren"))
			break
		Sleep 10
	}

return {"use":VOuse, "ingredient":VOingr}

VOAnzeigen:
	Run, % userMedDB
return

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
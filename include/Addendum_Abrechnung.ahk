; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                                                       	** Addendum_Abrechnung **
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;		Beschreibung:	    	Funktionen mit Regeln zu hausärztlichen Ziffer zur Vereinfachung des Abrechnungsprozeßes
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
; 		Inhalt:
;       Abhängigkeiten:		Addendum_DBASE, Addendum_Ini
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Addendum_Calc started:    	18.01.2021
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


class EBMDB {                                                     ; EBM Zifferndatenbank - eine Objektklasse für den Abrechnungshelfer

	; ich versuche diese Klasse unabhängig von Addendum.ahk zu bekommen

	; ------------------------------------------------------------------------------------------
	; Standarddaten festlegen, Verzeichnisse anlegen
	; ------------------------------------------------------------------------------------------;{
		__New(AddendumDBPath, StartQuartal, ArztID="", debug=0) {

				;~ RegExMatch(A_ScriptDir, ".*(?=AlbisOnWindows)", AddendumDir)
				;~ AddendumDir .= "AlbisOnWindows"

			; debug
				this.debug				:= debug

			; Parameter sichern
				this.StartQuartal	:= StartQuartal
				this.ArztID				:= ArztID

			; Verzeichnisse
				this.admDBPath		:= AddendumDBPath
				this.EBMDBPath		:= this.admDBPath "\PatData"
				this.AlbisDBPath	:= this._GetAlbisPath() "\db"

			; Arbeitsdateien
				this.EBMZFilePath 	:= this.EBMDBPath "\ZIFFERN.adm"
				this.BefundPath 	:= this.AlbisDBPath "\BEFUND.dbf"

			; Strings und Abrechnungsziffern
				this.30Spaces 		:= "                                   "
				this.8Spaces 			:= "        "
				this.Abr					:= { "Z01740":"1x im Arztfall"
											,	"Z01732":"alle 3 Jahre"
											,	"Z01745":"alle 3 Jahre"
											,	"Z01746":"alle 3 Jahre"
											,	"Z03360":"2x im Behandlungsfall"
											,	"Z03362":"1x im Quartal,nicht neben 03360"
											,	"Z03220":"Chroniker"
											,	"Z03221":"wenn 03220 abgerechnet"}

			; Unterverzeichnis im Addendum Datenbankpfad anlegen falls nicht vorhanden
				If !((err := this._CreateFilePath(this.EBMDBPath)) = 0)
					throw A_ThisFunc ": Das Anlegen eines notwendigen Dateipfades [" this.EBMDBPath "] ist fehlgeschlagen!`nWindows Fehlercode: " err

			; Indexdatei lesen
				this.BefundDBF  	:= this._GetBefundDBF(noIndex)

			; Addendums EBM Datenbank anlegen falls noch nicht vorhanden
				If !FileExist(this.EBMZFilePath)
					this.CreateDB(this.StartQuartal, this.ArztID)

			; Datensatzlänge lesen
				this.OpenDB()
				this.LenDataSet 	:= Trim(this.EBMZ._ReadString(4))
				this.recordsStart 	:= this.LenDataSet + 1
				this.encoding		:= "UTF-8"

			return
		}

	;}

	; ------------------------------------------------------------------------------------------
	; EBM Ziffern Datenbank anlegen
	; ------------------------------------------------------------------------------------------;{
		CreateDB(StartQuartal, ArztID="") {                                                                	;-- Neuerstellung: 	EBM Ziffernsammler für Hausärzte - Abrechnungshelfer

			; StartQuartal(Formatierung YYYYQ) in startrecord übersetzen
				If (StrLen(StartQuartal) > 0) && !this.BefundDBF.HasKey(StartQuartal)
					throw A_ThisFunc ": StartQuartal [" StartQuartal "] ist unbekannt!"
				startrecord := (StrLen(StartQuartal) = 0) ? 1 : Floor(this.BefundDBF[StartQuartal].1)

			; Abrechnungsziffern von allen Patienten ab dem StartQuartal sammeln
				DBZiffern 	:= GetEBMNumbers(startrecord)

			; ----------------------------------------------------------------------------------------------------------------------------------------------------------------------
			; 	Vergleich der Objekte Abr und DBZiffern.
			; ----------------------------------------------------------------------------------------------------------------------------------------------------------------------
			; 	bei einer Ziffer wie 01732, 01745, 01746, 01740 wird nur das letzte Abrechnungsquartal gespeichert
			; 	bei einer Ziffern wie 03220, 03221, 03360, 03362, werden die letzten 5 Quartale in denen diese abgerechnet wurden erfasst
			; 	es werden für alle Patienten die letzten 5 Quartale in denen ein Abrechnungsschein (Ziffern 0300x) angelegt wurde gesammelt
			;
				PatAbr := Object()
				For mIdx, m in DBZiffern
					For idx, lko in StrSplit(m.inhalt, "-")
						If this.Abr.HasKey("Z" lko) || RegExMatch(lko, "0300\d") {

							NR    	:= SubStr("         " m.PATNR, -8)
							LK	    	:= "Z" lko
							YQ   	:= SubStr(m.Datum, 1, 4) Ceil(SubStr(m.Datum, 5, 2)/3)
							MAXNR := m.PATNR > MAXNR ? m.PATNR : MAXNR

							If !IsObject(PatAbr[NR])
								PatAbr[NR] := Object()

							If RegExMatch(LK, "Z0300\d")
								LK := "Q"
							else If InStr("Z01732,Z01746,Z01745,Z01740", LK) {
								PatAbr[NR][LK] := m.Datum
								continue
							}

							If !InStr(SubStr(PatAbr[NR][LK], -4), YQ)
								PatAbr[NR][LK] := PatAbr[NR].HasKey(LK) ? (SubStr(PatAbr[NR][LK], -29) YQ) : SubStr(30Spaces YQ, -34)

					}

			; ----------------------------------------------------------------------------------------------------------------------------------------------------------------------
			; 	Speicherung der Daten in einem eigenen Format
			; ----------------------------------------------------------------------------------------------------------------------------------------------------------------------
			;	feste Spaltengrößen und damit feste Zeilenlänge wie in DBASE Dateien
			; 	damit kann bei Übergabe der Patienten Nr der zugehörige Datensatz direkt ausgelesen werden (sollte dadurch sehr schnell sein)
			;
				AbrDBF := FileOpen(EBMZFilePath, "w", "UTF-8")
				Loop % MAXNR {

					NR	:= A_Index
					PNR	:= SubStr("         " A_Index, -8)                                                                       ; PatNR                                                 	-	9 Zeichen
					QZ 	:= !PatAbr[NR].HasKey("Q") ? 30Spaces : PatAbr[NR]["Q"]	                         	; die letzten 5 Quartale des Patienten      	- 30 Zeichen
					Z1 	:= !PatAbr[NR].HasKey("Z03220") ? 30Spaces 	: PatAbr[NR]["Z03220"] 		; abgerechnete Chronikerziffer 5 Quartale - 30 Zeichen
					Z2 	:= !PatAbr[NR].HasKey("Z03221") ? 30Spaces 	: PatAbr[NR]["Z03221"] 		; abgerechnete Chronikerziffer 5 Quartale - 30 Zeichen
					Z3 	:= !PatAbr[NR].HasKey("Z03360") ? 30Spaces   	: PatAbr[NR]["Z03360"] 		; abgerechnete GB 5 Quartale - 30 Zeichen
					Z4 	:= !PatAbr[NR].HasKey("Z03362") ? 30Spaces   	: PatAbr[NR]["Z03362"] 		; abgerechnete GB 5 Quartale - 30 Zeichen
					Z5 	:= !PatAbr[NR].HasKey("Z01732") ? 8Spaces   	: PatAbr[NR]["Z01732"]
					Z6 	:= !PatAbr[NR].HasKey("Z01740") ? 8Spaces   	: PatAbr[NR]["Z01740"]
					Z7 	:= !PatAbr[NR].HasKey("Z01745") ? 8Spaces   	: PatAbr[NR]["Z01745"]
					Z8 	:= !PatAbr[NR].HasKey("Z01746") ? 8Spaces   	: PatAbr[NR]["Z01746"]
					;ToolTip, % NR ": " PatAbr[NR]["Z03220"]

					dataset := PNR "|" QZ "|" Z1 "|" Z2 "|" Z3 "|" Z4 "|" Z5 "|" Z6 "|" Z7 "|" Z8 "`n"
					If (A_Index = 1) {
						header := SubStr("    " StrLen(dataset), -4)
									. 	"|Q"          	SubStr("   " StrLen(QZ)	, -2)
									. 	"|Z03220" 	SubStr("   " StrLen(Z1) 	, -2)
									.	"|Z03221" 	SubStr("   " StrLen(Z2) 	, -2)
									.	"|Z03360" 	SubStr("   " StrLen(Z3) 	, -2)
									.	"|Z03362" 	SubStr("   " StrLen(Z4) 	, -2)
									.	"|Z01732" 	SubStr("   " StrLen(Z5) 	, -2)
									.	"|Z01740" 	SubStr("   " StrLen(Z6) 	, -2)
									.	"|Z01745" 	SubStr("   " StrLen(Z7) 	, -2)
									.	"|Z01746" 	SubStr("   " StrLen(Z8) 	, -2)
									.	"|"

						Loop % (StrLen(dataset) - StrLen(header) - 1)
							header := header " "

						AbrDBF.Write(header "`n")
					}

					AbrDBF.Write(dataset)

				}
				AbrDBF.Close()

		return PatAbr
		}

	; EBM Ziffern Datenbank aktualisieren
		UpDateDB(StartQuartal, ArztID="") {

			; StartQuartal(Formatierung YYYYQ) in startrecord übersetzen
				If (StrLen(StartQuartal) > 0) && !this.BefundDBF.HasKey(StartQuartal)
					throw A_ThisFunc ": StartQuartal [" StartQuartal "] ist unbekannt!"
				startrecord := (StrLen(StartQuartal) = 0) ? 1 : Floor(this.BefundDBF[StartQuartal].1)

			; Abrechnungsziffern von allen Patienten ab dem StartQuartal sammeln
				DBZiffern 	:= GetEBMNumbers(startrecord)

			; ----------------------------------------------------------------------------------------------------------------------------------------------------------------------
			; 	Vergleich der Objekte Abr und DBZiffern mit den zuvor gesammelten Daten.
			; ----------------------------------------------------------------------------------------------------------------------------------------------------------------------
			;


		}

	;}

	; ------------------------------------------------------------------------------------------
	; Addendums EBM Datenbank
	; ------------------------------------------------------------------------------------------;{

		OpenDB() {

			If !IsObject(this.EBMZ := FileOpen(this.EBMZFilePath, "rw", "UTF-8")) {
				throw "open database file: " this.filepath " failed!"
				ExitApp
			}

			this.EBMZ.Seek(this.recordsStart, 0)

		return this.EBMZ.Tell()
		}

		CloseDB() {

			filepos := this.EBMZ.Tell()
			this.EBMZ.Close()

		return filepos
		}

		ReadDBSet(PATNR) {

			; Fehlerbehandlung
				If !RegExMatch(PATNR, "\d+")
					throw A_ThisFunc ": Parameter PATNR can only contain digits!"

			; zum Datensatz springen
				If (PATNR > 0) {
					this.EBMZ.Seek(this.recordsStart + (this.lendataset * (PATNR-1)), 0)
					this.filepos := this.EBMZ.Tell()
				}

			; Datensatz laden
				bytes	:= this.dbf.RawRead(recordbuf, this.lendataset)
				set 	:= StrGet(&recordbuf, this.lendataset, this.encoding)

			; Objekt erstellen
				dataset := Object()
				Tmp := StrSplit(set, "|")
				dataset["PATID"]        	:= Trim(Tmp.1)
				dataset["Q"]	        	:= Trim(Tmp.2)
				dataset["Z03220"]		:= Trim(Tmp.3)
				dataset["Z03221"]		:= Trim(Tmp.4)
				dataset["Z03360"]		:= Trim(Tmp.5)
				dataset["Z03362"]		:= Trim(Tmp.6)
				dataset["Z01732"]		:= Trim(Tmp.7)
				dataset["Z01740"]		:= Trim(Tmp.8)
				dataset["Z01745"]		:= Trim(Tmp.9)
				dataset["Z01746"]		:= Trim(Tmp.10)

		return dataset
		}

	;}

	; ------------------------------------------------------------------------------------------
	; Albis Datenbankfunktionen
	; ------------------------------------------------------------------------------------------;{
		GetEBMNumbers(startpos) {

			befund  	:= new DBASE(this.BefundPath, this.debug)
			filepos  	:= befund.OpenDBF()
			ziffern    	:= befund.SearchFast({"KUERZEL": "lko"}, ["DATUM", "PATNR", "INHALT"], startpos)
			filepos  	:= befund.CloseDBF()

		return ziffern
		}
	;}

	; ------------------------------------------------------------------------------------------
	; Hilfsfunktionen
	; ------------------------------------------------------------------------------------------;{
		_GetBefundDBF(ByRef noIndex) {                                                                    	;-- Index der Befund.dbf lesen

			; lädt die Position für fileseek an denen sich der erste Quartaleintrag in der 'BEFUND.dbf' befindet
				If !IsObject(BefundDBF) {

					noIndex := false
					IndexFilePath := this.EBMDBPath "\DBIndex_BEFUND.json"
					If FileExis(IndexFilePath)
						BefundDBF := JSON.Load(FileOpen(IndexFilePath, "r", "UTF-8").Read())
					else {
						BefundDBF := this._CreateDBIndex("BEFUND", true)
						noIndex := true
					}

				}

		return BefundDBF
		}

		_CreateDBIndex(ReIndex:=false) {                                                                 	;-- Indexerstellung einer Albis-Datenbank durchführen

			; Index der DBASE Datei BEFUND.dbf wird erstellt
				dbfFile := new DBASE(this.EBMDBPath "\DBIndex_BEFUND.json", this.debug)
				DBFIndex := dbfFile.CreateIndex("quarter", this.EBMDBPath "\DBIndex_BEFUND.json", ReIndex)

		return DBFIndex
		}

		_CreateFilePath(path) {                                                                                     	;-- erstellt einen Dateipfad falls dieser noch nicht existiert

			If !InStr(FileExist(path "\"), "D") {
				FileCreateDir, % path
				If ErrorLevel
					return A_LastError
				else
					return 0
			}

		return 0
		}

		_GetAlbisPath() {                                                                                           	;-- liest das Albisinstallationsverzeichnis aus der Registry

				If (A_PtrSize = 8)
					SetRegView	, 64
				else
					SetRegView	, 32

				RegRead, PathAlbis, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-MainPath
				If (StrLen(PathAlbis) = 0)
					throw "Der Installationspfad von Albis konnte nicht aus der Windows Registry gelesen werden!"

		return PathAlbis
		}

		_ReadString(bytes) {                                                                                      	;-- reads bytes from database and returns the encoded string
			VarSetCapacity(buffin, bytes, 0)
			this.dbf.RawRead(buffin, bytes)
		return StrGet(&buffin, bytes, this.encoding)
	}
	;}

	; ------------------------------------------------------------------------------------------
	; Gui's für die Ersteinrichtung
	; ------------------------------------------------------------------------------------------;{
		Gui_Startquartal() {

			global
			local Cl:= Array()

			Text2 =
			(LTrim
			Der Addendum Abrechnungshelfer liest Daten aus den Albis Datenbankdateien.
			Einige Quartale sind aber möglicherweise nicht interessant weil diese z.B. Abrechnungsdaten
			ihrer/ihres Praxisvorgängers enthalten. Würden diese Daten miterfasst sind im Anschluß die
			Ausgaben fehlerhaft. Sie können hier das Start-Quartal und ihre Arzt-ID (Kürzel) eingeben.
			)

			Gui, SQ: new    	, HWNDhSQ -DPIScale
			Gui, SQ: Margin	, 10 , 5
			Gui, SQ: Color  	, cCCCCAA
			Gui, SQ: Font    	, s18 q5 bold underline, Calibri
			Gui, SQ: Add    	, Text, % "xm ym vsqT1"                                            	, % "Bitte wählen Sie ein Startquartal aus"
			Gui, SQ: Font    	, s9 q5 Normal, Calibri
			Gui, SQ: Add    	, Text, % "xm ym vsqT2"                                            	, % StrReplace(Text2, "`n")

		}

	;}


}




EBMdb_Create(StartQuartal="20094") {                                                                	;-- Neuerstellung: 	EBM Ziffernsammler für Hausärzte - Abrechnungshelfer

	; erstellt im Moment die Daten immer neu, überschreibt somit alles

	/*  	Parameterbeschreibung

			PatID:				- aktualisiert nur die Daten für diesen Patienten
			StartQuartal: 	- Formatierung YYYYQ
			rebuild:				- Addendums eigene Datenzusammenstellung neu erstellen lassen wenn true
									  ansonsten nur Veränderungen speichern

	*/

	; Variablen ;{
		static 30Spaces := "                                   "
		static 8Spaces 	:= "        "

		abr := GetAbrZiffern()
		EBMZFilePath := Addendum.DBPath "\PatData\ZIFFERN.adm"

		If !IsObject(BefundDBF)
			BefundDBF := GetBefundDBF(noIndex)
	;}

	; Daten von allen Patienten über alle Jahre sammeln
		befund  	:= new DBASE(Addendum.AlbisDBPath "\BEFUND.dbf", 0)
		filepos  	:= befund.OpenDBF()
		startpos  	:= Floor((BefundDBF[StartQuartal].1) ; .1 ist die Record Nr innerhalb der DBASE Datei (Addendum_DBASE arbeitet mit diesen Sprungposition)
		ziffern    	:= befund.SearchFast({"KUERZEL": "lko"}, ["DATUM", "PATNR", "INHALT"], startpos)
		filepos  	:= befund.CloseDBF()

	; Ziffern aus Abr mit Ziffern aus der Datenbank vergleichen
		PatAbr := Object()
		For mIdx, m in ziffern {

			For idx, lko in StrSplit(m.inhalt, "-")
				If Abr.HasKey("Z" lko) || RegExMatch(lko, "0300\d") {

					NR	:= SubStr("         " m.PATNR, -8)
					LK		:= "Z" lko
					YQ 	:= SubStr(m.Datum, 1, 4) Ceil(SubStr(m.Datum, 5, 2)/3)

					MAXNR := m.PATNR > MAXNR ? m.PATNR : MAXNR

					If !IsObject(PatAbr[NR])
						PatAbr[NR] := Object()

					If RegExMatch(LK, "Z0300\d")
						LK := "Q"
					else If InStr("Z01732,Z01746,Z01745,Z01740", LK) {
						PatAbr[NR][LK] := m.Datum
						continue
					}

					If !InStr(SubStr(PatAbr[NR][LK], -4), YQ) {
						PatAbr[NR][LK] := PatAbr[NR].HasKey(LK) ? (SubStr(PatAbr[NR][LK], -29) YQ) : SubStr(30Spaces YQ, -34)
					}
			}

		}

	; speichert Daten in einem eigenen Format
		AbrDBF := FileOpen(EBMZFilePath, "w", "UTF-8")
		Loop % MAXNR {

			NR	:= A_Index
			PNR	:= SubStr("         " A_Index, -8)                                                                       ; PatNR                                                 	-	9 Zeichen
			QZ 	:= !PatAbr[NR].HasKey("Q") ? 30Spaces : PatAbr[NR]["Q"]	                         	; die letzten 5 Quartale des Patienten      	- 30 Zeichen
			Z1 	:= !PatAbr[NR].HasKey("Z03220") ? 30Spaces 	: PatAbr[NR]["Z03220"] 		; abgerechnete Chronikerziffer 5 Quartale - 30 Zeichen
			Z2 	:= !PatAbr[NR].HasKey("Z03221") ? 30Spaces 	: PatAbr[NR]["Z03221"] 		; abgerechnete Chronikerziffer 5 Quartale - 30 Zeichen
			Z3 	:= !PatAbr[NR].HasKey("Z03360") ? 30Spaces   	: PatAbr[NR]["Z03360"] 		; abgerechnete GB 5 Quartale - 30 Zeichen
			Z4 	:= !PatAbr[NR].HasKey("Z03362") ? 30Spaces   	: PatAbr[NR]["Z03362"] 		; abgerechnete GB 5 Quartale - 30 Zeichen
			Z5 	:= !PatAbr[NR].HasKey("Z01732") ? 8Spaces   	: PatAbr[NR]["Z01732"]
			Z6 	:= !PatAbr[NR].HasKey("Z01740") ? 8Spaces   	: PatAbr[NR]["Z01740"]
			Z7 	:= !PatAbr[NR].HasKey("Z01745") ? 8Spaces   	: PatAbr[NR]["Z01745"]
			Z8 	:= !PatAbr[NR].HasKey("Z01746") ? 8Spaces   	: PatAbr[NR]["Z01746"]

			;ToolTip, % NR ": " PatAbr[NR]["Z03220"]

			dataset := PNR "|" QZ "|" Z1 "|" Z2 "|" Z3 "|" Z4 "|" Z5 "|" Z6 "|" Z7 "|" Z8 "`n"
			If (A_Index = 1) {
				LenDataSet :=
				header := SubStr("    " StrLen(dataset), -4)
							. 	"|Q"          	SubStr("   " StrLen(QZ)	, -2)
							. 	"|Z03220" 	SubStr("   " StrLen(Z1) 	, -2)
							.	"|Z03221" 	SubStr("   " StrLen(Z2) 	, -2)
							.	"|Z03360" 	SubStr("   " StrLen(Z3) 	, -2)
							.	"|Z03362" 	SubStr("   " StrLen(Z4) 	, -2)
							.	"|Z01732" 	SubStr("   " StrLen(Z5) 	, -2)
							.	"|Z01740" 	SubStr("   " StrLen(Z6) 	, -2)
							.	"|Z01745" 	SubStr("   " StrLen(Z7) 	, -2)
							.	"|Z01746" 	SubStr("   " StrLen(Z8) 	, -2)
							.	"|"

				Loop % (StrLen(dataset) - StrLen(header) - 1)
					adds .= "."

				AbrDBF.Write(header adds "`n")
			}

			AbrDBF.Write(dataset)

		}

		AbrDBF.Close()


return PatAbr
}

EBMdb_Refresh(StartQuartal) {                                                                               	;-- Hinzufügen:  	EBM Ziffernsammler für Hausärzte - Abrechnungshelfer

	; Variablen ;{
		static 30Spaces := "                                   "
		static 8Spaces 	:= "        "
		abr := GetAbrZiffern()
	;}

	; EBMdb vorhanden, wenn nicht erstellen
		EBMZFilePath := Addendum.DBPath "\PatData\ZIFFERN.adm"
		If !FileExist(EBMZFilePath) {
			PatAbr := EBMdb_Create(StartQuartal)
		return PatAbr
		}

	; lädt die Position für fileseek-Positionen an denen sich der jeweils erste Quartaleintrag in der 'BEFUND.dbf' befindet
		BefundDBF := GetBefundDBF(noIndex)

	; Startposition festlegen
		If (StartQuartal = "letztes") {
			For recQ, position in BefundDBF
				continue
			startrecord := position.1
		}
		else
			startrecord := BefundDBF[StartQuartal].1

	; Daten ab dem StartQuartal einsammeln
		befund  	:= new DBASE(Addendum.AlbisDBPath "\BEFUND.dbf", 0)
		filepos  	:= befund.OpenDBF()
		ziffern    	:= befund.SearchFast({"KUERZEL": "lko"}, ["DATUM", "PATNR", "INHALT"], startrecord)
		filepos  	:= befund.CloseDBF()

	; Ziffern aus Abr mit Ziffern aus der Datenbank vergleichen
		AbrDBF := FileOpen(EBMZFilePath, "rw", "UTF-8")
		PatAbr := Object()
		For mIdx, m in ziffern {

			For idx, lko in StrSplit(m.inhalt, "-")
				If Abr.HasKey("Z" lko) || RegExMatch(lko, "0300\d") {

					NR	:= SubStr("         " m.PATNR, -8)
					LK		:= "Z" lko
					YQ 	:= SubStr(m.Datum, 1, 4) Ceil(SubStr(m.Datum, 5, 2)/3)

					MAXNR := m.PATNR > MAXNR ? m.PATNR : MAXNR

					If !IsObject(PatAbr[NR])
						PatAbr[NR] := Object()

					If RegExMatch(LK, "Z0300\d")
						LK := "Q"
					else If InStr("Z01732,Z01746,Z01745,Z01740", LK) {
						PatAbr[NR][LK] := m.Datum
						continue
					}

					If !InStr(SubStr(PatAbr[NR][LK], -4), YQ) {
						PatAbr[NR][LK] := PatAbr[NR].HasKey(LK) ? (SubStr(PatAbr[NR][LK], -29) YQ) : SubStr(30Spaces YQ, -34)
					}
			}

		}


}

BuildRangeString(Zeitraum) {

	If RegExMatch(Zeitraum, "(?<SorL>[\<\>])*(?<isEq>\=)*(?<Q1>\d{2}.\d{4})*(?<D1>\d{8})*(?<Range>\-)*(?<Q2>\d{2}.\d{4})*(?<D2>\d{8})*", z) {

		;If (StrLen(Q1) > 0)


	}

}

Leistungskomplexe(PatID=0, Zeitraum, SaveToPath="") {

	; Parameter: 	Zeitraum - hilft nur die dazu passenden Eintragungen zu untersuchen
	;					- 	das übergebene Format muss dem DBASE Format entsprechen also YYYYMMDD
	;					-	es können aber auch nur die Daten eines Quartals ("01/2011") oder ein Quartalsbereich ("01/2011-03/2017")
	;						die Schreibweise "012020" respektive "012011-032017" wird auch akzeptiert
	;					-	stellen Sie ein größer als Zeichen (">") vor den Suchzeitraum werden alle Daten nach diesem Datum in die Suche einbezogen
	;					-	dementsprechend wird ein kleiner als Zeichen ("<") nur alle Daten vor diesem Datum durchsuchen
	;					-	ein "=" ist für die Tages- oder Quartalsgenaue Suche muss nicht angeben werden, verwenden Sie das "=" nach "<" oder "<"
	;						im Sinne von kleiner gleich oder größer gleich
	;					-	ein Ausrufezeichen "!" steht für den Suchausschluß, ein "!01/2011" würde also Datumszahlen untersuchen die nicht in diesem
	;						Quartal liegen. Das ganze funktioniert aber auch als Bereichsuche "!01/2011-03/2011"

		static AlbisDBPath := GetAlbisPath() "\db"
		static Abr	:= {	"01740":"1x im Arztfall"
        						"01732":"alle 3 Jahre"
        						"01746":"alle 3 Jahre"
        						"03360":"2x im Behandlungsfall"
        						"03362":"1x im Quartal,nicht neben 03360"
        						"03220":"Chroniker"
        						"03221":"wenn 03220 abgerechnet"}

	; get dates index if available for 'BEFUND.dbf'
		;BEFUNDidx := ReadDBASEIndex(AlbisDBPath, "BEFUND")

	; Zeitraum für den Suchfilter parsen
		; 20120102 = 20

	; Filtereinstellung für die Suche in der Albis BEFUND.dbf
		inpattern  	:= {	"KUERZEL"	: "lko"
							, 	"DATUM"	: :">=20100101&&<=20111231"}						; jedes Datum ab 2010 wird untersucht!!
							;, 	"DATUM"	: [1:{"match":"20100101", "condition":">="}]}						; jedes Datum ab 2010 wird untersucht!!

	; es können gezielt nur die Daten eines Patienten herausgesucht werden
		If PatID
			inpattern["PatNR"] := PatID

	; welche Werte die Suchfunktion zurückgeben soll
		outfields := ["DATUM", "PATNR", "INHALT"]		                                            	; DSVGO = Löschsymbol?

	; die Suche starten
		befund := new DBASE(AlbisDBPath "\BEFUND.dbf", 1)
		ziffern  	:= befund.SearchFields(inpattern, outpattern)

	; sucht nur nach bestimmten Ziffern
		For ZIndex, m in ziffern {

			If !IsObject(Tmp[(PatID := m.PATNR)])
				Tmp[PatID] := Object()

			;Tmp[PatID].Chroniker := PatDBF[PAtID].Chroniker

			If m.removed
				ZZ := "*Z"
			else
				ZZ := "Z"

		; trennt Ziffernfolgen auf um daraus einen key zu erzeugen
			For lidx, LZiffer in StrSplit(m.Inhalt, "-")
				If RegExMatch(LZiffer, "^\d+")
					For AbrZiffer, AbrRegel in Abr
						If RegExMatch(LZiffer, "^" AbrZiffer) {
							Tmp[PatID][ZZ RTrim(LZiffer, "H")] := m.Datum
							break
						}
		}

		PatExtra := Object()
		For PatExID, ZExtra in Tmp
			If (ZExtra.Count() > 0) {
				PatExtra[PatExID] := ZExtra
			}

		If (StrLen(SaveToPath) > 0)  ;## PatID <> 0 in extra Patientenordner speichern
			FileOpen(admDBPath "\PatData\PatExtra.json", "w", "UTF-8").Write(JSON.DUMP(PatExtra,,1))

return PatExtra
}

Hausbesuchsrechner() {

	HBKomplexe := { 	"K1"      	: ["01410", "01411", "01412", "01413", "01414", "01415", "01416", "01418"]            	; Hausbesuchsziffern
			    			, 	"K2"      	: ["97234", "97235", "97236", "97237", "97238", "97239"]            	; Wegegebühren
							, 	"K3"      	: ["03370", "03372"]                                                                    	; Palliativziffern
					    	,	"Zusatz" 	: {"01415" :	["37102", "37113", "03370", "03372", "03373", "37306", "01102", "03030"]
												  ,	"01410" :  ["03370", "03372(x:2)", "37305", "01102", "01205", "01210", "01207"]
												  ,	"01411" :  ["03370", "03373", "37306", "01102", "03030"]
												  ,	"01412" :  ["03370", "03373", "37306", "01102", "03030"]
												  ,	"01413" :  ["03370", "03372(x:2)", "37305", "01102"]
												  ,	"01418" :  ["03370", "03372", "03373", "37305", "01102", "01207"]}
				    		,	"Regel"  	: {"01413" : 	"-Wege"}
							,	"Texte"		: {"01410" :	 	"Besuch eines Kranken, wegen der Erkrankung ausgeführt."
												,   "01411" :	 	"Dringender Besuch wegen der Erkrankung, unverzüglich nach Bestellung ausgeführt."
												,   "01412" :		"Dringender Besuch wegen der Erkrankung, unverzüglich nach Bestellung ausgeführt"
														   		   . 	" zwischen 22 und 7 Uhr oder an Samstagen, Sonntagen und gesetzlichen Feiertagen"
														    	   . 	", am 24.12. und 31.12., zwischen 19 und 7 Uhr oder bei Unterbrechen der"
																   . 	" Sprechstundentätigkeit mit Verlassen der Praxisräume"
												,   "01413" :	 	"Besuch eines weiteren Kranken in derselben sozialen Gemeinschaft (z.B. Familie) "
																   . 	"und/oder in beschützenden Wohnheimen bzw. Einrichtungen bzw. Pflege- oder Altenheimen mit Pflegepersonal"
												,	"03372" :	 	"Zuschlag zu den Gebührenordnungspositionen 01410 oder 01413 für die palliativmedizinische Betreuung in der Häuslichkeit"
																   .	"`t- mehrfach abrechenbar je vollendete 15min"
												,	"03373" :	 	"Zuschlag zu den Gebührenordnungspositionen 01411, 01412 oder 01415 für die palliativmedizinische Betreuung in der Häuslichkeit"
																   .	"`t- je Besuch nur einem abrechenbar"
												,	"01102" :	 	"Inanspruchnahme des Arztes nach Vorbestellung an Samstagen zwischen 7 und 14 Uhr, 10,93 Euro"
												,	"01205" :		"Notfallpauschale im organisierten Not(-fall)dienst und für nicht an der vertragsärztlichen Versorgung teilnehmende Ärzte,"
																   .	" Institute und Krankenhäuser für die Abklärung der Behandlungsnotwendigkeit bei Inanspruchnahme`n"
																   .	"`t- zwischen 07:00 und 19:00 Uhr (außer an Samstagen, Sonntagen, gesetzlichen Feiertagen und am 24.12. und 31.12.)"
												,	"01207" :	  	"Notfallpauschale im organisierten Not(-fall)dienst und für nicht an der vertragsärztlichen Versorgung"
																   .  	" teilnehmende Ärzte, Institute und Krankenhäuser für die Abklärung der Behandlungsnotwendigkeit bei Inanspruchnahme"
																   .  	"`t- zwischen 19:00 und 07:00 Uhr des Folgetages`n`t- ganztägig an Samstagen, Sonntagen, gesetzlichen Feiertagen"
																   .  	" und am 24.12. und 31.12."}}

return HBKomplexe
}

; Hilfsfunktionen
VorquartaleR(Datum, Anzahl, returnFormat="String") {                                              	;-- erstellt einen formatierten String für den Vergleich

	; Rückgabeformat: String oder Array
		ret := returnFormat = "Array" ? true : false
		If ret
			VorQuartale := Array()

	; Berechnung der Vorquartale
		Loop % Anzahl {
			VorQuartal := TT_VorQuartal(Datum, "YYYYQ")
			If ret
				VorQuartale.Push(VorQuartal)
			else
				VorQuartale := VorQuartal . VorQuartale
			Datum := {"aktuell":("0" SubStr(VorQuartal, 5, 1) SubStr(VorQuartal, 3, 2))}
		}

return VorQuartale
}

CreateDBIndex(DB:="BEFUND", ReIndex:=false) {                                                  	;-- Indexerstellung einer Albis-Datenbank durchführen

	; ACHTUNG funktioniert wahrscheinlich nur bei der BEFUND.dbf im Moment, da nur nach Quartalen indiziert wird

	; prüft ob der Datenbankpfad angelegt, legt diesen an wenn nicht
		If !((err := CreateFilePath, Addendum.DBPath "\PatData")) = 0) {
			throw A_ThisFunc ": Das Anlegen des Dateipfades [" Addendum.DBPath "\PatData\" "] ist fehlgeschlagen!"
		}

	; Index der DBASE Datei BEFUND.dbf wird erstellt
		dbfFile := new DBASE(Addendum.AlbisDBPath "\" DB ".dbf", 1)
		DBFIndex := dbfFile.CreateIndex("quarter", Addendum.DBPath "\PatData\DBIndex_" DB ".json", ReIndex)

return DBFIndex
}

GetBefundDBF(ByRef noIndex) {

		static BefundDBF
		noIndex := false

	; lädt die Position für fileseek an denen sich der erste Quartaleintrag in der 'BEFUND.dbf' befindet
		If !IsObject(BefundDBF) {
			IndexFilePath := Addendum.DBPath "\PatData\DBIndex_BEFUND.json"
			If FileExis(IndexFilePath)
				BefundDBF := JSON.Load(FileOpen(IndexFilePath, "r").Read())
			else {
				BefundDBF := CreateDBIndex("BEFUND", true)
				noIndex := true
			}
		}

return BefundDBF
}

GetAbrZiffern() {

	static Abr	:= { "Z01740":"1x im Arztfall"
						,	"Z01732":"alle 3 Jahre"
						,	"Z01745":"alle 3 Jahre"
						,	"Z01746":"alle 3 Jahre"
						,	"Z03360":"2x im Behandlungsfall"
						,	"Z03362":"1x im Quartal,nicht neben 03360"
						,	"Z03220":"Chroniker"
						,	"Z03221":"wenn 03220 abgerechnet"}

return Abr
}


;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                                  	  (PSEUDO-) DATENBANKFUNKTION für die
;                                        ADDENDUM INTERNE PATIENTENNAMEN-DATENBANK UND BEFUND-DATENBANK
;
;                  			 + FUZZY STRING SEARCH UND NORMALE SUCHFUNKTIONEN FÜR OBJEKTBASIERTEN DATENBANK
;                                                                                  	------------------------
;                                                	FÜR DAS AIS-ADDON: "ADDENDUM FÜR ALBIS ON WINDOWS"
;                                                                                  	------------------------
;    		BY IXIKO STARTED IN SEPTEMBER 2017 - LAST CHANGE 15.02.2021 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                           	|                                          	|                                          	|                                          	|
; ▹PatDb                                 	PatDir                                    	class admDB

; ▹ReadGVUListe                   	IstChronischKrank           	    	InChronicList                  	    	IstGeriatrischerPatient

; ▹class AlbisDatenbanken		Leistungskomplexe            		DBASEStructs                       		GetDBFData
; 	 ReadPatientDBF					ReadDBASEIndex

; ▹StrDiff                                	DLD                                      	FuzzyFind

; ▹RxNames					    		EscapeStrRegEx            	    		class string

; ▹GetAlbisPath

;----------------------------------------------------------------------------------------------------------------------------------------------------------------------
;###############################################################################
;----------------------------------------------------------------------------------------------------------------------------------------------------------------------

;----------------------------------------------------------------------------------------------------------------------------------------------------------------------
; PATIENTEN DATENBANK INKL. SUCHFUNKTIONEN (Textdatei im Addendum Ordner)
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
PatDb(Pat, cmd:="")                                                       	{                 	;-- überprüft die Addendum Patientendatenbank und führt auch das alternative Tagesprotokoll

	; letzte Änderung: 13.02.2021

			If !IsObject(Pat)
				Exceptionhelper(A_ScriptName, "PatDb(Pat,cmd:="") {", "Die Funktion hat einen Fehler ausgelöst,`nweil kein Objekt übergeben wurde.", A_LineNumber)

		; leere PatID Übergabe vermeiden
			PatID := Pat.ID
			If (!RegExMatch(PatID, "^\d+$") || StrLen(PatID) = 0) 									;abbrechen falls keine PatID ermittelt werden kann
				return

		; Nn - Nachname, Vn - Vorname, Gt - Geschlecht, Gd - Geburtsdatum, Kk - Krankenkasse
			If !oPat.Haskey(PatID) {
				oPat[PatID] := {"Nn": Pat.Nn, "Vn": Pat.Vn, "Gt": Pat.Gt, "Gd": Pat.Gd, "Kk": Pat.Kk}
				admDB.SavePatDB(oPat, Addendum.DBPath "\Patienten.json")
				If Addendum.ShowTrayTips
					TrayTip, Addendum, % "neue PatID (Zähler: " oPat.Count() ") für die Addendumdatenbank:`n(" PatID ") " Pat.Nn "," Pat.Vn, 1
			}

		; ermittelt ob diese Patientenakte heute schon aufgerufen wurden
			For Index, TProtID in TProtokoll
				 If (TProtID = PatID)
					return 0

		; PatID wird dem Tagesprotokoll hinzugefügt und gespeichert
			TProtokoll.Push(PatID)
			IniAppend("<" PatID "(" A_Hour ":" A_Min ")>", Addendum.TPFullPath, SectionDate, compname)

return 1
}

PatDir(PatID)                                                                 	{               	;-- berechnet den Addendum-Datenpfad für einen Patienten

	; Addendum speichert zusätzliche Daten zum Patienten in einem extra Ordnerverzeichnis nach folgendem Schema:
	; 	- die Sortierung erfolgt anhand der Patienten ID von Albis
	;	- maximal 1000 Unterverzeichnisse pro Ordner, ein Ordner mit der Bezeichnung 8000 enthält alle Patientenunterordner zwischen 7001-8000
	;
	; die Funktion legt noch nicht vorhandene Ordner/Unterordner automatisch an

	BaseNum	:= Round((PatID / 1000) ) * 1000
	IDBase  	:= (PatID - BaseNum <= 0) ? BaseNum : BaseNum + 1000
	PatientDir 	:= Addendum.DBPath "\PatData\" IDBase "\" PatID
	If !InStr(FileExist(PatientDir), "D")
			FileCreateDir, % PatientDir

return PatientDir
}

class admDB                                                                 	{                 	;-- AddendumDB - Klasse für DB (Objekt) Handling (im Moment nur Suche)

	; das globale oPat-Objekt wird benötigt (wird per )

	; lädt Daten
		ReadPatDB(PatDBPath:="") {

			If !RegExMatch(PatDBPath, "\.json$")
				PatDBPath := Addendum.DBPath "\Patienten.json"

			this.PatDBPath := PatDBPath
			If !FileExist(PatDBPath)
				PatDB := ""
			else
				PatDB := JSONData.Load(PatDBPath,, "UTF-8")

		return PatDB
		}

	; speichert von Daten
		SavePatDB(PatDB, PatDBPath:="" ) {

			If !RegExMatch(PatDBPath, "\.json$")
				PatDBPath := Addendum.DBPath "\Patienten.json"

			For PatID, Pat in PatDB {
				If !Pat.HasKey("Nn") || !Pat.HasKey("Vn") || !Pat.HasKey("Gd") || !Pat.HasKey("Gt") || !Pat.HasKey("Kk") {
					throw A_ThisFunc ": Das übergebene Objekt enthält entweder keine Patientendaten oder ist defekt.`nPatID: " PatID
					return
				}
			}

			JSONData.Save(PatDBPath, PatDB, true,, 1, "UTF-8")

		}

	; GetExactMatches
		GetExactMatches(getString:="PatID|Gd", OnlyFirstMatch:=0, Nn:="", Vn:="", Gt:="", Gd:="", Kk:="") {

			matches	:= Object()
			getter    	:= StrSplit(getString, "|")

			For PatID, Pat in oPat
				If InStr(Pat["Nn"], Nn) && InStr(Pat["Vn"], Vn) && InStr(Pat["Gt"], Gt) && InStr(Pat["Gd"], Gd) && InStr(Pat["Kk"], Kk)			{
					Loop % getter.MaxIndex()
						If InStr(getter[A_Index], "PatID")
							matches[getter[A_Index]] := PatID
						else
							matches[getter[A_Index]] := Pat[getter[A_Index]]

						If OnlyFirstMatch
							break
					}

		return matches
		}

	; Patienten ID Suche per Stringvergleich
		MatchID(key, val) {

			; key:	kann sein Nn, Vn, Gd, Kasse
			; val: 	zu suchender Eintrag
			matches := Object()

			For PatID, Pat in oPat
				If (Pat[key] = val)
					matches[PatID] := Pat

		return matches
		}

	; Patienten ID Suche per String-Similarity Funktion
		StringSimilarityID(name1, name2, diffmin=0.09) {

			PatIDArr	:= Object()
			minDiff 	:= 100
			NVname 	:= RegExReplace(name1 . name2, "[\s]")
			VNname 	:= RegExReplace(name2 . name1, "[\s]")

			For PatID, Pat in oPat 		{

				If InStr(PatID, "MaxPat")
					break

				DbName	:= RegExReplace(Pat.Nn Pat.Vn, "[\s]")
				DiffA     	:= StrDiff(DBName, NVname)
				DiffB     	:= StrDiff(DbName, VNname)
				Diff        	:= DiffA <= DiffB ? DiffA : DiffB

				If (Diff <= diffmin)
					PatIDArr.Push(PatID)

				If (Diff < minDiff)
					minDiff	:= Diff, bestDiff := PatID, bestNn := Pat.Nn, bestVn := Pat.Vn, bestGd := Pat.Gd

			}

		return PatIDArr
		}

	; erweitertete Patienten ID Suche
		StringSimilarityEx(name1, name2, diffmin:=0.09) {

			PatIDs    	:= Object()
			minDiff 	:= 100
			NVname 	:= RegExReplace(name1 . name2, "[,.\-\s]")
			VNname 	:= RegExReplace(name2 . name1, "[,.\-\s]")

			For PatID, Pat in oPat 		{

				DbName	:= RegExReplace(Pat.Nn Pat.Vn, "[\s]")
				If (StrLen(dbName) = 0)
					continue

				DiffA	:= StrDiff(DBName, NVname), DiffB := StrDiff(DbName, VNname)
				Diff  	:= DiffA <= DiffB ? DiffA : DiffB

				If (Diff <= diffmin)
					PatIDs[PatID] := Pat

				If (Diff < minDiff)
					minDiff	:= Diff, bestDiff := PatID, bestNn := Pat.Nn, bestVn := Pat.Vn, bestGd := Pat.Gd

			}

			PatIDs.Diff := {"minDiff":minDiff, "bestID":bestDiff, "bestNn":bestNn, "bestVn":bestVn, "bestGd":bestGd}

			If (PatIDs.Count() > 1)
				return PatIDs

		return
		}

	; Sift-Ngram-Suche
		SiftNgram(Name) {

			static Haystack, Haystack_Init := true
			found 	:=Object()
			Needle	:= Array()

			If !IsObject(oPat)
				MsgBox, Kein Objekt!

			If Haystack_Init		{
				Haystack_Init := false
				For PatID, Pat in oPat
					Haystack .= Pat.Nn " " Pat.Vn "`n"
			}

			Needle.1	:= Name[3] " " Name[4]
			Needle.2	:= Name[4] " " Name[3]

		; from https://www.autohotkey.com/boards/viewtopic.php?f=76&t=28796 - Getting closest string
			Loop 2		{
				for key, element in Sift_Ngram(Haystack, Needle[A_Index], 0,, 2, "S")
					If (element.delta > 0.90) {
						;matches := this.MatchID()
						;found.Push(GetFromPatDb("PatID|Nn|Vn|Gd", 1, StrSplit(element.data, " ").1, StrSplit(element.data, " ").2))
					}

			}

	return idx = 0 ? 0 : found
	}

}

;}

;----------------------------------------------------------------------------------------------------------------------------------------------------------------------
; Sonstige Datenlisten
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
ReadGVUListe(path, Quartal) {                                                            	;-- Einlesen der manuell angelegten untersuchten Patienten

	If !FileExist(path "\" Quartal "-GVU.txt")
		return

	GVU := Object()
	For index, line in StrSplit(FileOpen(path "\" Quartal "-GVU.txt", "r").Read(), "`n", "`r") {

		If (StrLen(Line) = 0)
			continue

		UDatum	:= StrSplit(Line, ";").1
		UQuartal	:= StrReplace(StrSplit(Line, ";").2, "/")
		UPatID   	:= StrSplit(Line, ";").3
		URdyTxt 	:= StrSplit(Line, ";").4
		UReady 	:= StrLen(StrSplit(Line, ";").4) = 0 ? false : true

		If (UQuartal <> Quartal)
			continue

		GVU[UPatID] := {"UQuartal":UQuartal, "UDatum":UDatum, "UReady":UReady, "URdxTxt":URdyTxt}

	}

return GVU
}

IstChronischKrank(PatID)    	{

	For listID, PatIDchr in Addendum.Chroniker
		If (PatIDchr = PatID)
			return true

return false
}

InChronicList() {                                                                                	;-- Neuaufnahme in Chroniker Liste

		GruppenName := "Chroniker"

		PatID := AlbisAktuellePatID()

		For key, ChronikerID in Addendum.Chroniker
			If (PatID = ChronikerID)
				return

	; Nutzer abfragen ob Pat. aufgenommen werden soll
		hinweis := "Pat: " oPat[PatID].Nn ", " oPat[PatID].Vn ", geb. am: " oPc[PatID].Gd "`n"
		hinweis .= "ist nicht als Chroniker vermerkt.`nMöchten Sie automatisch alle Eintragungen`ninnerhalb von Albis vornehmen lassen?"

		MsgBox, 0x1024, Addendum für Albis on Windows, % hinweis, 10
		IfMsgBox, Yes
		If (PatID = AlbisAktuellePatID())		{
				ChronCb	:= false
				Indikation	:= false
				PatGruppe:= false

			; Ziffer der Liste hinzufügen und in die Datei speichern
				Addendum.Chroniker.Push(PatID)
				FileAppend, % PatID "`n", % Addendum.DBPath "\DB_Chroniker.txt"

			; Patientengruppierung vornehmen  ;{

				failed := false

				Albismenu("34362", "Patientengruppen für ahk_class #32770")
				ControlGet, result, List, , % "SysListView321", % "Patientengruppen für ahk_class #32770"

				If !InStr(result, GruppenName) {

							PraxTT("Patient ist nicht der Chronikergruppe zugeordnet", "3 2")

						; click auf NEU
							VerifiedClick("Button2", "Patientengruppen für ahk_class #32770")
							Sleep, 300
							WinWait, % "Patientengruppen ahk_class #32770",, 2

						; Click hat funktioniert
							If WinExist("Patientengruppen ahk_class #32770") {

								; vorhandene Gruppeneinträge auslesen und Position der Chronikergruppe finden
									hwnd := WinExist("Patientengruppen ahk_class #32770")
									ControlGet, result, List, , % "Listbox1", % "ahk_id " hwnd

									Loop, Parse, result, `n
										If InStr(A_LoopField, GruppenName) {
											ListboxRow := A_Index
											break
										}

									If (ListBoxRow > 0) 	{

										; den Listboxeintrag mit der Chronikergruppe auswählen
											VerifiedChoose("Listbox1", "ahk_id " hwnd, ListBoxRow)
											VerifiedClick("Button1",  "ahk_id " hwnd)

											PatGruppe := true

									} else
											failed := true

							} else
    								failed := true

				}

			; Fenster 'Patientengruppe für' schließen
				If WinExist("Patientengruppen für ahk_class #32770") {

						If failed {

						; drückt auf Abbrechen
							VerifiedClick("Button7", "Patientengruppen für ahk_class #32770", "", "", true)
							Sleep 200
							If WinExist("Patientengruppen für ahk_class #32770")
								WinClose, % "Patientengruppen für ahk_class #32770"

						} else {

						; OK - Button
							VerifiedClick("Button1", "Patientengruppen für ahk_class #32770", "", "", true)

						}

				}

			;}

			; Chroniker Häkchen setzen und bei Ausnahmeindikation (weitere Informationen) diese Ziffer hinzufügen ;{

				failed        	:= false
				hPersonalien := Albismenu("32774", "Daten von ahk_class #32770") ; Menu 'Personalien'

				If hPersonalien {

						If VerifiedCheck("Chroniker"                      	, "ahk_id " hPersonalien)
							ChronCb := true

					; weitere Information für nächsten Dialog drücken
						  VerifiedClick("Weitere In&formationen..."	, "ahk_id " hPersonalien)
						WinWait, % "ahk_class #32770", % "Adresse des", 2

						If (hInformationen := WinExist("ahk_class #32770 ahk_exe albis", "Ausnahmeindikation")) {

							; Feldinhalt wird ausgelesen, bei fehlender Eintragung wird die Ziffer hinzugefügt
								ControlGetText, ctext, Edit14, % "ahk_id " hInformationen
								If !InStr(ctext, "03220") {

										ctext := LTrim(ctext "-03220", "-")

										If VerifiedSetText("Edit14", cText, "ahk_class #32770", 200, "Ausnahmeindikation")
											indikation := true

										ControlFocus	, 					  % "Edit14", % "ahk_id " hInformationen
										ControlSend	, % "{Enter}"	, % "Edit14", % "ahk_id " hInformationen

								}

							; Fenster schliessen falls noch geöffnet
								If WinExist("ahk_class #32770", "Ausnahmeindikation") {
									ControlClick, Ok	, % "ahk_id " hInformationen
									WinWaitClose		, % "ahk_id " hInformationen,, 2
								}

						}

						; Fenster "Daten von " schliessen
							If WinExist("Daten von", "Anrede")
								VerifiedClick("Button30", "ahk_id " hPersonalien)
				}

			;}

				processed := "1. Chronikergruppe     : " (PatGruppe 	? "hinzugefügt" 	: "nicht hinzugefügt") "`n"
				processed .= "2. Chronikerhäkchen   : " (ChronCb 	? "gesetzt"     	: "nicht gesetzt") "`n"
				processed .= "2. Ausnahmeindikation: " (indikation 	? "Lk eingefügt"	: "Lk nicht eingefügt") "`n"
				PraxTT("Eingruppierung abgeschlossen.`n" processed, "6 3")

		}

	; PraxTT("Pat. konnte keiner Gruppe zugeordnet werden.`nSetze Funktion mit anderen Eintragungen fort!", "3 3")
	;PraxTT("Pat. konnte keiner Gruppe zugeordnet werden.`nSetze Funktion mit anderen Eintragungen fort!", "3 3")
}

IstGeriatrischerPatient(PatID)	{

	For listID, PatIDGB in Addendum.Geriatrisch
		If (PatIDGB = PatID)
			return true

return false
}
;}

;----------------------------------------------------------------------------------------------------------------------------------------------------------------------
; ALBIS DBASE DATENBANKEN - benötigt Addendum-DBASE.ahk
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
class AlbisDb { 		                             	 					         								;-- erweitert Addendum_DBASE um zusätzliche Funktionen

		; diese Klasse benötigt Addendum_DBASE, Addendum_Internal

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Initialisieren
	;----------------------------------------------------------------------------------------------------------------------------------------------
		__New(DBasePath:="", debug:=0) {

					this.debug 		:= debug

			; Albisdatenbankpfad oder ein alternativer Pfad für Testversuche
				If !DBasePath
					this.DBPathAlbis      	:= GetAlbisPath() "\db"
				else
					this.DBPathAlbis      	:= DBasePath
				If !FileExist(this.DBPathAlbis "\" this.dbName ".dbf")
					throw A_ThisFunc 	": Die Datenbankdatei <" this.dbName ">`n"
											.	"befindet sich nicht im Ordnerpfad:`n"
											.	"<" this.DBPathAlbis ">"

			; Addendum Verzeichnis und Datenbankverzeichnis
				RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)
				IniRead, admDBPath, % AddendumDir "\" AddendumIni, Addendum, AddendumDBPath
				admDBPath := StrReplace(admDBPath, "%AddendumDir%", AddendumDir)
				this.DBPathAddendum	:= AddendumDir "\logs'n'data\_DB"
				If !InStr(FileExist(this.DBPathAddendum "\"), "D") {
					throw A_ThisFunc 	": Addendum Pfad oder die Addendum.ini Datei ist nicht korrekt!`n"
											.	"Hinweis:  Falls noch nicht gemacht. Führen Sie zunächst das AddendumStarter "
											.	"Skript aus, damit grundlegende Einstellungen, sowie Unterverzeichnisse im "
											.	"Addenumhauptverzeichnis angelegt werden. Rufen Sie danach wieder diese Funktion auf."
				}


		}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Funktionen für Datenauswertungen
	;----------------------------------------------------------------------------------------------------------------------------------------------

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Hilfsfunktionen
	;----------------------------------------------------------------------------------------------------------------------------------------------
		CreateDatabaseIndex(dbName, IndexField:="", ReIndex:=true) {       	;-- Indexerstellung für Albis DBase Datenbanken

			/* BESCHREIBUNG

				Die Funktion erleichtert die Erstellung eines Datenbankindex durch hinterlegte Feldnamen, welche zur Indizierung herangezogen werden.
				Ist kein Feldname für eine Datenbank hinterlegt, kann diesen über den Parameter IndexField übergeben.
				Der übergebene Feldname wird von der Funktion vorrangig behandelt!
				Ist ein Feldname nicht vorhanden wird eine Fehlermeldung ausgeworfen.

			 */

			; Variablen                                                                    	;{
				static IndexFields := {"BEFUND" 	: "DATUM"    	, "BEFTEXTE"	: "DATUM"				, "LABORANF"	: "EINDATUM"
											, 	"LABBLATT"	: "DATUM"    	, "LABBUCH"	: "ABNAHMDATE"	, "LABREAD"  	: "EINDATUM"
											, 	"LABUEBPA"	: "ABDATUM"}
			;}

			; Dateinamen überprüfen                                               	;{
				dbName 	:= RegExReplace(dbName, "i)\.dbf$")
				If !FileExist(this.DBPathAlbis "\" dbName ".dbf")
					throw A_ThisFunc 	": Die Datenbankdatei <" dbName ">`n"
											.	"befindet sich nicht im Ordnerpfad:`n"
											.	"<" this.DBPathAlbis ">"
			;}

			; Verzeichnis für Daten zu den DBase Dateien anlegen    	;{
				If (this.CreateFilePath(this.DBPathAddendum "\DBase") != 0)
					throw A_ThisFunc ": Das Anlegen des Dateipfades [" this.DBPathAddendum "\DBase\] ist fehlgeschlagen!"
				this.DBaseIndexPath := this.DBPathAddendum "\DBase"
			;}

			; Information auslesen und Indizierbarkeit prüfen             	;{
				this.db := new DBASE(this.DBPathAlbis "\" dbName ".dbf", this.debug)
				If (StrLen(IndexField) > 0) {

					If !this.FieldNameExist(IndexField)
						this.ThrowError(	"Die Datenbank ["  dbName "] kann nicht indiziert werden!`n"
											.	"  Der Feldnahme [" IndexField "] ist nicht enthalten.")

				}
				else {

					If !IndexFields.HasKey(dbName)
						this.ThrowError(	"Die Datenbank ["  dbName "] kann nicht indiziert werden!`n"
											.	"  Zur Datenbank sind keine Information verfügbar.")

					IndexField := IndexFields[this.dbName]

					If !this.FieldNameExist(IndexField)
						this.ThrowError(	"Die Datenbank ["  dbName "] kann nicht indiziert werden!`n"
						.	"  Der voreingestellte Feldnahme [" IndexField "] existiert in dieser Datenbank nicht.")

				}
			;}

			; Index erstellen (wird als JSON String gespeichert)
				dbIndex := this.db.CreateIndex("quarter", this.DBaseIndexPath "\DBIndex_" this.dbName ".json", IndexField, ReIndex)
				this.db.CloseDBF()
				this.db := ""

		return dbIndex
		}

		CreateFilePath(path) {                                                      	;-- erstellt einen Dateipfad falls dieser noch nicht existiert

			If !InStr(FileExist(path "\"), "D") {
				FileCreateDir, % path
				If ErrorLevel
					return A_LastError
				else
					return 0
			}

		return 0
		}

		FieldNameExist(fieldLabel) {                                              	;-- ist der übergebene Feldname vorhanden?

			idxFieldExist := false

			If !IsObject(this.db.dbfields)
				this.ThrowError(	"Fehler beim Auslesen von Information aus ["  this.dbName "]!`n"
								. 	"Informationen zur Datenbankstruktur konnten nicht erzeugt werden..")

			For fLabel, fData in this.db.dbfields
				If (fLabel = fieldLabel) {
					idxFieldExist := true
					break
				}

		return idxFieldExist
		}

		ThrowError(msg) {	                                                        	;-- "wirft" Fehlermeldungen aus
			this.db.CloseDBF()
			this.db := ""
			throw A_ThisFunc ": " msg
		}

}


DBASEStructs(AlbisPath, DBStructsPath:="") {                                         	;-- analysiert alle DBF Dateien im albiswin\db Ordner

	; schreibt die ausgelesenen Strukturen als JSON formatierte Datei

		dbfStructs := Object()
		AlbisDBpath:= AlbisPath "\db"

	; zählt für die Dateianzeige zunächst einmal die vorhandenen Dateien im Verzeichnis
		Loop, Files, % AlbisDBpath "\*.dbf"
			dbfStructs[StrReplace(A_LoopFileName, ".dbf")] := Object()

	; für formatierte Ausgabe
		dbFilesMax := dbfStructs.Count()
		dbIL := StrLen(dbFilesMax) - 1

	; öffnet jede DBASE Datei und liest den Header der Datei aus
		dbNr := 0
		For dbfName, struct in dbfStructs {

			ToolTip, % "lese: " SubStr("000" (dbNr ++), -1 * dbIL) "/" dbFilesMax ": " dbfName

			FileGetSize, SizeDBF	, % AlbisDBPath "\" dbfName ".dbf"		, K
			FileGetSize, sizeDBT	, % AlbisDBPath "\" dbfName ".dbt" 	, K
			FileGetSize, sizeMDX	, % AlbisDBPath "\" dbfName ".mdx"	, K
			dbfSizeMax += sizeDBF
			dbtSizeMax += sizeDBT
			mdxSizeMax += sizeMDX

			dbfStructs[dbfShort].fields     	:= Object()
			dbfStructs[dbfShort].header   	:= Object()

			dbf  	:= new DBASE(AlbisDBPath "\" dbfName ".dbf", false)
			dbfStructs[dbfName].fields     	:= isObject(dbf.dbfields) ? dbf.dbfields : "error reading dbase file"
			dbfStructs[dbfName].header  	:= isObject(dbf.dbstruct) ? dbf.dbstruct : "error reading dbase file"
			dbf.Close()
			dbf := ""
		}
		ToolTip

	; speichern der Strukturdaten
		If RegExMatch(DBStructsPath, "[A-Z]\:\\.*\.json") {

			file := FileOpen(DBStructsPath, "w", "UTF-8")
			file.Write("; -------------------------------------------------------------------------------------------------`n")
			file.Write("; ***** Informationen DBASE Dateien *****"														                        	"`n")
			file.Write("; Verzeichnis  `t: " 	AlbisPath "\db"		                                                                                    	"`n")
			file.Write("; Erstellt am `t: " 	A_DD "." A_MM "." A_YYYY " um " A_Hour ":" A_Min "Uhr	                            	 `n")
			file.Write("; Dateien      `t: " 	dbFilesMax 																		                        	"`n")
			file.Write("; Dateigrößen`t" 	 																						                        	"`n")
			file.Write(";   .dbf       `t: " 		Round(dbfSizeMax/1024/1024	, 2)	" GB" 						                        	"`n")
			file.Write(";   .dbt       `t: " 		Round(dbtSizeMax/1024/1024	, 2)	" GB" 						                        	"`n")
			file.Write(";   .mdx       `t: " 	Round(mdxSizeMax/1024/1024, 2)	" GB" 		    			                        	"`n")
			file.Write(";   Summe     `t: " 	Round((dbfSizeMax+dbtSizeMax+mdxSizeMax)/1024/1024, 2) " MB"    			"`n")
			file.Write("; -------------------------------------------------------------------------------------------------`n")
			file.Write(JSON.Dump(dbfStructs,,2))
			file.Close()

	}

}

GetDBFData(DBpath,p="",out="",s=0,opts="",dbg=0,dbgOpts="") {     	;-- holt Daten aus einer beliebigen Datenbank

	; p     	- pattern as object
	; out  	- returned values as object
	; s     	- seek to position

	dbfile      	:= new DBASE(DBpath, dbg, dbgParams)
	startrec 	:= (s > 0 ? (s - (dbfile.headerlen + 1)) / dbf.lendataset : 0)

	res        	:= dbfile.OpenDBF()
	matches 	:= dbfile.Search(p, startrec, opts)
	res         	:= dbfile.CloseDBF()
	dbfile    	:= ""

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

Leistungskomplexe(PatID=0, StartQuartal="2009-4", SaveToPath="") {	;-- sammelt, erstellt eine Datenbank für den Abrechnungshelfer

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

		static Abr	:= { "Z01740":"1x im Arztfall"
        					,	"Z01732":"alle 3 Jahre"
        					,	"Z01745":"alle 3 Jahre"
        					,	"Z01746":"alle 3 Jahre"
        					,	"Z03360":"2x im Behandlungsfall"
        					,	"Z03362":"1x im Quartal,nicht neben 03360"
        					,	"Z03220":"Chroniker"
        					,	"Z03221":"wenn 03220 abgerechnet"}

		static 30Spaces := "                                   "
		static 8Spaces := "        "

		PatAbr := Object()

	; Quartalseinsprünge laden 'BEFUND.dbf'
		BefundDBF := JSON.Load(FileOpen(Addendum.DBpath "\PatData\BEFUNDDBF.json", "r").Read())

	; Daten von allen Patienten über alle Jahre sammeln
		befund  	:= new DBASE(Addendum.AlbisDBPath "\BEFUND.dbf", 0)
		filepos  	:= befund.OpenDBF()
		startpos  	:= Floor((BefundDBF[StartQuartal]-befund.recordsStart)/befund.LenDataSet)
		ziffern    	:= befund.SearchFast({"KUERZEL": "lko"}, ["DATUM", "PATNR", "INHALT"], startpos)
		filepos  	:= befund.CloseDBF()

	; Abr Ziffern mit gefundenem vergleichen
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
		AbrDBF := FileOpen(AddendumDir "\logs'n'data\_DB\PatData\ZIFFERN.adm", "w", "UTF-8")
		Loop % MAXNR {

			NR	:= A_Index
			PNR	:= SubStr("         " A_Index, -8)                                                                       ; PatNR                                                 	-	9 Zeichen
			Q 	:= !PatAbr[NR].HasKey("Q") ? 30Spaces : PatAbr[NR]["Q"]	                         	; die letzten 5 Quartale des Patienten      	- 30 Zeichen
			Z1 	:= !PatAbr[NR].HasKey("Z03220") ? 30Spaces 	: PatAbr[NR]["Z03220"] 		; abgerechnete Chronikerziffer 5 Quartale - 30 Zeichen
			Z2 	:= !PatAbr[NR].HasKey("Z03221") ? 30Spaces 	: PatAbr[NR]["Z03221"] 		; abgerechnete Chronikerziffer 5 Quartale - 30 Zeichen
			Z3 	:= !PatAbr[NR].HasKey("Z03360") ? 30Spaces   	: PatAbr[NR]["Z03360"] 		; abgerechnete GB 5 Quartale - 30 Zeichen
			Z4 	:= !PatAbr[NR].HasKey("Z03362") ? 30Spaces   	: PatAbr[NR]["Z03362"] 		; abgerechnete GB 5 Quartale - 30 Zeichen
			Z5 	:= !PatAbr[NR].HasKey("Z01732") ? 8Spaces   	: PatAbr[NR]["Z01732"]
			Z6 	:= !PatAbr[NR].HasKey("Z01740") ? 8Spaces   	: PatAbr[NR]["Z01740"]
			Z7 	:= !PatAbr[NR].HasKey("Z01745") ? 8Spaces   	: PatAbr[NR]["Z01745"]
			Z8 	:= !PatAbr[NR].HasKey("Z01746") ? 8Spaces   	: PatAbr[NR]["Z01746"]

			;ToolTip, % NR ": " PatAbr[NR]["Z03220"]

			dataset := PNR "|" Q "|" Z1 "|" Z2 "|" Z3 "|" Z4 "|" Z5 "|" Z6 "|" Z7 "|" Z8 "`n"
			If (A_Index = 1) {
				LenDataSet :=
				header := SubStr("    " StrLen(dataset), -4)
							. 	"|Q"          	SubStr("   " StrLen(Q)	, -2)
							. 	"|Z03220" 	SubStr("   " StrLen(Z1)	, -2)
							.	"|Z03221" 	SubStr("   " StrLen(Z2)	, -2)
							.	"|Z03360" 	SubStr("   " StrLen(Z3)	, -2)
							.	"|Z03362" 	SubStr("   " StrLen(Z4)	, -2)
							.	"|Z01732" 	SubStr("   " StrLen(Z5)	, -2)
							.	"|Z01740" 	SubStr("   " StrLen(Z6)	, -2)
							.	"|Z01745" 	SubStr("   " StrLen(Z7)	, -2)
							.	"|Z01746" 	SubStr("   " StrLen(Z8)	, -2)
							.	"|"

				Loop % (StrLen(dataset) - StrLen(header))
					adds .= "."

				AbrDBF.Write(header adds "`n")
			}

			AbrDBF.Write(dataset)

		}

		AbrDBF.Close()

return PatAbr
}

ReadPatientDBF(basedir="", infilter="", outfilter="", debug=0) {           	;-- gibt nur benötigte Daten der albiswin\db\PATIENT.DBF zurück

	; basedir kann man leer lassen, der Albishauptpfad wird aus der Registry gelesen
	; Rückgabeparameter ist ein Objekt mit Patienten Nr. und dazugehörigen Datenobjekten (die key's sind die Feldnamen in der DBASE Datenbank)

		PatDBF := Object()

	; kein basedir - sollte ..\albiswin\db\ enthalten dann wird hier versucht den Pfad aus der Windows Registry zu bekommen
		If (StrLen(basedir) = 0) || !InStr(FileExist(basedir), "D")
			basedir	:= GetAlbisPath() "\db"

	; für Abrechnungsüberprüfungen die geschätzt minimal notwendigste Datenmenge
		If !IsObject(infilter)
			infilter := ["NR", "NAME", "VORNAME", "GEBURT", "MORTAL", "LAST_BEH"]

	; liest alle Patientendaten in ein temp. Objekt ein
		database 	:= new DBASE(basedir "\PATIENT.dbf", debug)
		res        	:= database.OpenDBF()
		matches	:= database.GetFields(infilter)
		res         	:= database.CloseDBF()

	; temp. Objekt wird nach PATNR sortiert
		For idx, m in matches {

			strObj	:= Object()
			For key, val in m
				If (key <> "NR") && (StrLen(val) > 0)
					strObj[key] := val

			PatDBF[m.NR] := strObj

		}

return PatDBF
}

ReadDBASEIndex(admDBPath, DBASEName, ReIndex:=true) {               	;-- liest erstellte Indizes oder bei Bedarf eine Indexdatei

		DBASEName 	:= RegExReplace(DBASEName, "i)\.dbf$")
		DBFDataPath 	:= admDBPath "\DBase"
		DataFileName 	:= "DBIndex_" DBASEName ".json"
		DataFilePath		:= DBFDataPath "\" DataFileName

	; ruft die Indizierungsfunktion auf falls der Index noch nicht erstellt wurde
		If FileExist(DataFilePath)
			dbIndex	:= JSON.Load(FileOpen(DataFilePath, "r", "UTF-8").Read())
		else {

			AlbisDBPath := GetAlbisPath() "\db"
			TrayTip, Addendum für Albis on Windows, % "Die Datenbank " DBASEName " wird gerade indiziert."
			AlbisDB 	:= new AlbisDb()
			dbIndex	:= AlbisDB.CreateDatabaseIndex(DBASEName,, ReIndex)
			AlbisDB 	:= ""

		}

return dbIndex
}

;}

;----------------------------------------------------------------------------------------------------------------------------------------------------------------------
; FUZZY SUCHFUNKTIONEN
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------;{

; ---- FUZZY FUNCTIONS ---
; E - Ende des Wortes, A - Anfang des Wortes, M - in der Mitte des Wortes
; häufige Vertipper:
/*
zu schnell gedrückt:
ee 	-> e 	(E)   	-	Renee
ss 	-> s 		(E)		-	Pauluss

zwei Tasten gleichzeitig:
tr 		-> t 		(M)	-	Ritra (Rita)

zu schnell gedacht:
ile		-> iel	(M)	-	Danilea (Daniela)

phonetisch unklar:
ey 	-> ay 	(E)		-

wie wirds geschrieben?
dt 	-> d		(E)		-	Herr Brandt oder Brand?
*/

; phonetisch mögliche Ersetzungen für genaueren Vergleich:
/*
ue -> ü
ei, ai, ey, ay ​ 			in Leim, Mais, Speyer, Mayer
eu, äu 						​ in Heu, Läufer

*/

StrDiff(str1, str2, maxOffset:=5) {											            	    	;-- SIFT3 : Super Fast and Accurate string distance algorithm, Nutze ich um Rechtschreibfehler auszugleichen

	/*                              	DESCRIPTION

			By Toralf:
			Forum thread: http://www.autohotkey.com/forum/topic59407.html
			Download: https://gist.github.com/grey-code/5286786

			Basic idea for SIFT3 code by Siderite Zackwehdex
			http://siderite.blogspot.com/2007/04/super-fast-and-accurate-string-distance.html

			Took idea to normalize it to longest string from Brad Wood
			http://www.bradwood.com/string_compare/

			Own work:
			    - when character only differ in case, LSC is a 0.8 match for this character
			    - modified code for speed, might lead to different results compared to original code
			    - optimized for speed (30% faster then original SIFT3 and 13.3 times faster than basic Levenshtein distance)

			Dependencies. None

	*/

	if (str1 = str2)
		return (str1 == str2 ? 0/1 : 0.2/StrLen(str1))
	if (str1 = "" || str2 = "")
		return (str1 = str2 ? 0/1 : 1/1)
	StringSplit, n, str1
	StringSplit, m, str2
	ni := 1, mi := 1, lcs := 0
	while ((ni <= n0) && (mi <= m0))
	{
			if (n%ni% == m%mi%)
				lcs += 1
			else if (n%ni% = m%mi%)
				lcs += 0.8
			else {
				Loop, % maxOffset {
					oi := ni + A_Index, pi := mi + A_Index
					if ((n%oi% = m%mi%) && (oi <= n0)) {
						ni := oi, lcs += (n%oi% == m%mi% ? 1 : 0.8)
						break
					}
					if ((n%ni% = m%pi%) && (pi <= m0)) {
						mi := pi, lcs += (n%ni% == m%pi% ? 1 : 0.8)
						break
					}
				}
			}
			ni += 1
			mi += 1
	}
	return ((n0 + m0)/2 - lcs) / (n0 > m0 ? n0 : m0)
}

DLD(s, t) {																										;-- DamerauLevenshteinDistance
   m := strlen(s)
   n := strlen(t)
   if(m = 0)
      return, n
   if(n = 0)
      return, m
   d0_0 = 0
   Loop, % 1 + m
      d0_%A_Index% = %A_Index%
   Loop, % 1 + n
      d%A_Index%_0 = %A_Index%
   ix = 0
   iy = -1
   Loop, Parse, s
   {
      sc = %A_LoopField%
      i = %A_Index%
      jx = 0
      jy = -1
      Loop, Parse, t
      {
         a := d%ix%_%jx% + 1, b := d%i%_%jx% + 1, c := (A_LoopField != sc) + d%ix%_%jx%
            , d%i%_%A_Index% := d := a < b ? a < c ? a : c : b < c ? b : c
         if (i > 1 and A_Index > 1 and sc == tx and sx == A_LoopField)
            d%i%_%A_Index% := d < c += d%iy%_%ix% ? d : c
         jx++
         jy++
         tx = %A_LoopField%
      }
      ix++
      iy++
      sx = %A_LoopField%
   }
   return, d%m%_%n%
}

FuzzyFind(dict, query) {																					;-- Creates an array of match objects

	/*
		*
		  *
		   * LINK: https://github.com/bartlb/AHK_Libraries/blob/master/FuzzySearch/src/FuzzySearch.ahk
		   *
		   * Creates an array of match objects using a custom search technique based on a mix of
		   * fuzzy search and LCS algoritms.
		   *
		   * Before the search is initiated the query is split into individual characters, and then
		   * joined again by a regex wildcard -- though if query is blank or 'dict' is not an array
		   * the function will fail, returning 0. The search then begins by looping through the
		   * array of strings and attempting to validate a match by using a basic fuzzy search. If
		   * a match is found, the function will attempt to locate the best match for the query,
		   * rather than just submitting the first match (left-to-right); the 'best match' is
		   * determined by a weight system which holds matches at the begining of a string, new
		   * words (seen as characters separated by a standard word boundary '\b' or an underscore)
		   * and capitals as identifiers of a better match. The weight system is fairly basic. In
		   * order, matches are weighted as follows:
		   * 1. The percent of characters in the string that the query covers, as a whole number.
		   *    i.e. Round((QUERY_LENGTH / STRING_LENGTH) * 100)
		   * 2. The first character in the query matches the first character in a string (+50).
		   * 3. A character in the query matches the first letter of a new word in a string (+25).
		   * 4. A character in the quert matches any uppercase letter in a given string (+10).
		   * 5. Any other match (+1).
		   *
		   * @function  FuzzySearch
		   * @param    	[in]        	dict    	An array of strings to query against.
		   * @param    	[in]        	query   	Any query in the form of a string.
		   * @returns   	{object}          		An array of match objects consisting of the matched
		   *                              					string and information pertaining to where the match
		   *                              					occured within the string.
		  *
		*
	 */

  if (! IsObject(dict) || query == "")
    return 0
  matches := []
  for each, token in StrSplit(query)  {
    re_string .= (re_string ? ".*?" : "") "(" token ")"
  }
  for each, string in dict  {
    if (RegExMatch(string, "Oi)" re_string, re_obj)) {
      _match := {
      (Join
        "string": string,
        "tokens": [],
        "weight": Round((re_obj.Count() / StrLen(string)) * 100)
      )}
      loop % re_obj.Count()
      {
        m_offset    := re_obj.Pos(A_Index)
        m_restring  := (A_Index == 1 ? re_string : SubStr(m_restring, 7))
        _token      := { "id": "", "position": 0, "weight": 0 }

        while (RegExMatch(string, "Oi)" m_restring, m_obj, m_offset)) {
          _weight := m_obj.Pos(1) == 1 ? 50
                   : RegExMatch(SubStr(string, m_obj.Pos(1) - 1, 2)
                              , "i)(\b|(?<=_))" m_obj[1]) ? 25
                   : RegExMatch(SubStr(string, m_obj.Pos(1), 1), "\p{Lu}") ? 10 : 01
          if (_weight > _token.weight)
            _token.id := m_obj[1]
            , _token.weight   := _weight
            , _token.position := m_obj.Pos(1)

          m_offset := m_obj.Pos(1) + 1
        }
        _match.tokens.Insert(_token.position, _token.id)
        _match.weight += _token.weight
      }
      matches.Insert(_match)
    }
  }

  return matches
}

;}

;----------------------------------------------------------------------------------------------------------------------------------------------------------------------
; HILFSFUNKTIONEN - STRING HANDLING,
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
RegExStrings()                                                         	{                         	;-- erstellt die RegExString für die string Klassenbibliothek

	rxb := {"S"             	: "[\s,]+"
			, 	"Person1"  	: "[A-ZÄÖÜ][\pL]+[\s-]*([A-ZÄÖÜ][\pL-]+)*"
			, 	"Person2"  	: "[A-ZÄÖÜ][\pL]+([\-\s][A-ZÄÖÜ][\pL]+)*"
			,	"Date1"        	: "\d{1,2}\.\d{1,2}\.\d{2,4}"}

	rx := {	"fileExt"	    	: "\.[A-Za-z]{,3}$"
			, 	"umlauts"   	: {"Ä":"Ae", "Ü":"Ue", "Ö":"Oe", "ä":"ae", "ü":"ue", "ö":"oe", "ß":"sz"}
			, 	"Extract"      	: "^\s*(?<NN>" rxb.Person1 ")" rxb.S "(?<VN>" rxb.Person1 ")" rxb.S "(?<Title>.*)*"
			, 	"Names1"  	: "^\s*(?<NN>" rxb.Person1 ")" rxb.S "(?<VN>" rxb.Person1 ")" rxb.S
			, 	"Names2"  	: "^\s*(?<NN>" rxb.Person2 ")" rxb.S "(?<VN>" rxb.Person2 ")" rxb.S
			; Addendums Dateibezeichnungsregel für Befunde: z.B. "Muster-Mann, Max-Ali, Entlassungsbrief Kardiologie v. 10.10.2020-18.10.2020"
			, 	"CaseName"	: "^\s*(?<N>.*)?,\s*(?<I>.*)?\s+v\.\s*"
									. "(?<D1>\d+.\d+.\d*)(?<D2>\-\d+.\d+.\d+)*"
									. "(?<E>\s*\.[A-Za-z]+)*"
			; letztes Datum im Dateinamen (der letzte Datumsstring auf den weder ein Whitespace, Zahl, Punkt oder ein Minus folgt, oder am Ende des String steht
			,	"LastDate" 	: "(?<Date>\" rxb.Date1 ").?([^\d.\s\-]|$)"}

	For key, rxString in rxb
		rx[key] := rxString

return rx
}

class string                                                               	{                        	;-- wird alle String Funktionen von Addendum enthalten

	; benötigt Addendum_Datum.ak
	; letzte Änderung 17.02.2021

	; RegEx Strings erstellen
		static rx := RegExStrings()

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; diverse Stringfunktionen
	;----------------------------------------------------------------------------------------------------------------------------------------------
		EscapeStrRegEx(str)                           	{      	;-- formt einen String RegEx fähig um
			For idx, Char in StrSplit("[](){}*\/|-+.?^$")
				str := StrReplace(str, char , "\" char)
		return str
		}

		Flip(string)                                       	{        	;-- String reverse
			VarSetCapacity(new, n:=StrLen(string))
			Loop % n
				new .= SubStr(string, n--, 1)
		return new
		}

		Karteikartentext(str)                        	{        	;-- entfernt Namen, Dateierweiterung, Leerzeichen

			str := this.Replace.FileExt(str)                	; entfernen der Dateiendung
			str := this.Replace.Names(str)            	; Namen von Patienten entfernen
			str := RegExReplace(str, "^[\s,]*")        	; einzeln stehende Komma entfernen
		return RegExReplace(str, "\s{2,}", " ")       	; Leerzeichenfolgen kürzen
		}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Funktionen für die PDF Kategorisierung von Addendum-Gui.ahk
	;----------------------------------------------------------------------------------------------------------------------------------------------
		ContainsName(str)                         	{       	;-- enthält Personennamen?
			rxpos := RegExMatch(Str, this.rx.Names1)
			return (rxpos ? rxpos : RegExMatch(Str, this.rx.Names2))
		}

		isFullNamed(str)                            	{        	;-- Dateiname enthält Nachname,Vorname, Dokumentbezeichung und Datum

			;str := RegExReplace(str, this.rx.fileExt) ; neuer RegExString 'CaseName' kann mit Dateiendungen umgehen
			If RegExMatch(str, this.rx.CaseName, F) {
				Part1	:= RegExReplace(FN      	, "[\s,]+$")
				Part2	:= RegExReplace(FI       	, "[\s,]+$")
				Part3	:= RegExReplace(FD1 FD2, "^[\s,v\.]+")
				If (StrLen(Part1) > 0) && (StrLen(Part2) > 0) && (StrLen(Part3) > 0)
					return true
			}

		return false
		}

		class Get extends string                   	{        	;-- alle Stringextraktionen unter einem Namen

			DocDate(str)                                           	{       	;-- Dokumentdatum aus dem Dateinamen erhalten

				RegExMatch(str, this.rx.LastDate, Doc)
				If RegExMatch(DocDate, "^" this.rx.Date1 "$")
					return FormatDate(DocDate, "DMY", "dd.MM.yyyy")

			return
			}

			Names(str)                                             	{         	;-- Personennamen zurückgeben
				RegExMatch(Str, this.rx.Extract, Pat)
				return {"Nn":PatNN, "Vn":PatVN, "DokTitel":PatDokTitel}
			}

		}

		class Replace extends string             	{           ;-- alle Ersetzungsfunktionen unter einem Namen

			Names(str, rplStr:="")            	{         	;-- Personennamen ersetzen
				return RegExReplace(Str, this.rx.Names2, rplStr)
			}

			FileExt(str, ext:="", rplStr:="") 	{        	;-- Dateiendungen ersetzen
				ext := !ext ? "[a-z]+" : ext
				return RegExReplace(Str, "\." ext "$" , rplStr)
			}

		}

}

GetAlbisPath()                                                        	{                         	;-- liest das Albisinstallationsverzeichnis aus der Registry

		If (A_PtrSize = 8)
			SetRegView	, 64
		else
			SetRegView	, 32

		RegRead, PathAlbis, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-MainPath

		If (StrLen(PathAlbis) = 0)
			throw " Parameter 'basedir' enthält keine Pfadangabe"

return PathAlbis
}

;}



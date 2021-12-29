﻿;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                                                   Addendum_DB
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;
;	INHALTE:
;
;		+ einfache Datenbankfunktion für die Addendum interne Patientennamen und Befunddatenbank
;
; 		+ Fuzzy Stringmatching Algorithmus und einfache Suchfunktionen für die Datensuche in den Albis dBase Datenbanken
;
;		+ Klassenbibliothek als Erweiterung für Addendum_DBASE.ahk.
;				*	gedacht für die Auswertung von Patientendaten
;				*	nutzt ein hinterlegtes Regelsystem für die Erkennungen von Abrechnungspotentialen bei der Kassenabrechnung
;				*	derzeit vorwiegend für abrechnungsbasierte Auswertungen wie z.B. die quartalsweise Erfassungen aller Kassenscheine, aller
;				 	abgerechneten Leistungen oder den Behandlungstagen. Ermittlung aller Patienten mit abgerechneten
;					Chronikerpauschalen.
;              	*	Die Funktionen werden intensiv im Abrechnungsassistenten verwendet, welcher mittels regelbasierter Auswertung fehlende
;					Leistungskomplexe vorschlägt und auf Wunsch in der Patientenkarteikarte einträgt..
;
;
;                                                                                  	------------------------
;                                                	FÜR DAS AIS-ADDON: "ADDENDUM FÜR ALBIS ON WINDOWS"
;                                                                                  	------------------------
;    		BY IXIKO STARTED IN SEPTEMBER 2017 - LAST MODIFICATION 12.12.2021 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                           	|                                          	|                                          	|                                          	|
; ▹PatDb                                 	PatDir                                    	class admDB

; ▹ReadGVUListe                   	IstChronischKrank           	    	IstGeriatrischerPatient

; ▹class AlbisDb						class PatDBF                            	DBASEStructs                       		GetDBFData
; 	 Leistungskomplexe             	ReadPatientDBF			         		ReadDBASEIndex                     	Convert_IfapDB_Wirkstoffe

; ▹StrDiff                                	DLD                                      	FuzzyFind

; ▹class xstring

; ▹GetAlbisPath

;----------------------------------------------------------------------------------------------------------------------------------------------------------------------
;###############################################################################
;----------------------------------------------------------------------------------------------------------------------------------------------------------------------

;----------------------------------------------------------------------------------------------------------------------------------------------------------------------
; PATIENTEN DATENBANK INKL. SUCHFUNKTIONEN (Textdatei im Addendum Ordner)
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
PatDb(Pat, cmd:="")                                                       	{                 	;-- Addendum Patientendatenbank und alternatives Tagesprotokoll

	; letzte Änderung: 09.10.2021

	; Kontrolle des Objektes
		If (!IsObject(Pat) || !RegExMatch(Pat.ID, "^\d+$"))
			return

	; Nn - Nachname, Vn - Vorname, Gt - Geschlecht, Gd - Geburtsdatum, Kk - Krankenkasse
		PatID := Pat.ID
		If !oPat.Haskey(PatID) {
			oPat[PatID] := {"Nn": Pat.Nn, "Vn": Pat.Vn, "Gt": Pat.Gt, "Gd": Pat.Gd, "Kk": Pat.Kk}
			admDB.SavePatDB(oPat, Addendum.DBPath "\Patienten.json")
			If Addendum.ShowTrayTips
				TrayTip, Addendum, % "neue PatID (Zähler: " oPat.Count() ") für die Addendumdatenbank:`n(" PatID ") " Pat.Nn "," Pat.Vn, 1
		}

	; ermittelt ob diese Patientenakte heute schon aufgerufen wurden
		For Index, TProtID in TProtokoll
			 If (TProtID = PatID)
				return PatID

	; PatID wird dem Tagesprotokoll hinzugefügt und gespeichert
		TProtokoll.Push(PatID)
		IniAppend("<" PatID "(" A_Hour ":" A_Min ")>", Addendum.TPFullPath, SectionDate, compname)

return PatID
}

PatDir(PatID)                                                                 	{               	;-- erstellt den Datensicherungspfad für einen Patienten

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

class admDB                                                                 	{                 	;-- AddendumDB - Klasse für Objekt Handling

	; das globale oPat-Objekt wird benötigt (wird per )

	; lädt Daten
		ReadPatDB(PatDBPath:="") {

			If !RegExMatch(PatDBPath, "\.json$")
				PatDBPath := Addendum.DBPath "\Patienten.json"

			this.PatDBPath := PatDBPath
			If !FileExist(PatDBPath)
				PatData := ""
			else
				PatData := JSONData.Load(PatDBPath,, "UTF-8")

			; EMail Adressen hinzufügen
				tmp := ReadPatientDBF(Addendum.AlbisDBPath)
				For PatID, tmpPat in tmp
					If tmpPat.EMAIL && PatData.HasKey(PatID)
						PatData[PatID].EMAIL := tmpPat.EMAIL

		return PatData
		}

	; speichert Daten
		SavePatDB(PatData, PatDBPath:="" ) {

			If !RegExMatch(PatDBPath, "\.json$")
				PatDBPath := Addendum.DBPath "\Patienten.json"

			For PatID, Pat in PatData {
				If !Pat.HasKey("Nn") || !Pat.HasKey("Vn") || !Pat.HasKey("Gd") || !Pat.HasKey("Gt") || !Pat.HasKey("Kk") {
					throw A_ThisFunc ": Das übergebene Objekt enthält entweder keine Patientendaten oder ist defekt.`nPatID: " PatID
					return
				}
			}

			JSONData.Save(PatDBPath, PatData, true,, 1, "UTF-8")

		}

	; nur Nachname und Vorname als Array zurückgeben
		GetNamesArray() {
			names:=[]
			For PatID, Pat in oPat
				names.Push(Pat.Nn ", " Pat.Vn)
		return names
		}

	; GetExactMatches
		GetExactMatches(getString:="PatID|Gd", OnlyFirstMatch:=0, Nn:="", Vn:="", Gt:="", Gd:="", Kk:="") {

			matches	:= Object()
			getter    	:= StrSplit(getString, "|")

			For PatID, Pat in oPat {

				; ein Treffer reicht
				m1 := (StrLen(Nn) > 0)	&& RegExMatch(Pat.Nn	, "i)^" Nn) 	? 1 : 0                    	; Nachname
				m2 := (StrLen(Vn) > 0) 	&& RegExMatch(Pat.Vn	, "i)^" Vn) 	? 1 : 0                    	;
				m3 := (StrLen(Gt) > 0)	&& RegExMatch(Pat.Gt	, "i)^" Gt)  	? 1 : 0                    	;
				m4 := (StrLen(Gd) > 0)	&& RegExMatch(Pat.Gd	, "i)^" Gd) 	? 1 : 0                    	;
				m5 := (StrLen(Kk) > 0)	&& RegExMatch(Pat.Kk	, "i)^" Kk) 	? 1 : 0                    	;
				If (m1+m2+m3+m4+m5 = 0)
					continue

				matches[PatID] := Object()
				For idx, getValue in getter
					matches[PatID][getValue] := Pat[GetValue]

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

			; oPat:	muss im aufrufenden Skript global gemacht sein

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
class AlbisDB	{ 		                             	 					            		 			;-- erweitert Addendum_DBASE um zusätzliche Funktionen

		; diese Klasse benötigt Addendum_DBASE.ahk, Addendum_Datum.ahk, Addendum_Internal, class_JSON.ahk
		; PatDB wird als globales Objekt benötigt. Das Objekt muss Daten der PATIENT.dbf enthalten (verwende hierfür die Funktion ReadPatientDBF())
		; letzte Änderung: 09.12.2021

		/*  Notizen zu dBASE Feldern

				BEFAUSWA 	KUERZEL: 	wird hier nur als Zahl angegeben - möglichweise bekommt man den Klartext aus einer anderen Datenbank

				BEFDATA		TEXTDB:	welche Daten sind hier hinterlegt
									TYP:			steht wofür?
									DATA:		ist eine Zahl - hat welche Bedeutung?


		 */


	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Initialisieren
	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
		__New(DBasePath:="", debug:=0) {

				global PatDB

				this.debug 	:= debug

			; Albisdatenbankpfad oder ein alternativer Pfad für Testversuche
				this.DBPath  	:= !DBasePath ? GetAlbisPath() "\db" : DBasePath

			; Addendum Verzeichnis, Datenbankverzeichnis
				If !Addendum.Dir {
					RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)
					this.admDir := AddendumDir
				}
				else
					this.admDir := Addendum.Dir

			; Addendum Datenverzeichnis
				IniRead, admDBPath, % this.admDir "\Addendum.ini", Addendum, AddendumDBPath
				this.DBPathAddendum	:= StrReplace(admDBPath, "%AddendumDir%", this.admDir)
				If !this.CreateFilePath(this.DBPathAddendum) {
					this.ThrowError(A_ThisFunc 	": Addendum Pfad oder die Addendum.ini Datei ist nicht korrekt!`n"
											.	">>" A_ScriptDir " :: " this.DBPathAddendum "<<"
											.	"Hinweis:  Falls noch nicht gemacht. Führen Sie zunächst das AddendumStarter "
											.	"Skript aus, damit grundlegende Einstellungen, sowie Unterverzeichnisse im "
											.	"Addenumhauptverzeichnis angelegt werden. Rufen Sie danach wieder diese Funktion auf.")
				}
				this.PathDBaseData := this.DBPathAddendum "\DBASE"
				If !this.CreateFilePath(this.PathDBaseData) {
					this.ThrowError(A_ThisFunc 	": Speicherpfad für die Indexdateien konnte nicht angelegt werden`n"
											.	"[" this.PathDBaseData "]")
				}

			; weitere Objekte
				this.dbfStructs    	:= Object()
				this.EBM              	:= Object()

			; ――――――――――――――――――――――――――――――――――――――――――――――――――――――
			; EBM Regeln			        für Vorsorgeuntersuchungen oder Beratungen (aktuelle Regeln für Hausärzte von 2021)
			;
			;									zukünftige Änderungen sollten sich ohne größere Änderungen am Programmcode einpflegen lassen
			;                                 	VSFilter legt fest
			; ――――――――――――――――――――――――――――――――――――――――――――――――――――――
																			;	Vorsorgeuntersuchungen
				this.EBM.Filter        	:= {	"Vorsorgen"	: 	"rx:(01731|01732|01740M*|01745[HKM]*|01746M*|01747)"
										                                	; 	Chroniker Ziffern
								            		,	"Chroniker"	: 	"rx:(03220|03221)"
							                                 				; 	geriatrisches Basisassement
								            		,	"Geriatrisch"	:	"rx:(03360|03360)"
						                                 					; 	Quartalspauschalen
								            		,	"Pauschalen"	:	"rx:(0300[0-6])|03040)"}

				this.EBM.Vorsorgen 	:= {"EBMFilter"  	: "rx:(01731|01732|01740M*|01745[HKM]*|01746M*|01747|"
																	    	. "03220|03221|03360|03362|0300[0-6]|03040)"

												; - - - - EBM FILTERREGELN FÜR VORSORGE UNTERSUCHUNGEN ODER BERATUNGEN - - - -
												,	   "VSFilter"	: {1:{"sex":"M|W"	, "Min":"18"	, "Max":"34"  	, "repeat":"0"  	, "exam":"GVU"	}
																		,	2:{"sex":"M|W"	, "Min":"35"	, "Max":"999"	, "repeat":"36"	, "exam":"GVU"	}
																		,	3:{"sex":"M|W"	, "Min":"35"	, "Max":"999"	, "repeat":"36"	, "exam":"HKS"	}
																		,	4:{"sex":"M"   	, "Min":"55"	, "Max":"999"	, "repeat":"36"   	, "exam":"KVU"	}
																		,	5:{"sex":"M"   	, "Min":"65"	, "Max":"999"	, "repeat":"0"   	, "exam":"Aorta"	}
																		,	6:{"sex":"M|W" 	, "Min":"55"	, "Max":"999"	, "repeat":"0"		, "exam":"Colo"	}}

						                    	,		"KVU"       	: {"GO"	: "01731"          	, "Geschlecht":"M"  	, "W" : {"Alter":"55"     	, "Abstand":"36"	}}
						                    	,		"GVU"   	: {"GO"	: "01732"          	, "Geschlecht":"M|W"	, "W" : {"Alter":"35"      	, "Abstand":"36"	}
																																				, "E" 	: {"Alter":"18-34"	, "Abstand":"0"	}}
						                    	,		"Colo"    	: {"GO": "01740"            	, "Geschlecht":"M|W"	, "W" : {"Alter":"55"     	, "Abstand":"0"	}}
						                    	,		"HKS"    	: {"GO": "(01745|01746)"	, "Geschlecht":"M|W"	, "W" : {"Alter":"35"     	, "Abstand":"36"	}}
						                    	,		"Aorta"    	: {"GO": "01747"            	, "Geschlecht":"M"  	, "W" : {"Alter":"65"     	, "Abstand":"0"	}}
						                    	,		"Chr1"    	: {"GO": "03320"            	, "Geschlecht":"M|W"	, "W" : {"Alter":"0"        	, "Abstand":"3"	}}
						                    	,		"Chr2"    	: {"GO": "03321"            	, "Geschlecht":"M|W"	, "W" : {"Alter":"0"       	, "Abstand":"3"	}}
						                    	,	  	"#01731"	: "KVU"
						                    	,	  	"#01732"	: "GVU"
						                    	, 	 	"#01740"	: "Colo"
						                    	, 	 	"#01745"	: "HKS"
						                    	, 	 	"#01746"	: "HKS"
						                    	, 	 	"#01747"	: "Aorta"}

				this.EBM.Chroniker 	:= {	"CHR1"       	: {"GO": "03320", "Geschlecht":"M|W", "B" : {"Alter":"0", "Abstand":"3", "ICDListe":""}}
						                    	,	 	"CHR2"    	: {"GO": "03321", "Geschlecht":"M|W", "B" : {"Alter":"0", "Abstand":"3", "GO":"CHR1", "ICDListe":""}}
												,		"#03320"	: "CHR1"
												,		"#03321"	: "CHR2"}

				this.EBM.Geriatrisch	:= {"EBMFilter"  	: "rx:(0336[02])"}
				this.EBM.Pauschalen	:= "rx:(0300[0-6]|03040)"

			; zusätzliche Daten/Informationen
				this.VERSART            	:= ["Mitglied", "2", "Angehöriger", "4", "Rentner"]
				this.ScheinN            	:= {"Abrechnung":0, "1":1 , "Überweisung":2, "3":3, "Notfall":4, "PRIVAT":5}
				this.ScheinR             	:= {"#0":"Abrechnung", "#1":"1", "#2":"Überweisung", "#3":"3", "#4":"Notfall", "#5":"PRIVAT"}
				this.DateLabels      	:= "(STAT_VON|STAT_BIS|DATUM|EINLESTAG|AUS2DAT|GUELTIG|GUELTVON|"
							        			.	 "GUELTVON2|GUELTBIS|GUELTBIS2|GEBURT|VERSBEG)"

		}


	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Basisfunktionen
	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		DBASEStructs(DBASEName:="", DBStructsPath:="", debug=false)	                    	{ 	;-- analysiert alle oder eine DBase Datei/en

			; DBASEName 		wird hier ein Name übergeben, werden nur die Strukturdaten dieser Datei hinzugefügt/geändert
			; DBStructsPath: schreibt die ausgelesenen Strukturen als JSON formatierte Datei

				this.dbfiles         	:= Array()
				this.dbfSizeMax 	:= this.dbfSizeMax := this.mdxSizeMax := 0

			; zählt für die Dateianzeige zunächst einmal die vorhandenen Dateien im Verzeichnis
				Loop, Files, % this.DBpath "\*.dbf"
					this.dbfiles.Push(StrReplace(A_LoopFileName, ".dbf"))

			; für formatierte Ausgabe
				this.dbFilesMax := this.dbfiles.Count()
				dbIL := StrLen(this.dbFilesMax) - 1

			; öffnet jede DBASE Datei und liest den Header der Datei aus
				dbNr := 0
				getStructAll := StrLen(DBASEName) > 0 ? false : true
				For dbfNr, dbname in this.dbfiles {

						If (!getStructAll && DBASEName = dbName) || (getStructAll = true) {

							; Debug Ausgabe
								If debug && (Mod(dbfNr, 20) = 0)
									ToolTip, % "lese: " SubStr("000" (dbNr ++), -1 * dbIL) "/" this.dbFilesMax ": " dbname

							; Zeitstempel
								FileGetTime, accessed,  % this.DBPath "\" dbname ".dbf"	, A
								FileGetTime, modified,  % this.DBPath "\" dbname ".dbf"	, M

							; Datenbankgrößen
								FileGetSize, SizeDBF	, % this.DBPath "\" dbname ".dbf"		, K
								FileGetSize, sizeDBT	, % this.DBPath "\" dbname ".dbt" 	, K
								FileGetSize, sizeMDX	, % this.DBPath "\" dbname ".mdx"	, K
								this.dbfSizeMax 	+= sizeDBF
								this.dbfSizeMax 	+= sizeDBT
								this.mdxSizeMax	+= sizeMDX

							; Datenbankstruktur einlesen
								dbf  	:= new DBASE(this.DBPath "\" dbname ".dbf", false)
								If (dbf.Version = 0x8b || dbf.Version = 0x3) {

									If !this.dbfStructsPath(dbName)  ; erstellt ein sub Objekt mit dem Namen der Datenbank
										this.ThrowError(	"Ein weiteres Objekt für dbfStructs ["  dbName "] kann nicht angelegt werden!")

									this.dbfStructs[dbname].Nr              	:= dbfNr
									this.dbfStructs[dbname].dbfields      	:= isObject(dbf.dbfields) 	? dbf.dbfields 	: "error reading dbase file"
									this.dbfStructs[dbname].fields         	:= isObject(dbf.fields)    	? dbf.fields    	: "error reading dbase file"
									this.dbfStructs[dbname].header      	:= isObject(dbf.dbstruct)	? dbf.dbstruct 	: "error reading dbase file"
									this.dbfStructs[dbname].headerLen  	:= dbf.headerLen
									this.dbfStructs[dbname].lendataset  	:= dbf.lendataset
									this.dbfStructs[dbname].records      	:= dbf.records
									this.dbfStructs[dbname].lastupdate  	:= dbf.lastupdateDate
									this.dbfStructs[dbname].lastupdateE	:= dbf.lastupdateEng
									this.dbfStructs[dbname].sizeDBF     	:= sizeDBF
									this.dbfStructs[dbname].sizeDBT     	:= sizeDBT
									this.dbfStructs[dbname].sizeMDX     	:= sizeMDX
									this.dbfStructs[dbname].accessed     	:= accessed
									this.dbfStructs[dbname].modified     	:= modified
									this.dbfStructs[dbname].Version      	:= dbf.Version

								}
								dbf.Close()
								dbf := ""

						}

						If (!getStructAll && DBASEName = dbName)
							break

				}
				ToolTip

			; speichern der Strukturdaten
				If RegExMatch(DBStructsPath, "[A-Z]\:\\") && this.CreateFilePath(DBStructsPath) {

					If (!getStructAll)
						saveName := DBASEName
					else
						saveName := "AlbisDBASE"

					file := FileOpen(DBStructsPath "\" savename ".json", "w", "UTF-8")
					file.Write("; -------------------------------------------------------------------------------------------------"                                      	"`n")
					file.Write("; ***** Informationen DBASE Dateien *****"														                                            	"`n")
					file.Write("; Verzeichnis  `t: " 	AlbisPath "\db"		                                                                                                        	"`n")
					file.Write("; Erstellt am `t: " 	A_DD "." A_MM "." A_YYYY " um " A_Hour ":" A_Min "Uhr	"                                               	"`n")
					file.Write("; Dateien      `t: " 	dbFilesMax 																		                                            	"`n")
					file.Write("; Dateigrößen`t" 	 																						                                            	"`n")
					file.Write(";   .dbf       `t: " 		Round(this.dbfSizeMax/1024/1024	, 2)	" GB" 	            					                        	"`n")
					file.Write(";   .dbt       `t: " 		Round(this.dbfSizeMax/1024/1024	, 2)	" GB" 						                                    	"`n")
					file.Write(";   .mdx       `t: " 	Round(this.mdxSizeMax/1024/1024, 2)	" GB" 		    		                	                        	"`n")
					file.Write(";   Summe     `t: " 	Round((this.dbfSizeMax+this.dbfSizeMax+this.mdxSizeMax)/1024/1024, 2) " MB"    			"`n")
					file.Write("; -------------------------------------------------------------------------------------------------"                                      	"`n")
					file.Write(JSON.Dump(this.dbfStructs,,2))
					file.Close()

					this.DBStructsPath := DBStructsPath
			}

		return this.dbfstructs
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		GetDBFData(dbName, P="", O="", S=0, dg=0, Opt="")                                	{ 	;-- holt Daten aus einer beliebigen Datenbank

			/*	GetDBFData()

					benötigt Addendum_DBASE.ahk
					dbName	-	Name der DBASE Datenbank ohne Dateiendung
					P             	- 	pattern as object ({"Quartal":"119", "KUERZEL":"lk\w"})
					O           	- 	Array mit Feldnamen die zurück gegeben werden sollen
					S             	- 	String mit 2 Optionen:
							    			1. seekpos	= Start an dieser Dateiposition oder
							    			2. recordnr	= die Nummer des Datensatzes mit welchem der Lesevorgang startet
					dg	    	-	Debuglevel (0 = keine Ausgabe, 1-4 sind unterschiedliche Ausgabemodi)
					Opt	    	-	weitere Debug Optionen können hier übergeben werden:
										callback="Name der Funktion"    	- z.B. callback="Fortschrittsanzeige"
										gui="GuiName"                            - Ausgabe von Informationen direkt in eine Gui
									-	verbundene Daten aus der BEFTEXT.dbf automatisch anfügen:
										linkedData=true

			 */

			; entfernt angehängte Dateierweiterungen u.a.
				If !(dbName := this.DBaseFileExist(dbName))
					return 0

			; eine Callback Funktion für die Anzeige des Fortschritts kann eingerichtet werden
				RegExMatch(Opt " ", "callback\s*\=\s*(?<Func>.*?)([^\w]|\s|$)"	, callback)
				RegExMatch(Opt " ", "gui\s*\=\s*(?<Gui>.*?)([^\w]|\s|$)"        	, debug)

			; filepointer berechnen
				startrec	:=	RegExMatch(S, "seekpos\s*\=\s*(?<pos>\d+)", seek) 	? Floor(seekpos/dbf.lendataset)-1
								: 	RegExMatch(S, "recordnr\s*\=\s*(?<nr>\d+)", record)	? recordnr		:	0

			; P - enthaltene RegEx-Strings säubern
				For fLabel, value in P
					If RegExMatch(P[fLabel], "^\s*rx\:") {
						useRegExSearch := true
						break
					}
				;~ P[fLabel] := RegExReplace(P[fLabel], "^\s*rx\:")

			; Informationen der Datenbank bereitstellen
				dbf	:= new DBASE(this.DBpath "\" dbName ".dbf", dg)
				res 	:= dbf.OpenDBF()

			; Search [RegEx - Datenbanksuche durchführen] oder SearchFast [Instr - Datenbanksuche durchführen]
				matches := useRegExSearch ? dbf.Search(P, O, startrec,, callbackFunc) : dbf.SearchFast(P, O, startrec,, callbackFunc)
				;SciTEOutput(A_ThisFunc ": dbf.connected - " dbf.connected)

			; ein paar wichtige Filepositionen sichern
				this.DBEndFilePos := dbf.filepos
				this.DBEndRecord := dbf.breakrec
				res := dbf.CloseDBF()
				dbf := ""

				If !IsObject(O)
					return matches

			; Datenmenge schrumpfen
				;~ data := Array()
				;~ For midx, m in matches {
					;~ strObj := Object()
					;~ For cidx, k in O
						;~ strObj[k] := m[k]
					;~ data.Push(strObj)
				;~ }


			; specialist enhancement for the use with Albis databases, some values are not stored in full length in main database
				;~ If IsObject(dbf.connected) {
					;~ data := this.AppendMatches(data, dbf.connected)
				;~ }


		return matches
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		GetBEFTEXTData(TextDB, ReIndex:=false)                                                          	{	;-- verlinkte Daten aus BEFTEXT.dbf lesen

		  ; TEXTDB 	-	ein Object mit den laufenden Nummern zu verknüpften Text(Daten)inhalten (TEXTDB in BEFUND.dbf => LFDNR in BEFTEXTE.dbf)
		  ; Reindex 	-	um die Daten schneller auslesen zu können, wird die Datenbank zunächst indiziert. Sollte dieser Index
		  ;				    	einmal fehlerhaft sein, kann er über ein 'ReIndex=true' komplett neu erstellt werden. Ansonsten werden nur neue Sprungadressen hinzugefügt
		  ;

		  ; Datensammler anlegen
			xData := Object()

		  ; BEFTEXTE.dbf indizieren oder nur den Index laden ;{
			If (!IsObject(this.BEFIndex) || ReIndex)
				this.BEFIndex := this.TextDBIndex(ReIndex)
			;}

		  ; TEXTDB Array aufsteigend sortieren ;{
			;~ t := ""
			;~ For k,v in TextDB
				;~ t .= (k > 1 ? "`n" : "") v
			;~ Sort, t, N U
			;~ TextDB := Array()
			;~ For i, v in StrSplit(t, "`n")
				;~ TextDB.Push(v)
		  ;}

		  ; BEFTEXTE: Auslesen zusätzlicher Daten
			dbf        	:= new DBASE(Addendum.AlbisDBPath "\BEFTEXTE.dbf", 2)
			steps      	:= dbf.records/100
			filepos   	:= dbf.OpenDBF()

		; LFDNR finden und Datensätze zu einem String zusammenfügen
			MaxLFDNR := MaxLines := LFDNR_next := LFDNR_last := 0
			found := false
			For LFDNR, PatID in TextDB {

		      ; den nächsten indizierten Einsprung finden und den filepointer an diese Stelle versetzen
				For LFDNR_Indexed, seek in this.BEFIndex.LFDNR {

				  ; seekpos gefunden, Suche nach dem seekpos beenden
					If (LFDNR <= LFDNR_Indexed) {
						found := true
						startrecord := LFDNR < LFDNR_Indexed ? lastseek.1 : seek.1
						break
					}
					LFDNR_last 	:= LFDNR_Indexed
					lastseek      	:= seek

				}

			  ; LFDNR wurde nicht gefunden, dann wird der letzte indizierte Datensatz verwendet
				If !found
					startrecord := seek.1

			  ; FilePointer zur gefundenen seekpos verschieben
				dbf.__SeekToRecord(startrecord)

			  ; LFDNR in der BEFTEXT.dbf suchen
				found := false, xtra := ""
				Loop % (dbf.records-startrecord) {

				  ; einzelne Datensätze zu einem String zusammensetzen, wenn die Nummern übereinstimmen
					set := dbf.ReadRecord(["PATNR", "DATUM", "LFDNR", "POS", "TEXT"])
					If (LFDNR = set.LFDNR) {
						found	:=	true
						xtra   	.=	set.TEXT
						xlines	:=	set.POS + 1
					}
				 ; alle Datensätzen wurden erfasst, jetzt einfach
					else If (found && LFDNR <> set.LFDNR) {
						break
					}

				}

			  ; zusätzliche Daten im Sammler ablegen
				If found {
					If !IsObject(xData[PatID])
						xData[PatID] := Object()
					xData[PatID][LFDNR] := {"xtra" : xtra, "xlines" : xlines}
					MaxLines += xlines
					MaxLFDNR ++
					ToolTip, % "Extradata added:`n`n" MaxLines " lines for`n" xData.Count() " counts of patients with `n" MaxLFDNR " collected LFDNR's"
				}

			}

			filepos	:= dbf.CloseDBF()
			dbf    	:= ""

		return xData
		}


	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Funktionen für eine quartalsweise Auswertung
	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
		VorsorgeDaten(Abrechnungsquartal, ReIndex:=false, save:=true, ini:=false)     	{ 	;-- Datumserfassung abgerechneter Vorsorgekomplexe

			; sammelt Abrechnungsdaten zu den EBM Gebühren der Vorsorgen, Chroniker- und Geriatrische Pauschalen

			; wird benötigt um innerhalb dieses Zeitraumes das letzte Vorsorgedatum zu suchen (einmalige GVU)
				global PatDB
				static PatExtraFilePath := Addendum.DBPath "\PatData\PatExtra.ini"

			; gespeicherte Daten laden oder zunächst erstellen
				If !ReIndex {

				; VSData wurde schon erstellt, dann gleich hier zurück geben
					If IsObject(this.VSData)              ; ## Test
						return this.VSData

					;~ BaseQ := GetQuartalEx(Abrechnungsquartal, "QQYY")
					IniRead, lastUpdate    	, % this.PathDBaseData "\class_AlbisDB.ini", % A_ThisFunc, % "lastUpdate"
					IniRead, lastDBPos     	, % this.PathDBaseData "\class_AlbisDB.ini", % A_ThisFunc, % "lastFilePos"
					IniRead, lastDBRecord	, % this.PathDBaseData "\class_AlbisDB.ini", % A_ThisFunc, % "lastRecord"
					ReIndex := InStr(lastUpdate, "ERROR") ? true : ReIndex
					;~ If (!ReIndex) && (lastUpdate = BaseQ)
						;~ return JSONData.Load(this.PathDBaseData "\Vorsorgen.json", "", "UTF-8")
				}

			; abgerechnete Vorsorgen/Beratungen ohne Filterung zusammenstellen
				VSData	:= Object()
				For idx, m in this.GetDBFData("BEFGONR", {"GNR":this.EBM.Vorsorgen.EBMfilter}, ["PATNR","QUARTAL", "GNR", "DATUM", "ID"], 0) {

					PatID 	:= m.PATNR
					Patient	:= PatDB[PatID]
					YYQ 	:= this.SwapQuarterS(m.QUARTAL)
					RegExMatch(m.GNR, "\d+", GNR)

					If !IsObject(VSData[PatID])
						VSData[PatID] := {"GESCHL":this.Geschlecht(Patient.GESCHL), "GEBURT":Patient.GEBURT, "VS":{}}

					If RegExMatch(GNR, "0322(?<N>[01])", c)                	{   	; Chroniker Pauschalen

						If ini {
							IniWrite, % m.QUARTAL	, % PatExtraFilePath, % PatID , % "Chroniker" cN "_letztes_Quartal"
							IniWrite, % m.Datum 	, % PatExtraFilePath, % PatID , % "Chroniker" cN "_letztes_Datum"
							IniWrite, % c              	, % PatExtraFilePath, % PatID , % "Chroniker" cN "_letzte_Ziffer"
						}

						If !IsObject(VSData[PatID].chronisch)
							VSData[PatID].chronisch := Object()

						cN += 1    ; RegExMatch Ergebnis ist entweder 1 oder 2
						If !VSData[PatID].chronisch.HasKey(YYQ)
							VSData[PatID].chronisch[YYQ] := cN
						else
							VSData[PatID].chronisch[YYQ] += cN

					}
					else If RegExMatch(GNR, "0336(?<N>[02])", c)       	{		; Geriatsche Pauschalen

						If ini {
							IniWrite, % m.QUARTAL	, % PatExtraFilePath, % PatID , % "Geriatrisch" cN "_letztes_Quartal"
							IniWrite, % m.Datum  	, % PatExtraFilePath, % PatID , % "Geriatrisch" cN "_letztes_Datum"
							IniWrite, % c              	, % PatExtraFilePath, % PatID , % "Geriatrisch" cN "_letzte_Ziffer"
						}

						If !IsObject(VSData[PatID].Geriatrisch)
							VSData[PatID].Geriatrisch := Object()

						cN += 1
						If !VSData[PatID].Geriatrisch.HasKey(YYQ)
							VSData[PatID].Geriatrisch[YYQ] := cN
						else
							VSData[PatID].Geriatrisch[YYQ] += cN

					}
					else if RegExMatch(GNR, "(0300[0-6]|03040)", c)    	{		; Grundpauschalen, Ordinationsgebühr

						If !IsObject(VSData[PatID].Pauschale)
							VSData[PatID].Pauschale := Object()
						If !IsObject(VSData[PatID].Pauschale[YYQ])
							VSData[PatID].Pauschale[YYQ] := Object()
						VSData[PatID].Pauschale[YYQ].Push(c)

					}
					else {

						If ini {
							exam := this.EBM.Vorsorgen["#" GNR]
							IniWrite, % m.QUARTAL	, % PatExtraFilePath, % PatID , % exam "_letztes_Quartal"
							IniWrite, % m.Datum   	, % PatExtraFilePath, % PatID , % exam "_letztes_Datum"
							IniWrite, % GNR        	, % PatExtraFilePath, % PatID , % exam "_Ziffer"
						}

						VSData[PatID].VS[(this.EBM.Vorsorgen["#" GNR])] := {"Q":m.QUARTAL, "D":m.Datum}

					}

				}

			  ; Speichern der Daten um diese später laden zu können
				If save || (lastUpdate <> BaseQ) {
					JSONData.Save(this.PathDBaseData "\Vorsorgen.json", VSData, true,, 1, "UTF-8")
					IniWrite, % BaseQ               	, % this.PathDBaseData "\class_AlbisDB.ini", % A_ThisFunc, % "lastUpdate"
					IniWrite, % this.DBEndRecord	, % this.PathDBaseData "\class_AlbisDB.ini", % A_ThisFunc, % "lastRecord"
					IniWrite, % this.DBEndFilePos	, % this.PathDBaseData "\class_AlbisDB.ini", % A_ThisFunc, % "lastFilePos"
				}

				this.VSData := VSData

		return VSData
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		Chroniker(Abrechnungsquartal, ReIndex:=false, save:=true)                           	{	;-- erfasst alle Patienten mit Chronikerpauschalen

				global PatDB

			; VSData laden oder neu erstellen
				If !ReIndex && !IsObject(VSData) && FileExist(this.PathDBaseData "\Vorsorgen.json") {
					VSData := JSONData.Load(this.PathDBaseData "\Vorsorgen.json", "", "UTF-8")
						If !IsObject(VSData)
							this.ThrowError("VSData konnte nicht geladen werden!", A_LineNumber-2)
				} else
					VSData := this.VorsorgeDaten(Abrechnungsquartal, true, true)

			; Chronikerpauschalen sammeln
				For idx, m in this.GetDBFData("BEFGONR", {"GNR":"0322[01]"}, ["PATNR","QUARTAL", "GNR", "DATUM", "ID"], 0) {

					PatID := m.PATNR
					If !IsObject(VSData[PatID])
						VSData[PatID] := Object()
					If !IsObject(this.VSData[PatID].chronisch)
						VSData[PatID].chronisch := Array()

					quarter := this.SwapQuarter(m.QUARTAL)
					GNR := SubStr(m.GNR, 5, 1) + 1
					VSData[PatID].chronisch.Push(m.QUARTAL)

				}

				If save
					JSONData.Save(this.PathDBaseData "\Vorsorgen+Chroniker.json", VSData, true,, 1, "UTF-8")

				this.VSData := VSData

		return VSData
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		GBAsessement(ReIndex:=false, save:=true)                                                      	{	;-- erfasst alle eingetragenen geriatr. Basiskompl.

		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		Krankenscheine(Abrechnungsquartal, ksart="alle", pv=true, save=false)          	{	;-- Patienten mit angelegten Kassenscheinen

			; Abrechnungsquartal im Format QQYY
			; ksart  : 	Scheine als String [Komma getrennt] zb. Abrechnung, Notfallfall

				global PatDB

			; Klarnamen der Abrechnungsscheine werden für den Vergleich in das von der Datebank genutzte Format umgewandelt
				ksartNr	:= this.KrankenscheinArt(ksart)
				quartal := LTrim(Abrechnungsquartal, "0")

			; Krankenscheine zusammenstellen
				KSMatches   	:= this.GetDBFData("KSCHEIN", {"QUARTAL":quartal, "TYP":ksartNr}
																, ["PATNR", "QUARTAL", "TYP", "VERSART", "EINLESTAG"], 0)

				this.KScheine	:= Object()
				For idx, data in KSMatches {

					; wenn pv = false (Privatversicherung), werden Kassenscheine nicht erfasst
						PatID := data.PATNR
						If (!pv && (PatDB[PatID].PRIVAT <> "f"))   ; !RegExMatch(data.TYP, ksartNr) ||
							continue

						If !IsObject(this.KScheine[PatID]) {
							this.KScheine[PatID] := {"VERSART"    	: this.VERSART[data.VERSART]
															, 	 "EINLESTAG"	: data.EINLESTAG
															, 	 "KSCHEIN"    	: this.KrankenscheinArt(data.TYP)
															, 	 "KSCHEINTYP" 	: data.TYP
															, 	 "PRIVAT"       	: (PatDB[PatID].PRIVAT <> "f" ? 1 : 0)
															, 	 "GESCHL"       	: this.Geschlecht(PatDB[PatID].GESCHL)}
						}

				}

				If Save
					JSONData.Save(this.PathDBaseData "\KScheine_" quartal ".json", this.KScheine, true,, 1, "UTF-8")

		return this.KScheine
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		Abrechnungsscheine(Abrechnungsquartal, ksart="alle", pv:=false, save:=false)	{	;-- Krankenscheine und alle Abrechnungsziffern

			; Abrechnungsquartal im Format 0121 oder 121
			; pv = privatversichert

				global PatDB

			; alle angelegten Krankenscheine auslesen
				KScheine 	:= this.Krankenscheine(Abrechnungsquartal, ksart, pv, save)
				quartal  	:= LTrim(Abrechnungsquartal, "0")

			; alle abgerechneten Ziffern zusammentragen
				this.AbrScheine 	:= Object()
				matches            	:= this.GetDBFData("BEFGONR"
																		, {"QUARTAL": LTrim(Abrechnungsquartal, "0")}
																		, ["PATNR", "ARZTID", "QUARTAL", "GNR", "DATUM", "ID", "removed"], 0)

			; KScheine und Ziffern zusammenführen
				For idx, m in matches {

						PatID := m.PATNR
					; wenn pv = false (Privatversicherung), werden diese Kassenscheine nicht erfasst
						If (!pv && PatDB[PatID].PRIVAT <> "f")
							continue
					; pv = 2 dann werden keine Kassenscheine erfasst
						else if (pv = 2 && PatDB[PatID].PRIVAT = "f")
							continue

					; neue PatID - Object erweitern
						If !IsObject(this.AbrScheine[PatID]) {
							this.AbrScheine[PatID] := { "VERSART"     	: KScheine[PatID].VERSART
																, 	 "EINLESTAG"	: KScheine[PatID].EINLESTAG
																, 	 "KSCHEINTYP" 	: KScheine[PatID].TYP
																, 	 "PRIVAT"       	: KScheine[PatID].PRIVAT
																, 	 "GESCHL"       	: KScheine[PatID].Geschlecht
																,	 "GEBUEHR"    	: {}}  ; "ARZTID":0, "DATUM":[]
						}

					; Ziffer und Abrechnungsdatum hinzufügen
						If !this.AbrScheine[PatID].GEBUEHR.HasKey("#" m.GNR)
							this.AbrScheine[PatID].GEBUEHR["#" m.GNR] := Array()

						this.AbrScheine[PatID].GEBUEHR["#" m.GNR].Push({"DATUM":m.DATUM, "ARZTID":m.ARZTID, "removed":m.removed})

				}

			 ; Speichern
				If Save
					JSONData.Save(this.PathDBaseData "\AbrScheine_" quartal ".json", this.AbrScheine, true,, 1, "UTF-8")

		return this.AbrScheine
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		Behandlungstage(Abrechnungsquartal, ksart="alle", pv:=false, save:=false)     	{	;-- ermittelt alle Tage mit Eintragungen (Behandlungstage)

			; Abrechnungsquartal Format : QQYY z.B. 0121

				global PatDB

				quartal	:= LTrim(Abrechnungsquartal, "0")
				QZ     	:= SubStr(quartal, 1, 1)                   	; Quartalzahl
				YY     	:= SubStr(quartal, 2, 2)                 	; Jahr
				QM1 	:= SubStr("0" (((QZ-1)*3)+1), -1)    	; Monatszahl des 1.Monat im Quartal
				QM2 	:= SubStr("0" (QM1+1), -1)              	; Monatszahl des 2.Monat im Quartal
				QM3 	:= SubStr("0" (QM2+1), -1)              	; Monatszahl des 3.Monat im Quartal
				rxQ   	:= (YY > SubStr(A_YYYY, 3, 2) ? "19":"20") YY "(" QM1 "|" QM2 "|" QM3 ")\d\d"   ; 2021(01|02|03)\d\d


				KScheine := !IsObject(this.KScheine) ? this.Krankenscheine(Abrechnungsquartal, ksart, pv, false) : this.KScheine

			; alle Tage mit Kuerzeleintragungen finden (QUARTAL ist nicht verwendbar da manchmal nur eine 0 vergeben ist)
				matches   	:= this.GetDBFData("BEFUND", {"DATUM":"rx:" rxQ, "KUERZEL":"\w"}, ["PATNR", "DATUM", "KUERZEL", "INHALT", "removed"], 0)
				this.Behandlungstage := Object()
				For idx, m in matches {

						PatID := m.PATNR
					; wenn pv = false (Privatversicherung), werden diese Kassenscheine nicht erfasst
						If (!pv && PatDB[PatID].PRIVAT <> "f" || m.removed)
							continue

					; Objektstruktur anlegen
						If !IsObject(this.Behandlungstage[PATID])
							this.Behandlungstage[PATID] := Object()
						If !IsObject(this.Behandlungstage[PATID][m.Datum])
							this.Behandlungstage[PATID][m.Datum] := Array()

						this.Behandlungstage[PATID][m.Datum].Push({"KZL":m.KUERZEL, "INH":m.INHALT})

				}

			; Speichern
				If Save
					JSONData.Save(this.PathDBaseData "\BehTage_" quartal ".json", this.Behandlungstage, true,, 1, "UTF-8")

		return this.Behandlungstage
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		VorsorgeFilter(VSData, Abrechnungsquartal, ksart="alle", pv:=true, save:=true)	{	;-- welche Vorsorgen aktuell liquidiert werden können

		  ; benötigt PatDB als globales Objekt!
		  ; pv = privatversichert
			global PatDB

		  ; Fehlerhandling                                                                                                    	;{
			If !IsObject(VSData)
				Throw A_ThisFunc ": Parameter VSData ist kein Objekt!"
			If !RegExMatch(Abrechnungsquartal, "0*[1-4][0-9]{2}")
				Throw A_ThisFunc ": Parameter Abrechnungsquartal " Abrechnungsquartal " hat ein falsches Format oder ist leer!"
		  ;}

		  ; Patienten mit angelegten Abrechnungsscheinen im Abrechnungsquartal ermitteln
			KScheine := this.Krankenscheine(Abrechnungsquartal, ksart, pv, save)

		  ; Variablen                                                                                                            	;{
			VSCan          	:= Object()
			Quartal         	:= QuartalTage({"aktuell":SubStr("0" Abrechnungsquartal, -3)})
			lastQDay      	:= Quartal.DBaseEnd
			lastCmpQDay	:= lastQDay "000000"
			VSFilter         	:= this.EBM.Vorsorgen.VSFilter
		  ;}

		  ; die letzten 3 Quartale darstellen                                                                            	;{
			YYQ     	:= this.SwapQuarterS(LTrim(Abrechnungsquartal, "0"))              	; Jahr dann Quartal
			nYY      	:= SubStr(YYQ, 1, 2)
			nQ        	:= SubStr(YYQ, 3, 1)
			lQuarters	:= "("
			Loop 3 {
				nQ 	:= nQ-1=0 ? 4 : nQ-1
				nYY 	:= nQ=4 ? (nYY-1<0 ? 99 : nYY-1) : nYY
				lQuarters .= SubStr("0" nYY, -1) . nQ . "|"
			}
			lQuarters := RTrim(lQuarters, "|") . ")"
		  ;}

		  ; abrechenbare Untersuchungen mit Hilfe der VSFilter (Vorsorgefilter) finden
			For PatID, q in KScheine {
				If (!pv && !q.PRIVAT) && !RegExMatch(q.KSCHEINTYP, "[24]") {

				  ; --------------------------------------------------------------------------------------------
				  ; Alter und Geschlecht berechnen
				  ; --------------------------------------------------------------------------------------------
					age	:= this.ageYears(PatDB[PatID].GEBURT, lastQDay)
					sex	:= this.Geschlecht(PatDB[PatID].GESCHL)

				  ; --------------------------------------------------------------------------------------------
				  ; EBM Regeln für Vorsorgeuntersuchungen, Präventionsziffern anwenden
				  ; --------------------------------------------------------------------------------------------
					pending := Array()
					For ruleNR, rule in VSFilter {

					  ; das Alter des Patienten und das Geschlecht müssen passen
						If (age >= rule.Min && age <= rule.Max) && RegExMatch(sex, "i)(" rule.sex ")"){

							; Leistung (exam) ist nur einmal (!repeat) im Arztfall abrechenbar und wurde bisher nicht angesetzt
								If (!rule.repeat && !VSData[PatID].VS.HasKey(rule.exam)) {                                     			; z.B. rule.exam = Colo
									pending.Push(rule.exam)
								}

							; Leistung ist alle x-Monate abrechenbar
								else if rule.repeat {                                                                                                      	; wurde aber noch nie abgerechnet
									If !VSData[PatID].VS.HasKey(rule.exam) {
										pending.Push(rule.exam)
									}
									else {                                                                                                                    	; wurde abgerechnet - Abstand zwischen einer möglichen erneuten Abrechnung überprüfen
										NextPossDate := DateAddEx(VSData[PatID].VS[rule.exam].D, Floor(rule.repeat/12) "y")
										If (lastCmpQDay >= NextPossDate)
											pending.Push(rule.exam "-" NextPossDate)
									}
								}
							}
						}

				  ; --------------------------------------------------------------------------------------------
				  ; EBM Regeln für Chronikerpauschalen
				  ; --------------------------------------------------------------------------------------------
					If !VSData[PatID].chronisch.Count()                               	; es wurde noch nie eine Chronikerpauschale berechnet
						iscronicallysick := 0 , chronic := chronicQuarterMissed := "---"
					else {
					  ; chronic - 0 beide Ziffern angesetzt, 2 - es fehlt 03221, 1 - es fehlt 03220
					  ;				3 beide Ziffern fehlen
					  ; von den letzten 3 Quartalen muss der Patient in 2 Quartalen behandelt worden sein
						iscronicallysick := true
						crn        	:= VSData[PatID].chronisch[YYQ]
						chronic 	:= 3 - crn
						notmissed := 0
						For cQuarter, cCount in VSData[PatID].chronisch
							If RegExMatch(cQuarter, lQuarters)                   ; RegExString der Form 2021(01)|(02|(03)\d\d
								notmissed ++
						chronicQuarterMissed := 2 - (notmissed > 2 ? 2 : notmissed) ; 0 wenn beide Ziffern abgerechnet wurden
					}

				  ; --------------------------------------------------------------------------------------------
				  ; EBM Regel für geriatrische Basiskomplexe
				  ; --------------------------------------------------------------------------------------------
					geriatric := 0, isgeriatric := false
					If (VSData[PatID].Geriatrisch.Count()>0) {
						isgeriatric := true
						If !VSData[PatID].Geriatrisch.HasKey(YYQ)
							geriatric := 1
					}

				  ; --------------------------------------------------------------------------------------------
				  ; Grundpauschalen nicht eingetragen oder mehrfach
				  ; --------------------------------------------------------------------------------------------;{
					nobasefee := false
					If !IsObject(VSData[PatID].Pauschale[YYQ])
						nobasefee := 2
					else {
						feePointer:=0
						For feeIdx, fee in VSData[PatID].Pauschale[YYQ]
							If RegExMatch(fee, this.EBM.Pauschalen, c)
								feePointer ++
						nobasefee := feepointer = 2 ? 0 : feePointer > 2 ? feePointer-2 : feepointer
					}
				;}

				  ; --------------------------------------------------------------------------------------------
				  ; gesammelte Daten dem Objekt hinzufügen
				  ; --------------------------------------------------------------------------------------------
					If (pending.MaxIndex() > 0 || cronic || cronicQuartalMissed || geriatric || nobasefee) {

						VSCan[PatID] := {	"ALTER"                     	: age
												, 	"GESCHL"                 	: sex
												,	"NAME"                     	: PatDB[PatID].Name ", " PatDB[PatID].VORNAME
												,	"KSCHEINART"           	: this.KrankenscheinArt(q.KSCHEINTYP)
												,	"KSCHEINTYP"           	: q.KSCHEINTYP
												,	"VERSART"                  	: q.VERSART
												,	"PRIVAT"                     	: q.PRIVAT
												,	"VORSCHLAG"          	: (pending.MaxIndex() = 0 ? "" : pending)
												,	"CHRONISCHGRUPPE" 	: iscronicallysick
												,	"GERIATRISCHGRUPPE"	: isgeriatric
												,	"PAUSCHALE"       	     	: nobasefee}

						If VSCan[PatID].CHRONISCHGRUPPE {
							VSCan[PatID].CHRONISCH 	:= chronic
							VSCan[PatID].CHRONISCHQ	:= chronicQuarterMissed
						}
						If VSCan[PatID].GERIATRISCHGRUPPE
							VSCan[PatID].GERIATRISCH 	:= geriatric

					}

				}
			}

		; Statistik der KScheine in VorsorgeKandidaten hinterlegen (VSCandidates)
			VSCan["_Statistik"] := {"KScheine": KScheine.Count()}

		; Liste speichern
			If save
				JSONData.Save(this.PathDBaseData "\VorsorgeKandidaten-" LTrim(Abrechnungsquartal, "0") ".json", VSCan, true,, 1, "UTF-8")

		return VSCan
		}


	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Funktionen für andere Auswertungen
	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		Impfstatistik(Datum, ARZTID:="", Opt:="")                                                      	{	;-- Impfststatistik zu den COVID19-Schutzimpfungen

			; Impstatistik: im Moment nur Zählung eines Tages
			; B = Biontech, A - Astra ...
			; AG = Altersgruppe + O  ,     M  	,    J
			;								>59 , 18-59 , < 18

			global PatDB

		  ; Variablen
					vacfees := { "Z" 	:{"A|G|V":1, "B|H|W":2, "K|R|X":3}
									,  "#88331":"B", "#88332":"M", "#88333":"A", "#88334":"J", "#88335":"C"
									,  "#88323":"HB1", "#88324":"HB2"
									,  "#88350":"vCRT1", "#88351":"vCRT2", "#88352":"vCRT3"
									,  "#88370":"gCRT1", "#88371":"gCRT2" }
			vaccinations  := { "B"		:{"AGO":[0,0,0], "AGM":[0,0,0], "AGJ":[0,0,0], "NR":[0,0,0]}
									,  "M"	:{"AGO":[0,0,0], "AGM":[0,0,0], "AGJ":[0,0,0], "NR":[0,0,0]}
									,  "A"  	:{"AGO":[0,0,0], "AGM":[0,0,0], "AGJ":[0,0,0], "NR":[0,0,0]}
									,  "J"   	:{"AGO":[0,0,0], "AGM":[0,0,0], "AGJ":[0,0,0], "NR":[0,0,0]}
									,  "C" 	:{"AGO":[0,0,0], "AGM":[0,0,0], "AGJ":[0,0,0], "NR":[0,0,0]}
									,  "VS" 	: [0, 0, 0] 	; Zähler für 1., 2., Auffrischungsimpfungen allr Altersgruppen
									,  "VO"	: [0, 0, 0] 	; Zähler für 1., 2., Auffrischungsimpfungen > 60
									,  "VM"	: [0, 0, 0] 	; Zähler für 1., 2., Auffrischungsimpfungen 18-59
									,  "VJ"  	: [0, 0, 0] 	; Zähler für 1., 2., Auffrischungsimpfungen <18
									, "sum"	: 0
									, "patients" : {}}

		  ; Datum ins DBASE Format wandeln
             Datum := InStr(Datum, ".") ? ConvertToDBASEDate(Datum) : Datum

		  ; Quartal errechnen
			Quartal := GetQuartalEx(Datum, "QYY")

		  ; BEFGONR Index laden
			dbindex := this.IndexRead("BEFGONR")
			IF !dbindex[Quartal].1
				dbindex := this.IndexCreate("BEFGONR")

		  ; Datenbanksuche starten
			matches   	:= this.GetDBFData("BEFGONR"                                                                                       	; Name der Datenbank
														, (ARZTID ? {"DATUM":Datum, "ARZTID":ARZTID} : {"DATUM":Datum})  	; Suchparameter
														, ["ARZTID", "PATNR", "QUARTAL", "DATUM", "GNR", "SNR", "POS"]          	; Rückgabeparameter
														, (dbindex[Quartal].1 ? "recordnr=" dbindex[Quartal].1 : 0)                   	; Startindex Datensatz
														, 0                                                                                                      	; Debugoption
														, Opt)                                                                                                 	; weitere Optionen (Callback..)

		  ; Statistik anhand der Regeln erstellen
			For idx, m in matches {

				RegExMatch(m.GNR, "i)(?<fee>\d+)(?<Z>[A-Z])*(\(charge:)*(?<chr>[A-Z\d]+)*", vac)
				PatID := m.PATNR

				If RegExMatch(vacfees["#" vacfee], "i)^[BMAJC]$") {

					PatAge   	:= this.ageYears(PatDB[PatID].GEBURT, Datum)
					vacname 	:= vacfees["#" vacfee]
					patgroup 	:= PatAge > 59 ? "O" : PatAge < 18 ? "J" : "M"
					vacgroup 	:= "AG" . patgroup

					For extra, NR in vacfees.Z
						If RegExMatch(vacZ, "i)(" extra ")") {
							vacnr := NR
							break
						}

					vaccinations[vacname][vacgroup][vacnr] += 1
					vaccinations[vacname]["NR"][vacnr] += 1
					vaccinations["V" patgroup][vacnr] += 1
					vaccinations["VS"][vacnr] += 1
					vaccinations.sum += 1

					If !isObject(vaccinations.Patients[PatID]) {
						vaccinations.Patients[PatID] := { "vacnr"  	: 2^(vacnr-1)
																	, 	"vacdays"	: [(m.Datum ": " vacname vacnr)]}
					} else {
						vaccinations.Patients[PatID].vacnr += 2^(vacnr-1)
						vaccinations.Patients[PatID].vacdays.Push((m.Datum ": " vacname vacnr))
					}

				}

			}

		  ; Summen in den Gruppen berechnen
			For idx, vacname in StrSplit("BMAJC") {

				vac := vaccinations[vacname]
				For group, arr in vac
					Loop vac[group].MaxIndex()
						vac[group "SUM"] += vac[group][A_Index]

				vac.ALLSUM .= vac.AGOSUM + vac.AGMSUM + vac.AGJSUM

			}


		return vaccinations
		}

	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Funktionen um Albis Einstellungen auszulesen
	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		Karteikartenfilter()                                                                                           	{	;-- Karteikartenfilter auslesen

			this.KKFilter := Object()
			dbf     	:= new DBase(this.DBPath "\BEFTAG.dbf", false)
			res      	:= dbf.OpenDBF()
			beftag	:= dbf.GetFields("alle")
			res       	:= dbf.CloseDBF()
			dbf      	:= ""

			For idx, filter in beftag
				If !filter.removed
					this.KKFilter[filter.NAME] := {"Inhalt":StrReplace(filter.inhalt, ",", ", "), "Beschr":filter.beschr}

		return this.KKFilter
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		class Doctors extends AlbisDB                                                                         	{ 	;-- #Daten der erfassenden Ärzte auslesen

				Get(Doctor:="")                                                                                     	{	;-- Daten aller Ärzte oder nur eines Arztes erhalten

				  ; Doctor - kann die Arzt-ID, das Kürzel oder sein Name sein

					this.ID := Object()
					dbf     	:= new DBase(this.DBPath "\ERFASSER.dbf", false)
					res      	:= dbf.OpenDBF()
					match	:= dbf.GetFields("alle")
					res       	:= dbf.CloseDBF()
					dbf      	:= ""

				}

		}


	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Datenobjekte leeren
	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
		Empty(objectname)                                                                                        	{	;-- ein Objekt leeren aber nicht löschen
		return ObjRelease(this[objectname])
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		EmptyDBData()                                                                                              	{ 	;-- löscht alle Objekte welche ausgelesene Daten enthalten

			this.VSData   	:= ""
			this.KScheine 	:= ""
			this.AbrScheine	:= ""
			this.KKFilter   	:= ""

		}


	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Datenobjekte erhalten
	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
		dbIndex(dbName)                                                                                         	{
			return this.dbfStructs[dbname].dbIndex
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		GetEBMRule(RuleName)                                                                                 	{	;-- eine bestimmte EBM Regel erhalten
		return this.EBM[RuleName]
		}


	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Datenobjekte speichern (auch extern bearbeitete)
	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
		SaveVorsorgeKandidaten(VSCan, Abrechnungsquartal)                                   	{	;-- speichern im JSON Format
			quartal := LTrim(Abrechnungsquartal, "0")
			JSONData.Save(this.PathDBaseData "\VorsorgeKandidaten-" quartal ".json", VSCan, true,, 1, "UTF-8")
			FileGetSize, filesize, % this.PathDBaseData "\VorsorgeKandidaten-" quartal ".json"
		return fileSize
		}


	;――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Indizes erstellen/lesen
	;――――――――――――――――――――――――――――――――――――――――――――――――――――
		IndexCreate(dbName, IndexField="", IndexMode="", ReIndex=true)                  	{ 	;-- Indexerstellung für Albis DBase Datenbanken

			/* BESCHREIBUNG

				Die Funktion erleichtert die Erstellung eines Datenbankindex durch hinterlegte Feldnamen, welche zur Indizierung herangezogen werden.
				Ist kein Feldname für eine Datenbank hinterlegt, kann dieser über den Parameter IndexField übergeben.
				Der übergebene Feldname wird von der Funktion vorrangig behandelt!
				Ist ein Index-Feldname nicht vorhanden wird eine Fehlermeldung ausgeworfen.

			 */

			; Variablen                                                                    	;{
				; Standard Indexfelder
				static IndexFields := {"BEFUND" 	: {"IndexField":"DATUM"            	, "IndexMode":"DateToQuarter"}
											, 	"BEFTEXTE"	: {"IndexField":"DATUM"           	, "IndexMode":"DateToQuarter"}
											,	"BEFGONR"	: {"IndexField":"QUARTAL"        	, "IndexMode":""}
											, 	"LABORANF"	: {"IndexField":"EINDATUM"       	, "IndexMode":"DateToQuarter"}
											, 	"LABBLATT"	: {"IndexField":"DATUM"           	, "IndexMode":"DateToQuarter"}
											, 	"LABBUCH"	: {"IndexField":"ABNAHMDATE"	, "IndexMode":"DateToQuarter"}
											, 	"LABREAD"  	: {"IndexField":"EINDATUM"       	, "IndexMode":"DateToQuarter"}
											, 	"LABUEBPA"	: {"IndexField":"ABDATUM"        	, "IndexMode":"DateToQuarter"}}

				if !IndexMode
					IndexMode := IndexFields[dbName].IndexMode

				If !IndexField
					IndexField := IndexFields[dbName].IndexField

				If !IndexField
					throw A_ThisFunc	": Keine Indizierung ohne Angabe eines Indexfeldes [IndexField] möglich!"

			;}

			; Dateinamen überprüfen                                               	;{
				If !(this.dbName := this.DBaseFileExist(dbName))
					Throw A_ThisFunc	": Die Datenbank ["  dbName "] ist nicht vorhanden!"
			;}

			; Information auslesen und Indizierbarkeit prüfen             	;{
				dbf := new DBASE(this.DBPath "\" dbName ".dbf", this.debug)
				If !this.dbfStructsPath.HasKey(dbName) {
					this.dbfStructsPath(dbName) 		                                         	;-- legt das Objekt für die Abbildung der DBase Datei an
					this.dbfStructs[dbname].dbfields     	:= dbf.dbfields	        	;-- Feldnameninformationen hinzufügen
					this.dbfStructs[dbname].maxrecords	:= dbf.records	        	;-- Anzahl der Recordsets
					this.dbfStructs[dbname].lastupdate 	:= dbf.lastupdate        	;-- letzte Änderung
				}
				If (StrLen(IndexField) > 0) {
					If !this.FieldNameExist(IndexField)
						throw A_ThisFunc	": Die Datenbank ["  dbName "] kann nicht indiziert werden!`n"
										    	.	"  Der Feldnahme [" IndexField "] ist nicht enthalten."
				}
			;}

			; Index erstellen (wird als JSON String gespeichert)          	;{
				this.dbfStructs[dbname].dbIndex := dbf.CreateIndex(this.PathDBaseData "\DBIndex_" dbName ".json", IndexField, IndexMode, ReIndex)
				dbf.CloseDBF()
				dbf := ""
			;}


		return this.dbfStructs[dbname].dbIndex
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		IndexRead(dbName, IndexField:="", ReIndex:=false, IndexMode:="")                 	{ 	;-- Indizes laden oder bei Bedarf erstellen

				dbName       	:= RegExReplace(dbName, "i)\.dbf$")
				DataFileName 	:= "DBIndex_" dbName ".json"
				DataFilePath		:= this.PathDBaseData "\" DataFileName

			; ruft die Indizierungsfunktion auf falls ein Index noch nicht erstellt wurde
				If FileExist(DataFilePath) && !ReIndex
					dbIndex	:= JSONData.Load(DataFilePath, "", "UTF-8")
				else {
					TrayTip, Addendum für Albis on Windows, % "Die Datenbank " dbName " wird gerade indiziert.", 4
					dbIndex	:= this.IndexCreate(dbName, IndexField, IndexMode, ReIndex)
				}

				If !this.dbfStructsPath(dbName)
					this.ThrowError(	"Ein weiteres Objekt für dbfStructs ["  dbName "] kann nicht angelegt werden!")

				this.dbfStructs[dbName].dbIndex := dbIndex

		return dbIndex
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		TextDBIndex(Reindex:=false)                                                                            	{	;-- erstellt zwei Indizes der BEFTEXT.dbf

			DBIndex        	:= {"Quartal":{}, "LFDNR":{}}
			DbIndexFile1  	:= Addendum.DBPath "\DBase\DBIndex_BEFTEXTE.json"
			DbIndexFile2 	:= Addendum.DBPath "\DBase\DBIndex_BEFTEXTE_LFDNR-Sort.json"
			dbf               	:= new DBASE(Addendum.AlbisDBPath "\BEFTEXTE.dbf", 2)

			If !ReIndex && FileExist(DbIndexFile1) && FileExist(DbIndexFile2) {
				DBIndex.Quartal	:= JSONData.Load(DbIndexFile1, "", "UTF-8")
				DBIndex.LFDNR 	:= JSONData.Load(DbIndexFile2, "", "UTF-8")
			}

		  ; Indizierung wird auch ausgeführt wenn die letzte Indizierung mehr als 30 Tage her ist
			today := A_YYYY A_MM A_DD "000000"
			IndexFromDay := today
			IndexFromDay += 30, days
			IndexStartRecord := (DBIndex.Quartal.LastIndex && DBIndex.Quartal.LastIndex < IndexFromDay) ? DBIndex.Quartal.LastRecord : 0

		  ; Indexerstellung
			If !FileExist(DbIndexFile1) || !FileExist(DbIndexFile2) || IndexStartRecord || ReIndex {

				; Indexer erstellt zwei Index Objekte, eines mit Zuordnung zu LFDNR und das andere zum Quartal ;{

					 ; für Fortschrittsanzeige
						dbf.ShowAt	:= 100
						LFDNRsteps  	:= dbf.records < 1200 ? Floor(dbf.records/100) : 1200    ; Indexschritte (bei großen Datenbanken alle 1200 Datensätze)

					; Datenbankzugriff herstellen
						filepos       	:= dbf.OpenDBF()

					; Index wird vergrößert wenn der letzte gespeicherte Index ein Monat her ist
						If IndexStartRecord
							dbf._SeekToRecord(IndexStartRecord)

					; indizieren (liest Zeile für Zeile)
						Loop % (dbf.records - IndexStartRecord) {

							set := dbf.ReadRecord(["PATNR", "DATUM", "LFDNR", "POS", "TEXT"])

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
									DBIndex.LFDNR[set.LFDNR] := [dbf.recordnr, dbf.filepos]       	; Datensatz Nummer im Objekt als Einsprungpunkt sichern
							}

						}

				;}

				DBIndex.Quartal.LastIndex   	:= dbf.lastupdatedBase
				DBIndex.Quartal.LastRecord 	:= dbf.records
				DBIndex.Quartal.LastFilePos 	:= dbf.filepos

			  ; Datenbankzugriff beenden
				filepos	:= dbf.CloseDBF()
				dbf    	:= ""

			  ; Indizes speichern
				JSONData.Save(DbIndexFile1, DBIndex.Quartal	, true,, 1, "UTF-8")
				JSONData.Save(DbIndexFile2, DBIndex.LFDNR 	, true,, 1, "UTF-8")

			}

		return DBIndex
		}


	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Verschiedenes
	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
		BefundKuerzel(KUERZEL:="", staticObj:=false)                                                   	{	;-- Schriftart-, Farb-, ...einstellungen erhalten

			; verwendet die KUERZEL als Schlüsselnamen

				static BEFKUERZ

				dbf        	:= new DBASE(this.DBpath "\BEFKUERZ.dbf")
				res        	:= dbf.OpenDBF()
				matches 	:= dbf.GetFields()               			; Datenbanksuche durchführen
				res        	:= dbf.CloseDBF()
				dbf       	:= ""

				retObj   	:= Object()
				For idx, obj in matches {
					key := obj.Delete("KUERZEL")
					obj.Delete("removed")
					obj.Delete("recordnr")
					retObj[key] := obj
				}

		return retObj
		}


	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Hilfsfunktionen
	;―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		AppendMatches(matches, properties)                                                                	{	;-- appends connected database values

			; ## no tests until now ##
			; there are two features inside:
			; first: 		this function will automatically create an individual index based on given parameters
			;				if there's no index created at time, the function will not waste time create an index for first. It reads and parse data at same time it's indexing.
			; second: 	it loads the data you want to have
			; dependencies: class_DBASE and class_JSON.ahk

				static dbIndex, lastfilepath

			; extract data from properties                               	;{
				filepath     	:= properties.DBFilePath
				fieldFrom  	:= properties.link[1]              	; TEXTDB 	in BEFUND means
				fieldTo      	:= properties.link[2]                	; LFDNR 	in BEFTEXT
				fieldIndex  	:= properties.link[3]	            	; POS    	in BEFTEXT means the sequence of parts of the text to append (a second index)
				appendTo  	:= properties.append[1]          	; value of INHALT (BEFUND.dbf) must append with
				appendFrom	:= properties.append[2]			; values from TEXT (BEFTEXT.dbf)
			;}

			; create outfields parameter                                 	;{
				outfields.Push(fieldTo)
				outfields.Push(appendFrom)
				If fieldIndex
					outfields.Push(fieldIndex)
			;}

			; collect data to append                                     	;{
				option    	:= ""
				startrecord := 0
				appendTo := Array()
				connected 	:= new DBASE(filepath, 1)
				res             	:= connected.OpenDBF()

				For mIdx, m in matches {

					If (m[fieldFrom] = 0)
						continue

					pattern  	:= {(fieldTo):m[fieldFrom]}
					appends	:= connected.SearchFast(pattern, outfields, startrecord, (mIdx < 2 ? "" : "next") )
					startrecord := connected.foundrecord

					If Mod(mIdx, 20) = 0
						SciTEOutput(m[fieldFrom] ":" appends.MaxIndex() " | " startrecord " | " connected.filepos_start " | " connected.filepos_found)

					tmp:=Array()
					for aIdx, a in appends
						tmp[a[fieldIndex]+1] := a[appendFrom]

					for idx, t in tmp
						matches[mIdx][appenendTo] .= t

				}

				res            	:= connected.CloseDBF()
				connected := ""

			;}

			; save index for faster access next time
				JSONData.Save(A_Temp "\dump.json" , matches, true,, 1, "UTF-8")
				Run % A_Temp "\dump.json"

		return matches
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		dbfStructsPath(dbName)                                                                                	{	;-- legt das Objekt dbfStructsPath an

			 If !IsObject(this.dbfstructs[dbName])
				this.dbfStructs[dbName] := Object()

		return IsObject(this.dbfstructs[dbName]) ? true : false
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		CreateFilePath(path)                                                                                      	{ 	;-- erstellt Dateipfade falls diese nicht vorhanden sind

			If !InStr(FileExist(path "\"), "D") {
				FileCreateDir, % path
				return ErrorLevel ? 0 : 1
			}

		return 1
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		DBaseFileExist(dbName)                                                                                  	{	;-- Datenbank ist vorhanden?

			; Dateinamen überprüfen
				dbName 	:= RegExReplace(dbName, "i)\.dbf$")
				If !FileExist(this.DBPath "\" dbName ".dbf")
					Throw A_ThisFunc	": Die Datenbankdatei <" dbName ">`nbefindet sich nicht im Ordnerpfad:`n<" this.DBPath ">"

		return dbName
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		FieldNameExist(fieldLabel)                                                                              	{	;-- ist der übergebene Feldname vorhanden?

			idxFieldExist := false
			If !IsObject(this.dbfStructs[this.dbname].dbfields)
				Throw A_ThisFunc	": Fehler beim Auslesen von Information aus ["  this.dbName "]!`n"
								. 	"Informationen zur Datenbankstruktur konnten nicht erzeugt werden.."

			For fLabel, fData in this.dbfStructs[this.dbname].dbfields
				If (fLabel = fieldLabel) {
					idxFieldExist := true
					break
				}

		return idxFieldExist
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		SwapQuarter(QQYY)                                                                                     	{ 	;-- Quartal Forrmat QQYY zu YYQQ
			QQYY := SubStr("0" QQYY, -3)
		return SubStr(QQYY, 3, 2) SubStr(QQYY, 1, 2)
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		SwapQuarterS(QYY)                                                                                      	{	;-- Quartal Format QYY zu YYQ
		return SubStr(QYY, 2, 2) SubStr(QYY, 1, 1)
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		ageYears(birthday, thisday)                                                                             	{	;-- Altersjahrberechnung
			;Format für beide Daten ist YYYYMMDD
			timeDiff := HowLong(birthday, thisday)
			return timeDiff.Years
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		Geschlecht(AlbisKodierung)                                                                             	{	;-- Umkodieren der Geschlechtziffern in einen Buchstaben
			sex := AlbisKodierung
			return (sex = 1 ? "M": sex = 2 ? "W" : sex)
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		SortMatchesByKey(matches, SortByKey:="")                                                        	{	;-- Array in key:object() umwandeln

			If !SortByKey
				return matches

			reordered := Object()
			For mIndex, m in matches {

				key := m.Delete(SortByKey)
				If !IsObject(reordered[key])
					reordered[key] := Array()
				reordered[key].Push(m)

			}

		return reordered
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		KrankenscheinArt(ksart)                                                                                  	{	;-- Computer-Mensch-Übersetzer

			static ScheinN	:= {"Abrechnung":0, "1":1 , "Überweisung":2, "3":3, "Notfall":4, "Privat":5}
			static ScheinR   	:= {"#0":"Abrechnung", "#1":"1", "#2":"Überweisung", "#3":"3", "#4":"Notfall", "#5":"PRIVAT"}

			If RegExMatch(ksart, "\d")
				return ScheinR["#" ksart]

			If (ksart = "alle")
				return "\d"

			rxStr := "["
			For idx, ART in StrSplit(ksart, ",") {
				rxStr .= ScheinN[Trim(ART)]
				}

		return rxStr "]"
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		ThrowError(msg, errorLine:="")                                                                       	{ 	;-- "wirft" Fehlermeldungen aus
			this.db.CloseDBF()
			this.db := ""
			throw A_ThisFunc "," errorLine ": " msg
		}

}

class PatDBF	{                                                                                       	;-- PATIENT.dbf Handler

	;-- gibt nur benötigte Daten der albiswin\db\PATIENT.DBF zurück
		__New(basedir="", infilter="", opt="") {

			; basedir kann man leer lassen, der Albishauptpfad wird aus der Registry gelesen
			; Rückgabeparameter ist ein Objekt mit Patienten Nr. und dazugehörigen Datenobjekten (die key's sind die Feldnamen in der DBASE Datenbank)
			; lade nur Patienten die in den letzten 10 Jahren behandelt wurden
			; minDate := 20100101
			;
			; letzte Änderung: 27.12.2021

			; kein basedir - sollte ..\albiswin\db\ enthalten dann wird hier versucht den Pfad aus der Windows Registry zu bekommen
				this.Patients := Object()
				SplitPath, basedir,, basedir
				If (StrLen(basedir)=0 || !InStr(FileExist(basedir), "D")  || !RegExMatch(basedir, "\\db\\*.*\.*[a-z]*$")) {
					Addendum.AlbisDBPath := GetAlbisPath() "\db"
					SciTEOutput(	"BaseDir changed from " (StrLen(basedir)=0 ? " none existing path ('basedir' param was empty) ":" wrong dir <<" basedir ">>")
									. 	" to Albis DBASE path: "  Addendum.AlbisDBPath)
					this.basedir	:= Addendum.AlbisDBPath
				}
				else
					this.basedir := basedir

			; Optionen
				this.dbg         	:= RegExMatch(opt, "i)dbg(?<nr>\d+)", dbg)     	? dbgnr 	: false
				this.moreData	:= RegExMatch(opt, "i)moreData\s*\=\s*(true|1)") 	? true    	: false
				this.AllData   	:= RegExMatch(opt, "i)allData\s*=\s*(true|1)")     	? true     	: false
				this.minDate   	:= RegExMatch(opt, "i)minDate\s*=\s*(?<Date>(\d{1,2}\.\d{1,2}\.(\d{4}|\d{2}))|((19|20)[0-9]{2})((0[1-9])|(1[0-2]))(0[1-9]|1[0-9]|2[0-9]|3[01]))", min) ? minDate : 0
				SciTEOutput("this.minDate: " this.minDate)
				;~ this.minDate 	:= InStr(this.minDate, ".") ? ConvertToDBASEDate(this.minDate) : this.minDate
				If RegExMatch(this.minDate, "((?<Y1>\d{4}|\d{2})\.?<M1>\d{1,2}\.(?<D1>\d{1,2}))|((?<D2>\d{1,2})\.(?<M2>\d{1,2})\.(?<Y2>\d{4}|\d{2}))", _) {
					Yl := StrLen(_Y1) > 0 ? StrLen(_Y1) : StrLen(_Y2)
					If (Yl = 3)
						throw A_ThisFunc ": Falsches Datum (Jahr '" _Y1 _Y2  "') übergeben!"
					this.minDate := Yl=2 ? SubStr(A_YYYY, 1, 2) : ""
					this.minDate .= _Y1 && _M1 && _D1
					                	? 	_Y1 . SubStr("00" _M1, -1) . SubStr("00" _D1, -1)
				                  		: 	_Y2 . SubStr("00" _M2, -1) . SubStr("00" _D2, -1)
					SciTEOutput("this.minDate: " this.minDate)
				}

			; die minimal notwendigsten Daten
				If !this.moreData
					this.infilter	:= !IsObject(infilter) 	? ["NR", "NAME", "VORNAME", "GEBURT", "GESCHL", "MORTAL"
																	,	"TELEFON", "TELEFON2", "FAX", "LAST_BEH", "RKENNZ", "NAMENORM"] : infilter
				else
					this.infilter	:= !IsObject(infilter) 	? ["NR", "NAME", "VORNAME", "GEBURT", "GESCHL", "NAMENORM"
																	, 	"TELEFON", "TELEFON2", "FAX",	"HAUSARZT", "ARBEIT", "MORTAL", "LAST_BEH", "DEL_DATE", "RKENNZ"] : infilter

			; liest alle Patientendaten in ein temporäres Objekt
				database 	:= new DBASE(this.basedir "\PATIENT.dbf", this.debug)
				res        	:= database.OpenDBF()
				matches	:= database.GetFields(this.infilter)
				res         	:= database.CloseDBF()

			; temp. Objekt wird nach NR sortiert (Patientenindex)
				For idx, m in matches {

					; bestimmte Datensätze aussortieren (optional)
					If !this.AllData && InStr(m.RKENNZ, "*") || (m.LAST_BEH < this.minDate) || (!m.NAME && !m.VORNAME)
						continue

					; Sortieren und Namensnormierungen anlegen
					strObj := Object()
					For key, val in m
						If (key <> "NR" && StrLen(val) > 0) {
							strObj[key] := val
							If (key = "NAME")
								strObj.DiffN 	:= xstring.TransformGerman(val)
							else if (key = "VORNAME")
								strObj.DiffVN:= xstring.TransformGerman(val)
						}

					; nach Patientennummern indiziertes Objekt erstellen
					this.Patients[m.NR] := strObj

				}

			; Mail Adressen aus PATEXTRA beziehen
				database 	:= new DBASE(this.basedir "\PATEXTRA.dbf", this.debug)
				res        	:= database.OpenDBF()
				matches	:= database.GetFields(["NR", "POS", "TEXT"])      ; POS 93 enthält EMail, POS 97 Ausnahmekennziffern
				res         	:= database.CloseDBF()

			; und hinzufügen
				For idx, m in matches
					If (m.POS = "93") && this.Patients.HasKey(m.NR)
						If RegExMatch(m.Text, "i)^.*@.*\.[a-z]+$")
							this.Patients[m.NR].EMAIL := m.Text

		}

	;-- Patientendaten
		Get(PatID, key) {
		return this.Patients[PatID][key]
		}

		GetPatID(searchobj) {

		}

	; Patienten ID Suche per String-Similarity Funktion
		StringSimilarityID(name1, name2, diffmin:=0.09) {

			matches	:= Array()
			minDiff 	:= 100
			dname1	:= InStr(name1, " ") ? 1 : 0
			dname2	:= InStr(name2, " ") ? 1 : 0

		  ; Suche ohne Stringmatching als erstes
			For PatID, Patient in this.Patients {
				DbName1 := Patient.Name, DbName2 := Patient.VORNAME
				If (DbName1 = name1 && DbNAME2 = name2) || (DbName1 = name2 && DbNAME2 = name1) {
					;SciTEOutput("PatDB: " PatID)
					matches.Push(PatID)
					return matches.1
				}
			}

			name1  	:= xstring.TransformGerman(name1)              ; class string
			name2  	:= xstring.TransformGerman(name2)

			NVname 	:= RegExReplace(name1 . name2, "[\s\-]+")
			VNname 	:= RegExReplace(name2 . name1, "[\s\-]+")

			For PatID, Patient in this.Patients 		{

				dVorname	:= InStr(Patient["VORNAME"]	, " ")	? 1 : 0
				dName   	:= InStr(Patient["NAME"]      	, " ")	? 1 : 0
				DbName  	:= (dName ? StrSplit(Patient["DiffN"], " ").1 : Patient["DiffN"]) . (dVorname ? StrSplit(Patient["DiffVN"], " ").1 : Patient["DiffVN"])
				DbName	:= RegExReplace(DbName, "[\s\-]+")
				DiffA     	:= StrDiff(DBName, NVname)
				DiffB     	:= StrDiff(DbName, VNname)
				Diff        	:= DiffA <= DiffB ? DiffA : DiffB

				If (Diff <= diffmin)
					matches.Push(PatID)

				If (Diff < minDiff)
					minDiff	:= Diff, bestDiff := PatID, bestNn := Patient.Name, bestVn := Patient.VORNAME, bestGd := Patient.GEBURT

			}

		return matches.1
		}

	; erweitertete Patienten ID Suche
		StringSimilarityEx(name1, name2, diffmin:=0.09) {

			PatIDs    	:= Object()
			minDiff 	:= 100
			NVname 	:= RegExReplace(name1 . name2, "[,.\-\s]")
			VNname 	:= RegExReplace(name2 . name1, "[,.\-\s]")

			For PatID, Patient in this.Patients 		{

				DbName	:= RegExReplace(Patient.Nn Patient.Vn, "[\s]")
				If (StrLen(dbName) = 0)
					continue

				DiffA	:= StrDiff(DBName, NVname), DiffB := StrDiff(DbName, VNname)
				Diff  	:= DiffA <= DiffB ? DiffA : DiffB

				If (Diff <= diffmin)
					PatIDs[PatID] := Pat

				If (Diff < minDiff)
					minDiff	:= Diff, bestDiff := PatID, bestNn := Patient.Name, bestVn := Patient.VORNAME, bestGd := Patient.GEBURT

			}

			PatIDs.Diff := {"minDiff":minDiff, "bestID":bestDiff, "bestNn":bestNn, "bestVn":bestVn, "bestGd":bestGd}

			If (PatIDs.Count() > 1)
				return PatIDs

		return
		}

}

DBASEStructs(AlbisPath, DBStructsPath:="", debug=false) {                        	;-- analysiert alle DBF Dateien im albiswin\db Ordner

	; schreibt die ausgelesenen Strukturen als JSON formatierte Datei

		dbfStructs	:= Object()
		dbfiles    	:= Array()
		AlbisDBpath:= RegExReplace(AlbisPath, "i)\\db.*$") "\db"

	; zählt für die Dateianzeige zunächst einmal die vorhandenen Dateien im Verzeichnis
		Loop, Files, % AlbisDBpath "\*.dbf"
			dbfiles.Push(StrReplace(A_LoopFileName, ".dbf"))

	; für formatierte Ausgabe
		dbFilesMax := dbfiles.Count()
		dbIL := StrLen(dbFilesMax) - 1

	; öffnet jede DBASE Datei und liest den Header der Datei aus
		dbNr := 0
		For dbfNr, dbfName in dbfiles {

				If debug && (Mod(dbfNr, 20) = 0)
					ToolTip, % "lese: " SubStr("000" (dbNr ++), -1 * dbIL) "/" dbFilesMax ": " dbfName

			; Zeitstempel
				FileGetTime, accessed,  % AlbisDBPath "\" dbfName ".dbf"	, A
				FileGetTime, modified,  % AlbisDBPath "\" dbfName ".dbf"	, M

			; Datenbankgrößen
				FileGetSize, SizeDBF	, % AlbisDBPath "\" dbfName ".dbf"		, K
				FileGetSize, sizeDBT	, % AlbisDBPath "\" dbfName ".dbt" 	, K
				FileGetSize, sizeMDX	, % AlbisDBPath "\" dbfName ".mdx"	, K
				dbfSizeMax 	+= sizeDBF
				dbfSizeMax 	+= sizeDBT
				mdxSizeMax	+= sizeMDX

			; Datenbankstruktur

				dbf  	:= new DBASE(AlbisDBPath "\" dbfName ".dbf", false)

				If (dbf.Version = 0x8b || dbf.Version = 0x3) {
					dbfStructs[dbfName]                     	:= Object()
					dbfStructs[dbfName].Nr               	:= dbfNr
					dbfStructs[dbfName].dbfields        	:= isObject(dbf.dbfields) 	? dbf.dbfields 	: "error reading dbase file"
					dbfStructs[dbfName].fields         	:= isObject(dbf.fields)    	? dbf.fields    	: "error reading dbase file"
					dbfStructs[dbfName].header      	:= isObject(dbf.dbstruct)	? dbf.dbstruct 	: "error reading dbase file"
					dbfStructs[dbfName].headerLen  	:= dbf.headerLen
					dbfStructs[dbfName].lendataset  	:= dbf.lendataset
					dbfStructs[dbfName].records      	:= dbf.records
					dbfStructs[dbfName].lastupdate  	:= dbf.lastupdateDate
					dbfStructs[dbfName].lastupdateE	:= dbf.lastupdateEng
					dbfStructs[dbfName].sizeDBF     	:= sizeDBF
					dbfStructs[dbfName].sizeDBT     	:= sizeDBT
					dbfStructs[dbfName].sizeMDX     	:= sizeMDX
					dbfStructs[dbfName].accessed     	:= accessed
					dbfStructs[dbfName].modified     	:= modified
				}

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
			file.Write(";   .dbt       `t: " 		Round(dbfSizeMax/1024/1024	, 2)	" GB" 						                        	"`n")
			file.Write(";   .mdx       `t: " 	Round(mdxSizeMax/1024/1024, 2)	" GB" 		    			                        	"`n")
			file.Write(";   Summe     `t: " 	Round((dbfSizeMax+dbfSizeMax+mdxSizeMax)/1024/1024, 2) " MB"    			"`n")
			file.Write("; -------------------------------------------------------------------------------------------------`n")
			file.Write(JSON.Dump(dbfStructs,,2))
			file.Close()

	}

return dbfstructs
}

GetDBFData(DBpath,p="",out="",s=0,opts="",dbg=0,dbgOpts="") {         	;-- holt Daten aus einer beliebigen Datenbank

	/*
		p     	- pattern as object like
		out  	- returned values as object
		s     	- seek to position
		opts	-
	 */

	dbfile      	:= new DBASE(DBpath, dbg, dbgParams)
	startrec 	:= (s > 0 ? (s - (dbfile.headerlen + 1)) / dbfile.lendataset : 0)

	res        	:= dbfile.OpenDBF()
	matches 	:= dbfile.Search(p, startrec, opts)
	res         	:= dbfile.CloseDBF()
	dbfile    	:= ""

	If !IsObject(out)
		return matches

	data := Array()
	For midx, m in matches {
		strObj:= Object()
		For cidx, ckey in out
			strObj[ckey] := m[ckey]
		data.Push(strObj)
	}

return data
}

Leistungskomplexe(PatID=0, StartQuartal="2009-4", SaveToPath="") {   	;-- sammelt, erstellt eine Datenbank für den Abrechnungshelfer

	; im Grunde schon obsolete Funktion ist im Abrechnungsassistenten besser gelöst

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
		filepos   	:= befund.OpenDBF()
		startpos  	:= Floor((BefundDBF[StartQuartal]-befund.recordsStart)/befund.LenDataSet)
		ziffern    	:= befund.SearchFast({"KUERZEL": "lko"}, ["DATUM", "PATNR", "INHALT"], startpos)
		filepos   	:= befund.CloseDBF()

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
		AbrDBF := FileOpen(Addendum.Dir "\logs'n'data\_DB\PatData\ZIFFERN.adm", "w", "UTF-8")
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

ReadPatientDBF(basedir="", infilter="", opt="", minDate=20100101) {       	;-- gibt nur benötigte Daten der albiswin\db\PATIENT.DBF zurück

	; basedir:  	kann man leer lassen, der Albishauptpfad wird von der Funktion aus der Registry gelesen
	; mindate:  	lade nur Patienten die in den letzten 10 Jahren behandelt wurden
	; 					z.B. minDate := 20100101
	; Rückgabeparameter ist ein Objekt geordnet nach Patienten Nr. und den dazu gehörenden Daten (die Schlüsselnamen sind die Feldnamen in der DBASE Datenbank)

		local PatDB

		PatDB := Object()
		If (StrLen(basedir) = 0) || !InStr(FileExist(basedir), "D")
			basedir	:= GetAlbisPath() "\db"

	; für eine Abrechnungsüberprüfungen die geschätzt minimal notwendigste Datenmenge
		If !IsObject(infilter)
			infilter 	:= ["NR", "NAME", "VORNAME", "GEBURT", "GESCHL", "MORTAL", "LAST_BEH", "RKENNZ"]

	; Optionen
		debug := AllData := 0
		If RegExMatch(opt, "^\d$", dbg) || RegExMatch(opt, "i)debug(?<bg>\d)\s", d)
			debug := dbg
		If RegExMatch(opt, "i)allData")
			AllData := true

	; liest alle Patientendaten in ein temp. Objekt ein
		database 	:= new DBASE(basedir "\PATIENT.dbf", debug)
		res        	:= database.OpenDBF()
		matches	:= database.GetFields(infilter)
		res         	:= database.CloseDBF()

	; temp. Objekt wird nach NR (Patientennummerierung) sortiert
		For idx, m in matches {

			If !AllData
				If (InStr(m.RKENNZ, "*") || !m.LAST_BEH || m.LAST_BEH<minDatum || (!m.NAME && !m.VORNAME))
					continue

			obj := Object()
			For key, val in m
				If !RegExMatch(key, "i)^(NR|removed)$") ;&& (StrLen(val) > 0)
					obj[key] := val

			PatDB[m.NR] := obj

		}

	; Mail Adressen aus PATEXTRA beziehen
		database 	:= new DBASE(basedir "\PATEXTRA.dbf", debug)
		res        	:= database.OpenDBF()
		matches	:= database.GetFields(["NR", "POS", "TEXT"])      ; POS 93 enthält EMail, 97 Ausnahmekennziffern
		res         	:= database.CloseDBF()
		database 	:= ""

		For idx, m in matches
			If (m.POS = "93") && PatDB.haskey(m.NR)
				If RegExMatch(m.Text, "i)^.*@.*\.[a-z]+$")
					PatDB[m.NR].EMAIL := m.Text

	; mit opt und 'saveAs=FilePath' die extrahierte Daten speichern
	; durch Endungserkennung können die Daten im json und csv (; getrennt) gespeichert werden
		If RegExMatch(opt, "i)\bsaveAs\s*=\s*(?<path>[\._\-\pL\s\\:]+.*?\\)(?<name>.+)+?\.(?<ext>[\w_]+)", file) {
			If FilePathExist(filepath)
				If (FileExt = "json")
					JSONData.Save(filepath . filename "." FileExt, PatDB, true,, 1, "UTF-8")
		}


return PatDB
}

ReadDBASEIndex(admDBPath, DBASEName, ReIndex:=true) 	{              	;-- liest erstellte Indizes oder bei Bedarf eine Indexdatei

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
			dbIndex	:= AlbisDB.IndexCreate(DBASEName,, ReIndex)
			AlbisDB 	:= ""

		}

return dbIndex
}

Convert_IfapDB_Wirkstoffe(filepath, savePath:="")                	{              	;-- konvertiert die Ifap Wirkstoffdatenbank in eine reine Textdatei

	;base:= "M:\albiswin\ifapDB\ibonus3\Datenbank"
	If !(dbf := FileOpen(filePath, "r", "CP1252")) {
			MsgBox, % "Dbf - file read failed.`n" filePath
			return 0
	}

	If !(save := FileOpen( (!savePath ? A_ScriptDir : savePath) "\Wirkstoffe.txt", "w", "UTF-8")) {
		MsgBox, % "Can't open file to save data.`n" (!savePath ? A_ScriptDir : savePath) "\Wirkstoffe.txt"
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
      loop % re_obj.Count()      {

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
; HILFSFUNKTIONEN - STRING HANDLING
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
class xstring                                                              	{                          	;-- wird alle String Funktionen von Addendum enthalten

	; benötigt Addendum_Datum.ak
	; letzte Änderung 25.11.2021

		static rx
		RegExStrings()                                	{        	;-- erstellt die RegExString für die string Klassenbibliothek

			static __ := xstring.RegExStrings()

			rxb := {"S"             	: "[\s,]+"
					, 	"Person1"  	: "[A-ZÄÖÜ][\pL\-]+[\-\s]*([A-ZÄÖÜ][\pL\-]+)*"
					, 	"Person2"  	: "[A-ZÄÖÜ][\pL\-]+([\-\s]*[A-ZÄÖÜ][\pL\-]+)*"
					,	"Date1"        	: "(?<First>(\d{1,2}\.\d{1,2}\.(\d{4}|\d{2}))|(\d{1,2}\.\d{1,2}\.)|(\d{1,2}\.))[\s\-]*(?<Last>\d{1,2}\.\d{1,2}\.(\d{4}|\d{2}))"
					,	"fileExt"      	: "\.[A-Za-z]{,3}$"}

			this.rx := {	"fileExt"	    	: rxb.FileExt
				        	, 	"umlauts"   	: {"Ä":"Ae", "Ü":"Ue", "Ö":"Oe", "ä":"ae", "ü":"ue", "ö":"oe", "ß":"sz"}
				        	, 	"Extract"      	: "^\s*(?<NN>" rxb.Person1 ")" rxb.S "(?<VN>" rxb.Person1 ")" rxb.S "(?<Title>.*)*"
				        	, 	"Names1"  	: "^\s*(?<NN>" rxb.Person1 ")" rxb.S "(?<VN>" rxb.Person1 ")" rxb.S
				        	, 	"Names2"  	: "^\s*(?<NN>" rxb.Person2 ")" rxb.S "(?<VN>" rxb.Person2 ")" rxb.S
				        	; Addendums Dateibezeichnungsregel für Befunde: z.B. "Muster-Mann, Max-Ali, Entlassungsbrief Kardiologie v. 10.10.2020-18.10.2020"
				        	, 	"CaseName"	: "^\s*(?<N>.*)?\s*,\s*(?<I>.*)?\s*v\.\s*(?<D1>\d+\.\d*\.*\d*)*\s*\-*\s*(?<D2>\d+.\d+.\d+)*(?<E>\s*\.[A-Za-z]+)*$"
				        	,	"DocDates"  	: "i)" rxb.Date1 "?([.\sa-zß\-_]|$)"
				        	; letztes Datum im Dateinamen (der letzte Datumsstring auf den weder ein Whitespace, Zahl, Punkt oder ein Minus folgt, oder am Ende des String steht
				        	,	"LastDate" 	: "i)(?<Date>\d{1,2}\.\d{1,2}\.(\d{4}|\d{2})?)([^\-\d]|$)"}

			For key, rxString in rxb
				If !this.rx.haskey(key)
					this.rx[key] := rxString

		}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; diverse Stringfunktionen
	;----------------------------------------------------------------------------------------------------------------------------------------------
		EscapeStrRegEx(str)                           	{      	;-- formt einen String RegEx fähig um
			For idx, Char in StrSplit("[](){}*\/|-+.?!^$")
				str := StrReplace(str, char , "\" char)
		return str
		}

		Flip(str)                          	               	{        	;-- String reverse
			VarSetCapacity(new, n:=StrLen(str))
			Loop % n
				new .= SubStr(str, n--, 1)
		return new
		}

		Karteikartentext(str)                        	{        	;-- entfernt Namen, Dateierweiterung und überflüssige Leerzeichen
			str := this.Replace.FileExt(str)                    	; Dateiendungen entfernen
			str := this.Replace.Names(str)                	; Patientennamen entfernen
			str := RegExReplace(str, "^[\s,]*")            	; einzeln stehende Kommata entfernen
		return RegExReplace(str, "\s{2,}", " ")           	; Leerzeichenfolgen kürzen
		}

		RemovePunctuation(str)                  	{	      	;-- Interpunktion entfernen
		return RegExReplace(str, "[-,'\.;:_\+\*\(\)\{\}\[\]\/\\]")
		}

		TransformGerman(str)                     	{        	;-- für Stringmatching Algorithmen
			str   	:= StrReplace(str, "ä", "ae")
			str   	:= StrReplace(str, "ö", "oe")
			str   	:= StrReplace(str, "ü", "ue")
			str   	:= StrReplace(str, "ß", "ss")
		return str
		}

		MetaSoundex(str)                            	{         	;-- Soundex Umwandlung

			; Siehe: https://de.wikipedia.org/wiki/Soundex
			; https://github.com/Arthur-Akbarov/Configs/blob/fc57d892a0211ac912bdfa1fa3bc0c08b356d557/AutoHotKey/files/ac'tivAid/internals/ac'tivAid_func.ahk

			static msdx

			If !IsObject(msdx) {

				msdx := Object()
							                	; Doppelbuchstaben und Sonderzeichen vereinfachen
				msdx[1] := { "bb":"b", "cc":"c", "dd":"d", "ff":"f", "gg":"g", "hh":"h", "jj":"j", "kk":"k", "ll":"l", "mm":"m", "nn":"n", "pp":"p"
					    	      , "rr":"r", "ss":"s", "ß": "s", "tt":"t", "vv":"v", "ww":"w", "xx":"x", "zz":"z", "ú":"u", "ù":"u", "û":"u", "ü":"u", "á":"a"
							      , "à":"a", "â":"a", "å":"a", "ã":"a", "æ":"a", "ä":"a", "é":"e", "è":"e", "ê":"e", "ë":"e", "ó":"o", "ò":"o", "ô":"o"
							      , "ø":"o", "õ":"o", "ö":"o", "ý":"y", "í":"i", "ì":"i", "î":"i", "ï":"i", "ç":"c", "ÿ":"y", "ñ":"n", "ð":"d"}
												; Buchstabenkombinationen phonetish vereinfachen
				msdx[2] := { "v":"f", "w":"f", "sch":"s", "chs":"s", "cks":"s", "ch":"k", "cz":"s", "ks":"s", "ts":"s", "tz":"s", "sz":"s", "ck":"k"
								  , "dt":"t", "pf":"f", "ph":"f", "ch":"o", "qu":"kf", "ct":"kt"}
												; Buchstaben phonetisch Zahlen zuordnen
				msdx[3] := {"ie":"7", "ei":"7", "ai":"7", "ui":"7", "oi":"7", "a":"", "e":"", "i":"", "o":"", "u":"", "h":"", "w":"", "y":"", "b":"1"
								  , "f":"1", "p":"1", "v":"1", "c":"2", "g":"8", "j":"2", "k":"8", "q":"8", "ß":"2", "s":"2", "x":"2", "z":"2", "d":"3"
								  , "t":"3", "l":"4", "m":"5", "n":"5", "r":"6"}

			}

			str := this.RemovePunctuation(str)
			For each, replacers in msdx
				For chars, replacement in replacers
					str := StrReplace(str, chars, replacement)

		return str
		}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Funktionen für die PDF Kategorisierung von Addendum-Gui.ahk
	;----------------------------------------------------------------------------------------------------------------------------------------------
		ContainsName(str)                         	{       	;-- enthält Personennamen?
			rxpos := RegExMatch(Str, this.rx.Names1)
			return (rxpos ? rxpos : RegExMatch(Str, this.rx.Names2))
		}

		isFullNamed(str)                            	{        	;-- PDF-Bezeichnung enthält Nachname,Vorname, Dokumentbezeichung und Datum

			If RegExMatch(str, this.rx.CaseName, F) {
				Part1	:= RegExReplace(FN      	, "[\s,]+$")
				Part2	:= RegExReplace(FI       	, "[\s,]+$")
				Part3	:= RegExReplace(FD1 FD2, "^[\s,(v\.)]+")
				If (Part1 && Part2 && Part3)
					return true
			}

		return false
		}

		isNamed(str)                                 	{        	;-- Dateiname enthält Nachname,Vorname und eine Datumsangabe

			If RegExMatch(str, this.rx.CaseName, F) {
				Part1	:= RegExReplace(FN      	, "[\s,]+$")
				Part2	:= RegExReplace(FI       	, "[\s,]+$")
				Part3	:= RegExReplace(FD1 FD2, "^[\s,v\.]+")
				If (Part1 && Part3)
					return true
			}

		return false
		}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; class string extends
	;----------------------------------------------------------------------------------------------------------------------------------------------
		class get extends xstring                   	{        	;-- alle Stringextraktionen unter einem Namen

			DocDate(str)                                           	{       	;-- Dokumentdatum aus dem Dateinamen erhalten
				If RegExMatch(str, this.rx.LastDate, LastDoc)
					return FormatDate(LastDocDate, "DMY", "dd.MM.yyyy")
			}

			Names(str)                                             	{         	;-- Personennamen zurückgeben
				If RegExMatch(Str, this.rx.Extract, Pat)
					return {"Nn":PatNN, "Vn":PatVN, "DokTitel":PatDokTitel}
			}

		}

		class Replace extends xstring             	{           ;-- alle Ersetzungsfunktionen unter einem Namen

			Names(str, rplStr:="")                             	{         	;-- Personennamen ersetzen
				return RegExReplace(Str, this.rx.Names2, rplStr)
			}

			FileExt(str, ext:="", rplStr:="")                  	{        	;-- Dateiendungen ersetzen
				ext := !ext ? "[a-z]+" : ext
				return RegExReplace(Str, "i)\." ext "$" , rplStr)
			}

		}

}

GetAlbisPath()                                                        	{                           	;-- liest das Albisinstallationsverzeichnis aus der Registry

	SetRegView, % (A_PtrSize=8 ? 64 : 32)
	RegRead, PathAlbis, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-MainPath
	If (StrLen(PathAlbis) = 0)
		throw " Parameter 'basedir' enthält keine Pfadangabe"

return PathAlbis
}

;}


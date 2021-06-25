;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                                  	  (PSEUDO-) DATENBANKFUNKTION für die
;                                        ADDENDUM INTERNE PATIENTENNAMEN-DATENBANK UND BEFUND-DATENBANK
;
;                  			 + FUZZY STRING SEARCH UND NORMALE SUCHFUNKTIONEN FÜR OBJEKTBASIERTEN DATENBANK
;                                                                                  	------------------------
;                                                	FÜR DAS AIS-ADDON: "ADDENDUM FÜR ALBIS ON WINDOWS"
;                                                                                  	------------------------
;    		BY IXIKO STARTED IN SEPTEMBER 2017 - LAST CHANGE 05.06.2021 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                           	|                                          	|                                          	|                                          	|
; ▹PatDb                                 	PatDir                                    	class admDB

; ▹ReadGVUListe                   	IstChronischKrank           	    	InChronicList                  	    	IstGeriatrischerPatient

; ▹class AlbisDb						DBASEStructs                       		GetDBFData
; 	 Leistungskomplexe             	ReadPatientDBF			         		ReadDBASEIndex                     	Convert_IfapDB_Wirkstoffe

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

	; speichert von Daten
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
class AlbisDB	{ 		                             	 					         		 			;-- erweitert Addendum_DBASE um zusätzliche Funktionen

		; diese Klasse benötigt Addendum_DBASE.ahk, Addendum_Datum.ahk, Addendum_Internal, class_JSON.ahk
		; ausserdem wird PatDB als globales Objekt benötigt. Dieses muss Daten aus PATIENT.dbf enthalten
		; die Funktion ReadPatientDBF() liest die Daten in der benötigten Form ein

		/*  Notizen zu dBASE Feldern

				BEFAUSWA 	KUERZEL: 	wird hier nur als Zahl angegeben - möglichweise bekommt man den Klartext aus einer anderen Datenbank

				BEFDATA		TEXTDB:	welche Daten sind hier hinterlegt
									TYP:			steht wofür?
									DATA:		ist eine Zahl - hat welche Bedeutung?


		 */


	;――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Initialisieren
	;――――――――――――――――――――――――――――――――――――――――――――――――――――
		__New(DBasePath:="", debug:=0) {

					this.debug 		:= debug

			; Albisdatenbankpfad oder ein alternativer Pfad für Testversuche
				If !DBasePath
					this.DBPath      	:= GetAlbisPath() "\db"
				else
					this.DBPath      	:= DBasePath

			; Addendum Verzeichnis, Datenbankverzeichnis
				If !Addendum.Dir {
					RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)
					this.admDir := AddendumDir
				}
				else
					this.admDir :=  Addendum.Dir

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

			; EBM Regeln			                      	; Vorsorgeuntersuchungen oder -beratungen
				this.EBM.Filter        	:= {"Vorsorgen"	: 	"(01731|01732|01740M*|01745[HKM]*|01746M*|01747)"
										                                	; Chroniker Ziffern
								            		,	"Chroniker"	: 	"(03220|03221)"
							                                 				; geriatrisches Basisassement
								            		,	"Geriatrisch"	:	"(03360|03360)"
						                                 					; Quartalspauschalen
								            		,	"Pauschalen"	:	"(0300[0-6])|03040)"}

				this.EBM.Vorsorgen 	:= {"EBMFilter"	: "rx:(01731|01732|01740M*|01745[HKM]*|01746M*|01747|03220|03221|03360|03362|0300[0-6]|03040)"

												; - - - - EBM FILTERREGELN FÜR VORSORGE UNTERSUCHUNGEN ODER BERATUNGEN - - - -
												,	   "VSFilter"	: {1:{"sex":"M|W"	, "Min":"18"	, "Max":"34"  	, "repeat":"0"  	, "exam":"GVU"	}
																		,	2:{"sex":"M|W"	, "Min":"35"	, "Max":"999"	, "repeat":"36"	, "exam":"GVU"	}
																		,	3:{"sex":"M|W"	, "Min":"35"	, "Max":"999"	, "repeat":"36"	, "exam":"HKS"	}
																		,	4:{"sex":"M"   	, "Min":"55"	, "Max":"999"	, "repeat":"36"   	, "exam":"KVU"	}
																		,	5:{"sex":"M"   	, "Min":"65"	, "Max":"999"	, "repeat":"0"   	, "exam":"Aorta"	}
																		,	6:{"sex":"M|W" 	, "Min":"55"	, "Max":"999"	, "repeat":"0"		, "exam":"Colo"	}}

						                    	,		"KVU"       	: {"GO"	: "01731"          	, "Geschlecht":"M"  	, "W" : {"Alter":"55", "Abstand":"36"}}
						                    	,		"GVU"   	: {"GO"	: "01732"          	, "Geschlecht":"M|W"	, "W" : {"Alter":"35", "Abstand":"36"}
																																				, "E" 	: {"Alter":"18-34", "Abstand":"0"}}
						                    	,		"Colo"    	: {"GO": "01740"            	, "Geschlecht":"M|W"	, "W" : {"Alter":"55", "Abstand":"0"}}
						                    	,		"HKS"    	: {"GO": "(01745|01746)"	, "Geschlecht":"M|W"	, "W" : {"Alter":"35", "Abstand":"36"}}
						                    	,		"Aorta"    	: {"GO": "01747"            	, "Geschlecht":"M"  	, "W" : {"Alter":"65", "Abstand":"0"}}
						                    	,		"Chr1"    	: {"GO": "03320"            	, "Geschlecht":"M|W"	, "W" : {"Alter":"0", "Abstand":"3"}}
						                    	,		"Chr2"    	: {"GO": "03321"            	, "Geschlecht":"M|W"	, "W" : {"Alter":"0", "Abstand":"3"}}
						                    	,	  	"#01731"	: "KVU"
						                    	,	  	"#01732"	: "GVU"
						                    	, 	 	"#01740"	: "Colo"
						                    	, 	 	"#01745"	: "HKS"
						                    	, 	 	"#01746"	: "HKS"
						                    	, 	 	"#01747"	: "Aorta"}

				this.EBM.Chroniker 	:= {	"CHR1"    	: {"GO": "03320", "Geschlecht":"M|W", "B" : {"Alter":"0", "Abstand":"3", "ICDListe":""}}
						                    	,	 	"CHR2"    	: {"GO": "03321", "Geschlecht":"M|W", "B" : {"Alter":"0", "Abstand":"3", "GO":"CHR1", "ICDListe":""}}
												,		"#03320"	: "CHR1"
												,		"#03321"	: "CHR2"}

				this.EBM.Geriatrisch	:= {"EBMFilter"	: "rx:(0336[02])"}

				this.EBM.Pauschalen	:= "(0300[0-6]|03040)"

		}


	;――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Basisfunktionen
	;――――――――――――――――――――――――――――――――――――――――――――――――――――
		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		DBASEStructs(DBASEName:="", DBStructsPath:="", debug=false)		{ 	;-- analysiert alle oder eine DBase Datei/en

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

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		GetDBFData(dbName, P="", O="", S=0, dg=0, dgOpt="")            	{ 	;-- holt Daten aus einer beliebigen Datenbank

			; P     	- pattern as object
			; O   	- Feldnamen welche man zurück erhalten möchte
			; S     	- use seekpos = Start an dieser Dateiposition

			; entfernt angehängte Dateierweiterungen u.a.
				If !(dbName := this.DBaseFileExist(dbName))
					return 0

			; eine Callback Funktion für die Anzeige des Fortschritts kann eingerichtet werden
				RegExMatch(dgOpt " ", "callback\s*\=\s*(?<Func>.*?)([^\w]|\s|$)", callback)
				RegExMatch(dgOpt " ", "gui\s*\=\s*(?<Gui>.*?)([^\w]|\s|$)", debug)

			; Startposition ausrechnen
				If RegExMatch(S, "seekpos\s*\=\s*(?<pos>\d+)", seek)
					startrec	:= Floor(seekpos/dbf.lendataset) - 1
				else If RegExMatch(S, "recordnr\s*\=\s*(?<nr>\d+)", record)
					startrec := recordnr
				else
					startrec := 0

			; Informationen der Datenbank bereitstellen
				dbf        	:= new DBASE(this.DBpath "\" dbName ".dbf", dg)

			; Datenbank für Lesezugriff öffnen
				res        	:= dbf.OpenDBF()

			; Datenbanksuche durchführen
				matches 	:= dbf.Search(P, startrec, callbackFunc)

			; ein paar wichtige Filepositionen sichern
				this.DBEndFilePos := dbf.filepos
				this.DBEndRecord := dbf.breakrec
				dbf.CloseDBF()
				dbf := ""

				If !IsObject(O)
					return matches

			; Datenmenge schrumpfen
				data := Array()
				For midx, m in matches {
					strObj := Object()
					For cidx, k in O
						strObj[k] := m[k]
					data.Push(strObj)
				}

			; specialist enhancement for the use with Albis databases, some values are not stored in full length in main database
				;~ If IsObject(dbf.connected) {
					;~ data := this.AppendMatches(data, dbf.connected)
				;~ }


		return data
		}


	;――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Funktionen zur Datenauswertung
	;――――――――――――――――――――――――――――――――――――――――――――――――――――
		VorsorgeDaten(ReIndex:=false, save:=true, ini:=false)                  	{ 	;-- Abrechnungsdaten der letzten Vorsorgen

			; wird benötigt um innerhalb dieses Zeitraumes das letzte Vorsorgedatum zu suchen (einmalige GVU)
				PatExtraFilePath := "C:\tmp\Pat.Extra.ini"

			; Fehlerhandling
				If !this.DBaseFileExist("BEFGONR")	{
					this.ThrowError(A_ThisFunc 	": DBASE Datei ist nicht vorhanden`n[" "\BEFGONR.dbf" "]")
					return
				}

			; gespeicherte Daten laden oder zunächst erstellen
				If !ReIndex {
					BaseQ := GetQuartalEx("heute", "QQYY")
					IniRead, lastUpdate    	, % this.PathDBaseData "\class_AlbisDB.ini", % A_ThisFunc, % "lastUpdate"
					IniRead, lastDBPos     	, % this.PathDBaseData "\class_AlbisDB.ini", % A_ThisFunc, % "lastFilePos"
					IniRead, lastDBRecord	, % this.PathDBaseData "\class_AlbisDB.ini", % A_ThisFunc, % "lastRecord"
					ReIndex := InStr(lastUpdate, "ERROR") ? true : ReIndex
					If (!ReIndex && lastUpdate = BaseQ)
						return JSONData.Load(this.PathDBaseData "\Vorsorgen.json", "", "UTF-8")
				}

			; abgerechnete Vorsorgen/Beratungen ohne Filterung zusammenstellen
				VSData	:= Object()
				For idx, m in this.GetDBFData("BEFGONR", {"GNR":this.EBM.Vorsorgen.EBMfilter}, ["PATNR","QUARTAL", "GNR", "DATUM", "ID"], 0) {

					PatID := m.PATNR
					YYQ := this.SwapQuarterS(m.QUARTAL)
					RegExMatch(m.GNR, "\d+", GNR)

					If !IsObject(VSData[PatID])
						VSData[PatID] := {"GESCHL":this.Geschlecht(PatDB[PatID].GESCHL), "GEBURT":PatDB[PatID].GEBURT, "VS":{}}

					If RegExMatch(GNR, "0322(?<N>[01])", c)                	{

						If ini {
							IniWrite, % m.QUARTAL	, % PatExtraFilePath, % PatID , % "Chroniker" cN "_letztes_Quartal"
							IniWrite, % m.Datum 	, % PatExtraFilePath, % PatID , % "Chroniker" cN "_letztes_Datum"
							IniWrite, % c              	, % PatExtraFilePath, % PatID , % "Chroniker" cN "_letzte_Ziffer"
						}

						If !IsObject(VSData[PatID].chronisch)
							VSData[PatID].chronisch := Object()

						cN += 1
						If !VSData[PatID].chronisch.HasKey(YYQ)
							VSData[PatID].chronisch[YYQ] := cN
						else
							VSData[PatID].chronisch[YYQ] += cN

					}
					else If RegExMatch(GNR, "0336(?<N>[02])", c)       	{

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
					else if RegExMatch(GNR, "(0300[0-6]|03040)", c)    	{

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

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		Chroniker(ReIndex:=false, save:=true)                                       	{	;-- fügt alle Patient mit Chronikerziffern hinzu

				If !ReIndex && !IsObject(VSData) {
					VSData := JSONData.Load(this.PathDBaseData "\Vorsorgen.json", "", "UTF-8")
						If !IsObject(VSData)
						this.ThrowError("VSData konnte nicht erstellt werden!", A_LineNumber-2)
				} else
					VSData := this.VorsorgeDaten(true, true)

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

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		GBAsessement(ReIndex:=false, save:=true)                                 	{	;-- erfasst alle eingetragenen geriatr. Basiskompl.

		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		Krankenscheine(Abrechnungsquartal, pv:=true, save:=false)       	{	;-- Patienten mit angelegten Kassenscheinen

			; VERSART: 	1=Mitglied, 2= ?, 3=Angehöriger, 4= ?, 5=Rentner

			If !this.DBaseFileExist("KSCHEIN") {
				this.ThrowError(A_ThisFunc 	": DBASE Datei ist nicht vorhanden`n[" "\KSCHEIN.dbf" "]")
				return
			}

			quartal      	:= LTrim(Abrechnungsquartal, "0")
			this.KScheine	:= Object()
			matches    	:= this.GetDBFData("KSCHEIN", {"QUARTAL":quartal}, ["PATNR","TYP","VERSART","EINLESTAG"], 0)

			For idx, data in matches {

				; wenn pv = false (Privatversicherung), wenn diese Kassenscheine nicht erfasst
					PatID := data.PATNR
					If !pv
						If (PatDB[PatID].PRIVAT <> "f")
							continue

					If !IsObject(this.KScheine[PatID]) {
						this.KScheine[PatID] := {"VERSART"    	: data.VERSART
								        				, 	 "EINLESTAG"	: data.EINLESTAG
								    	    			, 	 "KSCHEINTYP" 	: data.TYP
								    	    			, 	 "PRIVAT"       	: (PatDB[PatID].PRIVAT <> "f" ? 1 : 0)
								    		    		, 	 "GESCHL"       	: this.Geschlecht(PatDB[PatID].GESCHL)}
					}

			}

			If Save
				JSONData.Save(this.PathDBaseData "\KScheine_" quartal ".json", this.KScheine, true,, 1, "UTF-8")

		return this.KScheine
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		Abrechnungsscheine(Abrechnungsquartal, pv:=false, save:=false)	{	;-- Krankenscheine und alle Abrechnungsziffern

			; Abrechnungsquartal im Format 0121 oder 121

			; Prüfen ob DBASE Datenbank verfügbar ist
				If !this.DBaseFileExist("BEFGONR") {
					this.ThrowError(A_ThisFunc 	": DBASE Datei ist nicht vorhanden`n[" "\BEFGONR.dbf" "]")
					return
				}

			; alle angelegten Krankenscheine auslesen
				KScheine := this.Krankenscheine(Abrechnungsquartal, pv, save)

			; alle abgerechneten Ziffern zusammentragen
				quartal              	:= LTrim(Abrechnungsquartal, "0")
				this.AbrScheine 	:= Object()
				matches            	:= this.GetDBFData("BEFGONR", {"QUARTAL": "rx:.*" quartal ".*"}, ["PATNR", "ARZTID",  "QUARTAL", "GNR", "DATUM", "ID", "removed"], 0)

				SciTEOutput("abrscheine: " matches.Count())

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

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		VorsorgeFilter(VSData, Abrechnungsquartal, pv:=true, save:=true) 	{	;-- welche Vorsorgen sind aktuell abrechenbar

			; benötigt PatDB als globales Objekt!

			; Fehlerhandling                                                                                                    	;{
				If !IsObject(VSData)
					Throw A_ThisFunc ": Parameter VSData ist kein Objekt!"
				If !RegExMatch(Abrechnungsquartal, "[0]*[1-4][0-2][0-9]")
					Throw A_ThisFunc ": Parameter Abrechnungsquartal hat ein falsches Format!"
			;}

			; Patienten mit angelegten Abrechnungsscheinen im übergebenen Abrechnungsquartals ermitteln
				KScheine := this.Krankenscheine(Abrechnungsquartal, pv, save)

			; Variablen                                                                                                            	;{
				VSCan          	:= Object()
				Quartal         	:= QuartalTage({"aktuell":Abrechnungsquartal})
				lastQDay      	:= Quartal.DBaseEnd
				lastCmpQDay	:= lastQDay "000000"
				VSFilter         	:= this.EBM.Vorsorgen.VSFilter
			;}

			; die letzten 3 Quartale darstellen                                                                             	;{
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

			; nach mögliche Untersuchungen mit Hilfe des VSFilter kategorisieren
				For PatID, q in KScheine {

					If (!pv && !q.PRIVAT) {

						; Alter und Geschlecht berechnen
							age	:= this.ageYears(PatDB[PatID].GEBURT, lastQDay)
							sex	:= this.Geschlecht(PatDB[PatID].GESCHL)
							;t := " [" PatID "] " PatDB[PAtID].NAme ", " PatDB[PatID].VORNAME " "

						; EBM Regeln anwenden
							If !(q.KSCHEINTYP = 4) {

								pending := Array()
								For ruleNR, rule in VSFilter {

								  ; das Alter des Patienten und das Geschlecht passen für die Regel
									If (age >= rule.Min && age <= rule.Max) && RegExMatch(sex, "i)(" rule.sex ")"){

										; nur einmal abrechenbar, wurde bisher nicht angesetzt
											If (!rule.repeat && !VSData[PatID].VS.HasKey(rule.exam)) {
												pending.Push(rule.exam)
											}
										; alle x-Monate abrechenbar
											else if rule.repeat {

												; wurde noch nie abgerechnet
													If !VSData[PatID].VS.HasKey(rule.exam) {
														pending.Push(rule.exam)
													}
												; wurde abgerechnet - Abstand zwischen einer möglichen erneuten Abrechnung überprüfen
													else {
														NextPossDate := DateAddEx(VSData[PatID].VS[rule.exam].D, Floor(rule.repeat/12) "y")   ; rechnet die Monatsangaben in Jahre um und addiert diese zum letzten Untersuchungsdatum dazu
														If (lastCmpQDay >= NextPossDate)
															pending.Push(rule.exam "-" NextPossDate)
													}

											}
									}

								}

							  ; EBM Regeln für Chroniker
								notmissed := 0
								For cQuarter, cCount in VSData[PatID].chronisch
									If RegExMatch(cQuarter, lQuarters)
										notmissed ++
								If (notmissed = 3)
									cronic := 2 - VSData[PatID].chronisch[YYQ]

							  ; EBM Regel für geriatrische Basiskomplexe
								geriatric := 0
								If (VSData[PatID].Geriatrisch.Count() > 1  &&  !VSData[PatID].Geriatrisch.HasKey(YYQ))
									geriatric := 1

							}

						; Grundpauschalen nicht eingetragen oder mehrfach
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

							If (pending.MaxIndex() > 0 || cronic || geriatric || nobasefee) {

								VSCan[PatID]	:= {	"ALTER"         	: age
															, 	"GESCHL"     	: sex
															,	"NAME"          	: PatDB[PatID].Name ", " PatDB[PatID].VORNAME
															,	"KSCHEINTYP"  	: q.KSCHEINTYP
															,	"VORSCHLAG" 	: (pending.MaxIndex() = 0 ? "" : pending)
															,	"CHRONISCH" 	: cronic
															,	"GERIATRISCH" 	: geriatric
															,	"PAUSCHALE" 	: nobasefee}

							}

					}

				}

			; Statistik der KScheine in VorsorgeKandidaten hinterlegen
				VSCan["_Statistik"] := {"KScheine": KScheine.Count()}

			; Liste speichern
				If save
					JSONData.Save(this.PathDBaseData "\VorsorgeKandidaten-" Abrechnungsquartal ".json", VSCan, true,, 1, "UTF-8")

		return VSCan
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		Karteikartenfilter()                                                                        	{	;-- angelegte Karteikartenfilter auslesen

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


	;――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Datenobjekte leeren
	;――――――――――――――――――――――――――――――――――――――――――――――――――――
		Empty(objectname)                                                                            	{	;-- ein Objekt leeren aber nicht löschen
		return ObjRelease(this[objectname])
		}

		EmptyDBData()                                                                                  	{ 	;-- löscht alle Objekte welche ausgelesene Daten enthalten

			this.VSData   	:= ""
			this.KScheine 	:= ""
			this.AbrScheine	:= ""
			this.KKFilter   	:= ""

		}

	;――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Datenobjekte erhalten
	;――――――――――――――――――――――――――――――――――――――――――――――――――――
		dbIndex(dbName)                                                                             	{
			return this.dbfStructs[dbname].dbIndex
		}

		GetEBMRule(RuleName)                                                                     	{	;-- eine bestimmte EBM Regel
		return this.EBM[RuleName]
		}

	;――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Datenobjekte speichern (auch extern bearbeitete)
	;――――――――――――――――――――――――――――――――――――――――――――――――――――
		SaveVorsorgeKandidaten(VSCan, Abrechnungsquartal)                       	{	;-- speichern im JSON Format
			JSONData.Save(this.PathDBaseData "\VorsorgeKandidaten-" Abrechnungsquartal ".json", VSCan, true,, 1, "UTF-8")
			FileGetSize, filesize, % this.PathDBaseData "\VorsorgeKandidaten-" Abrechnungsquartal ".json"
		return fileSize
		}


	;――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Indizes erstellen/lesen
	;――――――――――――――――――――――――――――――――――――――――――――――――――――
		IndexCreate(dbName, IndexField="", IndexMode="", ReIndex=true)      	{ 	;-- Indexerstellung für Albis DBase Datenbanken

			/* BESCHREIBUNG

				Die Funktion erleichtert die Erstellung eines Datenbankindex durch hinterlegte Feldnamen, welche zur Indizierung herangezogen werden.
				Ist kein Feldname für eine Datenbank hinterlegt, kann diesen über den Parameter IndexField übergeben.
				Der übergebene Feldname wird von der Funktion vorrangig behandelt!
				Ist ein Feldname nicht vorhanden wird eine Fehlermeldung ausgeworfen.

			 */

			; Variablen                                                                    	;{
				; Standard Indexfelder
				static IndexFields := {"BEFUND" 	: {"IndexKey":"DATUM"            	, "IndexMode":"DateToQuarter"}
											, 	"BEFTEXTE"	: {"IndexKey":"DATUM"           	, "IndexMode":"DateToQuarter"}
											,	"BEFGONR"	: {"IndexKey":"QUARTAL"        	, "IndexMode":""}
											, 	"LABORANF"	: {"IndexKey":"EINDATUM"        	, "IndexMode":"DateToQuarter"}
											, 	"LABBLATT"	: {"IndexKey":"DATUM"           	, "IndexMode":"DateToQuarter"}
											, 	"LABBUCH"	: {"IndexKey":"ABNAHMDATE"	, "IndexMode":"DateToQuarter"}
											, 	"LABREAD"  	: {"IndexKey":"EINDATUM"        	, "IndexMode":"DateToQuarter"}
											, 	"LABUEBPA"	: {"IndexKey":"ABDATUM"        	, "IndexMode":"DateToQuarter"}}

				if !IndexMode
					IndexMode := IndexFields[dbName].IndexMode

				If !IndexField
					IndexField := IndexFields[dbName].IndexKey

			;}

			; Dateinamen überprüfen                                               	;{
				If !(this.dbName := this.DBaseFileExist(dbName))
					Throw A_ThisFunc	": Die Datenbank ["  dbName "] ist nicht vorhanden!"
			;}

			; Information auslesen und Indizierbarkeit prüfen             	;{
				dbf := new DBASE(this.DBPath "\" this.dbName ".dbf", this.debug)
				If !this.dbfStructsPath.HasKey(this.dbName) {
					this.dbfStructsPath(this.dbName) 		                                     	;-- legt Objekt für die Abbildung der DBase Datei an
					this.dbfStructs[this.dbname].dbfields    	:= dbf.dbfields	        	;-- Feldnameninformationen hinzufügen
					this.dbfStructs[this.dbname].maxrecords	:= dbf.records	        	;-- Anzahl der Recordsets
					this.dbfStructs[this.dbname].lastupdate	:= dbf.lastupdate        	;-- letzte Änderung
				}
				If (StrLen(IndexField) > 0) {
					If !this.FieldNameExist(IndexField)
						Throw A_ThisFunc	": Die Datenbank ["  this.dbName "] kann nicht indiziert werden!`n"
										    	.	"  Der Feldnahme [" IndexField "] ist nicht enthalten."
				}
			;}

			; Index erstellen (wird als JSON String gespeichert)          	;{
				this.dbfStructs[this.dbname].dbIndex := dbf.CreateIndex(this.PathDBaseData "\DBIndex_" this.dbName ".json", IndexField, IndexMode, ReIndex)
				dbf.CloseDBF()
				dbf := ""
			;}


		return this.dbfStructs[this.dbname].dbIndex
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		IndexRead(dbName, IndexField:="", ReIndex:=false, IndexMode:="")      	{ 	;-- Indizes laden oder bei Bedarf erstellen

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


	;――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Hilfsfunktionen
	;――――――――――――――――――――――――――――――――――――――――――――――――――――
		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		AppendMatches(matches, properties)                                            	{	;-- appends database values, if connection to another db is known

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

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		dbfStructsPath(dbName)                                                            	{	;-- legt das Objekt dbfStructsPath an

			 If !IsObject(this.dbfstructs[dbName])
				this.dbfStructs[dbName] := Object()

		return IsObject(this.dbfstructs[dbName]) ? true : false
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		CreateFilePath(path)                                                                  	{ 	;-- erstellt einen Dateipfad falls dieser noch nicht existiert

			If !InStr(FileExist(path "\"), "D") {
				FileCreateDir, % path
				return ErrorLevel ? 0 : 1
			}

		return 1
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		DBaseFileExist(dbName)                                                              	{	;-- Datenbank ist vorhanden?

			; Dateinamen überprüfen
				dbName 	:= RegExReplace(dbName, "i)\.dbf$")
				If !FileExist(this.DBPath "\" dbName ".dbf")
					Throw A_ThisFunc	": Die Datenbankdatei <" dbName ">`nbefindet sich nicht im Ordnerpfad:`n<" this.DBPath ">"

		return dbName
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		FieldNameExist(fieldLabel)                                                          	{	;-- ist der übergebene Feldname vorhanden?

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

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		SwapQuarter(QQYY)                                                                 	{ 	;-- Quartal Forrmat QQYY zu YYQQ
			QQYY := SubStr("0" QQYY, -3)
		return SubStr(QQYY, 3, 2) SubStr(QQYY, 1, 2)
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		SwapQuarterS(QYY)                                                                  	{	;-- Quartal Format QYY zu YYQ
		return SubStr(QYY, 2, 2) SubStr(QYY, 1, 1)
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		ageYears(birthday, thisday)                                                         	{	;-- Altersjahrberechnung
			;Format für beide Daten ist YYYYMMDD
			timeDiff := HowLong(birthday, thisday)
			return timeDiff.Years
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		Geschlecht(AlbisKodierung)                                                         	{	;-- Umkodieren der Geschlechtziffern in einen Buchstaben
			sex := AlbisKodierung
			return (sex = 1 ? "M": sex = 2 ? "W" : sex)
		}

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		SortMatchesByKey(matches, SortByKey:="")                                    	{	;-- Array in key:object() umwandeln

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

		; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		ThrowError(msg, errorLine:="")                                                   	{ 	;-- "wirft" Fehlermeldungen aus
			this.db.CloseDBF()
			this.db := ""
			throw A_ThisFunc "," errorLine ": " msg
		}

}

class PatDBF	{                                                                                    	;-- PATIENT.dbf Handler

	;-- gibt nur benötigte Daten der albiswin\db\PATIENT.DBF zurück
		__New(basedir="", infilter="", opt="") {

			; basedir kann man leer lassen, der Albishauptpfad wird aus der Registry gelesen
			; Rückgabeparameter ist ein Objekt mit Patienten Nr. und dazugehörigen Datenobjekten (die key's sind die Feldnamen in der DBASE Datenbank)
			; lade nur Patienten die in den letzten 10 Jahren behandelt wurden
			; minDate := 20100101

			; kein basedir - sollte ..\albiswin\db\ enthalten dann wird hier versucht den Pfad aus der Windows Registry zu bekommen
				this.Patients := Object()
				If (StrLen(basedir) = 0) || !InStr(FileExist(basedir), "D")
					this.basedir	:= GetAlbisPath() "\db"

			; Optionen
				this.debug := this.AllData := 0
				If RegExMatch(opt, "^\d$", dbg) || RegExMatch(opt, "i)debug(?<bg>\d)\s", d)
					this.debug := dbg
				If RegExMatch(opt, "i)moreData")
					this.moreData := true
				If RegExMatch(opt, "i)allData")
					this.AllData := true

			; die minimal notwendigsten Daten
				If !this.moreData
					this.infilter		:= !IsObject(infilter) 	? ["NR", "NAME", "VORNAME", "GEBURT", "TELEFON", "GESCHL", "MORTAL", "TELEFON2", "FAX", "LAST_BEH", "RKENNZ", "NAMENORM"] : infilter
				else
					this.infilter		:= !IsObject(infilter) 	? ["NR", "NAME", "VORNAME", "GEBURT", "GESCHL", "NAMENORM"
																		, 	"TELEFON", "TELEFON2", "FAX"
																		,	"HAUSARZT", "ARBEIT"
																		, 	"MORTAL", "LAST_BEH", "DEL_DATE", "RKENNZ"] : infilter

			; liest alle Patientendaten in ein temp. Objekt ein
				database 	:= new DBASE(this.basedir "\PATIENT.dbf", this.debug)
				res        	:= database.OpenDBF()
				matches	:= database.GetFields(this.infilter)
				res         	:= database.CloseDBF()

			; temp. Objekt wird nach NR (Patientennummerierung) sortiert
				For idx, m in matches {

					If !this.AllData
						If (InStr(m.RKENNZ, "*") || (m.LAST_BEH = 0) || (!m.NAME && !m.VORNAME))
							continue

					strObj := Object()
					For key, val in m
						If (key <> "NR") && (StrLen(val) > 0) {

							strObj[key] := val
							If (key = "NAME")
								strObj.DiffN 	:= string.TransformGerman(val)
							else if (key = "VORNAME")
								strObj.DiffVN:= string.TransformGerman(val)

						}

					this.Patients[m.NR] := strObj

				}

			; Mail Adressen aus PATEXTRA beziehen
				database 	:= new DBASE(this.basedir "\PATEXTRA.dbf", this.debug)
				res        	:= database.OpenDBF()
				matches	:= database.GetFields("NR", "POS", "TEXT")      ; POS 93 enthält EMail, 97 Ausnahmekennziffern
				res         	:= database.CloseDBF()

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
					SciTEOutput("PatDB: " PatID)
					matches.Push(PatID)
					return matches.1
				}
			}

			name1  	:= string.TransformGerman(name1)              ; class string
			name2  	:= string.TransformGerman(name2)

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

DBASEStructs(AlbisPath, DBStructsPath:="", debug=false) {                    	;-- analysiert alle DBF Dateien im albiswin\db Ordner

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

ReadPatientDBF(basedir="", infilter="", outfilter="", opt="", minDate=20100101) {    	;-- gibt nur benötigte Daten der albiswin\db\PATIENT.DBF zurück

	; basedir kann man leer lassen, der Albishauptpfad wird aus der Registry gelesen
	; Rückgabeparameter ist ein Objekt mit Patienten Nr. und dazugehörigen Datenobjekten (die key's sind die Feldnamen in der DBASE Datenbank)
	; lade nur Patienten die in den letzten 10 Jahren behandelt wurden
	; minDate := 20100101

	; kein basedir - sollte ..\albiswin\db\ enthalten dann wird hier versucht den Pfad aus der Windows Registry zu bekommen
		PatDBF := Object()
		If (StrLen(basedir) = 0) || !InStr(FileExist(basedir), "D")
			basedir	:= GetAlbisPath() "\db"

	; für eine Abrechnungsüberprüfungen die geschätzt minimal notwendigste Datenmenge
		If !IsObject(infilter)
			infilter 	:= ["NR", "NAME", "VORNAME", "GEBURT", "GESCHL", "MORTAL", "LAST_BEH", "RKENNZ"]
		If !IsObject(outfilter)
			outfilter	:= ["NR", "NAME", "VORNAME", "GEBURT", "GESCHL", "MORTAL", "LAST_BEH", "RKENNZ"]

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
				If (InStr(m.RKENNZ, "*") || (m.LAST_BEH < minDatum) || (m.LAST_BEH = 0) || (!m.NAME && !m.VORNAME))
					continue

			strObj	:= Object()
			For key, val in m
				If (key <> "NR") && (StrLen(val) > 0)
					strObj[key] := val

			PatDBF[m.NR] := strObj

		}

	; Mail Adressen aus PATEXTRA beziehen
		database 	:= new DBASE(basedir "\PATEXTRA.dbf", this.debug)
		res        	:= database.OpenDBF()
		matches	:= database.GetFields("NR", "POS", "TEXT")      ; POS 93 enthält EMail, 97 Ausnahmekennziffern
		res         	:= database.CloseDBF()

		For idx, m in matches
			If (m.POS = "93") && PatDBF.HasKey(m.NR)
				If RegExMatch(m.Text, "i)^.*@.*\.[a-z]+$")
					PatDBF[m.NR].EMAIL := m.Text

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
			dbIndex	:= AlbisDB.IndexCreate(DBASEName,, ReIndex)
			AlbisDB 	:= ""

		}

return dbIndex
}

Convert_IfapDB_Wirkstoffe(filepath, savePath:="") {                               	;-- konvertiert die Ifap Wirkstoffdatenbank in eine reine Textdatei

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

		Karteikartentext(str)                        	{        	;-- entfernt Namen, Dateierweiterung und überflüssige Leerzeichen
			str := this.Replace.FileExt(str)                	; entfernen der Dateiendung
			str := this.Replace.Names(str)            	; Namen von Patienten entfernen
			str := RegExReplace(str, "^[\s,]*")        	; einzeln stehende Komma entfernen
		return RegExReplace(str, "\s{2,}", " ")       	; Leerzeichenfolgen kürzen
		}

		TransformGerman(str)                     	{        	;-- für Stringmatching Algorithmen

			str   	:= StrReplace(str, "ä", "ae")
			str   	:= StrReplace(str, "ö", "oe")
			str   	:= StrReplace(str, "ü", "ue")
			str   	:= StrReplace(str, "ß", "ss")

		return str
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

		isNamed(str)                                 	{        	;-- Dateiname enthält Nachname,Vorname und eine Datumsangabe

			;str := RegExReplace(str, this.rx.fileExt) ; neuer RegExString 'CaseName' kann mit Dateiendungen umgehen
			If RegExMatch(str, this.rx.CaseName, F) {
				Part1	:= RegExReplace(FN      	, "[\s,]+$")
				Part2	:= RegExReplace(FI       	, "[\s,]+$")
				Part3	:= RegExReplace(FD1 FD2, "^[\s,v\.]+")
				If (StrLen(Part1) > 0) && (StrLen(Part3) > 0)
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



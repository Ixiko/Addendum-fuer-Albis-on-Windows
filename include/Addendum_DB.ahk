;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                                  	  (PSEUDO-) DATENBANKFUNKTION für die
;                                        ADDENDUM INTERNE PATIENTENNAMEN-DATENBANK UND BEFUND-DATENBANK
;
;                  			 + FUZZY STRING SEARCH UND NORMALE SUCHFUNKTIONEN FÜR OBJEKTBASIERTEN DATENBANK
;                                                                                  	------------------------
;                                                	FÜR DAS AIS-ADDON: "ADDENDUM FÜR ALBIS ON WINDOWS"
;                                                                                  	------------------------
;    		BY IXIKO STARTED IN SEPTEMBER 2017 - LAST CHANGE 24.06.2020 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;----------------------------------------------------------------------------------------------------------------------------------------------------------------------
;###############################################################################
;----------------------------------------------------------------------------------------------------------------------------------------------------------------------

;----------------------------------------------------------------------------------------------------------------------------------------------------------------------
; PATIENTEN DATENBANK INKL. SUCHFUNKTIONEN
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------;{

ReadPatientDatabase(AddendumDBPath) {														;-- liest die .csv Datei Patienten.txt als Object() ein

	PatDB	:= Object()

	If FileExist(AddendumDBPath "\Patienten.txt")
	{
			; Einlesen der Datenbank als Textliste, Sortieren aufsteigend nach PatID, Aussortieren doppelter Einträge (später neue Einträge unter den Skripten kommunizieren?)
				FileRead    	,    PatDB_temp, % AddendumDBPath "\Patienten.txt"
				Sort          	,    PatDB_temp, N U
				FileDelete  	, % AddendumDBPath "\Patienten.txt"
				FileAppend	, % PatDB_temp, % AddendumDBPath "\Patienten.txt", UTF-8

			; Einlesen in ein Objekt
				;~ 1.PatID = key; 2.Nachname (Nn); 3.Vorname (Vn); 4.Geschlecht (Gt); 5.Geburtsdatum (Gd); 6.Krankenkasse (Kk); 7.letzteGVU (letzteGVU)
				Loop, Parse, PatDB_temp, `n, `r
				{
							If (StrLen(A_LoopField) = 0)
									continue
							Str := StrSplit(A_LoopField, ";", A_Space)
							PatID := Str[1]
							PatDB[PatID] := {"Nn": Str[2], "Vn": Str[3], "Gt": Str[4], "Gd": Str[5], "Kk": Str[6], "letzteGVU": Str[7]}
				}

				PatDB["MaxPat"] := maxPat := PatDB.Count()
	}

return PatDB
}

PatInDB(PatDb, Name) {						     													;-- für ScanPool - Fuzzy-Patientensuche

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Name[3] - ist Nachname, Name[4] sollte der Vorname sein
	;----------------------------------------------------------------------------------------------------------------------------------------------
		If !IsObject(PatDB)
					Exceptionhelper(A_ScriptName, "PatInDb(PatDB,cmd:="") {", "Die Funktion hat einen Fehler ausgelöst,`nweil kein Objekt übergeben wurde.", A_LineNumber)

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; versucht auch durch vertauschen von Vor- und Nachnamen den Patienten in der Datenbank zu finden
	;----------------------------------------------------------------------------------------------------------------------------------------------
		PatID:= FindPatData(PatDb, "PatID", "Nn", Name[3] , "Vn", Name[4] )
		If PatID
			return PatID
		PatID:= FindPatData(PatDb, "PatID", "Nn", Name[4] , "Vn", Name[3] )
		If PatID
			return PatID

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Fuzzy Teil
	;----------------------------------------------------------------------------------------------------------------------------------------------
		For key, DbPat in PatDb
		{
				If StrSplit(DbPat["Vn"], " ").MaxIndex() > 1
				{
						Loop, % StrSplit(DbPat["Vn"], " ").MaxIndex()		;geht jeden Vornamen durch und vergleicht ihn mit dem gesuchten Namen
						{
								nr:= A_Index
								Loop, % StrSplit(Name[4], " ").MaxIndex()
								{
										If ( StrDiff(DbPat["Nn"] StrSplit(DbPat["Vn"], " ")[nr], Name[3] StrSplit(Name[4], " ")[A_Index]) <= 0.2 ) || ( StrDiff(DbPat["Nn"] StrSplit(DbPat["Vn"], " ")[nr], Name[4] StrSplit(Name[3], " ")[A_Index]) <= 0.2 )
												suggestion.= key . ": " . DbPat["Nn"] . ", " . DbPat["Vn"] . "|"
								}
						}
				}
				else
				{
						If ( StrDiff(DbPat["Nn"] . DbPat["Vn"], Name[3] . Name[4]) <= 0.2 ) || ( StrDiff(DbPat["Nn"] . DbPat["Vn"], Name[4] . Name[3]) <= 0.2 )
								suggestion.= key . ": " . DbPat["Nn"] . ", " . DbPat["Vn"] . "|"
				}
		}

return RTrim(suggestion, "|")
}

PatInDB_SiftNgram(Name) {

		static Haystack, Haystack_Init:=0
		found:=Object(), Needle:= [], idx:= 0

		If !IsObject(oPat)
			MsgBox, Kein Objekt!

		If !Haystack_Init
		{
				Haystack_Init:= 1
				For PatNr, PatData in oPat
					Haystack.= PatData["Nn"] " " PatData["Vn"] "`n"
		}


		Needle[1]:= Name[3] " " Name[4]
		Needle[2]:= Name[4] " " Name[3]

	; from https://www.autohotkey.com/boards/viewtopic.php?f=76&t=28796 - Getting closest string
		Loop, 2
		{
				for key, element in Sift_Ngram(Haystack, Needle[A_Index], 0, , 2, "S")
				If (element.delta > 0.90) 	{
						idx ++
						element:= StrSplit(element.data, " ")
						found[idx] := GetFromPatDb("PatID|Nn|Vn|Gd", 1, element[1], element[2])
				}
		}

return (idx=0) ? 0 : found
}

GetFromPatDb(getString:="PatID|Gd", OnlyFirstMatch:=0, Nn:="", Vn:="", Gt:="", Gd:="", Kk:="") {

	returnObj	:= Object()
	getter    	:= StrSplit(getString, "|")

	For PatID, PatData in oPat
			If InStr(PatData["Nn"], Nn) && InStr(PatData["Vn"], Vn) && InStr(PatData["Gt"], Gt) && InStr(PatData["Gd"], Gd) && InStr(PatData["Kk"], Kk)
			{
					Loop, % getter.MaxIndex()
							If InStr(getter[A_Index], "PatID")
								returnObj[getter[A_Index]] := PatID
							else
								returnObj[getter[A_Index]] := PatData[getter[A_Index]]

					If OnlyFirstMatch
							break
			}

return returnObj
}

PatDb(Pat, cmd:="") {																			        ;-- überprüft die Addendum Patientendatenbank und führt auch das alternative Tagesprotokoll

			If !IsObject(Pat)
				Exceptionhelper(A_ScriptName, "PatDb(Pat,cmd:="") {", "Die Funktion hat einen Fehler ausgelöst,`nweil kein Objekt übergeben wurde.", A_LineNumber)


		;fehlerhafte Funktionsaufrufe finden noch immer statt, warum?
			PatID := Pat.ID
			If (!RegExMatch(PatID, "^\d+$") || StrLen(PatID) = 0) 									;abbrechen falls keine PatID ermittelt werden kann
					return

		;Nn - Nachname, Vn - Vorname, Gt - Geschlecht, Gd - Geburtsdatum, Kk - Krankenkasse
			If !oPat.Haskey(PatID) {

					oPat[PatID] := {"Nn": Pat.Nn, "Vn": Pat.Vn, "Gt": Pat.Gt, "Gd": Pat.Gd, "Kk": Pat.Kk}

					FileAppend, % PatID ";" Pat.Nn ";" Pat.Vn ";" Pat.Gt ";" Pat.Gd ";" Pat.Kk ";`n", % Addendum.DBPath "\Patienten.txt", UTF-8
					TrayTip, Addendum, % "neue PatID (Zähler: " oPat.Count() ") für die Addendumdatenbank:`n(" PatID ") " Pat.Nn "," Pat.Vn, 1

				; Praxomat zeigt den maximalen Index der Patienten in der 'Addendum-Patienten' Datenbank an
					If ( hPraxomatGui:= WinExist("PraxomatGui ahk_class AuthotkeyGui") ) {
							Send_WM_COPYDATA("PatDBCount|" oPat.Count() , hPraxomatGui)
							Sleep, 500
					}
			}

		;ermittelt ob diese Patientenakte heute schon aufgerufen wurden
			For Index, Value in TProtokoll
				 If InStr(Value, PatID)
								return

		;PatID wird dem Tagesprotokoll hinzugefügt und gespeichert
			TProtokoll.Push(PatID)
			IniAppend("<" PatID "(" A_Hour ":" A_Min ")>", Addendum.TPFullPath, SectionDate, compname)
			TrayTip, Tagesprotokoll, % "neue ID für das Tagesprotokoll: (" PatID ") " oPat.Nn ", " oPat.Vn ", " oPat.Gd "`nZähler: " TProtokoll.Count(), 1
			If ( hPraxomatGui := WinExist("PraxomatGui") )
					Send_WM_COPYDATA("TPCount|" TProtokoll.Count() , hPraxomatGui)

return
}

PatDBSave(AddendumDBPath) {                                                                   	;-- zum Sichern der Patientendatenbank

	dbFile:= FileOpen(AddendumDbPath "\Patienten.txt", "w", "UTF-8")
	For PatID, obj in oPat
	{
			line           	:=  PatID ";" obj.Nn ";" obj.Vn ";" obj.Gt ";" obj.Gd ";" obj.Kk ";" obj.letzteGVU
			dataToWrite	.=  RTrim(line, ";") "`n"
	}
	dbFile.Write(dataToWrite)
	dbFile.Close()

return
}

FindPatData(tPat, returnValue, pA1, pA2, pB1:="", pB2:="") {

	;zB. FindPatData(oPat, "PatID", "Nn", "Mustermann", "Vn", "Max" ) gibt die PatientenID zurück wenn ein Datensatz vorhanden ist
	;zB. FindPatData(oPat, "Birth"	, "Nn", "Mustermann", "Vn", "Max" ) gibt das Geburtsdatum zurück wenn ein Datensatz vorhanden ist
	;;Nn - Nachname, Vn- Vorname, Gt - Geschlecht, Gd - Geburtsdatum, Kk - Krankenkasse

	For FPD_PatID, PatData in tPat
	{
			If (PatData[(pA1)] = pA2)
			{
						If (pB1 = "Vn") && (StrSplit(PatData[pB2], " ").MaxIndex() > 1)
						{
								For k, v in StrSplit(PatData[pB2], " ")
									If v = pB2
										return ( InStr(returnValue, "PatID") ? k : PatData[(returnValue)] )
						}
						else if (PatData[(pB1)] = pB2)
						{
								return ( InStr(returnValue, "PatID") ? k : PatData[(returnValue)] )
								;~ If InStr(returnValue, "PatID")
										;~ return FPD_PatID
								;~ else
										;~ return PatData[(returnValue)]
						}
			}
		}

return 0
}

;}

;----------------------------------------------------------------------------------------------------------------------------------------------------------------------
; PDF DATENBANK / SCANPOOL
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------;{

ScanPoolArray(cmd, param:="", opt:="") {                    		;-- verarbeitet den files-Array der die Datei-Informationen des Befundordners bereit hält

/*							BESCHREIBUNG

		ein Array mit dem Namen 'ScanPool' muss superglobal gemacht werden am Anfang des Skriptes, wichtig egal was man mit dem Array dann macht: Er darf niemals in dieser Funktion gelöscht werden oder neu
		initialisiert werden.  Beispiel: ScanPool:="" , entfernt leert den Speicher den der Array besetzt, einen Array mit selbigen Namen zu initalisieren ' ScanPool:=[] ' ergibt nicht den selben Array.
		Ergebnis ist das der Array Name zwar global angelegt wurde, jetzt aber ausserhalb dieser Funktion leer ist.

		Beschreibung: param
		was param beinhalten kann, ist vom übergebenen Befehl (cmd) abhängig, z.B. ScanPoolArray("GetCell", "5|3") ; ermittle den Inhalt der 3.Spalte der 5.Zeile, weiteres siehe folgende Zeilen:
		cmd:    "Delete"			- löschen einer Datei innerhalb des ScanPool-Array, param - Name der zu entfernenden Datei
					"Sort"   			- sortiert die Dateien im Array und somit auch für die Anzeige im Listviewfenster
					"Load"			- lädt aus einer Datei den zuvor indizierten Ordner samt entsprechender Daten
					"Save"			- speichert die Daten auf Festplatte in eine Textdatei (mit '|' getrennte Speicherung der einzelnen Felder Dateiname,Größe,Seiten,Signiert ja=1;nein=Feld bleibt leer)
					"Rename" 		- Umbenennen einer Datei innerhalb des ScanPool-Array
					"ValidKeys"	- zählt die vorhandenen Datensätze, da auch manchmal leere Datensätze gespeichert wurden, werden nur nicht leere gezählt
					"Find"	    	- sucht nach einem Dateinamen und gibt den Index zurück
					"CountPages"- ermittelt die Gesamtzahl aller Seiten der Pdf-Dateien im BefundOrdner
					"Reports"     	- fuzzy Personennamen Suche, param - Nachname Vorname (des Patienten)
					"Signed"     	- erstellt einen Array in dem alle im Befundordner signierten Dateien aufgelistet sind (kontrolliert nicht ob neue Dateien hinzugefügt wurden)
					"NotSigned"	- wie "Signed" nur alle bekannt unsignierten um diese auf eine vorhandene Signatur zu prüfen
*/

	static Loaded := false
	static PdfIndexFile
	static newPDF

	columns:=[], res:=0, allfiles:=""
	FileEncoding, UTF-8

	;diese Zeilen sichern ab das zuerst der ScanPool-Array erstellt wird bevor die anderen Befehle aufgerufen werden können
	If !Loaded && StrLen(param) = 0 	{
			MsgBox, 0, Funktion ScanPoolArray, Dieser Funktion muss als erstes per 'Load' command`nder Pfad zur PdfIndex.txt Datei übergeben werden.
			ExitApp
			;Loaded:= ScanPoolArray("Load")
	}
  ;--------------------------------------------------------------------- Befehle -----------------------------------------------------------------------------------

	If Instr(cmd, "Delete") || Instr(cmd, "Remove")	{  	;param: gesuchter Dateiname			           			, Rückgabe: Wert       	- Seitenanzahl der entfernten Pdf-Datei

			param:= RegExReplace(param, "\.pdf$", "") ".pdf" ;muss man nicht immer dran denken die Dateiendung zu übergeben
			For key, val in ScanPool
				If Instr(val, param)
				{
					delpages := StrSplitEx(val, 3)
					ScanPool.Delete(key)
					break
				}

			FileCount:= CountValidKeys(ScanPool)						;ist sozusagen dann der "NEUE" MaxIndex()
			return delpages
	}
	else If Instr(cmd, "Load")         	{                         	;param: Pfad zur PDFIndex.txt Datei                      	, Rückgabe: Wert       	- Gesamtzahl der Befunde in der pdfIndex.txt Datei

			if FileExist(param)
			{
					PdfIndexFile := param
					FileRead, allfiles, % PdfIndexFile
					Sort, allfiles
					Loop, Parse, allfiles, `n, `r
							ScanPool.Push(A_LoopField)

					Loaded 	:= true											        	;zur Überprüfung - Load muss vor allen anderen Befehlen als erstes stattgefunden haben
					VarSetCapacity(allfiles, 0)
					MaxFiles:= ScanPool.MaxIndex()
					return MaxFiles
			}
			else
					return 0
	}
	else If Instr(cmd, "Rename")    	{                         	;param: Original-Dateiname, opt: neuer Name     	, Rückgabe: Wert       	- ist der Index im ScanPool-Array

			for key, val in ScanPool
			{
					If Instr(val, param)
					{
							columns:= StrSplit(ScanPool[key], "|")
							ScanPool[key]:= opt . "|" . columns[2] . "|" . columns[3]
							return key
					}
			}
			return 0
	}
	else If Instr(cmd, "Save")         	{                         	;param: und opt - unbenutzt					            	, Rückgabe: ErrorLevel	- erfolgreich = 1, Speicherung nicht möglich = 0

			File:= FileOpen(PdfIndexFile, "w", "UTF-8")

			For key, val in ScanPool
				If val != ""
					allfiles.= val "`n"

			allfiles:= RTrim(allfiles, "`n")
			File.Write(allfiles)
			File.Close()
			If !ErrorLevel
						return 1
			else
						return 0
	}
	else if Instr(cmd, "Sort")           	{                         	;param: und opt - unbenutzt			            			, Rückgabe: ohne      	- der ScanPool-Array wird sortiert

			For key, val in ScanPool
					allfiles.= val . "`n"
			Sort, allfiles
			allfiles:= RTrim(allfiles, "`n")
			ScanPool:= StrSplit(allfiles, "`n")
	}
	else if Instr(cmd, "ValidKeys")   	{                         	;param: und opt - unbenutzt				           			, Rückgabe: Wert       	- Anzahl der im Array gespeicherten Pdf-Dateien
		return CountValidKeys(ScanPool)
	}
	else If InStr(cmd, "Find")           	{                         	;param: gesuchter Dateiname		            			, Rückgabe: Wert       	- ist der Indexwert oder KeyIndex im ScanPool-Array

			for key, val in ScanPool
				If Instr(val, param)
			    		return key

			return 0
	}
	else If Instr(cmd, "CountPages")  {                        	;param: und opt - unbenutzt			           				, Rückgabe: Wert       	- Gesamtzahl aller Seiten in den Pdf-Dateien des Befundordners

			tpgs:= 0
			for key, val in ScanPool
					tpgs += (StrSplitEx(val, 3) = "" ) ? 0 : StrSplitEx(val, 3)

			return tpgs
	}
	else If InStr(cmd, "Reports")       	{                      	;param: Patientenname (Nachname, Vorname)     	, Rückgabe: Array      	- Pdf Befunde mit passendem Patientennamen

			Reports := []

			RegExMatch(StrReplace(param, "-"), "(?P<Nachname>[\w\p{L}]+)[\,\s]+(?P<Vorname>[\w\p{L}]+)", Such)
			SuchName := SuchNachname SuchVorname

			For key, val in ScanPool					;wenn keine PatID vorhanden ist, dann ist die if-Abfrage immer gültig (alle Dateien werden angezeigt)
			{
						if (StrLen(val) = 0)
							continue

						filename := StrReplace(StrSplit(val, "|").1, ".pdf")
						RegExMatch(StrReplace(filename, "-"), "(?P<Nachname>[\w\p{L}]+)[\,\s]+(?P<Vorname>[\w\p{L}]+)", pdf)

						a := StrDiff(SuchName, pdfNachname pdfVorname)
						b := StrDiff(SuchName, pdfVorname pdfNachname)

						If ((a < 0.12) || (b< 0.12))
								Reports.Push(filename)
			}

			return Reports
	}
	else If InStr(cmd, "Refresh")       	{                       	;param: unbenutzt                                                	, Rückgabe: Wert        	- Anzahl neuer Funde

			If InStr(param, "Tip") {
				PraxTT("ermittle neue Befunde im ScanPool", "0 3")
				Sleep, 500
			}

			If InStr(A_ScriptName, "ScanPool")
				newPDF:= RefreshPdfIndex(BefundOrdner)
			else
				newPDF:= RefreshPdfIndex(Addendum.BefundOrdner)

			If InStr(param, "Tip")
				PraxTT("", "off")

			return newPDF
	}
	else If InStr(cmd, "Signed")      	{                      	;param: und opt - unbenutzt						           	, Rückgabe: Array       	- signierte Befunde

			Signed := [], BOidx := 0
			for key, val in ScanPool
					If StrSplitEx(val, 4) = 1
					{
							BOidx ++
							Signed[BOidx]:= StrSplitEx(val, 1)
					}
			return Signed
	}
	else If InStr(cmd, "NotSigned") 	{                       	;param: Array (signierter Befunde "Signed")           	, Rückgabe: Array       	- unsignierte Befunde

			if !IsObject(param)
					return ScanPool

			NotSigned := [], BOidx := 0
			For key, val in ScanPool
					allfiles.= val . "`n"
			For key, val in param
					allfiles.= val . "`n"
			allfiles:= RTrim(allfiles, "`n")

			for key, val in ScanPool
				If InStr(val, param)
					If StrSplitEx(val, 4) = 1
					{
							BOidx ++
							SignedArr[BOidx]:= StrSplitEx(val, 1)
					}

			return SignedArr
	}

return "EndOfFunction"
}

CountValidKeys(arr) {                                                     	;-- zählt die gültigen Einträge im Array

	counter:=0, notValid:=""
	For key, val in arr
		If !(val = "")
			counter++

return counter
}

ReadDir(dir, ext) {                                                           	;-- liest ein Verzeichnis ein, ext=Dateiendung
	Loop, Files, % dir "\*." ext
		tlist .= A_LoopFileName "`n"
return tlist
}

ReadPdfIndex(PdfIndexFile) {	                                        	;-- erstellt das ScanPool Object

		;Teile der Variablen sind globale Variablen

		PageSum	:= 0
		FileCount	:= 0
		allfiles   	:= ""
		tidx       	:= 0

		If ( FileCount	:= ScanPoolArray("Load", PdfIndexFile) )					;erstellt den files Array aus der pdfIndex.txt Datei
				PageSum:= ScanPoolArray("CountPages")

		PdfDirList:= ReadDir(BefundOrdner, "pdf")
		RegExReplace(PdfDirList, "m)\n", "", filesInDir)

	;nicht mehr vorhandene Dateien aus dem Index nehmen
		For key, val in ScanPool
				If !InStr(PdfDirList, StrSplit(val, "|").1)
						ScanPool.Delete(key)

	;nach noch nicht aufgenommenen Dateien suchen
		Loop, Parse, PdfDirList, `n, `r
		{
				If !ScanPoolArray("Find", A_LoopField)
				{
						FileGetSize, FSize, % BefundOrdner "\" A_LoopField, K
						ScanPool.Push(A_LoopField . "|" . FSize . "|" . pages)
						continue
				}

		}

	;Sortieren der eingelesenen und aktualisierten Dateien
		ScanPoolArray("Sort")
		ScanPoolArray("Save")

return CountValidKeys(ScanPool)
}

RefreshPdfIndex(BefundOrdner) {	                                   	;-- frischt das ScanPool Object auf

	; globale Variabeln im aufrufenden Skript: ScanPool := Object()

		static newPDFs  	:= 0
		PageSum	:= 0
		FileCount	:= 0

	; Kopie des ScanPool-Objektes anlegen
		tmpObj		:= Object()
		tmpObj 	:= ScanPool

	; ScanPool-Objekt leeren
		ScanPool	:= Object()

	; alle pdf Dokumente einlesen
		PdfDirList:= ReadDir(BefundOrdner, "pdf")

	;nicht mehr vorhandene Dateien aus dem Index nehmen
		For key, val in tmpObj
			If FileExist(BefundOrdner "\" StrSplit(val, "|").1)
					ScanPool.Push(val)

	;nach noch nicht aufgenommenen Dateien suchen
		Loop, Parse, PdfDirList, `n, `r
		{
				If !ScanPoolArray("Find", A_LoopField)
				{
						FileGetSize, FSize, % BefundOrdner "\" A_LoopField, K
						ScanPool.Push(A_LoopField "|" FSize "|" pages)
						newPDFs ++
						continue
				}
		}

	;Sortieren der eingelesenen und aktualisierten Dateien
		ScanPoolArray("Sort")
		ScanPoolArray("Save")

return newPDFs
}

;}

;----------------------------------------------------------------------------------------------------------------------------------------------------------------------
; DBASE
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
ReadDbf(dbfPath, SaveTo:="", options:="") {                                                                           ;-- liest Datensätze aus einer DBASE Datei

		/*  Beschreibung

				◦ es werden als erstes Daten aus dem DBASE-Header gelesen
					1. die Anzahl der Datensätze
					2. Header Länge
					3. Länge eines Datensatzes

				◦ das Auslesen erfolgt Datensatzweise entweder in eine Datei (utf-8 konvertiert) oder in eine Variable

				~~~~~~~~~~~~~ PARAMETER ~~~~~~~~~~~~~~

				◦ SaveTo        	- 	ist der Parameter leer wird der Inhalt der DBASE Datei in eine Variable geladen
							         		! ACHTUNG: 	große DBASE Dateien sollten nicht in den RAM eingelesen werden, wenn man bei Autohotkey
									         						nicht vorher die	maximale Speichernutzung für eine Variable eingestellt hat (#MaxMem)
				        				- 	enthält SaveTo einen Dateipfad mit Dateinamen wird der Datenbankinhalt ohne RAM Zwischenspeicherung
						        			in eine .csv-Datei (mit Tabulatoren als Trennzeichen) konvertiert

				◦ options          	- 	ein Objekt bestehend aus einer variablen Anzahl an key:value Paaren
											◦ StartWithSet:      	Datensatznummer bei der mit dem Lesen begonnen werden soll
																		bei Übergabe einer 0 wird bei einer bereits geöffneten Datei ab dem nächsten Datensatz fortgesetzt
																		ein nicht vorhandener Parameter ist dasselbe als wenn eine 1 übergeben wird, der Lesezugriff erfolgt
																		ab dem ersten Datensatz der DBASE Datei
				                         	◦ MaxDataSets: 	 	maximale Anzahl zu lesender Datensätze, die Funktion wird bei Erreichen dieser Zahl beendet
									                         			0 oder nicht vorhanden = die Anzahl der gelesenen Datensätze wird nicht begrenzt
											◦ Search:            	nur die passenden Datensätze werden zurück gegeben
											◦ CloseAfterRead:	der Lesezugriff auf die DBASE Datei wird nach dem Auslesen der Daten beendet (true oder false)

				◦ Rückgabewert: 	UTF-8 String (0000x:Tabellenspalten getrennt durch Leerzeichen/Tabs `n)

				◦ ANMERKUNG: 	korrekte Umwandlung getestet an PatGRArt.dbf, Nummern.dbf, befgonr.dbf, Patient.dbf
											nicht funktionierend bei beftext.dbf

		 */

	; Variablen
		static WorkDb, ConvFile, appendix, CntDataSets, HeaderLen, LenDataSet, buffin, lastpos, lastDatensatz
		static MB := 1024*1024 ; maximale Dateigröße der csv Datei (hier 1 MB)
		static dbf

	; Falls DBASE Datei noch nicht geöffnet ist, jetzt den Lesezugriff erstellen
		If (WorkDb != dbfPath) {

			If !(dbf := FileOpen(dbfPath, "r", "CP1252")) {
				MsgBox, % "Dbf - file read failed."
				return 0
			}

		; DBASE Header auslesen, DBASE Pfad speichern
			CntDataSets 	:= SeekReadNum(dbf, 4	, 4, "Uint")     	; maximale Zahl der Datensätze
			HeaderLen  	:= SeekReadNum(dbf, 8	, 2, "Uint") 		; Header Länge
			LenDataSet	:= SeekReadNum(dbf, 10	, 2, "Uint")	    	; Länge eines Datensatzes
			WorkDb    	:= dbfPath
			VarSetCapacity(buffin, LenDataSet, 0) ; buffer vorbereiten
			lastpos      	:= 0
			lastDatensatz	:= 0
			pos				:= HeaderLen

		}

	; Konvertierungsdatei für Schreibzugriff öffnen
		If (StrLen(SaveTo) > 0) && (ConvFile != SaveTo) {
				ConvFile := SaveTo
				appendix := 1
				SaveToThis  := RegExReplace(SaveTo, "\.\w+$", "")
				If !(cfile := FileOpen(SaveToThis appendix ".csv", "w", "UTF-8")) {
					MsgBox, % SaveTo "`nfile write failed."
					return 0
				}
		}

	; Leseposition festlegen
		If (options.StartWithSet > 0)
			pos := options.StartWithSet * LenDataSet
		else ;if (!options.StartWithSet) && (lastpos = 0)
			pos := lastpos + LenDataSet

		data := ""
		;SciTEOutput(pos)

	; Datei Leseschleife
		while (!dbf.AtEOF) {

				datensatz := A_Index
				ToolTip, % "Datensatz: " lastDatensatz + A_Index "/" CntDataSets

				dbf.Seek(pos, 0)                                                	; Dateizeiger versetzen
				pos += LenDataSet                                           	; Leseposition + eine Datensatzlänge
				dbf.RawRead(buffin, LenDataSet)                        	; liest einen Datensatz
				string := StrGet(&buffin, LenDataSet, "cp1252")
				string := Trim(StrReplace(string, "`r`n", " "))
				string := StrReplace(string, "`n", " ")

				If (StrLen(String) = 0)
						continue

				If (StrLen(SaveToThis) > 0) {                               	; Daten in Konvertierungsdatei schreiben

						cfile.Write("(" SubStr("0000000" pos, -5) ") " string "`n")
						If (cfile.Length > MB) {
							SciTEOutput("Dateilänge: " cfile.Length)
							cfile.Close()
							appendix ++
							If (appendix > 3)
								break
							cfile := FileOpen(SaveToThis appendix ".csv", "w", "UTF-8")
						}

				}
				else
					data .= string "`n"

				If options.MaxDataSets && (A_Index >= options.MaxDataSets)
					break

		}

		lastDatensatz += Datensatz
		lastpos := pos

	; Dateizugriffe beenden
		If dbf.AtEOF || options.CloseAfterRead
		{
			data .= "###End of file"
			dbf.Close()
		}

		If (StrLen(SaveToThis) > 0)
				cfile.Close()

return data
}

SeekReadNum(file, pos, len, Type) {

	VarSetCapacity(buffin, len, 0)
	file.Seek(pos)
	file.RawRead(buffin, len)

return NumGet(buffin, 0, Type)
}

;}

;--------------------------------------------------------------------------------------------------------------------------------------------------------------------
; BILDER IM BEFUNDORDNER
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------;{


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

FuzzyNameMatch(Name1, Name2, diffmax := 0.12) {                                                         	;-- Fuzzy Suchfunktion für Vor- und Nachnamensuche

	surname1	:= Trim(Name1[1])
	prename1	:= Trim(Name1[2])
	surname2	:= Trim(Name2[3])
	prename2	:= Trim(Name2[4])

	;FileAppend, % "Name1: " prename1 ", " surname1 "    matching with Name2: " prename2 ", " surname2 "`n", %A_ScriptDir%\logs\MatchingMethodLog.txt
	;FileAppend, % StrDiff(surname1 . prename1, prename2 . surname2) "`n", %A_ScriptDir%\logs\MatchingMethodLog.txt
	;FileAppend, % "Name1: " prename1 ", " surname1 "    matching with Name2: " surname2 ", " prename2 "`n", %A_ScriptDir%\logs\MatchingMethodLog.txt
	;FileAppend, % StrDiff(surname1 . prename1, surname2 . prename2) "`n`n", %A_ScriptDir%\logs\MatchingMethodLog.txt

	if ( StrDiff(surname1 . prename1, surname2 . prename2) <= diffmax )  ||  ( StrDiff(surname1 . prename1, prename2 . surname2) <= diffmax )
			return 1

return 0
}

FuzzyCompareNameLists(NamesArr, oPat, delta) {											;#### überarbeiten!

	/* FUNCTION based on StringDifference (SIFT3) - or fuzzysearch, this is a FuzzyCompare function
			bei der manuellen Eingabe von Kundennamen entstehen zwangsläufig häufig andere
			Schreibweisen von Namen. Ein Computer ist auf einfache Weise aber auf exakte
			Daten angewiesen.
			Da die Fehlerhäufigkeit von Mensch zu Mensch unter selbst als normal empfundener
			Arbeitsbelastung stark variiert, bringt es nichts Lösungsversuche beim Menschen anzusetzen.

			Diese Funktion sollte dann eingesetzt werden wenn die Personendaten eingegeben
			werden und zu diesem Zeitpunkt nur umständlich oder kein Zugriff auf z.B.
			eine Kundendatenbank bestand. Um Menschen nicht mit den Unzulänglichkeiten
			des Computer zu quälen kann hier ein Ähnlichkeitsvergleich für eine ganze Liste an Namen
			durchgeführt werden.
			Die Funktion gibt die korrekt geschriebenen Namen zurück und weist jeweils die
			abweichenden Schreibweisen den korrekten Namen zu.
			So könnte man mit der Zeit eine Datenbank mit den häufigsten Abweichungen/Tipfehlern
			erstellen und bei der Eingabe die Namen simultan korrigieren. Dies hätte den Vorteil
			das man Neukunden besser richtig und auch schneller erfassen kann.

			When entering customer names manually, other spelling of names inevitably occurs.
			A computer is simply dependent on exact data. Since the frequency of errors varies
			greatly from person to person and workload, it does not take anything to do there.
			This function should be used when the personal data had to be entered and at that
			time only cumbersome or no access to e.g. a customer database existed. In order not
			to torment people with the inadequacies of the computer (we know which person is meant)
			, a similarity comparison can be made here for a whole list of names. The function
			returns the correctly written names and assigns the different spellings.
			So one could over time create a database with the most common deviations / typos
			and correct the names simultaneously. This has the advantage that one can better
			capture new customers correctly, e.g.

			Parameter:
			oPat 	-	is an Autohotkey key:value object
			delta	- 	StrDiff returns a count between 0.0 to 1.0 after matching the strings
						1.0 - means all characters are equal in order and in uppercase and lowercase
						minimal is 0.0 - there is no match
						for german pre- and surenames i think a 0.2 is a good result not to
						is a good result in which very rarely actually different names are still
						identified as belonging together.

			Result:		is a txtlist delimered by `n
	*/


	If InStr(SurPreNameStr, "`,")
		LName:= StrSplit(SurPreNameStr, "`,")
	else
		return 1

	LPatList:=""

	For PatID in oPat
	{
			sur:= oPat[PatID].Nn
			pre:= oPat[PatID].Vn

			distanceOne:= StrDiff(sur . pre, LName[1] . LNames[2])
			distanceTwo:= StrDiff(sur . pre, LNames[2] . LNames[1])

			If (distanceOne > delta) or (distanceTwo > delta)
			{
					LPatList.= sur . "`, " . pre . " (" . oPat[PatID].Gt . ") (" . oPat[PatID].Gd . ") (" . PatID . ")`n"
			}
	}

return RTrim(LPatList, "`n")				;letztes Trennzeichen entfernen
}

FuzzySearch(string1, string2) {

	lenl := StrLen(string1)
	lens := StrLen(string2)
	if(lenl > lens)
	{
		shorter := string2
		longer := string1
	}
	else if(lens > lenl)
	{
		shorter := string1
		longer := string2
		lens := lenl
		lenl := StrLen(string2)
	}
	else
		return StrDiff(string1, string2)
	min := 1
	Loop % lenl - lens + 1
	{
		distance := StrDiff(shorter, SubStr(longer, A_Index, lens))
		if(distance < min)
			min := distance
	}
	return min
}

StrDiff(str1, str2, maxOffset:=5) {												    	;-- SIFT3 : Super Fast and Accurate string distance algorithm, Nutze ich um Rechtschreibfehler auszugleichen

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

StrDiffFromWords(words, comparison, tolerance) {                         	;-- SIFT3-extended: unscharfe Suche nach dem Vorkommen einer Anzahl von Wörtern in einem Mehrwortstring (Versuch! nicht fehlerfrei)

	;Words 			- a string containing words (text) - e.g. a filename, Special characters in the string are filtered out by the RegExMatch
	;comparison 	- space delimited list of words to found
	;tolerance		- the accuracy of the search

	newpos:=1, found:=0, arr:=[]
	cpr:=StrSplit(comparison, " ")

	Loop {
			If (fpos:=RegExMatch(words, "O)([A-züöä]+)", out, newpos)) {
					arr.Push(out[1])
					newpos:=fpos+StrLen(out[1])
			}
	} until (fpos=0)

	Loop % mc:=cpr.Length()
	{
			cIdx:=A_Index
			Loop % arr.Length() {
					If (StrDiff(arr[A_Index], cpr[CIdx])>tolerance)
							found += 1
					If ( found=mc )
							return 1
			}
	}

	return 0
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
		   * @param     [in]      dict    An array of strings to query against.
		   * @param     [in]      query   Any query in the form of a string.
		   * @returns   {object}          An array of match objects consisting of the matched
		   *                              string and information pertaining to where the match
		   *                              occured within the string.
		  *
		*
	 */

  if (! IsObject(dict) || query == "")
    return 0
  matches := []
  for each, token in StrSplit(query)
  {
    re_string .= (re_string ? ".*?" : "") "(" token ")"
  }
  for each, string in dict
  {
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

FindPatientName(TextFilePath) {

	regEx[1]:= "\d\d\d\d\d\s[a-zA-ZäöüÄÖÜ]+"							;z.B. 13353 Berlin
	regEx[2]:= "geb\.\s*am\s*\d\d\.\d\d\.\d\d\d\d"					;z.B. geb. am 30.01.1954 aber auch geb.am30.01.1954
	regEx[3]:= "\swohnhaft\s"															;
	regEx[4]:= "Frau\s"							;
	regEx[5]:= "Herr\s"							;

}

;}

;----------------------------------------------------------------------------------------------------------------------------------------------------------------------
; HILFSFUNKTIONEN - STRING HANDLING,
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------;{

VorUndNachname(Name) {				                              		        				;-- teilt einen Komma-getrennten String und entfernt Leerzeichen am Anfang und Ende eines Namens
	Arr    	:=[]
	Arr[1]	:= StrSplitEx(Name, 1, ",")		;Trimmed den String
	Arr[2]	:= StrSplitEx(Name, 2, ",")
return Arr
}

StrSplitEx(str, nr=1, splitchar:="|") {                                                       		;-- Trim-Split mit Rückgabe eines Wertes (keine Array-Rückgabe)
	splitArr:= StrSplit(str, splitchar)
return Trim(splitArr[nr])
}

GetAddendumDbPath() {                                                                            	;-- liest den Pfad zum Datenbankordner aus der Addendum.ini

	If (AddendumDir = "") {
			AddendumDir:= FileOpen("C:\albiswin.loc\AddendumDir","r").Read()
			If !AddendumDir {
					MsgBox, 262144 , % "Addendum - " A_ScriptName,
					(LTrim
						Der Pfad zu den Dateien für Albis on Windows ist nicht hinterlegt.
						Bitte starten Sie das AddendumStarter-Skript aus dem Addendum-
						Hauptverzeichnis, damit alle notwendigen Dateien und Verzeichnisse
						lokalisiert werden können.
						Das Skript wird jetzt beendet!
					), 15
					ExitApp
			}
	}

	IniRead, AddendumDbPath, % AddendumDir . "\Addendum.ini", Addendum, AddendumDbPath
	If InStr(AddendumDbPath, "Error") {
			MsgBox, 262144 , % "Addendum - " A_ScriptName,
			(LTrim
				Es wurde noch kein Datenbankpfad durch das Hauptskript Addendum.ahk
				angelegt. Diese Funktion greift auf Dateien in diesem Ordner zu.
				Bitte starten Sie Addendum.ahk!
			)
			ExitApp
	} else {
		return StrReplace(AddendumDbPath, "%AddendumDir%", AddendumDir)
	}

}

;}



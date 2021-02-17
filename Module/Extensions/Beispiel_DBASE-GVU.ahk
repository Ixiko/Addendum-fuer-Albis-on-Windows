;	Beispielskript für die Verwendung von Addendum_DBASE.ahk
; 	findet Patienten bei denen eine Vorsogeuntersuchung abgerechnet werden kann

		#NoEnv
		#MaxMem 4096
		SetBatchLines, -1
		ListLines, Off

		; Skriptstartzeit
			starttime	:= A_TickCount

		SciTEOutput()

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Objekte / Variablen / Daten
	;-------------------------------------------------------------------------------------------------------------------------------------------;{

		; Arbeitsobjekte
		GVUVorschlag := Object()

		; wichtige Pfade
		RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)
		basedir         	:= "M:\albiswin\db"
		pathGVUListe 	:= AddendumDir "\Tagesprotokolle"

		; minimaler GVU Untersuchungsabstand
		minAbstand  	:= 3     	; Jahre
		minAlter            := 35    	; Abrechnung möglich ab dem Alter von

		; das aktuelle Abrechungsquartal
		Abrechnungsquartal 	:= "0420"
		AbrQ      	:= QuartalTage({"Aktuell":Abrechnungsquartal})

		; letztes GVU Datum muss älter als dieses sein
		MinJ	    	:= "20" SubStr("0" (SubStr(Abrechnungsquartal, 3, 2) - minAbstand), -1)
		MinM     	:= (SubStr(Abrechnungsquartal, 1, 2) * 3) - 1
		MinT        	:= 1
	;}

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; vorhandene Untersuchungen laden
	;-------------------------------------------------------------------------------------------------------------------------------------------;{
		GVUListe := ReadGVUListe(pathGVUListe, Abrechnungsquartal)
		SciTEOutput("GVUListe erfasst mit " GVUListe.Count() " Untersuchungen")
	;}

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Patientendaten aus der PATIENT.DBF laden
	;-------------------------------------------------------------------------------------------------------------------------------------------;{
		infilter 	:= ["NR", "NAME", "VORNAME", "GEBURT", "MORTAL", "LAST_BEH", "RECHART"]
		PatDBF	:= ReadPatientDBF(basedir, infilter,, 1)
		SciTEOutput("Anzahl Patienten Datensätze: " PatDBF.Count())

		;FileOpen(A_ScriptDir "\GVUPatienten.JSON", "w", "UTF-8").Write(JSON.DUMP(PatDBF,,2))
	;}

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



ExitApp

GetDBFData(DBASEfilepath, inpattern:="", outpattern:="", options:="", debug:=0) {

	dbfile      	:= new DBASE(DBASEfilepath, debug)
	res        	:= dbfile.OpenDBF()
	matches 	:= dbfile.Search(inpattern, 0, options)
	res         	:= dbfile.CloseDBF()
	dbfile    	:= ""

return matches
}

ReadPatientDBF(basedir, infilter="", outfilter="", debug=0) {                                   	;-- gibt nur benötigte Daten der PATIENT.DBF zurück

		PatDBF := Object()
		strObj	:= Object()

	; für Abrechnungsüberprüfungen die geschätzt minimal notwendigste Datenmenge
		If !IsObject(infilter)
			infilter := ["NR", "NAME", "VORNAME", "GEBURT", "MORTAL", "LAST_BEH"]

		database 	:= new DBASE(basedir "\PATIENT.dbf", debug)
		res        	:= database.OpenDBF()
		matches	:= database.GetFields(infilter)
		res         	:= database.CloseDBF()

		For idx, m in matches {

			strObj	:= Object()
			For key, val in m
				strObj[key] := val

			PatDBF[m.NR] := strObj

		}

return PatDBF
}

ReadGVUListe(path, Quartal) {                                                                                	;-- Einlesen der manuell angelegten untersuchten Patienten

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

GetQuartal(Datum, Format:="") {                                                                           	;--zum Errechnen zu welchem Quartal das übergebene Datum gehört

	; Funktionsbeschreibung:
	; Datum im folgenden Format ist erlaubt: 13.02.2017 oder 13.2.17 - in dieser Form der Übergabe müssen die Punkte im Übergabestring vorhanden sein
	; 14.5.2018: absofort ist die Übergabe des Begriffs heute möglich, es wird damit das aktuelle Quartal berechnet
	; Format ist ein Trennzeichen zwischen Quartal und Jahr z.B. Format="`/" dann wäre die Ausgabe so z.B. 01/18 , Format kann jedes beliebige Zeichen oder mehrere enthalten

	If InStr(Datum, "heute") {
		Monat	:= A_MM
		Jahr		:= SubStr(A_Year, 3, 2) ;die letzten zwei Zeichen
	} else {
		split		:= StrSplit(Datum, ".")
		Monat	:= split[2]
		Jahr		:= Substr(split[3], StrLen(split3) - 1, 2)
	}

	;MsgBox, % Monat ", " Jahr "`n" SubStr("0" . Ceil(Monat/3), -1) . Format . Jahr
return SubStr("0" . Ceil(Monat/3), -1) . Format . Jahr
}

HowLong(Date1,Date2) {

	; https://www.autohotkey.com/boards/viewtopic.php?t=54796

	year1 := SubStr(Date1, 1, 4), 	month1 := SubStr(Date1, 5, 2), 	day1 := SubStr(Date1, 7, 2)
	year2 := SubStr(Date2, 1, 4), 	month2 := SubStr(Date2, 5, 2), 	day2 := SubStr(Date2, 7, 2)
	month := leapyear(year1) ? [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31] : [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
	if (day1 > day2)
		day2 := day2 + month[month1], month2 := month2 - 1
	if (month1 > month2)
		year2 := year2 - 1, month2 := month2 + 12
	D:= day2 - day1, 	M:= month2 - month1, 	Y := year2 - year1

return {"years":Y, "months":M, "days":d}
}

leapyear(year) {
    if (Mod(year, 100) = 0)
        return (Mod(year, 400) = 0)
    return (Mod(year, 4) = 0)
}

ConvertDBASEDate(DBASEDate) {
	return SubStr(DBaseDate, 7, 2) "." SubStr(DBaseDate, 5, 2) "." SubStr(DBaseDate, 1, 4)
}

QuartalTage(Quartal) {

	/*  Funktion zur Berechnung von wichtigen Tagen eines Quartals im Jahr

			Die Funktion hat eine erweiterte Fähigkeit um das von Albis benutzte Datumsformat .

			Quartal als Objekt übergeben mit folgenden Schlüsseln:
			1) Quartal.Aktuell  	=	in der Form 0320 - Monat Jahr, möglich ist auch 03/20
												WICHTIG: Berechnungen sind damit nur für das aktuelle Jahr und dem Jahr davor möglich
	         		oder
			2) Quartal.TDatum	=	Tagesdatum. Wird dieser Parameter angegeben errechnet die Funktion zunächst das Quartal, dazu muss ein weiterer
			     								Parameter übergeben werden der das Format spezifiziert.
												TDatum kann folgendes Format haben 20201018 oder 18102020.
												Punkte zwischen Tag, Monat und Jahreszahl können auch übergeben werden (z.B. 18.10.2020).
				    							2a) .TFormat: 	YYYYMMDD o. DDMMYYYY - dies vermittelt die Reihenfolge des Zahlenformates
					    												DBASE ist als Wert auch möglich und entspricht dann YYYYMMDD

	 */

		If !IsObject(Quartal)	{
			throw Exception(A_ThisFunc " (" A_LineFile ") : Fehler beim Funktionsaufruf, der übergebene Parameter ist kein Objekt!")
			return 0
		}

		If Quartal.TDatum {

				Quartal.TDatum := StrReplace(Quartal.TDatum, ".")

			; kontrolliert die Validität des Parameter TFormat
				If !RegExMatch(Quartal.TFormat, "i)DBASE") {
					RegExMatch(Quartal.TFormat, "i)Y+", YStr)
					If StrLen(YStr) not in 2,4
					{
						throw Exception(A_ThisFunc " (" A_LineFile ") : Fehler beim Funktionsaufruf, die Länge des Parameter .TFormat ist unzulässig (2 oder 4 Zeichen)!`nDie benutzte Länge ist " StrLen(YStr))
						return 0
					}
				}

				If RegExMatch(Quartal.TFormat, "i)(?<Y>Y+)(?<M>M+)(?<D>D+)|DBASE", P) {
					If (P = "DBASE")
						PY := "YYYY", PM := "MM", PD := "DD"
					Jahr  	:= SubStr(Quartal.TDatum, 1                                         	, StrLen(PY))
					Monat	:= SubStr(Quartal.TDatum, StrLen(PY)                    	+ 1	, StrLen(PM))
				}
				else if RegExMatch(Quartal.TFormat, "i)(?<D>D+)(?<M>M+)(?<Y>Y+)|DBASE", P) {
					Monat	:= SubStr(Quartal.TDatum, StrLen(PD)                   	+ 1 	, StrLen(PM))
					Jahr  	:= SubStr(Quartal.TDatum, StrLen(PD) + StrLen(PM) 	+ 1	, StrLen(PY))
				}
				else {
					throw Exception(A_ThisFunc " (" A_LineFile ") : Error in function call, parameter .TFormat with invalid format!")
					return 0
				}

				Quartal.Quartal 	:= SubStr("0" Ceil(Monat/3), -1)
				Quartal.Jahr      	:= StrLen(PY) = 4 ? Jahr : (Jahr > SubStr(A_YYYY, 3, 2) ? (SubStr(A_YYYY, 1, 2) - 1) . Jahr : SubStr(A_YYYY, 1, 2) . Jahr)
				Quartal.Aktuell    	:= Quartal.Quartal (StrLen(PY) = 4 ? SubStr(Jahr, 3, 2) : Jahr)
		}
		else If RegExMatch(Quartal.aktuell, "(?<Q>\d{2}).*(?<Y>\d{2})", P) {
				If (PQ = 0 || PQ > 4) {
					throw Exception(A_ThisFunc " (" A_LineFile ") : Fehler beim Funktionsaufruf, das Quartal ist ungültig! (" PQ "?)")
					return 0
				}
				Quartal.Quartal   	:= PQ
				Quartal.Jahr      	:= PY > SubStr(A_YYYY, 3, 2) ? (SubStr(A_YYYY, 1, 2) - 1) . PY : SubStr(A_YYYY, 1, 2) . PY
		}
		else {
				throw Exception(A_ThisFunc " (" A_LineFile ") : Error in function call, parameter Quartal with invalid format!")
				return 0
		}

		Quartal.MBeginn    	:= SubStr("0" (Quartal.Quartal - 1) * 3 + 1	, -1)
		Quartal.MEnde      	:= SubStr("0" Quartal.MBeginn + 2            	, -1)
		Quartal.TDBeginn	   	:= "01." SubStr("0" Quartal.MBeginn, -1) "." Quartal.Jahr
		Quartal.TDEnde       	:= DaysInMonth(Quartal.Jahr Quartal.MEnde) "." Quartal.MEnde "." Quartal.Jahr
		Quartal.Monat       	:= Object()

		Loop 3 {

			dsim     	:= DaysInMonth(Quartal.Jahr (Quartal.MBeginn - A_Index + 1))
			dsgesamt	+= dsim
			wochen		:= Round(dsim/7)
			wogesamt	+= wochen
			FormatTime, NameLang, % (Quartal.MBeginn - A_Index + 1)	, MMMM
			FormatTime, NameKurz, % (Quartal.MBeginn - A_Index + 1)	, MMM
			Quartal.Monat.Push({"Tage":dsim, "Wochen":wochen, "NameLang":NameLang, "NameKurz":NameKurz})

		}

		Quartal.Tage        	:= dsgesamt
		Quartal.Wochen      	:= wogesamt
		Quartal.DBaseStart   	:= Quartal.Jahr Quartal.MBeginn "01"
		Quartal.DBaseEnd   	:= Quartal.Jahr Quartal.MEnde Quartal.Monat.1.Tage


return Quartal
}

DaysInMonth(date:="") {
    date := (date = "") ? (a_now) : (date)
    FormatTime, year  	, % date, yyyy
    FormatTime, month	, % date, MM
    month += 1                                                      	; goto next month
    if (month > 12)
        year += 1, month := 1                                  	; goto next year, reset month
    month := (month < 10) ? (0 . month) : (month)  	; 0 to 01
    new_date := year . month
    new_date += -1, days                                      	; minus 1 day
return subStr(new_date, 7, 2)
}



#Include %A_ScriptDir%\..\..\Include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#Include %A_ScriptDir%\..\..\lib\class_JSON.ahk


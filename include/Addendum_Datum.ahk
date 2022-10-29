; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                           	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                             	Funktionen für die Berechnung/Umwandlung von Tagesdaten, Quartalsdaten
;                               	by Ixiko started in September 2017 - last change 09.07.2022 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

; Datum berechnen
AddToDate(Feld="", val="", timeunits="") {                                                         	;-- addiert Tage bzw. eine Anzahl von Monaten zu einem Datum hinzu
	calcdate:= SubStr(Feld.Datum, 7, 4) . SubStr(Feld.Datum, 4, 2) . SubStr(Feld.Datum, 1, 2)
	calcdate += val, %timeunits%
	FormatTime, newdate, % calcdate, dd.MM.yyyy
	If IsObject(Feld)
		ControlSetText,, % newdate, % "ahk_id " Feld.hwnd
return newdate
}

DateDiff(fnTimeUnits, fnStartDate, fnEndDate) {                                                  	;-- berechnet Tagesdifferenzen zwischen zwei Tagen

	; deutsches Datumsformat dd.MM.yyyy wird automatisch umgerechnet
	; returns the difference between two timestamps in the specified units

	; declare local, global, static variables

	; set default return value
		TimeDifference := 0

	; Convert date from german date format to englisch dateformat (only DD.MM.YYYY to YYYY.DD.MM)
		fnStartDate	:= ConvertGerDateToEng(fnStartDate)
		fnEndDate 	:= ConvertGerDateToEng(fnEndDate)

	; validate parameters
		If fnTimeUnits not in YY,MM,DD,HH,MI,SS
			Throw Exception("fnTimeUnits were not valid")

		If fnStartDate is not date
			Throw Exception("fnStartDate was not a date")

		If fnEndDate is not date
			Throw Exception("fnEndDate was not a date")


	; initialise variables
		DatePadding := "00000101000000"
		StartDate   	:= fnStartDate	SubStr(DatePadding, StrLen(fnStartDate)+1) ; normalise start date to 14 digits
		EndDate    	:= fnEndDate 	SubStr(DatePadding, StrLen(fnEndDate  )+1) ; normalise end   date to 14 digits

	; for day or time, use native function
		If RegExMatch(fnTimeUnits, "i)(DD|HH|MI|SS)") 	{

			TimeDifference := EndDate
			TimeUnit := fnTimeUnits = "SS" ? "S"
                     :  fnTimeUnits = "MI" ? "M"
                     :  fnTimeUnits = "HH" ? "H"
                     :  fnTimeUnits = "DD" ? "D"
                     :                       ""

			EnvSub, TimeDifference, % StartDate, % TimeUnit
		}

		; for year or month
		FormatTime, StartYear   	, % StartDate, yyyy
		FormatTime, StartMonth	, % StartDate, MM
		FormatTime, StartDay    	, % StartDate, dd
		FormatTime, StartTime   	, % StartDate, HHmmss

		FormatTime, EndYear    	, % EndDate, yyyy
		FormatTime, EndMonth 	, % EndDate, MM
		FormatTime, EndDay      	, % EndDate, dd
		FormatTime, EndTime    	, % EndDate, HHmmss

		If RegExMatch(fnTimeUnits, "MM") {

			StartInMonths	:= (StartYear*12)+StartMonth
			EndInMonths  	:= (EndYear  *12)+EndMonth
			TimeDifference	:= EndInMonths-StartInMonths

			If (StartDate < EndDate)
				If (EndDay < StartDay)
					TimeDifference--

			If (StartDate > EndDate)
				If (EndDay > StartDay)
					TimeDifference++
		}

		If RegExMatch(fnTimeUnits, "YY") {

			TimeDifference := EndYear-StartYear

			If (StartDate < EndDate)
				If (EndMonth <= StartMonth)
					If (EndDay < StartDay)
						TimeDifference--

			If (StartDate > EndDate)
				If (EndMonth >= StartMonth)
					If (EndDay > StartDay)
						TimeDifference++
		}


	; return
Return TimeDifference
}

DaysBetween(FirstDate, LastDate) {                                                                    	;-- errechnet die Tage zwischen zwei Tagen

	FirstDate	:= FirstDate	. (StrLen(FirstDate) = 14 ? "" : SubStr("00000000000000", 1, StrLen(FirstDate)-13))
	diff       	:= LastDate 	. (StrLen(LastDate) = 14 ? "" : SubStr("00000000000000", 1, StrLen(LastDate)-13))
	EnvSub, diff, % FirstDate, Days

return diff
}

DaysInMonth(date:="") {                                                                                    	;-- errechnet die Anzahl der Tage des Monats
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

GetQuartal(Datum, Trenner:="") {                                                                      	;-- berechnet zu welchem Quartal das übergebene Datum gehört

	; Funktionsbeschreibung:
	; Datum: 	erlaubt ist "13.02.2017" oder "13.2.17" - in dieser Form der Übergabe müssen die Punkte im Übergabestring vorhanden sein
	; 				oder "heute" - es wird das aktuelle Quartal berechnet
	; Trenner: ist ein Trennzeichen zwischen Quartal und Jahr. Trenner:="/" würde z.B. 01/18 ergeben,
	; 				der Trenner kann jedes beliebige Zeichen oder auch mehrere enthalten

	If InStr(Datum, "heute")
		Monat	:= A_MM, Jahr	:= SubStr(A_Year, 3, 2) ; die letzten zwei Zeichen
	else
		Monat	:= StrSplit(Datum, ".").2, Jahr	:= Substr(StrSplit(Datum, ".").3, StrLen(StrSplit(Datum, ".").3)-1, 2)

	;RegExMatch(Format, "(?<Q>Q+)(?<Y>Y+)", C)

return SubStr("0" . Ceil(Monat/3), -1) . Trenner . Jahr
}

GetQuartalEx(Datum, Format:="QQYY") {                                                          	;-- flexiblere Ein-/Ausgabeformate als bei der GetQuartal

	; Funktionsbeschreibung:
	; Datum: 	erlaubt sind folgende Datenformate dd.MM.yyyy "13.02.2017" oder dd.M.yy "13.2.17" oder yyyyMMdd "20170213"
	; 				oder "heute" auch "today" - es wird das aktuelle Quartal berechnet
	; Format:	möglich ist auch QYYYY o. YYYYQ, für Ausgabereihenfolge und Anzahl der Zeichen z.B. 0118 o. 012018
	;				oder YYYY-Q - das Zeichen zwischen den Zahlen wird der Trenner 2018-1

	; return Format auswerten
		RegExMatch(Format, "(?<Q1>Q+)(?<T1>[^QY]*)(?<Y1>Y+)|(?<Y2>Y+)(?<T2>[^QY]*)(?<Q2>Q+)", c)
		LenQ 	:= StrLen(cQ1) > 0	? StrLen(cQ1)	: StrLen(cQ2)
		LenY 	:= StrLen(cY1) > 0 	? StrLen(cY1) 	: StrLen(cY2)
		If (LenQ>2)
			throw A_ThisFunc ": Das Rückgabeformat für die Quartalszahl ist auf maximal 2 Zeichen begrenzt!"

	; Monat und Jahr trennen
		If RegExMatch(Datum, "i)(heute|today)") {
			QMonat	:= A_MM
			QJahr	   	:= LenY < 3 ? SubStr(A_YYYY, 3, 2) : A_YYYY ; die letzten zwei Zeichen oder das 4stellige Jahr
		}
		else if RegExMatch(Datum, "^\d{8}$") {           ; yyyyMMdd
			QMonat 	:= SubStr(Datum, 5, 2)
			QJahr      	:= SubStr(Datum, (LenY < 3 ? 3 : 1), (LenY < 3 ? 2 :4))
		}
		else {
			 If !RegExMatch(Datum, "\d{1,2}\.(?<Monat>\d{1,2})\.(?<Jahr>\d{4}|\d{2})", D)
				throw A_ThisFunc ": Ein falsches Datumsformat wurde übergeben!`nrichtig ist: D(1-2)[.]M(1-2)[.]YYYY(2-4) oder yyyyMMdd"

			QMonat := DMOnat
			QJahr  	:= LenY<3 && StrLen(DJahr)=4 ? SubStr(DJahr, 3, 2) : DJahr
		}

	; Quartalszahl erstellen
		QZ := SubStr("00" Ceil(QMonat/3), -1*(LenQ-1))

	; Format QQYY
		;~ If (StrLen(cQ1) > 0)
			;~ return QZ . cT1 . Jahr

return cQ1 ? QZ . cT1 . QJahr : QJahr . cT2 . QZ
}

Trimester(month) {                                                                                            	;-- einfachste Funktion um die Trimesterzahl eines Monats im Quartal zu erhalten
	; zurückgegeben werden nur Werte zwischen 1 bis 3
return Mod(month+2, 3) + 1
}

HowLong(Date1, Date2) {                                                                                  	;-- berechnet die Anzahl der Jahre, Monate u. Tage zw. zwei Tagen

	; Format YYYYMMDD
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

Age(birthday, CalculationDate) {                                                                        	;-- Lebensalter berechnen

	; possible formats: d[d].M[M].YYYY or YYYY.M[M].d[d] or YYYYmmdd

	If 	RegExMatch(birthday, "^\s*(?<D>\d{1,2})\.(?<M>\d{1,2})\.(?<Y>\d{4})$", birth)
	|| RegExMatch(birthday, "^\s*(?<Y>\d{4})[.\-](?<M>\d{1,2})[.\-](?<D>\d{1,2})$", birth)
		birthday := birthY . SubStr("0" birthM, -1) . SubStr("0" birthD, -1)

	If RegExMatch(CalculationDate, "^\s*(?<D>\d{1,2})\.(?<M>\d{1,2})\.(?<Y>\d{4})$", Calc)
	|| RegExMatch(CalculationDate, "^\s*(?<Y>\d{4})[.\-](?<M>\d{1,2})[.\-](?<D>\d{1,2})$", Calc)
		CalculationDate := CalcY . SubStr("0" CalcM, -1) . SubStr("0" CalcD, -1)

	Age 	:= HowLong(birthday, CalculationDate)

return Age.Years
}

leapyear(year) {                                                                                                	;-- Schaltjahr
    if (Mod(year, 100) = 0)
        return (Mod(year, 400) = 0)
    return (Mod(year, 4) = 0)
}

QuartalTage(Quartal) {                                                                                     	;-- zur Berechnung von wichtigen Tagen eines Quartals im Jahr

	/*  Funktion zur Berechnung von wichtigen Tagen eines Quartals im Jahr

			Die Funktion hat eine erweiterte Fähigkeit um das von Albis benutzte Datumsformat .

			PARAMETER: 				Quartal:  immer ein Objekt

			1) Quartal.Aktuell  	=	in der Form 0320 - Monat Jahr, möglich ist auch 03/20
												WICHTIG: Berechnungen sind damit nur für das aktuelle Jahr und dem Jahr davor möglich
	         		oder
			2) Quartal.TDatum	=	Tagesdatum. Wird dieser Parameter angegeben errechnet die Funktion zunächst das Quartal, dazu muss ein weiterer
			     								Parameter übergeben werden der das Format spezifiziert.
												TDatum kann folgendes Format haben 20201018 oder 18102020.
												Punkte zwischen Tag, Monat und Jahreszahl können auch übergeben werden (z.B. 18.10.2020).
				    							2a) .TFormat: 	YYYYMMDD o. DDMMYYYY - dies vermittelt die Reihenfolge des Zahlenformates
					    												DBASE ist als Wert auch möglich und entspricht dann YYYYMMDD

			RÜCKGABEPARAMETER:

			key:value - Objekt mit folgendem Inhalt
			{
				"aktuell": "0322",                                	; aktuelles Quartal
				"DBaseEnd": "20220930",                   	; letztes Tagesdatum im DBASE schreibweise
				"DBaseStart": "20220701",                  	; erstes Tagesdatum im DBASE schreibweise
				"Jahr": "2022",                                     	; das aktuelle Jahr
				"MBeginn": "07",                                  	; Nr. des ersten Monats im Quartal
				"MEnde": "09",                                    	; Nr. des letzten Monats im Quartal
				"MMitte": "08",                                    	; Nr. des mittleren Monats im Quartal
				"Monat": {
					{
						"NameKurz": "Mai",
						"NameLang": "Mai",
						"Nr": 5,
						"Tage": "31",
						"Wochen": 4
					},
					{
						"NameKurz": "Jun",
						"NameLang": "Juni",
						"Nr": 6,
						"Tage": "30",
						"Wochen": 4
					},
					{                                          	; Nr. des Monats im Quartal
						"NameKurz": "Jul",                  	; verkürzte Schreibweise
						"NameLang": "Juli",                   	; ausgeschriebener Monatsname
						"Nr": 7,                                   	; Nr. des Monats im Quartal
						"Tage": "31",                            	; Anzahl der Tage in diesem Monat
						"Wochen": 4                          	; Anzahl der Wochen in diesem Monat
					}
				],
				"Quartal": 3,                                  	; Nr. des Quartal
				"Tage": 92,                                    	; Anzahl der Tage im Quartal
				"TDBeginn": "01.07.2022",             	; erstes Tagesdatum
				"TDEnde": "30.09.2022",                 	; letztes Tagesdatum
				"Wochen": 12                                 	; Anzahl der Wochen im Quartal
			}

			letzte Änderung: 09.07.2022

	 */

	; Quartal muss jetzt kein Objekt mehr sein. Übergebe ein Datum (ddMMyyyy) oder ein Quartal (QQYY) als String
		If !IsObject(Quartal) {

			If RegExMatch(Quartal, "^(?<Z>\d{1,2})[^\d]{0,1}(?<Y>\d{2})$", Q)
				Quartal := {"aktuell": Quartal}
			else if RegExMatch(Quartal, "^\d{8}$") {
				TFormat := GetDateFormat(Quartal)
				Quartal := {"TFormat":TFormat, "TDatum":Quartal}
			}
			else if (Quartal = "aktuell") {
				Quartal := {"TFormat":"DDMMYYYY", "TDatum":A_DD A_MM A_YYYY}
			}
			else {
				throw Exception(A_ThisFunc " (" A_LineFile-7 ") : Fehler beim Funktionsaufruf, der übergebene Parameter ist kein Objekt"
										. "`nund enthält auch sonst keine verwendbaren Daten!")
				return 0
			}

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
		else If RegExMatch(Quartal.aktuell, "^(?<Z>\d{1,2})[^\d]{0,1}(?<Y>\d{2})$", P) {
				If (PZ = 0 || PZ > 4) {
					throw Exception(A_ThisFunc " (" A_LineFile ") : Fehler beim Funktionsaufruf, das Quartal muss eine Zahl zwischen 1 bis 4 sein! (" PQ "?)")
					return 0
				}
				Quartal.Quartal   	:= PZ
				Quartal.Jahr      	:= PY > SubStr(A_YYYY, 3, 2) ? (SubStr(A_YYYY, 1, 2) - 1) . PY : SubStr(A_YYYY, 1, 2) . PY   ; vierstelliges Jahr erstellen
		}
		else {
				throw Exception(A_ThisFunc " (" A_LineFile ") : Error in function call, parameter Quartal with invalid format!")
				return 0
		}

		Quartal.Monat       	:= Array()
		Quartal.MBeginn    	:= SubStr("0" (Quartal.Quartal - 1) * 3 + 1	, -1)                                                	; erster Monat
		Quartal.MMitte       	:= SubStr("0" Quartal.MBeginn + 1            	, -1)                                                	; mittlerer Monat
		Quartal.MEnde      	:= SubStr("0" Quartal.MBeginn + 2            	, -1)                                                	; letzter Monat
		Quartal.TDBeginn	   	:= "01." SubStr("0" Quartal.MBeginn, -1) "." Quartal.Jahr                                    	; erster Tag im Quartal
		Quartal.TDEnde       	:= DaysInMonth(Quartal.Jahr Quartal.MEnde) "." Quartal.MEnde "." Quartal.Jahr	; letzter Tag im Quartal
		Loop 3 {                                                                                                                                            	; je Quartal Tage und Wochen, Monatsname

			mNR        	:= Quartal.MBeginn + (A_Index-1)
			dsim     	:= DaysInMonth(Quartal.Jahr mNR)
			dsgesamt	+= dsim
			wochen		:= Round(dsim/7)
			wogesamt	+= wochen
			FormatTime, Lang	, % Quartal.Jahr mNR, MMMM
			FormatTime, Kurz	, % Quartal.Jahr mNR, MMM
			Quartal.Monat[A_Index] := {"Nr":mNR, "Tage":dsim, "Wochen":wochen, "NameLang":Lang, "NameKurz":Kurz}

		}
		Quartal.Tage        	:= dsgesamt                                                                                                      	; Anzahl der Tage im Quartal
		Quartal.Wochen      	:= wogesamt                                                                                                       	; Anzahl der Wochen im Quartal
		Quartal.DBaseStart  	:= Quartal.Jahr Quartal.MBeginn "01"                                                                	; erster Tag im DBase-Format (yyyyMMdd)
		Quartal.DBaseEnd   	:= Quartal.Jahr Quartal.MEnde DaysInMonth(Quartal.Jahr Quartal.MEnde)          	; letzter Tag im DBase-Format (yyyyMMdd)

return Quartal
}

Vorquartal(Datum, retFormat:="YYYYQ") {                                                         	;-- gibt ein Vorquartal formatiert zurück

	; Beschreibung
	; Datum: 		Formatierung siehe Funktion QuartalTage()

	; Berechnung
		QT := QuartalTage(Datum)
		lastday := QT.DBASEStart
		lastday += -1, Days
		FormatTime, LastDayQBefore, % lastday, dd.MM.yyyy
		QT := QuartalTage({"TDatum":LastDayQBefore, "TFormat":"DDMMYYYY"})

	; Rückgabestring formatieren
		retYPos	:= RegExMatch(retFormat, "Y+", retY)
		retQPos:= RegExMatch(retFormat, "Q+", retQ)
		RegExMatch(retFormat, "[^YQ]+", Divider)
		QJ   	:= StrLen(retY) = 4 	? QT.Jahr       	: SubStr(QT.Jahr, 3, 2)
		QZ	:= StrLen(retQ) > 1 	? QT.Quartal	: SubStr(QT.Quartal, 2, 1)
		retStr:= retYPos > retQPos ? QZ . Divider . J : J . Divider . QZ

return retStr
}

DateValidator(dateString, interpolateCentury:="") {                                             	;-- prüft String auf enthaltenes Datum

	/*  	DateValidator() by Ixiko 2021

			* die Funktion ist für den Abgleich/Überprüfung von Datumszahlen aus OCR-Texten geschrieben worden

			erkennt ein Datum in deutscher Schreibweise:

				dd[.,;:]\s*(Monatsname lang|Montagsname kurz|MM)[.,;:]*\s+(yy|yyyy)  / Jahr 2 oder 4 stellig
				z.B. 12.Dezember 2020, 12. Dez. 20, 12.12.2020 oder 12.12.20
				die RegEx Strings habe ich auf einer Internetseite gefunden. Es lassen nicht plausible Daten wie z.B. der 32.04.2021
				als Datum ausschließen.

			Rückgabewert:

				ein Schreibweisen Tageszahl.Monatsname Jahreszahl wird in ein reines "Zahlen"-Datum umgewandelt
				aus dem 12.Dezember 2020 wird der 12.12.2020

			optionale Änderung:

				2-stellige Jahreszahlen können in 4-stellige Jahrezahlen umgerechnet werden, wenn der Parameter:
				interpolateCentury (ipolC) einen Integerwert enthält. Die Funktion prüft diesen Wert nicht auf Plausibilität.
				Nimmt also jeden Wert entgegen. Ausserdem wird die optionale Änderung nur für 2-stellige DateStr Jahreszahlen
				ausgeführt.

				Es wird das Jahr des validierten Strings wird mit dem ipolC Jahr verglichen.
				Dabei kommt es auf die Länge (Anzahl der Stellen) des ipoLC Jahres an!

				Zweistellige valStr Daten werden in 4 stellige Daten umgewandelt, wenn ipolC vier Stellen hat.
				Es wird nur interpoliert wenn ipolC mindestens eine Stelle mehr als der dateStr besitzt.
				z.B. dateString = 12.12.20 und interpolateCentury = 2021 ergibt den 12.12.2020, übergibt man
				interpolateCentury = 2019 ergibt dies den 12.12.1920. Es wird davon ausgegangen das es sich nicht um ein Datum
				in der Zukunft handelt.

				Was wird für interpolateCentury = 219 berechnet?
				Antwort: der 12.12.120

			**letzte Änderung: 06.11.2021

	 */

	; RegEx Strings
		static global rxMonths	:= "Jan*u*a*r*|Feb*r*u*a*r*|Mä*a*rz|Apr*i*l*|Mai|Jun*i*|Jul*i*|Aug*u*s*t*|Sept*e*m*b*e*r*|Okt*o*b*e*r*|Nov*e*m*b*e*r*|Deze*m*b*e*r*"
		static rxDateOCRrpl   	:= "[,;:\s]+"
		static rxDateValidator 	:= [ 	"^(?<D>[0]?[1-9]|[1|2][0-9]|[3][0|1]).(?<M>[0]?[1-9]|[1][0-2]).(?<Y>[0-9]{4}|[0-9]{2})$"
										    	,	"^(?<D>[0]?[1-9]|[1|2][0-9]|[3][0|1])[.;,\s]+(?<M>" rxMonths ")[.;,\s]+(?<Y>[0-9]{4}|[0-9]{2})$"]

	; OCR Korrektur ausführen und erste Evaluierung (Format ist korrekt) durchführen
		dateString := Trim(dateString)
		dateString := RegExReplace(dateString, rxDateOCRrpl, ".")
		If !RegExMatch(dateString, rxDateValidator[1], d)
			If !RegExMatch(dateString, rxDateValidator[2], d)
				return

	; geschriebenen Monat in Zahl umwandeln
		If RegExMatch(dM, "(" rxMonths ")") {
			For nrMonth, rxMonth in StrSplit(rxMonths, "|")
				If RegExMatch(dM, rxMonth)
					break
			dM := nrMonth
		}

	; das Jahrhundert interpolieren
		If RegExMatch(interpolateCentury, "^\d+$") && (StrLen(interpolateCentury) > StrLen(dY)) && (StrLen(dY) = 2) {
			refYear 	:= SubStr(interpolateCentury, -1)	; die letzten 2 Stellen
			refCentury	:= SubStr(interpolateCentury, 1, StrLen(interpolateCentury) - 2)
			dY          	:= (dY > refYear ? refCentury-1 : refCentury) dY
		}

return SubStr("0" dD, -1) "." SubStr("0" dM, -1) "." dY	; Rückgabe immer im Format dd.mm.yy oder dd.mm.yyyy
}

WeekDayNr(wday, short:=true) {                                                                       	;-- ##Wochentag als Zahl oder Kurzbezeichnung

	; wday - Zahl oder Kurzname des Wochentages
	; letzte Änderung: 05.04.2021

	static WDays := ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]
	If !RegExMatch(wday, "i)^[a-z]+$")
		return short ? SubStr(WDays[wDay], 1, 2) : WDays[wDay]
	else {
		For wdNr, day in WDays
			If RegExMatch(day, "i)^" wday)
				return wdNr
	}

return
}

DayOfWeek(dateStr, DayFormat:="short", format:="dd.MM.yyyy") {                          	;-- nutzt FormatTime anstatt eigene Berechnungen anzustellen

	static retFormat := {"short":"ddd", "full":"dddd", "ddd":"ddd", "dddd":"dddd"}

	dPos 	:= RegExMatch(format, "d+"	, d)
	mPos	:= RegExMatch(format, "M+" 	, m)
	yPos   	:= RegExMatch(format, "y+"  	, y)
	year  	:= SubStr(dateStr, yPos, StrLen(y))
	month	:= SubStr("0" SubStr(dateStr, mPos, StrLen(m)), -1)
	day   	:= SubStr("0" SubStr(dateStr, dPos, StrLen(d)), -1)

	FormatTime, DayOfWeek, % year month day "000000", % retformat[DayFormat]

return DayOfWeek
}
;GetWeekday("26.08.2022", "dd.MM.yyyy", true, "full")
GetWeekday(dateStr, format="dd.MM.yyyy", NameOfDay=true, DayFormat="full" ) {  	;-- Name oder Zahl des Wochentages vom übergebenen Datum

	; jNizM Funktion modifiziert (https://www.autohotkey.com/boards/viewtopic.php?t=3352)
	; letzte Änderung: 14.09.2021

	static WDays := ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]

	dPos 	:= RegExMatch(format, "d+"	, d)
	mPos	:= RegExMatch(format, "M+"	, m)
	yPos  	:= RegExMatch(format, "y+"	, y)

	d	:= SubStr("0" SubStr(dateStr, dPos, StrLen(d)), -1)
	m	:= SubStr("0" SubStr(dateStr, mPos, StrLen(m)), -1)
	y 	:= SubStr(dateStr, yPos, StrLen(y))

    if (m < 3)  {
        m += 12
        y 	-= 1
    }

    wd := mod(d+(2*m) + floor(6*(m+1)/10) + y + floor(y/4) - floor(y/100) + floor(y/400)+1, 7)  ; + 1 -  bei mir beginnt die Woche am Montag (erster Arbeitstag)

return !NameOfDay ? wd : (DayFormat="full" ? WDays[wd] : SubStr(WDays[wd], 1, 2))
}

WeekOfYear(dateStr) {                                                                                        	;-- die Nummer der Woche im Jahr

	timestamp	:= dateStr	. (StrLen(dateStr) = 14 ? ""	: SubStr("00000000000000", 1, StrLen(dateStr)-13))
	FormatTime weekOfYear, % timestamp, YWeek

return SubStr(weekOfYear, 5, 2)
}

DateAddEx(vDate, vDiff, AddOrSub:="add") {                                                      	;-- erweiterte Datumaddition

	; from jeeswig - https://www.autohotkey.com/boards/viewtopic.php?t=59825
	; vDiff expects 1 to 6 space-separated digit sequences (any non-spaces/non-digits are ignored)
	; e.g. MsgBox, % DateAddEx(20010101, "3y")		            			; 20040101000000
	; e.g. MsgBox, % DateAddEx(20010101, "3y 3m 3d")		         	; 20040404000000
	; e.g. MsgBox, % DateAddEx(20010101, "3y 3m 3d 8h 8m 8s")	; 20040404080808

	local
	static date := "date"

	vDate 	:= FormatTime(vDate, "yyyyMMddHHmmss")
	vMonth	:= SubStr(vDate, 5, 2)
	oTemp	:= StrSplit((RegExReplace(vDiff, "[^\d ]") . " 0 0 0 0 0"), " ")
	vDate  += (vMonth+oTemp.2 > 12) ? ((oTemp.2-12)*100000000 + (oTemp.1+1)*10000000000) : (oTemp.2*100000000 + oTemp.1*10000000000)

	Loop 3 {
		if vDate is %date%
			break
		vDate -= 1000000
	}

	;~ SciTEOutput("3: " vDate ", " oTemp.3*86400+oTemp.4*3600+oTemp.5*60+oTemp.6)

return AddOrSub = "add" ? DateAdd(vDate, oTemp.3*86400+oTemp.4*3600+oTemp.5*60+oTemp.6, "S") : DateSub(vDate, oTemp.3*86400+oTemp.4*3600+oTemp.5*60+oTemp.6, "S")
}

DateAdd(DateTime, Time, TimeUnits){                                                                	;-- Zeit zu einem Datum addieren
	EnvAdd DateTime, % Time, % TimeUnits
return DateTime
}

DateSub(DateTime, Time, TimeUnits){                                                                    	;-- Zeit von einem Datum abziehen
	EnvSub DateTime, % Time, % TimeUnits
return DateTime
}

Feiertage(jahr,land,timestrg) {                                                                               	;-- Feiertage berechnen

	/* Beschreibung

		;https://autohotkey.com/board/topic/97400-feiertage-berechnen-osterformel-nach-heiner-lichtenberg/
		;https://www.excel-coach.com/excel-und-die-feiertage/
		;https://www.dagmar-mueller.de/wdz/html/feiertagsberechnung.html
		;https://www.ferienwiki.de/feiertage/2019/de
		; ==============
		;timestrg:="yyyyMMdd"
		;timestrg:="longdate"

		;Als Test alle Feiertage in Deutschland mit Benennung
		msgbox, % feiertage("2019","dl","longdate")

		;Immer-Sonntage sind nicht dabei
		msgbox, % feiertage("2019","ni","yyyyMMdd")

		return

	 */

	; Ostersonntag
	ttos:="`nOstersonntag:                 "
		X := jahr                                              ; Jahreszahl, für die Ostern berechnet werden soll
		K  := Floor(X/100)                                      ; Säkularzahl
		M  := 15+Floor((3*K+3)/4)-Floor((8*K+13)/25)            ; säkulare Mondschaltung
		S  := 2-Floor((3*K+3)/4)                                ; säkulare Sonnenschaltung
		A  := Mod(X,19)                                         ; Mondparameter
		D  := Mod(19*A+M,30)                                    ; Keim für den ersten Vollmond im Frühling
		R  := Floor(D/29)+(Floor(D/28)-Floor(D/29))*Floor(A/11) ; kalendarische Korrekturgröße (beseitigt die Gaußschen Ausnahmeregeln)
		OG := 21+D-R                                            ; Ostergrenze (Märzdatum des Ostervollmonds)
		SZ := 7-Mod(X+Floor(X/4)+S,7)                           ; Erster Sonntag im März
		OE := 7-Mod(OG-SZ,7)                                    ; Entfernung des Ostersonntag von der Ostergrenze in Tagen
		OS := OG+OE                                             ; Märzdatum (ggf. in den April verlängert) des Ostersonntag, (32. März = 1. April usw.)
		Os:= x . (OS > 31 ? "04" SubStr("0"OS-31,-1) : "03" OS)
		FormatTime, oss,% os, % timestrg
	 ; ==============
	 ;Buß- und Bettag                                           ;Mittwoch vor dem 23. November
	  tbbt:="`nBuß- und Bettag:             "
	  md := OG+OE
	  tOs:= md > 31 ? SubStr("0"md-31,-1) :  md                   ;tag des Ostersonntags für die Berechnung des Buß und Bettags https://www.dagmar-mueller.de/wdz/html/feiertagsberechnung.html
	  tz:= md > 31 ?  Mod((30 - tos), 7)  :  Mod((33 - tos), 7)
	  bbt:="20191122"
	  bbt += -tz, days
	  FormatTime, bbt,% bbt, % timestrg
	 ; ==============
	 ; NeuJahr
	  tnj:="`nNeuJahr:                           "
	  nj:=  jahr . "0101"
	  FormatTime, nj,% nj, % timestrg
	 ; ==============
	 ; Heilige Drei Könige
	   thdk:="`nHeilige Drei Könige:        "
	   hdk:=  jahr . "0106"
	   FormatTime, hdk,% hdk, % timestrg
	 ; ==============
	 ; Internationaler Frauentag (nur Berlin)
	   tift:="`nInternationaler Frauentag:         "
	   ift:= jahr . "0308"
	   FormatTime, ift,% ift, % timestrg
	 ; ==============
	  ; Tag der Arbeit
	   ttda:="`nTag der Arbeit:                "
	   tda:=  jahr . "0501"
	   FormatTime, tda,% tda, % timestrg
	 ; ==============
	 ;  Augsburger Friedensfest
		taff:="`nAugsburger Friedensfest:   "
		aff:=  jahr . "0808"
		FormatTime, aff,% aff, % timestrg
	 ; ==============
	 ; Mariä Himmelfahrt
	   tmhf:="`nMariä Himmelfahrt:            "
	   mhf:=  jahr . "0815"
	   FormatTime, mhf,% mhf, % timestrg
	 ; ==============
	 ; Tag der Deutschen Einheit
	   ttdde:="`nTag der Deutschen Einheit:   "
	   tdde:=  jahr . "1003"
	   FormatTime, tdde,% tdde, % timestrg
	 ; ==============
	 ; Reformationstag
	   trt:="`nReformationstag:           "
	   rt:=  jahr . "1031"
	   FormatTime, rt,% rt, % timestrg
	 ; ==============
	 ; Allerheiligen
	   tah:="`nAllerheiligen:                "
	   ah:=  jahr . "1101"
	   FormatTime, ah,% ah, % timestrg
	 ; ==============
	 ; 1. Weihnachtsfeiertag
	   tw1:="`n1. Weihnachtsfeiertag:        "
	   w1:=  jahr . "1225"
	   FormatTime, w1,% w1, % timestrg
	 ; ==============
	 ; 1. Weihnachtsfeiertag
	   tw2:="`n1. Weihnachtsfeiertag:        "
	   w2:=  jahr . "1226"
	   FormatTime, w2,% w2, % timestrg
	 ; ==============
	; Karfreitag                      Freitag vor Ostersonntag                           = OS – 2
	tkfr:="`nKarfreitag:                        "
	  kfr:=os
	  kfr += -2, days
	  FormatTime, kfr,% kfr, % timestrg
	 ; ==============
	;Ostermontag                    Montag nach Ostersonntag                        = OS + 1
	  tosm:="`nOstermontag:                  "
	  osm:=os
	  osm += 1, days
	  FormatTime, osm,% osm, % timestrg
	 ; ==============
	;Christi Himmelfahrt          39 Tage nach Ostersonntag                       = OS + 39
	  tchf:="`nChristi Himmelfahrt:       "
	  chf:=os
	  chf += 39, days
	  FormatTime, chf,% chf, % timestrg
	 ; ==============
	;Pfingstsonntag                49 Tage nach Ostersonntag                        = OS + 49
	  tpfs:="`nPfingstsonntag:             "
	  pfs:=os
	  pfs += 49, days
	  FormatTime, pfs,% pfs, % timestrg
	 ; ==============
	;Pfingstmontag                 50 Tage nach Ostersonntag                        = OS + 50
	  tpfm:="`nPfingstmontag:              "
	  pfm:=os
	  pfm += 50, days
	  FormatTime, pfm,% pfm, % timestrg
	 ; ==============
	;Fronleichnam                   60 Tage nach Ostersonntag                       = OS + 60
	  tfl:="`nFronleichnam:               "
	  fl:=os
	  fl += 60, days
	  FormatTime, fl,% fl, % timestrg
	 ; ==============

	 ; alle (dl)
	  ;ft:= nj "," hdk "," ift "," kfr "," os  "," osm "," tda "," chf "," pfs  "," pfm "," fl "," aff "," mhf "," tdde "," rt "," ah "," bbt ","  w1 "," w2
	   ;ft:= tnj nj "," thdk hdk "," tift ift "," tkfr kfr "," ttos oss  "," tosm osm "," ttda tda "," tchf chf "," tpfs pfs  "," tpfm pfm "," tfl fl "," taff aff "," tmhf mhf "," ttdde tdde "," trt rt "," tah ah "," tbbt bbt ","  tw1 w1 "," tw2 w2
	   ;return ft

	if (land = "dl") ;test alle
	  ft:= tnj nj "," thdk hdk "," tift ift "," tkfr kfr "," ttos oss  "," tosm osm "," ttda tda "," tchf chf "," tpfs pfs  "," tpfm pfm "," tfl fl "," taff aff "," tmhf mhf "," ttdde tdde "," trt rt "," tah ah "," tbbt bbt ","  tw1 w1 "," tw2 w2
	else if (land = "bw") ;Feiertage Baden-Württemberg
	  ft:= nj "," hdk "," kfr "," osm "," tda "," chf "," pfm "," fl  "," tdde  "," ah  ","  w1 "," w2
	else if  (land = "by") ;Feiertage Bayern alle
	  ft:= nj "," hdk "," kfr "," osm "," tda "," chf "," pfm "," fl "," aff "," mhf "," tdde "," rt "," ah "," bbt ","  w1 "," w2
	else if  (land = "bya") ;Feiertage Bayern alle
	  ft:= nj "," hdk "," kfr "," osm "," tda "," chf "," pfm "," fl "," aff "," mhf "," tdde "," rt "," ah "," bbt ","  w1 "," w2
	else if  (land = "byk") ;Feiertage Bayern kath.
	  ft:= nj "," hdk "," kfr "," osm "," tda "," chf "," pfm "," fl "," mhf "," tdde "," rt "," ah "," bbt ","  w1 "," w2
	else if  (land = "byp") ;Feiertage Bayern prot.
	  ft:= nj "," hdk "," kfr "," osm "," tda "," chf "," pfm "," fl  "," tdde "," rt "," ah "," bbt ","  w1 "," w2
	else if  (land = "be") ;Feiertage Berlin
	  ft:= nj "," ift "," kfr "," osm "," tda "," chf "," pfm "," tdde ","  w1 "," w2
	else if  (land = "bb") ;Feiertage Brandenburg
	  ft:= nj "," kfr "," osm "," tda "," chf "," pfm "," tdde "," rt ","  w1 "," w2
	else if  (land = "hb") ;Feiertage Bremen
	  ft:= nj "," kfr "," osm "," tda "," chf "," pfm "," tdde "," rt ","  w1 "," w2
	else if  (land = "hh") ;Feiertage Hamburg
	  ft:= nj "," kfr "," osm "," tda "," chf "," pfm "," tdde "," rt ","  w1 "," w2
	else if  (land = "he") ;Feiertage Hessen
	  ft:= nj "," kfr "," osm "," tda "," chf "," pfm "," fl "," tdde ","  w1 "," w2
	else if  (land = "mv") ;Feiertage Mecklenburg-Vorpommern
	  ft:= nj "," kfr "," osm "," tda "," chf "," pfm "," tdde "," rt ","  w1 "," w2
	else if  (land = "ni") ;Feiertage Niedersachsen
	  ft:= nj "," kfr "," osm "," tda "," chf "," pfm "," tdde "," rt ","  w1 "," w2
	else if  (land = "nw") ;Feiertage Nordrhein-Westfalen
	   ft:= nj "," kfr "," osm "," tda "," chf "," pfm "," fl "," tdde "," ah ","  w1 "," w2
	else if  (land = "rp") ;Feiertage Rheinland-Pfalz
	  ft:= nj "," kfr "," osm "," tda "," chf "," pfm "," fl "," tdde "," ah ","  w1 "," w2
	else if  (land = "sl") ;Feiertage Saarland
	  ft:= nj "," kfr "," osm "," tda "," chf "," pfm "," fl "," mhf "," tdde "," ah ","  w1 "," w2
	else if  (land = "sn") ;Feiertage Sachsen
	  ft:= nj "," hdk "," kfr "," osm "," tda "," chf "," pfm "," tdde "," rt ","  w1 "," w2
	else if  (land = "st") ;Feiertage Sachsen-Anhalt
	  ft:= nj "," hdk "," kfr "," osm "," tda "," chf "," pfm "," tdde "," rt ","  w1 "," w2
	else if  (land = "sh") ;Feiertage Schleswig-Holstein
	  ft:= nj "," kfr "," osm "," tda "," chf "," pfm "," tdde "," rt ","  w1 "," w2
	else if  (land = "th") ;Feiertage Thüringen
	  ft:= nj "," kfr "," osm "," tda "," chf "," pfm "," tdde "," rt ","  w1 "," w2

 return ft
}


; Zeit berechnen
Clock() {                                                                                                        		;-- genaue Zeit in Millisekunden
return DllCall("msvcrt.dll\clock")
}

FormatSeconds(timestr, formatstring:="hh:mm:ss")  {                                          	;-- Sekunden in Stunden:Minuten:Sekunden umrechnen
    atime := A_YYYY A_MM A_DD "000000"
    atime += timestr, seconds
    FormatTime, hhmmss, % atime, % formatstring
return hhmmss
}

FormatTime(YYYYMMDDHH24MISS:="", Format:="") {                                        	;-- FormatTime wrapper
	local OutVar
	FormatTime OutVar, % YYYYMMDDHH24MISS, % Format
return OutVar
}

GetSeconds(timestr) {                                                                                        	;-- Sekunden berechnen von hhmmss
	; timestr format: hhmmss oder hh:mm:ss o. ....
	timestr := SubStr("000000" RegExReplace(timestr, "[^\d]"), -5)
return (SubStr(timestr,1,2)*3600)+(SubStr(timestr,3,2)*60)+SubStr(timestr,5,2)
}

TimerTime(TimeToStart) {                                                                                   	;-- berechnet die Millisekunden bis zu einer bestimmten Uhrzeit

	; berechnet die Millisekungen bis zu einer bestimmten Uhrzeit von der aktuellen Zeit
	; ist die Uhrzeit am aktuellen Tag schon vorbei, werden 24h auf die Startzeit gerechnet
	; TimeToStart Format: kann relativ frei geschrieben werden, zB. 16:10 Uhr, da Zeichen welche keine Zahl sind entfernt werden
	; es können Sekunden angegeben werden. Für die Berechnungen werden fehlende Sekunden als 00 ergänzt

	TimeNow  	:= GetSeconds(A_Hour A_Min A_Sec)
	TimeToStart 	:= RegExReplace(TimeToStart, "[^\d]")
	TimeToStart	:= SubStr(TimeToStart "0000", 1, 6	)	            	; fehlende Sekunden anhängen
	TimeToSet 	:= GetSeconds(TimeToStart)
	TimeToSet   	:= TimeToSet - TimeNow < 0 ? TimeToSet + GetSeconds("240000") : TimeToSet

return (TimeToSet - TimeNow) * 1000
}

TimeFormatEx(sec, ShowSeconds:=true) {                                                           	;-- Sekunden in hh:mm:ss

	H	:= SubStr("00" Floor(sec/3600), -1)
	M	:= SubStr("00" Floor((sec-(H*3600))/60), -1)
	S	:= SubStr("00" Floor(sec-(H*3600)-(M*60)), -1)

return H ":" M (ShowSeconds ? ":" S : "")
}

TimeDiff(time1, time2="now", output="m") {                                                       	;-- Zeitdifferenz zwischen Zeit1 und Zeit2 (max. ein Tag Differenz!)

	; Ausgabe in Minuten (output="m") oder Sekunden (output="s")

	if Instr(time2, "now") {
		time2 := A_Now
		FormatTime, time2,, HHmmss
	}

	time1 := Instr(time1, "000000") ? "240000" : time1

	if Instr(output, "m")
		return (SubStr(time1, 1, 2)*60 + SubStr(time1, 3, 2)) - (SubStr(time2, 1, 2)*60 + SubStr(time2, 3, 2))
	else if Instr(output,"s")
		return (SubStr(time1, 1, 2)*3600 + SubStr(time1, 3, 2) + SubStr(time1, 5, 2)) - ( SubStr(time2, 1, 2)*3600 + SubStr(time2, 3, 2)*60 + SubStr(time2, 5, 2))

return
}

GetTimestrings(ms, maxTime:="Auto") {                                                               	;-- Stunden, Minuten, Sekunden aus Millisekunden berechnen

	; maxTime = "Auto"

	If (maxTime >= 2 || maxTime = "Auto")  {
		hour	:= ms // 360000
		ms   	-= hour * 360000
	}
	If (maxTime >= 1 || maxTime = "Auto") {
		min	:= ms // 60000
		ms   	-= min * 60000
	}
	sec  	:= ms // 1000
	ms   	-= sec * 1000

return {"hour":SubStr("0" hour, -1), "min":SubStr("0" min, -1), "sec":SubStr("0" sec, -1), "msec":ms}
}


; Formatierung
FormatDate(timestr, timeformat:="DMY", returnformat:="dd.MM.yyyy") {

		static rxMDate := {"D"	: "(?<D>\d{1,2})"
								, 	"M"	: "(?<M>\d{1,2})"
								, 	"Y"	: "(?<Y>\d{1,4})"}

	; "timeformat" korrigieren
		timeformat := RegExReplace(timeformat, "[D]"	, "D")
		timeformat := RegExReplace(timeformat, "[M]"	, "M")
		timeformat := RegExReplace(timeformat, "[Y]" 	, "Y")

	; RegExMatch String anhand timeformat zusammenstellen
		TF	:= Object()
		TF[InStr(timeformat, "D")]	:= "D"
		TF[InStr(timeformat, "M")]	:= "M"
		TF[InStr(timeformat, "Y")]	:= "Y"

		rxMatch := rxMDate[TF.1] ".*?" rxMDate[TF.2] ".*?" rxMDate[TF.3]

	; fehlerhaften "timestr" verwerfen
		If !RegExMatch(timestr, rxMatch, T) || (StrLen(TY) = 1) || (StrLen(TY) = 3)
			return  ; "wrong timeformat"

	; Anzahl der Ziffern des Jahres im returnformat berichtigen
		RegExReplace(returnformat, "i)y", "", YearDigits)
		If (YearDigits != 4)
			returnformat := RegExReplace(returnformat, "i)y+", "yyyy")

	; Eingabedatum formatieren
		TD 	:= SubStr("0" TD	, -1)
		TM 	:= SubStr("0" TM	, -1)
		If (StrLen(TY) = 2) {   ; wendet sich gegen Datumszahlen in der Zukunft!
			YD	:= SubStr(A_YYYY, 3, 2) + 0
			YT	:= SubStr(A_YYYY, 1, 2) + 0
			TY 	:= (TY > YD ? YT-1 : YT) . TY
		}

	; Zeitstring formatieren
		FormatTime, formatedTime, % TY TM TD "000000", % returnformat

return formatedTime
}

FormatDateEx(datestr, dateformat:="DMY", returnformat:="dd.MM.yyyy") {          	;-- formatiert ein Datum in der gewünschten Folge samt Interpunktion

		static rxMDate := {"D"	: "(?<D>\d{1,2})"
								, 	"M"	: "(?<M>\d{1,2})"
								, 	"Y"	: "(?<Y>\d{1,4})"}

	; Delimeter erkennen z.B. ein Punkt oder minus
		RegExMatch(datestr, "O)\w+(\D)\w+(\D)\w", dmtr)
		For idx, interpunctation in dmtr
			datestr := StrReplace(datestr, interpunctuation)

	; dateformat korrigieren
		dateformat := RegExReplace(dateformat, "i)[D]"	, "D")
		dateformat := RegExReplace(dateformat, "i)[M]"	, "M")
		dateformat := RegExReplace(dateformat, "i)[Y]" 	, "Y")

	; RegExMatch String anhand dateformat zusammenstellen
		TF	:= Object()
		TF[InStr(dateformat, "D")]	:= "D"
		TF[InStr(dateformat, "M")]	:= "M"
		TF[InStr(dateformat, "Y")]	:= "Y"

		rxMatch := rxMDate[TF.1] ".*?" rxMDate[TF.2] ".*?" rxMDate[TF.3]

	; inkorrekte Jahreszahl - umwandeln abbrechen
		If !RegExMatch(datestr, rxMatch, T) || (StrLen(TY) = 1) || (StrLen(TY) = 3)
			return  ; "wrong dateformat"

	; Anzahl der Ziffern des Jahres in returnformat berichtigen
		RegExReplace(returnformat, "i)y", "", YearDigits)
		If !(YearDigits = 4)
			returnformat := RegExReplace(returnformat, "i)y+", "yyyy")

	; Eingabedatum formatieren
		TD 	:= SubStr("0" TD	, -1)
		TM 	:= SubStr("0" TM	, -1)
		If (StrLen(TY) = 2) {   ; wendet sich gegen Datumszahlen in der Zukunft!
			YD	:= SubStr(A_YYYY, 3, 2) + 0
			YT	:= SubStr(A_YYYY, 1, 2) + 0
			TY 	:= (TY > YD ? YT-1 : YT) . TY
		}

	; Zeitstring formatieren
		FormatTime, formatedDate, % TY TM TD "000000", % returnformat

return formatedDate
}

ShrinkTwoDatesStr(TwoDates) {                                                                          	;-- verkürzt ein von-bis-Datum logisch um Zeichen zu sparen

	; TwoDates 	-	muss ein String mit zwei durch ein Minus (-) getrennte Datumzahlen
	;						die beiden Datumzahlen müssen folgendes Format haben: dd.MM.yyyy
	;						aus 01.01.2022-09.01.2022 wird 01.-09.01.22
	;						aus 27.01.2022-12.02.2022 wird 27.01.-12.02.22
	;						aus 19.12.2021-03.01.2022 wird 19.12.21-03.01.22

	TwoDates := StrSplit(TwoDates, "-")
	day1 := StrSplit(TwoDates[1], ".")
	day2 := StrSplit(TwoDates[2], ".")
	return (	day1.3 <> day2.3 ? 	day1.1 "." day1.2 "." SubStr(day1.3, -1)
			:	day1.2 <> day2.2 ? 	day1.1 "." day1.2 :	day1.1 ".") "-" day2.1 "." day2.2 "." SubStr(day2.3, -1)
}

GetDateFormat(dateString) {                                                                               	;-- in welcher Reihenfolge liegen Jahr, Monat, Tag vor

	; Voraussetzung ist ein 8 stelliges Datum aus einem zweistelligen Tag und Monat plus einem vierstelligen Jahr

	If (StrLen(dateString) != 8 || RegExMatch(dateString, "[^\d]"))
		return

	RegExMatch(dateString, "^(?<dd>\d{2})(?<MM>\d{2})(?<yyyy>\d{4})$", t)
	RegExMatch(dateString, "^(?<yyyy>\d{4})(?<MM>\d{2})(?<dd>\d{2})$", y)
	dmismatch 	:= (tMM 	>	12) || (tdd	> 31) 	? true : false
	ymismatch	:= (yMM 	> 12) || (ydd	> 31)	? true : false

return (!dmismatch ? "DDMMYYYY" : !ymismatch ? "YYYYMMDD" : "") ; leer wenn weder das eine noch andere passt
}


; Konvertierung
ConvertGerDateToEng(dateStr) {                                                                         	;-- konvertiert vom deutschen ins englische Datumsformat
return SubStr(dateStr, 7, 4) . SubStr(dateStr, 4, 2) . SubStr(dateStr, 1, 2)
}

ConvertDBASEDate(DBASEDate) {                                                                        	;-- Datumskonvertierung von YYYYMMDD nach DD.MM.YYYY
	return SubStr(DBaseDate, 7, 2) "." SubStr(DBaseDate, 5, 2) "." SubStr(DBaseDate, 1, 4)
}

ConvertToDBASEDate(Date) {                                                                             	;-- Datumskonvertierung von DD.MM.YYYY nach YYYYMMDD
	RegExMatch(Date, "((?<Y1>\d{4})|(?<D1>\d{1,2})).(?<M>\d+).((?<Y2>\d{4})|(?<D2>\d{1,2}))", t)
return (tY1?tY1:tY2) . SubStr("00" tM, -1) . SubStr("00" (tD1?tD1:tD2), -1)
}


; Stempel
datestamp(nr:=1) {                                                                                            	;-- für Protokolle
	If (nr = 1)
		return (A_YYYY "-" A_MM "-" A_DD " " A_Hour ":" A_Min ":" A_Min)
	else
		return (A_Hour ":" A_Min ":" A_Sec)
}

TimeCode(DaT) {                                                                                                 	;-- für Addendum_Protokoll - true für [YYYY.MM.DD,] hh:mm:ss.ms
	TC := (DaT ? (A_YYYY "." A_MM "." A_DD ", ") : ("")) . A_Hour ":" A_Min ":" A_Sec "." A_MSec
return TC
}


; sonstiges
GetFileTime(filepath, WhichTime:="C", formatstring:="") {                                    	;-- formatiertes Datum aus der Dateierstellungszeit

	FileGetTime, FTIME , % filepath, % WhichTime
	FormatTime, RTIME, % FTIME, % (formatstring ? formatstring : "dd.MM.yyyy")

return RTIME
}




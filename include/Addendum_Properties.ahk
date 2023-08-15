#NoEnv
#Persistent
SetBatchLines, -1

AddendumProperties()

return

AddendumProperties(adm:="") {

	global

	;~ static apTabs := "Allgemeines|Addendum|Preise|AutoStartAbfrage|LaborAbruf|ScanPool|OCR|AutoDoc|Abrechnungshelfer|InfoFenster|QuickSearch|Laborjournal|AddendumMonitor|Telegram"
	static apTabs := "Allgemeines|Addendum|Labor|Befunde|Preise&Gebühren|Abrechnungshelfer|InfoFenster|Telegram"

	Gui, aP: new, HWNDhaP
	;~ Gui, aP: Color		, cWhite, cLightGrey
	Gui, aP: Margin	, 5, 5
	Gui, aP: Font     	, s10 q5 Normal cBlack, Arial
	Gui, aP: Add      	, Tab     	, % "x0 y0  w1000 h800 HWNDaPhTab vaPTabs gapGui_TabsHandler"       	, % apTabs


	; Tab Allgemeines
	Gui, aP: Tab  		, 1
	; PraxisNameStrassePLZOrtEMail1EMail2Telefon1Telefon1Fax1KVStempelMailStempelSprechstundeUrlaubArzt1NameArzt1LANRArzt1BSNRArzt1AbrechnungStandardArzt
	Gui, ap: Add     	, Progress	, % "x10 y40 w390 h150 vaPPGPraxis cGrey "
	Gui, ap: Add     	, GroupBox, % "x10 y40 w390 h150 vaPGBPraxis Backgroundtrans"	, % "Praxisdaten"
	Gui, ap: Add     	, Text	, % "x75 y70 w300 Center"	, % "Tragen Sie hier die Daten für ihre Praxis ein."
	Gui, ap: Add     	, Text	, % "x15 y+5 w60 Right"	, % "Name"
	Gui, ap: Add     	, Edit 	, % "x+5 w300 r1" 	, % Addendum.Praxis.Name
	Gui, ap: Add     	, Text	, % "x15 y+5 w60 Right"	, % "Straße"
	Gui, ap: Add     	, Edit 	, % "x+5 w300 r1" 	, % Addendum.Praxis.Strasse
	Gui, ap: Add     	, Text	, % "x15 y+5 w60 Right"	, % "PlZ"
	Gui, ap: Add     	, Edit 	, % "x+5 w60 r1"    	, % Addendum.Praxis.PLZ
	Gui, ap: Add     	, Text	, % "x+10 w20 Right"	, % "Ort"
	Gui, ap: Add     	, Edit 	, % "x+5 w205 r1" 	, % Addendum.Praxis.Ort



	Gui, aP: Show, AutoSize, % "Einstellungen - Addendum (" AddendumVersion " vom " DatumVom ")"

	return


}

apGui_TabsHandler(params*) {

}

aPGui_Close() {

	apGuiClose:
	apGuiEscape:
	ExitApp

}
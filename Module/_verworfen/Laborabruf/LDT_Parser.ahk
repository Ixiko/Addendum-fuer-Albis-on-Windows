;------- LDT Parser by Ixiko 2020
; Es darf nur der Zeichencode ISO 8859-15 verwendet werden. Cp 850 sieht geeignet aus.
/*
Feldteil 			Länge 			Bedeutung
Länge			3 Bytes			Angabe der Feldlänge
Kennung		4 Bytes			Feldkennung
Inhalt			Variabel		Daten
Ende				2 Bytes			Wert 13 = CR (Wagenrücklauf), gefolgt von Wert 10 = LF (Zeilenvorschub), dargestellt im Code ISO/IEC 6429
Für die Längenberechnung eines Feldes gilt die Regel:
Länge des Feldteils “Inhalt” + 9

*/

LDTKonstanten() {

	; LDT 1014.01
	LDT_FeldkennungCode := {
		"0101"	: "KBV-Prüfnummer",	                      	; KBV-Prüfnummer
		"0201"	: "BSNR",                                     		; Betriebs- (BSNR) oder Nebenbetriebsstätten-nummer (NBSNR)
		"0203"	: "NBSNR",                                  		; (N)BSNR-Bezeichnung
		"0205"	: "Straße",                                   		; Strasse der (N)BSNR
		"0211"	: "Arztname",                                 	;
		"0212"	: "LANR":,                                      	; Lebenslange Arztnummer (LANR)
		"0215"	: "PLZ":,                                          	; PLZ der (N)BSNR
		"0216"	: "Ort":,                                          	; Ort der (N)BSNR
		"3000"	: "Versichertenstatus",                     	; Versichertenstatus
		"3101"	: "Nachname",                               	; ! Nachname des Patienten !
		"3102"	: "Vorname",                                  	; ! Vorname des Patienten !
		"3103"	:"Geburtsdatum",                           	; ! Geburtsdatum des Patienten !
		"3110"	: "Geschlecht",                               	; ! Geschlecht des Patienten !
		"5001"	: "Gebührennummer",                      	;
		"8000"	: "Satzart",                                     	;
		"8100"	: "Satzlaenge",                               	;
		"8300"	: "Labor",                                    		;
		"8301"	: "AuftragsEingangsdatum",	            ; Eingangsdatum des Auftrags im Labor
		"8302"	: "Berichtsdatum",                        		; Berichtsdatum des Auftrages im Labor
		"8303"	: "Berichtszeit",                               	;
		"8310"	: "Anforderung-Ident",                 		; ! Nummer die von der Praxis zum Labor geht !
		"8311"	: "LaborAuftragsnummer",               	; Auftragsnummer des Labors
		"8312"	: "KundenNummer",                        	; Kunden- (Arzt-) Nummer
		"8320"	: "Laborname",                              	;
		"8321"	: "LaborStrasse",                               	; Strasse der Laboradresse
		"8322"	: "LaborPLZ",                       				; PLZ der Laboradresse
		"8323"	: "LaborOrt", 										; Ort der Laboradresse
		"8401"	: "Befundart",                 					; oder Status, dabei steht E bei meinem Labor für Endbefund
		"8403"	: "Gebührenordnung",                 		;
		"8406"	: "Kosten", 										; Kosten in cent
		"8410"	: "Test-Ident",										; oder Laborparameter (Kurzbezeichnung)
		"8411"	: "Testbezeichnung",                 			; Langbezeichnung des Laborparameter
		"8615"	: "Auftraggeber",                             	; ?obsolet? - da steht meine LANR
		"8420"	: "Ergebniswert",                 				; ! der Laborwert an sich !
		"8421"	: "Einheit",                 						; ! Laborparameter zugehörige Einheit !
		"8422"	: "Grenzwert-Indikator",                	 	;
		"8432"	: "Abnahme-Datum",                 			;
		"8433"	: "Abnahme-Zeit",                 				;
		"8460"	: "Normalwert-Text",                 			;
		"8461"	: "Normalwert-Untergrenze",        	 	; !
		"8462"	: "Normalwert-Obergrenze",           	; !
		"8470"	: "Testbezogene-Hinweise",            		;
		"8480"	: "Ergebnis-Text",                 				;
		"8609"	: "Abrechnungstyp",							; K = Kasse
		"8614"	: "Abrechnungsverantwortlichkeit",		; die Abrechnungsverantwortlichkeit zwischen Labor und Einweiser regelt (FK 8614).
		"8615"	: "Auftraggeber",                 				;
		"9103"	: "Erstellungsdatum",                 			;
		"9106"	: "Zeichensatz",									; verwendeter Zeichensatz
		"9202"	: "LDatenPaket",							    	; Gesamtlänge des Datenpaketes
		"9212"	: "LDTVersion"                 					;
		}

return LDT_FeldkennungCode
}


LDT2CSV(file) {

	static LDTKonstante

	labor := Object()

	If !IsObject(LDTKonstante)
		LDTKonstante := LDTKonstanten()

	;FileEncoding, cpp8859								;ISO 8859-15 Latin 9
	ldtfile:= Substr(file, 1, StrLen(file)-3) . "csv"

	If FileExist("Test.csv") {
			FileDelete, Test.csv
	}

	LDTDaten := FileOpen(file, "r", "CP28605").Read()

	Loop, Parse, LDTDaten, `n, `r
	{
			Laenge     	:= SubStr(A_LoopField, 1, 3)
			FeldKennung	:= SubStr(A_LoopField, 4, 4)
			Inhalt        	:= SubStr(A_LoopField, -8)

			If (FeldKennung == "8310")
				Auftragsnummer := Inhalt
			else if (FeldKennung == "8301")
				Abnahmedatum  := Inhalt
			else if (FeldKennung == "8302")
				Abnahmezeit  	  := Inhalt
			else if (FeldKennung == "3101")
				Nachname    	  := Inhalt
			else if (FeldKennung == "3102")
				Vorname       	  := Inhalt
			else if (FeldKennung == "3103")
				Geburtsdatum  	  := Inhalt
			else if (FeldKennung == "3110")
				Geschlecht       	  := Inhalt


	}

}


; LDT 1014.01
global LDT_FeldkennungCode:={
"0101"	: "KBV-Prüfnummer",	                      	; KBV-Prüfnummer
"0201"	: "BSNR",                                     		; Betriebs- (BSNR) oder Nebenbetriebsstätten-nummer (NBSNR)
"0203"	: "NBSNR",                                  		; (N)BSNR-Bezeichnung
"0205"	: "Straße",                                   		; Strasse der (N)BSNR
"0211"	: "Arztname",                                 	;
"0212"	: "LANR":,                                      	; Lebenslange Arztnummer (LANR)
"0215"	: "PLZ":,                                          	; PLZ der (N)BSNR
"0216"	: "Ort":,                                          	; Ort der (N)BSNR
"3000"	: "Versichertenstatus",                     	; Versichertenstatus
"3101"	: "Nachname",                               	; Nachname des Patienten
"3102"	: "Vorname",                                  	; Vorname des Patienten
"3103"	:"Geburtsdatum",                           	; Geburtsdatum des Patienten
"3110"	: "PatGeschlecht",                           	; Geschlecht des Patienten
"5001"	: "Gebührennummer",                      	;
"8000"	: "Satzart",                                     	;
"8100"	: "Satzlaenge",                               	;
"8300"	: "Labor",                                    		;
"8301"	: "AuftragsEingangsdatum",	            ; Eingangsdatum des Auftrags im Labor
"8302"	: "Berichtsdatum",                        		;
"8303"	: "Berichtszeit",                               	;
"8310"	: "Anforderung-Ident",                 		;
"8311"	: "LaborAuftragsnummer",               	; Auftragsnummer des Labors
"8312"	: "KundenNummer",                        	; Kunden- (Arzt-) Nummer
"8320"	: "Laborname",                              	;
"8321"	: "LaborStrasse",                               	; Strasse der Laboradresse
"8322"	: "LaborPLZ",                       				; PLZ der Laboradresse
"8323"	: "LaborOrt", 										; Ort der Laboradresse
"8401"	: "Befundart",                 					;
"8403"	: "Gebührenordnung",                 		;
"8406"	: "Kosten", 										; Kosten in cent
"8410"	: "Tes-Ident",										; oder Laborparameter
"8411"	: "Testbezeichnung",                 			;
"8420"	: "Ergebniswert",                 				;
"8421"	: "Einheit",                 						;
"8422"	: "Grenzwert-Indikator",                	 	;
"8432"	: "Abnahme-Datum",                 			;
"8433"	: "Abnahme-Zeit",                 				;
"8460"	: "Normalwert-Text",                 			;
"8461"	: "Normalwert-Untergrenze",        	 	;
"8462"	: "Normalwert-Obergrenze",           	;
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
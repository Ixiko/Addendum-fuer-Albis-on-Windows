; ----------------------------- Ausschluß von bestimmten Patientennummern (PatID)
	global excludeIDs
; ----------------------------- Worttrennzeichen
	global rx, ry, rz, rs, rp1, rp2
; ----------------------------- Datum
	global ThousandYear, HundredYear, rxDay, rxWDay, rxWDay2, rxMonths, rxMonths2, rxYears 	; Teile des Datum
	global rxDatum, rxDates                                                                                                         	; Komplett
	global rxGebAm, rxGebAm2                                                                                                    	; Startstrings
; ----------------------------- Eigen- oder Personennamen
	global rxPerson, rxPerson1, rxPerson2, rxPerson3, rxPerson4, rxName1, rxName2                      	; Teile
	global RxNames, RxNames2                                                                                                  	; Komplett
	global GName, rxAnrede, rxAnredeM, rxAnredeF, rxVorname, rxNachname, rxNameW				; Startstrings
; ----------------------------- sonstige
	global rxContinue
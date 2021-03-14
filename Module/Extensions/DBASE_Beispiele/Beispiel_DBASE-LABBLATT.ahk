
;	Beispielskript für die Verwendung von Addendum_DBASE.ahk
;	sucht nach einem Parameter (COV**) und zählt positive und negative Befunde

	#NoEnv
	SetBatchLines, -1

	; Skriptstartzeit
		starttime	:= A_TickCount

	; Albis-Datenbank Pfad
		RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)        ; Addendum-Verzeichnis
		RegRead, PathAlbis, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-MainPath
		basedir    	:= PathAlbis "\db"
	;-------------------------------------------------------------------------------------------------------------------------------------------
	; neues Datenbankobjekt anlegen. new DBASE(Datenbankpfad [, debug])
	; Datenbankpfad	- vollständiger Pfad zur Datenbank in der Form M:\albiswin\db\BEFUND.dbf
	; debug                	- Ausgabe von Werten zur Kontrolle des Ablaufes, Voreinstellung: keine Ausgabe
	;-------------------------------------------------------------------------------------------------------------------------------------------
		labblatt   	:= new DBASE(basedir "\labblatt.dbf", 4)
		res        	:= labblatt.OpenDBF()
		;matches 	:= labblatt.SearchExt({"AUSDATUM":"rx:202103(08|09|10|11|12|13)"}, ["ANFNR", "AUSDATUM", "PARAMBEZ"])
		matches 	:= labblatt.SearchExt({"PARAM":"rx:COV[A-Z\d-]+"}  ; , "ERG":"POSITIV" ; "DATUM":">20200101",
													,  	 ["ANFNR", "EINDATLDT3", "PARAM", "BERICHT", "ERGEBNIS", "BEFUND", "DATUM", "BERZEIT"],,, "MyToolTip")
		;~ matches 	:= labblatt.SearchExt({"PARAM":"rx:HEPE[A-Z\d-]+"}  ; , "ERG":"POSITIV" ; "DATUM":">20200101",
													;~ ,  	 ["PATNR", "ANFNR", "PARAM", "ERGEBNIS", "BEFUND", "DATUM"],,, "MyToolTip")

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Treffer nach Anforderungsnummer und Parameterbezeichnung sortieren
	;-------------------------------------------------------------------------------------------------------------------------------------------
		results	  	:= Object()
		Covids  	:= Object()
		Covids.positiv  	:= Object()
		Covids.negativ 	:= Object()
		For idx, obj in matches {

			If !results.haskey(obj.ANFNR)
				results[obj.ANFNR] := Object()
			results[obj.ANFNR][obj.PARAM] := obj.ERGEBNIS "|" obj.PATNR "|" obj.DATUM

			If InStr(obj.ERGEBNIS, "positiv") {

				If !Covids.positiv.haskey(obj.ANFNR)
					Covids.positiv[obj.ANFNR] := obj.ERGEBNIS "|" obj.PATNR "|" obj.DATUM

				If Covids.negativ.haskey(obj.ANFNR)
					Covids.negativ.RemoveAt(obj.ANFNR)
			}
			else {
				If !Covids.negativ.haskey(obj.ANFNR) && !Covids.positiv.haskey(obj.ANFNR)
					Covids.negativ[obj.ANFNR] := obj.ERGEBNIS "|" obj.PATNR "|" obj.DATUM
			}

		}

		FileOpen(A_Temp "\labblatt.json", "w", "UTF-8").Write(JSON.Dump(results,,1,"UTF-8"))
		FileOpen(A_Temp "\covids.json", "w", "UTF-8").Write(JSON.Dump(Covids,,1,"UTF-8"))

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; 2.Zeit - Dauer des Zählens
	;-------------------------------------------------------------------------------------------------------------------------------------------
		objSize1 	:= labblatt.GetCapacity()
		duration	:= A_TickCount - starttime

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Ausgabe in Scite
	;-------------------------------------------------------------------------------------------------------------------------------------------
		t := " records count:         `t"   	labblatt.records                                                	"`n"
		t .= " entrys found:            `t" 	matches.Count()                                            	"`n"
		t .= " Duration all:             `t" 	Round(duration  / 1000, 2) 	" seconds"            	"`n"
		t .= " records found/sec     `t"    Round(matches.Count()/(duration/1000))           	"`n"
		t .= " rec's searched/sec     `t"  	Round(labblatt.records/(duration/1000))           	"`n"
		t .= " ---------------------------------------------------------------------------------"       	"`n"
		t .= " Patients CoV-SARS2  "                                                                           	"`n"
		t .= " positiv:                    `t"  	Covids.positiv.Count()                                       	"`n"
		t .= " negativ:                   `t"  	Covids.negativ.Count()                                      	"`n"
		t .= " tests done:               `t"  	Covids.negativ.Count() + Covids.positiv.Count() 	"`n"


		SciTEOutput(RTrim(t, "`n"))

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Datenbank Objekt freigegeben. Damit werden alle gespeicherten Daten gelöscht und Speicher freigegeben
	;-------------------------------------------------------------------------------------------------------------------------------------------
		res          	:= labblatt.CloseDBF()
		labblatt	:= ""

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; extrahierte Daten ansehen
	;-------------------------------------------------------------------------------------------------------------------------------------------
		Run, % A_Temp "\labblatt.json"
		Run, % A_Temp "\COVIDS.json"

ExitApp

MyToolTip(p*) {

	ToolTip, % p.1 "/" p[2] " [" p[3] "]" , 500, 100, 13

}

LaborParameter() {

	labparam 	:= new DBASE(basedir "\LABPARAM.dbf", 4)
	res        	:= labparam.OpenDBF()
	matches 	:= labparam.GetFields("GetAll")

	PARAMs := ""
	Spaces := "                                                                                                                                                                   "
	For idx, obj in matches
		PARAMs .= 	obj.NAME 	SubStr(Spaces, 1, Floor(23-StrLen(obj.NAME)*1.3)) 	"`t|`t"
					 .  	obj.BEZEICH SubStr(Spaces, 1, Floor(71-StrLen(obj.BEZEICH)*1.3))	"`t|`t"
					 .  	obj.EINHEIT "`n"

	Sort, PARAMs, U

	FileOpen(A_Temp "\labparam.txt", "w", "UTF-8").Write(PARAMs)
	res        	:= labparam.CloseDBF()
	;Run, % A_Temp "\labparam.txt"

}

#Include %A_ScriptDir%\..\..\..\Include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\..\lib\SciTEOutput.ahk
#Include %A_ScriptDir%\..\..\..\lib\class_JSON.ahk

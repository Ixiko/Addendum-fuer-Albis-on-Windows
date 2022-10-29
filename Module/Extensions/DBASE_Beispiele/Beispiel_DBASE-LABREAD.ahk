
;	Beispielskript für die Verwendung von Addendum_DBASE.ahk
;	konvertiert die komplette Datenbank in einen JSON String

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
		labread   	:= new DBASE(basedir "\LABREAD.dbf", 4)
		res        	:= labread.OpenDBF()
		;matches 	:= labread.SearchExt({"AUSDATUM":"rx:202103(08|09|10|11|12|13)"}, ["ANFNR", "AUSDATUM", "PARAMBEZ"])
		matches 	:= labread.SearchExt({"AUSDATUM":">20200101", "PARAM":"rx:^COV[A-Z\d-]+"}  ; , "ERG":"POSITIV"
													,  	 ["ANFNR", "EINDATUM", "AUSDATUM", "PARAM", "BERICHT", "ERG", "BEFUND", "BERICHZEIT"],, "MyToolTip")

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Treffer nach Anforderungsnummer und Parameterbezeichnung sortieren
	;-------------------------------------------------------------------------------------------------------------------------------------------
		results	  	:= Object()
		For idx, obj in matches {
			If !results.haskey(obj.ANFNR)
				results[obj.ANFNR] := Object()
			results[obj.ANFNR][obj.PARAM] := obj.ERG "|" obj.BEFUND "|" obj.EINDATUM "|" obj.AUSDATUM "|" obj.BERICHZEIT
		}

		FileOpen(A_Temp "\labread.json", "w", "UTF-8").Write(JSON.Dump(results,,1,"UTF-8"))

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; 2.Zeit - Dauer des Zählens
	;-------------------------------------------------------------------------------------------------------------------------------------------
		objSize1 	:= labread.GetCapacity()
		duration	:= A_TickCount - starttime

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Ausgabe in Scite
	;-------------------------------------------------------------------------------------------------------------------------------------------
		t := " records count:         `t"   	labread.records                                                	"`n"
		t .= " entrys found:            `t" 	matches.Count()                                            	"`n"
		t .= " Duration all:             `t" 	Round(duration  / 1000, 2) 	" seconds"            	"`n"
		t .= " records found/sec     `t"    Round(matches.Count()/(duration/1000))           	"`n"
		t .= " rec's searched/sec     `t"  	Round(labread.records/(duration/1000))           	"`n"
		t .= " ---------------------------------------------------------------------------------"     	"`n"

		SciTEOutput(RTrim(t, "`n"))

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Datenbank Objekt freigegeben. Damit werden alle gespeicherten Daten gelöscht und Speicher freigegeben
	;-------------------------------------------------------------------------------------------------------------------------------------------
		res          	:= labread.CloseDBF()
		labread	:= ""

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; extrahierte Daten ansehen
	;-------------------------------------------------------------------------------------------------------------------------------------------
		Run, % A_Temp "\labread.json"

ExitApp

MyToolTip(p*) {

	ToolTip, % p.1 "/" p.2, 3000, 100, 13

}


#Include %A_ScriptDir%\..\..\..\Include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\..\lib\SciTEOutput.ahk
#Include %A_ScriptDir%\..\..\..\lib\class_JSON.ahk

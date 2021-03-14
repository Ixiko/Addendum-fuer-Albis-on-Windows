
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
		labbuch   	:= new DBASE(basedir "\LABBUCH.dbf", 4)
		res        	:= labbuch.OpenDBF()
		matches 	:= labbuch.GetFields("GetAll")

		FileOpen(A_Temp "\labbuch.json", "w", "UTF-8").Write(JSON.Dump(matches,,1))

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; 2.Zeit - Dauer des Zählens
	;-------------------------------------------------------------------------------------------------------------------------------------------
		duration	:= A_TickCount - starttime

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Ausgabe in Scite
	;-------------------------------------------------------------------------------------------------------------------------------------------
		t := " records count:     `t"   	labbuch.record                                                 	"`n"
		t .= " entrys found:        `t" 	matches.Count()                                            	"`n"
		t .= " Duration all:         `t" 	Round(duration  / 1000, 2) 	" seconds"            	"`n"
		t .= " records/second:    `t" 	Round(labbuch.record/(duration/1000))           	"`n"
		t .= " ---------------------------------------------------------------------------------"     	"`n"

		SciTEOutput(RTrim(t, "`n"))

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Datenbank Objekt freigegeben. Damit werden alle gespeicherten Daten gelöscht und Speicher freigegeben
	;-------------------------------------------------------------------------------------------------------------------------------------------
		res          	:= labbuch.CloseDBF()
		labbuch	:= ""

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; extrahierte Daten ansehen
	;-------------------------------------------------------------------------------------------------------------------------------------------
		Run, % A_Temp "\labbuch.json"

ExitApp


#Include %A_ScriptDir%\..\..\..\Include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\..\lib\SciTEOutput.ahk
#Include %A_ScriptDir%\..\..\..\lib\class_JSON.ahk

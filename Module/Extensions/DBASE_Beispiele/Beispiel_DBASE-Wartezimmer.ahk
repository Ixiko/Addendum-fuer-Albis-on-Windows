;	Beispielskript für die Verwendung von Addendum_DBASE.ahk
;

	#NoEnv
	#MaxMem 4096
	SetBatchLines, -1

	; Skriptstartzeit
		starttime	:= A_TickCount

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Objekte / Variablen
	;-------------------------------------------------------------------------------------------------------------------------------------------
		RegRead, PathAlbis, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-MainPath
		basedir    	:= PathAlbis "\db"
		Search  	:= Array()
		Search[1] 	:= {	"dbasefile"		: "WZIMMER"
							,	"output"     	: "Wartezimmer-Arzt-19112020"
							,	"showPatID"  	: true
							,	"text"          	: "Anzahl Patienten:     `t"
							,	"fields"      	: {"DATUM"   	: "rx:20201119"
													,	"RAUM"	    	: "Arzt"}}

		nr         	:= 1
		With      	:= Object()
		cSearch		:= Object()

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Ausgabedatei anlegen
	;-------------------------------------------------------------------------------------------------------------------------------------------
		filename 	:= A_Temp "\" search[nr].output ".json"
		;savefile 	:= FileOpen(filename, "w", "UTF-8")

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; neues Datenbankobjekt anlegen. new DBASE(Datenbankpfad [, debug])
	; Datenbankpfad	- vollständiger Pfad zur Datenbank in der Form M:\albiswin\db\BEFUND.dbf
	; debug                	- Ausgabe von Werten zur Kontrolle des Ablaufes, Voreinstellung: keine Ausgabe
	;-------------------------------------------------------------------------------------------------------------------------------------------
		wzimmer   	:= new DBASE(basedir "\" search[nr].DBASEfile ".dbf", false)

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Lesezugriff einrichten. Rückgabewert ist die Position des Dateizeigers
	;-------------------------------------------------------------------------------------------------------------------------------------------
		res        	:= wzimmer.OpenDBF()

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Suche starten. objekt.Search(StringObjekt [, exportPfad])
	; StringObjekt	- Objekt mit den Feldnamen als key und einem RegExString als value
	; exportPfad  	- optional ein Pfad - dann wird aber kein Array mit den extrahierten Zeilen zurückgegeben
	;-------------------------------------------------------------------------------------------------------------------------------------------
		matches 	:= wzimmer.Search(Search[nr].fields)

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Datenbankzugriff kann geschlossen werden
	;-------------------------------------------------------------------------------------------------------------------------------------------
		res         	:= wzimmer.CloseDBF()

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; 1.Zeit - Lesedauer aus der Datenbankdatei
	;-------------------------------------------------------------------------------------------------------------------------------------------
		duration2	:= A_TickCount - starttime

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Ausgabe in Scite
	;-------------------------------------------------------------------------------------------------------------------------------------------
		t := " records count:     `t"   	wzimmer.records                                             	"`n"
		t .= " entrys found:        `t" 	matches.Count()                                               	"`n"
		t .= " Duration all:         `t" 	Round(duration2 / 1000, 2) 	" seconds"            	"`n"
		t .= " records/second:    `t" 	Round(wzimmer.records/ (duration2/1000))     	"`n"
		t .= " ---------------------------------------------------------------------------------"     	"`n"
		t .= " Suchstring (RegEx):`t"  	wzimmer.SearchRegExStr                                  	"`n"
		t .= " Suchstring (Länge):`t"  	wzimmer.lastrxlen                                             	"`n"
		t .= " ---------------------------------------------------------------------------------"     	"`n"

		If search[nr].showPatID
			For PatID, index in With.matchcount
				t .= SubStr(" " PatID, -6) ": " index "x`n"

		SciTEOutput(RTrim(t, "`n"))

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Datenbank Objekt freigegeben. Damit werden alle gespeicherten Daten gelöscht und Speicher freigegeben
	;-------------------------------------------------------------------------------------------------------------------------------------------
		wzimmer := ""
		SciTEOutput(" Anzahl der Elemente vor Freigabe: " objSize1 ", danach: " (wzimmer.GetCapacity() = "" ? 0 : wzimmer.GetCapacity()))

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; extrahierte Daten ansehen
	;-------------------------------------------------------------------------------------------------------------------------------------------

		file := FileOpen(filename, "w", "UTF-8")
		file.Write(JSON.Dump(matches,,4))
		file.Close()

		Run, % filename

ExitApp



#Include %A_ScriptDir%\..\..\Include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#Include %A_ScriptDir%\..\..\lib\class_JSON.ahk

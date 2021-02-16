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
		basedir1 	:= "C:\tmp"
		basedir2 	:= "M:\albiswin\db"
		nr         	:= 1
		Search  	:= Array()
		Search[1] 	:= {	"dbasefile"		: "PATIENT"
							,	"output"     	: "Patientendatenbank"
							,	"showPatID"  	: false
							,	"saveJSON"  	: false
							,	"debug"     	: false
							,	"text"          	: "Anzahl Patienten:     `t"
							,	"get"           	: "ARBEIT"}

		Search[2] 	:= {	"dbasefile"		: "PATIENT"
							,	"output"     	: "Patientendatenbank"
							,	"showPatID"  	: false
							,	"saveJSON"  	: true
							,	"debug"     	: true
							,	"text"          	: "Anzahl Patienten:     `t"
							,	"fields"      	: {"NAME"   	: "Müller"}}

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Ausgabedatei anlegen
	;-------------------------------------------------------------------------------------------------------------------------------------------
		filename 	:= A_Temp "\" search[nr].output ".json"

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; neues Datenbankobjekt anlegen. new DBASE(Datenbankpfad [, debug])
	; Datenbankpfad	- vollständiger Pfad zur Datenbank in der Form M:\albiswin\db\BEFUND.dbf
	; debug                	- Ausgabe von Werten zur Kontrolle des Ablaufes, Voreinstellung: keine Ausgabe
	;-------------------------------------------------------------------------------------------------------------------------------------------
		Patient   	:= new DBASE(basedir1 "\" search[nr].DBASEfile ".dbf", search[nr].debug)

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Lesezugriff einrichten. Rückgabewert ist die Position des Dateizeigers
	;-------------------------------------------------------------------------------------------------------------------------------------------
		res        	:= Patient.Open()

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Suche starten. objekt.Search(StringObjekt [, exportPfad])
	; StringObjekt	- 	Objekt mit den Feldnamen als key und einem RegExString als value
	;							Übergabe eines leeren Objektes gibt alle Einträge zurück
	; exportPfad  	- 	optional ein Pfad - dann wird aber kein Array mit den extrahierten Zeilen zurückgegeben
	;-------------------------------------------------------------------------------------------------------------------------------------------
		matches 	:= Patient.Search(Search[nr].fields)

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Datenbankzugriff kann geschlossen werden
	;-------------------------------------------------------------------------------------------------------------------------------------------
		res         	:= Patient.Close()

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; 1.Zeit - Lesedauer aus der Datenbankdatei
	;-------------------------------------------------------------------------------------------------------------------------------------------
		duration1	:= A_TickCount - starttime

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; 1.Zeit - Lesedauer aus der Datenbankdatei
	;-------------------------------------------------------------------------------------------------------------------------------------------
		If (StrLen(search[nr].get) > 0) {
			retKey := search[nr].get
			For PatID, Pat in matches
				If (StrLen(Pat[retKey]) > 0)
					b .= Pat[retKey] ";" Pat.Name "," Pat.Vorname ";" PatID "`n"
			Sort, b
			b := RTrim("Patienten mit Eintragungen in [" retKey "]:`n" . b, "`n")
		}

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; 2.Zeit - Treffer suchen und sortieren
	;-------------------------------------------------------------------------------------------------------------------------------------------
		duration	:= A_TickCount - starttime
		duration2	:= duration - duration1

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Ausgabe in Scite
	;-------------------------------------------------------------------------------------------------------------------------------------------
		t := " records count:     `t"   	Patient.records                                              	"`n"
		t .= " entrys found:        `t" 	matches.Count()                                               	"`n"
		t .= " duration reading:`t"  	Round(duration1 	/ 1000, 2) 	" seconds"      	"`n"
		t .= " records/second:    `t" 	Round(Patient.records / (duration1/1000))      	"`n"
		t .= " duration result:    `t" 	Round(duration2	/ 1000, 3) 	" seconds"      	"`n"
		t .= " duration all:         `t" 	Round(duration 	/ 1000, 2) 	" seconds"       	"`n"
		t .= " ---------------------------------------------------------------------------------"     	"`n"
		t .= " Suchstring (RegEx):`t"  	Patient.SearchRegExStr                                   	"`n"
		t .= " Suchstring (Länge):`t"  	Patient.lastrxlen                                             	"`n"
		t .= " ---------------------------------------------------------------------------------"     	"`n"
		t .= " gefunden:           `t"  	(StrSplit(b, "`n").MaxIndex() - 1) " Einträge"	      	"`n"

		If search[nr].showPatID
			t .= b

		SciTEOutput(RTrim(t, "`n"))

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; extrahierte Daten ansehen
	;-------------------------------------------------------------------------------------------------------------------------------------------
		If search[nr].saveJSON {
			file := FileOpen(filename, "w", "UTF-8")
			file.Write(JSON.Dump(matches,,4))
			file.Close()
			Run, % filename
		} else {
			file := FileOpen(StrReplace(filename, ".json", ".txt"), "w", "UTF-8")
			file.Write(RTrim(b, "`n"))
			file.Close()
			Run, % StrReplace(filename, ".json", ".txt")
		}

ExitApp



#Include %A_ScriptDir%\..\..\Include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#Include %A_ScriptDir%\..\..\lib\class_JSON.ahk

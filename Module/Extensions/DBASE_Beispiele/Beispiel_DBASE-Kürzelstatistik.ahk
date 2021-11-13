;	Beispielskript für die Verwendung von Addendum_DBASE.ahk
;

	#NoEnv
	#MaxMem 4096
	SetBatchLines, -1
	ListLines, Off

	; Skriptstartzeit
		starttime	:= A_TickCount

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Objekte / Variablen
	;-------------------------------------------------------------------------------------------------------------------------------------------
		RegRead, PathAlbis, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-MainPath
		basedir    	:= PathAlbis "\db"

		nr         	:= 1
		With      	:= Object()
		cSearch		:= Object()
		kuerzel	             	:= Object()
		kuerzel.summe		:= Object()
		kuerzel.jahre 		:= Object()
		kuerzel.liste			:= Object()



	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Ausgabedatei anlegen
	;-------------------------------------------------------------------------------------------------------------------------------------------
		filename 	:= A_Temp "\Kürzelstatistik.json"

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; neues Datenbankobjekt anlegen. new DBASE(Datenbankpfad [, debug])
	; Datenbankpfad	- vollständiger Pfad zur Datenbank in der Form M:\albiswin\db\BEFUND.dbf
	; debug                	- Ausgabe von Werten zur Kontrolle des Ablaufes, Voreinstellung: keine Ausgabe
	;-------------------------------------------------------------------------------------------------------------------------------------------
		kstatistik   	:= new DBASE(basedir "\BEFUND.dbf", true)
		res        	:= kstatistik.OpenDBF()

		VarSetCapacity(dataset, kstatistik.lenDataset)

		; DATUM 10,8	|	KUERZEL 25,5
		while (!kstatistik.dbf.AtEOF) {

				If Mod(A_Index, 1000) = 0
					ToolTip, % A_Index "/" (kstatistik.records)

				bytes 	:= kstatistik.dbf.RawRead(dataset, kstatistik.lendataset)
				set     	:= StrGet(&dataSet, kstatistik.lendataset, "CP1252")

				spos		:= SubStr(set, 146	, 3) + 0
				date   	:= SubStr(set, 10	, 8)
				year  	:= SubStr(set, 10	, 4)
				month 	:= SubStr(set, 14	, 2)
				quartal	:= Ceil(month/3)
				;SciTEOutput("q: " month)
				krzl   	:= Trim(SubStr(set, 25, 5))

				If (StrLen(Trim(krzl)) = 0) || (year+0 < 2010) || (year+0 > 2020) || (lastkrzl = krzl && lastdate = date)
					continue

				lastkrzl	:= krzl
				lastdate	:= date

				If !kuerzel.summe.haskey(krzl)
					kuerzel.summe[krzl] := 1
				else
					kuerzel.summe[krzl] += 1


				If !IsObject(kuerzel.jahre[year]) {
					kuerzel.jahre[year] := Object()
					kuerzel.jahre[year]._quartale := Object()
				}

				If !IsObject(kuerzel.jahre[year]._quartale[quartal])
					kuerzel.jahre[year]._quartale[quartal] := Object()


				If !kuerzel.jahre[year].haskey(krzl)
					kuerzel.jahre[year][krzl] := 1
				else
					kuerzel.jahre[year][krzl] += 1

				If !kuerzel.jahre[year]._quartale[quartal].haskey(krzl)
					kuerzel.jahre[year]._quartale[quartal][krzl] := 1
				else
					kuerzel.jahre[year]._quartale[quartal][krzl] += 1


				If !IsObject(kuerzel.liste[krzl])
					kuerzel.liste[krzl] := Object()

				yearq := year "-" quartal
				If !kuerzel.liste[krzl].haskey(yearq)
					kuerzel.liste[krzl][yearq] :=1
				else
					kuerzel.liste[krzl][yearq] +=1

				If !kuerzel.liste[krzl].haskey(year)
					kuerzel.liste[krzl][year] :=1
				else
					kuerzel.liste[krzl][year] +=1

				If !kuerzel.liste[krzl].haskey("Q" quartal)
					kuerzel.liste[krzl]["Q" quartal] :=1
				else
					kuerzel.liste[krzl]["Q" quartal] +=1

				If !kuerzel.liste[krzl].haskey("M" month)
					kuerzel.liste[krzl]["M" month] :=1
				else
					kuerzel.liste[krzl]["M" month] +=1

				;~ If A_Index > 200000
					;~ break

		}


	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Datenbankzugriff kann geschlossen werden
	;-------------------------------------------------------------------------------------------------------------------------------------------
		res         	:= kstatistik.CloseDBF()

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; 1.Zeit - Lesedauer aus der Datenbankdatei
	;-------------------------------------------------------------------------------------------------------------------------------------------
		duration2	:= A_TickCount - starttime

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Datenbank Objekt freigegeben. Damit werden alle gespeicherten Daten gelöscht und Speicher freigegeben
	;-------------------------------------------------------------------------------------------------------------------------------------------
		rechnung := ""
		SciTEOutput(" Anzahl der Elemente vor Freigabe: " objSize1 ", danach: " (rechnung.GetCapacity() = "" ? 0 : rechnung.GetCapacity()))

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; extrahierte Daten ansehen
	;-------------------------------------------------------------------------------------------------------------------------------------------

		file := FileOpen(filename, "w", "UTF-8")
		file.Write(JSON.Dump(kuerzel,,2))
		file.Close()

		Run, % filename

ExitApp

#Include %A_ScriptDir%\..\..\Include\Addendum_Datum.ahk
#Include %A_ScriptDir%\..\..\Include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#Include %A_ScriptDir%\..\..\lib\class_JSON.ahk

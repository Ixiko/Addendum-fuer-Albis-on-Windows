; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                      letzte Vorsorge
;
;      Funktion:           	erstellt eine Liste mit den letzten Vorsorgeuntersuchungen der Patienten, diese werden dann im Infofenster angezeigt  (albiswin\db\Befund.dbf)
;
;		Hinweis:				Aufruf könnte einmal am Anfang des Quartals erfolgen. Skript ist als für unabhängigen Aufruf konzipiert und schreibt die Daten im binären Format!
;
;      Basisskript:          	Addendum.ahk, Addendum_Gui.ahk
;
;		Abhängigkeiten:	include\Addendum_DBASE.ahk !!
;
;	                    	Addendum für Albis on Windows
;                        	by Ixiko started in September 2017 - last change 28.12.2020 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	; Fileformat
		; Bytes 1-7 	Pat.Nr.
		; Bytes 8-16 	letzte GVU im Format YYYYMMDD

	; Einstellungen
		#NoEnv
		#KeyHistory, Off
		#MaxMem 4096

		SetBatchLines, -1
		ListLines    	, Off

	; wichtige Pfade
		RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)        ; Addendum-Verzeichnis
		RegRead, PathAlbis, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-MainPath
		basedir         	:= PathAlbis "\db"

	; große Patientenliste laden
		infilter 	:= ["NR", "NAME", "VORNAME", "GEBURT", "MORTAL", "LAST_BEH", "RECHART"]
		PatDBF	:= ReadPatientDBF(basedir, infilter,, 1)

	; analysiere Befund.dbf
		inpattern   	:= {"KUERZEL":"lko", "INHALT":"01732"}
		outpattern 	:= {"DATUM":"rx:(20|19)0[0-9]\d{4}"}                                       	; RegEx - alle Datumsformate vor 2010 sind ausgeschlossen
		GVUPatients	:= GetDBFData(basedir "\BEFUND.dbf", inpattern, outpattern,,0)
		For idx, m in GVUPatients
			PatDBF[m.PatNR].LastGVU := m.Datum

	; größte und kleinste PATID
		maxPatID := 100, minPatID:= 100
		For PatID, m in PatDBF
			if (PatID > maxPatID)
				maxPATID := PatID
			else if (PatID < minPatID)
				minPATID := PatID

	;letzte GVU Object
		lgv := FileOpen(AddendumDir "\logs'n'data\_DB\lastGVU.db", "w", "UTF-8")
		Loop, % (maxPatID-minPatID+1) {
			PatID := minPatID + A_Index
			towrite := SubStr("0000000" A_Index, -6) (StrLen(PatDBF[PatID].lastGVU) > 0 ? PatDBF[PatID].lastGVU : "00000000")
			lgv.Write(towrite)
		}
		lgv.Close()




ExitApp


GetDBFData(DBASEfilepath, inpattern:="", outpattern:="", options:="", debug:=0) {                	;-- holt Daten aus einer beliebigen Datenbank

	dbfile      	:= new DBASE(DBASEfilepath, debug)
	res        	:= dbfile.Open()
	matches 	:= dbfile.Search(inpattern, 0, options)
	res         	:= dbfile.Close()
	dbfile    	:= ""

return matches
}

#Include %A_ScriptDir%\..\..\..\Include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\..\Include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\..\Include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\..\Include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\..\Include\Addendum_Datum.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\..\Include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\..\Include\Addendum_PdfHelper.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_PraxTT.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\..\Include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\..\Include\Addendum_Window.ahk

#Include %A_ScriptDir%\..\..\..\lib\ACC.ahk
#Include %A_ScriptDir%\..\..\..\lib\class_JSON.ahk
#include %A_ScriptDir%\..\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\..\lib\ini.ahk
#Include %A_ScriptDir%\..\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\..\lib\Sift.ahk
#Include %A_ScriptDir%\..\..\..\lib\SciTEOutput.ahk
;}


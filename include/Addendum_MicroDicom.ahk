; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;     	Addendum_PDFReader
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;		Verwendung:	    	-	RPA Funktionen für MicroDicom einer kostenlosen Software zur Anzeige von DICOM Dateien
; 	                                -	optimiert für die MicoDicom Version 3.4+
;
;
;       Abhängigkeiten:   	-	\lib\SciteOutPut.ahk
;                                	-	\include\Addendum_Control.ahk
;
;
;      	MicroDicom - kostenloser DICOM Viewer für Windows
;   	Download Seite: https://www.microdicom.com/downloads.html
;
;
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Addendum_PDFReader last change:    	09.12.2020
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
return

MicroDicom_VideoExport(DestPath="") {                                      	;-- automatischer Export in ein Videoformat mit Dateinamenerstellung anhand von Dicom Tags

		StringCaseSense, Off

		DestPath :=  (StrLen(DestPath) = 0) ? Addendum.VideoOrdner : DestPath

	; window description for MicroDicom 64bit V2 'Export to video' dialog
		MDExport  	:= {	"Title"              	: {"class":"Export to video ahk_class #32770", "Type":"window"}
								, 	"Dest"             	: {"class":"Edit1"         	, "Type":"edit"}
								,	"Name"           	: {"class":"Edit2"         	, "Type":"edit"}
								, 	"Source"          	: {"class":"ComboBox1"	, "Type":"combobox"}
								, 	"Planes"           	: {"class":"Button4"     	, "type":"checkbox"}              	; export each planes
								, 	"SpFiles"          	: {"class":"Button5"     	, "type":"checkbox"}               	; Separate files for &every
								,	"WMV"             	: {"class":"Button7"     	, "type":"button"}                	; WMV Export (Checkbox)
								,	"AVI"                	: {"class":"Button8"     	, "type":"button"}                  	; AVI Export (Checkbox)
								,	"FRateDefault"  	: {"class":"Button9"     	, "type":"button"}                  	; Frame rate custom
								,	"FRateCustom"  	: {"class":"Button10"     	, "type":"button"}                  	; Frame rate custom
								,	"FPS"              	: {"class":"Edit3"         	, "Type":"edit"}                     	; frames per second
								,	"WMVQuality"   	: {"class":"Edit4"         	, "Type":"edit"}                     	; Exportquality
								,	"ComprYes"      	: {"class":"Button11"     	, "type":"button"}                	; AVI Compression Yes
								,	"ComprNo"      	: {"class":"Button12"    	, "type":"button"}                  	; AVI Compression No
								,	"SizeOrig"      	: {"class":"Button14"    	, "type":"button"}                  	; VideoSize Original
								,	"SizeAsScreen"  	: {"class":"Button15"    	, "type":"button"}                  	; Video Size As on screen
								,	"SizeOther"     	: {"class":"Button16"    	, "type":"button"}                  	; Video Size As on screen
								,	"SizeWidth"      	: {"class":"Edit5"         	, "Type":"edit"}                     	; Video Size Other Width
								,	"SizeHeight"      	: {"class":"Edit6"         	, "Type":"edit"}                     	; Video Size Other Height
								,	"Annotation"     	: {"class":"Button18"    	, "type":"checkbox"}              	; show annotations
								,	"AllOverlay"     	: {"class":"Button19"    	, "type":"button"}                  	; all overlay
								,	"Anonymous"    	: {"class":"Button20"    	, "type":"button"}                  	; anonymous overlay
								,	"NoOverlay"    	: {"class":"Button21"    	, "type":"button"}                  	; Without overlay
								,	"Export"         	: {"class":"Button22"    	, "type":"button"}                	; Export button
								,	"Cancel"         	: {"class":"Button23"    	, "type":"button"}}                	; Cancel button

	; smart RPA working object for function RPA() - sets all needed options in 'Export to video' dialog
		ExportRPA 	:= Array()
		ExportRPA[1]	:= ["Dest|" DestPath
					    		, "Name|%PatName%"		; wird ersetzt
					    		, "Source|i)All\spatients" 	; RegEx string
					    		, "Planes|check"
					    		, "SpFiles|uncheck"
					    		, "WMV|click"
					    		, "FRateCustom|click"
					    		, "FPS|10"
					    		, "WMVQuality|100"
					    		, "SizeOrig|click"
					    		, "Annotation|check"
					    		, "NoOverlay|click"
					    		, "Export|click"]

	; Handle des MicroDicom Hauptfensters
		If !(hMDicom := WinExist("MicroDicom viewer")) {
			PraxTT("Der automatische Videoexport ist fehlgeschlagen.", "4 0")
			return
		}

	; den Exportdialog erstmal wieder schließen
		If (hExport := WinExist(MDExport.title.class))
			VerifiedClick(MDExport.Cancel.class, hExport)

	; die Dicom Tags anzeigen
		If !(hDicomTags := MicroDicom_Invoke("DicomTags", hMDicom, "DICOM Tags ahk_class #32770", 6)) {
			SciTEOutput("Der automatische Videoexport ist fehlgeschlagen.")
			PraxTT("Der automatische Videoexport ist fehlgeschlagen.`nDer Dialog Dicom Tags konnte nicht aufgerufen werden", "4 0")
			return
		}

	; Patientenname, Geburtstag,Untersuchungsart und Untersuchungstag ermitteln (Auslesen der SysListview321 im Dicom Tags Dialog)
		ExportRPA[2] := MicroDicom_BuildFileName(hDicomTags)

	; Dicom Tags Dialog schliessen
		VerifiedClick("Button1", hDicomTags)
		WinWaitClose, % "ahk_id " hDicomTasgs,, 3

	; den Exportvorgang mit den ausgelesenen Daten starten
		If !MicroDicom_Invoke("ExportToVideo", hMDicom, MDExport.title.class, 6) {
			PraxTT("Der automatische Videoexport ist fehlgeschlagen.", "4 0")
			return
		}

	; den Inhalt der Dicom CD als Video exportieren
		RPA(MDExport, ExportRPA)

	; wartet bis zu 3min auf den Abschluss des Exportvorganges - ein Explorerfenster wird nach Abschluss angezeigt
		;WinWait, % "ahk_class CabinetWClass", % Addendum.VideoOrdner, 300  ; werden andere Threads noch ausgeführt wenn WinWait aktiv ist?

	; Schliessen und Auswerfen der CD nach Export
		;MsgBox, 4, Addendum für Albis on Windows, % ""

return
}

MicroDicom_Invoke(command, MicroDicomID="", WinToWait="", TimeToWait=6) {          	;-- Menuaufrufe für MicroDicom 64bit

		static MicroDicom := { 	"DicomTags":           	40018  	; View
										,	"ExportToImage":      	32790  	; File/Export/To a picture file
										,	"ExportToVideo":      	32806}	; File/Export/To a video file

	; wm_command code wird zurück gegeben
		If !MicroDicomID
			return MicroDicom[command]

	; Fenster ist schon geöffnet, Rückgabe des Handle
		If (StrLen(WinToWait) > 0) && (hwnd := WinExist(WinToWait))
			return hwnd

	; Senden des wm_command Befehls
		PostMessage, 0x111, % MicroDicom[command],,, % "ahk_id " MicroDicomID

	; wartet auf den sich öffnenden Dialog und gibt das Fensterhandle zurück
		If (StrLen(WinToWait) > 0) {
			WinWait, % WinToWait,, % TimeToWait
			return WinExist(WinToWait) 	; handle ist 0 bei nicht vorhandenem Fenster
		}


return ; gibt nichts zurück!
}

MicroDicom_BuildFileName(hDicomTags) {                                                                    	;-- erstellt einen sinnvollen Dateinamen für den Export

		ControlGet, tags, List,, SysListView321, % "ahk_id " hDicomTags

		RegExMatch(tags, "StudyDescription\s*\w+\s*\w+\s*\w+\s*([\w\d\s]+)"                  	, description)
		RegExMatch(tags, "i)PatientName\s*\w+\s*\w+\s*\w+\s*(?<Name>[\p{L}\s^]+)"	, Patient)
		RegExMatch(tags, "i)PatientBirthDate\s*\w+\s*\w+\s*\w+\s*(?<BD>\d+)"              	, Pat)
		RegExMatch(tags, "i)StudyDate\s*\w+\s*\w+\s*\w+\s*(?<D>\d+)"                         	, St)

		PatientName := RegExReplace(PatientName, "(\p{L})\^(\p{L})", "$1, $2" )
		PatientName := RegExReplace(PatientName, "(\p{L})\s+(\p{L})", "$1, $2" )
		PatientName := StrReplace(PatientName, "^")
		PatientName := Trim(StrReplace(PatientName, "`n"))
		PatientBirthDate := SubStr(PatBD, 7,2) "." SubStr(PatBD, 5,2) "." SubStr(PatBD, 1,4)
		StudyDate := SubStr(StD, 7,2) "." SubStr(StD, 5,2) "." SubStr(StD, 1,4)

return "Name|" PatientName " (" PatientBirthdate ") - " StrReplace(description1, "`n") " vom " StudyDate " - "
}

RPA(WinObj, steps, stepsPause=50) {                                                                              	;-- universal prototype: 'macro like' robot process automation for standard system controls in windows

	; universal function to check or uncheck checkboxes, radiobuttons or to select entries in combo/listboxes and to set text entries in edit controls
	; function is made to be used on standard windows controls
	; the main goal is to give users the possibility to have their own settings for windows
	;
	; dependancies: Addendum_Control.ahk (VerifiedSetText/Click/Check/Choose)

	If !IsObject(WinObj)
		throw Exception("parameter 'WinObject': is not an object", "RPA")

	If !IsObject(steps)
		throw Exception("parameter 'steps': is not an object", "RPA")

	If !(hRPAWin := WinExist(WinObj.title.class))
		throw Exception("Could not find window:`n" WinObj.title.class, "RPA")

	For idx, step in steps	{

		on := StrSplit(step, "|").1, do := StrSplit(step, "|").2

		switch WinObj[on].type
		{
				case "edit":
					VerifiedSetText(WinObj[on].class, do, hRPAWin)

				case "button":
					VerifiedClick(WinObj[on].class, hRPAWin)

				case "checkbox":
					VerifiedCheck(WinObj[on].class, hRPAWin,,, InStr(do, "uncheck") ? false:true)

				case "combobox":
					VerifiedChoose(WinObj[on].class, hRPAWin, do)
		}

		Sleep, % stepsPause

	}

}


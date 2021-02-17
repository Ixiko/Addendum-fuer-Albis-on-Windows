; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . .                                                                                                                                                          	. . . . . . . . . .
; . . . . . . . . .                                                 	ADDENDUM  ABRECHNUNGSHELFER                                            	. . . . . . . . . .
											                			 Version:="0.98a" , vom:="01.02.2021"
; . . . . . . . . .                                                                                                                                                         	. . . . . . . . . .
; . . . . . . . . .  ROBOTIC PROCESS AUTOMATION FOR THE GERMAN MEDICAL SOFTWARE "ALBIS ON WINDOWS"	. . . . . . . . . .
; . . . . . . . . .         BY IXIKO STARTED IN SEPTEMBER 2017 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE         	. . . . . . . . . .
; . . . . . . . . .                                                                                                                                                         	. . . . . . . . . .
; . . . . . . . . .        RUNS WITH AUTOHOTKEY_H AND AUTOHOTKEY_L IN 32 OR 64 BIT UNICODE VERSION         	. . . . . . . . . .
; . . . . . . . . .                            PLEASE USE ONLY THE NEWEST VERSION OF AUTOHOTKEY!!!                              	. . . . . . . . . .
; . . . . . . . . .                                                                                                                                                         	. . . . . . . . . .
; . . . . . . . . .                         SORRY THIS SOFTWARE IS ONLY AVAIBLE IN GERMAN LANGUAGE                           	. . . . . . . . . .
; . . . . . . . . .                                                                                                                                                         	. . . . . . . . . .
; . . . . . . . . .         FOR THE BEST VIEW USE ⊡ FUTURA BK BT ⊡ (SORRY I'DONT LIKE MONOSPACE FONTS)          	. . . . . . . . . .
; . . . . . . . . .                                                                                                                                                         	. . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .


/* Erwarten Sie keine Henne die goldene Eier am Stück legt

	Dieses Skript gehört zu den ersten Skripten die ich begonnen habe. Die Benutzeroberfläche ist wenig anwenderfreundlich.
	Der Programmcode ist grausam zu lesen, so daß ich mich selbst jedesmal in diesen einarbeiten muss, bevor ich Änderungen vor-
	nehmen kann. Vieles sollte einfach nur schnell und irgendwie funktionieren um nicht wieder tagelang mehrere Seiten füllende
	freie Statistiken von Albis "Klick für Klick" durchgehen wollte. Deshalb gibt esso gut wie keine Beschreibungen wie der Abrechnungs-
	helfer überhaupt funktioniert. Da ich halbwegs weiß in welcher Reihenfolge ich die 'Makro' genannten Skriptteile aufrufen muss,
	funktioniert aus meiner Sicht vieles wirklich gut. Jemand der sich hier "DEN ABRECHNUNGSHELFER" vorgestellt hat, kann nur
	enttäuscht werden.

	Ich habe von Anfang an zwar versucht Addendum nicht nur für meine Praxis zu programmieren, doch übersteigt der Aufwand
	für eine allgemeine Verwendung die zeitlichen Möglichkeiten dieses Ein-Mann-Projektes.
	Deshalb ist immer noch viel Handarbeit für andere Anwender notwendig. Doch kann ich versprechen das es sich auch für Sie
	zeitlich und finanziell lohnen wird, das Abrechnungshelfer Skript zu verwenden.

	Und noch eines. Der Abrechnungshelfer ist nicht mein "Lieblingsprojekt".

*/

;-------------------------------------------------------------------------------------------------------------------------------------------------------------
; SKRIPTEINSTELLUNGEN
;-------------------------------------------------------------------------------------------------------------------------------------------------------------;{
	#NoEnv
	#Persistent
	#MaxMem 4096                                                           	; brauchen wir wirklich bis zu 4GB für eine einzelne Variable?

	SetBatchLines            	, -1
	SetWinDelay              	, -1
	SetControlDelay         	, -1

	FileEncoding             	, UTF-8
	DetectHiddenWindows	, On

	Menu, Tray, Icon, % "hIcon: " Create_Abrechnungshelfer_ico(true)

	; DPI-Skalierung einschalten
	Result := DllCall("User32\SetProcessDpiAwarenessContext", "UInt" , -1)
;}

;-------------------------------------------------------------------------------------------------------------------------------------------------------------
; VARIABLEN
;-------------------------------------------------------------------------------------------------------------------------------------------------------------;{
	; Addendum Stamm- und Datenverzeichnis auslesen (hat nichts mit den Albisverzeichnissen zu tun!)
	global 	AddendumDir
				RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)
	global 	AlbisWorkDir
	global 	TagesprotokollPfad
	global 	AddendumDbPath 	:= GetAddendumDbPath()
	global 	AddendumID          	:= GetAddendumID()
	global 	Statistik                   	:= Object()
	global 	QData     	            	:= Object()
	global 	RegExMining            	:= Object()
	global 	PatDB                      	:= Object()
	global 	RegExFind               	:= Array()
	global 	compname             	:= StrReplace(A_ComputerName, "-")                    	; der Name des Computer auf dem das Skript läuft

;}

;-------------------------------------------------------------------------------------------------------------------------------------------------------------
; EINSTELLUNGEN AUS DER ADDENDUM.INI HOLEN UND VARIABLEN DEFINIEREN
;-------------------------------------------------------------------------------------------------------------------------------------------------------------;{

	; Gui Einstellungen
		IniRead, BgColor       	, % AddendumDir "\Addendum.ini", Abrechnungshelfer	, HgFarbe
		If InStr(BgColor, "Error") {
			IniRead, BgColor    	, % AddendumDir "\Addendum.ini", Addendum          	, DefaultBgColor
			IniWrite, % BgColor	, % AddendumDir "\Addendum.ini", Abrechnungshelfer	, HgFarbe
		}

		IniRead, BgColor2     	, % AddendumDir "\Addendum.ini", Abrechnungshelfer	, HgFarbe2
		If InStr(BgColor2, "Error") {
			IniRead, BgColor2 	, % AddendumDir "\Addendum.ini", Addendum          	, DefaultBgColor2
			IniWrite, % BgColor2	, % AddendumDir "\Addendum.ini", Abrechnungshelfer	, HgFarbe2
		}

		IniRead, FontColor    	, % AddendumDir "\Addendum.ini", Abrechnungshelfer	, FontFarbe
		If InStr(FontColor, "Error") {
			IniRead, FontColor	, % AddendumDir "\Addendum.ini", Addendum          	, DefaultFntColor
			IniWrite, % FontColor, % AddendumDir "\Addendum.ini", Abrechnungshelfer	, FontFarbe
		}

		IniRead, FontSize       	, % AddendumDir "\Addendum.ini", Abrechnungshelfer	, FontGroesse
		If InStr(FontSize, "Error") {
			IniRead, FontSize		, % AddendumDir "\Addendum.ini", Addendum          	, StandardFontSize
			IniWrite, % Fontsize	, % AddendumDir "\Addendum.ini", Abrechnungshelfer	, FontGroesse
		}

		IniRead, Font                	, % AddendumDir "\Addendum.ini", Abrechnungshelfer	, Schriftart
		If InStr(Font, "Error") {
			IniRead, Font          	, % AddendumDir "\Addendum.ini", Addendum          	, StandardFont
			IniWrite, % Font        	, % AddendumDir "\Addendum.ini", Abrechnungshelfer	, Schriftart
		}

		IniRead, AlbisWorkDir	, % AddendumDir "\Addendum.ini", Albis                    	, AlbisWorkDir
		If InStr(AlbisWorkDir, "Error") {
			;IniRead, Font, % AddendumDir . "\Addendum.ini", Addendum, StandardFont
			;IniWrite, % Font,  % AddendumDir . "\Addendum.ini", Addendum, Abrechnungshelfer-Schriftart
		}

		IniRead, Fensterposition	,  % AddendumDir "\Addendum.ini", % compname, % "Abrechnungshelfer_Position"
		If InStr(Fensterposition, "Error") {
			guiPos:= "AutoSize"
		}	else	{
			factor := A_ScreenDPI / 96
			pos := StrSplit(Fensterposition, "|")
			If (pos.1 + pos.3 > A_ScreenWidth)
				pos.1 := Floor((A_ScreenWidth/2)  - (pos.3/2) * factor)
			If (pos.2 + pos.4 > A_ScreenHeight)
				pos.2 := Floor((A_ScreenHeight/2) - (pos.4/2) * factor)
			guiPos := "x" pos.1 " y" pos.2 " w" Round(pos.3 * factor) " h" Round(pos.4 * factor)
		}

		IniRead, lTprotfile    		, % AddendumDir "\Addendum.ini", Abrechnungshelfer	,  letzte_Tagesprotokolldatei
		IniRead, lProtokollNr		, % AddendumDir "\Addendum.ini", Abrechnungshelfer	,  letzte_Makrofunktion

	; Tagesprotokollpfad überprüfen
		TagesprotokollPfad	:= AddendumDir "\Tagesprotokolle\TP-Abrechnungshelfer"
		If !InStr(FileExist(TagesprotokollPfad), "D")
			FileCreateDir, % TagesprotokollPfad

	; Ergebnislistenpfad überprüfen
		AlbisListenPfad			:= AddendumDir "\Tagesprotokolle\Listen für Albis"
		If !InStr(FileExist(AlbisListenPfad), "D")
			FileCreateDir, % AlbisListenPfad

	; fertige Makro's für die Auswertungen als AHK-Objekt
		Makro:= Object()
		Makro.push({"DispName"	: "01. Tagesprotokoll erstellen"
							, "label"              	: "M10"
							, "Beschreibung"	: "hilft bei der Erstellung von ein oder mehreren Tagesprotokollen"
							, "Parameter"     	: false})
		Makro.push({"DispName"	: "02. ein Tagesprotoll parsen"
							, "label"     	: "M1"
							, "Parameter"	: false})
		Makro.push({"DispName"	: "03. freie Statistik"
							, "label"			: "M2"
							, "Parameter": true})
		Makro.push({"DispName"	: "04. Patienten für die Vorsorgeliste suchen"
							, "label"      	: "M3"
							, "Parameter"	: true})
		Makro.push({"DispName"	: "05. nach fehlenden GB Ziffern suchen"
							, "label"     	: "M4"
							, "Parameter": true})
		Makro.push({"DispName"	: "06. Ziffernstatistiken erstellen"
							, "label"     	: "M5"
							, "Parameter": false})
		Makro.push({"DispName"	: "07. Ziffernstatistik parsen"
							, "label"       	: "M6"
							, "Parameter"	: false})
		Makro.push({"DispName"	: "08. Ziffernstatistiken zusammen führen"
							, "label"  		: "M7"
							, "Parameter"	: true})
		Makro.push({"DispName"	: "09. fehlende Chronikerziffern"
							, "label"     	: "M8"
							, "Parameter": true})
		Makro.push({"DispName"	: "10. neue Chroniker finden"
							, "label"     	: "M9"
							, "Parameter": true})
		Makro.push({"DispName"	: "11. Praxisstatistiken "
							, "label"             	: "M16"
							, "Beschreibung"	: "Erstellt eine Statistik über alle Quartale zu Arbeitsunfähigkeitsbescheinigungen (Dauer, Anzahl der Patienten, Anzahl der Erst- und Folgebescheinigungen)"
							, "Parameter"      	: false})
		Makro.push({"DispName"	: "12. Dauerdiagnosen sortiert anzeigen"
							, "label"      	: "M11"
							, "Parameter": true})
		Makro.push({"DispName"	: "13. Querverweise für Tagesprotokolle erstellen"
							, "label    " 	: "M12"
							, "Parameter": false})
		Makro.push({"DispName" 	: "14. zähle Überweisungen aller Quartale"
							, "label"             	: "M13"
							, "Parameter"     	: false
							, "Beschreibung"	: "zählt die Anzahl der Überweisungen pro gewähltem Quartal."})
		Makro.push({"DispName" 	: "15. Aktenkürzel und Nutzerkategorisierung"
							, "label"             	: "M14"
							, "Parameter"     	: false
							, "Beschreibung"	: "findet heraus welche Kürzel benutzt werden und stellt diese übersichtlich zusammen und vergibt die Mögllichkeit die Kürzel zu kategorisieren oder genauer zu benennen.."})
		Makro.push({"DispName" 	: "16. Fachgruppenbezeichnungen ermitteln"
							, "label"             	: "M15"
							, "Parameter"     	: false
							, "Beschreibung"	: "sucht die Namen der Fachgruppen heraus."})
		Makro.push({"DispName" 	: "17. Rezepte- und Medikamentenstatistik"
							, "label"             	: "M18"
							, "Parameter"     	: false
							, "Beschreibung"	: "kleine Rezeptstatistik der gewählten Quartale"})
		;Makro.push({"DispName": "", "label":""})

	 ; verschieden Datenobjekte
		CICD    	:= ChronikerICDListe(A_ScriptDir "\data\ICD-chronisch_krank.txt")   	; Chroniker ICD-Schlüsselnummern lesen
		Chroniker	:= ChronikerListe(AddendumDbPath "\DB_Chroniker.txt")                	; Patienten-IDs der als Chroniker vermerkten Patienten lesen
		PatDB   	:= ReadPatientDatabase(AddendumDir "\logs'n'data\_DB")                	; Patienten-Datenbank einlesen
		if !IsObject(PatDB)
			MsgBox, % A_ScriptName, Die Patientendatenbank konnte nicht erstellt/gelesen werden.

;}

;-------------------------------------------------------------------------------------------------------------------------------------------------------------
; ALLE AUSWERTBAREN TAGESPROTOKOLLDATEIEN IM TAGESPROTOKOLLPFAD ERMITTELN
;-------------------------------------------------------------------------------------------------------------------------------------------------------------;{
	Tagesprotokolle	:= ReadTPDir(TagesprotokollPfad, "*_TP-AbrH", "txt", "R")

;}

;-------------------------------------------------------------------------------------------------------------------------------------------------------------
; REGEX-STRINGS IN EINEM ARRAY ZUM PARSEN EINES ALBIS TAGESPROTOKOLLS
;-------------------------------------------------------------------------------------------------------------------------------------------------------------;{

	; zum Entfernen von Strings aus dem Albis-Tagesprotokoll
		RegExTp	 := Object()
		RegExTP[1] := {regex:"m)\\H\\N\+\d+"														, replace:""}		; für (\H\N+00000)12345AOK
		RegExTP[2] := {regex:"m)\\R\+\d+"       														, replace:""}		; ZeilenCode oder Datenbank-Adresse könnte das sein
		RegExTP[3] := {regex:"m)\s{5}[A-Z]{1,2}\s+"												, replace:""}		; das Arztkürzel entfernen (ich brauche dieses nicht bei der Auswertung)
		RegExTP[4] := {regex:"`am)^\s+(?=\d\d\.\d\d\.\d\d\d\d)"							, replace:""}
		RegExTP[5] := {regex:"`am)\r\n"				                                        			, replace:"`n"}
		RegExTP[6] := {regex:"`am)\n"				                                            			, replace:"`r`n"}
		;RegExTP[7] := {regex:"m)\\HKeine\sDiagnosen\svorhanden\!"	                   	, replace:""}
		;RegExTP[5] := {regex:"m)\H"																			, replace:""}

	; zum Identifizieren bestimmter Zeilen
		RegExFind     	:= Array()
		RegExFind[1] 	:= "(?<=\()\d+(?=\)\s\d\d\.\d\d.\d\d\d\d\,)"						; identifiziert eine Zeile als Patientenzeile und liest die Patienten ID aus
		RegExFind[2]  	:= "G[VU][VU].H[KS][KS]"											   			; GVU/HKS - findet dies auch wenn Buchstaben verdreht oder doppelt sind
		RegExFind[3]  	:= "^.*(?=\s\(\d+\))" 												   			; Patientenzeile findet den Namen ( >> nur nutzen wenn die Zeile über RegExFind[1] als Patientenzeile identifiziert wurde!)
		RegExFind[4]  	:= "\d\d\.\d\d\.\d\d\d\d"											   			; Patientenzeile findet das Geburtsdatum (>> nur nutzen wenn die Zeile über RegExFind[1] als Patientenzeile identifiziert wurde!)
		RegExFind[5] 	:= "lko\s+"                                                                        	; findet lko 03001 bis 03009 - z.B. um zu klären ob ein Abrechnungsschein angelegt wurde
		RegExFind[6] 	:= "^\d\d\.\d\d\.\d\d\d\d"                                              	; erkennt einen Datumseintrag, z.B. den Tag des ersten Karteikarteneintrages
		RegExFind[7] 	:= "i)^(Behandlung|anamnestisch)\s+"                              	; identifiziert die Felder mit den Dauerdiagnosen
		RegExFind[8] 	:= "\{(?<ICD>[A-Z]\d+[\d\.\-\*]+)[LRGB]*\}*"                     	; ICD Feld - ohne L,R,G
		RegExFind[9] 	:= "\slko\s"                                                                        	; findet ausschließlich 'lko' für Leistungskomplex
		RegExFind[10]	:= "^(?<Nachname>[A-ZÄÖÜ][a-z\-\sßäöü]+)\s*\,\s*(?<Vorname>.*)?\((?<ID>\d+)\)\s+(?<Geburtstag>\d\d\.\d\d\.\d\d\d\d)\s*\,\s*(?<Kasse>.*)?\(\)" ; <<< nicht fertig!
		;Amus, Christian (29649) 04.05.1977, BARMER(IK:100580002, VNR: V123456789

	; um Daten zu extrahieren ;{
		RegExMining := Object()
		RegExMining["all"] := "(?<caveVon>^Cave\!\:)|"
		RegExMining["all"] .= "(?<Dauerdiagnosen>^Dauerdiagnosen\:)|"
		RegExMining["all"] .= "(?<Dauermedikamente>^Dauermedikamente\:)|"
		RegExMining["all"] .= "(?<Behandlung>^\s*Behandlung\s)|"
		RegExMining["all"] .= "(?<Anamnestisch>^\s*anamnestisch\s)|"
		RegExMining["all"] .= "(?<ddoc>^\s*ddoc\s)|"
		RegExMining["all"] .= "(?<Anamnese>^\s*anam\s)|"
		RegExMining["all"] .= "(?<Befund>^\s*bef\s)|"
		RegExMining["all"] .= "(?<Kassenrezept>^\s*medrp\s)|"
		RegExMining["all"] .= "(?<info>^\s*info\s)|"
		RegExMining["all"] .= "(?<Bemerkung>^\s*bem\s)|"
		RegExMining["all"] .= "(?<Therapie>^\s*ther\s)|"
		RegExMining["all"] .= "(?<Ueberweisung>^\s*füb\s)|"
		RegExMining["all"] .= "(?<Arbeitsunfaehigkeit>^\s*fau\s)|"
		RegExMining["all"] .= "(?<Leistungskomplex>^\s*lk[oü]\s)|"
		RegExMining["all"] .= "(?<Diagnose>^\s*dia\s)|"
		RegExMining["all"] .= "(?<Altanamnese>^\s*z\s)|"
		RegExMining["all"] .= "(?<LaborPathologisch>^\s*labor\s)|"
		RegExMining["all"] .= "(?<Privatziffern>^\s*lp\s)|"
		RegExMining["all"] .= "(?<>^\s*\s)|"
		RegExMining["all"] .= "(?<>^\s*\s)|"
		RegExMining["all"] .= "(?<>^\s*\s)|"
		RegExMining["all"] .= "(?<>^\s*\s)|"
		RegExMining["caveVon"]         	:= {"start"     	: "^Cave\!\:"
															  , "stoplabel"	: ["RecordDate"]
															  , "parent" 	: ""
															  , "mining"	: "lines complete"}
		RegExMining["RecordDate"]    	:= {"start"     	: "(?<start>\d\d.\d\d.\d\d\d\d)\s+(?<part>.*)"
															 , "stoplabel"	: ["RecordDate", "newLine"]
															 , "parent"   	: ""
															 ,  "mining"	: ["start", "part"]}
		RegExMining["Dauerdiagnosen"]	:= {"start"     	: "^Dauerdiagnosen\:"
															 , "stoplabel"	: ["newline"]
															 , "parent"   	: ""
															 ,  "mining"	: ["newKey"]}
		RegExMining["Behandlung"]    	:= {"start"     	: "^Behandlung\s+(?<text>.*)"
															 , "stoplabel"	: ["newline", "anamnestisch"]
															 , "parent"  	: ["Dauerdiagnosen", "RecordDate"]
															 , "mining" 	: ["text"]}
		RegExMining["Anamnestisch"] 	:= {"start"     	: "^Anamnestisch\s+(?<text>.*)"
                                                    		 , "stoplabel"	: ["newline", "Behandlung"]
															 , "parent"   	: ["Dauerdiagnosen", "RecordDate"]
															 ,  "mining"	: ["text"]}
		; findet nur lko
		;}

	; für Ziffernstatistik ;{
		RegExFS  	 := Array()
		RegExFS[1] := {regex:"`am)--------\|---------\|-----------\|-----------\|\r\n"	                            	, replace:"\b"}
		RegExFS[2] := {regex: "`am)\\R\+[0-9]{10}\\N\+[0-9]{10}\s*"	                                            	, replace:""}      	;für \R+0241954868\N+0000000001 in der Ziffernstatistik,
		RegExFS[3] := {regex: "`am)\\P\s*"	                                                                                    	, replace:"\b"}
		RegExFS[4] := {regex: "`am)erstellt\sam\s\d\d\.\d\d\.\d\d\d\d\.\s*\,\s*um\s*\d\d\.\d\d\s*Uhr"	, replace:"\b"}
		RegExFS[5] := {regex: "`am)Zeitraum\:\s*Quartal\s\d\/\d\d\s*"                                              	, replace:"\b"}
		RegExFS[6] := {regex: "`am)Ärzte\:\s*.*\;\s*.*\w+\-\w+"                                                          	, replace:"\b"}
		RegExFS[7] := {regex: "`am)\\HZiffernstatistik"                                                                          	, replace:"\b"}
		RegExFS[8] := {regex: "`am)Ausgabe\:"                                                                                  	, replace:"\b"}

		RegExZSData := "\s+(?P<GoNr>\d+)\s+\|\s+(?P<Anzahl>\d+)\s+\|\s+(?P<Punkte>[\d\.]+)\s+\w*\|\s+(?P<Euro>[\d\.]+)\s*\w*\"
		RegExZSData .= "|\s+(?P<Prozent>[\d\.]+)\s*\%*\|\s+(?P<Fachgruppe>[\d\.]+)\s*\%*\|\s+(?P<Abweichung>[\d+\.]+)\s*\%*\|\s+(?P<Leistungstext>[\w\s]+)"
	;}

	; Multiline RegEx - Parsen eines Tagesprotokoll #### FUNKTIONIERT NICHT UM ES NUTZEN ZU KÖNNEN ;{
		Segment:= Object()		; Multi line RegEx
		; findet z.b. Mustermann, Markus (12345) ........ bis zum nächsten Nachnamen, Vornamen (PatientenID)
		Segment.PatIdent               	:= "m)[\w\p{L}]+\,\s[\w\p{L}]+\s\(\d+\)"
		; findet z.B. Spahn, Jens (00001) 30.02.1968, PLEITEKASSE(IK:00700700, VNR: P666999666)
		Segment.Patientenzeile         	:= "is)[\w\p{L}]+\,\s[\w\p{L}]+\s\(\d+\)\s\d+\.\d+\.\d+\,\s[\w\p{L}]+\([\w\s]+\:\d+\,\s\w+\:\s[\w]+\)[\r\n]"
		Segment.Patient                 	:= "is)[\w\p{L}-]+\,\s[\w\p{L}]+\s\(\d+\)\s\d+\.\d+\.\d+\,\s[\w\p{L}\s]+\(IK\:\d+\,\sVNR\:\s[\w]+\)[\r\n]([\w\s\p{L}\[\]\(){}=µ&*!,.+/\--:;\^]+[\r\n])+(?=[\w\p{L}-]+\,\s[\w\p{L}]+\s\(\d+\))"
		; alles unter Dauerdiagnosen
		Segment.Dauerdiagnosen  	:= "is)Dauerdiagnosen\:\s*[\r\n]+((Behandlung\s*[\w\s.,{};\p{L}\[\]-]+[\r\n]+)*|(anamnestisch[\w\s.,{};\p{L}\[\]-]+[\r\n]+)*|(\s+[\w\s,{};.\p{L}]+[\r\n]+)*)+"
		Segment.Dauermedikamente	:= ""
		Segment.Behandlungstag   	:= ""
	;}

;}

;-------------------------------------------------------------------------------------------------------------------------------------------------------------
;  ABRECHNUNGSHELFER - GUI ---- Labels und alle Funktionen
;-------------------------------------------------------------------------------------------------------------------------------------------------------------;{

	;-----------------------------------------------------------------------------------------------------------------------------------------------
	;  EINIGE NOTWENDIGE VARIABLEN FÜR DIE GUI                                                                                                 	;{
		global Protokolle, rxTP, hAbrH

		TVParent    	:= Object()
		TPTV         	:= []
		TVChecked	:= false
		TVExpand  	:= false

		MinW	:= 1360	; minimale Gui Breite
		MinH  	:= 370 	; minimale Gui Höhe
		xMargin:= 10
		yMargin:= 5

		; Info-Bereich (Progress Texte)
		global txtLines	:= 14
		;MsgBox, % screenDims(1).DPI / 96
		; Default Breiten bestimmter Steuerelemente
        wMW	:= 320 	; Breite des Makroauswahl-Elementes (listview)
        wProto	:= 150 	; Breite des Protokollauswahl-Elementes (treeview)
		wInfo	:= 380 	; Breite der Infogruppe (mehrzeilige Textanzeige)
		hMPA	:= MinH	; Anfangshöhe aller 3 oberen Elemente -- wird später berechnet

		wMinus := 10

		; Berechnung der Breite des Auswertungselementes und x Position der Infogruppe
		If InStr(guipos, "AutoSize") {
			wAusw	:= 330                                             	; Breite des Auswertung-Elementes (edit)
			xInfo  	:= wMW + 5 + wProto + 5 + wAusw	; x Position der Infogruppe
		} 	else 	{
			RegExMatch(guipos, "i)\sw(?<W>\d+)\s+h(?<H>\d+)"	, rxTP)
			xInfo 	:= rxTPW - wInfo - xMargin
			wAusw	:= xInfo - wMinus - wProto - 5 - wMW - 5 - 2 * xMargin
		}

		; Format der Variable aktQuartal: [QQJJ] z.B. 0319
		RegExMatch(Tagesprotokolle[Tagesprotokolle.MaxIndex()].filename, "\d\d(?<_Jahr>\d\d)\s*\-\s*(?<_Quartal>\d+)", letztes_Abrechnungs)
		aktQuartal := "0" letztes_Abrechnungs_Quartal letztes_Abrechnungs_Jahr

		;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	;  GUI                                                                                                                                                                      ;{
		Gui, rxTP: New, +OwnDialogs HWNDhAbrH +MinSize%MinW%x%MinH%  ; +AlwaysOnTop +Resize -DPIScale
		Gui, rxTP: Margin, % xMargin, % yMargin
		Gui, rxTP: Color, % "c" BgColor

		Gui, rxTP: Font, % DPI("s" (FontSize + 3) " c" FontColor " q5"), % Font
		Gui, rxTP: Add, Text, % DPI("xm+5 ym w" (wMW + 5 + wProto + 5 + wAusw +5 ) " BackgroundTrans Center +0x200 vT1"), % "Wähle links das gewünschte Makro und rechts die auszuwertenden Quartale"
		Gui, rxTP: Add, Progress, % DPI("xm y" CPoser("rxTP", "T1", "B4") " w" (wMW + 5 + wProto + 5 + wAusw ) " h" 22 " c" BgColor2 " vrxPrg1 HWNDhrxTPPrg1"), % 100

		Gui, rxTP: Font, % DPI("s" (FontSize - 2) " c" FontColor " q5 Normal"), % Font
		Gui, rxTP: Add, Text, % DPI("x" CPoser("rxTP", "rxPrg1", "E-30") " y3 BackgroundTrans Right +Wrap vT1a"), % "aktuelles`nQuartal"
		Gui, rxTP: Font, % DPI("s" (FontSize + 3) " c" FontColor " q5"), % Font
		Gui, rxTP: Add, Text, % DPI("x" CPoser("rxTP", "T1a", "R8") " y" CPoser("rxTP", "T1a", "T5") " BackgroundTrans Center +0x200 vaktQ"), % SubStr(aktQuartal,1,2) "/" SubStr(aktQuartal,3,2)

	; testweise zur Größenbestimmung, danach werden diese gelöscht
		Gui, rxTP: Font, % DPI("s" (FontSize - 2)	 " Normal cBlack q5")	, % Font
		Gui, rxTp: Add, Button, % "xm ym grxToggleProtokolle vProtokollButton"	, % " alle "
		Gui, rxTp: Add, Button, % "xm ym grxToggleExpand     vExpandButton"  	, % "Aufklappen"
		GuiControlGet, ProtokollButton	, rxTP: Pos, ProtokollButton
		GuiControlGet, ExpandButton	, rxTP: Pos, ExpandButton
		ButtonsH := CPoser("rxTP", "T1", "B25") + 2*yMargin + 5 + ProtokollButtonH + 5 + ExpandButtonH

	;----------------------------------------------------------------------------------------------------------------------------------------------
	;  MAKROAUSWAHL
		Makroauswahl:=""
		For MNumber, Auswahl in Makro
			Makroauswahl.= Auswahl.DispName "|"
		Makroauswahl:= RTrim(Makroauswahl "|")
		Gui, rxTP: Font, % "s" (FontSize - 1) " cBlack q5 ", % Font
		Gui, rxTP: Add, ListBox, % "xm y" CPoser("rxTP", "rxPrg1", "B2") " w" wMW " h" hMPA - ButtonsH " vMakrowahl HWNDhMakrowahl AltSubmit", % Makroauswahl

	;----------------------------------------------------------------------------------------------------------------------------------------------
	;  PROTOKOLLE                                                                                     --- ** NEU ALS TREEVIEW ** - Jahr / Quartale
		Gui, rxTP: Font, % "s" (FontSize - 2) " cBlack q5", % Font
		Gui, rxTP: Add, Treeview, % "x" CPoser("rxTP", "Makrowahl", "R5") " y" CPoser("rxTP", "rxPrg1", "B2") " w" wProto " h" (hMPA - ButtonsH - 10) " vProtokolle grxProtokoll AltSubmit Checked HWNDhProtokolle"
		;SetExplorerTheme(hProtokolle)
		TVNum:= 0
		For index, protokoll in Tagesprotokolle		{

			Jahr  	:= protokoll.Jahr
			Quartal	:= protokoll.Quartal

			If !TPTV[Jahr]
				TPTV[Jahr] := TV_Add(Jahr,, "bold")

			TPTV[Jahr][Quartal]:= TV_Add("Quartal " Quartal, TPTV[Jahr])

		}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	;  AUSWERTUNGEN (vAuswertungen)
		Gui, rxTP: Add, Edit, % "x" CPoser("rxTP", "Protokolle", "R5") " y" CPoser("rxTP", "rxPrg1", "B2") " w" wAusW " h" (hMPA - ButtonsH) " vAuswertungen HWNDhAuswertungen" , Datei

	;----------------------------------------------------------------------------------------------------------------------------------------------
	;  BESCHRIFTUNG
		Gui, rxTP: Font, % DPI("s" FontSize +1  " Bold cBlack q5"), % Font
		Gui, rxTP: Add, Text, % DPI("x" CPoser("rxTP", "Makrowahl"  	, "L0")   	" y" CPoser("rxTP", "rxPrg1", "T2") " w" wMW  	" BackgroundTrans Center vT" txtLines + 3)	, % "Auswertungsmakro's"
		Gui, rxTP: Add, Text, % DPI("x" CPoser("rxTP", "Protokolle"   	, "L0")   	" y" CPoser("rxTP", "rxPrg1", "T2") " w" wProto	" BackgroundTrans Center vT" txtLines + 4)	, % "Tagesprotokolle"
		Gui, rxTP: Add, Text, % DPI("x" CPoser("rxTP", "Auswertungen", "L0") 	" y" CPoser("rxTP", "rxPrg1", "T2") " w" wAusw	" BackgroundTrans Center vT" txtLines + 5)	, % "Informationen"

	;----------------------------------------------------------------------------------------------------------------------------------------------
	;  START BUTTON (vMakroButton, vAQButton, vReloadButton)
		Gui, rxTP: Font, % "s" (FontSize)  	 " Normal cBlack q5"   	, % Font
		Gui, rxTp: Add, Button, % "xm                                             y" CPoser("rxTP", "Makrowahl"    	, "B5") " gRunMacro  vMakroButton"                                    	, % "Makro starten"
		Gui, rxTP: Font, % "s" (FontSize - 2)	 " Normal cBlack q5" 	, % Font
		GuiControl, rxTp: Move, ProtokollButton	, % "x" CPoser("rxTP", "Protokolle" , "L0") " y" CPoser("rxTP", "Makrowahl"    	, "B5")
		GuiControl, rxTp: Move, ExpandButton	, % "x" CPoser("rxTP", "Protokolle" , "L0") " y" CPoser("rxTP", "ProtokollButton" 	, "B5")
		Gui, rxTp: Add, Button, % "x" CPoser("rxTP", "ProtokollButton", "R5")    " y" CPoser("rxTP", "ProtokollButton", "T0")  	" grxAktuellesQuartal    vakQButton"            	, % "akt.Quartal"
		Gui, rxTP: Font, % "s" (FontSize)       " Normal cBlack q5"    	, % Font
		Gui, rxTp: Add, Button, % "x+910                                                      y" CPoser("rxTP", "Makrowahl", "B5")      	" grxTPReload        	vReloadButton"             	, % "Reload Skript"
		GuiControlGet, RLBtn            	, rxTP: Pos, ReloadButton

		;Gui, rxTp: Add, Button, % "xm y" CPoser("rxTP", "Makrowahl", "B10") " gRestartMacro"	, % "letztes Makro wiederholen"

	;----------------------------------------------------------------------------------------------------------------------------------------------
	;  VORSCHAUBEREICH (Fortschritt-Info-Bereich (Progress Texte)) (vrxGroup1)
		Gui, rxTP: Font, % DPI("s" FontSize " Normal cWhite q5"), % Font
		AuswRechts := CPoser("rxTP", "Auswertungen", "R5")
		emptyLine := " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
		Gui, rxTP: Add, Text, % DPI("x" AuswRechts " y" CPoser("rxTP", "T1", "B36") " w" (wInfo-15 ) " BackgroundTrans +0x200 vT3")                                               	, % emptyLine
		Loop, % (txtLines - 1)
		Gui, rxTP: Add, Text, % DPI("x" AuswRechts " y" CPoser("rxTP", "T" (A_Index+2), "B0") " w" (wInfo-15) " BackgroundTrans +0x200 vT" (A_Index+3))                	, % emptyLine
		Gui, rxTp: Add, GroupBox, % DPI("x" (AuswRechts - 10) " y" CPoser("rxTP", "T1", "B20") " h" CPoser("rxTP", "T" txtLines-1 , "B10")  " w" wInfo + 5 " +0x00008000 vrxGroup1")

	;----------------------------------------------------------------------------------------------------------------------------------------------
	;  ANZEIGEN UND CONTROLS AUTOMATISCH SETZEN/ÄNDERN/POSITIONIEREN
		Gui, rxTP: Show, % DPI(guipos) " Hide", Abrechnungshelfer - Addendum für Albis on Windows

		WinGetPos,			,			, GuiW	 , GuiH	, % "ahk_id " hAbrH
		WinGetPos, AlbisX	, AlbisY	, AlbisW , AlbisH, % "ahk_class OptoAppClass"

		GuiControl, rxTP: Move, Makrowahl   	, % "h"	(hMPA	- ButtonsH)
		GuiControl, rxTP: Move, Protokolle     	, % "h"	(hMPA	- ButtonsH - 10)
		GuiControl, rxTP: Move, Auswertungen	, % "w"	(MinW	- wInfo - 20 - wMW - wProto - 65) "h" (hMPA - ButtonsH)
		GuiControl, rxTP: Move, rxPrg1          	, % "w"	(MinW	- wInfo - 20 - 3 * xMargin)
		GuiControl, rxTP: Move, ReloadButton	, % "x"	(GuiW	- RLBtnW - 5)

		;DllCall("UxTheme.dll\SetWindowTheme", "Ptr", hrxTPPrg1, "WStr", "Telegram", "WStr", "")
		WinSet, ExStyle, -0x20000, % "ahk_id " hrxTPPrg1

		Gui, rxTP: Show, % DPI(guipos) " Hide"    	, Abrechnungshelfer - Addendum für Albis on Windows

		GuiControlGet, ExpandButton , rxTP: Pos, ExpandButton
		GuiH	:= CPoser("rxTP", "ExpandButton", "B" yMargin)
		m     	:= GetWindowSpot(hAbrH)
		SetWindowPos(hAbrH, m.X, m.Y, m.W, GuiH)
		Gui, rxTP: Show, , Abrechnungshelfer - Addendum für Albis on Windows

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; LETZTES MAKRO VORWÄHLEN
		If !InStr(lProtokollNr, "ERROR") and (lProtokollNr <> "")
			GuiControl, rxTP: Choose, Makrowahl, % lProtokollNr

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; ZULETZT GENUTZTES PROTOKOLL VORWÄHLEN
		GuiControlGet, MB, rxTP: Pos, MakroButton
		GuiControlGet, T, rxTP: Pos, % "T" txtLines + 3

	; Gui neuzeichnen lassen
		Gui, rxTP: +Resize
		WinSet, Redraw,, % "ahk_id " hAbrH

	; für die Interskript-Kommunikation (Datentausch)
		OnMessage(0x4a, "Receive_WM_COPYDATA")


		gosub, rxToggleProtokolle

return ;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	;  GUI LABELS                                                                                                                                                           ;{
rxProtokoll:            	;{                                                                                 	bearbeitet die Nutzerauswahl innerhalb der Protokoll-Treeview

		Gui, rxTP: Submit, NoHide
		MouseGetPos, mx, my
		TVEvent     	:= A_GuiEvent
		TVEventinfo  	:= A_EventInfo
		TvE:= TV_Element(TVEventInfo)

	; Nutzer hat innerhalb der Treeview einen Klick mit der linken Maustaste gemacht
		If RegExMatch(TVEvent, "^S") || InStr(TVEvent, "Normal") {

	    	; Parent-Häkchen wurde gesetzt, dann alle Childs abhaken und umgekehrt
				If (TvE.isParent && TvE.isChecked)
					For index, child in TvE.Childs
						TV_Modify(child.ID, "Check")
				else If (TvE.isParent && !TvE.isChecked)
					For index, child in TvE.Childs
						TV_Modify(child.ID, "-Check")

			 ; ein Child-Häkchen wurde gesetzt, dann die Parent Checkbox abhaken, ist kein Child mehr abgehakt auch das Parent-Häkchen entfernen
				childsChecked := 0
				If (TvE.isChild && TvE.isChecked)
					TV_Modify(TvE.ParentID, "Check")
				else If (TvE.isChild && !TvE.isChecked)
					For index, child in TvE.Childs
						If child.isChecked
							childsChecked ++

		TV_GetText(ts, TvE.ParentID)
		GuiControl, rxTP:, T5, % "ParentID: " TvE.ParentID ", text: " TvE.ParentText ", otext:" ts
		GuiControl, rxTP:, T6, % "isChild: " (TvE.isChild ? "true":"false")
		GuiControl, rxTP:, T7, % "isChecked: " (TvE.isChecked ? "true":"false")

				If (childsChecked = 0)
					TV_Modify(TvE.ParentID, "-Check")
		}

		Protokolle:= ProtokollAuswahl("rxTP", "Protokolle")
		GuiControl, rxTP:, T15, % "Anzahl ausgewählter Protokolle: " (Protokolle.MaxIndex() >=1 ? Protokolle.MaxIndex() : "0")

	/* Info's für Treeview Funktionskontrolle

	GuiControl, rxTP:, T11, % "Child-Anzahl: " TvE.Childs.MaxIndex()
	GuiControl, rxTP:, T12, % "Child Checkboxes: " (childIsChecked) ;= false ? "false":"true")
	GuiControl, rxTP:, T13, % "Anzahl ausgewählter Protokolle: " (Protokolle.MaxIndex() >=1 ? Protokolle.MaxIndex() : "0")
	GuiControl, rxTP:, T3, % "is parent: " (TvE.isParent ? "Yes" : "No") ", is child: " (TvE.isChild ? "Yes" : "No")
	GuiControl, rxTP:, T4, % "is checked: " (TvE.isChecked ? "Yes" : "No")
	GuiControl, rxTP:, T5, % "ElementID: " TVEventInfo
	GuiControl, rxTP:, T6, % "ParentID: " TvE.ParentID
	GuiControl, rxTP:, T7, % "FirstChildID: " TvE.FirstChildID
	GuiControl, rxTP:, T8, % "childs are checked: " (childIsChecked=1 ? "Yes" : "No")

	 */

return ;}

rxToggleProtokolle: 	;{                                                                                 	wählt alle Protokolle aus oder ab

	TVChecked	:= !TVChecked
	Option      	:= TVChecked ? "Check" : "-Check"
	TVID         	:= TV_GetNext()

	If TVChecked
		GuiControl, rxTp:, ProtokollButton, % "keine"
	else
		GuiControl, rxTp:, ProtokollButton, % "alle"

	Loop	{
		TV_Modify(TVID, Option)
		TvE 	:= TV_Element(TVID)
		For index, child in TvE.Childs
			TV_Modify(child.ID, Option)
		TVID := TV_GetNext(TVID)
	} until !TVID


	Protokolle:= ProtokollAuswahl("rxTP", "Protokolle")
	GuiControl, rxTP:, T15, % "Anzahl ausgewählter Protokolle: " (Protokolle.MaxIndex() >=1 ? Protokolle.MaxIndex() : "0")

return ;}

rxToggleExpand:   	;{                                                                                 	entfaltet oder faltet alle Treeview-Abschnitte

	TVExpand	:= !TVExpand
	TVID        	:= TV_GetNext()

	If TVExpand
		GuiControl, rxTp:, ExpandButton, % "Einklappen"
	else
		GuiControl, rxTp:, ExpandButton, % "Ausklappen"

	Loop, % TV_GetCount()	{
		If TVExpand && TV_GetChild(TVID)
			TV_Modify(TVID, "Expand")
		else if !TVExpand && TV_GetChild(TVID)
			TV_Modify(TVID, "-Expand")
		TVID := TV_GetNext(TVID)
	}

	Protokolle:= ProtokollAuswahl("rxTP", "Protokolle")
	GuiControl, rxTP:, T15, % "Anzahl ausgewählter Protokolle: " (Protokolle.MaxIndex() >=1 ? Protokolle.MaxIndex() : "0")

return ;}

rxTPReload:           	;{                                                                                  	das Abrechnungshelferskript neu starten
	a:= GetWindowSpot(hAbrH)
	factor := (96 / screenDims(1).DPI)
	a:= GetWindowSpot(hAbrH)
	IniWrite, % a.X "|" a.Y "|" a.CW - 21 "|" a.CH,  % AddendumDir "\Addendum.ini", % compname, Abrechnungshelfer_Position
	Reload
return ;}

rxMakroWahl:        	;{

return ;}

rxAktuellesQuartal:	;{                                                                                 	wählt nur das aktuelle Quartal zur Bearbeitung aus

		aktQuartal	:= GetQuartal("heute")
		sJahr        	:= "20" SubStr(aktQuartal, 3, 2)
		sTVText     	:= "Quartal " SubStr(aktQuartal, 1, 2)

		TVID         	:= TV_GetNext()
		Option      	:= "-Check"

		Loop {

				TvE 	:= TV_Element(TVID)                    	; Treeview-Element Objekt

				If !InStr(TvE.Text, sJahr)	{
					TV_Modify(TVID, Option)
					TV_Modify(TVID, "-Expand")
				} else {
					TV_Modify(TVID, "Check")
					TV_Modify(TVID, "Expand")
				}

				If TvE.isParent
					ParentText:= TvE.Text

				For index, child in TvE.Childs
				{
						If (InStr(ParentText, sJahr) && InStr(child.Text, sTVText))
							TV_Modify(child.ID, "Check")
						else
							TV_Modify(child.ID, Option)
				}

				TVID := TV_GetNext(TVID)

		} until !TVID

		Protokolle:= ProtokollAuswahl("rxTP", "Protokolle")
		GuiControl, rxTP:, T15, % "Anzahl ausgewählter Protokolle: " (Protokolle.MaxIndex() >=1 ? Protokolle.MaxIndex() : "0")

return ;}

Hotkeys:                	;{                                                                                    	hier ist der Neustart-Hotkey versteckt
^#!a::Reload
return ;}

RunMacro:            	;{                                                                                  	ein vorgefertigtes Makro neu starten

	Gui, rxTp: Submit, NoHide

	Protokolle	:= ProtokollAuswahl("rxTP", "Protokolle")
	tprotfile 	:= TagesprotokollPfad "\" Protokolle[1] "_TP-AbrH.txt"

	GuiControl, rxTP:, T15, % "Anzahl ausgewählter Protokolle: " 	(Protokolle.MaxIndex() >=1 ? Protokolle.MaxIndex() : "0")
	GuiControl, rxTP:, T16, % "erstes ausgewähltes Protokoll: "  	Protokolle[1]

	If Makro[Makrowahl].Parameter {

		If !Protokolle.MaxIndex()
				return
			/*
		;~ Gui, rxTp: ListView, Protokolle
		;~ protokollNr	:= LV_GetNext(0,"F")
		;~ tprotname 	:= Tagesprotokolle[protokollNr].filename
		;~ tprotfile			:= RTrim(Tagesprotokolle[protokollNr].path "\" tprotname, "\")
		;ToolTip, % A_GuiEvent "`n" A_GuiControl "`n" Makrowahl "`n" protokollNr "`n" tprotfile, 1, -200, 4
		;IniWrite, % tprotname	, % AddendumDir . "\Addendum.ini", Abrechnungshelfer,  letzte_Tagesprotokolldatei
		 */
	}

	IniWrite, % Makrowahl  	, % AddendumDir "\Addendum.ini", Abrechnungshelfer,  letzte_Makrofunktion
	gosub % Makro[Makrowahl].Label

return ;}

RestartMacro:        	;{                                                                                 	ein Makro erneut starten - UNBENUTZT!

return ;}

rxTpGuiSize:          	;{                                                                                     Steuerelemente bei Änderung der Fenstergröße neuzeichnen
	Critical, Off	; erst Critical Off soll Critical On dann schneller machen oder zuverlässiger machen (hab vergessen woher ich das habe)
	if A_EventInfo = 1 || !WinExist(hAbrH) ; Das Fenster wurde minimiert.  Keine Aktion notwendig.
		return
	Critical
	rxTPW := A_GuiWidth, rxTPH:= A_GuiHeight, prevX:= rxTPW - wInfo - xMargin
	wAusw	:= prevX - wMinus - wProto - 5 - wMW - 5 - 2*xMargin
	hAusw	:= rxTPH - ButtonsH
	GuiControl, rxTP: -Redraw, Makrowahl
	GuiControl, rxTP: -Redraw, Protokolle
	Loop, % txtLines + 1
		GuiControl, rxTP: Move, % "T" (A_Index + 1 ), % DPI("x" prevX)
	GuiControl, rxTP: Move, rxGroup1            	, % DPI("x" 	prevX - 10)
	GuiControl, rxTP: Move, rxPrg1                 	, % DPI("w"	prevX - xMargin - 21)
	GuiControl, rxTP: Move, Makrowahl        	, % DPI("h" 	hAusw)
	GuiControl, rxTP: Move, Protokolle         	, % DPI("h" 	hAusw)
	GuiControl, rxTP: Move, Auswertungen    	, % DPI("w" 	wAusw " h" hAusw)
	GuiControl, rxTP: Move, MakroButton     	, % DPI("y" 	(MBy := CPoser("rxTP", "Makrowahl"  	  , "B5")))
	GuiControl, rxTP: Move, ProtokollButton 	, % DPI("x" 	CPoser("rxTP", "Protokolle"    	  , "L0")	" y" MBy)
	GuiControl, rxTP: Move, ExpandButton    	, % DPI("x" 	CPoser("rxTP", "ProtokollButton", "L0")	" y" CPoser("rxTP", "ProtokollButton"	, "B5"))
	GuiControl, rxTP: Move, akQButton        	, % DPI("x" 	CPoser("rxTP", "ProtokollButton", "R5")	" y" MBy)
	GuiControl, rxTP: Move, ReloadButton    	, % DPI("x" 	(rxTPW - RLBtnW - xMargin - 5)         	" y" MBy)
	;GuiControl, rxTP: Move, % "T" txtLines+5	, % "w" 	wAusw
	Critical, Off
	SetTimer, rxTPGuiRedraw, -100
return
rxTPGuiRedraw:
	GuiControlGet, MWahl, rxTP: Pos, Makrowahl
	GuiControl, rxTP: MoveDraw, Protokolle         	, % "h" 	MWahlH
	GuiControl, rxTP: MoveDraw, Auswertungen    	, % "w" 	wAusw " h" MWahlH
	GuiControl, rxTP: +Redraw, Makrowahl
	GuiControl, rxTP: +Redraw, Protokolle
	WinSet, Redraw,, % "ahk_id " hAbrH
return ;}

rxTPGuiClose:       	;{                                                                                 	Speichern von Einstellungen wenn der Abrechnungshelfer beendet wird
rxTPGuiEscape:
	;WinGetPos, GuiX, GuiY, GuiW, GuiH, % "ahk_id " hAbrH
	factor := (96 / screenDims(1).DPI)
	a:= GetWindowSpot(hAbrH)
	IniWrite, % a.X "|" a.Y "|" a.CW - 21 "|" a.CH,  % AddendumDir "\Addendum.ini", % compname, Abrechnungshelfer_Position
ExitApp ;}

;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	;  GUI FUNKTIONEN                                                                                                                                                 ;{
UpdateTagesprotokolle(Tagesprotokolle) {                                                      	;-- frischt die Anzeige der Tagesprotokolle auf

	Gui, rxTP: TreeView, Protokolle
	TV_Delete()
	Sleep, 1000
	For index, protokoll in Tagesprotokolle
	{
				Jahr  	:= protokoll.Jahr
				Quartal	:= protokoll.Quartal
				If !TPTV[Jahr]
					TPTV[Jahr]     	:= TV_Add(Jahr,, "bold")

				TPTV[Jahr][Quartal]	:= TV_Add("Quartal " Quartal, TPTV[Jahr])
	}
}

CPoser(guiname, distControl, dist:="B10") {                                                   	;-- zum Positionieren von Gui-Controls

	;L - Left, R - Right, B - Bottom , T - Top : use "B+10" for a positing under your prefered control with an extra distance of 10 pixel
	;only one direction per function call!

	RegExMatch(dist, "\-*\d+"	, inc)
	RegExMatch(dist, "\w"     	, dir)

	GuiControlGet, pos, %guiname%: Pos, %distControl%
	;MsgBox, % inc "`n" dir "`n" posX "`n" posY "`n" posW "`n" posH

	If InStr(dir, "B")
		return posY + posH + inc
	else if Instr(dir, "T")
		return posY + inc
	else if Instr(dir, "R")
		return posX + posW + inc
	else if Instr(dir, "L")
		return posX - inc
	else if InStr(Dir, "E") ; end
		return posX + posW - inc
	else if InStr(Dir, "B") ; beginning
		return posX + inc
}

TV_Element(TVID) {                                                                                          	;-- returns an object with reference data of the treeview element

	/*   Function description -- TV_Element() --
	-------------------------------------------------------------------------------------------------------------------
	Description     	:	returns an object with reference data of the Treeview element
								This function uses existing Autohotkey functions to determine the structural
								references of a treeview element.
								.Text         	: text/name of the element
								.isChild     	: true if it is a child, false if it is not
								.ParentID  	: the ID number of the parent element (only if .isChild is true)
								.ParentText	: the text/name of the parent element (only if .isChild is true)
								.isParent    	: true if it is a parent, false if it is not
								.Childs      	: enumerated Object() - containing data of childs
													  .ID          	- the element ID
													  .Text       	- displayed text of this child
													  .isChecked	- true if checked
													  .isBold    	- true if bold
													  .isParent   	- true if containing child elements
								.isChecked	: true if Checkbox is set
								.isBold      	: true if bold
								.NextID     	: ID number of the next item below (or 0 if none)
								.NextText   	: text/name of the next item below (or empty if none)
								.PrevID      	: ID number of the sibling above (or 0 if none)
								.PrevText   	: text/name of the sibling above (or empty if none)
	Link               	:
	Author           	:	Ixiko
	Date             	:	13.12.2019
	AHK-Version 	:	tested on AHK_L & AHK_H v1.1.30.03
	License          	:	unlicensed
	Syntax           	:	TV_Element(handle of the specified the element)
	Parameter(s)  	:	TVID = handle of the specified the element
	Return value   	:	key-value-Object
	Remark(s)      	:	does not make a complete tree traversal!
	Dependencies	:	none
	KeyWords     	:	gui, treeview
	-------------------------------------------------------------------------------------------------------------------
	|	EXAMPLE(s)
	-------------------------------------------------------------------------------------------------------------------
	*/

	TvE           	:= Object() ; means TvElement
	TvE.Childs 	:= Object()

	TV_GetText(Text	, TVID)
	TvE.Text         	:= Text
	TvE.isChild    	:= TV_GetParent(TVID)	? true : false
	TvE.isParent    	:= TV_GetChild(TVID) 	? true : false
	TvE.isChecked	:= TV_Get(TVID, "C") 	? true : false
	TvE.isBold        	:= TV_Get(TVID, "B")  	? true : false
	TvE.NextID       	:= TV_GetNext(TVID)
	TvE.PrevID       	:= TV_GetPrev(TVID)

	If TvE.isChild     	{
		TvE.ParentID   	:= TV_GetParent(TVID)
		TV_GetText(ParentText, TvE.ParentID)
		TvE.ParentText 	:= ParentText
		TvE.FirstChildID:= TV_GetChild(TvE.ParentID)
	}
	If TvE.isParent    	{
		TvE.FirstChildID:= TV_GetChild(TVID)
	}
	If TvE.FirstChildID	{
		TV_GetText(Text, childID:= TvE.FirstChildID)
		TvE.Childs.Push({"ID"            	: childID
								 , "Text"        	: Text
								 , "isChecked": (TV_Get(childID, "C")	? true : false)
							 	 , "isBold"     	: (TV_Get(childID, "B") 	? true : false)
								 , "isParent"   	: (TV_GetChild(childID)	? true : false)})
		While, (childID:= TV_GetNext(childID)){
			TV_GetText(Text, childID)
			TvE.Childs.Push({	"ID"            	: childID
									, 	"Text"        	: Text
									, 	"isChecked"	: (TV_Get(childID, "C")	? true : false)
									,	"isBold"     	: (TV_Get(childID, "B") 	? true : false)
									,	"isParent"   	: (TV_GetChild(childID)	? true : false)})
		}
	}
	If TvE.NextID     	{
			TV_GetText(Text, TvE.NextID)
			TvE.NextText:= Text
	}
	If TvE.PrevID      	{
			TV_GetText(Text, TvE.PrevID)
			TvE.PrevText:= Text
	}

return TvE
}

ProtokollAuswahl(GuiName, TreeViewName) {                                               	;-- erstellt ein Object das alle ausgewählten Protokolle enthält

	TP := Object()
	Gui, % GuiName ": Default"
	Gui, % GuiName ": TreeView", % TreeViewName

	; first element in treeview
	TVID	 := TV_GetNext()
	while % (TVID := TV_GetNext(TVID, "C")) 	{
		If !TV_GetChild(TVID) {
				TV_GetText(ChildText	 ,	TVID)
				TV_GetText(ParentText, 	TV_GetParent(TVID))
				RegExMatch(ChildText, "Quartal\s+\d(?<Nr>\d)", Q)
				TP.Push(ParentText "-" QNr)
		}
	}

return TP
}

Fortschrittsanzeige(cmd) {

	Gui, rxTP: Default

	If RegExMatch(cmd, "i)(löschen)|(refresh)|(erase)|(leeren)")
		Loop % TextLines - 1
			GuiControl, rxTP: , % "T" A_Index + 3
}

QuartalsKalender(Quartal, Jahr, Urlaubstage) {                                                	;~~~ zeigt den Kalender des Quartals


	; blendet Feiertage nach Bundesland aus

	;--------------------------------------------------------------------------------------------------------------------
	; Variablen
	;--------------------------------------------------------------------------------------------------------------------;{
		static Cok, Can, Hinweis, MyCalendar, hCal, CalBorder, FormularT
		static feldB:= Object()
		global newCal, hmyCal
		FormularT	:= ""
		feldB:= feld

	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	Berechnung der anzuzeigenden Monate
	;--------------------------------------------------------------------------------------------------------------------;{

		Q := Array()
		lastmonth := (Quartal - 1) * 3 + 3
		;~ Loop 3 {
			;~ Q[A_Index] := {"start":(Jahr (lastmonth - 2) "01"), "end":()}
		;~ }

		FormatTime, begindate, % year . month . day, yyyyMMdd
		FormatTime, datum, % datum, yyyyMMdd
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	Kalender Gui
	;--------------------------------------------------------------------------------------------------------------------;{
		Gui, newCal: Destroy
		Gui, newCal: +Owner +AlwaysOnTop HWNDhCal

	;-:verschiedene Darstellungen je nach geöffnetem Formular
		Loop, % Formular.MaxIndex()
			If InStr(feldB.WinTitle, Formular[A_Index])							;der Hinweis wird nur angezeigt, wenn Anfangs- und Enddatum gewählt werden können
			{
				Gui, newCal: Add, Text, vHinweis Center, (halte Shift um Anfangs-und Enddatum zu wählen)
				Gui, newCal: Add, MonthCal, % "vMyCalendar Multi R3 W-3 4 HWNDhmyCal"	, % begindate
				FormularT:= Formular[A_Index]
				break
			}

		If FormularT = ""
			Gui, newCal: Add, MonthCal, % "vMyCalendar R3 W-3 4 HWNDhmyCal"        	, % begindate

		GuiControl, newCal:, MyCalendar, % datum

		Gui, newCal: Add, Button, default      	gCalendarOK    	vCok, Datum übernehmen
		Gui, newCal: Add, Button, xp+200 yp 	gCalendarCancel	vCan, Abbrechen

		Gui, newCal: Show, AutoSize Hide , Addendum für Albis on Windows - erweiterter Kalender

	;-:Größen ermitteln
		WinA       	:= GetWindowInfo(hCal)
		WinB       	:= GetWindowInfo(feldB.hWin)
		GuiControlGet, MyCal, newCal: Pos, MyCalendar
		GuiControlGet, Cok, newCal: Pos
		GuiControlGet, Can, newCal: Pos
		ControlGetPos, feldX, feldY, feldW, feldH,, % "ahk_id " feld.hwnd
		feldX += WinB.WindowX
		SendMessage 0x101F, 0, 0,, % "ahk_id " hmyCal ; MCM_GETCALENDARBORDER
		CalBorder:= ErrorLevel

	;-:Position von Gui-Elementen anpassen
		GuiControl, newCal: MoveDraw, Hinweis, % "x" MyCalX " w" MyCalW
		GuiControl, newCal: MoveDraw, Cok, % "x" (MyCalX + 2*CalBorder)
		GuiControl, newCal: MoveDraw, Can, % "x" (MyCalX + MyCalW - CanW - 2*CalBorder)

		If InStr(FormularT, "Muster 12a")
				newCalX:= feldX + feldW//2 - WinA.windowW//2, newCalY:= feldY + feldH + 10
		else if InStr(FormularT, "Muster 1a")
				newCalX:= WinB.WindowX + WinB.windowW + 5	, newCalY:= WinB.WindowY + WinB.windowH//2 - WinA.windowH//2
		else if InStr(FormularT, "Muster 20") || InStr(FormularT, "Muster 1a")
				newCalX:= WinB.WindowX - WinA.windowW - 5 	, newCalY:= WinB.WindowY + WinB.windowH//2 - WinA.windowH//2
		else
				newCalX := feldX + feldW + 10, newCalY:= feldY - 10

	;-:wenn Gui die Bildschirmhöhe übersteigt, muss die y-Position angepasst werden
		newCalY := newCalY + WinA.windowH > A_ScreenHeight ? A_ScreenHeight - WinA.windowH - 30 : newCalY

		Gui, newCal: Show, % "x" newCalX " y" newCalY , Addendum für Albis on Windows - erweiterter Kalender

Return ;}

CalendarOK: 				                                                                                                              	 ;{
	Gui, newCal: Submit
	Gui, newCal: Destroy

	If !InStr(MyCalendar,"-")  {                                                                                               	; just in case we remove MULTI option
		FormatTime, CalendarM1, %MyCalendar%, dd.MM.yyyy
		CalendarM2:= CalendarM1
	}
	Else	{
		 FormatTime, CalendarM1, % StrSplit(MyCalendar,"-").1, dd.MM.yyyy
		 FormatTime, CalendarM2, % StrSplit(MyCalendar,"-").2, dd.MM.yyyy
	}

	If InStr(FormularT, "Muster 1a") && (CalendarM1 != CalendarM2)	{
		ControlSetText, Edit1, %CalendarM1%, % "ahk_id " feldB.hWin
		ControlSetText, Edit2, %CalendarM2%, % "ahk_id " feldB.hWin
	}
	else If (InStr(FormularT, "Muster 12a") || InStr(FormularT, "Muster 21")) && (CalendarM1 != CalendarM2)	{
		ControlSetText, Edit6, %CalendarM1%, % "ahk_id " feldB.hWin
		ControlSetText, Edit7, %CalendarM2%, % "ahk_id " feldB.hWin
	}
	else if InStr(FormularT, "Muster 20") && (CalendarM1 != CalendarM2)	{
		editgroup   	:= []
		editgroup[1]	:= [3,4]
		editgroup[2]	:= [7,8]
		editgroup[3]	:= [11,12]
		editgroup[4]	:= [15,16]
		Loop, % editgroup.MaxIndex()
			If (feldB.classnn = editgroup[A_Index, 1]) || (feldB.classnn = editgroup[A_Index, 2])	{
				ControlSetText, % "Edit" editgroup[A_Index, 1] 	 , %CalendarM1%, % "ahk_id " feldB.hWin
				ControlSetText, % "Edit" editgroup[A_Index, 2] 	 , %CalendarM2%, % "ahk_id " feldB.hWin
				ControlSetText, % "Edit" editgroup[A_Index+1, 1] , %CalendarM1%, % "ahk_id " feldB.hWin
				break
			}
	}
	else
		ControlSetText,, %CalendarM1%, % "ahk_id " feldB.hwnd

	CalendarM1:=CalendarM2:=0

CalendarCancel:
newCalGuiClose:
newCalGuiEscape:
	Sleep 150
	Gui, newCal: Destroy
	WinActivate, % "ahk_id " feldB.hactive
Return
;}

}


;}

;}

;-------------------------------------------------------------------------------------------------------------------------------------------------------------
; MAKRO's -
;-------------------------------------------------------------------------------------------------------------------------------------------------------------;{

;weitere Ideen
; Eintragungen in Akten ohne abgerechnete Ziffern
;	labor - Urin - ohne Ziffer 32030 am selben Tag
;   Kuerzel: mindestens 1 von diesen Kuerzeln:
;			anam, bef, fkh, fhv13 (Erstverordnung), ther - ohne 03230 am selbsten Tag wenn nicht schon 0300x abgerechnet wurde
;   Kuerzel: mindestens 2 von diesen Kuerzeln:
;			 fausw, fau, RR, info (mehr als 10 Zeichen) - ohne 03230 am selbsten Tag wenn nicht schon 0300x abgerechnet wurde
; Nutzer: Beschreibungen der Kürzel anlegen lassen für angenehmeres Suchen

M1:                                                  	;	Tagesprotokoll konvertieren                   	;{

		Gui, rxTP: Default

	; Hinweisdialog anzeigen
		MsgBox, 262144 , % "Addendum - " A_ScriptName,
		(LTrim
			!HINWEIS!  !HINWEIS!  !HINWEIS!  !HINWEIS!  !HINWEIS!  !HINWEIS!

		Dieses Makro bereitet ein in Albis auf die Festplatte gespeichertes
		Tagesprotokoll für die Verarbeitung durch die Skriptmakros vor.
		Wählen Sie im folgenden Dateidialog das gewünschte Protokoll aus.
		Nach Umwandlung wird dieses in folgendem Ordner gespeichert:

		"%TagesprotokollPfad%"
		)

	; Dateidialog anzeigen
		FileSelectFile, tprotfile,, % TagesprotokollPfad, Tagesprotoll auswählen ;{
		If (StrLen(tprotFile) = 0)		{
			MsgBox,, Addendum für Albis on Windows - Abrechnungshelfer,
			(LTrim
			Es wurde keine Datei ausgewählt!
			Die Konvertierung wird abgebrochen!
			), 3
			GuiControl,rxTP:, T3, % "Es wurde keine Datei zum Konvertieren ausgewählt."
			return
		}
		SplitPath, tprotfile, tprotfilename,,, tprotfilenameNoExt
		tpFile := FileOpen(tprotfile, "r").Read()

		If !InStr(tpfile, "\HTagesprotokoll") {
			MsgBox, 262144 , % "Addendum - " A_ScriptName,
			(LTrim
							!!! KEIN TAGESPROTOKOLL !!!

			Die ausgewählte Datei ist kein Tagesprotokoll.
			Tagesprotokolle werden über das Albismenu:
			Statistik/Tagesprotokoll... erstellt.
			Speichern Sie das Protokoll über das Menu:
			Patient/Speichern unter... oder drücken Sie
			Strg+U.
			), 3
			tpfile := ""
			GuiControl,, T3, % "Datei: " tprotfilenameNoExt " ist kein Tagesprotokoll!"
			return
		}
		;}

	; Datei parsen
		tpfile:= RegExParseTPFile(tprotfile, TagesprotokollPfad, RegExTP)

	; Listviewinhalt auffrischen
		Tagesprotokolle:= ReadTPDir(TagesprotokollPfad, "*_TP-AbrH", "txt", "R")
		UpdateTagesProtokolle(Tagesprotokolle)

	; Albis Tagesprotokoll löschen
		FileDelete, % tprotfile

	; Fertigstellung anzeigen
		Gui, rxTP: Default
		GuiControl,, T3, % "Tagesprotokoll: " tprotfilenameNoExt " wurde erstellt"
		GuiControl,, T4, % "Das Original des Tagesprotokoll wurde gelöscht."
		MsgBox,, Addendum für Albis on Windows,
		(LTrim
		Das Tagesprotokoll ist bearbeitet und wurde gespeichert.
		Das Original >Albis< Tagesprotokoll wurde gelöscht.
		), 6

return
;}
M2:                                                  	;	freie Statistik                                          	;{

	;Variablen
		Suchmuster	:= []
		regExFound	:= 0
		ind				:= 0
		Patidx			:= 0
		; dia = \{F\d+\.* | - lko = 351\d\d

		;GuiFreieStatistik()
		InputBox, rxStr, Abrechnungshelfer, Hier eingeben wonach Sie suchen:,, 600, 300


		return
	; im Moment noch eingeschränkte Suchfunktion, sobald einmalig die Suche erfolgreich war wird mit dem nächsten Patienten weiter gemacht!!
		IniRead, tpfilepath,  % AddendumDir "\Addendum.ini", Abrechnungshelfer, letzte_Tagesprotokolldatei
		If InStr(tpfilepath, "Error") || !tpfilepath {
			FileSelectFile, tpfilepath,, % AddendumDir "\Tagesprotokolle\" A_YYYY, Tagesprotokoll auswählen
			IniWrite, % tpfilepath,  % AddendumDir "\Addendum.ini", Abrechnungshelfer, letzte_Tagesprotokolldatei
		}

	; Tagesprotokoll wird eingelesen
		FileRead, tpfile, % tpfilepath

	; Auslesen aller gespeicherten Suchmuster
		Suchmuster:= ReadSearchPatterns("freieStatistik")

	; Gui für Regex wird erstellt
		;Ergebnisse:= GuiFreieStatistik()


	;Beispiele:
	;
	; i)lko.*0322. - findet alle Patienten mit Ziffer 03220 oder 03221
	;regExStr		:= "i)" . RegExReplace(Suchstring, "\s+", "`.`*")

	GuiControl,, T3, % "Suchmuster: '" regExStr "'"

	;If FileExist(AddendumDir . "\Tagesprotokolle\" . A_YYYY . "\AbrH-freieStatistik.txt")
	;		FileDelete, % AddendumDir . "\Tagesprotokolle\" . A_YYYY . "\AbrH-freieStatistik.txt"

	File := FileOpen(AddendumDir "\Tagesprotokolle\" A_YYYY "\AbrH-freieStatistik.txt", "w")
	File.Write("Ergebnisse für: " Suchmuster ", Datum: " A_DD . "." . A_MM . "." A_YYYY . ", " . A_Hour . ":" . A_Min . "`n")

	Loop, Parse, tpfile, `n
	{
				If RegExMatch(A_LoopField, RegExFind[1], PatID) 				{

						;MsgBox, % PatIDStored ", " PatID
						If (regExFound = 2) and (A_Index > 1)
							File.Write("kein Treffer bei Patient: " . PatIDStored . ";" . PatName . ";" . PatBirth . "`n")

						PatIDStored:= PatID
						RegExMatch(A_LoopField, RegExFind[3], PatName)
						RegExMatch(A_LoopField, RegExFind[4], PatBirth)
						regExFound:= 0
						PatIdx ++
						GuiControl,, T3, % "akt. Patient   : " PatName
						GuiControl,, T4, % "Patientenzahl: " PatIdx
						continue

				}

				If !regExFound and RegExMatch(A_LoopField, regExStr, fdummy)	{				;"i)lko.*0322"

					regExFound:= 1
					ind++
					GuiControl,, T5, % "Treffer       : " PatName
					GuiControl,, T6, % "Trefferzahl : " ind
					File.Write(PatIDStored . ";" . PatName . ";" . PatBirth . "`n")

				}


	}

	File.Close()

	MsgBox, % ind " Patienten zur Suche gefunden."
	; ^.*[^(lko)].*
return

GuiFreieStatistik() {

	global

	oFSk:= Object()
	oFsk.Controls:= {1:"Name", 2:"Kuerzel,Inhalt"}

	m:=GetWindowSpot(hAbrH)

	Gui, FSk: new, +HWNDhFsk +Owner%hAbrH% -DPIScale +Border
	Gui, Fsk: Color, % "c" BgColor
	Gui, Fsk: Margin, 0 , 0
	Gui, FSk: Add, Progress, % "x-2 y-2 w800 h28 vFskPrgUS c" BgColor2 " HWNDhFskPrgUS" , % 100
	Gui, FSk: Font, % "S13 q5 c" BGColor, Futura Md Bt
	Gui, FSk: Add, Text, x10 y1 w800 BackgroundTrans Center vFskUeberschrift, % "≡≡≡  ≡≡≡  DIE NEUE FREIE STATISTIK!  Einfacher und schneller bessere Ergebnisse erzielen!  ≡≡≡  ≡≡≡"

	Gui, FSk: Font, S10 q5 cWhite, Futura Bk Bt
	Gui, FSk: Add, Text, % "x10 y" CPoser("FSk", "FskPrgUS", "B10") "BackgroundTrans", Optionen


;-: REIHE 1 ---
	Gui, FSk: Font, S10 q5 italic cWhite, Futura Bk Bt
	Gui, FSk: Add, Text, % "x80 y" CPoser("FSk", "FskPrgUS", "B10") " w250 Center", % "Nachname,Vorname"
	Gui, FSk: Add, Text, x+2 w100	    	Center, % "Geburtsdatum"
	Gui, FSk: Add, Text, x+2 w75     		Center, % "Gesch."
	Gui, FSk: Add, Text, x+2 w75    		Center, % "ID"
	Gui, FSk: Add, Text, x+2 w250	    	Center, % "Krankenkasse"
	Gui, FSk: Font, S10 q5 Normal, Futura Bk Bt
	Gui, FSk: Add, Edit, x80 y+0 r1 w250 vName
	Gui, FSk: Add, Edit, x+2 r1 w100 vGeburtsdatum
	Gui, FSk: Add, ComboBox, x+2 r4 w75 vGeschlecht, % "alle|weiblich|männlich|unbestimmt"
	Gui, FSk: Add, Edit, x+2 r1 w75 vPatID
	Gui, FSk: Add, ComboBox, x+2 r6 w250 vKrankenkasse, % "|alle|gesetzl. KK|private KK|AOK|BARMER|TKK"

;-: REIHE 2 --
	Gui, FSk: Font, S10 q5 italic cWhite, Futura Bk Bt
	Gui, FSk: Add, Text, x80 y+5 w150     	Center, % "Altersgrenzen"
	Gui, FSk: Add, Text, x+2  w250         	Center, % "Behandlungszeitraum"
	Gui, FSk: Add, Text, x+2  w250         	Center, % "Arbeitgeber/Beruf"
	Gui, FSk: Add, Text, x+2  w250         	Center, % "caveVon Einträge"
	Gui, FSk: Font, S10 q5 Normal, Futura Bk Bt
	Gui, FSk: Add, ComboBox, x80 y+0 r6 w150 vAltersgrenzen           	, % "|Rentner|Säuglinge|KleinKinder|Kinder|Jugendliche|Erwachsene|10-20 Jahre|20-30 Jahre"
	Gui, FSk: Add, ComboBox, x+2  r6 w250 vBehandlungszeitraum	, % "alle gewählten Quartale|aktuelles Quartal|0117|0118-0318"
	Gui, FSk: Add, Edit, x+2 r1 w250 vArbeitgeber	, % ""
	Gui, FSk: Add, Edit, x+2 r1 w250 caveVon      	, % ""

;-: REIHE 3 --
	Gui, FSk: Font, S10 q5 italic cWhite, Futura Bk Bt
	Gui, FSk: Add, Text, x40  y+5 w200 Center, % "Karteikartenkürzel"
	Gui, FSk: Add, Text, x+0 w200 Center, % "Inhaltssuche"
	Gui, FSk: Font, S10 q5 Normal, Futura Bk Bt
	Gui, FSk: Add, ComboBox, x80 y+0 r6 w200 vKuerzel1, % "anam|bef|info|lp|scan|bild|lko|dia|ther|Dauerdiagnose|medrp|medpn|fhv13|fhv18|"
	Gui, FSk: Add, Edit, x+0 r1 w600 vInhalt1, % ""

	Gui, FSk: Font, S10 q5 Normal, Futura Bk Bt
	Gui, FSk: Add, Button, x80 y+5 vMehr gFSKGuiInteraction, % "▼"

;-: Starten speichern
	Gui, FSk: Add, Progress, % "x70 y+15 w800 h50 vFskPrg1 c" BGColor2 , % 100
	GuiControlGet, FskPrg, Fsk: Pos, FskPrg1
	Gui, FSk: Font, S10 q5 Normal, Futura Bk Bt
	Gui, FSk: Add, Button    	, % "x80 y" FskPrgY + 10 " vFskStarten"    	, % "Suche starten"
	Gui, FSk: Add, Button    	, x+25                       	vFskSpeichern	, % "Eingaben speichern"
	Gui, FSk: Add, Button    	, x+25                      	vFskAbbruch 	, % "Beenden"
	Gui, FSk: Font, S8 q5 Normal italic cBlack, Futura Bk Bt
	Gui, FSk: Add, Text	    	, % "x" CPoser("FSk", "FSkAbbruch", "R35") " y" FskPrgY + 2 " w450 Left Backgroundtrans vFskTSE", % "gespeicherte Sucheingaben"
	Gui, FSk: Add, ComboBox	, % "x" CPoser("FSk", "FSkAbbruch", "R35") " y" CPoser("FSk", "FSkTSE", "B1") " w450 r6 vFskAuswahl " 	, % ""

;-: Fenster versteckt zeichnen, Styles, Positionen der Steuerelemente werden angepasst
	Gui, FSk: Show, % "x" m.X +20  " y" m.Y + 20 " AutoSize Hide", Abrechnungshelfer -  freie Statistik
	;Gui, FSk: Show, % "x" m.X +20  " y" m.Y + 20 " AutoSize Hide", Abrechnungshelfer -  freie Statistik

	WinSet, Style 	, 0x94020000, % "ahk_id " hFsk
	WinSet, ExStyle	, 0x00020100, % "ahk_id " hFsk
	WinGetPos,,, ww, wh, % "ahk_id " hFsk
	GuiControl, Fsk: MoveDraw, FskPrg1         	, % "w" ww-140
	GuiControl, Fsk: MoveDraw, FskUeberschrift	, % "w" ww-10
	GuiControl, Fsk: MoveDraw, FskPrgUS       	, % "w" ww+4
	GuiControl, Fsk: Disable, FskPrg1
	GuiControl, Fsk: Disable, FskPrgUS
	GuiControlGet, p, Fsk: Pos, FskAuswahl
	GuiControlGet, b, Fsk: Pos, FskAbbruch
	;GuiControl, Fsk: MoveDraw, FskAuswahl, % "y" pY + 6 " h" bH + 8

;-: Fenstergröße anpassen
	Gui, FSk: Show, % "x" m.X + Floor(m.W//2) - Floor(ww//2) " y" m.Y + m.H - 10 " h" wh - 20 " Hide"

;-: Fenster mit Animation anzeigen ;{
	FADE     		:= 524288
	SHOW   		:= 131072
	HIDE      		:= 65536
	FADE_SHOW:= GetHex(FADE + SHOW) 	; Converts to 0xa0000
	FADE_HIDE	:= GetHex(FADE + HIDE)     	; Converts to 0x90000
	Duration 		:= 500                             	; Duration of Animation in milliseconds
	AnimateWindow(hFsk, Duration, FADE_SHOW) ;}

	OnMessage(0x201, "FSK_LBUTTONDOWN")

	;Gui, ▲▼▲►◄
	;≡⁞‖⓪①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑰⑳⓫⓬⓭⓮⓯⓱⓲⓳⓴❶❷❸❺❻❼❽❾❿
	;□◎△▽※◎△▽※♫♪♠♣☼♀◎△▽※♫♪♠♣☼♀☺☻◦◙◘●○▬▫▪▓▒░▌
	;OnMessage("0x200", "MouseMove")
	;OnMessage("0x201", "MouseMove")
	;OnMessage("0x202", "MouseMove")
return

FSKGuiInteraction: ;{

		Gui, FSk: Submit, NoHide
		Fc := A_GuiControl

		If RegExMatch(Fcl, "\w+CB(\d+)", used) {

			ctrls:= StrSplit(oFsK.Controls[used1])
			GuiControlGet, ischecked, FSk:, Fcl
			If isChecked
				Option:= "Enable"
			else
				Option:= "Disable"
			Loop, % ctrls.MaxIndex()
				GuiControl, % "FSk: " Option, % ctrls[A_Index]

		}

		If InStr(Fc, "Mehr") && (InStr(A_GuiEvent, "S") || InStr(A_GuiEvent, "Normal")) 		{
			Gui, Fsk: Default
			GuiControlGet, pl, Fsk:Pos, Mehr
			GuiControlGet, pr, Fsk:Pos, FskPrg1
			GuiControlGet, pS, Fsk:Pos, FSKStarten
			GuiControlGet, pA, Fsk:Pos, FskAuswahl
			GuiControlGet, pT, Fsk:Pos, FskTSE
			WinGetPos, wx, wy, ww, wh, % "ahk_id " hFsk
			Loop 	{
				schub := A_Index * 2.6
				;DllCall("Sleep","UInt", 1)
				GuiControl, Fsk: MoveDraw, Mehr                      	, % "y" plY + schub
				GuiControl, Fsk: MoveDraw, FskPrg1                  	, % "y" prY + schub
				GuiControl, Fsk: MoveDraw, FskStarten                	, % "y" pSY + schub
				GuiControl, Fsk: MoveDraw, FskSpeichern           	, % "y" pSY + schub
				GuiControl, Fsk: MoveDraw, FskAbbruch            	, % "y" pSY + schub
				GuiControl, Fsk: MoveDraw, FskAuswahl            	, % "y" pAY + schub
				GuiControl, Fsk: MoveDraw, FskTSE	, % "y" pTY + schub
				Gui, Fsk: Show, % "AutoSize"
			} until schub >=35
			WinSet, Redraw,, % "ahk_id " hFsk
		}

return ;}

}

FSK_LBUTTONDOWN(wparam, lparam, msg, hWMLbd) {				;-- Klick mit der li. Maustaste verschiebt das Fenster
	 global hFsk, hFskPrgUS
	if (hWMLbd in %hFsk%,%hFskPrgUS%)
		PostMessage, 0xA1, 2,,, A 					; WM_NCLBUTTONDOWN
return
}

MouseMove(lParam,wParam){

	static thisTime:= 0
	global hFsk, Fsk, Mehr, FskStarten, FskSpeichern, FskAbbruch, FskTSE, FskPrg1

	If (A_GuiControl="Mehr") && (thisTime = 0)
		thisTime:= A_TickCount

	;ToolTip, % A_GuiEvent

	If (thisTime > 0 )
		If (A_TickCount - thisTime > 1500)		{

			Gui, Fsk: Default
			GuiControlGet, pl, Fsk:Pos, Mehr
			GuiControlGet, pr, Fsk:Pos, FskPrg1
			GuiControlGet, pS, Fsk:Pos, FSKStarten
			GuiControlGet, pA, Fsk:Pos, FskAuswahl
			GuiControlGet, pT, Fsk:Pos, FskTSE
			WinGetPos, wx, wy, ww, wh, % "ahk_id " hFsk
			thisTime:= 0

			Loop 34	{
				;DllCall("Sleep","UInt", 1)
				GuiControl, Fsk: MoveDraw, Mehr                      	, % "y" plY + A_Index
				GuiControl, Fsk: MoveDraw, FskPrg1                  	, % "y" prY + A_Index
				GuiControl, Fsk: MoveDraw, FskStarten                	, % "y" pSY + A_Index
				GuiControl, Fsk: MoveDraw, FskSpeichern           	, % "y" pSY + A_Index
				GuiControl, Fsk: MoveDraw, FskAbbruch            	, % "y" pSY + A_Index
				GuiControl, Fsk: MoveDraw, FskAuswahl            	, % "y" pAY + A_Index
				GuiControl, Fsk: MoveDraw, FskTSE	, % "y" pTY + A_Index
				Gui, Fsk: Show, % "AutoSize"
			}

			WinSet, Redraw,, % "ahk_id " hFsk

		}


	;else
	;	thisTime:= 0

}

;}
M3:                                                  	;	Pat. für GVU suchen                              	;{

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; VARIABLEN
	;-------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		StarteAbPatient  	:= 0                                             	; ab welchem Patient die Suche begonnen oder fortgesetzt werden soll
		MinGVUAbstand	:= 36                                              	; Angabe in Monate (hat sich geändert am 01.01.2020 auf 36 Monate - leider)
		MinAlter            	:= 35                                            	; einmalig bis zum 35.LJ
		MinAlterIntervall	:= 35
		qname                	:= "0420"
		maxMonth         	:= "12"
		maxYear           	:= "17"
		QYear              	:= "20" SubStr(qname, 3, 2)
		LastQMonth      	:= (SubStr(qname, 1, 2) - 1) * 3 + 1 + 2
		LastQDay          	:= DaysInMonth( QYear LastQMonth ) "." SubStr("0" LastQMonth, -1) "." QYear

		tpfilesDir            	:= AddendumDir "\Tagesprotokolle"
		tprotfile             	:= Tagesprotokollpfad "\" QYear "-" SubStr(qname,2,1) "_TP-AbrH.txt"
		qliste                 	:= tpfilesDir "\" qname "-GVU.txt"

		Ignore               	:= MaxLines := MaxPatient := PatIndex := 0
		Vorsorge_IDs     	:= PatIDStored := []
		ScheinVorhanden	:= false
		GVUNaechsteZeile:= false

		PatIndex            	:= 0

		Vorsorgen				:= Object()
	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; VORBEREITUNGEN
	;-------------------------------------------------------------------------------------------------------------------------------------------------------------------;{

		GuiControl, rxTP: , T3, % "Suche nach potentiellen Vorsorgeuntersuchungen"
		GuiControl, rxTP: , T4, % "Quartal: " qname " - letzter Tag: " LastQDay
		GuiControl, rxTP: , T5, % "<><><><><><><><><><><><><><><>"

	; die mit Praxomat angelegte ID Liste der Vorsorgepatienten wird eingelesen (diese Patienten werden nicht vorgeschlagen!)
		Vorsorge_IDs := VorsorgelistenIDs2Arr(qliste)
		GuiControl, rxTP: , T10, % "Gesamtzahl Gesundheitsvorsorge-Untersuchungen: " SubStr("000" Vorsorge_IDs.MaxIndex(), -2)

	; Tagesprotokoll wird in eine Variable eingelesen
		FileRead, tprot, % tprotfile

	; ermitteln der Zeilen- und Patientenanzahl in der Datei
		TP := TPMax(tprot)
	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; DURCHSUCHT DAS AUSGEWÄHLTE TAGESPROTOKOLL
	;-------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		Loop, Parse, tprot, `n, `r
		{
				GuiControl, rxTP: , T6 	, % "Zeilen bearbeitet      : " SubStr("00000" A_Index, -4) "/" SubStr("00000" TP.Lines, -4)

			; ========================================
			; =============   Patienten-ID erkennen    ============= ;{
				If RegExMatch(A_LoopField, RegExFind[1], PatID)	{

						GuiControl, rxTP: , T10	, % "Patienten bearbeitet : " SubStr("0000" PatIndex, -3) "/" SubStr("0000" TP.Patients, -3)

					; vor Wechsel zu anderem Patienten: prüfen welche Daten zum bisherigen Patienten vorhanden sind
						If (PatIndex >= StarteAbPatient)			{

								Gui, rxTP: Default
								GuiControl, rxTP: , T7 	, % "aktuelle Patienten-ID: " PatIDStored ", Alter: " ageStored
								GuiControl, rxTP: , T8 	, % "Schein vorhanden    : " (ScheinVorhanden = true ? "Ja" : "Nein")
								GuiControl, rxTP: , T9 	, % "letzte GVU               : " ((StrLen(year) > 0) ? (SubStr("0" month, -1) "/" SubStr("0" year, -1)) : ("--/--"))
								letzteGVU := StrLen(year) > 0 ? (SubStr("0" month, -1) "/" SubStr("0" year, -1)) : (0)

								If (PatID <> PatIDStored) && ScheinVorhanden	{

										If GVUStored && (ageStored > MinAlterIntervall) {

											GVUAbstand:= (maxMonth + maxYear * 12) - (month + year * 12)

											If (GVUAbstand >= MinGVUAbstand)	{

												For PatientenID, Patient in PatDB
													If (PatientenID = PatIDStored)	{
														Anzeige:= "(" PatIDStored ") " Patient.Nn ", " Patient.Vn " (" Patient.Gt "), " Patient.Gd " (" ageStored " Jahre), letzte GVU: " letzteGVU
														break
													}

											  ; Daten werden in ein Objekt geschrieben
												Vorsorgen.Push({"PatID":PatIDStored, "Alter":ageStored, "Name":Anzeige , "letzteGVU": letzteGVU})
												GuiControl, rxTP: , T11 	, % "Vorschläge: " Vorsorgen.MaxIndex()

											}

										}
										;else if (ageStored >= MinAlter)
										;		gosub GVUPatientVorschlag


								}

							; Patientenalter ausrechnen
								RegExMatch(A_LoopField, RegExFind[4], PatBirth)
								age	:= Floor(DateDiff( "YY", PatBirth, LastQDay))

							; ist der Patient in der GVU-Liste, weiter zum nächsten
								If inArr(Vorsorge_IDs, PatID) || (age < MinAlter)
								{
										GoNextPatient:= true
										continue
								}
						}

					; Patienten-ID sichern, Patienten-Index erhöhen, Variablen zurücksetzen
						PatIDStored := PatID, ageStored := age, GoNextPatient:= false, GVU := GVUStored:= month:= year := ""
						GVUNaechsteZeile:= ScheinVorhanden := false
						PatIndex		++

				}
				else if GoNextPatient
						continue

				;}

			; ========================================
			; ========   findet cave von! Eintragung GVU/HKS......   ======== ;{
				If !GoNextPatient
					If RegExMatch(A_LoopField, "i)G[VU]+.H[KS]+.[KVU]*\s*") || GVUNaechsteZeile		{

						If !RegExMatch(A_LoopField, "i)G[VU]+.H[KS]+.[KVU]*\s*(\d+).(\d+)") && !GVUNaechsteZeile {
							GVUNaechsteZeile := true, GVUAlles := A_LoopField
							continue
						}

						GVUAlles .= A_LoopField

						If RegExMatch(GVUAlles, "i)G[VU]+.H[KS]+.[KVU]*\s*(\d+).(\d+)", match)
							GVUStored := match, month := SubStr("0" match1, -1), year := SubStr("0" match2, -1), GVUAlles := "", GVUNaechsteZeile := false

					}
			;}

			; ========================================
			; ====== prüft auf Vorhandensein eines Abrechnungsscheines ====== ;{
				If !GoNextPatient
					If !ScheinVorhanden && RegExMatch(A_LoopField, RegExFind[5], lko)
							ScheinVorhanden := true
			;}

		}
	;}

		gosub GVUVorschlagGui

return

GVUPatientVorschlag(PatID, Vorsorgen) {

		If AlbisAkteOeffnen(PatID)		{
				MsgBox, 262148, % "Addendum - " A_ScriptName, % "Patient in die GVU Liste aufnehmen?`nID: " PatIDStored "`nlQ: " GVUAlles "`nQ:  " month "/" year "`n`nOk für nächsten Patienten!"
				AlbisAkteSchliessen(PatIDStored)
			; Vorsorgeliste wird aufgefrischt
				Vorsorge_IDs	:= VorsorgelistenIDs2Arr(qliste)
		}
		Gui, GV: Default
		GuiControl, GV: , Patient, % "Gesamtzahl Gesundheitsvorsorge-Untersuchungen: " SubStr("000" Vorsorge_IDs.MaxIndex(), -2)

return
}

GVUVorschlagGui: ;{

		VNr := 1
		a:= GetWindowSpot(AlbisWinID())

		Gui, GV: New, +HWNDhGV +AlwaysOnTop
		Gui, GV: Font, S14 q5, Futura Bk Bt

		;Gui, GV: Add, GroupBox, x10 y0 w105 h35
		Gui, GV: Add, Text, xm ym w100 Left BackgroundTrans                            	, Patient     :

		;Gui, GV: Add, GroupBox, x115 y0 w410 h35
		Gui, GV: Add, Text, x115 ym w600 Center BackgroundTrans vGVPatient

		Gui, GV: Font, S10 q5, Futura Bk Bt

		;Gui, GV: Add, GroupBox, x10 y+0 w105 h30
		Gui, GV: Add, Text, xm y42 w150 Left BackgroundTrans                            	, nächster Patient:

		;Gui, GV: Add, GroupBox, x115 y32 w410 h30
		Gui, GV: Add, Text, x115 y42 w600 Center BackgroundTrans vGVPatientNext

		Gui, GV: Font, S16 q5, Futura Md Bt
		Gui, GV: Add, Button, x220 y65 vGVBackward gGVButtonPress, ⏪
		Gui, GV: Add, Button, x+20 vGVForward gGVButtonPress, ⏩
		Gui, GV: Show, % " AutoSize Hide"
		GuiControl, GV:, GVPatient, % Vorsorgen[VNr].Anzeige
		ToolTip, %  Vorsorgen[VNr].Anzeige

		g:= GetWindowSpot(hGV)
		;Gui, GV: Show, % "x" (a.X + a.W - g.W - 30) " y" (a.Y + a.H - g.H - 30), % "Patienten ohne Vorsorge in den letzten 2 Jahren (Gesamt:" Vorsorgen.MaxIndex() ")"
		Gui, GV: Show, % "x" (500) " y" (800), % "Patienten ohne Vorsorge in den letzten 2 Jahren (Gesamt:" Vorsorgen.MaxIndex() ")"

		AlbisAkteOeffnen(Vorsorgen[VNr].PatID)

return

GVButtonPress: ;{

	If InStr(A_GuiControl, "GVForward") && (VNr <= Vorsorgen.MaxIndex())
	{
			AlbisAkteSchliessen(Vorsorgen[VNr].PatID)
			VNr ++
			Gui, GV: Default
			GuiControl, GV:, GVPatient, % Vorsorgen[VNr].Name
			If VNr+1 <= Vorsorgen.MaxIndex()
				GuiControl, GV:, GVPatientNext, % Vorsorgen[VNr+1].Name
			else
				GuiControl, GV:, GVPatientNext, % ""

			AlbisAkteOeffnen(Vorsorgen[VNr].PatID)

	}
	else If InStr(A_GuiControl, "GVBackward") && (VNr > 1)
	{
			AlbisAkteSchliessen(Vorsorgen[VNr].PatID)
			VNr --
			Gui, GV: Default
			GuiControl, GV:, GVPatient, % Vorsorgen[VNr].Name
			GuiControl, GV:, GVPatientNext, % Vorsorgen[VNr+1].Name
			AlbisAkteOeffnen(Vorsorgen[VNr].PatID)
	}

return ;}

GVGuiClose:
GVEscape:

	Gui, GV: Destroy

return
;}


;}
M4:                                                  	;	fehlende geriatrische Basiskomplexe       	;{

	; Variablen definieren																															;{
		Gb					:= []								;Array der die PatientenID's geriatrischer Patienten enthält
		noGb				:= []
		noGbIndex		:= 0
		newGBIndex  	:= 0
		GbIndex	    	:= 0
		GBFound  		:= 0
		PatIndex			:= 0
		actGeriatrisch	:= 0
		PatID				:= 0
		Last_PatID			:= 0
		Last_Name		:= ""
		Last_Birth			:= ""
	;}

		;SciTEOutput("PNr: " Protokolle[1])
		If       	(Protokolle.MaxIndex() > 1) {
			MsgBox, 1, Addendum für Albis on Windows, Hier bitte nur ein Protokoll zur Auswertung auswählen!
			return
		}
		else If	(Protokolle.MaxIndex() = 0) {
			MsgBox, 1, Addendum für Albis on Windows, Ein Protokoll zur Auswertung auswählen!
			return
		}

		For idx, protokoll in Tagesprotokolle
			If InStr(protokoll.filename, Protokolle[1])
				break

		tprotfile := TagesprotokollPfad "\" protokoll.filename

	; PatientenIDs der geriatrischen Patienten einlesen (Dateipfad: logs'n'data\_DB\geriatrische_Patienten.txt) 	;{
		If !FileExist(AddendumDbPath "\DB_Geriatrisch.txt") 		{

			MsgBox,
			(LTrim
				Es wurde noch keine Datenbankdatei für geriatrische Patienten erstellt.
				Schritt 1:
				Die Erstellung einer solchen Datei benötigt ein oder mehrere Tagesprotokolle.
				Zunächst brauchen wir möglichst einen sehr langen Zeitraum 1-2 Jahre oder mehr.
				Albis braucht mehrere Stunden für ein Tagesprotokoll von circa 2 Jahren!
				Erstellen Sie das Protokoll z.B. nach der Sprechstunde.
				Schritt 2:
				Starten Sie das Abrechnungshelfer-Skript und drücken dort auf folgenden Button:
				"                                 nur ein Tagesprotoll parsen                                            "
				Schritt 3:
				drücken Sie nach dem Parsen auf den Button:
				"                                         freieStatistik                                                           "
				sie ....... #### muss ich noch weiter schreiben
			)
			return

		} else {

			FileRead, tmpVar, % AddendumDbPath "\DB_Geriatrisch.txt"
			Loop, Parse, tmpVar, `n, `r
			{
				RegExMatch(A_LoopField, "^\d+", tID)
				gb[A_Index] := tID
			}

			GuiControl,, T2, % "Insgesamt " Gb.MaxIndex() " geriatrische Patienten sind in der Datenbank."

		;leeren der Variablen
			VarSetCapacity(tmpVar, 0)
			VarSetCapacity(tID, 0)

		}
	;}

	; Einlesen des zu untersuchenden Tagesprotokolles																					;{

		FileOutExt   	:= StrReplace(protokoll.filename, ".txt", "")
		gbfehlend 	:= AlbisListenPfad "\" FileOutExt "_fehlende-geriatr.Ziffern.txt"
		gbProtokoll 	:= AlbisListenPfad "\" FileOutExt "_GB-Protokoll.txt"

		If FileExist(gbfehlend) {
			MsgBox, 4, Addendum für Albis on Windows, Das ausgewählte Protokoll wurde schon ausgewertet!`nSoll die Auswertung wiederholt werden?
			IfMsgBox, No
				goto fehlendeGBZiffer_Bearbeiten
		}

		FileAlbis		:= FileOpen(gbfehlend  	, "w", "CP1252")
		FileProt			:= FileOpen(gbProtokoll	, "w", "UTF-8")
		FileAlbis.Write("\P`n")

	;}

	; schauen ob ein Patient in geriatrischen Datendatei steht und ob eine Ziffer abgerechnet wurde				;{
		FileRead, tprot, % tprotfile
		Loop, Parse, tprot, `n, `r
		{
			;Patientenbereich gefunden
				If RegExMatch(A_LoopField, RegExFind[1], PatID) {

					; Name und Geburtsdatum
						RegExMatch(A_LoopField, RegExFind[3], PatName)
						RegExMatch(A_LoopField, RegExFind[4], PatBirth)

					; prüft wenn keine GB-Ziffer gefunden wurde ob der Patient in der Liste der geriatrischen Patienten eingetragen ist
						If      	(GBFound = 0) && (PatIndex > 0)	{

							If inArr(Gb, PatIDStored)	{
								noGbIndex ++, noGB[noGbIndex] := PatIDStored, name := StrSplit(Last_Name, ",")
								GuiControl, rxTP:, T6, % "Patient       : " Last_Name " ohne 03360 oder 03362"
								GuiControl, rxTP:, T7, % "fehlend bei: " noGbIndex
								FileAlbis.Write("\N+00000" . PatIDStored . "  " . PatIDStored . " | " . name[1] . SubStr("                                   ", 1, 20-StrLen(name[1])) . " | " . name[2] . "`n")
							}

						}
						else If	(GBFound = 1) && (PatIndex > 0)	{

							If !inArr(Gb, PatIDStored)		{
								GB.Push(PatIDStored)
								GBIndex := GB.MaxIndex(), newGBIndex ++, name := StrSplit(Last_Name, ",")
								FileAppend, % PatIDStored ";" Last_Name ";" Last_Birth, % AddendumDbPath . "\DB_Geriatrisch.txt"	;sichern des neuen Patienten
								GuiControl, rxTP:, T3, % "Insgesamt " Gb.MaxIndex() " geriatrische Patienten sind in der Datenbank."
								GuiControl, rxTP:, T6, % "Patient       : " Last_Name " neu aufgenommen als geriatrischer Pat."
								GuiControl, rxTP:, T7, % ""
							}

						}

					; dies ist nur für die Protokolldatei
						If (PatIndex > 0)	{

							inDB	:= inArr(Gb, PatIDStored), inDBText := "geriatrischer Patient: " inDB
							If inDB
								GbText	:=  "GB-Komplex: " GBFound, actGeriatrisch ++
							else
								GbText	:= GBFound = 1 	? "GB-Komplex: 1" : ""

							FileProt.Write(PatIDStored " | " inDBText " | " GbText " | " PatZeile "`n")

						}

						PatIDStored := PatID, Last_Name := PatName, Last_Birth := PatBirth, PatZeile := A_LoopField, GBFound := 0, PatIndex ++
						GuiControl, rxTP:, T4, % "akt. Patient   : " PatName " (" PatID ")"
						GuiControl, rxTP:, T5, % "Patientenzahl: " PatIndex
						continue
				}

			; Sucht die Ziffer 03360 oder 03362
				If !GBFound && RegExMatch(A_LoopField, "i).*lko.*0336")
					GBFound:= 1

		}

		FileAlbis.Close()
		FileProt.Close()

		GuiControl,, T3, % "Insgesamt " Gb.MaxIndex() " geriatrische Patienten sind in der Datenbank."
		GuiControl,, T4, % actGeriatrisch " geriatrische Patienten fanden sich in der untersuchten Datei."
		GuiControl,, T5, % StrReplace(tprotfile, AddendumDir)
		GuiControl,, T6, % " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
		GuiControl,, T7, % "Bei " (noGbIndex=0 ? "KEINEM" : noGbIndex) " der Patienten fehlten ein oder zwei Ziffern."
		GuiControl,, T8, % (newGBIndex=0 ? "Keine neuen" : newGBIndex=1 ? " 1 neuer": newGBIndex " neue") " geriatr.Pat. für die Datenbank."
		GuiControl,, T9, % "Die Ausgabedateien wurden geschrieben."

		;SavePatList(GB, FullFilePath, BackupPath:="")


	;}

	fehlendeGBZiffer_Bearbeiten:

	MsgBox, 4, Addendum für Albis on Windows, Patienten mit fehlenden GB-Komplexen wurden rausgesucht!`nMöchten Sie die Datei in Albis öffnen?
	IfMsgBox, No
	{
			MsgBox, 4, Addendum für Albis on Windows, Oder sollen die Ziffern jetzt durch das Skript eingetragen werden?
			;IfMsgBox, Yes
			;	AlbisGBZiffernEintragen(gbfehlend)
			Loop % noGb.MaxIndex()	{

				AlbisAkteOeffnen(noGB[A_Index])
				MsgBox, 4, Addendum für Albis on Windows, Drücke auf 'Ja' um den nächsten Patienten zu bearbeiten`noder auf 'Nein' um den Vorgang abzubrechen.
				IfMsgBox, No
					break

			}
			return
	}

	AlbisDateiAnzeigen(gbfehlend)

return

;}
M5:                                                  	;	Ziffern Statistik erstellen                          	;{

		Quartal     	:= Object()

	; Ziffernstatistiken aller Jahre erstellen
		startJahr    	:= 10
		startQuartal	:= 1
		endJahr    	:= 20
		endQuartal	:= 1

	Loop, % (endJahr - startJahr)
	{
			JIndex := A_Index - 1
			Loop, 4
			{
					Jahr		    	    	:= startJahr + JIndex
					FullFilePath        	:= TagesprotokollPfad "\Auswertungen\20" Jahr "-" A_Index "_Ziffernstatistik.txt"

					If !FileExist(FullFilePath) {
						Quartal.Aktuell	:= "0" . A_Index . Jahr
						Quartal         	:= QuartalTage(Quartal)
						savedAs         	:= AlbisErstelleZiffernStatistik("", Quartal.TDBeginn "," Quartal.TDEnde, "", "", FullFilePath)
					}
			}

	}

	MsgBox, 4, % A_ScriptName, Soll die erstellte Datei gleich kovertiert werden?
	If MsgBox, No
		return

;return ;}
M6:                                                  	;	Ziffern Statistik konvertieren                    	;{

	GuiControl,rxTP: , T3, % "Ziffernstatistik konvertieren....."
	ZFStats			:= Object()
	new_Statistik	:= ""
	Quartal     	:= aktQuartal
	nr              	:= 0
	line           	:= ""
	Statistikfile 	:= TagesprotokollPfad "\Auswertungen\20" SubStr(Quartal,3,2) "-" SubStr(Quartal,2,1) "_Ziffernstatistik.txt"

	ZFStats[Quartal]:= Object()
	If !FileExist(Statistikfile)	{
			SplitPath, Statistikfile, outFileName
			MsgBox, 1, % A_ScriptName, % "Eine Auswertungsdatei mit dem Name:`n" outFileName "`nexistiert nicht. Der Vorgang wird abgebrochen", 5
			return
	}

	FileRead, statistik, % Statistikfile
	;MsgBox, % Statistikfile "`n" StrLen(statistik)
	Loop, Parse, statistik, `n, `r
	{
			If RegExMatch(A_LoopField, RegExZSData, stat)			{
					GuiControl,rxTP: , T5, % "Nr: " A_Index " - " statGoNr
					statGoNr:= "#" statGoNr
					If ZFStats[Quartal].HasKey(statGoNr)					{
							ZFStats[Quartal][statGoNr].Anzahl      	:= +statAnzahl
							ZFStats[Quartal][statGoNr].Punkte      	:= +statPunkte
							ZFStats[Quartal][statGoNr].Euro         	:= +statEuro
							ZFStats[Quartal][statGoNr].Prozent     	:= +statProzent
							ZFStats[Quartal][statGoNr].Fachgruppe	:= +statFachgruppe
							ZFStats[Quartal][statGoNr].Abweichung	:= +statAbweichung
					}
					else					{
							ZFStats[Quartal][statGoNr] := Object()
							ZFStats[Quartal][statGoNr] := {"Anzahl":statAnzahl, "Punkte":statPunkte, "Euro":statEuro, "Prozent":statProzent, "Fachgruppe":statFachgruppe, "Abweichung":statAbweichung, "Leistungstext": statLeistungstext}
					}
			Sleep, 100
			}
	}

	MsgBox, % ZFStats[Quartal]["#02300"].Euro

	line:= Quartal "`n"
	For Quartal, Ziffer in ZFStats
		For GoNr, Data in Ziffer		{
			line .= GoNr ";" SubStr("00000" Data.Anzahl, -3) ";" SubStr("00000" Data.Punkte, -4) ";" SubStr("00000" Data.Euro, -4) ";" SubStr("00000" Data.Prozent, -3) ";" SubStr("00000" Data.Fachgruppe, -3) ";" SubStr("00000" Data.Abweichung, -3) ";" Data.Leistungstext "`n"
			GuiControl,rxTP: , T6, % "Data.Anzahl: " Data.Anzahl
		}

	FileDelete          	, % TagesprotokollPfad "\" Quartal "-ZF.csv"
	FileAppend, % line	, % TagesprotokollPfad "\" Quartal "-ZF.csv"


return
;}
M7:                                                  	;	Ziffern Statistiken zusammenfuehren       	;{

	Lk 		:= []
	Ziff	    := []
	ind:= 0

	ZfStatFolder:=  "M:\Praxis\Skripte\Skripte Neu\Addendum für AlbisOnWindows\logs'n'data\_DB\Ziffernstatistik\"

	ZfStDatei:= []
	ZfStDatei[1]:= 	"Ziffernstatistik_1801.txt"
	ZfStDatei[2]:= 	"Ziffernstatistik_1802.txt"
	ZfStDatei[3]:= 	"Ziffernstatistik_1803.txt"
	ZfStDatei[4]:= 	"Ziffernstatistik_1804.txt"
	ZfStDatei[5]:= 	"Ziffernstatistik_1901.txt"
	ZfStDatei[6]:= 	"Ziffernstatistik_1902.txt"
	ZfStDatei[7]:= 	"Ziffernstatistik_1903.txt"
	ZfStDatei[8]:= 	"Ziffernstatistik_1904.txt"
	ZfStDatei[9]:= 	"Ziffernstatistik_2001.txt"
	ZfStDatei[10]:= "Ziffernstatistik_2002.txt"
	ZfStDatei[11]:= "Ziffernstatistik_2003.txt"

	Loop, % ZfStDatei.MaxIndex()
	{
			FileRead, ZfSt, % ZfStatFolder . ZfStDatei[A_Index]
			ind++

			Loop, Parse, ZfSt, `n, `r
			{
					 If A_Index = 1
					{
							RegExMatch(A_LoopField, "\d\/\d\d", Quartal)
							Quartal:= StrReplace(Quartal, "/", "")
							;ToolTip, % quartal
							continue
					}
					fields:= StrSplit("'" . A_LoopField, "|")
					Ziff[A_Index]	:= fields[1]
					Lk[(Fields[1])] .= ";" . fields[2]
			}

			For key, value in Lk
					If !inArr(ziff, key)
							Lk[(key)] .= ";0"

			ziff:="", ziff:=[]
	}

	File:= FileOpen(ZfStatFolder . "gesamtStatistik.txt", "w")
	File.Write("Ziffer;01/18;02/18;03/18;04/18;01/19;02/19`n")
	For key, val in Lk
			File.Write(key . val . "`n")
	File.Close()

	MsgBox, alle Statistiken zusammen geführt

return
;}
M8:                                                  	;	fehlende Chroniker Ziffern suchen          	;{

	; Variablen definieren                                                                                                                        	;{
		Gb            	:= Array()                                            	;Array der Patienten ID's enthält
		noGb        	:= Object()
		noGbIndex	:= newGBIndex := regExFound := regExIdx := PatIndex := PatIdx := ziff1 := ziff2 := lkoDa := actChroniker := 0
		Arbeitsfrei 	:= ""

		;SciTEOutput("PNr: " Protokolle[1])
		If (Protokolle.MaxIndex() > 1) {
			MsgBox, 1, Addendum für Albis on Windows, Hier bitte nur ein Protokoll zur Auswertung auswählen!
			return
		}
		else If (Protokolle.MaxIndex() = 0) {
			MsgBox, 1, Addendum für Albis on Windows, Ein Protokoll zur Auswertung auswählen!
			return
		}

		For idx, protokoll in Tagesprotokolle
			If InStr(protokoll.filename, Protokolle[1])
				break

	; Dateipfade
		TpFileName	:= Protokolle[1] "_TP-AbrH"
		Tprotfile     	:= TagesprotokollPfad "\" protokoll.filename
		gbfehlend 	:= AlbisListenPfad     	"\" TpFileName "_fehlende-ChronikerZiffer.txt"
		gbProtokoll 	:= AlbisListenPfad     	"\" TpFileName "_Chroniker-Protokoll.txt"
		gbMakro   	:= TagesprotokollPfad 	"\" TpFileName "_Chroniker-Automation.txt"

		GuiControl, rxTP: , T14, % "File       : " TpFileName

	;}

	; prüft ob eine Auswertung vorhanden ist                                                                                             	;{

		If FileExist(gbMakro) 	{

				aktuelle_Makro_Datei:= StrReplace(TpFileName, "_TP-AbrH")
				MsgBox, 4, Addendum für Albis on Windows - Abrechnungshelfer,
				(LTrim
				Für die Datei:

				>> %aktuelle_Makro_Datei%  <<

				gibt es eine fertige Auswertung.
				Soll diese jetzt für die automatisierte Eintragung
				der fehlenden Ziffern verwendet werden?
				)
				IfMsgBox, Yes
				{
						AlbisFehlendeLkoEintragen(gbMakro)
						return
				}

		}

	;}

	; Patienten ID's der Chroniker einlesen, Urlaubs-/Feiertage/Praxisschließtage einlesen                            	;{

		Chroniker	:= ChronikerListe(AddendumDbPath "\DB_Chroniker.txt")                	; Patienten-IDs der als Chroniker vermerkten Patienten lesen
		PatDB   	:= ReadPatientDatabase(AddendumDir "\logs'n'data\_DB")                	; Patienten-Datenbank einlesen

		GuiControl, rxTP: , T13, % "bek. Chroniker: " chroniker.MaxIndex()

		If FileExist(AddendumDbPath "Praxisschließtage.txt") {
			IniRead, Arbeitsfrei, % AddendumDbPath "\Praxisschließtage.txt", % Protokoll[1]
		}

	;}

	; Schreibzugriffe herstellen																	                        							;{

			; Listendatei im Albisformat
			FileAlbis		:= FileOpen(gbfehlend  	, "w", "CP1252")
			FileAlbis.Write("\P`n")   ; damit beginnen Textdateien die sich mit Albis öffnen lassen

			; Datendatei für die Automatisierung der Eintragungen im ini-Format
			FileProt			:= FileOpen(gbProtokoll	, "w", "UTF-8")
			If !FileExist(gbMakro) {

				; wird mit der Struktur einer ini Datei angelegt. So funktioniert dies wie eine einfache Datenbank. Inhalte lassen sich einfach prüfen
				; und neue Inhalte können hinzugefügt werden.
				FileMakro		:= FileOpen(gbMakro   	, "w", "UTF-8")
				FileMakro.WriteLine("; Addendum Abrechnungshelfer Automatisierungsdatei  >> fehlende Chronikerziffern <<")
				FileMakro.WriteLine("[Abrechnung]")
				FileMakro.WriteLine("Quartal=" Protokolle[1])
				FileMakro.WriteLine("[Patientenliste]")
				FileMakro.WriteLine("; bearbeitet|PatID|PatName|Behandlungstage|Ziff1|Ziff2")
				FileMakro.Close()

			} else {

				; zählt wieviele Patienten in der Datei gelistet sind
				Loop {

					IniRead, val, % gbMakro, Patientenliste, % "Patient" A_Index
					If InStr(val, "ERROR")
						break
					else if (StrLen(val) = 0)
						continue

					noGbIndex ++
					data := StrSplit(val, "|")
					bearbeitet := (data[1] = 1) ? (bearbeitet + 1) : bearbeitet

					noGB.Push({	"bearbeitet"        	: data[1]
									, 	"PatID"              	: data[2]
									, 	"PatName"         	: data[3]
									, 	"Behandlungstage": data[4]
									, 	"Ziff1"                	: data[5]
									, 	"Ziff2"                	: data[6]})

				}

			}

	;}

	; schauen ob ein Patient in Chroniker Datendatei steht und ob eine Ziffer abgerechnet wurde		    		;{
		ziff2 := ziff1 := 0, ziff := ""

		; zeilenweises Auswerten des Tagesprotokolles
		FileRead, tprot, % Tprotfile
		Loop, Parse, tprot, `n, `r
		{

			; Patienten ID gefunden
				If RegExMatch(A_LoopField, RegExFind[1], PatID) {

				; Daten des vorherigen Patienten werden gespeichert
					If !regExFound && (PatIndex > 0) && inDB && lkoDa 	{         	; fehlende Chronikerziffer, Pat. ist gelistet in der Chronikerdatei

						; verhindert das Überschreiben von Daten
							inNoGB := false
							For index, data in noGB
								  If (data.PatID = PatIDStored) {
										inNoGB := true
										break
								  }

						; Hinzufügen eines neuen Patienten
							If !inNoGB {

								noGbIndex++
								ziff := RTrim( ((ziff1 = 0 ? "03220," : "") . (ziff2 = 0 ? "03221" : "")), ",")
								noGB.Push({	"PatID"               	: PatIDStored
												, 	"PatName"         	: PatName
												, 	"Behandlungstage": RTrim(BehTage, ";")
												, 	"Ziff1"                	: ziff1
												, 	"Ziff2"                	: ziff2})

								IniWrite, % "0|" PatIDStored "|" PatName "|" RTrim(BehTage, ";") "|" Ziff1 "|" Ziff2, % gbMakro , % "patient" noGbIndex

								FileAlbis.Write(	"\N+00000" PatIDStored "  " PatIDStored " | "
														. StrSplit(PatName, ",").1
														. SubStr("                                   ", 1, 20 - StrLen(StrSplit(PatName, ",").1))
														. "`t | " StrSplit(PatName, ",").2
														. " | fehlend: " ziff "`n")

								GuiControl,, T5, % "Patient       : " PatName " ohne " ziff
								GuiControl,, T6, % "fehlend bei: " noGbIndex

							}

					}
					else if regExFound && (PatIndex > 0) && !inDB && lkoDa { 	; Chronikerziffer(n) ist vorhanden, Pat. ist aber noch nicht als Chroniker vermerkt

						newGBIndex++
						FileAppend, % PatIDStored ";" PatName ";" PatBirth "`n",  % AddendumDbPath "\DB_Chroniker.txt", UTF-8

					}

				; zeigt die Gesamtzahl aller Patienten des Quartals. Patienten ohne Gebührennummer werden nicht gezählt
					If lkoDa {
						GuiControl,, T3, % "akt. Patient   : " PatName
						GuiControl,, T4, % "Patientenzahl: " PatIndex++
					}

				; Daten des aktuellen Patienten werden gesichert, Variablen werden zurückgesetzt
					PatIDStored	:= PatID
					If (inDB := inArr(Chroniker, PatIDStored))
						actChroniker++

				; der Name des Patienten und sein Geburtsdatum stehen in derselben Zeile wie die Patienten ID
					RegExMatch(A_LoopField, RegExFind[3], PatName)
					RegExMatch(A_LoopField, RegExFind[4], PatBirth)

				; Variablen zurücksetzen
					lkoDa := ziff1 := ziff2 := 0
					BehTagLast := BehTage := "", ziff := ""

					continue
				}

			; behandlungsrelevante Karteikartenkürzel erkannt, Tagesdatum wird als Behandlungsdatum aufgenommen
				If RegExMatch(A_LoopField, "\sanam|AISC|bef|bem|bild|dia|fau|fausw|fhp|fhv13|fhv18|fkb|fkh|füb|fvap|labor|lko|lle|medrp|medp|medis|medh|info|spiro\s") 	{

					If RegExMatch(A_LoopField, "^\d\d\.\d\d\.\d\d\d\d", BehDatum) 	; Behandlungsdatum gefunden
							BehTage .= BehDatum ";", BehTagLast := BehDatum
					If RegExMatch(A_LoopField, "\s(lko)\s")                                         	; abgerechnete Leistungen gefunden
							lkoDa := 1

				}

			; Chroniker Ziffern erkennen und behandeln
				If (ziff1 = 0) 	&& (regExFound:= RegExMatch(A_LoopField, "lko.*03220"))
					ziff1 := 1
				if (ziff2 = 0)	&& (regExFound:= RegExMatch(A_LoopField, "lko.*03221"))
					ziff2 := 1

			; Zähler für vollständigen Chronikerkomplex
				If regExFound && (ziff1 = 1) && (ziff2 = 1)
					GuiControl,, T7, % "03220 & 03221 Zähler: " (regExIdx++)

		}

	; erstellt eine Liste von Patienten mit fehlenden Chronikerziffern in Form einer Ini-Textdatei
		For index, obj in noGB
			FileMakro.WriteLine("patient" index "=0|" obj.PatID "|" obj.PatName "|" obj.Behandlungstage "|" obj.Ziff1 "|" obj.Ziff2)

		FileAlbis.Close()
		FileProt.Close()

	;}

	; Auswertung und Zusammenfassung anzeigen						    															;{

		GuiControl,, T3, % "Insgesamt " Gb.MaxIndex() " Chroniker sind in der Datenbank."
		GuiControl,, T4, % actChroniker " Chroniker fanden sich in der untersuchten Datei."
		GuiControl,, T5, % " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
		GuiControl,, T6, % "Bei " (noGbIndex=0 ? "KEINEM" : noGbIndex) " der Patienten fehlten ein oder zwei Ziffern."
		GuiControl,, T7, % (newGBIndex=0 ? "Keine neuen" : newGBIndex=1 ? " 1 neuer" : newGBIndex " neue") " Chroniker für die Datenbank."
		GuiControl,, T8, % "Die Ausgabedateien wurden geschrieben."

	;}

	; automatisierte Eintragung der Komplexe starten                                                                                	;{
		MsgBox, 4, Addendum für Albis on Windows, Sollen die fehlenden Ziffern durch das Skript eingetragen werden?
		IfMsgBox, Yes
		{
			res := AlbisFehlendeLkoEintragen(gbMakro, Arbeitsfrei)
			If (res=0)
				MsgBox, 4, Addendum für Albis on Windows, Funktionsaufruf ist fehlgeschlagen.
		}

	;}

return

;}
M9:                                                  	;	neue Chroniker finden                           	;{

	; letzte Änderung am 06.07.2020
	;
	; nutzt eine Liste von ICD-Schlüsselnummern, die nach Einschätzung der AG medizinische Grouperanpassung chronische Krankheiten kodieren (Bewertungsausschuss nach § 87 Absatz 1 SGB V)
	; damit werden gesetzlich korrekt Patienten identifiziert bei den der Chronikerkomplex angesetzt werden kann
	;
	; Das Makro öffnet auf Wunsch die erstellte Datei mit den über die ICD-Schlüsselnummern vorausgewählten Patienten in Albis.
	;
	; Es ist schwer möglich Patienten korrekt der Gruppe der "chronisch Kranken" zuzuordnen. Sie selbst wissen, das weitere Bedingungen erfüllt sein müssen.
	; Schritt 1: Gehen Sie also durch die Vorauswahl der Patienten. Geben Sie die oder den Chronikerkomplex(e) bei jedem von Ihnen als korrekt erkannten Patienten ein.
	; Schritt 2: Erstellen Sie über das Makro 'Tagesprotokoll erstellen' für das gerade untersuchte Quartal erneut ein Tagesprotokoll
	; Schritt 3:	Starten Sie das Makro 'fehlende Chronikerziffern'. Patienten die sich nicht in der Chronikerliste aufgeführt sind, bei den ein Chronikerkomplex abgerechnet wurde,
	;				werden dieser Liste hinzugefügt.


	; Variablen
		inDB := 0
		newChr := oldChr := 0

		Chroniker	:= ChronikerListe(AddendumDbPath "\DB_Chroniker.txt")                	; Patienten-IDs der als Chroniker vermerkten Patienten lesen
		PatDB   	:= ReadPatientDatabase(AddendumDir "\logs'n'data\_DB")                	; Patienten-Datenbank einlesen

		Fortschrittsanzeige("leeren")

	; Dateipfade und Schreibzugriffe erstellen												;{

		tprotfile := Tagesprotokolle[(Tagesprotokolle.MaxIndex())].filename ; das letzte Tagesprotokoll und wenn erstellt das aktuell zu bearbeitende
		RegExMatch(tprotfile, "\d+\-\d", tprotDate)
		ChronikerVorschlag 	:= AlbisListenPfad "\NeueChroniker-Quartal(" tprotDate ").txt"
		If FileExist(ChronikerVorschlag) {
				MsgBox, 4, Addendum für Albis on Windows, % "Möchten Sie das Quartal " tprotDate " nocheinmal untersuchen lassen?"
				IfMsgBox, No
					goto neueChronikerAnzeigen
		}

		FileAlbis:= FileOpen(ChronikerVorschlag, "w", "CP1252")
		FileAlbis.Write("\P`n")

	;}

	; erstellt eine Datenbank mit den Dauerdiagnosen der Patienten				;{

		Loop, Parse, % FileOpen(TagesprotokollPfad "\" tprotfile, "r", "UTF-8").Read(), `n, `r
		{

				If RegExMatch(A_LoopField, RegExFind[1], PatID) {                                 	; neue Patienten-ID, Name, und Geburtstag (RegExFind[1])

						PatIDStored	:= PatID
						ICDMatch 	:= false
						If inDB			:= inArr(Chroniker, PatID)                    ; Pat ist schon als Chroniker vermerkt?
								oldChr++

						RegExMatch(A_LoopField, RegExFind[3], PatName)
						RegExMatch(A_LoopField, RegExFind[4], PatBirth)

						GuiControl,, T4, % "Patient: "                  	PatName
						GuiControl,, T5, % "ID: "                          	PatID
						GuiControl,, T6, % "Chroniker: "               	(inDB ? "ja":"nein")
						GuiControl,, T7, % "bekannte Chroniker: "	oldChr
						GuiControl,, T8, % "neue Chroniker: "    	newChr

						continue

				}

				If inDB || ICDMatch                                                                             	; fortfahren bis zum nächsten Patienten der nicht in der Chronikerliste steht
					continue

				If RegExMatch(A_LoopField, RegExFind[8], dummy)  {                          	; Untersuchen von Diagnosen auf Vorkommen einer Chroniker ICD

					Diagnosen := RegExGetAll(A_LoopField, RegExFind[8], "ICD") 	; gibt ein Array mit allen ICD-Codes der Zeile zurück
					;SciTEOutput(Diagnosen.MaxIndex())
					For index, ICDCode in Diagnosen
						If inArr(CICD, ICDCode) {

							ICDMatch := true
							GuiControl,, T8, % "neue Chroniker: " 	(newChr ++)
							GuiControl,, T9, % "CDiagnose: "       	ICDCode
							FileAlbis.Write("\N+00000" PatIDStored "  " PatIDStored " | " StringSpacer(StrSplit(PatName, ",").1, 30) " | " StringSpacer(StrSplit(PatName, ",").2, 30) " | " PatBirth "`n")

							break
						}
					continue

		    	}

		}

		FileAlbis.Close()

	neueChronikerAnzeigen:
		MsgBox, 4, Addendum für Albis on Windows, % "Möchten Sie die erstellte Liste jetzt mit Albis ansehen?"
		IfMsgBox, Yes
			AlbisDateiAnzeigen(ChronikerVorschlag)

	return	;}


;}
M10:                                                	;	Tagesprotokolle erstellen                        	;{

	PQuartal:="", Prepend:= ""
	statistics:= Object()

	If !WinExist("ahk_class OptoAppClass")	{
		MsgBox, 1, Addendum - Abrechnungshelfer, Auf diesem Rechner ist Albis nicht gestartet!`NDas gewählte Makro kann nicht ausgeführt werden!
		return
	}

	QuartalsEingabe:

	InputBox, Quartal, Addendum für Albis on Windows,
			(LTrim
			%Prepend% Geben Sie ein
			als einzelnes Quartal: 03/18 o. 0318 o. 318,
			als Quartalsbereich: `t01/17-04/19 o. 0117-0419,
			als Datumsbereich:  `t18.01.2019-20.3.2019
			alle oder *              `t: Erstellung aller Abrechnungsquartale
			),,,,,,,, % PQuartal
	If ErrorLevel
		return

	If RegExMatch(Quartal, "i)\*|(\salle\s)") {

			Loop 10		{

				LoopJahr:= SubStr("00" 9+A_Index, -1)
				Loop 4 				{

					PQuartal       	:= ""
					Prepend        	:= ""
					LoopQuartal 	:= SubStr("00" A_Index, -1)
					Quartal         	:= LoopQuartal LoopJahr
					protokollDatei	:= TagesprotokollPfad "\20" LoopJahr "-" A_Index "_TP-AbrH.txt"

					If !FileExist(protokollDatei) 	{

						; Einlesen und konvertieren eines Protokolles
							Startzeit 	:= A_TickCount
							newTPFile	:= AlbisErstelleTagesprotokoll(Quartal, TagesprotokollPfad, 1)
							newTPText	:= RegExParseTPFile(newTPFile, TagesprotokollPfad, RegExTP, true)

						; Listviewinhalt auffrischen, wenn eine Datei erfolgreich erstellt werden konnte
							If (StrLen(newTPText) > 0) 	{

								Erstellungszeit:= A_TickCount - Startzeit
								FileAppend, % Quartal ";Erstellungszeit: " Erstellungszeit "`n", % TagesprotokollPfad "\Statistiken_und_Fehler.log"
								statistics.Push({"Quartal": quartal, "Erstellungszeit": Erstellungszeit})
								Tagesprotokolle:= ReadTPDir(TagesprotokollPfad, "*_TP-AbrH", "txt", "R")
								UpdateTagesProtokolle(Tagesprotokolle)
								WinActivate, % "ahk_id " hAbrH

							}
							else	{

								FileAppend, % Quartal ";Rückgabe leer. Protokoll konnte nicht konvertiert werden.`n" , % TagesprotokollPfad "\Statistiken_und_Fehler.log"
								PraxTT("Protokolldatei: 20" LoopJahr "-" A_Index " konnte nicht konvertiert werden.`nFahre mit dem nächsten Quartal fort.", "2 1")
								Sleep, 1000
							}
					}
					else 	{

							PraxTT("Protokolldatei: 20" LoopJahr "-" A_Index " ist vorhanden.`nFahre mit dem nächsten Quartal fort.", "1 1")
							;Sleep, 1000

					}
				}
			}
	}
	else {

		Startzeit         	:= A_TickCount
		gosub M10A
		EndZeit         	:= A_TickCount
		Erstellungszeit	:= (EndZeit - Startzeit)/1000
		;FormatTime, EZeit, Erstellungszeit,
		GuiControl,rxTP: , T14, % "Quartal: " Quartal " , Erstellungszeit: " Erstellungszeit "s"
		statistics.Push({"Quartal": quartal, "Erstellungszeit": Erstellungszeit})

	}

return

M10A: ;{

	PQuartal := RegExReplace(Quartal, "\s*")
	PQuartal := StrReplace(PQuartal, "/")
	;PQuartal := StrReplace(PQuartal, " ")
	If      	RegExMatch(PQuartal, "^\d{3,4}$")	{

			QNr := SubStr(PQuartal, StrLen(PQuartal) - 2, 1)
			If QNr not in 1,2,3,4
			{
				prepend:="! UNGÜLTIGE QUARTALSZAHL ! - es sind nur Ziffern von 1-4 erlaubt!`n"
				goto Quartalseingabe
			}

	}
	else if	RegExMatch(PQuartal, "^(\d{3,4})\-(\d{3,4})$", Q)	{

			QNr1 := SubStr(Q1, StrLen(Q1) - 2, 1)
			QNr2 := SubStr(Q2, StrLen(Q2) - 2, 1)
			If QNr1 not in 1,2,3,4
			{
				prepend:="! UNGÜLTIGE QUARTALSZAHL ! - es sind nur Ziffern von 1-4 erlaubt!`n"
				goto Quartalseingabe
			}
			else If QNr2 not in 1,2,3,4
			{
				prepend:="! UNGÜLTIGE QUARTALSZAHL ! - es sind nur Ziffern von 1-4 erlaubt!`n"
				goto Quartalseingabe
			}

	}
	else	{

			prepend:="! KEIN GÜLTIGES FORMAT !`n"
			goto Quartalseingabe

	}

	newTPFile := AlbisErstelleTagesprotokoll(Quartal, TagesprotokollPfad, 1)
	If (newTPFile = "invoke error") {
		MsgBox, 0x40000, Tagesprotokoll erstellen, Ups, da ist was schief gegangen....
		return
	}

	RegExParseTPFile(newTPFile, TagesprotokollPfad, RegExTP, true)

  ; Listviewinhalt auffrischen
	Tagesprotokolle:= ReadTPDir(TagesprotokollPfad, "*_TP-AbrH", "txt", "R")
	UpdateTagesProtokolle(Tagesprotokolle)


	WinActivate, % "ahk_id " hAbrH


return ;}

;}
M11:                                                	;	Dauerdiagnosen anzeigen                     	;{

		Diagnosen    	  := Object()
		found            	  := 0
		GesamtICD     	  := 0
		NaechsterPatient := 0
		ReadNextLine	  := 0

		;FileRead, tprot, % tprotfile
		file := FileOpen(tprotfile, "r")
		tprot:= File.Read()
		File.Close()

		Loop, Parse, tprot, `n, `r
		{
			; Patienten-ID, Name, Geburtstag wird ermittelt
				If RegExMatch(A_LoopField, RegExFind[1], PatID)
				{
						PatIDStored    	  := PatID
						DauerDia         	  := ""
						NaechsterPatient := 0
						RegExMatch(A_LoopField, RegExFind[3], PatName)
						RegExMatch(A_LoopField, RegExFind[4], PatBirth)
						GuiControl,, T3, % "Patient: " PatName
						GuiControl,, T4, % "PatientID: " PatID
						continue
				}

			; der Code darunter wird erst ausgewertet, wenn ein neuer Patient gefunden wurde
				If NaechsterPatient
						continue

			; eine leere Zeile markiert das Ende des Dauerdiagnosenbereiches,
				If DauerDia && (A_LoopField = "") && !ReadNextLine
						NaechsterPatient := 1, continue

			; Dauerdiagnosenbereich gefunden
				If RegExMatch(A_LoopField, "(Behandlung|anamnestisch)\s+(.*)", DauerDia) || (ReadNextLine = 1)
				{
							If !RegExMatch(DauerDia2, "(?<=\{)([\+\*\!]*[A-Z]\d*\.*\d*)[\-\*]*.*([ZGAV])[\*]*(?=\})", ICD)
    						{
									ReadNextLine:= 1
									DDtmp:= DauerDia2
									continue
							}

							If ReadNextLine
							{
									RegExMatch(A_LoopField, ".*(?=\{)", lastPartDia)
									DauerDia2:= DDTmp " " Trim(lastPartDia)
									DDTmp  	:= ""
									ReadNextLine:= 0
							}

							ICD := RegExReplace(ICD1, "[\+\*\!]") ICD2
							DauerDia:= TrimDiagnose(DauerDia2)

						; neues Objekt anlegen
							If !Diagnosen.HasKey(ICD)
							{
									Diagnosen[(ICD)] := {"Counter" : 1, "Text" : {1:(DauerDia)}, "PatID" : {1:(PatIDStored)}}
									GesamtICD ++
							}
							else
							{
									Diagnosen[(ICD)].Counter ++
									GesamtICD ++
								; Diagnosentext hinterlegen, wenn dieser noch nicht vorhanden ist
									found := 0
									For key, Dia in Diagnosen[(ICD)]["Text"]
											If InStr(Dia, DauerDia)
													found:= 1, break
									If !found
											Diagnosen[(ICD)]["Text"].Push(DauerDia)

								; PatID hinzufügen
									found := 0
									For key, ID in Diagnosen[(ICD)]["PatID"]
											If InStr(ID, PatIDStored)
													found:= 1, break
									If !found
											Diagnosen[(ICD)]["PatID"].Push(PatIDStored)
							}

						; bischen was anzeigen
							GuiControl,, T6, % "ICD: " ICD
							GuiControl,, T7, % "Text: " DauerDia
							GuiControl,, T8, % "ICD gefunden: " GesamtICD
							GuiControl,, T9, % "Store: " Diagnosen[(ICD)].Counter

				}
		}

		RegExMatch(tprotFile, "\d+\-(\d+|I|II|III|IV)", tprotDate)
		AutoListview("ICD|Gesamtzahl|ICD Text|PatID", "Anzeige aller Dauerdiagnosen " tprotDate, "8,5,60,27", "Text,Integer,Text,Text")

		Gui, ALV:Default
		Gui, ALV: Listview, DasLV
		GuiControl, ALV: -Redraw, DasLV

		FileDelete, % RTrim(TagesprotokollPfad, "\") "\" tprotDate "-Dauerdiagnosen.txt"
		tmpVar:=""
		For ICD, Data in Diagnosen
		{
				tmerge:=""
				For idx, Text in Data["Text"]
				{
						tmerge.= Text "; "
						tmpVar.= Text " {" ICD "}`n"
				}
				pmerge:=""
				For idx, ID in Data["PatID"]
						pmerge.= ID "; "
				LV_Add("", ICD, Data.Counter, RTrim(tmerge, "; "), RTrim(pmerge, "; "))
		}
		GuiControl, ALV: +Redraw, DasLV
		FileAppend, % tmpVar, % RTrim(TagesprotokollPfad, "\") "\Auswertung_Dauerdiagnosen_" tprotDate ".txt", UTF-8
		tmpVar:=""
		;DD.Close()

return
;}
M12:                                                	;	Gesamtdaten-Querverweise erstellen      	;{

	qv     	:= Object()
	PatID	:= ""
	PatAll	:= 0

	For tprotIndex, tprot in Tagesprotokolle 	{

			fileread, tp, % tprot.path "\" tprot.filename
			NoExt:= StrReplace(tprot.filename, ".txt", "")
			GuiControl,rxTP:, T3, % "Datei: " NoExt
			GuiControl,rxTP:, T4, % "DateiNr: " A_Index

			Loop, Parse, tp, `n, `r
			{
					tpZeile:= A_Index
					If RegExMatch(A_LoopField, RegExFind[1], ID) 					{

							If !IsObject(qv[ID])
								PatAll ++

							GuiControl,rxTP:, T5, % "PatID's gesamt: " PatAll
							GuiControl,rxTP:, T6, % "PatID: " ID
							GuiControl,rxTP:, T7, % "Start: " tpZeile

							If (PatID <> ID) && (StrLen(PatID) != 0) 							{

									qv[PatID][NoExt].End:= tpZeile - 1
									GuiControl,rxTP:, T8, % "PatID last: " PatID
									GuiControl,rxTP:, T9, % "Ende: " tpZeile - 2

							}

							PatID:= ID

							If !IsObject(qv[PatID])
								qv[PatID]:= Object()

							qv[PatID][NoExt]:= {"Start":tpZeile, "End":0}
					}
			}
			qv[PatID][NoExt].End:= tpZeile - 2
	}

	q := Chr(0x22)
	qvFile := FileOpen(AddendumDbPath "\querverweise.json", "w", "UTF-8")

	PatAllLen := StrLen(PatAll) - 1

	For PatID, obj in qv	{

			GuiControl,rxTP:, T10, % "Speichere PatID: " SubStr("00000" PatID, -4) " (" SubStr("000000" A_Index, -1*PatAllLen) "/" PatAll ")"
			qvFile.WriteLine(q SubStr("00000" PatID, -4) q ":{")

			t := ""
			For tprotfile, Data in obj 			{

					t.= "`t" q tprotfile q ":[`n"
					t.= "`t`t" q "Start" q ":" Data.Start ",`n"
					t.= "`t`t" q "End"  q ":" Data.End "`n"
					t.= "`t`t],`n"

			}
			qvFile.WriteLine(RTrim(t, ",`n"))

			If A_Index = qv.MaxCount()
				qvFile.WriteLine("}")
			else
				qvFile.WriteLine("},")
	}

	qvFile.Close()


return ;}
M13:                                                 	;	zähle alle Überweisungen über alle Jahre	;{


	Ueberweisungen                    	:= Object()	; Überweisungen
	Ueberweisungen.Max             	:= Object() 	; Gesamtzahl der Überweisungen  (Quartal, Jahr) + über alle Jahre
	Ueberweisungen.Fachgruppen	:= Object()	; child Object Fachgruppen inklusive Zählung (gesamt und Quartal, Jahr)
	Ueberweisungen.Patienten       	:= Object()	; Altersgruppenstastiken der Überwesungen, Anzahl Pat. pro Quartal + Jahr

	FGMax 	:= 0
	qw        	:= 1080
	fk          	:= 1
	alle       	:= 1

	Gui, rxS: New, +HWNDhrxS -DPIScale +MinSize1200x530 +Resize   ; +AlwaysOnTop
	Gui, rxS: Margin, 5, 5
	Gui, rxS: Color, % "cFFFFAA"
	Gui, rxS: Font, % "s" 14 " cBlack bold" " q5", % Font
	Gui, rxS: Add, Text, % "xm ym w1200 Center", % "PFERDE-RENNEN - Fachgruppenüberweisungen"
	Gui, rxS: Font, % "s" 11 " cBlack bold" " q5", % Font
	Gui, rxS: Add, Text, % "y+-5 w1200 BackgroundTrans Center", % "--------------------------------------------------------------------------------------------------------------------"
	Gui, rxS: Font, % "s" 8 " cBlack" " q5", % Font
	Gui, rxS: Show, % "x30 y20 w1500 h1000", % "Fachgruppen-Rennen"
	Gui, rxS: Font, % "s" 9 " cBlack normal" " q5", % Font

	For index, protokoll in Tagesprotokolle 	{

		Jahr  	:= protokoll.Jahr
		Quartal	:= protokoll.Quartal "/" Jahr
		TPFullPath:= protokoll.path "\" protokoll.filename
		FileRead, tpDatei, % TPFullPath
		;ToolTip, % TPFullPath ",   " StrLen(tpDatei)
		Gui, rxTP: Default
		GuiControl,, T3, % "Quartal: " Quartal
		Loop, Parse, tpDatei, `n, `r
		{
			;If RegExMatch(A_LoopField, RegExFind[1], PatID) {
					;If !IsObject(Ueberweisungen.Pat[Quartal])
			;		Ueberweisungen.Pat := {"Quartal":
			; }

			If RegExMatch(A_LoopField, "^\s+füb\s+([\w\sÄÖÜäöüß\-\.]+),*|\(|$", FG)	{

					FG2:= Trim(FG1)
					FG1:= Trim(StrReplace(FG2, ".", " "))
					FG1:= Trim(StrReplace(FG1, "-", " "))
					FG1:= RegExReplace(FG1, "i)^FA\s+fü*r*", "")
					FG1:= RegExReplace(FG1, "i)^Facharzt\s+fü*r*", "")
					FG1:= RegExReplace(FG1, "i)^Arzt\s+fü*r*", "")
					FG1:= RegExReplace(FG1, "Diagnostik", "")
					FG1:= RegExReplace(FG1, "ambulante\s*", "")
					FG1:= RegExReplace(FG1, "Poliklinik\s*", "")
					FG1:= RegExReplace(FG1, "Allge*m[ie]*n"                             	, "Allgemein")
					FG1:= RegExReplace(FG1, ".*Allgemeinm.*"                          	, "Allgemeinmedizin")
					FG1:= RegExReplace(FG1, ".*Allge*meinar.*"                        	, "Allgemeinmedizin")
					FG1:= RegExReplace(FG1, ".*Hausar.*"                               	, "Allgemeinmedizin")
					FG1:= RegExReplace(FG1, "Allerlogie"                                	, "Allergologie")
					FG1:= RegExReplace(FG1, ".*Allerg.*"                                	, "Allergologie")
					FG1:= RegExReplace(FG1, ".*Allegie.*"                                	, "Allergologie")
					FG1:= RegExReplace(FG1, ".*Augenarzt.*"                         	, "Augenheilkunde")
					FG1:= RegExReplace(FG1, ".*A*Auge*n.*"                          	, "Augenheilkunde")
					FG1:= RegExReplace(FG1, "^Chr*iru.*"                             	, "Chirurgie")
					FG1:= RegExReplace(FG1, "Chirurige"                               	, "Chirurgie")
					FG1:= RegExReplace(FG1, "i).*Allgemeinch.*"                     	, "Chirurgie")
					FG1:= RegExReplace(FG1, "i).*Fußch.*"                              	, "Chirurgie")
					FG1:= RegExReplace(FG1, "i).*Unfallch.*"                           	, "Chirurgie")
					FG1:= RegExReplace(FG1, "i).*Handch.*"                           	, "Chirurgie")
					FG1:= RegExReplace(FG1, "i).*Handsprech.*"                        	, "Chirurgie")
					FG1:= RegExReplace(FG1, "i).*metabolische\sChi.*"              	, "Chirurgie")
					FG1:= RegExReplace(FG1, "i).*chirurgische\sAmb.*"              	, "Chirurgie")
					FG1:= RegExReplace(FG1, ".*MIC.*"                                  	, "Chirurgie")
					FG1:= RegExReplace(FG1, "i).*Neuroch.*"                          	, "Chirurgie/Neurochirugie")
					FG1:= RegExReplace(FG1, "i).*Neuch.*"                             	, "Chirurgie/Neurochirugie")
					FG1:= RegExReplace(FG1, ".*MKG.*"                                 	, "Chirurgie/MKG")
					FG1:= RegExReplace(FG1, "i).*Gesichtsch.*"                          	, "Chirurgie/MKG")
					FG1:= RegExReplace(FG1, "i).*Kieferch.*"                              	, "Chirurgie/MKG")
					FG1:= RegExReplace(FG1, "i).*Kinderch.*"                            	, "Chirurgie/Kinder")
					FG1:= RegExReplace(FG1, "i).*plastisch*"                            	, "Chirurgie/plastische Chirurgie")
					FG1:= RegExReplace(FG1, "i)Diabetolo*gi*e"                        	, "Diabetologie")
					FG1:= RegExReplace(FG1, "i).*Dia*be*t.*"                             	, "Diabetologie")
					FG1:= RegExReplace(FG1, "i).*Daibet.*"                              	, "Diabetologie")
					FG1:= RegExReplace(FG1, "i).*Fußamb.*"                             	, "Diabetologie")
					FG1:= RegExReplace(FG1, "i).*Gefäßch.*"                          	, "Chirurgie/Gefäßchirurgie")
					FG1:= RegExReplace(FG1, "i).*Gefäßambulanz.*"                	, "Chirurgie/Gefäßchirurgie")
					FG1:= RegExReplace(FG1, "i).*Gefch.*"                              	, "Chirurgie/Gefäßchirurgie")
					FG1:= RegExReplace(FG1, "i).*sensor*mot.*"                         	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Aortenspre.*"                          	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Add*ipositas.*"                       	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Bauchz.*"                            	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Darmkrebsz.*"                        	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Darmsp.*"                           	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*DRK\s+Kr.*"                          	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*für\s+seltene\s+Erkr.*"          	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Ventil.*"                              	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Lipid.*"                               	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Mukov.*"                              	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Naturheil.*"                            	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Urtikaria.*"                           	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Leberz.*"                             	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Brustz.*"                              	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Fachamb.*"                            	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Ernährungs.*"                         	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Infektions.*"                           	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Leberz.*"                             	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Immunolog.*"                         	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Muskelerkran.*"                      	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Tran*splantation.*"                 	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Transfusions.*"                    	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i).*Zentrum f Gefäßm.*"              	, "Fachzentren")
					FG1:= RegExReplace(FG1, "i)Gyn[äa]koogie"                        	, "Gynäkologie")
					FG1:= RegExReplace(FG1, "i)Gyn$"                                    	, "Gynäkologie")
					FG1:= RegExReplace(FG1, "i)Gyn\s+.*"                                	, "Gynäkologie")
					FG1:= RegExReplace(FG1, "i).*Gyn[äaö]kol.*"                      	, "Gynäkologie")
					FG1:= RegExReplace(FG1, "i).*Beckenbo.*"                          	, "Gynäkologie")
					FG1:= RegExReplace(FG1, "i).*Kinderw.*"                           	, "Gynäkologie")
					FG1:= RegExReplace(FG1, "i).*Haut.*"                               	, "Dermatologie")
					FG1:= RegExReplace(FG1, "i).*Dermat.*"                              	, "Dermatologie")
					FG1:= RegExReplace(FG1, "i).*HNO.*"                               	, "HNO")
					FG1:= RegExReplace(FG1, "i).*Phoniat.*"                              	, "HNO")
					FG1:= RegExReplace(FG1, "i).*Humangen.*"                         	, "Humangenetik")
					FG1:= RegExReplace(FG1, "i).*Genetik.*"                           	, "Humangenetik")
					FG1:= RegExReplace(FG1, "i)Internist"                                	, "Innere Medizin")
					FG1:= RegExReplace(FG1, "i)^\s*Innere$"                        	, "Innere Medizin")
					FG1:= RegExReplace(FG1, "i).*Innere M.*"                         	, "Innere Medizin")
					FG1:= RegExReplace(FG1, "i).*Ka*rdi*ol.*"                           	, "Innere Medizin/Kardiologie")
					FG1:= RegExReplace(FG1, "i).*Herz.*"                                 	, "Innere Medizin/Kardiologie")
					FG1:= RegExReplace(FG1, "i).*Schrittm.*"                              	, "Innere Medizin/Kardiologie")
					FG1:= RegExReplace(FG1, "i).*Defisp.*"                               	, "Innere Medizin/Kardiologie")
					FG1:= RegExReplace(FG1, "i).*Rhythmus.*"                           	, "Innere Medizin/Kardiologie")
					FG1:= RegExReplace(FG1, "i).*Endokr.*"                            	, "Innere Medizin/Endokrinologie")
					FG1:= RegExReplace(FG1, "i).*Gastroe*n*t*e*r*o*.*"             	, "Innere Medizin/Gastroenterologie")
					FG1:= RegExReplace(FG1, "i).*Endosk.*"                            	, "Innere Medizin/Gastroenterologie")
					FG1:= RegExReplace(FG1, "i).*Gastrent.*"                          	, "Innere Medizin/Gastroenterologie")
					FG1:= RegExReplace(FG1, "i).*Hepato.*"                             	, "Innere Medizin/Gastroenterologie")
					FG1:= RegExReplace(FG1, "i).*H[äae]*m[ao]*to.*"                 	, "Innere Medizin/Hämatologie und Onkologie")
					FG1:= RegExReplace(FG1, "i).*Onkol.*"                             	, "Innere Medizin/Hämatologie und Onkologie")
					FG1:= RegExReplace(FG1, "i).*Tumorsp.*"                             	, "Innere Medizin/Hämatologie und Onkologie")
					FG1:= RegExReplace(FG1, "i).*Hämostase.*"                         	, "Innere Medizin/Hämostaseologie")
					FG1:= RegExReplace(FG1, "i).*Gerr*inn*u.*"                         	, "Innere Medizin/Hämostaseologie")
					FG1:= RegExReplace(FG1, "i).*Nefro.*"                              	, "Innere Medizin/Nephrologie")
					FG1:= RegExReplace(FG1, "i).*Niere.*"                              	, "Innere Medizin/Nephrologie")
					FG1:= RegExReplace(FG1, "i).*Nephro.*"                              	, "Innere Medizin/Nephrologie")
					FG1:= RegExReplace(FG1, "i).*Pneumologi*e.*"                     	, "Innere Medizin/Pulmologie")
					FG1:= RegExReplace(FG1, "i).*Pul*o*molo.*"                         	, "Innere Medizin/Pulmologie")
					FG1:= RegExReplace(FG1, "i).*Lungen.*"                            	, "Innere Medizin/Pulmologie")
					FG1:= RegExReplace(FG1, "i).*Rh*eumat.*"                        	, "Innere Medizin/Rheumatologie")
					FG1:= RegExReplace(FG1, "i).*Rh*eumak.*"                        	, "Innere Medizin/Rheumatologie")
					FG1:= RegExReplace(FG1, "i).*Schlaf.*"                              	, "Innere Medizin/Schlafmedizin")
					FG1:= RegExReplace(FG1, "i).*Somnog.*"                             	, "Innere Medizin/Schlafmedizin")
					FG1:= RegExReplace(FG1, "i).*Somnol.*"                             	, "Innere Medizin/Schlafmedizin")
					FG1:= RegExReplace(FG1, "i).*Labor.*"                              	, "Labormedizin")
					FG1:= RegExReplace(FG1, "i).*Epileps.*"                            	, "Neurologie")
					FG1:= RegExReplace(FG1, "i).*Neu*ru*o*logi*e*.*"                	, "Neurologie")
					FG1:= RegExReplace(FG1, "i).*Neuropraxis.*"                    	, "Neurologie")
					FG1:= RegExReplace(FG1, "i).*Demenz.*"                          	, "Neurologie")
					FG1:= RegExReplace(FG1, ".*EMG.*"                                  	, "Neurologie")
					FG1:= RegExReplace(FG1, "i).*Neurolo.*"                            	, "Neurologie")
					FG1:= RegExReplace(FG1, "i).*MS.*(Z*C*entrum|Beratung|Ambulanz).*", "Neurologie")
					FG1:= RegExReplace(FG1, "i).*Nukle*a*.*"                            	, "Neurologie")
					FG1:= RegExReplace(FG1, "^PT.*"                                    	, "Psychiatrie")
					FG1:= RegExReplace(FG1, "i)^PIA\s*.*"                                	, "Psychiatrie")
					FG1:= RegExReplace(FG1, "i).*Psychi*o*at.*"                       	, "Psychiatrie")
					FG1:= RegExReplace(FG1, ".*KJP.*"                                   	, "Kinder und Jugend Psychiatrie")
					FG1:= RegExReplace(FG1, "i).*Kinderar.*"                             	, "Pädiatrie")
					FG1:= RegExReplace(FG1, "i).*Kinderh.*"                              	, "Pädiatrie")
					FG1:= RegExReplace(FG1, "i).*Kinderkl.*"                             	, "Pädiatrie")
					FG1:= RegExReplace(FG1, "i).*Neuropä.*"                            	, "Pädiatrie")
					FG1:= RegExReplace(FG1, ".*SPZ.*"                                    	, "Pädiatrie")
					FG1:= RegExReplace(FG1, "i).*Path*ol.*"                            	, "Pathologie")
					FG1:= RegExReplace(FG1, "i).*Patolo.*"                             	, "Pathologie")
					FG1:= RegExReplace(FG1, "i).*Physiothera.*"                         	, "Physikalische Rehabiliative Medizin")
					FG1:= RegExReplace(FG1, "i).*Ph*ysia*ka.*"                          	, "Physikalische Rehabiliative Medizin")
					FG1:= RegExReplace(FG1, "i).*Rehabilit.*"                          	, "Physikalische Rehabiliative Medizin")
					FG1:= RegExReplace(FG1, "i).*Proktol.*"                            	, "Proktologie")
					FG1:= RegExReplace(FG1, "i).*Ph*s*ychot.*"                          	, "Psychotherapie")
					FG1:= RegExReplace(FG1, "i).*Psycholog.*"                           	, "Psychotherapie")
					FG1:= RegExReplace(FG1, "i).*Ps*ychot.*"                           	, "Psychotherapie")
					FG1:= RegExReplace(FG1, "i).*Psyscho.*"                           	, "Psychotherapie")
					FG1:= RegExReplace(FG1, "i).*Physchol.*"                           	, "Psychotherapie")
					FG1:= RegExReplace(FG1, "i).*Pscholo.*"                           	, "Psychotherapie")
					FG1:= RegExReplace(FG1, "i).*Ul*rolo.*"                               	, "Urologie")
					FG1:= RegExReplace(FG1, "UIrologie"                               	, "Urologie")
					FG1:= RegExReplace(FG1, "i).*Ortho*i*pädi*e*.*"                  	, "Orthopädie")
					FG1:= RegExReplace(FG1, "i).*Schulter.*"                            	, "Orthopädie")
					FG1:= RegExReplace(FG1, "i).*Gelenk.*"                            	, "Orthopädie")
					FG1:= RegExReplace(FG1, "i).*Knies.*"                               	, "Orthopädie")
					FG1:= RegExReplace(FG1, "i).*Wirbelsäulen.*"                   	, "Orthopädie")
					FG1:= RegExReplace(FG1, "i)WS\s*Zentrum"                        	, "Orthopädie")
					FG1:= RegExReplace(FG1, "i).*Radiologi.*"                          	, "Radiologie")
					FG1:= RegExReplace(FG1, "i).*Radoplogi.*"                          	, "Radiologie")
					FG1:= RegExReplace(FG1, ".*MRT.*"                                    	, "Radiologie")
					FG1:= RegExReplace(FG1, ".*PET.*"                                    	, "Radiologie")
					FG1:= RegExReplace(FG1, "i).*Mammo.*"                             	, "Radiologie")
					FG1:= RegExReplace(FG1, "i).*Röntg.*"                              	, "Radiologie")
					FG1:= RegExReplace(FG1, "i).*Neurora.*"                             	, "Radiologie")
					FG1:= RegExReplace(FG1, "i).*Schmerz.*"                            	, "Schmerztherapie")
					FG1:= RegExReplace(FG1, "i)^Merz.*"                              	, "Schmerztherapie")
					FG1:= RegExReplace(FG1, "i).*Str*ah*l*en.*"                         	, "Strahlentherapie")
					FG1:= RegExReplace(FG1, "i).*Charr*ite.*"                           	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Charit.*"                               	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Chrit.*"                                 	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Charite\s*AOZ.*"                  	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Chiropraktiker.*"                   	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Chirotherapie.*"                   	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Dr\s*Müller.*"                        	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Fuß.*"                                    	, "sonstige")
					FG1:= RegExReplace(FG1, ".*PRM.*"                                    	, "sonstige")
					FG1:= RegExReplace(FG1, ".*BFW.*"                                     	, "sonstige")
					FG1:= RegExReplace(FG1, ".*MVZ.*"                                     	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Hennigsdorf.*"                      	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Podi*ologie.*"                        	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Hellios\s*KH\s*Buch.*"            	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Khs\s*Buch.*"                         	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Urlaubsort.*"                         	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Pflegeheim.*"                         	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*1\s*Hilfe.*"                          	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Maria\s*Heims.*"                   	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Rettungsstelle.*"                   	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Weaning.*"                         	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Wening.*"                           	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Rheinische.*"                         	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Dresden.*"                          	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Ergoth.*"                             	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Sportar.*"                             	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Sportmed.*"                          	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Sonogra.*"                             	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Sono\sAbd.*"                        	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*ungezielt.*"                         	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*unbestimmt.*"                        	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Ultrascha.*"                           	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Gercke.*"                            	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Unfallkran.*"                          	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Behring.*"                           	, "sonstige")
					FG1:= RegExReplace(FG1, "i).*Psychosom.*"                         	, "sonstige")
					If StrLen(FG1) = 0
						continue
					vFG := StrReplace(FG1, "/", "")
					vFG := StrReplace(vFG, " ", "")
					Gui, rxS: Default
					;ToolTip, % "FG2: " FG2 "`nFG1: " FG1 "`nvFG: " vFG

					If !Ueberweisungen.Fachgruppen.HasKey(vFG)					{

							Ueberweisungen.Fachgruppen[vFG] := 1
							Gui, rxS: Default
							If (FGMax = 0) {
								Gui, rxS: Add, Text, % "xm y+20 w260 ", % FG1
							} else {
								Gui, rxS: Add, Text, % "xm w260 ", % FG1
							}
							Gui, rxS: Add, Text, % "x+5 cBlue Right vU" vFG, % "       1"
							Gui, rxS: Add, Progress, % "x+5 w" qw " h9 cBlue v" vFG, 1
							FGMax ++

					} else {
							UZ := Ueberweisungen.Fachgruppen[vFG] += 1
							vonHundert:= Floor(UZ*100/(qw*fk))

							GuiControl, rxS:, % "U" vFG	, % UZ ;SubStr("0000" UZ, -4)
							GuiControl, rxS:, % vFG     	, % vonHundert
					}

					Ueberweisungen.Fachgruppen[Max] := Ueberweisungen.Fachgruppen[vFG] > Ueberweisungen.Fachgruppen[Max] ? Ueberweisungen.Fachgruppen[vFG] : Ueberweisungen.Fachgruppen[Max]
					Ueberweisungen.alle += 1
					alle += 1

					If (Ueberweisungen.Fachgruppen[Max] > (qw*fk))
						fk ++

					;~ For FG, Ubw in Ueberweisungen.Fachgruppen 	{
						;~ vonHundert:= Floor(Ubw/(qw*fk)) * 100
						;~ vFG := RegExReplace(FG, "\s", "")
						;~ GuiControl, rxS:, % vFG, % vonHundert
					;~ }

					Gui, rxTP: Default
					GuiControl,, T4, % "verschiedene Fachgruppen: " FGMax
					GuiControl,, T5, % "Überweisungen insgamt    : " alle
			}

		}
	}

	FileDelete, % A_ScriptDir "\Fachgruppen.txt"
	For FG, Ubw in Ueberweisungen.Fachgruppen
		FileAppend, % FG "`t: " Ubw "`n", % A_ScriptDir "\Fachgruppen.txt"

return ;}
M14:                                                 	;	Kürzelliste und Nutzerkategorisierung      	;{

	KuerzelListe    	:= Object()
	kCount          	:= 0
	ExamQuartals	:= 0
	;~ KzlKategorien	:= {"Informationen":"-"
								;~ , "Leistungen": ["EBM", "GOÄ", "BG", "LABOR":["EBM", "GOÄ"]]
								;~ , "Überweisung" : "-"
								;~ , "Anamnese" : "-"
								;~ , "Befunde" : "-"
								;~ , "Fremdbefund" : "-"}

							;	, "Formularwesen":["REHA-Antrag"]}

	Fortschrittsanzeige("leeren")
	GuiControl, rxTP: , T3	, % "Ermitteln aller Aktenkürzel und Hinzufügen zur Kürzelliste"
	GuiControl, rxTP: , T4	, % "<><><><><><><><><><><><><><><>"

	Loop, 10	{

			JahrIndex	:= A_Index - 1
			QJahr   	:= "20" (10 + JahrIndex)

			Loop, 4			{

					ExamQuartals ++

					aktQuartal:= SubStr("0" A_Index, -1) (10 + JahrIndex)
					GuiControl, rxTP: , T5	, % "aktuell untersuchtes Quartal: " aktQuartal

					QData   	:= TPObject(aktQuartal, PatIDFilter, false, false)
					QKuerzel	:= QuartalsKuerzel_ermitteln(QData)

					For Albiskuerzel, val in QKuerzel			{
						;	GuiControl, rxTP: , T7	, % "akutelles Kürzel: " Albiskuerzel
						;	Sleep, 500
						If !KuerzelListe.HasKey(Albiskuerzel)
							kuerzelliste[Albiskuerzel] := "---"
					}

					GuiControl, rxTP: , T6	, % (Kuerzelliste.Count() - kCount) " Kürzel wurden der Liste hinzugefügt."
					kCount:= Kuerzelliste.Count()
			}
	}

	bytes := SerDes(Kuerzelliste, TagesprotokollPfad "\Aktenkuerzel.json", 3)

	GuiControl, rxTP: , T5	, % "Anzahl der untersuchten Quartale: " ExamQuartals
	GuiControl, rxTP: , T6	, % "Anzahl gefundener Aktenkürzel    : " Kuerzelliste.Count()

return ;}
M15:                                                	;	Fachgruppenthesaurus erstellen             	;{

	TVZ:=0
	FGMax 	:= 0
	FgTv := []

	If FileExist(A_ScriptDir "\data\Fachgruppen_Thesaurus.json")	{

			Fachgruppen	:= Object()	; child Object Fachgruppen inklusive Zählung (gesamt und Quartal, Jahr)
			Fachgruppen	:= new JSONFile(A_ScriptDir "\data\Fachgruppen_Thesaurus.json")

			Gui, FGT: New
			Gui, FGT: Add, TreeView	, % "xm ym w250 h900 gFachgruppenwahl vFachgruppen"
			For FgID, Fachname in Fachgruppen.Object().IDName			{
					If !isObject(FgID)
						FgTv[FgID]	:= TV_Add(Fachname)
					else					{
						FgTv[FgID]	:= TV_Add(Fachname.FG)
						FgTv[FgID]	:= []
						For FgSubID, Subspezialisierung in Fachname.FgID.Object(	)
							FgTv[FgID][FgSubID] := TV_Add(Subspezialisierung, FgTv[FgID])
					}
			}

			For ThWort, FgID in Fachgruppen.Object().Thesaurus			{

			}

			Gui, FGT: Add, Button, % "x+5  ym gStoreName"    	, % "<"
			Gui, FGT: Add, Button, % "y+5       gRemoveName"	, % ">"
			Gui, FGT: Add, ListView, % "x+5 ym w300 h900 gFachbezeichner vFachbezeichner" , % "in den Tagesprotokollen verwendete Bezeichnungen"

			If FileExist(TagesprotokollPfad "\Fachgruppenbezeichner.txt")			{
					FileRead, t, % TagesprotokollPfad "\Fachgruppenbezeichner.txt"
					Loop, Parse, t, `n, `r
					{
							RegExMatch(A_LoopField, "(.*)?\s+\[", bez)
							Bezeichner := Trim(bez1)
							For FgID, Facharzt in Fachgruppen.Object().IDName
								If (Facharzt = Bezeichner)
									continue
							For ThWort, FgID in Fachgruppen.Object().Thesaurus
								If (ThWort = Bezeichner)
									continue
							LV_Add("", Bezeichner)
					}
					t := bez := ""
			}
			Gui, FGT: Show, AutoSize, Fachgruppen Thesaurus
	}
	else	{
			MsgBox, % "Addendum - " A_ScriptName, % "Datei: '\data\Fachgruppen_Thesaurus.json' ist nicht vorhanden.`nDas Makro kann nicht ausgeführt werden."
			return
	}


	return
	For index, protokoll in Tagesprotokolle	{
			Jahr      	:= protokoll.Jahr
			Quartal	:= protokoll.Quartal "/" Jahr
			TPFullPath:= protokoll.path "\" protokoll.filename
			FileRead, tpDatei, % TPFullPath
			Gui, rxTP: Default
			GuiControl,, T3, % "Quartal: " Quartal
			l:=""
			Loop, Parse, tpDatei, `n, `r
			{
					If RegExMatch(A_LoopField, "\s*füb\s+([A-Za-zÄÖÜäöüß\s\/\.]+)[\,\(]", FG)					{

							FG:= StrReplace(Trim(FG1), " ", "-")
							l.= FG1 "`n"
							If !Fachgruppen.HasKey(FG)
									Fachgruppen[FG] := 0

							Fachgruppen[FG] += 1

							GuiControl,, T4, % "Fachgruppenbezeichnungen: " Fachgruppen.MaxCount()
					}
			}
	}


	Sort, l, U

	;FileDelete,  % TagesprotokollPfad "\Fachgruppen_Thesaurus.dic"

	 t:=""
	For key, val in Fachgruppen
		t.=  key "`t[" val "]`n"

	;FileAppend, % t, % TagesprotokollPfad "\Fachgruppenbezeichner.txt", UTF-8


return

StoreName: ;{

return ;}

RemoveName: ;{

return ;}

Fachgruppenwahl: ;{

return ;}

Fachbezeichner: ;{

return ;}

;}
M16:                                                	;	Praxisstatistiken                                     	;{

	ZiffernFilter	:= ["0300[1-5]", "03040"]
	PatIDFilter 	:= "1,2,17844" ; diese Patienten ID's werden ignoriert

	Statistik["AU"]	                                    	:= Object()
	Statistik["AU"]["Bescheinigungen"]           	:= 0
	Statistik["AU"]["Erstbescheinigungen"]     	:= 0
	Statistik["AU"]["Folgebescheinigungen"]   	:= 0
	Statistik["AU"]["Dauer_Alle"]                  	:= 0
	Statistik["AU"]["Patienten"]                      	:= 0
	Statistik["AU"]["Schein_Quote"]               	:= 0
	Statistik["AU"]["Patientenzahl"]                	:= 0

	Loop 11 {                                                                 	; 11 Jahre

		JahrIndex	:= A_Index - 1
		QJahr   	:= "20" (10 + JahrIndex)

		Loop, 4 {

				; AHK-Objekt zur Auswertung erstellen und anschliessend im JSON Format speichern
					aktQuartal 	:= SubStr("0" A_Index, -1) (10 + JahrIndex)
					QDFilePath	:= TagesprotokollPfad "\Quartalsdaten_" QJahr "-" SubStr("0" A_Index, -1) ".json"
					If !FileExist(TagesprotokollPfad "\" QuartalZuProtokollname(aktQuartal)) ;|| FileExist(QDFilePath)    	; Fortsetzen wenn Quartal nicht existiert oder schon konvertiert wurde
						continue
					QData      	:=	TPObject2(aktQuartal, PatIDFilter, false)
					bytes         	:=	SerDes(QData, QDFilePath, 3)
					bytesSum 	  +=	bytes

				; Anlegen eines neuen Quartal-Objektes
					NodeName := QJahr "[" A_Index "]"
					If !IsObject(Statistik[(NodeName)])
						Statistik[(NodeName)] := {AU:{}, Labor:{}, Rezepte:{}, Ueberweisungen:{}}

				; Aufruf der Funktion zur statistischen Auswertung von AU-Bescheinigungen
					AU_Statistik(QData, aktQuartal)
		}

	}

	;bytes	     	:= SerDes(Statistik, TagesprotokollPfad "\Auswertungen\AU Statistiken 2010-2019.json", 3)
	GuiControl, rxTP: , T3	, % "Auswertung der AU-Statistiken der Jahre 2010-2020 sind in den"
	GuiControl, rxTP: , T4	, % "Tagesprotokollpfad \Auswertungen geschrieben worden."
	GuiControl, rxTP: , T5	, % "Insgesamt " bytesSum " bytes im JSON-Format wurden geschrieben."



return ;}
M17:                                                  	;	Kürzel zählen                                        	;{

return ;}
M18:                                                	;	Rezepte zählen                                         	;{

	Rezepte		:= Object()
	Rezepte.medrp  	:= Object()
	Rezepte.medhm 	:= Object()
	Rezepte.medp    	:= Object()
	Rezepte.medbm	:= Object()

	rxRezept 	:= "^(?<Datum>\d{2}\.\d{2}\.\d{4})*\s*(?<Art>medrp|medhm|medp|medbm)\s*(?<Med>[A-Za-z\s]+)"
	RzZaehler	:= 0
	RzMedrp	:= 0
	RzMedhm	:= 0
	RzMedp	:= 0
	RzMedbm	:= 0
	RzLast   	:= ""
	medrp  	:= []
	medhm 	:= []
	medp    	:= []
    medbm 	:= []

	For index, protokoll in Tagesprotokolle 	{

		Jahr  	:= protokoll.Jahr
		Quartal	:= protokoll.Quartal "/" Jahr
		TPFullPath:= protokoll.path "\" protokoll.filename
		tpfile := FileOpen(TPFullPath, "r").Read()
		;ToolTip, % TPFullPath ",   " StrLen(tpDatei)
		Gui, rxTP: Default
		GuiControl,, T3, % "Quartal: " Quartal
		Loop, Parse, tpfile, `n, `r
		{
				lTP := A_LoopField
				If RegExMatch(lTP, "^(?<Datum>\d{2}\.\d{2}\.\d{4})\s+", "Akt") {
					RegExMatch(AktDatum, "\d{4}", Jahr)
				}

				If RegExMatch(lTP, rxRezept, Rz) {

					If (RzArt = "medrp") {
						RzMedRp ++
						GuiControl,, T4, % "medrp: " RzMedrp
					} else if (RzArt = "medbm") {
						RzMedbm ++
						GuiControl,, T5, % "medbm: " RzMedbm
					} else if (RzArt = "medhm") {
						RzMedhm ++
						GuiControl,, T6, % "medhm: " RzMedhm
					} else if (RzArt = "medp") {
						RzMedp ++
						GuiControl,, T7, % "medp: " RzMedp
					}

					GuiControl,, T8, % "gesamt: " RzMedrp + RzMedbm + RzMedhm + Rzmedp


				}

		}


	}

return
;}
;}



;-------------------------------------------------------------------------------------------------------------------------------------------------------------
; FUNKTIONEN
;-------------------------------------------------------------------------------------------------------------------------------------------------------------;{
QuartalsKuerzel_ermitteln(DataObj) {

	SammelObjekt := Object()

	For PatID, QObj in DataObj
		For Datum, TObj in DataObj[PatID]["BTag"] {
				If IsObject(TObj) {
						For Kuerzel, val in TObj	{
							If (StrLen(Kuerzel) > 0) && !SammelObjekt.HasKey(Kuerzel)
								SammelObjekt[Kuerzel] := "---"
						}
				}
		}

return SammelObjekt
}

QuartalZuProtokollname(QuartalsJahr) {

	Jahr    	:= SubStr(QuartalsJahr, -1)
	Quartal	:= LTrim(SubStr(QuartalsJahr, 1, StrLen(QuartalsJahr) - 2), "0") ; damit geht 0119 oder 119

return "20" Jahr "-" Quartal "_TP-AbrH.txt"
}



;"Tagesprotokoll": "32802" - für Sendmessage
;----------------------------------------------------------------------------------------------------------------------------------------------
; RegEx Funktionen
;---------------------------------------------------------------------------------------------------------------------------------------------- ;{
RegExParseTPFile(tprotfile, regextFilePath, RegExTP, deleteOrigin= false) {                  	;-- entfernt per RegEx für die weitere Verarbeitung unnötige Zeichen aus einem Tagesprotokol

	; tprotfile   	- 	String mit Pfadangabe zum zu parsenden Tagesprotokoll
	; regextFile	- 	Pfad- und Dateiname unter der das veränderte Tagesprotokoll gespeichert werden soll
	; RegExTP  	- 	ein indiziertes Key, value Objekt bestehend aus einem RegEx-String und den ersetzenden Zeichen - am Anfangsbereich des Skriptes zu finden
	;           	    	wenn schon ein neues Tagesprotokoll erstellt wurde, wird dieses in eine Variable eingelesen und von der Funktion zurück gegeben

	 ; Dateinamen aus dem Inhalt der Datei generieren
		SplitPath, tprotfile,,,, fileNoExt
		regextFile := RTrim(regextFilePath , "\") "\" fileNoExt "_TP-AbrH.txt"

	; Einlesen des Originals und Umwandeln von ANSI nach UTF-8
		TPText := ReadAndConvertFile(tprotfile, "CP1252")
		If (StrLen(TPText) > 0) {

			; wenn noch nicht geparsed
				If RegExMatch(TPText, "m)\\H\\N\+\d+")
					For index, arg in RegExTP
						TPText := RegExReplace(TPText, arg.regex, arg.replace)

			; Datei unter übergebenem Namen schreiben
				File := FileOpen(regextFile, "w", "UTF-8")
				File.Write(TPText)
				File.Close()

			; das Albisprotokoll löschen
				If deleteOrigin
					FileDelete, % tprotfile

			return TPText
		}

return
}

RegExGetAll(str, rxStr, ValName) {

	matches := Array()
	pos := 0

	While (pos := RegExMatch(str, rxStr, matched_, pos + 1))
		matches.Push(matched_%ValName%)

return matches
}

TrimDiagnose(str) {                                                                                               	;-- ändert vom Nutzer geänderte Diagnosentexte in eine maschinenlesbarere Form um
	str:= RegExReplace(str, "\,\s*(li|re|bds)\s*\.")
	str:= RegExReplace(str, "\,\s*G\.\s*")
	str:= RegExReplace(str, "\,*\s*Z\.n*\.*\s*")
	str:= RegExReplace(str, "\,\s*(L|R|B)\.\s*")
	str:= RegExReplace(str, "\{[\w\.\+\*\-\!]+\}\;*")
	str:= RegExReplace(str, "(Behandlung|anamnestisch)\s+")
	str:= RegExReplace(str, "\([\w\.\+\s\,\-\*]+\)")
	str:= RegExReplace(str, "ED(\.|\:)*\s*\d*[\^\.\s]*\d\d\d\d")
	str:= RegExReplace(str, "ED(\.|\:)*\s*\d\d")
	str:= RegExReplace(str, "\sED\s")
	str:= RegExReplace(str, "(nach)*\s(re|li)\.")
	str:= RegExReplace(str, "(nach)*\s(rechts|links)")
	;str:= RegExReplace(str, "\s\-\s")
	str:= RegExReplace(str, "(LWK|BWK|HWK|LWS|BWS|HWS)\s*[\/|\-|\,\d]+S*\d*")
	str:= RegExReplace(str, "(C|L)[\/|\-|\,\d]+S*\d*")
	str:= RegExReplace(str, "i)\sOP\s")
	str:= RegExReplace(str, "^\d+[\.\+]")
	str:= RegExReplace(str, "\d+[\.\+x]\s*(vor)*")
	str:= RegExReplace(str, "\d+[x\,\d]*\s*(mm|cm|\%)")
	str:= RegExReplace(str, "\,*\s*\d+[\-\s]\d+")
	str:= RegExReplace(str, "(Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember)")
	str:= RegExReplace(str, "(Jan|Feb|Mär|Apr|Mai|Jun|Jul|Aug|Sep|Okt|No|Dez)\.\-*(Jan|Feb|Mär|Apr|Mai|Jun|Jul|Aug|Sep|Okt|No|Dez)*\.*")
	str:= RegExReplace(str, "Schulunfall")
	str:= RegExReplace(str, "\d\d\.\d\d\.\d\d\d\d")
	str:= RegExReplace(str, "\d\d[\.\-\^\/]\s*\d\d\d\d")
	str:= RegExReplace(str, "\d\d[\.\-\^\/]\s*\d\d")
	str:= RegExReplace(str, "\d\d\d\d\/*\d*")
	str:= RegExReplace(str, "(im\s)*\d+\s*LJ\.*(\s|$)")
	str:= RegExReplace(str, "\d+(\/|\^)\d+\s.*")
	str:= RegExReplace(str, "M\d+\/\d+\,*[R\d]*")
	str:= RegExReplace(str, "((c|p)*T[1234a]).*")
	str:= RegExReplace(str, "\(\.*\)")
	str:= RegExReplace(str, "\,*\s+(re|li)[\>\<]+(re|li)\.*")
	str:= RegExReplace(str, "\[.*\]")
	str:= RegExReplace(str, "\d+x")
	str:= RegExReplace(str, "seit(\sKindheit)*")
	str:= RegExReplace(str, "\#")
	str:= RegExReplace(str, "i)\,*\sZ\.*(\s|$)")
	str:= RegExReplace(str, "bek\.")
	str:= RegExReplace(str, "Größe")
	str:= RegExReplace(str, "(Beginn)\.*")
	str:= RegExReplace(str, "Gleason\s.*")
	str:= RegExReplace(str, "\s(re|li)$")
	str:= RegExReplace(str, "\,*und$")
	str:= RegExReplace(str, "\**")
	str:= RegExReplace(str, "an\sden\s.*")
	str:= RegExReplace(str, "der\sbeiden\s.*")
	str:= RegExReplace(str, "\:\s.*")
return Trim(str)
}


;}

;----------------------------------------------------------------------------------------------------------------------------------------------
; Datei / Daten Funktionen
;---------------------------------------------------------------------------------------------------------------------------------------------- ;{
TagesprotokollPfad(AddendumDir) {                                                                      	;-- liest den Tagesprotokollpfad ein

	IniRead, tpfilepath,  % AddendumDir "\Addendum.ini", Abrechnungshelfer, letzte_Tagesprotokolldatei
	If InStr(tpfilepath, "Error")	{

				FileSelectFile, tpfilepath,, % AddendumDir . "\Tagesprotokolle\" . A_YYYY, Tagesprotokoll auswählen
				IniWrite, % tpfilepath,  % AddendumDir . "\Addendum.ini", Abrechnungshelfer, letzte_Tagesprotokolldatei

	} else	{

				SplitPath, tpfilepath, lastprot
				MsgBox, 4, Addendum für Albis on Windows, % "Möchten Sie das letzte Tagesprotokoll`n" lastprot "`nfür diese Suche nutzen?"
				IfMsgBox, No
				{
						FileSelectFile	, tpfilepath,	, % AddendumDir "\Tagesprotokolle\" A_YYYY, Tagesprotokoll auswählen
						IniWrite     	, % tpfilepath	, % AddendumDir "\Addendum.ini", Abrechnungshelfer, letzte_Tagesprotokolldatei
				}

	}

return tpfilepath
}

TPMax(TProt) {                                                                                                      	;-- gibt die Anzahl der Zeilen der Datei und die Anzahl aller Patienten im Tagesprotokoll zurück

	Patients	:= 0
	fileLine	:= StrSplit(Tprot, "`n", "`r")

	Loop % fileLine.MaxIndex()
		If RegExMatch(fileLine[A_Index], RegExFind[1])
			Patients ++

return {"Lines": fileLine.MaxIndex(), "Patients":Patients}
}

ReadTPDir(dir, filepattern, ext, mode) {                                                                   	;-- liest ein Verzeichnis ein, ext=Dateiendung

	tlist:= Array()

	Loop, Files, % dir "\" filepattern "." ext, % mode
	{
		RegExMatch(A_LoopFileFullPath, "(?<Jahr>\d{4})\-(?<Quartal>\d)", TP)
		tlist.Push({"Jahr": TPJahr, "Quartal": "0" TPQuartal, "path": A_LoopFileDir, "filename": A_LoopFileName, "displayname": StrReplace(A_LoopFileName, "_TP-AbrH.txt", "")})
	}

return tlist
}

ReadAndConvertFile(FilePath, ReadEncoding="CP1252", WriteEncoding="") {          	;-- konvertiert ANSI nach UTF-8-BOM (besser!)

	; konvertiert ohne Umweg über die Erstelung einer Datei (Albis schreibt ANSI Dateien im Codepageformat CP1252)

	If FileExist(FilePath)	{

		txt := FileOpen(FilePath, "r", ReadEncoding).Read()
		If (StrLen(WriteEncoding) > 0) 	{
			SplitPath, FilePath,, Dir, fExt, fName
			f:= FileOpen(FilePath, "w", WriteEncoding)
			f.Write(txt)
			f.Close
		}

	}

return txt
}

TPFilesListe(tpfilesDir) {


}

SavePatList(Arr, FullFilePath, BackupPath:="") {                                                       	;-- zum Sichern der GB-Liste oder Chroniker-Liste (erstellt gleichzeitig Backups)

	If !IsObject(Arr)
		return

	; besser ein Backup anlegen
		SplitPath, FullFilePath,,, FileListNoExt, FileListExt
		If (StrLen(BackupPath) > 0) {

				BackupFile:= BackupPath "\" FileListNoExt "-" A_Year A_MM A_DD
				incr := 0
				while, FileExist(BackupFile . (incr = 0 ? "" : "-" incr) . FileListExt)
						incr ++
				FileCopy, % FullFilePath, % BackupFile . (incr = 0 ? "" : "-" incr) . FileListExt, 1

		}

	; Array in String umwandeln, sortieren und die Datei speichern
		StrToWrite:=""
		Loop, % Arr.MaxIndex()
			StrToWrite .= Arr[A_Index] "`n"
		Sort, StrToWrite, N U
		File:= FileOpen(FullFilePath, "w")
		File.WriteLine(StrToWrite)
		File.Close()

return ErrorLevel
}

ChronikerListe(FullFilePath) {                                                                                  	;-- Chroniker-Datei einlesen

	Chroniker := Array()

	Loop, Parse, % FileOpen(FullFilePath, "r", "UTF-8").Read(), `n, `r
		If RegExMatch(A_LoopField, "^\d+", PatID)
				Chroniker.Push(PatID)

return Chroniker
}

ChronikerICDListe(FullFilePath) {                                                                            	;-- ICD-Schlüsselnummerliste für Zuordnung zu chronische Erkrankungen

	CICD := Array()

	Loop, Parse, % FileOpen(FullFilePath, "r", "UTF-8").Read(), `n, `r
		If RegExMatch(A_LoopField, "^A-Z\d+[\.\d\-\*]*", ICD)
				CICD.Push(ICD)

return CICD
}

AU_Statistik(DataObj, aktQuartal) {                                                                       	; sucht nach AU-Scheinen (Kürzel fau) und stellt statistische Daten zusammen

	Jahr          		:= "20" SubStr(aktQuartal, -1)
	Quartal           	:= LTrim(SubStr(aktQuartal, 1, StrLen(aktQuartal) - 2), "0")
	QJahr				:= Jahr "[" Quartal "]"
	EQMonat      	:= SubStr("0" ((Quartal - 1) * 3 + 1), -1)                                                    	; erster Monat des Quartals
	EQTag           	:= "01"
	QStartTag	     	:= EQTag "." EQMonat "." Jahr

	AUPatientenListe	:= ""
	AUPatienten       	:= 0
	AUScheineGesamt	:= 0
	AUScheineErst   	:= 0
	AUScheineFolge	:= 0
	AUDauerAlle     	:= 0


	For PatID, QObj in DataObj
	{
			If (StrLen(AUDauer) > 0) {

					; in die Berechnung der Gesamt-AU-Zeiten (Dauer) fließen nur die Tage der letzten AU-Bescheinigung des Patienten ein, denn mit auf jeder Bescheinigung ist die
					; Dauer seit dem ersten Tag der Krankschreibung vermerkt. MIt der Funktion DateDiff und der folgenden If Abfrage wird sicher gestellt das nur die Tage die in den
					; Zeitraum des aktuellen Quartals fallen summiert werden:
					; 	lag der erste Tag der AU vor dem 1.Tag des Quartals wird Tagesdiff eine negative Zahl sein, diese Differenz wird dann von der ausgelesenen Dauer (auf der
					;	Folgeverordnung) abgezogen und ergibt somit die Krankheitstage die in dieses und in das nächste Quartal fallen (und hier ergibt sich eine Doppelanrechnung
					; 	von Tagen!!!)

					Tagesdiff:= DateDiff("dd", QStartTag, AUVon) + 0
					If (Tagesdiff < 0)
						AUDauer := AUDauer + Tagesdiff

					AUDauerAlle += AUDauer
					AUDauer := ""
			}
			lastPatID:= PatID

			For Datum, TObj in DataObj[PatID]["BTag"]
			{
					For kuerzel, val in TObj
					{
							If RegExMatch(Kuerzel, "^fau")
							{
									If PatID not in %AUPatientenListe%
										If StrLen(AUPatientenListe) = 0
											AUPatientenListe .= PatID
										else
											AUPatientenListe .= "," PatID

									AUText:= val.1
									RegExMatch(AUText, "^s*(?<Dauer>\d+)\s*[Tage]+\,\s*(?<Von>\d{2}\.\d{2}\.\d{4})\s*[a-z]*\s*(?<Bis>\d{2}\.\d{2}\.\d{4})\s*\((?<Nr>\w+)\)", AU)

									AUScheineGesamt ++
									If InStr(AUNr, "Erstbesch")
										AUScheineErst ++
									else
										AUScheineFolge ++
							}
					}

			}
	 }


	Statistik[QJahr]["AU"]["Bescheinigungen"]                        	:= AUScheineGesamt
	Statistik[QJahr]["AU"]["Erstbescheinigungen"]                  	:= AUScheineErst
	Statistik[QJahr]["AU"]["Folgebescheinigungen"]               	:= AUScheineFolge
	Statistik[QJahr]["AU"]["Dauer_Alle"]                        	      	:= AUDauerAlle
	Statistik[QJahr]["AU"]["Patienten"]                                   	:= AUPatientenGesamt                      	:= StrSplit(AUPatientenListe, ",").MaxIndex()
	Statistik[QJahr]["AU"]["Schein_Quote"]                            	:= Round(AUPatientenGesamt / Statistik[QJahr]["Patientenzahl"] * 100, 1) "%"
	Statistik[QJahr]["AU"]["Scheine_pro_Patient"]                   	:= Round(AUScheineGesamt / AUPatientenGesamt, 1 )
	Statistik[QJahr]["AU"]["Dauer_Durchschnitt"]                    	:= AUDauerSchnitt                            	:= Round(AUDauerAlle / AUPatientenGesamt, 1)
	Statistik[QJahr]["AU"]["Dauer_In_Wochen"]                     	:= AUDauerInWochen                       	:= Round(AUDauerAlle / 7, 1)
	Statistik[QJahr]["AU"]["Dauer_In_Monaten"]                    	:= AUDauerInMonaten                      	:= Round(AUDauerAlle / 30, 1)
	Statistik[QJahr]["AU"]["Dauer_In_Jahren"]                       	:= AUDauerInJahren                         	:= Round(AUDauerAlle / 365, 1)
	Statistik[QJahr]["AU"]["FormularAufwendungsSekunden"]	:= AUformularAufwendungsSekunden	:= AUScheineGesamt * 36                            	;(geschätzt 36s pro AU-Formular)
	Statistik[QJahr]["AU"]["FormularAufwendungsStunden"]  	:= AUformularAufwendungsStunden  	:= Round(AUScheineGesamt * 36 / 3600, 2)  	;(geschätzt 36s pro AU-Formular)

Ausgabe =
(
Statistik zu Arbeitsunfähigkeitsbescheinigungen
für das Quartal %aktQuartal%
---------------------------------------------------------

Anzahl arbeitsunfähiger Patienten               `t: %AUPatientenGesamt%

Anzahl AU-Bescheinigungen gesamt           `t: %AUScheineGesamt%
davon sind Erstbescheinigungen                `t: %AUScheineErst%
davon sind Folgebescheinigungen             `t: %AUScheineFolge%

Mindestzahl geleistete Unterschriften          `t: %UnterschriftenGesamt%
geschätzte Zeit f. Formularerstellung          `t: 1 AU-Schein circa 36s
gerechnet auf akt. AU Zahlen                    `t: %AuformularAufwendungsStunden% Stunden
... da hätten wir glatt zwei Tage schliessen können!

insgesamt bescheinigte AU-Zeiten (nur Tage
innerhalb des aktuellen 	oder folgenden Quartals)
                                               Tage      `t: %AUDauerAlle%
                                              Wochen  `t: %AUDauerInWochen%
										       Monate  `t: %AUDauerInMonaten%
												Jahre     `t: %AUDauerInJahren%

durchschnittl. bescheinigte AU-Tage`t: %AUDauerSchnitt%

Verursachter gesellschaftlicher Gesamtschaden:
`t-- nicht berechenbar

Verursachter Gewinn an Lebenserwartung für die Patient:
`t-- werden wir auch nicht herausbekommen

Verursachte Dankbarkeit Arbeitnehmer:
`t-- ENORM!

Verursachte Dankbarkeit Arbeitgeber:
`t-- es tut sich was!

)

	;Statistik[QJahr][Quartal]["Ausgabe"] := Ausgabe

	Statistik["AU"]["Bescheinigungen"]           	+=	AUScheineGesamt
	Statistik["AU"]["Erstbescheinigungen"]      	+= 	AUScheineErst
	Statistik["AU"]["Folgebescheinigungen"]   	+= 	AUScheineFolge
	Statistik["AU"]["Dauer_Alle"]                  	+= 	AUDauerAlle
	Statistik["AU"]["Patienten"]                      	+= 	AUPatientenGesamt
	Statistik["AU"]["Patientenzahl"]                	+= 	Statistik[QJahr]["Patientenzahl"]
	Statistik["AU"]["Schein_Quote"]               	:= 	Round(Statistik["AU"]["Patienten"] / Statistik["AU"]["Patientenzahl"] * 100, 1)


}

Ueberweisungen_Statistik(DataObj, aktQuartal) {

}

VorsorgelistenIDs2Arr(file) {                                                                                   	;-- liest eine vorhandene Gesundheitsvorsorgeliste (Praxomat) in ein indizierten Array

	btID		:= []
	VIndex 	:= 0

	If FileExist(file)
	{
			FileRead, vpFile, % file

			Loop, Parse, vpFile, `n, `r
     			If RegExMatch(A_LoopField, "(?<=\d\d\/\d\d;)\d+", Oid)
				{
					VIndex ++
					btID[VIndex] := Oid
				}

			return btID
	}
	else
			return ""
}

ReadPatientDatabase(PatDBPath) {										                        				;-- liest die .csv Datei Patienten.txt als Object() ein

	PatDB := Object()
	If !RegExMatch(PatDBPath, "\.json$")
		PatDBPath := Addendum.DBPath "\Patienten.json"

	If !FileExist(PatDBPath)
		return PatDB

return JSONData.Load(PatDBPath,, "UTF-8")
}

GetAddendumDbPath() {                                                                                       	;-- liest den Pfad zum Datenbankordner aus der Addendum.ini

	If (AddendumDir = "") {
			AddendumDir:= FileOpen("C:\albiswin.loc\AddendumDir","r").Read()
			If !AddendumDir {
					MsgBox, 262144 , % "Addendum - " A_ScriptName,
					(LTrim
						Der Pfad zu den Dateien für Albis on Windows ist nicht hinterlegt.
						Bitte starten Sie das AddendumStarter-Skript aus dem Addendum-
						Hauptverzeichnis, damit alle notwendigen Dateien und Verzeichnisse
						lokalisiert werden können.
						Das Skript wird jetzt beendet!
					), 15
					ExitApp
			}
	}

	IniRead, AddendumDbPath, % AddendumDir . "\Addendum.ini", Addendum, AddendumDbPath
	If InStr(AddendumDbPath, "Error") {
			MsgBox, 262144 , % "Addendum - " A_ScriptName,
			(LTrim
				Es wurde noch kein Datenbankpfad durch das Hauptskript Addendum.ahk
				angelegt. Diese Funktion greift auf Dateien in diesem Ordner zu.
				Bitte starten Sie Addendum.ahk!
			)
			ExitApp
	} else {
		return StrReplace(AddendumDbPath, "%AddendumDir%", AddendumDir)
	}

}

;}

;----------------------------------------------------------------------------------------------------------------------------------------------
; Objekt Funktionen
;----------------------------------------------------------------------------------------------------------------------------------------------
inArr(arr, Sstr) {                                                                                                     	;-- sucht nach einem Wert in einem indizierten Array

	;~ Loop, % arr.MaxIndex()
	For key, val in arr
		If (val = Sstr)
			return true

return false
}

ArrIn(Str, arr) {                                                                                                      	;-- gibt true oder false
}

TPObject2(ConvQuartal, PatIDFilter:="", debug:=true, ShowProgress:=true) {            	;-- konvertiert Tagesprotokolle (Abrechnungshelfer-Format!) in ein AHK-Objekt

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Variablen festlegen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		global 	MainNode, Datum, Subnode

		QData    	:= Object()
		TPFilepath	:= TagesprotokollPfad "\" QuartalZuProtokollname(ConvQuartal)
		QJahr      	:= "20" SubStr(ConvQuartal, -1)
		Quartal    	:= LTrim(SubStr(ConvQuartal, 1, StrLen(ConvQuartal) - 2), "0")                         	; damit geht 0119 oder 119
		LQMonat 	:= (Quartal - 1) * 3 + 1 + 2                                                                          	; letzter Monat des Quartals
		LQTag     	:= DaysInMonth(QJahr LQMonat) "." SubStr("0" LQMonat, -1) "." QJahr
		tprot        	:= FileOpen(TPFilepath, "r").Read()
		TP          	:= TPMax(tprot)                                                                                                ; zählt vorab die Zeilen und gibt die Anzahl gefundener PatientenID's zurück

		Statistik[QJahr][Quartal]["Patientenzahl"] := TP.Patients
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Fortschrittsanzeige
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		If ShowProgress		{
			Gui, rxTP: default
			GuiControl, rxTP: , T3, % "Tagesprotokoll (" ConvQuartal ") für Verarbeitung strukturieren"
			GuiControl, rxTP: , T4, % "Zeilenzahl     `t: " 	TP.Lines
			GuiControl, rxTP: , T5, % "Patientenzahl  `t: " 	TP.Patients
			GuiControl, rxTP: , T6, % "<><><><><><><><><><><><><><><>"
		}

		If debug
			debugFile:= FileOpen(A_ScriptDir "\Debug-" ConvQuartal ".txt", "w")
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Konvertieren des Tagesprotokolls
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		For lineNr, line in StrSplit(tprot, "`n", "`r") {

			;----------------------------------------------------------------------------------------------------------------------------------------------
			; Fortschrittsanzeige:   	nicht jede bearbeitete Zeile quittieren, dies verlangsamt unnötig die Verarbeitungsgeschwindigkeit
			;----------------------------------------------------------------------------------------------------------------------------------------------
				If ShowProgress	&& (Mod(lineNr, 200) = 0) {
					GuiControl, rxTP: , T7 	, % "Zeilen bearbeitet  `t: " 	SubStr("00000" lineNr, -4) "/" SubStr("00000" TP.Lines, -4)
					GuiControl, rxTP: , T10 	, % "nodebuffer Länge`t: " 	SubStr("00000" StrLen(nodebuffer), -4)
					GuiControl, rxTP: , T11 	, % "Knotennamen         `t: "	MainNode "/" (Datum ? Datum "/" : "") (SubNode ? SubNode "/" : "")
				}

			; ---------------------------------------------------------------------------------------------------------------------------------------------
			; nächste Zeile:            	bei zutreffender Bedingung die aktuelle Zeile nicht verwerten
			;----------------------------------------------------------------------------------------------------------------------------------------------
				If GoNextPatient || (StrLen(line) = 0) || RegExMatch(line, "\\[HP]|Keine Eintragung innerhalb der Selektion vorhanden") ; z.B. \HKeine Diagnosen vorhanden! oder \HKeine Leistungen vorhanden!
					continue

			; ---------------------------------------------------------------------------------------------------------------------------------------------
			; neuer Patient: 	         	Daten des Patienten werden extrahiert, nodebuffer sichern
			;----------------------------------------------------------------------------------------------------------------------------------------------
				If RegExMatch(line, ".*\,\s+.*\(\d+\).*\,\s+.*\(.*\,\s+.*\)") {

					; nodebuffer speichern
						If PatID
							AddToTPObject2(PatID, nodebuffer)

					; Patientendaten ermitteln
						GoNextPatient	:= false
						RegExMatch(line         	, "(?<Name>.*?)\((?<ID>\d+)\)\s+(?<Geb>\d{2}\.\d{2}\.\d{4})\s*\,\s*(?<Kasse>.*)"	, Pat)
						RegExMatch(PatKasse 	, "(?<name>.*?)\(IK\:\s*(?<IK>\w+)\,\s+VNR\:\s*(?<VNR>\w+)"                                	, Kassen)

					; PatientenID-Filter um bestimmte Akten nicht zu untersuchen (z.B. weil eine Akte für Notizen genutzt wird)
						If PatID in %PatIDFilter%
						{
							GoNextPatient := true
							continue
						}

					; Patienten-Index erhöhen, Variablen zurücksetzen, QData Objekt erweitern
						PatIndex ++
						MainNode := Datum := Subnode := nodebuffer := ""
						QData[PatID] := {	"Nachname"  	: Trim(StrSplit(PatName, ",").1)
												,	"Vorname"        	: Trim(StrSplit(PatName, ",").2)
												,	"GebDatum"  	: PatGeb
												,	"Krankenkasse"	: {"Name": Kassenname, "IK": KassenIK, "VNR": KassenVNR}
												,	"Dia"             	: {}
												,	"Ziffern"         	: {}
												,	"ZLabor"        	: {}
												,	"PathoLabor" 	: {}
												,	"DDia"          	: {}
												,	"DMed"         	: {}
												,	"Cave"           	: {}
												,	"BTag"             	: {}}

					; Fortschrittsanzeige
						If ShowProgress						{
							GuiControl, rxTP: , T8 	, % "aktuelle Patienten-ID`t: " PatID
							GuiControl, rxTP: , T9 	, % "Patienten bearbeitet `t: " PatIndex
						}
						If debug
							debugFile.WriteLine("NEUER PATIENT : " PatID)

						continue
				}

			; ---------------------------------------------------------------------------------------------------------------------------------------------
			; NewNode:              	per RegEx alle Informationen der aktuellen Zeile erhalten
			;----------------------------------------------------------------------------------------------------------------------------------------------
			; ^((?<MN1>\d\d\.\d\d\.\d\d\d\d)|(?<MN2>[A-ZÄÖÜ][a-zäöüß!]+\:))*\s*(?<SubNode>^[A-Za-z\dÄÖÜäöü]+\s{2,}|\s{2,}[A-Za-z\dÄÖÜäöü]+)*(?<Value>.*)
				RegExMatch(line, "^\d\d\.\d\d\.\d\d\d\d"  	, NewMN1)
				RegExMatch(line, "^[A-ZÄÖÜ][a-zäöüß!]+\:"	, NewMN2)
				If NewMN1
					line := RegExReplace(line, "^" NewMN1 "\s*")
				else If NewMN2
					line := RegExReplace(line, "^" NewMN2 "\s*")
				RegExMatch(line,"^(?<SubNode>[A-Za-z\dÄÖÜäöü]+\s{2,})*\s*(?<Value>.*)", New)

			;----------------------------------------------------------------------------------------------------------------------------------------------
			; neuer Hauptknoten:  	Knotenname festlegen oder wenn kein Hauptknoten Daten an nodebuffer angehängen
			;----------------------------------------------------------------------------------------------------------------------------------------------
				If (StrLen(NewMN1) > 0) {
					NewMainNode	:= "BTag"
					NewDatum      	:= NewMN1
					GuiControl, rxTP: , T12 	, % "NewMainNode    `t: " 	NewMainNode ", " NewSubNode
					GuiControl, rxTP: , T13 	, % "NewMN               `t: " 	(NewMN1 ? NewMN1 : " ---------- ") ", " (NewMN2 ? NewMN2 : " ---------- ")
					;ListVars
				}
				else If (StrLen(NewMN2) > 0) {
					NewMainNode	:= InStr(NewMN2, "Dauerdiagnosen") ? "DDia" : InStr(NewMN2, "Dauermedikamente") ? "DMed" : InStr(NewMN2, "Cave!") ? "Cave" : ""
					NewSubNode	:= ""
					NewDatum      	:= ""
					GuiControl, rxTP: , T12 	, % "NewMainNode    `t: " 	NewMainNode ", " NewSubNode
					GuiControl, rxTP: , T13 	, % "NewMN               `t: " 	(NewMN1 ? NewMN1 : " ---------- ") ", " (NewMN2 ? NewMN2 : " ---------- ")
					;ListVars
				}
				else
					nodebuffer .= NewValue

			;----------------------------------------------------------------------------------------------------------------------------------------------
			; nodebuffer:             	speichern bei neuem Subknotennamen oder Hauptknotennamen
			;                                	Festlegung der aktuellen Knotennamen
			;----------------------------------------------------------------------------------------------------------------------------------------------
				NewSubNode := Trim(NewSubNode)
				If (StrLen(NewSubNode) > 0) || (StrLen(NewMainNode) > 0) {

					; nodebuffer speichern
						AddToTPObject2(PatID, nodebuffer)

					; Knotennamen, nodebuffer setzen
						nodebuffer  	:= NewValue
						MainNode	:= StrLen(NewMainNode)	> 0       	? NewMainNode	: MainNode
						Datum      	:= StrLen(NewDatum)      	> 0      	? NewDatum	    	: MainNode = "BTag" ? Datum : ""
						SubNode     	:= StrLen(NewSubNode) 	> 0      	? NewSubNode 	: ""
						GuiControl, rxTP: , T11 	, % "Knotennamen         `t: "	MainNode "/" (Datum ? Datum "/" : "") (SubNode ? SubNode "/" : "")
						Sleep 2000

					; QData Objekt - erweitern je nach Bedarf
						If (MainNode = "BTag") && !IsObject(QData[PatID][MainNode][Datum])
							QData[PatID][MainNode][Datum] := Object()
						If (MainNode = "BTag") && !IsObject(QData[PatID][MainNode][Datum][SubNode])
							QData[PatID][MainNode][Datum][SubNode] := Object()

				}

		}
	;}

		If debug
			DebugFile.Close()

return QData
}
AddToTPObject2(PatID, nodebuffer) {

	; globale Funktion auf das QData Objekt

	; Variablen
		global 	MainNode, Datum, Subnode
		static 	trenner := {"lko":"-", "lkü":"-", "lp":"-", "lle":"-", "DMed":";", "DDia":";", "Cave":";", "labor":";"}

		StrDelimiter := trenner.HasKey(SubNode) ? trenner[SubNode] : trenner.HasKey(MainNode) ? trenner[MainNode] : ""

	; Daten sichern
		If (StrLen(StrDelimiter) > 0) {                    ; Daten sind trennbar

			For idx, val in StrSplit(nodebuffer, StrDelimiter) {

				val := Trim(val)
				If (StrLen(val) = 0)
					continue

				If     	  (MainNode = "DDia")
					QData[PatID][MainNode].Push({"BhArt":SubNode, "ICD":val})
				else if (MainNode = "DMed") || (MainNode = "Cave")
					QData[PatID][MainNode].Push(val)
				else
					QData[PatID][MainNode][Datum][SubNode].Push(val)

				If debug
					debugFile.WriteLine("`t`t>write to: " . MainNode . "/" . (Datum ? Datum . "/" : "") . SubNode ": " . val)

			}

		}
		else {                                                	; Daten sind nicht trennbar

			QData[PatID][MainNode][Datum][SubNode].Push(nodebuffer)

		}


}

TPObject(ConvQuartal, IDFilter:="", debug:=true, ShowProgress:=true) {                	;-- konvertiert Tagesprotokolle (Abrechnungshelfer-Format!) in ein AHK-Objekt

	; Variablen festlegen

		global QData:= Object()
		global nodebuffer, node, Subnode, Databuffer, Kuerzel

		TPFilepath     	:= TagesprotokollPfad "\" QuartalZuProtokollname(ConvQuartal)
		QJahr          		:= "20" SubStr(ConvQuartal, -1)
		Quartal           	:= LTrim(SubStr(ConvQuartal, 1, StrLen(ConvQuartal) - 2), "0")                         	; damit geht 0119 oder 119
		LQMonat      	:= (Quartal - 1) * 3 + 1 + 2                                                                          	; letzter Monat des Quartals
		LQTag           	:= DaysInMonth(QJahr LQMonat) "." SubStr("0" LQMonat, -1) "." QJahr

	; vorgeparstes Tagesprotokoll einlesen
		FileRead, tprot, % TPFilepath

	; zählt vorab die Zeilen und gibt die Anzahl gefundener PatientenID's zurück
		TP := TPMax(tprot)
		Statistik[QJahr][Quartal]["Patientenzahl"] := TP.Patients

	; Fortschrittsanzeige
		If ShowProgress		{

			Gui, rxTP: default
			GuiControl, rxTP: , T3, % "Tagesprotokoll (" ConvQuartal ") für Verarbeitung strukturieren"
			GuiControl, rxTP: , T4, % "Zeilenzahl     `t: " TP.Lines
			GuiControl, rxTP: , T5, % "Patientenzahl  `t: " TP.Patients
			GuiControl, rxTP: , T6, % "<><><><><><><><><><><><><><><>"

		}

		If debug
			debugFile:= FileOpen(A_ScriptDir "\Debug-" ConvQuartal ".txt", "w")

	; Parsen des als Variable eingelesenen Tagesprotokolles
		Loop, Parse, tprot, `n, `r
		{

			; nicht jede bearbeitete Zeile quittieren, dies verlangsamt unnötig die Verarbeitungsgeschwindigkeit
				If ShowProgress		{
					If (Mod(A_Index, 50) = 0) || (A_Index+100 > TP.Lines)
						GuiControl, rxTP: , T7 	, % "Zeilen bearbeitet  `t: " SubStr("00000" A_Index, -4) "/" SubStr("00000" TP.Lines, -4)
				}

			; ---------------------------------------------------------------------------------------------------------------------------------------------
			; leere Zeile mit der nächsten Zeile weiter machen
				If (StrLen(A_LoopField) = 0) || RegExMatch(A_LoopField, "\\[HP]") ; z.B. \HKeine Diagnosen vorhanden!
					continue

			; ---------------------------------------------------------------------------------------------------------------------------------------------
			; Anfang eines Patientenbereiches wird erkannt und die Daten des Patienten werden extrahiert
				If RegExMatch(A_LoopField, ".*\,\s+.*\(\d+\).*\,\s+.*\(.*\,\s+.*\)") {

						GoNextPatient	:= false
						RegExMatch(A_LoopField, "(?<Name>.*?)\((?<ID>\d+)\)\s+(?<Geb>\d{2}\.\d{2}\.\d{4})\s*\,\s*(?<Kasse>.*)"	, Pat)
						RegExMatch(PatKasse	 , "(?<name>.*?)\(IK\:\s*(?<IK>\w+)\,\s+VNR\:\s*(?<VNR>\w+)"                               	, Kassen)

					; PatientenID Filter um bestimmte Akten nicht zu untersuchen (z.B. weil eine Akte für Notizen genutzt wird)
						If PatID in %IDFilter%
						{
							GoNextPatient := true
							continue
						}

					; Patienten-ID sichern, Patienten-Index erhöhen, Variablen zurücksetzen
						QData[PatID]                    	:= Object()
						QData[PatID].Nachname   	:= Trim(StrSplit(PatName, ",").1)
						QData[PatID].Vorname      	:= Trim(StrSplit(PatName, ",").2)
						QData[PatID].GebDatum   	:= PatGeb
						QData[PatID].Krankenkasse  	:= {"Name": Kassenname, "IK": KassenIK, "VNR": KassenVNR}
						QData[PatID].Dia               	:= Object()
						QData[PatID].Ziffern            	:= Object()
						QData[PatID].ZLabor            	:= Object()
						QData[PatID].DDia            	:= {"anamnestisch": {}, "Behandlung": {}}
						QData[PatID].DMed           	:= Object()
						QData[PatID].Cave            	:= Object()
						QData[PatID].BTag             	:= Object()

						Datum:= Kuerzel:= Databuffer 	:= ""
						node:= Subnode:= SubDiagnose:= nodebuffer	:= ""
						Dauerinfo := true
						PatIndex ++

					; Fortschrittsanzeige
						If ShowProgress						{
							GuiControl, rxTP: , T8 	, % "aktuelle Patienten-ID`t: " PatID
							GuiControl, rxTP: , T9 	, % "Patienten bearbeitet `t: " PatIndex
						}
						If debug
							debugFile.WriteLine("NEUER PATIENT : " PatID)

				}

			; ---------------------------------------------------------------------------------------------------------------------------------------------
			; IDFilter:                    	 ist dieser aktiv, dann wird der Rest der Abfragen bis zu nächsten PatientenID übergangen
				If GoNextPatient
					continue

			; ---------------------------------------------------------------------------------------------------------------------------------------------
			; Datum:                    	 am Zeilenanfang markiert den Beginn der Quartalsdaten und steht für den Behandlungstag
				If RegExMatch(A_LoopField, "^\d\d\.\d\d\.\d\d\d\d") {

					; speichert Daten (Cave!:)
						If DauerInfo {
							Dauerinfo := false
							If (StrLen(nodebuffer) > 0) {
								val:= StrSplit(RTrim(nodebuffer, ";"), ";")
								Loop, % val.MaxIndex()			{
									toStore := Trim(val[A_Index])
									QData[PatID][node].Push(toStore)
									If debug
										debugFile.WriteLine("`t`t>write to: " node ", subnode: " Subnode ", buffer: " toStore)
								}
							}
							nodebuffer:= node:= subnode := ""
						}

					; Sichern gefundener Daten
						If (StrLen(Databuffer) > 0) {
							AddToTPObject(PatID, Datum, Kuerzel, Databuffer)
							Kuerzel:= Databuffer := ""
						}

					; neues Datum auslesen
						RegExMatch(A_LoopField, "^\d\d\.\d\d\.\d\d\d\d", Datum)
						QData[PatID]["BTag"][Datum]:= Object()

						If debug
							debugFile.WriteLine("`t- neues Datum : " Datum)

					; eine Datumszeile ohne Eintragung, es wird mit der nächsten Zeile weiter gemacht
						If InStr(A_LoopField, "Keine Eintragung innerhalb der Selektion vorhanden")
							continue

				}

			; ---------------------------------------------------------------------------------------------------------------------------------------------
			; Patientendaten:       	Dauermedikamente, Dauerdiagnosen und Cave! werden verarbeitet
				If DauerInfo {

				  ; findet Dauerdignose: oder Dauermedikamente: oder Cave!:
				  ; speichert zuvor gefundene Daten unter dem bisherigen Knotenpunkt
				  ; legt den aktuellen Knotenpunkt fest
					If RegExMatch(A_LoopField, "^[A-ZÄÖÜ][a-zäöüß!]*\:\s*$", Dauer) {

						; nodebuffer	- Inhalt speichern und leeren im Anschluß
							If (StrLen(nodebuffer) > 0) {

								For idx, val in StrSplit(RTrim(nodebuffer, ";"), ";") {
									QData[PatID][node].Push(Trim(val))
									If debug
										debugFile.WriteLine("`t`t>write to: " node ", subnode: " Subnode ", buffer: " val)
								}
								nodebuffer := ""

							}

						; node      	- Knotenpunktname festlegen
							If InStr(Dauer1, "Dauerdiagnosen")
								node := "DDia"
							else If InStr(Dauer1, "Dauermedikamente")
								node := "DMed"
							else If InStr(Dauer1, "Cave")
								node := "Cave"

							If debug
								debugFile.WriteLine("`t`t*found node : " node " at line: " A_Index ", orig.RegMatch: " Dauer1)

    						continue

					}

				  ; Knotenpunkt - Dauerdiagnosen
					If InStr(node, "DDia") {

						; matched anamnestisch oder Behandlung + Diagnosentext
							RegExMatch(A_LoopField, "^(?<nodeN>\w+)*\s+(?<Diagnose>.*)", Sub)
							Subnode 	 := (StrLen(SubnodeN) > 0) ? SubnodeN : Subnode
							nodebuffer .= Trim(SubDiagnose) " "

						; eine Diagnose endet mit einem Semikolon, endet die aktuelle Zeile dann können die Daten gespeichert werden
							If RegExMatch(nodebuffer, "\;\s*$") {

								tmpArr := StrSplit(Trim(nodebuffer), ";")
								For idx, thisdiagnose in tmpArr
									QData[PatID]["DDia"][Subnode].Push(thisdiagnose)
								nodebuffer := ""

								If debug
									debugFile.WriteLine("`t`t->write to: " node ", subnode: " Subnode ", buffer: " Trim(nodebuffer))

							}

					} ; ENDE DDIA
					else If InStr(node, "DMed") {

						RegExMatch(A_LoopField, "^(?<node>\w+)\s+(?<Diagnose>.*)", Sub)
						nodebuffer := SubDiagnose

						If debug
							debugFile.WriteLine("`t`t+add to: " node ", subnode: " Subnode ", buffer: " nodebuffer)

				}
				else {
						nodebuffer .= A_LoopField
				}

				continue
				}

			; ---------------------------------------------------------------------------------------------------------------------------------------------
			; Fortsetzungsbereich: 	Datenpuffer befüllen wenn ein Fortsetzungsbereich erkannt wurde
				RegExMatch(A_LoopField, "^\s+", ws)
				trailing_WhiteSpaces := StrLen(ws)
				;if RegExMatch(A_LoopField, "^\s{13,}(.*)") && !RegExMatch(A_LoopField, "\s{2,12}([A-Za-zÄÖÜäöüß]+\s*[A-Za-zÄÖÜäöüß]+)\s{2,}.+")	{
				if (trailing_WhiteSpaces > 12) {

						RegExMatch(A_LoopField, "\s+(?<Eintragung>.*)", A_)
						Databuffer .= A_Eintragung "|"

						If debug
							debugFile.WriteLine("`t`t" Datum ", " Kuerzel ": " Databuffer)
				}
			; ---------------------------------------------------------------------------------------------------------------------------------------------
			; Aktenkürzel:            	prüft und ermittelt in der Zeile das Vorhandensein eines Aktenkürzel
				else  {	;RegExMatch(A_LoopField, "\s{2,12}([A-Za-zÄÖÜäöüß]+)\s{2,}.+")

					; 1. Sichern der bisher gesammelten Daten
						AddToTPObject(PatID, Datum, Kuerzel, Databuffer)

					; 2. Ermitteln des neuen Kürzel
						If RegExMatch(A_LoopField, "^\d{2}\.\d{2}\.\d{4}")
							RegExMatch(A_LoopField, "\s{2,}(?<Kuerzel>[A-Za-zÄÖÜäöüß\s]+?)\s{2,}(?<Eintragung>.*)", A_)
						else
							RegExMatch(A_LoopField, "\s+(?<Kuerzel>[A-Za-zÄÖÜäöüß\s]+?)\s{2,}(?<Eintragung>.*)", A_)

						Kuerzel      	:= Trim(A_Kuerzel)
						Databuffer	:= LTrim(A_Eintragung)

						If debug
							debugFile.WriteLine("`t- neues Kürzel : " Kuerzel "`n`t`t" Datum ", " Kuerzel ": " Databuffer)

				}
		}

		If debug
			DebugFile.Close()

return QData
}
AddToTPObject(PatID, Datum, Kuerzel, Databuffer) {                                              	; gehört zu TPObject()

		global QData

		;Leistungskomplexe: werden an den Behandlungstagen und zusätzlich gemeinsam unter .Ziffern gesammelt
		If RegExMatch(Kuerzel, "lk[onü]")	{

				Databuffer:= StrReplace(Databuffer, "|", "")
				Ziffern    	:= StrSplit(Databuffer, "-")

				If !IsObject(QData[PatID]["BTag"][Datum][Kuerzel])
					QData[PatID]["BTag"][Datum][Kuerzel] := Object()

				Loop % Ziffern.MaxIndex()		{
						ziffer := Trim(Ziffern[A_Index])
						If (StrLen(ziffer) > 0)						{
								QData[PatID]["BTag"][Datum][Kuerzel].Push(ziffer)

								For ZiffIndex, QZiffer in QData[PatID]["Ziffern"]		{
										RegExMatch(QZiffer, "(?<Komplex>\d+)\(*x*\:*(?<Faktor>\d+)*\)*", Q)
										RegExMatch(ziffer, "(?<Komplex>\d+)\(*x*\:*(?<Faktor>\d+)*\)*", T)
										If (QKomplex = TKomplex)		{
												QFaktor := QFaktor=""	? 1: QFaktor
												TFaktor	 :=  TFaktor=""	? 1: TFaktor
												QData[PatID]["Ziffern"][ZiffIndex]:= QKomplex "(x:" QFaktor + TFaktor ")"
												ZiffMatch:= true
												break
										}
								}

								If !ZiffMatch
									QData[PatID]["Ziffern"].Push(ziffer)
						}
				}

		}
		; Laborleistungen:
		else If InStr(Kuerzel, "lle")	{

				Databuffer:= StrReplace(Databuffer, "|", "")
				Ziffern    	:= StrSplit(Databuffer, "-")

				If !IsObject(QData[PatID]["BTag"][Datum]["lle"])
					QData[PatID]["BTag"][Datum]["lle"] := Object()

				Loop, % Ziffern.MaxIndex()
				{
						ziffer := Trim(Ziffern[A_Index])
						If (StrLen(ziffer) > 0)
						{
								QData[PatID]["BTag"][Datum]["lle"].Push(ziffer)

								For ZIndex, LBZiffer in QData[PatID]["ZLabor"]
								{
										RegExMatch(LBZiffer, "(?<Komplex>\d+)\s*\(*x*\:*(?<Faktor>\d+)*\)*", Q)
										RegExMatch(ziffer, "(?<Komplex>\d+)\s*\(*x*\:*(?<Faktor>\d+)*\)*", T)
										If (QKomplex = TKomplex)
										{
													QFaktor := QFaktor=""	? 1: QFaktor
													TFaktor	 :=  TFaktor=""	? 1: TFaktor
													QData[PatID]["ZLabor"][ZIndex]:= QKomplex "(x:" QFaktor + TFaktor ")"
													ZiffMatch:= true
													break
										}
								}

								If !ZiffMatch
									QData[PatID]["ZLabor"].Push(ziffer)
						}
				}

		}
		; Diagnosen: werden an den Behandlungstagen und zusätzlich gemeinsam unter .Dia gesammelt
		else if InStr(Kuerzel, "dia")	{

				Databuffer := StrReplace(Databuffer, "|", "")
				Diagnosen := StrSplit(Databuffer, ";")

				If !IsObject(QData[PatID]["BTag"][Datum]["dia"])
					QData[PatID]["BTag"][Datum]["dia"] := Object()

				Loop, % Diagnosen.MaxIndex()
				{
						Diagnosentext := Trim(Diagnosen[A_Index])
						If (StrLen(Diagnosentext) > 0)
						{
								QData[PatID]["BTag"][Datum]["dia"].Push(Diagnosentext)
								QData[PatID]["Dia"].Push(Diagnosentext)
						}
				}

		}
		; Labor: path. Werte welche in der Patientenkartei sichtbar sind (nicht Laborblatt!)
		else if InStr(Kuerzel, "labor")	{

				If !IsObject(QData[PatID]["BTag"][Datum]["labor"])
					QData[PatID]["BTag"][Datum]["labor"] := Object()

				Databuffer	:= StrReplace(Databuffer, "|", "")
				LBParameter := StrSplit(Databuffer, ";")

				Loop, % LBParameter.MaxIndex()
				{
						LBParam := Trim(LBParameter[A_Index])
						If (StrLen(LBParam) > 0)
							QData[PatID]["BTag"][Datum]["labor"].Push(LBParam)
				}

		}
		; Daten zu allen anderen Kuerzeln werden als einzelner String im Object hinterlegt
		else {

				If StrLen(Trim(Databuffer)) > 0
				{
						If !IsObject(QData[PatID]["BTag"][Datum][Kuerzel])
							QData[PatID]["BTag"][Datum][Kuerzel] := Object()

						Databuffer := RTrim(Databuffer, "|")
						Databuffer := Trim(Databuffer)
						QData[PatID]["BTag"][Datum][Kuerzel].Push(Databuffer)
				}
		}

}

;----------------------------------------------------------------------------------------------------------------------------------------------
; String Funktionen
;----------------------------------------------------------------------------------------------------------------------------------------------
StringSpacer(Str, StrWidth) {                                                                                   	;-- fügt einem String Leerzeichen hinzu (nur Space!!)
return SubStr(Str "                                                      ", 1, StrWidth)
}

;----------------------------------------------------------------------------------------------------------------------------------------------
; Objekte anzeigen
;----------------------------------------------------------------------------------------------------------------------------------------------
AutoListview(ColStr, Bezeichner, ColSize,ColType:="") {                                           	;-- eine Gui für die schnelle Anzeige von Daten

	;ColSize - als Prozentangabe der Gesamtbreite des Listview

	static Init, colmax, DasLV, hALV, ALV, hDasLV, ColSize_stored
	ColSize_stored:= ColSize

	If !Init {

			Init:= 1
			Sysget, Mon1, Monitor, 1

			If (Mon1Right>1920) {
					gx:= 2500, gy:= 50, gw:= 1000, gh:= 1024
					GMinW:= 400, GMinh:= 400 ;, gmaxw:= Floor(Mon1Right/3), gmaxh:= (Mon1Bottom - 100)
					gfontSize:= 11
					Rows:= 40
			} else {
					gx:= 1200, gy:= 50, gw:= 600, gh:= 400
					GMinW:= 400, GMinh:= 400 ;, gmaxw:= Floor(Mon1Right/3), gmaxh:= (Mon1Bottom - 100)
					gfontSize:= 8
					Rows:= 20
			}
	}

			;Größe des ListView-Controls
			LVw:= gw-30, LVh:= gh-10

			Gui, ALV:NEW, +ToolWindow
			Gui, ALV:Font, % "s" gFontSize , Futura Bk Bt
			Gui, ALV:Margin, 0, 0
			Gui, ALV:+LastFound +AlwaysOnTop +Resize +HwndhALV +MinSize%Gminw%x%Gminh% +Owner%hAbrH%
			Gui, ALV:Add, Listview, % "xm ym w" gw " h" gh " r" Rows " vDasLV HWNDhDasLV BackgroundAAAAFF Grid" , % ColStr 					;
			Gui, ALV:Show, % "x" gx " y" gy , % Bezeichner " - Addendum für Albis on Windows"
			Gui, ALV:Default

			Loop, Parse, ColSizeStored, `,
				DllCall("SendMessage", "uint", hDasLV, "uint", 4126, "uint", A_Index-1, "int", Floor(l.W * (A_LoopField / 100))) 	;sets the column width

			Loop, Parse, ColType, `,
				LV_ModifyCol(A_Index, A_LoopField)

Return

ALVGuiSize:
	Critical, Off
	if A_EventInfo = 1  ; Das Fenster wurde minimiert.  Keine Aktion notwendig.
		return
	Critical
	ALVw:= A_GuiWidth, ALVh:= A_GuiHeight, a:=""
	GuiControl, ALV: MoveDraw, %hDasLV%, % "w" . ALVw . " h" . ALVh
	l:= GetWindowSpot(hDasLV)
	Loop, Parse, ColSize_Stored, `,
		DllCall("SendMessage", "uint", hDasLV, "uint", 4126, "uint", A_Index-1, "int", Floor(l.W * (A_LoopField / 100))) 	;sets the column width
	Critical, Off
return

ALVGuiClose:
	Gui, ALV: Destroy
	ScriptMem()
return
}

ShowArray(Arr, Option := "w800 h500", GuiNum:= 90                                          	;-- aus dem Autohotkey Forum zur Anzeige eines Array's oder Objektes als Listview (debuggen z.B.)
, Colums:= "Nr|Werte", statustext:=" - MaxIndex des angezeigten Array") {

		static initGui:= []

		If !(init[GuiNum]) {
					Gui, %GuiNum%: Margin, 5, 5
					Gui, %GuiNum%: Add, ListView, %Option%, % "Nr|Werte" ;Columns
					Gui, %GuiNum%: Add, Statusbar
					init[GuiNum]:= 1
		}

		Gui, %GuiNum%: default
		;LV_Add("")


		Loop, % arr.MaxIndex()
			LV_Add("", A_Index, arr[A_Index])

		loop % LV_GetCount("Column")
			LV_ModifyCol(A_Index, "AutoHdr")

		SB_SetText("   " LV_GetCount() " " statustext)
		Gui, %GuiNum%: Show,, Array
}

;----------------------------------------------------------------------------------------------------------------------------------------------
; sonstige Funktionen
;----------------------------------------------------------------------------------------------------------------------------------------------
SetExplorerTheme(HCTL) {                                                                                        	;-- HCTL : handle of a ListView or TreeView control
   If (DllCall("GetVersion", "UChar") > 5) {
      VarSetCapacity(ClassName, 1024, 0)
      If DllCall("GetClassName", "Ptr", HCTL, "Str", ClassName, "Int", 512, "Int")
         If (ClassName = "SysListView32") || (ClassName = "SysTreeView32")
            Return !DllCall("UxTheme.dll\SetWindowTheme", "Ptr", HCTL, "WStr", "Explorer", "Ptr", 0)
   }
   Return False
}

ReadSearchPatterns(Section) {

	Loop {

			IniRead, SuchMuster,% A_ScriptDir "\Abrechnungshelfer.ahk", % Section, % "String" A_Index
			If !InStr(SuchMuster, "Error")
					Suchmuster[A_Index]:=  SuchMuster[A_Index]
			else
					break
	}

return Suchmuster
}

SaveSearchPattern(Section, Suchmuster, StrNr="") {

	If (StrNr = "" )
		Loop	{
			IniRead, dummy, % A_ScriptDir "\Abrechnungshelfer.ahk", % Section, % "String" A_Index
			If InStr(dummy, "Error") {
				StrNr := A_Index
				break
			}
		}

	IniWrite, % Suchmuster, % A_ScriptDir "\Abrechnungshelfer.ahk", % Section, % "String" StrNr

return ErrorLevel
}

Create_Abrechnungshelfer_ico(NewHandle := False) {                                           	;-- erstellt das ICON für die Taskbar
Static hBitmap := 0
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 2680 << !!A_IsUnicode)
B64 := "iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAEQwAABEMBS7v+bwAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAdVSURBVGiBzVptUFTXGX7OuZeP3WUXFncXJSJ7V60fwQbkaxrUWNBMbGwiBKnNCLRpE2PtD6WjadqZDPnTD7WmkzRJ06TTorXUr2BMzUwTNNqJQgKCFltpFHZBDMouCKwsLLv3nP5QHOJe5O5dNH3+7Xue877PM3vPxz3nEiigcj+PNvmxBjIrZFzOJUSYQSmNVeLeKzDGhsHlbkqFTwmhNbphHN6wgQTu5JE7A69U8WLG5d9QKsxy97QHXa4mcXDQjUBg5P4ov4WoqFgYTVZI9sVBq80hylzuoBC2VJSTmvG82wYqKzlNcGA7B37SdqmefXz8berxuO6r6IlgtUpYnv8smz07l3KCHYNt+GllJWHAOAO/3c13Ms4qjtW+SRob3v3q1N4FWTlFKCjYyEHpzopSsg24ZeCVKl4MggO1H72O/1fxY8jKKcKKFZvAOYoqykkNqdzPo40++WJ7W8PMgwd+TqeymMGQCElaDLuUicTEmdDpE6DTmyAHR+EbGkBvbyeczka0t30Kr7dXdd61Jb9kdinjsnE0ai7ZtYeXgPF9f3z7h5iqZ95mc+DhvPWYN38ZCAmZJ0IgsyAu/Ps46uqq0evpnJRvtUr4wbPvABxrRcJY0TW3M+DxuKIiFS5GxWDlyh/j6w+tUiV8DAIVkbboUSxcmI9PPtmN+rpqMMYm5LvdTrjd7UGLxV4oMhbM7XA1Ryw+MXEmCosqYbVJmnNQQcSyR56BXpeA2trX78p1uZpFc2LKN0QQMck76NZcFADM5mQ8vX4X4uKmRZRnDLG6uEk53kE3KBVsIqVUF8kiFRc3Dd9Zt/2u4m/c6MWli3Xounweg94eMDmIOJMV5oRkzJ+/DEnT54Zdd3R0GJRQg6hZ+S18a/VWJJhnKLb5/T6cPPEOzp39ALIcsgsAANSd/iusNgn5BRshSZlh14/IwINpBXA4shXbPB4XDh14CdevX5k0j7vHiX3V25C26FGsWrUlLA2aDQhiNPILNiq29fVdQfXerRga6gsr5/mWD3HD68HX5i1V3UezgQULlsNgMIfEZRbE4ZqXwxY/BperCd1X/6uar3nlzcxcoxg/03gYPdfatKYFAPhHhlRzNRkwm5MxI3leSJwxGQ2fHdSSUjM0GZg16yHFuMvZhEjXlHChyUBKarpivKOjKSIxWqBqEOctWY9Fix67/dtoVF60cnJLkJHxRNgiOjvP4oOjO8PuB6g0IDlyJlysxsNgMAOG8EU0NGgfN5MaoFRAQsJ0jIx4AQAEFDGxoSoZYxgdVT97jMeli/Wa+gEqDDAm43evltz+bTYnY8PGPSG8a1c/R9WfN2kWohVhD+KoaJ1i3D/qi1iMFoRtgBJBMc5kOWIxWqBqEFssdlisqQCA+PgkRY7RaMH8BY+EVdzn60dnx7mw+twJVQaWLC2bVJzVJmFN4UthFT99am/EBlQ9Qg88sDCiIhOhtfVkxDkm/QcEKuLUqb98KZa3pBRGoyWEW19Xjf7+q6oKyywY8aYPUGFAZkGcbf77l2IzU9KQlrYyhDs42BPCvdfQtBfq6mxRjNulrIjEaIEmA52dZxXjs+fkwmBIjEhQuNBkoK/vCrq6zofEBSoiJ7c4YlHhQPMb2bnmo4rx7OynYLHYtaYFcHP/pZqrtciF1pPwej2hCQURRcUvQ6czacqbOC0Fq7/9gmq+ZgPBgB/Hjr2pLCJxJtY9vQMmk011PkIIFj5YgPLvvQFT/HTV/SI6Tm/9zwk4nWcU25KS5uD7z/we6RmrIdCJZ2tCCGalpqO07FU88eTPEBOjhxwcVa0h4pO594/8AqVlr8FsTg5p0+nj8diqLchbUoa2tnp0f9GKoaHroITCaLLBYknBnLl5IYtiUA7DAOfMFx0Vq9dqwDfUj31/24bSstcUz4mAm6+g6emPIz39cVU5AwH/pJyYGD0Yk4coY8GeOJM1LNF3ov96N/bu2TxlFyRBFY9QnNECztk1KlCxTpIylU9ew0BfXxeq/rQJLS0fgnOuKQfnHC3/+geO1ypPDuMh2bOClAqnRUJojdUqfddqleB2OzUVHkMgMIKj7/8aDZ8dxNKl5Zgz92F1V0xyAK0X/onGxkPo/mLyY0WbzQGLNVUkBIdJ5X4eHTcsf+5qO5NyYP+LU3rJp9cnwC5lwm7PQHz8dOgNCYiJ1sPv92F4eBBudzsud51Hh6sZw74B1XlL1v2KpdoXdxj9wryb16y7eRGAQ8c+egMNDYem0sOUIyd3LfILngfnWFNRTt6jALCljLwLjp35K57n2dlPfdUaJ0ROTjG+mf8cB7C9opy8B4xbBwaceME0m/KClT/aapey2ImP/0AjHRNTBZvNgeX5zzGHI5sC2DHQjhfH2kJG2K4qXsi5vItSwe72OAMuZ1OUd6AHo/f5Y4/oqFiYTEmwSxlBi1USZSY7RUHYvLmUHBnPU5wi3nqLR/li8SQHK5SZnEuJkEwpVT4QukfgnPlkJndTQagXCK3R+XBE6XOb/wE5IowLirfeCwAAAABJRU5ErkJggg=="
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
VarSetCapacity(Dec, DecLen, 0)
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
DllCall("Gdiplus.dll\GdipCreateHICONFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
Return hBitmap
}

ScriptMem() {                                                                                                        	;-- gibt belegten Speicher frei

	static PID
	IfEqual, PID
	{
		DHW := A_DetectHiddenWindows
		TMM := A_TitleMatchMode
		DetectHiddenWindows, On
		SetTitleMatchMode	 , 2
		WinGet, PID, PID, % "\" A_ScriptName " ahk_class AutoHotkey"
		DetectHiddenWindows, % DHW
		SetTitleMatchMode	 , % TMM
	}
	h:=DllCall("OpenProcess", "UInt", 0x001F0FFF, "Int", 0, "Int", pid)
	DllCall("SetProcessWorkingSetSize", "UInt", h, "Int", -1, "Int", -1)
	DllCall("CloseHandle", "Int", h)

}

MessageWorker: ;{

	If InStr(InComing, "AutoLogin Disabled")
		MessageReceived:= InComing
	else if InStr(InComing, "PatData|")
		MessageReceived:= InComing

return ;}
;}

;-------------------------------------------------------------------------------------------------------------------------------------------------------------
; INCLUDES --- benötigte BIBLIOTHEKEN
;-------------------------------------------------------------------------------------------------------------------------------------------------------------;{
	#Include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Datum.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Menu.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_PdfHelper.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_PraxTT.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Screen.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Window.ahk

	#include %A_ScriptDir%\..\..\lib\class_JSON.ahk
	#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
	#Include %A_ScriptDir%\..\..\lib\Sift.ahk
	#Include %A_ScriptDir%\..\..\lib\ACC.ahk
	#Include %A_ScriptDir%\..\..\lib\ini.ahk
	#Include %A_ScriptDir%\..\..\lib\SerDes.ahk
	#Include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
	#Include %A_ScriptDir%\..\..\lib\class_JSONFile.ahk
	#Include %A_ScriptDir%\..\..\lib\JSON_FromObj.ahk
	#Include %A_ScriptDir%\..\..\lib\TVH.ahk
	#Include %A_ScriptDir%\..\..\lib\RMO.ahk
	#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk

;}

;-------------------------------------------------------------------------------------------------------------------------------------------------------------
; SUCHMUSTER
;-------------------------------------------------------------------------------------------------------------------------------------------------------------;{
;[freieStatistik]


;}




; Tastaturkürzelskript für Albis - nutzt Autohotkey_L - www.autohotkey.com
; entweder kompilieren oder nach Installation von Autohotkey die mit einem beliebigen Namen versehene Datei
; in den Autostart-Ordner kopieren, dann startet das Skript mit Windows und ein paar Bedienungen der Albisoberfläche sind wieder konsistenter

; Was kann das Skript?
;
;



; SKRIPTEINSTELLUNGEN
#NoEnv
#Persistent
#SingleInstance, Force
#InstallKeybdHook
#MaxThreads, 100
#MaxThreadsBuffer, On
#MaxHotkeysPerInterval 99000000
#HotkeyInterval 99000000
#MaxThreadsPerHotkey, 2
#KeyHistory 0

ListLines Off

SetTitleMatchMode, 2		;Fast is default
SetTitleMatchMode, Fast		;Fast is default

CoordMode, Mouse	, Screen
CoordMode, Pixel 	, Screen
CoordMode, ToolTip	, Screen
CoordMode, Caret	, Screen
CoordMode, Menu	, Screen

SetKeyDelay		, -1, -1
SetBatchLines	, -1
SetWinDelay		, -1
SetControlDelay, -1
SendMode		, Input
AutoTrim			, On
FileEncoding		, UTF-8


Hotkey, IfWinActive, ahk_class OptoAppClass			            			;~~ ALBIS Hotkey - bei aktiviertem Fenster
Hotkey, $^c			    		, AlbisKopieren					        			;= Albis: selektierten Text ins Clipboard kopieren
Hotkey, $^x			    		, AlbisAusschneiden			        		   	;= Albis: selektierten Text ausschneiden
Hotkey, $^v			    		, AlbisEinfuegen                                 	;= Albis: Clipboard Inhalt (nur Text) in ein editierbares Albisfeld einfügenHotkey, !Down      			, AlbisKarteiKarteSchliessen        			;= Albis: per Kombination die aktuelle Karteikarte schliessen
Hotkey, !Up               		, AlbisNaechsteKarteiKarte	        			;= Albis: per Kombination eine geöffnete Kartekarte vorwärts
Hotkey, !Right    				, Laborblatt    										;= Albis: Laborblatt des Patienten anzeigen
Hotkey, !Left  					, Karteikarte      									;= Albis: Karteikarte des Patienten anzeigen
Hotkey, IfWinActive

Hotkey, IfWinExist, ahk_class OptoAppClass                              		;~~ mehrmonatiges Kalender-Addon
Hotkey, $F9			    		, MinusEineWoche                                	;= Albis: addiert eine Woche einem Datum hinzu, z.B. im AU Formular
Hotkey, $F10		        	, PlusEineWoche                                	;= Albis: subtrahiert eine Woche von einem Datum, z.B. im AU Formular
Hotkey, $F11		    		, MinusEinMonat                                	;= Albis: addiert 4x7 Tage (4 Wochen) einem Datum hinzu, z.B. im AU Formular
Hotkey, $F12		        	, PlusEinMonat                                    	;= Albis: subtrahiert 4x7 Tage (4 Wochen) von einem Datum, z.B. im AU Formular
Hotkey, IfWinExist

Hotkey, IfWinActive, Dauermedikamente ahk_class #32770    		;~~ Sortieren im Dauermedikamentenfenster
Hotkey, #				    		, DauermedBezeichner		        			;= Albis: zum Kategorisieren der Medikamente (übersichtlichere Darstellung, besser als mit dem BMP)
Hotkey, ^Up                	, DauermedRauf                                   	;= Albis: schiebt die ausgewählte Zeile um eine Zeile nach oben
Hotkey, ^Down              	, DauermedRunter                         		;= Albis: schiebt die ausgewählte Zeile um eine Zeile nach unten
Hotkey, IfWinActive

Hotkey, IfWinActive, Dauerdiagnosen von ahk_class #32770         	;~~ Sortieren im Dauerdiagnosenfenster
Hotkey, ^Up                	, DiagnoseRauf                                    	;= Albis: schiebt eine Diagnose höher
Hotkey, ^Down              	, DiagnoseRunter                                  	;= Albis: zieht eine Diagnose eine Zeile runter
Hotkey, IfWinActive

Hotkey, IfWinActive, Cave! von ahk_class #32770                         	;~~ Sortieren im Cave! von Dialog
Hotkey, ^Up                	, CaveVonRauf                                    	;= Albis: schiebt eine Diagnose höher
Hotkey, ^Down              	, CaveVonRunter                                  	;= Albis: zieht eine Diagnose eine Zeile runter
Hotkey, IfWinActive

return

AlbisHeuteDatum:         	;{ 	das Albis Programmdatum auf heute setzen
	AlbisSetzeProgrammDatum(A_DD "." A_MM "." A_YYYY)
Return
;}
AlbisKopieren:               	;{ 	Strg	+	c                                                                	(wie sonst überall in Windows nur nicht in Albis)

	AlbisCText:= AlbisCopyCut()
	If StrLen(AlbisCText) > 0
	{
			clipboard := AlbisCText
			GdiSplash(SubStr(AlbisCText, 1, 30) . "`.`.`.", 1)
	}
	else
			Tooltip("!Nichts kopiert!", 1)
return
;}
AlbisAusschneiden:        	;{ 	Strg	+	x                                                                	(wie sonst überall in Windows nur nicht in Albis)

	AlbisCText:= AlbisCopyCut()
	If StrLen(AlbisCText) > 0
	{
			clipboard := AlbisCText
			GdiSplash(SubStr(AlbisCText, 1, 30) . "`.`.`.", 1)
			SendInput, {DEL}
	}
	else
			Tooltip("!Nichts ausgeschnitten!", 1)

return
;}
AlbisEinfuegen:             	;{ 	Strg	+	v                                                                	(wie sonst überall in Windows nur nicht in Albis)
	If (AlbisCText <> clipboard) {
        FileAppend, %clipboard% `n, logs'n'data\CopyAndPaste.log
	}
	SendRaw, % Clipboard
return
;}
AlbisMarkieren:              	;{ 	Strg	+ a
	AlbisSelectAll()
return ;}
AlbisKarteikarteSchliessen:	;{ 	Shift	+	Pfeil nach unten
	IfWinNotActive, ahk_class OptoAppClass
			WinActivate, ahk_class OptoAppClass
	Send, {LControl Down}{F4}{LControl Up}
return
;}
AlbisNaechsteKarteikarte:	;{ 	Shift	+	Pfeil nach rechts
	SendMessage, 0x224,,,, % "ahk_id " AlbisGetMDIClientHandle()	; WM_MDINEXT
return
;}
DauermedRauf:             	;{ 	Strg	+ Pfeil nach oben
	PostMessage, 0x00F5,,, Button1,  Dauermedikamente ahk_class #32770
	SetTimer, DauermedFokus, -50
Return
;}
DauermedRunter:          	;{ 	Strg	+ Pfeil nach unten
	PostMessage, 0x00F5,,, Button2,  Dauermedikamente ahk_class #32770
	SetTimer, DauermedFokus, -50
return
;}
DiagnoseRauf:              	;{ 	Strg	+ Pfeil nach oben
	PostMessage, 0x00F5,,, Button1,  Dauerdiagnosen von ahk_class #32770
	SetTimer, DauerDiagnoseFokus, -50
Return ;}
DiagnoseRunter:            	;{ 	Strg	+ Pfeil nach unten
	PostMessage, 0x00F5,,, Button2,  Dauerdiagnosen von ahk_class #32770
	SetTimer, DauerDiagnoseFokus, -50
return ;}
CaveVonRauf:               	;{ 	Strg	+ Pfeil nach oben  stark Verbesserungswürdig
	PostMessage, 0x00F5,,, Button1,  Cave! von ahk_class #32770
	SetTimer, DauerDiagnoseFokus, -50
Return ;}
CaveVonRunter:            	;{ 	Strg	+ Pfeil nach unten
	PostMessage, 0x00F5,,, Button2,  Cave! von ahk_class #32770
	SetTimer, DauerDiagnoseFokus, -50
return ;}
TTipOff:                        	;{
	ToolTip,,,, % TT
return
;}
DauermedFokus:           	;{ 	Eingabefokus auf das Listview im Dauermedikamentenfenster
	If WinExist("Dauermedikamente ahk_class #32770")
			ControlFocus, SysListView321, Dauermedikamente ahk_class #32770
return
;}
DauerDiagnoseFokus:   	;{ 	Eingabefokus auf das Listview im Dauerdiagnosenfenster
	If WinExist("Dauerdiagnosen von ahk_class #32770")
			ControlFocus, SysListView321, Dauerdiagnosen von ahk_class #32770
return
;}
CaveVonFokus:             	;{ 	Eingabefokus auf das Listview im Cave! von fenster
	If WinExist("Dauerdiagnosen von ahk_class #32770")
			ControlFocus, SysListView321, Cave! von ahk_class #32770
return
;}
PlusEineWoche:             	;{ 	F10                                                                                     - in allen Datumsfeldern in Albis
	If IsObject(Feld:= IsDatumsfeld())
			AddToDate(Feld, 7, "days")
	else
			SendInput, {F10}
Return
;}
MinusEineWoche:          	;{ 	F9                                                                                       - in allen Datumsfeldern in Albis
	If IsObject(Feld:= IsDatumsfeld())
			AddToDate(Feld, -7, "days")
	else
			SendInput, {F9}
Return
;}
PlusEinMonat:                	;{ 	F12                                                                                     - in allen Datumsfeldern in Albis
	If IsObject(Feld:= IsDatumsfeld())
			AddToDate(Feld, 4*7, "days")
	else
			SendInput, {F12}
Return
;}
MinusEinMonat:            	;{ 	F11                                                                                     - in allen Datumsfeldern in Albis
	If IsObject(Feld:= IsDatumsfeld())
			AddToDate(Feld, -4*7, "days")
	else
			SendInput, {F11}
Return
;}
Kalender:                      	;{ 	Shift	+ F3
	If IsObject(feld:= IsDatumsfeld())
	{
				If !InStr(GetProcessNameFromID(feld.hwin), "albis")
				{
						SendInput, {Shift Down}{F3}{Shift Up}
						return
				}

			;Kalender Gui erstellen
				Calendar(feld)
	}
	else
	{
				SendInput, {Shift Down}{F3}{Shift Up}
	}
return
;}
DruckeLabor:                	;{ 	Alt 	+ F8
	AlbisDruckeLaborblatt(1, "Standard", "")
return
;}
Laborblatt:                    	;{ 	Alt 	+ Rechts
	If !InStr(AlbisGetActiveWindowType(), "Patientenakte") or InStr(AlbisGetActiveWindowType(), "Laborblatt")
			return
	AlbisZeigeLaborBlatt()
return
;}
Karteikarte:                   	;{ 	Alt 	+ Links
	If !InStr(AlbisGetActiveWindowType(), "Patientenakte") or InStr(AlbisGetActiveWindowType(), "Karteikarte")
			return
	AlbisZeigeKarteikarte()
return
;}


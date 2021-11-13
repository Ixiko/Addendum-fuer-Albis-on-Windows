;###############################################################
;------------------------------------------------------------------------------------------------------------------------------------
;----------------------------------------------	Addendum für Albis on Windows ---------------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------
;------------------------------------------------------ Modul: Schulzettel -------------------------------------------------------
															         Version:= "V0.91"
;------------------------------------------------------------------------------------------------------------------------------------
;{      			Ein Skript zur einfachen Erstellung einer Schulbefreiung und/oder Sportbefreiung
;
;	Keine Schulbefreiung  mit der Hand  schreiben auch  wenn es  schneller geht.  Dieses Skript versucht eine
;	annährend so schnell so sein. Es bietet die Möglichkeit einen ansprechenden Schulzettel zu haben, indem
;	man ein vorbereitet Word-Dokument das sicherlich schon in Albis hinterlegt ist genutzt wird.
;
;	Warum also nicht einfach nur Word nutzen, wozu noch ein Skript?
;	Word bietet kein Kalenderfeld welches man einfach in ein Worddokument einbinden könnte! So tippt man
;	meistens doch die Daten von Hand ein und muss erst nachschauen welches Datum z.B. der Freitag hat.
;
;	Dieses Skript kann aber noch mehr. Es verändert die Eintragungen ganz smart je nachdem ob man nur
;	einen Tag auswählt oder einen Datumsbereich.
;	Beispiele:
;
;		1)	ausgewählt wurde nur ein Tag dann würde folgendes auf dem Schulzettel erscheinen:
;			Mario ist am Mo, dem 18.02.2019 schulunfähig erkrankt.
;
;		2) ein Datumsbereich 18.02.2019 bis 22.02.2019
;			Mario ist vom Mo, dem 18.02. bis zum Fr, dem 22.02.2019 schulunfähig erkrankt.
;			(wie sie sehen wird der Übersichtlichkeit wegen beim ersten Datum das Jahr entfernt,
;			 aber nur wenn zwischen beiden Tagen kein Jahreswechsel statt findet)
;
;	Ok, mit der Hand geht's schneller. Was habe ich sonst noch vom Einsatz des Skriptes ?
;	Vielleicht überzeugt es Sie das das Skript in die Patientakte die Zeiten der Schul- und/oder Sportbefreiung
;	einträgt. Es berechnet zusätzlich die Anzahl der Fehltage (ohne Wochenende). Ganz nett zu wissen wieviele
;	Tage der Mario in diesem Jahr wieder der Schule ferngeblieben ist!
;
;}
;------------------------------------------------------------------------------------------------------------------------------------
;------------------------------------- written by Ixiko -this version is from 19.10.2021 ------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------
;------------------------------- please report errors and suggestions to me: Ixiko@mailbox.org ------------------------
;---------------------------- use subject: "Addendum" so that you don't end up in the spam folder ---------------------
;------------------------------------------------------------------------------------------------------------------------------------
;--------------------------------- GNU Lizenz - can be found in main directory  - 2017 ----------------------------------
;------------------------------------------------------------------------------------------------------------------------------------
;###############################################################


;------------------------------------------- automatischer Ausführungsbereiches -------------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------;{
	;--------------------------------------------------------------------------------------------------------------------
	;	Skript Einstellung
	;-------------------------------------------------------------------------------------------------------------------- ;{
		#Persistent
		#NoEnv
		#SingleInstance, Force
		;#NoTrayIcon
		SetBatchLines        	, -1
		SetWinDelay         	, -1
		SetControlDelay    	, -1
		CoordMode, ToolTip	, Screen
		FileEncoding         	, UTF-8

	; Tray Icon erstellen
		If (hIcon := Schulzettel_ico())
			Menu, Tray, Icon, % "hIcon: " hIcon
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	globale Variablen festlegen
	;-------------------------------------------------------------------------------------------------------------------- ;{
		;- die zwei wichtigsten Variablen erstellen -
		global AddendumDir
		global AlbisWinID  	:= WinExist("ahk_class OptoAppClass")
		global CompName	:= StrReplace(A_ComputerName, "-")
		RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	Schulzetteltexte
	;-------------------------------------------------------------------------------------------------------------------- ;{
		Schulbefreiung    	:= "<Vorname> ist <Zeitraum> schulunfähig erkrankt."
		Sportbefreiung    	:= "Anschließend ist <KontextAnrede> <Zeitraum> vom Sportunterricht zu befreien."
		am				    	:= "am M1"
		vombis		    		:= "vom M1 bis zum M2"
		timeFormat	    	:= "ddd.', dem' dd.MM.yyyy"
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	andere Variablen, Drucker ermitteln
	;-------------------------------------------------------------------------------------------------------------------- ;{
	  ;- Vorname und Nachname des Patienten ermitteln -
		namen := StrSplit(AlbisCurrentPatient(), ",")
	  ;- den Drucker ermitteln den der Skript ausführende Clientrechner normalerweise nutzt - evtl. Nutzer fragen -
		IniRead, Schulzettel_DefaultPrinter, % AddendumDir "\Addendum.ini", % CompName, Schulzettel_DefaultPrinter
		If (DefaultPrinter = "ERROR")
				DefaultPrinter := DefaultPrinter("Schulzettel")
		IniRead, SchulzettelDoc, % AddendumDir "\Addendum.ini", % "Addendum", SchulzettelDokument
		;- Ersteinstellungen vervollständigen
		If InStr(SchulzettelDoc, "ERROR")		{
			stammverzeichnis:
			IniRead, AlbisWorkDir, % AddendumDir "\Addendum.ini", Albis, AlbisWorkDir
			If InStr(AlbisWorkDir, "ERROR")				{
				FileSelectFolder, AlbisWorkDir,, 0, % "Bitte wählen Sie zunächst das Albis Stammverzeichnis (meist \\Server\albiswin)"
				If ErrorLevel						{
					MsgBox, 4, Addendum für Albis on Windows, % "Sie haben nichts ausgewählt.`nSoll das Skript abgebrochen werden?"
					IfMsgBox, Yes
						ExitApp
				}

				AlbisWorkDir := RegExReplace(AlbisWorkDir, "\\$")
				If !InStr(FileExist(AlbisWorkDir "\TVL") , "D")						{
					PraxTT("Dies ist nicht das Albisstammverzeichnis.", "6 3")
					goto stammverzeichnis
				}

				IniWrite, % AlbisWorkDir, % AddendumDir "\Addendum.ini", Albis, AlbisWorkDir
			}

			schulzettelDoc:
			FileSelectFile, schulzettelDoc, 1, % AlbisWorkDir "\TVL", % "Bitte wählen Sie hier Ihr Word-Dokument für den Schulzettel aus!", *.doc;*.docx
			If ErrorLevel				{
				MsgBox, 4, Addendum für Albis on Windows, % "Sie haben nichts ausgewählt.`nSoll das Skript abgebrochen werden?"
				IfMsgBox, Yes
						ExitApp
			}

			msgText =
			(LTrim
				Beachten Sie bitte das in Ihrer Schulzettel Vorlage
				die richtigen Seiteneinstellungen abgespeichert sind.
				Das Skript kann die Größe des Ausdruckes (z.B. A5)
				nicht wissen!
			)

			MsgBox, 4, Addendum für Albis on Windows, % "Gewähltes Dokument:`n" SchulzettelDoc "`nMit 'Ja' wird dieses Dokument in die Einstellungen übernommen.`n`n" msgText
			IfMsgBox, No
					ExitApp

			IniWrite, % SchulzettelDoc, % AddendumDir "\Addendum.ini", % "Addendum", SchulzettelDokument

		}

		;GetInstalledPrinters("|",True)
	;}

	;- Schulbefreiung Gui erstellen
		gosub Calendar

return
;}


;------------------------------------------------------------------------------------------------------------------------------------
;---------------------------------------------------- Kalender Gui ---------------------------------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------;{
Calendar:	         	;{ Gui für die schnelle Eingabe zum Erstellen einer Schulbefreiung

	init := DllCall("LoadLibrary", "Str", "Msftedit.dll", "Uint"), MODULEID := 091009

	;--------------------------------------------------------------------------------------------------------------------
	;	Berechnungen für Datumsfokus
	;-------------------------------------------------------------------------------------------------------------------- ;{
		FormatTime, datum, % A_Now, % timeFormat
		day := SubStr(datum, 1, 2), month := SubStr(datum, 3, 2), year := SubStr(datum, 5, 4)
		if (month-1 < 1)
			month := SubStr( "00" (12+month-1), -1), year -= 1
		else
			month := SubStr("00" (month-1), -1)
		FormatTime, begindate, % year month day, yyyyMMdd
	;}
	;--------------------------------------------------------------------------------------------------------------------
	;	Erste Textersetzungen
	;-------------------------------------------------------------------------------------------------------------------- ;{
		ScBf := StrReplace(schulbefreiung, "<Vorname>", Trim(namen.2))
		ScBf := StrReplace(ScBf, "<Zeitraum>", StrReplace(am, "M1", datum))
		If InStr(daten[3], "w")
			SpBf := StrReplace(Sportbefreiung, "<KontextAnrede>", "sie")
		else
			SpBf := StrReplace(Sportbefreiung, "<KontextAnrede>", "er")
		SpBf := StrReplace(SpBf, "<Zeitraum>", StrReplace(am, "M1", datum))
	;}
	;--------------------------------------------------------------------------------------------------------------------
	;	Kalender Gui
	;-------------------------------------------------------------------------------------------------------------------- ;{
		nCalBg:= "AAFFAA"
		Gui, newCal: Destroy
		Gui, newCal: +Owner +AlwaysOnTop HWNDhCal
		Gui, newCal: Color, % "c" nCalBg

	;-:Anlegen des Kalenders 							;{
		Gui, newCal: Add, Progress	, % "vProgr1		HWNDhProgr1									x0 y0 w600 h60 cCCFFCC", 100
		WinSet, ExStyle, -0x00020000, ahk_id %hProgr1%
		Gui, newCal: Font, s14
		Gui, newCal: Add, Text			, % "vHinweis1 															x5 y5 Backgroundtrans Center", % "Wähle den Datumsbereich für Schul- und/oder Sportunfähigkeit"
		Gui, newCal: Font, s10
		Gui, newCal: Add, Text			, % "vHinweis2 															Backgroundtrans Center"	, % "halte Shift um Anfangs-und Enddatum zu wählen"
		Gui, newCal: Font, s12
		Gui, newCal: Add, Checkbox	, % "vChb1 		HWNDhChb1 		gEnabler 				checked"							, % "Befreiung vom Unterricht"
		WinSet, ExStyle, 0x00000020, ahk_id %hChb1%
	;-: Kalender Schulunfähig
		Gui, newCal: Add, MonthCal	, % "vmyCal1 	HWNDhmyCal1 	gDatePicker			Multi R1 W-3 4"    				, % begindate
		WinSet, ExStyle, 0x00000020, ahk_id %hmyCal1%
		SendMessage 0x100A,, % Format("{:u}", "0x" . nCalBg),, ahk_id %hmyCal1% ; MCM_SETCOLOR

		Gui, newCal: Add, Edit			, % "vZsz1 		HWNDhZsz1 									w100 h60"						, % ScBf
		Gui, newCal: Add, Checkbox	, % "vChb2 		HWNDhChb2 		gEnabler"														, % "Sportbefreiung etc."
		WinSet, ExStyle, 0x00000020, ahk_id %hChb2%
	;-: Kalender Sportbefreiung
		Gui, newCal: Add, MonthCal	, % "vmyCal2 	HWNDhmyCal2	gDatePicker			Multi R1 W-3 4 	Disabled"	, % begindate
		WinSet, ExStyle, 0x00000020, ahk_id %hmyCal2%
		SendMessage 0x100A,, % Format("{:u}", "0x" . nCalBg),, ahk_id %hmyCal2% ; MCM_SETCOLOR

		Gui, newCal: Add, Edit			, % "vZsz2 		HWNDhZsz2 									w100 h60 		Disabled"	, % SpBf
		Gui, newCal: Add, Button		, % "vCok	     								gCalendarOK    	default"								, % "Schulzettel drucken"
		Gui, newCal: Add, Button		, % "vCan 										gCalendarCancel	xp+200 yp"						, % "Abbrechen"

		;}
	;-:verstecktes Anzeigen des Kalenders			;{
		Gui, newCal: Show, AutoSize Hide , Addendum für Albis on Windows - erweiterter Kalender
		;}
	;-:Größen ermitteln 									;{
		oWinA:= GetWindowInfo(hCal)
		GuiControlGet	, MyCal1, newCal: Pos			, myCal1
		GuiControlGet	, Hw	 , newCal: Pos				, Hinweis1
		GuiControlGet	, Cok , newCal: Pos
		GuiControlGet	, Can , newCal: Pos
		SendMessage 0x101F, 0, 0,, % "ahk_id " hmyCal1 ; MCM_GETCALENDARBORDER
		CalBorder:= ErrorLevel
		;}
	;-:Position von Gui-Elementen anpassen 	;{
		GuiControl				, newCal: MoveDraw		, Hinweis1		, % "x" MyCal1X 										" w" MyCal1W
		GuiControl				, newCal: MoveDraw		, Hinweis2		, % "x" MyCal1X " y" (HwY+HwH+0) 			" w" MyCal1W
		GY:= HwY+HwH+0, GWidth:= MyCal1W
		;Schulbefreiung
		GuiControlGet	, Hw	, newCal: Pos				, Hinweis2
		GX:= LeftX:= MyCal1X + 2*CalBorder +4
		GuiControl				, newCal: MoveDraw		, Chb1			, % "x" LeftX 		" y" (HwY + HwH + 20)
		GuiControlGet	, Hw	, newCal: Pos				, Chb1
		GuiControl				, newCal: MoveDraw		, myCal1		, % 					" y" (HwY + HwH +   5)
		GuiControlGet	, Hw	, newCal: Pos				, myCal1
		GuiControl				, newCal: MoveDraw		, Zsz1			, % "x" LeftX		" y" (HwY + HwH +   5) 	" w" MyCal1W - 5*CalBorder - 2
		GuiControlGet	, Hw	, newCal: Pos				, Zsz1
		;Sportbefreiung
		GuiControl				, newCal: MoveDraw		, Chb2			, % "x" LeftX 		" y" (HwY + HwH + 10)
		GuiControlGet	, Hw	, newCal: Pos				, Chb2
		GuiControl				, newCal: MoveDraw		, myCal2		, % 					" y" (HwY + HwH +   5)
		GuiControlGet	, Hw	, newCal: Pos				, myCal2
		GuiControl				, newCal: MoveDraw		, Zsz2			, % "x" LeftX 		" y" (HwY + HwH +   5) 	" w" MyCal1W - 5*CalBorder - 2
		GuiControlGet	, Hw	, newCal: Pos				, Zsz2
		GH:= HwY + HwH - GY
		;Ok und Abbruch
		GuiControl				, newCal: MoveDraw		, Cok			, % "x" LeftX 		" y" (HwY + HwH + 15)
		GuiControl				, newCal: MoveDraw		, Can			, % "x" LeftX + HwW - CanW " y" (HwY+HwH+15)
		GuiControlGet	, Hw	, newCal: Pos				, Cok
	;}
	;-:Fenster positionieren und anzeigen 		;{
		oWinA:= GetWindowInfo(hCal)
		GuiControlGet	, Tw	, newCal: Pos				, Hinweis2
		GuiControl				, newCal: MoveDraw		, Progr1		, % "x0 y0 w" oWinA.WindowW " h" TwY + TwH + 10
		;GuiControlGet	, Hw	, newCal: Pos            	, MyCal2

		MouseGetPos, mx, my
		SysGet, mon, MonitorWorkArea
		If (my + oWinA.WindowH > monBottom)
				Gui, newCal: Show, % "x" (mx  - 10)  " y" (monBottom 		- oWinA.WindowH 	 - 50 )	" h" HwY + HwH + 10, % "Addendum für Albis on Windows - Schulzettel"
		else
				Gui, newCal: Show, % "x" (mx  - 10)  " y" (monBottom//2 - oWinA.WindowH//2		)	" h" HwY + HwH + 10, % "Addendum für Albis on Windows - Schulzettel"
	;}
	;}

Return
;}

Enabler:				;{ schaltet die zwei Kalender frei, je nachdem was gebraucht wird

		ControlGet, isChecked, Checked,,, % "ahk_id " h%A_GuiControl%
		RegExMatch(A_GuiControl, "\d", nr)
 		If isChecked		{
			Control, Enable,,, % "ahk_id " hmyCal%nr%
			Control, Enable,,, % "ahk_id " hZsZ%nr%
		}
		else		{
			Control, Disable,,, % "ahk_id " hmyCal%nr%
			Control, Disable,,, % "ahk_id " hZsZ%nr%
		}

return
;}

DatePicker:			;{ zeigt bei jedem Auswählen eines Datums oder Datumbereiches zur Kontrolle sofort den Text an der auf der Schulbefreiung eingetragen wird

		Gui, newCal: Submit, NoHide
		;- wenn sich die Tage unterscheiden dann wird equal später auf 0 gesetzt -
		equal := true
		;- ermittelt die Nummer des auslösenden Controls -
		RegExMatch(A_GuiControl, "\d", nr)
		FormatTime, CalM1, % StrSplit(%A_GuiControl%,"-").1, % timeFormat
		FormatTime, CalM2, % StrSplit(%A_GuiControl%,"-").2, % timeFormat
		;- liest hier den zugehörigen Text aus dem Editcontrol mit der zuvor ermittelten Nummer ein -
		GuiControlGet, inhalt1,, Zsz%nr% ;%nr%

	;--------------------------------------------------------------------------------------------------------------------
	;	liegt der Befreiungszeitraum innerhalb desselben Jahres, entfernt er beim vom Datum das Jahr
	;	sieht optisch besser aus
	;--------------------------------------------------------------------------------------------------------------------;{
		If ( CalM1 != CalM2 )		{
			equal:= false
			RegExMatch(CalM1,"\d+$", year1)
			RegExMatch(CalM2,"\d+$", year2)
			If ( year1 = year2 )
				CalM1:= RTrim(StrReplace(CalM1, year1), A_Space)
		}
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	sucht nach den Key-Worten: am, vom, bis und ersetzt diese Bereiche durch <Zeitraum>
	;	danach wird dort wieder entsprechend dem Ausgewählten der Text smart ersetzt
	;--------------------------------------------------------------------------------------------------------------------;{
		If RegExMatch(inhalt1, "i)ist(\s\w+\s|\s+)am")
			inhalt:= RegExReplace(inhalt1, "i)am\s\w+\.\,\sdem\s[\d\.]+", "<Zeitraum>")
		else if RegExMatch(inhalt1, "i)ist(\s\w+\s|\s+)vom")
			inhalt:= RegExReplace(inhalt1, "i)vom\s\w+\.\,\sdem\s[\d\.]+\sbis\szum\s\w+\.\,\sdem\s[\d\.]+", "<Zeitraum>")
		;ToolTip, % A_GuiControl "`n" CalM1 "`n" CalM2 "`n" inhalt, % oWinA.WindowX + oWinA.WindowW + 5, % oWinA.WindowY
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	<Zeitraum> wird mit den neuen Tagen ersetzt, nutzt als Template die zuvor definierten
	;	Variablen am und vombis
	;--------------------------------------------------------------------------------------------------------------------;{
		If ( equal = true )
			GuiControl,, % zsz%nr%, % StrReplace(inhalt, "<Zeitraum>", StrReplace(am, "M1", CalM1))
		else
			GuiControl,, % zsz%nr%, % StrReplace(inhalt, "<Zeitraum>", StrReplace(StrReplace(vombis, "M1", CalM1), "M2", CalM2))
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	enthält die Sportbefreiung eines der Stichwörter: "Anschließend",  wird automatisch das
	;	Anfangsdatum auf den folgenden Werktag gesetzt im Kalender für die Sportbefreiung gesetzt
	;--------------------------------------------------------------------------------------------------------------------;{

	;}

return
;}

CalendarOK:		;{ Schulzettel drucken wurde gewählt - es werden alle Daten zusammengestellt und dann das Dokument ausgedruckt

		Gui, newCal: Submit, NoHide+
	;--------------------------------------------------------------------------------------------------------------------
	;	liest beide Kalender aus und formatiert das Datum ins deutsche Format
	;--------------------------------------------------------------------------------------------------------------------;{
		FormatTime, CalM1, % StrSplit(MyCal1,"-").1, % timeFormat
		FormatTime, CalW1, % StrSplit(MyCal1,"-").1, dd.MM.yyyy
		FormatTime, CalM2, % StrSplit(MyCal1,"-").2, % timeFormat
		FormatTime, CalW2, % StrSplit(MyCal1,"-").2, dd.MM.yyyy
		FormatTime, CalM3, % StrSplit(MyCal2,"-").1, % timeFormat
		FormatTime, CalW3, % StrSplit(MyCal2,"-").1, dd.MM.yyyy
		FormatTime, CalM4, % StrSplit(MyCal2,"-").2, % timeFormat
		FormatTime, CalW4, % StrSplit(MyCal2,"-").2, dd.MM.yyyy
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	liest den Inhalt der Edit-Felder, wenn die zugehörige Checkbox gesetzt wurde, erstellt
	;	Eintragungen für die Akte je nach gewählten Daten
	;--------------------------------------------------------------------------------------------------------------------;{
		Entschuldigung:=[], Akteneintrag:="", checkedWas:= []
		Loop, 2		{
			ControlGet, isChecked, Checked,,, % "ahk_id " hChb%A_Index%
			If !isChecked{
				checkedWas[A_Index] 		:= 0
				Entschuldigung[A_Index]	:= ""
			}
			else		{

				checkedWas[A_Index] 		:= 1
				GuiControlGet, tmptext,, Zsz%A_Index%
				Entschuldigung[A_Index]	:= tmptext

				If (A_Index=1) {
					If (CalM1=CalM2){
						Akteneintrag .= "Schulbefreiung (1 Tag) am " CalM1
					}
					else	{
						Tage:= DateDiff("DD", CalW2, CalW1) + 1
						RegExMatch(CalM1,"\d+$", year1)
						RegExMatch(CalM2,"\d+$", year2)
						If ( year1 = year2 )
								CalM1:= RTrim(StrReplace(CalM1, year1), A_Space)
						Akteneintrag .= "Schulbefreiung für " Tage " Tage vom " CalM1 " bis " CalM2
					}
				}
				else if (A_Index=2)				{
					If (CalM3=CalM4)							{
						If (checkedWas[1] = 1)
							Akteneintrag .= ", Sportbefreiung nur am " CalM3
						else
							Akteneintrag .=   "Sportbefreiung nur am " CalM3
					}
					else							{
						Tage:= DateDiff("DD", CalW4, CalW3) + 1
						If (checkedWas[1] = 1)
							Akteneintrag .= ", Sportbefreiung für " Tage " Tage vom " CalM3 " bis " CalM4
						else
							Akteneintrag .=   "Sportbefreiung für " Tage " Tage vom " CalM3 " bis " CalM4
					}
				}
			}
		}
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	prüft ob Albis läuft und wenn das es auch antwortet, schliessen blockierender Fenster
	;--------------------------------------------------------------------------------------------------------------------;{
		Gui, newCal: Show, Hide                                                                                                                                                        	;erst jetzt wird die Gui versteckt
		AlbisStatus()
		AlbisIsBlocked(AlbisWinID)                                                                                                                                                     	;autoclose blockierender Fenster ist default
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	öffnet das Schulzetteldokument in Word, setzt Eintragungen für die Akte
	;--------------------------------------------------------------------------------------------------------------------;{
		;- Menuaufruf Extern/Arztbrief -
			hVorlagen := Albismenu(32829, "Vorlagen", 5, 1)
			Sleep, 2000
			VerifiedSetText("Edit1", SchulzettelDoc, hVorlagen, 200)
			VerifiedSetText("Edit6", Akteneintrag , hVorlagen, 200)
			VerifiedClick("Button18", hVorlagen,,,1)
		;- wartet auf das sich öffnende Wordfenster -
			WinWaitActive, ahk_class OpusApp, 10
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	Verbindung zu Word herstellen und falls mehrere Word-Dokumente geöffnet sind,
	;	muss das richtige Dokument gefunden werden
	;--------------------------------------------------------------------------------------------------------------------;{
			oWord 	:= ComObjActive("Word.Application")
			docNr	:= 0

		;- einen Augenblick warten damit Albis Zeit hat seine Platzhalter auszufüllen
			Sleep, 2000

		;- das richtige Word Dokument heraussuchen und nach vorne holen
			Loop, % oWord.Documents.Count			{
					Content:= oWord.Documents.Item(A_Index).Content.Text
				;- Übereinstimmung anhand der speziellen Schlüsselwörter und des Patientennamen wird überprüft
					If ( InStr(Content, "<Entschuldigung1>") and InStr(Content, namen[1]) and InStr(Content, namen[2]) )					{
						;- gefunden, dann wird genau dieses Dokument zum aktiven Dokument gemacht
							oWord.Documents(A_Index).Activate
							docNr:= A_Index
							WinActivate, % oWord.Documents(A_Index).Name
							break
					}
			}

		;- falls kein passendes Dokument geöffnet wird, erfolgt ein Hinweis und das Skript-Gui wird wieder angezeigt
			If !docNr			{
					MsgBox, 4, Addendum für Albis on Windows,
													(LTrim
													Scheinbar konnte das Dokument `"%SchulzettelDoc%`"
													nicht aufgerufen werden.
													Möchten Sie es noch einmal probieren?
													Abbrechen - beendet das Skript, die gemachten Einstellungen
													für diesen Patienten werden aber gespeichert.
													)
					IfMsgBox, Yes
							return
					IfMsgBox, Cancel
					{
							toWrite:= % namen[1] "," namen[2] "`n" MyCal1 "`n" MyCal2 "`n" Entschuldigung[1] "`n" Entschuldigung[2]
							FileAppend, % toWrite, % A_ScriptDir "\letzte_SchulbefreiungSicherung.txt"
							ExitApp
					}

			}
		;}

	;--------------------------------------------------------------------------------------------------------------------
	;	im Dokument werden die Schlüsselwörter durch die erstellten Texte ersetzt und das
	;	Dokument wird gedruckt, gespeichert und geschlossen, das Skript wird anschliessend beendet
	;--------------------------------------------------------------------------------------------------------------------;{
		MSWord_FindAndReplace(oWord, "<Entschuldigung1>", Entschuldigung[1])
		MSWord_FindAndReplace(oWord, "<Entschuldigung2>", Entschuldigung[2])

		;oWord.ActivePrinter := "Microsoft Print to PDF" ;Testeinstellung
		oWord.ActivePrinter := DefaultPrinter       	; Einstellen des Druckers
		oWord.DisplayAlerts := 0                        	; Hinweisfenster ausschalten, verhindert z.B. das "die Ränder sind zu schmal" den Ablauf unterbricht
		oWord.ActiveDocument.PrintOut             	; Druck starten
		Sleep, 2000
		oWord.ActiveDocument.Save()                  	; Dokument speichen
		oWord.ActiveDocument.close                 	; Dokument schliessen
		oWord.Quit
	;}
;}

CalendarCancel:				;{
newCalGuiClose:
newCalGuiEscape:
	Sleep 150
	Gui, newCal: Destroy
ExitApp


;}

;}


;------------------------------------------------------------------------------------------------------------------------------------
;------------------------------- Funktionen welche nur von diesem Skript benötigt werden -------------------------------
;------------------------------------------------------------------------------------------------------------------------------------;{
Edit_GetMargins(hEdit,ByRef r_LeftMargin="",ByRef r_RightMargin="") {
    Static EM_GETMARGINS:=0xD4
    SendMessage EM_GETMARGINS,0,0,,ahk_id %hEdit%
    r_LeftMargin :=ErrorLevel & 0xFFFF      ;-- LOWORD of result
    r_RightMargin:=ErrorLevel>>16           ;-- HIWORD of result
    Return r_LeftMargin  ; ##### --------------------------------------- Really?
}

MSWord_FindAndReplace(obj, search, replace) {
	obj.Selection.Find.ClearFormatting
	obj.Selection.Find.Replacement.ClearFormatting
	obj.Selection.Find.Execute( search, 0, 0, 0, 0, 0, 1, 1, 0, replace, 2)
}

MSWord_SelectionSetup(obj, setup:="PageSetup", ValueSet:="TopMargin,20,LeftMargin,20,RightMargin,20,BottomMargin,20") {

	;Parameters:
	;
	; setup 			- can be every command MS Word is using after the Selection command
	;

	/*
		oWord.Selection.PageSetup.TopMargin := 20
		oWord.Selection.PageSetup.LeftMargin := 20
		oWord.Selection.PageSetup.RightMargin := 20
		oWord.Selection.PageSetup.BottomMargin := 20

		oWord.Selection.Font.Bold :=   1
		oWord.Selection.Font.Size :=   18
	*/

	Set:=[], Cnt:=1

	Set:= StrSplit(ValueSet, ",")
	Loop
	{
			obj.Selection[(Set[A_Index])] := Set[A_Index+1]
			cnt += 2
	} until (Cnt=Set.MaxIndex())


}

DefaultPrinter(forDocument := "") {

	InstalledPrinter	:= GetInstalledPrinters("|",True)
	printerCount		:= StrLen(InstalledPrinter) - StrLen(StrReplace(InstalledPrinter, "|")) + 1
	Sort, InstalledPrinter, D|

	Gui DP: New, +hWndhPrinter +AlwaysOnTop -DPIScale +Owner -Caption
	Gui DP: Color, 0xD2D8EA
	Gui DP: Add, Progress	, % "vDPProgr1		HWNDhDPProgr1				x0 y0 w500 h60 cRED", 100
	WinSet, ExStyle, -0x00020000, ahk_id %hDPProgr1%
	Gui DP: Font, s30 bold cWhite, Futura Md Bk
	Gui DP: Add, Text, hWndhTxt2 x0 y5 w500 h60 +Center BackgroundTrans, Achtung!
	Gui DP: Font, s20 Norm q5 cBlack, Futura Bk BT
	Gui DP: Add, Text, x0 y70 w500 +Center BackgroundTrans, Es muss noch ein Standard-Drucker`nfür diesen Computer festgelegt werden!
	Gui DP: Font, s16 Norm q5 cBlack, Futura Bk BT
	Gui DP: Add, Text, x0 y150 w500 +Center BackgroundTrans, folgende Drucker sind installiert:
	Gui DP: Add, ListBox, x25 y180 w450 r%printerCount% BackgroundTrans vDefaultPrinter, % InstalledPrinter
	GuiControlGet	, Hw	, DP: Pos				, DefaultPrinter
	;GuiControl				, DP: MoveDraw	, Chb1			, % "x" LeftX 		" y" (HwY + HwH + 20)
	Gui DP: Add, Text, % "x0 y" (HwY + HwH + 5 ) " w500 +Center BackgroundTrans vChoose", wähle den Drucker und drücke dann Ok.
	GuiControlGet	, Hw	, DP: Pos				, Choose
	Gui, DP: Add, Button		, % "x25    y" ( HwY + HwH + 20 ) " vDPCok gDPOK	default"	, % "OK"
	Gui, DP: Add, Button		, %  "x200 y" ( HwY + HwH + 20 ) " vDPCan gDPCancel"		, % "Abbruch"
	GuiControlGet	, Hw	, DP: Pos				, DPCan
	GuiControl				, DP: MoveDraw	, DpCan, % "x" (500 - 25 - HwW)
	Gui DP: Show, w500, Addendum für Albis on Windows - Standard-Client-Drucker festlegen

Return

DPOK:
	Gui, DP: Submit, Hide
	Gui, DP: Destroy
	IniWrite, % DefaultPrinter, % AddendumDir "\Addendum.ini", % CompName, % (forDocument = "" ? "DefaultPrinter" : forDocument "_DefaultPrinter")

return DefaultPrinter

DPCancel:
ExitApp
}

Schulzettel_ico(NewHandle := False) {
Static hBitmap := Schulzettel_ico()
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "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kQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBwrUDc69D////////////////////////////////////////////////////+//6/2qlYnyBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkQRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBOmRO31p/+//7////////////////////////////////////////////////////////////////4+/WSwG5AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkQRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDQ5MD////////////////////////////////////r8+Tt9ef///////////////////////////////////+EuFtAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkQRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB+tVO41qBdoidAkQBAkQDB26z////////////////////////////D3K9hpCxAkQBAkQBmpzPM4rv///////////////////////////+z05pAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkQRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDy+O7///+Ww3NAkQBAkQCrz4/////////////////////W58huqz1AkQBAkQBAkQBAkQBAkQBAkQB0r0bb6s////////////////////+jyoRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkQRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDg7db///9yrkNAkQBAkQCVwnL////////////l8NyBt1dBkQFAkQBAkQBAkQBJlgxIlgtAkQBAkQBAkQBBkgKGuV3o8uD///////////+SwG1AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkQRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDN4rz///9lpjFAkQBAkQCAtlX////r8+SOvmhEkwVAkQBAkQBAkQBjpS/I4Lb+//76/PjF3bFjpS9AkQBAkQBAkQBGlAiXw3Ty+O7///99tFFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQC516L///9ZnyFAkQBAkQBipC2Ww3NHlQlAkQBAkQBAkQBXnh+516L9/v3////////////////9/v2516JWnh5AkQBAkQBAkQBNmBGZxHdKlw1AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCly4f///9Xnh9AkQBAkQBAkQBAkQBAkQBAkQBOmROqzo76/Pj////////////////////////////////5/PemzIhLlw9AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCYw3X///9yrkNAkQBAkQBAkQBAkQBHlQqbxnr0+fD////////////////////////////////////////////////w9uuSwG1EkwVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCZxHf///+EuFtAkQBAkQBDkwSMvWbs9OX////////////////////////////////////////////////////////////////j79p+tVNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCNvWf///90r0ZAkQBzrkTd7NL////////////////////////////////////////////////////////////////////////////////W58huqz1AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCQv2v///+rz4/U5sb////////////////////////////////////////////////////////////////////////////////////////////+//6/2qpaoCNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCdxnz////////////////////////////////////////////////////////////////////////////////////////////////////////////////7/Pmrz49NmRJAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQBAkQBAkQBNmBGmzIj6/Pj////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////y+O6ZxHdGlAhAkQBAkQBAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQBFlAeWw3Py+O7////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////n8d+Dt1lBkQFAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQDB26z///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+516JAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQCGuV3p8uL////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////m8N10r0VAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQBAkQBBkgJ6sk3U5sX////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////Z6cyGul5CkgNAkQBAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBmpzO516L8/fv////////////////////////////////////////////////////////////////////////////////////////////////////////9/v3G3rNrqjlAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBZnyGmzIjz+O/////////////////////////////////////////////////////////////////////////////////////////1+fKt0JFXnh9AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBKlw6TwW/p8uH////////////////////////////////////////////////////////////////////////n8d+ZxHdNmBFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkgJ6sk3U5sb////////////////////////////////////////////////////////a6c2Gul5DkwRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBuqz3C3K79/v3////////////////////////////////9/vzB26xrqjpAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBZoCKt0JH3+vT////////////////2+vOz05ldoidAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBNmBGRwGzF3bHM4ruYxHZMmBBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQFAkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQFGkAxAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBHkA1WjitAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBWjipyimBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBwiVwAAABVjihAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBQjiGFhoUAAACCh35MjxhAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBIkQ9/h3kAAAAAAAAAAACCh35VjilAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBOjxx+h3cAAAAAAAAAAAAAAAAAAAAAAABziWFXjSxGkA5AkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBEkAlSjyJsilSFhoUAAAAAAAAAAADwAAAAAAcAAMAAAAAAAwAAgAAAAAABAACAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAAAAAACAAAAAAAEAAMAAAAAAAwAA8AAAAAAHAAA="
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
;}


;------------------------------------------------------------------------------------------------------------------------------------
;------------------------------------------------------- Includes ------------------------------------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------;{
#include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
#include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
#include %A_ScriptDir%\..\..\include\Addendum_Datum.ahk
#include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
#include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
#include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk
#include %A_ScriptDir%\..\..\include\Addendum_PraxTT.ahk
#include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
#include %A_ScriptDir%\..\..\include\Addendum_Screen.ahk
#include %A_ScriptDir%\..\..\include\Addendum_Window.ahk

#include %A_ScriptDir%\..\..\lib\ACC.ahk
#include %A_ScriptDir%\..\..\lib\Gdip_All.ahk
#include %A_ScriptDir%\..\..\lib\ini.ahk
#include %A_ScriptDir%\..\..\lib\printer.ahk
#Include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#Include %A_ScriptDir%\..\..\lib\Sift.ahk

;}


/* aussortierter Code
		ValueSet	=
							(LTrim Join
							LineNumbering.Active,False
							,Orientation,wdOrientPortrait
							,TopMargin,CentimetersToPoints(1)
							,BottomMargin,CentimetersToPoints(1)
							,LeftMargin,CentimetersToPoints(1.5)
							,RightMargin,CentimetersToPoints(1.3)
							,Gutter,CentimetersToPoints(0)
							,HeaderDistance,CentimetersToPoints(1.27)
							,FooterDistance,CentimetersToPoints(1.27)
							,PageWidth,CentimetersToPoints(14.8)
							,PageHeight,CentimetersToPoints(21)
							,FirstPageTray,wdPrinterLowerBin
							,OtherPagesTray,wdPrinterLowerBin
							,SectionStart,wdSectionNewPage
							,OddAndEvenPagesHeaderFooter,False
							,DifferentFirstPageHeaderFooter,False
							,VerticalAlignment,wdAlignVerticalTop
							,SuppressEndnotes,False
							,MirrorMargins,False
							,TwoPagesOnOne,False
							,BookFoldPrinting,False
							,BookFoldRevPrinting,False
							,BookFoldPrintingSheets,1
							,GutterPos,wdGutterPosLeft
							)
		;MSWord_SelectionSetup(oWord, "PageSetup", ValueSet)
		;_ := ComObjMissing()
*/

/* Original VBS Code

Sub Macro1()


    With Options
        .UpdateFieldsAtPrint = False
        .UpdateLinksAtPrint = False
        .DefaultTray = "Use printer settings"
        .PrintBackground = True
        .PrintProperties = False
        .PrintFieldCodes = False
        .PrintComments = False
        .PrintHiddenText = False
        .PrintDrawingObjects = True
        .PrintDraft = False
        .PrintReverse = False
        .MapPaperSize = False
        .PrintOddPagesInAscendingOrder = False
        .PrintEvenPagesInAscendingOrder = False
    End With
    With ActiveDocument
        .PrintPostScriptOverText = False
        .PrintFormsData = False
    End With
    Application.PrintOut FileName:="", Range:=wdPrintAllDocument, Item:= _
        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
        ManualDuplexPrint:=False, Collate:=True, Background:=True, PrintToFile:= _
        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
        PrintZoomPaperHeight:=0
End Sub


*/


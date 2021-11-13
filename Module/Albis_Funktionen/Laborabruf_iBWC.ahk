; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                          	** elektronischer Laborabruf ** Automatisierung für infoBoxWebClient.exe
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;
;		Beschreibung: 		Automatisierung für 'vianova Webclient' und Labordatenimport
;
; 		Inhalt:					- Skript startet und überwacht den vianova Webclient für den elektronischen LDT-Transfer
;									- erkennt dabei selbstständig ein neues Datenfile und
; 									- führt automatisch den Labordatenimport durch
;									- nach erfolgreichem Import öffnet es das Laborbuch und
;									- stößt den Import der Daten in die Patientenkartei an (Addendum übernimmt den Rest)
;
;       Abhängigkeiten: 	siehe includes
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Laborabruf_iBWC.ahk last change:    	30.09.2021
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  ; -------------------------------------------------------------------------------------------------------------------------------------------------------
  ; Einstellungen
  ; -------------------------------------------------------------------------------------------------------------------------------------------------------;{
	#NoEnv
	#Persistent
	#SingleInstance              	, Force
	#MaxThreads                  	, 100
	#MaxThreadsBuffer       	, On
	#KeyHistory						, Off

	SetBatchLines					, -1
	ListLines    						, On
	FileEncoding                 	, UTF-8
  ;}

  ; -------------------------------------------------------------------------------------------------------------------------------------------------------
  ; Variablen
  ; -------------------------------------------------------------------------------------------------------------------------------------------------------;{
	global amsg
	global adm            	:= Object()
	global qq              	:= Chr(0x20)
	global ThreadControl:= true
	global compname := StrReplace(A_ComputerName, "-")                    	; der Name des Computer auf dem das Skript läuft

	adm.Labor            	:= Object()
	adm.scriptname    	:= RegExReplace(A_ScriptName, "\.ahk$")
	adm.ID                 	:= GetAddendumID()
  ;}

  ; -------------------------------------------------------------------------------------------------------------------------------------------------------
  ; Pfade und Einstellungen laden
  ; -------------------------------------------------------------------------------------------------------------------------------------------------------;{

		LaborAbrufGui()

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Pfad zu Addendum und den Einstellungen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)
		adm.Dir := AddendumDir
		If !FileExist(adm.ini	:= AddendumDir "\Addendum.ini") {
			MsgBox, 1024, % adm.scriptname, % (amsg .= datestamp(2) "|`t - Die Einstellungsdatei ist nicht vorhanden!`n`t[" adm.ini "]")
			ExitApp
		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Ini initialisieren
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If (StrLen(workini:= IniReadExt(adm.ini)) = 0) {
			MsgBox, 1024, % adm.scriptname, % "Es gab ein unerwartetes Problem beim Zugriff auf die ini Datei!"
			ExitApp
		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Albis Installationspfad, Datenbankpfad
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
		adm.AlbisPath   	:= AlbisPath
		adm.AlbisDBPath 	:= AlbisPath "\db"

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Backup-Ort für empfangene Labordateien auslesen. Bei Bedarf werden Verzeichnisse angelegt
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		adm.DBPath := IniReadExt("Addendum", "AddendumDBPath")           	; Datenbankverzeichnis
		If InStr(adm.DBPath, "ERROR") {
			MsgBox, 1024, % adm.scriptname, % "In der Addendum.ini ist kein Pfad für`n"
																. 	"den Speicherort der Addendumdaten hinterlegt!`n"
																.	"[Addendum/AddendumDBPath=.....]"
			ExitApp
		}
		If !FilePathCreate(adm.DBPath "\Labordaten") {
			MsgBox, 1024, % adm.scriptname, % "Der Pfad für Labordateien konnte nicht angelegt werden.`n`t[" adm.DBPath "\Labor" "]`n"
			ExitApp
		}
		If !FilePathCreate(adm.DBPath "\Labordaten\LDT") {
			MsgBox, 1024, % 	adm.scriptname, % 	"Der Backup-Pfad für die Labordaten-Dateien konnte nicht angelegt werden.`n"
							                                	.    	"`t[" adm.DBPath "\Labor" "]`n"
			ExitApp
		}

		; neues Tagesdatum in die Logdatei schreiben
		FileAppend	, % "---------------------------------------- " A_DD "." A_MM "." A_YYYY "," A_Hour ":" A_Min " ----------------------------------------`n"
							, % adm.LogFilePath := adm.DBPath "\Labordaten\LaborabrufLog.txt"

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Skripteinstellungen und Laborabrufeinstellungen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

		; externes Programm das für den Abruf ausgeführt werden muss
			adm.Labor.ExecuteFile           	:= IniReadExt("LaborAbruf"	, "Laborabruf_Extern"                 	, ""                	)

		; Laborabruf soll bis zum Laborbuch durchgeführt werden
			adm.Labor.ZeigeLabJournal  	:= IniReadExt("LaborAbruf"	, "Laborabruf_Zeige_Laborbuch"   	, "nein"            	)

		; Speicherort der LDT Dateien (elektronisches Labordatenformat)
			adm.Labor.LDTDirectory         	:= IniReadExt("LaborAbruf"	, "LDTDirectory"                         	, "C:\Labor"    	)

		; Name ihres Labors
			adm.Labor.LbName              	:= IniReadExt("LaborAbruf"	, "LaborName"                           	, ""                 	)

		; Karteikartenkürzel für Eintragungen von Laborwerten, -kommentaren
			adm.Labor.Laborkuerzel        	:= IniReadExt("LaborAbruf"	, "Aktenkuerzel"                          	, "labor"        	)

		; Alarmierungsgrenze in Prozent oberhalb der Normgrenzen (eigentlich Laborjournal.ahk)
			adm.Labor.Alarmgrenze     		:= IniReadExt("LaborAbruf"	, "Alarmgrenze"                        	    , "30%"          	)

		; Übertragen der Laborwerte aus dem Laborbuch ins Laborblatt der Patienten
			adm.Labor.AutoImport		     	:= IniReadExt(compname 	, "Laborimport_automatisieren"      	, "ja"          	)

		; - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : -
		; Telegram - Mitteilungen
		; - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : - - : -
			If (adm.Labor.TGramMsg      	:= IniReadExt("LaborAbruf"	, "TGramVersand"                       	, "Nein"	        	)) {

				; die Nummer ihres Telegram Bots und + oder - Tel z.b. "1+Tel"
					adm.Labor.TGramOpt   	:= IniReadExt("LaborAbruf"	, "TGramOpt"        	                	, "+Tel -Crypt"  	)
					If !RegExMatch(adm.Labor.TGramOpt, "i)([\+\-][a-z]+)\V*([\+\-][a-z]+)")
						Addendum.Labor.TGramMsg := false

				; BotNr muss eine Zahl sein
					adm.Labor.TGramBotNr	:= IniReadExt("LaborAbruf"	, "TGramBotNr"     	                	, "1"      	    	)
					If !RegExMatch(adm.Labor.TGramBotNr, "\d+")
						Addendum.Labor.TGramMsg := false

				; Telegram-Bot-Daten laden
					If Addendum.Labor.TGramMsg {

						BotNr    	:= adm.Labor.TGramBotNr
						BotName 	:= IniReadExt("Telegram", "Bot" BotNr)
						If (InStr(BotName, "Error") || StrLen(BotName) = 0) {
							adm.Labor.TGramMsg	:= false
						} else {
							adm.Telegram            	:= Array()
							adm.Telegram[BotNr]   	:= {	"Token"         	: IniReadExt("Telegram", BotName "_Token")
																	, 	"ChatID"          	: IniReadExt("Telegram", BotName "_ChatID")
																	,	"Active"           	: IniReadExt("Telegram", BotName "_Active", 0)
																	,	"LastMsg"        	: IniReadExt("Telegram", BotName "_LastMsg", "--")
																	,	"LastMsgTime"  	: IniReadExt("Telegram", BotName "_LastMsgTime", "000000")}
						}

					}

			}

		; Laborname ist unbekannt. Skript holt sich den Namen aus einer Albisdatenbank.
			If (StrLen(adm.Labor.LbName) = 0 || InStr(adm.Labor.LbName, "ERROR")) {
				labore := GetDBFData(adm.AlbisDBPath "\LABSTAMM.DBF" , {"ID":"rx:\d+"}, ["ID", "NAME", "PFAD"], 0)
				If (labore.Count() > 1) {
					For idx, labor in labore
						laborliste .= (idx > 1 ? "|" : "") labor.NAME
					LNr := ListBoxAbfrage(Laborliste, "Laborabruf", labore.Count() " Labore gefunden. Wähle ein Labor aus!")
					If (StrLen(LNr) = 0) {
						MsgBox, 1024, % adm.scriptname, % "Sie haben kein Labor ausgewählt!`nDas Skript wird beendet.", 12
						ExitApp
					}
				}
				else
					LNr := 1

				adm.Labor.LbName      	:= labore[LNr].NAME
				adm.Labor.LDTDirectory	:= labore[LNr].PFAD
				IniWrite, % adm.Labor.LbName     	, % adm.ini, % "Laborabruf", % "LaborName"
				IniWrite, % adm.Labor.LDTDirectory	, % adm.ini, % "Laborabruf", % "LDTDirectory"
			}

		; Programm für den Labodatendownload bekannt/vorhanden?
			If (StrLen(adm.Labor.ExecuteFile) = 0 || InStr(adm.Labor.ExecuteFile, "ERROR")) {
				MsgBox, 1024, % adm.scriptname, % " Das Programm für den Download`nder Labordaten ist nicht vorhanden!`n"
																	.	"`t[" adm.Labor.ExecuteFile "]", 12
				FileSelectFile, exeFile, 3, % "C:\", % "Wählen Sie die InfoBoxWebClient.exe aus", ausführbare Dateien (*.exe)
				if ErrorLevel {
					MsgBox, 1024, % adm.scriptname, % " Sie haben nichts ausgewählt!`n"
					ExitApp
				}
				If FileExist(exeFile) {
					adm.Labor.ExecuteFile := exeFile
					IniWrite, % exeFile, % adm.ini, Laborabruf, Laborabruf_Extern
				}
			}
			If !FileExist(adm.Labor.ExecuteFile) && !debug {
				FileAppend, % datestamp(2) " | infoBoxWebClient.exe`t- ist nicht vorhanden. " adm.Labor.ExecuteFile "`n", % adm.LogFilePath
				MsgBox, 1024, % adm.scriptname, %	"[infoBoxWebClient]`nDas auszuführende externe Programm für den`n"
																	.	"elektronischen Datenabruf ist nicht vorhanden`n[" adm.Labor.ExecuteFile "]", 12
				ExitApp
			}

		; Auto Labordatenimport nach erfolgreichem Download von Labordaten ausführen?
			If (StrLen(adm.Labor.AutoImport) = 0 || InStr(adm.Labor.AutoImport, "ERROR")) {
				adm.Labor.AutoImport := false
				MsgBox, 4, % adm.scriptname, % 	"Möchten Sie neue Labordaten automatisch in das`n"
																.	" Laborblatt übertragen ['alle übertragen' im Laborbuch]?"
				IfMsgBox, Yes
					adm.Labor.AutoImport := true
				IniWrite, % (adm.Labor.AutoImport ? "ja" : "nein"), % adm.ini, % compname, % "Laborimport_automatisieren"
			}

	;}

  ; -------------------------------------------------------------------------------------------------------------------------------------------------------
  ; Laborabruf - Automatisierung (infoBoxWebClient und Labordatenimportieren)
  ; -------------------------------------------------------------------------------------------------------------------------------------------------------;{
		; Albis ist nicht gestartet: Abbruch des Laborabrufes
			If !WinExist("ahk_class OptoAppClass") {
				FileAppend	, % (amsg .= datestamp(2) "|" adm.scriptname "()`t- Albis ist nicht gestartet`n")
									, % adm.LogFilePath
				ExitApp
			}

		; Automatisierungsfunktionen von Addendum.ahk per IPC pausieren!
		; Addendum.ahk automatisiert den Laborabruf zum Teil und würde wenn die Funktionen
		; nicht angehalten werden in fehlerhafter RPA-Ausführung enden
			If !(writeStr := LaborAbrufToggle("off")) {
				MsgBox, 1024, % adm.scriptname, % (amsg := writeStr), 6
				FuncExitApp()
			}

		; prüfe evtl. noch geöffnete Laborabruf Dialogfenster (z.B. wurden vom Anwender nicht geschlossen)
			;~ If WinExist("Labordaten ahk_class #32770")


		; 1. externes Programm für Download der LDT Datei starten
			If !infoBoxWebClient()
				FuncExitApp()

		; 2. LDT Datei importieren
			If !AlbisLaborImport(adm.Labor.LbName)
				FuncExitApp()

		; 3. Laborbuch aufrufen
			adm.hLabbuch := AlbisLaborbuchOeffnen()

		; 4. alle ins Laborblatt übertragen
			If adm.hLabbuch {

				; AutoLaborAbruf wiederherstellen
					If !adm.Labor.AbrufStatus && adm.Labor.AutoAbruf
						LaborAbrufToggle("on")

				; Information ins Logfile schreiben falls der Labordatenimport ausgestellt ist
				; ansonsten Laborblattübertragung (LaborImport - nutzt ein anderes Logfile) starten
					amsg .= datestamp(2) "|" adm.scriptname "`t - warte auf Dialog ' Laborbuch übertragen '`n"
					if adm.Labor.AutoImport && adm.Labor.AbrufStatus && !adm.Labor.AutoAbruf  {
						writeStr := "|Laborabruf`t- Labordatenimport nicht möglich! AutoAbruf ist in Addendum ausgeschaltet!"
						FileAppend, % (amsg .= datestamp(2) "|" writeStr "`n"), % adm.LogFilePath
					} else {
						result	:= AlbisLaborAlleUebertragen()
						writeStr := "|Laborabruf`t- Laborbuch 'alle übertragen' ausgeführt [" result "]"
						FileAppend, % (amsg .= datestamp(2) "|" writeStr "`n"), % adm.LogFilePath
					}

			} else {
				writeStr := "|Laborabruf`t- Aufruf des Laborbuch fehlgeschlagen. Labordatenimport musste abgebrochen werden.`n"
				FileAppend, % (amsg .= datestamp(2) "|" writeStr "`n"), % adm.LogFilePath
			}

		; Automatisierungsfunktionen im Hauptskript wieder anschalten
			RestoreLaborAutomationSettings:
			FuncExitApp("normal")

	;}

  ; -------------------------------------------------------------------------------------------------------------------------------------------------------
  ; -------------------------------------------------------------------------------------------------------------------------------------------------------
   ExitApp
  ; -------------------------------------------------------------------------------------------------------------------------------------------------------
  ; -------------------------------------------------------------------------------------------------------------------------------------------------------

  ; -------------------------------------------------------------------------------------------------------------------------------------------------------
  ; RPA Funktionen
  ; -------------------------------------------------------------------------------------------------------------------------------------------------------;{
infoBoxWebClient() {                                                               	;-- Programm für elektronischen Labordatenempfang

	; die Funktion startet die Client.exe, wartet auf die Beendigung des Abrufes und legt ein Dateibackup
	; der empfangenen .LDT Datei an und schließt das externe Programmfenster im Anschluß

	; externes Programm starten
		vianovaClass := "vianova infoBox-webClient" ;" ahk_class WindowsForms10.Window.8"
		amsg .= datestamp(2) "|" A_ThisFunc "() `t- starte externes Programm für den Laborabruf`n"
		Run, % adm.Labor.ExecuteFile
		WinWait, % vianovaClass,, 10
		If !(hVianova := WinExist(vianovaClass)) {
			FileAppend	, % (amsg .= datestamp(2)  "|" A_ThisFunc "() `t- infoBoxWebClient: Fenster nicht gefunden`n")
								, % adm.LogFilePath
			return 0
		}

	; Abruf Button drücken
		ctrlClass := Controls("", "ControlFind, Abruf, Button, return classNN", hVianova)
		VerifiedClick(ctrlClass, hVianova)

	; wartet bis keine Daten mehr hinzukommen
		amsg .= datestamp(2) "|" A_ThisFunc "() `t- rufe Labordaten ab`n"
		adm.Labor.Daten := false
		Loop {
			oAcc := Acc_Get("Object", "4.1.4.5.4.1.4.4.4" , 0, "ahk_id " hVianova)
			t := ""
			Loop % oAcc.accChildCount()	{

				 oAcc := Acc_Get("Object", "4.1.4.5.4.1.4.4.4." A_Index, 0, "ahk_id " hVianova)
				 t .= oAcc.accValue(2) "`n"

				 If RegExMatch(t, "(Dateien\serhalten)|(Sichere.*ins\sArchiv)")
					adm.Labor.Daten := true

				If RegExMatch(t, "Ende\sDateien\sabrufen")
					break

			}

			If (A_Index > 600) || adm.Labor.Daten || RegExMatch(t, "Ende\sDateien\sabrufen")
				break
			sleep 50

		}

	; Dateieingang prüfen
		FileEncoding, CP28605   ; spezielles LDT Encoding (kopierte Datei ist sonst nicht lesbar)
		Loop, Files, % adm.Labor.LDTDirectory "\*.LDT"
		{
				FileGetTime, LDTTime, % A_LoopFileFullPath, M
				LDTTime := SubStr(LDTTime, 1 , 8)
				ANow := A_YYYY . A_MM . A_DD

			; eine neue Datei ist da
				If (ANow = LDTTime) {
					datestamp := (A_YYYY "-" A_MM "-" A_DD " " A_Hour "_" A_Min "_" A_Min)
					FileCopy, % A_LoopFileFullPath, % adm.DBPath "\Labordaten\LDT\" datestamp . A_LoopFileName
					writeText .= A_LoopFileName ", "
				}

		}

	; Logdaten schreiben
		IniWrite	, % datestamp(1) , % adm.ini, % "LaborAbruf", % "Letzter_Abruf"
		If (StrLen(writeTExt) > 0) {
			IniWrite			, % datestamp(1) , % adm.ini, % "LaborAbruf", % "Letzter_Abruf_mit_Daten"
			FileAppend	, % (amsg .= datestamp(2)  "|" A_ThisFunc "() `t- # Daten vorhanden [" RTrim(writeText, ", ") "]`n")
								, % adm.LogFilePath, UTF-8
			LabDataReceived := true
		}	else {
			IniWrite			, % datestamp(1) , % adm.ini, % "LaborAbruf", % "Letzter_Abruf_ohne_Daten"
			FileAppend	, % (amsg .= datestamp(2) "|" A_ThisFunc "() `t- # Keine Daten vorhanden`n")
								, % adm.LogFilePath, UTF-8
			LabDataReceived := false
		}

	; externes Programmfenster schließen
		amsg .= datestamp(2) "|" A_ThisFunc "() `t- WebClient Fenster wird geschlossen`n"
		sleep 500
		Controls("","reset","")
		hVianova := WinExist(vianovaClass)
		ctrlClass := Controls("", "ControlFind, Schließen, Button, return classNN", hVianova)
		WinActivate, % vianovaClass
		WinWaitActive, % vianovaClass,, 2
		If !VerifiedClick(ctrlClass, hVianova) {
			sleep 500
			VerifiedClick(ctrlClass, hVianova)
		}


return LabDataReceived
}


;}

  ; -------------------------------------------------------------------------------------------------------------------------------------------------------
  ; Helfer
  ; -------------------------------------------------------------------------------------------------------------------------------------------------------;{
ListBoxAbfrage(liste, GTitle:="", GText:="")         	{

	global

	inputDone := false

	Gui, LBW: Default
	Gui, +AlwaysOnTop +LastFound -DPIScale +HWNDhLBW
	Gui, Color, cFFCCAA, c777744

	Gui, Font, s11 bold	, Calibri
	Gui, Add, Text   	, xm ym w270                                           	, % GText

	Gui, Font, s11 cWhite Normal, Calibri
	Gui, Add, ListBox	, y+10 w270 r5 0x100 vLBR AltSubmit       	, % liste

	Gui, Font, s11 cBlack Normal, Calibri
	Gui, Add, Button	, xm y+10  	vLBWOK 		gBtnHandler      	, OK
	Gui, Add, Button	, x+20      	vLBWCancel 	gLBWGuiEscape	, Abbruch

	Gui, Show, AutoSize, Laborabruf

	while !inputdone || WinExist(GTitle " ahk_class AutoHotkeyGUI") {
		Sleep 10
		Loop 5
			Sleep 2
	}

	Gui, LBW: Submit, Hide
	Gui, LBW: Destroy

return LBR

LBWGuiClose:
LBWGuiEscape:
	inputDone := true
	LBR := ""
	Gui, LBW: Destroy
return

BtnHandler:

	Gui, LBW: Submit, Hide
	Gui, LBW: Destroy
	inputDone := true

return LBR
}

; - - - - -
LaborAbrufGui(Title:="")                                	{

		global

	; IPC Gui erstellen
		GuiTitle  	:= (StrLen(Title) = 0 ? adm.scriptname : Title) " - IPC Gui"
		AlbisWinID:= WinExist("ahk_class OptoAppClass")
		AlbisPos	:= GetWindowSpot(AlbisWinID)
		LCw      	:= 850, LCh := 300
		amsg    	:= "----------------------------------------------------------" 	"`n"
		amsg     	.= datestamp(1) "  " adm.scriptname " gestartet"         	"`n"
		amsg    	.= "----------------------------------------------------------" 	"`n`n"

		Gui, LC: New, +hwndhLabCall +ToolWindow -Caption +AlwaysOnTop
		Gui, LC: Margin, 0, 0
		Gui, LC: Color, c172842, c172842
		Gui, LC: Font, s8 q5 cWhite, Consolas
		Gui, LC: Add, Edit     	, % "xm ym 	w" LCw " h" LCh " 	vLCdbg	hwndhLCdbg"                                       	, % amsg
		Gui, LC: Add, Progress	, % "xm y+0 	w" LCw " h20 		vLCPrg 	hwndhLCPrg Background172842 cWhite"	, 100

		Gui, LC: Show, % "x" (AlbisPos.X+7) " y" (AlbisPos.Y+AlbisPos.H-LCh-30) " w" LCw " h" LCh+20 " NA", % GuiTitle ;AlbisPos.W-LCw-
		WinSet, Trans, 220, % "ahk_id " hLabCall
		OnMessage(0x4a, "Receive_WM_COPYDATA")
		blockSendWMCOPYDATA := false

	; keine Nachrichten senden, wenn Addendum.ahk nicht läuft
		if !adm.ID {
			blockSendWMCOPYDATA := true
			return 0
		}
		else if blockSendWMCOPYDATA
			return 0

	; Autoupdate Debug Gui
		func_Autoupdate := Func("AutoUpdateLCdbg")
		SetTimer, % func_Autoupdate, 50

}

AutoUpdateLCdbg()                                        	{

	global LC, LCdbg, hLCdbg
	static vOutput, RoundZ

	If (StrLen(amsg) = 0)
		return

	vOutput .= RTrim(RegExReplace(amsg, "\s+\t", "  "), "`n") "`n"
	amsg := ""
	GuiControl, LC:, %hLCdbg%, % vOutput

}

LaborAbrufToggle(Switch="off", Title:="")        	{

	; sendet eine Nachricht an Addendum.ahk, um die Laborabrufautomatisierung an- oder auszuschalten
	; beim ersten Aufruf wird eine versteckte Gui erstellt (die ist notwendig für die Interprocesskommunikation)
	;
	; Switch:                        	off   	- 	Befehl zum Ausschalten senden
	;				on/restore	[destroy]	- 	Befehl zum Anschalten oder was dasselbe ist, restore um das Laborabruf flag wiederherzustellen
	;				      								optional oder auch allein kann destroy übergeben werden um die IPC-MessageGui zu schliessen
	; 													und OnMessage zu beenden

		global

	; sendet eine Nachricht an Addendum.ahk
		if RegExMatch(Switch, "i)(on|restore|off)") {

			; ThreadControl wird hier immer true gesetzt
				ThreadControl := true, eLvl := 1

			; Labor AutoAbruf Funktionalität von Addendum ausschalten
				If InStr(Switch, "off") {
					adm.Labor.AutoAbruf := ""
					writeStr := A_ThisFunc "()  - Addendum hat nicht geantwortet! Die Automatisierungsfunktionen konnten nicht angehalten werden."
					eL := Send_WM_COPYDATA("AutoLaborAbruf|Status|" hLabCall, adm.ID)
					ThreadControl := true
					amsg .= datestamp(2) "|" A_ThisFunc "()  - Laborabruf Abfrage des Labor AutoAbruf Status von Addendum`n"
					while, (ThreadControl && (A_Index <= 100)) {
						ToolTip, % "Laborabruf wartet auf Addendum: " ThreadControl "`nAddendumID: " adm.ID ", " hLabCall "`nEL: " eL "`nA_Index: " A_Index "/" 100
						Sleep 50
					}
					ToolTip
					If ThreadControl {
						writeStr := A_ThisFunc "()  - Labor AutoAbruf Status wurde nicht empfangen`n"
						FileAppend, % (amsg .= datestamp(2) "|" StrReplace(writeStr, "`n", " ")) "`n", % adm.LogFilePath
						eLvl := 0
					}
					else
						amsg .= datestamp(2) "|" A_ThisFunc "()  - Der Labor AutoAbruf ist: " (adm.Labor.AutoAbruf ? "an" : "aus") "`n"

					If adm.Labor.AutoAbruf {
						eL := Send_WM_COPYDATA("AutoLaborAbruf|aus|" hLabCall, adm.ID)
						amsg .= datestamp(2) "|" A_ThisFunc "()  - Der AutoLaborabruf wurde vorübergehend abgeschaltet.`n"
					}
					adm.Labor.AbrufStatus := false
				}
			; AutoLaborAbruf Funktionalität von Addendum einschalten oder wiederherstellen
				else if RegExMatch(Switch, "i)(on|restore)") {
					If (adm.Labor.AutoAbruf) || RegExMatch(Switch, "i)on")
						eL := Send_WM_COPYDATA("AutoLaborAbruf|an|" hLabCall, adm.ID)
				}

			; auf Antwort von Addendum.ahk warten (max. 5s wartet das Skript hier auf eine Änderung der ThreadControl Variable)
			; die Messageworker() Funktion ändert die Variable auf false sobald eine Antwort von Addendum.ahk eingetroffen ist
				ThreadControl := true
				while, (ThreadControl && (A_Index <= 100))
					Sleep 50

			; keine Antwort, dann notieren und Fehler zurückmelden
				If ThreadControl {
					adm.Labor.AbrufStatus := false
					writeStr := A_ThisFunc "()  - Addendum hat nicht geantwortet! Automatisierungsfunktionen konnten nicht wiederhergestellt werden.`n"
					FileAppend, % (amsg .= datestamp(2) "|" StrReplace(writeStr, "`n", " ")) "`n", % adm.LogFilePath
					eLvl := 0
				} else {
					adm.Labor.AbrufStatus := true
					amsg .= datestamp(2) "|" A_ThisFunc "()  - Der AutoLaborabruf ist wieder eingeschaltet.`n"
				}

		}

	; IPC-MessageGui schliessen, OnMessage beenden
		If RegExMatch(Switch, "i)destroy")
			return 1


return eLvl
}

; - - - -
FuncExitApp(ExitReason:="failure", wTime:=20)	{           ; wTime = Wartezeit bis zum Beenden in Sekunden

		global func_Autoupdate, LC, hLCPrg

		If (ExitReason = "failure")
			LaborAbrufToggle("restore destroy")

		OnMessage(0x4a, "")                                    	; RPC aus
		SetTimer, % func_Autoupdate, Delete              	; Gui Autoupdate aus

		wTime := wTime*1000,	Interrupts := 100
		Loop, % Floor(wTime/Interrupts) {
			;~ ToolTip % 100 - Floor((100*(A_Index*interrupts))/wTime)
			GuiControl, LC:, %hLCPrg%, % 100 - Floor((100*(A_Index*interrupts))/wTime)
			sleep % Interrupts
		}
		ExitApp

}

MessageWorker(msg)                                      	{   		;--  verarbeitet die eingegangen Nachrichten

		recv := {	"txtmsg"		: (StrSplit(msg, "|").1)
					, 	"opt"     	: (StrSplit(msg, "|").2)
					, 	"fromID"	: (StrSplit(msg, "|").3)}

		;SciTEOutput("`nLabor Incoming: `nmsg: " recv.txtmsg "`nfrom: " recv.fromID "`nThreadControl: " ThreadControl "`n")
	; main script receive's data from this thread and calls to continue the OCR process
		If InStr(recv.txtmsg, "Laborabruf angehalten")            	; Thread waits until ThreadControl is false
			ThreadControl := false
		else If InStr(recv.txtmsg, "Laborabruf fortgesetzt")                ; to stop on emergency
			ThreadControl := false
		else If InStr(recv.txtmsg, "AutoLaborAbruf Status") {
			SciTEOutput(" recv.Opt: "  recv.Opt )
			adm.Labor.AutoAbruf := recv.Opt ? true : false
			ThreadControl := false
		}

		PraxTT(recv.txtmsg, "6 1")

return
}


;}

  ; -------------------------------------------------------------------------------------------------------------------------------------------------------
  ; Includes
  ; -------------------------------------------------------------------------------------------------------------------------------------------------------;{
#Include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Datum.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PDFHelper.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PraxTT.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Window.ahk

#Include %A_ScriptDir%\..\..\lib\acc.ahk
#Include %A_ScriptDir%\..\..\lib\class_JSON.ahk
#Include %A_ScriptDir%\..\..\lib\class_Neutron.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\lib\Sift.ahk
;}

  /*  weitere Fenster für die Automatisierung

xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
------------------------------------------------------------
ALBIS
ahk_class #32770
Soll der unbekannte Parameter
P1
übernommen werden?
xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
Dateiübernahme
ahk_class #32770
------------------------------------------------------------
Im Zielverzeichnis liegen noch LDT-Dateien !

Soll der Abruf erfolgen ?
xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

 */



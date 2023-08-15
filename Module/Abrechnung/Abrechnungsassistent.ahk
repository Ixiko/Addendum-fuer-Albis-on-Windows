; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;                                                                    	💰	Addendum Abrechnungsassistent	💰
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;      Funktion: 				⚬	prüft auf Berechnungsfähigkeit von Vorsorgeuntersuchungs- und beratungsziffern (01732,01745/56, 01740,
;										01747, 01731) nach den EBM-Regeln entsprechend Alter und Abrechnungsfrequenz
;                                	⚬	Abrechnungsüberprüfung für Chronikerpauschalen (eingegeben wurden)
;
;
;		Hinweis:				⚬	Gui friert manchmal ein. [--- aus class_LV_Colors.ahk --]
;											"Um den Verlust von Gui-Ereignissen und Meldungen zu vermeiden, muss der Message-Handler eventuell
; 											auf "kritisch" gesetzt werden. Dies kann erreicht werden, indem die Instanzeigenschaft 'Critical' auf den
; 											gewünschten Wert gesetzt wird (z. B. MyInstance.Critical := 100). Neue Instanzen sind standardmäßig
;                                   		auf 'Critical, Off' eingestellt. Auch wenn es manchmal nötig ist, können ListViews oder das gesamte Gui
;                                   		unter (un?)bestimmten Umständen nicht mehr reagieren, wenn Critical gesetzt ist und die ListView ein g-Label hat."
;
;		Abhängigkeiten:	⚬	siehe includes
;
;	                    			Addendum für Albis on Windows
;                        			by Ixiko started in September 2017 - last change 03.07.2023 - this file runs under Lexiko's GNU Licence
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
; weitere Regeln: 	ICD T14.x -> Wundversorgung
;							Laborleistung wie BSG, Glucose, INR
;                         	Impfleistung <--> Diagnose , wenn  eines von beidem fehlt


  ; Einstellungen   	;{
	#NoEnv
	#Persistent
	#MaxMem                    	, 256
	#KeyHistory                  	, 0

	SetBatchLines                	, -1
	SetControlDelay            	, -1
	SetWinDelay                    	, -1

	SetTitleMatchMode         	, RegEx
	SetTitleMatchMode         	, Slow
	DetectHiddenText          	, On
	DetectHiddenWindows    	, On

	ListLines                        	, Off

	cJSON.EscapeUnicode   	:= "UTF-8"

  ; startet Windows Gdip
   	If !(pToken:=Gdip_Startup()) {
		MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
		ExitApp
	}

  ; Tray-Icon darstellen
	If (hIconAbrAssist := Create_Abrechnungsassistent_ico())
		Menu, Tray, Icon, % "hIcon: " hIconAbrAssist

  ;}

  ; Daten laden    	;{

	global Albis, PatDB, Addendum, habr
	global VSEmpf := Object()

  ; Albis Datenbankpfad / Addendum Verzeichnis
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
	RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)

  ; Zugriff auf die Addendum.ini festlegen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	If !(workini := IniReadExt(AddendumDir "\Addendum.ini")) {
		MsgBox, % "Einstellungen aus der Addendum.ini konnten nicht geladen werden."
		ExitApp
	}

  ; AlbisPfade
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Albis := GetAlbisPaths()                                                                                         	; include/Addendum_Internal.ahk

  ; Addendum
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Addendum                     	:= Object()
	Addendum.Dir                	:= AddendumDir
	Addendum.Ini                  	:= Addendum.Dir "\Addendum.ini"                                  	; (Netzwerk) Pfad zu den Addendum Einstellungen
	Addendum.DBPath          	:= Addendum.Dir "\logs'n'data\_DB"                               	; Pfad zu Daten von Addendum
	Addendum.DBasePath      	:= Addendum.Dir "\logs'n'data\_DB\DBase"                    	; Pfad zu Daten über einige Albis dBase Dateien
	Addendum.AlbisDBPath   	:= Albis.db                                                                   	; Datenbankverzeichnis von Albis

	Addendum.compname    	:= StrReplace(A_ComputerName, "-")                                	; der Name des Computer auf dem das Skript läuft
	Addendum.ThousandYear	:= SubStr(A_YYYY, 1, 2)                                                	; "20"
	Addendum.IndisToRemove	:= IniReadExt("Abrechnungshelfer", "AusIndis_ToRemove")
	Addendum.IndisToRemove	:= IniReadExt("Abrechnungshelfer", "AusIndis_ToRemove")
	Addendum.IndisToHave 	:= IniReadExt("Abrechnungshelfer", "AusIndis_ToHave")
	Addendum.ChangeCave 	:= IniReadExt("Abrechnungshelfer", "Abrechnungsassistent_CaveBearbeiten"	, "Nein")
	Addendum.letztesQuartal 	:= LTrim(IniReadExt("Abrechnungshelfer", "Abrechnungsassistent_letztesArbeitsquartal"), "0")
	Addendum.abrOnTop    	:= IniReadExt(Addendum.compname "Abrechnungsassistent_AlwaysOnTop", "Nein")
	;~ SciTEOutput(Addendum.compname ", " Addendum.abrOnTop)
	Addendum.abrOnTop     	:= InStr(Addendum.abrOntop, "ERROR") ? false : true
	Addendum.abrOnPos    	:= IniReadExt(Addendum.compname "Abrechnungsassistent_TilingModus", "F")
	;~ SciTEOutput(Addendum.abrOnPos)
	Addendum.abrOnPos     	:= InStr(Addendum.abrOnPos, "ERROR") ? "F" : Addendum.abrOnPos
	Addendum.KeinBMP       	:= IniReadExt("Addendum", "KeinBMP", "Ja")
	Addendum.ShowIndis       	:= IniReadExt("Addendum", "Zeige_Ausnahmeindikationen", "Nein")
	Addendum.Debug          	:= true

  ; Befundkürzel laden
	aDB := new AlbisDB(Addendum.AlbisDBPath)
	Addendum.BefKuerzel := aDB.BefundKuerzel()
	;~ SciTEOutput(Addendum.BefKuerzel.Count())

  ; freie Tage und Praxisurlaub werden aus der Addendum.ini gelesen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Urlaub	:= IniReadExt("Allgemeines", "Urlaub")
	If (StrLen(Urlaub) > 0 && Urlaub <> "ERROR") {
		Addendum.Praxis := Object()
		Addendum.Praxis.Urlaub := vacation.ConvertHolidaysString(Urlaub)
	}

  ; Patientendaten laden
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	SplashText(StrReplace(A_ScriptName, ".ahk"), "Lade Patientendatenbank...")
	infilter := ["NR", "VNR", "VKNR", "PRIVAT", "ANREDE", "ZUSATZ", "NAME", "VORNAME", "GEBURT", "TITEL", "PLZ", "ORT", "STRASSE"
				, 	"TELEFON", "GESCHL", "SEIT", "ARBEIT", "HAUSARZT", "GEBFREI", "LESTAG", "GUELTIG", "FREIBIS", "KVK"
				, 	"TELEFON2", "TELEFAX", "LAST_BEH", "DEL_DATE", "MORTAL", "HAUSNUMMER", "RKENNZ"]
	PatDB := ReadPatientDBF(Addendum.AlbisDBPath, infilter, "EMail=0 allData")
	SplashText(Skriptname := StrReplace(A_ScriptName, ".ahk"), "Patientendatenbank mit " PatDB.Count() " Namen geladen", 3)

  ; Albisdatenbank - Objekt für Zugriff und Datenauswertung erstellen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	global DBAccess := new AlbisDb(Addendum.AlbisDBPath, 0)

  ; aktuelles Quartal ins Format QYY konvertieren (z.B. 421)
  ; als Abrechnungsquartal wird das Vorquartal genommen (wenn noch keine 3 Wochen innerhalb des aktuellen Quartal vergangen sind)
  ; - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -  - -
	Tagesdatum     	:= A_DD "." A_MM "." A_YYYY
	abrQuartaljahr 	:= GetQuartalEx(Tagesdatum, "QYY")

  ; noch in den ersten Wochen eines neuen Quartal?
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	ReIndex := false
	If (Trimester(A_MM)=1 && A_DD <= 21)  {
		aktQuartal       	:= SubStr(abrQuartaljahr, 1, 1)
		aktJahr         	:= SubStr(abrQuartaljahr, 2, 2)
		abrQuartal      	:= aktQuartal-1=0 ? 4 : aktQuartal-1
		abrJahr         	:= SubStr("0" (aktQuartal-1=0 ? aktJahr-1 : aktJahr), -1)
		abrQuartaljahr 	:= abrQuartal abrJahr
	}

	If (!RegExMatch(Addendum.letztesQuartal, "^[1-4][0-9]{2}$") || Addendum.letztesQuartal <> abrQuartaljahr)
		ReIndex := true

	If ReIndex {
		Abrechnungsquartal := Addendum.letztesQuartal := abrQuartaljahr
		SaveAbrechnungsquartal(Abrechnungsquartal)
	}
	else
		Abrechnungsquartal := Addendum.letztesQuartal

	;}

  ; uxtheme Änderung
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	uxtheme := DllCall("GetModuleHandle", "str", "uxtheme", "ptr")
	SetPreferredAppMode := DllCall("GetProcAddress", "ptr", uxtheme, "ptr", 135, "ptr")
	DllCall(SetPreferredAppMode, "int", 1) ; Dark

  ; Albis nach vorne holen
	SplashText(Skriptname, "Der Abrechnungsassistent wird gleich angezeigt....")
	AlbisActivate(1)

  ; Abrechnungsassistent anzeigen
	AbrAssistGui(Abrechnungsquartal, ReIndex)

return

  ; Hotkey's          	;{
#IfWinActive, Abrechnungsassistent ahk_class AutoHotkeyGUI
~ESC:: ;{
	saveGuiPos(habr, "Abrechnungsassistent")
	ExitApp
return ;}
#IfWinActive
~^!b:: ;{
	saveGuiPos(habr, "Abrechnungsassistent")
	Reload
return ;}

;}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; grafische Anzeige
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
AbrAssistGui(Abrechnungsquartal, ReIndex=false, Anzeige=true)                        	{

		global

	; Variablen und Einstellungen   	;{

		static EventInfo, empty := "- - - - - - -", LVCol
		static LVClear 	:= LVw := 0                                                                     	; (flag) 1 wenn beide Listviews keine Daten enthalten (sauber sind)
		static FSize    	:= A_ScreenHeight > 1080 ? 11 : 10                      	; FSize - ist abhängig von der Monitorauflösung

		local hAlbis, winpos5

		abrQuartal 	:= Abrechnungsquartal
		MarginX    	:= MarginY := 5
		GuiH	     	:= 1080

	; Anpassungen für 2k bzw. (>)4k Monitore
		If (A_ScreenHeight > 1080)
			LVCol	:= FSize>10 ? [60, 120, 30, 25, 30, 30, 30, 25, 25, 130, 60, 60]	: [60, 130, 30, 25, 40, 40, 38, 40, 45, 90, 60, 60]
		else
			LVCol	:= FSize>10 ? [55, 130, 30, 25, 32, 32, 35, 35, 40, 80, 50 , 50]	: [55, 120, 30, 25, 32, 32, 35, 35, 40, 120, 60 , 60]

	; Listviewbreite, Breite der Gui
		vertScrollw := GetScrollInfo(1, 0)  	; SB_Vert = 1 ; SB_HORZ = 0
		For colNr, colsize in LVCol
			LVw += colsize
		guiW := LVw + vertScrollw + 2*MarginX

	; Euro Preise der Gebührennummern (wo speichert Albis diese Daten?)
		Euro := {"#01731"        	: 15.82
					, "#01732"          	: 36.73
					, "#01745"          	: 27.80
					, "#01746"          	: 23.55
					, "#01737"          	: 6.34
					, "#01740"          	: 12.75
					, "#01747"          	: 9.01
					, "#03001"        	: 25.03
					, "#03002"        	: 15.80
					, "#03003" 	    	: 12.84
					, "#03004"         	: 16.46
					, "#03005" 	    	: 22.25
					, "#03040" 	    	: 15.55
					, "#03220"          	: 14.65
					, "#03221"          	: 4.51
					, "#03230"          	: 14.42
					, "KVU"              	: 15.82
					, "GVU"             	: 35.82
					, "HKS1"            	: 27.80
					, "HKS2"            	: 22.96
					, "GB1"              	: 12.57
					, "GB2"              	: 19.36
					, "Colo"             	: 12.75
					, "Aorta"            	: 9.01
					, "Pausch0-4"        	: 25.03
					, "Pausch5-18"      	: 15.80
					, "Pausch19-54" 	: 12.68
					, "Pausch55-75" 	: 16.46
					, "Pausch76-" 		: 22.25
					, "PauschZ"    		: 15.35
					, "Chron1"         	: 14.46
					, "Chron2"         	: 4.450
					, "SPES"          	  	: 6.34}

	; Komplexe laden/erstellen
		SplashText("Vorbereitung", "Abrechnungsziffern aller Quartale laden ...")
		candidates := VorsorgeKomplexe(Abrechnungsquartal, ReIndex, true)                           		; lädt oder erstellt alle notwendigen Daten
		ObjectReplace(candidates)                                                                                              	; übergibt das Objekt "candidates" an das super-globale "VSEmpf" Objekt

	; Farbeinstellungen des Albisnutzer für Kürzel aus der Datenbank lesen
		befkuerzel := DBAccess.BefundKuerzel()

	; EBM Regeln laden
		SplashText("Vorbereitung", "lade Regeln für die EBM-Abrechnung...")
		If !IsObject(VSRule := DBAccess.GetEBMRule("Vorsorgen"))
			throw A_ThisFunc ": EBM Vorsorgeregeln konnten nicht übermittelt werden!"

	; letzte Fenstereinstellungen laden   	;{
		workpl	:= WorkPlaceIdentifier()
		winpos 	:= IniReadExt(Addendum.compname, "Abrechnungsassistent_Pos_" workpl)
		winpos 	:= winpos ~= "i)^(ERROR|\s+|)$" ? "xCenter yCenter " : winpos
		options	:= IniReadExt(Addendum.compname, "Abrechnungsassistent_Optionen" , "1|1|1")
		options := options ~= "i)^(ERROR|\s+|)$" ? StrSplit("1|1|1", "|") : StrSplit(options, "|")
	;}

	; Albis Fensterposition                     	;{
		AlbisFensterPosition:
		wp := PosStringToObject(winpos)
		If (hAlbis := AlbisWinID()) {

			; Monitordimensionen auf dem Albis geöffnet ist
				ap     	:= GetWindowSpot(hAlbis)
				Mon		:= ScreenDims(monIndex := GetMonitorIndexFromWindow(AlbisWinID()))
				TBarH	:= TaskbarHeight(monIndex)

			; Albisfenster verschieben wenn es zu weit rechts ist, die Gui soll rechts daneben passen
				If (ap.X > 0 || ap.Y > 0)
					SetWindowPos(AlbisWinID(), 0, 0, ap.W, ap.H)
		}
		else {

			; Albis starten ?
				MsgBox, 0x4	, % StrReplace(A_ScriptName, ".ahk")
							,	% "Albis muss gestartet sein, wenn Sie alle Funktionen`n"
							.   	"nutzen wollen.`n Albis jetzt starten?"
				IfMsgBox, Yes
				{
					nil := 0
					WinWait, ahk_class OptoAppClass,, 80
					If !WinExist("ahk_class OptoAppClass")
						ExitApp
					goto AlbisFensterPosition
				}

			; Monitorgrößen des ersten Monitor
				Mon		:= ScreenDims(1)
				TBarH	:= TaskbarHeight(1)

		}

	  ; Backup der Albisfensterposition
		wpAlbis := ap

	;}

	; andere Variablen                        	;{
		gcolor              	:= "666666"    ; Hintergrund
		tcolor               	:= "BCCCE4"	; Listview
		tcolorD           	:= "3B8600"		; Markierung
		tcolorP            	:= "333333"		; Unterteilung
		tcolorM            	:= "DADEE7"   	; Vorschläge
		tcolorI            	:= "5ED000"  	; Symbole
		tcolorBG       	:= "D0C1C1"
		tcolorBG1      	:= "BCCCE4"
		tcolorBG2      	:= "F8F5F5"
		TextColor1     	:= "65DF00"
		TextColor2     	:= "D9E137"
		SelectionColor1	:= "0x0078D6"
		abrTabNames	:= "Vorsorge-Guru|ICD-Krauler|Großer-Geriater|freier Statistiker|Einstellungen"
	;}

	;}

	; Gui                                       	;{

		Gui, abr: new    	, % "HWNDhabr  " (ontop ? " +AlwaysOnTop " : "")   ;
		Gui, abr: Color  	, % "c" gcolor , % "c" tcolor
		Gui, abr: Margin	, 5, 5
		Gui, abr: Font    	, % "s" FSize-1 " q5 Normal c" TextColor1 , Futura Bk Bt
		Gui, abr: Add     	, Tab   	, % "x1 y1  	w" LvW " h" 25 " HWNDabrHTab vabrTabs gabr_Tabs "                   	, % abrTabNames


	;-: TAB 1 - VORSORGE-GURU                                                 	;{
		Gui, abr: Tab    	, 1

	  ;-: Kopf                                                                                  ;{
		Gui, abr: Font    	, % "s" FSize  " c" TextColor1
		Gui, abr: Add    	, Text        	, % "xm+5 ym+30    	BackgroundTrans     	"                                            	, % "Quartal"
		Gui, abr: Font    	, % "s" FSize-1 " cBlack"
		Gui, abr: Add    	, Edit           	, % "x+5 yp+-9 w" (FSize-1)*5 " 	              	vabrQ           	gabrHandler"	, % Abrechnungsquartal
		GuiControlGet, cp, abr: Pos, abrQ
		ButtonH :=  cpH-2
		Gui, abr: Add    	, Button        	, % "x" cpX+cpW " y" cpY+1 " h" cpH-2 "	      	vabrQUp       	gabrHandler"   	, % "▲"
		Gui, abr: Add    	, Button        	, % "x+0 h" cpH-2 "                        	         	vabrQDown   	gabrHandler"   	, % "▼"

		Gui, abr: Font    	, % "s" FSize " c" TextColor1
		Gui, abr: Add    	, Text        	, % "x+20     	BackgroundTrans             	"                                               	, % "mögliches EBM plus: "
		Gui, abr: Font    	, % "s" FSize " c" TextColor2
		Gui, abr: Add    	, Text        	, % "x+0       	BackgroundTrans  				vEBMEuroPlus"                        	, % "00000.00 €  "

		Gui, abr: Font    	, % "s" FSize " c" TextColor1
		Gui, abr: Add    	, Text        	, % "x+10       	BackgroundTrans             	"                                               	, % "Krankenscheine: "
		Gui, abr: Font    	, % "s" FSize " c" TextColor2
		Gui, abr: Add    	, Text        	, % "x+0       	BackgroundTrans  				vEBMScheine"                         	, % VSEmpf["_Statistik"].KScheine

	;-:      🔐
		Gui, abr: Font	, % "s" FSize+1  " c" TextColor1
		Gui, abr: Add	, Button  	, % "x" LvW+MarginX-40 " y" cpY "	w20 h20 hwndabrHOnPos vabrOnPos	gabrOn", % (	Addendum.abrOnPos="R"	? "◨"
																																												: 	Addendum.abrOnPos="L"	? "◧" : "🗗")
		Gui, abr: Add	, Button  	, % "x+2                                      	w20 h20 hwndabrHOnTop vabrOnTop	gabrOn", % (Addendum.abrOnTop    	? "🔐" : "🔏")
		Gui, abr: Add	, Button  	, % "x+2                                      	w20 h20 hwndabrHReload vabrReload	gabrReload", ♻
		AddTooltip(abrHOnTop, "AlwaysOnTop ist " 	(Addendum.abrOnTop      	? "eingeschaltet" : "ausgeschaltet"))
		AddTooltip(abrHOnPos,  "Fensteraufteilung: "	(Addendum.abrOnPos="R" 	? "Albis | Abrechnungsassistent"
														              	:	 Addendum.abrOnPos="L" 	? "Abrechnungsassistent | Albis" : "keine"))
		AddTooltip(abrHReload, "drücken um Skript neu zu laden")
		WinSet, Style 	, +0x800000, % "ahk_id " abrHReload
		WinSet, Style 	, +0x800000, % "ahk_id " abrHOnTop
		WinSet, Style 	, +0x800000, % "ahk_id " abrHOnPos
		WinSet, ExStyle	, +0x000000, % "ahk_id " abrHReload
		WinSet, ExStyle	, +0x000000, % "ahk_id " abrHOnTop
		WinSet, ExStyle	, +0x000000, % "ahk_id " abrHOnPos

	;}

	  ;-: Section 1                                                                           	;{
		Gui, abr: Add  	, Progress   	, % "x14 y+10 w290 h2 -E0x20000 vabrPS1 HWNDabrHPS1 c" tcolorP " Background" tcolorP, 100
	  ;}

	  ;-: Refresh / ReIndex                                                              	;{
		Gui, abr: Font	, % "s" FSize+6 " q5 Normal c" tcolorI
		Gui, abr: Add	, Text        	, % "xm+5 y+5 BackgroundTrans vabrSymbol1 "                                           	, % "♻"
		GuiControlGet, cp, abr: Pos, abrSymbol1
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal c" TextColor1
		Gui, abr: Add	, Button        	, % "x+2 y" cpY+2 " h" ButtonH "       	vabrRefresh   	gabrHandler"               	, % "Liste auffrischen"

		Gui, abr: Font	, % "s" FSize+6 " q5 Normal c" tcolorI
		Gui, abr: Add	, Text        	, % "x+15 y" cpY-1 " BackgroundTrans"                                                         	, % "🔄"
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal c" TextColor1
		Gui, abr: Add	, Button        	, % "x+2 y" cpY+2 " h" ButtonH " hwndabrHRidx vabrReIndex gabrHandler"   	, % "Liste erneuern"
		AddTooltip(abrHRidx, "Erstellt die Quartalsdaten neu!`nACHTUNG: markierte Einträge gehen verloren!")

		Gui, abr: Font	, % "s" FSize+6 " q5 Normal c" tcolorI
		Gui, abr: Add	, Text        	, % "x+20 y" cpY-1 " vabrHide BackgroundTrans"                                           	, % "☑"
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal c" TextColor1
		Gui, abr: Add	, Button        	, % "x+2 y" cpY+2 " h" ButtonH "         	       	vabrBearbeitet  	gabrHandler"   	, % "Markierte ausblenden"

		xhwnd := " hwndabrHunhide "
		GuiControlGet, cp, abr: Pos, abrHide
		Gui, abr: Font	, % "s" FSize+18 " q5 Normal c" tcolorI
		Gui, abr: Add	, Text        	, % "x+10 y" cpY-12 " BackgroundTrans"                                                       	, % "◻"
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal c" TextColor1 , Futura Bk Bt
		Gui, abr: Add	, Button        	, % "x+2  y" cpY+2 " h" ButtonH " " xhwnd " 	vabrUnHide  gabrHandler"       	, % "alle einblenden"

		WinSet, Redraw,, % "ahk_id " abrHPS1

	;}

	  ;-: Optionen                                                                         	;{
		CBOpt := "BackgroundTrans Checked"
		Gui, abr: Add	, Checkbox   	, % "xm+5 	y+10 "             	CBOpt " 	vabrYoungStars	gabrHandler"    	, % "Alter 18 bis 34 anzeigen"
		GuiControlGet, cp, abr: Pos, abrYoungStars
		Gui, abr: Add	, Checkbox   	, % "x+20 "                        	CBOpt "	vabrOldStars 	gabrHandler"     	, % "Alter ab 35 anzeigen"
		Gui, abr: Add	, Checkbox   	, % "x+20 "                        	CBOpt " 	vabrAngels    	gabrHandler"    	, % "Verstorbene anzeigen"
		Gui, abr: Add	, Text        	, % "x+40           	BackgroundTrans  	vabrLVRows"                               	, % "[0000 Patienten]"

	;}

	  ;-: Patientenliste / Listview                                                    	;{
		LVOptions3 	:= "vabrLV gabrLVHandler HWNDabrhLV Checked Grid -E0x200 -LV0x10 AltSubmit -Multi "
		LVLabels3 	:= "ID|Name, Vorname|Alt.|G|Gv|Hk|Kv|C|A|Chroniker|GB|Pauschale"
		LVwidth     	:= 0
		For idx, w in LVCol
			LVwidth += w

		Gui, abr: Font	, % "s" FSize-2 " q5 Normal cBlack"
		Gui, abr: Add	, ListView    	, % "xm+5 y+5 w" LVwidth+20 " r15 " LVOptions3                                    	, % LVLabels3
		LV_ModifyCol(1 	, LVCol.1  	" Right Integer")
		LV_ModifyCol(2 	, LVCol.2  	" Left Text")
		LV_ModifyCol(3 	, LVCol.3  	" Center Integer")
		LV_ModifyCol(4 	, LVCol.4  	" Center Text")
		LV_ModifyCol(5 	, LVCol.5  	" Center Text")
		LV_ModifyCol(6 	, LVCol.6  	" Center Text")
		LV_ModifyCol(7 	, LVCol.7  	" Center Text")
		LV_ModifyCol(8 	, LVCol.8  	" Center Text")
		LV_ModifyCol(9 	, LVCol.9  	" Center Text")
		LV_ModifyCol(10	, LVCol.10	" Left Text")
		LV_ModifyCol(11	, LVCol.11	" Left Text")
		LV_ModifyCol(12	, LVCol.12	" Left Text")
		;}

	  ;-: Patienten und Untersuchungsauswahl                               	;{
		GuiControlGet, cp, abr: Pos, abrLV
		PUpX := cpX + cpW + 10
		GuiControlGet, cp, abr: Pos, abrPS1
		PUpY := cpY + 10
		Gui, abr: Font	, % "s" FSize   	" q5 Normal c" TextColor1 , Futura Bk Bt
		Gui, abr: Add	, Text        	, % "xm+5 y+1 BackGroundTrans"                                                       	, % "gewählte(r) Patient(in): "
		Gui, abr: Font	, % "s" FSize+1 " q5 Bold  c" TextColor1
		Gui, abr: Add	, Text        	, % "x+0 w" FSize*30 "  	vabrPat"                                                           	, % " "
		;}

	  ;-: Checkboxen für Untersuchungen                                      	;{
		abrCheckbox := Object()
		Gui, abr: Font	, % "s" FSize-1   	" q5 Normal c" TextColor1
		Loop % VSRule.VSFilter.MaxIndex() {
			fee := VSRule.VSFilter[A_Index].exam
			If !abrCheckbox.haskey(fee) {
				Gui, abr: Add	, Checkbox   	, % (A_Index = 1 ? "xm+5 y+1" : "x+25")  " vabr_" fee                	, % fee
				GuiControlGet, cp, abr: Pos, % "abr_" fee
				abrCheckbox[fee] := cpX
			}
		}
		;}

	  ;-: zeigt den ersten möglichen Abrechnungstag an                   	;{
		floop := 1
		Gui, abr: Font, % "s" FSize-3 " q5 Normal c" TextColor1
		For fee, feeX in abrCheckbox {
			Gui, abr: Add, Text, % (floop = 1 ? "xm+5 y+2" : "x" feeX " y" cpY)  " vabr_Date" fee " "            	, % " [00.00.0000] " ; BackgroundTrans
			GuiControlGet, cp, abr: Pos, % "abr_Date" fee
			;~ GuiControl, abr:, % "abr_Date" fee, % ""
			floop ++
		}
		;}

	  ;-: zeigt alle Vorschläge in einer Liste an                               	;{
		CLV := New LV_Colors(abrhLV)
		CLV.Critical := 100
		CLV.SelectionColors(selectionColor1)
		gosub abrVorschlaege
	  ;}

	  ;-: Section 2                                                                        	;{

	  ;-: Euro Summe* aller EBM Gebühren anzeigen  und
	  ;-: Zahl der angezeigten Patienten                                    		;{
		; (*vorausgesetzt alle Vorschläge würden zur Abrechnung kommen)
		abrEBMPlus(Round(EBMPlus, 2))
		;}

	  ;-: Quartalkalender für spätere Abrechnungstage hinzufügen 	;{
		QuartalsKalender(Abrechnungsquartal, 0, "abr", TextColor1)
	  ;}

	;}

	;}

	;-: TAB 2 - ICD-Krauler                                                            	;{
		Gui, abr: Tab  	, 2

		GuiControlGet, cp, abr: Pos, % "abrCal"
		Gui, abr: Add	, Edit  	, % "xm+5 ym+35  w" (FSize-1)*5 " 	                        	vabrQ2           	gabrHandler"	, % Abrechnungsquartal
		Gui, abr: Add	, Button	, % "x+5                                                                   	vabrICDKrauler gabrHandler"  	, % "nach Chronikern kraulen lassen"
		LVOptions4 	:= "vabrICDLV gabrICDLVHandler HWNDabrICDhLV Checked Grid -E0x200 -LV0x10 AltSubmit -Multi "
		LVLabels4 	:= "    ID|Name, Vorname|Geb.Datum|chron.Pauschale"

		Gui, abr: Font	, % "s" FSize-2 " q5 Normal cBlack"
		Gui, abr: Add	, ListView    	, % "xm+5 y+5 w" LVwidth+20 " r55 " LVOptions4                                              	, % LVLabels4
		LV_ModifyCol(1 	, 80  	" Right Integer")
		LV_ModifyCol(2 	, 160  	" Left Text")
		LV_ModifyCol(3 	, 90  	" Lef Text")
		LV_ModifyCol(4 	, 200  	" Left Text")

	;}

	;-: TAB 3 - Der große Geriater                                                	;{
		Gui, abr: Tab  	, 3
	;}

	;-: TAB 4 - Der freie Statistiker                                                  	;{
		Gui, abr: Tab  	, 4

		Gui, abr: Font  	, % "s" FSize+2 " q5 Normal c" TextColor1 , Futura Bk Bt
		Gui, abr: Add	, Text	, % "xm+5 ym+40 ", F I L T E R  für den freien Statistiker

	;-: ArztID
		Gui, abr: Font  	, % "s" FSize " q5 Normal c" TextColor1
		Gui, abr: Add	, Text	, % "xm+5 y+15 vabrTArztID"					                                                                        , ArztID
		Gui, abr: Font  	, % "s" FSize " q5 Normal cBlack"
		Gui, abr: Add	, Edit  	, % "x+2 w" FSize*3 "  vabrArztID HWNDabrHArztID  	gabrHandler"	                 	, % "1  "
		AddTooltip(abrHArztID, "Möchtest Du die ArztID nicht filtern, leere das Feld oder setzte eine 0 ein.")
		cp := GuiControlGet("abr", "POS", "abrTArztID")
		dp := GuiControlGet("abr", "POS", "abrArztID")
		GuiControl, abr: Move, abrArztID, % "y" cp.Y+Floor(cp.H/2) - Floor(dp.H/2)
		dp := GuiControlGet("abr", "POS", "abrArztID")

	;-: Datum von
		Gui, abr: Font  	, % "s" FSize " q5 Normal c" TextColor1
		Gui, abr: Add	, Text	, % "x" dp.X+dp.W+10 " y" cp.Y                                                                                       	, Datum von-bis
		Gui, abr: Font  	, % "s" FSize " q5 Normal cBlack"
		Gui, abr: Add	, Edit  	, % "x+3 y" dp.Y " w" FSize*9 "  vabrFSDvon HWNDabrHFSDvon          	gabrHandler"	  	, % ""
		AddTooltip(abrHFSDvon, "Tragen hier das Datum ein von dem aus Du frühestens mit der Suche starten willst.`n"
											. "Du kannst hier ein Datum einsetzen in der Form dd.mm.YYYY oder einfach nur`n "
											. "ein Quartal.")
		dp := GuiControlGet("abr", "POS", "abrFSDvon")

	;-: Datum bis
		Gui, abr: Font  	, % "s" FSize " q5 Normal c" TextColor1
		Gui, abr: Add	, Text	, % "x" dp.X+dp.W+10 " y" cp.Y                                                                                       	, bis
		Gui, abr: Font  	, % "s" FSize " q5 Normal cBlack"
		Gui, abr: Add	, Edit  	, % "x+3 y" dp.Y " w" FSize*9 "  vabrFSDbis HWNDabrHFSDbis          	gabrHandler"	  	, % ""
		AddTooltip(abrHFSDbis, "Tragen hier das Datum bis zu dem Du suchen möchtest. Du kannst hier ebenso ein`n"
											. "Datum, sowie ein Quartal eintragen. Lass das Feld frei wenn bis zum letzten eingetragen`n "
											. "Tag gesucht werden soll.`n`nTip: leere beide Felder und die Suche findet nur im aktuellen Quartal statt.")

	;-: Alter
		Gui, abr: Font  	, % "s" FSize " q5 Normal c" TextColor1
		Gui, abr: Add	, Text	, % "x" dp.X+dp.W+10 " y" cp.Y                                                                                       	, Alter
		Gui, abr: Font  	, % "s" FSize " q5 Normal cBlack"
		;~ Gui, abr: Add	, Edit  	, % "x+3 y" dp.Y " w" FSize*9 "  vabrFSDbis HWNDabrHFSDbis          	gabrHandler"	  	, % ""
		;~ AddTooltip(abrHFSDbis, "Tragen hier das Datum bis zu dem Du suchen möchtest. Du kannst hier ebenso ein`n"
											;~ . "Datum, sowie ein Quartal eintragen. Lass das Feld frei wenn bis zum letzten eingetragen`n "
											;~ . "Tag gesucht werden soll.`n`nTip: leere beide Felder und die Suche findet nur im aktuellen Quartal statt.")

	;-: Karteikartenkürzel
		BefKzl := ""
		For KZL, data in Addendum.BefKuerzel
			Befkzl .= KZL "|"

     ;-: 1
		Gui, abr: Font  	, % "s" FSize-2 " q5 Normal c" TextColor1
		Gui, abr: Add	, Text	, % "xm+5 y+10"                                                                                                            	, % "Kürzel"
		Gui, abr: Font  	, % "s" FSize-1 " q5 Normal cBlack"
		Gui, abr: Add 	, ComboBox, % "xm+5 y+3 w" 7*(FSize-1) " r10 vabrFSKzl1 gabrHandler"                                     	, % RTrim(BefKzl, "|")
		cp := GuiControlGet("abr", "POS", "abrFSKzl1")
		Gui, abr: Add	, Edit  	, % "x+2 y" cp.Y " w550 h" cp.H "                       	vabrFStE1              	gabrHandler"		, % ""

		Gui, abr: Add 	, ComboBox, % "xm+5 y+3 w" 7*(FSize-1) " r10 vabrFSKzl2 gabrHandler"                                     	, % RTrim(BefKzl, "|")
		cp := GuiControlGet("abr", "POS", "abrFSKzl2")
		Gui, abr: Add	, Edit  	, % "x+2 y" cp.Y " w550 h" cp.H "                       	vabrFStE2              	gabrHandler"		, % ""

		Gui, abr: Add 	, ComboBox, % "xm+5 y+3 w" 7*(FSize-1) " r10 vabrFSKzl3 gabrHandler"                                     	, % RTrim(BefKzl, "|")
		cp := GuiControlGet("abr", "POS", "abrFSKzl3")
		Gui, abr: Add	, Edit  	, % "x+2 y" cp.Y " w550 h" cp.H "                       	vabrFStE3              	gabrHandler"		, % ""



	;}

	;-: TAB 5 - Einstellungen                                                         	;{
		Gui, abr: Tab  	, 5
		Gui, abr: Tab
	;}

	;-: GUI - Position anpassen                                                     	;{
		abrP:= GetWindowSpot(habr)
		aP 	:= GetWindowSpot(AlbisWinID())
		GuiControlGet, cp, abr: Pos, abrLV
		DeltaH := abrP.H - abrP.CH
		GuiW   := cpX + cpW + 5

	  ; Progressbalken
		GuiControl, abr: Move, abrPS1 	, % "x4 w"  GuiW - 5
		GuiControl, abr: Move, abrTabs	, % "w" GuiW - 5

		WinSet, Style 	, -0x1     	, % "ahk_id " abrHPS1
		WinSet, ExStyle	, -0x20000	, % "ahk_id " abrHPS1
		WinSet, Style 	, -0x1     	, % "ahk_id " abrHPS2
		WinSet, ExStyle	, -0x20000	, % "ahk_id " abrHPS2

		GuiControlGet, cp, abr: Pos, abrChrOut
		GuiControlGet, dp, abr: Pos, abrTabs
		GuiControl, abr: Move, abrChrOut, % "h" cpH-DeltaH
		GuiControl, abr: Move, abrLvBT  , % "w645 h690"

		Control, % (options.1 ? "Check":"UnCheck"),, Button9 	, % "ahk_id " habr
		Control, % (options.2 ? "Check":"UnCheck"),, Button10	, % "ahk_id " habr
		Control, % (options.3 ? "Check":"UnCheck"),, Button11	, % "ahk_id " habr

		Gui, abr: Show, % winpos , Abrechnungsassistent
		Gui, abr: Show, % " w" GuiW " h" ap.H

     ; Backup der Abrechnungsfensterposition
		wpAbr := GetWindowSpot(habr)
		SciTEOutput("w" wpAbr.w " h" wpAbr.H ", " ap.H)

		SplashTextOff
	;}

	;-: Letzte Aufteilung Albis|Abrechnungsassistent eintrichten       	;{
		If (Addendum.abrOnPos ~= "i)(R|L)")
			abrGui_WindowTiling(Addendum.abrOnPos)
	;}

	#IfWinActive Abrechnungsassistent ahk_class AutoHotkeyGUI
		Hotkey, ~Enter       	, abrHotkeyHandler
		Hotkey, ~NumPadEnter 	, abrHotkeyHandler
		Hotkey, ~Up          	, abrHotkeyHandler
		Hotkey, ~Down       	, abrHotkeyHandler
	#If

	; umschalten auf anderen Tab
	;~ SendMessage, 0x1330, 3 ,,, % "ahk_id " abrHTab

return ;}

abrHotkeyHandler:                     	;{

	If !WinActive("Abrechnungsassistent ahk_class AutoHotkeyGUI")
		return

	Gui, abr: ListView, abrLV
	selectedRow := LV_GetNext(0, "F")
	LV_GetText(PatID    	, selectedRow, 1)
	LV_GetText(PatName	, selectedRow, 2)

	Switch A_ThisHotkey {

		 Case "Enter":
			AlbisActivate(2)
			AlbisAkteOeffnen(PatName, PatID)

		Case "Down":
			ControlSend,, {Down}	, % "ahk_id " abrhLV

		Case "Up":
			ControlSend,, {Up}   	, % "ahk_id " abrhLV

	}

return ;}

abrLVHandler:                            	;{

		Critical

		If (RegExMatch(A_GuiEvent, "(K|f|C)") || A_EventInfo = 0)
			return

	; ausgewählten Patient ermitteln
		Gui, abr: ListView, abrLV
		LV_GetText(tmpPatID   		, A_EventInfo, 1)
		LV_GetText(tmpPatName	, A_EventInfo, 2)

		;~ SciTEOutput(tmpPatID " / " tmpPatName)

		If (A_GuiEvent <> "A")
			If (!tmpPatID && !tmpPatName) || (lvLastPatID = tmpPatID) || !IsObject(VSEmpf[tmpPatID])
				return

		lvLastPatID :=PatID := tmpPatID
		lvLastPatName := PatName := tmpPatName

	; Karteikarte nach Doppelklick anzeigen
		If RegExMatch(A_GuiEvent, "i)(DoubleClick|A)")
			AlbisAkteOeffnen(PatName, PatID)

	; Namen und Features setzen
		GuiControl, abr:, abrPat, % "[" PatID "] " PatName
		For feeName, X in abrCheckbox              	{
			GuiControl, abr: Hide  	, % "abr_" feeName
			GuiControl, abr:          	, % "abr_" feeName    	, 0
			GuiControl, abr:          	, % "abr_Date" feeName, % ""
		}

		For feeIDX, fee in VSEmpf[PatID].Vorschlag 	{
			RegExMatch(fee, "i)^(?<Name>[a-z]+)\-*(?<PossDate>\d+)*", fee)
			feePossDate := ConvertDBASEDate(feePossDate)
			GuiControl, abr: Show   	, % "abr_" feeName
			GuiControl, abr:         	, % "abr_" feeName     	, 1
			GuiControl, abr:         	, % "abr_Date" feeName, % "> " feePossDate
		}

	; Behandlungstage des Patienten anzeigen
		Behandlungstage(PatID, VSEmpf[PatID].BehTage)

		;~ SciTEOutput(cJSON.Dump(VSEmpf[PatID].BehTage, 1))

return ;}

abrICDLVHandler:                     	;{

	Critical

	If (RegExMatch(A_GuiEvent, "(K|f|C)") || A_EventInfo = 0)
		return

  ; ausgewählten Patient ermitteln
	Gui, abr: ListView, abrICDLV
	LV_GetText(ICDPatID   		, A_EventInfo, 1)
	LV_GetText(ICDPatName	, A_EventInfo, 2)

  ; Karteikarte nach Doppelklick anzeigen
	If RegExMatch(A_GuiEvent, "i)(DoubleClick|A)")
		AlbisAkteOeffnen(ICDPatName, ICDPatID)

return
;}

abrHandler:                              	;{

  ; Gui-Abfrage (ohne Critical funktioniert es nicht!)
	Critical

  ; Gui Steuerelementvariablen abfragen
	Gui, abr: Default
	Gui, abr: Submit, NoHide

  ; unter anderen Variablenbezeichnungen hinterlegen
	Abrechnungsquartal := abrQ
	Q  	:= SubStr(abrQ, 1, 1)
	YY	:= SubStr(abrQ, 2, 2) + 0

  ; alles nullen
	;~ GuiControl, abr:, EBMEuroPlus	, % "00000.00 Euro"
	;~ GuiControl, abr:, EBMScheine	, % ""

	If   	  RegExMatch(A_GuiControl, "(abrQUp|abrQDown)")                             	{      	; Quartal - vor und zurück

		If !LVClear
			LVClear := abrGui_LVClear()

	  ; Abrechnungsquartal berechnen
		Q := Q + (QUD := A_GuiControl="abrQUp" ? 1 : -1)
		If (Q = 0)
			Q := 4, YY += -1*QUD
		else if (Q = 5)
			Q := 1, YY += QUD
		YY := SubStr( "0" (YY > 99 ? 0 : YY < 0 ? 99 : YY), -1)

	  ; Abrechnungsquartal stellen
		GuiControl, abr:, abrQ, % (Abrechnungsquartal := Q YY)

	  ; Kalenderanzeige ändern
		GuiControl, abr:, abrCal, % Addendum.ThousandYear . YY . SubStr( "0" ((Q+2)*3)+1, -1) 	. "01"
		GuiControl, abr:, abrCal, % Addendum.ThousandYear . YY . SubStr( "0" ((Q-1)*3)+1, -1) 	. "01"

	  ; Daten nullen
		GuiControl, abr:, EBMEuroPlus	, % "00000.00 Euro"
		GuiControl, abr:, EBMScheine	, % ""

	  ; Nutzerauswahl anpassen (als ob man einen Nutzer auswählen könnte - aber bitte schön: schlaustes Marketingsprech!)
		If FileExist(Addendum.DBasePath "\Q" YY Q "\VorsorgeKandidaten_" Q YY ".json") {
			GuiControl, abr: Enable         	, abrRefresh
			GuiControl, abr: Enable        	, abrReIndex
			GuiControl, abr:, abrRefresh   	, % "Liste laden"
			GuiControl, abr:, abrReIndex 	, % "Liste erneuern"
		}
		else {
			GuiControl, abr: Disable        	, abrRefresh
			GuiControl, abr: Enable        	, abrReIndex
			GuiControl, abr:, abrRefresh   	, % "noch keine Daten"
			GuiControl, abr:, abrReIndex 	, % "Liste erstellen"
		}

	}
	else if RegExMatch(A_GuiControl, "i)(abrReIndex|abrRefresh)")                     	{      	; Abrechnungsquartal - Daten zusammenstellen, speichern und anzeigen

	  ; Listview-Steuerelemente zurücksetzen
		LVClear := abrGui_LVClear()

	; der Nutzer hat seine Wahl getroffen: Auffrischen (Laden) oder die Daten neu erstellen
		ReIndex := A_GuiControl = "abrReIndex" ? true : false
		SplashText("Vorbereitung", "Abrechnungsdaten werden " 	(ReIndex ? "neu erstellt ...":"geladen ..."))
		abrEBMPlus()
		candidates := VorsorgeKomplexe(Abrechnungsquartal, ReIndex, true)

	; entfernt alle Daten aus VSEmpf und ersetzt diese durch neue Daten (VSEmpf := Vorsorge.... führt zum Verlust der globalen Eigenschaft, das Objekt ist nicht vorhanden)
		ObjectReplace(candidates)

	; Auswahlmöglichkeiten ändern
		GuiControl, abr: Enable     	, abrReIndex
		GuiControl, abr: Enable      	, abrRefresh
		GuiControl, abr:, abrRefresh	, % "Liste laden"
		GuiControl, abr:, abrReIndex	, % "Liste erneuern"

	; Abrechnungsvorschläge jetzt anzeigen
		SplashText("neues Quartal", "Abrechnungsdaten des Quartal " Q "/" YY " wurden "  (A_GuiControl = "abrReIndex" ? "reindiziert" : "aufgefrischt"), 6)
		If Addendum.Debug
			SciTEOutput("Abrechnungsdaten des Quartal " Q "/" YY " wurden "  (A_GuiControl = "abrReIndex" ? "reindiziert" : "aufgefrischt"))
		gosub abrVorschlaege
		abrEBMPlus(Round(EBMPlus, 2))

	; aktuelles Abrechnungsquartal in der Addendum.ini sichern
		SaveAbrechnungsquartal(Abrechnungsquartal)

	}
	else if (A_GuiControl = "abrBearbeitet")                                           	{     	; bearbeitete Einträge - verstecken

			Gui, abr: Default
			Gui, abr: ListView, abrLV

		; alle Abgehakten registrieren
			row := rowF := 0, RowsToDelete := []
			while (row := LV_GetNext(row, "C")) {
				rowF := !rowF ? row : rowF
				LV_GetText(PatID, row, 1)
				LV_GetText(PatName, row, 2)
				RowsToDelete.InsertAt(1, {"deleteRow":row, "hideID":PatID, "hideName":PatName})
			}

		; Markieren und Löschen (von unten nach oben))
			;GuiControl, -Redraw, % abrhLV
			Gui, abr: Default
			For idx, to in RowsToDelete {
				VSEmpf[to.hideID].Hide := true
				Gui, abr: ListView, abrLV
				LV_Delete(to.deleteRow)
			}

	  ; Karteikarte schließen bringt Geschwindigkeitsvorteile
			AlbisActivate(1)
			For each, to in RowsToDelete {
				AlbisMDIChildWindowClose(to.hideName)
				while (AlbisMDIChildHandle(to.hideName) && A_Index<30)
					Sleep 10
			}


		; verändertes Objekt speichern
			Gui, abr: Default
			Gui, abr: Submit, NoHide
			FileSize := DBAccess.SaveVorsorgeKandidaten(VSEmpf, abrQ)

		; Färbung der Zellen auffrischen
			Loop % LV_GetCount() {
				row := A_Index
				Loop 5 {
					LV_GetText(CellText, row, 4+A_Index)
					CLV.Cell(row, 4+A_Index, "0x" (CellText ? tcolorM : tcolor), 0x000000)
				}
			}

			;GuiControl, +Redraw, % abrhLV
			;WinSet, Redraw,, % "ahk_id " habr

		;-: Zahl der angezeigten Patienten
			Sleep 200
			abrEBMPlus(Round(EBMPlus, 2))

	}
	else if (A_GuiControl = "abrUnHide")                                               	{	     	; versteckte Einträge - wieder anzeigen

			fn_msbx := Func("MsgBoxMove").Bind("Frage", "Wollen Sie wirklich die ausgeblendeten Fälle")
			SetTimer, % fn_msbx, -100
			MsgBox, 0x1004, Frage, % "Wollen Sie wirklich die ausgeblendeten Fälle wieder sehen?"
			IfMsgBox, No
				return

		; Objektflags auf false setzen
			For PatID, vs in VSEmpf
				If VSEmpf[PatID].Hide
					VSEmpf[PatID].Hide := false

		; verändertes Objekt speichern
			FSize := DBAccess.SaveVorsorgeKandidaten(VSEmpf, abrQuartal)

		; Inhalt neu anzeigen über universal Label
			gosub abrVorschlaege

	}
	else if RegExMatch(A_GuiControl, "(abrYoungStars|abrAngels|abrOldStars)")          	{      	; Filterhandler
		gosub abrVorschlaege
	}
	else if (A_GuiControl = "abrICDKrauler")                                           	{      	; sucht neue Pat. für die Abrechnung der Chronikerpauschalen
		Gui, abr: Submit, NoHide
		chrPausch := ICDKrauler(abrQ)
		For PatID, Pauschale in chrPausch
			LV_Add("", PatID, PatDB[PatID].Name ", " PatDB[PatID].Vorname, ConvertDBASEDate(PatDB[PatID].GEBURT), Pauschale)

	}

return ;}

abrVorschlaege:                        	;{

		; global abrhLV, abrLVm, CLV

		SplashText(Skriptname " .... bitte warten ...", "gefilterte Daten werden aufgelistet ...")

		Gui, abr: Default
		Gui, abr: ListView, abrLV
		Gui, abr: Submit, NoHide

		GuiControlGet, ShowYoungStars	, abr:, % "abrYoungStars"
		GuiControlGet, ShowAngels      	, abr:, % "abrAngels"
		GuiControlGet, ShowOldStars      	, abr:, % "abrOldStars"

		CLV.Clear(1, 1)
		LV_Delete()
		GuiControl, -Redraw, % abrhLV

		EBMplus := 0
		If Addendum.Debug
			SciTEOutput("`nAbrechnungsquartal " SubStr(Abrechnungsquartal, 1, 1) "/" SubStr(Abrechnungsquartal, -1))
		For PatID, vs in VSEmpf {

		  ; nicht anzeigen wenn eine Bedingung erfüllt ist
			If !RegExMatch(PatID, "^\d+$") || VSEmpf[PatID].Hide || (!ShowYoungStars && vs.Alter < 35)
				|| (!ShowOldStars && vs.Alter > 34) || (!ShowAngels && PatDB[PatID].Mortal)
				continue

		  ; #fehlerhaft einsortierte Patienten aus einem anderen Quartal nicht anzeigen (haben aus irgendeinem Grund auch eine nicht passende PatID)
			If IsObject(vs.BehTage) {

				qmismatch :=  false
				For Tag, Behandlung in vs.BehTage {
					If RegExMatch(Tag, "[\d\.]+") && (GetQuartalEx(Tag, "QYY") != Abrechnungsquartal) {
						qmismatch := true
						break
					}
				}
				If qmismatch
					continue

			}

		  ; Vorschläge vorschlagen ∶-)   	;{
			e := []
			For idx, exam in vs.Vorschlag 	{       ; Vorsorgeuntersuchungen/Prävention
				e[idx] := false
				RegExMatch(exam, "i)(?<rdination>[A-Z]+)\-*(?<Date>\d+)*", o)
				Switch ordination  	{
					Case "GVU":
						e.1 := true
						EBMplus += Euro.GVU
					Case "HKS":
						e.2 := true
						EBMplus += Euro.HKS2
					Case "KVU":
						e.3 := true
						EBMplus += Euro.KVU
					Case "Colo":
						e.4 := true
						EBMplus += Euro.Colo
					Case "Aorta":
						e.5 := true
						EBMplus += Euro.Aorta
				}
			}
			If vs.CHRONISCHGRUPPE                                 	{     	; Chronikerpauschalen
				EBMPlus += vs.chronisch 	= 1 	? EBM.Chron1 	: 0
				EBMPlus += vs.chronisch 	= 2 	? EBM.Chron2 	: 0
				EBMPlus += vs.chronisch 	= 3 	? EBM.Chron2 + EBM.Chron1 : 0
			}
			If (vs.GERIATRISCHGRUPPE  && vs.Alter >= 55) 	{   	; Geriatrische Basiskomplexe
				EBMPlus += vs.geriatrisch = 1 	? EBM.GB1   	: 0
				EBMPlus += vs.geriatrisch = 2 	? EBM.GB2   	: 0
				EBMPlus += vs.geriatrisch = 3 	? EBM.GB2 + EBM.GB1 : 0
			}
			;PauschGO := vs.Alter
			;~ EBMPlus += vs.Pauschale = 1 	? EBM.PauschGO	: 0
			;~ EBMPlus += vs.Pauschale = 2 	? EBM.PauschZ	: 0
			;~ EBMPlus += vs.Pauschale	= 3 	? EBM.PauschGO + EBM.PauschZ : 0
		  ;}

		    f := (vs.chronisch ? 1 : 0) + (vs.chronischQ ? 1 : 0)
			Loop 5
				f += e[A_Index] ? 1 : 0

			If (f) {

				; Zeile anfügen
					Mortal 	:= PatDB[PatID].Mortal ? "♱" : ""

					If vs.CHRONISCHGRUPPE                              	{     	; Chronikerpauschalen
						Chron	:= vs.chronisch	= 1 	? "03220" : vs.chronisch 	= 2 ? "03221" : vs.chronisch	= 3 ? "03220/1"
						Chron	.= vs.chronischQ		? " [" vs.chronischQ " VQ" (vs.chronischQ = 1 ? " fehlt]" : " fehlen]") : ""
					}
					If (vs.GERIATRISCHGRUPPE  && vs.Alter >= 55) {   	; Geriatrische Basiskomplexe
						GB   	:=  vs.Alter >= 55 && vs.geriatrisch = 1 ? "03360"
									:	  vs.Alter >= 55 && vs.geriatrisch = 2 ? "03362"
									: 	  vs.Alter >= 55 && vs.geriatrisch = 3 ? "03360/2"
									: 	  vs.Alter < 55 ? "" : ""
					}
					Pausch	:= vs.Pauschale     	?            	vs.Pauschale          	: ""

					LV_Add("Checked"
								, PatID                                      ; 1
								, Mortal . vs.Name                 	; 2
								, vs.Alter                                   ; 3
								, vs.Geschl                            	; 4
								, e.1 ? "X":""                              	; 5
								, e.2 ? "X":""                              	; 6
								, e.3 ? "X":""                              	; 7
								, e.4 ? "X":""                              	; 8
								, e.5 ? "X":""                              	; 9
								, Chron                                   	; 10
								, GB                                      	; 11
								, Pausch)                                  	; 12

			}

		}

	; belegte Zellen in der Zeile umfärben
		Loop % LV_GetCount() {
			row := A_Index
			Loop 8 {
				col := 4 + A_Index
				LV_GetText(txt, row, col)
				If !txt
					CLV.Cell(row, col, "0x" tcolorM, 0x000000)
			}
		}

		GuiControl, +Redraw, % abrhLV
		WinSet, Redraw,, % "ahk_id " abrhLV

		SplashTextOff

	;-: Zahl der angezeigten Patienten
		GuiControl, abr:, abrLVRows, % "[" LV_GetCount() " Patienten]"

	;-: Euro Summe aller EBM Gebühren anzeigen (vorausgesetzt sie übernehmen die Vorschläge zu 100%)
		GuiControl, abr:, EBMEuroPlus, % Round(EBMPlus, 2) " €"
		GuiControl, abr: Focus, EBMEuroPlus

	;-: Listview-Status zurücksetzen
		LVClear := 0

return ;}

abr_Tabs:                                  	;{

return ;}

abrOn:                                      	;{
; 🔐🔏


	Switch A_GuiControl {

		Case "abrOnTop":

			Addendum.abrOntop := !Addendum.abrOntop
			GuiControl, abr:, abrOntop, % (Addendum.abrOntop ? "🔐" : "🔏")
			WinSet, AlwaysOnTop, % (Addendum.abrOntop ? "on" : "off"), % "ahk_id " habr
			AddTooltip(abrHOnTop, "AlwaysOnTop ist " (Addendum.abrOnTop ? "eingeschaltet" : "ausgeschaltet"))
			GuiControlGet, otp, abr: Pos, abrOntop
			ToolTip, % "AllwaysOnTop ist " (Addendum.abrOntop ? "eingeschaltet":"ausgeschaltet"), otpX-50, otpY-10, 3

		Case "abrOnPos":

			Addendum.abrOnPos := Addendum.abrOnPos="R" ? "L" : Addendum.abrOnPos="L" ? "F" : "R"
			GuiControl, abr:, abrOnPos, % (Addendum.abrOnPos="R" ? "◨" : Addendum.abrOnPos="L" ? "◧" : "🗗")
			aufteilung := Addendum.abrOnPos="R" ? "Albis | Abrechnungsassistent" : Addendum.abrOnPos="L" ? "Abrechnungsassistent | Albis" : "keine"
			AddTooltip(abrHOnPos,  "Fensteraufteilung: " aufteilung)
			GuiControlGet, otp, abr: Pos, abrOnPos
			ToolTip, % "Fensteraufteilung: " aufteilung, otpX-70, otpY-10, 3
			abrGui_WindowTiling(Addendum.abrOnPos)

	}

	;~ otp := GetWindowSpot(hotip := WinExist("A"))
	;~ onround := 0
	SetTimer, onAbrTimerOff, -3000

return
onAbrTimerOff:
	ToolTip,,,, 3
return
onRoundTimer:
	onround ++
	otp.X -= 1
	otp.Y -= 1
	SetWindowPos(hotip, otp.X, otp.Y, otp.W, otp.H)

	If (onround = 50) {
		SetTimer, onRoundTimer, off
		ToolTip,,,, 3
	}

return ;}

abrReload:
abrGuiClose:
abrGuiEscape:                           	;{

	Gui, abr: Default
	Gui, abr: Submit, NoHide
	SaveAbrechnungsquartal(abrQ)
	SaveGuiPos(habr, "Abrechnungsassistent")

	If (A_GuiControl = "abrReload") {
		Reload
		return
	}

ExitApp ;}
}

abrGui_WindowTiling(Tiling)                                                                                 	{ 		;-- Fensterverteilung (für Monitore mit hoher Auflösung)

	/* Hinweise

		◨		Das Albisfenster wird bei ausgewähltem Modus "R" [Symbol ◨] oder "L" [Symbol ◧] genau neben dem Fenster des Abrechnungsassistenten positioniert.
				Die Größe des Albisfenster wird dem verbliebenen Platz vom Monitorrand bis zum Abrechnungsassistenten und Taskleiste angepasst.

		🗗		Bei Auflösungen in der horizontalen von kleiner/gleich 1920 Pixel ist der Fensteraufteilungsmodus nicht zu empfehlen! Hier ist es besser einen
				zweiten Monitor anzuschließen oder den Modus des Assistenten auf "Immer im Vordergrund (AlwaysOnTop) [Symbol 🔐]".

				letzte Änderung: 02.07.2022

	*/

	global wpAlbis, wpAbr, hAbr

	abrWin  	:= GetWindowSpot(hAbr)
	albisWin 	:= GetWindowSpot(hAlbis := AlbisWinID())
	albisMon	:= ScreenDims(monIndex := GetMonitorIndexFromWindow(hAlbis))
	TBarH   	:= TaskbarHeight(monIndex)

  ; Pixel-Bereich für die Albisbreite (Breite des Monitorauflösung - Breite des Abrechnungsassistentenfensters)
	MonLA		:= albisMon.W - abrWin.CW + 15

	If (Tiling="R") {
		SetWindowPos(hAbr	, albisMon.W-abrWin.W+3, 0, abrWin.W, abrWin.H,, 0x1)
		SetWindowPos(hAlbis, -8, 0, MonLA, albisMon.H-TBarH)
	}
	else If (Tiling="L") {
		SetWindowPos(hAbr	, 0, 0, abrWin.W, abrWin.H,, 0x1)
		SetWindowPos(hAlbis, abrWin.W+1, 0, MonLA, albisMon.H-TBarH)
	}
	else {	; stellt die ursprüngliche Position wieder her
		SetWindowPos(hAbr	, wpAbr.X, wpAbr.Y, wpAbr.W, wpAbr.H,, 0x1)
		SetWindowPos(hAlbis, wpAlbis.X, wpAlbis.Y, wpAlbis.W , wpAlbis.H)
	}

}

abrGui_LVClear()                                                                                              	{

	global BLV, CLV, abr, abrBT, abrBTage, abrCalStart, abrLV, abrLvT

	GuiControl, abr:, abrBT     	, % "--.--.----"
	GuiControl, abr:, abrBTage	, % "--"
	GuiControl, abr: Disable, abrCalStart
	Gui, abr: ListView, abrLV
	CLV.Clear(1, 1)
	LV_Delete()
	Gui, abr: ListView, abrLvBT
	BLV.Clear(1, 1)
	LV_Delete()

return 1
}

abrEBMPlus(Plus:=0)                                                                                          	{  	;-- frischt nur ein paar Steuerelemente auf

	global abr, abrLV, abrLVRows, EBMEuroPlus, Abrechnungsquartal, abrQuartal, VSEmpf, Euro

	Gui, abr: Default
	Gui, abr: ListView, abrLV
	Gui, abr: Submit, NoHide

	GuiControlGet, ShowYoungStars	, abr:, % "abrYoungStars"
	GuiControlGet, ShowAngels      	, abr:, % "abrAngels"
	GuiControlGet, ShowOldStars      	, abr:, % "abrOldStars"

	EBMPlus := 0
	For PatID, vs in VSEmpf {

	  ; wird für die Berechnung verwendet wenn eine Bedingung erfüllt ist
		If !RegExMatch(PatID, "^\d+$") || VSEmpf[PatID].Hide || (!ShowYoungStars && vs.Alter < 35)
			|| (!ShowOldStars && vs.Alter > 34) || (!ShowAngels && PatDB[PatID].Mortal)
			continue

	  ; #fehlerhaft einsortierte Patienten aus einem anderen Quartal nicht anzeigen (haben aus irgendeinem Grund auch eine nicht passende PatID)
		If IsObject(vs.BehTage) {

			qmismatch :=  false
			For Tag, Behandlung in vs.BehTage {
				If Tag && (GetQuartalEx(Tag, "QYY") != Abrechnungsquartal) {
					qmismatch := true
					break
				}
			}
			If qmismatch
				continue

		}

	  ; Vorschläge vorschlagen ∶-)
		e := []
		For idx, exam in vs.Vorschlag 	{       ; Vorsorgeuntersuchungen/Prävention
			e[idx] := false
			RegExMatch(exam, "i)(?<rdination>[A-Z]+)\-*(?<Date>\d+)*", o)
			Switch ordination  	{
				Case "GVU":
					e.1 := true
					EBMplus += Euro.GVU
				Case "HKS":
					e.2 := true
					EBMplus += Euro.HKS2
				Case "KVU":
					e.3 := true
					EBMplus += Euro.KVU
				Case "Colo":
					e.4 := true
					EBMplus += Euro.Colo
				Case "Aorta":
					e.5 := true
					EBMplus += Euro.Aorta
			}
		}
		If vs.CHRONISCHGRUPPE     	{     	; Chronikerpauschalen
			EBMPlus += vs.chronisch 	= 1 	? EBM.Chron1 	: 0
			EBMPlus += vs.chronisch 	= 2 	? EBM.Chron2 	: 0
			EBMPlus += vs.chronisch 	= 3 	? EBM.Chron2 + EBM.Chron1 : 0
		}
		If vs.GERIATRISCHGRUPPE     	{   	; Geriatrische Basiskomplexe
			EBMPlus += vs.geriatrisch = 1 	? EBM.GB1   	: 0
			EBMPlus += vs.geriatrisch = 2 	? EBM.GB2   	: 0
			EBMPlus += vs.geriatrisch = 3 	? EBM.GB2 + EBM.GB1 : 0
		}


	}

	GuiControl, abr:, EBMEuroPlus, % (!EBMPlus ? "00.00" : Round(EBMPlus, 2)) " €"
	GuiControl, abr: Focus, EBMEuroPlus

	Gui, abr: ListView, abrLV
	KS := LV_GetCount()
	GuiControl, abr:, abrLVRows, % "[" (!KS ? "0000" : KS) " Patienten]"

return KS
}


ObjectReplace(NewObject)                                                                                 	{

	; Warum diese Funktion?
	; --------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; 	Wenn ein globales Objekt wie VSEmpf ersetzt wird durch Übergabe eines neuen Inhaltes ‹ VSEmpf := VorsorgeKomplexe("222", 1,1) ›
	; 	oder ‹ VSEmpf := NewObject › wird man beim Auslesen von Objektinhalten außerhalb der aktuellen Funktion feststellen das das Objekt
	;	keinen Inhalt mehr hat. Ich vermute das der bisherge Pointer durch einen neuen Pointer ersetzt wird. Dieser wird aber dem globalen Objekt
	;	nicht übergeben.
	;	Daher muß zwingend jeder einzelne Schlüssel (key) entfernt werden, wenn man sich nicht sicher ist das alle überschrieben werden und


	; NewObject - muss die neuen Daten für VSEmpf (Vorsorgeempfehlungen) enthalten

	global VSEmpf

	objindex := 0
	If Addendum.Debug
		SciTEOutput("entferne Daten von " VSEmpf.Count() " Patienten. Objektkapazität: " VSEmpf.GetCapacity())

  ; VSEmpf leeren
	For PatID, obj in VSEmpf {
		If IsObject(VSEmpf[PatID]) {
			objindex ++
			;~ VSEmpf[PatID].SetCapacity(0)   ; wozu oder besser?
			tmp := VSEmpf.Delete(PatID)
		}
	}

	MaxElements := NewObject.GetCapacity()
	VSEmpf.SetCapacity(MaxElements)

	If Addendum.Debug {
		SciTEOutput("MaxElements in NewObject: " MaxElements )
		SciTEOutput("VSEmpf enthält leer: " VSEmpf.Count() " Schlüssel")
	}

	For key, obj in NewObject
		VSEmpf[key] := obj

	If Addendum.Debug {
		SciTEOutput("VSEmpf enthält neu befüllt: " VSEmpf.Count() " Schlüssel. Es wurden " objindex " Schlüssel entfernt.")
		clipboard := "VSEmpf leer:`n " clip1 "`n-----------------------------------------------------------`n" cJSON.Dump(VSEmpf, 1)
	}

}

QuartalsKalender(Abrechnungsquartal,Urlaubstage,guiname="QKal",TextColor1=""){    	;-- zeigt den Kalender des Quartals

	; Todo: blendet Feiertage nach Bundesland aus/ein (### ist nicht programmiert! ####)

	;--------------------------------------------------------------------------------------------------------------------
	; Variablen
	;--------------------------------------------------------------------------------------------------------------------;{
		global QKal, hmyCal, hCal, habr, gcolor, tcolor, CalQK, abrQ
		global Cok, Can, Hinweis, abrCal, CalBorder, abrCalDay, abrCalStart, abrCalBTage
		global abrCheckbox, VSRule, CV, PUpX, PUpY, abrLvBT, abrhLvBT, StaticBT, abrhLV
		global KK, lvLastPatID, lvLastPatName, lvPatID, lvPatname, BLV, abrYear
		global BefKuerzel, DateChoosed, abrDayChoosed, abrHCalStart

		static gname, lastPrgDate, lastBTEvent, lastBTInfo

		If (gname <> guiname)
			gname := guiname
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	Berechnung der Monate
	;--------------------------------------------------------------------------------------------------------------------;{
		aQZ	:= Ceil(A_MM/3)*3
		QZ 	:= SubStr(Abrechnungsquartal, 1, 1)
		YY   	:= SubStr(Abrechnungsquartal, 2, 2)
		; Abrechnungsquartal kann nicht nicht in der Zukunft liegen, also liegt es vor dem aktuellen Quartal
		; (rechnet von 2000er auf 1900er Jahre, wenn man es braucht)
		Jahr := (SubStr(A_YYYY, 3, 2)=YY && QZ > aQZ) || SubStr(A_YYYY, 3, 2)<YY ? "19" YY : "20" YY

		firstmonth	:= ((QZ-1)*3)+1
		lastmonth := FirstMonth+2
		lastdays  	:= DaysInMonth(Jahr . SubStr("0" lastmonth, -1) 	. "01")
		FormatTime, begindate	, % Jahr . SubStr("0" firstmonth, -1) 	. "01"    	, yyyyMMdd
		FormatTime, lastdate 	, % Jahr . SubStr("0" lastmonth, -1)		. lastdays	, yyyyMMdd

	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	Kalender
	;--------------------------------------------------------------------------------------------------------------------;{
		lp 	:= GetWindowSpot(abrhLV)
		ap	:= GetWindowSpot(AlbisWinID())

		GuiControlGet, cp, abr: Pos, CalQK
		Gui, abr: Font     	, % "s11 w600 c" TextColor1, Futura Mk Md
		Gui, abr: Add     	, Text		, % "xm+5 y+1 vCalQK"                             	, % "QUARTALSKALENDER"
		GuiControlGet, cp, abr: Pos, CalQK

		Gui, abr: Font     	, % "s8 Normal italic c" TextColor1, Futura Bk Bt
		Gui, abr: Add     	, Text		, % "x+2 y" cpY+1                                        	, % "[" Quartal SubStr(Jahr, -1) "]"

		Gui, abr: Font     	, % "s12 w600 c" TextColor1, Futura Mk Md
		Gui, abr: Add     	, Text		, % "x+1 y" cpY-4                                         	, % ":"

		Gui, abr: Font     	, % "s10 Bold c" TextColor1, Futura Bk Bt
		Gui, abr: Add     	, Text		, % "x+8 y" cpY+1                                        	, % SubStr(ConvertDBASEDate(begindate), 1,6) "-" ConvertDBASEDate(lastdate)

		CalOptions := " AltSubmit +0x50010051 vabrCal HWNDhmyCal gCalHandler"
		Gui, abr: Add     	, MonthCal, % "xm+4 y+2 W-3 4 w521 "  CalOptions   	, % begindate
		GuiControlGet, cp, abr: Pos, abrCal
		cX 	:= cpX + cpW + 5
		cY		:= cpY + cpH

		Gui, abr: Font     	, % "s10 Normal c" TextColor1
		Gui, abr: Add     	, Text		, % "x" cX " y" cpY  " BackgroundTrans"           	, % "gewählter Tag: "
		Gui, abr: Font     	, % "s11 Bold c" TextColor1
		Gui, abr: Add     	, Text		, % "x" cX " y+5 vabrDayChoosed"                 	, % "  --.--.----  "
		Gui, abr: Font     	, % "s10 Normal c" TextColor1
		opt := "hwndabrHCalStart gCalHandler"
		Gui, abr: Add     	, Button		, % "x" cX " y+5 vabrCalStart " opt                	, % "Ziffernauswahl`neintragen"
		GuiControl, abr: Disable, abrCalStart

		gcolor              	:= "0x00333333"    	; Hintergrund
		tcolor               	:= "0x00BCCCE4"		; Listview
		tcolorD           	:= "0x003B8600"		; Markierung
		tcolorP            	:= "0x00333333"		; Unterteilung
		tcolorM            	:= "0x00DADEE7"   	; Vorschläge
		tcolorI            	:= "0x005ED000"  	; Symbole 3399FF
		tcolorBG       	:= "0x00D0C1C1"
		tcolorBG1       	:= "0x000078D7"		; Zeilenauswahl im Listview
		tcolorBG2      	:= "0x00F8F5F5"
		tcolorBG2       	:= "0x003399FF"
		TextColor1     	:= "0x0065DF00"
		SelectionColor1	:= "0x000078D6"

	; Kalenderfarben
		SendMessage, 0x100A	, 0, % gcolor                      	,, % "ahk_id " hmyCal   	; MCSC_BACKGROUND 	background color displayed between months.
		SendMessage, 0x100A	, 1, % tcolorBG                     	,, % "ahk_id " hmyCal   	; MCSC_MONTHBK   		background color displayed within the month.
		SendMessage, 0x100A	, 2, % TextColor1                 	,, % "ahk_id " hmyCal   	; MCSC_TEXT 					color used to display text within a month.
		SendMessage, 0x100A	, 3, % gcolor                     	,, % "ahk_id " hmyCal   	; MCSC_TITLEBK       		background color displayed in the calendar's title.
		SendMessage, 0x100A	, 4, % tcolorI                      	,, % "ahk_id " hmyCal   	; MCSC_TITLETEXT       		color used to display text within the calendar's title.
		SendMessage, 0x100A	, 5, % SelectionColor1			,, % "ahk_id " hmyCal   	; MCSC_TRAILINGTEXT		color used to display header day and trailing day text.
		SendMessage, 0x100A	, 5, 0xFFAA99, SysMonthCal321, % "ahk_id " habr   	; MCSC_TRAILINGTEXT		color used to display header day and trailing day text.
															; 	Header and trailing days are the days from the previous and following months that appear on the current month calendar.

		WinSet, Style, 0x50010051, % "ahk_id " hmyCal

	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	Behandlungstage
	;--------------------------------------------------------------------------------------------------------------------;{

		Gui, abr: Font     	, % "s9 Normal c" TextColor1, Futura Mk Md
		Gui, abr: Add     	, Text  		, % "x" cX " y" cpY  "BackGroundTrans vStaticBT"                                    	, % "Behandlungstage: "
		GuiControlGet, dp, abr: Pos, staticBT

		Gui, abr: Font     	, % "s10 Normal c" TextColor1, Futura Bk Bd
		Gui, abr: Add    	, Text      	, % "x" dpx " y+2 vabrBTage"                                                              	, % "          ------         "
		;~ GuiControl, abr: Move, staticBT, % "x" (nx := 5+lp.W-cpW-30)

		GuiControlGet, cp, abr: Pos, abrCal
		LVOptions 	:= " -Hdr vabrLvBT gCalHandler HWNDabrhLvBT -E0x200 -LV0x10 AltSubmit -Multi"
		LVColumns	:= "Datum|Kürzel|Inhalt"
		Gui, abr: Font     	, % "s10 Normal cBlack", Futura Bk Bt
		;~ Gui, abr: Add     	, ListView  	, % "xm+5 y" cY+2 " w" lp.W " h" (ap.H)-(cpY-cpH) " "  LVOptions      	, % LVColumns
		Gui, abr: Add     	, ListView  	, % "xm+5 y" cY+2 " w968 h1060 "  LVOptions      	, % LVColumns

		SciTEOutput("w: " lp.W " h" (ap.CH)-(cpY-cpH))

		LV_ModifyCol(1, 50)
		LV_ModifyCol(2, 65)
		LV_ModifyCol(3, 500)
		WinSet, Style, 0x5021800D, % "ahk_id " abrhLvBT   ; No_HScroll

		BLV := new LV_Colors(abrhLvBT)
		BLV.Critical := 200
		BLV.SelectionColors("0x0078D6")

		Redraw(habr)
		UpdateWindow(habr)

	;}

return

CalHandler: ;{

		Critical
		If (lastBTEvent = A_GuiEvent && lastBTInfo = A_EventInfo)
			return
		lastBTEvent := A_GuiEvent, lastBTInfo := A_EventInfo

		Gui, abr: Default
		Gui, abr: Submit, NoHide
		GuiControlGet, daychoosed, abr:, abrDayChoosed

	;---------------------------------------------------------------------
	; Eintragen starten
	;---------------------------------------------------------------------
		If (A_GuiControl = "abrCalStart")  {
			fn_msbx := Func("MsgBoxMove").Bind("Frage", "ausgewähltes Datum")
			SetTimer, % fn_msbx, -100
			MsgBox, 0x1004, Frage, % "ausgewähltes Datum: " daychoosed "`nZiffernauswahl dort eintragen?"
			IfMsgBox, No
				return
			ZiffernauswahlEintragen(daychoosed, lvLastPatID, lvLastPatName)
			BlockInput, Off
			HandsOff(false)
		}

	;---------------------------------------------------------------------
	; angeklicketes Datum sichern
	;---------------------------------------------------------------------
		else If (A_GuiControl = "abrCal") {
			If RegExMatch(A_GuiEvent,"i)(Normal|I|1)") {
				Gui, abr: Submit, NoHide
				DateChoosed 	:= ConvertDBASEDate(StrSplit(abrCal, "-").1)
				GuiControl, abr:, abrDayChoosed, % DateChoosed
				GuiControl, abr: Enable, abrCalStart
			}
			return
		}

	;---------------------------------------------------------------------
	; alternativ das Datum der angeklickten Zeile aus der Karteikarte
	;---------------------------------------------------------------------
		else If (A_GuiControl = "abrLvBT") {

			If (A_GuiEvent = "I") {
				Gui, abr: Default
				Gui, abr: ListView, abrLvBT

				sDate := "", sRow := A_EventInfo + 1
				while (!sDate && sRow>0)
					LV_GetText(sDate, (sRow := sRow - 1), 1)

				If sDate {
					GuiControl, abr:      	, abrDayChoosed	, % (DateChoosed := sDate . Addendum.ThousandYear . SubStr(abrQ, 2, 2))
					GuiControl, abr:      	, abrCal      	, % ConvertToDBASEDate(DateChoosed)
					GuiControl, abr: Enable	, abrCalStart
				}
			}

		}

return ;}
}

Behandlungstage(PatID, BehTage)                                                                        	{		;-- Karteikartensimulation

		global abr, abrLV, abrLvBT, abrhLvBT, abrBTage, BLV, BefKuerzel, tcolorBG1, tcolorBG2, abrYear
		global VSEmpf
		cols := [], altFlag := true

		Gui, abr:  Default
		Gui, abr: ListView, abrLvBT

		GuiControl, abr:, abrDayChoosed, % "--.--.----"
		GuiControl, abr:, abrBTage, % BehTage.Count()
		BLV.Clear(1, 1)
		LV_Delete()

		;~ For DATUM, Leistung in BehTage {
			;~ abrYear := SubStr(ConvertDBASEDate(DATUM), 7, 4)
			;~ break
		;~ }

		For DATUM, Leistung in BehTage {

		  ; die Kürzel in diesselbe Reihenfolge wie in der Albiskarteikarte bringen
			KKOrder := []
			For row, col in Leistung
				If !IsObject(KKOrder[col.ORD])
					KKOrder[col.ORD] := [col]
				else
					KKOrder[col.ORD].Push(col)

		  ; sortierte Daten ausgeben
			BTRow :=	0
			For order, rowdata in KKOrder
				For each, col in rowdata {

					If (col.KZL = "lle")
						continue

					BTRow ++
					cols.1   	:= BTRow = 1 ? SubStr(ConvertDBASEDate(DATUM), 1, 6) : ""
					cols.2   	:= col.KZL
					cols.3   	:= col.INH
					bgColor 	:= "0x" (altFlag ? tcolorBG1 : tcolorBG2)
					txtColor	:= RGBtoBGR(Format("0x{:06X}", BefKuerzel[col.KZL].Color))
					LV_Add("", cols*)
					BLV.Row(LV_GetCount()   	, bgColor, 0x000000)
					BLV.Cell(LV_GetCount(), 2	, bgColor, txtColor)
					BLV.Cell(LV_GetCount(), 3	, bgColor, txtColor)

				}

			altFlag := !altFlag

		}

		If (row := LV_GetNext(0, "F"))
			LV_Modify(row, "-Focus -Select") ;, LV_Modify(row, "-Select")

		WinSet, Style, 0x5021800D, % "ahk_id " abrhLvBT   ; No_HScroll

return SubStr(DATUM, 1, 4)  ; Abrechnungsjahr wird zurückgegeben  ;abrYear
}


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Automatisierung
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ZiffernauswahlEintragen(DateChoosed, lvPatID, lvPatName)                                 	{   	;-- Karteikartenautomatisierung

		global KK, VSEmpf, VSRule, abrCheckbox, habr, habrLV
		static AlbisClass := "ahk_class OptoAppClass"
		static ksleep1 := 175, ksleep2 := 300, ksleep3 := 350

		dbg := 1

		SetTimer, BlockInputOff, -30000
		fn_msgMove := Func("MsgBoxMove").bind("Frage", "")
		;~ HandsOff(true, 600, A_ScriptDir "\Einstellungen\handsoff2.png")

		SciTEOutput(" [ - - - - - - " A_Hour ":" A_Min ":" A_Sec "- - - - - - - ]"   )

	; Karteikarte des gewählten Patienten anzeigen                                        	;{
		startT := Clock()
		PraxTT("Die Karteikarte von [" lvPatID "] " lvPatname "`n"
				. (AlbisGetActiveWinTitle()~="i)" lvPatID "\s*\/\s*" lvPatname ? " wird bereits angezeigt!" : " wird geöffnet!"), "3 0")
		If Addendum.Debug
			SciTEOutput("start" dbg++ ": " Round(Clock()-startT) " ms  [Start]")

		PatID := AlbisAktuellePatID()
		If (PatID <> lvPatID) {
			If !AlbisAkteOeffnen(lvPatName, lvPatID) {
				MsgBox, 0x1, % A_ScriptName, % "Die Karteikarte von [" lvPatID "] " lvPatname "`nließ sich nicht öffnen!", 3
				return
			}
			PatID := lvPatID
		}
		If Addendum.Debug
			SciTEOutput("start" dbg++ ": " Round(Clock()-startT) " ms zum Öffnen der Karteikarte")
	;}

	; Patient der aktuellen Karteikarte ist gelistet und hat offene Komplexziffern? 	;{
		PraxTT("Lese Programmdatum und ermittle Aufgaben...", "3 0")
		PrgDate := AlbisLeseProgrammDatum()
		res := AlbisSchliesseProgrammDatum()
		If (!VSEmpf[PatID].Vorschlag.Count() && !VSEmpf[PatID].chronisch) {
			PraxTT(	"Bei Patient [" PatID "] " PatDB[PatID].Name ", " PatDB[PatID].VORNAME
					. 	" fehlt keine Eintragung.`nWähle einen anderen Patienten!", "5 1")
			return 0
		}
	;}

	; gewähltes Datum liegt an einem Wochenende 	                                        	;{
		PraxTT("prüfe Abrechnungsdatum....", "3 0")
		If Addendum.Debug
			SciTEOutput("start" dbg++ ": " Round(Clock()-startT) " ms [Aufgaben ermittelt]")
		If RegExMatch(DayOfWeek(DateChoosed, "short"), "^(Sa|So)") {
			MsgBox, 0x1024, % StrReplace(A_ScriptName, ".ahk")
						, % 	"Das gewählte Datum (" DateChoosed ") fällt auf ein Wochenende.`n"
						.		"Soll dieses Datum trotzdem verwendet werden?"
			IfMsgBox, No
				return 0
		}
	;}

	; Gui Daten                                                                                               	;{
		Gui, abr: Default
		Gui, abr: Submit, NoHide
	;}

	; EBM-Ziffern zusammen stellen , Tagesregeln anwenden                              	;{
		If Addendum.Debug
			SciTEOutput("start:" dbg++ ": " Round(Clock()-startT) " ms [Abrechnungsdatum geprüft]")
		PraxTT("Stelle Abrechnungsziffern zusammen", "3 0")
	    DDateChoosed := ConvertToDBaseDate(DateChoosed)
		EBMFees       	:= GetEBMFees(PatID)
		EBMToAdd     	:= ""
		EBMAddAt      	:= {}
		HKSFormular 	:= false

		If VSEmpf[PatID].CHRONISCHGRUPPE
			EBMChrn 	:= VSEmpf[PatID].chronisch = 1	? "03220-" : VSEmpf[PatID].chronisch = 2 ? "03221-" : VSEmpf[PatID].chronisch = 3 ? "03220-"

		If VSEmpf[PatID].GERIATRISCHGRUPPE
			EBMGB 	:= VSEmpf[PatID].geriatrisch = 1	? "03360-" : "03362-"

	  ; Nutzer bei doppelten Ansatz des GB Komplexes im Quartal warnen
		rx := StrReplace(EBMChrn, "-", "|")
		For itemNr, item in VSEmpf[PatID].BehTage[DDateChoosed]
			If (item.KZL = "lko") && RegExMatch(item.INH, "(" EBMChrn ")", lko) {
					examdays := []
					For examDate, examination in VSEmpf[PatID].BehTage
						examdays.Push[examDate]

					GEBUEHRwarner("GB", lko1, DateChoosed, examdays)
			}

		rxEBMGB := StrReplace(EBMChrn, "-", "|")
		EBMToAdd .= EBMChrn EBMGB

		If (VSEmpf[PatID].Vorschlag.Count() > 0)
			For fee, X in abrCheckbox {
				GuiControlGet, ischecked, abr:, % "abr_" fee
				EBMNr := ischecked ? EBMFees[fee] "-" : ""
				If RegExMatch(EBMNr, "(01745|01746)") || RegExMatch(fee, "HKS")
					HKSFormular := true
				EBMToAdd .= EBMNr
			}

		If !(EBMtoAdd := RTrim(RegExReplace(EBMtoAdd, "[\-]{2,}", "-"), "-")) {
				BlockInput, Off
				MsgBox, EMBtoAdd  ist leer!
		}
		If Addendum.Debug
			SciTEOutput("start" dbg++ ": "  Round(Clock()-startT) " ms [EBM Ziffern zusammen gestellt]")

	;}

	; Ziffern eintragen                                                                                        	;{
		SciTEOutput("beginne Eintragungen")
		; aktuelles Programmdatum lesen

			startT := Clock()

		   ; das gewählte Abrechnungsdatum eintragen (Dialog wird automatisch geschlossen)
			lastPrgDate := AlbisSetzeProgrammDatum(DateChoosed)
			Sleep % ksleep1
			res := AlbisKarteikarteAktivieren()
			Sleep % ksleep1
			If Addendum.Debug
				SciTEOutput("start" A_LineNumber ": "  Round(Clock()-startT) " ms [ausgelesen Programmdatum (" lastPrgDate ") und Aktivierung der Karteikarte (" res ")")
		; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
		; Karteikarte zeigen, für Eingaben aktivieren und den Tastaturfokus setzen
		; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
			If !(res := AlbisGetFocus("Karteikarte",, Kkarte))
				If !(res := AlbisKarteikarteAktivieren()) {
					BlockInput, Off
					HandsOff(false)
					MsgBox, 0x1000, Fehler, Die Karteikarte konnte nicht angezeigt werden!
					BlockInput, On
					HandsOff(true)
				}	else
					Sleep ksleep3

			;~ AlbisActivate(1)
			hInput := AlbisKarteikarteEingabe("Kürzel")
			AlbisGetFocus("Karteikarte",,Kkarte)
			If (Kkarte.subFocus != "Kürzel") {
				PraxTT("Fehler!`nDie Eingabefelder der Karteikarte konnten nicht erstellt werden!`nBitte jetzt eine Zeile manuell öffnen, setzen Sie das Caret ins Kürzelfeld und drücken Sie anschließend auf q.", "20 1")
				KeyWait, q, D
				SendInput, {BS}
				PraxTT("", "off")
			} 	else
				Sleep % ksleep3

			res := AlbisGetFocus("Karteikarte",, Kkarte)
			If Addendum.Debug
				SciTEOutput("start" A_LineNumber ": "  (stepB := Clock()-startT) " ms ["  Kkarte.subFocus "]")

		; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
		; Zeilendatum überprüfen
		; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
			ZDatum := AlbisZeilenDatumLesen(80, false)
			If (ZDatum <> DateChoosed) {

				res := AlbisKarteikarteEingabe("Datum")
				AlbisGetFocus("Karteikarte",,Kkarte)

				SciTEOutput("Zeilendatum stimmt nicht überein (" ZDatum "<>" DateChoosed ")")
				SciTEOutput("subfocus: " Kkarte.subFocus)

				If !AlbisZeilendatumSetzen(DateChoosed)  {
					AlbisGetFocus("Karteikarte",,Kkarte)
					If (Kkarte.subFocus <> "Datum")
						res := AlbisKarteikarteEingabe("Datum")
					SendInput, {Home}                                     	; an den Anfang
					SendInput, {LShift Down}{End}{Shift Up}   	; alles markieren
					Sleep % ksleep2
					Send, % "{Raw}" DateChoosed
					Sleep % ksleep3
					SendInput, {TAB}
					Sleep % ksleep2
				}

				If Addendum.Debug
		     		SciTEOutput("start" A_LineNumber ": "  Round(Clock()-startT) " ms")

				hInput := AlbisKarteikarteEingabe("Kuerzel")
			}

		; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
		; lko
		; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺;{
			;~ AlbisActivate(1)
			AlbisGetFocus("Karteikarte",,Kkarte)

			If (Kkarte.subFocus = "Kürzel")
				If !VerifiedSetText("", "lko", "ahk_id " Kkarte.hCaret) {
					For kbindex, kbkey in StrSplit("lko") {
						SendInput, % kbkey
						Sleep 75
					}
					Sleep % ksleep1
					SendInput, {TAB}
					Sleep % ksleep1
				}
			AlbisGetFocus("Karteikarte",, Kkarte)
			If Addendum.Debug
				SciTEOutput("start" A_LineNumber ": " (stepA := Clock()-startT) " ms (" stepA-stepB " ms) [sub: " Kkarte.subFocus "]")
			Sleep % ksleep1
		;}

		; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
		; und zum nächsten Feld vorrücken [INHALT]
		; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺;{
			;~ AlbisActivate(1)
			AlbisGetFocus("Karteikarte",, Kkarte)
			If (Kkarte.subFocus <> "Inhalt") {
				SendInput, {TAB}
				Sleep % ksleep2
				AlbisGetFocus("Karteikarte",, Kkarte)
				If (Kkarte.subFocus <> "Inhalt")
					If !(hInput := AlbisKarteikarteEingabe("Inhalt")) {
						SendInput, {TAB}
						Sleep % ksleep2
						while !AlbisGetFocus("Karteikarte", "Inhalt") && (A_Index < 5)
							Sleep % ksleep1
					}
			}
			If !AlbisGetFocus("Karteikarte", "Inhalt", Kkarte)
				If (Kkarte.subFocus = "Kürzel")
					SendInput, {TAB}
					Sleep % ksleep2

			AlbisGetFocus("Karteikarte", "", Kkarte)
			If Addendum.Debug
				SciTEOutput("start" A_LineNumber ": " (stepB := Clock()-startT) " ms (" stepB-stepA " ms) [sub: " Kkarte.subFocus "]")
			If (Kkarte.subFocus <> "Inhalt") {
				BlockInput, Off
				HandsOff(false)
				PraxTT("Fehler!`nBitte versetzen Sie den Cursor für die Eingabe der EBM-Gebühren!`nund drücken Sie anschließend auf die Taste >> q <<.", "20 1")
				KeyWait, q, D
				PraxTT("", "off")
				SendInput, {BS}
				BlockInput, On
			}
			Sleep % ksleep1
			;}

		; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
		; String senden
		; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺;{
			Send, % "{Raw}" EBMtoAdd
			Sleep % ksleep2

			SendInput, {TAB}
			Sleep % ksleep2
			SendInput, {TAB}
			Sleep % ksleep2
			Loop {

				ControlGetFocus, fcs, % AlbisClass
			  ; wenn die Karteikarte Eingaben annehmen kann, wird dies durch Senden von Escape beendet
				If (fcs <> KK.Inhalt) {
					while AlbisKarteiKarteEingabeStatus() {
						SendInput, {Escape}
						Sleep 30
						SendInput, {Escape}
						Sleep 30
						If (A_Index > 5)
							break
					}
					break

			 ; abbrechen wenn es zu lange dauert
				}
				else If  (A_Index > 40)
					break

				Sleep 50

			}

			AlbisGetFocus("Karteikarte",, Kkarte)
			If Addendum.Debug
				SciTEOutput("04: " (stepA := Clock()-startT) " ms [sub: " Kkarte.subFocus "]")
		;}

	;}

	; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	; Hautkrebscreening-Formular
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺;{
		If RegExMatch(EBMtoAdd, "(01745|01746)")
			if (hEHKS := AlbisMenu(34505, "Hautkrebsscreening - Nichtdermatologe ahk_class #32770", 3))
				AlbisHautkrebsScreening()   ; alles leer lassen, alles super!
		If Addendum.Debug
				SciTEOutput("05: " (stepB := Clock()-startT) " ms (" stepB-stepA " ms)  [eHKS]")
	;}

	; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	; auf letztes Datum zurücksetzen und Ausnahmeindikationen korrigieren
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺;{
		AlbisSetzeProgrammDatum(lastPrgDate)
		ausIndi := AusIndisKorrigieren(VSEmpf[PatID].CHRONISCHGRUPPE
													, VSEmpf[PatID].GERIATRISCHGRUPPE
													, Addendum.KeinBMP
													, Addendum.ShowIndis )
		If Addendum.Debug
			SciTEOutput("06: " stepA := Round(Clock()-startT) " ms (" stepA-stepB " ms) [AusIndis]")
	;}

	; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	; CAVE Dokumentation - Änderungen vornehmen
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺;{
		examDate := RegExReplace(DateChoosed, "\d+\.(\d+)\.\d\d(\d+)", "$1^$2")
		If RegExMatch(EBMtoAdd, "O)(01732|0174[56]|01731)\-*.*(01732|0174[56]|01731)*\-*.*(01732|0174[56]|01731)*", exams) {
			GVUHKSKVU := RegExMatch(EBMtoAdd, "01732") ? "G"              	: ""
			GVUHKSKVU .= RegExMatch(EBMtoAdd, "(01745|01746)") ? "H"  	: ""
			GVUHKSKVU .= RegExMatch(EBMtoAdd, "01731") ? "K"               	: ""
			GVUHKSKVU := ", " LTrim(GVUHKSKVU, "/") " " examDate
		}
		COLO	:= RegExMatch(EBMtoAdd, "01740") ? "01740"  	: ""
		AORTA	:= RegExMatch(EBMtoAdd, "01747") ? ", A"        	: ""
		ClipBoard := Trim(Colo . AORTA . GVUHKSKVU , ", ")

		If !(hCave := Albismenu(32778, "Cave! von ahk_class #32770", 4)) {
			return 0
		}
		CLines := AlbisGetCave(false)
		Vorsorge := Impfung := Array()
		For lineNr, line in StrSplit(CLines, "`n") {
			appendix := "\s*((?<nr>\-[\d\w]+)\s*((?<Monat>\d+)[^\w](?<Jahr>\d+)|(?<Jahr>\d+))"
			If RegExMatch(line, "i)(?<stoff>Td|TdPP|TdPert|Masern|Pocken|Tet|BCG|COV19|COV|Pneu|GSI|MMR|Polio|Po|Pert|FSME|Hep[AB]{1,2}|Hep|Hib)")
				Impfung.Push([lineNr, line])
			If RegExMatch(line, "i)(01740|AORTA|A|GVU|HKS|KVU|GHK|GH|GB|Colo|Gastro|PP)") {
				Vorsorge.Push([lineNr, line])
			}
		}

		If Addendum.ChangeCave {
			AlbisCaveSetCellFocus(hCave, lineNr-1, "Beschreibung")
			while (WinExist("Cave! von ahk_class #32770") && A_Index < 10)
				Sleep % ksleep3
		}

	;}

	; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	; CAVE Dokumentation - schliessen
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺;{
		If WinExist("Cave! von ahk_class #32770") {
			If !VerifiedClick("OK", "Cave! von ahk_class #32770")
				If WinExist("Cave! von ahk_class #32770")
					If !VerifiedClick("OK", "Cave! von ahk_class #32770") {
						WinClose, % "Cave! von ahk_class #32770",, 1
						If WinExist("Cave! von ahk_class #32770") {
							BlockInput, Off
							HandsOff(false)
							MsgBox, 0x1000, % StrReplace(A_ScriptName, ".ahk"), % "Bitte schließen Sie das Cave! von Fenster manuell!"
							BlockInput, On
							HandsOff(true)
						}
					}
		}
	;}

	; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	; Impfbuch Hinweis in die Karteikarte zur Erinnerung einschreiben
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺;{
		If !Impfung.Count() {

		; aktuelles Programmdatum lesen
			infoDate := "31.12." SubStr(DateChoosed, 7, 4)+1
			AlbisSetzeProgrammDatum(infoDate)

		; setzen des Zeilendatum vorbereiten
			AlbisActivate(1)
			hInput := AlbisKarteikarteEingabe("Datum")
			AlbisGetFocus("Karteikarte",, Kkarte)
			If (Kkarte.subFocus <> "Datum") {
				BlockInput, Off
				MsgBox, 0x1000, Fehler, Die Eingabefelder der Karteikarte konnten nicht erstellt werden!`n Bitte machen Sie es manuell.
				BlockInput, On
			} 	else
				Sleep % ksleep3

		; Zeilendatum ändern
			If AlbisZeilendatumSetzen(infoDate) {

			; Kürzelfeld aktivieren
	    		AlbisActivate(1)
				hInput := AlbisKarteikarteEingabe("Kürzel")
				AlbisGetFocus("Karteikarte",,Kkarte)
				If (Kkarte.subFocus <> "Kürzel") {
					BlockInput, Off
					MsgBox, 0x1000, Fehler, Die Eingabefelder der Karteikarte konnten nicht erstellt werden!`n Bitte machen Sie es manuell.
					BlockInput, On
				} 	else
					Sleep % ksleep3

			  ; Kürzelfeld befüllen
				If !VerifiedSetText("", "Impf", "ahk_id " Kkarte.hCaret) {
					Send, % "{RAW}Impf"
					Sleep % ksleep2
					SendInput, {TAB}
					Sleep % ksleep2
					SendInput, {TAB}
					Sleep % ksleep2
				}

			  ;~ ; Kürzelfeld prüfen
				;~ ControlGetText, INHALT,, % "ahk_id " Kkarte.hCaret
				;~ If (INHALT <> "Impf") {
					;~ BlockInput, Off
					;~ HandsOff(false)
					;~ MsgBox, 0x1000, Fehler, % "Das Kürzel 'Impf' konnte nicht eingegeben werden.`n Bitte machen Sie es manuell."
					;~ BlockInput, On
				;~ }
				;~ SendInput, {TAB}
				;~ Sleep % ksleep2

			;~ ; Inhaltsfeld aktivieren (falls nicht aktiviert)
				;~ AlbisGetFocus("Karteikarte",, Kkarte)
				;~ If (Kkarte.subFocus <> "Inhalt") {
					;~ SendInput, {TAB}
					;~ Sleep % ksleep2
					;~ while !AlbisGetFocus("Karteikarte", "Inhalt") && (A_Index < 5)
						;~ Sleep % ksleep1
				;~ }

			;~ ; Inhaltsfeld befüllen
				;~ Send, % "{Raw}     "
				;~ Sleep % ksleep1
				;~ SendInput, {BS 5}
				;~ Sleep % ksleep1
				;~ Send, % "{Raw}Impfausweis mitbringen"
				;~ Sleep % ksleep2
				;~ SendInput, % "{TAB}"
				;~ Sleep % ksleep2

			}

		}
		else {
			FileAppend, % Impfung, % Addendum.DBPath "\sonstiges\Impfungen.txt", UTF-8
		}

	;}

		SendInput, {Escape}
		Sleep % ksleep1
		SendInput, {Escape}

	; TI ZIffern
	; 01647,01650,01660,01670,01671,01672,01710B,01710C,01710D,30700

	; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	; MsgBX - melde das Du das fertig hast
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺;{
		If Addendum.ChangeCave {
			If IsObject(ausIndi) {

				Alt 	:= ausIndi.Alt
				RegExReplace(Alt, "\-",, AltC)
				Neu	:= ausIndi.neu
				RegExReplace(Neu, "\-",, NeuC)
				inter := Alt ? (AltC > 0 ? "n" : "") " ersetzt mit" : "n erstmalig eingetragen"
				msg := "Ausnahmeindikation" (AltC > 0 ? "en" : "") ":`n" Alt "`nwurde" inter "`n" Neu

			}
			else {

				msg := "Ausnahmeindikation" (AltC > 0 ? "en" : "") ":`n" Alt "`nwurde" (AltC > 0 ? "n" : "") " nicht geändert"
				MsgBox	, 0x1024
						, % A_ScriptName
						, % "Bitte alles überprüfen!`n`n" msg
						, 20
			}
		}


	;}

	SciTEOutput("Fertig")

BlockInputOff:
	BlockInput, Off
	HandsOff(false)
return
}

AusIndisKorrigieren(Chroniker=false, GB=false, KeinBMP=true, ShowIndis=true)   	{     	;-- lädt Daten für AlbisAusIndisKorrigieren()

	static AusIndisToRemove, AusIndisToHave

	If !AusIndisToRemove {
		AusIndisToRemove := RegExReplace(Addendum.IndisToRemove, "^\s\(")
		AusIndisToRemove := RegExReplace(AusIndisToRemove, "^\s\)")
		AusIndisToRemove := RegExReplace(AusIndisToRemove, "[\,\.\s\;\:\-\+]", "|")
		AusIndisToRemove := "(" RegExReplace(AusIndisToRemove, "[\\]{2,}", "|")  ")"
	}
	If !IsObject(AusIndisToHave) {
		AusIndisToHave := []
		For idx, EBMZiffer in StrSplit(Addendum.IndisToHave, "|")
			AusIndisToHave.Push(EBMZiffer)
	}

	AusIndis := AlbisAusIndisKorrigieren(AusIndisToRemove, AusIndisToHave, "ChronischKrank=" Chroniker " Geriatsch=" GB " KeinBMP=" KeinBMP)

	If ShowIndis {
		If AusIndis.neu {
			RegExReplace(AusIndis.Alt, "\-",, AltC)
			RegExReplace(AusIndis.Neu, "\-",, NeuC)
			inter := AusIndis.Alt ? (AltC > 0 ? "n" : "") " ersetzt mit" : "n erstmalig eingetragen"
			msg := "Ausnahmeindikation" (AltC > 0 ? "en" : "") ":`n" AusIndis.Alt "`nwurde" inter "`n" AusIndis.Neu
		}
		else
			msg := "Ausnahmeindikation" (AltC > 0 ? "en" : "") ":`n" AusIndis.alt "`nwurde" (AltC > 0 ? "n" : "") " nicht geändert"

		MsgBox	, 0x1024, % A_ScriptName, %	"Bitte überprüfen!`n"
																. 	"⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘`n"
																.	msg "`n"
																.	"⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘", 10
	}

}

GEBUEHRWarner(note, param*)                                                                         	{    	;-- erstellt Nutzerabfragen

	global abr, habr, abrHCalStart

	ctrlp  	:= GetWindowSpot(abrHCalStart)
	abrp  	:= GetWindowSpot(habr)

	Switch note {

		Case "EBM":
			 ibxW := 300, ibxH := 120,
			warn := EBM = "03220" ? "03221" : "03220"
			warn := EBM = "03360" ? "03362" : "03360"
			question := 	"Die Ziffer (" param.1 ") kann am selben Tag wie die Ziffer (" warn ")`n"
							. 	"nur mit einer Tagtrennung (z.B. " warn "(um:18:20)" ") abgerechnet werden.`n`n"
							. 	"Sie können jetzt eine Uhrzeit eingeben? "
			InputBox, Tagtrennung, Tagtrennung vornehmen, % question,, ibxW, ibxH, abrp.W-ibxW, ctrlp.Y-ibxH,,, % A_Hour-1 ":" A_Min
			If !Tagtrennung {
				MsgBox, 0x1004, % 	"Ohne Angabe einer Uhrzeit ist eine Tagtrennung nicht möglich.`n"
											.	"Soll diese ohne Zusatzangaben zur Abrechnung gebrachte werden?`n"
											.	"Falls nein, können Sie die Ziffer (" param.1 ") jederzeit erneut eintragen lassen."
				IfMsgBox, Yes
					override := 1
				IfMsgBox, No
					override := 0
				IfMsgBox, Cancel
					return "EBM1"

			}
	}

	;~ MsgBox, 0x1004, Abrechnungskonflikt, %
 }

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Datensammler
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
VorsorgeKomplexe(Abrechnungsquartal, ReIndex:=false, SaveResults:="")              	{     	;-- ermittelt alle Abrechnungsdaten des Quartals

		global DBAccess

		cJSON.EscapeUnicode := "UTF-8"

	; nur bei leerem SaveResults-Parameter ReIndex verwenden
		SaveResults       	:= SaveResults ? SaveResults : ReIndex
		YearQuarter      	:= SubStr(Abrechnungsquartal, 2, 2) . SubStr(Abrechnungsquartal, 1, 1)
		saveCandidates  	:= true
		quarterPath         	:= Addendum.DBasePath "\Q" YearQuarter
		CandidatesPath 	:= quarterPath "\VorsorgeKandidaten_" Abrechnungsquartal ".json"

	; Albisdatenbank Objekt erstellen
		If Addendum.Debug
			SciTEOutput("DBAccess.created: " DBAccess.created)
		If !DBAccess.created
			DBAccess 	:= new AlbisDb(Addendum.AlbisDBPath, 0)

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; nicht neu indizieren, dann gespeicherte Daten laden
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If !ReIndex && FileExist(CandidatesPath) {
			SplashText("bitte warten ...", "lade Vorsorge Kandidaten")
			candidates := cJSON.Load(FileOpen(CandidatesPath, "r", "UTF-8").Read())
			SplashText("bitte warten ...", "Vorsorge Kandidaten wurden geladen (" candidates.Count() ")", 6)
			return candidates
		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; keine Daten vorhanden oder neu indizieren wurde übermittelt
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

	  ; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	  ; Krankenscheine:         	Abrechnungsart: ohne Überweiser, Notfallscheine oder Private
	  ; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺;{
		If (!IsObject(candidates) || ReIndex || !FileExist(quarterPath "\KScheine_" Abrechnungsquartal ".json")) {
			SplashText("bitte warten ... 1/7", "Abrechnungsscheine des Quartal " Abrechnungsquartal " werden gelesen ...")
			If Addendum.Debug
				SciTEOutput("bitte warten ... 1/7, Abrechnungsscheine des Quartal " Abrechnungsquartal " werden gelesen ...")
			KScheine	:= DBAccess.Krankenscheine(Abrechnungsquartal, "Abrechnung,Überweisung,Notfall", false, SaveResults)
			If Addendum.Debug
				SciTEOutput(KScheine.Count() " Abrechnungsscheine wurden im Quartal " Abrechnungsquartal " ermittelt")
		}
	  ;}

	  ; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	  ; Behandlungstage:     	des Quartals für jeden Patienten zusammenstellen
	  ; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺;{
		If (!IsObject(candidates) || ReIndex || !FileExist(quarterPath "\BehTage_" Abrechnungsquartal ".json")) {
			SplashText("bitte warten... 2/7", "Behandlungstage des Quartal " Abrechnungsquartal " ermitteln ...")
			If Addendum.Debug
				SciTEOutput("bitte warten... 2/7, Behandlungstage des Quartal " Abrechnungsquartal " ermitteln ...")
			BehTage := DBAccess.Behandlungstage(Abrechnungsquartal, "Abrechnung,Überweisung,Notfall", false, SaveResults)
			If Addendum.Debug
				SciTEOutput(BehTage.Count() " Patienten mit Behandlungstagen im Quartal " Abrechnungsquartal " ermittelt ...")
		}
		else If (!ReIndex && FileExist(quarterPath "\BehTage_" Abrechnungsquartal ".json")) {
			If Addendum.Debug
				SciTEOutput("bitte warten... 2/7, Behandlungstage im Quartal " Abrechnungsquartal " werden geladen ...")
			BehTage := cJSON.Load(FileOpen(quarterPath "\BehTage_" Abrechnungsquartal ".json", "r", "UTF-8").Read())
		}
	  ;}

	  ; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	  ; Arbeitstage:               	alle Tage mit Eintragungen
	  ; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺;{
		Arbeitstage := Object()
		For PatID, QData in BehTage
			For BTag, Tageseintraege in QData
				Tage .= BTag "`n"
		Sort, Tage, U
		For each, Tagesdatum in StrSplit(Tage, "`n")
			Arbeitstage[Tagesdatum] := 1
	  ;}

	  ; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	  ; Vorsorgedaten:         	von allen Patienten die Vorsorgeziffern und Abrechnungstage finden
	  ; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺;{
		If (!IsObject(candidates) || ReIndex || !FileExist(Addendum.DBasePath "\Vorsorgen.json")) {
			SplashText("bitte warten ... 3/7", "EBM Ziffern aller Quartale werden gesammelt ...")
			VSDaten 	:= DBAccess.VorsorgeDaten(Abrechnungsquartal, ReIndex, SaveResults)
		}
	  ;}

	  ; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	  ; Vorsorgekandidaten:  	finden
	  ; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺;{

		; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
		; letzte Vorsorgekandidatendatei laden, um bearbeitete Datensätze zu übernehmen
		; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
			VSEmpfOld := ""
			If FileExist(CandidatesPath) {
				SplashText("bitte warten ... 4/7", "Vorsagekandidaten Datenbackup wird geladen  ...")
				VSEmpfOld := cJSON.Load(FileOpen(CandidatesPath, "r", "UTF-8").Read())
			}

		;~ If (!IsObject(candidates) || ReIndex ) {        ; || !FileExist(CandidatesPath)

		; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
		; Vorsorgekandidaten neu erstellen
		; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
			SplashText("bitte warten ... 5/7", "EBM Ziffern aller Quartale werden gefiltert ...")
			candidates := DBAccess.VorsorgeFilter(VSDaten, Abrechnungsquartal, "Abrechnung,Überweisung,Notfall", false, (ReIndex ? false : true))

		; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
		; Daten aus anderer Kandidatendatei mit neuen Daten vergleichen
		; (im Moment nur key: "Hide" - manuell ausgeblendete Patienten)
		; SciTEOutput(" [" PatID "] " PatDB[PatID].NAME ", " PatDB[PatID].VORNAME)
		; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
			If IsObject(VSEmpfOld) {
				SplashText("bitte warten ... 6/7", "Anzeigeeinstellungen werden geladen...")
				saveCandidates := true
				For PatID, VS in VSEmpfOld {
					If !candidates.haskey(PatID)
						candidates[PatID] := Object()
					If VSEmpfOld[PatID].haskey("Hide")
						candidates[PatID].Hide := VSEmpfOld[PatID].Hide
				}
			}

		; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
		; Behandlungstage zuorden
		; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
			If IsObject(BehTage) {
				SplashText("bitte warten ... 7/7", "Behandlungstage werden geladen...")
				saveCandidates := true
				For PatID, Tage in BehTage
					If !IsObject(candidates[PatID].BehTage)
						candidates[PatID].BehTage := Tage ;BehTage[PatID]
			}

		;~ }
	  ;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; keine Daten vorhanden oder neu indizieren wurde übermittelt
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		SplashText("bitte warten ...", "Speicher wird bereinigt......")

	  ; Objekte freigeben
		DBAccess.EmptyDBData()

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Kandidaten sichern wenn Änderungen gemacht wurden
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If saveCandidates
			FileOpen(CandidatesPath, "w", "UTF-8").Write(cJSON.Dump(candidates, 1))

	GuiControl, abr:, EBMScheine, % candidates["_Statistik"].KScheine "  "

return candidates
}

ImpfBoss()                                                                                                         	{

	geimpft =
	(
		Td-w 07, Polio w 15, FSME-w 15,
		Tdp-w15, GSI 14
		TdPP-w 15, Td-3 05, HepAB-3 09
	)
}

GetEBMFees(PatID)                                                                                               	{

	global VSEmpf, VSRule

	If !IsObject(VSRule := DBAccess.GetEBMRule("Vorsorgen"))
		throw A_ThisFunc ": EBM Vorsorgeregeln konnten nicht übermittelt werden!"

	EBMFees := Object()
	For feeNr, fee in VSEmpf[PatID].Vorschlag {
		fee := RegExReplace(fee, "\-\d+$", "")
		EBMFees[fee] := VSRule[fee].GO
	}

	If EBMFees.haskey("HKS") && EBMFees.haskey("GVU")
		EBMFees["HKS"] := "01746"
	else if EBMFees.haskey("HKS") && !EBMFees.haskey("GVU")
		EBMFees["HKS"] := "01745"

	cJSON.Dump(EBMFees, 1)

return EBMFees
}

ICDKrauler(Abrechnungsquartal)                                                                           	{     	;-- sucht Patienten für die Abrechnung der Chroniker-Ziffern

	; sucht alle ICD Abrechnungsdiagnosen eines Quartals heraus und vergleicht diese mit einer ICD Liste
	; 2 Datenbankdateien müssen für die Suche ausgelesen werden: 1.BEFUND.dbf und 2.BEFTEXTE.dbf
	; die Eingrenzung der Suche innerhalb der BEFUND.dbf über das Abrechnungsquartal ist nicht anwendbar,
	; Albis hinterlegt nicht immer die Quartalbezeichnung. Meist steht dort eine 0 oder eine andere Zahl
	; das Abrechnungsquartal wird in ein Datum in Form eines RegEx-String umgewandelt aus
	; 0320 wird 2020(07|08|09)\d\d. Damit sind alle Tage des Quartals erfasst.

		static CRICD
		static file_chronkrank
		static wwwCRICD   	:= "https://raw.githubusercontent.com/Ixiko/Addendum-fuer-Albis-on-Windows/meta/include/Daten/ICD-chronisch_krank.json"

		If !file_chronkrank
			file_chronkrank	:= Addendum.Dir "\include\Daten\ICD-chronisch_krank.json"

	; ICD Vergleichsliste laden
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		If !FileExist(file_chronkrank) {

			MsgBox, 0x1004, Daten vermisst, % 	"Der ICD-Krauler benötigt eine`n"
																. 	"Liste von ICD-Schlüsselnummern, die nach Einschätzung`n"
																. 	"der " q "AG medizinische Grouperanpassung" q "`n"
																. 	"chronische Krankheiten kodieren.`n"
																.	"Diese Datei ist nicht vorhanden.`n`n"
																. 	"Das Skript kann die Datei herunterladen.`n"
																.	wwwCRICD "`n"
																. 	"Drücken Sie auf Ja wenn sie einverstanden sind."
			IfMsgBox, No
				return

			DownloadCRICD:
			downloadToFile := true
			while downloadToFile {
				URLDownloadToFile, % wwwCRICD, % file_chronkrank
				If !FileExist(file_chronkrank) {
					MsgBox, 0x1004, Daten vermisst, % "Der Download ist fehlgeschlagen!`n`nMöchten Sie es noch einmal versuchen?"
					IfMsgBox, No
						return
				}
				else
					downloadToFile := false
			}
		}

	; prüft die ICD Vergleichsliste auf das richtige Format
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		CRICD :=cJSON.Load(FileOpen(file_chronkrank, "r", "UTF-8").Read())
		For letter, ICDcollection in CRICD {

				If !RegExMatch(letter, "^[A-Z]$") {
					FileFormatMismatch := true
					break
				}

				If !FileFormatMismatch
					For icdIndex, icd in ICDcollection
						If !RegExMatch(ICD, "^[A-Z]\d{1,2}\.*[\d\*\-\!]*$") {
							FileFormatMismatch := true
							SciTEOutput("mismatch icd: " ICD)
							break
						}

				If FileFormatMismatch {
					CRICD := ""
					MsgBox, 0x1004, Datenformat, % "Die Liste mit ICD-Schlüsselnummern hat ein falsches`n"
																	.  "Format. Die Funktion kann nicht ausgeführt werden.`n"
																	.  "Soll die Orignaldatei aus dem Internet geladen werden?`n" . wwwCRICD
					IfMsgBox, Yes
						ICDKrauler(Abrechnungsquartal)
					IfMsgBox, No
						return
				}

		}


	; Datenbankzugriff
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		chrPausch := Object()
		chrPauschale := ["03220", "03221"]

		albis          	:= new AlbisDB(Addendum.AlbisDBPath, "debug=0 callback=ICDKraulerCallback")
		abrScheine 	:= albis.Abrechnungsscheine(Abrechnungsquartal)           ; nur Abrechnungssscheine
		chroniker  	:= albis.ChronischKrank(Abrechnungsquartal, "reNew=false save=true load=true EmptyStatics=true")
		albis         	:= ""

		For PatID, chrICDs in chroniker {

			If (abrScheine[PatID].PRIVAT || abrScheine[PatID].KSCHEIN  != "Abrechnung")
				continue

			chrPStatus1 := !abrScheine[PatID].GEBUEHR["#" chrPauschale.1] ? 0x0 : 0x1
			chrPStatus2 := !abrScheine[PatID].GEBUEHR["#" chrPauschale.2] ? 0x0 : 0x2
			If (chrPStatus1 && chrPStatus2)
				continue

			output := !chrPStatus1 && !chrPStatus2 ? "03220/1" : !chrPStatus1 ? "03220" : !chrPStatus2 ? "03221" : "- - - -"
			chrPausch[PatID]  := output

		}

return chrPausch
}

ICDKraulerGui(ShowGui:=true)                                                                           	{

	static hQS, hQSi, QSDBS, QShDBS
	static QSiT, QSi, kraulInit := true

	op := GetWindowSpot(AlbisWinID())
	QSiW := Floor(op.CW//4), QSiH := Floor(op.CH//4)

	If kraulInit {

		Gui, QSi: new	, -Caption +ToolWindow  HWNDhQSi -DPIScale +Owner%hQS% ; +AlwaysOnTop
		Gui, QSi: Color	, c000000
		Gui, QSi: Font	, s26 q5 bold, Futura Bk Bt
		Gui, QSi: Add	, Text    	, % "x0 y0 	w" QSiW " cDarkRed Center vQSiT"    	, % "Indexerstellung`nbitte warten...."
		Gui, QSi: Add	, Progress	, % "y+5 	w" QSiW " Center vQSPG1"                	, % "0"

		WinSet, Trans, 200, % "ahk_id " hQSi
		;~ GuiControlGet, cp, QSi: Pos, QSiT
		;~ GuiControl, QSi: Move, QSiT, % "y" Floor(op.CH/2)-Floor(QSiH/2)

		kraulInit := false

	}

	Gui, QSi: Show, %  "y" Floor(op.CH/2)-Floor(QSiH/2) " w" QSiW " h" QSiH " NoActivate " (ShowGui ? "" : "Hide") 	, ICDKrauler Fortschritt   ; "x" op.X " y" op.Y

}

ICDKraulerCallback(itemIndex, maxItems)                                                          	{

	ToolTip, % "item: " SubStr("00000000" itemIndex, -1*(StrLen(maxItems)-1)) "/" SubStr("00000000" maxItems, -1*(StrLen(maxItems)-1)) ,  % A_ScreenWidth//2-100, % A_ScreenHeight//2, 15

return
}

CoronaEBM(Abrechnungsquartal)                                                                         	{   	;-- alle Coronaabstrich und alle Behandlungsziffern finden

	COVID     	:= Object()
	abrYear      	:= "2021"
	aDB          	:= new AlbisDB(Addendum.AlbisDBPath)
	abrScheine  	:= aDB.Abrechnungsscheine(Abrechnungsquartal, false, true)
	withCorona  	:= aDB.GetDBFData(  "BEFUND"
									     			  , {"KUERZEL":"rx:.*(dia|labor).*"
														, "INHALT":"rx:.*(COVIPC|CoV|SARS|COVID|Spezielle|COVSEQ|SARS\-CoV\-2\-Mutation|"
														. "U07\.[12]|U99\.\d|N501Y|delH69|E484K).*"
														, "DATUM":"rx:" abrYear "\d\d\d\d"}
								    				  , [ "PATNR", "DATUM", "KUERZEL", "INHALT"],,1 )

	; nur die Patienten-ID's zusammentragen
		For idx, set in withCorona {
			If !COVID.haskey(set.PATNR)
				COVID[set.PATNR] := {"Matches":[], "GNROK":false}
			COVID[set.PATNR].Matches.Push({"TEXT": (set.KUERZEL " | " set.INHALT), "DATUM":set.DATUM})
		}

	; fehlende GNR suchen
		GNRMustHave := ["#32006", "#88240"]
		For PatID, COVIData in COVID {
			GNRMatched := GNRMustHave.MaxIndex()
			For GNR, gnrData in abrScheine[PatID].GEBUEHR
				For lkoIndex, lko in GNRMustHave
					If (lko = GNR) {
						GNRMatched --
						If (GNRMatched = 0)
							break
					}
			COVID[PatID].GNROK := GNRMatched = 0 ? true : false
		}

	; Karteikarte für Karteikarte anzeigen
		NotFirstPat := 1
		For PatID, data in COVID {
			PatName := PatDB[PatID].NAME ", " PatDB[PatID].VORNAME
			If (NotFirstPat > 1) {
				MsgBox, 0x1024, % "Weiter?", % "nächster Patient (" NotFirstPat "/" COVID.Count()
					. "):`n[" PatID "] " PatName "`nAbrechnung: " (data.GNROK ? "ist ok":"da fehlt was")
				IfMsgBox, No
					break
			}
			AlbisAkteOeffnen(PatName, PatID)
			NotFirstPat ++
		}

		aDB:= ""
		abrScheine := withCorona := COVID := ""

return
}

NoSchein(Abrechnungsquartal)                                                                            	{  	;-- soll Pat. mit möglichen Leistungen aber ohne angelegten Schein finden



}

; - - - -
; vorbereitete Funktionen für Auswertungen und Fehlersuche
; - - - -
eHKSFormularOderEBM(Abrechnungsquartal)                                                    	{  	;-- sucht nach Hautkrebsscreeningziffer u. fehlendem eHKS Formular

	; sucht nach abgerechneter Hautkrebsscreeningziffer und fehlendem eHKS Formular und umgekehrt

	quartal := LTrim(Abrechnungsquartal, "0")
	QZ := SubStr(quartal, 1, 1)
	QY := SubStr(quartal, 2, 2)

	If !FileExist(fpath := Addendum.DBasePath "\Q" QY QZ "\BehTage_" QY QZ ".json") {
		MsgBox, Abrechnungsdatei ist noch nicht erstellt worden.`nDie Funktion bricht ab!
		return
	}

	behTage := cJSON.Load(FileOpen(fpath, "r", "UTF-8").Read())

	eHKSF := []
	ebm := eHKS := 0

	For PatID, BTage in behTage {

		If (lastPatID != PatID) {

			If (ebm || eHKS)
				If (ebm != eHKS)
					eHKSF.Push(lastPatID)

			lastPatID := PatID
			ebm := eHKS := 0

		}

		For BTag, Eintraege in BTage
			For each, Eintrag in Eintraege {

				If (EBM && eHKS)
					break

				If (Eintrag.kzl = "lko" && RegExMatch(Eintrag.inh, "\b01746\b"))
					EBM := 1
				else if (Eintrag.kzl = "fhksn")
					eHKS := 1

			}

	}



return eHKSF
}


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Hilfsfunktionen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - -;{
class vacation                                                                                                    	{    	;-- Funktionsklasse Datumsberechnungen (Urlaub)

	; ACHTUNG: benötigt Addendum.Praxis.Urlaub/Sprechstunde als globales Objekt
	;
	; Funktionen:    1. 	Daten aus Addendum.ini werden geparst. die Urlaubstage können mit relativ freier Schreibweise in der Ini-Datei eingetragen sein
	;                          	einzelne Tage 	:  	in der Form 01.01.2022 auch als 1.1.22
	;                           Datumbereiche	:	05.06.2022-06.06.2022 oder 05.06.-06.06.2022 oder 5.-6.6.22 sowie Jahreswechsel
	;                       	Achtung        	:	zwingend zur Unterscheidung sind Punkte und Minuszeichen. Leerzeichen, andere Zeichen können zwischen den Datumzahlen
	;
	; letzte Änderung: 18.12.2021

	ConvertHolidaysString(holidays:="")             	{                                     	;-- Ini-String

		/*  parst einen String welcher Urlaubstage enthält

			Beispielformat: 01.03.2021, 08.10.-14.10.2021, 25.-26.12.2021, 29.-30.12.

			- die einzelnen Tage oder Bereiche müssen durch ein Komma getrennt sein
			- fehlt eine Jahreszahl wird diese mit dem aktuellen Jahr gleichgesetzt
			- ACHTUNG: auch wenn eine recht flexible Schreibweise des Datum möglich ist,
								 werden fehlerhafte Datumzahlen nicht erkannt

		*/

			Debug := false

		  ; Urlaubszeiten mit einem Datum in der Vergangenheit werden aussortiert.
			static rxUrlaub

			If !rxUrlaub {
				rxUrlaub1	:= "(?<StartD>\d{1,2})\.(?<StartM>\d{1,2})*\.*(?<StartY>\d{2,4})*"
				rxUrlaub2	:= "\s*-\s*(?<EndD>\d{1,2})\.(?<EndM>\d{1,2})\.*(?<EndY>\d{2,4})*"
				rxUrlaub 	:= "(" rxUrlaub1 rxUrlaub2 ")|(" StrReplace(rxUrlaub1, "Start") ")"
			}

			spos := 1, AToday := A_YYYY . A_MM . A_DD, AYearNow := SubStr(A_YYYY, 1, 2)
			vacations := Array()

		  ; Abbruch wenn, nichts oder ein falsche String übergeben wurde
			If !holidays
				return

		 ; Datum für Datum extrahieren
			while (spos := RegExMatch(holidays, rxUrlaub, Px, spos)) {

			  ; Stringposition weiterrücken
				spos  	+= StrLen(Px)

			  ; nur ein Datum
				PxStartD  	:= PxD	? PxD	: PxStartD
				PxStartM  	:= PxM	? PxM 	: PxStartM
				PxStartY  	:= PxY	? PxY 	: PxStartY

			  ; Debug
				Debug && PxD && PxM ? SciTEOutput("PxDMY: " PxD "." PxM "." PxY)

			 ; eine zweistellige Jahreszahl auf eine vierstellige Jahreszahl ändern, leere Jahreszahlen ersetzen
				PxStartY	:= StrLen(PxStartY)	= 2 	? AYearNow . PxStartY 	: PxStartY	? PxStartY 	: PxEndY 	? PxEndY 	: A_YYYY
				PxEndY 	:= StrLen(PxEndY)	= 2 	? AYearNow . PxEndY  	: PxEndY	? PxEndY	: PxEndD 	? A_YYYY	: ""

			  ; fehlenden Monat ersetzen
				PxStartM 	:= PxStartM ? PxStartM : PxEndM

			  ; Debug
				Debug ? SciTEOutput("Px: " px "`nPxStart: " PxStartD "." PxStartM "." PxStartY "`nPxEnd: " PxEndD "." PxEndM "." PxEndY )

			  ; Monate und Tage 2-stellig auffüllen
				PxStartM 	:= SubStr("00" PxStartM	, -1)
				PxStartD 	:= SubStr("00" PxStartD	, -1)
				PxEndM   	:= SubStr("00" PxEndM	, -1)
				PxEndD   	:= SubStr("00" PxEndD	, -1)

			  ; die Formatierung wird auf YYYYMMDD geändert
				AYearS 	:= (!PxStartY && !PxEndY)	? A_YYYY : PxStartY ? PxStartY 	: PxEndY 	? PxEndY   ; Jahr
				AYearE 	:= (!PxStartY && !PxEndY)	? A_YYYY : PxEndY ? PxEndY 	: PxStartY 	? PxStartY   ; Jahr
				ADayE  	:= (PxEndM 	&& PxEndD)	? AYearE . PxEndM . PxEndD : ""
				;~ AYearS   	:= PxStartM < PxEndM ? AYears + 1
				ADayS  	:= AYearS . PxStartM . PxStartD

			  ; ein Datumsbereich wird anders hinterlegt, als ein einzelner freier Tag
				If (ADayS && ADayE && !this.PeriodExists(ADayS, ADayE))
					vacations.Push({"IsPeriod":true	, "firstday":ADayS, "lastday":ADayE})
				else If (ADayS && !ADayE && !this.PeriodExists(ADayS))
					vacations.Push({"IsPeriod":false	, "day":ADayS})

				PxD := PxM := PxY := PxStartD := PxStartM := PxStartY := PxEndD := PxEndM := PxEndY := ""

			  ; Debug
				Debug ? SciTEOutput("`n")

			}

	return vacations
	}

	PeriodExists(firstday, lastday:="")                	{                                    	;-- Datum oder Datumsbereich suchen

		If !firstday && !lastday
			return false

		For hdNR, holidays in Addendum.Praxis.Urlaub
			If (firstday && lastDate) {
				If (firstday=holidays.firstday && lastday=holidays.lastday)
					return hdNR
			}
			else {
				If (!holiday.IsPeriod && firstday=holidays.day)
					return hdNR
			}

	return false
	}

	DateIsHoliday(datestring)                            	{                                     	;-- ermittelt ob ein übergebenes Datum innerhalb eines Praxisurlaub liegt

	  ; automatisch 4-stelliges Jahresformat (funktioniert bis 2099)
		If RegExMatch(datestring, "(?<D>\d{1,2})\.(?<M>\d{1,2})\.(?<Y>(\d{2}|\d{4}))", t)
			datestring := SubStr(SubStr(A_YYYY, 1, 2) . tY, -2) . tM . tD

		For hdNR, holidays in Addendum.Praxis.Urlaub
			If holidays.IsPeriod && (datestring>=holidays.firstday && datestring<=holidays.lastday)
				return hdNR
			else if (!holidays.IsPeriod && datestring=holidays.day)
				return hdNR

	return false
	}

	isConsultationTime(date:="", time:="", ByRef whynotwork:="") {            	;-- prüft ob zu diesem Zeitpunkt Sprechstunde (Konsultationszeit) ist

	  ; benötigt wird das durch .ConvertHolidaysString() erstellte Objekt in Addendum.Praxis.Urlaub

	  ; über die ByRef Variable "whynotwork" wird zurückgegeben welcher Status am untersuchten Zeitpunkt vorliegt,
	  ; wie 	1= regulärer arbeitsfreier Tag
	  ; 		2= Urlaubstag/Feiertag
	  ;			3= außerhalb der Sprechzeiten  ###(a ist vor, b ist nach der Sprechstunde und c ist zwischen den Sprechstunden)

	  ; letzte Änderung: 15.01.2022

		static reasons := {"nw1a":"regulär freier Arbeitstag", "nw1b":"arbeitsfreies Wochende", "nw2":"Urlaub", "nw3a":"vor der Sprechstunde", "nw3b":"nach der Sprechstunde", "nw3c":"zwischen den Sprechstunden" }

	  ; aktueller Tag oder/und Uhrzeit einstellen wenn sie leer sind
		time	:= time	? time	: A_Hour ":" A_Min
		date	:= date	? date	: A_DD "." A_MM "." A_YYYY

	  ; Datum/Zeitstring konvertieren
		RegExMatch(time, "(?<Hour>\d{1,2})\s*[:\.\-]\s*(?<Min>\d{1,2})\s*[:\.\-]*\s*(?<Sek>\d{1,2})*", t)
		time := SubStr("00" tHour, -1) . SubStr("00" tMin, -1)
		time := SubStr(time, 1, 4)
		If !RegExMatch(date, "\d+.\d+.\d+")
			date := A_DD "." A_MM "." A_YYYY
		date := ConvertToDBASEDate(date)


	  ; Wochentag berechnen
		dow	:= DayOfWeek(date, "full", "yyyyMMdd")

	 ; an diesem ist Wochentag ist keine Sprechstunde
		If !Addendum.Praxis.Sprechstunde.haskey(dow) {
			whynotwork := {"reason":reasons["nw1" (RegExMatch(dow, "(Samstag|Sonntag)") ? "b":"a")], "level":1, "weekday":dow}
			return false
		}
		else If this.DateIsHoliday(date) {
			whynotwork := {"reason":reasons["nw2"], "level":2, "weekday":dow}
			return false
		}

	  ; Zeitpunkt liegt während oder außerhalb der Sprechstunde
		isduring  	:= false
		conTimes	:= RegExReplace(Addendum.Praxis.Sprechstunde[dow], "(\s*,\s*)|(\d)\s+(\d)", "$2|$3)")
		For conNr, timestr in StrSplit(conTimes, "|") {
			If RegExMatch(timestr, "(?<Hour1>\d{1,2})\s*:\s*(?<Min1>\d{1,2})\s*\-\s*(?<Hour2>\d{1,2})\s*:\s*(?<Min2>\d{1,2})", con) {
				conStart:= SubStr("00" conHour1, -1) . SubStr("00" conMin1, -1)
				conEnd	:= SubStr("00" conHour2, -1) . SubStr("00" conMin2, -1)
				If (time>=conStart && time<=conEnd)                                          	; liegt in der Sprechstunde
					isduring := true
			}
		}

	  ; außerhalb der Sprechzeiten whynotwork
		If !isduring {
			RegExMatch(conTimes, "^\s*(?<Hour1>\d{1,2})\s*:\s*(?<Min1>\d{1,2}).*?(?<Hour2>\d{1,2})\s*:\s*(?<Min2>\d{1,2})\s*$", con)
			conDayS 	:= SubStr("00" conHour1, -1) . SubStr("00" conMin1, -1)
			conDayE	:= SubStr("00" conHour2, -1) . SubStr("00" conMin2, -1)
			sublevel  	:= time<conDayS ? "a" : time>conDayE ? "b" : "c"
			whynotwork := {"reason":reasons["nw3" sublevel], "level":2, "sublevel": sublevel, "weekday":dow}
		}


	return isduring
	}

}

SaveAbrechnungsquartal(Abrechnungsquartal:="")                                                	{  	;-- speichert das aktuelle bearbeitete Abrechnungsquartal in der Addendum.ini

	global abrQ

	Gui, abr: Default
	Gui, abr: Submit, NoHide

	If !RegExMatch(Abrechnungsquartal, "^[1-4][0-9]{2}$")
		Abrechnungsquartal := LTrim(abrQ, "0")

	IniWrite, % Abrechnungsquartal, % Addendum.ini, Abrechnungshelfer, Abrechnungsassistent_letztesArbeitsquartal

return Abrechnungsquartal
}

SaveGuiPos(hwnd, Fenstername)                                                                        	{

	global abr, abrOntop, alwtop, habr, abrHReload, abrHOnTop, ontop, abrYoungStars, abrOldStars, abrAngels

	Gui, abr: Submit, NoHide

	props 	:= alwtop ? true : false
	win   	:= GetWindowSpot(hwnd)

	If (win.X > 0 && win.Y > 0) {
		workpl	:= WorkPlaceIdentifier()
		winPos 	:= "x" win.X " y" win.Y
		IniWrite, % winpos	, % Addendum.Ini, % Addendum.compname, %  Fenstername "_Pos_" workpl
	}

	IniWrite, % abrYoungStars "|" abrOldStars "|" abrAngels	, % Addendum.Ini, % Addendum.compname, %  Fenstername "_Optionen"
	IniWrite, % Addendum.abrOnTop ? "Ja":"Nein"            	, % Addendum.Ini, % Addendum.compname, %  Fenstername "_AlwaysOnTop"
	IniWrite, % Addendum.abrOnPos                                	, % Addendum.Ini, % Addendum.compname, %  Fenstername "_TilingModus"

}

WorkPlaceIdentifier()                                                                                           	{  	;-- Identifikationsstring für die Arbeitsumgebung

  ; so lassen sich einfacher unterschiedliche Monitorarbeitsablätze identifizieren und die Gui kann individueller positioniert werden

	SysGet, MonCount, MonitorCount
	Loop % MonCount {
		SysGet, Mon, Monitor, % A_Index
		MonCRC .= MonRight "|" MonBottom "|"
	}
return LTrim(CRC32(RTrim(MonCRC, "|")), "^0x")
}

CRC32(str, enc = "UTF-8")                                                                                 	{
    l := (enc = "CP1200" || enc = "UTF-16") ? 2 : 1, s := (StrPut(str, enc) - 1) * l
    VarSetCapacity(b, s, 0) && StrPut(str, &b, floor(s / l), enc)
    CRC32 := DllCall("ntdll.dll\RtlComputeCrc32", "UInt", 0, "Ptr", &b, "UInt", s)
    return Format("{:#X}", CRC32)
}

MessageWorker(InComing)                                                                                	{   	;-- verarbeitet Nachrichten

		recv := {	"txtmsg"		: (StrSplit(InComing, "|").1)
					, 	"opt"     	: (StrSplit(InComing, "|").2)
					, 	"fromID"	: (StrSplit(InComing, "|").3)}

	; Laborabruf_*.ahk
		; per WM_COPYDATA die Daten im Laborjournal auffrischen (Skript wird neu gestartet)
		if RegExMatch(recv.txtmsg, "i)\s*reload") {
			Send_WM_COPYDATA("reload Laborjournal||" labJ.hwnd, recv.fromID)
			Reload
		}

return
}

RGBtoBGR(cRGB)                                                                                                	{
	RegExMatch(cRGB, "i)^(?<Prepend>(0x|#))*(?<Red>[A-F\d]{2})(?<Green>[A-F\d]{2})(?<Blue>[A-F\d]{2})", c)
return cPrepend . cBlue . cGreen . cRed
}

PosStringToObject(string)                                                                                    	{    	;-- ini-String to object

	p := Object()
	For wIdx, coord in StrSplit("XYWH") {
		RegExMatch(string, "i)" coord "(?<Pos>\d+)", w)
		p[coord] := wPos
	}

return p
}

SplashText(title, msg:="", offtime:="")                                                                   	{   	;-- zeigt nur ein Splashtext Hinweis an

	global habr, abrhLV

	hWin := abrhLV ? abrhLV : AlbisWinID()

	SplashTextOn, 300, 20, % title, % msg

	If (hSTO := WinExist(title " ahk_class AutoHotkey2")) {

		stp 	:= GetWindowSpot(hSTO)
		abrP := GetWindowSpot(hWin)
		SetWindowPos(hSTO, abrP.X+(abrP.W//2-stp.W//2), abrP.Y+(abrP.H//2-stp.H//2), stp.W, stp.H+30)

		WinSet, Style 	, 0x50000000, % "ahk_id " hSTO
		WinSet, ExStyle	, 0x00000208, % "ahk_id " hSTO

	}

	If (RegExMatch(offtime, "^\d+$") || !msg)
		SetTimer, SplashTOff, % -1 * (!msg ? 1 : offtime*1000)

return

SplashTOff:
	;~ SetTimer, SplashTOff, Off
	SplashTextOff
return
}

WorkInProgress()                                                                                               	{		;--

	global 	habr, abrTxtPRG
	static 	BlindWin    	:= false
	static 	WS_POPUP 	:= 0x80000000
	static 	WS_CHILD 	:= 0x40000000

	SetWinDelay, -1

	If BlindWin {
		BlindWin := !BlindWin
		Gui, 1: Destroy
		Gui, 2: Destroy
		return
	}

	abrW := GetWindowSpot(habr)
	Gui, 1: +LastFound -Caption -Border +hWndhGui1 +Owner +AlwaysOnTop
	Gui, 1: Color, 000000
	;WinSet, TransColor, 000000, % "ahk_id " hGui1
	Parent_ID := WinExist()

	Gui, 2: margin,1,1
	Gui, 2: -Caption -Border +hWndhGui2 +%WS_CHILD% -%WS_POPUP%
	Gui, 2: Color, c000001
	;~ WinSet, TransColor, off  	, % "ahk_id " hGui2
	;~ WinSet, TransColor, c000001	, % "ahk_id " hGui2
	Gui, 1: Font, s28 bold,  Segoe UI
	Gui, 1: Add, Text, % "xm y10 w" abrW.CW " Center cDarkBlue", % "der Abrechnungsassistent arbeitet..."
	Gui, 2: Font, s10, Segoe UI
	Gui, 2: Add, Edit, % "xm+20 y+10 r30 w" abrW.CW-40 " ReadOnly BackgroundTrans vabrTxtPRG hwndabrHTxtPRG"
	Gui, 2: Font, s10 bold, Segoe UI
	;Gui, 2: Add, Button, y+15 x205 w80 h30 gSubmit Default, Submit
	Gui, 2: +LastFound
	Child_ID := WinExist()
	;DllCall("SetParent", "uint",  Child_ID, "uint", Parent_ID)
	WinSet, Style, 0x50011804	, % "ahk_id " abrHTxtPRG
	WinSet, Style, 0x00000000	, % "ahk_id " abrHTxtPRG

	Gui, 1: Show, % "x" abrW.X+abrW.BW " y" abrW.Y+abrW.BH+20 " w" abrW.CW-1 " h" abrW.CH
	EnableBlur(hGui1)
	Gui, 2: Show, % "x0 y80 w" abrW.CW " h" abrW.CH-80
	WinSet, Redraw,, % "ahk_id " hGui2
	WinSet, Redraw,, % "ahk_id " hGui1

	BlindWin := !BlindWin

	Pause

}

EnableBlur(hWnd)                                                                                              	{  	;-- Fenster unscharf erscheinen lassen

  ;Function by qwerty12 and jNizM (found on https://autohotkey.com/boards/viewtopic.php?t=18823)

  ;WindowCompositionAttribute
  WCA_ACCENT_POLICY := 19

  ;AccentState
  ACCENT_DISABLED := 0,
  ACCENT_ENABLE_GRADIENT := 1,
  ACCENT_ENABLE_TRANSPARENTGRADIENT := 2,
  ACCENT_ENABLE_BLURBEHIND := 3,
  ACCENT_INVALID_STATE := 4

  accentStructSize := VarSetCapacity(AccentPolicy, 4*4, 0)
  NumPut(ACCENT_ENABLE_BLURBEHIND, AccentPolicy, 0, "UInt")

  padding := A_PtrSize == 8 ? 4 : 0
  VarSetCapacity(WindowCompositionAttributeData, 4 + padding + A_PtrSize + 4 + padding)
  NumPut(WCA_ACCENT_POLICY, WindowCompositionAttributeData, 0, "UInt")
  NumPut(&AccentPolicy, WindowCompositionAttributeData, 4 + padding, "Ptr")
  NumPut(accentStructSize, WindowCompositionAttributeData, 4 + padding + A_PtrSize, "UInt")

  DllCall("SetWindowCompositionAttribute", "Ptr", hWnd, "Ptr", &WindowCompositionAttributeData)

}

HandsOff(showGui, width:="", ImagePath:="")                                                    	{ 		;-- Nutzerhinweis einblenden

	global abr, habr
	Static gvisible := false, firstrun := true, hHoff, minusX
	static wpos := Object()

	If firstrun {

		mdiclient := GetWindowSpot(AlbisMDIChildGetActive())
		x := mdiclient.X + Floor(mdiclient.W/2), y := mdiclient.Y + Floor(mdiclient.H/2)

		Gui, HOff: new	, +LastFound -Caption +E0x80000 +Owner +AlwaysOnTop +hwndhHoff +ToolWindow +OwnDialogs
		Gui, HOff: Show, NoActivate, HandsOff

		if !FileExist(ImagePath)
			return

		DllCall("gdiplus\GdipCreateBitmapFromFile", WStr, ImagePath, PtrP, pBitmap)
		DllCall("gdiplus\GdipGetImageWidth", Ptr, (pBitmap := Gdip_ResizeBitmap(pBitmap, width, 400, 1,, 1)), UIntP, w)
		DllCall("gdiplus\GdipGetImageHeight", Ptr, pBitmap, UIntP, h)

		x -= Floor(width), y -= Floor(h)

		hdc        	:= CreateCompatibleDC()
		hbm        	:= CreateDIBSection(w, h)
		obm        	:= SelectObject(hdc, hbm)
		pGraphics	:=Gdip_GraphicsFromHDC(hdc)

		Gdip_SetSmoothingMode(pGraphics, 4)
		Gdip_SetInterpolationMode(pGraphics, 7)
		Gdip_GraphicsClear(pGraphics)

		UpdateLayeredWindow(hHoff, hdc, 0, 0, w,h)
		Gdip_DrawImage(pGraphics, pBitmap, 0, 0, w, h, 0, 0, w, h)
		UpdateLayeredWindow(hHoff, hdc)

		firstrun := false

	}

	Gui, HOff: Show, % " NoActivate Hide"

	If (showGui = false) {
		SetTimer, HandsOffAni, Off
		return
	}
	else if (showGui = true) {

		abrp		:= GetWindowSpot(habr), hffp := GetWindowSpot(hHoff)
		wPos		:= {"X":abrp.X+1, "Y":abrp.Y+(abrp.H//2-hffp.H//2), "W":hffp.W, "H":hffp.H, "maxL":abrp.X-hffp.W}
		minusX	:= 5
		SciTEOutput("`n {winpos: x" wpos.X " y" wpos.Y " w" wPos.W " h" wPos.H " maxL" wpos.maxL "}")

		SetWindowPos(hHoff, wPos.X, wPos.Y, wPos.W, wPos.H)
		Gui, HOff: Show	, % winpos " NoActivate "
		SetWindowPos(hHoff, wPos.X, wPos.Y, wPos.W, wPos.H)

		;~ SetTimer, HandsOffAni, 10

	}

return
;  {winpos: x1815 y487 w380 h400 maxL1434}
HandsOffAni:

	delta := wPos.X-wPos.maxL
	If (delta <= 50)
		minusX := Floor(delta/10) = 0 ? 1 : Floor(delta/10)
	wPos.X -= minusX
	IF Mod(wPos.X, 10)=0
		SciTEOutput(wPos.X)

	SetWindowPos(hHoff, wPos.X, wPos.Y, wPos.W, wPos.H)

	If (wPos.X <= wPos.maxL)
		SetTimer, HandsOffAni, Off


return

}

MsgBoxMove(msgboxTitle, msgBoxText)                                                              	{    	;-- verschiebt einen MsgBox Dialog

	global abr, habr, abrCalStart, abrHCalStart, abrHunhide, abrHLV

	hmsbx 	:= WinExist(msgboxTitle " ahk_class #32770", msgboxText)

	If InStr(msgboxText, "ausgeblendeten") {
		xhwnd 	:= abrHunhide
		xtop  	:= false
	}
	else if InStr(msgboxText, "ausgewähltes Datum") {
		xhwnd 	:= abrHCalStart
		xtop  	:= true
	}

	msbx    	:= GetWindowSpot(hmsbx)
	abrp     	:= GetWindowSpot(habr)
 	ctrlp      	:= GetWindowSpot(xhwnd)

	msbx.X  	:= ctrlp.X-(msbx.W//2-ctrlp.W//2)
	msbx.X 		:= msbx.X+msbx.W > abrp.X+abrp.W ? (abrp.X+abrp.W)-msbx.W : msbx.X   ; MsgBox soll innerhalb des Elternfensters bleiben
	msbx.X		:= msbx.X < abrp.X ? abrp.X : msbx.X
	msbx.Y  	:= xtop ? ctrlp.Y-msbx.H : ctrlp.Y+ctrlp.H

	SetWindowPos(hmsbx, msbx.X, msbx.Y, msbx.W, msbx.H)

}

Create_Abrechnungsassistent_ico(NewHandle := False)                                      	{
VarSetCapacity(B64, 8724 << !!A_IsUnicode)
B64 := "AAABAAEALS0AAAEAGAB4GQAAFgAAACgAAAAtAAAAWgAAAAEAGAAAAAAAAAAAAGAAAABgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABziWJbjTZLjxZDkAZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBDkAZLjxZbjTZ0iWMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAfIhyUY4hQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAUY4hfIhzAAAAAAAAAAAAAAAAAAAAAHWJZUWQCUCRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAEuXD1uhJGmpN26sPl+jKU6ZE0CRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAEWQCXWJZQAAAAAAAAAAAAB7iHJEkAhAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQFqqTibxXnI4Lb1+fL////////////////////////7/frU5saky4Z0r0ZFlAdAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBFkAl8iHMAAAAAAAAAUo4iQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEARpQIhrpe3ezS////////////////////////////////////////////////////////7/bppsyIVJ0bQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAUI4hAAAAAHOJYkCRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAIa6XvH37P////////////////////////////////////////////////////////////////////////r8+KTKhUWUB0CRAECRAECRAECRAECRAECRAECRAECRAECRAHSJYwBbjTVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBVnRzJ4Lf////////////////////////////////////////////////////////////////////////////////////////f7NRoqDVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBcjDYAS48UQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAWZ8h5vDd////////////////////////////////////////////////////////////////////////////////////////////////+/z5erNOQJEAQJEAQJEAQJEAQJEAQJEAQJEAS48VAEOQBkCRAECRAECRAECRAECRAECRAECRAGCkK+z05v///////////////////////////////////////////+z05r/aqr/aqtHlwv////////////////////////////////////////////v8+XqzTkCRAECRAECRAECRAECRAECRAEOQBwBAkQBAkQBAkQBAkQBAkQBAkQBAkQBNmBHm8N3////////////////////////////////////////////////F3bFAkQBAkQB4sUv////////////////////////////////////////////////6/PhlpzJAkQBAkQBAkQBAkQBAkQBAkQAAQJEAQJEAQJEAQJEAQJEAQJEAQJEAudei////////////////////////////////////////////////////v9qqQJEAQJEAdrBI////////////////////////////////////////////////////4O3VRJQGQJEAQJEAQJEAQJEAQJEAAECRAECRAECRAECRAECRAECRAGioNf7//v////////////////////////////////////////P476DIgGKlLkGSAkCRAECRAECRAFOcGYW5XN3r0f///////////////////////////////////////////5LAbkCRAECRAECRAECRAECRAABAkQBAkQBAkQBAkQBAkQBAkQC/2qn///////////////////////////////////////+/2qpMmBBAkQBAkQBIlgtfoylipS5NmRJAkQBAkQBBkgKVwnL9/v3////////////////////////////////////o8uBBkgJAkQBAkQBAkQBAkQAAQJEAQJEAQJEAQJEAQJEASZYM+Pv2////////////////////////////////////yeC3QZICQJEASZYMudei+/z5////////////0eTBYqUuQJEAQJEAosqD////////////////////////////////////////bas8QJEAQJEAQJEAQJEAAECRAECRAECRAECRAECRAHGtQv///////////////////////////////////////2uqOUCRAECRALbVnf///////////////////////+vz5EOTBECRAEuXD/v8+f///////////////////////////////////57HfUCRAECRAECRAECRAABAkQBAkQBAkQBAkQBAkQCRwGz///////////////////////////////////////9rqjlZoCJZoCLr8+T///////////////////////////9coSZAkQBAkQDn8d////////////////////////////////////+82aZAkQBAkQBAkQBAkQAAQJEAQJEAQJEAQJEAQJEAncZ8////////////////////////////////////////////////////////////////////////////////6vPjRJMFQJEARZQH9vrz////////////////////////////////////yuC4QJEAQJEAQJEAQJEAAECRAECRAECRAECRAECRAJPBb/////////////////////////////////////////////////////////////////////////r8+LnXolmfIUCRAECRAJPBb////////////////////////////////////////8LcrUCRAECRAECRAECRAABAkQBAkQBAkQBAkQBAkQCDuFr////////////////////////////////////////////////////////7/frR5cKfyH9wrUBHlQpAkQBAkQBBkQGPv2r7/fr///////////////////////////////////////+v0ZRAkQBAkQBAkQBAkQAAQJEAQJEAQJEAQJEAQJEAYaQs////////////////////////////////////////////////7vXok8FvSZYMQJEAQJEAQJEAQJEAVp4ejL1l2urO////////////////////////////////////////////////k8FvQJEAQJEAQJEAQJEAAECRAECRAECRAECRAECRAEGRAe/26f///////////////////////////////////////9PmxEuXD0CRAECRAE+aFIS4W7PTmuTv2////////////////////////////////////////////////////////////2KlLkCRAECRAECRAECRAABAkQBAkQBAkQBAkQBAkQBAkQC21Z3////////////////////////////////////8/ftWnh5AkQBAkQCfyH/4+/b////////////////////////////////////////////////////////////////////t9edBkQFAkQBAkQBAkQBAkQAAQJEAQJEAQJEAQJEAQJEAQJEAaKg1////////////////////////////////////3evRQJEAQJEAYaQs////////////////////////9vrzpsyIs9OZvNil////////////////////////////////////o8qEQJEAQJEAQJEAQJEAQJEAAECRAECRAECRAECRAECRAECRAECRANTmxf///////////////////////////////+Pv2kCRAECRAFmgIv///////////////////////73Zp0CRAECRAG+sP/////////////////////////////////z9+1WdHECRAECRAECRAECRAECRAABAkQBAkQBAkQBAkQBAkQBAkQBAkQBqqTj+//7///////////////////////////////9oqDZAkQBAkQCSwG3z+O/////////////X6MlTnBlAkQBAkQC31p////////////////////////////////+x0pdAkQBAkQBAkQBAkQBAkQBAkQAAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAuNag////////////////////////////////5vDdXKEmQJEAQJEAR5UJYaQsbas8WaAiQJEAQJEAQJEAjb1n/v/+////////////////////////////8ffsTpkTQJEAQJEAQJEAQJEAQJEAQJEAAECRAECRAECRAECRAECRAECRAECRAECRAEuXD+Xw3P////////////////////////////////j79arOjWKkLUGRAUCRAECRAECRAEeVCXOuRMrguP///////////////////////////////////4W5XECRAECRAECRAECRAECRAECRAECRAABAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBmpzP5/Pf///////////////////////////////////////+sz5BAkQBAkQCFuVz8/fv///////////////////////////////////////+u0JNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQAAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAerJN+/36////////////////////////////////////rdCSQJEAQJEAj79q////////////////////////////////////////zOK7RJMFQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAIC2Vvv9+v////////////////////////////////r8+PL47vL47vj79f///////////////////////////////////8jftUaUCECRAECRAECRAECRAECRAECRAECRAECRAECRAABAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBwrUDx9+z///////////////////////////////////////////////////////////////////////////+31p9ElAZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQAAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAS5cPca1BhLhbd7FKc65E1ObF////////////////////////////////////////////////////////////////9Pnwg7dZQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAAECRAECRAECRAECRAECRAECRAECRAECRAGqpOPj79v///////////////9/s1OHu1/3+/f////////////////////////////////////////////////v8+a/RlFOcGkCRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAABAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBDkwTg7dbR5MGaxXiMvWabxnrC3K7v9un///+/2qmmzIimzIimzIimzIimzIimzIimzIimzIimzIimzIidxnxWnh5AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQAAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEATpkTQJEAQJEAQJEAQJEAQJEAQZEBcK1AibthQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAEGSAmGkLI6+aLzZpvX58uXw3LHSl7PTmbPTmbPTmbPTmbPTmbPTmbPTmbPTmV+jKUCRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAABDkAZAkQBAkQBAkQBAkQBAkQB9tFGpzYymzIit0JLF3rLb6s/0+fD////////6/Piw0ZVoqDX///////////////////////////////////+sz5BAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBDkAYAS48UQJEAQJEAQJEAQJEAQJEATJgQ7vXo////////////////////7vXooMiAUpsYQJEAlcJx////////////////////////////////////7/bpRJMFQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAS48VAFqNM0CRAECRAECRAECRAECRAECRAHqzTtDkwNfoysXesqHJgnCtQEKSA0CRAECRAECRAMvhuv///////////////////////////////////////3qzTkCRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAFyNNQBziWFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBIlgv6/Pj////////////////////////////////////////A26tAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBziWIAAAAAUI4fQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAfLRQ////////////////////////////////////////////+/z5TpkTQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAUY4gAAAAAAAAAHuIcUSQCECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRALPTmv///////////////////////////////////////////////4++aUCRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAEWQCXyIcgAAAAAAAAAAAAB0iGREkAhAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBeoihmpzNmpzNmpzNmpzNurD5zrkRzrkRzrkRzrkRzrkRzrkRzrkRhpCxAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBEkAh1iWUAAAAAAAAAAAAAAAAAAAAAe4hyUI4gQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAQJEAUY4ge4hyAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHOJYVuNNUuPFEORBkCRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAECRAEORBkuPFVuNNXOJYgAAAAAAAAAAAAAAAAAAAAD4AAAAAPgAAOAAAAAAOAAAwAAAAAAYAACAAAAAAAgAAIAAAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAAACAAAgAAAAAAIAADAAAAAABgAAOAAAAAAOAAA+AAAAAD4AAA="
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return -1
VarSetCapacity(Dec, DecLen, 0)
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return -2
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
DllCall("gdiplus\GdipCreateHICONFromBitmap", A_PtrSize ? "UPtr" : "UInt", pBitmap, A_PtrSize ? "UPtr*" : "uint*", hIcon)
DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
Return hIcon
}

WinCollect(hparent:="", wait := 500)                                                                 	{

	winc := Object()
	start := Clock()
	hparent := !hparent ? AlbisWinID() : hparent

	while (PopWinID	:= GetHex(DLLCall("GetLastActivePopup", "uint", hparent)) ) {

		If (PopWinID = hparent)
			break

		title 	:= WinGetTitle(PopWinID)
		class := WinGetClass(PopWinID)
		text 	:= WinGetText(PopWinID)
		winc[hwnd] := {"titte":title , "class":class, "text":text, "count":1}
		If (lastCount<>winc[hwnd].count) {
			If (lastround = 2)
				break
			lastround := !lastround ? 1 : 2
		}

		If (Clock()-start)/1000 > wait
			break
	}

return winc.Count() > 0 ? winc : ""
}

GetScrollInfo(SB, ByRef SI)                                                                                  	{
	VarSetCapacity(SI, 28, 0) ; SCROLLINFO
	NumPut(28, SI, 0, "UInt")
	NumPut(0x17, SI, 4, "UInt") ; SIF_ALL = 0x17
	Return DllCall("User32.dll\GetScrollInfo", "Ptr", This.HWND, "Int", SB, "Ptr", &SI, "UInt")
}

AddTooltip(p1,p2:="",p3="")                                                                              	{

	;------------------------------
	;
	; Function: AddTooltip v2.0
	;
	; Description:
	;
	;   Add/Update tooltips to GUI controls.
	;
	; Parameters:
	;
	;   p1 - Handle to a GUI control.  Alternatively, set to "Activate" to enable
	;       the tooltip control, "AutoPopDelay" to set the autopop delay time,
	;       "Deactivate" to disable the tooltip control, or "Title" to set the
	;       tooltip title.
	;
	;   p2 - If p1 contains the handle to a GUI control, this parameter should
	;       contain the tooltip text.  Ex: "My tooltip".  Set to null to delete the
	;       tooltip attached to the control.  If p1="AutoPopDelay", set to the
	;       desired autopop delay time, in seconds.  Ex: 10.  Note: The maximum
	;       autopop delay time is ~32 seconds.  If p1="Title", set to the title of
	;       the tooltip.  Ex: "Bob's Tooltips".  Set to null to remove the tooltip
	;       title.  See the *Title & Icon* section for more information.
	;
	;   p3 - Tooltip icon.  See the *Title & Icon* section for more information.
	;
	; Returns:
	;
	;   The handle to the tooltip control.
	;
	; Requirements:
	;
	;   AutoHotkey v1.1+ (all versions).
	;
	; Title & Icon:
	;
	;   To set the tooltip title, set the p1 parameter to "Title" and the p2
	;   parameter to the desired tooltip title.  Ex: AddTooltip("Title","Bob's
	;   Tooltips"). To remove the tooltip title, set the p2 parameter to null.  Ex:
	;   AddTooltip("Title","").
	;
	;   The p3 parameter determines the icon to be displayed along with the title,
	;   if any.  If not specified or if set to 0, no icon is shown.  To show a
	;   standard icon, specify one of the standard icon identifiers.  See the
	;   function's static variables for a list of possible values.  Ex:
	;   AddTooltip("Title","My Title",4).  To show a custom icon, specify a handle
	;   to an image (bitmap, cursor, or icon).  When a custom icon is specified, a
	;   copy of the icon is created by the tooltip window so if needed, the original
	;   icon can be destroyed any time after the title and icon are set.
	;
	;   Setting a tooltip title may not produce a desirable result in many cases.
	;   The title (and icon if specified) will be shown on every tooltip that is
	;   added by this function.
	;
	; Remarks:
	;
	;   The tooltip control is enabled by default.  There is no need to "Activate"
	;   the tooltip control unless it has been previously "Deactivated".
	;
	;   This function returns the handle to the tooltip control so that, if needed,
	;   additional actions can be performed on the Tooltip control outside of this
	;   function.  Once created, this function reuses the same tooltip control.
	;   If the tooltip control is destroyed outside of this function, subsequent
	;   calls to this function will fail.
	;
	; Credit and History:
	;
	;   Original author: Superfraggle
	;   * Post: <http://www.autohotkey.com/board/topic/27670-add-tooltips-to-controls/>
	;
	;   Updated to support Unicode: art
	;   * Post: <http://www.autohotkey.com/board/topic/27670-add-tooltips-to-controls/page-2#entry431059>
	;
	;   Additional: jballi.
	;   Bug fixes.  Added support for x64.  Removed Modify parameter.  Added
	;   additional functionality, constants, and documentation.
	;
	;-------------------------------------------------------------------------------
    Static hTT

          ;-- Misc. constants
          ,CW_USEDEFAULT:=0x80000000
          ,HWND_DESKTOP :=0

          ;-- Tooltip delay time constants
          ,TTDT_AUTOPOP:=2
                ;-- Set the amount of time a tooltip window remains visible if
                ;   the pointer is stationary within a tool's bounding
                ;   rectangle.

          ;-- Tooltip styles
          ,TTS_ALWAYSTIP:=0x1
                ;-- Indicates that the tooltip control appears when the cursor
                ;   is on a tool, even if the tooltip control's owner window is
                ;   inactive.  Without this style, the tooltip appears only when
                ;   the tool's owner window is active.

          ,TTS_NOPREFIX:=0x2
                ;-- Prevents the system from stripping ampersand characters from
                ;   a string or terminating a string at a tab character.
                ;   Without this style, the system automatically strips
                ;   ampersand characters and terminates a string at the first
                ;   tab character.  This allows an application to use the same
                ;   string as both a menu item and as text in a tooltip control.

          ;-- TOOLINFO uFlags
          ,TTF_IDISHWND:=0x1
                ;-- Indicates that the uId member is the window handle to the
                ;   tool.  If this flag is not set, uId is the identifier of the
                ;   tool.

          ,TTF_SUBCLASS:=0x10
                ;-- Indicates that the tooltip control should subclass the
                ;   window for the tool in order to intercept messages, such
                ;   as WM_MOUSEMOVE.  If this flag is not used, use the
                ;   TTM_RELAYEVENT message to forward messages to the tooltip
                ;   control.  For a list of messages that a tooltip control
                ;   processes, see TTM_RELAYEVENT.

          ;-- Tooltip icons
          ,TTI_NONE         :=0
          ,TTI_INFO         :=1
          ,TTI_WARNING      :=2
          ,TTI_ERROR        :=3
          ,TTI_INFO_LARGE   :=4
          ,TTI_WARNING_LARGE:=5
          ,TTI_ERROR_LARGE  :=6

          ;-- Extended styles
          ,WS_EX_TOPMOST:=0x8

          ;-- Messages
          ,TTM_ACTIVATE      :=0x401                    ;-- WM_USER + 1
          ,TTM_ADDTOOLA      :=0x404                    ;-- WM_USER + 4
          ,TTM_ADDTOOLW      :=0x432                    ;-- WM_USER + 50
          ,TTM_DELTOOLA      :=0x405                    ;-- WM_USER + 5
          ,TTM_DELTOOLW      :=0x433                    ;-- WM_USER + 51
          ,TTM_GETTOOLINFOA  :=0x408                    ;-- WM_USER + 8
          ,TTM_GETTOOLINFOW  :=0x435                    ;-- WM_USER + 53
          ,TTM_SETDELAYTIME  :=0x403                    ;-- WM_USER + 3
          ,TTM_SETMAXTIPWIDTH:=0x418                    ;-- WM_USER + 24
          ,TTM_SETTITLEA     :=0x420                    ;-- WM_USER + 32
          ,TTM_SETTITLEW     :=0x421                    ;-- WM_USER + 33
          ,TTM_UPDATETIPTEXTA:=0x40C                    ;-- WM_USER + 12
          ,TTM_UPDATETIPTEXTW:=0x439                    ;-- WM_USER + 57

    ;-- Save/Set DetectHiddenWindows
    l_DetectHiddenWindows:=A_DetectHiddenWindows
    DetectHiddenWindows On

    ;-- Tooltip control exists?
    if not hTT
        {
        ;-- Create Tooltip window
        hTT:=DllCall("CreateWindowEx"
            ,"UInt",WS_EX_TOPMOST                       ;-- dwExStyle
            ,"Str","TOOLTIPS_CLASS32"                   ;-- lpClassName
            ,"Ptr",0                                    ;-- lpWindowName
            ,"UInt",TTS_ALWAYSTIP|TTS_NOPREFIX          ;-- dwStyle
            ,"UInt",CW_USEDEFAULT                       ;-- x
            ,"UInt",CW_USEDEFAULT                       ;-- y
            ,"UInt",CW_USEDEFAULT                       ;-- nWidth
            ,"UInt",CW_USEDEFAULT                       ;-- nHeight
            ,"Ptr",HWND_DESKTOP                         ;-- hWndParent
            ,"Ptr",0                                    ;-- hMenu
            ,"Ptr",0                                    ;-- hInstance
            ,"Ptr",0                                    ;-- lpParam
            ,"Ptr")                                     ;-- Return type

        ;-- Disable visual style
        ;   Note: Uncomment the following to disable the visual style, i.e.
        ;   remove the window theme, from the tooltip control.  Since this
        ;   function only uses one tooltip control, all tooltips created by this
        ;   function will be affected.
		;;;;;        DllCall("uxtheme\SetWindowTheme","Ptr",hTT,"Ptr",0,"UIntP",0)

        ;-- Set the maximum width for the tooltip window
        ;   Note: This message makes multi-line tooltips possible
        SendMessage TTM_SETMAXTIPWIDTH,0,A_ScreenWidth,,ahk_id %hTT%
        }

    ;-- Other commands
    if p1 is not Integer
        {
        if (p1="Activate")
            SendMessage TTM_ACTIVATE,True,0,,ahk_id %hTT%

        if (p1="Deactivate")
            SendMessage TTM_ACTIVATE,False,0,,ahk_id %hTT%

        if (InStr(p1,"AutoPop")=1)  ;-- Starts with "AutoPop"
            SendMessage TTM_SETDELAYTIME,TTDT_AUTOPOP,p2*1000,,ahk_id %hTT%

        if (p1="Title")
            {
            ;-- If needed, truncate the title
            if (StrLen(p2)>99)
                p2:=SubStr(p2,1,99)

            ;-- Icon
            if p3 is not Integer
                p3:=TTI_NONE

            ;-- Set title
            SendMessage A_IsUnicode ? TTM_SETTITLEW:TTM_SETTITLEA,p3,&p2,,ahk_id %hTT%
            }

        ;-- Restore DetectHiddenWindows
        DetectHiddenWindows %l_DetectHiddenWindows%

        ;-- Return the handle to the tooltip control
        Return hTT
        }

    ;-- Create/Populate the TOOLINFO structure
    uFlags:=TTF_IDISHWND|TTF_SUBCLASS
    cbSize:=VarSetCapacity(TOOLINFO,(A_PtrSize=8) ? 64:44,0)
    NumPut(cbSize,      TOOLINFO,0,"UInt")              ;-- cbSize
    NumPut(uFlags,      TOOLINFO,4,"UInt")              ;-- uFlags
    NumPut(HWND_DESKTOP,TOOLINFO,8,"Ptr")               ;-- hwnd
    NumPut(p1,          TOOLINFO,(A_PtrSize=8) ? 16:12,"Ptr")
        ;-- uId

    ;-- Check to see if tool has already been registered for the control
    SendMessage
        ,A_IsUnicode ? TTM_GETTOOLINFOW:TTM_GETTOOLINFOA
        ,0
        ,&TOOLINFO
        ,,ahk_id %hTT%

    l_RegisteredTool:=ErrorLevel

    ;-- Update the TOOLTIP structure
    NumPut(&p2,TOOLINFO,(A_PtrSize=8) ? 48:36,"Ptr")
        ;-- lpszText

    ;-- Add, Update, or Delete tool
    if l_RegisteredTool
        {
        if StrLen(p2)
            SendMessage
                ,A_IsUnicode ? TTM_UPDATETIPTEXTW:TTM_UPDATETIPTEXTA
                ,0
                ,&TOOLINFO
                ,,ahk_id %hTT%
         else
            SendMessage
                ,A_IsUnicode ? TTM_DELTOOLW:TTM_DELTOOLA
                ,0
                ,&TOOLINFO
                ,,ahk_id %hTT%
        }
    else
        if StrLen(p2)
            SendMessage
                ,A_IsUnicode ? TTM_ADDTOOLW:TTM_ADDTOOLA
                ,0
                ,&TOOLINFO
                ,,ahk_id %hTT%

    ;-- Restore DetectHiddenWindows
    DetectHiddenWindows %l_DetectHiddenWindows%

    ;-- Return the handle to the tooltip control
    Return hTT
    }

ObjFullyClone(obj)                                                                                            	{	 	;-- ersetzt das alte Objekt mit dem neuen????

	; https://www.autohotkey.com/board/topic/103411-cloned-object-modifying-original-instantiation/

    nobj := ObjClone(obj)
    for k,v in nobj
        if IsObject(v)
            nobj[k] := ObjFullyClone(v)
    return nobj
}

;}



;
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Includes
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
#Include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DATUM.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PDFHelper.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PraxTT.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Window.ahk

#Include %A_ScriptDir%\..\..\lib\acc.ahk
#Include %A_ScriptDir%\..\..\lib\class_JSON.ahk
#Include %A_ScriptDir%\..\..\lib\class_cJSON.ahk
#Include %A_ScriptDir%\..\..\lib\class_Loaderbar.ahk
#Include %A_ScriptDir%\..\..\lib\class_LV_Colors.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_all.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\lib\Sift.ahk
;}





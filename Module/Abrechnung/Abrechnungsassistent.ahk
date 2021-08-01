; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;                                      	💰	Addendum Abrechnungsassistent	💰
;
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;      Funktion: 				 ⚬	prüft auf Abrechenbarkeit von Vorsorgeuntersuchungs- und beratungsziffern (01732,01745/56, 01740
;										, 01747, 01731) nach den allgemeinen EBM-Regeln entsprechend Alter und Abrechnungsfrequenz
;                                	 ⚬ schaut ob die Chronikerziffern (eingegeben wurden)
;
;
;		Hinweis:					Gui friert manchmal ein.
;										<--- aus class_LV_Colors.ahk -->
;										Um den Verlust von Gui-Ereignissen und Meldungen zu vermeiden, muss der Message-Handler eventuell
; 										auf "kritisch" gesetzt werden. Dies kann erreicht werden, indem die Instanzeigenschaft 'Critical' auf den
; 										gewünschten Wert gesetzt wird (z. B. MyInstance.Critical := 100). Neue Instanzen sind standardmäßig
;                                   	auf 'Critical, Off' eingestellt. Auch wenn es manchmal nötig ist, können ListViews oder das gesamte Gui
;                                   	unter (un?)bestimmten Umständen nicht mehr reagieren, wenn Critical gesetzt ist und die ListView ein g-Label hat.
;
;		Abhängigkeiten:	siehe includes
;
;	                    			Addendum für Albis on Windows
;                        			by Ixiko started in September 2017 - last change 10.07.2021 - this file runs under Lexiko's GNU Licence
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
; weitere Regeln: 	ICD T14.x -> Wundversorgung
;							Laborleistung wie BSG, Glucose, INR

  ; Einstellungen   	;{
	#NoEnv
	#Persistent
	#MaxMem                    	, 256
	#KeyHistory                  	, 0

	SetBatchLines                	, -1
	SetWinDelay                    	, -1
	SetControlDelay            	, -1

	SetTitleMatchMode        	, 2	              	;Fast is default
	SetTitleMatchMode        	, Fast        		;Fast is default

	ListLines                        	, On

  ; startet Windows Gdip
   	If !(pToken:=Gdip_Startup()) {
		MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
		ExitApp
	}

  ; Tray Icon erstellen
	;~ If (hIconLabJournal := Create_Laborjournal_ico())
    	;~ Menu, Tray, Icon, % "hIcon: " hIconLabJournal
	;}

  ; Variablen      	;{
	global PatDB
	global Addendum
	global DBAccess
	global habr

  ; Albis Datenbankpfad / Addendum Verzeichnis
	RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
	RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)

	Addendum                 	:= Object()
	Addendum.Dir            	:= AddendumDir
	Addendum.Ini              	:= AddendumDir "\Addendum.ini"
	Addendum.DBPath      	:= AddendumDir "\logs'n'data\_DB"
	Addendum.DBasePath  	:= AddendumDir "\logs'n'data\_DB\DBase"
	Addendum.AlbisDBPath	:= AlbisPath "\DB"
	Addendum.compname	:= StrReplace(A_ComputerName, "-")                                	; der Name des Computer auf dem das Skript läuft

	ReIndex := false

	SciTEOutput()

  ; Patientendatenbank
	infilter := ["NR", "VNR", "VKNR", "PRIVAT", "ANREDE", "ZUSATZ", "NAME", "VORNAME", "GEBURT", "TITEL", "PLZ", "ORT", "STRASSE"
				, 	"TELEFON", "GESCHL"
				, 	"SEIT", "ARBEIT", "HAUSARZT", "GEBFREI", "LESTAG", "GUELTIG", "FREIBIS", "KVK", "TELEFON2", "TELEFAX", "LAST_BEH"
				, 	"DEL_DATE", "MORTAL",	"HAUSNUMMER", "RKENNZ"]
	PatDB := ReadPatientDBF(Addendum.AlbisDBPath, infilter)

  ; aktuelles Quartal im rechnerischen Format YYQQ berechnen (z.B. 2102)
	aktuellesQuartal := GetQuartalEx(A_DD "." A_MM "." A_YYYY, "QQYY")

  ; Addendum.ini Verfügbarkeit prüfen
	If !FileExist(Addendum.ini) {
		MsgBox, % "Einstellungen konnten nicht geladen werden.`n" Addendum.ini
		ExitApp
	}

  ; meine Notizmethode
	IniRead, ChangeCave, % Addendum.ini, Abrechnungshelfer, Abrechnungsassistent_CaveBearbeiten, % "Nein"
	If InStr(ChangeCave, "ERROR")
		IniWrite, % (ChangeCave := "Nein"), % Addendum.ini, Abrechnungshelfer, Abrechnungsassistent_CaveBearbeiten

	Addendum.ChangeCave := (ChangeCave = "Ja") ? true : false
	SciTEOutput(ChangeCave "`n" Addendum.ini)

  ; letztes Arbeitsquartal aus der Addendum.ini lesen
	IniRead, letztesQuartal, % Addendum.ini, Abrechnungshelfer, Abrechnungsassistent_letztesArbeitsquartal
	letztesQuartal := Trim(letztesQuartal)
	If !RegExMatch(letztesQuartal, "^0[1-4][0-9]{2}$")
		ReIndex := true, Abrechnungsquartal := letztesQuartal := aktuellesQuartal

  ; Abrechnungsquartale vergleichen und Nutzer bei Unterschieden befragen
	If (letztesQuartal <> aktuellesQuartal) {
		MsgBox, 0x4	, % StrReplace(A_ScriptName, ".ahk")
							,	% "Wollen Sie mit dem zuletzt bearbeiteten`n"
							. 	"Quartal <" letztesQuartal "> weiter arbeiten?`n"
							.	"(Wenn nicht wird das aktuelle Quartel verwendet.)"
		IfMsgBox, Yes
			Abrechnungsquartal := letztesQuartal
		IfMsgBox, No
			Abrechnungsquartal := aktuellesQuartal, ReIndex := true

	}
	;}

  ; Albis nach vorne holen
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
		static EventInfo, empty := "- - - - - - -"
		static FSize  	:= A_ScreenHeight > 1080 ? 11 : 10
		static LVw  	:= A_ScreenHeight > 1080 ? 800 : 720
		static LVCol	:= A_ScreenHeight > 1080 ? [60, 130, 30, 30, 40, 40, 38, 40, 45, 110, 80 ,80] : [55, 130, 30, 30, 35, 35, 38, 40, 45, 90, 80 ,80]
		static abrUD_o
		static abrQuartal, abrQuartalold

		local hAlbis, winpos

		abrQuartal := abrUD_o := Abrechnungsquartal

	; Euro Preise der Gebührennummern
		Euro := {"#01731"        	: 15.82
					, "#01732"       	: 35.82
					, "#01745"       	: 27.80
					, "#01746"       	: 22.96
					, "#01737"       	: 6.34
					, "#01740"       	: 12.75
					, "#01747"       	: 9.01
					, "#03320"       	: 14.28
					, "#03321"       	: 4.39
					, "KVU"              	: 15.82
					, "GVU"             	: 35.82
					, "HKS1"            	: 27.80
					, "HKS2"            	: 22.96
					, "GB1"             	: 12.57
					, "GB2"             	: 19.36
					, "Colo"             	: 12.75
					, "Aorta"            	: 9.01
					, "Chron1"         	: 14.46
					, "Chron2"         	: 4.450
					, "Pausch0-4"    	: 25.03
					, "Pausch5-18"  	: 15.80
					, "Pausch19-54"	: 12.68
					, "Pausch55-75"	: 16.46
					, "Pausch76-"		: 22.25
					, "PauschZ"    		: 15.35
					, "SPES"          	  	: 6.34}


	; Komplexe laden/erstellen
		SplashText("Vorbereitung", "Abrechnungsziffern aller Quartale laden ...")
		VSEmpf	:= VorsorgeKomplexe(Abrechnungsquartal, ReIndex)                           		; erstellt das DBAccess Objekt

	; Farbeinstellungen des Albisnutzer für Kürzel aus der Datenbank lesen
		befkuerzel := DBAccess.BefundKuerzel()

	; EBM Regeln laden
		SplashText("Vorbereitung", "lade EBM Regeln ...")
		VSRule 	:= DBAccess.GetEBMRule("Vorsorgen")
		If !IsObject(VSRule)
			throw A_ThisFunc ": EBM Vorsorgeregeln konnten nicht übermittelt werden!"

	;}

	; letzte Fensterposition laden   	;{
		workpl	:= WorkPlaceIdentifier()
		IniRead, winpos, % Addendum.ini, % Addendum.compname, % "Abrechnungsassistent_Pos_" workpl
		If (InStr(winpos, "ERROR") || StrLen(winpos) = 0)
			winpos := "xCenter yCenter "
	;}

	; Albis Fensterposition              	;{
		AlbisFensterPosition:
		wp := PosStringToObject(winpos)
		GuiH	:= 1080
		If (hAlbis := AlbisWinID()) {

				ap := GetWindowSpot(hAlbis)

			; Monitorgrößen in welchem sich Albis befindet
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

	;}

	; andere Variablen                	;{
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
		SelectionColor1	:= "0x0078D6"
		abrTabNames	:= "Ziffern-Guru|ICD-Krauler|Großer-Geriater"
	;}

	; Gui                                       	;{

		Gui, abr: new	, % "AlwaysOnTop HWNDhabr -DPIScale"
		Gui, abr: Color	, % "c" gcolor , % "c" tcolor
		Gui, abr: Margin, 5, 5
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal c" TextColor1 , Futura Bk Bt
		Gui, abr: Add  	, Tab     	, % "x1 y1  	w" 640 " h" 25 " HWNDabrHTab vabrTabs gabr_Tabs"             	, % abrTabNames

	;-: TAB 1 - ZIFFERN-GURU                                                        	;{
		Gui, abr: Tab  	, 1

	  ;-: Kopf                                                                                  ;{
		Gui, abr: Font	, % "s" FSize  " q5 Normal c" TextColor1  , Futura Bk Bt
		Gui, abr: Add	, Text        	, % "xm+5 ym+30    	BackgroundTrans     	"                                            	, % "Quartal"
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal cBlack"
		Gui, abr: Add	, Edit           	, % "x+5 yp+-9 w" (FSize-1)*5 " 	              	vabrQ           	gabrHandler"	, % Abrechnungsquartal
		GuiControlGet, cp, abr: Pos, abrQ
		ButtonH :=  cpH-2
		Gui, abr: Add	, Button        	, % "x" cpX+cpW " y" cpY+1 " h" cpH-2 "	      	vabrQUp       	gabrHandler"   	, % "▲"
		Gui, abr: Add	, Button        	, % "x+0 h" cpH-2 "                        	         	vabrQDown   	gabrHandler"   	, % "▼"

		Gui, abr: Font	, % "s" FSize " q5 Normal c" TextColor1 , Futura Bk Bt
		Gui, abr: Add	, Text        	, % "x+30     	BackgroundTrans             	"                                               	, % "mögliches EBM plus: "
		Gui, abr: Add	, Text        	, % "x+0       	BackgroundTrans  				vEBMEuroPlus"                        	, % "00000.00 Euro  "
		Gui, abr: Add	, Text        	, % "x+20     	BackgroundTrans             	"                                               	, % "Krankenscheine: "
		Gui, abr: Add	, Text        	, % "x+0       	BackgroundTrans  				vEBMScheine"                         	, % VSEmpf["_Statistik"].KScheine "  "
	;}

	  ;-: Section 1                                                                           	;{
		Gui, abr: Add  	, Progress   	, % "x14 y+10 w290 h2 -E0x20000 vabrPS1 HWNDabrHPS1 c" tcolorP " Background" tcolorP, 100
	  ;}

	  ;-: Refresh / ReIndex                                                              	;{
		Gui, abr: Font	, % "s" FSize+6 " q5 Normal c" tcolorI
		Gui, abr: Add	, Text        	, % "xm+5 y+5 BackgroundTrans vabrSymbol1 "                             	, % "♻"
		GuiControlGet, cp, abr: Pos, abrSymbol1
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal c" TextColor1
		Gui, abr: Add	, Button        	, % "x+2 y" cpY+2 " h" ButtonH "       	vabrReFresh   	gabrHandler"   	, % "Liste auffrischen"

		Gui, abr: Font	, % "s" FSize+6 " q5 Normal c" tcolorI
		Gui, abr: Add	, Text        	, % "x+15 y" cpY-1 " BackgroundTrans"                                                           	, % "🔄"
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal c" TextColor1
		Gui, abr: Add	, Button        	, % "x+2 y" cpY+2 " h" ButtonH "          	       	vabrReIndex   	gabrHandler"    	, % "Index erneuern"

		Gui, abr: Font	, % "s" FSize+6 " q5 Normal c" tcolorI
		Gui, abr: Add	, Text        	, % "x+20 y" cpY-1 " vabrHide BackgroundTrans"                                           	, % "☑"
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal c" TextColor1
		Gui, abr: Add	, Button        	, % "x+2 y" cpY+2 " h" ButtonH "         	       	vabrBearbeitet  	gabrHandler"    	, % "Markierte ausblenden"

		GuiControlGet, cp, abr: Pos, abrHide
		Gui, abr: Font	, % "s" FSize+18 " q5 Normal c" tcolorI
		Gui, abr: Add	, Text        	, % "x+10 y" cpY-12 " BackgroundTrans"                                                   	, % "◻"
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal c" TextColor1 , Futura Bk Bt
		Gui, abr: Add	, Button        	, % "x+2  y" cpY+2 " h" ButtonH "       	vabrUnHide  	gabrHandler"    	, % "alle einblenden"

		WinSet, Redraw,, % "ahk_id " abrHPS1

	;}

	  ;-: Optionen                                                                         	;{
		CBOpt := "BackgroundTrans Checked"
		Gui, abr: Add	, Checkbox   	, % "xm+5 	y+10 "             	CBOpt " 	vabrYoungStars	gabrHandler"    	, % "Alter 18 bis 34 anzeigen"
		GuiControlGet, cp, abr: Pos, abrYoungStars
		Gui, abr: Add	, Checkbox   	, % "x+20 "                        	CBOpt "	vabrOldStars 	gabrHandler"     	, % "Alter ab 35 anzeigen"
		Gui, abr: Add	, Checkbox   	, % "x+20 "                        	CBOpt " 	vabrAngels    	gabrHandler"    	, % "Verstorbene anzeigen"
		Gui, abr: Add	, Text        	, % "x+40           	BackgroundTrans  	vabrLVRows"                              	, % "[0000]"
	;}

	  ;-: Patientenliste / Listview                                                    	;{
		LVOptions3 	:= "vabrLV gabrLVHandler HWNDabrhLV Checked Grid -E0x200 -LV0x10 AltSubmit -Multi "
		LVLabels3 	:= "ID|Name, Vorname|Alt.|G|GVU|HKS|KVU|Colo|Aorta|Chroniker|GB|Pauschale"
		LVwidth     	:= 0
		For idx, w in LVCol
			LVwidth += w

		Gui, abr: Font	, % "s" FSize-2 " q5 Normal cBlack"
		Gui, abr: Add	, ListView    	, % "xm+5 y+5 w" LVwidth+20 " r15 " LVOptions3                                                   	, % LVLabels3
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
		Gui, abr: Add	, Text        	, % "xm+5 y+1 BackGroundTrans"                                                       	, % "gewählter Patient: "
		Gui, abr: Font	, % "s" FSize+1 " q5 Bold  c" TextColor1
		Gui, abr: Add	, Text        	, % "x+0 w" FSize*20 "  	vabrPat"                                                              	, % " "
		;}

	  ;-: Checkboxen für Untersuchungen                                        	;{
		abrCheckbox := Object()
		Gui, abr: Font	, % "s" FSize-1   	" q5 Normal c" TextColor1
		Loop % VSRule.VSFilter.MaxIndex() {
			feat := VSRule.VSFilter[A_Index].exam
			If !abrCheckbox.haskey(feat) {
				Gui, abr: Add	, Checkbox   	, % (A_Index = 1 ? "xm+5 y+1" : "x+25")  " BackgroundTrans vabr_" feat " HWNDabrHwnd" 	, % feat
				GuiControlGet, cp, abr: Pos, % "abr_" feat
				abrCheckbox[feat] := cpX
			}
		}
		;}

	  ;-: zeigt frühest möglichen Abrechnungstag an                      	;{
		floop := 1
		Gui, abr: Font	, % "s" FSize-3    	" q5 Normal c" TextColor1
		For feat, X in abrCheckbox {
			Gui, abr: Add	, Text   	, % (floop = 1 ? "xm+5 y+2" : "x" X " y" cpY)  " BackgroundTrans vabr_Date" feat " " 	, % " [00.00.0000] "
			GuiControlGet, cp, abr: Pos, % "abr_Date" feat
			GuiControl, abr:, % "abr_Date" feat, % ""
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
	 ;}

	  ;-: Zahl der angezeigten Patienten                                        	;{
		GuiControl, abr:, abrLVRows, % "[" LV_GetCount() "]"
	;}

	  ;-: Euro Summe aller EBM Gebühren anzeigen                       	;{
		; (vorausgesetzt sie übernehmen die Vorschläge zu 100%)
		GuiControl, abr:, EBMEuroPlus, % Round(EBMPlus, 2) " €"
		GuiControl, abr: Focus, EBMEuroPlus
		;}

	  ;-: Quartalkalender für spätere Abrechnungstage hinzufügen  	;{
		QJahr := ("20" SubStr(Abrechnungsquartal, 3, 2) > A_YYYY ? "19" : "20") SubStr(Abrechnungsquartal, 3, 2)
		QuartalsKalender(SubStr(Abrechnungsquartal, 1, 2), QJahr , 0, "abr", TextColor1)
	  ;}



	;}

	;-: TAB 2 - ICD-Krauler                                                            	;{
		Gui, abr: Tab  	, 2

		GuiControlGet, cp, abr: Pos, % "MyCalendar"
		Gui, abr: Add	, Edit  	, % "xm+5 ym+35  w" (FSize-1)*5 " 	                        	vabrQ2           	gabrHandler"	, % Abrechnungsquartal
		Gui, abr: Add	, Button	, % "x+5                                                                   	vabrICDKrauler gabrHandler"  	, % "nach Chroniker kraulen lassen"
		Gui, abr: Add	, Edit 	, % "xm+5 y+10 w610 h" cpY+cpH " vabrChrOut"                                                       	, % "~~~~~"

	;}

	;-: TAB 3 - Der große Geriater                                                	;{
		Gui, abr: Tab  	, 3
	;}

		Gui, abr: Tab

	;-: GUI - Position anpassen                                                     	;{
		abrP := GetWindowSpot(habr)
		GuiControlGet, cp, abr: Pos, abrLV
		DeltaH := abrP.H - abrP.CH
		GuiW := cpX + cpW + 5

	  ; Progressbalken
		GuiControl, abr: Move, abrPS1, % "x4 w"  GuiW - 5
		GuiControl, abr: Move, abrTabs	, % "w" GuiW - 5

		WinSet, Style 	, -0x1       	, % "ahk_id " abrHPS1
		WinSet, ExStyle	, -0x20000	, % "ahk_id " abrHPS1
		WinSet, Style 	, -0x1       	, % "ahk_id " abrHPS2
		WinSet, ExStyle 	, -0x20000	, % "ahk_id " abrHPS2

		GuiControlGet, cp, abr: Pos, abrChrOut
		GuiControlGet, dp, abr: Pos, abrTabs
		GuiControl, abr: Move, abrChrOut, % "h" cpH-DeltaH

		Gui, abr: Show, % winpos " NoActivate", Abrechnungsassistent


		SplashTextOff
	;}

	;~ #IfWinActive Abrechnungsassistent ahk_class AutoHotkeyGUI
		;~ Hotkey, Enter             	, abrHotkeyHandler
		;~ Hotkey, NumPadEnter 	, abrHotkeyHandler
		;~ Hotkey, Up                 	, abrHotkeyHandler
		;~ Hotkey, Down          	 , abrHotkeyHandler
	;~ #If

return ;}

abrHotkeyHandler:                     	;{

	Gui, abr: ListView, abrLV
	selectedRow := LV_GetNext(0, "F")
	LV_GetText(PatID    	, selectedRow, 1)
	LV_GetText(PatName	, selectedRow, 2)

	Switch A_ThisHotkey {

		 Case "Enter":
			AlbisActivate(2)
			AlbisAkteOeffnen(PatName, PatID)

		Case "Down":
			ControlSend,, {Down}	, % "ahk_id " AbrhLV

		Case "Up":
			ControlSend,, {Up}   	, % "ahk_id " AbrhLV
	}

return ;}

abrLVHandler:                            	;{

		If (RegExMatch(A_GuiEvent, "(K|f|C)") || A_EventInfo = 0)
			return

	; Patient lesen
		Gui, abr: ListView, abrLV
		LV_GetText(tmpPatID   		, A_EventInfo, 1)
		LV_GetText(tmpPatName	, A_EventInfo, 2)

		If !(A_GuiEvent = "A")
			If (!tmpPatID && !tmpPatName) || (lvLastPatID = tmpPatID) || !IsObject(VSEmpf[tmpPatID])
				return
		else if (A_GuiEvent = "A" && !IsObject(VSEmpf[tmpPatID]))
			return

		lvLastPatID := PatID := tmpPatID
		lvLastPatName := PatName := tmpPatName

	; Karteikarte nach Doppelklick anzeigen
		If RegExMatch(A_GuiEvent, "(DoubleClick|A)") {
			AlbisAkteOeffnen(PatName, PatID)
		}

	; Namen und Features setzen
		GuiControl, abr:, abrPat, % "[" PatID "] " PatName
		For feat, featposX in abrCheckbox              	{
			GuiControl, abr: Hide     	, % "abr_" feat
			GuiControl, abr:             	, % "abr_" feat    	, 0
			GuiControl, abr:             	, % "abr_Date" feat, % ""
		}

		For featIDX, feat in VSEmpf[PatID].Vorschlag 	{
			RegExMatch(feat, "i)^(?<Name>[a-z]+)\-*(?<PossDate>\d+)*", Feat)
			GuiControl, abr: Show     	, % "abr_" featName
			GuiControl, abr:            	, % "abr_" FeatName     	, 1
			GuiControl, abr:             	, % "abr_Date" FeatName	, % FeatPossDate ? " > " ConvertDBASEDate(FeatPossDate) : ""
		}

	; Behandlungstage des Patienten anzeigen
		Behandlungstage(PatID)

return ;}

abrHandler:                              	;{

	If RegExMatch(A_GuiControl, "(abrQUp|abrQDown)")                                 	{         	; Quartalzähler vor und zurück

		Gui, abr: Submit, NoHide

		abrQQYY	:= abrQ                                                ; Edit-Feld mit Quartal
		abrQQ 	:= SubStr(abrQQYY, 1, 2) + 0
		abrYY    	:= SubStr(abrQQYY, 3, 2) + 0
		QUD     	:= A_GuiControl = "abrQUp" ? 1 : -1
		abrQQ 	:= abrQQ + QUD

		If (abrQQ = 0)
			abrQQ := 4, abrYY := abrYY + QUD
		else if (abrQQ = 5)
			abrQQ := 1, abrYY := abrYY + QUD
		abrYY       	:= abrYY > 99 ? 0 : abrYY < 0 ? 99 : abrYY
		QQYY      	:= SubStr( "0" abrQQ, -1) SubStr( "0" abrYY,  -1)
		GuiControl, abr:, abrQ, % QQYY

	}
	else if RegExMatch(A_GuiControl, "(abrReIndex|abrRefresh)")                        	{         	; übernimmt das eingestellte Abrechungsquartal

		Gui, abr: Submit, NoHide
		abrQuartal := SubStr("0" abrQ, -3)
		abrQuartalold := abrQuartal
		SaveAbrechnungsquartal(abrQuartal)
		SplashText("Vorbereitung", "Abrechnungsziffern aller Quartale laden ...")
		VSEmpf	:= VorsorgeKomplexe(abrQuartal, (A_GuiControl = "abrReIndex" ? true : false))
		gosub abrVorschlaege

	}
	else if (A_GuiControl = "abrBearbeitet")                                                      	{         	; bearbeitete Karteikarten als versteckt markieren

			Gui, abr: Default
			Gui, abr: ListView, abrLV

		; alle Abgehakten registrieren
			row := rowF := 0, RowsToDelete := []
			while (row := LV_GetNext(row, "C")) {
				rowF := !rowF ? row : rowF
				LV_GetText(PatID, row, 1)
				RowsToDelete.InsertAt(1, {"Delete":row, "Hide":PatID})
			  ; Karteikarte schließen bringt Geschwindigkeitsvorteile
				AlbisAkteSchliessen2(PatID "/")         ; damit es nicht als ein Handle erkannt wird
			}

		; Markieren und Löschen (von unten nach oben))
			;GuiControl, -Redraw, % abrhLV
			For idx, to in RowsToDelete {
				VSEmpf[to.Hide].Hide := true
				LV_Delete(to.Delete)
			}

		; verändertes Objekt speichern
			FileSize := DBAccess.SaveVorsorgeKandidaten(VSEmpf, abrQuartal)

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
			Gui, abr: Default
			Gui, abr: ListView, abrLV
			GuiControl, abr:, abrLVRows, % "[" LV_GetCount() "]"



	}
	else if (A_GuiControl = "abrUnHide")                                                         	{	     	; versteckte Karteikarten wieder anzeigen

		; Objektflags auf false setzen
			For PatID, vs in VSEmpf
				If VSEmpf[PatID].Hide
					VSEmpf[PatID].Hide := false

		; verändertes Objekt speichern
			FSize := DBAccess.SaveVorsorgeKandidaten(VSEmpf, abrQuartal)

		; Inhalt neu anzeigen über universal Label
			gosub abrVorschlaege

	}
	else if RegExMatch(A_GuiControl, "(abrYoungStars|abrAngels|abrOldStars)") 	{      	; Filterhandler
		gosub abrVorschlaege
	}
	else if (A_GuiControl = "abrICDKrauler")                                                       	{
		Gui, abr: Submit, NoHide
		chronischKrauler(abrQ2)
	}

return ;}

abrVorschlaege:                        	;{

		SplashText("bitte warten ...", "gefilterte Daten werden aufgelistet ...")

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
		For PatID, vs in VSEmpf {

		  ; nicht anzeigen bei diesen Bedingungen
			If !RegExMatch(PatID, "^\d+$")
				continue
			else if VSEmpf[PatID].Hide
				continue
			else if VSEmpf[PatID].Hide
				continue
			else if !ShowYoungStars	&& (vs.Alter < 35)
				continue
			else if !ShowOldStars 	&& (vs.Alter > 34)
				continue
			else if !ShowAngels    	&& PatDB[PatID].Mortal
				continue

		  ; Vorschläge vorschlagen ∶-)   	;{
			e := []
			For idx, exam in vs.Vorschlag {
				e[idx] := false
				RegExMatch(exam, "i)(?<rdination>[A-Z]+)\-*(?<Date>\d+)*", o)
				Switch ordination {
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
			EBMPlus += vs.chronisch 	= 1 	? EBM.Chron1 	: 0
			EBMPlus += vs.chronisch 	= 2 	? EBM.Chron2 	: 0
			EBMPlus += vs.chronisch 	= 3 	? EBM.Chron2 + EBM.Chron1 : 0
			EBMPlus += vs.geriatrisch = 1 	? EBM.GB1   	: 0
			EBMPlus += vs.geriatrisch = 2 	? EBM.GB2   	: 0
			EBMPlus += vs.geriatrisch = 3 	? EBM.GB2 + EBM.GB1 : 0
			;PauschGO := vs.Alter
			;~ EBMPlus += vs.Pauschale = 1 	? EBM.PauschGO	: 0
			;~ EBMPlus += vs.Pauschale = 2 	? EBM.PauschZ	: 0
			;~ EBMPlus += vs.Pauschale	= 3 	? EBM.PauschGO + EBM.PauschZ : 0
		  ;}

		    f := 	vs.chronisch  	? 1 : 0
			f +=	vs.chronischQ 	? 1 : 0
			Loop 5
				f += e[A_Index] ? 1 : 0

			If (f > 0) {

				; Zeile anfügen
					Mortal 	:= PatDB[PatID].Mortal ? "♱" : ""

					Chron	:= vs.chronisch	= 1 	? "03220" : vs.chronisch 	= 2 ? "03221" : vs.chronisch	= 3 ? "03220+.21"
					Chron	.= vs.chronischQ		? " VQ"	vs.chronischQ (vs.chronischQ = 1 ? " fehlt" : " fehlen") 	: ""
					GB   	:= vs.geriatrisch = 1	? "03360" : vs.geriatrisch	= 2 ? "03362" : vs.geriatrisch= 3 ? "03360+.62"
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

				; Zellen färben
					;~ For eIdx, isEBM in e
						;~ If isEBM
						 ;~ CLV.Cell(LV_GetCount(), 4+eIdx	, "0x" tcolorM, 0x000000)

					;~ If vs.chronisch || vs.chronischQ
						 ;~ CLV.Cell(LV_GetCount(), 10      	, "0x" tcolorM, 0x000000)

					;~ If vs.geriatrisch
						 ;~ CLV.Cell(LV_GetCount(), 11      	, "0x" tcolorM, 0x000000)

					;~ If vs.Pauschale
						 ;~ CLV.Cell(LV_GetCount(), 12      	, "0x" tcolorM, 0x000000)


			}

		}

		Loop, % LV_GetCount() {
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
		GuiControl, abr:, abrLVRows, % "[" LV_GetCount() "]"

	;-: Euro Summe aller EBM Gebühren anzeigen (vorausgesetzt sie übernehmen die Vorschläge zu 100%)
		GuiControl, abr:, EBMEuroPlus, % Round(EBMPlus, 2) " €"
		GuiControl, abr: Focus, EBMEuroPlus

return ;}

abr_Tabs:                                  	;{

return ;}

abrGuiClose:
abrGuiEscape:                           	;{
	SaveGuiPos(habr, "Abrechnungsassistent")

ExitApp ;}
}

QuartalsKalender(Quartal, Jahr, Urlaubstage, guiname:="QKal", TextColor1:="") 	{     	;-- zeigt den Kalender des Quartals

	; blendet Feiertage nach Bundesland aus (### ist nicht programmiert! ####)

	;--------------------------------------------------------------------------------------------------------------------
	; Variablen
	;--------------------------------------------------------------------------------------------------------------------;{
		global QKal, hmyCal, hCal, habr, gcolor, tcolor, CalQK
		global Cok, Can, Hinweis, MyCalendar, CalBorder, MyCalDay, MyCalStart, MyCalBTage
		global VSEmpf, abrCheckbox, VSRule, CV, PUpX, PUpY, abrLvBT, abrhLvBT, StaticBT, abrhLV
		global KK, lvLastPatID, lvLastPatName, lvPatID, lvPatname, BLV, abrYear
		global BefKuerzel

		static gname, lastPrgDate

		If (gname <> guiname)
			gname := guiname
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	Berechnung der anzuzeigenden Monate
	;--------------------------------------------------------------------------------------------------------------------;{
		Q := Array()
		FirstMonth:= SubStr("0" ((Quartal-1)*3)+1 , -1)
		lastmonth := SubStr("0" FirstMonth+2    	, -1)
		FormatTime, begindate	, % Jahr . FirstMonth	. "01"                                                	, yyyyMMdd
		FormatTime, lastdate 	, % Jahr . lastmonth	. DaysInMonth(Jahr . lastmonth . "01")	, yyyyMMdd
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	Kalender Gui
	;--------------------------------------------------------------------------------------------------------------------;{
		CalOptions := " AltSubmit +0x50010051 vMyCalendar HWNDhmyCal gCalHandler"
		lp 	:= GetWindowSpot(AbrhLV)
		ap	:= GetWindowSpot(AlbisWinID())

		GuiControlGet, cp, abr: Pos, CalQK
		Gui, abr: Font     	, % "s11 w600 c" TextColor1, Futura Md Bk
		Gui, abr: Add     	, Text		, % "xm+5 y+1 vCalQK"                             	, % "QUARTALSKALENDER"
		GuiControlGet, cp, abr: Pos, CalQK

		Gui, abr: Font     	, % "s8 Normal italic c" TextColor1
		Gui, abr: Add     	, Text		, % "x+2 y" cpY+1                                        	, % "[" Quartal SubStr(Jahr, -1) "]"

		Gui, abr: Font     	, % "s13 w600 c" TextColor1
		Gui, abr: Add     	, Text		, % "x+1 y" cpY-4                                         	, % ":"

		Gui, abr: Font     	, % "s10 Normal c" TextColor1
		Gui, abr: Add     	, Text		, % "x+8 y" cpY+1                                        	, % SubStr(ConvertDBASEDate(begindate), 1,6) "-" ConvertDBASEDate(lastdate)
		Gui, abr: Add     	, MonthCal, % "xm+5 y+2 W-3 4 w521"  CalOptions   	, % begindate
		GuiControlGet, cp, abr: Pos, MyCalendar
		cX 	:= cpX + cpW + 5
		cY		:= cpY + cpH

		Gui, abr: Font     	, % "s10 Normal c" TextColor1
		Gui, abr: Add     	, Text		, % "x+5"  " BackgroundTrans"                       	, % "gewählter Tag: "
		Gui, abr: Add     	, Text		, % "x+5  vMyCalDay"                                  	, % "--.--.----           "
		Gui, abr: Add     	, Button		, % "x" cX " y+5 vMyCalStart gCalHandler"   	, % "Ziffernauswahl eintragen"
		GuiControl, abr: Disable, MyCalStart

		Gui, abr: Font     	, % "s10 Normal c" TextColor1
		Gui, abr: Add     	, Text  		, % "x" cX " y" cY-10 "BackGroundTrans vStaticBT" 	, % "Behandlungstage"
		GuiControlGet, cp, abr: Pos, staticBT
		GuiControl, abr: Move, staticBT, % "x" 5+lp.W-cpW

		LVOptions 	:= " -Hdr vabrLvBT gCalHandler HWNDabrhLvBT -E0x200 -LV0x10 AltSubmit -Multi"
		LVColumns	:= "Datum|Kürzel|Inhalt"
		Gui, abr: Font     	, % "s10 Normal cBlack"
		Gui, abr: Add     	, ListView  	, % "xm+5 y+2 w" lp.W " h" ap.CH+1*ap.BH-cpY-cpH " "  LVOptions                                          	, % LVColumns

		WinSet, Style, 0x50010051, % "ahk_id " hmyCal
		LV_ModifyCol(1, 50)
		LV_ModifyCol(2, 65)
		LV_ModifyCol(3, lp.CW-115)
		WinSet, Style, 0x5021800D, % "ahk_id " abrhLvBT   ; No_HScroll

		BLV := New LV_Colors(abrhLvBT)
		;BLV.Critical := 100
		BLV.SelectionColors("0x0078D6")

	; Kalenderfarben
		SendMessage, 0x100A	, 0, 0x666666 	,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 1, 0xAA0000 	,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 2, 0x0101AA 	,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 3, 0x01FF01   ,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 4, % "0x" RGBtoBGR(gcolor),, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 5, 0xFFAA99, SysMonthCal321, % "ahk_id " habr
		Sleep 100
		SendMessage, 0x100A	, 1, 0xAA0000 	,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 2, 0x0101AA 	,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 3, 0x01FF01 	,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 4, % "0x" RGBtoBGR(gcolor),, % "ahk_id " hmyCal

Return ;}

CalHandler: ;{

		Gui, abr: Default
		Gui, abr: Submit, NoHide

	; angeklicketes Datum sichern                                                                    	;{
		If (A_GuiControl = "MyCalendar") {
			If RegExMatch(A_GuiEvent,"i)(Normal|I|1)") {
				DateChoosed 	:= ConvertDBASEDate(StrSplit(MyCalendar,"-").1)
				GuiControl, abr:, MyCalDay, % DateChoosed
				GuiControl, abr: Enable, MyCalStart
			}
		}
	;}

	; das Eintragen starten
		If (A_GuiControl = "MyCalStart")
			ZiffernauswahlEintragen(DateChoosed, lvLastPatID, lvLastPatName)

		If (A_GuiControl = "abrLvBT") && RegExMatch(A_GuiEvent, "(DoubleClick|A)") {
			Gui, abr: ListView, abrLvBT
			LV_GetText(DateChoosed, A_EventInfo, 1)
			;~ WorkInProgress()
			ZiffernauswahlEintragen(DateChoosed abrYear, lvLastPatID, lvLastPatName)
			;~ WorkInProgress()
		}

return ;}


}

Behandlungstage(PatID)                                                                                      	{

		global VSEmpf, abrLV, abrLvBT, abrhLvBT, BLV, BefKuerzel, tcolorBG1, tcolorBG2, abrYear

		Gui, abr: ListView, abrLvBT

		BLV.Clear(1, 1)
		LV_Delete()

		BehTage := VSEmpf[PatID].BehTage
		For DATUM, Leistung in BehTage {
			abrYear := SubStr(ConvertDBASEDate(DATUM), 7, 4)
			break
		}

		cols := [], altFlag := true, rows := 0
		For DATUM, Leistung in BehTage {

			DATUM := SubStr(ConvertDBASEDate(DATUM), 1, 6)
			BTrow := 1, rows ++
			For row, col in Leistung {
				If (col.KZL = "lle")
					continue
				cols.1 	:= BTRow = 1 ? DATUM : ""
				cols.2	:= col.KZL
				cols.3	:= col.INH
				BTRow ++
				LV_Add("", cols*)
				bgColor 	:= "0x" (altFlag ? tcolorBG1 : tcolorBG2)
				txtColor		:= RGBtoBGR(Format("0x{:06X}", BefKuerzel[col.KZL].Color))
				BLV.Row(LV_GetCount() 	, bgColor, 0x000000)
				BLV.Cell(LV_GetCount(), 2	, bgColor, txtColor)
				BLV.Cell(LV_GetCount(), 3	, bgColor, txtColor)
			}
			altFlag := !altFlag

		}

		If (row := LV_GetNext(0, "F")) {
			LV_Modify(row, "-Focus")
			LV_Modify(row, "-Select")
		}

		WinSet, Style, 0x5021800D, % "ahk_id " abrhLvBT   ; No_HScroll

return abrYear
}

ZiffernauswahlEintragen(DateChoosed, lvPatID, lvPatName)                                 	{   	;-- Karteikartenautomatisierung

		global KK, VSEmpf, VSRule, abrCheckbox, habr, habrLV

	; die Karteikeite des ausgewählten Patienten geöffnet                                   	;{
		PatID := AlbisAktuellePatID()
		If (PatID <> lvPatID)
			If !AlbisAkteOeffnen(lvPatName, lvPatID) {
				MsgBox, 0x1, % A_ScriptName, % "Die Karteikarte von " lvPatname "`nließ sich nicht öffnen.", 3
				return
			}
	;}

	; ist aktuelle eine Karteikarte geöffnet?                                                          	;{
		AlbisAnzeige := AlbisGetActiveWindowType()
		If !InStr(AlbisAnzeige, "Karteikarte") {
			PraxTT("Es wird keine Karteikarte in Albis angezeigt`nWähle zunächst einen Patienten mit Doppelklick!", "5 1")
			return 0
		}
	;}

	; Karteiblatt anzeigen                                                                                 	;{
		If !RegExMatch(AlbisAnzeige, "i)Karteikarte\s*$")
			If !AlbisKarteikarteZeigen() {
					MsgBox, 0x1, % A_ScriptName, % "Es konnte nicht auf das Karteikartenblatt umgeschaltet werden.", 3
					return
				}
	;}

	; Patient der aktuellen Karteikarte ist gelistet und hat offene Komplexziffern? 	;{
		PrgDate   	:= AlbisLeseProgrammDatum()

		If (!VSEmpf[PatID].Vorschlag.Count() && !VSEmpf[PatID].chronisch) {
			PraxTT(	"Bei Patient [" PatID "] " PatDB[PatID].Name ", " PatDB[PatID].VORNAME
					. 	" fehlt keine Eintragung.`nWähle einen anderen Patienten!", "5 1")
			return 0
		}
	;}

	; gewähltes Datum liegt an einem Wochenende 	                                        	;{
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

	; EBM-Ziffern zusammen stellen                                                                  	;{
		EBMFees       	:= GetEBMFees(PatID)
		EBMToAdd     	:=""
		HKSFormular 	:= false
		If (VSEmpf[PatID].Vorschlag.Count() > 0)
			For feat, Xpos in abrCheckbox {
				GuiControlGet, ischecked, abr:, % "abr_" feat
				EBMNr := ischecked ? EBMFees[feat] "-" : ""
				If RegExMatch(EBMNr, "(01745|01746)") || RegExMatch(feat, "HKS")
					HKSFormular := true
				EBMToAdd .= EBMNr
			}

		EBMtoAdd := RTrim(RegExReplace(EBMtoAdd, "[\-]{2,}", "-"), "-")
		;SciTEOutput("EBM String: " EBMtoAdd " einzutragen am " DateChoosed)
	;}

	; Ziffern eintragen                                                                                        	;{

		; aktivieren
			AlbisActivate(1)

		; PopUp-Fenster schliessen
			;AlbisCloseLastActivePopups(AlbisWinID())

		; aktuelles Programmdatum lesen
			lastPrgDate := AlbisSetzeProgrammDatum(DateChoosed)

		; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
		; Karteikarte anzeigen
		; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
			If InStr(AlbisGetActiveWindowType(), "Karteikarte")
				If !AlbisKarteikarteZeigen() {
					MsgBox, Die Kartekarte ließ sich nicht Anzeigen
					return
				}

		; Karteikarte aktivieren (anklicken)
			AlbisKarteikarteAktivieren()

		; Karteikartenzeile lesen
			KKarte	:= AlbisGetActiveControl("content")

		; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
		; ist kein Eingabefokus eröffnet, wird die Eingabe eröffnet
		; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
			If !IsObject(Kkarte) {
				SendInput, {INSERT}
				Sleep 50
				KKarte	:= AlbisGetActiveControl("content")    ; jetzt kann die Zeile gelesen werden
			}

		; Zeilendatum lesen
			ZDatum 	:= AlbisLeseZeilenDatum()
			If (ZDatum <> DateChoosed) {
				AlbisKarteikarteFocus("Datum")
				If !VerifiedSetText(KK.DATUM, DateChoosed, "ahk_class optoAppClass") {
					SendInput, {Home}
					SendInput, {LShift Down}{End}{Shift Up}
					Sleep 100
					Send, % DateChoosed
					Sleep 100
					SendInput, {TAB}
				}
			}

		; Tastatureingabe ins Kürzelfeld
			AlbisKarteikarteFocus("Kuerzel")

		; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
		; lko
		; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
			If (StrLen(KKarte.kuerzel) = 0) {

			 If (Kkarte.FName <> "Kuerzel")                                                 ; FName = Focused Name
				If !VerifiedClick(KK.kuerzel,,, AlbisWinID())
					If !VerifiedSetFocus(KK.kuerzel, AlbisWinID(),,, false) {
						ControlGetPos, ctrlX, ctrlY,,, % KK.KUERZEL, % "ahk_class optoAppClass"
						ControlClick, % KK.KUERZEL, % "ahk_class OptoAppCLass",, Left, 1, NA
						ControlGetFocus, fc, A
						If !InStr(KK.KUERZEL, fc) || !InStr(fc, KK.KUERZEL)
							SciTEOutput("nix is :" fc)
							;MsgBox, bekomme das Kürzelfeld nicht fokussiert.
					}

			ControlGetFocus, fcs, A
			If (fcs<>KK.kuerzel) {
				SciTEOutput(fcs " <> " KK.KUERZEL)
			}
			If !VerifiedSetText(KK.kuerzel, "lko", AlbisWinID) {
				Send, % "lko"
				Sleep, 100
			}

		}

		; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
		; und zum nächsten Feld vorrücken [INHALT]
		; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
			SendInput, {TAB}
			Sleep 100
			Loop {
				ControlGetFocus, fcs, % "ahk_class OptoAppClass"
				If (fcs = KK.Inhalt) || (A_Index > 40)
					break
				Sleep 100
			}

		; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
		; String senden
		; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
			If StrLen(EBMtoAdd) = 0
				MsgBox, EMBtoAdd  ist leer!
			If !VerifiedSetText(KK.Inhalt, EBMtoAdd, AlbisWinID)
				SendInput, % "{Raw}" EBMtoAdd
			Sleep 100
			SendInput, {TAB}
			Sleep 100
			Loop {
				ControlGetFocus, fcs, % "ahk_class OptoAppClass"
				If (fcs <> KK.Inhalt) {
				  ; Karteikarte kann Eingaben aufnehmen, dann durch Senden eines Escape Tastendruckes beenden
					while AlbisKarteiKarteEingabeStatus() {
						SendInput, {Escape}
						Sleep 250
						If (A_Index > 5) {
							breakWhile := true
							break
						}
					}
					break

			 ; Abbruch wenn es zu lange dauert
				} else If  (A_Index > 40)
					break
					Sleep 50
			}

	; Hautkrebscreening-Formular                                                                    	;{
		If RegExMatch(EBMtoAdd, "(01745|01746)")
			if (hEHKS := AlbisMenu(34505, "Hautkrebsscreening - Nichtdermatologe ahk_class #32770", 3))
				AlbisHautkrebsScreening()   ; alles leer lassen, alles super!
	;}

	; auf letztes Datum zurücksetzen und Ausnahmeindikationen korrigieren       	;{
		AlbisSetzeProgrammDatum(lastPrgDate)
		ausIndi := AlbisAusIndisKorrigieren()
	;}

	;}

	; CAVE Dokumentation - Standard: ausgestellt                                              	;{
		If Addendum.ChangeCave {

			examDate := RegExReplace(DateChoosed, "\d+\.(\d+)\.\d\d(\d+)", "$1^$2")
			If RegExMatch(EBMtoAdd, "O)(01732|0174[56]|01731)\-*.*(01732|0174[56]|01731)*\-*.*(01732|0174[56]|01731)*", exams) {
				GVUHKSKVU := RegExMatch(EBMtoAdd, "01732") ? "GVU"  	: ""
				GVUHKSKVU .= RegExMatch(EBMtoAdd, "01745") ? "HKS"  	: RegExMatch(EBMtoAdd, "01746") ? "/HKS" : ""
				GVUHKSKVU .= RegExMatch(EBMtoAdd, "01731") ? "/KVU" 	: ""
				GVUHKSKVU := ", " LTrim(GVUHKSKVU, "/") " " examDate
			}
			COLO	:= RegExMatch(EBMtoAdd, "01740") ? "01740"  	: ""
			AORTA	:= RegExMatch(EBMtoAdd, "01747") ? ", (AORTA)" 	: ""

			ClipBoard := Trim(Colo . AORTA . GVUHKSKVU , ", ")

			If !(hCave := Albismenu(32778, "Cave! von ahk_class #32770", 4))
				return 0
			Lines := AlbisGetCave(false)
			For lineNR, line in StrSplit(Lines, "`n") {
				If RegExMatch(line, "i)(Td|Pneu|GSI|MMR|Polio|Pert|FSME)")
					Impfung .= lineNr "|"
				If RegExMatch(line, "i)(GVU|AORTA|HKS|KVU|01740|GB)")
					Vorsorge .= lineNr "|"
			}

			AlbisCaveSetCellFocus(hCave, lineNr-1, "Beschreibung")

			while WinExist("Cave! von ahk_class #32770") {
				Sleep 300
			}

			If !Impfung {

				infoDate := "31.12." SubStr(DateChoosed, 7, 4)

			; aktuelles Programmdatum lesen
				lastPrgDate := AlbisSetzeProgrammDatum(infoDate)
				If AlbisZeilendatumSetzen(infoDate) {

				  ; Karteikartenzeile lesen
					Send, % "Impf"
					Sleep 200
					SendInput, % "{TAB}"
					Sleep 200
					;AlbisKarteikarteFocus("Inhalt")
					Send, % "Impfausweis mitbringen"
					Sleep 200
					SendInput, % "{TAB}"
					Sleep 1000

				}

			}

		}
	;}

	; MsgBX - melde das Du das fertig hast                                                       	;{
		If IsObject(ausIndi) {
			Alt 	:= ausIndi.Alt
			RegExReplace(Alt, "\-",, AltC)
			Neu	:= ausIndi.neu
			RegExReplace(Neu, "\-",, NeuC)
			inter := Alt ? (AltC > 0 ? "n" : "") " ersetzt mit" : "n erstmalig eingetragen"
			msg := "Ausnahmeindikation" (AltC > 0 ? "en" : "") ":`n" Alt "`nwurde" inter "`n" Neu
		}
		else
			msg := "Ausnahmeindikation" (AltC > 0 ? "en" : "") ":`n" Alt "`nwurde" (AltC > 0 ? "n" : "") " nicht geändert"
			MsgBox	, 0x1024
					, % A_ScriptName
					, % "Bitte alles überprüfen!`n`n" msg
					, 20
	;}



return
}

GetEBMFees(PatID)                                                                                               	{

	global	VSEmpf, VSRule

	If !IsObject(VSRule := DBAccess.GetEBMRule("Vorsorgen"))
		throw A_ThisFunc ": EBM Vorsorgeregeln konnten nicht übermittelt werden!"

	EBMFees := Object()
	For featNr, feat in VSEmpf[PatID].Vorschlag {
		feat := RegExReplace(feat, "\-\d+$", "")
		EBMFees[feat] := VSRule[feat].GO
	}

	If EBMFees.haskey("HKS") && EBMFees.haskey("GVU")
		EBMFees["HKS"] := "01746"
	else if EBMFees.haskey("HKS") && !EBMFees.haskey("GVU")
		EBMFees["HKS"] := "01745"

	;~ JSONData.Save("C:\tmp\tmp.json", EBMFees, true,, 1, "UTF-8")

return EBMFees
}

VorsorgeKomplexe(Abrechnungsquartal, ReIndex:=false, SaveResults:="")            	{     	;-- berechnet die Vorsorgekandidaten

		static PathVSCandidates

	; nur bei leerem SaveResults-Parameter ReIndex verwenden
		SaveResults	:= SaveResults ? SaveResults : ReIndex
		quartal      	:= LTrim(Abrechnungsquartal, "0")

		saveVSCandidates := false
		PathVSCandidates := Addendum.DBasePath "\VorsorgeKandidaten-" quartal ".json"

	; Albisdatenbank Objekt erstellen
		If !IsObject(DBAccess)
			DBAccess 	:= new AlbisDb(Addendum.AlbisDBPath, 0)

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; nicht neu indizieren, dann gespeicherte Daten laden
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If !ReIndex && FileExist(PathVSCandidates) {
			SplashText("bitte warten ...", "lade Vorsorge Kandidaten")
			VSEmpf := JSONData.Load(PathVSCandidates, "", "UTF-8")
			SplashText("bitte warten ...", "Vorsorge Kandidaten geladen")
		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; keine Daten vorhanden oder neu indizieren wurde übermittelt
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	  ; Krankenscheine:         	Abrechnungsart: ohne Überweiser, Notfallscheine oder Private
	  ; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		If !IsObject(VSEmpf) || ReIndex || !FileExist(Addendum.DBasePath "\KScheine_" quartal ".json") {
			SplashText("bitte warten ... 1/5", "Abrechnungsscheine des Quartals werden gelesen ...")
			KScheine	:= DBAccess.KScheine(Abrechnungsquartal, "Abrechnung", false, SaveResults)
		}

	  ; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	  ; Behandlungstage:     	des Quartals für jeden Patienten zusammenstellen
	  ; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		If !IsObject(VSEmpf) || ReIndex || !FileExist(Addendum.DBasePath "\BehTage_" quartal ".json") {
			SplashText("bitte warten... 2/5", "Behandlungstage im Quartal ermitteln ...")
			BehTage	:= DBAccess.Behandlungstage(Abrechnungsquartal, "Abrechnung", false, true)
		}
		else
			BehTage	:= JSONData.Load(Addendum.DBasePath "\BehTage_" quartal ".json", "", "UTF-8")

	  ; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	  ; Vorsorgedaten:         	von allen Patienten die Vorsorgeziffern und Abrechnungstage finden
	  ; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		If !IsObject(VSEmpf) || ReIndex || !FileExist(Addendum.DBasePath "\Vorsorgen.json") {
			SplashText("bitte warten ... 3/5", "EBM Ziffern aller Quartale werden gesammelt ...")
			VSDaten 	:= DBAccess.VorsorgeDaten(Abrechnungsquartal, ReIndex, SaveResults)
		}

	  ; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	  ; Vorsorgekandidaten:  	finden
	  ; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		If !IsObject(VSEmpf) || ReIndex || !FileExist(PathVSCandidates) {

			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
			; Vorsorgekandidatendatei laden um die verarbeiteten Datensätze zu behalten
			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
				VSEmpfOld := ""
				If FileExist(PathVSCandidates)
					VSEmpfOld := JSONData.Load(PathVSCandidates, "", "UTF-8")

			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
			; Vorsorgekandidaten neu erstellen
			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
				SplashText("bitte warten ... 4/5", "EBM Ziffern aller Quartale werden gefiltert ...")
				VSEmpf := DBAccess.VorsorgeFilter(VSDaten, Abrechnungsquartal, "Abrechnung", false, (ReIndex ? false : true))

			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
			; Daten aus alter Kandidatendatei mit neuen Daten vergleichen
			; (im Moment nur key: "Hide" - manuell ausgeblendete Patienten)
			; SciTEOutput(" [" PatID "] " PatDB[PatID].NAME ", " PatDB[PatID].VORNAME)
			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
				If IsObject(VSEmpfOld) {
					saveVSCandidates := true
					For PatID, vs in VSEmpfOld {
						VSEmpf[PatID].Hide := VSEmpfOld[PatID].Hide
					}
				}

			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
			; Behandlungstage zuorden
			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
				If IsObject(BehTage) {
					saveVSCandidates := true
					For PatID, Tage in BehTage {
						If !IsObject(VSEmpf[PatID].BehTage)
							VSEmpf[PatID].BehTage := BehTage[PatID]
					}
				}

			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
			; weitere Abrechnungsprüfungen
			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
				;~ For PatID in VSEmpf {
					;~ leistungen := diagnosen := ""
					;~ For BehDatum, ke in VSEmpf[PatID].BehTage {

						;~ If ke.KZL =

					;~ }
				;~ }

		}


	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; VSEmpf sichern wenn Änderungen gemacht wurden
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If saveVSCandidates
			JSONData.Save(PathVSCandidates, VSEmpf, true,, 1, "UTF-8")

return VSEmpf
}

chronischKrauler(Abrechnungsquartal)                                                                 	{     	;-- sucht Patienten für die Abrechnung der Chroniker-Ziffern

	; sucht alle ICD Abrechnungsdiagnosen eines Quartals heraus und vergleicht diese mit einer ICD Liste
	; 2 Datenbankdateien müssen für die Suche ausgelesen werden: 1.BEFUND.dbf und 2.BEFTEXTE.dbf
	; die Eingrenzung der Suche innerhalb der BEFUND.dbf über das Abrechnungsquartal ist nicht anwendbar,
	; Albis hinterlegt dort nicht manchmal die Quartalbezeichnung meist steht dort eine 0 oder eine andere Zahl
	; das Abrechnungsquartal wird in ein Datum in Form eines RegEx-String umgewandelt aus
	; 0320 wird 2020(07|08|09)\d\d. Damit sind alle Tage des Quartals erfasst.

		static CRICD

	; ICD Liste für den Vergleich laden
		If !IsObject(CRICD)
			CRICD := JSONData.Load(Addendum.Dir "\include\Daten\ICD-chronisch_krank.json", "", "UTF-8")

	; dafür sorgen das Zugriff auf die Albisdatenbanken besteht (class AlbisDB in Addendum_DB.ahk)
		If !IsObject(DBAccess)
			DBAccess 	:= new AlbisDb(Addendum.AlbisDBPath, 0)

	; Quartal in RegExString umwandeln
		Quartal := QuartalTage({"aktuell": Abrechnungsquartal})
		rxQuarterDays := "rx:" Quartal.Jahr "(" Quartal.MBeginn "|" Quartal.MMitte "|" Quartal.MEnde "|" ")\d\d"

	; BEFUND.dbf Suche nach Kürzel 'dia' im Datumsbereich, Erfassung der Patienten, Diagnosen und den
	; Verbindungen zu den Einträgen in der BEFTEXT.dbf
		dia := DBAccess.GetDBFData("BEFUND", {"KUERZEL":"dia", "DATUM":rxQuarterDays}, ["PATNR", "DATUM", "INHALT", "TEXTDB"])

	; Ergebnisse der Suche anzeigen
		t:=""
		GuiControl, QS:, abrChrOut, % t
	  For Index, set in dia {
			t .= set.PATNR "`t" SubStr(ConvertDBASEDate(set.Datum), 1, 6) " " set.INHALT "`n"
			GuiControl, QS:, abrChrOut, % t
	  }

}

SaveAbrechnungsquartal(Abrechnungsquartal)                                                   	{  	;-- speichert das aktuell eingestellte Quartal

	IniWrite, % Abrechnungsquartal, % Addendum.ini, Abrechnungshelfer, Abrechnungsassistent_letztesArbeitsquartal

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

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Hilfsfunktionen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - -;{
SaveGuiPos(hwnd, Fenstername)                                                                        	{

	win 	:= GetWindowSpot(hwnd)
	If (win.X > 0 && win.Y > 0) {
		workpl	:= WorkPlaceIdentifier()
		winPos 	:= "x" win.X " y" win.Y
		IniWrite, % winpos	, % Addendum.Ini, % Addendum.compname, %  Fenstername "_Pos_" workpl
	}

}

WorkPlaceIdentifier()                                                                                           	{  	;-- Identifikationsstring für die Arbeitsumgebung
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

MessageWorker(InComing)                                                                                	{   	;-- verarbeitet die eingegangen Nachrichten

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

SplashText(title, msg)                                                                                            	{
	SplashTextOn, 300, 20, % title, % msg
	hSTO := WinExist("Vorbereitung ahk_class AutoHotkey2")
	WinSet, Style 	, 0x50000000, % "ahk_id " hSTO
	WinSet, ExStyle	, 0x00000208, % "ahk_id " hSTO
}

EnableBlur(hWnd)                                                                                              	{

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

Thread_CoronaEBM:        ;{  alle Coronaabstrich, alle Behandlungsziffern Corona (88240, 32006) finden

	aDB          	:= new AlbisDB(Addendum.AlbisDBPath)
	abrScheine  	:= aDB.Abrechnungsscheine(Abrechnungsquartal, false, true)
	withCorona  	:= aDB.GetDBFData(  "BEFUND"
									     			  , {	"KUERZEL":"rx:(dia|labor)"
														, 	"INHALT":"rx:.*(COVIPC|CoV|SARS|COVID|Spezielle|COVSEQ|SARS\-CoV\-2\-Mutation|U07\.[12]|U99\.\d|N501Y|delH69|E484K).*"
														, 	"DATUM":"rx:2021\d\d\d\d"}
								    				  , ["PATNR", "DATUM", "KUERZEL", "INHALT"],,1 )

	; nur die Patienten-ID's zusammentragen
		COVID := Object()
		For idx, set in withCorona {
			If !COVID.haskey(set.PATNR) {
				COVID[set.PATNR] := {"Matches":[], "GNROK":false}
			}
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
						If GNRMatched = 0
							break
					}

			If (GNRMatched = 0)
				COVID[PatID].GNROK := true

		}

	; Karteikarte für Karteikarte anzeigen
		NotFirstPat := 1
		For PatID, data in COVID {
			PatName := PatDB[PatID].NAME ", " PatDB[PatID].VORNAME
			If (NotFirstPat > 1) {
				MsgBox, 0x1024, % "Weiter?", % "nächster Patient (" NotFirstPat "/" COVID.Count() "):`n[" PatID "] " PatName "`nAbrechnung: " (data.GNROK ? "ist ok":"da fehlt was")
				IfMsgBox, No
					break
			}
			AlbisAkteOeffnen(PatName, PatID)
			NotFirstPat ++
		}

		aDB:= ""
		abrScheine := withCorona := COVID := ""

return   ;}
;}

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
#Include %A_ScriptDir%\..\..\lib\class_LV_Colors.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\lib\Sift.ahk
;}





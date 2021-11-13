; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;                                      	💰	Addendum Abrechnungsassistent	💰
;
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;      Funktion: 				 ⚬	prüft auf Berechnungsfähigkeit von Vorsorgeuntersuchungs- und beratungsziffern (01732,01745/56, 01740,
;										01747, 01731) nach den EBM-Regeln entsprechend Alter und Abrechnungsfrequenz
;                                	 ⚬ Abrechnungsüberprüfung für Chronikerpauschalen (eingegeben wurden)
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
;                        			by Ixiko started in September 2017 - last change 07.10.2021 - this file runs under Lexiko's GNU Licence
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

	If (hIconAbrAssist := Create_Abrechnungsassistent_ico())
		Menu, Tray, Icon, % "hIcon: " hIconAbrAssist

	;}

  ; Variablen      	;{
	global PatDB, Addendum, DBAccess, habr

  ; Albis Datenbankpfad / Addendum Verzeichnis
	RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
	RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)

  ; Addendum.ini Verfügbarkeit prüfen
	If !(workini := IniReadExt(AddendumDir "\Addendum.ini")) {
		MsgBox, % "Einstellungen konnten nicht geladen werden.`n" Addendum.ini
		ExitApp
	}


  ; Addendum
	Addendum                     	:= Object()
	Addendum.Dir                	:= AddendumDir
	Addendum.Ini                  	:= AddendumDir "\Addendum.ini"
	Addendum.DBPath          	:= AddendumDir "\logs'n'data\_DB"
	Addendum.DBasePath      	:= AddendumDir "\logs'n'data\_DB\DBase"
	Addendum.AlbisDBPath   	:= AlbisPath "\DB"
	Addendum.compname    	:= StrReplace(A_ComputerName, "-")                                	; der Name des Computer auf dem das Skript läuft
	Addendum.ThousandYear	:= "20"
	Addendum.IndisToRemove	:= IniReadExt("Abrechnungshelfer", "AusIndis_ToRemove")
	Addendum.IndisToHave 	:= IniReadExt("Abrechnungshelfer", "AusIndis_ToHave")
	Addendum.letztesQuartal 	:= IniReadExt("Abrechnungshelfer", "Abrechnungsassistent_letztesArbeitsquartal")
	Addendum.ChangeCave 	:= IniReadExt("Abrechnungshelfer", "Abrechnungsassistent_CaveBearbeiten", "Nein")
	Addendum.KeinBMP       	:= IniReadExt("Addendum", "KeinBMP", "Ja")
	Addendum.ShowIndis       	:= IniReadExt("Addendum", "Zeige_Ausnahmeindikationen", "Nein")

	ReIndex := false

  ; Patientendatenbank
	SplashText(StrReplace(A_ScriptName, ".ahk"), "Lade Patientendatenbank...")
	infilter := ["NR", "VNR", "VKNR", "PRIVAT", "ANREDE", "ZUSATZ", "NAME", "VORNAME", "GEBURT", "TITEL", "PLZ", "ORT", "STRASSE"
				, 	"TELEFON", "GESCHL"
				, 	"SEIT", "ARBEIT", "HAUSARZT", "GEBFREI", "LESTAG", "GUELTIG", "FREIBIS", "KVK", "TELEFON2", "TELEFAX", "LAST_BEH"
				, 	"DEL_DATE", "MORTAL",	"HAUSNUMMER", "RKENNZ"]
	PatDB := ReadPatientDBF(Addendum.AlbisDBPath, infilter)
	SplashText(StrReplace(A_ScriptName, ".ahk"), "Patientendatenbank geladen")

  ; Albisdatenbank Objekt erstellen
	DBAccess 	:= new AlbisDb(Addendum.AlbisDBPath, 0)

  ; aktuelles Quartal ins Format YYQQ konventieren (z.B. 2102)
	aktuellesQuartal := GetQuartalEx(A_DD "." A_MM "." A_YYYY, "QQYY")
	letztesQuartal := Trim(Addendum.letztesQuartal)
	If !RegExMatch(letztesQuartal, "^0[1-4][0-9]{2}$") {
		Abrechnungsquartal := letztesQuartal := aktuellesQuartal
		SaveAbrechnungsquartal(Abrechnungsquartal)
		ReIndex := true
	}
	Abrechnungsquartal := letztesQuartal

	;}

  ; Albis nach vorne holen
	SplashText(StrReplace(A_ScriptName, ".ahk"))
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
		static FSize, LVw, LVCol
		static abrUD_o, c:=0
		static abrQuartal

		local hAlbis, winpos

		FSize  	:= A_ScreenHeight > 	1080 ? 11 : 10
		LVw  	:= A_ScreenHeight > 	1080 ? 800 : 720
		If A_ScreenHeight > 	1080
			LVCol	:= FSize>10 	? 	[60, 120, 35, 25, 40, 40, 38, 40, 45, 130, 60, 80] : 	[60, 130, 30, 30, 40, 40, 38, 40, 45, 90, 60, 60]
		else
			LVCol	:= FSize>10  	?	[55, 130, 30, 30, 35, 35, 38, 40, 45, 90, 60 , 60] : [55, 110, 30, 40, 35, 35, 38, 40, 45, 130, 60 , 80]

		abrQuartal := abrUD_o := Abrechnungsquartal

	; Euro Preise der Gebührennummern
		Euro := {"#01731"        	: 15.82
					, "#01732"           	: 35.82
					, "#01745"           	: 27.80
					, "#01746"           	: 22.96
					, "#01737"           	: 6.34
					, "#01740"           	: 12.75
					, "#01747"           	: 9.01
					, "#03320"           	: 14.28
					, "#03321"           	: 4.39
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
					, "Pausch0-4"        	: 25.03
					, "Pausch5-18"      	: 15.80
					, "Pausch19-54" 	: 12.68
					, "Pausch55-75" 	: 16.46
					, "Pausch76-" 		: 22.25
					, "PauschZ"    		: 15.35
					, "SPES"          	  	: 6.34}

	; Komplexe laden/erstellen
		SplashText("Vorbereitung", "Abrechnungsziffern aller Quartale laden ...")
		VSEmpf	:= VorsorgeKomplexe(Abrechnungsquartal, ReIndex)                           		; lädt oder erstellt alle notwendigen Daten

	; Farbeinstellungen des Albisnutzer für Kürzel aus der Datenbank lesen
		befkuerzel := DBAccess.BefundKuerzel()

	; EBM Regeln laden
		SplashText("Vorbereitung", "lade EBM Regeln ...")
		VSRule 	:= DBAccess.GetEBMRule("Vorsorgen")
		If !IsObject(VSRule)
			throw A_ThisFunc ": EBM Vorsorgeregeln konnten nicht übermittelt werden!"

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

			; Monitordimensionen auf dem Albis geöffnet ist
				ap := GetWindowSpot(hAlbis)
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

	; andere Variablen                 	;{
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
		abrTabNames	:= "Vorsorge-Guru|ICD-Krauler|Großer-Geriater|Einstellungen"
	;}

	;}

	; Gui                                       	;{

		Gui, abr: new	, % "AlwaysOnTop HWNDhabr -DPIScale"  ;
		Gui, abr: Color	, % "c" gcolor , % "c" tcolor
		Gui, abr: Margin, 5, 5
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal c" TextColor1 , Futura Bk Bt
		Gui, abr: Add  	, Tab     	, % "x1 y1  	w" LvW " h" 25 " HWNDabrHTab vabrTabs gabr_Tabs"             	, % abrTabNames

	;-: TAB 1 - VORSORGE-GURU                                                 	;{
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
		Gui, abr: Add	, Text        	, % "xm+5 y+5 BackgroundTrans vabrSymbol1 "                                           	, % "♻"
		GuiControlGet, cp, abr: Pos, abrSymbol1
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal c" TextColor1
		Gui, abr: Add	, Button        	, % "x+2 y" cpY+2 " h" ButtonH "       	vabrRefresh   	gabrHandler"               	, % "Liste auffrischen"

		Gui, abr: Font	, % "s" FSize+6 " q5 Normal c" tcolorI
		Gui, abr: Add	, Text        	, % "x+15 y" cpY-1 " BackgroundTrans"                                                         	, % "🔄"
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal c" TextColor1
		Gui, abr: Add	, Button        	, % "x+2 y" cpY+2 " h" ButtonH "          	       	vabrReindex   	gabrHandler"   	, % "Index erneuern"

		Gui, abr: Font	, % "s" FSize+6 " q5 Normal c" tcolorI
		Gui, abr: Add	, Text        	, % "x+20 y" cpY-1 " vabrHide BackgroundTrans"                                           	, % "☑"
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal c" TextColor1
		Gui, abr: Add	, Button        	, % "x+2 y" cpY+2 " h" ButtonH "         	       	vabrBearbeitet  	gabrHandler"   	, % "Markierte ausblenden"

		GuiControlGet, cp, abr: Pos, abrHide
		Gui, abr: Font	, % "s" FSize+18 " q5 Normal c" tcolorI
		Gui, abr: Add	, Text        	, % "x+10 y" cpY-12 " BackgroundTrans"                                                       	, % "◻"
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal c" TextColor1 , Futura Bk Bt
		Gui, abr: Add	, Button        	, % "x+2  y" cpY+2 " h" ButtonH "       	vabrUnHide  	gabrHandler"        	, % "alle einblenden"

		WinSet, Redraw,, % "ahk_id " abrHPS1

	;}

	  ;-: Optionen                                                                         	;{
		CBOpt := "BackgroundTrans Checked"
		Gui, abr: Add	, Checkbox   	, % "xm+5 	y+10 "             	CBOpt " 	vabrYoungStars	gabrHandler"    	, % "Alter 18 bis 34 anzeigen"
		GuiControlGet, cp, abr: Pos, abrYoungStars
		Gui, abr: Add	, Checkbox   	, % "x+20 "                        	CBOpt "	vabrOldStars 	gabrHandler"     	, % "Alter ab 35 anzeigen"
		Gui, abr: Add	, Checkbox   	, % "x+20 "                        	CBOpt " 	vabrAngels    	gabrHandler"    	, % "Verstorbene anzeigen"
		Gui, abr: Add	, Text        	, % "x+40           	BackgroundTrans  	vabrLVRows"                                 	, % "[0000]"
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
	 ;}

	  ;-: Zahl der angezeigten Patienten                                        	;{
		GuiControl, abr:, abrLVRows, % "[" LV_GetCount() "]"
	;}

	  ;-: Euro Summe aller EBM Gebühren anzeigen                       	;{
		; (vorausgesetzt sie übernehmen die Vorschläge zu 100%)
		GuiControl, abr:, EBMEuroPlus, % Round(EBMPlus, 2) " €"
		GuiControl, abr: Focus, EBMEuroPlus
		;}

	  ;-: Quartalkalender für spätere Abrechnungstage hinzufügen 	;{
		QJahr := ("20" SubStr(abrQuartal, 3, 2) > A_YYYY ? "19" : "20") SubStr(abrQuartal, 3, 2)
		QuartalsKalender(SubStr(abrQuartal, 1, 2), QJahr , 0, "abr", TextColor1)
	  ;}

	;}

	;-: TAB 2 - ICD-Krauler                                                            	;{
		Gui, abr: Tab  	, 2

		GuiControlGet, cp, abr: Pos, % "abrCal"
		Gui, abr: Add	, Edit  	, % "xm+5 ym+35  w" (FSize-1)*5 " 	                        	vabrQ2           	gabrHandler"	, % Abrechnungsquartal
		Gui, abr: Add	, Button	, % "x+5                                                                   	vabrICDKrauler gabrHandler"  	, % "nach Chroniker kraulen lassen"
		Gui, abr: Add	, Edit 	, % "xm+5 y+10 w610 h" cpY+cpH " vabrChrOut"                                                       	, % "~~~~~"

	;}

	;-: TAB 3 - Der große Geriater                                                	;{
		Gui, abr: Tab  	, 3
	;}

	;-: TAB 4 - Einstellungen                                                         	;{
		Gui, abr: Tab  	, 4
		Gui, abr: Tab
	;}

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

		Critical

		If (RegExMatch(A_GuiEvent, "(K|f|C)") || A_EventInfo = 0)
			return

	; eingestellten Patient auslesen
		Gui, abr: ListView, abrLV
		LV_GetText(tmpPatID   		, A_EventInfo, 1)
		LV_GetText(tmpPatName	, A_EventInfo, 2)

		If !(A_GuiEvent = "A")
			If (!tmpPatID && !tmpPatName) || (lvLastPatID = tmpPatID) || !IsObject(VSEmpf[tmpPatID])
				return

		lvLastPatID:=PatID:=tmpPatID,  lvLastPatName:=PatName:=tmpPatName

	; Karteikarte nach Doppelklick anzeigen
		If RegExMatch(A_GuiEvent, "(DoubleClick|A)")
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
			GuiControl, abr: Show  	, % "abr_" feeName
			GuiControl, abr:         	, % "abr_" feeName     	, 1
			GuiControl, abr:         	, % "abr_Date" feeName, % "> " feePossDate
		}

	; Behandlungstage des Patienten anzeigen
		Behandlungstage(PatID, VSEmpf[PatID].BehTage)

return ;}

abrHandler:                              	;{

  ; Gui-Abfrage (ohne Critical funktioniert es nicht!)
	Critical

	Gui, abr: Default
	Gui, abr: Submit, NoHide

	If     	  RegExMatch(A_GuiControl, "(abrQUp|abrQDown)")                            	{         	; Quartalzähler vor und zurück

		QQ 	:= SubStr(abrQ, 1, 2) + (QUD := A_GuiControl = "abrQUp" ? 1 : -1)
		YY   	:= SubStr(abrQ, 3, 2) + 0

		If (QQ = 0)
			QQ := 4, YY += QUD
		else if (QQ = 5)
			QQ := 1, YY += QUD
		YY 	:= YY > 99 ? 0 : YY < 0 ? 99 : YY
		GuiControl, abr:, abrQ, % SubStr( "0" QQ, -1) SubStr( "0" YY,  -1)
		GuiControl, abr:, abrCal, % Addendum.ThousandYear YY . SubStr( "0" ((QQ+2)*3)+1, -1) 	. "01"
		GuiControl, abr:, abrCal, % Addendum.ThousandYear YY . SubStr( "0" ((QQ-1)*3)+1, -1) 	. "01"

	}
	else if RegExMatch(A_GuiControl, "(abrReindex|abrRefresh)")                        	{         	; eingestelltes Abrechungsquartal speichern und anzeigen

	  ; Steuerelemente zurücksetzen
		GuiControl, abr:, abrBT     	, % "--.--.----"
		GuiControl, abr:, abrBTage	, % "--"
		GuiControl, abr: Disable, abrCalStart
		Gui, abr: ListView, abrLV
		CLV.Clear(1, 1)
		LV_Delete()
		Gui, abr: ListView, abrLvBT
		BLV.Clear(1, 1)
		LV_Delete()

	; neu einlesen oder indizieren
		SplashText("Vorbereitung", "Abrechnungsziffern aller Quartale " (A_GuiControl = "abrReindex" ? "werden indiziert ..." : "laden ..."))
		Sleep 2000
		VSEmpf	:= VorsorgeKomplexe(abrQ, (A_GuiControl = "abrReindex" ? true : false), true)
		gosub abrVorschlaege
		SplashText("neues Quartal", "Abrechnungsziffern des Quartal " abrQuartal " wurden "  (A_GuiControl = "abrReIndex" ? "reindiziert" : "geladen"), 6)

		SaveAbrechnungsquartal(abrQ)

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
	else if (A_GuiControl = "abrICDKrauler")                                                      	{      	; sucht neue Pat. für die Abrechnung der Chronikerpauschalen
		Gui, abr: Submit, NoHide
		ICDKrauler(abrQ)
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

		  ; nicht anzeigen wenn eine Bedingung erfüllt ist
			If !RegExMatch(PatID, "^\d+$") || VSEmpf[PatID].Hide || (!ShowYoungStars && vs.Alter < 35) || (!ShowOldStars && vs.Alter > 34) || (!ShowAngels && PatDB[PatID].Mortal)
				continue

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

					If vs.CHRONISCHGRUPPE     	{     	; Chronikerpauschalen
						Chron	:= vs.chronisch	= 1 	? "03220" : vs.chronisch 	= 2 ? "03221" : vs.chronisch	= 3 ? "03220/1"
						Chron	.= vs.chronischQ		? " [" vs.chronischQ " VQ" (vs.chronischQ = 1 ? " fehlt]" : " fehlen]") : ""
					}
					If vs.GERIATRISCHGRUPPE     	{   	; Geriatrische Basiskomplexe
						GB   	:= vs.geriatrisch = 1	? "03360" : vs.geriatrisch	= 2 ? "03362" : vs.geriatrisch= 3 ? "03360+.62"
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

	; belegte Zellen in der Zeile umfärben
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

	Gui, abr: Default
	Gui, abr: Submit, NoHide
	SaveAbrechnungsquartal(abrQ)
	SaveGuiPos(habr, "Abrechnungsassistent")

ExitApp ;}
}

QuartalsKalender(Quartal, Jahr, Urlaubstage, guiname:="QKal", TextColor1:="") 	{     	;-- zeigt den Kalender des Quartals

	; Todo: blendet Feiertage nach Bundesland aus/ein (### ist nicht programmiert! ####)

	;--------------------------------------------------------------------------------------------------------------------
	; Variablen
	;--------------------------------------------------------------------------------------------------------------------;{
		global QKal, hmyCal, hCal, habr, gcolor, tcolor, CalQK, abrQ
		global Cok, Can, Hinweis, abrCal, CalBorder, abrCalDay, abrCalStart, abrCalBTage
		global abrCheckbox, VSRule, CV, PUpX, PUpY, abrLvBT, abrhLvBT, StaticBT, abrhLV
		global KK, lvLastPatID, lvLastPatName, lvPatID, lvPatname, BLV, abrYear
		global BefKuerzel, DateChoosed

		static gname, lastPrgDate, lastBTEvent, lastBTInfo

		If (gname <> guiname)
			gname := guiname
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	Berechnung der Monate
	;--------------------------------------------------------------------------------------------------------------------;{
		FirstMonth:= SubStr("0" ((Quartal-1)*3)+1 , -1)
		LastMonth := SubStr("0" FirstMonth+2    	, -1)
		FormatTime, begindate	, % Jahr . FirstMonth	. "01"                                                	, yyyyMMdd
		FormatTime, lastdate 	, % Jahr . lastmonth	. DaysInMonth(Jahr . lastmonth . "01")	, yyyyMMdd
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	Kalender
	;--------------------------------------------------------------------------------------------------------------------;{
		CalOptions := " AltSubmit +0x50010051 vabrCal HWNDhmyCal gCalHandler"
		lp 	:= GetWindowSpot(AbrhLV)
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
		Gui, abr: Add     	, MonthCal, % "xm+5 y+2 W-3 4 w521"  CalOptions   	, % begindate
		GuiControlGet, cp, abr: Pos, abrCal
		cX 	:= cpX + cpW + 5
		cY		:= cpY + cpH

		Gui, abr: Font     	, % "s10 Normal c" TextColor1
		Gui, abr: Add     	, Text		, % "x" cX " y" cpY  " BackgroundTrans"           	, % "gewählter Tag: "
		Gui, abr: Font     	, % "s11 Bold c" TextColor1
		Gui, abr: Add     	, Text		, % "x" cX " y+5 vabrCalDay"                         	, % "--.--.----                             "
		Gui, abr: Font     	, % "s10 Normal c" TextColor1
		Gui, abr: Add     	, Button		, % "x" cX " y+5 vabrCalStart gCalHandler"   	, % "Ziffernauswahl eintragen"
		GuiControl, abr: Disable, abrCalStart

	; Kalenderfarben
		SendMessage, 0x100A	, 0, 0x666666                     	,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 1, 0xAA0000                     	,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 2, 0x0101AA                     	,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 3, 0x01FF01                       ,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 4, % "0x" RGBtoBGR(gcolor)	,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 5, 0xFFAA99, SysMonthCal321, % "ahk_id " habr
		Sleep 200
		SendMessage, 0x100A	, 1, 0xAA0000                     	,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 2, 0x0101AA                     	,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 3, 0x01FF01                     	,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 4, % "0x" RGBtoBGR(gcolor)	,, % "ahk_id " hmyCal

		WinSet, Style, 0x50010051, % "ahk_id " hmyCal
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	Behandlungstage
	;--------------------------------------------------------------------------------------------------------------------;{

		Gui, abr: Font     	, % "s11 Normal c" TextColor1, Futura Mk Md
		Gui, abr: Add     	, Text  		, % "x" cX " y" cY-15 "BackGroundTrans vStaticBT"                                                                        	, % "Behandlungstage: "
		GuiControlGet, cp, abr: Pos, staticBT
		GuiControl, abr: Move, staticBT, % "x" (nx := 5+lp.W-cpW-30)

		Gui, abr: Font     	, % "s11 Normal c" TextColor1, Futura Bk Bd
		Gui, abr: Add    	, Text      	, % "x" nx+cpW " y" cpY " vabrBTage"                                                                                          	, % "--"

		LVOptions 	:= " -Hdr vabrLvBT gCalHandler HWNDabrhLvBT -E0x200 -LV0x10 AltSubmit -Multi"
		LVColumns	:= "Datum|Kürzel|Inhalt"
		Gui, abr: Font     	, % "s10 Normal cBlack", Futura Bk Bt
		Gui, abr: Add     	, ListView  	, % "xm+5 y+2 w" lp.W " h" ap.CH+1*ap.BH-cpY-cpH " "  LVOptions                                          	, % LVColumns

		LV_ModifyCol(1, 50)
		LV_ModifyCol(2, 65)
		LV_ModifyCol(3, lp.CW-115)
		WinSet, Style, 0x5021800D, % "ahk_id " abrhLvBT   ; No_HScroll

		BLV := new LV_Colors(abrhLvBT)
		BLV.Critical := 200
		BLV.SelectionColors("0x0078D6")

	;}

Return

CalHandler: ;{

		Critical

		Gui, abr: Default
		Gui, abr: Submit, NoHide

		If (lastBTEvent = A_GuiEvent && lastBTInfo = A_EventInfo)
			return
		lastBTEvent := A_GuiEvent, lastBTInfo := A_EventInfo

		;~ ToolTip, % A_GuiControl "`nEvent: " A_GuiEvent "|" lastBTEvent " `ninfo: " A_EventInfo "|" lastBTInfo, 1600, 1, 15
	; Eintragen starten
		If (A_GuiControl = "abrCalStart")
			ZiffernauswahlEintragen(DateChoosed, lvLastPatID, lvLastPatName)

	; angeklicketes Datum sichern
		If (A_GuiControl = "abrCal") {
			If RegExMatch(A_GuiEvent,"i)(Normal|I|1)") {
				DateChoosed 	:= ConvertDBASEDate(StrSplit(abrCal, "-").1)
				GuiControl, abr:, abrCalDay, % DateChoosed
				GuiControl, abr: Enable, abrCalStart
			}
			return
		}

	; alternativ das Datum der angeklickten Zeile aus der Karteikarte
		If (A_GuiControl = "abrLvBT") {

			Gui, abr: ListView, abrLvBT

			If (lastSRow = A_EventInfo && A_GuiEvent = "Normal") {  ; Doppelklick
				ZiffernauswahlEintragen(DateChoosed, lvLastPatID, lvLastPatName)
			}

			lastSRow := A_EventInfo
			sDate := "", sRow := A_EventInfo + 1
			while (!sDate && sRow>0)
				LV_GetText(sDate, sRow--, 1)
			If sDate {
				GuiControl, abr:           	, abrCalDay	, % (DateChoosed := sDate . Addendum.ThousandYear . SubStr(abrQ, 3, 2))
				GuiControl, abr:             	, abrCal      	, % DateChoosedListVars
				GuiControl, abr: Enable	, abrCalStart
			}

		}

return ;}
}

Behandlungstage(PatID, BehTage)                                                                        	{		;-- Karteikartensimulation

		global abr, abrLV, abrLvBT, abrhLvBT, abrBTage, BLV, BefKuerzel, tcolorBG1, tcolorBG2, abrYear
		global VSEmpf
		cols := [], altFlag := true, rows := 0

		Gui, abr:  Default
		Gui, abr: ListView, abrLvBT
		GuiControl, abr:, abrBTage, % BehTage.Count()
		BLV.Clear(1, 1)
		LV_Delete()

		For DATUM, Leistung in BehTage {
			abrYear := SubStr(ConvertDBASEDate(DATUM), 7, 4)
			break
		}

		For DATUM, Leistung in BehTage {

			DATUM := SubStr(ConvertDBASEDate(DATUM), 1, 6), BTrow := 1, rows ++
			For row, col in Leistung {
				If (col.KZL = "lle")
					continue
				cols.1 	:= BTRow = 1 ? DATUM : ""
				cols.2	:= col.KZL
				cols.3	:= col.INH
				BTRow ++
				LV_Add("", cols*)
				bgColor:= "0x" (altFlag ? tcolorBG1 : tcolorBG2)
				txtColor	:= RGBtoBGR(Format("0x{:06X}", BefKuerzel[col.KZL].Color))
				BLV.Row(LV_GetCount() 	, bgColor, 0x000000)
				BLV.Cell(LV_GetCount(), 2	, bgColor, txtColor)
				BLV.Cell(LV_GetCount(), 3	, bgColor, txtColor)

			}
			altFlag := !altFlag
		}

		If (row := LV_GetNext(0, "F"))
			LV_Modify(row, "-Focus"), LV_Modify(row, "-Select")

		WinSet, Style, 0x5021800D, % "ahk_id " abrhLvBT   ; No_HScroll

return abrYear
}

ZiffernauswahlEintragen(DateChoosed, lvPatID, lvPatName)                                 	{   	;-- Karteikartenautomatisierung

		global KK, VSEmpf, VSRule, abrCheckbox, habr, habrLV
		static AlbisClass := "ahk_class OptoAppClass"

		SetTimer, BlockInputOff, 30000
		HandsOff(true, 600, A_ScriptDir "\Einstellungen\handsoff2.png")

	; die Karteikeite des ausgewählten Patienten geöffnet                                   	;{
		PatID := AlbisAktuellePatID()
		If (PatID <> lvPatID) {
			If !AlbisAkteOeffnen(lvPatName, lvPatID) {
				BlockInput, Off
				HandsOff(false)
				MsgBox, 0x1, % A_ScriptName, % "Die Karteikarte von " lvPatname "`nließ sich nicht öffnen.", 3
				return
			}
			PatID := lvPatID
		}
	;}

	; ist aktuelle eine Karteikarte geöffnet?                                                          	;{
		AlbisAnzeige := AlbisGetActiveWindowType()
		If !InStr(AlbisAnzeige, "Karteikarte") {
			BlockInput, Off
			HandsOff(false)
			PraxTT("Es wird keine Karteikarte in Albis angezeigt`nWähle zunächst einen Patienten mit Doppelklick!", "5 1")
			return 0
		}
	;}

	; Karteiblatt anzeigen                                                                                 	;{
		If !RegExMatch(AlbisAnzeige, "i)Karteikarte\s*$")
			If !AlbisKarteikarteZeigen() {
				BlockInput, Off
				HandsOff(false)
				MsgBox, 0x1, % A_ScriptName, % "Es konnte nicht auf das Karteikartenblatt umgeschaltet werden.", 3
				return 0
			}
	;}

	; Patient der aktuellen Karteikarte ist gelistet und hat offene Komplexziffern? 	;{
		PrgDate   	:= AlbisLeseProgrammDatum()
		If (!VSEmpf[PatID].Vorschlag.Count() && !VSEmpf[PatID].chronisch) {
			BlockInput, Off
			HandsOff(false)
			PraxTT(	"Bei Patient [" PatID "] " PatDB[PatID].Name ", " PatDB[PatID].VORNAME
					. 	" fehlt keine Eintragung.`nWähle einen anderen Patienten!", "5 1")
			return 0

		}
	;}

	; gewähltes Datum liegt an einem Wochenende 	                                        	;{
		If RegExMatch(DayOfWeek(DateChoosed, "short"), "^(Sa|So)") {
			BlockInput, Off
			HandsOff(false)
			MsgBox, 0x1024, % StrReplace(A_ScriptName, ".ahk")
						, % 	"Das gewählte Datum (" DateChoosed ") fällt auf ein Wochenende.`n"
						.		"Soll dieses Datum trotzdem verwendet werden?"
			IfMsgBox, No
				return 0
			BlockInput, On
			HandsOff(true)
		}
	;}

	; Gui Daten                                                                                               	;{
		Gui, abr: Default
		Gui, abr: Submit, NoHide
	;}

	; EBM-Ziffern zusammen stellen                                                                  	;{
		EBMFees       	:= GetEBMFees(PatID)
		EBMToAdd     	:= ""
		HKSFormular 	:= false

		If VSEmpf[PatID].CHRONISCHGRUPPE
			EBMToAdd .= VSEmpf[PatID].chronisch = 1	? "03220-" : VSEmpf[PatID].chronisch = 2 ? "03221-" : VSEmpf[PatID].chronisch = 3 ? "03220-"

		If VSEmpf[PatID].GERIATRISCHGRUPPE
			EBMToAdd .= VSEmpf[PatID].geriatrisch = 1	? "03360-" : "03362-"

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
				HandsOff(false)
				MsgBox, EMBtoAdd  ist leer!
				HandsOff(true)

		}
		;SciTEOutput("EBM String: " EBMtoAdd " einzutragen am " DateChoosed)
	;}

	; Ziffern eintragen                                                                                        	;{

		; aktivieren
			AlbisActivate(1)

		; aktuelles Programmdatum lesen
			lastPrgDate := AlbisSetzeProgrammDatum(DateChoosed)

		; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
		; Karteikarte anzeigen
		; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
			If InStr(AlbisGetActiveWindowType(), "Karteikarte")
				If !AlbisKarteikarteZeigen() {
					BlockInput, Off
					HandsOff(false)
					MsgBox, Die Karteikarte hat sich nicht anzeigen lassen.
					return 0
				}

		; Karteikartenzeile lesen
			AlbisKarteikarteAktivieren()
			KKarte	:= AlbisGetActiveControl("content")

		; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
		; Eingabefokus setzen
		; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
			If !IsObject(KKarte) {
				SendInput, {INSERT}
				Sleep 50
				KKarte	:= AlbisGetActiveControl("content")    ; jetzt kann die Zeile gelesen werden
			}

		; Zeilendatum lesen
			ZDatum := AlbisLeseZeilenDatum()
			If (ZDatum <> DateChoosed) {
				AlbisKarteikarteFocus("Datum")
				If !VerifiedSetText(KK.DATUM, DateChoosed, AlbisClass) {
					SendInput, {Home}                                     	; an den Anfang
					SendInput, {LShift Down}{End}{Shift Up}   	; alles markieren
					Sleep 100
					Send, % DateChoosed
					Sleep 100
					SendInput, {TAB}
				}
			}

		; Eingabefokus weiterrücken ins Kürzelfeld
			AlbisKarteikarteFocus("Kuerzel")

		; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
		; lko
		; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
			If (StrLen(KKarte.kuerzel) = 0) {

			 If (Kkarte.FName <> "Kuerzel")                                                 ; FName = Focused Name
				If !VerifiedClick(KK.kuerzel,,, AlbisWinID())
					If !VerifiedSetFocus(KK.kuerzel, AlbisWinID(),,, false) {
						ControlGetPos   	, ctrlX, ctrlY,,, % KK.KUERZEL	, % AlbisClass
						ControlClick      	, % KK.KUERZEL                   	, % AlbisClass,, Left, 1, NA
						ControlGetFocus	, fc                                      	, % AlbisClass
						If !InStr(KK.KUERZEL, fc) || !InStr(fc, KK.KUERZEL)
							SciTEOutput("nix is :" fc)
							;MsgBox, bekomme das Kürzelfeld nicht fokussiert.
					}

			ControlGetFocus, fcs, % AlbisClass
			If (fcs<>KK.kuerzel)
				SciTEOutput(fcs " <> " KK.KUERZEL)
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
				ControlGetFocus, fcs, % AlbisClass
				If (fcs = KK.Inhalt) || (A_Index > 40)
					break
				Sleep 100
			}

		; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
		; String senden
		; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
			If !VerifiedSetText(KK.Inhalt, EBMtoAdd, AlbisWinID) {
				SendInput, % "{Raw}" EBMtoAdd
				Sleep 150
			}
			SendInput, {TAB}
			Sleep 100
			SendInput, {TAB}
			Sleep 100
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

	; Hautkrebscreening-Formular                                                                    	;{
		If RegExMatch(EBMtoAdd, "(01745|01746)")
			if (hEHKS := AlbisMenu(34505, "Hautkrebsscreening - Nichtdermatologe ahk_class #32770", 3))
				AlbisHautkrebsScreening()   ; alles leer lassen, alles super!
	;}

	; auf letztes Datum zurücksetzen und Ausnahmeindikationen korrigieren       	;{
		AlbisSetzeProgrammDatum(lastPrgDate)
		ausIndi := AusIndisKorrigieren(VSEmpf[PatID].CHRONISCHGRUPPE
													, VSEmpf[PatID].GERIATRISCHGRUPPE
													, Addendum.KeinBMP
													, Addendum.ShowIndis )
	;}

	;}

	; CAVE Dokumentation - Änderungen vornehmen                                         	;{
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

		If !(hCave := Albismenu(32778, "Cave! von ahk_class #32770", 4)) {
			BlockInput, Off
			HandsOff(false)
			return 0
		}
		CLines := AlbisGetCave(false)
		For lineNr, line in StrSplit(CLines, "`n") {
			If RegExMatch(line, "i)(Td|TdPP|Masern|Pocken|Tet|BCG|COV|Pneu|GSI|MMR|Polio|Pert|FSME|Hep|Hib|Po)")
				Impfung .= lineNr "|"
			If RegExMatch(line, "i)(01740|AORTA|GVU|HKS|KVU|GB|Colo|Gastro|PP)")
				Vorsorge .= lineNr "|"
		}

		If Addendum.ChangeCave {
			AlbisCaveSetCellFocus(hCave, lineNr-1, "Beschreibung")
			while (WinExist("Cave! von ahk_class #32770") && A_Index < 10)
				Sleep 300
		}

	;}

	; CAVE Dokumentation - schliessen                                                            	;{
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

	; Impfbuch Hinweis in die Karteikarte zur Erinnerung einschreiben                 	;{
		If !Impfung {

		; aktuelles Programmdatum lesen
			infoDate := "31.12." SubStr(DateChoosed, 7, 4)
			lastPrgDate := AlbisSetzeProgrammDatum(infoDate)
			If AlbisZeilendatumSetzen(infoDate) {

			  ; Karteikartenzeile beschreiben
				Send, % "Impf"
				Sleep 200
				SendInput, % "{TAB}"
				Sleep 200
				ControlGetFocus, TabbedFocus, % AlbisClass
				newTabbed := TabbedFocus
				;AlbisKarteikarteFocus("Inhalt")
				Send, % "Impfausweis mitbringen"
				Sleep 200
				SendInput, % "{TAB}"
				while (newTabbed = TabbedFocus && A_Index < 200) {
					Sleep 20
					ControlGetFocus, newTabbed, % AlbisClass
				}
			}

		}
		else {

			FileAppend, % Impfung, % Addendum.DBPath "\sonstiges\Impfungen.txt", UTF-8

		}

	;}

	SendInput, {Escape}
	Sleep 100
	SendInput, {Escape}

	; MsgBX - melde das Du das fertig hast                                                       	;{
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
				BlockInput, Off
				HandsOff(false)
				msg := "Ausnahmeindikation" (AltC > 0 ? "en" : "") ":`n" Alt "`nwurde" (AltC > 0 ? "n" : "") " nicht geändert"
				MsgBox	, 0x1024
						, % A_ScriptName
						, % "Bitte alles überprüfen!`n`n" msg
						, 20
			}
		}
	;}

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

ImpfBoss()                                                                                                         	{

	geimpft =
	(
		Td-w 07, Polio w 15, FSME-w 15,
		Tdp-w15, GSI 14
		TdPP-w 15, Td-3 05, HepAB-3 09
	)
}

GetEBMFees(PatID)                                                                                               	{

	global	VSEmpf, VSRule

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

	;~ JSONData.Save("C:\tmp\tmp.json", EBMFees, true,, 1, "UTF-8")

return EBMFees
}

VorsorgeKomplexe(abrQuartal, ReIndex:=false, SaveResults:="")                            	{     	;-- ermittelt alle Abrechnungsdaten des Quartals

		static PathVSCandidates

	; nur bei leerem SaveResults-Parameter ReIndex verwenden
		SaveResults       	:= SaveResults ? SaveResults : ReIndex
		quartal               	:= LTrim(abrQuartal, "0")
		YearQuarter      	:= SubStr(quartal, 2, 2) SubStr(quartal, 1, 1)
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
			SplashText("bitte warten ...", "Vorsorge Kandidaten wurden geladen", 4)
			return VSEmpf
		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; keine Daten vorhanden oder neu indizieren wurde übermittelt
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	  ; Krankenscheine:         	Abrechnungsart: ohne Überweiser, Notfallscheine oder Private
	  ; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		If (!IsObject(VSEmpf) || ReIndex || !FileExist(Addendum.DBasePath "\KScheine_" quartal ".json")) {
			SplashText("bitte warten ... 1/7", "Abrechnungsscheine des Quartals werden gelesen ...")
			KScheine	:= DBAccess.Krankenscheine(abrQuartal, "Abrechnung", false, SaveResults)
		}

	  ; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	  ; Behandlungstage:     	des Quartals für jeden Patienten zusammenstellen
	  ; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		If (!IsObject(VSEmpf) || ReIndex || !FileExist(Addendum.DBasePath "\BehTage_" quartal ".json")) {
			SplashText("bitte warten... 2/7", "Behandlungstage im Quartal ermitteln ...")
			BehTage	:= DBAccess.Behandlungstage(abrQuartal, "Abrechnung", false, true)
		}
		else
			BehTage	:= JSONData.Load(Addendum.DBasePath "\BehTage_" quartal ".json", "", "UTF-8")

	  ; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	  ; Vorsorgedaten:         	von allen Patienten die Vorsorgeziffern und Abrechnungstage finden
	  ; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		If (!IsObject(VSEmpf) || ReIndex || !FileExist(Addendum.DBasePath "\Vorsorgen.json")) {
			SplashText("bitte warten ... 3/7", "EBM Ziffern aller Quartale werden gesammelt ...")
			VSDaten 	:= DBAccess.VorsorgeDaten(abrQuartal, ReIndex, SaveResults)
		}

	  ; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	  ; Vorsorgekandidaten:  	finden
	  ; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		If (!IsObject(VSEmpf) || ReIndex || !FileExist(PathVSCandidates)) {

			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
			; letzte Vorsorgekandidatendatei laden um die bearbeiteten Datensätze zu behalten
			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
				VSEmpfOld := ""
				If FileExist(PathVSCandidates) {
					SplashText("bitte warten ... 4/7", "Vorsagekandidaten Datenbackup wird geladen  ...")
					VSEmpfOld := JSONData.Load(PathVSCandidates, "", "UTF-8")
				}

			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
			; Vorsorgekandidaten neu erstellen
			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
				SplashText("bitte warten ... 5/7", "EBM Ziffern aller Quartale werden gefiltert ...")
				VSEmpf := DBAccess.VorsorgeFilter(VSDaten, abrQuartal, "Abrechnung", false, (ReIndex ? false : true))

			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
			; Daten aus anderer Kandidatendatei mit neuen Daten vergleichen
			; (im Moment nur key: "Hide" - manuell ausgeblendete Patienten)
			; SciTEOutput(" [" PatID "] " PatDB[PatID].NAME ", " PatDB[PatID].VORNAME)
			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
				If IsObject(VSEmpfOld) {
					SplashText("bitte warten ... 6/7", "Anzeigeeinstellungen werden geladen...")
					saveVSCandidates := true
					For PatID, vs in VSEmpfOld
						VSEmpf[PatID].Hide := VSEmpfOld[PatID].Hide
				}

			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
			; Behandlungstage zuorden
			; ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧ ‧
				If IsObject(BehTage) {
					SplashText("bitte warten ... 7/7", "Behandlungstage werden geladen...")
					saveVSCandidates := true
					For PatID, Tage in BehTage
						If !IsObject(VSEmpf[PatID].BehTage)
							VSEmpf[PatID].BehTage := BehTage[PatID]
				}

		}

		SplashText("bitte warten ...", "")

	  ; Objekte freigeben
		DBAccess.EmptyDBData()

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; VSEmpf sichern wenn Änderungen gemacht wurden
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If saveVSCandidates
			JSONData.Save(PathVSCandidates, VSEmpf, true,, 1, "UTF-8")

return VSEmpf
}

ICDKrauler(abrQuartal)                                                                                     	{     	;-- sucht Patienten für die Abrechnung der Chroniker-Ziffern

	; sucht alle ICD Abrechnungsdiagnosen eines Quartals heraus und vergleicht diese mit einer ICD Liste
	; 2 Datenbankdateien müssen für die Suche ausgelesen werden: 1.BEFUND.dbf und 2.BEFTEXTE.dbf
	; die Eingrenzung der Suche innerhalb der BEFUND.dbf über das Abrechnungsquartal ist nicht anwendbar,
	; Albis hinterlegt dort nicht manchmal die Quartalbezeichnung meist steht dort eine 0 oder eine andere Zahl
	; das Abrechnungsquartal wird in ein Datum in Form eines RegEx-String umgewandelt aus
	; 0320 wird 2020(07|08|09)\d\d. Damit sind alle Tage des Quartals erfasst.

		static CRICD
		static file_chronkrank
		static wwwCRICD   	:= "https://raw.githubusercontent.com/Ixiko/Addendum-fuer-Albis-on-Windows/meta/include/Daten/ICD-chronisch_krank.json"

		If !file_chronkrank
			file_chronkrank	:= Addendum.Dir "\include\Daten\ICD-chronisch_krank.json"

	; ICD Vergleichsliste laden
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		If !IsObject(CRICD) {

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
						MsgBox, 0x1004, Daten vermisst, % "Der Download ist fehlgeschlagen.`n Möchten Sie es noch einmal versuchen`n"
						IfMsgBox, No
							return
					}
					else
						downloadToFile := false
				}
			}

			CRICD := JSONData.Load(file_chronkrank, "", "UTF-8")

		  ; prüft das Objekt auf den richtigen Inhalt
			For letter, ICDcollection in CRICD {

					SciTEOutput("letter: " letter)
					If !RegExMatch(letter, "^[A-Z]$") {
						SciTEOutput("letter mismatch: " letter)
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
							gotoDownload := true
						IfMsgBox, No
							return
					}

			}

			If gotoDownload
				ICDKrauler(abrQuartal)

		}

	; Datenbankzugriff
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		If !IsObject(DBAccess)
			DBAccess 	:= new AlbisDb(Addendum.AlbisDBPath, 0)

	; Quartal in RegExString umwandeln
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		Quartal := QuartalTage({"aktuell": abrQuartal})
		rxQuarterDays := "rx:" Quartal.Jahr "(" Quartal.MBeginn "|" Quartal.MMitte "|" Quartal.MEnde "|" ")\d\d"

	; BEFUND.dbf Suche nach Kürzel 'dia' im Datumsbereich, Erfassung der Patienten, Diagnosen und den
	; Verbindungen zu den Einträgen in der BEFTEXT.dbfm, Chroniker heraussuchen
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		dia    	:= DBAccess.GetDBFData("BEFUND", {"KUERZEL":"dia", "DATUM":rxQuarterDays}, ["PATNR", "DATUM", "INHALT", "TEXTDB"])
		ill      	:= DBAccess.Chroniker(abrQuartal, true, false)

		;~ JSONData.Save(A_Temp "\temp.json", dia, true,, 1, "UTF-8")
		;~ Run, % A_Temp "\temp.json"
		;~ DBAccess := ""
		;~ return

	; TEXTDB Werte in einem Objekt erfassen
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		TextDB := Object()
		For each, set in dia
			If RegExMatch(Trim(set.TextDB), "^\d+$")
				TextDB[set.TextDB] := set.PATNR

	; zusätzliche Daten komplett auslesen
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		linked 	:= DBAccess.GetBEFTEXTData(TextDB)

	; neues Objekt sortiert nach PATNR mit allen Diagnosedaten erstellen
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		qDia := Object()
		For LFDNR, PATNR in TextDB {

		}

	; Hinweisfenster
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
	;~ ICDKraulerGui()


	; Ergebnisse anzeigen
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺

	  For matchNr, set in dia {
			;~ t .= set.PATNR "`t" SubStr(ConvertDBASEDate(set.Datum), 1, 6) " " set.INHALT "`n"
			If RegExMatch(set.INHALT, "\{(?<CD>[A-Z]\d+\.*[\d\-]+).*?G\}")
				nil := 0
	  }


	;~ ICDKraulerGui(false)

}

ICDKraulerGui(ShowGui:=true) {

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

ICDKraulerCallback() {

}

CoronaEBM(abrQuartal)                                                                                    	{   	;-- alle Coronaabstrich und alle Behandlungsziffern finden

	COVID     	:= Object()
	abrYear      	:= "2021"
	aDB          	:= new AlbisDB(Addendum.AlbisDBPath)
	abrScheine  	:= aDB.Abrechnungsscheine(abrQuartal, false, true)
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

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Hilfsfunktionen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - -;{
SaveAbrechnungsquartal(Quartal:="")                                                           	{  	;-- speichert das aktuell eingestellte Quartal

	If !RegExMatch(Quartal, "0*[1-4]\d{2}") {
		Gui, abr: Default
		Gui, abr: Submit, NoHide
		Quartal := SubStr("0" abrQ, -3)
	}
	IniWrite, % Quartal, % Addendum.ini, Abrechnungshelfer, Abrechnungsassistent_letztesArbeitsquartal

return Quartal
}

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

SplashText(title, msg:="", offtime:="")                                                                   	{
	SplashTextOn, 300, 20, % title, % msg
	hSTO := WinExist("Vorbereitung ahk_class AutoHotkey2")
	WinSet, Style 	, 0x50000000, % "ahk_id " hSTO
	WinSet, ExStyle	, 0x00000208, % "ahk_id " hSTO
	If RegExMatch(offtime, "^\d+$")
		SetTimer, SplashTOff, % -1*(offtime*1000)
	else If (StrLen(msg) = 0)
		gosub SplashTOff
return
SplashTOff:
	SetTimer, SplashTOff, Off
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

HandsOff(showGui, width:="", ImagePath:="")                                                    	{ 		;-- Nutzerhinweis einblenden

	Static gvisible:=false, firstrun:=true

	If firstrun {

		mdiclient := GetWindowSpot(AlbisMDIChildGetActive())
		x := mdiclient.X + Floor(mdiclient.W/2), y := mdiclient.Y + Floor(mdiclient.H/2)

		Gui, HOff: new	, +LastFound -Caption +E0x80000 +Owner +AlwaysOnTop +hwndhHoff +ToolWindow +OwnDialogs
		Gui, HOff: Show	, NoActivate, HandsOff

		if !FileExist(ImagePath)
			return

		DllCall("gdiplus\GdipCreateBitmapFromFile", WStr, ImagePath, PtrP, pBitmap)
		DllCall("gdiplus\GdipGetImageWidth", Ptr, (pBitmap := Gdip_ResizeBitmap(pBitmap, width, 400, 1,, 1)), UIntP, w)
		DllCall("gdiplus\GdipGetImageHeight", Ptr, pBitmap, UIntP, h)

		x -= Floor(width/2), y -= Floor(h/2)

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
		Gui, HOff: Show	, % "x" x " y" y " NoActivate "
		return

	}

	Gui, HOff: Show	, % " NoActivate " (!showGui ? "Hide" : "")

return

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
#include %A_ScriptDir%\..\..\lib\GDIP_all.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\lib\Sift.ahk
;}





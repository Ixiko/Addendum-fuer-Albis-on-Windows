; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;                                      	💰	Addendum Abrechnungsassistent
;                  	     				💰	Ersatz für den Addendum Abrechnungshelfer? Nein, noch nicht! Ersatz für den Albis Abrechnungsassistenten? Durchaus.
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;      Funktion: 				 ⚬	prüft auf Abrechenbarkeit von Vorsorgeuntersuchungs- und beratungsziffern (01732,01745/56, 01740, 01747, 01731) nach
;										den allgemeinen EBM-Regeln entsprechend Alter und Abrechnungsfrequenz
;                                	 ⚬ schaut ob die Chronikerziffern (eingegeben wurden)
;                                	 ⚬ wird zunehmend Funktionen für das Auffinden bestimmter Abrechnungspositionen bei entsprechenden Konstellationen bekommen
;
;
;		Hinweis:
;
;		Abhängigkeiten:	siehe includes
;
;	                    			Addendum für Albis on Windows
;                        			by Ixiko started in September 2017 - last change 09.04.2021 - this file runs under Lexiko's GNU Licence
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣

  ; Einstellungen
	#NoEnv
	#Persistent
	#KeyHistory, Off

	SetBatchLines, -1
	ListLines    	, Off

	global PatDB
	global Addendum

  ; startet Windows Gdip
   	If !(pToken:=Gdip_Startup()) {
		MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
		ExitApp
	}

  ; Tray Icon erstellen
	;~ If (hIconLabJournal := Create_Laborjournal_ico())
    	;~ Menu, Tray, Icon, % "hIcon: " hIconLabJournal

  ; Albis Datenbankpfad / Addendum Verzeichnis
	RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
	RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)

	Addendum                 	:= Object()
	Addendum.Dir            	:= AddendumDir
	Addendum.Ini              	:= AddendumDir "\Addendum.ini"
	Addendum.DBPath      	:= AddendumDir "\logs'n'data\_DB"
	Addendum.DBasePath  	:= AddendumDir "\logs'n'data\_DB\DBase"
	Addendum.AlbisDBPath	:= AlbisPath "\DB"
	Addendum.compname	:= StrReplace(A_ComputerName, "-")                                                                 	; der Name des Computer auf dem das Skript läuft

	SciTEOutput()

  ; Patientendatenbank
	infilter := [	"NR", "VNR", "VKNR", "PRIVAT", "ANREDE", "ZUSATZ", "NAME", "VORNAME", "GEBURT", "TITEL"
					, 	"PLZ", "ORT", "STRASSE", "TELEFON", "GESCHL", "SEIT", "ARBEIT", "HAUSARZT", "GEBFREI", "LESTAG", "GUELTIG"
					, 	"FREIBIS", "KVK", "TELEFON2", "TELEFAX", "LAST_BEH", "DEL_DATE", "MORTAL", "HAUSNUMMER", "RKENNZ"]
	PatDB 	    	:= ReadPatientDBF(Addendum.AlbisDBPath, infilter)

  ; Abrechnungsassistent anzeigen
	AbrAssistGui("0121")


return

#IfWinActive, Abrechnungsassistent ahk_class AutoHotkeyGUI
~ESC:: ;{
	SaveWinPositions()
	ExitApp
return ;}
#IfWinActive
~^!b:: ;{
	SaveWinPositions()
	Reload
return ;}

SaveWinPositions() {

	global habr, hCal, hmyCal

	saveGuiPos(habr, "Abrechnungsassistent")
	saveGuiPos(hCal, "Kalender")

}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; grafische Anzeige
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
AbrAssistGui(Abrechnungsquartal, Anzeige=true) {

		global

	; Variablen                             	;{
		static EventInfo, empty := "- - - - - - -"
		static FSize := (A_ScreenHeight > 1080 ? 11 : 10)
		static abrUD_o
		static abrQuartal
		abrQuartal := abrUD_o := Abrechnungsquartal

		local hAlbis, winpos

	; Komplexe laden/erstellen
		SplashText("Vorbereitung", "Abrechnungsziffern aller Quartale laden ...")
		VSEmpf	:= VorsorgeKomplexe(Abrechnungsquartal, false)

	; EBM Regeln laden
		SplashText("Vorbereitung", "lade EBM Regeln ...")
		VSRule 	:= DBAccess.GetEBMRule("Vorsorgen")
		If !IsObject(VSRule)
			throw A_ThisFunc ": EBM Vorsorgeregeln konnten nicht übermittelt werden!"

	; Euro Preise der Gebührennummern
		Euro := {"#01731"	: 15.82
					, "#01732"	: 35.82
					, "#01745"	: 27.80
					, "#01746"	: 22.96
					, "#01740"	: 12.75
					, "#01747"	: 9.01
					, "#03320"	: 14.28
					, "#03321"	: 4.39
					, "KVU"      	: 15.82
					, "GVU"      	: 35.82
					, "HKS1"     	: 27.80
					, "HKS2"     	: 22.96
					, "Colo"      	: 12.75
					, "Aorta"     	: 9.01}
	;}

	; letzte Fensterposition laden   	;{
		IniRead, winpos, % Addendum.ini, % Addendum.compname, Abrechnungsassistent_Position
		If (InStr(winpos, "ERROR") || StrLen(winpos) = 0)
			winpos := "w620 h800"
	;}

	; Albis Fensterposition              	;{
		wp := PosStringToObject(winpos)
		If (hAlbis := AlbisWinID()) {
			ap := GetWindowSpot(hAlbis)
			winpos := "x" ap.X+ap.W " y" (ap.Y < 0 ? 0 : ap.Y) " AutoSize NA"
		}
	;}

	; andere Variablen                	;{
		gcolor          	:= "BBBB99"
		tcolor           	:= "e4e4c4"
		tcolorD        	:= "646464"
		tcolorM        	:= "949494"
		abrTabNames := "Ziffern-Guru|ICD-Krauler|Großer-Geriater"
	;}

	; Gui                                       	;{

		Gui, abr: new	, % "+AlwaysOnTop HWNDhabr -DPIScale"
		Gui, abr: Color	, % "c" gcolor , % "c" tcolor
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal cBlack", Futura Bk Bt
		Gui, abr: Add  	, Tab     	, % "x3 y3  	w" 635 " h" wp.H-5 " HWNDabrHTab vabrTabs gabr_Tabs"             	, % abrTabNames

	;-: TAB 1 - ZIFFERN-GURU     	;{
		Gui, abr: Tab  	, 1

	  ;-: Kopf                                                                                  ;{
		Gui, abr: Add	, Text        	, % "xm ym+30    	BackgroundTrans             	"                                            	, % "Quartal"
		Gui, abr: Add	, Edit           	, % "x+10    w" (FSize-1)*5 " 	                     	vabrQ           	gabrHandler"	, % Abrechnungsquartal
		GuiControlGet, cp, abr: Pos, abrQ
		Gui, abr: Add	, Button        	, % "x" cpX+cpW " y" cpY " h" cpH "	            	vabrQUp       	gabrHandler"   	, % "▲"
		Gui, abr: Add	, Button        	, % "x+0 h" cpH "                        	            	vabrQDown   	gabrHandler"   	, % "▼"

		Gui, abr: Font	, % "s" FSize " q5 Normal cBlack", Futura Bk Bt
		Gui, abr: Add	, Text        	, % "x+30     	BackgroundTrans             	"                                               	, % "mögliches EBM plus: "
		Gui, abr: Add	, Text        	, % "x+0       	BackgroundTrans  				vEBMEuroPlus"                        	, % "00000.00 Euro  "
		Gui, abr: Add	, Text        	, % "x+20     	BackgroundTrans             	"                                               	, % "Krankenscheine: "
		Gui, abr: Add	, Text        	, % "x+0       	BackgroundTrans  				vEBMScheine"                         	, % VSEmpf["_Statistik"].KScheine "  "
	;}

	  ;-: Section 1                                                                           	;{
		Gui, abr: Add  	, Progress   	, % "x8 y+10 w600 h6 -E0x20000 vabrPS1 HWNDabrHPS1 c" tcolorD , 100
	  ;}

	  ;-: Refresh / ReIndex                                                              	;{
		Gui, abr: Font	, % "s" FSize+6 " q5 Normal c" tcolorD, Futura Bk Bt
		Gui, abr: Add	, Text        	, % "xm y+8 "                                                                                   	, % "♻"
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal cBlack", Futura Bk Bt
		Gui, abr: Add	, Button        	, % "x+2                                  	       	vabrReFresh   	gabrHandler"   	, % "Liste auffrischen"
		Gui, abr: Font	, % "s" FSize+6 " q5 Normal c" tcolorD, Futura Bk Bt
		Gui, abr: Add	, Text        	, % "x+30 "                                                                                   	, % "🔄"
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal cBlack", Futura Bk Bt
		Gui, abr: Add	, Button        	, % "x+2                             	       	vabrReIndex   	gabrHandler"   	, % "Index erneuern"

		Gui, abr: Font	, % "s" FSize+6 " q5 Normal c" tcolorD, Futura Bk Bt
		Gui, abr: Add	, Text        	, % "x+50 vabrHide"                                                                       	, % "☑"
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal cBlack", Futura Bk Bt
		Gui, abr: Add	, Button        	, % "x+2                            	       	vabrBearbeitet  	gabrHandler"   	, % "Markierte ausblenden"
		GuiControlGet, cp, abr: Pos, abrHide
		Gui, abr: Font	, % "s" FSize+10 " q5 Normal c" tcolorD, Futura Bk Bt
		Gui, abr: Add	, Text        	, % "x" cpX " y" cpY+cpH+3                                                          	, % "◻"
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal cBlack", Futura Bk Bt
		Gui, abr: Add	, Button        	, % "x+2                            	       	vabrUnHide  	gabrHandler"   	, % "alle einblenden"
	;}

	  ;-: Optionen                                                                         	;{
		CBOpt := "BackgroundTrans Checked"
		Gui, abr: Add	, Checkbox   	, % "xm 	y+10 "               	CBOpt " 	vabrYoungStars	gabrHandler"  	, % "Alter 18 bis 34 anzeigen"
		GuiControlGet, cp, abr: Pos, abrYoungStars
		Gui, abr: Add	, Checkbox   	, % "x+20 "                        	CBOpt "	vabrOldStars 	gabrHandler"  	, % "Alter ab 35 anzeigen"
		Gui, abr: Add	, Checkbox   	, % "x+20 "                        	CBOpt " 	vabrAngels    	gabrHandler"   	, % "Verstorbene anzeigen"
		Gui, abr: Add	, Text        	, % "x+40           	BackgroundTrans  	vabrLVRows"                          	, % "[0000]"
	;}

	  ;-: Patientenliste / Listview                                                    	;{
		LVOptions3 := "vabrLV gabrLVHandler HWNDabrhLV Checked Grid -E0x200 -LV0x10 AltSubmit -Multi "
		Gui, abr: Font	, % "s" FSize-1 " q5 Normal", Futura Bk Bt
		Gui, abr: Add	, ListView    	, % "xm y+5 w620 h564 " LVOptions3                                                 	, % "ID|Name, Vorname|Alter|M/W|GVU|HKS|KVU|Colo|Aorta|03220/1"
		LV_ModifyCol(1 	, "70 	Right Integer")
		LV_ModifyCol(2 	, "145 	Left Text")
		LV_ModifyCol(3 	, "40 	Center Integer")
		LV_ModifyCol(4 	, "40 	Center Text")
		LV_ModifyCol(5 	, "40 	Center Text")
		LV_ModifyCol(6 	, "40 	Center Text")
		LV_ModifyCol(7 	, "40 	Center Text")
		LV_ModifyCol(8 	, "40 	Center Text")
		LV_ModifyCol(9 	, "45 	Center Text")
		LV_ModifyCol(10	, "100 	Left Text")
		LV_ModifyCol(11	, 0)
		;}

	  ;-: Patienten und Untersuchungsauswahl                               	;{
		Gui, abr: Font	, % "s" FSize " q5 Normal cBlack", Futura Bk Bt
		Gui, abr: Add	, Text        	, % "xm y+8  "             	, % "gewählter Patient: "
		Gui, abr: Add	, Text        	, % "x+0 w" FSize*20 "  	vabrPat"    	, % " "
		;}

	  ;-: Checkboxen für Untersuchungen                                        	;{
		abrCBX := Object()
		Gui, abr: Font	, % "s" FSize-1   	" q5 Normal cBlack", Futura Bk Bt
		Loop % VSRule.VSFilter.MaxIndex() {
			feat := VSRule.VSFilter[A_Index].exam
			If !abrCBX.haskey(feat) {
				Gui, abr: Add	, Checkbox   	, % (A_Index = 1 ? "xm y+5" : "x+25")  " BackgroundTrans vabr_" feat " HWNDabrHwnd" 	, % feat
				GuiControlGet, cp, abr: Pos, % "abr_" feat
				abrCBX[feat] := cpX
			}
		}
		;}

	  ;-: zeigt später den Tag der frühesten Abrechnung an           	;{
		floop := 1
		Gui, abr: Font	, % "s" FSize-3    	" q5 Normal cBlack", Futura Bk Bt
		For feat, X in abrCBX {
			Gui, abr: Add	, Text   	, % (floop = 1 ? "xm y+2" : "x" X " y" cpY)  " BackgroundTrans vabr_Date" feat " " 	, % " [00.00.0000] "
			GuiControlGet, cp, abr: Pos, % "abr_Date" feat
			GuiControl, abr:, % "abr_Date" feat, % ""
			floop ++
		}
		;}

	  ;-: zeigt alle Vorschläge in einer Liste an                               	;{
		CLV := New LV_Colors(abrhLV)
		CLV.SelectionColors("0x" gcolor)
		gosub abrGuiVorschlaege
	  ;}

	  ;-: Section 2                                                                        	;{
		Gui, abr: Add  	, Progress   	, % "x0 y+5 w620 h3 -E0x20000 vabrPS2 HWNDabrHPS2 c" tcolorD , 100
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
		QuartalsKalender(SubStr(Abrechnungsquartal, 1, 2), QJahr , 0, "abr")
	  ;}

	;}

	;-: TAB 2 - ICD-Krauler          	;{
		Gui, abr: Tab  	, 2

		Gui, abr: Add	, Edit  	, % "xm ym+35  w" (FSize-1)*5 " 	                        	vabrQ2           	gabrHandler"	, % Abrechnungsquartal
		Gui, abr: Add	, Button	, % "x+5                                                               	vabrICDKrauler gabrHandler"  	, % "nach Chroniker kraulen lassen"
		Gui, abr: Add	, Edit 	, % "xm y+10 w610 h" wp.H-90 " vabrChrOut"                                                       	, % "~~~~~"

	;}

	;-: TAB 3 - Der große Geriater 	;{
		Gui, abr: Tab  	, 3
	;}

		Gui, abr: Tab
		Gui, abr: Show, % winpos " Hide" , Abrechnungsassistent

		abrP := GetWindowSpot(habr)
		GuiControl, abr: Move, abrPS1, % "x4 w" abrP.CW-18
		GuiControl, abr: Move, abrPS2, % "x4 w" abrP.CW-18
		WinSet, ExStyle, -0x20000, % "ahk_id " abrHPS1
		WinSet, ExStyle, -0x20000, % "ahk_id " abrHPS2


		Gui, abr: Show,, Abrechnungsassistent
	;}

	SplashTextOff

return

abrLVHandler: ;{


	If Instr(A_GuiEvent	, "DoubleClick") {

		; Datenbankname lesen
			Gui, abr: Default
			Gui, abr: ListView, abrLV

			LV_GetText(PatID, A_EventInfo, 1)
			LV_GetText(PatName, A_EventInfo, 2)
			GuiControl, abr:, abrPat, % "[" PatID "] " PatName

			For feat, featHwnd in abrCBX {
				GuiControl, abr:Disable, % "abr_" feat
				GuiControl, abr:, % "abr_" feat, 0
				GuiControl, abr:, % "abr_Date" feat, % ""
			}
			For featIDX, feat in VSEmpf[PatID].Vorschlag {
				RegExMatch(feat, "i)^(?<Name>[a-z]+)\-*(?<PossDate>\d+)*", Feat)
				GuiControl, abr:Enable, % "abr_" featName
				GuiControl, abr:, % "abr_" FeatName, 1
				GuiControl, abr:, % "abr_Date" FeatName, % ">" ConvertDBASEDate(FeatPossDate)
			}

			AlbisAkteOeffnen(PatName, PatID)

		}
return ;}

abrHandler:                    	;{

  ; Quartale vor und zurück schalten
	If RegExMatch(A_GuiControl, "(abrQUp|abrQDown)")                                 	{

		Gui, abr: Submit, NoHide
		abrQQYY:= abrQ
		abrQQ	:= SubStr(abrQQYY, 1, 2) + 0
		abrYY	:= SubStr(abrQQYY, 3, 2) + 0
		QUD 	:= A_GuiControl = "abrQUp" ? 1 : -1

		abrQQ := abrQQ + QUD
		If (abrQQ = 0)
			abrQQ := 4, abrYY := abrYY + QUD
		else if (abrQQ = 5)
			abrQQ := 1, abrYY := abrYY + QUD
		abrYY	:= abrYY > 99 ? 0 : abrYY < 0 ? 99 : abrYY
		QQYY	:= SubStr( "0" abrQQ, -1) SubStr( "0" abrYY,  -1)
		GuiControl, abr:, abrQ, % QQYY

	}
	else if (A_GuiControl = "abrReIndex")                                                         	{

		Gui, abr: Submit, NoHide
		Abrechnungsquartal := SubStr("0" abrQ, -3)
		SplashText("Vorbereitung", "Abrechnungsziffern aller Quartale laden ...")
		VSEmpf	:= VorsorgeKomplexe(Abrechnungsquartal, true)
		gosub abrGuiVorschlaege

	}
	else if (A_GuiControl = "abrRefresh")                                                          	{

		Gui, abr: Submit, NoHide
		Abrechnungsquartal := SubStr("0" abrQ, -3)
		SplashText("Vorbereitung", "Abrechnungsziffern des Quartals auffrischen ...")
		VSEmpf	:= VorsorgeKomplexe(Abrechnungsquartal, false)
		gosub abrGuiVorschlaege

	}
	else if (A_GuiControl = "abrBearbeitet")                                                      	{

			Gui, abr: Default
			Gui, abr: ListView, abrLV

		; alle abgehakten als versteckt markieren
			row := 0
			Loop {
				If !(row := LV_GetNext(row, "C"))
					break
				LV_GetText(PatID, row, 1)
				If RegExMatch(PatID, "^\d+$")
					VSEmpf[PatID].Hide := true
			}

		; verändertes Objekt speichern
			FSize := DBAccess.SaveVorsorgeKandidaten(VSEmpf, abrQuartal)

		; Inhalt neu anzeigen über universal Label
			gosub abrGuiVorschlaege

	}
	else if (A_GuiControl = "abrUnHide")                                                         	{

		; Objektflags auf false setzen
			For PatID, vs in VSEmpf
				If VSEmpf[PatID].Hide
					VSEmpf[PatID].Hide := false

		; verändertes Objekt speichern
			FSize := DBAccess.SaveVorsorgeKandidaten(VSEmpf, abrQuartal)

		; Inhalt neu anzeigen über universal Label
			gosub abrGuiVorschlaege

	}
	else if RegExMatch(A_GuiControl, "(abrYoungStars|abrAngels|abrOldStars)") 	{
		gosub abrGuiVorschlaege
	}
	else if (A_GuiControl = "abrICDKrauler")                                                       	{
		Gui, abr: Submit, NoHide
		chronischKrauler(abrQ2)
	}

return ;}

abrGuiVorschlaege:       	;{

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

		  ; Vorschläge vorschlagen (was auch sonst) ;{
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
		  ;}

		    f := vs.chronisch
			Loop 5
				f += e[A_Index] ? 1 : 0

			If (f > 0) {

				; Zeile anfügen
					Mortal 	:= PatDB[PatID].Mortal ? "♱" : ""
					lastcell := vs.chronisch	? "Chr: " vs.chronisch : ""
					lastcell .= vs.geriatrisch	? ", GB: " vs.geriatrsch : ""
					lastcell .= vs.Pauschale	? ", P: " vs.Pauschale : ""
					lastcell := LTrim(lastcell, ", ")
					LV_Add("Checked"	, PatID, Mortal . vs.Name, vs.Alter, vs.Geschl
									, e.1 ? "X":""
									, e.2 ? "X":""
									, e.3 ? "X":""
									, e.4 ? "X":""
									, e.5 ? "X":""
									, lastcell)

				; Zellen färben
					For eIdx, isEBM in e
						If isEBM
							 CLV.Cell(LV_GetCount(), 4+eIdx, "0x" tcolorM, 0x000000)

					If vs.chronisch
						 CLV.Cell(LV_GetCount(), 10, "0x" tcolorM, 0x000000)
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

abr_Tabs:

return

abrGuiClose:
abrGuiEscape: ;{
	SaveGuiPos(habr, "Abrechnungsassistent")

ExitApp ;}
}

QuartalsKalender(Quartal, Jahr, Urlaubstage, guiname:="QKal") {                                                	;-- zeigt den Kalender des Quartals

	; blendet Feiertage nach Bundesland aus (### ist nicht programmiert! ####)

	;--------------------------------------------------------------------------------------------------------------------
	; Variablen
	;--------------------------------------------------------------------------------------------------------------------;{
		global QKal, hmyCal, hCal, habr, gcolor, tcolor, CalQK
		global Cok, Can, Hinweis, MyCalendar, CalBorder
		global VSEmpf, abrCBX, VSRule, CV

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
		CalOptions := " vMyCalendar HWNDhmyCal gCalHandler"
		lp := GetWindowSpot(AbrhLV)

		Gui, abr: Font     	, s12 w600 cBlack, Futura Md Bk
		Gui, abr: Add     	, Text		, % "xm+5 y+8 vCalQK"                                       	, % "QUARTALSKALENDER"
		GuiControlGet, cp, abr: Pos, CalQK
		Gui, abr: Font     	, s7 Normal italic Black, Futura Bd Bt
		Gui, abr: Add     	, Text		, % "x+2 y" cpY+1                                                	, % "[" Quartal SubStr(Jahr, -1) "]"
		Gui, abr: Font     	, s14 w600 Black, Futura Md Bk
		Gui, abr: Add     	, Text		, % "x+1 y" cpY-4                                                	, % ":"
		Gui, abr: Font     	, s11 Normal Black, Futura Bd Bt
		Gui, abr: Add     	, Text		, % "x+8 y" cpY+1                                                	, % ConvertDBASEDate(begindate) "-" ConvertDBASEDate(lastdate)
		Gui, abr: Add     	, MonthCal, % "xm y+2 W-3 4 w620 "  CalOptions                         	, % begindate

	; Kalenderfarben
		SendMessage, 0x100A	, 0, 0x666666 ,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 1, 0xAA0000 ,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 2, 0x0101AA ,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 3, 0x01FF01 ,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 4, % "0x" RGBtoBGR(gcolor),, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 5, 0xFFAA99, SysMonthCal321, % "ahk_id " habr
		Sleep 100
		SendMessage, 0x100A	, 1, 0xAA0000 ,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 2, 0x0101AA ,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 3, 0x01FF01 ,, % "ahk_id " hmyCal
		SendMessage, 0x100A	, 4, % "0x" RGBtoBGR(gcolor),, % "ahk_id " hmyCal

Return ;}

CalHandler: ;{

	; angeklicketes Datum sichern                                                                    	;{
		Gui, abr: Default
		Gui, abr: Submit, NoHide
		DateChoosed := ConvertDBASEDate(StrSplit(MyCalendar,"-").1)
	;}

	; ist aktuelle eine Karteikarte geöffnet?                                                          	;{
		If !InStr(AlbisGetActiveWindowType(), "Karteikarte") {
			PraxTT("Es wird keine Karteikarte in Albis angezeigt`nWähle zunächst einen Patienten (Doppelklick)!", "5 1")
			return 0
		}
	;}

	; Patient der aktuellen Karteikarte ist gelistet und hat offene Komplexziffern? 	;{
		PatID := AlbisAktuellePatID()
		If (VSEmpf[PatID].Vorschlag.Count() = 0) && !VSEmpf[PatID].chronisch {
			PraxTT("Bei Patient [" PatID "] " PatDB[PatID].Name ", " PatDB[PatID].VORNAME " fehlt keine Eintragung.`nWähle einen anderen Patienten!", "5 1")
			return 0
		}
		EBMFees := GetEBMFees(PatID)
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

	; Nutzer nach den einzutragenden Ziffern befragen                                     	;{
		Gui, abr: Default
		Gui, abr: Submit, NoHide
	;}

	; EBM-Ziffern zusammen stellen                                                                  	;{
		EBMToAdd :="", HKSFormular := false
		If (VSEmpf[PatID].Vorschlag.Count() > 0)
			For feat, Xpos in abrCBX {
				GuiControlGet, ischecked, abr:, % "abr_" feat
				EBMNr := ischecked ? EBMFees[feat] "-" : ""
				If RegExMatch(EBMNr, "(01745|01746)") || RegExMatch(feat, "HKS")
					HKSFormular := true
				EBMToAdd .= EBMNr
			}

		EBMtoAdd := RTrim(RegExReplace(EBMtoAdd, "[\-]{2,}", "-"), "-")
	;}

	; Ziffern eintragen                                                                                        	;{
		lastPrgDate := AlbisSetzeProgrammDatum(DateChoosed)
		AlbisKarteikarteFocus("Kürzel")
		Send, % "lko"
		Sleep 50
		Send, {TAB}
		Sleep 100
		Send, % EBMtoAdd
		Sleep 100
		send, {TAB}
		Sleep 100
		If (VSEmpf[PatID].chronisch = 1) {
			Send, % "lko"
			Sleep 50
			Send, {TAB}
			Sleep 50
			Send, % "03221"
			Sleep 100
			Send, {TAB}
			Sleep 50
		}
		Send, {Esc}
	;}

	; Hautkrebscreening war dabei ? Dann erstelle bitte auch das Formular                         	;{
		If RegExMatch(EBMtoAdd, "(01745|01746)")
			if (hEHKS := AlbisMenu(34505, "Hautkrebsscreening - Nichtdermatologe ahk_class #32770", 3)) {
				AlbisHautkrebsScreening()   ; alles leer lassen, alles super!
			}
	;}

	; auf letztes Datum zurücksetzen und Ausnahmeindikationen korrigieren       	;{
		AlbisSetzeProgrammDatum(lastPrgDate)
		AlbisAusIndisKorrigieren()
	;}

	; MsgBX - melde das Du das fertig hast                                                       	;{
		MsgBox, 0x1024, % A_ScriptName, % "Bitte alles überprüfen!", 20
	;}

return ;}

CalendarOK: 				                                                                                                              	 ;{
	Gui, %gname%: Default
	Gui, %gname%: Submit

	If !InStr(MyCalendar,"-") {                                                                                                	; just in case we remove MULTI option
		FormatTime, CalendarM1, %MyCalendar%, dd.MM.yyyy
		CalendarM2:= CalendarM1
	}	Else {
		 FormatTime, CalendarM1, % StrSplit(MyCalendar,"-").1, dd.MM.yyyy
		 FormatTime, CalendarM2, % StrSplit(MyCalendar,"-").2, dd.MM.yyyy
	}

	CalendarM1:=CalendarM2:=0
return

CalendarCancel:
QKalGuiClose:
QKalGuiEscape:
	SaveGuiPos(hCal, "Kalender")
	Gui, %gname%: Destroy

Return
;}

}

GetEBMFees(PatID) {

	global	VSEmpf, DBAccess

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

return EBMFees
}

VorsorgeKomplexe(Abrechnungsquartal, ReIndex:=false) {

	global DBAccess

	DBAccess 	:= new AlbisDb(Addendum.AlbisDBPath, 0)

	If !ReIndex && FileExist(Addendum.DBasePath "\VorsorgeKandidaten-" Abrechnungsquartal ".json") {
		SplashText("bitte warten ...", "lade Vorsorge Kandidaten")
		VSEmpf := JSONData.Load(Addendum.DBasePath "\VorsorgeKandidaten-" Abrechnungsquartal ".json", "", "UTF-8")
		SplashText("bitte warten ...", "Vorsorge Kandidaten geladen")
	}
	If !IsObject(VSEmpf) || ReIndex {

		SplashText("bitte warten ...", "EBM Ziffern aller Quartale werden gesammelt ...")
		VSDaten 	:= DBAccess.VorsorgeDaten(ReIndex, ReIndex ? true : false)

		SplashText("bitte warten ...", "EBM Ziffern aller Quartale werden gefiltert ...")
		VSEmpf   	:= DBAccess.VorsorgeFilter(VSDaten, Abrechnungsquartal, false, ReIndex ? true : false)

		SplashText("bitte warten ...", "Filter ist angewendet worden ...")

	}

return VSEmpf
}

chronischKrauler(Abrechnungsquartal) {                                                                                       ;--

	; sucht alle ICD Abrechnungsdiagnosen eines Quartals heraus und vergleicht diese mit einer ICD Liste
	; 2 Datenbankdateien müssen für die Suche ausgelesen werden: 1.BEFUND.dbf und 2.BEFTEXTE.dbf
	; die Eingrenzung der Suche innerhalb der BEFUND.dbf über das Abrechnungsquartal ist nicht anwendbar,
	; Albis hinterlegt dort nicht manchmal die Quartalbezeichnung meist steht dort eine 0 oder eine andere Zahl
	; das Abrechnungsquartal wird in ein Datum in Form eines RegEx-String umgewandelt aus
	; 0320 wird 2020(07|08|09)\d\d. Damit sind alle Tage des Quartals erfasst.

		global DBAccess
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

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Hilfsfunktionen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
SaveGuiPos(hwnd, Fenstername) {

	win 			:= GetWindowSpot(hwnd)
	winPos	 	:= "x" win.X " y" win.Y " w" win.CW " h" win.CH
	IniWrite, % winpos	, % Addendum.Ini, % Addendum.compname, %  Fenstername "_Position"

}

MessageWorker(InComing) {                                                                                    	;-- verarbeitet die eingegangen Nachrichten

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

RGBtoBGR(cRGB) {
	RegExMatch(cRGB, "i)^(?<Prepend>(0x|#))*(?<Red>[A-F\d]{2})(?<Green>[A-F\d]{2})(?<Blue>[A-F\d]{2})", c)
return cPrepend . cBlue . cGreen . cRed
}

PosStringToObject(string) {

	p := Object()
	For wIdx, coord in StrSplit("XYWH") {
		RegExMatch(string, "i)" coord "(?<Pos>\d+)", w)
		p[coord] := wPos
	}

return p
}

SplashText(title, msg) {
	SplashTextOn, 300, 20, % title, % msg
	hSTO := WinExist("Vorbereitung ahk_class AutoHotkey2")
	WinSet, Style 	, 0x50000000, % "ahk_id " hSTO
	WinSet, ExStyle	, 0x00000208, % "ahk_id " hSTO
}


Thread_CoronaEBM:        ;{  alle Patienten mit welche einen Abstrich erhalten haben oder wegen Corona behandelt wurden. Fehlende Ziffern (88240, 32006) finden

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


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Includes
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
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





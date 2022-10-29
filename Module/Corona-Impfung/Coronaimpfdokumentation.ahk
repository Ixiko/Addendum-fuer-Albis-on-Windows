; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;                                    Addendum Assistent für die Corona Impfdokumentation
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;      Funktion: 				 ⚬	Erstellung und Druck aller Formulare (liegen als umgewandelte PDF -> Word-Dokumente vor)
;                                	 ⚬	Anlegen eines Abrechnungsscheines für Privatversicherte für die Abrechnung der Impfungen über die KV
;                                	 ⚬	Eintragen aller Abrechnungsziffern, Kasse sowie Privat
;                                	 ⚬	Eintragen der Impfdiagnose
;                                	 ⚬	erkennt geimpfte Personen (insofern Impfziffern eingetragen wurden)
;
;
;
;		Hinweis:
;
;		Abhängigkeiten:		siehe includes
;
;      	begonnen:       	    	01.05.2021
; 		letzte Änderung:	 	11.05.2021
;
;	  	Addendum für Albis on Windows by Ixiko started in September 2017
;      	- this file runs under Lexiko's GNU Licence
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣

  ; Einstellungen 	;{
	global starttime    	:= A_TickCount

	#NoEnv
	#Persistent
	#KeyHistory, Off

	SetBatchLines, -1
	ListLines    	, Off

	If !pToken := Gdip_Startup() {
		MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
		ExitApp
	}

  ;}

  ; Variablen      	;{
	global PatDB           							 	; Patientendatenbank
	global Addendum
	global CIProps       							  	; Gui Einstellungen und anderes
	global KK 				:= {"Arzt":"Edit4", "Datum":"Edit5", "Kuerzel":"Edit6", "Inhalt":"RichEdit20A1"}

  ; startet Windows Gdip
   	If !(pToken:=Gdip_Startup()) {
		MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
		ExitApp
	}

  ; Tray Icon erstellen
	If (hIconImpfen := Create_Impfen_ico())
    	Menu, Tray, Icon, % "hIcon: " hIconImpfen

  ; Albis Datenbankpfad / Addendum Verzeichnis
	RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
	RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)

	Addendum                 	:= Object()
	Addendum.Dir            	:= AddendumDir
	Addendum.Ini              	:= AddendumDir "\Addendum.ini"
	Addendum.DBPath      	:= AddendumDir "\logs'n'data\_DB"
	Addendum.DBasePath  	:= AddendumDir "\logs'n'data\_DB\DBase"
	Addendum.AlbisDBPath	:= AlbisPath "\DB"
	Addendum.AlbisBriefe	:= AlbisPath "\Tvl"
	Addendum.compname	:= StrReplace(A_ComputerName, "-")                                                                 	; der Name des Computer auf dem das Skript läuft
	Addendum.propsPath 	:= A_ScriptDir "\Coronaimpfdokumentation.json"

	SciTEOutput()

  ; Patientendatenbank
	;outfilter 	:= ["NR", "PRIVAT", "ANREDE", "NAME", "VORNAME", "GEBURT", "PLZ", "ORT", "STRASSE", "TELEFON", "GESCHL", "TELEFON2", "TELEFAX", "LAST_BEH", "DEL_DATE", "MORTAL", "HAUSNUMMER"]
	;PatDB 		:= ReadPatientDBF(Addendum.AlbisDBPath,, outfilter)

  ; Gui und notwendige Daten laden z.B. die Chargennummern
	If FileExist(Addendum.propsPath) {
		CIProps := JSONData.Load(Addendum.propsPath, "", "UTF-8")
	} else {
		CIProps := {"CNR":[], "GuiPos":{}}
		; Beispielbefüllung
		CIProps.CNR[14] := ["ET3045",""]
		CIProps.CNR[15] := ["EX3510", "ABW7189"]
		CIProps.CNR[16] := ["EX3510",""]
	}

 ; hier im Moment die Dokumentnamen eintragen
 ; müssen in albiswin\tvl vorliegen ;{
	mRNA_Anamnese 	:= "Corona_Anamnese-mrna.docx"
	mRNA_Aufklaerung 	:= "Corona_Aufklaerungsbogen-mrna.docx"
	mRNA_Einwilligung	:= "Corona_Einwilligung_mRNA-Impfstoff.docx"
	Vektor_Anamnese 	:= "Corona_Anamnese-Vektor.docx"
	Vektor_Aufklaerung 	:= "Corona_Aufklaerungsbogen-Vektor.docx"
	Vektor_Einwilligung	:= "Corona_Einwilligung_Vektor-Impfstoff.docx"
	;}

	If !IsObject(CIProps.Boegen) {
		CIProps.Boegen     	:= {	"mRNA" : {"Einwilligung":mRNA_Einwilligung, "Anamnese":mRNA_Anamnese, "Aufklärung":mRNA_Aufklaerung}
									        ,	"Vektor" : {"Einwilligung":Vektor_Einwilligung, "Anamnese":Vektor_Anamnese, "Aufklärung":Vektor_Aufklaerung}}
	}

  ; Impfleistungen Kasse  ### NICHT ÄNDERN!!
	If !IsObject(CIProps.Kasse) {
		CIProps.Kasse      		:= {	"#1"	: "88331"
											, 	"#2"	: "88333"
											, 	"#3"	: "88332"
											, 	"#4"	: "88334"
											,	"Hausbesuch"  	: ["88323", "88324"]
											, 	"Indikation"    	: {"Normal"    	: ["A", "B"]
														        		,	"Pflegeheim"	: ["G", "H"]
														        		,	"Beruf"          	: ["V", "W"]}}
	}

  ; Template Diagose für die Eintragungen
	If !CIProps.haskey(Diagnose)
		CIProps.Diagnose := "Impfung gegen COVID-19 #Impfstoffbezeichnung# Ch.: #Chargennummer# {U11.9G};"

 ; Impfstoffe erscheinen hier in eigener Sortierung
	;If !IsObject(CIProps.Impfstoffe) {
		CIProps.Impfstoffe 	:= {"#1": {"Hersteller"   	: "BionTech/Pfizer"
										        	,	"Handelsname"	: "Comirnaty"
										        	,	"SNR"            	: "88331"
													,	"Art"                 	: "mRNA"
													,	"Anamnese"		: mRNA_Anamnese
													,	"Aufklaerung"	: mRNA_Aufklaerung
													,	"Einwilligung"	: mRNA_Einwilligung
										        	,	"REF"             	: "XIMPFCORONA1"}
											, "#2": {"Hersteller"    	: "AstraZeneca"
													,	"Handelsname"	: "Vaxzevria"
													,	"SNR"            	: "88333"
													,	"Art"                 	: "Vektor"
													,	"Anamnese"		: Vektor_Anamnese
													,	"Aufklaerung"	: Vektor_Aufklaerung
													,	"Einwilligung"	: Vektor_Einwilligung
													,	"REF"             	: "XIMPFCORONA2"}
											, "#3": {"Hersteller"    	: "Moderna"
													, 	"Handelsname"	: "Moderna"   ;Covid-19 Vaccine
													,	"SNR"            	: "88332"
													,	"Art"                 	: "mRNA"
													,	"Anamnese"		: mRNA_Anamnese
													,	"Aufklaerung"	: mRNA_Aufklaerung
													,	"Einwilligung"	: mRNA_Einwilligung
													,	"REF"                	: "XIMPFCORONA3"}
											, "#4": {"Hersteller"		: "Johnson & Johnson"
													,	"Handelsname"	: "Covid-19 Vaccine Johnson & Johnson"
													,	"SNR"            	: "88334"
													,	"Art"                 	: "Vektor"
													,	"Anamnese"		: Vektor_Anamnese
													,	"Aufklaerung"	: Vektor_Aufklaerung
													,	"Einwilligung"	: Vektor_Einwilligung
													,	"REF"             	: "XIMPFCORONA4"}}
	; }

  ;}

  ; Gui aufrufen
	CoronaImpfung()

return

CoronaImpfung() {

; GUI                                        	;{

	; VARIABLEN                                                               	;{

		global

		static CNRListe	:= []
		static TitleW 		:= 580
		static col1W 		:= 190
		static chNrW 	:= 14*11
		static tr          	:= " Right "
		static tc         	:= " Center "
		static bgt       	:= " BackgroundTrans ", OCW := "cWhite", OCB := "cBlack"
		static ex           	:= " +0x00000000 "
		static tm            := " +0x02000000 +E0x20 "
		static tl            	:= " +E0x00080000 "
		static gH       	:= " gCIButtonHandler "
		static gCHR      	:= " AltSubmit gCIChargennummern "
		static gTD       	:= " AltSubmit gCITagesdatum "
		static btn          	:= ["CIDate", "CIBTN", "CIBTE", "CIBT1", "CIBT2", "CIAZN", "CIAZE", "CIAZ1", "CIAZ2"]
		static disButtons	:= ["CIBT", "CIAZ"]
		static CNRCtrls 	:= ["CIBTE", "CIAZE"]


	;}

	; aktuellen Fokus des Albisfenster merken                    	;{
		ControlGetFocus, hCtrlActive, ahk_class OptoAppClass
	;}

	; aktuelle Kalenderwoche berechnen                             	;{
		Tagesdatum := A_DD "." A_MM "." A_YYYY
		If RegExMatch(Tagesdatum, "(?<day>\d\d)\.(?<month>\d\d)\.(?<year>\d\d\d\d)", _) {
			guiDate	:= _year _month _day
			weekNr	:= SubStr("0" WeekOfYear(guiDate), -1)
		}
	;}

	; ComboxStrings erstellen                                          	;{
		For CNRWeek, VacList in CIProps.CNR {
			For VacNr, CNRs in VacList
				For idx, CNR in StrSplit(CNRs, "|")
					CNRListe[VacNr] .= CNRWeek ": " CNR "|"
		}
		Loop % CIProps.CNR.Count() {
			liste := RTrim(CNRListe[A_Index], "|")
			Sort, liste, D| U N R
			CNRListe[A_Index] := liste
		}
	;}

	; aktuelle Daten vom Albisfenster                               	;{
		PatID     	:= AlbisAktuellePatID()
		PatName 	:= AlbisCurrentPatient()
		PatGeburt	:= AlbisPatientGeburtsdatum()
	;}

	; Nachfragen bei als verstorben markierten Patienten  	;{
		If PatDB[PatID].MORTAL {
			MsgBox, 0x1024, % RegExReplace(A_ScriptName, "\.*$"), % "Möchten Sie wirklich bei einem`nverstorbenen Patienten eine Impfdokumentation anlegen?"
			IfMsgBox, No
				ExitApp
		}
	;}

	; Impfliste aus Albisdatenbankeinträgen erstellen        	;{
		Impflinge	:= new ImpfListe
		ImpfNr  	:= Impflinge.Vaccinations(PatID)
		ImpfTag	:= Impflinge.VaccinationDates(PatID)
		Impfstoff	:= Impflinge.Vaccine(PatID)
		Impfzahl	:= Impflinge.VaccinationsCount()

		JSONData.Save(A_Temp "\Impflinge.json", Impflinge, true,, 1, "UTF-8")
		;~ Run % A_Temp "\Impflinge.json"

	;}

	; GUI ERSTELLEN                                                	    	;{

	;-: Steuerelemente ;{
		Gui, CI: New 	, -DPIScale +HWNDhCI +AlwaysOnTop +OwnDialogs
		Gui, CI: Margin	, 0, 0
		Gui, CI: Font		, % "s13 cWhite", % "Futura Md Bt"
		;Gui, CI: Add  	, Progress	, % "x0 y0 w" TitleW+10 " h50 C007300 vCIPGS1 " ex                      	, 100
		Gui, CI: Add		, Picture	, % "x0 y0 w" TitleW+10 " h50 0xE vCIPGS1 HWNDCIhPGS1"
		GdipBgr(CIhPGS1, 0xff007300)
		Gui, CI: Add  	, Text    	, % "x5 y5 w" TitleW " " tc bgt " vCITL1" tm                                          	, % "IMPFUNG GEGEN COVID-19"
		Gui, CI: Font		, % "s12 cWhite", % "Futura Bk Bt"
		Gui, CI: Add  	, Text    	, % "y+0 w" TitleW " " tc bgt " vCITL2 " tm                                         	, % "[" weekNr ". Kalenderwoche " _year "]"

		;{
		GuiControlGet, cp, CI: Pos, CIPGS1
		PG2Y := cpY + cpH
		;}

		;Gui, CI: Add  	, Progress	, % "x0 y" PG2Y " w" TitleW+10 " h60 C0CF09C vCIPGS2 " ex            	, 100
		Gui, CI: Add  	, Picture	, % "x0 y" PG2Y " w" TitleW+10 " h60 0xE vCIPGS2 HWNDCIhPGS2"
		GdipBgr(CIhPGS2, 0xffC0CF09C)
		Gui, CI: Font		, % "s12 cBlack"	, % "Futura Bk Bt"
		Gui, CI: Add  	, Text    	, % "x5  	y" PG2Y+5 " w" col1W  " " tr bgt tm                                  	, % "Patient:"
		Gui, CI: Add  	, Text    	, % "x+5 " bgT tm                                                                             	, % "[" PatID "] " PatName ",  Impfungen: " ImpfNr
		Gui, CI: Add  	, Text    	, % "x5  	y+7 w" col1W  " " tr bgt tm                                                	, % "Tagesdatum:"
		Gui, CI: Font		, % "s11 cBlack"	, % "Futura Bk Bt"
		Gui, CI: Add  	, Edit	    	, % "x+5 yp-3 w95 Center vCIDate " gTD bgt tm                                	, % Tagesdatum

		;{
		GuiControlGet, cp, CI: Pos, CIPGS2
		PG2Y := cpY+cpH
		;}

		;Gui, CI: Add  	, Progress    	, % "x0 y" PG2Y " w" TitleW+10 " h95 C16A8D9 vCIPGS3 " ex			, 100
		Gui, CI: Add  	, Picture     	, % "x0 y" PG2Y " w" TitleW+10 " h95 0xE vCIPGS3 HWNDCIhPGS3"
		GdipBgr(CIhPGS3, 0xff16A8D9)
		Gui, CI: Font		, % "s11"    	, % "Futura Bk Bt"
		Gui, CI: Add  	, Text        	, % "x5 y" PG2Y+20 " w" col1W " vCIBTN " tr bgt tm                       	, % "Biontech [Comirnaty Ch.:"
		GuiControlGet, cp, CI: Pos, CIBTN
		Gui, CI: Add  	, ComboBox 	, % "x+5 y" cpY-3 " w" chNrW " vCIBTE hwndCIhBTE " tm gCHR      	, % CNRListe.1
		GuiControl, CI: Choose, CIBTE, 1
		Gui, CI: Add  	, Text        	, % "x+2 y" cpY " " bgt tm                                                              	, % "]"

		Gui, CI: Add  	, Button     	, % "x+5 y" cpY-6 " vCIBT1" gH tc                                                	, % "1.Impfdosis   "
		Gui, CI: Add  	, Button     	, % "x+5 y" cpY-6 " vCIBT2" gH tc                                                	, % "2.Impfdosis   "

		Gui, CI: Add  	, Text        	, % "x5 y+10 w" col1W " vCIAZN " tr bgt tm                                   	, % "AstraZeneca [Vaxzevria Ch.:"
		GuiControlGet, cp, CI: Pos, CIAZN
		Gui, CI: Add  	, ComboBox 	, % "x+5 y" cpY-3 " w" chNrW " vCIAZE hwndCIhAZE " tm gCHR   	, % CNRListe.2
		GuiControl, CI: Choose, CIAZE, 1
		Gui, CI: Add  	, Text        	, % "x+2 " bgt tm                                                                          	, % "]"
		Gui, CI: Add  	, Button     	, % "x+5 y" cpY-6 " vCIAZ1" gH tc                                                	, % "1.Impfdosis   "
		Gui, CI: Add  	, Button     	, % "x+5 y" cpY-6 " vCIAZ2" gH tc                                                 	, % "2.Impfdosis   "

		;{
		GuiControlGet, cp, CI: Pos, CIPGS3
		PG2Y := cpY+cpH
		;}

		;Gui, CI: Add  	, Progress    	, % "x0 y" PG2Y " w" TitleW+10 " h35 C0E6D8C vCIPGS4 " ex    		, 100
		Gui, CI: Add  	, Picture     	, % "x0 y" PG2Y " w" TitleW+10 " h35 0xE vCIPGS4 HWNDCIhPGS4"
		GdipBgr(CIhPGS4, 0xff0E6D8C)
		Gui, CI: Font		, % "s11 cWhite"	, % "Segoe UI"
		Gui, CI: Add   	, Text            , % "x5 y" PG2Y+5 " w" col1W " " tr bgt tm                               		, % "durchgeführte Impfungen:"
		Gui, CI: Add   	, Text            , % "x+5 " bgt tm                                                                  	     	, % Impfzahl
		Gui, CI: Font		, % "s8 cBlack"	, % "Segoe UI"
		Gui, CI: Add   	, Text            , % "x" TitleW+10-85 " y" PG2Y+20 " " bgt tm                            	 	, % "Ladezeit: "
		Gui, CI: Add   	, Text            , % "x+1 vCILTime " bgt tm                                                     		, % "0.00s  "

		;}

	;-: Steuerelementinteraktionen einschränken wenn schon geimpft oder der Impfabstand zu gering ist
		Loop % ImpfNr {
			Nr := A_Index
			For btnIndex, btnName in disButtons {
				; wurde geimpft
					GuiControl, CI: Disable, % (btnName . Nr)
					If (btnIndex = Impfstoff[Nr])
						GuiControl, CI: Text, % (btnName . Nr), % ConvertDBASEDate(ImpfTag[Nr])
					else
						GuiControl, CI: Hide, % (btnName . Nr)
			}
		}

	;-: Wiederholungsimpfungen
		gosub CITagesdatum

		Gui, CI: Show	, AutoSize,	 Impfassistent

		OnMessage(0x47	, "GuiPosChanging") ; WM_WINDOWPOSCHANGED
		OnExit("Feierabend")


		hEdit1 := GetComboBoxEdit(CIhAZE)
		hEdit2 := GetComboBoxEdit(CIhBTE)
		GuiControl, CI:, CILTime, %  Round((A_TickCount-starttime)/1000, 2) "s"
		;}

	;}


return ;}

; -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

CIChargennummern:                 	;{                 validiert die Eingaben der Chargennummern, zeigt vor jeder Chargennummer die zugehörige Kalenderwoche

		Gui, CI: Submit, NoHide

		chwnd   	:= (A_GuiControl = "CIBTE" ? hEdit2 : hEdit1)
		charpos  	:= CaretPos(chwnd)-1
		CharPos	:= CharPos = 1 ? 5 : CharPos
		chnr      	:= %A_GuiControl%
		chnr      	:= Format("{:U}", chnr)

		;SciTEOutput(A_GuiControl ": " A_EventInfo ", " A_GuiEvent ", chnr: " chnr)
		If RegExMatch(chnr, "^\d+$") {
			GuiControl, CI: Choose, % A_GuiControl, % chnr
			return
		}

		RegExMatch(chnr, "O)^(?<KW>\d+)*\s*:*\s*(?<CNRA>[A-Z]+)*(?<CNRB>\d+)*", rx)
		If (CharPos > 4)
			chnr := (!rx.KW ? weekNr : rx.KW) ": " rx.CNRA rx.CNRB
		If (A_GuiControl = "CIBTE") {
			ControlSetText,, % chnr, % "ahk_id " hEdit2
			SetCaretPos(hEdit2, CharPos)
		}
		else If (A_GuiControl = "CIAZE") {
			ControlSetText,, % chnr, % "ahk_id " hEdit1
			SetCaretPos(hEdit1, CharPos)
		}

return ;}

CIButtonHandler:                       	;{

	; auf welchen Button wurde gedrückt?   											;{
		thisguicontrol := A_GuiControl
		Gui, CI: Submit, NoHide
		RegExMatch(thisguicontrol, "Oi)(?<name>[A-Z]+)(?<nr>\d)", Btn)
	;}

	; wird nur ausgeführt wenn das Datumformat stimmt 						;{
		If !RegExMatch(CIDate, "\d\d\.\d\d\.\d\d\d\d") {
			MsgBox, 0x1000, % StrReplace(A_ScriptName, ".ahk"), % "Das Tagesdatum hat ein falsches Format.`nEs sollte so aussehen dd.MM.YYYY"
			return
		}
	;}

	; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Daten für Impfdokumentation zusammenstellen
	; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		For ImpfstoffNr, GuiButton in disButtons
			If (GuiButton = Btn.name)
				break

		vaccination                	:= CIProps.Impfstoffe["#" ImpfstoffNr]
		vaccination.CNR 	    	:= CNR.Nr
		vaccination.Diagnose 	:= CIProps.Diagnose
		vaccination.NR           	:= Btn.Nr

		vaccdocs :=[	vaccination["Anamnese"]
						, 	vaccination["Aufklaerung"]
						, 	vaccination["Einwilligung"] ]
	;}

	; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Chargennummer lesen
	; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		CNRBtn 	:= Btn.name "E", chwnd := (CNRBtn = "CIBTE" ? hEdit2 : hEdit1)

		ControlGet, CNR, Line, 1,, % "ahk_id " chwnd
		CNR      	:= ValidateChargenNummer(CNR)

	; Chargennummer ist nicht korrekt oder leer
		If !CNR.Nr {
			MsgBox, 0x1004, % RegExReplace(A_ScriptName, "\.*$"), % "Ohne Chargennummer können`nkeine Gebühren abgerechnet werden?", 6
		}
	; Chargennummer speichern
		else {

			; KW der Chargennummer und die aktuelle KW unterscheiden sich
				If (CNR.KW <> weekNr) {

					; gespeicherte Chargennummer und die Neue unterscheiden sich
						If (CIProps.CNR[CNR.KW][ImpfstoffNr] <> CNR.Nr) {

								msg := "Soll die Chargennummer [ " CIProps.CNR[CNR.KW][ImpfstoffNr] " ] der " CNR.KW ".KW`n"
										.	" mit [ " CNR.Nr " ] ersetzen werden?`n"
										.	"Wenn nicht wird die Ch.Nr. für die aktuelle Woche gesichert."

								MsgBox, 0x1003, % RegExReplace(A_ScriptName, "\.*$"), % msg
								IfMsgBox, Cancel
									return
								IfMsgBox, Yes
								{
										If !IsObject(CIProps.CNR[weekNr])
											CIProps.CNR[weekNr] := Array()
										CIProps.CNR[weekNr][ImpfstoffNr] := CNR.Nr
								}
								IfMsgBox, No
								{
										If !IsObject(CIProps.CNR[CNR.KW])
											CIProps.CNR[CNR.KW] := Array()
										CIProps.CNR[CNR.KW][ImpfstoffNr] := CNR.Nr
								}

						}
					; Chargennummer wurde geändert
						else {

								msg := "Soll die Chargennummer [ " CNR.Nr " ]`n"
										. 	"für die aktuelle " weekNr ".KW übernommen werden?`n"
								MsgBox, 0x1004, % RegExReplace(A_ScriptName, "\.*$"), % msg
								IfMsgBox, No
									return
								If !IsObject(CIProps.CNR[weekNr])
									CIProps.CNR[weekNr] := Array()
								CIProps.CNR[weekNr][ImpfstoffNr] := CNR.Nr

						}

				}
			; KW Chargennummer entspricht der KW des Gui-Tagesdatum
				else {

					If !IsObject(CIProps.CNR[weekNr])
						CIProps.CNR[weekNr] := Array()
					CIProps.CNR[weekNr][ImpfstoffNr] := CNR.Nr

				}

			JSONData.Save(Addendum.propsPath, CIProps, true,, 1, "UTF-8")
		}
	;}

	; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Daten für Impfdokumentation zusammenstellen
	; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		For ImpfstoffNr, GuiButton in disButtons
			If (GuiButton = Btn.name)
				break

		vaccination                	:= CIProps.Impfstoffe["#" ImpfstoffNr]
		vaccination.CNR 	    	:= CNR.Nr
		vaccination.Diagnose 	:= CIProps.Diagnose
		vaccination.NR           	:= Btn.Nr

		vaccdocs :=[	vaccination["Anamnese"]
						, 	vaccination["Aufklaerung"]
						, 	vaccination["Einwilligung"] ]

		;}

		SciTEOutput("INr: " ImpfstoffNr ", " Btn.name "," vaccination.Handelsname "," vaccdocs.1 "," vaccdocs.2   )

 ;}

GuiDokumentwahl:                    	;{

		CIPos := GetWindowSpot(hCI)
		Gui, CID: New 		, -Border -DPIScale +HWNDhCID +Owner%hCI%
		Gui, CID: Margin	, 10, 10
		Gui, CID: Color 	, c16A8D9, c0E6D8C
		Gui, CID: Font		, % "s20 cBlack underline", % "Futura Md Bt"
		Gui, CID: Add    	, Text, % "xm ym w" CIPos.CW-20 " " bgt tc          	, % "             Auswahl Impfdokumentation             "
		Gui, CID: Show		, Hide
		Gui, CID: Font		, % "s12 cSeaBlue normal", % "Futura Md Bt"
		Gui, CID: Add    	, Text, % "xm y+0  w" CIPos.CW-20 " " bgt tc         	, % "für " Vaccination.Art "-Impfstoff [" Vaccination.Handelsname " von " Vaccination.Hersteller "]"

		Gui, CID: Font		, % "s14 cBlack underline", % "Futura Md Bt"
		halfW := Floor(CIPos.W/2-10)
		Gui, CID: Add    	, Text, % "xm y+15 vCIDDPrintT  " bgt tc                                                                    	, % "Dokumente                    "
		Gui, CID: Show  	, AutoSize Hide
		GuiControlGet, fp, CID: Pos, CIDDPrintT

		CIChecked := []
		Gui, CID: Font		, % "s11 cBlack normal", % "Futura Bk Bt"
		Gui, CID: Add    	, Checkbox, % "xm+10 y+5 	vCIDD1 HWNDCIhCheck" bgt                                         	, % "Anamnesebogen "        	;Vaccination.Art "-Imfpstoff"
		CIChecked.Push(CIhCheck)
		Gui, CID: Add    	, Checkbox, % "y+5 				vCIDD2 checked HWNDCIhCheck" bgt                             	, % "Aufklärungsmerkblatt " 	;Vaccination.Art "-Imfpstoff"
		CIChecked.Push(CIhCheck)
		Gui, CID: Add    	, Checkbox, % "y+5 				vCIDD3 checked HWNDCIhCheck" bgt                            	, % "Einwilligungsbogen "   	;Vaccination.Art "-Imfpstoff"
		CIChecked.Push(CIhCheck)
		If (Btn.Nr < 2) {
			Gui, CID: Add  	, Checkbox, % "y+5 	    		vCIDD4 checked HWNDCIhCheck" bgt                            	, % "Termin 2.Impfung "     	;Vaccination.Art "-Imfpstoff"
			CIChecked.Push(CIhCheck)
		}

		Gui, CID: Font		, % "s14 cBlack underline", % "Futura Md Bt"
		Gui, CID: Add    	, Text		, % "x" (halfW+10) " y" fpY " Right vCIDAW " bgt                                           	, % "                         Auswahl"

		Gui, CID: Font		, % "s11 cBlack normal", % "Futura Bk Bt"
		Gui, CID: Add    	, Button 	, % "y+5 h28 vCIDDPrint 	gCIDBTNHandler"                                                	, % "nur Dokumente drucken"
		Gui, CID: Add    	, Button 	, % "y+5 h28 vCIDLKO  	gCIDBTNHandler"                                                	, % "nur Leistungen eintragen"
		Gui, CID: Add    	, Button 	, % "y+5 h28 vCIDKDok 	gCIDBTNHandler"                                                	, % "Impfung komplett vorbereiten"
		; ist nur aktiviert wenn die Chargennummer dieser Woche eingeben wurde
		If !CNR.Nr {
			GuiControl, CI: Disable, CIDKDok
			GuiControl, CI: Disable, CIDLKO
		}

		;{
		GuiControlGet, fp, CID: Pos, CIDAW
		GuiControl, CID: Move, CIDAW 	, % "x" CIPos.W-20-fpW
		GuiControlGet, fp, CID: Pos, CIDDPrint
		GuiControl, CID: Move, CIDDPrint	, % "x" CIPos.W-20-fpW
		GuiControlGet, fp, CID: Pos, CIDLKO
		GuiControl, CID: Move, CIDLKO	, % "x" CIPos.W-20-fpW
		GuiControlGet, fp, CID: Pos, CIDKDok
		GuiControl, CID: Move, CIDKDok	, % "x" CIPos.W-20-fpW
		WinSet, Style, 0x94000000, % "ahk_id " hCID
		;}

		Gui, CID: Show		, % "x" CIPos.X+3 " y" CIPos.Y+CIPos.H-3 " w" CIPos.CW-CIPos.BW*2-20 " AutoSize"	, Impf-Dokumentation


return ;}

CIDBTNHandler: 							;{             	Automatisierung der Impfdokumentation

	; wird nur ausgeführt wenn das Datumformat stimmt
		CIDGuiControl := A_GuiControl
		Gui, CI: Submit, NoHide
		If !RegExMatch(CIDate, "\d\d\.\d\d\.\d\d\d\d") {
			MsgBox, 0x1000, % StrReplace(A_ScriptName, ".ahk"), % "Das Tagesdatum hat ein falsches Format.`nEs sollte so aussehen dd.MM.YYYY"
			return
		}

		Loop 3 {
			ControlGet, isChecked, Checked,,, % "ahk_id " CIChecked[A_Index]
			SciTEOutput("ischecked " A_Index "(" CIChecked[A_Index] ") :" isChecked)
			If !isChecked
				vaccdocs[A_Index] := ""
		}
		If CIDD4
			vaccdocs[4] := "Corona_" Btn.Nr ".Impftermin.docx"

	;{
		;WinHideRoll(hCID, 600, "BT")
		;~ Gui, CID: Show, Hide
		;~ CIPos 	:= GetWindowSpot(hCI)
		;~ SetWindowPos(hCI, A_ScreenWidth - CIPos.W - 2, A_ScreenHeight - CIPos.H - 30, CIPos.W, CIPos.H)
		;WinSet, Style, 0x94000000, % "ahk_id " hCID
		;}


	; ab hier wird die Automatisierung gestartet
		Switch CIDGuiControl {

			Case "CIDDPrint":
				; druckt alle ausgewählten Dokumente
					PraxTT("1/1 Drucke die Impfdokumente", "20 1")
					PrintDocuments(PatID, PatName, PatGeburt, vaccdocs)


			Case "CIDLKO":
				; prüft ob ein Abrechnungsschein angelegt ist
					PraxTT("1/3 Prüfe auf angelegten Abrechnungsschein", "10 1")
					If !AbrScheinAngelegt(CIDate)
						return

				; Abrechnungsziffern eintragen
					PraxTT("1/1 Abrechnungsziffern werden eingetragen", "20 1")
					AlbisCImpfDokumentation(CIDate, vaccination)


			Case "CIDKDok":

				; prüft ob ein Abrechnungsschein angelegt ist
					PraxTT("1/3 Prüfe auf angelegten Abrechnungsschein", "10 1")
					If !AbrScheinAngelegt(CIDate)
						return

				; druckt alle ausgewählten Dokumente
					PraxTT("2/3 Drucke die Impfdokumente", "20 1")
					PrintDocuments(PatID, PatName, PatGeburt, vaccdocs)

				; Abrechnungsziffern eintragen
					PraxTT("3/3 Abrechnungsziffern werden eingetragen", "20 1")
					AlbisCImpfDokumentation(CIDate, vaccination)

		}

		CIPos     	:= GetWindowSpot(hCI)
		CIDPos 	:= GetWindowSpot(hCID)
		SetWindowPos(hCI, A_ScreenWidth - CIPos.W - 2, A_ScreenHeight - CIPos.H - CIDPos.H - 30, CIPos.W, CIPos.H)
		Gui, CID: Show
		;~ WinShowRoll(hCID, 600, "TB")
		;CIGui_RePaint()

return ;}

CIDGuiClose: 								;{
CIDGuiEscape:
	ExitApp
return
;}

GuiShowCID:                            	;{
	CIPos := GetWindowSpot(hCI)
	Gui, CID: Show, % "x" CIPos.X " y" CIPos.Y+CIPos.H-3 " Hide"
	WinShowRoll(hCID, 300, "TB")
return ;}

CITagesdatum:                      		;{

	; wird nur ausgeführt wenn das Datumformat stimmt
		Gui, CI: Submit, NoHide
		If !RegExMatch(CIDate, "\d\d\.\d\d\.\d\d\d\d") {
			MsgBox, 0x1000, % StrReplace(A_ScriptName, ".ahk"), % "Das Tagesdatum hat ein falsches Format.`nEs sollte so aussehen dd.MM.YYYY"
			return
		}

	; Datum umformatieren und damit das Wochendatum bestimmen ;{
		;RegExMatch(CIDateEntry, "(?<Day>\d\d)\.(?<Month>\d\d)\.(?<Year>\d\d\d\d)", g)
		gDay 	:= SubStr(CIDate, 1, 2)
		gMonth	:= SubStr(CIDate, 4, 2)
		gYear 	:= SubStr(CIDate, 7, 4)
		guiDate	:= gYear . gMonth . gDay
		weekNr	:= SubStr("0" WeekOfYear(guiDate), -1)
	;}

	; Kalenderwoche anzeigen ;{
		GuiControl, CI:, CITL2, % "[" weekNr ". Kalenderwoche " gYear "]"
		GuiControlGet, chwnd, CI: Hwnd, CITL2
		Sleep 100
		WinSet, Style    	, 0x50000001 , % "ahk_id " chwnd
		Sleep 100
		SendMessage, 0x031A,,,, % "ahk_id " chwnd
	;}

	; die Chargennummern werden anhand der Wochenkennung ausgewählt ;{
		For ctrlIndex, vctrl in CNRCtrls {
			cbox	:= GuiComboBoxPos("CI", vctrl, "^" weekNr ":")
			GuiControl, CI: Choose, % vctrl, % cbox.Pos
		}
	;}

	; Buttons 2.Impfdosis werden aktiviert oder inaktiviert ;{
		If (ImpfNr > 0) {

			ImpfIntervalle 	:= Impflinge.VaccinationIntervals(PatID, guiDate)
			ImpfWieder		:= Impflinge.NextVaccinationDate(PatID)

			For btnIndex, btnName in disButtons {

				iv := LTrim(ImpfIntervalle[btnIndex], "#") ; # Array speichert weder 0 noch negative Zahlen
				;SciTEOutput("iv: " iv ", d: " ImpfWieder[btnIndex])
				If (iv = 0)   {                                 	; kann geimpft werden
					GuiControl, CI: Text   	, % (btnName . "2"), % "2.Impfdosis"
					GuiControl, CI: Enable	, % (btnName . "2")
				}
				else if (iv < 0) {                             	; sollte noch gewartet werden
					GuiControl, CI: Text   	, % (btnName . "2"), % "ab " ConvertDBASEDate(ImpfWieder[btnIndex])
					GuiControl, CI: Disable	, % (btnName . "2")
				}
				else if (iv > 0) {                             	; Impfzeitraum überschritten
					GuiControl, CI: Text   	, % (btnName . "2"), % "2.Impfdosis +" iv
					GuiControl, CI: Enable	, % (btnName . "2")
				}

			}

		}
		else if (ImpfNr = 0) {

			For btnIndex, btnName in disButtons
				GuiControl, CI: Disable	, % (btnName . "2")

		}
	;}

return ;}

CIGuiClose:                          		;{
CIGuiEscape:
	Gdip_Shutdown(pToken)
	Gui, CI: Destroy
	Gui, CID: Destroy
	ExitApp
return ;}

}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Automatisierung
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
AlbisCImpfDokumentation(DATUM, vaccination) 	{

	; Impfung gegen COVID-19 #Impfstoffbezeichnung# Ch.: #Chargennummer# {U11.9G};

	; Interaktion über Maus und Tastatur blockieren
		;~ BlockInput, On

	; Aktivieren und blockierende Fenster entfernen
		AlbisActivate(2)
		AlbisIsBlocked()

	; Vorbereiten der Karteikarten
		DatumZuvor:= AlbisSetzeProgrammDatum(Datum)
		AlbisKarteikarteAktivieren()
		AlbisKarteikarteFocus(KK.KUERZEL)

	; handle des Kuerzel Steuerelementes merken
		ControlGet, hKuerzel, HWND,, % KK.KUERZEL, ahk_class OptoAppClass

	; Impfdiagnose erstellen
		Diagnose := StrReplace(vaccination.Diagnose, "#Impfstoffbezeichnung#", vaccination.Handelsname)
		Diagnose := StrReplace(Diagnose, "#Chargennummer#", vaccination.CNR)

	; Impfdiagnose senden
		ControlFocus, % KK.Kuerzel, ahk_class OptoAppClass
		VerifiedSetText(KK.Kuerzel, "dia", "ahk_class OptoAppClass")
		ControlFocus, % KK.Inhalt, ahk_class OptoAppClass
		VerifiedSetText(KK.Inhalt, Diagnose, "ahk_class OptoAppClass")
		ControlFocus, % KK.Inhalt, ahk_class OptoAppClass
		Send, {Tab}{Tab}{Tab}
		Sleep 500


		ControlGetFocus, cfocus, ahk_class OptoAppClass
		If (cfocus != KK.KUERZEL) {
			SendInput, {Tab}
			Sleep 300
			AlbisKarteikarteFocus(KK.KUERZEL)
		}

	; Leistungen eintragen
		ControlFocus, % KK.Kuerzel, ahk_class OptoAppClass
		VerifiedSetText(KK.Kuerzel, "lko", "ahk_class OptoAppClass")
		ControlFocus, % KK.Inhalt, ahk_class OptoAppClass
		VerifiedSetText(KK.Inhalt, vaccination.SNR (vaccination.NR = 1 ? "A":"B"), "ahk_class OptoAppClass")
		ControlFocus, % KK.Inhalt, ahk_class OptoAppClass
		Send, {Tab}{Tab}{Tab}
		Sleep 300

		ControlGetFocus, cfocus, ahk_class OptoAppClass
		If (cfocus = KK.INHALT) {
			ControlFocus, % KK.Inhalt, ahk_class OptoAppClass
			Send, {Tab}{Tab}{Tab}
			Sleep 1000
			AlbisKarteikarteFocus(KK.KUERZEL)
		}

		SendInput, {Escape}
		Sleep 300
		SendInput, {Escape}
		Sleep 300

	; Karteikarte wieder aktivieren und Datum zurücksetzen
		AlbisKarteikarteAktivieren()
		AlbisSetzeProgrammDatum(DatumZuvor)

	; Interaktion über Maus und Tastatur wieder freigeben
		;~ BlockInput, Off


}

AbrScheinAngelegt(Tagesdatum ) 							{

	Quartal := QuartalTage({"TDatum":StrReplace(Tagesdatum, "."), "TFormat":"ddMMYYYY"})
	abrExist := AlbisAbrechnungsscheinVorhanden(Quartal.Quartal "/" SubStr(Quartal.Jahr, -1))
	If       	(AlbisVersArt() = 2 && !abrExist) {
		MsgBox, 0x1004, % RegExReplace(A_ScriptName, "\.*$"), % "Die KVK ist noch nicht eingelesen worden`n"
					.  "bzw. ein Abrechnungsschein wurde noch nicht angelegt!`n"
					.	"Wiederholen Sie den Vorgang wenn Sie soweit sind"
		return 0
	}
	else If 	(AlbisVersArt() = 1 && !abrExist) {
		If !AlbisAbrScheinCOVIDPrivat(Quartal) {
			MsgBox, 0x1004, % RegExReplace(A_ScriptName, "\.*$"), % "Das Anlegen des Abrechnungsscheines ist fehlgeschlagen.`n"
				.  "Bitte lege den Schein manuell an!`nWiederhole diesen Schritt nachdem fertig bist`n`n"
				.	"Möchtest Du die Anleitung für das Anlegen dieses speziellen Scheines sehen?"
			IfMsgBox, Yes
				Run, % A_ScriptDir "\CoronaImpfdoku_resources\Abrechnung COVID-19 PRivatversicherung.svg"
			return 0
		}

	}

return 1
}

PrintDocuments(PatID, PatName, PatGeburt, docs) {

	; ### sichern welche Impfdokumente schon gedruckt wurden
		static tags := ["$Nachname#", "$Vorname#" , "$gebdatum#"]

	; Dokumente aufrufen und drucken
		wdApp := new MSWord_ComObject()
		Loop % docs.Count() {
			If docs[A_Index] && FileExist(Addendum.AlbisBriefe "\" docs[A_Index]) {

				SciTEOutput(docs[A_Index])
				wdApp.OpenDocument(Addendum.AlbisBriefe "\" docs[A_Index])
				wdApp.FindAndReplace("$Nachname#" 	, StrSplit(PatName, ", ").1)
				wdApp.FindAndReplace("$Vorname#"    	, StrSplit(PatName, ", ").2)
				wdApp.FindAndReplace("$gebdatum#"  	, PatGeburt)
				wdApp.PrintOut("PRT-SP1")
				wdApp.CloseDocument()

			}
		}




}
;}

class MSWord_ComObject                                 	{


		__New(objerror:=1, wdVisible:=1) {

			ComObjError(objerror)
			this.objerror := objerror
			try
				this.wdApp := ComObjActive("Word.Application")
			catch
				this.wdApp := ComObjCreate("Word.Application")

			this.wdApp.Visible := wdVisible

		}

		Disconnect() {
			this.wdApp := ""
		}

		FindAndReplace(search, replace) {
			this.wdApp.Selection.Find.Execute( search, 0, 0, 0, 0, 0, 1, 1, 0, replace, 2)
		}

		OpenDocument(filePath) {
			this.ActiveDocument 	:= this.wdApp.Documents.Open(filePath)
			this.ActiveFullName 	:= this.wdApp.ActiveDocument.FullName
		}

		CloseDocument(SaveChanges:=false) {
			this.wdApp.Documents.Close(SaveChanges ? 1 : 0)
		}

		PrintOut(printer) {

			this.wdApp.ActivePrinter := printer

			PageSetup := this.wdApp.Selection.PageSetup
			PageSetup.Orientation 									:= wdOrientPortrait	    	:=	0
			PageSetup.FirstPageTray 								:= wdPrinterDefaultBin  	:=	0
			PageSetup.OtherPagesTray 							:= wdPrinterDefaultBin  	:=	0
			PageSetup.SectionStart       							:= wdSectionContinuous 	:=	0
			PageSetup.OddAndEvenPagesHeaderFooter 	:= False
			PageSetup.DifferentFirstPageHeaderFooter 	:= False
			PageSetup.VerticalAlignment 							:= wdAlignVerticalTop		:=	0
			PageSetup.SuppressEndnotes 						:= False
			PageSetup.MirrorMargins 								:= False
			PageSetup.TwoPagesOnOne 			     			:= False
			PageSetup.BookFoldPrinting 							:= False
			PageSetup.BookFoldRevPrinting 						:= False
			PageSetup.BookFoldPrintingSheets 				:= 1
			Pagesetup.GutterPos 										:= wdGutterPosLeft			:= 	0

			; expression.PrintOut(Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName
			;	    					, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight)

			Background			    	:= 0
			Append				     		:= 0
			Range							:= ""
			OutputFileName			:= ""
			From							:= 0
			To								:= 0
			Item					     		:= ""
			Copies				     		:= 1
			Pages							:= 0
			PageType				    	:= wdPrintAllPages := 0
			PrintToFile				    	:= 0
			Collate					    	:= 1
			FileName				    	:= "doku"
			ActivePrinterMacGX   	:= ""
			ManualDuplexPrint		:= 1
			PrintZoomColumn 		:= 0
			PrintZoomRow				:= 0
			PrintZoomPaperWidth	:= 0
			PrintZoomPaperHeight	:= 0

			this.wdApp.PrintOut(Background,,,,,,,Copies)

			/*
			;~ this.wdApp.PrintOut(Background
										;~ , Append
										;~ , Range
										;~ , OutputFileName
										;~ , From
										;~ , To
										;~ , Item
										;~ , Copies
										;~ , Pages
										;~ , PageType
										;~ , PrintToFile
										;~ , Collate
										;~ , FileName
										;~ , ActivePrinterMacGX
										;~ , ManualDuplexPrint
										;~ , PrintZoomColumn
										;~ , PrintZoomRow
										;~ , PrintZoomPaperWidth
										;~ , PrintZoomPaperHeight)

				 */

		}

}

class ImpfListe                                                   	{                                                             	;	erleichtert das Handling einer Impfliste für Corona Impfungen

		__New()                                                    	{			;	Anlegen einer Impfling-Liste

			this.Vaccines 	:= {"#88331":"1", "#88333":"2", "Biontech":"88331", "AstraZeneca":"88333"}  ; 1 = Biontech , 2 = AstraZeneca
			this.VacPeriods	:= {1:"6", 2:"8-12"}
			this.Impflinge 	:= this.CoronaImpfungen()

		}

		CoronaImpfungen()                                   	{        	;  Daten aller Geimpften einlesen


			aDB          	:= new AlbisDB(Addendum.AlbisDBPath)
			BefundIndex	:= aDB.IndexRead("BEFUND",, false)
			CImpfungen 	:= aDB.GetDBFData("BEFUND"
															, 	{"KUERZEL":"rx:.*(lko|dia).*", "INHALT":"(8833[1-4]|U11\.9G)"}
															, 	["PATNR", "DATUM", "KUERZEL", "INHALT", "removed"]
															, 	"recordnr=" BefundIndex["#20212"].1
															, 	0)

			;JSONData.Save(A_Temp "\CImpfungen.json", CImpfungen, true,, 1, "UTF-8")
			;~ Run % A_Temp "\CImpfungen.json"

			; nur die Patienten-ID's zusammentragen
				this.Impflinge := Object()
				EBMziffer := 0
				For idx, set in CImpfungen {

						If set.removed
							continue

						PatID := set.PATNR
						If !this.Impflinge.haskey(PatID)
							this.Impflinge[PatID] := {"Impftage":{}}
						If !this.Impflinge[PatID].haskey(set.DATUM)
							this.Impflinge[PatID].ImpfTage[set.DATUM] := {}

						If RegExMatch(set.INHALT, "i)(88331|88332|88333|88334)[A-Z]", EBM) {
							ImpfStatus := RegExMatch(set.INHALT, "(A|G|V)") ? 1 : 2
							this.Impflinge[PatID].Status := ImpfStatus
							this.Impflinge[PatID].ImpfTage[set.DATUM].gebuehr := EBM
						}
						If RegExMatch(set.INHALT, "i)\s+(?<Stoff>\w+)\s(?<NR>\w+)\s*\{", C) {
							this.Impflinge[PatID].ImpfTage[set.DATUM].CStoff 	:= CStoff
							this.Impflinge[PatID].ImpfTage[set.DATUM].CNR 	:= CNR
						}

				}



	return this.Impflinge
	}

		CoronaImpfungen2()                                   	{        	;  Daten aller Geimpften einlesen

			starttime    	:= A_TickCount
			aDB          	:= new AlbisDB(Addendum.AlbisDBPath)
			BefundIndex	:= aDB.IndexRead("BEFUND",, false)
			CImpfungen 	:= aDB.GetDBFData("BEFUND"
															, 	{"KUERZEL":"rx:.*(lko|dia).*", "INHALT":"(8833[1-4]|U11\.9G)"}
															, 	["PATNR", "DATUM", "KUERZEL", "INHALT", "RECID", "DATETIME", "removed"]
															, 	"recordnr=" BefundIndex["#20212"].1
															, 	0)

			;JSONData.Save(A_Temp "\CImpfungen.json", CImpfungen, true,, 1, "UTF-8")
			;Run % A_Temp "\CImpfungen.json"

			; nur die Patienten-ID's zusammentragen
				this.Impflinge := Object()
				EBMziffer := 0
				For idx, set in CImpfungen {

						If set.removed
							continue

						PatID := set.PATNR
						If !IsObject(this.Impflinge[PatID])
							this.Impflinge[PatID] := {"Impftage":{}}
						If !IsObject(this.Impflinge[PatID][set.DATUM])
							this.Impflinge[PatID].ImpfTage[set.DATUM] := {}

						If RegExMatch(set.KUERZEL, "i)lko") {

							ImpfStatus := RegExMatch(set.INHALT, "(A|G|V)") ? 1 : 2
							If RegExMatch(set.INHALT, "i)(88331|88332|88333|88334)[A-Z]", EBM) {
								this.Impflinge[PatID].Status := ImpfStatus
								this.Impflinge[PatID].ImpfTage[set.DATUM].gebuehr := EBM
							}
						}
						else if RegExMatch(set.KUERZEL, "i)dia") {

							If RegExMatch(set.INHALT, "i)\s+(?<Stoff>\w+)\s(?<NR>\w+)\s*\{", C) {
								this.Impflinge[PatID].ImpfTage[set.DATUM].CStoff 	:= CStoff
								this.Impflinge[PatID].ImpfTage[set.DATUM].CNR 	:= CNR
							}
						}
				}

			SciTEOutput("time: " Round((A_TickCount-starttime)/1000, 2) "s")

	return this.Impflinge
	}


	; Objektinformationen
		VaccinationsStatistic()                                  	{         	;	erstellt eine Gesamtstatistik aller Impfungen

			If IsObject(this.stats)
				return this.stats

			this.stats := {"countAll":0, "vaccines":[0,0,0,0], "days":0, "weeks":0, "months":0}
			For PatID, Impf in this.Impflinge {

				this.stats.countAll += (Impf.Status ? Impf.Status : 1)


			}


		return this.stats
		}

		VaccinationsCount() 		                        		{			; 	Anzahl der Impfungen bisher
			stats := this.VaccinationsStatistic()
		return stats.countAll
		}


	; Patienteninformationen
		PatExist(PatID)                                           	{   		;	PatID ist in der Impfliste

			If this.Impflinge.haskey(PatID)
				return true

		return false
		}

		Vaccinations(PatID)                                    	{			;	wieviele Impfungen hat der Patient erhalten

			If !this.PatExist(PatID)
				return 0

		return this.Impflinge[PatID].Status
		}

		VaccinationDates(PatID)                             	{			;	gibt die Tage zurück an denen geimpft wurde

			If !this.PatExist(PatID)
				return 0

			VacDays := []
			For VacDate, data in this.Impflinge[PatID].ImpfTage
				VacDays.Push(VacDate)

		return VacDays
		}

		Vaccine(PatID)                                           	{			;	gibt die/den Impfstoffnamen die geimpft wurden zurück

			If !this.PatExist(PatID)
				return 0

			Vacs := []
			For VacDate, data in this.Impflinge[PatID].ImpfTage {
				RegExMatch(data.gebuehr, "\d+", gebuehr)
				Vacs.Push(this.Vaccines["#" gebuehr])
			}

		return Vacs
		}


	; Berechnungen
		NextVaccinationDate(PatID)                         	{			;	Rückgabe des frühstmöglichen Datums der Wiederholungsimpfung nach Impfstoff

			FirstVacDate 	:= this.VaccinationDates(PatID).1
			nxtVDays   	:= []
			For vaccineNr, period in this.VacPeriods {
				RegExMatch(period, "(?<from>\d+)\-*(?<till>\d+)*", W)
				nextdate := RTrim(DateAdd(FirstVacDate, Wfrom*7, "days"), "0")
				nxtVDays.Push(nextdate)
			}

		return nxtVDays
		}

		VaccinationIntervals(PatID, DateNow)          	{			;	eingestelltes Datum wird mit Impfintervall verglichen

			; Rückgabeparameter: 	0 				= kann geimpft werden
			; 									kleiner 0 	= Anzahl der Wochen bis zur nächsten Impfung
			; 									größer 0 	= Anzahl der Wochen die der Impfzeitraum überschritten wurde

			FirstVacDate 	:= this.VaccinationDates(PatID).1
			If RegExMatch(DateNow, "\d{2}\.\d{2}\.\d{4}")
				DateNow 	:= ConvertToDBASEDate(DateNow)

			;SciTEOutput("fd: " FirstVacDate "`ntd: " DateNow )
			intervals := []
			For vaccineNr, period in this.VacPeriods {

				RegExMatch(period, "(?<from>\d+)\-*(?<till>\d+)*", W)
				days 	:= DaysBetween(FirstVacDate, DateNow)
				weeks 	:= Floor(days/7)
				;SciTEOutput("nr: " vaccineNr ", d: " days ", w: " weeks)

				If (weeks < Wfrom)
					intervals.Push("#" -1*(Wfrom-weeks))
				else if (weeks >= Wfrom) {

					If !WTill
						intervals.Push("#0")
					else if WTill && (weeks <= WTill)
						intervals.Push("#0")
					else if WTill && (weeks > WTill)
						intervals.Push("#" WTill - weeks)

				}

			}

		return intervals
		}

}

ShowImpflinge(Impflinge)                                   	{

	For PatID, impfdaten in Impflinge {
		tage := ""
		For DATUM, o in Impfdaten
			tage := ConvertDBASEDate(DATUM) " EBM: " o.EBM ", "
		SciTEOutput(" [" PatID "] " PatDB[PatID].NAme ", " PatDB[PatID].Vorname "`t Status: " Impfdaten.Status ". | " RTrim(Tage, ", "))
	}
	SciTEOutput(Impflinge.Count())

}


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Hilfsfunktionen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
Feierabend(ExitReason, ExitCode) 							{

	OnExit("Feierabend", 0)
	JSONData.Save(Addendum.propsPath, CIProps, true,, 1, "UTF-8")
	Gui, CI: Destroy
	Gui, CID: Destroy
	ExitApp

}

ValidateChargenNummer(chnr) 							{                                                  ; prüft die übergebene Chargennummer

	; String sollte aus Großbuchstaben und anschließend Zahlen bestehen
	; gibt die Ch.Nr. im Objekt zurück key (CNR)

	obj := Object()
	chnr := Format("{:U}", chnr)
	RegExMatch(chnr, "O)^(?<KW>\d+)*\s*:*\s*(?<NRa>[A-Z]+)(?<NRb>\d+)", rx)
	obj.NR  	:=  !rx.NRa || !rx.NRb ? "" : rx.NRa rx.NRb
	obj.KW 	:= rx.KW
	obj.NRa 	:= rx.NRa
	obj.NRb 	:= rx.NRb

return obj
}

dBaseProgress(index, maxIndex, len, matchcount)	{

	global	DBSearchW
	;prgPos := Floor((index*100)/maxIndex)
	prgPos := "Datensätze gefunden: " matchcount
	ToolTip, % prgPos
	SetTimer, dBaseProgressOff, -2000

return
dBaseProgressOff:
	ToolTip
return
}

CIGui_RePaint()                                                    	{

		global hCI

		WinGet, CList, ControlListHwnd, % "ahk_id " hCI
		For idx, hwnd in StrSplit(CList, "`n") {
			WinGetClass, classnn, % "ahk_id " hwnd
			If !RegExMatch(classnn, "(msctls_progress)") {
				WinSet, Top    	,, % "ahk_id " hwnd
			}
		}
		Sleep 200
		For idx, hwnd in StrSplit(CList, "`n") {
			WinGetClass, classnn, % "ahk_id " hwnd
			If RegExMatch(classnn, "ComboBox")
				SendMessage, 0x0007,,,, % "ahk_id " hwnd
		}


}

GuiPosChanging(lParam, wParam, msg, hwnd) 		{
	global hCI, hCID
	Static x_offset := A_PtrSize*2 ; Support for x64 AHK


		If (msg = 0x46) {
			;SetTimer, GuiShowCID, -500
		}
		If (msg = 0x47) {
			CIPos := GetWindowSpot(hCI)
			Gui, CID: Show, % "x" CIPos.X+3 " y" CIPos.Y+CIPos.H-3
		}

}

GuiComboBoxPos(guiname, vctrl, string)            	{

	global hCtrlActive, CITL, CIDate, CIBTN, CIBTE, CIBT1, CIBT2, CIAZN, CIAZE, CIAZ1, CIAZ2

	cbox := {}
	cbox.pos := 0

	Gui, % guiname ": Default"
	GuiControlGet, chwnd, % guiname ": Hwnd", % vctrl
	ControlGet, CBList, List,,, % "ahk_id " chwnd
	For cbPos, content in StrSplit(CBList, "`n") {
		cbox.list .= (cbPos > 1 ? "|" : "") content
		If RegExMatch(content, string)
			cbox.pos := cbPos
	}

return cbox
}

SetCaretPos(HWND, caretPos) 								{
   ; EM_SETSEL = 0x00B1 -> msdn.microsoft.com/en-us/library/bb761661(v=vs.85).aspx
   Return DllCall("User32.dll\SendMessage", "Ptr", HWND, "UInt", 0x00B1, "Ptr", caretPos , "Ptr", caretPos, "Ptr")
}

GetComboBoxEdit(hCombo) 									{

	;-- Define/Populate the COMBOBOXINFO structure
	cbSize:=(A_PtrSize=8) ? 64:52
	VarSetCapacity(COMBOBOXINFO,cbSize,0)
	NumPut(cbSize,COMBOBOXINFO,0,"UInt")                ;-- cbSize

	;-- Get ComboBox info
	DllCall("GetComboBoxInfo","Ptr",hCombo,"Ptr",&COMBOBOXINFO)

Return NumGet(COMBOBOXINFO,(A_PtrSize=8) ? 48:44,"Ptr")
}

GdipBgr(hgc,c1,c2=0,Txt="",TxtOpt="", Fnt="")	{

	global

	If !TxtOpt
		TxtOpt="x0p y15p s60p Center cff000000 r4 Bold"
	If !Fnt
		Fnt="Arial"

	ctrlp := GetWindowSpot(hgc)

	pBrushFront	:= Gdip_BrushCreateSolid(c1)
	;pBrushBack	:= Gdip_BrushCreateSolid(c2)
	pBitmap 		:= Gdip_CreateBitmap(ctrlp.W, ctrlp.H)
	G 				:= Gdip_GraphicsFromImage(pBitmap)

	Gdip_SetSmoothingMode(G, 4)
	Gdip_FillRectangle(G, pBrushFront, 0, 0, ctrlp.W, ctrlp.H)
	;Gdip_FillRoundedRectangle(G, pBrushFront, 4, 4, (Posw-8)*(Percentage/100), Posh-8, (Percentage >= 3) ? 3 : Percentage)
	If Txt
		Gdip_TextToGraphics(G, (Txt != "") ? Txt : Round(Percentage) "`%", TxtOpt, Fnt, ctrlp.W, ctrlp.H)

	hBitmap := Gdip_CreateHBITMAPFromBitmap(pBitmap)
	SetImage(hgc, hBitmap)

	Gdip_DeleteBrush(pBrushFront)
	;Gdip_DeleteBrush(pBrushBack)
	Gdip_DeleteGraphics(G), Gdip_DisposeImage(pBitmap), DeleteObject(hBitmap)

Return, 0
}

; AnimateWindow Wrapper Functions                 	;{

WinShowFade(winId, speed := 200){
    DllCall("AnimateWindow", "UPtr", winID, "Int", speed, "UInt", 0x00080000)
}
WinHideFade(winID, speed := 200){
    DllCall("AnimateWindow", "UPtr", winID, "Int", speed, "UInt", 0x00080000|0x00010000)
}
WinShowPop(winID, speed := 100){
    DllCall("AnimateWindow", "UPtr", winID, "Int", speed, "UInt", 0x00000010)
}
WinHideShrink(winID, speed := 100){
    DllCall("AnimateWindow", "UPtr", winID, "Int", speed, "UInt", 0x00010000|0x00000010)
}
WinShowSlide(winId, speed := 200, dir := "LR"){
    dirCode := dir = "RL" ? 0x00000002 : dir = "TB" ? 0x00000004 : dir = "BT" ? 0x00000008 : 0x00000001
    DllCall("AnimateWindow", "UPtr", winID, "Int", speed, "UInt", 0x00020000|0x00040000|dirCode)
}
WinHideSlide(winId, speed := 200, dir := "LR"){
    dirCode := dir = "RL" ? 0x00000002 : dir = "TB" ? 0x00000004 : dir = "BT" ? 0x00000008 : 0x00000001
    DllCall("AnimateWindow", "UPtr", winID, "Int", speed, "UInt", 0x00010000|0x00040000|dirCode)
}
WinShowRoll(winId, speed := 200, dir := "LR"){
    dirCode := dir = "RL" ? 0x00000002 : dir = "TB" ? 0x00000004 : dir = "BT" ? 0x00000008 : 0x00000001
    DllCall("AnimateWindow", "UPtr", winID, "Int", speed, "UInt", 0x00020000|dirCode)
}
WinHideRoll(winId, speed := 200, dir := "LR"){
    dirCode := dir = "RL" ? 0x00000002 : dir = "TB" ? 0x00000004 : dir = "BT" ? 0x00000008 : 0x00000001
    DllCall("AnimateWindow", "UPtr", winID, "Int", speed, "UInt", 0x00010000|dirCode)
}
;}

Create_Impfen_ico()                                          	{

VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACHh4fT09P////////////Pz88AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACHh4fT09P////////////Pz88AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADU1NT////////////Nzc0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACHh4fm5ub////////Nzc0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACOjo7j4+PLy8sAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB///////8AAD///////wAAn///////AACP//////8AAMP//////wAA4P//////AADwP/////8AAPgOH////wAA+AQP////AAD8AAf///8AAPwAA////wAA/AAB////AAD+AAD///8AAP8AAH///wAA/wAAP///AAD+AAAf//8AAPwAAA///wAA/AAAB///AAD8AAAD//8AAP4AAAH//wAA/wAAAP//AAD/gAAAf/8AAP/AAAA//wAA/+AAAB//AAD/8AAAD/8AAP/4AAAH/wAA//wAAAP/AAD//gAAAf8AAP//AAAA/wAA//+AAAD/AAD//8AAAP8AAP//4AAB/wAA///wAAP/AAD///gAB/8AAP///AAH/wAA///+AAP/AAD///8AAfAAAP///4BA4AAA////wOBAAAD////h8AEAAP/////4AwAA//////wHAAD//////A8AAP/////4HwAA//////A/AAD/////8H8AAP/////g/wAA//////H/AAA="
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return -1
VarSetCapacity(Dec, DecLen, 0)
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return -2
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
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\lib\Sift.ahk
;}





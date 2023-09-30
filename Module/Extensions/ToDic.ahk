; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                   TEXT TO DICTIONARY INTERFACE
;
;      Funktion:
;
;      	Basisskript:	-	keines
;		 Abhängigkeiten: 	-	siehe #includes
;
;	    	Addendum für Albis on Windows
;     	by Ixiko started in September 2017 - last change 17.08.2023 -
;				this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;

	; Skripteinstellungen                                        	;{

		#NoEnv
		#KeyHistory               	, 0
		#MaxThreads                	, 100
		#MaxThreadsBuffer         	, On
		#MaxMem                    	, 4096

		SetBatchLines              	, -1
		;ListLines                 	, Off
		SetControlDelay            	, -1
		SetWinDelay               	, -1
		AutoTrim                   	, On
		FileEncoding              	, UTF-8

	;}

	; Variablen                                                  	;{

		global q:=Chr(0x22)
		global adm := Object()
		global stats := Object()
		global FilesP, FilesB, FilesT, FilesO
		global Medical, Others, MedicalCountAtStart, OthersCountAtStart
		global ngrams
		global codetesting := 0

		MFreq	:= {}, maxMFreq	:= 0, minMFreq := 10000000
		SFreq	:= {}, maxSFreq	:= 0, minSFreq := 10000000

		RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)
		adm := AddendumBaseProperties(AddendumDir)
		adm.albis := GetAlbisPaths()


		Menu, Tray,  Icon, % adm.Dir "\assets\ModulIcons\Praxomat.ico"
		cJSON.EscapeUnicode := "UTF-8"


		adm.basePath        := adm.DBPath "\Dictionary"
		adm.FilesProcessed  := adm.basePath "\FilesP.json"
		adm.medicalWords    := adm.basePath "\Medical.json"
		adm.othersWords     := adm.basePath "\Others.json"
		adm.PathWordEndings := adm.basePath "\WordEndings.json"
		adm.FilesFromAlbis  := adm.basePath "\FilesAlbis.json"
		adm.OCRspath        := adm.basePath "\FilesAlbisUnprocessed.json"

	;}

	; Wörterbücher laden                                        	;{

		Medical	:= FileExist(adm.medicalWords)  	? cJSON.Load(FileOpen(adm.medicalWords  	, "r"	, "UTF-8").Read())	: Object()
		Others 	:= FileExist(adm.othersWords)    	? cJSON.Load(FileOpen(adm.othersWords   	, "r"	, "UTF-8").Read()) 	: Object()
		tempArr	:= FileExist(adm.FilesProcessed) 	? cJSON.Load(FileOpen(adm.FilesProcessed	, "r"	, "UTF-8").Read())	: Array()
		FilesP	:= RemoveDriveLetter(tempArr)
		tempArr := ""

		SciTEOutput(cjSON.Dump(adm, 1))

		MedicalCountAtStart	:= Medical.Count()
		OthersCountAtStart 	:= Others.Count()


	; Autobackup Medical Dictionary
		If !InStr(FileExist(adm.basePath "\backup"), "D")
			FileCreateDir, % adm.basePath "\backup"
		if !FileExist(backupFilePath := adm.basepath "\backup\" A_YYYY A_MM A_DD A_Hour "-" (Floor(A_Min/15)+1) "_Medical.json")
			FileOpen(backupFilePath , "w", "UTF-8").Write(cJSON.Dump(Medical, 1))

		;SciTEOutput("backup file: " backupFilePath)
		;~ endings := WordEndings(Medical)
		;~ ExitApp

		removedOthers := streetadds := 0
		SciTEOutput("Others Wörterbuch enthält: " Others.Count() " Worte")
		SciTEOutput("entferne Doubletten und ergänze Straßennamen ... ")

	; Others - doppelte Worte entfernen
	; Medical - str auf straße ändern
		For word, wcount in Medical {

			if (wcount = "")
				Medical[word] := wcount := 1

			If Others.haskey(word) {
				Medical[word] += Others[word]
				wcount := Medical[word]
				Others.Delete(word)
				removedOthers ++
			}

			If (word ~= "(\p{Ll}{2,})str(asse)*\s*$") {
				street := RegExReplace(word, "(\p{Ll}{2,})str(asse)*\s*$", "$1straße")       	; Hauptstr -> Hauptstraße
				Medical[street] := (Medical.haskey(street) ? Medical[street] : 0) + Medical[word]
				Medical.Delete(word)
				streetadds ++
				streetnames .= word ", "
			}

			maxMFreq := wcount > maxMFreq	? wcount : maxMFreq
			minMFreq := wcount > minMFreq	? wcount : minMFreq
			MFreq[wcount] := !MFreq.haskey(wcount) ? 1 : MFreq[wcount] += 1

		}

		SciTEOutput((removedOthers ? removedOthers " Worte aus d. Others WB entfernt. Wörterbuchgröße nach Bereinigung: " Others.Count()  " Worte)" : "keine zu bereinigenden Einträge gefunden"))
		SciTEOutput((streetadds    ? streetadds    " Straßennamen ergänzt ... " streetnames  : ""))
		justreturn := true
		gosub SaveData1
		justreturn := false

		;~ freqs := []


		; ___________________________________________

		For word, wcount in Others {
			maxSFreq	:= wcount > maxSFreq	? wcount : maxSFreq
			minSFreq 	:= wcount > minSFreq 	? wcount : minSFreq
			SFreq[wcount] := !SFreq.haskey(wcount) ? 1 : SFreq[wcount] += 1
		}
		;~ SciTEOutput(PrintFrequenciesCount(MFreq))
		;~ SciTEOutput(PrintFrequenciesCount(SFreq))

		;~ Sci
		;~ ExitApp

	;}

	;ngrams := RunCrazyNGrams("3,4,5", true)
	;clean := RemoveDoublettes(adm.medicalWords, adm.othersWords)
	;adm.medicalWords 	:= clean.objA
	;adm.othersWords  	:= clean.objB
	;~ gosub AlbisPDFOdner
	;ReIndexMedicalList()

	FilesT := BefundOrdner(FilesP)

	DicGui()
	CheckSyntaxGui()

return

AutoWortChecker() {

	; Medical ist global

	wordEndings := {"de" : ["ung(?<end>en)"			  	, "([knrt])(?<end>en)"		, "[knrt](?<end>er)"		, "[kmnrt](?<end>es)"
												, "([dgknrt])(?<end>e)"		, "(le)(?<end>n)"	  			, "[gn](?<end>s)"	    	, "ium"									, "ie"
												, "isch"							  	, "ion"]}
	wordBindings := {"de" : ["[tg](s)", "[l](en)", "lf(e)", "f(d)", "f(p)", "f(sy)", "ions", "s(th|[bdep])"]}

	refObj := Object()

	For word, wcount in Medical {

		;~ Sub


	}


}

AlbisPDFOdner:	; für nachträgliches OCR -- Albis PDF Ordner , PDF -> Text	;{

		noadm.OCRspath := adm.DBPath "\Dictionary\albisDir_NoOCR_PDF.json"
		adm.OCRspath   := adm.DBPath "\Dictionary\albisDir_OCR_PDF.json"
		FilesT     := FileExist(adm.OCRspath) ? cJSON.Load(FileOpen(adm.OCRspath, "r"	, "UTF-8").Read()) : Array()
		FilesNT	   := Array()

;~ return


		; SciTEOutput(cJSON.dump(FilesT, 1))

		;~ FilesT		:= Array()
		startT       	:= A_TickCount
		Loop, Files, % adm.albis.Briefe "\*.pdf", R
		{

		  ; Verlauf
			If Mod(A_Index, 100) = 0 {
				t 	:= Floor((A_TickCount - startT) / 1000)
				m	:= Floor(t / 60)
				s	:= t - m*60
				ToolTip, % A_Index " PDF-Dateien`nZeit: (" t ") " m ":" SubStr("0" s, -1), 1400, 1, 13
			}

			fpath := RegExReplace(A_LoopFileFullPath, "^\w+?(:|\\)")
			If inArr(fpath, FilesT) || inArr(fpath, FilesNT)
				continue

			If !PDFisSearchable(A_LoopFileFullPath)
				FilesNT.Push(fpath)
			else
				FilesT.Push(fpath)

		}

		ToolTip,,,,13

		; Speichern
		FileOpen(adm.DBPath "\sonstiges\albisDir_NoOCR_PDF.json", "w", "UTF-8").Write(cJSON.Dump(FilesNT))
		FileOpen(adm.DBPath "\sonstiges\albisDir_OCR_PDF.json"  , "w", "UTF-8").Write(cJSON.Dump(FilesT))

;}

ObjectSaveBackup(object, filepath, backupPath:="\backup") {

	static nobackupDir

	if nobackupDir
		return

	SplitPath, filepath, fdir, fname, fExt, fOutExt

	if !isObject(object) {
		name := %object%
		SciTEOutput("Es wurde kein Backup der Daten von " fOutExt " erstellt, da es kein Daten-Objekt ist.")
		return
	}

	backupPath := backupPath ~= "i)^[A-Z]+(:|\\)" ? backupPath : fdir "\" LTrim(backupPath, "\")
	if !Instr(FileExist(backupPath), "D") {
		SciTEOutput("Das Verzeichnis für den Backup existiert nicht:`n - " backupPath)
		nobackupDir := true
		return
	}

	if !FileExist(backupFilePath := backupPath "\" fOutExt "_" A_YYYY A_MM A_DD A_Hour A_min "." fExt)
			FileOpen(backupFilePath , "w", "UTF-8").Write(cJSON.Dump(object, 1))

}
PrintFrequenciesCount(Frequencies, maxHeight:=50) {

	static pointsA := "▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉▉"
	                   ;1                                                                         ;50
	static pointsB := "▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏▏"

	For freq, wordCount in Frequencies
		maxWordCount := maxWordCount >= wordCount ? maxWordCount : wordCount
	SciTEOutput(MaxWordCount)

	For freq, wordCount in Frequencies {

		pMaxWord    := (wordCount*100/maxWordCount)*100
		perc        := Floor((pMaxWord*100)/maxWordCount)
		drawLen     := Floor(500*(perc/100))+1
		drawSmall   := drawLen - Floor(drawLen/7)*7
		drawWide    := Floor(drawLen/7)

		t .= Format("{:7}", Round(freq)) " " SubStr(pointsA, 1, drawWide) Substr(pointsB, 1, drawSmall) " " wordCount  "`n"

	}

return t
}
BefundOrdner(FilesP) {	; nicht bearbeitete Textdateien zusammenstellen	;{

	FilesT := Array()

	if !adm.ScanPath
		return

	SciTEOutput("bisher verarbeitete Dokumente: " FilesP.Count() "`nstelle Dokumente im Posteingang zusammen (FilesT)" )
	Loop, Files, % adm.ScanPath "\*.txt", R
	{

		If (Mod(A_Index, 200) < 2) {
			MouseGetPos, mx, my
			ToolTip, % FilesT.Count() " unverarbeitete Dateien von " A_Index " gesamt.`naktuelle Datei: " A_LoopFileFullPath, % mx-20 , % my+50, 5
		}

		fpath := RegExReplace(A_LoopFileFullPath, "^\w+?(:|\\)")
		If inArr(fpath, FilesP)
			continue

		FilesT.Push(fpath)

	}

 ; ToolTip Nr. 5 in 10s ausschalten
	TTOff(5, 10000)
	SciTEOutput("FilesT: " FilesT.Count() " unverarbeitete Posteingangsbefunde" )

return FilesT
}
;}

TTOff(TTNr, Delay) {
	fn := Func("TTNrOff").Bind(TTNr)
	SetTimer, % fn, % -1*Delay
}
TTNrOff(TTNr) {
	ToolTip,,,, % TTNr
}

SaveData1:
SaveData2: ;{

	;RunCrazyNGrams("3,4", true)
	cJSON.EscapeUnicode := "UTF-8"

	MedicalNew	:= Medical.Count() - MedicalCountAtStart
	OthersNew 	:= Others.Count()  - OthersCountAtStart

	If IsObject(Medical) 	&& (MedicalNew || justreturn)
		FileOpen(adm.medicalWords, "w", "UTF-8").Write(cJSON.Dump(Medical, 1))
	If IsObject(Others) 	&& (OthersNew || justreturn) {
		if FileExist(adm.othersWords)
			ObjectSaveBackup(Others, adm.othersWords)
		FileOpen(adm.othersWords, "w", "UTF-8").Write(cJSON.Dump(Others, 1))
	}
	If IsObject(FilesP)   	&& FilesP.Count() {
		if FileExist(adm.FilesProcessed)
			ObjectSaveBackup(FilesP, adm.FilesProcessed)
		FileOpen(adm.FilesProcessed	, "w", "UTF-8").Write(cJSON.Dump(FilesP, 1))
	}
	;~ If IsObject(FilesB)   	&& FilesB.Count()
		;~ FileOpen(adm.FilesFromAlbis	, "w", "UTF-8").Write(cJSON.Dump(FilesB, 1))
	;~ If IsObject(FilesT)   	&& FilesT.Count()
		;~ FileOpen(adm.OCRspath      	, "w", "UTF-8").Write(cJSON.Dump(FilesT, 1))

	If justreturn
		return

   ; einwenig Statistik
	statstxt := FNR+FNR_Last " Dateien ausgewertet.`n  - Medical-Dict: +" MedicalNew " Worte.`n  - Others: +" OthersNew  " Worte"
	LastError := A_LastError, EL := ErrorLevel

	FileAppend, % A_YYYY "-" A_MM "-" A_DD ", " A_Hour ":" A_Min " "
						. RegExReplace(StrReplace(statstxt, "`n"), "\s+\-\s") "`n"  , % adm.DBPath "\Dictionary\ToDic-Statistik.txt", UTF-8
	SciTEOutput(statstxt "`n  - EL: " EL ", A_LastError: " A_LastError "`n  - adm.DBPath: " adm.DBPath)


ExitApp ;}

CheckSyntaxGui() {

	static ngOut, ngInput, lastOK

	WinGetPos, x, y, w, h, % "Wörterbucheditor ahk_class AutoHotkeyGUI"
	Gui, ngram: new, -DPIScale
	Gui, ngram: Font, s11 q5 cBlack, Arial
	Gui, ngram: Add, Edit, xm ym   	w500 h600 	vngOut
	Gui, ngram: Add, Edit, xm y+10 	w500 r1    	vngInput gngInputValidate -AltSubmit
	Gui, ngram: Show, % "x" x+w+5 " y" y " AutoSize NA", ngram-Syntaxchecker

	Hotkey, IfWinActive, ngram-Syntaxchecker ahk_class AutoHotkeyGUI
	Hotkey, ~Enter, ngramSyntaxValidator
	Hotkey, IfWinActive


return
ngramSyntaxValidator: ;{

	Gui, ngram: Submit, NoHide

	ngLen := StrLen(ngInput)
	out   := ngInput "`n"

	Loop 3 {

		NgramSize := 2+A_Index
		Loop, % ngLen - (NgramSize - 1)  {

			;If ngrams[NgramSize].haskey(sub := SubStr(ngInput, A_Index, NgramSize))
			sub := SubStr(ngInput, A_Index, NgramSize)
			If ngrams[NgramSize].haskey(sub)
				out .= sub "`t+ [" ngrams[NgramSize][sub].count "]`n"
			else {
				out .= sub "`t-`n"
				Gui, ngram: Color, cFF4444
			}
		}
		out .= "`n"
	}

	GuiControl, ngram:, ngOut, % RTrim(out, "`n")


return ;}

ngInputValidate: ;{ korrigiert während der Eingabe

	Gui, ngram: Submit, NoHide
	If RegExMatch(ngInput, "i)[^\pL]") {
		If !lastOK
			return
		lastOK := false
		Gui, ngram: Color, cFF4444
		Sleep 150
		GuiControl, ngram:, ngInput, % RegExReplace(ngInput, "i)[^\pL]")
	} else {
		Gui, ngram: Color, cFFFFFF
		lastOK := true
	}

return ;}

ngramGuiClose:
ngramGuiEscape:
MsgBox, 0x1004, % A_ScriptName, % "Das Schließen dieses Fensters beendet das Skript.`n Sicher?"
IfMsgBox, No
	return
ExitApp
}

DicEditorGui() {

	global

	Gui, EDT: new, -DPIScale
	Gui, EDT: Font, % "s" (A_ScreenWidth>1920 ? 12 : 11) " q5 cBlack", Consolas
	Gui, EDT: Add, ListView	, % "xm 	y18 	w500 r40 AltSubmit vEDT_LV gEDT_LV"	, Wort|GZ
	; ----

return
EDT_LV:
return
}

DicGui() {

	global
	static 	FNR := 1, DicInit := true

	col1W1	:= 380, col1W2 := 80
	col2W1	:= 100, col2W2 := 370, col2W3 := 60
	Lv1W  	:= col1W1 + col1W2 + 40
	Lv2W  	:= col2W1 + col2W2 + col2W3 + 20

	Gui, Dic: new, -DPIScale
	Gui, Dic: Font, % "s" (A_ScreenWidth>1920 ? 11 : 10) " q5 cBlack", Consolas
	Gui, Dic: Add, ListView	, % "xm 	y18 	w" Lv1W " r40 AltSubmit vDIC_LVTXT1 	gDIC_LV"	, Wort|GZ
	Gui, Dic: Add, ListView	, % "x+0 	    	w" Lv1W " r40 AltSubmit vDIC_LVTXT2 	gDIC_LV"	, Wort|GZ
	Gui, Dic: Add, ListView	, % "x+0 	    	w" Lv2W " r40 AltSubmit vDIC_LVTXT3 	gDIC_LV"	, WBuch|Textwort|Z
	Gui, Dic: Add, ListView	, % "x+5 	  	 	w" Lv2W " r40 AltSubmit vDIC_LVDIC  	gDIC_LV"	, WBuch|Med.Wörterbuch|Z

	Gui, Dic: Font, s8 q5 cBlack, Arial
	GuiControlGet, cp2,	 Dic: Pos, DIC_LVTXT2
	GuiControlGet, cp3,	 Dic: Pos, DIC_LVTXT3
	Gui, Dic: Add, Edit    	, % "x+5 		 	w500 h" cp2H    "  vDIC_TXT   "                           	, % ""

	Gui, Dic: Font, s9 q5 cBlack, Arial
	Gui, Dic: Add, Text		, % "xm+20 y0 "	     		                        		    	            		, % "Medizin"
	Gui, Dic: Add, Text		, % "x" cp2X+20 " y0 "	  		                       		    	         		, % "andere"
	Gui, Dic: Add, Text		, % "x" cp3X+20 " y0 "	  		                       		    	         		, % "unbekannt/ngram"

	GuiControlGet, cp, DIC: pos, DIC_LVDIC
	Gui, Dic: Font, s8 q5 cBlack, Arial
	Gui, Dic: Add, Text		, % "x" cpX-100 "	y0"				                        		  			  				, % "Datei:"
	Gui, Dic: Font, s8 q5 cBlack, Arial
	Gui, Dic: Add, Text		, % "x+2 y0 w570             	vDIC_FN"		                          			, % adm.DriveLetter . FilesT[FNR]

	GuiControlGet, cp, DIC: pos, DIC_TXT
	Gui, Dic: Font, s8 q5 cBlack, Arial
	Gui, Dic: Add, Text		, % "x" cpX " y0 w400   	Center  	vDIC_WD"     	             			, % "[unverarbeitet: " " " FilesT.Count() " (Nr: " SubStr("         " FNR, -5) ")] "
	Gui, Dic: Add, Text		, % "x+10         	Center             	vDIC_WB"			          				, % "[00000 Worte] "
	;~ Gui, Dic: Add, Text		, % "x" cpX " y" cpY+cpH+2 " Center vDIC_WB"			       				, % "[00000 Worte] "

	; Statistik
	Gui, Dic: Font, s10 q5 cBlack, Arial
	Gui, Dic: Add, Text		, % "x" cpX " y" cpY+cpH+10 " w300 Center"		    			              			, % "Gesamtstatistik"
	Gui, Dic: Font, s8 q5 cBlack, Arial
	Gui, Dic: Add, Edit		, % "Y+5   w300  h120 vDIC_GS ReadOnly"					                    				, % "-------"

	GuiControlGet, cp, DIC: pos, DIC_WD
	Y := cp2Y + cp2H + 20
	Gui, Dic: Font, s10 q5 cBlack, Arial
	Gui, Dic: Add, Button 	, % "xm     y" Y      "	vDIC_BBK   	gDIC_BTNS"	                         	, % "<<"
	Gui, Dic: Add, Button 	, % "x+10              	vDIC_BFW   	gDIC_BTNS"	                         	, % ">>"
	Gui, Dic: Add, Button 	, % "x+30              	vDIC_SPDF  	gDIC_BTNS"	                         	, % "PDF-Datei ansehen"

	GuiControlGet, cp, DIC: pos, DIC_WD
	Gui, Dic: Add, Button 	, % "xm     y+10       	vDIC_BOK  	gDIC_BTNS hwndDIC_hOK"            		, Übernehmen
	Gui, Dic: Add, Button 	, % "x+10 	        		vDIC_BCL  	gDIC_BTNS"	                         	, Abbrechen + Sichern
	Gui, Dic: Add, Button 	, % "x+100 	        		vDIC_BXT  	gDIC_BTNS"	                        	, Beenden

	GuiControlGet, cp, DIC: pos, DIC_GS
	Gui, Dic: Add, Button 	, % "xm     y+10       	vDIC_EDT  	gDIC_BTNS"                         		, Wörterbücher`nbearbeiten

	Gui, Dic: ListView, % "DIC_LVTXT1"
	LV_ModifyCol(1, col1W1)
	LV_ModifyCol(2, col1W2 " Integer")

	Gui, Dic: ListView, % "DIC_LVTXT2"
	LV_ModifyCol(1, col1W1)
	LV_ModifyCol(2, col1W2 " Integer")

	Gui, Dic: ListView, % "DIC_LVTXT3"
	LV_ModifyCol(1, col2W1)
	LV_ModifyCol(2, col2W2)
	LV_ModifyCol(3, col2W3 " Integer")

	Gui, Dic: ListView, DIC_LVDIC
	LV_ModifyCol(1, col2W1)
	LV_ModifyCol(2, col2W2)
	LV_ModifyCol(3, col2W3 " Integer")

	Gui, Dic: Show, AutoSize, Wörterbucheditor [Substantive]


	If DicInit {
		res := 0
		while !res
			res := LoadText()
		WBStats()
	}

	Hotkey, IfWinActive, Wörterbucheditor ahk_class AutoHotkeyGUI
	Hotkey, Enter 	, addToDic
	Hotkey, Space	, ClickWord
	Hotkey, IfWinActive

return

ClickWord:

return

DIC_LV: ;{


	Critical

	;SciTEOutput(A_GuiControl "   " A_GuiControlEvent "  " A_EventInfo)
	; Daten der angeklickten Zeile auslesen
		row   	:= A_EventInfo
		LvCtrl	:= A_GuiControl
		rword 	:= []
		Gui, Dic: Default
		Gui, Dic: ListView, % LvCtrl
 		Loop % (LvCtrl = "DIC_LVTXT3" ? 3 : 2) {
			LV_GetText(rtxt, row, A_Index)
			rword[A_Index + (LvCtrl = "DIC_LVTXT3" ? 0 : 1)] := rtxt
		}

	; ins gegenüberliegende Listview Steuerelement verschieben
		If (A_GuiControlEvent = "Normal") {

			If (LvCtrl = "DIC_LVTXT3") {

				Gui, Dic: ListView, DIC_LVTXT3
				LV_Delete(row)

				Gui, Dic: ListView, DIC_LVDIC
				LV_Add("", rword.1, rword.2, rword.3)

			} else If (LvCtrl = "DIC_LVTXT2") {

				Gui, Dic: ListView, DIC_LVTXT2
				LV_Delete(row)

				Gui, Dic: ListView, DIC_LVDIC
				LV_Add("", "[andere]", rword.2, rword.3)

			} else if (LvCtrl = "DIC_LVDIC") {

				Gui, Dic: ListView, DIC_LVDIC
				LV_Delete(row)

				Gui, Dic: ListView, %  "DIC_LVTXT" (InStr(rword.1, "Medizin") ? "1" : InStr(rword.1, "andere") ? "2" : "3")
				If RegExMatch(rword.1, "i)(Medizin|andere)")
					LV_Add("", rword.2, rword.3)
				else
					LV_Add("", rword.1, rword.2, rword.3)

			}

			LV_ModifyCol(1, "SortDesc", "Sort")

			Gui, Dic: ListView, DIC_LVTXT2
			GuiControl, Dic:, DIC_WD, % "[" FNR "/" FilesT.MaxIndex() "] [" LV_GetCount() " Worte] "

			Gui, Dic: ListView, DIC_LVDIC
			GuiControl, Dic:, DIC_WB, % "[" LV_GetCount() " Worte]"

		}

	; aus Medical nach Others verschieben
		else If (A_GuiControlEvent = "R") || (A_GuiControlEvent = "RightClick") {

			If RegExMatch(rword.1, "Medizin\s+\[(?<count>\d+)\]", w) {

				MsgBox, 4, % RegExReplace(A_ScriptName, "\.[a-z]+$"), % "< " rword.2 " >`naus dem Medizinwörterbuch entfernen?"
				IfMsgBox, No
					return

				; ins andere Wörterbuch verschieben
					Medical.Delete(rword.2)
					Others[rword.2] :=  !Others.haskey(rword.2) ? wcount : Others[rword.2] += wcount

				; Listviewanzeige ändern
					Gui, Dic: ListView, % LvCtrl
					If (LvCtrl = "DIC_LVTXT2")
						 LV_Modify(row,,  "andere  [" wcount "]")
					else
						 LV_Add("",  "andere  [" wcount "]", rword.2, rword.3)

			}


		}

return ;}

addToDic: ;{

	ControlClick,, % "ahk_id " DIC_hOK

return ;}

DIC_BTNS: ;{

	If (A_GuiControl = "DIC_BOK")  {
		addToDictionary()
		WBStats()
		res := 0
		while !res
			res := LoadText()

	} else if (A_GuiControl = "DIC_BXT") {
		Gui, Dic: Destroy
		gosub SaveData2
		ExitApp
	} else if (A_GuiControl = "DIC_SPDF") {
		If FileExist(popfile)
			Run % popfile
		else
			MsgBox, % "Datei: " popfile " wurde nicht gefunden."
	} else if (A_GuiControl = "DIC_EDT") {
			DicEditorGui()
	}



return ;}

DicGuiClose:
DicGuiEscape:
	FilesT.Push(popfile)
	gosub SaveData2
return
}

addToDictionary() {


	global Dic, DIC_LVDIC, DIC_LVTXT, DIC_WB, FNR, popfile
	;global medicalWords, othersWords, FilesProcessed
	global Medical, Others, FilesP, FilesB, FilesT, AlbisBriefe

	; Medizinwörterbuch ergänzen
		Gui, Dic: ListView, DIC_LVDIC
		Loop, % LV_GetCount() {

			LV_GetText(rword  	, A_Index, 2)
			LV_GetText(rwcount 	, A_Index, 3)

			Medical[rword] := !Medical.haskey(rword) ? rwcount : Medical[rword] += rwcount
			If Others.haskey[rword] {
				Medical[rword] += Others[rword]
				Others.Delete(rword)
			}

		}
		LV_Delete()

	; Others Wörterbuch - Zähler erhöhen
		Gui, Dic: ListView, DIC_LVTXT2
		Loop % LV_GetCount() {
			LV_GetText(rword  	, A_Index, 2)
			Others[rword] := !Others.haskey(rword) ? 1 : Others[rword] += 1
		}
		LV_Delete()

	; Others Wörterbuch ergänzen mit den verbliebenen Worten in unbekannt/ngram
		tCol := 1
		Gui, Dic: ListView, DIC_LVTXT3
		Loop % LV_GetCount() {
			LV_GetText(rword  	, A_Index, 2)
			LV_GetText(rwcount  	, A_Index, 3)
			Others[rword] := !Others.haskey(rword) ? 1 : Others[rword] += rwcount
			 t .= rword " [" Others[rword] "]" (tCol=3 ? "`n" : " | ")
			 tCol += tCol = 3 ? -2 : 1
		}
		LV_Delete()

	If dbg
		SciTEOutput("unbekannt->Others:`n"t)

	; alle 3 Listviews leeren
		Loop 3 {
			Gui, Dic: ListView, % "DIC_LVTXT" A_Index
			LV_Delete()
		}
		Gui, Dic: ListView, % "DIC_LVDIC"
		LV_Delete()

	GuiControl, Dic:, DIC_WB, % "[           ]"

  ; bearbeitete Datei sichern
	If !AlbisBriefe
		FilesP.Push(popfile)
	else
		FilesB.Push(popfile)


}

LoadText() {

	global Dic, DIC_WD, DIC_FN, DIC_LVTXT, FNR, FNR_Last, FilesT, FilesB, FilesP, popfile, AlbisBriefe
	global thisText, wsCount, dbg := true

	; Ende wenn keine Dateien mehr zu bearbeiten sind, weitere Dateien aus dem Albis Stammverzeichnis laden
		If (FilesT.Count() = 0) {

			albisverzeichnis := InStr(FileExist(adm.Albis.Briefe), "D") ? true : false
			If AlbisBriefe {
				MsgBox, 0x1, Alle aktuellen Befunde sind ausgewertet, % "Keine Dateien mehr zum Bearbeiten vorhanden.`n"
						                                          		;~ .  ( albisverzeichnis ? "Weitermachen mit den Befunden in albiswin\Briefe?" : "")
				return 2
			}

			AlbisBriefe := true
			FNR_Last	:= FNR
			FileIndex	:= FNR := FileFoundIndex := 0
			tempArr  	:= FileExist(adm.FilesFromAlbis) ? cJSON.Load(FileOpen(adm.FilesFromAlbis, "r", "UTF-8").Read()) : Array()
			FilesB  	:= RemoveDriveLetter(tempArr)
			FilesT		:= Array()

			SciTEOutput("FilesB (bereits erfasste Dokumente im Verzeichnis 'albiswin\Briefe'): " FilesB.Count())

			If !codetesting {

				Loop, Files, % adm.albis.Briefe "\*.*", R
				{

					If A_LoopFileExt not in pdf,doc,txt
						continue

					If (Mod(A_Index, 1000) < 3) {
						RegExMatch(A_LoopFileFullPath, "\\Briefe\\\K(?<path>.*)\\(?<fn>[\pL\d\s]+\.(doc|pdf|txt))", rx_)
						MouseGetPos, mx, my
						ToolTip, % FilesT.Count() " unverarbeitete Dateien`n" FileFoundIndex " bearbeitete Dateien`n"
									. "aktuelles Verzeichnis: " rx_path "\" rx_fn, % mx + -20 , % my + 50, 5
					}

					fpath := RegExReplace(A_LoopFileFullPath, "^\w+?(:|\\)")
					If inArr(fpath, FilesP) {
						FileFoundIndex ++
						continue
					}

					FilesT.Push(fpath)

				}

				SciTEOutput(FilesT.Count() " Dateien im Verzeichnis 'albiswin\Briefe' erfasst")
				ToolTip,,,, 5

			}
			else
				FilesT := FilesB

		}

	; letzte Datei wird geladen und unter Processed gespeichert
		popfile := adm.DriveLetter . FilesT.pop()

	; Statistiken + aktuelle Datei
		GuiControl, Dic:, DIC_WD	, % "[unverarbeitet: " FilesT.Count() " (Nr: " SubStr("         " FNR, -5) ")] "
		If !popfile || !FileExist(popfile) {
			dbg ? SciTEOutput("Datei nicht vorhanden: " (popfile ? popfile : "--empty--")) : ""
			return 0
		}

		FNR++
		GuiControl, Dic:, DIC_FN, % popfile


	; Worte werden untersucht
		words := CollectWords(popfile,, errormsg)
		If (!IsObject(words) || !words.Count() || errormsg) {
			SciTEOutput(Format("{:6}", FNR) (errormsg ? errormsg "`n  " popfile : " - collected words object ist leer"))
			return 0
		}
		if thisText
			GuiControl, Dic:, DIC_TXT , % thisText
		GuiControl, Dic:, DIC_WB	, % "[" wsCount " Worte] "
		thisText := ""

	; Worte anzeigen
		For word, wdata in words {

			If (word ~= "^[\d\-\(\)\/\\]+$")
				continue

			if (wdata.list=2)
				continue

			col1 := ( wdata.list=1 ? "unbekannt"
				  		: wdata.list=2 ? "Medizin"
				  		: wdata.list=3 ? "andere"
					      						 : "ngram")  ; 4 = ngram

			Gui, Dic: ListView, % "DIC_LVTxt" ( wdata.list=1 || wdata.list=4 ? "3" :  wdata.list=2 ? "1" : "2")
			If (wdata.list=1 || wdata.list=4)
				LV_Add("", col1, word, wdata.count()) ; LV_Add("", col1, word, wdata.count) <--??
			else {
				RegExReplace(word,"\p{Lu}",, LenUpperChars)
				If (LenUpperChars < StrLen(word)/3)                 ; nicht anzeigen wenn Großbuchstaben mehr als 1/3 aller Buchstaben ausmachen
					LV_Add("",  word, wdata.listcount) ;, wdata.listcount)
			}
		}

	; alle 3 Listviews sind leer, weiter mit der nächsten Datei
		twordCount := 0
		Loop 3 {
			Gui, Dic: ListView, % "DIC_LVTxt" A_Index
			twordCount += LV_GetCount()
		}
		If !twordCount {
			;~ LoadText()
			return 0
		}


		Gui, Dic: ListView, % "DIC_LVTxt1"
		LV_ModifyCol(2, "SortDesc")
		Gui, Dic: ListView, % "DIC_LVTxt2"
		LV_ModifyCol(1, "SortDesc")
		Gui, Dic: ListView, % "DIC_LVTxt3"
		LV_ModifyCol(1, "SortDesc")

		GuiControl, Dic: Focus, DIC_LVTXT3

return 1
}

WBStats() {

	static baseTxt
	If !baseTxt {
		baseTxt =
		( LTrim
		  Medizinisches Wörterbuch
		  Neu jetzt:  `t#1#
		  Neu heute:`t#2#
		  Gesamt:    `t#3#
		  anderes Wörterbuch
		  Neu jetzt:  `t#4#
		  Neu heute:`t#5#
		  Gesamt:    `t#6#
		)

		stats.MedicalOld 	:= Medical.Count()
		stats.OthersOld 	:= Others.Count()

	}

	statTxt := StrReplace(baseTxt	, "#2#", Medical.Count()-stats.MedicalOld)
	statTxt := StrReplace(statTxt	, "#3#", Medical.Count())
	statTxt := StrReplace(statTxt	, "#5#", Others.Count()-stats.OthersOld)
	statTxt := StrReplace(statTxt	, "#6#", Others.Count())

	GuiControl, DIC:, DIC_GS, % statTxt

}

RunCrazyNGrams(NgramSizes:="2,3,4", save:=true) {

	ngrams := Object()

	For idx, NgramSize in StrSplit(NgramSizes, ",", A_Space)
		ngrams[NgramSize] :=  CrazyNGrams(NgramSize, save)

return ngrams
}

CrazyNGrams(NgramSize, save:=true) {

	static tmp
	ngrams := Object()
	VouwelNgrams := Object()
	wpos := 0

	; Vorbereitungen
	If !IsObject(tmp) {

		tmp	:= Array()
		For word, wdata in Medical {
			word := RegExReplace(word, "[A-ZÄÖÜ][A-ZÄÖÜ]", "-")
			word := RegExReplace(word, "[a-zäöüß][A-ZÄÖÜ]", "-")
			splitstring := StrSplit(word, "-")
			If IsObject(splitstring)
				For strIdx, strWord in splitstring
					If !RegExMatch(strWord, "^[A-ZÄÖÜ]+$")
						tmp.Push(strWord)
			else
				If !RegExMatch(word, "^[A-ZÄÖÜ]+$")
					tmp.Push(word)
		}

	}

	For idx, word in tmp {

		If Mod(idx, 500) = 0
			ToolTip, % "ngrams[" NgramSize "] "  idx "/" tmp.Count()
		If (StrLen(word) < NgramSize)
			continue

		occurSubs := {}
		Loop, % (StrLen(word) - (NgramSize-1)) {

			sub := Format("{:L}" , SubStr(word, A_Index, NgramSize))
			If InStr(sub, "-") || RegExMatch(sub, "\d")
				continue

			If !ngrams.haskey(sub)
				ngrams[sub] := {"count":1}
			else
				ngrams[sub].count += 1

		; Vorkommen und Position(en) innerhalb des Strings
			StrReplace(word, sub,, occurrenceCount)
			If !ngrams[sub].haskey("occurrence" occurrenceCount)
				ngrams[sub]["occurrence" occurrenceCount] := 1
			 else
				ngrams[sub]["occurrence" occurrenceCount] += 1

			If !occurSubs.haskey(sub) {
				occurSubs[sub] := 1
				Loop % occurrenceCount {
					wpos:=InStr(word, sub,,1, A_Index)
					If (wpos = 1)
						occurPos := "start"
					else If (wpos+StrLen(sub) = StrLen(word))
						occurPos := "end"
					else
						occurPos := "mid"
					If !ngrams[sub].haskey("pos_" occurPos)
						ngrams[sub]["pos_" occurPos] := 1
					else
						ngrams[sub]["pos_" occurPos] += 1
				}
			}


			If RegExMatch(SubStr(sub, NgramSize-1, 1), "i)[aeouiäöü]") {
				If !VouwelNgrams.haskey(sub)
					VouwelNgrams[sub] := {"count":1}
				else
					VouwelNgrams[sub].count += 1
			}


		}

	}

	ToolTip

	If save {
		JSONData.Save(adm.DBPath "\Dictionary\MedicalNgrams-" NgramSize ".json", ngrams, true,, 1, "UTF-8")
		JSONData.Save(adm.DBPath "\Dictionary\MedicalVouwelNgrams-" NgramSize ".json", VouwelNgrams, true,, 1, "UTF-8")
	}
	;Run % adm.DBPath "\Dictionary\MedicalVouwelNgrams-" NgramSize ".json"

return ngrams
}

ReIndexMedicalList() {   ;-- ### obsolet

	If !IsObject(Medical)
		return

	For word, wcount in Medical
		Medical[word] := 0

	For fidx, filepath in FilesP {

		IF (Mod(fidx, 100) = 0)
			ToolTip, % A_ThisFunc ": " fidx "/" FilesP.MaxIndex() "`n" filepath, 4200, 50
		text           	:= FileOpen(filepath, "r", "UTF-8").Read()
		spos          	:= 1

		while (spos := RegExMatch(text, "\s(?<Name>[A-ZÄÖÜ][\pL]+(\-[A-ZÄÖÜ][\pL]+)*)", P, spos)) {
			spos += StrLen(PName)
			If Medical.haskey(PName)
				Medical[PName] += 1
		}

	}

	FileOpen(adm.medicalWords, "w", "UTF-8").Write(cJSON.Dump(Medical, 1))
	;~ JSONData.Save(adm.medicalWords	, Medical	, true,, 1, "UTF-8")

}

CollectWords(filepath, debug:=false, byref errorlist:="") {

	;   phlie 				phie        	,   formuflar 		formular   	, 🗸Aiforderung    	Anforderung   ,	🗸Aipträumen   	Alpträumen
	;   Akkordiohn		Akkordlohn   	,   Akkurdlohn  	Akkordlohn
	; 🗸Aklenzeichen 	Aktenzeichen 	,   Nierenlaqer 	Nierenlager	,	🗸Laqer          	Lager       	,	🗸Doniinikos   	Dominikus
	; 🗸klinlk        	klinik       	, 🗸Dlagnose      	Diagnose
	; 🗸Leislung     	Leistung     	,	🗸Yberweisung  	Überweisung	,	🗸protruslon     	protrusion  	,	🗸glelten      	gleiten
	; 🗸Multl        	Multi        	, 🗸rhvthmus     	rhythmus
	;   zeil         	zeit         -> wann ist aber zeile gemeint?
	; 🗸FPisode      	Episode     	,	🗸koagulalion  	koagulation	,	🗸Operatlon       	Operation
	; 🗸frcquenz      	frequenz			,	🗸Vundauflage  	Wundauflage
	;  Mldazolam, Gastroskople, Stimmllppe 	Stimmlippe

	; RegEx Satztrennung: (\p{Lu}\p{Ll}.*?\p{Ll}\.)(\s|$) -> $1\n

	global AlbisBriefe, FilesB, FilesP, thisText, wsCount

				            		; Wortanfang
	static rxchanger := {	"\sAif"    : " Anf"	   			, "\sAip":" Alp"    		      	    		, "\sYbe"     	: " Übe"			     	, "\sFp":" Ep", "\sVund":" Wund"
						      			; überall im Wort
						       		, 	"aqe"    : "age"	    		, "nii"           	: "mi"    	    		, "islu"      	: "istu"    	    	, "i)ultl":"ulti"							    	, "i)hvth":"hyth"
						       		, 	"nlk"    : "nik"     			, "(\sD|d)j"      	: "$1i"  			  		, "i)aklen"   	: "akten"		      	, "i)lelten":"leiten"					    	, "i)laqe":"lage"
						       		,	"i)slon"   : "sion"		  		, "([^aeiouäüö])vv"	: "$1w" 	         	, "q([^u])"   	: "g$1"		         	, "(\p{Ll})B(\p{Ll})":"$1ß$2"	    	, "i)rcq":"req"
						       		,	"i)klomie" : "ktomie"      	, "i)(\p{Ll})clorf"	: "$1dorf"        	, "i)storung" 	: "störung"        	, "i)slände":"stände"              	, "i)otoge":"ologe"
						       			; Wortende
						       		, "nr(\s|\.)"  : "nummer "  	, "\-Teststr\s"   	: "Teststreifen " 	, "Teststr\s" 	: "Teststreifen " 	, "lalion\s":"lation " 	   		   	, "([ak])tlon\s":"$1tion "
						       		,	"i)arztin\s" : "ärztin "		, "ungs[\s\.,:;]" 	: "ung "
						       		,	"rlum\s"     : "rium "    	, "salion\s"      	: "sation "        	, "\srichle\s"	: "richte "		    	, "(\p{Ll})B\s":"$1ß "     }
	static vouwels    	:= "aeiouäöü"
	static consonants 	:= "qwrtzpsdfghjklxcvbnm"
	static clock        := "\d{1,2}\s*[\:\;\.]\s*\d{1,2}"
	static ocrcorrection:= {"rt":"h", "ri":"n", "x":"x", "x":"x", "x":"x"}
	static Delimiters 	:= "\-\/"
	PNames    	:= Object()
	MFreqLimit	:= 3
	OFreqLimit 	:= 2
	spos       	:= 1
	errorlist   := ""

	if !FileExist(filepath) {
		errorlist := "file not exists"
		return
	}

  ; Text aus der Datei laden
	If RegExMatch(filePath, "i)\.(pdf|doc|docx)$")  {

		try {
		  textO := IFilter(filepath)
		}
		catch e {
			textO := ""
			errorlist := "IFilter: " cJSON.dump(e)
			return
		}

	}
	else if RegExMatch(filePath, "i)\.txt$")
		textO := FileOpen(filepath, "r", "UTF-8").Read()
	else {
		Splitpath, filePath, fdir, fName, fOutExt, fExt
		errorlist := "unknown file extension: ." fExt
		return
	}


	If (StrLen(Trim(textO)) < 3 ) {
		errorlist := "text is read from file is empty"
		return
	}

  ; einfache Text/OCR Fehlerkorrekturen
	textO := RegExReplace(" " textO " ", "[\n\r\f]+", " ")
	text  := StrReplace(textO, "ı", "i")                                                                   	; ein i anstatt ı
	text  := RegExReplace(text, "(\p{Ll})(\d+)\s", "$1 $2 ")                                               	; Nummer6 -> Nummer 6
	text  := RegExReplace(text, "\d{1,2}\s*\.\s*\d{1,2}\s*\.\s*(\d{4}|\d{2})", " ")                        	; entfernt Datumszahlen
	text  := RegExReplace(text, "i)(" clock ")(\s*\-\s*" clock ")*(\s*Uhr)*\s", " ")                       	; entfernt Zahlenwerte und Uhrzeiten (12:30 Uhr* o. 12.30-14.30 Uhr)
	;~ text := RegExReplaceAll(text, "\s(-[\p{Ll}\-]*|\p{Ll})\s", "  ")                                    	; entfernt Folgen v. Kleinbuchstaben nach Leerzeichen
																																																		  		;	gefolgt v. einem Minus ' -abcdef'
	text  := RegExReplace(text, "\s\p{Ll}\.\s", " ")                                                       	; einbuchstabige Abkürzungen > v.
	text  := RegExReplaceAll(text, "(\s|\()\s*\d+\s*[\)\/\:]*[\s\d\:\.\,\-\/]+(\s|\))", " ")               	; entfernt Telefonnummern > (040) 1234 56 78
																																																			  	;	aber auch 040 - 1234 56 78
	text  := RegExReplace(text, "(\p{Ll})\-\s+(\p{Ll}{2,})", "$1$2")                                       	; mittel- gradige -> mittelgradige
	text  := RegExReplace(text, "(\d+\-|\p{Lu}\p{Ll}+\d*\-)\s+(\p{Lu}\p{Ll}{2,})", "$1-$2")                	; COVID-19-Schutzmaßnahmen-  Ausnahmeverordnung
	text  := RegExReplace(text, "(\p{Ll}{2,}|\p{Lu}\p{Ll})(\p{Lu}\p{Ll}{2,})", "$1 $2")                    	; trennt: rechtskonvexeSkoliose - aber auch > rechtskonveXeSkoliose
	text  := RegExReplace(text, "(\p{Lu}{2,}|[^\pL]\p{Lu})(\p{Lu}\p{Ll})", "$1 $2")                        	; MNeurologie -> M Neurologie
	text  := RegExReplace(text, "(\p{Ll})q\s", "$1g ")                                                     	; ausgiebiq -> ausgiebig
	text  := RegExReplace(text, "(\p{Ll}\d*)\s+(\-\p{Lu})", "$1$2")                                        	; Sauerstoff -Sättigung -> Sauerstoff-Sättigung
	text  := RegExReplace(text, "\s", "   ")                                                               	; Leerzeichen verdoppeln

	rtext := text
	Loop {
		ltext := rtext
		rtext := RegExReplace(rtext, "\s\p{Lu}{6,}\s", "  ")                                                	; DAEXXXMM
		rtext := RegExReplace(rtext, "\s[\\/\-\s]+\s", "  ")                                                 	; entfernt  - -D
		rtext := RegExReplace(rtext, "\s(\d+)(\p{Lu}\p{Ll}{2,})\s", "  $1 $2  ")                             	; 89Fortbildung 89 Fortbildung
		rtext := RegExReplace(rtext, "(\p{Ll}{3,})\/(\p{Lu}\p{Ll}{2,})", "$1 / $2")                         	; Unterschrift/Stempel -> Unterschrift / Stempel
		rtext := RegExReplace(rtext, "(\s|[^\pL\d])\d+[\.\,;]\d+(\s|[^\pL\d])", " $1  $2 ")                  	; entfernt Werte wie 114,01
		rtext := RegExReplace(rtext, "\s\d+\s*\-\s*\d+(\s*\-\s*\d+)*", "  ")                                	; entfernt 0-0-1-0
		rtext := RegExReplace(rtext, "\s\pL{1,2}\s", "  ")                                                   	; entfernt Worte mit max 2 Buchstaben
		rtext := RegExReplace(rtext, "\s\-\p{Lu}+\s", "  ")                                                  	; entfernt -JS
		rtext := RegExReplace(rtext, "\s\-*(\p{Lu}+\d+|\d+\p{Lu}+)[\p{Lu}\d]*\s", "  ")                      	; entfernt -J7S
		rtext := RegExReplace(rtext, "\-{2,}", "-")                                                          	; --- -> -
		rtext := RegExReplace(rtext, "\s\d+\s", "  ")                                                        	; entfernt alle noch einzeln stehenden  Zahlen
	} until (rtext = ltext)
	text := RegExReplace(rtext, "\s{2,}", " ")

	For rxmstring, rxrplstring in rxchanger
		text := RegExReplace(text, rxmstring, rxrplstring)

	text := RegExReplace(text, "(\p{Lu}\p{Ll}{2,})stra*s*s*e*\s", "$1straße ")                            	; Hauptstr     	-> Hauptstraße
	text := RegExReplace(text, "(\s|\-)Str(\.|asse)*\s", "$1Straße ")                                      	; Adalbert-Sauerampfer-Str.     	-> Adalbert-Sauerampfer-Straße
	text := StrReplace(text, "Teststraße", "Teststreifen")                                                	; Teststraße     	-> Teststreifen
	text := RegExReplace(text, "[^\pL\d\-\s\/]", " ")                                                      	; alle Zeichen ausser diesen entfernen
	text := RegExReplace(text, "\s{2,}", " ")

	If (StrLen(Trim(text)) = 0) {
		If dbg
			SciTEOutput("File: " filepath " contains no text after RegExReplace" )
		return
	}

	tLen := StrLen(text)
	RegExReplace(text, "\s",, wsCount)
	steps := Round(wsCount/20) + 1
	ToolTip, % "Pos: " "0/" wsCount, 900, 1

	outwords := knownwords := ""

	For wpos, TName in StrSplit(text, A_Space) {

			TName := Trim(TName, "-")
			TName := Trim(TName, "/")
			If !TName
				continue

			TName := RegExReplace(TName, "str$", "straße")
			TName := TName ~= "Teststraße" ? StrReplace(TName, "Teststraße", "Teststreifen") : TName

		; durch / oder - getrennte Worte einzeln untersuchen und aufnehmen
			XNames := Array()

			noTName := ""
			If (TName ~= "\-" || TName ~= "\/") {

				dName := RegExReplace(TName, "\d", "", dCount)
				dName := RegExReplace(dName, "(\-|\/)", "")
				If (dCount >= StrLen(dName))
					noTName := "A"

				tmpwords := RegExReplace(TName, "[\-\/]", " ")
				tmpwords := RegExReplace(tmpwords, "\s{2,}", " ")
				For each, tmpword in StrSplit(tmpwords, A_Space)
					If tmpword {
						If (tmpword ~= "\d\pL") {
							noTName .=  "B"
							outwords .= "Uu- " tmpword  "`n"
						} else
							XNames.Push(tmpword)
					}

			}
			If (TName ~= "\d\p{Ll}") || (TName ~= "(\/|\-)\d+$") || noTName {
				outwords .= "UD- " TName  "`n"
				;~ SciTEOutput("UD " TName " -- " noTName "`n")
			}
			else
				XNames.Push(TName)

			;~ If (XNames.Count()=0)  ;
					;~ continue

		; Progress
			If (Mod(wpos, steps) < 1)
				ToolTip, % "Pos: " wpos "/" wsCount, 900, 1

		; Worte in
			For each, PName in XNames {

				; sehr kurze Worte und reine Zahlen oder Zahlen-Bindestrichworte nicht untersuchen
					If (StrLen(PName) < 4) || (PName ~= "^[\d\-\/\p{Lu}ß]+$") {
							continue
					}

				; aussortieren von Worten mit mehr als 4 Konsonanten in Folge, bestimmte konsonantenfolgen werden mit einem Sternchen ersetzt und so nur als einzelner Konsonant gezählt
					If RegExMatch(PName, "i)[" consonants "]{4,}", match) {
						match := RegExReplace(match, "i)sch", "*")
						match := RegExReplace(match, "i)ch"	, "*")
						match := RegExReplace(match, "i)str"	, "*")
						match := RegExReplace(match, "i)st"	, "*")
						match := RegExReplace(match, "i)pr"	, "*")
						match := RegExReplace(match, "i)[^s]s", "*")
						If (StrLen(match)>4) {
							outwords .= "B- " PName " [" match "]`n"
							continue
						}
					}

				; aussortieren von Worten mit mehr als 3 Vokalen in Folge (ohne e!)
					If RegExMatch(PName, "i)[aeiouäöü]{4,}", match) {
						outwords .= "C- " PName " [" match "]" "`n"
						continue
					}

					takeword := 0
					takename := "F"

				; wenn ein Wort in einem der beiden Wörterbücher vorkommt
					if Medical.haskey(PName) {                                            	; Wörterbuch selbst
						Medical[PName] += 1
						takeword := (Medical[PName] <= MFreqLimit) 	? 2 : 0   ; Wort anzeigen wenn es noch nicht häufig gesehen wurde
						takename := "m"
						knownwords .= "🗸 " PName "`n"
					}
					else if Others.haskey(PName) {                                         ; Others
						Others[PName] += 1
						takeword := (Others[PName] <= OFreqLimit) ? 3 : 0
						takename := "o"
						knownwords .= "🗸 " PName "`n"
					}

				; ausgeschlossene Buchstabenkombinationen - kommen zunächst ins Others Wörterbuch
					else If RegExMatch(PName, "([\p{Ll}\-]\p{Lu}{2,}|^[\p{Lu}ß\-]+$)") {                                    ; men-LLweis oder ^ABCDEF$
						Others[PName] := !Others.haskey[PName] ? 1 : Others[PName]+1
						takeword := (Others[PName] <= OFreqLimit) ? 3 : 0
						takename := "k"
					}
				;
					else {

						If !InStr(PName, "-")
							If RegExReplace(PName,"\p{Lu}",, LenUpperChars)	 	;
								If (LenUpperChars > StrLen(PName)/3)   {              ; gleich ins Others WBuch, falls Großbuchstaben mehr als 1/3 aller Buchstaben ausmachen
									Others[PName] := !Others.haskey[PName] ? 1 : Others[PName]+1
									outwords .= "E- " PName "`n"
									continue
							}

						takeword := ngramWordCheck(PName, 5) ? 1 : 4
						takename := "n"
					}

				If !takeword {
					outwords .= takename  "- " takeword ": " PName "`n"
					continue
				}

				tw3:= takeword
				If !PNames.haskey(PName)
					PNames[PName] := {"count"    	: 1
												, 	 "list"      	: takeword
												, 	 "listcount"	: (takeword=2 ? Medical[PName]
																	:  takeword=3 ? Others[PName]    : 0)}
				else {
					PNames[PName].count += 1
				}

		}

	}

	thisText := text "`n- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - `n" textO
	thisText .= "`n+ + + + + + + + + + + + + + +`n" RegExReplace(outwords, "[\n\r]{2,}", "`n")
	thisText .= "`n~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~`n" knownwords

	ToolTip, % "Position: " spos "/" tLen, 900, 1
	TTOff(1, 5000)

return PNames
}

RegExReplaceAll(text, rxReplaceString, rxReplaceWith) {

	If !rxReplaceString
		return text

	Loop {                                                                                                                          	; -e
		text := RegExReplace(text, rxReplaceString, rxReplaceWith, rplCount)
	} until (rplCount = 0)

return text
}

Old_ngramWordCheck(word, depth:=5) {

	If (StrLen(word)-depth >= 0)
	Limit := depth - 2
	hits := 0
	Loop, % Limit {
		ngramLen := A_Index + 2
		If (StrLen(word) >= ngramLen && ngramSyntaxValidate(word, ngramLen))
			hits ++
	}

return hits=Limit ? true : false
}

ngramWordCheck(word, depth:=5, Limit:=2) {

	sword 	:= word
	words	:= StrSplit(sword, "-")
	hits     	:= 0
	For each, word in words {

		wLen 	:= StrLen(word)
		If (wLen <= 2)
			continue
		ngramLen	:= depth >= wLen ? wLen-1 : depth
		ngramsPerWord := wLen - (ngramLen-1)
		ngramToFailMax := Floor(ngramsPerWord/ngramLen) = 0 ? 1 : 2
		nongramhits := ngramSyntaxValidate(word, ngramLen)
		hits += nongramhits < ngramToFailMax ? 1 : 0

	}

	If (dbg && hits < words.Count())
		SciTEOutput( "hits: " SubStr("00" hits, -1) ", perWord: " ngramsPerWord " : " nongramhits " nonhits , FailMax: " ngramToFailMax " -- " word)

return hits>=words.Count() ? true : false
}

ngramSyntaxValidate(word, NgramSize, updateNgrams:=false) {

	global ngrams
	If (!IsObject(ngrams) || updateNgrams)
		ngrams := RunCrazyNGrams("3,4,5", false)

	nongramhits := ngramsPerWord := 0
	For idx, sword in StrSplit(word, "-") {
		ngramsPerWord += StrLen(sword) - (NgramSize-1)
		Loop, % ngramsPerWord
			If !ngrams[NgramSize].haskey(SubStr(sword, A_Index, NgramSize) )
				nongramhits ++
	}
				;~ return false

return nongramhits
}

WordEndings(dictObject, EndingsObject:="") {

	/*

		definitiver Endungsausschluß bei :  äch,
												sicher wenn davor ein                         , aber keine Endung wenn davor
		endung ist : 	s			+		e,f,de,m,n,rn,x
										-		al,s

							se			+
										-		ü,a

							e			+		ab,äch,uch,en,od,gt,hm,ich,ig,kt,lt,nd,nnt,r,t,und,ß
										-		ag,am,eg,er,hm,ht,ll,nd,ng,nt,om,os,rb,st,ug

							en		+		ab,äch,ft,hm,hr,kt,on,uch,und,ung,re,t
										-		ht,ieg,it,od,om,nd,ing,sch,

							er			+		en,ig,t,
										-		ht,

							es			+		ef,er,hr,ll,nd,und,st,ß
							ig			+		ünd(>und),
							ung		-
							n 			+    	er,hme,nte,ose,te,uge
										-		ite
							sein		+		en

	 */


	If FileExist(adm.PathWordEndings) && !IsObject(EndingsObject)
		EndingsObject := cJSON.Load(FileOpen(adm.PathWordEndings, "r", "UTF-8").Read())

	SciTEOutput("splitting dictionary words....")

  ; split words containing char "-" and normalize umlauts like porter stemmer
	tmpObj := Object()
	For word, wcount in dictObject {

		words := StrSplit(word, "-")
		If !words.Count()
			SciTEOutput(word)

		For each, strippedword in words {
			Normalized := NormalizeUmlauts(strippedword)
			If !IsObject(tmpObj[Normalized])
				tmpObj[Normalized] :=  {"s":strippedword, "o":[word], "c": wcount}
			else {
				 tmpObj[Normalized].c += wcount
				 ofound := false
				 For oIndex, originalword in tmpObj[Normalized].o
					If (originalword = word) {
						ofound := true
						break
					}
				If !ofound
					tmpObj[Normalized].o.Push(word)
			}
		}
		;~ else {
			;~ Normalized := NormalizeUmlauts(word)
			;~ tmpObj[word] := wcount
		;~ }

	}

	SciTEOutput("splitting from original " dictObject.Count() " words to " tmpObj.Count() " words")

  ;
	If dbg {
		FileOpen(A_Temp "\tmpObj.json", "w", "UTF-8").Write(cJSON.Dump(tmpObj, 1))
		Run % A_Temp "\tmpObj.json"
	}




}

NormalizeUmlauts(word) {

	static normals := {"ß":"ss","ae": "ä","oe": "ö","(?=[^qaeou])ue":"ü"}
	For char, Normalizer in normals
		word := RegExReplace(word, char, Normalizer)

return word
}

ToLetterBasedDict(dictObject) {

	static subs

	LetterBased := Object()

	;~ For word, wcount in dictObject {

		;~ One := SubSttr(word, 1, 1)

		;~ If !IsObject(LetterBased[One]) {
			;~ LetterBased[One] := {(word) : wcount}
		;~ }

		;~ else if {

		;~ }


	;~ }

}

TLBSubWords(subobject, word, subword) {

		static subs := Array()
		static lword


		;~ If (word != lastword) {
			;~ lastword 	:= word
			;~ subs     	:= Array()
		;~ }
		;~ subwordMatched := false

		;~ For subword, nextSubObject in SubObject {
			;~ If IsObject(nextSubObject) {
				;~ If (SubStr(word, 1, StringLen(subword)) = subword) {
					;~ subwordMatched := true
					;~ subs.InsertAt(1, subword)
				;~ }

			;~ }

		;~ }


}

class stemmer{  ;;;???

	; for german stemming algorithm variant
	; http://snowball.tartarus.org/algorithms/german/stemmer.html

		__New() {

			this.rx := {}
			this.rx.PreStem	:= {	"ae"                   	: "ä"
										,	"oe"                   	: "ö"
										,	"(?=[^qaeou])ue"	: "ü"
										,	"(?=[aeou])u(?!e)" 	: "U"
										,	"y"                       	: "Y"
										,	"ß"                    	: "ss"  }
			this.rx.Umlauts	:= {"ä":"a","ö":"o","ü":"u","ß":"ss"  }
			this.rx.Endings  	:= "[bdfghklmnrt]"
			this.rx.Suffix    	:= "(ung|keit|ig|em|er|ern|en|es|est|ig|isch|ik|lich|heit)$"
			;							"ie|ion|mus|um|enz|ose|eg|ng|us|at"

		}

}

IFilter(file, searchstring:="") {                                                                            	;-- Textsuche/Textextraktion in PDF/Word Dokumenten

	static cchBufferSize              	:= 4 * 1024
	static resultstriplinebreaks   	:= false

	static CHUNK_TEXT 	:= 1
	static STGM_READ 	:= 0
	static IFILTER_INIT_CANON_PARAGRAPHS	    	:= 1
	static IFILTER_INIT_HARD_LINE_BREAKS        		:= 2
	static IFILTER_INIT_CANON_HYPHENS 	        	:= 4
	static IFILTER_INIT_CANON_SPACES	            	:= 8
	static IFILTER_INIT_APPLY_INDEX_ATTRIBUTES	:= 16
	static IFILTER_INIT_APPLY_OTHER_ATTRIBUTES	:= 32
	static IFILTER_INIT_INDEXING_ONLY           		:= 64
	static IFILTER_INIT_SEARCH_LINKS	                	:= 128
	static IFILTER_INIT_APPLY_CRAWL_ATTRIBUTES	:= 256
	static IFILTER_INIT_FILTER_OWNED_VALUE_OK	:= 512
	static IFILTER_INIT_FILTER_AGGRESSIVE_BREAK	:= 1024
	static IFILTER_INIT_DISABLE_EMBEDDED	       	:= 2048
	static IFILTER_INIT_EMIT_FORMATTING	        	:= 4096

	static S_OK   	:= 0
	static FILTER_S_LAST_TEXT 	     	:= 268041
	static FILTER_E_NO_MORE_TEXT 	:= -2147215615

	if (!A_IsUnicode)
		throw A_ThisFunc ": The IFilter APIs appear to be Unicode only. Please try again with a Unicode build of AHK."

	if (!file || !FileExist(file))
		return -99

	SplitPath, file,,, ext
	VarSetCapacity(FILTERED_DATA_SOURCES, 4*A_PtrSize, 0)
	NumPut(&ext, FILTERED_DATA_SOURCES,, "Ptr")
	VarSetCapacity(FilterClsid, 16, 0)

	; Adobe workaround
	if (job := DllCall("CreateJobObject", "Ptr", 0, "Str", "filterProc", "Ptr"))
		DllCall("AssignProcessToJobObject", "Ptr", job, "Ptr", DllCall("GetCurrentProcess", "Ptr"))

	FilterRegistration := ComObjCreate("{9E175B8D-F52A-11D8-B9A5-505054503030}", "{c7310722-ac80-11d1-8df3-00c04fb6ef4f}")
	if (DllCall(NumGet(NumGet(FilterRegistration+0)+3*A_PtrSize), "Ptr", FilterRegistration, "Ptr", 0, "Ptr", &FILTERED_DATA_SOURCES, "Ptr", 0, "Int", false, "Ptr", &FilterClsid, "Ptr", 0, "Ptr*", 0, "Ptr*", IFilter) != 0 ) ; ILoadFilter::LoadIFilter
		throw A_ThisFunc ": can't load IFilter"

	ObjRelease(FilterRegistration)

	if (DllCall("shlwapi\SHCreateStreamOnFile", "Str", file, "UInt", STGM_READ, "Ptr*", iStream) != 0 )
		throw A_ThisFunc ": can't open filepath"

	PersistStream := ComObjQuery(IFilter, "{00000109-0000-0000-C000-000000000046}")
	if (DllCall(NumGet(NumGet(PersistStream+0)+5*A_PtrSize), "Ptr", PersistStream, "Ptr", iStream) != 0 ) ; ::Load
		throw A_ThisFunc ": can't load filestream"

	ObjRelease(iStream)

	status := 0
	MY_IFILTER := IFILTER_INIT_CANON_PARAGRAPHS | IFILTER_INIT_HARD_LINE_BREAKS | IFILTER_INIT_CANON_HYPHENS | IFILTER_INIT_CANON_SPACES
	if (DllCall(NumGet(NumGet(IFilter+0)+3*A_PtrSize), "Ptr", IFilter, "UInt", MY_IFILTER, "Int64", 0, "Ptr", 0, "Int64*", status) != 0 ) ; IFilter::Init
		throw A_ThisFunc ": can't init IFilter"

	VarSetCapacity(STAT_CHUNK, A_PtrSize == 8 ? 64 : 52)
	VarSetCapacity(buf, (cchBufferSize * 2) + 2)

	while (DllCall(NumGet(NumGet(IFilter+0)+4*A_PtrSize), "Ptr", IFilter, "Ptr", &STAT_CHUNK) == 0) { ; ::GetChunk
		if (NumGet(STAT_CHUNK, 8, "UInt") & CHUNK_TEXT) {
			while (DllCall(NumGet(NumGet(IFilter+0)+5*A_PtrSize), "Ptr", IFilter, "Int64*", (siz := cchBufferSize), "Ptr", &buf) != FILTER_E_NO_MORE_TEXT) { ; ::GetText
				IFText := StrGet(&buf,, "UTF-16")
				if (resultstriplinebreaks)
					IFText := StrReplace(IFText, "`r`n")
				If searchstring && (strpos := RegExMatch(IFText, searchstring))
					break

			}
		}
	}

	ObjRelease(PersistStream)
	ObjRelease(iFilter)
	if (job)
		DllCall("CloseHandle", "Ptr", job)

	;~ If (strpos > 0)
		;~ return strpos

return searchstring ? strpos : IFText
}

PDFisSearchable(pdfFilePath)	{                                                                              	;-- durchsuchbare PDF Datei?

	; letzte Änderung 03.12.2022 : Recursion limit exceed prevention

	If !IsObject(fobj := FileOpen(pdfFilePath, "r", "CP1252"))
		return 0

	filepos := 0
	while !fobj.AtEOF {
		If RegExMatch(fobj.ReadLine(), "i)(Font|ToUnicode|Italic)") {
			filepos := fobj.pos
			break
		}
		else If RegExMatch(line, "Length\s(?<seek>\d+)", file)     	; binären Inhalt überspringen
			fobj.seek(fileseek, 1)                                                   	; •1 (SEEK_CUR): Current position of the file pointer.
	}

	fobj.Close()

return filepos
}

inArr(str, arr) {

  For k,v in arr
    if (str=v)
      return k   ; Position zurückgeben

return 0
}

RemoveDoublettes(objA, objB) {                                                                             	;-- entfernt doppelte Einträge

	removed := 0
	objBClone := objB.Clone()

	For word, wcount in objA {
		ToolTip, % "objB:      " objB.Count() "`nremoved: " removed "`nobjA:      " objA.Count()
		If objB.haskey(word) {
			objBClone.Delete(word)
			removed ++
		}
	}

	ToolTip

return {"objA": objA, "objB":objBClone}
}

RemoveDriveLetter(filesArr) {

	MouseGetPos, mx, my

	steps := Floor(filesArr.Count() / 100)
	showPrg := filesArr.Count() > 20000 ? true : false

	for findex, filePath in filesArr {
		filesArr[findex] := RegExReplace(filepath, "^\w+?(:|\\)")
		if (showPrg && Mod(findex, steps) = 0)
			ToolTip, % "Verzeichnisse bearbeitet: " findex "/" filesArr.Count(), % mx , % my, 13

	}

	ToolTip,,,,13

return filesArr
}

; ab hier Funktionen für ini-Read und AddendumBaseProperties
IniReadExt(SectionOrFullFilePath, Key:="", DefaultValue:="", convert:=true) {            	;-- eigene IniRead funktion für Addendum

	; beim ersten Aufruf der Funktion !nur! Übergabe des ini Pfades mit dem Parameter SectionOrFullFilePath
	; die Funktion behandelt einen geschriebenen Wert der "ja" oder "nein" ist, als Wahrheitswert, also true oder false
	; UTF-16 in UTF-8 Zeichen-Konvertierung
	; Der Pfad in Addendum.Dir wird einer anderen Variable übergeben. Brauche dann nicht immer ein globales Addendum-Objekt
	; letzte Änderung: 31.01.2021

		static admDir
		static WorkIni

	; Arbeitsini Datei wird erkannt wenn der übergebene Parameter einen Pfad ist
		If RegExMatch(SectionOrFullFilePath, "^[A-Z]\:.*\\")	{
			If !FileExist(SectionOrFullFilePath)	{
				MsgBox,, % "Addendum für AlbisOnWindows", % "Die .ini Datei existiert nicht!`n`n" WorkIni "`n`nDas Skript wird jetzt beendet.", 10
				ExitApp
			}
			WorkIni := SectionOrFullFilePath
			If RegExMatch(WorkIni, "[A-Z]\:.*?AlbisOnWindows", rxDir)
				admDir := rxDir
			else
				admDir := Addendum.Dir
			return WorkIni
		}

	; Workini ist nicht definiert worden, dann muss das komplette Skript abgebrochen werden
		If !WorkIni {
			MsgBox,, Addendum für AlbisOnWindows, %	"Bei Aufruf von IniReadExt muss als erstes`n"
																			. 	"der Pfad zur ini Datei übergeben werden.`n"
																			.	"Das Skript wird jetzt beendet.", 10
			ExitApp
		}

	; Section, Key einlesen, ini Encoding in UTF.8 umwandeln
		IniRead, OutPutVar, % WorkIni, % SectionOrFullFilePath, % Key
		If convert
			OutPutVar := StrUtf8BytesToText(OutPutVar)

	; Bearbeiten des Wertes vor Rückgabe
		If InStr(OutPutVar, "ERROR")
			If (StrLen(DefaultValue) > 0) { ; Defaultwert vorhanden, dann diesen Schreiben und Zurückgeben
				OutPutVar := DefaultValue
				IniWrite, % DefaultValue, % WorkIni, % SectionOrFullFilePath, % key
				If ErrorLevel
					TrayTip, % A_ScriptName, % "Der Defaultwert <" DefaultValue "> konnte geschrieben werden.`n`n[" WorkIni "]", 2
			}
			else return "ERROR"
		else if InStr(OutPutVar, "%AddendumDir%")
				return StrReplace(OutPutVar, "%AddendumDir%", admDir)
		else if RegExMatch(OutPutVar, "%\.exe$") && !RegExMatch(OutPutVar, "i)[A-Z]\:\\")
				return GetAppImagePath(OutPutVar)
		else if RegExMatch(OutPutVar, "i)^\s*(ja|nein)\s*$", bool)
				return (bool1= "ja") ? true : false

return Trim(OutPutVar)
}

StrUtf8BytesToText(vUtf8) {                                                                                       	;-- Umwandeln von Text aus .ini Dateien in UTF-8
	if A_IsUnicode 	{
		VarSetCapacity(vUtf8X, StrPut(vUtf8, "CP0"))
		StrPut(vUtf8, &vUtf8X, "CP0")
		return StrGet(&vUtf8X, "UTF-8")
	} else
		return StrGet(&vUtf8, "UTF-8")
}

GetAppImagePath(appname) {                                                                                 	;-- Installationspfad eines Programmes

	headers:= {	"DISPLAYNAME"                  	: 1
					,	"VERSION"                         	: 2
					, 	"PUBLISHER"             	         	: 3
					, 	"PRODUCTID"                    	: 4
					, 	"REGISTEREDOWNER"        	: 5
					, 	"REGISTEREDCOMPANY"    	: 6
					, 	"LANGUAGE"                     	: 7
					, 	"SUPPORTURL"                    	: 8
					, 	"SUPPORTTELEPHONE"       	: 9
					, 	"HELPLINK"                        	: 10
					, 	"INSTALLLOCATION"          	: 11
					, 	"INSTALLSOURCE"             	: 12
					, 	"INSTALLDATE"                  	: 13
					, 	"CONTACT"                        	: 14
					, 	"COMMENTS"                    	: 15
					, 	"IMAGE"                            	: 16
					, 	"UPDATEINFOURL"            	: 17}

   appImages := GetAppsInfo({mask: "IMAGE", offset: A_PtrSize*(headers["IMAGE"] - 1) })
   Loop, Parse, appImages, "`n"
	If Instr(A_loopField, appname)
		return A_loopField

return ""
}

GetAppsInfo(infoType) {

	static CLSID_EnumInstalledApps := "{0B124F8F-91F0-11D1-B8B5-006008059382}"
        , IID_IEnumInstalledApps     	:= "{1BC752E1-9046-11D1-B8B3-006008059382}"

        , DISPLAYNAME            	:= 0x00000001
        , VERSION                    	:= 0x00000002
        , PUBLISHER                  	:= 0x00000004
        , PRODUCTID                	:= 0x00000008
        , REGISTEREDOWNER    	:= 0x00000010
        , REGISTEREDCOMPANY	:= 0x00000020
        , LANGUAGE                	:= 0x00000040
        , SUPPORTURL               	:= 0x00000080
        , SUPPORTTELEPHONE  	:= 0x00000100
        , HELPLINK                     	:= 0x00000200
        , INSTALLLOCATION     	:= 0x00000400
        , INSTALLSOURCE         	:= 0x00000800
        , INSTALLDATE              	:= 0x00001000
        , CONTACT                  	:= 0x00004000
        , COMMENTS               	:= 0x00008000
        , IMAGE                        	:= 0x00020000
        , READMEURL                	:= 0x00040000
        , UPDATEINFOURL        	:= 0x00080000

   pEIA := ComObjCreate(CLSID_EnumInstalledApps, IID_IEnumInstalledApps)

   while DllCall(NumGet(NumGet(pEIA+0) + A_PtrSize*3), Ptr, pEIA, PtrP, pINA) = 0  {
      VarSetCapacity(APPINFODATA, size := 4*2 + A_PtrSize*18, 0)
      NumPut(size, APPINFODATA)
      mask := infoType.mask
      NumPut(%mask%, APPINFODATA, 4)

      DllCall(NumGet(NumGet(pINA+0) + A_PtrSize*3), Ptr, pINA, Ptr, &APPINFODATA)
      ObjRelease(pINA)
      if !(pData := NumGet(APPINFODATA, 8 + infoType.offset))
         continue
      res .= StrGet(pData, "UTF-16") . "`n"
      DllCall("Ole32\CoTaskMemFree", Ptr, pData)  ; not sure, whether it's needed
   }
   Return res
}

AddendumBaseProperties(AddendumDir) {                                                               	;-- Basiseinstellungen für jedes Addendum-Modul

	; Hauptverzeichnis und ini Datei
		props          	:= Object()
		props.Dir      	:= AddendumDir
		props.Ini      	:= props.Dir "\Addendum.ini"

	; Log Dateipfad und Datenbankverzeichnis
		workini         := IniReadExt(props.Ini)
		props.LogPath 	:= IniReadExt("Addendum", "AddendumLogPath"	,, true)
		props.DBPath  	:= IniReadExt("Addendum", "AddendumDBPath" 	,, true)
		props.ScanPath 	:= IniReadExt("Scanpool", "Befundordner"  	,, true)

		RegExMatch(AddendumDir, "i)^(?<Letter>\w+?(:|\\))", Drive)
		if !Instr(FileExist(props.ScanPath), "D")
			props.ScanPath := RegExReplace(props.ScanPath, "i)^\w+(:|\\)", DriveLetter)
		if !Instr(FileExist(props.ScanPath), "D") {
			SciTEOutput("Der Befundordner wurde nicht gefunden: " RegExReplace(props.ScanPath, "i)^\w+(:|\\)"))
			props.ScanPath := ""
		}
		props.DriveLetter := DriveLetter



	; Log Dateipfad prüfen und evtl. anlegen
		If !RegExMatch(props.LogPath, "i)[A-Z]\:\\")
			props.LogPath := props.Dir "\logs'n'data"
		If !InStr(FileExist(props.LogPath), "D")
			FileCreateDir, % props.LogPath

	; Datenbankverzeichnis prüfen und evtl. anlegen
		If !RegExMatch(props.DBPath	, "i)[A-Z]\:\\")
			props.DBPath := props.Dir "\logs'n'data\_DB"
		If !InStr(FileExist(props.DBPath), "D")
			FileCreateDir, % props.DBPath

return props
}

GetAlbisPaths() {                                                                       	;-- ermittelt das Albisverzeichniss, sowie Unterverzeichnisse im albiswin Ordner

	nr := 1
	SetRegView	, % (A_PtrSize = 8 ? 64 : 32)

	regPathAlbis1 := "HKEY_CURRENT_USER\SOFTWARE\ALBIS\Albis on Windows\Albis_Versionen"

	Loop {

		nr := A_Index
		RegRead, MainPath 	, % regPathAlbis1, % nr "-MainPath"
		RegRead, LocalPath 	, % regPathAlbis1, % nr "-LocalPath"
		RegRead, Exe       	, % regPathAlbis1, % nr "-Exe"

		If  !(MainPath ~= "i)albis_demo") && InStr(FileExist(MainPath), "D") && FileExist(MainPath "\" Exe ){
			albisfound := true
			break
		}
		else if (A_Index > 10)
			break
	}


 ; HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows,
	If !albisfound {
		RegRead, MainPath		, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
		RegRead, LocalPath	, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, % LocalPath-2
		IF FileExist(MainPath "\albisCS.exe")
			Exe := "albisCS.exe"
		albisfound := true
	}

 ; Abbruch
	If !albisfound {
		throw "Der Dateipfad mit einer albis.exe oder albisCS.exe konnte nicht bestimmt werden"
		return
	}

	albisPaths := {"MainPath":MainPath, "LocalPath":LocalPath, "Exe":Exe, "Briefe":MainPath "\Briefe", "db":MainPath "\db", "Vorlagen":MainPath "\tvl"}

	RegExMatch(MainPath, "i)^(?<server>[A-Z]+?)((?<drive>:)|(?<smb>\\))", albis_)
	if albis_drive
		albisPaths["Drive"] := albis_server
	else if albis_smb
		albisPaths["smb"] 	:= albis_smb

return albisPaths
}


#include %A_ScriptDir%\..\..\lib\class_cJSON.ahk
#include %A_ScriptDir%\..\..\lib\class_JSON.ahk
#include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                                                     TEXT TO DICTIONARY INTERFACE
;
;      Funktion:
;
;      	Basisskript:    		-	keines
;		Abhängigkeiten: 	-	siehe #includes
;
;	                    				Addendum für Albis on Windows
;                        				by Ixiko started in September 2017 - last change 27.02.2021 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;

	; Skripteinstellungen                                        	;{

		#NoEnv
		#KeyHistory                   	, 0
		#MaxThreads                  	, 100
		#MaxThreadsBuffer       	, On
		#MaxMem                    	, 256

		SetBatchLines                	, -1
		;ListLines                        	, Off
		SetControlDelay               	, -1
		SetWinDelay                 	, -1
		AutoTrim                       	, On
		FileEncoding                 	, UTF-8

	;}

	; Variablen                                                       	;{
		global q:=Chr(0x22)
		global adm := Object()
		global stats := Object()
		global FilesP, FilesT, FilesO
		global ngrams

		RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)
		adm := AddendumBaseProperties(AddendumDir)
		Menu, Tray,  Icon, % adm.Dir "\assets\ModulIcons\Praxomat.ico"
	;}

	; Wörterbücher laden                                     	;{
		global medicalWords 	:= adm.DBPath "\Dictionary\Medical.json"
		global othersWords   	:= adm.DBPath "\Dictionary\Others.json"
		global FilesProcessed  	:= adm.DBPath "\Dictionary\FilesP.json"

		global Medical, Others, FilesP, FilesT

		MFreq 	:= {}, maxMFreq	:= 0, minMFreq:= 10000000
		SFreq	:= {}, maxSFreq	:= 0, minSFreq	:= 10000000

		Medical	:= FileExist(medicalWords) 	? JSONData.Load(medicalWords	, "", "UTF-8")	: Object()
		Others   	:= FileExist(othersWords)    	? JSONData.Load(othersWords  	, "", "UTF-8") 	: Object()
		FilesP    	:= FileExist(FilesProcessed)    	? JSONData.Load(FilesProcessed	, "", "UTF-8")	: Array()

		removedOthers := 0
		For word, wcount in Medical {

			If Others.haskey(word) {
				Medical[word] += Others[word]
				Others.Delete(word)
				removedOthers ++
			}

			maxMFreq:= wcount > maxMFreq	? wcount : maxMFreq
			minMFreq	:= wcount > minMFreq	? wcount : minMFreq
			MFreq[wcount] := !MFreq.haskey(wcount) ? 1 : MFreq[wcount] += 1
		}

		SciTEOutput((removedOthers ? removedOthers " identical words removed from others dictionary" : ""))

		freqs := []
		;~ points := "▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀"
		;~ For num, freq in MFreq
			;~ SciTEOutput( "|" num " " SubStr(points, 1, Floor(freq/8)) " " freq)

		; ___________________________________________

		For word, wcount in Others {
			maxSFreq	:= wcount > maxSFreq	? wcount : maxSFreq
			minSFreq 	:= wcount > minSFreq 	? wcount : minSFreq
			SFreq[wcount] := !SFreq.haskey(wcount) ? 1 : SFreq[wcount] += 1
		}

		;~ For num, freq in SFreq
			;~ SciTEOutput( "|" num " " SubStr(points, 1, Floor(freq/8)) " " freq)


	;}

	;ngrams := RunCrazyNGrams("3,4,5", true)
	;clean := RemoveDoublettes(medicalWords, othersWords)
	;medicalWords 	:= clean.objA
	;othersWords  	:= clean.objB
	;gosub AlbisPDFOdner
	gosub BefundOrdner
	;ReIndexMedicalList()

	DicGui()
	gosub CheckSyntax

return

AlbisPDFOdner:	; Albis PDF Ordner , PDF -> Text	;{

		noOCRspath := adm.DBPath "\Dictionary\albisDir_NoOCR_PDF.json"
		OCRspath		:= adm.DBPath "\Dictionary\albisDir_OCR_PDF.json"

		;~ If FileExist(noOCRspath)
			;~ FilesNT := JSONData.Load(noOCRspath,, "UTF-8")
		;~ else
			;~ FilesNT	:= Array()

		If FileExist(OCRspath)
			FilesT 	:= JSONData.Load(OCRspath,, "UTF-8")
		else
			FilesT	:= Array()

	return

		;~ FilesT:= Array()
		;~ For oidx, ofile in FilesO {
			;~ FileFound := false
			;~ For pidx, pfile in FilesP
			;~ If (pfile = ofile) {
				;~ FileFound := true
				;~ break
			;~ }
			;~ If !FileFound
				;~ FilesT.Push(ofile)

			;~ If Mod(oidx, 100) = 0
				;~ ToolTip, % "erstelle Dateiliste`n" oidx "/" FilesO.Count() "`nDateien aufgenommen`n" FilesT.Count()

		;~ }
		;~ FilesNT := FilesO := ""
		;~ ToolTip


		FilesT		:= Array()
		startT       	:= A_TickCount
		Loop, Files, % "M:\albiswin\Briefe\*.pdf", R
		{

			If Mod(A_Index, 100) = 0 {
				t 	:= Floor((A_TickCount - startT) / 1000)
				m	:= Floor(t / 60)
				s	:= t - m*60
				ToolTip, % A_Index " PDF-Dateien`nZeit: (" t ") " m ":" SubStr("0" s, -1), 1400, 1, 13
			}

			If inArr(A_LoopFileFullPath, FilesT) || inArr(A_LoopFileFullPath, FilesNT)
				continue

			If !PDFisSearchable(A_LoopFileFullPath) {
				FilesNT.Push(A_LoopFileFullPath)
				continue
			}

			FileFound := false
			For fidx, filepath in FilesP
				If (filepath = A_LoopFileFullPath) {
					FileFound := true
					break
				}
			If FileFound
				continue

			FilesT.Push(A_LoopFileFullPath)

		}

		ToolTip,,,,13
		JSONData.Save(adm.DBPath "\sonstiges\albisDir_NoOCR_PDF.json", FilesNT, true,, 1, "UTF-8")
		JSONData.Save(adm.DBPath "\sonstiges\albisDir_OCR_PDF.json", FilesT, true,, 1, "UTF-8")

return 	;}

BefundOrdner:	; nicht bearbeitete Textdateien zusammenstellen	;{

		SciTEOutput("FilesP: " FilesP.Count())

		FilesT		:= Array()
		Loop, Files, % "M:\Befunde\Text\*.txt", R
		{

			FileFound := false
			;~ If !PDFisSearchable(A_LoopFileFullPath)
				;~ continue

			For fidx, filepath in FilesP
				If (filepath = A_LoopFileFullPath) {
					FileFound := true
					break
				}
			If FileFound
				continue

			FilesT.Push(A_LoopFileFullPath)

		}
return 	;}

SaveData1: ;{
SaveData2:

	;RunCrazyNGrams("3,4", true)
	JSONData.Save(medicalWords	, Medical	, true,, 1, "UTF-8")
	JSONData.Save(othersWords 	, Others	, true,, 1, "UTF-8")
	JSONData.Save(FilesProcessed	, FilesP   	, true,, 1, "UTF-8")
	JSONData.Save(OCRspath    	, FilesT   	, true,, 1, "UTF-8")


ExitApp ;}

CheckSyntax: ;{

	WinGetPos, x, y, w, h, % "Wörterbucheditor ahk_class AutoHotkeyGUI"
	Gui, ngram: new, -DPIScale
	Gui, ngram: Font, s11 q5 cBlack, Arial
	Gui, ngram: Add, Edit, xm ym    	w500 h600 	vngOut
	Gui, ngram: Add, Edit, xm y+10 	w500 r1    	vngInput gngInputValidate -AltSubmit
	Gui, ngram: Show, % "x" x+w+5 " y" y " AutoSize NA", ngram-Syntaxchecker

	Hotkey, IfWinActive, ngram-Syntaxchecker ahk_class AutoHotkeyGUI
	Hotkey, ~Enter, ngramSyntaxValidator
	Hotkey, IfWinActive


return ;}
ngramSyntaxValidator: ;{

	Gui, ngram: Submit, NoHide
	out := ngInput "`n"
	Loop 3 {

		NgramSize := 2+A_Index
		Loop, % StrLen(ngInput) - (NgramSize - 1)  {

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
ngInputValidate: ;{

	Gui, ngram: Submit, NoHide
	If RegExMatch(ngInput, "i)[^a-zäöüß]") {
		If !lastOK
			return
		lastOK := false
		Gui, ngram: Color, cFF4444
		Sleep 150
		GuiControl, ngram:, ngInput, % RegExReplace(ngInput, "i)[^a-zäöüß]")
	} else {
		Gui, ngram: Color, cFFFFFF
		lastOK := true

	}

return
ngramGuiClose:
ngramGuiEscape:
ExitApp ;}

DicGui() {

	global
	static 	FNR := 1, DicInit := true

	col1W1	:= 280, col1W2 := 60
	col2W1	:= 80, col2W2 := 270, col2W3 := 40
	Lv1W  	:= col1W1 + col1W2 + 40
	Lv2W  	:= col2W1 + col2W2 + col2W3 + 20

	Gui, Dic: new, -DPIScale
	Gui, Dic: Font, s11 q5 cBlack, Arial
	Gui, Dic: Add, ListView	, % "xm 	y18 	w" Lv1W " r35 AltSubmit vDIC_LVTXT1 	gDIC_LV"	, Wort|GZ
	Gui, Dic: Add, ListView	, % "x+0 	    	w" Lv1W " r35 AltSubmit vDIC_LVTXT2 	gDIC_LV"	, Wort|GZ
	Gui, Dic: Add, ListView	, % "x+0 	    	w" Lv2W " r35 AltSubmit vDIC_LVTXT3 	gDIC_LV"	, WBuch|Textwort|Z
	Gui, Dic: Add, ListView	, % "x+5 		 	w" Lv2W " r35 AltSubmit vDIC_LVDIC  	gDIC_LV"	, WBuch|Med.Wörterbuch|Z

	GuiControlGet, cp2,	 Dic: Pos, DIC_LVTXT2
	GuiControlGet, cp3, 	 Dic: Pos, DIC_LVTXT3
	Gui, Dic: Font, s9 q5 cBlack, Arial
	Gui, Dic: Add, Text		, % "xm+20 y0 "	     		                        		    							, % "Medizin"
	Gui, Dic: Add, Text		, % "x" cp2X+20 " y0 "	  		                       		    							, % "andere"
	Gui, Dic: Add, Text		, % "x" cp3X+20 " y0 "	  		                       		    							, % "unbekannt/ngram"

	GuiControlGet, cp, DIC: pos, DIC_LVDIC
	Gui, Dic: Font, s10 q5 cBlack, Arial
	Gui, Dic: Add, Text		, % "x" cpX "	y0"				                        									, % "Datei:"
	Gui, Dic: Add, Text		, % "x+5 	Center	vDIC_WD"			                                 				, % "[" FNR "/" FilesT.MaxIndex() "] [0000 Worte] "
	Gui, Dic: Add, Text		, % "x" cpX " y" cpY+cpH+2 " Center vDIC_WB"			       				, % "[0000 Worte] "

	; Statistik
	Gui, Dic: Add, Text		, % "x+55 w300 Center"				   									 			, % "Gesamtstatistik"
	Gui, Dic: Font, s8 q5 cBlack, Arial
	Gui, Dic: Add, Edit		, % "Y+5   w300  h120 vDIC_GS ReadOnly"									, % "-------"

	GuiControlGet, cp, DIC: pos, DIC_WD
	Gui, Dic: Font, s8 q5 cBlack, Arial
	Gui, Dic: Add, Text		, % "xm y" cp2Y+cp2H+5 " w500 	vDIC_FN"		    						, % FilesT[FNR]

	Y := cp2Y + cp2H + 20
	Gui, Dic: Font, s10 q5 cBlack, Arial
	Gui, Dic: Add, Button 	, % "xm     y" Y      "	vDIC_BBK    	gDIC_BTNS"	                         	, % "<<"
	Gui, Dic: Add, Button 	, % "x+10              	vDIC_BFW   	gDIC_BTNS"	                         	, % ">>"

	GuiControlGet, cp, DIC: pos, DIC_WD
	Gui, Dic: Add, Button 	, % "xm     y+10       	vDIC_BOK	gDIC_BTNS hwndDIC_hOK"		, Übernehmen
	Gui, Dic: Add, Button 	, % "x+10 	     		vDIC_BCL 	gDIC_BTNS"	                         	, Abbrechen + Sichern
	Gui, Dic: Add, Button 	, % "x+100 	    		vDIC_BXT  	gDIC_BTNS"	                        	, Beenden

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
		LoadText()
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
	;~ WinActivate, Wörterbucheditor ahk_class AutoHotkeyGUI
	;~ addToDictionary()
	;~ WBStats()
	;~ LoadText()
return ;}



DIC_BTNS: ;{

	If (A_GuiControl = "DIC_BOK")  {
		addToDictionary()
		WBStats()
		LoadText()
	} else if (A_GuiControl = "DIC_BXT") {
		Gui, Dic: Destroy
		gosub SaveData2
		ExitApp
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
	global medicalWords, othersWords, FilesProcessed
	global Medical, Others, FilesP, FilesT

	; Medizinwörterbuch ergänzen
		Gui, Dic: ListView, DIC_LVDIC
		Loop, % LV_GetCount() {

			rword := [], row := A_Index
			Loop 3 {
				LV_GetText(rtxt, row, A_Index)
				rword[A_Index] := rtxt
			}

			Medical[rword.2] := !Medical.haskey(rword.2) ? rword.3 : Medical[rword.2] += rword.3
			If Others.haskey[rword.2] {
				Medical[rword.2] += Others[rword.2]
				Others.Delete(rword.2)
			}

			;~ For medWord, mWCount in Others
				;~ If (medword = rword.2) {
					;~ Others.Delete(medword)
					;~ break
				;~ }

		}
		LV_Delete()

	; Others Wörterbuch ergänzen
		Gui, Dic: ListView, DIC_LVTXT2
		Loop, % LV_GetCount() {

			rword := [], row := A_Index
			Loop 3 {
				LV_GetText(rtxt, row, A_Index)
				rword[A_Index] := rtxt
			}

			Others[rword.2] := !Others.haskey(rword.2) ? 1 : Others[rword.2] += 1

		}
		Loop 3 {
			Gui, Dic: ListView, % "DIC_LVTXT" A_Index
			LV_Delete()
		}
		Gui, Dic: ListView, % "DIC_LVDIC"
		LV_Delete()

	GuiControl, Dic:, DIC_WB, % "[           ]"

	FilesP.Push(popfile)

	;JSONData.Save(medicalWords	, Medical	, true,, 1, "UTF-8")
	;JSONData.Save(othersWords 	, Others	, true,, 1, "UTF-8")
	;JSONData.Save(FilesProcessed	, FilesP   	, true,, 1, "UTF-8")

}

LoadText() {

	global Dic, DIC_WD, DIC_FN, DIC_LVTXT, FNR, FilesT, popfile

	; Ende wenn keine Dateien mehr zu bearbeiten sind
		If (FilesT.Count() = 0) {
			MsgBox, Keine Dateien mehr zum Bearbeiten vorhanden
			return
		}

	; letzte Datei wird geladen und unter Processed gespeichert
		popfile := FilesT.pop()
		FNR++
		GuiControl, Dic:, DIC_WD	, % "[" FNR "/" FilesT.MaxIndex() "] [" words.MaxIndex() " Worte] "
		GuiControl, Dic:, DIC_FN	, % popfile

	; Worte werden untersucht
		words := CollectWords(popfile)
		If (words.Count() = 0)
			LoadText()

	; Worte anzeigen
		For word, wdata in words {
			col1 := ( wdata.list=1 ? "unbekannt"
						: wdata.list=2 ? "Medizin"
						: wdata.list=3 ? "andere"
											  : "ngram")  ; 4 = ngram

			Gui, Dic: ListView, % "DIC_LVTxt" ( wdata.list=1 || wdata.list=4 ? "3" :  wdata.list=2 ? "1" : "2")
			If (wdata.list=1 || wdata.list=4)
				LV_Add("", col1, word, wdata.count)
			else
				LV_Add("",  word, wdata.listcount) ;, wdata.listcount)
		}

		Gui, Dic: ListView, % "DIC_LVTxt1"
		LV_ModifyCol(1, "SortDesc")
		Gui, Dic: ListView, % "DIC_LVTxt2"
		LV_ModifyCol(1, "SortDesc")
		Gui, Dic: ListView, % "DIC_LVTxt3"
		LV_ModifyCol(1, "SortDesc")

		GuiControl, Dic: Focus, DIC_LVTXT3

return

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

ReIndexMedicalList() {

	For word, wcount in Medical
		Medical[word] := 0

	For fidx, filepath in FilesP {

		ToolTip, % fidx "/" FilesP.MaxIndex() "`n" filepath, 4200, 50
		text           	:= FileOpen(filepath, "r", "UTF-8").Read()
		spos          	:= 1

		while (spos := RegExMatch(text, "\s(?<Name>[A-ZÄÖÜ][\pL]+(\-[A-ZÄÖÜ][\pL]+)*)", P, spos)) {
			spos += StrLen(PName)
			If Medical.haskey(PName)
				Medical[PName] += 1
		}

	}

	JSONData.Save(medicalWords	, Medical	, true,, 1, "UTF-8")

}

CollectWords(filepath, debug:=false) {

	; phlie 				phie
	; formuflar 		formular
	; Aiforderung	Anforderung
	; Aipträumen		Alpträumen
	; Akkordiohn		Akkordlohn
	; Akkurdlohn		Akkordlohn
	; Aklenzeichen	Aktenzeichen

	vouwels := "aeiouäöü"
	ocrcorrection := {"rt":"h", "ri":"n", "x":"x", "x":"x", "x":"x"}

	PNames    	:= Object()
	MFreqLimit	:= 1
	OFreqLimit 	:= 2
	spos          	:= 1

	If RegExMatch(filePath, "\.txt$")
		text          	:= FileOpen(filepath, "r", "UTF-8").Read()
	else if RegExMatch(filePath, "\.pdf$") {
		text       	:= IFilter(filepath)
		If debug
			SciTEOutput(filepath ": " StrLen(text))
	}

	text := RegExReplace(text, "([a-zäöüß])([A-ZÄÖÜ][a-zäöü])", "$1 $2")
	text := RegExReplace(text, "[\n\r\f]", " ")

	tLen := StrLen(text)

	while (spos := RegExMatch(text, "\s(?<Name>[A-ZÄÖÜ][\pL]+(\-[A-ZÄÖÜ][\pL]+)*)", P, spos)) {

			takeword := 1
			ToolTip, % "Position: " spos "/" tLen, 900, 1
			spos += StrLen(PName)

		; wenn ein Wort in einem der beiden Wörterbücher vorkommt
			if Medical.haskey(PName) {
				Medical[PName] += 1
				takeword := (Medical[PName] < MFreqLimit) ? 0 : 2
			}
			else if Others.haskey(PName) {
				Others[PName] += 1
				takeword := (Others[PName] <= OFreqLimit) ? 3 : ngramWordCheck(PName, 5) ? 4 : 3
			}

		; ausgeschlossene Buchstabenkombinationen - kommen gleich ins Others Wörterbuch
			If RegExMatch(PName, "[a-zäöüß\-][A-ZÄÖÜ]{2,}") {
				Others[PName] := !Others.haskey[PName] ? 1 : Others[PName]+1
				continue
			}

		;~ ; ngram Evaluierung unbekannter Wörter
			;~ If (takeword = 99)
				;~ takeword := 1			 ; bisher unbekannt
				;~ If ngramWordCheck(PName, 5)
					;~ takeword := 1            ; ngrams haben nicht funktioniert
				;~ else

		tw3:= takeword
		; takeword 1 = unbekannt, 2 = andere
			;~ If (takeword < 99)
			If !PNames.haskey(PName)
				PNames[PName] := {"count"    	: 1
											, 	 "list"      	: takeword
											, 	 "listcount"	: (takeword=2 ? Medical[PName]
																:  takeword=3 ? Others[PName] : 0)}
			else
				PNames[PName].count += 1


	}

	;SciTEOutput(" - - - - - - - - - - - - - - - - - - - - - - - - - -`n")

return PNames
}

ngramWordCheck(word, depth:=5) {

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

ngramSyntaxValidate(word, NgramSize, updateNgrams:=false) {

	global ngrams
	If !IsObject(ngrams) || updateNgrams
		ngrams := RunCrazyNGrams("3,4,5", false)

	For idx, sword in StrSplit(word, "-")
		Loop, % StrLen(sword) - (NgramSize-1)
			If !ngrams[NgramSize].haskey( SubStr(sword, A_Index, NgramSize) )
				return false

return true
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

IFilter(file, searchstring:="") {                                                                                 	;-- Textsuche/Textextraktion in PDF/Word Dokumenten

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

	if (!file)
		return -99

	If !PDFisSearchable(file) {
		MsgBox, PDF ist nicht durchsuchbar
		return
	}

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
				text := StrGet(&buf,, "UTF-16")
				if (resultstriplinebreaks)
					text := StrReplace(text, "`r`n")
				If searchstring && (strpos := RegExMatch(text, searchstring))
					break

			}
		}
	}

	ObjRelease(PersistStream)
	ObjRelease(iFilter)
	if (job)
		DllCall("CloseHandle", "Ptr", job)

	If (strpos > 0)
		return strpos

return searchstring ? "" : Text
}

PDFisSearchable(pdfFilePath)	{                                                                              	;-- durchsuchbare PDF Datei?

	; letzte Änderung 08.02.2021 : neuer Matchstring

	If !(fobj := FileOpen(pdfFilePath, "r", "CP1252"))
		return 0

	while !fobj.AtEof {
		line := fobj.ReadLine()
		If RegExMatch(line, "i)(Font|ToUnicode|Italic)") {
			filepos := fobj.pos
			fobj.Close()
			return filepos
		}
		else If RegExMatch(line, "Length\s(?<seek>\d+)", file)     	; binären Inhalt überspringen
			fobj.seek(fileseek, 1)                                                   	; •1 (SEEK_CUR): Current position of the file pointer.
	}

	fobj.Close()

return 0
}

inArr(str, arr) {
  For k,v in arr
    If (str=v)
      return true
return false
}

RemoveDoublettes(objA, objB) {                                                                             	;-- d

	removed := 0

	For word, wcount in objA {
		ToolTip, % "objB:      " objB.Count() "`nremoved: " removed "`nobjA:      " objA.Count()
		If objB.haskey(word) {
			objB.Delete(word)
			removed ++
		}
	}

return {"objA": objA, "objB":objB}
}

return


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
		props             	:= Object()
		props.Dir         	:= AddendumDir
		props.Ini          	:= props.Dir "\Addendum.ini"

	; Log Dateipfad und Datenbankverzeichnis
		workini := IniReadExt(props.Ini)
		props.LogPath	:= IniReadExt("Addendum", "AddendumLogPath"	,, true)
		props.DBPath  	:= IniReadExt("Addendum", "AddendumDBPath" 	,, true)

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



#include %A_ScriptDir%\..\..\lib\class_JSON.ahk
#include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
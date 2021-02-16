; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . .                                                                                                                                                                                                                                   	. . . . . . . . . .
; . . . . . . . . . .                                                                                  	  ADDENDUM AUTO-ICD-DIAGNOSEN                                                                                     	. . . . . . . . . .
; . . . . . . . . . .                                                                                              letzte Änderung 26.09.2019                                                                                           	. . . . . . . . . .
; . . . . . . . . . .                                                                                                                                                                                                                                   	. . . . . . . . . .
; . . . . . . . . . .                                                  - effektive Eingabe von Diagnosen in Albis on Windows durch Autohotkey Hotstrings                                                     	. . . . . . . . . .
; . . . . . . . . . .                                                  - Hotstrings sind so kurz gehalten das diese nicht doppelt vorhanden sind                                                                     	. . . . . . . . . .
; . . . . . . . . . .                                                  - keine Verwendung kryptischer Kürzel                                                                                                                         	. . . . . . . . . .
; . . . . . . . . . .                                                  - sofortige ICD-Diagnose Expandierung nach Eingabe der passenden Buchstaben                                                         	. . . . . . . . . .
; . . . . . . . . . .                                                  - Abfrage der Seitenlokalisation bei entsprechenden Diagnosen                                                                                    	. . . . . . . . . .
; . . . . . . . . . .                                                  - Mehrfachauswahl von Diagnosen möglich                                                                                                                 	. . . . . . . . . .
; . . . . . . . . . .                                                  - Diagnosenketten z.B. bei Diabetes mit Folgekomplikationen                                                                                       	. . . . . . . . . .
; . . . . . . . . . .                                                  - Synonyme für Diagnosenerweiterung - Fachsprache und Laienbegriffe                                                                        	. . . . . . . . . .
; . . . . . . . . . .                                                  - automatisches Eintragen von Leistungskomplexen bei bestimmten Diagnosen                                                             	. . . . . . . . . .
; . . . . . . . . . .                                                  - integrierte Oberfläche zum komfortablen Anlegen eigener Hotstrings                                                                            	. . . . . . . . . .
; . . . . . . . . . .                                                                                                                                                                                                                                   	. . . . . . . . . .
; . . . . . . . . . .                                                  *****************************************************************************************                                                     	. . . . . . . . . .
; . . . . . . . . . .                                                 *         Achtung: dieses Skript benötigt die SetWinEventHooks aus Addendum.ahk  	*                                                   	. . . . . . . . . .
; . . . . . . . . . .                                                 *                           und Funktionen aus Addendum_Functions.ahk                         	*                                                   	. . . . . . . . . .
; . . . . . . . . . .                                                  *****************************************************************************************                                                     	. . . . . . . . . .
; . . . . . . . . . .                                                                                                                                                                                                                                   	. . . . . . . . . .
; . . . . . . . . . .                                                                                                                                                                                                                                   	. . . . . . . . . .
; . . . . . . . . . .                                        ROBOTIC PROCESS AUTOMATION FOR THE GERMAN MEDICAL SOFTWARE "ALBIS ON WINDOWS"                                    	. . . . . . . . . .
; . . . . . . . . . .                                               BY IXIKO STARTED IN SEPTEMBER 2017 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE                                             	. . . . . . . . . .
; . . . . . . . . . .                                          	 RUNS WITH AUTOHOTKEY_H  AND  AUTOHOTKEY_L  IN  32 OR 64 BIT  UNICODE VERSION                                          	. . . . . . . . . .
; . . . . . . . . . .                                                                     THIS SOFTWARE IS ONLY AVAIBLE IN GERMAN LANGUAGE                                                                    	. . . . . . . . . .
; . . . . . . . . . .                                                                                                                                                                                                                                   	. . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
;RegEx AutoDiagnose einsetzen : (\:R\*\:)(\w*)(\:\:)(.*)(\};) -> :XR*:\2::AutoICD("\4}", "")
^#!s::
HotstringStatitistik(AddendumDir "\Module\Addendum\Include\ICD-Hotstrings.ahk")
return
:XR*:.ADL::SciteAutoDiagnose(Clipboard)

#include %A_LineFile%\..\ICD-Hotstrings.ahk

;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; SENDEN VON ICD-DIAGNOSEN AN ALBIS
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
AutoICD(DX, Options := "") {                                                                                                 	;-- Auto-ICD-Diagnosen

	; Options: S	- Seitenangabe zur Diagnose notwendig
	;           	M	- Mehrfachauswahl von Diagnosen möglich

	/* todo
			4. Hotstring Gui - zum nachträglichen Editieren erstellen!
			7. evtl.RegEx Hotstrings - dynamic Hotstrings implementieren
	*/

	; Variablen, Liste und Optionen sichern
	; ------------------------------------------------------------------------ ;{
		global hACB, hACBLb, ACB, ACBLb
		static DiaTmp, DiaAnzahl, goLeft, tDX, Seite, Multi, lko, Ziffern, Freitext, Ftextc, lastPos
		static FilterWords, lastwords, aT, GuiRun
		static oACB := Object()

	; kurz die Tastatureingabe blockieren damit der User nach Hotstringauslösung nicht sofort weiter schreiben kann
		;BlockInputShort(500)

	; Verzweigung je nach Inhalt des DX Parameter
		If IsObject(DX)
		{
				AutoPrediction(DX, Options)
		}
		else If !RegExMatch(DX, "\{.*\}")	; Filterworte - alles ohne ICD Code
		{
				GuiRun         	:= 1
				FilterWords   	.= Options "|"							; z.B. :XR*::inkompl:AutoICD("inkompletter", "inkomplett|unvollständig") - die Worte in Options sollten Wortstämme sein
				lastwords      	.= DX "|"
				ADInterception	:= true
				SendRaw, % DX
				Send, {Space}
				return
		}
		else
		{
				GuiRun         	:= 1
				tDX	            	:= DX
				Freitext				:= InStr(DX, "###", "", 1, Ftextc) > 0 ? true : false
				Seite	            	:= InStr(Options, "S")  ? true : false
				Multi	            	:= InStr(Options, "M") ? true : false
				lko	            	:= RegExMatch(Options, "^(M|S)*(M|S)*L") ? true : false
				If lko
					RegExMatch(Options, "(L)[\d|]+", Ziffern)

				ADInterception	:= false                                      	; eventuell muss diese flag noch anders zurückgesetzt werden können
		}

	; zuvor eingegebene Worte aus dem RichEdit Control entfernen
		If StrLen(lastwords) > 0
			RemoveLastWords(lastwords)
		else
			lastwords:= ""

	;}

	; ------------------------------------------------------------------------
	; Zeilen für die Ausgabe zählen - maximal 10 gleichzeitig!
	; ------------------------------------------------------------------------ ;{
		If !InStr(DX, "|")
		{
				Diagnose:= DX
				gosub Seitenangabe
				return
		}
		else
		{
			;- Erstellen des Listboxinhaltes, ist die Liste leer, dann gibt es die Diagnose für diese Wortkombination nicht oder noch nicht
				PrevDia  	 := FilterDiagnosen(DX, FilterWords)
				If !PrevDia
					PrevDia := FilterDiagnosen(DX, "")

			;- zählt die Anzahl der Diagnosen für die Listbox
				Loop, Parse, PrevDia, `|
					If A_Index < 11
						rows:= A_Index
					else
						break

			;- Filterliste und letzte Worte leeren
				FilterWords:= ""

			;- Gui aufrufen
				GuiRun := 1
				TTipC(A_ThisHotkey ", `"" SubStr(PrevDia, 1, 50) "`", r: " rows ", m: " Multi ", ac: " activeControl.hwnd, "SP1")
				AutoICDCompleteGui(PrevDia, rows,  "x" A_CaretX " y" A_CaretY, Multi, activeControl.hwnd)
				Multi:= false
		}
	;}

return

	; ------------------------------------------------------------------------
	; Ausgabe der Diagnosen nach Albis
	; ------------------------------------------------------------------------ ;{
	  AutoCBLBres:
	; ausgewählten Eintrag oder Einträge aus der Liste heraussuchen
	; ------------------------------------------------------------------------ ;{
		Gui, ACB: Submit
		If GuiRun = 1 	;1.Durchlauf: Eintragen der Diagnose
		{
				ACBLb      	:= RegExReplace(ACBLb, "\s*\(.*\)")
				diaWahl    	:= StrSplit(ACBLb, "|")
				diaAuswahl	:= StrSplit(tDX, "|")
				Diagnose  	:= ""

				For Each, Item in diaAuswahl
					Loop % diaWahl.MaxIndex()
						If InStr(Item, diaWahl[A_Index]) = 1
							If RegExMatch(A_LoopField, "(?<=\>).*", str)
								Diagnose .= RegExReplace(str, "\}\s*\;\s*", "}; ")
							else
								Diagnose .= RegExReplace(Item, "\}\s*\;\s*", "}; ")
		}
		else if GuiRun = 2 	;2. Ablauf: Eintragen der Seitenangabe
		{
				Diagnose	:= ACBLb
				lastpos     	:= 0
		}

	;}

	;------------------------------------------------------------------------
	  Seitenangabe:
	; Eingabefokus zurückgeben ins RichEdit-Control von Albis
	; ------------------------------------------------------------------------ ;{
		AlbisActivate(1)
		ControlFocus,, % "ahk_id " activeControl.hwnd
		If WinExist("AutoDiagnose")
			Gui, ACB: Destroy
	;}

	; ------------------------------------------------------------------------
	; Seitenangabe ergänzen lassen, falls notwendig
	; ------------------------------------------------------------------------ ;{
		If Seite
		{
			; auf false setzen damit dieser Abschnitt beim 2.Durchlauf nicht aufgerufen wird
				Seite 	:= false
				GuiRun := 2
				startpos	:= DiaAnzahl:= 0		; mit 0 initialisieren

			; sucht die letzte Diagnose in einer Diagnosereihe heraus und fragt dann dort nur nach der Seite
				Loop
					If startpos:= RegExMatch(Diagnose, "O)([G|A|V|Z|\s]\})", "", startpos +1)
						lastpos:= startpos, DiaAnzahl:= A_Index
					else
						break

			; temporär ein 'L" an die gefundene Position einfügen und den Diagnosentext senden
				DiaTmp := Diagnose := Trim(SubStr(Diagnose, 1, lastpos - 1) "L" SubStr(Diagnose, lastpos, StrLen(Diagnose)))
				SendRaw, % Diagnose:= RegExReplace(Diagnose, ";\s*$", "; ")

			; den Cursor zu dieser Position bringen, die Caret-Position speichern und dann erst das 'L' selektieren
				MoveCursorSelect("Left " (goLeft:= StrLen(Diagnose) - lastpos), "Left 1")

			; Gui für die Auswahl der Seitenangabe erstellen
				AutoICDCompleteGui("L|R|B| ", 4, "x" A_CaretX " y" A_CaretY , 0, activeControl.hwnd)
				return
		}

		if StrLen(Diagnose) > 1 			;  !Seite &&
		{
				DiaTmp:= Diagnose:= RegExReplace(Diagnose, "\};*\s*$", "}; ")
				SendRaw, % Diagnose
		}
		else If StrLen(Diagnose) = 1
		{
				; setzt bei allen Diagnosen in einer Diagnosenkette die Seitenangabe hinzu
					DiaSend := ""
					Loop, Parse, DiaTmp, `}, A_Space
						DiaSend .= RegExReplace(A_LoopField, "[RLBG\s]*([GAVZ]\})", Diagnose "$1; ")

				; Text im Albiscontrol ändern
					CText := Trim(StrReplace(ControlGetText("", "ahk_id " activeControl.hwnd), DiaTmp, DiaSend)) " "
					VerifiedSetText("", CText, activeControl.hwnd)

				; Cursor an die letzte Eingabeposition setzen
					SetKeyDelay, 0, 0
					Send, % "{Right " (goRight:= StrLen(CText) - InStr(CText, DiaSend) + 1) "}"
					SetKeyDelay, 20, 50

					DiaTmp:=DiaSend:= CText:= ""
		}

		If (sPos:= InStr(Diagnose, "###"))
		{
				MoveCursorSelect("Left " (goLeft:= StrLen(Diagnose) - sPos - 2), "Left 3")
				r:= GetWindowSpot(activeControl.hwnd)
				ToolTip, % "Bitte den Freitext ergänzen!", % A_CaretX - 40, % r.Y - 25, 5
				SetTimer, TTOff, -3000
		}

		If lko
		{
				Ziffern:= StrSplit(Ziffern, "|")
				Loop % Ziffern.MaxIndex()
					ziffmsg .= A_Index
				PraxTT ("Zu der Diagnose: " Diagnose "`nkönnen folgende Leistungskomplex eingetragen werden.`n`n" ziffmsg "`nAuswahl per Zifferntasten solange dieser Dialog zu sehen ist.", "10 5")
				lko := false
		}

	;}

return ;}

ACBListbox:                                                                       	;{
	if (A_GuiEvent="DoubleClick")
		gosub AutoCBLBres
return ;}

TTOff:                                                                               	;{
	ToolTip,,,, 5
return ;}
}

FilterDiagnosen(DX, Filter:="") {                                                                                             	;-- entfernt Diagnosen aus der Diagnosenliste welche nicht die Filterworte enthalten

	; entfernt alle Diagnosen aus der Diagnosenliste welche nicht die Filterworte enthalten
		If (Filter <> "")
		{
				Filter:= StrSplit(RTrim(Filter, "|"), "|")

				Loop, Parse, DX, `|
					Loop, % Filter.MaxIndex()
						If InStr(A_LoopField, Filter[A_Index])
						{
								tmpDX.= A_LoopField "|"
								continue
						}

				DX := RTrim(tmpDX, "|")
		}

		Loop, Parse, DX, `|
			If InStr(A_LoopField, ">")
			{
					RegExMatch(A_LoopField, "(?<=\>)(.*)", TmpDiaListe)
					RegExMatch(A_LoopField, "^.*(?=\>)", TmpPrevDia)
					PrevDia .= TmpPrevDia " (Diagnosenreihe)|"
			}
			else
					PrevDia .= A_LoopField "|"

return RTrim(PrevDia, "|")
}

AutoPrediction(DX, trunk) {                                                                                                     	;-- die smartere Auto-ICD Funktion soll das werden

	;For hotstring in DX[trunk]
	;	Hotstring()

}

AutoICDCompleteGui(Content, rows, GuiOptions, Multi:=false, hfocus:=0                              	;-- zeigt mehrere Diagnosen oder Anderes zur Auswahl an
, FontOptions:= "s11 Normal q5", FontName:= "Futura Bk Bt") {

		global hACB, hACBLb, ACB, ACBLb
		static DpiF
		static	MultiInfo:= "mehrere Diagnosen auswählbar (Strg + Mausklick)"

		If !dpiF
			dpiF:= screenDims().DPI / 96

	; ------------------------------------------------------------------------
	; AutoComplete Gui erstellen
	; ------------------------------------------------------------------------ ;{
		guiWidth1	:= LBEX_CalcIdealWidthEx(0, MultiInfo "|" MultiInfo, "|", FontOptions, FontName, 1)
		guiWidth	:= LBEX_CalcIdealWidthEx(0, Content, "|", FontOptions, FontName, rows)
		guiWidth	:= guiWidth1 > guiWidth ? guiWidth1 : guiWidth

		Gui, ACB: New 	, -SysMenu -Caption +AlwaysonTop +ToolWindow +HWNDhACB ;0x98200000
		Gui, ACB: Margin	, 0, 0

		If Multi
		{
				Gui, ACB: Color, c172842
				RegExMatch(FontOptions, "(?<=s)\d+", fsize)
				smallFont:= RegExReplace(FontOptions, "(?<=s)\d+", fsize - 1)
				Gui, ACB: Font 	, % smallFont , % FontName
				Gui, ACB: Add	, Text, % "w" guiWidth " HWNDhTACB cWhite CENTER", mehrere Diagnosen auswählbar (Strg + Mausklick)
				t:= GetWindowSpot(hTACB)
		}

		Gui, ACB: Font  	, % FontOptions " cBlack", % FontName
		Gui, ACB: Add  	, ListBox, % (!Multi ? "" : "xm y" T.H+2 " ") "Choose1 HWNDhACBLb vACBLb gACBListbox" (!Multi ? "" : " Multi")  " r" rows " w" guiWidth, % Content	;+0x8
		Gui, ACB: Show	, % GuiOptions " NA Hide", AutoDiagnose

	; Gui positionieren
		a:= GetWindowSpot(hACB), c:= GetWindowSpot(hfocus)
		aX:= A_CaretX

		If (aX + a.W > A_ScreenWidth)
			aX:= c.X

		If InStr(Content, "L|R|B")
			aY:= c.Y + c.H - a.H
		else
			aY:= c.Y - a.H

		Gui, ACB: Show, % "x" aX " y" A_CaretY-20 " NA"

	;}

		If Multi
			GuiControl, ACB: Focus, ACBLb
		else
		{
			WinActivate, ahk_class OptoAppClass
			ControlFocus,, % "ahk_id  " activeControl.hwnd
		}
		AutoICDCompleteGuiHotkeys("On")

return hACB

ACBMoveDown: ;{
	ControlSend,, {Down}, % "ahk_id " hACBLb
return
ACBMoveUp:
	ControlSend,, {Up}, % "ahk_id " hACBLb
return ;}
ACBCheck: ;{
	MouseGetPos, mx, my, hWin
	If (hWin = hACB)
	{
		MouseClick, Left, % mx, % my
		return
	}
;}
ACBGuiClose: ;{
		If InStr(A_ThisHotkey, "RButton")
			Send, {RButton}
		Gui, ACB: Destroy
		AutoICDCompleteGuiHotkeys("Off")
		hACB:= hACBLb:= ""
return ;}

}

AutoICDCompleteGuiHotkeys(status:="On") {                                                                         	;-- Hotkeys für die Nutzereingaben im AutoICDCompleteGui

	Hotkey, IfWinExist	, AutoDiagnose ahk_class AutoHotkeyGUI
	Hotkey, Enter     	, AutoCBLBres   	, % status
	Hotkey, Down   	, ACBMoveDown	, % status
	Hotkey, Up    		, ACBMoveUp   	, % status
	Hotkey, Esc	    	, ACBGuiClose  	, % status
	Hotkey, Left	    	, ACBGuiClose  	, % status
	Hotkey, Right	    	, ACBGuiClose   	, % status
	Hotkey, LButton  	, ACBCheck       	, % status
	Hotkey, RButton  	, ACBGuiClose   	, % status
	Hotkey, IfWinExist

}

DiagnoseUndZiffer(Diagnose, Ziffer, Seitenangabe) {                                                              	;-- Eintragen von Diagnose und Leistungskomplex

	SendRaw, % Diagnose
	Send, {Tab}
	SendRaw, % "lko"
	Send, {Tab}
	SendRaw, % "02300"
	Send, {Tab}
	If Seitenangabe
			Send, {ShiftUp}{Left 2}{ShiftDown}

return
}

ADControlGet(hwnd) {                                                                                                          	;-- Funktion speichert Inhalt des RichEditControls von Albis

	ControlGetText, activeControlText,, % "ahk_id " hwnd
	activeControl.diaText	:= activeControlText
	activeControl.hwnd   	:= hwnd

return 0
}

;}

;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; ERSTELLUNG EINES NEUEN HOTSTRINGS
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
SciteAutoDiagnose(str) {                                                                                                         	;-- sendet den erstellten Hotstring in eine Scintilla Fenster

	;str:= "Lungenemphysem {J43.9G}; "					; zu Testzwecken
	If !RegExMatch(str, "\{[A-Z]\d+\.*\d*")
	{
			PraxTT("Das Clipboard enthält keine ICD-Diagnose`nbitte wählen Sie den Diagnosenstring`nnoch einmal aus und kopieren diesen mit`nStrg+c ins Clipboard.", "5 5")
			return
	}
	hotstring                            	:= {"gui":{}, "alternative":{}, "suggest":{}, "synICD": [], "addICD": []}
	hotstring.suggest.words     	:= Object()
	hotstring.suggest.suggestions	:= Object()
	hotstring.suggest.conflicts   	:= Object()
	hotstring.fcontrol                	:= GetFocusedControl()
	hotstring.ACaretX                	:= A_CaretX
	hotstring.ACaretY                	:= A_CaretY
	hotstring.string                    	:= TrimDiagnose(str)
	hotstring.icd                     		:= rxMatch(hotstring.string, "\{([A-Z]\d+\.*\d*).*\}", "1")
	hotstring.StammICD         	 	:= rxMatch(hotstring.icd, "[A-Z]\d+", "")
	hotstring.HotStringName     	:= rXMatch(hotstring.string, "^\s*[\s\-\|\p{L}\.\d\w\[\]\,\!]+", "")
	hotstring.diagnose                	:= regExMatch(hotstring.string, "\{[\w\.]*\}") ? TrimDiagnose(hotstring.string) : ""
	hotstring.HotStringName     	:= RegExReplace(hotstring.HotStringName, "(\[\s*[A-Z\d\.\!\*]+\s*\])*")
	hotstring.HotStringName     	:= RegExReplace(hotstring.HotStringName, "(\[{[A-Z\d\.\!\*]+\])(\{[A-Z0-9\.\!\*]+\})")

	ADHotstring()
	ADGui()

return
}

SciteFormatAutoDiagnosen() {                                                                                                	;-- Autoindent, Autocorrections to AutoDiagnose.ahk

	q	:= Chr(0x22)
	sci:= Object()
	sci:= Sci_GetCurTextLine()
	hSci:= sci.hScintilla

	If !RegExMatch(sci.curText, "^\:[XR\*C]+\:[A-Za-z\p{L}]+\:\:")
	{
			PraxTT("Du hast keine Hotstringzeile ausgewählt!", "5 4")
			return
	}

	If RegExMatch(sci.curText, "\;\{")
	{
			PraxTT("Diese Hotstringzeile wurde schon formatiert!", "5 4")
			return
	}
	;MsgBox, % "(?<=\(\" q ")[A-Za-z\sßÄÜÖäüöÉÈéè-\|]+"
	RegExMatch(sci.curText, "(?<=\(\" q ")[A-Za-z\s-\|\p{L}\.0-9\{\}]+", HotStringName)
	HotStringName:= RegExReplace(HotStringName, "\{[A-Z0-9\.\!\*]+\}")
	If InStr(HotStringName, "|")
	{

			tmpStr:= StrSplit(HotStringName, "|"), HotStringName:= ""
			Loop, % tmpStr.MaxIndex()
			{
				first:= StrSplit(tmpStr[A_Index], " ").1
				If !InStr(HotStringName, first)
				HotStringName .= first " / "
			}
			HotStringName:= RTrim(HotStringName, " / ")
	}

		startpos:= sci.Pos
	; geht zum Ende des Hotstring
		EndofHotstring:= sci.PosFromLine + InStr(sci.curText, "::") + 1
		time:= 200
		Sci_GotoPos(hSci, EndofHotstring)
		Sleep, % time
	; verschiebt alles was nach :: kommt auf eine neue Zeile
		Send, {Enter}
		Sleep, % time
	; geht wieder zurück zum Ende des Hotstring
		Sci_GotoPos(hSci, EndofHotstring)
	; sendet die Zeichenfolge die ein Folding in Scite ermöglicht
		SendRaw, % "                                                                                                                                                                                                                                                           "
		Send, `t
		SendRaw, % "`;`{ " HotStringName
		Sleep, % time
	; sucht in den Zeilen oberhalb nach dem Vorkommen von ;`{ um dann daran die editierte Zeile auszurichten
		endline:= Sci_LineFromPos(hSci, Sci_GetTextLength(hSci))
		up:= 1
		curline := sci.curline
		while, (curline > 1) and (curline < endline)
		{
				;Send, % (up = 1) ? "{Up}" : "{Down}"
				Send, {Up}
				Sleep, % time
				sci:= Sci_GetCurTextLine()
				If (brk1:= InStr(sci.curText, "`;`{"))
				{
						go:= StrLen(sci.curText) - brk1 + 3
						Send, {End}
						Send, % "{Left " go "}"
						Sleep, % time
						Send, % "{Down " A_Index "}"
						Sleep, % time
						sci:= Sci_GetCurTextLine()
						DelX:= InStr(sci.StrR, "`;`{") - 3
						If DelX > 0
							Send, % "{BackSpace " DelX "}"
						Sleep, % time
						break
				}
				If A_Index > 10
						break
				curline:= sci.curline
		}
	; jetzt noch das return ;`} einfügen
		Send, {Down}
		Sleep, % time
		Send, {End}
		Sleep, % time
		Send, {Enter}
		Sleep, % time
		SendRaw, % "return `;`}"
		Sleep, % time
	; zurück auf die Anfangszeile und zusammenfalten
		Sci_GotoPos(hSci, startpos)
		;SendMessage, 225, % Sci_LineFromPos(hSci, startpos), 0,, ahk_id %hSci%
}

ADLoadHotstrings(AutoDiagnosenFile) {                                                                                   	;-- loads hotstrings and ICD's from my AutoDiagnose.ahk file

	StrObj	:= Object()
	StrObj	:= {"hotstring":{}, "diagnose":{}}
	        q	:= Chr(0x22)

	If !FileExist(AutoDiagnosenFile)
	{
			MsgBox, 1, Addendum für Albis on Windows, % "Die Hotstringdatei:`n" AutoDiagnosenFile "`nexistiert nicht."
			return
	}

	FileRead, file, % AutoDiagnosenFile
	Loop, Parse, file, `n, `r
		If RegExMatch(A_LoopField, "\:.*\:(.*)\:\:", str)
		{
				StrObj.hotstring.Push({"string":str1, "line": A_Index})
		}
		else if RegExMatch(A_LoopField, "^AutoICD")
		{
				hotstringFileLine:= A_Index
				RegExMatch(A_LoopField, "O)([A-Z]\d+\.*\d*)", icd)
				Loop, % icd.Count()
					StrObj.diagnose.Push({"icd":icd[A_Index], "line": hotstringFileLine})
		}

return StrObj
}

ADHotstring() {                                                                                                                      	;-- finds the shortest possible hotstring

	/* FUNKTIONSBECHREIBUNG: Berechnung der kürzesten Buchstabenanzahl für das Anlegen effizienter ICD-Hotstring

		----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
																										BENÖTIGTE DATENDATEIEN
		----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			* Diagnosen.txt und Diagnosen_Zusatz

			* der Ursprung der Dateien sind die Datendateien zur Deutschen Fassung der ICD-10 die vom DIMDI herausgegeben werden

			* Link zu den ICD-10-Dateien            	: https://www.dimdi.de/dynamic/de/klassifikationen/downloads/
			* von diesem Skript benutzte Fassung 	: ICD-10-GM/Version 2019/ICD-10-GM 2019 Alphabet EDV-Fassung TXT (CSV)

			* Der Inhalt der Datei wurde verändert	: Die erste Zeile der Datei enthält die Anzahl der Diagnosentext und welche Version der ICD-10-GM vorliegt.
															    		  Die weiteren Zeilen enthalten jeweils einen ICD-Diagnosentext der direkt in Albis nutzbar ist (Albis formatted ICD-Code).
															    		  Den Diagnosen wurde das G für (G)esichert hinzugefügt, da dies die häufigste Option ist.
															    		  Was nicht zur Darstellung als ICD-Diagnose geeignet war, wurde in die Datei "Diagnosen_Zusatz.txt" ausgelagert.
															    		  Die Dateien befinden sich im selben Verzeichnis wie diese Skriptdatei.
															    		  Der ursprüngliche Name wurde auf Diagnosen.txt geändert um bei zukünftlichen Änderungen der ICD-Fassungen
															    		  in diesem Skript keine Änderungen vornehmen zu müssen.


		----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
																												FUNKTIONSWEISE
		----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			1. Die Funktion CapitelWords (siehe oberhalb) gibt einen Array zurück, welcher nur die Worte aus dem Diagnosetext enthält, die Groß geschrieben sind.
				Die Idee dahinter war, sich das Anlegen/Nutzen eines Wörterbuches zu ersparen. Das DIMDI (Deutsche Institut für Medizinische Dokumentation und Information)
				hat z.T. aber auch Adjektive der Großschreibung zugeführt. So das ein einfaches Erkennen von Eigen-/Krankheitsnamen nicht immer möglich ist.

			2.	prüft dabei das keine Hotstrings doppelt angelegt werden
				sind alle Buch ......

*/

	; die Funktion erstellt nur autoexpandierende Hotstrings (Option: *)
	; Parameter	: "hotstring" ist ein globales Objekt und muss nicht übergeben werden
	; Rückgabe	: schreibt Daten in das globale "hotstring" - Objekt

		If InStr(Trim(hotstring.HotStringName), " ")
			hotstring.suggest.words	:= CapitelWords(hotstring.HotStringName)
		else
			hotstring.suggest.words[1]:= hotstring.HotStringName

	; lädt die bestehende Hotstrings enthaltende Datei in das globale AutoDiagosen Objekt
		AutoDiagnosen:= ADLoadHotstrings(A_ScriptDir "/include/ICD-Hotstrings.ahk")
		If !IsObject(AutoDiagnosen)
			return

	; überprüft ob für die neue Diagnose schon ein Hotstring angelegt wurde
		For Index, ICDCode in AutoDiagnosen.diagnose
			If InStr(ICDCode.icd, hotstring.icd)
			{
					PraxTT("Für die ICD-Diagnose:`n" hotstring.string "`nexistiert ein Hotstring`nin Zeile: " ICDCode.line ".", "6 5")
					return ""
			}

	; Berechnung der kürzesten Buchstabenanzahl für den neuen Hotstring
		For wIndex, word in hotstring.suggest.words
			Loop, % StrLen(word) - 4
			{
					abrv      	:= SubStr(word, 1, A_Index + 4)
					addChar	:= false

					For index, value in AutoDiagnosen.hotstring
						If RegExMatch(value.string, "^" abrv)
						{
								addChar:= true
								If StrLen(word) = StrLen(abrv)
								{
										hotstring.suggest.conflicts.Push(value.string "`t| " AutoDiagnosen.diagnose[A_Index].icd "`t| " value.line "`n")
										abrv:= hotstring.HotStringName
										addChar:= false
								}
								break
						}

					If !addChar
					{
							hotstring.suggest.suggestions.Push(abrv)
							break
					}
			}

return
}

ADGui() {                                                                                                                              	;-- shows the suggestions for hotstring abbreviations

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Variablen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		static	  vX    	:= 5                     	; x Position der Vorschläge
				, gW    	:= 1200               	; Gesamtbreite der Gui
				, vW    	:= 180                  	; Breite des Hotstring Vorschlagbereiches
				, cW    	:= 190                    	; Breite des Konfliktanzeigebereiches
				, iw     	:= 400 - 4               	; Breite de ICD 1. Listview
				, aMinH:= 400                 	; minimale Höhe des Anzeigebereiches
				, cH    	:= 150                 	; Höhe des Konfliktanzeigebereiches
				, cX  	 	:= vW + 20          	; x Position des Konfliktanzeigebereiches
				, iX      	:= cX + cW + 10  	; x Position der 1. Listview
				, zX  		:= iX + iW + 10   	; x Position der 2. Listview
				, zw     	:= gW - iX - iW - 4  	; Breite der 2. Listview
				, q    	:= Chr(0x22)          	; -> das ist ein "
				, ICDAltStatus := 0

		static	 T1, T2, T3, T4, T5, T6, T7, T8, AConflict, hAConflict, Prg1, Prg4, Prg5, hPrg4, hPrg5
				, AEdit1, hAEdit1, Achk1, Achk2, hICDAlt1, hICDAlt2, Ready, TheString, hTheString

		global AHG, LvSynICD, LvStammICD, hLvSynICD, hLvStammICD
		hotstring.warn	:= 0
		hinweis1       	:= "Hotstrings Komma getrennt eingeben, begrenzen mit * (Ausschl*ag,Dermati*tis)"
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Gui
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		Progress, P0 B2 cW202842 cBFFFFFF cTFFFFFF zH25 w400 WM400 WS500,  ....lade Daten, Addendum für AlbisOnWindows - AutoICD, Lade ICD Code, Futura Bk Bt
		Gui, AHG: new, +HWNDhAHG

	;-: Diagnosenanzeige
		Gui, AHG: Add, Progress, % "xm ym      w" gw " h35 c172842"                                                                                                       	, 100
		Gui, AHG: Font, s14 cWhite
		Gui, AHG: Add, Text  	 , % "xm ym+5 w" gW " vT1 Center BackgroundTrans"                      	                                                     	, % hotstring.string

	;-: Bezeichner - Hotstring-Vorschläge, Hotstringkonflikte, Synonym ICD's, Stamm ICD's
		GuiControlGet, t, AHG: Pos, T1
		Gui, AHG: Font, s10 Normal
		Gui, AHG: Add, Progress	, % "xm     	y" tY + tH + 10 " w" vw	" vPrg1	h25		c972842	-E0x00020000 "                                	, 100
		Gui, AHG: Add, Progress	, % "x" cX " 	y" tY + tH + 10 " w" cw	"       	h25		c972842	-E0x00020000 "                                	, 100
		Gui, AHG: Add, Progress	, % "x" iX "  	y" tY + tH + 10 " w" iw	"       	h25		c972842	-E0x00020000 "                                	, 100
		Gui, AHG: Add, Progress	, % "x" zX " 	y" tY + tH + 10 " w" zw	"       	h25		c972842	-E0x00020000 "                                	, 100
		Gui, AHG: Add, Text     	, % "xm     	y" tY + tH + 15 " w" vw	" vT2 	Center	BackgroundTrans"                                             	, Hostring Vorschläge
		Gui, AHG: Add, Text     	, % "x" cX " 	y" tY + tH + 15 " w" cw	" vT3 	Center	BackgroundTrans"                                             	, Hotstring Konflikte
		Gui, AHG: Add, Text     	, % "x" iX "  	y" tY + tH + 15 " w" iw	" vT4 	Center	BackgroundTrans"                                             	, % "Synonym ICD's (" hotstring.icd ")"
		Gui, AHG: Add, Text     	, % "x" zX " 	y" tY + tH + 15 " w" zw	" vT5 	Center	BackgroundTrans"                                             	, % "Stamm ICD's (" hotstring.StammICD ".+)"

	;-: Hotstring Vorschläge - Checkboxen
		GuiControlGet, t, AHG: Pos, Prg1
		hotstring.vY	:= cY:= iY:=aY:= ty + tH +10					;aY = y Position des Auswahlbereiches (Vorschlag, Konflikt, Listviews)
		hotstring.vW	:= vW
		ADGuiAddCB()
		vY:= hotstring.vY

	;-: Mehrfachauswahl, Seitenangabe
		Gui, AHG: Font, s10 Normal q5 cBlack
		Gui, AHG: Add, Checkbox , % "xm+2 yp+10 vAchk1 gAHGOptions"                                                                                              	, % "Seitenangabe (re./li./bds.)"
		Gui, AHG: Add, Checkbox , % "xm+2 yp+10 vAchk2 gAHGOptions"                                                                                              	, % "Mehrfachauswahl"
		Gui, AHG: Add, Button   	, % "xm+2 yp+10 vReady gAHGReady"                                                                                                  	, % "FERTIG"

	;-: Hotstring Konfliktanzeige (Editfeld)
		GuiControlGet, t, AHG: Pos, Ready
		Gui, AHG: Font, s8 c660000
		aH:= tY + tH - aY < aMinH ? aMinH : tY + tH - aY	; berechnet die Höhe des Auswahlbereiches
		;EditOption1:= "t40 t50 t70 -0x00200040 ReadOnly WantTab"
		EditOption1:= "t40 t50 -0x00200040 ReadOnly "
		Gui, AHG: Add, Edit	   		, % "x" cX+1  	 " y" aY " w" cW-2 " h" aH " vAConflict HWNDhAConflict " EditOption1                                 	, % "Hotstring `t| Zeile `r`n"
		GuiControl, AHG: Move, Ready, % "y" aY + aH - tH
		GuiControlGet, t, AHG: Pos, Ready
		GuiControl, AHG: Move, Achk2, % "y" tY - tH - 10
		GuiControlGet, t, AHG: Pos, Achk2
		GuiControl, AHG: Move, Achk1, % "y" tY - tH - 10

	;-: Listview's - Synonym ICD's und Stamm ICD's
		Gui, AHG: Font, s10 Normal q5 cBlack
		LVOption1:= "LV0x00000431 Grid -Hdr -WantF2 r16"
		Gui, AHG: Add, ListView 	, % "x" iX+2	" y" aY " w" iW-4  " h" aH " vLvSynICD     	gAHGLVHandler1 HWNDhLVSynICD	  	 " LVOption1	, Synonym ICD
		Gui, AHG: Add, ListView 	, % "x" zX+2	" y" aY " w" zW-4 " h" aH " vLvStammICD	gAHGLVHandler2 HWNDhLVStammICD " LVOption1	, Stamm ICD
		WinSet, ExStyle, 0x0, % "ahk_id " hPrg5

	;-: Listview's werden mit Daten befüllt
		ADZusatzICD()                                                                                                                                                                          	; Listview Controls werden mit Daten befüllt

	;-: Hinweis bei nicht vorhandenem ICD-Code in der aktuellen ICD-Ausgabe
		If !hotstring.synICD.MaxIndex()
		{
				ICDAltStatus:= 1
				Gui, AHG: Font, s18 bold q5 c660000
				GuiControl, AHG: Hide, LvSynICD
				Gui, AHG: Add, Text, % "x" iX+2 " y" aY " w" iW-4	" HWNDhICDAlt1 h36 BackgroundTrans Center Border", % "ACHTUNG!"
				Gui, AHG: Font, s12 bold q5 c660000
				noICDtext:= "`nDer ICD-Code`n{" hotstring.icd "}`nfindet sich nicht in der`nICD-10-GM Version `n(" StrReplace(hotstring.ICDVersion,"ICD10GM") ")`n`nMöglicherweise ist dieser ICD-Code veraltet."
				Gui, AHG: Add, Text, % "x" iX+2 " y" aY+35 " w" iW-4	" h" aH-35 " HWNDhICDAlt2 BackgroundTrans Center Border", % noICDtext
		}

	;-: Synonymhotstrings - manueller Eingabebereich
		GuiControlGet, t, AHG: Pos, LvSynICD
		Y:= tY+tH
		Gui, AHG: Add, Progress	, % "xm y" Y+10 " w" gW-2 " h52 c972842 -E0x00020000 vPrg4 HWNDhPrg4"                                   	, 100                  	;{
		WinSet, ExStyle, 0x0, % "ahk_id " hPrg4                                                                                                                                	;
		Gui, AHG: Font, s12 bold q5 cWhite ;}
		Gui, AHG: Add, Text     	, % "xm y" Y+15 " w" gW-2 " vT6 Center BackgroundTrans"                                                                   	, % "Synonyme"  	;{
		GuiControlGet, t, AHG: Pos, T6
		Y:= tY+tH
		Gui, AHG: Font, s9 bold q5 cWhite ;}
		Gui, AHG: Add, Text     	, % "xm y" Y+2 	 " w" gW-2 " vT7 Center BackgroundTrans"                                                                    	, % hinweis1      	;{
		GuiControlGet, p, AHG: Pos, Prg4
		GuiControlGet, t, AHG: Pos, T7
		prg4H:= tY + tH - pY + 5
		GuiControl, AHG: MoveDraw, Prg4, % "h" Prg4H
		Gui, AHG: Font, s11 Normal cBlack                          ;}                                                                                                                             	; Synonym Editfeld
		Gui, AHG: Add, Edit      	, % "xm+1 y" pY+Prg4H+10 " w" gW-4 " r4 vAEdit1 HWNDhAEdit1 gAHGConflictCheck"                     	, % ""

	;-: Vorschau Funktionsaufruf
		GuiControlGet, t, AHG: Pos, AEdit1
		Gui, AHG: Font, s12 Normal cWhite                                                                                                                                            	; Progress Vorschau
		Gui, AHG: Add, Progress	, % "xm 	 y" tY+tH+10 " w" gW-2 " h32 vPrg5 c172842 -E0x00020000 HWNDhPrg5"                        	, 100
		Gui, AHG: Add, Text 	  	, % "xm 	 y" tY+tH+15 " w" gW-2  		" vT8 Center BackgroundTrans"                                                 	, Vorschau - ICD Diagnosen - Funktionsaufruf ;{
		GuiControlGet, t, AHG: Pos, Prg5
		Gui, AHG: Font, s12 Normal cBlack     ;}
		Gui, AHG: Add, Edit			, % "xm+1	 y" tY+tH+10 " w" gW-4    	" vTheString HWNDhTheString gAHGTheString r6"                     	, % "AutoICD(" q hotstring.string q ", " q q ")"

		Gui, AHG: Show, AutoSize, AutoHotstring - Hotstring übernehmen & Konflikte beseitigen
		Progress, Off

		hotstring.gui.hTheString     	:= hTheString
		hotstring.gui.hAEdit1         	:= hAEdit1
		hotstring.gui.hLVSynICD    	:= GetHex(hLVSynICD)
		hotstring.gui.hLVStammICD	:= GetHex(hLVStammICD)


return ;}

AHGReady: ;{

	; Warnungsdialog bei Konflikten anzeigen
		If hotstring.Warn = 1
		{
				MsgBox, 2, Addendum für Albis on Windows, % "Es bestehen noch Hotstringkonflikte!`nMöchten Sie den neuen Hotstring dennoch erstellen?"
				IfMsgBox, No
					return
		}

	; Haupt-Hotstring ermitteln
		Loop, % hotstring.suggest.suggestions.MaxIndex()
		{
				GuiControlGet, status, AHG:, % "Schk" A_Index
				GuiControlGet, value, AHG:, % "SEdit" A_Index
				If Status
					If StrLen(value) > 0
						hotstring.mainStr := value
		}

	; alternative Hotstrings ermitteln
		ControlGetText, value,, % "ahk_id " hAEdit1
		If StrLen(value) > 0
			Loop, Parse,  value, `,
				hotstring.alternative.Push({"short":Trim(StrSplit(A_LoopField, "*").1), "description":Trim(StrReplace(A_LoopField, "*"))})

	; Daten aus der Gui sammeln
		Gui, AHG: Submit
		hotstring.FunctionText:= TheString
		Gui, AHG: Destroy

	; String Sendefunktion aufrufen
		ADSend()

return ;}

AHGchkToggle: ;{

		;Gui, AHG: Submit, NoHide
		RegExMatch(A_GuiControl, "\d+", nr)
		Loop, % hotstring.suggest.suggestions.MaxIndex()
		{
				GuiControl, AHG: , % "Schk" A_Index, % (nr = A_Index ? 1 : 0)
				GuiControl, % "AHG: " (nr = A_Index ? "Enable" : "Disable"), % "SEdit" A_Index
		}

		gosub AHGConflictCheck

return ;}

AHGLVHandler1: ;{

	Gui, AHG: Submit, NoHide
	Gui, AHG: Default
	Gui, AHG: ListView, LvSynICD

	If InStr(A_GuiEvent, "DoubleClick")
	{
			LV_GetText(text, A_EventInfo)
			text:= RegExReplace(text, "\{.*\}")
			oText:= RegExReplace(AEdit1, "\,*\s*$")
			sText:= StrReplace(oText, "\*")

			If !oText
				ControlSetText,, % text, % "ahk_id " hAEdit1
			else
				ControlSetText,, % oText "," text, % "ahk_id " hAEdit1
	}

return
;}

AHGLVHandler2: ;{

	Gui, AHG: Submit, NoHide
	Gui, AHG: Default
	Gui, AHG: ListView, LvStammICD

	If InStr(A_GuiEvent, "DoubleClick")
	{
			LV_GetText(text, A_EventInfo)
			If !InStr(TheString, text)
			{
					hotstring.FunctionText := RegExReplace(TheString, "(AutoICD\(\" q ".*)(\" q "\,\s)(\" q ")(.*)(\" q "\))", "$1|" RegExReplace(text, "G*}", "G}") "$2$3$4$5")
					ControlSetText,, % hotstring.FunctionText , % "ahk_id " hTheString
			 }
			 else
			{
					hotstring.FunctionText:= RegExReplace(TheString, "\|" RegExReplace(text, "G*}", "G}"))
					ControlSetText,, % hotstring.FunctionText , % "ahk_id " hTheString
			}
	}

return
;}

AHGTheString: ;{

	;if hotstring.synICD.MaxIndex()
	;	return

	Gui, AHG: Submit, NoHide
	TheString:= RemoveDoubles(TheString, "(){}[];., ")

	If (ICDAltStatus = 1) && !InStr(TheString, hotstring.icd)
	{
		Control, Hide,,, % "ahk_id " hICDAlt1
		Control, Hide,,, % "ahk_id " hICDAlt2
		ICDAltStatus := 0
	}
	else if (ICDAltStatus = 0) && InStr(TheString, hotstring.icd)
	{
		Control, Show,,, % "ahk_id " hICDAlt1
		Control, Show,,, % "ahk_id " hICDAlt2
		GuiControl, AHG: , ICDAlt, % "`nDer ICD-Code`n{" hotstring.icd "}`nfindet sich nicht in der`nICD-10-GM Version `n(" StrReplace(hotstring.ICDVersion,"ICD10GM") ")`n`nMöglicherweise ist dieser ICD-Code veraltet."
		ICDAltStatus := 1
	}

return
;}

AHGOptions: ;{

	Gui, AHG: Submit, NoHide
	TheString:= RemoveDoubles(TheString, "(){}[];., ")
	hotstring.Option:= (Achk2 = 1 ? "M" : "")
	hotstring.Option.= (Achk1 = 1 ? "S" : "")
	hotstring.FunctionText := RegExReplace(TheString, "(AutoICD\(\" q ".*)(\" q "\,\s)(\" q ")(.*)(\" q "\))", "$1$2$3" hotstring.Option "$5")
	ControlSetText,, % hotstring.FunctionText, % "ahk_id " hTheString

return ;}

AHGConflictCheck: ;{

	Gui, AHG: Submit, NoHide
	tocheck:=""

	; prüft immer auch die Vorschläge
		Loop, % hotstring.suggest.suggestions.MaxIndex()
		{
				GuiControlGet, status, AHG:, % "Schk" A_Index
				GuiControlGet, value, AHG:, % "SEdit" A_Index
				If Status
					If StrLen(value) > 0
					{
						tocheck.= value
						break
					}
		}

	; Erstellen des Prüf-Arrays
		;RegExMatch(AEdit1, "O)(\([\w\d\s-_]+\))", Match)
		tocheck         	.= "," AEdit1
		checkStrings   	:= StrSplit(RTrim(tocheck, ","), ",")
		;hotstring.notice	:= "Hotstring `t| Zeile`n`n ---------------`t| --------`n"
		hotstring.notice	:= "Hotstring `t| Zeile `r`n"
		foundconflict 	:= 0

	; Prüfalgorithmus
		Loop % checkStrings.MaxIndex()
		{
				If InStr(checkStrings[A_Index], "*")
					short:= StrSplit(checkStrings[A_Index], "*").1
				else
					short:= checkStrings[A_Index]

				hotstringStatusText := short " `t| ok `r`n"

			; it adds one char per loop and compare if this shortcut exists, if exists one more round is necessary
				Loop, % StrLen(short) - 4
				{
						abrv      	:= SubStr(short, 1, A_Index + 4)
						addChar	:= false

						For index, value in AutoDiagnosen.hotstring
							If RegExMatch(value.string, "i)^" abrv)
							{
									addChar:= true
									If StrLen(short) = StrLen(abrv)
									{
											;hotstringStatusText := short "`t| " AutoDiagnosen.diagnose[A_Index].icd "`t| " value.line "`n"
											hotstringStatusText := short " `t| " value.line " `r`n"
											addChar     	:= false
											foundconflict	++
									}
									break
							}

							If !addChar
								break
				}

				hotstring.notice .= hotstringStatusText
		}

	;ToolTip, % hotstring.notice
	; Warnung wird bei erkanntem Konflikt gesetzt
		If foundconflict > 0
			hotstring.Warn:= 1
		else
			hotstring.Warn:= 0

	; Anzeige der Konflikte
		ControlSetText,, % hotstring.notice, % "ahk_id " hAConflict

return
;}

AHGGuiClose: ;{
AHGGuiEscape:
	Gui, AHG: Destroy
return ;}

}

ADGuiAddCB() {                                                                                                                   	;-- fügt der Gui die Hotstringvorschläge hinzu

		global
		hotstring.notice  := ""

		Gui, AHG: Default
		Gui, AHG: Font, s10 cBlack

	; Checkboxes und Editfelder für die berechneten Hotstringvorschläge erstellen
		For index, value in hotstring.suggest.suggestions
		{
				Gui, AHG: Add, Radio   	, % "x-7 y" (hotstring.vY+4) (A_Index=2 ? " Checked " : " ") " Right gAHGchkToggle vSchk" A_Index " HWNDhSchk" A_Index, % ""
				GuiControlGet, AHGr, AHG: Pos, % "Schk" A_Index
				Gui, AHG: Add, Edit			, % "x" (AHGrX + AHGrW + 4) " y" hotstring.vY " r1 w" (hotstring.vW - AHGrW + 10) (A_Index=2 ? "" : " Disabled") " gAHGConflictCheck vSEdit" A_Index , % value
				hotstring.vY += ControlRowInc("SEdit" A_Index)
		}

	; wenn nur ein automatischer Vorschlag vorhanden ist, muss Feld1 aktiviert werden
		If hotstring.suggest.suggestions.MaxIndex() = 1
		{
			GuiControl, AHG:, Schk1, 1
			GuiControl, AHG: Enable, SEdit1
		}

	; den String für die Anzeige im Hotstringkonflikt Editfeld erstellen, Hotstring Warnflag setzen
		For index, value in hotstring.suggest.conflicts
		{
				hotstring.notice .= value "`n"
				hotstring.Warn := 1
		}

return
}

ADZusatzICD() {                                                                                                                    	;-- befüllt die beiden Listviewfelder

	; Parsen der Diagnosen.txt Datei
		FileRead, zICD, % A_ScriptDir "\include\Diagnosen.txt"
		Loop, Parse, zICD, `n, `r
		{
				If (A_Index = 1)
				{
					RegExMatch(A_LoopField, "(?<=Diagnosen\:\s)\d+", LpFields)
					RegExMatch(A_LoopField, "(?<=Version\:).*", ICD10Version)
					hotstring.ICDVersion:= Trim(ICD10Version)
				}

				Progress, % Ceil(90 * (A_Index/LpFields))

				If InStr(A_LoopField, hotstring.StammICD)
					If InStr(A_LoopField, hotstring.icd)
					{
						hotstring.synICD.Push(A_LoopField)
					}
					else
					{
						hotstring.addICD.Push(A_LoopField)
					}
		}

	; Synonym-ICD Listview befüllen
		Gui, AHG: Default
		Gui, AHG: ListView, LvSynICD
		Loop % hotstring.synICD.MaxIndex()
		{
			If !InStr(hotstring.synICD[A_Index], hotstring.diagnose)
				LV_Add("", hotstring.synICD[A_Index])
			Progress, % 90 + Ceil(5 * (A_Index / hotstring.synICD.MaxIndex()))
		}

	; Stamm-ICD Listview befüllen
		Gui, AHG: Default
		Gui, AHG: ListView, LvStammICD
		Loop % hotstring.addICD.MaxIndex()
		{
			LV_Add("", hotstring.addICD[A_Index])
			Progress, % 95 + Ceil(5 * (A_Index / hotstring.addICD.MaxIndex()))
		}

}

ADSend() {                                                                                                                             	;-- sendet den fertigen Hotstring

	WinActivate, % "ahk_id " hotstring.ancestor
	ControlFocus,, % "ahk_id " hotstring.fcontrol
	MouseClick, Left, % hotstring.ACaretX, % hotstring.ACaretY

	; Haupt Hotstring
		SendRaw, % ":XR*:" hotstring.mainStr "::                                                                                                                                                               "
		Send, {Tab}
		SendRaw, % "`;`{ " hotstring.HotStringName "`n"

	; alternative oder Synonym Hotstrings
		For index, alternative in hotstring.alternative
			If StrLen(alternative.short) > 0
				If RegExMatch(alternative.short, "^[A_ZÖÄÜ]{2,}") 	; mehrere Großbuchstaben hintereinander - Groß-Kleinschreibungssensitivität hinzufügen
					SendRaw, % ":XRC*:" alternative.short "::                                                                                                `t`; " alternative.description "`n"
				else
					SendRaw, % ":XR*:" alternative.short "::                                                                                                  `t`; " alternative.description "`n"

	; Funktionsaufruf
		SendRaw	, % hotstring.FunctionText "`nreturn"
		Send 		, {Tab}
		SendRaw	, % "`;`}"
		Send     	, % "{Home}{Up " 2 + index "}{Right 12}"

Return
}

ADICDCheck(str, mode:= 1) {                                                                                               	;-- prüft übergeben String ob dieser ein ICD Code enthält

	; Parameter:  	mode = 1 - fragt nach ob
	ICDStrings:= Object()

	If RegExMatch(str, "\{[A-Z0-9\.\!\*]+\}")
	{

	}
return
}

;}

;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; STRING FUNKTIONEN
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
CapitelWords(str) {                                                                                                                 	;-- finds words with capitel chars at beginning
		words:= []
		pos	 := 0
		while, pos:= RegExMatch(str, "([A-ZÄÜÖ][\w\p{L}-]+)", word, pos + 1)
				words.Push(word)
return words
}

RateCapitelWords(str) {                                                                                                            	;-- finds words with capitel chars at beginning and and rate them based on their sentence position
		words:= Object()
		word	:= StrSplit(str, " ")
		Cnr	:= 0

	; Capital words get a rate of 3 at beginning
		Loop, % word.MaxIndex()
			If RegExMatch(word[A_Index], "([A-Z][\w\p{L}-]+)", word)
			{
					Cnr ++
					words[Cnr]:= {"Word": word, "Rate": 3}
			}
			else if RegExMatch(word[A_Index], "([a-z][\w\p{L}-]+)", word)
			{
					If RegExMatch(word[A_Index], "(der)|(die)|(das)|(des)|(von)")
							words[(Cnr)].Rate -= 1
			}

return words
}

RemoveDoubles(Str, DoublesToRemove) {

	Loop % StrLen(DoublesToRemove)
	{
			char:= SubStr(DoublesToRemove, A_Index, 1)
			str:= StrReplace(str, char char, char)
	}

return str
}

TrimDiagnose(str) {                                                                                                                  	;-- entfernt Seitenangaben, ICD Code und ', .G"
	str:= StrReplace(str, ";", " ")
	str:= StrReplace(str, "  ", " ")
	str:= RegExReplace(str, "\,\s*(li|re|bds)\s*\.")
	str:= RegExReplace(str, "\,\s*G\s*\.")
	str:= RegExReplace(str, "\,*\s*Z\.n*\.*\s*")
	str:= RegExReplace(str, "\,\s*(L|R|B)\.\s*")
	str:= RegExReplace(str, "(.*\{[A-Z]\d+\.*\d*)(L|R|B)*(G|V|A|Z)*(\})","${1}G${4}")
return Trim(str)
}

RemoveLastWords(lastwords) {                                                                                                	;-- entfernt die letzten Worte aus einem Edit oder RichEdit-Control

		;hActiveDiaControl 		:= GetFocusedControl()
		;ctext      	:= ControlGetText(hActiveDiaControl )
		ctext      	:= activeControl.diaText
		words   	:= StrSplit(lastwords, "|")

		; die RichEdit Befehle bringen Albis machmal zum Absturz oder funktionieren nicht
		Loop % words.MaxIndex()
		{
				CaretPos	:= CaretPos(activeControl.hwnd)
				idx			:= 1 + words.MaxIndex() - A_Index
				word	    	:= words[idx]
				if word = ""
						break
				wordStart	:= InStr(CText, word)
				wordEnd	:= wordstart + StrLen(word)
				goCaret	:= wordEnd - CaretPos
				;ToolTip, % "word: " word "`nwordstart: " wordstart "`nwordend: " wordEnd "`nCaretPos: " CaretPos "`ngoCaret: " goCaret, 800, 500, 5

				If goCaret != 0
				{
						If goCaret < 0
							Send, % "{Right " (goCaret * -1) "}"
						else
							Send, % "{Left " goCaret "}"
				}

				Send, % "{BackSpace " StrLen(word) +1 "}"
		}
}

rXMatch(Haystack, Needle, OutVarReturnNr:= "", startpos:= 1) {                                            	;-- nicht nutzbar für MatchObjekt (option: 'O)')
	RegExMatch(Haystack, Needle, Match, startpos)
return Match%OutVarReturnNr%
}

;}

;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; HILFS FUNKTIONEN
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
CaretPos(ControlId) {                                                                                                              	;-- Get start and End Pos of the selected string - Get Caret pos if no string is selected
	;https://autohotkey.com/boards/viewtopic.php?p=27979#p27979
	DllCall("User32.dll\SendMessage", "Ptr", ControlId, "UInt", 0x00B0, "UIntP", Start, "UIntP", End, "Ptr")
	SendMessage, 0xB1, -1, 0, , % "ahk_id" ControlId
	DllCall("User32.dll\SendMessage", "Ptr", ControlId, "UInt", 0x00B0, "UIntP", CaretPos, "UIntP", CaretPos, "Ptr")
	if (CaretPos = End)
	SendMessage, 0xB1, % Start, % End, , % "ahk_id" ControlId	;select from left to right ("caret" at the End of the selection)
		else
	SendMessage, 0xB1, % End, % Start, , % "ahk_id" ControlId	;select from right to left ("caret" at the Start of the selection)
	CaretPos++	;force "1" instead "0" to be recognised as the beginning of the string!
return, CaretPos
}

LBEX_CalcIdealWidthEx(hLB, Content:= "", Delimiter:= "|", FontOptions:= ""                           		;-- zum Berechnen der optimalen Breite einer Listbox
, FontName:= "", rows:=0) {

	DestroyGui := MaxW := 0
	static SM_CVXSCROLL

	If !SM_CVXSCROLL
		SysGet, SM_CVXSCROLL, 2	; width of a vertical scrollbar

	If !hLB
	{
		  If (Content = "")
			 Return -1
		  Gui, LB_EX_CalcContentWidthGui: Font, % FontOptions, % FontName
		  Gui, LB_EX_CalcContentWidthGui: Add, ListBox, % "+HWNDhLB " (rows = 0 ? "" : "r" rows), % Content
		  DestroyGui := True
	}

	ControlGet, Content, List,,, % "ahk_id " hLB
	Items := StrSplit(Content, "`n")
	SendMessage, 0x31, 0, 0,, % "ahk_id " hLB                     	; WM_GETFONT
	hFont	:= ErrorLevel
	hDC  	:= DllCall("User32.dll\GetDC", "Ptr", hLB, "UPtr")
	DllCall("Gdi32.dll\SelectObject", "Ptr", hDC, "Ptr", hFont)
	VarSetCapacity(SIZE, 8, 0)
	For Each, Item In Items
	{
		DllCall("Gdi32.dll\GetTextExtentPoint32", "Ptr", HDC, "Ptr", &Item, "Int", StrLen(Item), "UIntP", Width)
		MaxW := Width > MaxW ? Width : MaxW
	}

	DllCall("User32.dll\ReleaseDC", "Ptr", hLB, "Ptr", hDC)

	; einfachste Umsetzung um einfach nur herauszufinden ob eine vertikale Scrollbar existiert
	If rows = 0
		SBVERT_Exist:= 0
	else if Items.MaxIndex() > rows
		SBVERT_Exist:= 1

   If (DestroyGui)
      Gui, LB_EX_CalcContentWidthGui: Destroy

Return MaxW + (SBVERT_Exist = true ? SM_CVXSCROLL : 0) + 8 ; + 8 for the margins
}

ControlRowInc(ControlID, S:="H", addM:= 5 ) {                                                                      	;-- Gui helper function
	GuiControlGet, c, AHG: Pos, % ControlID
return (S = "H" ? cH + addM : cW + addM)
}

BlockInputShort(time) {                                                                                                             	;-- verhindert nach Zeitvorgabe Tasten- und Mauseingaben
	; ein Verzögerung scheint notwendig zu sein damit der Keyboardhook noch alle Eingaben löschen kann
	static OnTime
	OnTime:= time
	SetTimer, BlockDelay, -100
return
BlockDelay:
	;BlockInput, On
	SetTimer, BlockInputOff, % "-" OnTime
return
BlockInputOff:
	BlockInput, Off
return
}

NeueDiagnose(ADInterception) {                                                                                              	;-- verhindert das Hotstrings beim Korrigieren einer Diagnose  ausgelöst werden

		; Funktion liest den Inhalt des RichEditControls aus und ermittelt die Position des Eingabecursor (Caret)
		; steht der Cursor innerhalb eines Diagnosetextes oder genau am Anfang wird false zurückgegeben, ansonsten true
		; dies verhindert das ein Hotstring während der manuellen Änderung einer Diagnosebezeichnung ausgelöst wird

		static lastactive
		;ToolTip, % "Zeile: " A_LineNumber "`nADInterception: " ADInterception, 800, 600, 7
		Critical

		hactiveID    	:= AlbisGetActiveControl("hwnd")
		If (hactiveID <> lastactive)
				ADInterception := 0, lastactive := hactiveID

		If (ADInterception = true)
				return true

		thisHotString	:= RegExReplace(A_ThisHotkey, "\:.*\:")
		controlText	:= ControlGetText("", "ahk_id " hactiveID)
		CaretPos		:= CaretPos(hactiveID) - StrLen(ThisHotString)
		TextUpToCP	:= SubStr(controlText, 1, CaretPos) SubStr(controlText, CaretPos + StrLen(thisHotString), 1)
		If RegExMatch(TextUpToCP, "(;\s*$)|(^\s*$)") || (controlText = "")
				return true

return false
}

MoveCursorSelect(Move, Select) {                                                                                             	;-- Cursor verschieben und anschließend eine Anzahl an Zeichen selektieren
	Send, % "{" Move "}"
	Sleep, 200
	Send, % "{Shift Down}{" Select "}{Shift Up}"
	Sleep, 300
	Send, % "{Shift Up}"
	Send, % "{LControl Up}"
	Send, % "{Control Up}"
}

HotstringStatitistik(file) {                                                                                                         	;-- zeigt die Anzahl der Abkürzungen und Diagnosen an

	abbreviations := 0
	diagnosis		 := 0

	FileRead,f, % file
	Loop, parse, f, `n, `r
	{
			If RegExMatch(A_LoopField, "\:\w+\:\:")			;\w+\:\:
					abbreviations ++, continue
			else if RegExMatch(A_LoopField, "(\{[\w\.]*\})")
			{
					RegExReplace(A_LoopField, "(\{[\w\.]*\})", "", rplcount)
					diagnosis += rplcount
			}
	}

	text =
	(LTrim
	    HOTSTRINGSTATISTIK FÜR AUTO-ICD-DIAGNOSEN
	---------------------------------------------------------------------
	Abkürzungen: `t%abbreviations%
	Diagnosen:      `t%diagnosis%

	Möchten Sie mehr sehen, dann drücken Sie auf 'Ja'
	)
	MsgBox, 4, Addendum für Albis on Windows, % text

}

ToolTipDebug(line, str1:="", str2:="", str3:="") {

	static t

	t.= line ": " (str1!="" ? str1 : "--") (str2!="" ? ", " str2 : ", --") (str3!="" ? ", " str3 : ", --") "`n"
	ToolTip, % t, 3000, 50, 3

}

TTipC(TTipStr, comp, win:="ahk_class OptoAppClass", x:=1200, y:=0, MaxLines:= 3) {

	static t

	If !InStr(compname, comp)
		return

	t.= TTipStr "`n"
	p:= InStr(t,"`n",,1,c)
	If c>MaxLines
	{
		Loop, Parse, t, `n
			If A_Index > 1
					n.= A_LoopField "`n"
		t:= n
	}

	hwnd:= WinExist(win)
	a:= GetWindowSpot(hwnd)
	ToolTip, % TTipStr, % a.X + x, % a.Y + y, 19
	SetTimer, TTipCOff, -20000

return

TTipCOff:
	ToolTip,,,, 19
return
}
;}

/*

	If hotstring.conflicts
		Loop, % hotstring.conflicts
		{
				GuiControlGet, status, AHG:, % "Kchk" A_Index
				If Status
				{
						GuiControlGet, value, AHG:, % "KEdit" A_Index
						If Trim(value) != conflicts[A_Index]
							hotstring["conflict"].Push({"old":conflicts[A_Index], "new":value})
				}
		}



	If hotstring.suggest.suggestions.MaxIndex()
		Loop, % hotstring.suggest.suggestions.MaxIndex()
		{
				GuiControlGet, status, AHG:, % "Schk" A_Index
				GuiControlGet, value, AHG:, % "SEdit" A_Index
				If Status
					If StrLen(value) > 0
						hotstring.mainStr := value
		}

		;GuiControlGet, status, AHG:, % "Achk2"
		;hotstring.Option:= (status = 1 ? "M" : "")

		;GuiControlGet, status, AHG:, % "Achk1"
		;hotstring.Option.= (status = 1 ? "S" : "")

*/

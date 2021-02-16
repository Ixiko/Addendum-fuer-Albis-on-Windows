; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . .
; . . . . . . . . . .   ADDENDUM AUTO-ICD-DIAGNOSEN  Versuch 3
; . . . . . . . . . .   letzte Änderung 12.09.2020
; . . . . . . . . . .
; . . . . . . . . . .	ROBOTIC PROCESS AUTOMATION FOR THE GERMAN MEDICAL SOFTWARE "ALBIS ON WINDOWS"
; . . . . . . . . . .	BY IXIKO STARTED IN SEPTEMBER 2017 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE
; . . . . . . . . . .	RUNS WITH AUTOHOTKEY_H AND AUTOHOTKEY_L IN 32 OR 64 BIT UNICODE VERSION
; . . . . . . . . . .	THIS SOFTWARE IS ONLY AVAIBLE IN GERMAN LANGUAGE
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .

	/*  	AUTO-ICD-Diagnosen Versuch 3 *
		. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .

			AUTO-ICD-Diagnosen ist ein Textexpander der schon bei Erkennung eines Teils einer Krankheitsbezeichnung die richtige ICD-Diagnose findet.
			Es ist kein Autokorrekturvorgang wie in einem Textverarbeitungsprogramm.
			Ein Beispiel. Sie wollen die Diagnose für Pityriasis versicolor eingeben. Das Skript erkennt auch Synonyme. So wird es auch auf Kleienpilzflechte,
			Kleienflechte oder Kleieflechte reagieren und den Diagnosentext "Pityriasis versicolor {B36.0G}" ausgeben.

		. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .

			VERSUCH 3: Was ist anders?
			----------------------------------------------------------------------------------------------------------------------
			1.	Die Wortvervollständigung erfolgt nicht mehr durch Hotstrings, da diese nicht flexibel genug sind.

			2. Es werden zwei Datenquellen verwendet

				Datei 1 - 	ICDHotstrings.txt - enthält Kurzworte und Diagnosen und es wird eine Liste aller aktuellen ICD-10-GM Kode's geladen.
				Datei 2 - 	ICD-10-GM Klassifikation vom Bundesinstitut  für Arzneimittel und Medizinprodukte (ehemals DIMDI)
								(ICD-10-GM 2020 Alphabet EDV-Fassung TXT)
								Link: https://www.dimdi.de/dynamic/de/klassifikationen/downloads/?dir=icd-10-gm
								Das Skript verwendet eine modifizierte Originaldatei.
								Es wurden alle Einräge in ein Format für die direkte Verwendung mit Albis umgewandelt. (  Diagnosentext {[A-Z][00-99][.0-99]}  )


			3.	Das Skript erkennt automatisch eine Diagnosezeile anhand des Karteikartenkürzel 'dia'. Daraufhin wird ein Inputhook gestartet.
				Bei Erfolg wird oberhalb der Eingabezeile in Albis ein Fenster mit dem Inhalt "Textexpander für Diagnosen verfügbar"
				In diesem kleinen Fenster werden auch weitere Optionen angezeigt.

			4.	Wird das Caret versetzt wird (Tasten: Links, Rechts, Tab, Enter, Escape, Linke Maustaste) wird der Eingabepuffer gelöscht.
				Das bis dahin eingegebene bleibt in der Zeile stehen und wird durch einen neue gestarteten Inputhook nicht weiter verwendet!

			5.	Der Aufbau dieser Datei hat ein striktes Schema. Es gehören immer zwei Zeilen zusammen:

				Zeile 1	- 	enthält die Abkürzungen, dies sind die kürzest möglichen Wortanfänge für Krankheiten. Die Verkürzungen ergeben nicht
								immer einen Sinn und stehen oftmals für alternative Krankheitsbezeichnungen (auch nicht lateinische) im Sinne eines Thesaurus.
								Es kann mehrere Verkürzungen zu nur einer Diagnose geben oder nur eine Verkürzung zu mehreren Diagnosen.
				Zeile 2	-	enthält eine oder mehrere nach ICD-10-GM 19, getrennt durch ein '|' .
								Gehören mehrere Diagnosen zu einem Kürzel wird ein Auswahldialog geöffnet, aus dem eine bis alle gewünschten
								Diagnosen gewählt und eingefügt werden können.

			6.	Nach dem das Skript die gewählten die Diagnosen eingetragen hat, läuft der Inputhook weiter. Dieser endet immer erst dann wenn
			das Karteikartenfeld für Diagnosen verlassen/beendet wird (z.B. TAB).

			7.	Das Skript ist ohne weitere Bibliotheken verwendbar. Alle notwendigen Funktionen sind enthalten.
				Zur Ausführung kann jede Unicode Autohotkey(L/H).exe ab Version 1.1.32.01 verwendet werden.

		. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
		. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .

	*/

	; TODO
	;
	; Einstellungen 	- Diagnosen editieren
	; Datenbank   	- ?
	; Seitenangaben
	; Ergänzungen nach Eingabe
	;

	;{ SKRIPTEINSTELLUNGEN

		#NoEnv
		#Persistent
		#InstallKeybdHook
		#SingleInstance               	, Force
		#MaxThreads                  	, 100
		#MaxThreadsBuffer       	, On
		#MaxHotkeysPerInterval	, 99000000
		#HotkeyInterval              	, 99000000
		#MaxThreadsPerHotkey 	, 3
		#KeyHistory                  	, 0

		SetBatchLines, -1
		SetKeyDelay, -1, -1
		CoordMode, ToolTip	, Screen
		CoordMode, Caret	, Screen
		CoordMode, Mouse	, Screen

		OnExit("DasEnde")

	;}

	;{ VARIABLEN SETZEN / DATEN LADEN

		global CTX    	:= Object()	; Einstellungen
		global hsDia 	:= Array()
		global hsAbbr 	:= Object()
		global hsICD 	:= Array()
		global Albis    	:= Object()
		global SacHook
		global ShellHook
		global fchwnd, FClass, waitcheck := false

		CTX.Hook               	:= Object()
		CTX.Hook.Modes   	:= ["wartet", "Kürzelliste", "ICD Text", "ICD Code", "pausiert"]
		CTX.Hook.inputmode := 0
		CTX.ModusKeys     	:= "-#"
		CTX.Queue           	:= Object()
		CTX.Undo               	:= Object()
		CTX.Debug           	:= false
		;CTX.Debug           	:= true

		;SciTEOutput("----------------------------------------------------------------------------")
		;~ Run, % ConvertDiagnosis()
		;~ ExitApp

		; Tray Icon erstellen und anzeigen
			hIconCTEXEP := Create_ConTextExpander_ico()
			If hIconCTEXEP
				Menu, Tray, Icon				, % "hIcon: " hIconCTEXEP
			else {
				MsgBox, % hIconCTEXEP
				ExitApp
			}

		; Diagnosen und Abbreviationen laden
			LoadDiagnosen()                   	; ICD Diagnosen/Daten laden

	;}

	;{ HOOK'S STARTEN

		; Eingabe-Hook für Albis starten
			Albis.FokusHook	:= AlbisFocusHook(AlbisPID())

		; Hotkey für Skriptneustart
			fReload	:= Func("ReloadScript")
			fExitApp:= Func("HotkeyEnde")
			Hotkey, IfWinActive, ahk_class OptoAppClass			            			;~~ ALBIS Hotkey - bei aktiviertem Fenster
			Hotkey, ^!t	, % fReload
			Hotkey, ^!d	, % fExitApp
			Hotkey, IfWinActive

		; Input Hook einrichten	;{
			SacHook := InputHook("V I1", "{Esc}{Tab}{Enter}{{}{}}{;}{Up}{Down}{Left}{Right}{Delete}{Home}{End}")
			SacHook.OnChar      	:= Func("SacChar")
			SacHook.OnKeyDown	:= Func("SacKeyDown")
			SacHook.OnEnd        	:= Func("SacEnd")
			SacHook.KeyOpt("{Backspace}", "N")
		;}

		; Hotkey's installieren 	;{
			fHelp    	:= Func("sacHelp")
			fMEnd   	:= Func("SacMouseEnd")
			fOptions	:= Func("sacProps")
			fTPause 	:= Func("sacTogglePause")
			fUndo   	:= Func("sacUndo")
			Hotkey, IfWinExist	, ConTextExpander ahk_class AutoHotkeyGUI
			Hotkey, $^h     	, % fHelp
			Hotkey, ~LButton	, % fMEnd
			Hotkey, $^o	    	, % fOptions
			Hotkey, $^p	    	, % fTPause
			;Hotkey, $^z	    	, % fUndo
		;}

	;}

	; DEBUGANZEIGE
		If CTX.Debug
			SetTimer, SacHookInprogress, 3000

return

SacHookInprogress: ;{

	ToolTip, % "Tastatur-Hook: " (SacHook.InProgress ? "running" : "stopped")  ", InputMode: " (CTX.Hook.inputmode) "`nwaitcheck: " waitcheck, 800, 1, 15

return ;}

ReloadScript() {

	UnhookWinEvent(Albis.FokusHook.hEvH, Albis.FokusHook.HPA)
	StopInputHook("Reload")
	Reload

}

;-----------------------------------------------------------------------------------------------------------------------------------------------
; INPUTHOOK FUNKTIONEN
;-----------------------------------------------------------------------------------------------------------------------------------------------;{
;----- Diagnosenexpander
StartInputHook() {

		global

		CTX.Hook.inputmode := 0
		SacHook.Start()

	; Mini-Gui einblenden das ein Inputhook gestartet wurde
		f := GetWindowSpot(CTX.Hook.HookHwnd)
		CTX.Hook.CtrlYH := f.Y + f.H

	; Gui zeichnen
		Gui, sacInfo: new, % "-Caption -DPIScale +ToolWindow +HWNDhsacGUI" ; +Parent" Albis.IH_InfoParent "
		Gui, sacInfo: Margin, 10, 1
		;Gui, sacInfo: Color, c172842, c172842
		Gui, sacInfo: Color, cCCCCCC, cFFFFFF
		Gui, sacInfo: Font, s11 q5 bold cBlack, Futura MD Bt
		Gui, sacInfo: Add, Text  	, % "xm ym+3 vsacProp Center BackgroundTrans +HWNDhsacProp"
												, % "Textexpander für " CTX.Hook.context " (" CTX.Hook.Modes[CTX.Hook.inputmode+1] ")"

		Gui, sacInfo: Font, s9 q5 cBlack, Futura Bk Bt
		Gui, sacInfo: Add, Text   	, % "x+20	ym-2    	   	vsacC1    	HWNDhsacC1	BackgroundTrans"             	, % "Hilfe:"
		Gui, sacInfo: Add, Picture	, % "x+2	ym+1 	h11	vsacC2 	HWNDhsacC2	BackgroundTrans AltSubmit"	, % A_ScriptDir "\res\StrgWx13.png"
		Gui, sacInfo: Font, s9 q5 cBlack, Futura Bk Bt
		Gui, sacInfo: Add, Text   	, % "x+0 	ym-3    	   	vsacC3	   	HWNDhsacC3	BackgroundTrans"             	, % "+"
		Gui, sacInfo: Add, Picture	, % "x+-2 	ym+1	h11	vsacC4 	HWNDhsacC4 	BackgroundTrans AltSubmit"	, % A_ScriptDir "\res\H13x13.png"

		Gui, sacInfo: Font, s10 q5 cBlack, Futura Bk Bt
		Gui, sacInfo: Add, Text   	, % "x+20	ym+13	    	vsacC5    	HWNDhsacC5 	BackgroundTrans"             	, % "Pause:"
		Gui, sacInfo: Add, Picture	, % "x+2	ym+16	h11	vsacC6 	HWNDhsacC6	BackgroundTrans AltSubmit"	, % A_ScriptDir "\res\StrgWx13.png"
		Gui, sacInfo: Font, s9 q5 cBlack	, Futura Bk Bt
		Gui, sacInfo: Add, Text   	, % "x+0 	ym+12	    	vsacC7	   	HWNDhsacC7	BackgroundTrans"             	, % "+"
		Gui, sacInfo: Add, Picture	, % "x+-2 	ym+16	h11	vsacC8 	HWNDhsacC8 	BackgroundTrans AltSubmit"	, % A_ScriptDir "\res\P.png"

		xPlus := [2, 0, 0, 5, 2, 0, 0, 5]
		Loop 8 {

			nr := 9 - A_Index

			If nr in 4,8
			{
				GuiControlGet, d	, sacInfo: Pos, % "sacC" nr
				GuiControl, sacInfo: MoveDraw	, % "sacC" nr     	, % "x" f.W - dW - xPlus[nr]
				GuiControlGet, c	, sacInfo: Pos, % "sacC" nr
			}
			else {
				GuiControlGet, d	, sacInfo: Pos, % "sacC" nr
				GuiControl, sacInfo: MoveDraw	, % "sacC" nr     	, % "x" cX - dW - xPlus[nr]
				GuiControlGet, c	, sacInfo: Pos, % "sacC" nr
			}
		}

		Gui, sacInfo: Show, % "x" (f.X) " y" (f.Y - f.H + 2) " w" (f.W) " h" f.H - 4 " NA", % "ConTextExpander"

		sacInfoChange()

		;WinSet, Transparent, 200    , % "ahk_id " hsacGui
		WinSet, AlwaysOnTop, On  , % "ahk_id " hsacGui
		WinSet, Top                    	,,	% "ahk_id " hsacGui

	; registriert einen Shellhook nur wenn ein Inputhook gestartet wurde
		ShellHook := new ShellHook(hsacGui, SacHook, ["0x4","0x7","0xD","0xE","0x8004"])


return
}

StopInputHook(Reason) {

	;~ If !SacHook.InProgress
		;~ return

	SciTEOutput("sacHook stopped: " Reason)

	SacHook.Stop()

	waitcheck := false
	CTX.Hook.HookHwnd  	:= 0
	CTX.Hook.inputmode	:= 0
	CTX.LastHWND         	:= 0
	ShellHook.Deregister(Reason)

	ToolTip,,,, 3
	ToolTip,,,, 4
	ToolTip,,,, 5

	Gui, sacInfo:	Destroy
	If WinExist("AutoDiagnose ahk_class AutohotkeyGUI") {
		Gui, ACB: Destroy
		AutoCompleteGuiHotkeys("Off")
	}

	Sleep 500

}

EmptyInputHook(ih="") {

	SacHook.Stop()
	SacHook.Start()

	ToolTip,,,, 3
	ToolTip,,,, 5

	CTX.Hook.inputmode := 0
	sacInfoChange()


}

sacHelp() {

}

sacProps() {

	SciTEOutput("Du willst Einstellungen!")
return
}

SacInProgress(ih) {
	SciTEOutput("inprogress: " sacHook.InProgress)
return sacHook.InProgress
}

SacChar(ih, char, autoappend=false) {  ; Wird aufgerufen, wenn ein Zeichen zu SacHook.Input hinzugefügt wird.

		static Diag

		prefix 	:= LTrim(ih.Input)
		pLen 	:= StrLen(prefix)

		SacTTDebug("M" CTX.Hook.inputmode ", L" pLen ", p '" prefix "' ~ " sacHook.input)

		If autoappend
			return

		If   	 (CTX.Hook.inputmode = 0) 	&&	(pLen = 1)	{     	; Modus: '.' - Suche in den Diagnosetexten , '#' - Suche in ICD-Code

				CTX.Hook.inputmode := InStr(CTX.ModusKeys, prefix) + 1
				If (CTX.Hook.inputmode > 1) {
					sacHook.Stop()
					Send, {BackSpace}
					sacHook.Start()
					pLen	:= StrLen(ih.Input)
					sacInfoChange()
				}

		}
		else if (CTX.Hook.inputmode = 1) 	&&	(pLen > 1)	{     	; Kürzel-Modus. erst ab dem 2. Zeichen beginnt die Überprüfung

			; BESTIMMTE WORTE UND ZEICHEN WERDEN AUS DER EINGABE ENTFERNT
				prefix := TrimPrefix(prefix)
				pLen := StrLen(prefix)
				pUC := RegExMatch(prefix, "^[A-ZÄÖÜ\d]+$") ? true : false                ; Eingabe besteht ausschließlich aus Großbuchstaben?

			; EINE NULL WIRD AUS DEM INPUTHOOK ENTFERNT
				If (char = "0") {
					SacRemoveChar()
					return
				}

			; DIAGNOSENLISTE ERSTELLEN WENN ALLES GROßBUCHSTABEN, Zahleneingaben führen zur Ausgabe der ICD-Diagnose
				If (!pUC && !RegExMatch(char, "[1-9]")) || (pUC = true) || (Diag.MaxIndex() = 0){

					; ABKÜRZUNGEN DURCHSUCHEN
						Diag	:= Array()
						For abbr, diapos in hsAbbr {

							; Treffer wenn beide Worte aus Großbuchstaben bestehen (Kurzhotstring z.B. VI = Virusinfekt)
								If  (pUC && (prefix == abbr))  {
									SacModeReset()
									AutoDiagnose(hsDia[diapos], ih, abbr)
									return
								}

							; Matchlist erstellen
								aLen	:= StrLen(abbr), abbrp := SubStr(abbr, 1, pLen), abbrpL := aLen - pLen
								If (pUC && (prefix == abbrp)) || (!pUC && (prefix = abbrp))
									Diag.Push(abbr)

						}

				}

			; MATCHLIST ANZEIGEN ODER ICD-DIAGNOSE EXPANDIEREN
				n  := Diag.MaxIndex()
				If      	(n = 1) {                                                                           	; nur noch eine mögliche Abkürzung vorhanden, dann haben wir schon einen Treffer

						SacModeReset()
						diapos := hsAbbr[Diag[1]]
						AutoDiagnose(hsDia[diapos], ih, Diag[1])
						return

				}
				else If 	(n > 1) {                                                                              	; Anzeige der möglichen Treffer

					; WURDE EINE ZAHL ZWISCHEN 1-9 GETIPPT? DANN WIRD DIE GEWÄHLTE DIAGNOSE ENTFALTET
						If RegExMatch(SubStr(prefix, -1) , "[1-9]", ItemChoosed) {

							SciTEOutput("ItemChoosed: " ItemChoosed)

							If (ItemChoosed <= n)  {                                    	; Expandieren da Treffer

									SacModeReset()
									diapos := hsAbbr[Diag[ItemChoosed]]
									;AutoDiagnose(hsDia[diapos], ih, prefix . ItemChoosed)
									AutoDiagnose(hsDia[diapos], ih, Diag[ItemChoosed])
									return

							}
							else if RegExMatch(SubStr(prefix, -1) , "0") {           	; Zeichen '0' löschen

									SacRemoveChar()

							}

						}

					; --- SCHREIBT DAS ERSTE WORT AUS WENN KEIN ANDERES MEHR IN FRAGE KOMMT
						else if !RegExMatch(StrSplit(Diag[1], " ").1, "^[A-ZÄÖÜ]+\d*\s|$") {

							differents := false
							RegExMatch(Diag[1], "i)^([\wäöüß]+)\s*", word)

							SciTEOutput(word1 " != " prefix)

							If (word1 != prefix)	{                                     	; vergleicht die Anfänge

								Loop % (n - 1) {
									RegExMatch(Diag[A_Index + 1], "i)^([\wäöüß]+\s*)", compare)
									If (word1 != compare1) {
										differents := true
										break
									}
								}

							}

							; dem Inputhook hinzufügen
								If !differents {
									ih.MinSendLevel := 0                        ; SendEvent wird als Tastatureingabe gewertet
									SendEvent, % SubStr(word1, -1 * (StrLen(word1) - StrLen(prefix) -1))
									ih.MinSendLevel := 1                    	; die nächsten SendEvent's beeinflussen nicht den Inputhook
								}

						}

					; TREFFER ANZEIGEN
    					SacShowHits(Diag)

				}
				else
					ToolTip,,,, 3

		}
		else if (CTX.Hook.inputmode = 2) 	&&	(pLen > 3)	{     	; ICD Text - Modus. erst ab dem 2. Zeichen beginnt die Überprüfung

			; BESTIMMTE WORTE UND ZEICHEN WERDEN AUS DER EINGABE ENTFERNT
				prefix := TrimPrefix(prefix)

				If (Diag.MaxIndex() > 1) && RegExMatch(SubStr(prefix, -1) , "\d", ItemChoosed)
					prefix := SubStr(prefix, 1, StrLen(prefix) - 1)

				words  := StrSplit(Trim(prefix), " ")

			; DIAGNOSENTEXTE DURCHSUCHEN
				If StrLen(ItemChoosed) = 0 {


					; WENN EINES DER SUCHWORTE WENIGER ALS 4 ZEICHEN HAT DANN WIRD NICHT GESUCHT
						Loop, % words.MaxIndex()
							If (StrLen(words[A_Index]) < 3)
								return

						Diag	:= Array()
						For ICDindex, ICDCode in hsICD {

							matches := 0
							For wordIndex, word in words
								If InStr(ICDCode, word)
									matches ++
								else break

							If (matches = words.MaxIndex())
								Diag.Push(ICDCode)

						}

				}


				n  := Diag.MaxIndex()
				If      	(n = 1) {                                                                                                    	; nur noch eine mögliche Abkürzung vorhanden, dann haben wir schon einen Treffer

					SacModeReset()
					AutoDiagnose(Diag[1], ih, prefix)
					return

				}
				else If 	(n > 1) {                                                                                                   	; Anzeige der möglichen Treffer

					SacShowHits(Diag, "`n")
					If RegExMatch(ItemChoosed, "[1-9]") {                                                                   	; eine Zahl zwischen 1 - 9 beendet die Auswahl und sendet die gezeigte Diagnose
						SacModeReset()
						AutoDiagnose(Diag[ItemChoosed], ih, prefix . ItemChoosed)
						return
					}

				}

		}
		else if (CTX.Hook.inputmode = 3) 	&&	(pLen > 0)	{     	; ICD Code - Modus. erst ab dem 1. Zeichen beginnt die Überprüfung

				If 		((pLen < 4) && !RegExMatch(prefix, "i)^[A-Z]\d{0,2}(?=$)"))
					|| ((pLen > 4) && !RegExMatch(prefix, "i)^[A-Z]\d{2}\.\d{1,2}$", match)) {
					ToolTip, % "ICD Code: FALSCHES FORMAT!`n" match, % A_CaretX + 20, % A_CaretY, 5
				}
				else
					ToolTip,,,, 5

				If (pLen < 2)
					return

				Diag	:= "", matches := 0
				For ICDindex, ICDCode in hsICD
					If InStr(ICDCode, prefix) && (matches < 250) {
							matches ++
							Diag .= ICDCode "|"
							if (matches >= 250)
								break
					}

				Diag := RTrim(Diag, "|")
				If (StrLen(Diag) > 0)
					AutoCompleteGui(RTrim(Diag, "`n"), "x" A_CaretX " y" CTX.Hook.CtrlYH + 1, 1,,, prefix)

		}
		else if (CTX.Hook.inputmode > 0) 	&&	(pLen < 2)	{    	; Eingabemodus zurücksetzen, ToolTip aus

				If (pLen = 1) {
					ToolTip,,,, 3
					ToolTip,,,, 5
				}
				If (pLen = 0)
					CTX.Hook.inputmode := 0

				sacInfoChange()

		}


		SacTTDebug("M" CTX.Hook.inputmode ", L" pLen ", p '" prefix "' ~ " sacHook.input)

}

SacTTDebug(DebugString) {

	If !CTX.Debug
		return

	ToolTip, % DebugString, 750, 1052, 4

}

SacRemoveChar(charCount=1) {

		SacHook.MinSendLevel := 0                        ; SendEvent wird als Tastatureingabe gewertet
		SendEvent, % "{BackSpace" ( charCount=1 ? "" : " " charCount) "}"
		SacHook.MinSendLevel := 1                    	; die nächsten SendEvent's beeinflussen nicht den Inputhook

}

SacModeReset() {

	ToolTip,,,, 3
	CTX.Hook.inputmode := 0
	sacInfoChange()

}

SacShowHits(arr, sep=", ") {                                                                                                            	;-- Treffer zeigen

	n := arr.MaxIndex()
	Loop, % (n > 9 ? 9 : n)
		tt .= A_Index ": " arr[A_Index] (n = A_Index ? "" : sep)

	ToolTip	, % ((n > 9) ? (tt "  (+" (n - 9) ")") : tt), % A_CaretX - 10, % A_CaretY + 30, 3
	;~ hwnd := WinExist("A")
	;~ t := GetWindowSpot(hwnd)
	;~ If (CTX.Hook.CtrlYH + t.H > A_ScreenHeight)
		;~ SetWindowPos(hwnd, t.X, CTX.Hook.CtrlYH - t.Y - 10, t.W, t.H)

}

SacKeyDown(ih, vk, sc) {

    if (vk = 8) ; Backspace
        SacChar(ih, "")

}

SacMouseEnd() {

	If !SacHook.InProgress
		return


	MouseGetPos,,, hwin, hControl, 3

	If (GetHex(hControl) = GetHex(CTX.Hook.HookHwnd)) {

		sacHook.Stop()
		sacHook.Start()
		SacModeReset()
		ToolTip,,,, 4
		ToolTip,,,, 5
		;SciTEOutput("focus not lost")

	}
	else {
		StopInputHook("Mouseclick")
	}

return
}

SacTogglePause() {

	static lastmode

	if sacHook.InProgress {
		lastmode := CTX.Hook.inputmode
		CTX.Hook.inputmode := 4
		ToolTip,,,, 3
		ToolTip,,,, 4
		ToolTip,,,, 5
		SacHook.Stop()
		Gui, sacInfo: Color, cBB0000, cBB0000
		sacInfoChange()
	}
	else {
		CTX.Hook.inputmode := lastmode
		SacHook.Start()
		Gui, sacInfo: Color, c172842, c172842
		sacInfoChange()
	}

}

SacEnd(ih) {


	If (sacHook.EndReason != "Endkey")
		return

	;SciTEOutput(A_LineNumber ": EndReason: " sacHook.EndReason ", Key: "  sacHook.Endkey)

	If RegExMatch(sacHook.Endkey, "i)Escape|Tab") 	{

		StopInputHook("Endkey pressed: "  sacHook.Endkey)

	}
	else {

		sacHook.Stop()
		sacHook.Start()
		SacModeReset()
		ToolTip,,,, 3
		ToolTip,,,, 5

	}

return
}

sacInfoChange() {

	global hsacProp

	info := "Textexpander für " CTX.Hook.context " (" CTX.Hook.Modes[CTX.Hook.inputmode + 1] ")"

	SendMessage, 0x31, 0, 0,, % "ahk_id " hsacProp                     	; WM_GETFONT
	hFont	:= ErrorLevel

	hDC  	:= DllCall("User32.dll\GetDC", "Ptr", hsacProp, "UPtr")
	DllCall("Gdi32.dll\SelectObject", "Ptr", hDC, "Ptr", hFont)

	VarSetCapacity(SIZE, 8, 0)
	DllCall("Gdi32.dll\GetTextExtentPoint32", "Ptr", HDC, "Ptr", &info, "Int", StrLen(info), "UIntP", Width)
	DllCall("User32.dll\ReleaseDC", "Ptr", hsacProp, "Ptr", hDC)

	GuiControl, sacInfo: MoveDraw	, sacProp, % "w" Width
	GuiControl, sacInfo: 					, sacProp, % info
	Gui, sacInfo: Show, % " NA"

}

sacUndo() {

	;WM_UNDO = 0x0304
	ControlGetFocus, focused, A
	;SciTEOutput(focused)
	If RegExMatch(focused, "i)Edit|RichEdit|ComboBox")
		PostMessage, 0x0304,,,, % "ahk_id " CTX.Hook.HookHwnd

}

;----- Autovervollständigungsvorschläge für andere Kürzel
InitAutoSuggestion(context) {

		If !IsObject(CTX.ASG) {
			CTX.ASG   	:= Object()
			CTX.ASG.dic	:= Object()
		}
		If !IsObject(CTX.ASG[context])
			CTX.ASG.dic[context] := Array()

	; Input Hook einrichten	;{
		CTX.ASG.IH := InputHook("V I1", "{Tab}{Escape}")
		CTX.ASG.IH.OnKeyDown	:= Func("AsgKeyDown")
		CTX.ASG.IH.OnEnd        	:= Func("AsgEnd")
		CTX.ASG.IH.KeyOpt("{.}{Enter}{{}{}}{(}{)}{[}{]}{?}{!}", "N")
	;}

}

AsgKeyDown(ih, vk, sc) {

		static t

		keyname := GetKeyName(Format("vk{:x}sc{:x}", VK, SC))
		t := "VK: " VK "      SC: " SC "        KN: " keyname
		;ToolTip, % t, % A_CaretX - 5, % A_CaretY - 25, 13

}

AsgEnd(ih) {

		CTX.ASG.IH.Stop()

}
;}

;-----------------------------------------------------------------------------------------------------------------------------------------------
; FUNKTION ZUM SENDEN DER ICD-DIAGNOSEN AN ALBIS
;-----------------------------------------------------------------------------------------------------------------------------------------------;{
AutoDiagnose(DX, ih, abbr) {					                     				                        	            				;-- Auto-ICD-Diagnosen

	; abbr : das Diagnosenkürzel durch das die Funktion aufgerufen wurde, wird für den Diagnosenfilter benötigt

	static CaretMode

	CaretMode := A_CoordModeCaret
	CoordMode, Caret, Screen
	SendMode, Event
	SetKeyDelay, -1, -1
	BlockInput, Send

	thisInput := ih.Input
	SacHook.Stop()

	Send, % "{LShift Down}{Left " StrLen(thisInput) "}{LShift Up}"
	;Sleep 2
	Send, % "{BackSpace}"

	DX := DiagnosenFilter(DX, abbr)

	If !InStr(DX, "|") {

		DX := RegExReplace(DX, "G*\}", "G}")
		Send, % "{Text}" DX ";"
		If InStr(DX, "###")
			DiagnoseErgaenzen("###")

	}
	else	{

		DX := AutoCompleteGui(DX, "x" A_CaretX " y" CTX.Hook.CtrlYH + 1, 1)

	}

	SacHook.Start()   ; Hook neu starten
	CoordMode, Caret, % CaretMode

	;CheckICDHotStrings(StrSplit(DX, "|"))

return
}

DiagnosenFilter(DX, abbr:="") {                                                                                                     	;-- entfernt bestimmte Diagnosen

	; die Diagnosenabkürzung (abbr) ist ausschlaggebend dafür welche Diagnosen herausgefiltert werden
	; beginnt abbr mit kleinen Buchstaben, auch wenn das zweite Wort mit einem großen Buchstaben anfängt, wird auch nichts gefiltert
	; nur wenn abbr mit einem Großbuchstaben beginnt, im Sinne eines Eigenwortes, werden nur die Diagnose genommen die abbr enthalten

		If  (StrLen(abbr) = 0) || RegExMatch(abbr, "^[A-ZÄÖÜ]+\d*$") || RegExMatch(abbr, "^[a-zäöüß\d]+")
			return DX

	; Leerzeichen durch Punkt ersetzt. In RegEx findet der Punkt jedes Zeichen. "i)" nicht sensitive Suche für Groß- oder Kleinschreibung nutzen
		DXnew := ""
		;abbr := "i)" StrReplace(abbr, " ", ".")
		abbr := Trim(abbr)
		If RegExMatch(abbr, "([a-zäüöß])([A-ZÄÖÜ])")                                    	; KniePrellung -> Knie Prellung (alle Diagnosen welche das Wort 'Knie'
			abbr := RegExReplace(abbr, "([a-zäüöß])([A-ZÄÖÜ])", "$1 $2")    	; enthalten werden herausgesucht)
		If RegExMatch(abbr, "i)[\wäüöß]+\s[\wäöüß]+")                              	; das erste Wort der Abkürzung ist der Suchparameter
			abbr := "i)" StrSplit(abbr, " ").1

		Loop, Parse, DX, `|
			If RegExMatch(A_LoopField, abbr)
				DXnew .= A_LoopField "|"

		DXnew := RTrim(DXnew, "|")

return StrLen(DXnew) > 0 ? DXnew : DX 	; verhindert bei fehlerhaften Daten in der ICDHotstrings-Datei das keine Diagnosen zurückgegeben werden
}

AutoCompleteGui(Content, GuiOptions, hfocus=0, FontOptions="", FontName="", inputstr="") {       	;-- zeigt mehrere Diagnosen oder anderes zur weiteren Auswahl an

		local guiWidth, fsize, t, a, c, aX, aY
		static DpiF, rows, hfCtrl, CaretX, CaretY, ReplaceInput

		ReplaceInput := inputstr

        FontOptions := !FontOptions	? "s10 Normal q5"	: FontOptions
        FontName    := !FontName	? "Futura Bk Bt"   	: FontName

		hfCtrl := GetHex(GetFocusedControl())
		rows := StrSplit(Content, "|").MaxIndex()
		rows := rows > 10 ? 10 : rows         	; maximale Zeilen begrenzen

	; ------------------------------------------------------------------------
	; AutoComplete Gui erstellen
	; ------------------------------------------------------------------------ ;{

		guiWidth := LBEX_CalcIdealWidthEx(0, Content, "|", FontOptions, FontName, rows)
		guiWidth := guiWidth < 250 ? 250 : guiWidth
		RegExMatch(GuiOptions, "(?<=x)\d+", aX)
		RegExMatch(GuiOptions, "(?<=y)\d+", aY)
		CaretX := aX, CaretY := aY

		Gui, ACB: New 	, -SysMenu -Caption +AlwaysonTop +ToolWindow +HWNDhACB ;0x98200000
		Gui, ACB: Margin	, 0, 0
		Gui, ACB: Color, c172842

		RegExMatch(FontOptions, "(?<=s)\d+", fsize)
		Gui, ACB: Font 	, % RegExReplace(FontOptions, "(?<=s)\d+", fsize - 2), % FontName
		Gui, ACB: Add	, Text, % "x2 y2	HWNDhTACB1 vTACB1 cWhite", % "Diagnosen: " StrSplit(Content, "|").MaxIndex()
		Gui, ACB: Add	, Text, % "x+10	HWNDhTACB2 vTACB2 cWhite", % "[Auswahl: Strg/Shift+Maus, Übernehmen: Enter]"
		GuiControlGet, cp, ACB: Pos, TACB2
		GuiControl, ACB: Move, TACB2, % "x" guiWidth - cpW - 5

		t:= GetWindowSpot(hTACB1)

		Gui, ACB: Font  	, % FontOptions, % FontName
		Gui, ACB: Add  	, ListBox, % "x2 y" (t.Y + t.H + 2) " w" guiWidth " Choose1 Sort HWNDhACBLb vACBLb Multi r" rows , % Content	;+0x8 gACBListbox
		Gui, ACB: Show	, % GuiOptions " AutoSize NA Hide", % "AutoDiagnose"

	; Gui positionieren
		a:= GetWindowSpot(hACB), c:= GetWindowSpot(WinExist("A"))
		aX := (a.X + a.W > A_ScreenWidth) ? (c.W - a.W - 5  )	: (aX)
		aY := (a.Y + a.H > A_ScreenHeight) ? (aY  - a.H  - 20)	: (aY)

		Gui, ACB: Show, % "x" aX " y" aY " NA"

	;}

		AutoCompleteGuiHotkeys("On")

return hACB

ACBMoveDown:            	;{
	ControlSend,, {Down}	, % "ahk_id " hACBLb
	;LBCaretPos 	:= LBEX_GetFocus(hACBLb)
	;LBEX_SetFocus(hACBLb, LBCaretPos + 1)
return
ACBMoveUp:
	ControlSend,, {Up}   	, % "ahk_id " hACBLb
return ;}
ACBChoose:                     ;{

	Gui, ACB: Submit, NoHide
	LBCaretPos 	:= LBEX_GetFocus(hACBLb)
	LBSelItem  	:= LBEX_GetText(hACBLb, LBCaretPos)
	tmpLB := "", ItemExist := false
	Loop, Parse, ACBLb, `|
		If (A_LoopField = LBSelItem) {
			LBEX_SetSel(hACBLb, A_Index, false)
			ItemExist := true
			break
		} else {
			LBEX_SetSel(hACBLb, A_Index, true)
		}

	;~ If !ItemExist
		;~ GuiControl, ACB: Choose, ACBLb, % LBCaretPos

	;SciTEOutput(ACBLb)
	;GuiControl, ACB: Choose, ACBLb,
return ;}
ACBCheck:                   	;{

	CoordMode, Mouse, Screen
	MouseGetPos,,, hWin
	;SciTEOutput(GetHex(hWin) ", " WinGetClass(hWin))
	If (GetDec(hWin) = GetDec(hACB))    ; Mausklick außerhalb des Fensters, schließt die Auswahl
		return

;}
ACBGuiClose:               	;{
ACBGuiEscape:

	Gui, ACB: Destroy
	AutoCompleteGuiHotkeys("Off")
	hACB:= hACBLb:= ""

return ;}
AutoCBLBres:                 	;{ Diagnosenauswahl einfügen

	Gui, ACB: Submit
	MouseClick, Left, % CaretX + 10, % CaretY, 1, 0
	Sleep, 200

	BlockInput, Send

	If (StrLen(ReplaceInput) > 0) {

		Send, % "{LShift Down}{Left " StrLen(ReplaceInput) "}{LShift Up}"
		Sleep 20
		Send, % "{BackSpace}"

	}

	Send, % "{Text}" StrReplace(ACBLb, "|", ";") ";"
	If InStr(ACBLb, "###")
		DiagnoseErgaenzen("###")

	If (StrLen(ReplaceInput) > 0)
		EmptyInputHook()

return ACBLb ;}
}

AutoCompleteGuiHotkeys(status="On") {                                                                                       	;-- Hotkeys für die Nutzereingaben im AutoCompleteGui

	Hotkey, IfWinExist	, AutoDiagnose ahk_class AutoHotkeyGUI
	Hotkey, Enter     	, AutoCBLBres   	, % status
	Hotkey, Down   	, ACBMoveDown	, % status
	Hotkey, Up    		, ACBMoveUp   	, % status
	Hotkey, Space  		, ACBChoose    	, % status
	Hotkey, ~Esc     	, ACBGuiClose  	, % status
	Hotkey, ~Tab    	, ACBGuiClose  	, % status
	Hotkey, ~Home  	, ACBGuiClose  	, % status
	Hotkey, ~End    	, ACBGuiClose  	, % status
	Hotkey, ~Left	   	, ACBGuiClose  	, % status
	Hotkey, ~Right    	, ACBGuiClose   	, % status
	Hotkey, ~LButton 	, ACBCheck       	, % status
	Hotkey, ~RButton 	, ACBGuiClose   	, % status
	Hotkey, IfWinExist

}

DiagnoseErgaenzen(str) {

	BlockInput, Send

	Sleep, 200
	CText    	:= ControlGetText("", "ahk_id " CTX.Hook.HookHwnd)
	strPos    	:= InStr(CText, str)
	If !strPos
		return
	CaretPos 	:= CaretPos(CTX.Hook.HookHwnd)
	moveCrt	:= strPos - CaretPos

	If (moveCrt < 0)
		Send, % "{Left " Abs(moveCrt) "}"
	else if (moveCrt > 0)
		Send, % "{Right " Abs(moveCrt) "}"

	Send, % "{Shift Down}{Right " StrLen(str) "}{Shift Up}"

return
}

DiagnoseUndZiffer(Diagnose, Ziffer, Seitenangabe) {                                                                      	;-- Eintragen von Diagnose und Leistungskomplex

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

RemoveInput(ih, ihRestart=false) {

	SendMode, Event
	BlockInput, Send

	thisInput := ih.Input
	sacHook.Stop()

	Send, % "{LShift Down}{Left " StrLen(thisInput) "}{LShift Up}"
	Sleep 50
	Send, % "{BackSpace}"

	If ihRestart
		sacHook.Start()

}

LoadDiagnosen() {                                                                                                                          	;-- lädt die Daten

    hstrFile := FileOpen(A_ScriptDir "\ICDHotstrings.txt","r").Read()
    idx := -1
    Loop, Parse, hstrFile, `n, `r
    {

        If (StrLen(A_LoopField) = 0)
            continue

        If !RegExMatch(A_LoopField, "\{") {
            abbrev := StrSplit(A_LoopField, "|")
            continue
        }

        diapos := hsDia.Push(A_LoopField)

		;~ For idx, dia in StrSplit(A_LoopField, "|")
			;~ hsICD.Push(dia)

        For idx2, abbreviation in abbrev {

			If hsAbbr.HasKey(abbreviation)
				conflicts .= abbreviation "`n"

			hsAbbr[abbreviation] := diapos

		}

    }

	conflicts := RTrim(conflicts, "`n")
	If StrLen(conflicts) > 0
		MsgBox, %  " ------------- ACHTUNG -------------`n "
						. "Es gibt doppelt belegte Abkürzungen!`n" conflicts

	RegExMatch(A_ScriptDir, "i)(?<Dir>.*)\\Module", data)
	hstrFile := FileOpen(dataDir "\include\daten\icd10gm2019alpha_edvtxt_20181005.txt","r", "CP1252").Read()
    hsICD := StrSplit(hstrFile, "`n", "`r")

}

ConvertDiagnosis() {

	dia := Object()
	dia.hot := Object()
	dia.hot.abbr   	:= Object()
	dia.diagnose	:= Array()

	hstrFile := FileOpen(A_ScriptDir "\ICDHotstrings.txt","r").Read()
    idx := -1
    Loop, Parse, hstrFile, `n, `r
    {

        If (StrLen(A_LoopField) = 0)
            continue

        If !RegExMatch(A_LoopField, "\{") {
            abbrev := StrSplit(A_LoopField, "|")
            continue
        }

        diapos := ICD.code.Push(A_LoopField)
        For idx2, abbreviation in abbrev {

			If ICD.abbrevation.HasKey(abbreviation)
				conflicts .= abbreviation "`n"

			ICD.hotstrings[abbreviation] := diapos

		}

    }

	jsonStr := JSON.Dump(ICD,"", 4)
	save := FileOpen(A_ScriptDir "\ICDHotstrings.json", "w", "UTF-8")
	save.Write(jsonStr)
	save.Close()

return A_ScriptDir "\ICDHotstrings.json"
}

CheckICDHotStrings(aICD) {

	For i, ICDCode in aICD {

		If CTX.Queue.HasKey(ICDCode) {


		} else {

				CTX.Queue[ICDCode] := {"DayFirst":A_DD, "MonthFirst":A_MM, "YearFirst":A_YYYY, "Count":1}


		}


	}


}

;}

;-----------------------------------------------------------------------------------------------------------------------------------------------
; HOOK FUNKTIONEN
;-----------------------------------------------------------------------------------------------------------------------------------------------;{
;----- WINEVENTHOOK
GetActiveFocusHook() {                                                                                                         	;-- Hook der bei jeder Änderung des Eingabefokus aktiv wird

	HookProcAdr   	:= RegisterCallback("ActiveFocusEventProc", "F")
	hWinEventHook	:= SetWinEventHook( 0x8005, 0x8005, 0, HookProcAdr, 0, 0, 0x0003)

return {"HPA": HookProcAdr, "hEvH": hWinEventHook}
}

ActiveFocusEventProc() {

}

AlbisFocusHook(AlbisPID) {                                                                                                      	;-- Hook ausschließlich für EVENT_OBJECT_FOCUS Nachrichten wird gestartet

	; benötigt um Änderung des Eingabefocus innerhalb von Albis zu erkennen
	HookProcAdr   	:= RegisterCallback("AlbisFocusEventProc", "F")
	hWinEventHook	:= SetWinEventHook( 0x8005, 0x8005, 0, HookProcAdr, AlbisPID, 0, 0x0003)

return {"HPA": HookProcAdr, "hEvH": hWinEventHook}
}

AlbisFocusEventProc(hHook, event, hwnd) {                                                                            	;-- behandelt EVENT_OJECT_FOCUS Nachrichten von Albis

	If (GetDec(hwnd) = 0) || waitCheck
		return 0

	FClass := WinGetClass(GetDec(hwnd))
	;SciTEOutput("# 1: " FClass ", fc: " GetHex(hwnd) "   , lasthwnd: " CTX.LastHWND "     , waitcheck: " waitcheck)
	If InStr(FClass, "RichEdit20A") {
		waitcheck := true
		fchwnd := Format("0x{:x}", hwnd)
		If (GetHex(CTX.LastHWND) != fchwnd) {
			SetTimer, moreChecks, -100
			CTX.Hook.ClassF := FClass
			CTX.LastHWND := GetHex(fchwnd)
		}
	}

return 0
}

moreChecks: ;{

	Critical
	hparent	:= GetHex(DllCall("user32\GetAncestor", "Ptr", CTX.LastHWND, "UInt", 1, "Ptr"))
	PClass 	:= Control_GetClassNN(AlbisWinID(), hparent)

	If InStr(CTX.Hook.ClassF, "RichEdit20A") && InStr(PClass, "#327702") {

			context := ControlGetText("Edit3", "ahk_id " hparent)
			If (context = "dia") {

				ControlGet, hsacParent, HWnd,, Edit3, % "ahk_id " hparent
				CTX.Hook.context    		:= "Diagnosen"
				CTX.Hook.WinID        	:= AlbisWinID()
				CTX.Hook.HookHwnd 	:= GetHex(CTX.LastHWND)
				CTX.Hook.hparent      	:= GetHex(hparent)
				CTX.Hook.InfoParent  	:= GetHex(hsacParent)
				CTX.Hook.PID            	:= AlbisPID()

				StartInputHook()

				If CTX.Debug
					SciTEOutput("# 2:   CT: " CTX.Hook.context  "`n`tCF: " CTX.Hook.ClassF "`n`tfc: " CTX.Hook.HookHwnd "`n`tPC: " PClass "   ----> sacHook started" )

			}
			else if RegExMatch(context, "aem|aem|bef|info|mit") {

				If SacHook.InProgress
					StopInputHook("FocusEventProc")

				ControlGet, hsacParent, HWnd,, Edit3, % "ahk_id " hparent
				ControlGetText, starttext,, % "ahk_id " CTX.LastHWND
				CTX.Hook.context    		:= context
				CTX.Hook.WinID        	:= AlbisWinID()
				CTX.Hook.HookHwnd 	:= GetHex(CTX.LastHWND)
				CTX.Hook.hparent      	:= GetHex(hparent)
				CTX.Hook.InfoParent  	:= GetHex(hsacParent)
				CTX.Hook.PID            	:= AlbisPID()
				CTX.Hook.textbuf         	:= starttext

				;InitAutoSuggestion(context)
				;CTX.ASG.IH.Start()

				;ToolTip, % "AutoSuggestion Hook ist an", % A_CaretX - 5, % A_CaretY - 25, 13
				If CTX.Debug
					SciTEOutput("# 2:   CT: " CTX.Hook.context  "`n`tCF: " CTX.Hook.ClassF "`n`tfc: " CTX.Hook.HookHwnd "`n`tPC: " PClass "   ----> sacHook started" )

			}


	} else {

			If SacHook.InProgress
				StopInputHook("FocusEventProc")

	}

	waitcheck      	:= false

return ;}

; -----
SetWinEventHook(eventMin, eventMax, hmodWinEventProc, lpfnWinEventProc, idProcess, idThread, dwFlags) {
	return DllCall("SetWinEventHook", "uint", eventMin, "uint", eventMax, "Uint", hmodWinEventProc	, "uint", lpfnWinEventProc, "uint", idProcess, "uint", idThread, "uint", dwFlags)
}

UnhookWinEvent(hWinEventHook, HookProcAdr) {
	DllCall( "UnhookWinEvent", "Ptr", hWinEventHook )
	DllCall( "GlobalFree", "Ptr", HookProcAdr ) ; free up allocated memory for RegisterCallback
}
;------ SHELLHOOK

class ShellHook {

	;==================================================================
	; SHELL HOOK v1.0.1
	; Author: Daniel Shuy
	; changes by: Ixiko 07.09.2020
	;
	; Attaches a Shell Hook to the Windows API to listen for Shell events


	;{ hookCode Constants (refer to WinUser.h)
	;==================================================================
	; see https://msdn.microsoft.com/en-us/library/windows/desktop/ms644989(v=vs.85).aspx
	static HSHELL_WINDOWCREATED     		:=   1    ; A top-level, unowned window has been created. The window exists when the system calls this hook.
	static HSHELL_WINDOWDESTROYED 		:=   2	; A top-level, unowned window is about to be destroyed. The window still exists when the system calls this hook.
	static HSHELL_ACTIVATESHELLWINDOW	:=   3 	; The shell should activate its main window.
	static HSHELL_WINDOWACTIVATED      	:=   4	; The activation has changed to a different top-level, unowned window.
	static HSHELL_GETMINRECT 		            	:=   5	; A window is being minimized or maximized. The system needs the coordinates of the minimized rectangle for the window.
															            		; If the Shell procedure handles the WM_COMMAND message, it should not call CallNextHookEx.
	static HSHELL_REDRAW 	        		    	:=   6 	; The title of a window in the task bar has been redrawn.
	static HSHELL_TASKMAN 	        		    	:=   7	; The user has selected the task list. A shell application that provides a task list should return TRUE to prevent Windows from starting its task list.
	static HSHELL_LANGUAGE 	    		    	:=   8	; Keyboard language was changed or a new keyboard layout was loaded.
	static HSHELL_SYSMENU 	        		    	:=   9	; A window's system menu is brought up by right clicking its taskbar button
	static HSHELL_ENDTASK 	        		    	:= 10	; A window is forced to exit
	static HSHELL_ACCESSIBILITYSTATE          	:= 11 	; The accessibility state has changed.
	static HSHELL_APPCOMMAND 	    	    	:= 12	; The user completed an input event (for example, pressed an application command button on the mouse
																	    		; or an application command key on the keyboard), and the application did not handle the WM_APPCOMMAND message generated by that input.
	static HSHELL_WINDOWREPLACED     		:= 13	; A top-level window is being replaced. The window exists when the system calls this hook.
	static HSHELL_WINDOWREPLACING     		:= 14 	; A window is replacing the top-level window.
	static HSHELL_MONITORCHANGED     		:= 16	; A window is moved to a different monitor.
	static HSHELL_HIGHBIT                     		 = 0x8000
	static HSHELL_FLASH := HSHELL_REDRAW|HSHELL_HIGHBIT	; A window is flashed.
	static HSHELL_RUDEAPPACTIVATED := HSHELL_WINDOWACTIVATED|HSHELL_HIGHBIT	; Not sure what is the difference from HSHELL_WINDOWACTIVATED.
	;==================================================================
	;}

	__New(hWnd, ih, DeregisterEvents) {

		/* parameters

			hwnd                	:	hwnd of script gui
			ih                     	:	used inputhook in StartInputHook()
			DeregisterEvents	:	array of hex-values for ShellProc() function
											hex-values will be compared with the incoming EventID from Shell Hook

		 */

		If (this.hwnd = hwnd)
			this.Deregister("exists")

		this.hWnd	:= hWnd
		If IsObject(DeregisterEvents)
			this.DeregisterMessages := DeregisterEvents

		this.Register()

	}

	Register() {                                        	; Register Shell Hook
		DllCall("RegisterShellHookWindow", UInt, this.hWnd)
		hook := DllCall("RegisterWindowMessage", Str, "SHELLHOOK")
		OnMessage(hook, ObjBindMethod(this, "ShellProc"))
		OnExit(ObjBindMethod(this, "Deregister"), OnExitConst.ON_EXIT_ADD_AFTER)
	}

	Deregister(eventID="empty") {               	; Deregister Shell Hook

		; parameter eventID is only used for debugging purposes

		DllCall("DeregisterShellHookWindow", UInt, this.hWnd)
		;SciTEOutput("Shellhook closed on eventID: " eventID)
		this.DeregisterMessages := ""
		this.hwnd := ""
	}

	ShellProc(hookEvent, id) {                  	; inspect hook queue

		If IsObject(this.DeregisterMessages) {
			For idx, eventID in this.DeregisterMessages
				If (GetHex(hookEvent) = eventID) {
					If SacHook.InProgress
						StopInputHook(eventID "(ShellProc)")
					break
				}
		}

		return
	}

}

class OnExitConst {

	;==================================================================
	; ONEXIT CONSTANTS v1.0.0
	; Author: Daniel Shuy
	;
	; A library of constants for the OnExit function/command
	;
	; AddRemove Constants
	; see https://autohotkey.com/docs/commands/OnExit.htm
	;==================================================================


	static ON_EXIT_ADD_AFTER   	:= 1 	; Call the function after any previously registered functions
	static ON_EXIT_ADD_BEFORE 	:= -1	; Call the function before any previously registered functions
	static ON_EXIT_REMOVE      		:= 0		; Do not call the function

}

;}

;-----------------------------------------------------------------------------------------------------------------------------------------------
; STRING FUNKTIONEN
;-----------------------------------------------------------------------------------------------------------------------------------------------;{
CapitelWords(str) {                                                                                                                 	;-- finds words with capitel chars at beginning
	words := [], pos := 0
	while, pos:= RegExMatch(str, "([A-Z][\w\p{L}-]+)", word, pos + 1)
		words.Push(word)
return words
}

RateCapitelWords(str) {                                                                                                             	;-- finds words with capitel chars at beginning and rate them based on their sentence position

		words:= Object()
		word	:= StrSplit(str, " ")
		Cnr	:= 0

	; Capital words get a rate of 3 at beginning
		Loop, % word.MaxIndex()
			If RegExMatch(word[A_Index], "([A-Z][\w\p{L}-]+)", word)	{
					Cnr ++
					words[Cnr]:= {"Word": word, "Rate": 3}
			}
			else if RegExMatch(word[A_Index], "([a-z][\w\p{L}-]+)", word)	{
				If RegExMatch(word[A_Index], "(der)|(die)|(das)|(des)|(von)")
					words[Cnr].Rate -= 1
			}

return words
}

TrimDiagnose(str) {                                                                                                                  	;-- entfernt Seitenangaben, ICD Code und ', .G"
	str:= RegExReplace(str, "(\,\s*(li|re|bds)\.)", "")
	str:= RegExReplace(str, "(\,\s*G\.)", "")
	str:= RTrim(str, ";")
	str:= RegExMatch(str, "(\{[\w\.]*\})") ? str : ""
return str
}

RemoveLastWords(lastwords) {                                                                                                	;-- entfernt die letzten Worte aus einem Edit oder RichEdit-Control

		fchwnd		:= GetFocusedControl()
		ctext      	:= ControlGetText(fchwnd)
		words   	:= StrSplit(lastwords, "|")

		; die RichEdit Befehle bringen Albis machmal zum Absturz oder funktionieren nicht
		Loop, % words.MaxIndex()
		{
				CaretPos	:= CaretPos(fchwnd)
				idx			:= 1 + words.MaxIndex() - A_Index
				word	:= words[idx]
				if word = ""
						break
				wordStart	:= InStr(CText, word)
				wordEnd	:= wordstart + StrLen(word)
				goCaret	:= wordEnd - CaretPos
				;ToolTip, % "word: " word "`nwordstart: " wordstart "`nwordend: " wordEnd "`nCaretPos: " CaretPos "`ngoCaret: " goCaret, 800, 500, 5

				If goCaret <> 0
				{
						If goCaret < 0
							Send, % "{Right " (goCaret * -1) "}"
						else
							Send, % "{Left " goCaret "}"
				}

				Send, % "{BackSpace " StrLen(word) +1 "}"
		}
}

TrimPrefix(prefix) {
	prefix := RegExReplace(prefix, "der|des|von|eines|\-", " ")
return LTrim(RegExReplace(prefix, "\s{2,}", " "))
}

;}

;-----------------------------------------------------------------------------------------------------------------------------------------------
; HILFSFUNKTIONEN
;-----------------------------------------------------------------------------------------------------------------------------------------------;{
SciTEOutput(Text:="", Clear=false, LineBreak=true, Exit=false) {                                            	;-- modified version for Addendum für Albis on Windows

	; last change 17.08.2020

	; some variables
		static	LinesOut           	:= 0
			, 	SCI_GETLENGTH	:= 2006
			,	SCI_GOTOPOS		:= 2025

	; gets Scite COM object
		try
			SciObj := ComObjActive("SciTE4AHK.Application")           	;get pointer to active SciTE window
		catch
			return                                                                            	;if not return

	; move Caret to end of output pane to prevent inserting text at random positions
		SendMessage, 2006,,, Scintilla2, % "ahk_id " SciObj.SciteHandle
		endPos := ErrorLevel
		SendMessage, 2025, % endPos,, Scintilla2, % "ahk_id " SciObj.SciteHandle

	; shows count of printed lines in case output pane was erased
		If InStr(Text, "ShowLinesOut") {
			SciObj.Output("SciteOutput function has printed " LinesOut " lines.`n")
			return
		}

	; Clear output window
		If Clear || ((StrLen(Text) = 0) && (LinesOut = 0))
			SendMessage, SciObj.Message(0x111, 420)

	; send text to SciTE output pane
		If (StrLen(Text) != 0) {
			Text .= (LineBreak ? "`r`n": "")
			SciObj.Output(Text)
			LinesOut += StrSplit(Text, "`n", "`r").MaxIndex()
		}

		If Exit {
			MsgBox, 36, Exit App?, Exit Application?                         	;If Exit=1 ask if want to exit application
			IfMsgBox,Yes, ExitApp                                                       	;If Msgbox=yes then Exit the appliciation
		}

}

GetCaretPos() {

	static Size:=8+(A_PtrSize*6)+16, hwndCaret:=8+A_PtrSize*5
	Static CaretX:=8+(A_PtrSize*6), CaretY:=CaretX+4

	VarSetCapacity(Info, Size, 0), NumPut(Size, Info, "Int")
	DllCall("GetGUIThreadInfo", "UInt", 0, "Ptr", &Info), x:=y:=""
	if !(HWND:=NumGet(Info, hwndCaret, "Ptr"))
		return 0

	x:=NumGet(Info, CaretX, "Int"), y:=NumGet(Info, CaretY, "Int")
	VarSetCapacity(pt, 8), NumPut(y, NumPut(x, pt, "Int"), "Int")
	DllCall("ClientToScreen", "Ptr", HWND, "Ptr", &pt)

return {"X":NumGet(pt, 0, "Int"), "Y":NumGet(pt, 4, "Int")}
}

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
return CaretPos
}

LBEX_CalcIdealWidthEx(hLB, Content="", Delim="|", FontOptions="", FontName="", rows=0) {	;-- zum Berechnen der optimalen Breite einer Listbox

		DestroyGui := MaxW := 0
		static SM_CVXSCROLL

		If !SM_CVXSCROLL
			SysGet, SM_CVXSCROLL, 2                                        	; width of a vertical scrollbar

		If !hLB	{

			  If (StrLen(Content) = 0)
				 Return -1

			  Gui, LB_EX_CalcContentWidthGui: New	, % "+Delimiter" Delim
			  Gui, LB_EX_CalcContentWidthGui: Font	, % FontOptions, % FontName
			  Gui, LB_EX_CalcContentWidthGui: Add	, ListBox, % "+HWNDhLB " (rows = 0 ? "" : "r" rows), % Content
			  DestroyGui := True

		}

		ControlGet, Content, List,,, % "ahk_id " hLB                        ; Inhalt ermitteln
		Items := StrSplit(Content, "`n")

		SendMessage, 0x31, 0, 0,, % "ahk_id " hLB                     	; WM_GETFONT
		hFont	:= ErrorLevel

		hDC  	:= DllCall("User32.dll\GetDC", "Ptr", hLB, "UPtr")
		DllCall("Gdi32.dll\SelectObject", "Ptr", hDC, "Ptr", hFont)

		VarSetCapacity(SIZE, 8, 0)
		For Each, Item In Items	{
			DllCall("Gdi32.dll\GetTextExtentPoint32", "Ptr", HDC, "Ptr", &Item, "Int", StrLen(Item), "UIntP", Width)
			MaxW := Width > MaxW ? Width : MaxW
		}

		DllCall("User32.dll\ReleaseDC", "Ptr", hLB, "Ptr", hDC)

	; einfachste Umsetzung um einfach nur herauszufinden ob eine vertikale Scrollbar existiert
		SBVERT_Exist := (Items.MaxIndex() > rows) ? true : false

		If (DestroyGui)
			Gui, LB_EX_CalcContentWidthGui: Destroy

Return MaxW + (SBVERT_Exist = true ? SM_CVXSCROLL : 0) + 8 ; + 8 for the margins
}

LBEX_GetFocus(HLB) {
   Static LB_GETCARETINDEX := 0x019F
   SendMessage, % LB_GETCARETINDEX, 0, 0,, % "ahk_id " HLB
Return (ErrorLevel + 1)
}

LBEX_GetText(HLB, Index) {
   Static LB_GETTEXT := 0x0189
   Len := LBEX_GetTextLen(HLB, Index)
   If (Len = -1)
      Return ""
   VarSetCapacity(Text, Len << !!A_IsUnicode, 0)
   SendMessage, % LB_GETTEXT, % (Index - 1), % &Text, , % "ahk_id " . HLB
   Return StrGet(&Text, Len)
}

LBEX_GetTextLen(HLB, Index) {
   Static LB_GETTEXTLEN := 0x018A
   SendMessage, % LB_GETTEXTLEN, % (Index - 1), 0, , % "ahk_id " . HLB
   Return ErrorLevel
}

LBEX_SetFocus(HLB, Index) {
   Static LB_SETCARETINDEX := 0x019E
   SendMessage, % LB_SETCARETINDEX, % (Index - 1), 0, , % "ahk_id " . HLB
   Return (ErrorLevel + 1)
}

LBEX_SetSel(HLB, Index, Select := True) {

	; ===========================================================
	; SetSel          	Selects an item in a multiple-selection list box and scrolls the item into view, if necessary.
	; Parameters:  	Index    -  	If set to 0, the selection is added to or removed from all items.
	;                    	Select   -  	True to select the items, or False to deselect them.
	;                                   	Default: True - select.
	; Return values:	If the function succeeds, the return value is True; otherwise 0.
	; Remarks:      	Use this function only with multiple-selection list boxes.
	; ===========================================================

   Static LB_SETSEL := 0x0185
   SendMessage, % LB_SETSEL, % Select, % (Index - 1), , % "ahk_id " . HLB
   Return (ErrorLevel + 1)
}

ControlRowInc(ControlID, S:="H", addM:= 5 ) {                                                                      	;-- Gui helper function
	GuiControlGet, c, AHG: Pos, % ControlID
return (S = "H" ? cH + addM : cW + addM)
}

BlockInputShort(time) {                                                                                                             	;-- verhindert nach Zeitvorgabe Tasten- und Mauseingaben
	BlockInput, On
	SetTimer, BlockInputOff, % "-" time
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

GetDec(hwnd) {                                                                                                                    	;-- Umwandlung Hexadezimal nach Dezimal
return Format("{:u}", hwnd)
}

GetHex(hwnd) {                                                                                                                    	;-- Umwandlung Dezimal nach Hexadezimal
return Format("0x{:x}", hwnd)
}

GetClassName(hwnd) {                                                                                                         	;-- returns HWND's class name without its instance number, e.g. "Edit" or "SysListView32"
	VarSetCapacity( buff, 256, 0 )
	DllCall("GetClassName", "uint", hwnd, "str", buff, "int", 255 )
return buff
}

GetClassNN(HW, HC) {                                                                                                         	;-- HW : Window's HWND, HC : Control's HWND
   VarSetCapacity(CL,256,0), HA := 0
   , DllCall("GetClassNameW", "UInt",HC, "Str",CL, "Int",255)
   While HA := DllCall("FindWindowEx", "UInt",HW, "UInt",HA, "UInt",&CL, "UInt",0)
      If (HA = HC)
         Return CL . A_Index
   Return False
}

GetFocusedControl()  {                                                                                                             	;-- retrieves the ahk_id (HWND) of the active window's focused control.

   ; This script requires Windows 98+ or NT 4.0 SP3+.
   guiThreadInfoSize := 8 + 6 * A_PtrSize + 16
   VarSetCapacity(guiThreadInfo, guiThreadInfoSize, 0)
   NumPut(GuiThreadInfoSize, GuiThreadInfo, 0)
   ; DllCall("RtlFillMemory" , "PTR", &guiThreadInfo, "UInt", 1 , "UChar", guiThreadInfoSize)   ; Below 0xFF, one call only is needed
   if (DllCall("GetGUIThreadInfo" , "UInt", 0, "PTR", &guiThreadInfo) = 0) {   ; Foreground thread
			ErrorLevel := A_LastError   ; Failure
			Return 0
   }

Return NumGet(guiThreadInfo, 8+A_PtrSize, "Ptr") ; *(addr + 12) + (*(addr + 13) << 8) +  (*(addr + 14) << 16) + (*(addr + 15) << 24)
}

GetFocusedControlHwnd(hwnd:="A") {
	ControlGetFocus, FocusedControl, % (hwnd = "A") ? "A" : "ahk_id " hwnd
	ControlGet, FocusedControlId, Hwnd,, %FocusedControl%, % (hwnd = "A") ? "A" : "ahk_id " hwnd
return FocusedControlId
}

GetWindowSpot(hWnd) {                                                                                                          	;-- like GetWindowInfo, but faster because it only returns position and sizes
    NumPut(VarSetCapacity(WINDOWINFO, 60, 0), WINDOWINFO)
    DllCall("GetWindowInfo", "Ptr", hWnd, "Ptr", &WINDOWINFO)
    wi := Object()
    wi.X   	:= NumGet(WINDOWINFO, 4	, "Int")
    wi.Y   	:= NumGet(WINDOWINFO, 8	, "Int")
    wi.W  	:= NumGet(WINDOWINFO, 12, "Int") 	- wi.X
    wi.H  	:= NumGet(WINDOWINFO, 16, "Int") 	- wi.Y
    wi.CX	:= NumGet(WINDOWINFO, 20, "Int")
    wi.CY	:= NumGet(WINDOWINFO, 24, "Int")
    wi.CW 	:= NumGet(WINDOWINFO, 28, "Int") 	- wi.CX
    wi.CH  	:= NumGet(WINDOWINFO, 32, "Int") 	- wi.CY
	wi.S   	:= NumGet(WINDOWINFO, 36, "UInt")
    wi.ES 	:= NumGet(WINDOWINFO, 40, "UInt")
	wi.Ac	:= NumGet(WINDOWINFO, 44, "UInt")
    wi.BW	:= NumGet(WINDOWINFO, 48, "UInt")
    wi.BH	:= NumGet(WINDOWINFO, 52, "UInt")
	wi.A    	:= NumGet(WINDOWINFO, 56, "UShort")
    wi.V  	:= NumGet(WINDOWINFO, 58, "UShort")
Return wi
}

SetWindowPos(hWnd, x, y, w, h, hWndInsertAfter := 0, uFlags := 0x40) {                                		;--works better than the internal command WinMove - why?

	; https://docs.microsoft.com/en-us/windows/desktop/api/winuser/nf-winuser-setwindowpos
	SWP_ASYNCWINDOWPOS	:= 0x4000	;This prevents the calling thread from blocking its execution while other threads process the request.
	SWP_DEFERERASE                	:= 0x2000	;Prevents generation of the WM_SYNCPAINT message.
	SWP_DRAWFRAME            	:= 0x0020	;Draws a frame (defined in the window's class description) around the window.
	SWP_FRAMECHANGED     	:= 0x0020	;Applies new frame styles set using the SetWindowLong function.
	SWP_HIDEWINDOW         	:= 0x0080	;Hides the window
	SWP_NOACTIVATE   	        	:= 0x0010	;Does not activate the window.
	SWP_NOCOPYBITS            	:= 0x0100	;Discards the entire contents of the client area.
	SWP_NOMOVE                 	:= 0x0002	;Retains the current position (ignores X and Y parameters).
	SWP_NOOWNERZORDER 	:= 0x0200	;Does not change the owner window's position in the Z order.
	SWP_NOREDRAW             	:= 0x0008	;Does not redraw changes.
	SWP_NOREPOSITION        	:= 0x0200	;Same as the SWP_NOOWNERZORDER flag.
	SWP_NOSENDCHANGING	:= 0x0400	;Prevents the window from receiving the WM_WINDOWPOSCHANGING message.
	SWP_NOSIZE                       	:= 0x0001	;Retains the current size (ignores the cx and cy parameters).
	SWP_NOZORDER              	:= 0x0004	;Retains the current Z order (ignores the hWndInsertAfter parameter).
	SWP_SHOWWINDOW        	:= 0x0040	;Displays the window.

    Return DllCall("SetWindowPos", "Ptr", hWnd, "Ptr", hWndInsertAfter, "Int", x, "Int", y, "Int", w, "Int", h, "UInt", uFlags)
}

Control_GetClassNN(hWnd, hCtrl) {
	; SKAN: www.autohotkey.com/forum/viewtopic.php?t=49471
 WinGet, CH, ControlListHwnd, ahk_id %hWnd%
 WinGet, CN, ControlList, ahk_id %hWnd%
 LF:= "`n",  CH:= LF CH LF, CN:= LF CN LF,  S:= SubStr( CH, 1, InStr( CH, LF hCtrl LF ) )
 StringReplace, S, S,`n,`n, UseErrorLevel
 StringGetPos, LP, CN, `n, L%ErrorLevel%
 Return SubStr( CN, LP+2, InStr( CN, LF, 0, LP+2 ) -LP-2 )
}

ControlGetText(Control="", WinTitle="", WinText="", ExcludeTitle="", ExcludeText="") {                	;-- ControlGetText als Funktion
	ControlGetText, v, %Control%, %WinTitle%, %WinText%, %ExcludeTitle%, %ExcludeText%
	Return v
}

WinGetClass( hwnd ) {                                                                                                            	;-- schnellere Fensterfunktion
	if (hwnd is not Integer)
			hwnd :=GetDec(hwnd)
	VarSetCapacity(sClass, 80, 0)
	DllCall("GetClassNameW", "UInt", hWnd, "Str", sClass, "Int", VarSetCapacity(sClass)+1)
	wclass := sClass
	sClass =
	Return wclass
}

AlbisGetActiveControl(cmd, hwnd=0) {                                                                                     	;-- Contextsensitive Ermittlung des Eingabefokus

	/* 		FUNKTIONSBESCHREIBUNG: AlbisGetActiveControl() letzte Änderung: 20.06.2020
	   *
		* Steuerelemente in Fenster können die gleichen Namen besitzen. Um ein Steuerelement exakt zu identifizieren ermittelt diese Funktion die Eltern(Parent)-Elemente
		* wieviele Eltern ermittelt werden müssen, hängt von der jeweiligen Fensterstruktur ab
		*
		* wozu braucht man das: z.B. um Contextabhängige Hotstrings, Hotkeys oder andere Dateneingaben zu ermöglichen
		*
		* --------------------------------------------------------------------------------------------------------------------------------
		* Bezeichner	|            Ebenenelemente      	|                          Elternelemente						   	   	        	*
		* Bezeichner	|      ClassNN1,ClassNN2,... 	|ParentClass1|ParentClass2			|ParentClass3                	*
		* Dokument	|Edit1,Edit2,Edit3,RichEdit20A1	|#32770		 |AfxFrameOrView	|AfxMDIFrame              	*
		* --------------------------------------------------------------------------------------------------------------------------------
		* 		Bezeichner 			- ihre Bezeichnung für den Steuerelement-Zusammenhang
		*		Ebenenelemente	- eine durch Komma getrennte Liste an ClassNN Namen von Steuerelementen, diese Steuerelemente haben einen inhaltlichen Zusammenhang z.B.
		*		Elternelemente		- eine durch | getrennte Liste an ClassNN Namen von Eltern-Steuerelementen, diese lassen die eindeutige Identifikation der Ebenenelemente zu
		*
		* --------------------------------------------------------------------------------------------------------------------------------
		* WICHTIGES                                                                                                                                        	*
		* --------------------------------------------------------------------------------------------------------------------------------
		* Der Aufruf der Funktion holt immer das Albisfenster in den Vordergrund!

		* - 07.01.2020: AfxFrameOrView und AfxMDIFrame wurden intern in Albis umbenannt, aus AfxFrameOrView90 und AfxMidiFrame90 wurde AfxFrameOrView140 und AfxMidiframe140
		* - 11.07.2020: AfxFrameOrView und AfxMDIFrame wurden intern in Albis umbenannt, aus AfxFrameOrView90 und AfxMidiFrame90 wurde AfxFrameOrView140 und AfxMidiframe140
	   *
	 */

		static Init, Ident
		idx := 2

	; hier kommen noch Identifikationsstrings später hinzu , im Moment brauche ich diese Funktion nur für Eintragungen im Dokumentfenster
		If !Init 		{


				s	:= StrSplit("Dokument|Edit1,Edit2,Edit3,RichEdit20A1|#32770,AfxFrameOrView,AfxMDIFrame", "|")		;Dokument ist meine Bezeichnung für die 4-Controls (Edit1...)
				Ident := Object()
				Ident	:= { "Name" 	: s[1]
							,	"Controls"	: StrSplit(s[2], ",")
							,	"Parents" 	: StrSplit(s[3], ",") }
				Init 	:= true

		}

	; Ermitteln des aktuellen ControlFocus hwnd und classNN
		hFocused	:= !hwnd ? GetFocusedControlHwnd(AlbisWinID()) : hwnd
		cNN	    	:= Control_GetClassNN(AlbisWinID(), hFocused)
		firstparent	:= DllCall("user32\GetAncestor", "Ptr", hFocused, "UInt", 1, "Ptr")
		;SciTEOutput("fc: " GetHex(hFocused) ", cNN: " cNN ", parent: (" GetHex(firstparent) ") " GetClassNN(AlbisWinID(), firstparent))
		;SciTEOutput("fc: " hFocused ", cNN: " cNN ", parent: (" firstparent ") " GetClassNN(AlbisWinID(), firstparent))

		If InStr(cmd, "hwnd")
			return GetHex(hFocused)
		else if InStr(cmd, "classNN")
			return cNN

	; Übereinstimmungen suchen
		matched := false
		For idx, classNN in Ident.Controls
			If InStr(classNN, cNN) {
				 matched := true
				 break
			}

		If matched		{

				aacCounter	:= 0
				CHwnd     	:= hFocused
				IdentsMax  	:= Ident.Parents.MaxIndex()

				Loop % IdentsMax {
					Chwnd := DllCall("user32\GetAncestor", "Ptr", Chwnd, "UInt", 1, "Ptr")
					If InStr(WinGetClass(CHwnd), Ident.Parents[A_Index])
						aacCounter++
				}

			; keine Übereinstimmungskette gefunden
				If (aacCounter != IdentsMax)
					return false

			; ab hier können spezielle weitere Befehle programmiert werden
				If InStr(cmd, "content")     	{

					If InStr(IdentData[1], "Dokument") {

						ControlGetText, RichEdit, RichEdit20A1, % "ahk_id " firstparent
						Loop, 3
							ControlGetText, Edit%A_Index%, % "Edit" A_Index, % "ahk_id " firstparent

						return {"FHwnd":GetHex(hFocused), "FClassNN":cNN, "fpar":firstparent,"Identifier":IdentData[1], "Edit1":Edit1, "Edit2":Edit2, "Edit3":Edit3, "RichEdit":RichEdit}

					}

				}
				else if InStr(cmd, "contraction") 	{

					If InStr(IdentData[1], "Dokument")
						return ControlGetText("Edit3", "ahk_id " firstparent)

				}
				else if InStr(cmd, "identify")     	{
					return Ident.Name "|" cNN
				}

		}

return 0
}

AlbisWinID() {                            			                                                                                	;-- gibt die ID des übergeordneten Albisfenster zurück
return GetHex(WinExist("ahk_class OptoAppClass"))
}

AlbisPID() {                            				                                                                                	;-- gibt die Prozeß-ID des Albisprogrammes aus
	WinGet, AlbisPID, PID, ahk_class OptoAppClass
return AlbisPID
}

class JSON {

	class Load extends JSON.Functor {

		Call(self, ByRef text, reviver:="") 	{

			this.rev := IsObject(reviver) ? reviver : false
			this.keys := this.rev ? {} : false

			static quot := Chr(34), bashq := "\" . quot
			     , json_value := quot . "{[01234567890-tfn"
			     , json_value_or_array_closing := quot . "{[]01234567890-tfn"
			     , object_key_or_object_closing := quot . "}"

			key := ""
			is_key := false
			root := {}
			stack := [root]
			next := json_value
			pos := 0

			while ((ch := SubStr(text, ++pos, 1)) != "") {
				if InStr(" `t`r`n", ch)
					continue
				if !InStr(next, ch, 1)
						this.ParseError(next, text, pos)

				holder := stack[1]
				is_array := holder.IsArray

				if InStr(",:", ch) {
						next := (is_key := !is_array && ch == ",") ? quot : json_value

				} else if InStr("}]", ch) {

						ObjRemoveAt(stack, 1)
						next := stack[1]==root ? "" : stack[1].IsArray ? ",]" : ",}"

				} else {

					if InStr("{[", ch) {

						static json_array := Func("Array").IsBuiltIn || ![].IsArray ? {IsArray: true} : 0

						(ch == "{")
							? ( is_key := true
							  , value := {}
							  , next := object_key_or_object_closing )
						; ch == "["
							: ( value := json_array ? new json_array : []
							  , next := json_value_or_array_closing )

						ObjInsertAt(stack, 1, value)

						if (this.keys)
							this.keys[value] := []

					} else {

						if (ch == quot) {
							i := pos
							while (i := InStr(text, quot,, i+1)) {
								value := StrReplace(SubStr(text, pos+1, i-pos-1), "\\", "\u005c")

								static tail := A_AhkVersion<"2" ? 0 : -1
								if (SubStr(value, tail) != "\")
									break
							}

							if (!i)
								this.ParseError("'", text, pos)

							  value := StrReplace(value,  "\/",  "/")
							, value := StrReplace(value, bashq, quot)
							, value := StrReplace(value,  "\b", "`b")
							, value := StrReplace(value,  "\f", "`f")
							, value := StrReplace(value,  "\n", "`n")
							, value := StrReplace(value,  "\r", "`r")
							, value := StrReplace(value,  "\t", "`t")

							pos := i ; update pos

							i := 0
							while (i := InStr(value, "\",, i+1)) {
								if !(SubStr(value, i+1, 1) == "u")
									this.ParseError("\", text, pos - StrLen(SubStr(value, i+1)))

								uffff := Abs("0x" . SubStr(value, i+2, 4))
								if (A_IsUnicode || uffff < 0x100)
									value := SubStr(value, 1, i-1) . Chr(uffff) . SubStr(value, i+6)
							}

							if (is_key) {
								key := value, next := ":"
								continue
							}

						} else {
							value := SubStr(text, pos, i := RegExMatch(text, "[\]\},\s]|$",, pos)-pos)

							static number := "number", integer :="integer"
							if value is %number%
							{
								if value is %integer%
									value += 0
							}
							else if (value == "true" || value == "false")
								value := %value% + 0
							else if (value == "null")
								value := ""
							else
								this.ParseError(next, text, pos, i)

							pos += i-1
						}

						next := holder==root ? "" : is_array ? ",]" : ",}"
					}

					is_array? key := ObjPush(holder, value) : holder[key] := value

					if (this.keys && this.keys.HasKey(holder))
						this.keys[holder].Push(key)
				}

			}

			return this.rev ? this.Walk(root, "") : root[""]
		}

		ParseError(expect, ByRef text, pos, len:=1) {

			static quot := Chr(34), qurly := quot . "}"

			line := StrSplit(SubStr(text, 1, pos), "`n", "`r").Length()
			col := pos - InStr(text, "`n",, -(StrLen(text)-pos+1))
			msg := Format("{1}`n`nLine:`t{2}`nCol:`t{3}`nChar:`t{4}"
			,     (expect == "")     ? "Extra data"
			    : (expect == "'")    ? "Unterminated string starting at"
			    : (expect == "\")    ? "Invalid \escape"
			    : (expect == ":")    ? "Expecting ':' delimiter"
			    : (expect == quot)   ? "Expecting object key enclosed in double quotes"
			    : (expect == qurly)  ? "Expecting object key enclosed in double quotes or object closing '}'"
			    : (expect == ",}")   ? "Expecting ',' delimiter or object closing '}'"
			    : (expect == ",]")   ? "Expecting ',' delimiter or array closing ']'"
			    : InStr(expect, "]") ? "Expecting JSON value or array closing ']'"
			    :                      "Expecting JSON value(string, number, true, false, null, object or array)"
			, line, col, pos)

			static offset := A_AhkVersion<"2" ? -3 : -4
			throw Exception(msg, offset, SubStr(text, pos, len))
		}

		Walk(holder, key) {

			value := holder[key]
			if IsObject(value) {
				for i, k in this.keys[value] {
					; check if ObjHasKey(value, k) ??
					v := this.Walk(value, k)
					if (v != JSON.Undefined)
						value[k] := v
					else
						ObjDelete(value, k)
				}
			}

			return this.rev.Call(holder, key, value)
		}
	}

	class Dump extends JSON.Functor {

		Call(self, value, replacer:="", space:="") {

			this.rep := IsObject(replacer) ? replacer : ""

			this.gap := ""
			if (space) {
				static integer := "integer"
				if space is %integer%
					Loop, % ((n := Abs(space))>10 ? 10 : n)
						this.gap .= " "
				else
					this.gap := SubStr(space, 1, 10)

				this.indent := "`n"
			}

			return this.Str({"": value}, "")
		}

		Str(holder, key) {

			value := holder[key]

			if (this.rep)
				value := this.rep.Call(holder, key, ObjHasKey(holder, key) ? value : JSON.Undefined)

			if IsObject(value) {
				static type := A_AhkVersion<"2" ? "" : Func("Type")
				if (type ? type.Call(value) == "Object" : ObjGetCapacity(value) != "") {
					if (this.gap) {
						stepback := this.indent
						this.indent .= this.gap
					}

					is_array := value.IsArray
					if (!is_array) {
						for i in value
							is_array := i == A_Index
						until !is_array
					}

					str := ""
					if (is_array) {
								Loop, % value.Length() {
									if (this.gap)
										str .= this.indent

									v := this.Str(value, A_Index)
									str .= (v != "") ? v . "," : "null,"
								}

					} else {

						colon := this.gap ? ": " : ":"
						for k in value {
							v := this.Str(value, k)
							if (v != "") {
								if (this.gap)
									str .= this.indent

								str .= this.Quote(k) . colon . v . ","
							}
						}
					}

					if (str != "") {
						str := RTrim(str, ",")
						if (this.gap)
							str .= stepback
											}

					if (this.gap)
						this.indent := stepback

					return is_array ? "[" . str . "]" : "{" . str . "}"

				}

			} else ; is_number ? value : "value"
				return ObjGetCapacity([value], 1)=="" ? value : this.Quote(value)
		}

		Quote(string) {

			static quot := Chr(34), bashq := "\" . quot

			if (string != "") {
				  string := StrReplace(string,  "\",  "\\")
				; , string := StrReplace(string,  "/",  "\/") ; optional in ECMAScript
				, string := StrReplace(string, quot, bashq)
				, string := StrReplace(string, "`b",  "\b")
				, string := StrReplace(string, "`f",  "\f")
				, string := StrReplace(string, "`n",  "\n")
				, string := StrReplace(string, "`r",  "\r")
				, string := StrReplace(string, "`t",  "\t")

				static rx_escapable := A_AhkVersion<"2" ? "O)[^\x20-\x7e]" : "[^\x20-\x7e]"
				while RegExMatch(string, rx_escapable, m)
					string := StrReplace(string, m.Value, Format("\u{1:04x}", Ord(m.Value)))

			}

			return quot . string . quot
		}
	}

	Undefined[] {

		get {
			static empty := {}, vt_empty := ComObject(0, &empty, 1)
			return vt_empty
		}
	}

	class Functor {

		__Call(method, ByRef arg, args*) {
			if IsObject(method)
				return (new this).Call(method, arg, args*)
			else if (method == "")
				return (new this).Call(arg, args*)
		}
	}
}

;}

;{ ICON / EXITFUNKTION
Create_ConTextExpander_ICO() {

DecLen := 0
pBitmap := 0
VarSetCapacity(Base64, 2168 << !!A_IsUnicode)
B64 := "iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAACXBIWXMAAARDAAAEQwFLu/5vAAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAABeVJREFUaIHtWllsVFUY/s65dzboDKXtTGnZZqYYiiEhQBdMBEqpD2oibQWMUawRbUV8gBpx1yY+mEAthETZhLCEEFtDC6lRFGkRImUtDw2C2s40lW5T7DIwM5259xwfoLXtLLfT1i6E7+ne83//d/7v3nvuOXchCICCYq42dCMTMstiXE4lRIijlGoDcf8vMMbc4HITpcJFQmipzo2yvDziG8gjAxu2H+KrGZe/pFSY5Witk+z2a2JXlwM+n2d0Kn8AlUoLvcEIi3mRZDRZRZnL9RTC5vwcUtqX12ugoIDTSCu2cuCd2r+qWMWZfbStzT6qRQeD0WhBWvobLCEhlXKCbV21eL+ggDCgj4Edh3kh4yz/l9O7yJXLx8eu2hBISsnGypUbOCgtzF9HtgAPDGw/xFeDoOT0z19hvBbfg6SUbGRkbATnyM7PIaWkoJir9S75z7rayzO+K/mIhiMWFTUdy5avh1anH1IxHrcTZyv3o739dlh5a9Z+wcyWhQ16r+ox0dCNTBA6q7JiX9gFPP3sFsycOT/svL6YHBGFo0c2hZVTWbGXrp/zzWyXFqsoYSy71VHnG8qAjYyc1m//p1M7saMoE81Nt/y4zc1/YEdRJk79uHOARnzY/TocNjgcdRIHy6KMSan19mpV2CoB4JO64fE4wZjsF2NMhsfjhCSNzO3Ybq8WZSY/QUHEWGeXY0RERxPOLgcoFUyUUqob7UlqJOD1ukEJnSwGCur10YiOnh0w8c6dejidd0KKl5dvg1rVf+XhHcRB0uuNiI6eGVa/AQ0sWPAMpsUnos1h94t5fR5FAz6vC4z1X7bIkt8yxg+RkbEwWxb7tccYzWhuvInz54/4xQIaACG4eaMCNTWnFTsNhKzszxA//fF+bY2Nv+PwwbdD5jU01KChocavff78DERODXy3CmviGo94ZGCsMeENBB7Eg4QkdffbX7TwOSRYUzA1aoYfNypqBjKzPsWUKf2XHwM1wsWwzsClqmIwxnr34+LnInHecuh0Bj+uVqtH4rzliIuf29vGGMPlSyXDKWF4Z6C6uhw3b52DRjNpSPnd3S64XZ3DKWF4BgDA7eocdhHDwcM/iFdlfQJLgOm9B1evluHc2YMAAIPBhNde3xtS78D+PHR1tgAAnlz2KpKSMoNy7XVXUVb2eUg9RQNqlRZabfBHRlHQ9G4TQkNyAYCS/97kqER1SL5Ko1Mqb+JfQo8MjDUUx8Dde+3oaG8KGvd4nL3bjMshuQAg93ledrudIfn37v6jVJ6ygR++L1QU6YGzy4Hdu14eNL/qwjFUXTg2aH4gPPyXkFqtA6XBaZLsheS7vyAjhECjiQip1919F5zz+52rNBAFdVAuYxK8XndIPeWJLPNjJMxZEjR+saoYFWf2AAAMhlhs2Hg0pN7ur19CR0czAGDp0hykLnkhKLe29iJKvv0wpN6Ev4QeGRhrTHgDioO4paUWghj8ThHuu/3+uY2w268Fjbc21ypqiJwzl1qlDfpI9evZA0OrbhC4Xl2O69XlQ8rVaCaBMfmeyJjUGmEwmgcS4uITIcnKrwP7QqeLVORYE1LhcnWEpRsXnwi3u6tfW4Q+BpyzFlGg4gWLZfF0AL3fCG7/fQNzE5fCbF4UVkcAFI+oyWQNW7Onpr6wmJMkSoXfREJoqdFoedFotMDhsAEAbLYrsNmuDKmj0YDJZEWMcbZICMpohxYnZC7Xp63IZcqp4wNp6blMZrJN58JJWrCWeAUi5CfMSaHJyc+PdW2KSEldA6s1mVIibM7LIz4KAJtfIcfBUZie8SYfzyZSUlZjRXouB7A1P4ecAPrMA502vGdIoHzlU2+9a7YkscqKvbRnTIw1TCYr0tJzmdWaTAFs66zDBz0xv589ig7xLM7lIkoFs6PN5rPbrqmcna2D+kQ0klCrtDAYYmG2LJRijBZRZrJNFIRNm9aRk315fgYAYM8ernJpsYqDZclMTqVEiKeUKr/jGEFwzlwyk5uoIFQJhJbqXDgZ6HebfwF3/iuO6A270AAAAABJRU5ErkJggg=="


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
;----------------------------------------------
HotkeyEnde() {

	static HotkeyEnde := false

	If HotkeyEnde
		DasEnde("Hotkey für Beendung des Skriptes durch User", 0)
	else
		HotkeyEnde := true

}
DasEnde(ExitReason, ExitCode) {

	OnExit("DasEnde", 0)

	hook := Albis.Hook["Fokus"]
	UnhookWinEvent(hook.hEvH, hook.HPA)

ExitApp
}
;}


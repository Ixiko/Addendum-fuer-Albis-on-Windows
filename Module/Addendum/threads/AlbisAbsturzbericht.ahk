﻿; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                 	Addendum: AlbisAbsturzbericht.ahk
;
;
;      Funktion:           	Zusatzskript für Addendum.ahk, das nach Abfangen des Albis Absturzberichtfensters ausgeführt wird und den Dialog ohne Nutzerinteraktion
;									behandeln kann.
;									Ein Aufruf des Skriptes mit Doppelklick oder als parameterloser cmdline Befehl zeigt eine Gui zum Ändern der Einstellungen an.
;
;
;		Abhängigkeiten:	siehe includes
;
;	                    			Addendum für Albis on Windows
;                        			by Ixiko started in September 2017 - last change 15.01.2022 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


;{  Skripteinstellungen
#NoEnv
#Persistent
#KeyHistory, Off
SetBatchLines, -1
FileEncoding, UTF-8
DetectHiddenText, On
DetectHiddenWindows, On
SetTitleMatchMode, RegEx
SetTitleMatchMode, slow
AutoTrim, Off
CoordMode, ToolTip, Screen
;}

;{  Variablen
	global q                     	:= Chr(0x22)                           	; ist dieses Zeichen -> "
	global Addendum      	:= Object()
	global AddendumDir
	global aabProps
	global compname := StrReplace(A_ComputerName, "-")


	If !RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir) {
		RegExMatch(A_ScriptDir, ".*(?=\\Addendum\sfür\sAlbisOnWindows)", AddendumDir)
		AddendumDir .= "\Addendum für AlbisOnWindows"
	}
	Addendum.Dir               	:= AddendumDir
	Addendum.DBPath 			:= AddendumDir "\logs'n'data\_DB"

	aabProps := {"restart"    	: IniRead(A_ScriptDir "\AlbisAbsturzbericht.ini", "Properties", "Albis_neustarten", 1)
						, "anonym"	: IniRead(A_ScriptDir "\AlbisAbsturzbericht.ini", "Properties", "Anonym_senden", 0)
						, "remark"  	: IniRead(A_ScriptDir "\AlbisAbsturzbericht.ini", "Properties", "Problembeschreibung")
						, "actionnr"   	: IniRead(A_ScriptDir "\AlbisAbsturzbericht.ini", "Properties", "Aktion_ausfuehren", 1)
						, "sendmail"  	: IniRead(A_ScriptDir "\AlbisAbsturzbericht.ini", "Properties", "EMail_Benachrichtigung", 1)
						, "email"    	: IniRead(A_ScriptDir "\AlbisAbsturzbericht.ini", "Properties", "Email_Adresse")}
;}


  ; Skript zeigt bei parameterlosem Aufruf eine Gui für Einstellungen an
	If !A_Args.Count()
		AlbisAbsturzberichtEinstellungen()
	else
		AlbisAbsturzberichtSenden(A_Args*)

return

^!ö::
	Reload
return
^Esc::ExitApp


AlbisAbsturzberichtEinstellungen()                       	{

	; Generated by AutoGUI 2.5.4
	global

	local Index, Text, opt, opt1, opt2, opt3, cp, cpX, cpY, cpW, cpH, bgt, gL, guiH, Y
	static aabText, remarks, actionNr, hwnd, hEMail, hClick, DefaultMail, hprogress1, hprogress2
	static newactionNr, newaction, email

	gL 	:= " gaabHandler"
	guiH := 440
	bgt 	:= " BackgroundTrans "
	aabText =
	(LTrim
	ALBIS(TM) hat ein Problem festgestellt und muss beendet werden.|
	Falls Sie Ihre Arbeit noch nicht gespeichert hatten, können Daten möglicherweise verloren gegangen sein.|
	Programm neu starten|
	Dieses Problem bitte an ALBIS und auch an Ihren ALBIS Servicepartner berichten.|
	Ihre Problemmeldung wird dazu beitragen, die Qualität von ALBIS zu verbessern. Zur optimalen Problemanalyse
	ist eine kurze Beschreibung des Problems empfehlenswert.|
	Sie können Ihre Daten auch anonym versenden, allerdings ist dann keine Rückmeldung an Sie und keine
	Problembehebung durch Ihren zuständigen Albis Servicepartner möglich.|
	Um zu sehen, welche Daten Ihr Bericht enthält,|
	klicken Sie hier.|
	Anonym senden|
	Bemerkung / Problembeschreibung (freiwillig)|
	&Problembericht senden|
	&Speichern unter|
	&Beenden|
	Einstellungen - Albis Absturzbericht|
	Nehmen Sie hier Einstellungen für den Versand des Absturzberichtes vor.
	Nutzen Sie dazu links das geklonte Dialogfenster.|
	☑ Programm neu starten|
	⇦|Albis wird nach Berichterstellung / Beendigung des Dialoges:|neu gestartet|nicht neu gestartet|
	☑ Anonym senden|der Bericht wird anonym versendet:|ja|nein|
	geben Sie Ihren Text (optional) im Eingabefeld (links) ein. z.B:|
	Keine Angabe von Gründen möglich, da der Absturz nach rund $TimeIdlePhysical# min in unbeaufsichtigem
	Zustand erfolgt war (unbeaufsichtigt = letzte physisch feststellbare Eingabe).|
	🆗 Problembericht senden / Speichern unter / Beenden|
	Klicken Sie auf einen der Buttons um die abschließende Aktion für diesen Dialog auszuwählen.
	Aktion gewählt: |Problembericht senden|Speichern unter|Beenden|
	@|Benachrichtigung per E-Mail|E-Mailadresse eintragen wenn Sie selbst benachrichtigt werden wollen|
	)

	aabText := StrSplit(StrReplace(aabText, "`n", " "), "|")
	For Index, Text in aabText
		aabText[Index] := LTrim(aabText[Index], A_Space)


	;-: GUI Clone 	;{
	Gui, aab2: new, hwndhaab

	Gui aab2: Add, Picture  	, % "x390 y10 w32 h32 0x6 0x5000000E                           "	, % Addendum.Dir "\Module\Addendum\res\AlbisTM.jpg"

	Gui aab2: Font, s8 bold q5, MS Shell Dlg 2
	Gui aab2: Add, Text        	, % "x14 	y10 w365 h33 0x50020000                              "	, % aabText.1
	Gui aab2: Font, Normal
	Gui aab2: Add, Text        	, % "x30 	y63 w383 h36 0x50020000                              "	, % aabText.2
	opt := "0x50010003 vaabCB1 " gL
	Gui aab2: Add, CheckBox	, % "x30 	y96 w131 h16 +Checked  " opt                         	, % aabText.3
	Gui aab2: Font, bold
	Gui aab2: Add, Text        	, % "x30 	y119 w383 h31 0x50020000                            "	, % aabText.4
	Gui aab2: Font, Normal
	Gui aab2: Add, Text        	, % "x30 	y153 w383 h44 0x50020000                            "	, % aabText.5
	Gui aab2: Add, Text        	, % "x30 	y198 w383 h42 0x50020000                            "	, % aabText.6
	Gui aab2: Add, Text        	, % "x30 	y247 w228 h13 0x50020000                            "	, % aabText.7
	Gui aab2: Add, Text        	, % "x278 	y247 w75 h13 0x50020100                               "	, % aabText.8
	opt := " vaabCB2 " gL
	Gui aab2: Add, CheckBox	, % "x30 	y267 w101 h16 0x50010003 " opt                    	, % aabText.9
	Gui aab2: Add, Text        	, % "x30 	y288 w216 h13 0x50020000                            "	, % aabText.10
	opt := " +Multi 0x50211004 "
	Gui aab2: Add, Edit        	, % "x30 	y307 w383 h42 " opt " vaabremark                   "	, % aabProps.remark

	opt1 := " vaabBtn1 hwndaabHBtn1"
	opt2 := " vaabBtn2 hwndaabHBtn2"
	opt3 := " vaabBtn3 hwndaabHBtn3"
	Gui aab2: Add, Button     	, % "x30 	y361 w123 h26 0x50010000 " opt1 gL              	, % aabText.11
	Gui aab2: Add, Button     	, % "x161 	y361 w123 h26 0x50010000 " opt2 gL              	, % aabText.12
	Gui aab2: Add, Button     	, % "x290 	y361 w123 h26 0x50010000 " opt3 gL              	, % aabText.13

	;}

	;-: Hinweise 		;{
	Gui aab2: Add, Progress  	, % "x432 	y0 	w5 h" guiH-10 " CE0E0E0  hwndhprogress1" 	, 100

	Gui aab2: Font, s9 bold, Arial
	Gui aab2: Add, Text        	, % "x450	y8 w400 Center " bgt                                         	, % aabText.15

	Gui aab2: Font, s9 Normal, Arial
	Gui aab2: Add, Text        	, % "x450	y54 w400 " bgt                                          	    	, % aabText.16     	; Programm neu starten
	Gui aab2: Add, Text        	, % "x450	y+0 w15   " bgt "  hwndhClick"                         		, % aabText.17        	; ⇦
	Gui aab2: Add, Text        	, % "x465	yp w385 " bgt                                          	        	, % aabText.18
	Gui aab2: Font, bold
	Gui aab2: Add, Text        	, % "x465	y+0 w385 " bgt " vaabChange1"                           	, % aabText.19       	; neu gestartet

	Gui aab2: Font, s10 Normal
	Gui aab2: Add, Text        	, % "x450	y+10 w400 " bgt                                          	    	, % aabText.21    		; Anonym senden
	Gui aab2: Font, s9 Normal
	Gui aab2: Add, Text        	, % "x450	y+0 w15     " bgt                                                 	, % aabText.17        	; ⇦
	Gui aab2: Add, Text        	, % "x465	yp               " bgt                                              		, % aabText.22       	; der Bericht ...
	Gui aab2: Font, bold
	Gui aab2: Add, Text        	, % "x465	y+0            " bgt " vaabChange2"                     		, % aabText.24       	; ja/nein (23/24)

	Gui aab2: Font, Normal
	Gui aab2: Font, s10 Normal
	Gui aab2: Add, Text        	, % "x450	y+10 w390 " bgt                                             	 	, % "🔤 " aabText.10	; Bemerkung
	Gui aab2: Font, s9 Normal
	Gui aab2: Add, Text        	, % "x450	y+0 w15     " bgt                                              		, % aabText.17        	; ⇦
	Gui aab2: Add, Text        	, % "x465	yp w390      " bgt                                              		, % aabText.25       	;
	Gui aab2: Font, s8 Normal
	Gui aab2: Add, Text        	, % "x465 	y+3 w390   " bgt                                              		, % aabText.26       	; Editfeld

	Gui aab2: Font, s10 Normal
	Gui aab2: Add, Text        	, % "x450	y+10 w390 " bgt                                             	 	, % aabText.27    		; Buttons
	Gui aab2: Font, s9 Normal
	Gui aab2: Add, Text        	, % "x450	y+0 w15     " bgt                                          	    	, % aabText.17        	; ⇦
	Gui aab2: Add, Text        	, % "x465	yp w385      " bgt "  vaabChoosed"                    		, % aabText.28       	; gewählt

	GuiControlGet, cp, aab2: Pos, aabChoosed
	Y:=cpY+2, opt := "+Border Center vaabChange3"

	Gui aab2: Font, s9 bold
	Gui aab2: Add, Progress  	, % "x465 	y+2	w150 h18 CE0E0E0  hwndhprogress2    	" 	, 100                       	; grau hinterlegt wie ein Button
	Gui aab2: Add, Text        	, % "x465	yp  w150 " bgt "    " opt                                         	, % aabText.29       	; Aktion ausgewählt

	Gui aab2: Font, s8 Normal
	Gui aab2: Add, Text        	, % "x450	y+10 w13 h18  " bgt " 0x50800000	"                	, % aabText.32    		; Benachrichtigung per Mail
	Gui aab2: Font, s10 Normal
	Gui aab2: Add, Text        	, % "x465	yp+1 w385      " bgt                                         	 	, % aabText.33       	;
	Gui aab2: Font, s8 Normal
	Gui aab2: Add, Checkbox 	, % "x465	y+2               " bgt " vaabSendMail" gL           	, % aabText.34       	;
	Gui aab2: Font, s9 Normal
	Gui aab2: Add, Edit       	, % "x465	y+2 w386    " bgt " vaabEMail hwndhEMail"     		, % email

	;}

	;-: Änderungen	;{

	; Gui Buttons - Speichern und Beenden ohne Speichern
	Gui aab2: Font, s9 Normal
	Gui aab2: Add, Button    	, % "x465	y+20 	 vaabPropsSave " gL                        	, % "Einstellungen speichern"
	Gui aab2: Add, Button    	, % "x+5	yp     	vaabPropsExit " gL                          	, % "Beenden ohne Speichern"

	; Beenden Button rechts ausrichten
	GuiControlGet, cp, aab2: Pos, aabPropsExit
	GuiControl, aab2: Move, aabPropsExit, % "x" 870-cpW-20

	; Edit Banner setzen
	DefaultMail := "name@mailadress.org"
	SendMessage, 0x1501, 1, &DefaultMail,, % "ahk_id " hEMail

	; letzte Einstellungen wiederherstellen
	GuiControl, aab2: , aabSendMail	, % aabProps.sendmail
	GuiControl, aab2: , aabEMail    	, % aabProps.email
	GuiControl, % "aab2: Enable" aabProps.sendmail, aabEMail

	; Aktions-Button gedrückt erscheinen lassen
	hwnd := "aabHBtn" (actionNr := aabProps.actionNr)
	hwnd := %hwnd%
	WinSet, ExStyle, 0x200, % "ahk_id " hwnd
	WinSet, Style, 0x58010F00, % "ahk_id " aabHBtn2
	Loop 3
		If (A_Index <> 2) {
			hwnd := "aabHBtn" A_Index
			hwnd := %hwnd%
			WinSet, Style, 0x50010300, % "ahk_id " hwnd
		}


	;}

	Gui aab2: Show, w870 h%guiH%, % "ALBIS(TM) (" aabText.14 ")"


Return

aabHandler:

	Critical

	Gui, aab2: Default
	Gui, aab2: Submit, NoHide

	If RegExMatch(A_GuiControl, "i)aabBtn(?<Nr>\d)", newaction) {

		actionNr := newactionNr=2 ? actionNr : newActionNr
		Loop 3 {
			hwnd := "aabHBtn" A_Index
			hwnd := %hwnd%
			If (A_Index <> 2) {
				WinSet, Style, 0x50010300, % "ahk_id " hwnd                                          	; Steuerelement Style einmal ändern
				WinSet, ExStyle, % (actionNr=A_Index ? 0x200 : 0x0), % "ahk_id " hwnd  	; ExStyle auf ClientEdge ändern
				WinSet, Style, 0x50010F00, % "ahk_id " hwnd                                          	; Style zurücksetzen, jetzt erst wird der Button neu gezeichnet
			}
		}

		ControlClick,, % "ahk_id " hEMail  ,, Left                                                      	; entfernt den Fokusrahmen aus dem Button
		ControlClick,, % "ahk_id " hClick  ,, Left                                                      	; entfernt den Fokusrahmen aus dem Button
		GuiControl, aab2: , aabChange3, % aabText[28+actionNr]

	}
	else if (A_GuiControl = "aabCB1") {
		GuiControl, aab2: Text, aabChange1, % (aabCB1 ? aabText.19 : aabText.20)
	}
	else if (A_GuiControl = "aabCB2") {
		GuiControl, aab2: Text, aabChange2, % (aabCB2 ? aabText.23 : aabText.24)
	}
	else if (A_GuiControl = "aabSendMail") {
		GuiControl, % "aab2: Enable" aabSendMail, aabEMail
	}
	else if (A_GuiControl = "aabPropsSave") {
		aabremarks := RegExReplace(aabremarks, "[\n]+", "\n")
		aabremarks := RegExReplace(aabremarks, "[\r]+", "\r")
		IniWrite, % aabCB1        	, % A_ScriptDir "\AlbisAbsturzbericht.ini", Properties, Albis_neustarten
		IniWrite, % aabCB2       	, % A_ScriptDir "\AlbisAbsturzbericht.ini", Properties, Anonym_senden
		IniWrite, % aabremark   	, % A_ScriptDir "\AlbisAbsturzbericht.ini", Properties, Problembeschreibung
		IniWrite, % actionNr      	, % A_ScriptDir "\AlbisAbsturzbericht.ini", Properties, Aktion_ausfuehren
		IniWrite, % aabSendMail 	, % A_ScriptDir "\AlbisAbsturzbericht.ini", Properties, EMailbenachrichtigung
		IniWrite, % aabEMail        	, % A_ScriptDir "\AlbisAbsturzbericht.ini", Properties, Email_Adresse
		aabProps := {"restart"    	: aabCB1
							, "anonym"	: aabCB2
							, "remark"  	: aabremark
							, "action"   	: actionNr
							, "sendmail"  	: aabSendMail
							, "email"     	: aabEMail}
		ExitApp
	}
	else if (A_GuiControl = "aabPropsExit")
		ExitApp



return

aab2GuiEscape:
aab2GuiClose:

ExitApp
return


}

AlbisAbsturzberichtSenden(params*)                     	{

	crashsend := {"anonym"     	: "Button1"
						, "senden"      	: "Button2"
						, "speichern"    	: "Button3"
						, "beenden"    	: "Button4"
						, "starten"      	: "Button5"
						, "DSVGO"       	: "static8"
						, "Bemerkung"  	: "Edit1"}

	hWin     	:= params.1
	crashtime	:=params.2
	idletime	:=params.3

	ControlSetText, % crashsend.Bemerkung, % aabProps.remark, % "ahk_id " hWin


}


IniRead(filepath, Section, Key:="", Default:="")    	{
	IniRead, inival, % filepath, % Section, % Key, % Default
return !InStr(inival, "ERROR") ? inival : ""
}

;{ INCLUDES
#Include %A_ScriptDir%\..\..\..\include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Window.ahk

#Include %A_ScriptDir%\..\..\..\lib\Gdip_All.ahk
#Include %A_ScriptDir%\..\..\..\lib\SciteOutput.ahk
#Include %A_ScriptDir%\..\..\..\lib\RemoteBuf.ahk

;}


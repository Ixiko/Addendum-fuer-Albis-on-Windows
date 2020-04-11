#NoEnv
SetBatchLines, -1
FileEncoding, UTF-8
global AddendumDir		:= FileOpen("C:\albiswin.loc\AddendumDir","r").Read()

<<<<<<< Updated upstream
gosub Thread20
SetTimer, Thread20, 2000

=======
gosub Thread41
;SetTimer, Thread20, 2000

;AlbisActivate(1)

;q::gosub Thread28
return
ExitApp

^!ö::
Reload
return

Thread44:                                                                 	;{

return ;}

Thread43:                                                                 	;{

return ;}

Thread42:                                                                 	;{

return ;}

Thread41:                                                                 	;{

	AlbisRezeptAutIdemTest()

return

AlbisRezeptAutIdemTest() {                                                                                            	;-- setzt automatisch ein 'aut idem' - Häkchen

	; !!Voraussetzung!! - In den Dauermedikamenten ist im letzten Dosisfeld (Nacht) ein 'A' vermerkt
		static Dosisfeld  	:= [6,12,18,24,30,36]  	; Edit
		static RezeptFeld	:= [2,8,14,20,26,32]    	; Edit
		static Idemfeld   	:= [52,53,54,55,56,57]	; Button
		static f, blinks, hdc, hbm, obm, pGraphics, hMuster16

		global hRezeptOverlay, AutIdem

		Felder       	:= Object()
		hMuster16	:= GetHex(WinExist("Muster 16 ahk_class #32770"))
		RPos         	:= GetWindowSpot(hMuster16)
		hActive      	:= GetHex(WinActive("A"))
		f                	:= 1
		blinks        	:= 0

	/* Untersucht die Steuerelemente im Rezeptformular, setzt bei Bedarf das Aut-Idem Häkchen
		; die Aut-Idem Steuerelemente im Rezeptfenster sind keine Standard-Checkboxen, mit dem Control-Befehl läßt sich der Status nicht auslesen
		; es wird deshalb mit SendMessage gearbeitet
	*/
		Loop, % Dosisfeld.MaxIndex() {

				ControlGetText, Nacht       	, % "Edit"    Dosisfeld[A_Index], % "ahk_id " hMuster16
				ControlGetText, Medikament	, % "Edit"  Rezeptfeld[A_Index], % "ahk_id " hMuster16
				ControlGet, hAutIdem, hwnd,, % "Button" Idemfeld[A_Index], % "ahk_id " hMuster16
				SendMessage, 0xF2, 0, 0,, % "ahk_id " hAutIdem      ; BM_GETSTATE
				isChecked := ErrorLevel
				;z.= "Zeile " A_Index " - isChecked: " (isChecked ? "true":"false") ", Medikament: " Medikament "`n"
				;ToolTip, % z, 1200, 5, 15

				If RegExMatch(Nacht, "[A]") && !isChecked && StrLen(Medikament) > 0
				{
						Control, Check,,, % "ahk_id " hAutIdem
						t.= Medikament "`n"
						ControlGetPos, ax,ay,aw,ah, % "Button" Idemfeld[A_Index]	, % "ahk_id " hMuster16
						ControlGetPos, rx, ry, rw, rh, % "Edit"    Rezeptfeld[A_Index]	, % "ahk_id " hMuster16
						Felder.Push({"Idemfeld":Idemfeld[A_Index]
											, "ax": (ax + 1)	, "ay": (ay+1)		, "aw": (aw - 2)	, "ah": (ah - 2)
											, "rx" : (rx + 1)  	, "ry" : (ry +1) 	, "rw":  (rw - 3)  	, "rh":  (rh - 3)})
				}
		}

	; Anzeige eines Hinweisfensters das ein Aut-Idem Kreuz gesetzt wurde
		If (Felder.MaxIndex() > 0) {

				RPos         	:= GetWindowSpot(hMuster16)
				Font          	:= "Futura Bk Bt"
				Options    	:= "cFFFFFFFF italic s9 s5"

				Gui, AutIdem: New	 , -Caption +AlwaysOnTop +E0x080800AC +HWNDhRezeptOverlay ;Parent%hMuster16%
				Gui, AutIdem: Margin , 0, 0
				Gui, AutIdem: Show	 , % "x" RPos.X " y" RPos.Y " NoActivate"
				WinSet, ExStyle	, 0x080800AC	, % "ahk_id " hRezeptOverlay
				WinSet, Style 	, 0x54020000	, % "ahk_id " hRezeptOverlay
				Winset, Disable	,                   	, % "ahk_id " hRezeptOverlay
				WinSet, Top   	,                    	, % "ahk_id " hRezeptOverlay

				If !pToken
						pToken	:= Gdip_Startup()

				If !hFamily := Gdip_FontFamilyCreate(Font)  {
					MsgBox, % " font: " Font " does not Exist on this system"
					ExitApp
				  }

				hbm      	:= CreateDIBSection(RPos.W, RPos.H)
				hdc       	:= CreateCompatibleDC()
				obm      	:= SelectObject(hdc, hbm)
				pGraphics	:= Gdip_GraphicsFromHDC(hdc)
				pBrush  	:= Gdip_BrushCreateSolid(0x33FF0000)

				Gdip_SetSmoothingMode(pGraphics, 5)
				Gdip_SetInterpolationMode(pGraphics, 7)

				ControlGetPos, tx,ty,tw,th	, % "Button51" 	, % "ahk_id " hMuster16
				ControlGetPos, ex,,ew,   	, % "Edit2"	    	, % "ahk_id " hMuster16
				Gdip_FillRoundedRectangle(pGraphics, pBrush, tx+tw+3, ty-2, ex+ew-tx-tw-5, 24, 3)
				Gdip_TextToGraphics(pGraphics, "'aut idem' Kreuze wurden automatisch gesetzt !", "x" tx+tw+5 " y" ty 	" w" ex+ew-tx-tw-10 " Center cFF000000 s18 s5", Font)

				For i, feld in Felder
				{
						Gdip_FillRoundedRectangle(pGraphics, pBrush, feld.ax	, feld.ay	, feld.aw	, feld.ah	, 2)
						Gdip_FillRoundedRectangle(pGraphics, pBrush, feld.rx	, feld.ry 	, feld.rw	, feld.rh 	, 2)
						Gdip_TextToGraphics(pGraphics, "aut idem"   	, "x" feld.rx+1 " y" feld.ry-1                	" w" feld.rw-3 " Right " Options, Font)
						Gdip_TextToGraphics(pGraphics, "Medikament"	, "x" feld.rx+1 " y" feld.ry+feld.rh-11	" w" feld.rw-3 " Right " Options, Font)
				}

				UpdateLayeredWindow(hRezeptOverlay, hdc, RPos.X, RPos.Y , RPos.W, RPos.H)
				;SetTimer, AutIdemBlinking, 600

		}
		else
			ExitApp


return

AutIdemBlinking: ;{

	If !WinExist("ahk_id " hMuster16)
		gosub AutIdemGuiClose

	f:= (-1)*f
	alpha:= f > 0 ? 0 : 255
	Blinks ++

	Loop, 10 {
		alpha += f*25
		UpdateLayeredWindow(hRezeptOverlay, hdc, RPos.X, RPos.Y , RPos.W, RPos.H, alpha)
		Sleep, 10
	}

	If !WinExist("ahk_id " hMuster16)
		gosub AutIdemGuiClose

	If Blinks > 40
		gosub AutIdemGuiClose

return ;}



AutIdemGuiClose:
  SelectObject(hdc, obm)
	DeleteObject(hbm)
	DeleteDC(hdc)
	Gdip_DeleteGraphics(pGraphics)
return

}

;}

Thread40:                                                                 	;{ Rezeptfenster Dauermedikamente auslesen

		hMuster16	:= WinExist("Muster 16 ahk_class #32770")
		AlbisRezept_DauermedikamenteAuslesenTmp(hMuster16)

return

AlbisRezept_DauermedikamenteAuslesenTmp(hMuster16) {

		caveBereich	:= false
		;Med          	:= {"Cave":{1:{}}, "Dauer":{1:""}}
		Med          	:= Object()
		Med.Cave   	:= Object()
		Med.Dauer  	:= Object()

	; Listbox als Text einlesen, parsen und ein Objekt mit den Daten erstellen
		ControlGet, medlist, List,, ListBox1, % "ahk_id " hMuster16
		Med.Lb := medList
		Loop, Parse, medList, `n, `r
		{
				If RegExMatch(A_LoopField, "#+\s+([A-Za-zÄÖÜäöüß]+)\s", Gruppierung) {
						If InStr(Gruppierung1, "Problem") || InStr(Gruppierung1, "Allergie")
							caveBereich := true
						else
							caveBereich := false
						continue
				}

			/*  Problemmedikamentenbereich erkannt dann

				 Ich nutze die ersten Zeilen des Dauermedikamentenfenster um dort Medikamente anzuzeigen, auf welche der Patient allergisch reagiert hat
				 oder ich notiere dort Medikamente, welche anderweitig Probleme verursacht haben oder aber verursachen könnten.
				 Beispiel für das Aussehen im Dauermedikamentenfenster:

					############## Problemmedikamente #########*
					Penicillin,Amoxicillin,Cotrim, Cefpodoxim - Quincke Ödem*
					Diclofenac, Etoricoxib, Lidocain - Quincke Ödem*
					ACC, Imeron 350(Coro),  ACE-Hemmer - Quincke Ödem*
					Ciprofloxacin (Fluorchinolone) - zentralnervöse Störung + Psychose!!*
					Metoprolol nicht absetzen - Tachykardie sonst*
					Tomaten, Nüsse* - Allergie bis hinzu Quincke Ödem*
					############### Schilddrüse  ###############*
					L Thyroxin 75 Henning TAB N3 100 St (1---)
					...
					...
					...

				** Bei manchen Patienten kommen dermassen viele Medikamente zusammen. Dies ist der Grund warum ich vom Bundeseinheitlichen Medikationsplan abrate.
					Dieser hat keine Möglichkeit die vielen Allergien anzuzeigen. Ich drucke Medikamente als Stammblatt für den Patient. Das geht erstens sehr schnell. Zweitens
					sind alle Zeilen enthalten und drittens kann man die Dauerdiagnosen ebenfalls ausdrucken. Das Eintragen der Problemmedikamente unter den Dauermedikamten
					hat außerdem den entscheidenden Vorteil das diese beim Erstellen eines Briefes (z.B. Krankenhauseinweisung mit kurzer Epikrise) vom Krankenhausarzt an
					prominenter Stelle gelesen werden können.

			 */
				If caveBereich
				{
						RegExMatch(A_LoopField, "(?<Medikament>[A-Za-zÄÖÜäöüß\-\,\s\(\))]+)\s\-\s(?<Text>.*)", cave)
						caveText	:= Trim(caveText)
						caveText	:= RTrim(caveText, "*")
						If StrLen(caveText) = 0
							continue

						medis      	:= StrSplit(caveMedikament, ",")
						caveExist 	:= false

						For index, cave in Med.Cave
							If InStr(cave.text, caveText)
							{
									Loop % medis.MaxIndex()
										cave.Medikament.Push(Trim(medis[A_Index]))

									caveExist := true
									break
							}

							hinE .= caveText " " (hinweisExist ? "true":"false") "`nMed.Cave: " Med.Cave.MaxIndex() "`n"
							ToolTip, % hinE

							If !caveExist {
								Med.Cave.Push({"medikament": medis, "text": caveText})
							}
				}
			; die eigentlichen Medikamente werden extra gesichert
				else
					Med.Dauer.Push(A_LoopField)

		}

		For i, hinweis in Med.cave
		{
			t.= hinweis.text "`n"
			Loop, % hinweis.Medikament.MaxIndex()
					t.= " - " Trim(hinweis.Medikament[A_Index]) "`n"
		}

		MsgBox, % t

		For i, val in Med.Dauer
			z.= val "`n"

		FileAppend, % z , % A_ScriptDir "\Dauermedikamente.txt"
		MsgBox, % z

return med
}
;}

Thread39:                                                                 	;{ Test Overlay Gui für Rezepte

		f:= 1

		hMuster16	:= WinExist("Muster 16 ahk_class #32770")
		RPos         	:= GetWindowSpot(hMuster16)
		ControlGet, hEdit2, HWND,, Edit2, % "ahk_id " hMuster16
		ControlGetPos, ax,ay,aw,ah, % "Button52" 	, % "ahk_id " hMuster16
		ControlGetPos, x, y, w, h	, % "Edit2"     	, % "ahk_id " hMuster16

		;Gui, AutIdem: New	 , -Caption +OwnDialogs +Owner +ToolWindow +E0x80020 +LastFound +AlwaysOnTop +HWNDhRezeptOverlay
		;Gui, AutIdem: New	 , -Caption +ToolWindow +E0x80020 +LastFound +AlwaysOnTop +HWNDhRezeptOverlay +Parent%hMuster16%
		Gui, AutIdem: New	 , -Caption +ToolWindow +AlwaysOnTop +E0x080800AC +HWNDhRezeptOverlay ;Parent%hMuster16%
		Gui, AutIdem: Margin , 0, 0
		Gui, AutIdem: Show	 , % "x" RPos.X " y" RPos.Y " NoActivate"
		WinSet, ExStyle	, 0x080800AC	, % "ahk_id " hRezeptOverlay
		WinSet, Style 	, 0x54020000	, % "ahk_id " hRezeptOverlay
		Winset, Disable	,                   	, % "ahk_id " hRezeptOverlay

		;ToolTip, % "w" RPos.W " h" RPos.H ", x" x " r" r " y" y " w" w " h" h
		;Pw	:= 50

		pToken 	:= Gdip_Startup()

		Font   	:= "Futura Bk Bt"
		Options:= "cFFFFFFFF italic s9 s5"
		If !hFamily := Gdip_FontFamilyCreate(Font)  {
			MsgBox, % " font: " Font " does not Exist on this system"
			ExitApp
		  }


		hbm := CreateDIBSection(RPos.W, RPos.H)
		hdc := CreateCompatibleDC()
		obm := SelectObject(hdc, hbm)
		pGraphics := Gdip_GraphicsFromHDC(hdc)

		Gdip_SetSmoothingMode(pGraphics, 5)
		Gdip_SetInterpolationMode(pGraphics, 7)
		;Gdip_GraphicsClear(pGraphics,0x00000000)
		pBrush	:= Gdip_BrushCreateSolid(0x66FF0000)
		;pPen 	:= Gdip_CreatePen(0xFF880000, 3)
		;pPen1 	:= Gdip_CreatePen(0xFF880000, 1)
		Px 	:= ax - 6
		Py 	:= ay+Floor(ah//2)
		;Pfeil
		;Gdip_DrawLine(pGraphics, pPen, Px, Py   	, Px-20, Py -10)
		;Gdip_DrawLine(pGraphics, pPen, Px, Py  	, Px-20, Py+10)
		;Gdip_DrawLine(pGraphics, pPen, Px, Py   	, Px, Py)
		Gdip_FillEllipse(pGraphics, pBrush, ax-4, ay-4, aw+8, ah+8)
		Gdip_FillRoundedRectangle(pGraphics, pBrush, x+1, y+1, w-3, h-3, 2)
		Gdip_TextToGraphics(pGraphics, "aut idem", "x" x+1 " y" y " w" w-3 " Right " Options, Font)
		Gdip_TextToGraphics(pGraphics, "Medikament", "x" x+1 " y" y+h-12 " w" w-3 " Right " Options, Font)
		UpdateLayeredWindow(hRezeptOverlay, hdc, RPos.X, RPos.Y , RPos.W, RPos.H)
		SetTimer, BlinkOverlay, 450

return

GuiClose:
  SelectObject(hdc, obm)
	DeleteObject(hbm)
	DeleteDC(hdc)
	Gdip_DeleteGraphics(pGraphics)
	Gdip_Shutdown(pToken)
  ExitApp
return

BlinkOverlay:

	f:= (-1)*f
	alpha:= f > 0 ? 0 : 255

	Loop, 10 {

		alpha += f*25
		UpdateLayeredWindow(hRezeptOverlay, hdc, RPos.X, RPos.Y , RPos.W, RPos.H, alpha)
		Sleep, 10

	}

return

GetImage(hwnd) {
  SendMessage, 0x173,,,, ahk_id %hwnd%
  return (ErrorLevel = "FAIL" ? 0 : ErrorLevel)
}

;}

Thread38:                                                                 	;{ Aut Idem Kreuz setzen per SendMessage

	hMuster16	:= WinExist("Muster 16 ahk_class #32770")
	ControlGet, hwnd, hwnd,, Button53, % "ahk_id " hMuster16
	SendMessage, 0xF2, 0, 0,, % "ahk_id "hwnd                                                          ; BM_GETSTATE
	isChecked := ErrorLevel
	If !isChecked
		Control, Check,,, % "ahk_id " hwnd

return ;}

Thread37:                                                                 	;{ Steuerelementkoordinaten im Rezeptformular

	CoordMode, ToolTip, Screen

	; !!Voraussetzung!! - In den Dauermedikamenten ist im letzten Dosisfeld (Nacht) ein 'A' vermerkt
	Dosisfeld   	:= [6,12,18,24,30,36]  	; Edit
	RezeptFeld	:= [2,8,14,20,26,32]    	; Edit
	Idemfeld   	:= [52,53,54,55,56]      	; Button

	Felder       	:= Object()
	hMuster16	:= WinExist("Muster 16 ahk_class #32770")
	RPos         	:= GetWindowSpot(hMuster16)
	hMonitor   	:= MonitorFromWindow(AlbisWinID())
	ControlGet, hWerbung, HWND,, Shell Embedding1, % "ahk_id " hMuster16
	ControlGetPos, r, y, w, h, % "Edit2"    	, % "ahk_id " hMuster16

	t.= "Rezeptposition per GetWindowSpot()`n"
	t.= "x: " RPos.X " ,y: " RPos.Y " ,w: " RPos.W ",h: " RPos.H "`n"
	t.= "Steuerelementposition per ControlGetPos`n"
	t.= "x: " r " ,y: " y " ,w: " w " ,h: " h "`n"
	t.= "Werbesteuer-Element: " GetHex(hWerbung)

	;MsgBox, % t
	hActive      	:= GetHex(WinActive("A"))

	;If (hMuster16 != hActive)
	;	return

	; Untersucht die Steuerelemente im Rezeptformular, setzt bei Bedarf das Aut-Idem Häkchen
		; die Aut-Idem Steuerelemente im Rezeptfenster sind keine Standard-Checkboxen, mit dem Control-Befehl läßt sich der Status nicht auslesen
		; es wird deshalb mit SendMessage gearbeitet
		Loop, % Dosisfeld.MaxIndex()
		{
				ControlGetText, Nacht       	, % "Edit"    Dosisfeld[A_Index], % "ahk_id " hMuster16
				ControlGetText, Medikament	, % "Edit"  Rezeptfeld[A_Index], % "ahk_id " hMuster16
				ControlGet, hAutIdem, hwnd,, % "Button" Idemfeld[A_Index], % "ahk_id " hMuster16
				ControlGetPos, x, y, w, h     	, % "Button" Idemfeld[A_Index], % "ahk_id " hMuster16
				SendMessage, 0xF2, 0, 0,, % "ahk_id " hAutIdem      ; BM_GETSTATE
				isChecked := ErrorLevel
				ToolTip, % A_Index ": " (isChecked ? "ja":"nein"), % RPos.x+x+w+3, % RPos.y+y-3, 5
				Sleep, 1500
				If RegExMatch(Nacht, "[A]") && !isChecked && StrLen(Medikament) > 0
				{
						Control, Check,,, % "ahk_id " hAutIdem
						t.= Medikament "`n"
						ControlGetPos, x,,,,    	  % "Button" Idemfeld[A_Index]	, % "ahk_id " hMuster16
						ControlGetPos, r, y, w, h, % "Edit"    Dosisfeld[A_Index]   	, % "ahk_id " hMuster16
						Felder.Push({"Idemfeld":Idemfeld[A_Index], "x":(x-2), "y":(y-2), "w":(x+r+w+2), "h":(h+2)})
				}
		}

return ;}

Thread36:                                                                 	;{ Rezeptformular - Editfeld ansprechen, Caret an eine Position setzten einen Bereich auswählen

	global UserInput:= false

	hMuster16:= WinExist("Muster 16 ahk_class #32770")
	WinActivate   	, % "ahk_id " hMuster16
	WinWaitActive	, % "ahk_id " hMuster16
	ControlFocus 	, % "Edit8", % "ahk_id " hMuster16
	ControlGetText	, text, % "Edit8", % "ahk_id " hMuster16

	select	:= "..."
	cpos 	:= InStr(text, select) - 1
	turn  	:= 0
	back 	:= false
	mLen	:= StrLen(text) - 1
	SelW 	:= StrLen(select)

	ih := InputHook("L1 V M")
	ih.NotifyNonText	:= true
	ih.VisibleText     	:= true
	ih.OnKeyDown  	:= Func("OnKeyDown")
	ih.Start()

	Loop 50
	{
		SendMessage 0xB1, % cpos, % cpos+3, Edit8, % "ahk_id " hMuster16 ; EM_SETSEL
		if InterruptableSleep(150)
			break
		SendMessage 0xB1, % cpos, % cpos, Edit8, % "ahk_id " hMuster16 ; EM_SETSEL
		if InterruptableSleep(150)
			break
	}

	ih.Stop()

	SendMessage 0xB1, % cpos, % cpos+3, Edit8, % "ahk_id " hMuster16 ; EM_SETSEL

return

InterruptableSleep(time) {

	Loop, % time/5
	{
			Sleep, 5
			If UserInput
				break
	}

	If UserInput
		return true
}

OnKeyDown() {
UserInput:=true
}

	Loop
	{
		cpos:= cpos > -1 && cpos < (mLen-SelW) && !back ? cpos += 1 : cpos -= 1
		back:= cpos > (mLen-SelW-1) ? true : cpos < 1 ? false : back
		turn:= cpos > (mLen-SelW-1) ? turn += 1 : turn
		If turn > 10
			break
		SendMessage 0xB1, % cpos, % cpos+3, Edit8, % "ahk_id " hMuster16 ; EM_SETSEL
		If cpos=1
			Sleep, 100
		else if cpos=(mLen-SelW-1)
			Sleep, 100
		else
			Sleep, 50
	}

;}

Thread35:                                                                 	;{ OnClipBoardChange - Adresserkennung für schnelle Adress-Sammlung
	#Persistent
	global data   	:= Object()
	data.Adresses	:= Object()
	global rx
	;global ShowInfos:= false
	global ShowInfos:= true

	If !IsObject(rx) {
		rx:= Object()

		;rx.hits:= {"address": 4, "ahk": 2}
		rx.hits:= Object()
		rx.hits.address	:= 2
		rx.hits.ahk     	:= 2

		;{
		Flow := "break|byref|catch|class|continue|else|exit|exitapp|finally|for|global|gosub|goto|if|ifequal|ifexist|ifgreater|ifgreaterorequal|ifinstring|ifless|iflessorequal|ifmsgbox|ifnotequal|ifnotexist|ifnotinstring|ifwinactive|ifwinexist|ifwinnotactive|ifwinnotexist|local|loop|onexit|pause|return|settimer|sleep|static|suspend|throw|try|until|var|while"
		Commands := "autotrim|blockinput|clipwait|control|controlclick|controlfocus|controlget|controlgetfocus|controlgetpos|controlgettext|controlmove|controlsend|controlsendraw|controlsettext|coordmode|critical|detecthiddentext|detecthiddenwindows|drive|driveget|drivespacefree|edit|envadd|envdiv|envget|envmult|envset|envsub|envupdate|fileappend|filecopy|filecopydir|filecreatedir|filecreateshortcut|filedelete|fileencoding|filegetattrib|filegetshortcut|filegetsize|filegettime|filegetversion|fileinstall|filemove|filemovedir|fileread|filereadline|filerecycle|filerecycleempty|fileremovedir|fileselectfile|fileselectfolder|filesetattrib|filesettime|formattime|getkeystate|groupactivate|groupadd|groupclose|groupdeactivate|gui|guicontrol|guicontrolget|hotkey|imagesearch|inidelete|iniread|iniwrite|input|inputbox|keyhistory|keywait|listhotkeys|listlines|listvars|menu|mouseclick|mouseclickdrag|mousegetpos|mousemove|msgbox|outputdebug|pixelgetcolor|pixelsearch|postmessage|process|progress|random|regdelete|regread|regwrite|reload|run|runas|runwait|send|sendevent|sendinput|sendlevel|sendmessage|sendmode|sendplay|sendraw|setbatchlines|setcapslockstate|setcontroldelay|setdefaultmousespeed|setenv|setformat|setkeydelay|setmousedelay|setnumlockstate|setregview|setscrolllockstate|setstorecapslockmode|settitlematchmode|setwindelay|setworkingdir|shutdown|sort|soundbeep|soundget|soundgetwavevolume|soundplay|soundset|soundsetwavevolume|splashimage|splashtextoff|splashtexton|splitpath|statusbargettext|statusbarwait|stringcasesense|stringgetpos|stringleft|stringlen|stringlower|stringmid|stringreplace|stringright|stringsplit|stringtrimleft|stringtrimright|stringupper|sysget|thread|tooltip|transform|traytip|urldownloadtofile|winactivate|winactivatebottom|winclose|winget|wingetactivestats|wingetactivetitle|wingetclass|wingetpos|wingettext|wingettitle|winhide|winkill|winmaximize|winmenuselectitem|winminimize|winminimizeall|winminimizeallundo|winmove|winrestore|winset|winsettitle|winshow|winwait|winwaitactive|winwaitclose|winwaitnotactive"
		Functions := "abs|acos|array|asc|asin|atan|ceil|chr|comobjactive|comobjarray|comobjconnect|comobjcreate|comobject|comobjenwrap|comobjerror|comobjflags|comobjget|comobjmissing|comobjparameter|comobjquery|comobjtype|comobjunwrap|comobjvalue|cos|dllcall|exception|exp|fileexist|fileopen|floor|func|getkeyname|getkeysc|getkeystate|getkeyvk|il_add|il_create|il_destroy|instr|isbyref|isfunc|islabel|isobject|isoptional|ln|log|ltrim|lv_add|lv_delete|lv_deletecol|lv_getcount|lv_getnext|lv_gettext|lv_insert|lv_insertcol|lv_modify|lv_modifycol|lv_setimagelist|mod|numget|numput|objaddref|objclone|object|objgetaddress|objgetcapacity|objhaskey|objinsert|objinsertat|objlength|objmaxindex|objminindex|objnewenum|objpop|objpush|objrawset|objrelease|objremove|objremoveat|objsetcapacity|onmessage|ord|regexmatch|regexreplace|registercallback|round|rtrim|sb_seticon|sb_setparts|sb_settext|sin|sqrt|strget|strlen|strput|strsplit|substr|tan|trim|tv_add|tv_delete|tv_get|tv_getchild|tv_getcount|tv_getnext|tv_getparent|tv_getprev|tv_getselection|tv_gettext|tv_modify|tv_setimagelist|varsetcapacity|winactive|winexist|_addref|_clone|_getaddress|_getcapacity|_haskey|_insert|_maxindex|_minindex|_newenum|_release|_remove|_setcapacity"
		Keynames := "alt|altdown|altup|appskey|backspace|blind|browser_back|browser_favorites|browser_forward|browser_home|browser_refresh|browser_search|browser_stop|bs|capslock|click|control|ctrl|ctrlbreak|ctrldown|ctrlup|del|delete|down|end|enter|esc|escape|f1|f10|f11|f12|f13|f14|f15|f16|f17|f18|f19|f2|f20|f21|f22|f23|f24|f3|f4|f5|f6|f7|f8|f9|home|ins|insert|joy1|joy10|joy11|joy12|joy13|joy14|joy15|joy16|joy17|joy18|joy19|joy2|joy20|joy21|joy22|joy23|joy24|joy25|joy26|joy27|joy28|joy29|joy3|joy30|joy31|joy32|joy4|joy5|joy6|joy7|joy8|joy9|joyaxes|joybuttons|joyinfo|joyname|joypov|joyr|joyu|joyv|joyx|joyy|joyz|lalt|launch_app1|launch_app2|launch_mail|launch_media|lbutton|lcontrol|lctrl|left|lshift|lwin|lwindown|lwinup|mbutton|media_next|media_play_pause|media_prev|media_stop|numlock|numpad0|numpad1|numpad2|numpad3|numpad4|numpad5|numpad6|numpad7|numpad8|numpad9|numpadadd|numpadclear|numpaddel|numpaddiv|numpaddot|numpaddown|numpadend|numpadenter|numpadhome|numpadins|numpadleft|numpadmult|numpadpgdn|numpadpgup|numpadright|numpadsub|numpadup|pause|pgdn|pgup|printscreen|ralt|raw|rbutton|rcontrol|rctrl|right|rshift|rwin|rwindown|rwinup|scrolllock|shift|shiftdown|shiftup|space|tab|up|volume_down|volume_mute|volume_up|wheeldown|wheelleft|wheelright|wheelup|xbutton1|xbutton2"
		Builtins := "base|clipboard|clipboardall|comspec|errorlevel|false|programfiles|true"
		Keywords := "abort|abovenormal|activex|add|ahk_class|ahk_exe|ahk_group|ahk_id|ahk_pid|all|alnum|alpha|altsubmit|alttab|alttabandmenu|alttabmenu|alttabmenudismiss|alwaysontop|and|autosize|background|backgroundtrans|base|belownormal|between|bitand|bitnot|bitor|bitshiftleft|bitshiftright|bitxor|bold|border|bottom|button|buttons|cancel|capacity|caption|center|check|check3|checkbox|checked|checkedgray|choose|choosestring|click|clone|close|color|combobox|contains|controllist|controllisthwnd|count|custom|date|datetime|days|ddl|default|delete|deleteall|delimiter|deref|destroy|digit|disable|disabled|dpiscale|dropdownlist|edit|eject|enable|enabled|error|exit|expand|exstyle|extends|filesystem|first|flash|float|floatfast|focus|font|force|fromcodepage|getaddress|getcapacity|grid|group|groupbox|guiclose|guicontextmenu|guidropfiles|guiescape|guisize|haskey|hdr|hidden|hide|high|hkcc|hkcr|hkcu|hkey_classes_root|hkey_current_config|hkey_current_user|hkey_local_machine|hkey_users|hklm|hku|hotkey|hours|hscroll|hwnd|icon|iconsmall|id|idlast|ignore|imagelist|in|insert|integer|integerfast|interrupt|is|italic|join|label|lastfound|lastfoundexist|left|limit|lines|link|list|listbox|listview|localsameasglobal|lock|logoff|low|lower|lowercase|ltrim|mainwindow|margin|maximize|maximizebox|maxindex|menu|minimize|minimizebox|minmax|minutes|monitorcount|monitorname|monitorprimary|monitorworkarea|monthcal|mouse|mousemove|mousemoveoff|move|multi|na|new|no|noactivate|nodefault|nohide|noicon|nomainwindow|norm|normal|nosort|nosorthdr|nostandard|not|notab|notimers|number|off|ok|on|or|owndialogs|owner|parse|password|pic|picture|pid|pixel|pos|pow|priority|processname|processpath|progress|radio|range|rawread|rawwrite|read|readchar|readdouble|readfloat|readint|readint64|readline|readnum|readonly|readshort|readuchar|readuint|readushort|realtime|redraw|regex|region|reg_binary|reg_dword|reg_dword_big_endian|reg_expand_sz|reg_full_resource_descriptor|reg_link|reg_multi_sz|reg_qword|reg_resource_list|reg_resource_requirements_list|reg_sz|relative|reload|remove|rename|report|resize|restore|retry|rgb|right|rtrim|screen|seconds|section|seek|send|sendandmouse|serial|setcapacity|setlabel|shiftalttab|show|shutdown|single|slider|sortdesc|standard|status|statusbar|statuscd|strike|style|submit|sysmenu|tab|tab2|tabstop|tell|text|theme|this|tile|time|tip|tocodepage|togglecheck|toggleenable|toolwindow|top|topmost|transcolor|transparent|tray|treeview|type|uncheck|underline|unicode|unlock|updown|upper|uppercase|useenv|useerrorlevel|useunsetglobal|useunsetlocal|vis|visfirst|visible|vscroll|waitclose|wantctrla|wantf2|wantreturn|wanttab|wrap|write|writechar|writedouble|writefloat|writeint|writeint64|writeline|writenum|writeshort|writeuchar|writeuint|writeushort|xdigit|xm|xp|xs|yes|ym|yp|ys|__call|__delete|__get|__handle|__new|__set"

		;}

		rx.ahk:= Object()
		rx.ahk.Push({"usetoIdent": true, "addToCount": true	, "needle": "\:\="})	                                 	; 1
		rx.ahk.Push({"usetoIdent": true, "addToCount": true	, "needle": "\%\w+\%"})                         	; 2
		rx.ahk.Push({"usetoIdent": true, "addToCount": true	, "needle": "\%\s\w+"})                          	; 3
		rx.ahk.Push({"usetoIdent": true, "addToCount": true	, "needle": "ahk_\w+"})                          	; 4
		rx.ahk.Push({"usetoIdent": true, "addToCount": true	, "needle": "\w+\(.?\)\s*\{"})                  	; 5
		rx.ahk.Push({"usetoIdent": true, "addToCount": true	, "needle": "\w+\[.?\]"})                          	; 6
		rx.ahk.Push({"usetoIdent": true, "addToCount": true	, "needle": "(``n)|(``r)|(``t)"})                  	; 7
		rx.ahk.Push({"usetoIdent": true, "addToCount": true	, "needle": "\s\&\&\s"})                           	; 8
		rx.ahk.Push({"usetoIdent": true, "addToCount": true	, "needle": "\s*\+\+\s"})                        	; 9
		rx.ahk.Push({"usetoIdent": true, "addToCount": false, "needle": Flow "|" commands "|" Functions "|" keynames "|" Builtins "|" keywords})

		rx.address:= Object()
		rx.address.Push({"usetoIdent": false	, "addToCount": false	, "needle": "(?<Take>[A-Za-zßäöü]+[\s\-][A-Za-zßäöü]+(\s([A-Z][\w\-]+))*)"	, "str": "name"	, "pos": 1})
		rx.address.Push({"usetoIdent": false	, "addToCount": true	, "needle": "(?<Take>[A-Za-zßäöü]+\s*[a-zßäöü\s\.]+\d+[A-Za-z]*)"          	, "str": "street" 	, "pos": 2})
		rx.address.Push({"usetoIdent": true	, "addToCount": true	, "needle": "(?<Take>\d{5}\s[A-Za-zßäöü]+(\s([A-Z][\w\-]+))*)"                    	, "str": "zip   "    	, "pos": 3})
		rx.address.Push({"usetoIdent": true	, "addToCount": true	, "needle": "[T][elefon:.]*+\s([\(+]*\d+[\)-/]*\s*[\d\s-/]*)"                                	, "str": "tel    "    	, "pos": 4})
		rx.address.Push({"usetoIdent": true	, "addToCount": true	, "needle": "(.*?)\:.*?((\(\+)|(0)[\)\d\+]{3,5}([\s]\d{2})+)"                             	, "str": "tel    "    	, "pos": 4})
		rx.address.Push({"usetoIdent": true	, "addToCount": true	, "needle": "([F][ax:.]*)+\s(?<Take>[\(+]*\d+[\)-/]*\s*[\d\s-/]*)"                     	, "str": "fax   "    	, "pos": 5})
		rx.address.Push({"usetoIdent": true	, "addToCount": true	, "needle": "([E\-]*[M][ail:.]*)+\s(?<Take>[A-Za-z_-]+@\w+\.\w{2,5})"          	, "str": "mail "  	, "pos": 6})

	}

	OnClipboardChange("ClipChanged")

return

LoadAdresses(path) {

	If FileExist(jfile)
	{
			;FileRead, f, % jfile
			addresses	:= new JSONFile(jfile)

			MsgBox, % "JSON: " JDia.Ulcus.ve.dia
			;MsgBox, % JDia.HasKey("Patient")
			;MsgBox, % JDia["Patient"]["Neu"]
	}


}

ClipChanged(Type) {

	address     	:= Object()
	matchedList	:= "0"
	ClipAddress	:= false

	If (Type=1) {

			clip:= Clipboard
			TextType := IdentifyText(clip)
			;TrayTip, Clipboard, % "neue Textdaten`n" TextType, 1

			If InStr(TextType, "address")
			{
					clip:= StrReplace(clip, "`t", "")
					clips:= StrSplit(clip, "`n")

					Loop, Parse, clip, `n, `r
					{
							for i, rxmatch in rx[textType]
							{
									If i not in %matchedList%
										If RegExMatch(A_LoopField, rxmatch.needle, match )
										{
												address[rxmatch.pos]:= {"Bezeichner": rxmatch.str, "inhalt": matchTake}
												matchedList .= "," i
												If !ClipAddress
													If i in 2,3,4,5
														ClipAddress:= true
												break
										}
							}
					}

					If !ClipAddress
						return

					For row, rxResult in address
						t.= RxResult.bezeichner "`t: " RxResult.inhalt "`n"

					TrayTip, Neue Adresse
					MsgBox, % t
					Sleep, 8000
					ToolTip
			}
			else if InStr(TextType, "ahk")
			{

					;TrayTip, Clipboard, % "Autohotkey script detected`nClick tray to save to file!", 10, 0x11
					MsgBox, 0, % "Autohotkey script detected", % "Click ok to save to file!", 10
					IfMsgBox, No
						return

			}

	}

}

IdentifyText(Text) {

	hitscore:= Object()

	t.= SubStr(text, 1, 30) "`n`n"

	For TextType in rx
	{
			hits := 0, x:=""
			If InStr(TextType, "hits") || (StrLen(TextType) =0)
				continue

			Notify("`t" TextType ": min hits needed (" rx.hits[TextType] ")")
			For n, rxmatch in rx[TextType]
			{
				RegExReplace(Text, rxmatch.needle, "", Count)
				If (rxmatch.usetoIdent && Count > 0)
					If rxmatch.addToCount
							hits += Count
					else
							hits ++

				Notify("`t" (hits=hits_o ? "false" : "true") "`thits: " hits "`tneedle no.: " SubStr( "00" n, -1) "`tmatch object: " (IsObject(match) ? "true" : "false"))
				hits_o:= hits
			}

			If (hits >= rx.hits[TextType])
				hitscore.Push({"TextType": TextType, "hits": hits})

			hitso .= hits ", "
	}

	Notify( "hits: " hitso "`r`n---------------------------------------------------------------`r`n")

	If hitscore.MaxIndex() = 1
		return hitscore[1].TextType
	else If hitscore.MaxIndex() > 1
	{
			Loop, % hitscore.MaxIndex()
				hitscoreMax:= hitscore[A_Index].hits > hitscoreMax ? A_Index : hitscoreMax

			return hitscore[hitscoreMax].TextType
	}
	else
		return "kein bestimmbarer Text type"
}

;}

Thread34:                                                                 	;{ alle UTF-8 Zeichen ausgeben

	/*
			࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊࿊
			⊶⊷⊶⊷⊶⊷⊶⊷⊶⊷⊶⊷⊶⊷⊶⊷⊶⊷⊶⊷⊶⊷
			⊞⊞⊞⊞⊞⊞⊞⊞⊞⊞⊞⊞⊞⊞⊞⊞⊞⊞⊞⊞⊞⊞⊞
			⊡⊡⊡⊡⊡⊡⊡⊡⊡⊡⊡⊡⊡⊡⊡⊡⊡⊡⊡⊡⊡⊡⊡
			⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆⋆
			┌──────────────────┰─────────────────┒
			╭╮
			alphabet:= AÄaäBbCcDdEeFfGgHhIiJjKkLlMmNnOÖoöPpQqRrSsTtUÜuüVvWwXxYyZz
	*/
	;MsgBox, % A_AppData
	zeichen := FileOpen(A_ScriptDir "\Zeichenausgabe.txt", "w", "UTF-8")
	Loop, 160
	{
		Pos				:= A_Index ;+ 32
		hex             	:= GetHex(pos)
		;hexAusgabe	:= "0x" SubStr(000 SubStr(hex, 3, StrLen(hex) - 2), -2)
		hexAusgabe	:= hex
		zeichen.WriteLine(SubStr("000" pos, -2) "(" hexAusgabe ") = " Chr(hex))
	}

	zeichen.Close()
	Run, % "notepad.exe " A_ScriptDir "\Zeichenausgabe.txt"

return ;}

Thread33:                                                                   	;{ Commandline bei COM winmgmts

	SetTitleMatchMode, 2		;Fast is default
	SetTitleMatchMode, Slow		;Fast is default
	DetectHiddenWindows, on
	hwnd:= WinExist("Addendum Info ahk_exe AutoHotkey")
	for Prozess in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process WHERE Name = Autohotkey")
        If RegExMatch(Prozess.commandline, "AddendumInfo\.ahk\""$")
			 MsgBox, % Prozess.ProcessID "´n" Prozess.commandline


return ;}

Thread32:                                                                 	;{ StrDiff Test


	t:=  "Hartwig, Hartwig"                                           	" " StrDiff("Hartwig", "Hartwig") "`n"
	t.=  "Hartwig, Lassig"                                              	" " StrDiff("Hartwig", "Lassig") "`n"
	t.=  "Hans-Joachim, Hartwig"                                    	" " StrDiff("Hans-Joachim", "Hartwig") "`n"
	t.=  "Hans-Joachim, Lassig"                                    	" " StrDiff("Hans-Joachim", "Lassig") "`n"
	t.=  "Hans-Joachim, Hans"                                      	" " StrDiff("Hans-Joachim", "Hans") "`n"
	t.=  "Hans-Joachim, Hans Joachim"                       	" " StrDiff("Hans-Joachim", "Hans Joachim") "`n"
	t.=  "Hartwig Hans Joachim, Hartwig Hans"             	" " StrDiff("Hartwig " . "Hans Joachim"	, "Hartwig " . "Hans") "`n"
	t.=  "Hartwig Hans Joachim, Lassig Hans Joachim"    	" " StrDiff("Hartwig "  . "Hans Joachim"	, "Lassig " 	 . "Hans Joachim") "`n"

	MsgBox, % t
return ;}

Thread31:                                                                 	;{ Abrechnung vorbereiten automatisieren und Ergebnisse auswerten

	; Dialog Abrechnung vorbereiten

return

AlbisAbrechnungVorbereiten() {                                 	;-- öffnet, setzt Einstellungen und startet den Abrechnung vorbereiten Dialog

		PraxTT("Der Einstellungsdialog für die Abrechnungs-`nvorbereitung wird aufgerufen und bearbeitet.", "15 3")
		WinDialog:= "Abrechnung KVDT vorbereiten ahk_class #32770"
		Albismenu(32944, WinDialog)
		VerifiedClick("Button9", WinDialog) 	; AODT
		VerifiedClick("Button13", WinDialog)	; GNR-Regelwerkskontrolle
		;VerifiedClick("Button14", WinDialog)	; KRW-Regelprüfung
		;VerifiedClick("Button15", WinDialog)	; Obligat
		;VerifiedClick("Button16", WinDialog)	; Fakultativ
		VerifiedClick("Button20", WinDialog)	; Scheine ohne &Einlesedatum
		VerifiedClick("Button23", WinDialog)	; Patienten mit mehreren Scheinen
		Sleep, 10000
		VerifiedClick("Button1", WinDialog)		; OK
		PraxTT("Abrechnung vorbereiten`nWarte auf die Fertigstellung der Daten....", "0 3")
		; msctls_progress321
		WinWait, bitte warten ahk_class #32770,
		while, WinExist("bitte warten ahk_class #32770")
		{
				Sleep, 50		;Prüfe Komplexe...Erzeuge Liste...
		}

return
}
 ;}

Thread30:                                                                 	;{ selbst bei blockierendem Popup-Fenster läßt sich in einer Karteikarte noch etwas auslesen

	ControlGetText, t, Edit2, ahk_class OptoAppClass
	MsgBox, % t

return ;}

Thread29:                                                                 	;{ Auslesen einer Listbox und Auswahl eines Listboxeintrages

	entries := AlbisReadFromListbox("GNR-Vorschlag zur Befundung", 1, 0)
	MsgBox, % entries
	hwnd:= GetHex(WinExist("GNR-Vorschlag zur Befundung ahk_class #32770"))
	SendMessage, 0x186, 1, 0, ListBox1, % "ahk_id " hwnd
	MsgBox, % GetHex(ErrorLevel) "`n" hwnd

return ;}

Thread28:                                                                 	;{ andere Methode um in einer Akte die Tastatureingabe vorzubereiten

	AlbisWinID:= AlbisWinID()
	hCLient:= AlbisGetActiveMDIChild()
	clientT:= WinGetTitle(hClient)
	hMDIFrame:= Controls("AfxMDIFrame90", "ID", hClient)
	hAfxFrame:= Controls("AfxFrameOrView90", "ID", hMDIFrame)
	hEingabeWin:= Controls("#32770", "ID", hAfxFrame)
	ControlSend,, {End} , % "ahk_id " hEingabeWin
	Sleep, 50
	ControlSend,, {Down} , % "ahk_id " hEingabeWin
	ControlFocus, Edit2,  % "ahk_id " hEingabeWin
	ControlSetText, Edit2, % "31.12.2019", % "ahk_id " hEingabeWin

	ToolTip, % "active MDI Client handle: " hClient "`nClient Title: " clientT "`nhMdiFrame: " hMDIFrame "`nhAfxFrame: " hAfxFrame "`nEingabeWin: " hEingabeWin "`nSendMessage ErrorLevel: " Err

return ;}

Thread27:                                                                 	;{ AlbisFristenRechner

	AlbisFristenGuiTest()

return

AlbisFristenGuiTest() {

	static Start, Ende, hStart, hEnde, hFrist, Termine, Frist, hOver
	static AUSeitO, AUBisO
	static cTT    	:= Object()
	static Fristen	:= Object()

	CoordMode, Pixel, Screen

	ControlGetPos, tX, tY,, tH, Button8, Muster 1a ahk_class #32770
	hAU      	:= WinExist("Muster 1a ahk_class #32770")
	AU        	:= GetWindowSpot(hAU)
	Fristen   	:= AlbisFristenRechnerTest()
	FristX    	:= AU.BW
	FristY    	:= AU.BH
	FristW    	:= AU.CW - AU.BW*2
	FristH    	:= 20

	Gui, Frist: New, -Caption -DPIScale -AlwaysOnTop +HWNDhFrist Parent%hAU% +0x00000004  ;0x00000004 = NoParentNotify
	Gui, Frist: Margin, 0, 0
	Gui, Frist: Color, % "c172842"   ;"c128576"
	Gui, Frist: Font, s8 q5 cWhite, Futura Md Bt
	Gui, Frist: Add, Text, % "x4 y3	BackgroundTrans vTermine                    	", % "Krankengeldzahlung "
	Gui, Frist: Font, s8 q5 Normal cWhite, Futura Bk Bt
	Gui, Frist: Add, Text, % "x+0 w" FristW " BackgroundTrans vStart		+HWNDhStart	", % Fristen.Anzeige
	Gui, Frist: Show, % "x" FristX " y" FristY " w" FristW " h" FristH, AdmFristen

	SetTimer, AlbisFristenUpdateTest, 300

return

AlbisFristenUpdateTest:

	If !WinExist("Muster 1a ahk_class #32770")
	{
			Gui, Frist: Destroy
			SetTimer, AlbisFristenUpdateTest, Off
			return
	}

	ControlGetText, AUSeit, Edit1, Muster 1a
	ControlGetText, AUBis, Edit2, Muster 1a
	If (AUSeitO = AUSeit) && (AUBisO = AUBis)
    		 return
	AUSeitO 	:= AUSeit
	AUBisO 	:= AUBis
	Fristen:= AlbisFristenRechnerTest(AUSeit, AUBis)
	ControlSetText,, % Fristen.Anzeige, % "ahk_id " hStart
	WinSet, Redraw,, % "ahk_id " hFrist

return

}

AlbisFristenRechnerTest(AUSeit:="", AUBis:="") {                                                                                       	;-- zeigt Datum des Beginns und des Ende der Krankengeldfortzahlung an

		info =
		(
		Entgeltfortzahlung
		Der Ereignistag für die Berechnung der Fristen für die Entgeltfortzahlung wird wie folgt bestimmt:
		Wenn die Arbeitsunfähigkeit schon vor Beginn der Arbeit eingetreten ist,
		ist der Ereignistag der Tag vor dem 1. Arbeitsunfähigkeitstag.
		Hat der Arbeitnehmer am ersten Tag der Arbeitsunfähigkeit noch gearbeitet,
		entspricht der Ereignistag dem 1. Arbeitsunfähigkeitstag

		Krankengeld
		1. Der Anspruch auf Krankengeld entsteht bei stationärer Behandlung von ihrem Beginn an und im Übrigen von dem Tag
			der ärztlichen Feststellung der Arbeitsunfähigkeit an. Die Frist für den Beginn auf Krankengeld kann vom Beginn der Frist
			auf Entgeltfortzahlung abweichen. Anrechenbare Vorerkrankungen sind in den Ausführungen nicht berücksichtigt.
		2. Die Krankenversicherung übernimmt ab der 6. Woche die Krankengeldfortzahlung.
		3. Grundsätzlich gilt, dass das Krankengeld wegen derselben Erkrankung erst einmal relativ lange läuft – nämlich 78 Wochen
			oder 19,5 Monate lang innerhalb von drei Jahren (§ 48 SGB V). Dabei müssen Sie nicht am Stück krankgeschrieben sein.
			Die Zeiträume werden zusammengezählt.
		4. Auf das Krankengeld müssen Sie keine Steuern zahlen. Es unterliegt jedoch dem Progressionsvorbehalt (§ 32b EStG).
			Dadurch wird das Krankengeld zum versteuernden Einkommen hinzugerechnet. Der somit ermittelte höhere Steuersatz
			wird auf das zu versteuernde Einkommen angewandt. So vermeidet der Fiskus, dass Versicherte, die Krankengeld bezogen
			haben, einen geringeren Steuersatz haben als Versicherte, die kein Krankengeld bekommen haben.
		)

		bgStr1        	:= "ab d."
		endStr1		:= ""
		bgStr2			:= "bis zum"
		endstr2			:= ""

		Fristen			:= Object()

		If !AUSeit || !AUBis
		{
				If !WinExist("Muster 1a ahk_class #32770")
				{
					 return 0
				 }
				ControlGetText, AUSeit, Edit1, Muster 1a
				ControlGetText, AUBis, Edit2, Muster 1a
				If (AUSeitO = AUSeit) && (AUBisO = AUBis)
					 return
				AUSeitO 	:= AUSeit
				AUBisO 	:= AUBis
		}

		startdate:= StrSplit(AUSeit, ".").3 StrSplit(AUSeit, ".").2 StrSplit(AUSeit, ".").1
		enddate:= StrSplit(AUBis, ".").3 StrSplit(AUBis, ".").2 StrSplit(AUBis, ".").1
		KrankengeldStart	:= startdate + 0
		KrankengeldEnde	:= startdate + 0
		KrankengeldStart 	+= 41 	, days
		KrankengeldEnde 	+= (78*7), days
		FormatTime, KrankengeldStart, % KrankengeldStart, dd.MM.yyyy
		FormatTime, KrankengeldEnde, % KrankengeldEnde, dd.MM.yyyy
		FormatTime, Heute, , dd.MM.yyyy
		TageBisBeginn	:= DateDiff("dd", Heute, KrankengeldStart) + 0
		TageBisAblauf	:= DateDiff("dd", Heute, KrankengeldEnde) + 0
		KgStartVorbei	:= TageBisBeginn	< 0 ? 1 : 0
		KgEndeVorbei	:= TageBisAblauf	< 0 ? 1 : 0
		TageBisBeginn	:= TageBisBeginn	< 0 ? -1 * TageBisBeginn : TageBisBeginn
		TageBisAblauf	:= TageBisAblauf 	< 0 ? -1 * TageBisAblauf : TageBisAblauf

		If (TageBisBeginn >= 21) && !KGStartVorbei
		{
				TageBKg   	:= Mod(TageBisBeginn, 7)
				WochenBKg	:= Floor((TageBisBeginn - TageBkg)/7)
				ZeitBKg      	:= WochenBKg = 0 ? "(" : WochenBkg = 1 ? "(1 Woche": "(" WochenBKg " Wochen"
				ZeitBKg      	.= TageBKg = 0 ? ") " : TageBkg = 1 ? "  u. 1 Tag) ": " u. " TageBKg " Tage) "
		}
		else if (TageBisBeginn < 21) && !KGStartVorbei
				ZeitBKg      	:= "(" TageBisBeginn " Tage) "

		if KGStartVorbei
		{
				bgStr1			:= "seit d."
				endstr1			:= ""
		}

		If (TageBisAblauf >= 21) && !KgEndeVorbei
		{

				TageAKg   	:= Mod(TageBisAblauf, 7) + 0
				WochenAKg	:= Floor((TageBisAblauf - TageAKg)/7) + 0
				ZeitAKg      	:= WochenAKg = 0 ? "(" : WochenAkg = 1 ? "(1 Woche": "(" WochenAKg " Wochen"
				ZeitAKg      	.= TageAKg = 0 ? ") " : TageAkg = 1 ? " u. 1 Tag) ": " u. " TageAKg " Tage) "
		}
		else if (TageBisAblauf < 21) && !KgEndeVorbei
				ZeitAKg      	:= "(" TageBisAblauf " Tage)"

		If KgEndeVorbei
		{
				ZeitAKg     	:= ""
				ZeitBKg     	:= ""
				bgStr1			:= ""
				endStr1			:= ""
				bgStr2			:= "ist am"
				endstr2			:= "ausgelaufen"
		}

		Fristen.KgStart   	:= Trim(KrankengeldStart)
		Fristen.KgEnde  	:= Trim(KrankengeldEnde)
		Fristen.ZeitBKg  	:= ZeitBKg
		Fristen.ZeitAKg  	:= ZeitAKg
		Fristen.BKgStr1  	:= bgStr1
		Fristen.BKgStr2  	:= endStr1
		Fristen.AKgStr1  	:= bgStr2
		Fristen.AKgStr2  	:= endStr2
		Fristen.info        	:= info
		Fristen.Anzeige   	:= Fristen.BKgStr1 " " Fristen.KgStart " " Fristen.ZeitBKg Fristen.AKgStr1 " " Fristen.KgEnde " " Fristen.ZeitAKg " " Fristen.AKgStr2
		;PraxTT("AU von " AUSeit " bis zum " AUBis  "`n`nDie Krankengeldzahlung " bgStr1 ZeitBKg " am " KrankengeldStart endStr1 "`nund " bgStr2 ZeitAKg endStr2 " am " KrankengeldEnde, "25 -2")

return Fristen
}

;}

Thread26:                                                                  	;{ get scrollbar width

	hWin:= WinExist("WinSpy - Tree")
	ControlGet, Chwnd, HWND, , SysTreeView321, % "ahk_id " hWin
	scb:= GetScrollInfo(Chwnd, 1)
	MsgBox, % "hWin: " GetHex(hWin) ", hControl: " GetHex(Chwnd) "`n" "Min: " scb.Min ", Max: " scb.Max ", Page: "  scb.Page ", Pos: " scb.Pos

return

GetScrollInfo(hWnd, fnBar := 1) {
    Local o := {}
    NumPut(VarSetCapacity(SCROLLINFO, 28, 0), SCROLLINFO, 0, "UInt")
    NumPut(0x1F, SCROLLINFO, 4, "UInt") ; fMask: SIF_ALL
    DllCall("GetScrollInfo", "Ptr", hWnd, "Int", fnBar, "Ptr", &SCROLLINFO)
    o.Min  := NumGet(SCROLLINFO, 8, "Int")
    o.Max  := NumGet(SCROLLINFO, 12, "Int")
    o.Page := NumGet(SCROLLINFO, 16, "UInt")
    o.Pos  := NumGet(SCROLLINFO, 20, "Int")
    Return o
}
;}

Thread25:	                                                                	;{ Scite - Zeile in der das Caret steht einlesen

	hSci:= GetFocusedControl()
	ToolTip, % Sci_LineFromPos(hSci, Sci_GetCurPos(hSci))
	SendMessage, 0x111, 225,,, ahk_id %hSci%
>>>>>>> Stashed changes

;q::gosub Thread17
return
<<<<<<< Updated upstream
ExitApp
=======
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
;}

Thread24:                                                                		;{ RichEdit in Albis manipulieren ohne einen Absturz zu verursachen

	;Hypertonie {I10.0G}; Testwort ; Diabetes mellitus Typ 2 {E11.90G};

	fchwnd:= AlbisGetActiveControl("hwnd")
	sel:= RE_GetSel(fchwnd)
	len:= RE_GetTextLength(fchwnd)
	MsgBox, % sel.S ", " sel.E "`nTextLength: " len "`n" GetClassName(fchwnd)
	ExitApp
	SendRaw, % "Hypertonie {I10.0G} Testwort Diabetes mellitus Typ 2 {E11.90G} "
	suchwort:= "Testwort"
	fchwnd:= AlbisGetActiveControl("hwnd")
	If !InStr(GetClassName(fchwnd), "RichEdit")
			return
	lwPos:= RE_FindText(fchwnd, suchwort, "WHOLEWORD")
	MsgBox, % lwpos
	if lwpos = -1
		return
	ExitApp
;	RichEdit_SetSel(fchwnd, lwpos, lwpos + StrLen(suchwort))
;	RichEdit_ReplaceSel(fchwnd, "")

return



;}

Thread23:                                                                 	;{ Diagnosen trimmen

	MsgBox, % TrimDiagnose("Nicht näher bezeichnete durch multiplen Substanzgebrauch und Konsum anderer psychotroper Substanzen bedingte psychische und Verhaltensstörung, li., Z.n., G. {F19.9Z};")

return

TrimDiagnose(str) {
	str:= RegExReplace(str, "\,\s*(li|re|bds)\s*\.")
	str:= RegExReplace(str, "\,\s*G\.\s*")
	str:= RegExReplace(str, "\,\s*Z\.n\.\s*")
	str:= RegExReplace(str, "\{[\w\.]+\}\;*")
return Trim(str)
}   ;}

Thread22:                                                                 	;{ JSON-Objekt

	;jfile:= AddendumDir "\include\AlbisMenu.json"
	jfile1:= AddendumDir "\include\RezeptHelferDb_Test.json"
	jfile2:= AddendumDir "\include\RezeptHelferDb_Test2.json"
	If FileExist(jfile1)
	{
			;FileRead, f, % jfile
			Rezepte:= Object()
			Rezepte	:= new JSONFile(jfile1)

			For i, val in Rezepte.Object()
				t.= val.bezeichner

			MsgBox, % t
			;MsgBox, % "JSON: " Rezepte[1].Beschreibung "`n" Rezepte[1].Medikamente[1]

			Rezepte.Push({"Bezeichner":"Diabetes standard", "Beschreibung":"alle Diabetes kack Medikamente", "Medikamente":{}, "Zusaetze":{}, "RezeptTyp":"", "Optionen":""})

			;Rezepte.Save(true)
			;MsgBox, % JDia.HasKey("Patient")
			;MsgBox, % JDia["Patient"]["Neu"]
	}

return
;}

Thread21:                                                                 	;{ WORD Automation

	;--------------------------------------------------------------------------------------------------------------------
	;	Verbindung zu Word herstellen und falls mehrere Word-Dokumente geöffnet sind,
	;	muss das richtige Dokument gefunden werden
	;--------------------------------------------------------------------------------------------------------------------;{
			oWord 	:= ComObjActive("Word.Application")
			docNr	:= 0

			Loop, % oWord.Documents.Count
			{
					Content:= oWord.Documents.Item(A_Index).Content.Text
				;- Übereinstimmung anhand der speziellen Schlüsselwörter und des Patientennamen wird überprüft
					If ( InStr(Content, "<Entschuldigung1>") and InStr(Content, namen[1]) and InStr(Content, namen[2]) )
					{
						;- gefunden, dann wird genau dieses Dokument zum aktiven Dokument gemacht
							oWord.Documents(A_Index).Activate
							docNr:= A_Index
							WinActivate, % oWord.Documents(A_Index).Name
							break
					}
			}

		;- falls kein passendes Dokument geöffnet wird, erfolgt ein Hinweis und das Skript-Gui wird wieder angezeigt
			If !docNr
			{
					MsgBox,1, Addendum für Albis on Windows,
													(LTrim
													Scheinbar konnte das Dokument `"%schulbefreiungDoc%`"
													nicht aufgerufen werden.
													Möchten Sie es noch einmal probieren?
													Abbrechen - beendet das Skript, die gemachten Einstellungen
													für diesen Patienten werden aber gespeichert.
													)
					IfMsgBox, Yes
							return
					IfMsgBox, Cancel
					{
							toWrite:= % namen[1] "," namen[2] "`n" MyCal1 "`n" MyCal2 "`n" Entschuldigung[1] "`n" Entschuldigung[2]
							FileAppend, % toWrite, % A_ScriptDir "\letzte_SchulbefreiungSicherung.txt"
							ExitApp
					}

			}
		;}

	;--------------------------------------------------------------------------------------------------------------------
	;	im Dokument werden jetzt die Schlüsselwörter durch die erstellten Texte ersetzt und das
	;	Dokument wird gedruckt, gespeichert und geschlossen, das Skript wird anschliessend beendet
	;--------------------------------------------------------------------------------------------------------------------;{
		Entschuldigung:= []
		Entschuldigung[1]:= "Markus ist am 23.8.2019 schulunfähig erkrankt"

		MSWord_FindAndReplace(oWord, "<Entschuldigung1>", Entschuldigung[1])
		;MSWord_FindAndReplace(oWord, "<Entschuldigung2>", Entschuldigung[2])
		ValueSet	=
							(LTrim Join
							LineNumbering.Active,False
							,Orientation,wdOrientPortrait
							,TopMargin,CentimetersToPoints(1)
							,BottomMargin,CentimetersToPoints(1)
							,LeftMargin,CentimetersToPoints(1.5)
							,RightMargin,CentimetersToPoints(1.3)
							,Gutter,CentimetersToPoints(0)
							,HeaderDistance,CentimetersToPoints(1.27)
							,FooterDistance,CentimetersToPoints(1.27)
							,PageWidth,CentimetersToPoints(14.8)
							,PageHeight,CentimetersToPoints(21)
							,FirstPageTray,wdPrinterLowerBin
							,OtherPagesTray,wdPrinterLowerBin
							,SectionStart,wdSectionNewPage
							,OddAndEvenPagesHeaderFooter,False
							,DifferentFirstPageHeaderFooter,False
							,VerticalAlignment,wdAlignVerticalTop
							,SuppressEndnotes,False
							,MirrorMargins,False
							,TwoPagesOnOne,False
							,BookFoldPrinting,False
							,BookFoldRevPrinting,False
							,BookFoldPrintingSheets,1
							,GutterPos,wdGutterPosLeft
							)
		;MSWord_SelectionSetup(oWord, "PageSetup", ValueSet)
		;_ := ComObjMissing()
		oWord.ActivePrinter := "Microsoft Print to PDF"
		ActiveDocument.PrintOut
		;oWord.ActiveDocument.PrintOut("","wdPrintAllDocument","_wdPrintDocumentContent",1,"","wdPrintAllPages",_,1,1,false,0,_,0,0,0,0,0,0)

		;oWord.ActiveDocument.Save()
		;oWord.ActiveDocument.close
	;}


return

MSWord_FindAndReplace(obj, search, replace) {
	obj.Selection.Find.ClearFormatting
	obj.Selection.Find.Replacement.ClearFormatting
	obj.Selection.Find.Execute( search, 0, 0, 0, 0, 0, 1, 1, 0, replace, 2)
}

MSWord_SelectionSetup(obj, setup:="PageSetup", ValueSet:="TopMargin,20,LeftMargin,20,RightMargin,20,BottomMargin,20") {

	;Parameters:
	;
	; setup 			- can be every command MS Word is using after the Selection command
	;

	/*
		oWord.Selection.PageSetup.TopMargin := 20
		oWord.Selection.PageSetup.LeftMargin := 20
		oWord.Selection.PageSetup.RightMargin := 20
		oWord.Selection.PageSetup.BottomMargin := 20

		oWord.Selection.Font.Bold :=   1
		oWord.Selection.Font.Size :=   18
	*/

	Set:=[], Cnt:=1

	Set:= StrSplit(ValueSet, ",")
	Loop
	{
			obj.Selection[(Set[A_Index])] := Set[A_Index+1]
			cnt += 2
	} until (Cnt=Set.MaxIndex())


}


;}
>>>>>>> Stashed changes

Thread20:                                                                    ;{ AlbisGetActiveWindowType Überarbeitung

	ToolTip, % AlbisGetActiveWindowType() "`n" AlbisGetActiveWinTitle(), 1700, 150, 11


return
;}

Thread19:                                                                    ;{ WinTitel Übergabe an Funktion

	MsgBox, % tmpVerifiedClick("", "", "", "0x234FDAD")
	MsgBox, % tmpVerifiedClick("Button1", "Notepad")
	MsgBox, % tmpVerifiedClick("", "34123243254")
	MsgBox, % tmpVerifiedClick("", "0x234FDAD")


return
tmpVerifiedClick(CName, WinTitle:="", WinText:="", WinID:="") {									                                     				;-- 4 verschiedene Methoden um auf ein Control zu klicken


	; leeren des Fenster-Titel und Textes wenn ein Handle übergeben wurde
		if WinID {
				WinTitle:= "ahk_id " WinID
				WinText:= ""
		} else if RegExMatch(WinTitle, "^0x[\w]+$") {
				WinTitle	:= RegExMatch(WinTitle, "^0x[\w]+$")	? ("ahk_id " WinTitle)	: (WinTitle)
		} else if RegExMatch(WinTitle, "^\d+$", digits) {
				WinTitle	:= StrLen(WinTitle) = StrLen(digits)     	? ("ahk_id " digits)  	: (WinTitle)
		}

return WinTitle
}
;}

Thread18:                                                                    ;{ Laborabfrage

			Laborname:= "IMD"
			AlbisActivate(2)
			ControlGet, currSel, Choice,, ListBox1, Labor auswählen ahk_class #32770
			ToolTip, % currSel
			If !InStr(currSel, Laborname)
			{
					Control, ChooseString, % Laborname, ListBox1, Labor auswählen ahk_class #32770
					sleep, 200
					If A_Index > 10
					{
								MsgBox, % "Das eingestellte Labor: " Laborname "`nkonnte nicht ausgewählt werden."
								return
					}
			}

			err:= VerifiedClick("Button1", "Labor auswählen ahk_class #32770")
			ToolTip, % currSel "`n" err

			WinWait, ALBIS, Keine Datei(en) im Pfad, 6
			If ErrorLevel = 0
			{
					err:= VerifiedClick("Button1", "ALBIS", "Keine Datei(en) im Pfad")
					LaborDaten_Status:= 0
					return
			}
return ;}

Thread17:                                                                 	;{ Caret Position im Edit-/RichEdit-Control

	hactiveID  	:= AlbisGetActiveControl("hwnd")
	controlText:= ControlGetText("", "ahk_id " hactiveID)
	CurPos   	:=DllCall("User32.dll\SendMessage", "Ptr", hactiveID, "UInt", 0x7D8,"UInt")
	CharPos   	:=DllCall("User32.dll\SendMessage", "Ptr", hactiveID, "UInt", 0x0D6,"UInt")
	Var3     	:=DllCall("User32.dll\SendMessage", "Ptr", hactiveID, "UInt", 0x0B7,"UInt")
	WordPosL	:=DllCall("User32.dll\SendMessage", "Ptr", hactiveID, "UInt", 0x8DA, "UInt", CurPos, "UInt", 1, "UInt")
	WordPosR	:=DllCall("User32.dll\SendMessage", "Ptr", hactiveID, "UInt", 0x8DB, "UInt", CurPos, "UInt", 1, "UInt")
	;DllCall("User32.dll\SendMessage", "Ptr", hactiveID, "UInt", 0x00B1, "Ptr",  WordPosR+1 , "Ptr",WordPosL-1 , "Ptr")
	MsgBox, % hactiveID "`n" controlText "`nCharPos: " CharPos ", CurPos: " CurPos ", Var: " Var3 "`nWordPosL: " WordPosL ", WordPosR: " WordPosR

return ;}

Thread16:																		;{ SciteZoom auslesen und setzen

	ComObjError(0)
	scite := ComObjActive("SciTE4AHK.Application")
	ComObjError(1)
	if (!scite) ;SciTE not running, empty list
			return
	;SciTEHwnd:= oSci.SciTEHandle
	ControlGet, hScintilla1, hwnd,, Scintilla1, ahk_class SciTEWindow
	ControlGet, hScintilla2, hwnd,, Scintilla2, ahk_class SciTEWindow
	SCI_SETZOOM:=2373, SCI_GETZOOM:=2374
	SendMessage, SCI_GETZOOM, 0, 0,, % "ahk_id " hScintilla1
	ZoomScintilla1:= ErrorLevel
	SendMessage, SCI_GETZOOM, 0, 0,, % "ahk_id " hScintilla2
	ZoomScintilla2:= ErrorLevel
	MsgBox, % hScintilla1 ": " ZoomScintilla1 "`n" hScintilla2 ": " ZoomScintilla2
	SendMessage, SCI_SETZOOM, % ZoomScintilla1 + 4, 0,, % "ahk_id " hScintilla1
	SendMessage, SCI_SETZOOM, % ZoomScintilla2 + 4, 0,, % "ahk_id " hScintilla2
	Sleep, 4000
	SendMessage, SCI_SETZOOM, % ZoomScintilla1, 0,, % "ahk_id " hScintilla1
	SendMessage, SCI_SETZOOM, % ZoomScintilla2, 0,, % "ahk_id " hScintilla2

return
;}

Thread15:																		;{ Focused Control

	;AlbisActivate(2)
	;ControlCmd("Edit1", "SetFocus", "Speichern unter ahk_class #32770")
	;MsgBox, % GetHex(AlbisGetSpecificHMDI("Tagesprotokoll"))

	;AlbisActivate(1)
	;hSMDI:= AlbisGetSpecificHMDI("19933")
	;MsgBox, % hSMDI "`n" WinGetTitle(hSMDI)
	;Sleep, 1000
	;AlbisCloseMDITab("14012")
	FocusedControl:= Chwnd:= GetFocusedControlHwnd()
		cNN	:= Control_GetClassNN(AlbisWinID(), FocusedControl)
	hwnd:= GetFocusedControl()
	ControlGetFocus, cFocus, ahk_class OptoAppClass
	ToolTip, % GetHex(FocusedControl) "`n" cNN "`n" cFocus "`n" GetHex(hwnd) "`n" AlbisWinID() "`n" A_CaretX ", " A_CaretY

	;topwindow:= GetHex(DllCall("GetLastActivePopup", "uint", AlbisWinID()))
	;MsgBox, % AlbisWinID() "`n" topwindow "`n" WinGetTitle(topwindow) "`n" WinGetClass(topwindow)

return
;}

Thread14:																		;{ Quartalsaktuelles Tagesprotokoll erstellen

	_AlbisErstelleTagesprotokoll("04.01.2016-06.01.2016", AddendumDir "\Tagesprotokolle\TP-Abrechnungshelfer", 1)

	Sleep, 12000

return

_AlbisErstelleTagesprotokoll(period:="", SaveFolder:="", CloseProtokoll:=1) {

	;Parameter period	: 	    1. leerer String                         	- setzt den Zeitraum auf den ersten und letzten Tag des aktuellen Quartals
	;									2. mmyy (z.B. 0219)                 	- es wird ein Tagesprotokoll mit dem ersten und letzten Tag des übergebenen Quartals erstellt
	;									3. "dd.dd.dddd[,/-]dd.dd.dddd"	- von bis Datumsübergabe ist auch als String mit zwei Tagesdaten welche durch ein "-" oder "," getrennt sind möglich
	;
	; Beispiel               :		AlbisErstelleTagesprotokoll("04.01.2016-06.01.2016", AddendumDir "\Tagesprotokolle\TP-Abrechnungshelfer", 1)

		Quartal		:= Object()

	; TAGESPROTOKOLLFENSTER AUFRUFEN
		hTprotWin	:= Albismenu(32802, "Tagesprotokoll")

	; PARAMETER PARSEN
		If !InStr(period, ".") {
				Quartal.actual	      	:= GetQuartal("heute")							; um das Jahrtausend korrekt zu bestimmen
				Quartal.aktQuartal	:= SubStr(Quartal.actual, 1, 2)
				Quartal.aktJahr			:= SubStr(Quartal.actual, 3, 2)

				If period = ""
						period	        	:= GetQuartal("heute")
				Quartal.FuncQuartal	:= SubStr(period, 1, 2)
				Quartal.FuncJahr		:= SubStr(period, 3, 2)

				Quartal.LYear			:= Quartal.aktJahr < Quartal.FuncJahr ? ("19" Quartal.FuncJahr) : ("20" Quartal.FuncJahr)
				Quartal.FirstMonth	:= (Quartal.FuncQuartal - 1) * 3 + 1
				Quartal.LastMonth 	:= SubStr("0" . Quartal.FirstMonth + 2, -1)
				FirstDay		        	:= "01." SubStr("0" . Quartal.FirstMonth, -1) "." Quartal.LYear
				LastDay	               	:= _days_in_month( Quartal.LYear Quartal.LastMonth ) "." Quartal.LastMonth "." Quartal.LYear
		} else {
				If !RegExMatch(period, "^\d\d\.\d\d\.\d\d\d\d", FirstDay) or !RegExMatch(period, "(?<=-|,)\d\d\.\d\d\.\d\d\d\d", LastDay)
				{
						throw A_ThisFunc " (" A_LineFile ") : Error in function call, a date must be passed in the following format - """"dd.mm.yyyy[-/,]dd.mm.yyyy"""
						return 1
				}
		}

	; FELDER VON UND BIS MIT DEN JEWEILS ÜBERGEBEN DATEN FÜLLEN
		If !ControlCmd("Edit1", "SetText " FirstDay, hTprotWin)
				PraxTT("Das Anfangsdatum (von) konnte nicht gesetzt werden.", "6 3"), return 1
		If !ControlCmd("Edit2", "SetText " LastDay, hTprotWin)
				PraxTT("Das Enddatum (bis) konnte nicht gesetzt werden.", "6 3"), return 1

	; EINEN MAUSCLICK AN DEN 'OK' BUTTON DES FENSTER SENDEN
		If !ControlCmd("", "ControlClick, OK, Button", hTprotWin)
				ControlCmd("Button31", "Click use ControlClick", hTprotWin)
		WinWait, % "Bitte warten... ahk_class #32770",, 10
		If ErrorLevel and WinExist("Tagesprotokoll") {
				PraxTT("Der ausgelöste Click auf den 'OK' Button`nim Tagesprotokoll-Dialog ist fehlgeschlagen.`nEs wurde als letztes ein simulierter Mausklick versucht.", "6 3")
				WinClose, Tagesprotokoll ahk_class #32770
				return 1
		} else if ErrorLevel and !WinExist("Tagesprotokoll") {
				PraxTT("Es ist ein unbekannter Fehler aufgetreten,`nwelcher die Erstellung des Tagesprotokoll verhindert hat.", "6 3")
				return 1
		}

	; DATEN ZUM TAGESPROTOKOLLFENSTER IN CONTROLCMD LEEREN
		ControlCmd("","Reset", hTprotWin)

	; AUF DAS FERTIG ERSTELLTE TAGESPROTOKOLL WARTEN
		while, !WinExist("ALBIS - [Tagesprotokoll") {
				Sleep, 500
			; läßt den Loop pausieren
				If WinExist("bitte warten")
					WinWaitClose, % "Bitte warten... ahk_class #32770"
		}

	; SPEICHERN DES ANGEZEIGTEN TAGESPROTOKOLL AUF FESTPLATTE
		If SaveFolder {
			; Tagesprotokoll speichern und die Tagesprotokollausgabe schließen
				hSaveAsWin	:= Albismenu(33014, "Speichern unter")
				WinWaitActive, % "Speichern unter ahk_class #32770",, 10
				If ErrorLevel
					MsgBox, Erwartetes Speichern unter Fenster konnte nicht abgefangen werden!

			; ermitteln ob ein Tagesprotokoll mit selben Datumsangaben schon einmal erstellt wurde
				If FileExist(SaveFolder "\" Quartal.LYear "(" SubStr(FirstDay, 1, 6) "-" SubStr(LastDay, 1, 6) ".txt")
					tprotFileExists:=1

			; Ordner und Dateinamen im Speichern unter Dialog eintragen
				ControlCmd("Edit1", "SetText, " SaveFolder "\" Quartal.LYear "(" SubStr(FirstDay, 1, 6) "-" SubStr(LastDay, 1, 6) ").txt", hSaveAsWin)

			; Speichern unter drücken
				ControlCmd("Button2", "Click use Controlclick", hSaveAsWin)

			; Warten auf einen eventuellen "Speichern unter bestätigen" - Dialog
				If tprotFileExist
				{
					WinWaitActive, % "Speichern unter bestätigen ahk_class #32770",, 20
					ControlCmd("Button1", "Click use ControlClick", "Speichern unter bestätigen ahk_class #32770")
				}
		}

	; SCHLIESSEN DES PROTOKOLLS, WENN PARAMETER GESETZT
		If CloseProtokoll
			AlbisCloseMDITab("Tagesprotokoll")

return 0		; 0 = erfolgreicher Funktionsablauf,
}

_days_in_month(date:="") {
    date := (date = "") ? (a_now) : (date)
    FormatTime, year,  % date, yyyy
    FormatTime, month, % date, MM
    month += 1                 ; goto next month
    if (month > 12)
        year += 1, month := 1  ; goto next year, reset month
    month := (month < 10) ? (0 . month) : (month)  ; 0 to 01
    new_date := year . month
    new_date += -1, days       ; minus 1 day
    return subStr(new_date, 7, 2)
}


;}

Thread13:																		;{ Albis AutoLogin

	AlbisActivate(2)
	VerifiedSetText("Edit1", "Text1", "ALBIS - Login ahk_class #32770", 200)
	VerifiedSetText("Edit2", "Text2", "ALBIS - Login ahk_class #32770", 200)

return
;}

Thread12:																		;{ MCS VianovaBox

	MCS_WinClass  	:= "ahk_class WindowsForms10.Window.8.app.0.378734a"
	MCSinfoBoxID    	:= WinExist("MCS vianova infoBox-webClient")
	ACC_Init()
	;oAcc := Acc_Get("Object", "4.1.4.5.4.1.4.4.4.1.2" , 0, "ahk_id " MCSinfoBoxID)

	Loop
		{
				oAcc := Acc_Get("Object", "4.1.4.5.4.1.4.4.4.2", 0, "ahk_id " MCSinfoBoxID)
				If Instr(oAcc.accValue(2), "Keine Daten zum Abruf") {
						msg:="Es liegen keine neuen Labordaten vor!`nDer Laborabruf wird nicht fortgesetzt."
						Laborabruf_Daten := 0
						break
				} else if RegExMatch(oAcc.accValue(2), "(?<=^Sichere\s<).*(?=>\sins\sArchiv)", LDTFile) {
						LaborAbruf_Daten := 1
						msg:="Es liegen neue Labordaten vor!`nDer Laborabruf wird fortgesetzt."
						break
				}
				sleep 50
				If A_Index > 400
						break
		}

	MsgBox, % "oACC: " msg "`nID: " MCSinfoBoxID "`nValue: " oAcc.accValue(2) ", " oAcc.accValue(1)

return
;}

Thread11:																		;{ Albis COM

	SetWorkingDir, C:
	albis:= AlbisApplication()

	MsgBox, % albis.hello()

	ExitApp

	AlbisApplication() {
		try
		{
			If (impl:=ComObjCreate("{9ACC7108-9C10-4A49-A506-0720E0AACE32}","{6263C698-9393-4377-A6CC-4CB63A6A567A}"))
				return impl
			throw "IAlbisApplication Interface failed to initialize."
		}
		catch e
			MsgBox, 262160, IAlbisApplication Error, % IsObject(e)?"IAlbisApplication Interface is not registered.":e.Message
	}
return
;}

Thread10:																		;{ FoxitSignaturSetzen

	;Start gdi+
	If !pToken := Gdip_Startup()
	{
		MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
		ExitApp
	}

	ReaderID:= WinActive("A")
	WinGetTitle, t, ahk_id %ReaderID%
	ToolTip, % ReaderID ", " t
	FoxitReader_SignaturSetzen(ReaderID, "classFoxitReader")
	;FoxitInvoke("Place_Signature", ReaderID)

return



FoxitReader_SignaturSetzen(ReaderID, PDFReaderWinClass) {									;-- ruft Signatur setzen auf und zeichnet eine Signatur in die linke obere Ecke des Dokumentes
		/*
		CoordMode, ToolTip, Screen
		CoordMode, Pixel, Screen
		CoordMode, Mouse, Screen
		;----------------------------------------------------------------------------------------------------------------------------------------------
		; Ermitteln des Childhandles (Docwindow) im Foxitreader (Bildschirmposition des Dokuments)
		;----------------------------------------------------------------------------------------------------------------------------------------------
			hDocWnd		:= FindChildWindow({ID:ReaderID}, {Class:"FoxitDocWnd1"}, "On")
			hDocParent	:= FindChildWindow({ID:ReaderID}, {Class:"AfxFrameOrView140su1"}, "On")
			;ControlClick,, % "ahk_id " hDocWnd

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; FoxitReader vorbereiten für das Platzieren der Signatur
		;----------------------------------------------------------------------------------------------------------------------------------------------
			WinActivate, % "ahk_id " ReaderID
			WinWaitActive, % "ahk_id " ReaderID,,2

			;result2:= FoxitInvoke("SinglePage"			, hDocParent)
			;result2:= FoxitInvoke("SinglePage"			, ReaderID)
			;Sleep, 250
			;result3:= FoxitInvoke("Fit_Page"				, ReaderID)
			;sleep, 250
			;result4:= FoxitInvoke("firstPage"				, ReaderID)
			;sleep, 150

			PraxTT("sende Befehl: 'Signatur platzieren'", "12 2")
			WinGetPos, Fwx, Fwy	, Fww, Fwh , % "ahk_id " hDocWnd
			MouseClick, Left, % Fwx + 10, % Fwy + 10, 1, 0
			*/


			;hRec:= DrawRectangle(Fwx,Fwy,Fww,Fwh)
			;Sleep, 1000
			;WinHide, ahk_id %hRec%
			step:= 10
			hDocWnd		:= FindChildWindow({ID:ReaderID}, {Class:"FoxitDocWnd1"}, "On")
			WinGetPos, Fwx, Fwy	, Fww, Fwh , % "ahk_id " hDocWnd

			newx:=Fwx
			Loop, % (Fww//step)
			{
						newx += step
						PixelGetColor, color, % newx, % Fwy + 30, Alt
						If GetDec(color) > 16700000
						   	break
			}

			neww:= Fww
			Loop, % (Fww//step)
			{
						neww -=  step
						PixelGetColor, color, % neww, % Fwy + 10
						ToolTip, % Color
						Sleep, 100
						If GetDec(color) > 16700000
							break
			}

			y:=Fwy
			Loop, % (Fwh//step)
			{
						y += step
						PixelGetColor, color, % newx + 10, % y
						If GetDec(color) > 16700000
						{
							newy:= y
							break
						}
			}

			h:= Fwh
			Loop, % (Fwh//step)
			{
						h -=  step
						PixelGetColor, color, % newx + 10, % h
						If GetDec(color) > 16700000
						{
							newh:= h - newy
							break
						}
			}

			ListVars
			hRec:= DrawRectangle(newx,newy,neww,newh)
			Sleep, 4000
			return
		;----------------------------------------------------------------------------------------------------------------------------------------------
		; Signatur setzen Menupunkt aufrufen
		;----------------------------------------------------------------------------------------------------------------------------------------------
			result5:= FoxitInvoke("Place_Signature", ReaderID)
			PraxTT("Befehl: 'Signatur platzieren' gesendet", "12 2")
			sleep, 550

			;FileAppend, % "ReaderID: " ReaderID ", 1: " result1 ", 2: " result2 ", 3: " result3 ", 4: " result4 ", 5: " result5 "`n", %A_ScriptDir%\signLog.txt

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; sucht nach der linken oberen Ecke der Pdf Seite der 1.Seite, um dort im Anschluss die Signatur zu erstellen
		;----------------------------------------------------------------------------------------------------------------------------------------------
		/*
			PraxTT(" suche nach dem Signierbereich des Dokumentes.", "12 2")
				tryCount:=0, basetolerance := 0
			;!NICHT ÄNDERN! dieser String wird für 'feiyus' FindText() Funktion benötigt, er entspricht der linken oberen Ecke des PDF-Frame (AfxWnd100su4) ;{
				TopLeft:="|<>*210$71.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzw0000003zzzzs0000007zzzzk000000DzzzzU000000Tzzzz0000000zzzzy0000001zzzzw0000003zzzzs0000007zzzzk000000DzzzzU000000Tzzzz0000000zzzzy0000001zzzzw0000003zzzzs0000007zzzzk000000DzzzzU000000Tzzzz0000000zzzzy0000001"
			;}
			FindSignatureRange:
			;sucht hier im Prinzip nach einem Bild (entspricht der linken oberen Ecke des PDF Preview Bereiches) und nicht nach einem Text.
				if (ok:=FindText(Fwx, Fwy, Fww, Fwh, basetolerance, 0, TopLeft))
				{
					PraxTT("Signierbereich gefunden.", "4 2")
					X:=ok.1.1, Y:=ok.1.2, W:=ok.1.3, H:=ok.1.4, Comment:=ok.1.5, X+=W//2, Y+=H//2
					MouseClickDrag, Left, % x, % y, % x + 50, % y + 50, 0
				}
				else
				{
						sleep, 100
						tryCount++, basetolerance += 0.1
						If (tryCount < 20)
							goto FindSignatureRange
						else
							return 0
				}
		*/

return 1
}
FoxitInvoke(command, FoxitID:="") {		                                                            	;-- wm_command wrapper for FoxitReader Version:  9.1

		/* DESCRIPTION of FUNCTION:  FoxitInvoke() by Ixiko (version 2018/11/8)

		---------------------------------------------------------------------------------------------------
												a WM_command wrapper for FoxitReader V9.1 by Ixiko
																		...........................................................
													 Remark: maybe not all commands are listed at now!
		---------------------------------------------------------------------------------------------------
				by use  of a valid FoxitID, this function will post your command to FoxitReader
			                                             otherwise this function returns the command code
																		...........................................................
			Remark: You have to control the success of the postmessage command yourself!
		---------------------------------------------------------------------------------------------------
						I intentionally use a text first and then convert it to a -Key: Value- object,
                                                         so you can swap out the object to a file if needed
		---------------------------------------------------------------------------------------------------
		      EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES

		FoxitInvoke("Show_FullPage")                       FoxitInvoke("Place_Signature", FoxitID)
		.................................................          ...............................................................
		this one only returns the Foxit                  sends the command "Place_Signature" to
        command-code                                      your specified FoxitReader process using
																	 parameter 2 (FoxitID) as window handle.
																	        command-code will be returned too
		---------------------------------------------------------------------------------------------------

	*/

	static FoxitCommand:= [], run:=0

	If !run
	{
		FoxitCommands =
		(Comments
			SaveAs                                                             	: 1299
			Close                                                               	: 57602
			Hand                                                               	: 1348       	;Home - Tools
			Select_Text                                                       	: 46178     	;Home - Tools
			Select_Annotation                                             	: 46017     	;Home - Tools
			Snapshot                                                         	: 46069     	;Home - Tools
			Clipboard_SelectAll                                          	: 57642     	;Home - Tools
			Clipboard_Copy                                               	: 57634     	;Home - Tools
			Clipboard_Paste                                               	: 57637     	;Home - Tools
			Actual_Size                                                      	: 1332       	;Home - View
			Fit_Page                                                           	: 1343       	;Home - View
			Fit_Width                                                         	: 1345        	;Home - View
			Reflow                                                             	: 32818     	;Home - View
			Zoom_Field                                                      	: 1363       	;Home - View
			Zoom_Plus                                                       	: 1360       	;Home - View
			Zoom_Minus                                                   	: 1362       	;Home - View
			Rotate_Left                                                      	: 1340       	;Home - View
			Rotate_Right                                                    	: 1337       	;Home - View
			Highlight                                                         	: 46130     	;Home - Comment
			Typewriter                                                        	: 46096     	;Home - Comment, Comment - TypeWriter
			Open_From_File                                              	: 46140     	;Home - Create
			Open_Blank                                                     	: 46141     	;Home - Create
			Open_From_Scanner                                       	: 46165     	;Home - Create
			Open_From_Clipboard                                    	: 46142     	;Home - Create - new pdf from clipboard
			PDF_Sign                                                      	: 46157     	;Home - Protect
			Create_Link                                                     	: 46080     	;Home - Links
			Create_Bookmark                                            	: 46070     	;Home - Links
			File_Attachment                                               	: 46094     	;Home - Insert
			Image_Annotation                                            	: 46081     	;Home - Insert
			Audio_and_Video                                             	: 46082     	;Home - Insert
			Comments_Import                                          	: 46083     	;Comments
			Highlight                                                         	: 46130     	;Comments - Text Markup
			Squiggly_Underline                                          	: 46131     	;Comments - Text Markup
			Underline                                                         	: 46132     	;Comments - Text Markup
			Strikeout                                                          	: 46133     	;Comments - Text Markup
			Replace_Text                                                    	: 46134     	;Comments - Text Markup
			Insert_Text                                                       	: 46135     	;Comments - Text Markup
			Note                                                                	: 46137     	;Comments - Pin
	    	File                                                                  	: 46095     	;Comments - Pin
	    	Callout                                                            	: 46097     	;Comments - Typewriter
	    	Textbox                                                            	: 46098     	;Comments - Typewriter
	    	Rectangle                                                         	: 46101     	;Comments - Drawing
	    	Oval                                                                	: 46102     	;Comments - Drawing
	    	Polygon                                                           	: 46103     	;Comments - Drawing
	    	Cloud                                                              	: 46104     	;Comments - Drawing
	    	Arrow                                                              	: 46105     	;Comments - Drawing
	    	Line                                                                 	: 46106     	;Comments - Drawing
	    	Polyline                                                            	: 46107     	;Comments - Drawing
	    	Pencil                                                              	: 46108     	;Comments - Drawing
	    	Eraser                                                              	: 46109     	;Comments - Drawing
	    	Area_Highligt                                                   	: 46136     	;Comments - Drawing
	    	Distance                                                          	: 46110     	;Comments - Measure
	    	Perimeter                                                         	: 46111     	;Comments - Measure
	    	Area                                                                	: 46112     	;Comments - Measure
	    	Stamp				                                               	: 46149     	;Comments - Stamps , opens only the dialog
	    	Create_custom_stamp                                      	: 46151     	;Comments - Stamps
	    	Create_custom_dynamic_stamp                        	: 46152     	;Comments - Stamps
	    	Summarize_Comments                                   	: 46188     	;Comments - Manage Comments
	    	Import                                                             	: 46083     	;Comments - Manage Comments
	    	Export_All_Comments                                      	: 46086     	;Comments - Manage Comments
	    	Export_Highlighted_Texts                                  	: 46087     	;Comments - Manage Comments
	    	FDF_via_Email                                                 	: 46084     	;Comments - Manage Comments
	    	Comments                                                       	: 46088     	;Comments - Manage Comments
	    	Comments_Show_All                                        	: 46089     	;Comments - Manage Comments
	    	Comments_Hide_All                                        	: 46090     	;Comments - Manage Comments
	    	Popup_Notes                                                   	: 46091     	;Comments - Manage Comments
	    	Popup_Notes_Open_All                                    	: 46092     	;Comments - Manage Comments
	    	Popup_Notes_Close_All                                    	: 46093     	;Comments - Manage Comments
			firstPage                                                          	: 1286       	;View - Go To
			lastPage                                                           	: 1288       	;View - Go To
        	nextPage                                                         	: 1289       	;View - Go To
        	previousPage                                                   	: 1290       	;View - Go To
        	previousView                                                   	: 1335       	;View - Go To
        	nextView                                                          	: 1346       	;View - Go To
        	ReadMode                                                       	: 1351       	;View - Document Views
        	ReverseView                                                    	: 1353       	;View - Document Views
        	TextViewer                                                       	: 46180     	;View - Document Views
        	Reflow                                                             	: 32818     	;View - Document Views
        	turnPage_left                                                   	: 1340       	;View - Page Display
        	turnPage_right                                                 	: 1337       	;View - Page Display
        	SinglePage                                                      	: 1357       	;View - Page Display
        	Continuous                                                      	: 1338       	;View - Page Display
        	Facing                                                             	: 1356       	;View - Page Display - two pages side by side
        	Continuous_Facing                                           	: 1339       	;View - Page Display - two pages side by side with scrolling enabled
        	Separate_CoverPage                                        	: 1341       	;View - Page Display
        	Horizontally_Split                                             	: 1364       	;View - Page Display
        	Vertically_Split                                                  	: 1365       	;View - Page Display
        	Spreadsheet_Split                                             	: 1368       	;View - Page Display
        	Guides                                                            	: 1354       	;View - Page Display
        	Rulers                                                              	: 1355       	;View - Page Display
        	LineWeights                                                    	: 1350       	;View - Page Display
        	AutoScroll                                                        	: 1334       	;View - Assistant
        	Marquee                                                          	: 1361       	;View - Assistant
        	Loupe                                                              	: 46138     	;View - Assistant
        	Magnifier                                                         	: 46139     	;View - Assistant
        	Read_Activate                                                  	: 46198     	;View - Read
        	Read_CurrentPage                                           	: 46199     	;View - Read
        	Read_from_CurrentPage                                  	: 46200     	;View - Read
        	Read_Stop                                                       	: 46201     	;View - Read
        	Read_Pause                                                     	: 46206     	;View - Read
			Navigation_Panels                                           	: 46010      	;View - View Setting
			Navigation_Bookmark                                     	: 45401      	;View - View Setting
			Navigation_Pages                                            	: 45402     	;View - View Setting
			Navigation_Layers                                           	: 45403     	;View - View Setting
			Navigation_Comments                                    	: 45404     	;View - View Setting
			Navigation_Appends                                       	: 45405     	;View - View Setting
			Navigation_Security                                        	: 45406     	;View - View Setting
			Navigation_Signatures                                    	: 45408     	;View - View Setting
			Navigation_WinOff                                          	: 1318       	;View - View Setting
			Navigation_ResetAllWins                                	: 1316       	;View - View Setting
			Status_Bar                                                        	: 46008       	;View - View Setting
			Status_Show                                                    	: 1358       	;View - View Setting
			Status_Auto_Hide                                             	: 1333       	;View - View Setting
			Status_Hide                                                     	: 1349       	;View - View Setting
			WordCount                                                     	: 46179     	;View - Review
			Form_to_sheet                                                 	: 46072     	;Form - Form Data
			Combine_Forms_to_a_sheet                            	: 46074     	;Form - Form Data
			DocuSign                                                         	: 46189     	;Protect
			Login_to_DocuSign                                           	: 46190     	;Protect
			Sign_with_DocuSign                                         	: 46191     	;Protect
			Send_via_DocuSign                                          	: 46192     	;Protect
			Sign_and_Certify                                              	: 46181     	;Protect
			-----_-------------                                              	: 46182     	;Protect
			Place_Signature                                              	: 46183     	;Protect
			Validate                                                           	: 46185     	;Protect
			Time_Stamp_Document                                    	: 46184     	;Protect
			Digital_IDs                                                       	: 46186     	;Protect
			Trusted_Certificates                                          	: 46187     	;Protect
			Email                                                               	: 1296       	;Share - Send To - same like Email current tab
			Email_All_Open_Tabs                                      	: 46012       	;Share - Send To
			Tracker                                                            	: 46207       	;Share - Tracker
			User_Manual                                                   	: 1277       	;Help - Help
			Help_Center                                                    	: 558        	;Help - Help
			Command_Line_Help                                       	: 32768       	;Help - Help
			Post_Your_Idea                                                	: 1279        	;Help - Help
			Check_for_Updates                                          	: 46209       	;Help - Product
			Install_Update                                                  	: 46210       	;Help - Product
			Set_to_Default_Reader                                    	: 32770       	;Help - Product
			Foxit_Plug-Ins                                                   	: 1312        	;Help - Product
			About_Foxit_Reader                                          	: 57664       	;Help - Product
			Register                                                           	: 1280        	;Help - Register
			Open_from_Foxit_Drive                                   	: 1024        	;Extras - maybe this is not correct!
			Add_to_Foxit_Drive                                           	: 1025        	;Extras - maybe this is not correct!
			Delete_from_Foxit_Drive                                   	: 1026        	;Extras - maybe this is not correct!
			Options                                                           	: 243           	;the following one's are to set directly any options
			Use_single-key_accelerators_to_access_tools  	: 128           	;Options/General
			Use_fixed_resolution_for_snapshots                	: 126           	;Options/General
			Create_links_from_URLs                                   	: 133           	;Options/General
			Minimize_to_system_tray                                 	: 138           	;Options/General
			Screen_word-capturing                                    	: 127           	;Options/General
			Make_Hand_Tool_select_text                          	: 129           	;Options/General
			Double-click_to_close_a_tab                           	: 91             	;Options/General
			Auto-hide_status_bar                                      	: 162           	;Options/General
			Show_scroll_lock_button                                 	: 89             	;Options/General
			Automatically_expand_notification_message   	: 1725         	;Options/General - only 1 can be set from these 3
			Dont_automatically_expand_notification         	: 1726         	;Options/General - only 1 can be set from these 3
			Dont_show_notification_messages_again     		: 1727         	;Options/General - only 1 can be set from these 3
			Collect_data_to_improve_user_&experience    	: 111           	;Options/General
			Disable_all_features_which_require_internet 		: 562           	;Options/General
			Show_Start_Page                                             	: 160           	;Options/General
			Change_Skin                                                   	: 46004
			Filter_Options                                                  	: 46167       	;the following are searchfilter options
			Whole_words_only                                          	: 46168       	;searchfilter option
			Case-Sensitive                                                 	: 46169       	;searchfilter option
			Include_Bookmarks                                         	: 46170       	;searchfilter option
			Include_Comments                                          	: 46171       	;searchfilter option
			Include_Form_Data                                          	: 46172       	;searchfilter option
			Highlight_All_Text                                           	: 46173       	;searchfilter option
			Filter_Properties                                              	: 46174       	;searchfilter option
			Print                                                                	: 57607
			Properties                                                        	: 1302         	;opens the PDF file properties dialog
			Mouse_Mode                                                  	: 1311
			Touch_Mode                                                   	: 1174
			predifined_Text                                               	: 46099
			set_predefined_Text                                        	: 46100
			Create_Signature                                             	: 26885     	;Signature
			Draw_Signature                                               	: 26902     	;Signature
			Import_Signature                                            	: 26886     	;Signature
			Paste_Signature                                               	: 26884     	;Signature
			Type_Signature                                                	: 27005     	;Signature
			Pdf_Sign_Close                                                	: 46164     	;Pdf-Sign
		)

		FoxitCommands:= StrReplace(FoxitCommands, A_Space, "")

		Loop, Parse, FoxitCommands, `n
		{
				line	:= RegExReplace(A_LoopField, ";.*", "")
				split	:= StrSplit(line, ":")
				FoxitCommand[Trim(split[1])]:= Trim(split[2])
		}

		run:= 1
	}

	If FoxitID {
		SendMessage, 0x111, % FoxitCommand[command],,, ahk_id %FoxitID%
		return ErrorLevel
	}
	else
		return FoxitCommand[command]


}
DrawRectangle(x,y,width,height) {

		Gui, rec: -Caption +E0x80000 +LastFound +AlwaysOnTop +ToolWindow +OwnDialogs  +Disabled +E0x8000000
		;Gui, rec: Add, Picture, % "x" 0 " y" 0 " w" Width " h" Height
		Gui, rec: Show, x%x% y%y% NA

		hwnd1 := WinExist()
		hbm := CreateDIBSection(width, Height)
		hdc := CreateCompatibleDC()
		obm := SelectObject(hdc, hbm)
		G := Gdip_GraphicsFromHDC(hdc)
		Gdip_SetSmoothingMode(G, 4)
		pBrush := Gdip_BrushCreateSolid(0x660000ff)
		Gdip_FillRectangle(G, pBrush, 0, 0, width, Height)
		Gdip_DeleteBrush(pBrush)
		UpdateLayeredWindow(hwnd1, hdc, 0, 0, Width, Height)
		SelectObject(hdc, obm)
		DeleteObject(hbm)
		DeleteDC(hdc)
		Gdip_DeleteGraphics(G)
		WinMove, ahk_id %hwnd1%,, %x%, %y%

return hwnd1
}
;}

Thread9:																		;{ schöne Hotkey Gui


	;--Anzeige aller benutzten Hotkeys mit Beschreibung durch auslesen des Skriptes
	script		:= "M:\Praxis\Skripte\Skripte Neu\Addendum für AlbisOnWindows\Module\Addendum\Addendum.ahk"
	activeApp := ActiveContext()

ViewKeyList(script)

return

HumanReadableHK( key ) {
;a function to take care of replacing symbols in hotkeys with their words
  key:= 	 StrReplace(key, "+"			, "Shift + "			)
  key:=  RegExReplace(key, "#(?!=\s)", "Win + "			)
  key:= 	 StrReplace(key, "!"				, "Alt links + "	)
  key:= 	 StrReplace(key, "LAlt"			, "Alt links + "	)
  key:= 	 StrReplace(key, "^"			, "Strg + "			)
  key:= 	 StrReplace(key, "LControl"	, "Strg + "			)
  key:= RegExReplace(key, "\s\&\s"	, " + ")
  key:= RegExReplace(key, "\*(?=\^|\!|\+|\#|\$|\~|\w)"	, "")
  key:= RegExReplace(key, "\$(?=\^|\!|\+|\#|\*|\~|\w)"	, "")
  key:= RegExReplace(key, "\~(?=\^|\!|\+|\#|\$|\*|\w)"	, "")
  return key
}

ViewKeyList(script) {

	KeyList:= Object()
	KeyList:= FormatKeyList(script)
	activeContext:= activeContext()

	For idx, key in KeyList
		If !KeyList.HasKey(key.cat)
			KeyList[Key.cat] := 1
		else
			KeyList[Key.cat] += 1

	;MsgBox, % KeyList["Albis"] "`n" KeyList["Überall"] "`n" KeyList["SciTE"] "`n" KeyList[1]["cat"] ":" KeyList[1]["hk"] "/" KeyList[1]["des"] "/" KeyList[1]["test"]

	Gui, KL: New
	Gui, KL: Add, ListView,
}

FormatKeyList(file, srt:= False) {										; this does the actual formatting. Pass the text to be formatted

	fileread, scriptlines, % file
	f:= FileOpen(A_ScriptDir "\Hotkey.txt", "w")
	KeyList := Object(), idx:= 0

	Loop, Parse, scriptlines, `n, `r
	{
			RegExMatch(A_LoopField, "i)(?<=\sHotkey\,).*(?=\s\,\s\w)", hk)
			if hk
			{
				RegExMatch(A_LoopField, "i)(?<=\s\;\=\s).*(?=\:\s)", cat)
				RegExMatch(A_LoopField, "i)(?<=\s\;\=\s).*", des)
				hk:= Trim(hk), des:= Trim(des), cat:= Trim(cat)
			}

			If cat && des && hk
			{
				idx ++
				keyname		:= HumanReadableHK(Trim(hk))
				description	:= StrReplace(des				, cat	, "")
				description  	:= StrReplace(description	, "`:"	, "")
				f.Write(A_Index ": Nr." idx " - " cat "|" keyname "|" description "`n")
				KeyList[(idx)]	:= {"cat": cat, "hk": keyname, "des": description, "test": "warum"}
			}

			hk:=des:=cat:=""
	}

	f.Close()

return KeyList
}

activeContext() {																	;-- gibt einen Identifikator für das aktive Fenster zurück

	If InStr(activeClass	:= WinGetClass(WinExist("A")), "OptoAppClass")
					return "Albis"
	else if InStr(activeClass, "Scite")
					return "Scite"
	else if InStr(activeClass, "Qt5QWindowIcon") and InStr(WinGetTitle(WinExist("A")), "Telegram")
					return "Telegram"

}

;}

Thread8Start:																;{ Foxit Reader

MsgBox, % activeWT "`n" WinExist("ahk_class OptoAppClass")
WinActivate, % "ahk_id " WinExist("ahk_class OptoAppClass") ;ahk_class OptoAppClass

PDFReaderWinClass:= "classFoxitReader"

gosub Thread8
;^F12::gosub Thread8

ExitApp

Thread8:																		;{ neues FoxitFenster abfangen

		PraxTT("Warte auf neues " PDFReaderName " - Fenster.", "30 2")
		WinWaitActive, % "ahk_class " PDFReaderWinClass,, 30
		WinGetTitle, ActiveReaderTitle, "ahk_class " PDFReaderWinClass
		ReaderID := WinExist(ActiveReaderTitle " ahk_class " PDFReaderWinClass)
		MsgBox, % ReaderID

		SendMessage 0x10, 0, 0,, ahk_id %ReaderID% 			; WM_Close
		WinWaitClose,  % "ahk_id " ReaderID,, 10
		If ErrorLevel
		{
			SendMessage 0x2, 0, 0,, ahk_id %ReaderID% 	; WM_Destroy
			WinWaitClose, % "ahk_id " ReaderID,, 10
			If ErrorLevel
				MsgBox, Ach jet nisch
		}


return
;}
;}

Thread7: 																		;{ Patientenakte zeigen

		If !InStr(activeWT:= AlbisGetActiveWindowType(), "Patientenakte")
						MsgBox , Keine geöffnete Patientenakte

		MsgBox, % activeWT "`n" WinExist("ahk_class OptoAppClass")

		If !InStr(activeWT, "Karteikarte")
		{
			PostMessage, 0x111, 33033,,, ahk_class OptoAppClass
			sleep, 300
					;	If !InStr(AlbisGetActiveWindowType(), "Karteikarte")
					;	{
					;			MsgBox, 1, Addendum für Albis on Windows, % "Die Karteikarte des aktuellen Patienten ließ sich nicht öffnen.`nDie Funktion wird jetzt beendet!", 6
					;			return 0
					;	}

		}
		WinActivate, ahk_class SciTEWindow

return
;}

Thread6: 																		;{ GetFont from focused AlbisControl
;WinActivate, ahk_class OptoAppClass

;^F12::gosub GetFontControl

return

GetFontControl:
hwin:= WinExist("A")
ControlGetFocus, className, ahk_id %hWin%
ControlGet, hwnd, HWND,, % className,  ahk_id %hWin%
;Control_GetFont(hwnd, FName, FSize, FStyle)
RichEdit_GetCharFormat(hwnd, FName, FStyle, tColor, bColor, "NOSELECTION")
;MsgBox, % "Class:  " className "`nhwnd:  " hwnd "`nFont:   " FName "`nSize:   " FSize "`nStyle:   " FStyle ;"`n" bColor
MsgBox, % "Class:  " className "`nhwnd:  " hwnd "`nFont:    " FName "`nStyle:    " FSytle "`ntColor:  " tColor "`nbColor:  " bColor
WinActivate, ahk_class SciTEWindow
ExitApp

Control_GetFont(hWnd, ByRef Name, ByRef Size, ByRef Style, IsGDIFontSize := 0) {

    SendMessage 0x31, 0, 0, , ahk_id %hWnd% ; WM_GETFONT
    If (ErrorLevel == "FAIL") {
        Return
    }

    hFont := Errorlevel
    VarSetCapacity(LOGFONT, LOGFONTSize := 60 * (A_IsUnicode ? 2 : 1 ))
    DllCall("GetObject", "Ptr", hFont, "Int", LOGFONTSize, "Ptr", &LOGFONT)

    Name := DllCall("MulDiv", "Int", &LOGFONT + 28, "Int", 1, "Int", 1, "Str")

    Style := Trim((Weight := NumGet(LOGFONT, 16, "Int")) == 700 ? "Bold" : (Weight == 400) ? "" : " w" . Weight
    . (NumGet(LOGFONT, 20, "UChar") ? " Italic" : "")
    . (NumGet(LOGFONT, 21, "UChar") ? " Underline" : "")
    . (NumGet(LOGFONT, 22, "UChar") ? " Strikeout" : ""))

    Size := IsGDIFontSize ? -NumGet(LOGFONT, 0, "Int") : Round((-NumGet(LOGFONT, 0, "Int") * 72) / A_ScreenDPI)
}
RichEdit_GetCharFormat(hCtrl, ByRef Face="", ByRef Style="", ByRef TextColor="", ByRef BackColor="", Mode="SELECTION")  {

	static EM_GETCHARFORMAT=1082, SCF_SELECTION=1, SCF_NOSELECTION=0
  		  , CFM_CHARSET:=0x8000000, CFM_BACKCOLOR=0x4000000, CFM_COLOR:=0x40000000, CFM_FACE:=0x20000000, CFM_OFFSET:=0x10000000, CFM_SIZE:=0x80000000, CFM_WEIGHT=0x400000, CFM_UNDERLINETYPE=0x800000
		  , CFE_HIDDEN=0x100, CFE_BOLD=1, CFE_ITALIC=2, CFE_LINK=0x20, CFE_PROTECTED=0x10, CFE_STRIKEOUT=8, CFE_UNDERLINE=4, CFE_SUPERSCRIPT=0x30000, CFE_SUBSCRIPT=0x30000
		  , CFM_ALL2=0xFEFFFFFF, COLOR_WINDOW=5, COLOR_WINDOWTEXT=8
		  , styles="HIDDEN BOLD ITALIC LINK PROTECTED STRIKEOUT UNDERLINE SUPERSCRIPT SUBSCRIPT"

	VarSetCapacity(CF, 84, 0), NumPut(84, CF), NumPut(CFM_ALL2, CF, 4)
	SendMessage, EM_GETCHARFORMAT, SCF_%Mode%, &CF,, ahk_id %hCtrl%

	Face := DllCall("MulDiv", "UInt", &CF+26, "Int",1, "Int",1, "str")

	Style := "", dwEffects := NumGet(CF, 8, "UInt")
	Loop, parse, styles, %A_SPACE%
		if (CFE_%A_LoopField% & dwEffects)
			Style .= A_LoopField " "
    s := NumGet(CF, 12, "Int") // 20,  o := NumGet(CF, 16, "Int")
	Style .= "s" s (o ? " o" o : "")

	oldFormat := A_FormatInteger
    SetFormat, integer, hex

	if (dwEffects & CFM_BACKCOLOR)
		 BackColor := "-" DllCall("GetSysColor", "int", COLOR_WINDOW)
	else BackColor := NumGet(CF, 64), BackColor := (BackColor & 0xff00) + ((BackColor & 0xff0000) >> 16) + ((BackColor & 0xff) << 16)

	if (dwEffects & CFM_COLOR)
		 TextColor := "-" DllCall("GetSysColor", "int", COLOR_WINDOWTEXT)
	else TextColor := NumGet(CF, 20), TextColor := (TextColor & 0xff00) + ((TextColor & 0xff0000) >> 16) + ((TextColor & 0xff) << 16)

    SetFormat, integer, %oldFormat%
}


;}

Thread5: 																		;{ AutoCompleteLb Test - erfolgreich
	oACB:= Object()

	List = Mama|Papa|Bruder|Schwester|Opa|Omi|Onkel|Tante|Cousin|Cousine|Neffe|Nichte

	ControlGetFocus, cf, % "ahk_id " hwnd:= WinExist("A")
	ControlGetPos, cfX, cfY, cfW, cfH, % cf,  % "ahk_id " hwnd
	CrX:= A_CaretX, CrY:= A_CaretY

	Hotkey, IfWinExist, LbAutoComplete
	Hotkey, Enter		, AutoCBLBres
	Hotkey, Tab		, AutoCBLBres
	Hotkey, Down	, ACBMoveDown
	Hotkey, Up		, ACBMoveUp
	Hotkey, Esc		, ACBGuiClose
	Hotkey, LButton	, ACBGuiClose
	Hotkey, RButton	, ACBGuiClose


	Gui, ACB: New	, +Disabled -SysMenu -Caption +AlwaysonTop +hwndhACB ;0x98200000
	Gui, ACB: Add	, ListBox, % "Choose1 vACBLb r6", % List											;gAutoComplete
	Gui, ACB: Show, % "NA AutoSize"
	oACB	:= GetWindowInfo(hACB)
	ypos		:= CrY + 25 + oACB.WindowH
	ypos		:= ypos > A_ScreenHeight ? CrY - oACB.WindowH -5 : ypos
	Gui, ACB: Show, % "x" CrX - oACB.WindowW//2 " y" ypos " NA", LbAutoComplete

	ControlFocus, % cf, ahk_id %hwnd%

return

ACBMoveUp:
	ControlSend,, {Up}, ahk_id %hACB%
return

ACBMoveDown:
	ControlSend,, {Down}, ahk_id %hACB%
return

AutoCBLBres:
	Gui, ACB: Submit, NoHide
	Gui, ACB: Destroy
	Sleep, 200
	SendRaw, % ACBLb
ACBGuiClose:
	Gui, ACB: Destroy
	If InStr(A_ThisHotkey, "LButton")
			Send, {LButton}
	else If InStr(A_ThisHotkey, "RButton")
			Send, {RButton}
ExitApp
;}

Thread4:                                                                   	;{ Common SaveAs Dialog

WinActivate, Speichern unter ahk_class #32770
WinWaitActive, Speichern unter ahk_class #32770,,2
TBId:= ControlCmd("ToolbarWindow324","ID", "Speichern unter ahk_class #32770", true, true, "slow")
Sleep 100
ControlCmd("ToolbarWindow324","Click use Space", "Speichern unter ahk_class #32770", true, true, "slow")
MsgBox, % ControlCmd("Edit2","GetText", "Speichern unter ahk_class #32770", true, true, "slow")

ExitApp
;}

Thread3:                                                                     	;{ RegEx

T:= "Markus Dente Colloskopiebefund vom 04.11.2011.pdf - Foxit Reader"
MsgBox, % RegExReplace(T, "i)(?<=\.pdf).*", "#")


ExitApp
;}

Thread2:                     ;{
BefundOrdner	:= "M:\Befunde"
AlbisWinDir		:= "M:\albiswin"
WinActivate, Speichern unter ahk_class #32770
FoxitSpeichern:= WinExist("Speichern unter ahk_class #32770")
SELFLAG_TAKEFOCUS		 := 0x1
SELFLAG_TAKESELECTION := 0x2

accPath:= Object()
accPath["SpeichernUnter"]:= {"FilePathClickable": "4.13.4.1.4.5.4.1.4.1.4.1.4", "FilePathShort": "4.13.4.1.4.5.4.1.4.1.4.1.4.1.4", "Dateiname":"4.1.4.1.4.2.3.2.4.1.4"}

oAcc := Acc_Get("Object", accPath.SpeichernUnter.FilePathClickable, 0, "ahk_id " FoxitSpeichern)
SaveDir:= oAcc.accName(0)
If !InStr(SaveDir, BefundOrdner) and !InStr(SaveDir, AlbisWinDir "\Briefe")
{
		ToolTip, JUHU
		hwnd:= GetChildHWND(FoxitSpeichern, "Edit1")
		ControlGetText, filename, Edit1, ahk_id %FoxitSpeichern%
		If !InStr(filename, BefundOrdner)
				VerifiedSetText("Edit1", "M:\Befunde\" filename, "ahk_id " FoxitSpeichern, 200)
}

MsgBox, % "hwnd: " GetHex(FoxitSpeichern) "`nnSaveDir: " SaveDir "`nfilename: " filename "`nTb321: " oAcc.accName(0) "`nEdit1-hwnd: " GetHex(hwnd)


ExitApp
;}

Thread1:	;{
ControlGetText, txt1,	, % "ahk_id " hEdit1		:= 	ControlFind("Edit1"						, "ID", SpEHookHwnd1)
ControlClick,				, % "ahk_id " hTBW324	:= 	ControlFind("ToolBarWindow324"	, "ID", SpEHookHwnd1)
ControlGetText, txt2,	, % "ahk_id " hEdit2		:=	ControlFind("Edit2"						, "ID", SpEHookHwnd1)
If !InStr(txt2, BefundOrdner) and !RegExMatch(txt1, "(?<=\s)([P]?\d?\d\d\d\d\d[A-Z][A-Z])(?=\.)", match)
					ControlSetText,, % BefundOrdner , % "ahk_id " hEdit2
MsgBox, % txt1 "`n" txt2 "`n" match

return
;}



<<<<<<< Updated upstream


#Include %A_ScriptDir%\..\..\..\include\AddendumFunctions.ahk
#Include %A_ScriptDir%\..\..\..\include\AddendumDB.ahk
#Include %A_ScriptDir%\..\..\..\include\Gdip_All.ahk
#Include %A_ScriptDir%\..\..\..\include\FindText.ahk
#Include %A_ScriptDir%\..\..\..\Module\Albis_Funktionen\ScanPool\lib\ScanPool_PdfHelper.ahk
#Include %A_ScriptDir%\..\..\..\include\ACC.ahk
#Include %A_ScriptDir%\..\..\..\include\ini.ahk
#Include %A_ScriptDir%\..\..\..\include\Sift.ahk
#include %A_ScriptDir%\..\..\..\include\Gui\PraxTT.ahk
=======
;#Include %A_ScriptDir%\..\..\..\include\Addendum_Functions.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Menu.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Misc.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_PdfHelper.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Window.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_StackifyGui.ahk
#include %A_ScriptDir%\..\..\..\include\Gui\PraxTT.ahk
#Include %A_ScriptDir%\..\..\..\lib\Gdip_All.ahk
#Include %A_ScriptDir%\..\..\..\lib\FindText.ahk
#Include %A_ScriptDir%\..\..\..\lib\ACC.ahk
#Include %A_ScriptDir%\..\..\..\lib\ini.ahk
#Include %A_ScriptDir%\..\..\..\lib\class_JSONFile.ahk
#Include %A_ScriptDir%\..\..\..\lib\Sci.ahk
#Include %A_ScriptDir%\..\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\..\lib\Sift.ahk
#Include %A_ScriptDir%\..\..\..\lib\class_GuiControlTips.ahk
;#Include %A_ScriptDir%\..\..\..\lib\class_CtlColors.ahk
>>>>>>> Stashed changes
;#include %A_ScriptDir%\..\lib\ScanPool_PdfHelper.ahk


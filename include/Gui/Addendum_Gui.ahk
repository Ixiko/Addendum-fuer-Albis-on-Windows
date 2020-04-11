;###############################################################
;------------------------------------------------------------------------------------------------------------------------------------
;--------------------------------------------- Addendum für AlbisOnWindows -----------------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------
;----------------------------------------------------- zusätzliche Gui's- ---------------------------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------


SprechzimmerAuffuellen(nr) {  																						;--die Schwestern vergessen immer ein Sprechzimmer mit Formularen und Arbeitsmaterialien zu befüllen

;{diese GUI ist für eine Auflösung von 1920x1080 am Besten geeignet
	Global aufLogo1,aufLogo2,aufLogo3, aufBuchstabeK, Schwesterchen1, Doc

	ich = %A_ScriptDir%\assets\ich.png
	schwester = %A_ScriptDir%\assets\Schwester1.png

	SysGet, Mon1, Monitor, 1
	LX:= Mon1Left, LY:=Mon1Top, LW:=Mon1Right, LH:=Mon1Bottom

	Gui, auf: New, +HWNDLhauf -Caption AlwaysOnTop
	Gui, auf: Color, c172842
	Gui, auf: Margin, 0, 0
	Gui, auf: Add, Picture, x70 y30 vSchwesterchen1 gaufClick, %schwester%
	Gui, auf: Add, Picture, % "x" (LW-340) " y30 vDoc gaufClick", %ich%
	Gui, auf: Font, s72 q4 cWhite, Futura Bk Md
	Gui, auf: Add, Text, % "x380 y70 w" (LW-760) " Center", Liebe Schwester1,
	Gui, auf: Font, s30 q4 cWhite, Futura Bk Md
	Gui, auf: Add, Text, % "x380 y200 w" (LW-760) " Center", heute ist Donnerstag! Und ich möchte Dich`nan das Auffüllen der Sprechzimmer errinnern!

	Gui, auf: Font, s14 q4 cWhite, Futura Bk Bt
		static xIt:= 120, nextx:= 500, yIt:=390, nexty:=24, boxes:=0
	Loop, Read, %A_ScriptDir%\DieListe.txt
	{
		If (A_LoopReadLine = "--") {
					xIt += nextx
					yIt := 390
					continue
			}

		If (A_LoopReadLine <> "") {
					boxes ++
					Gui, auf: Add, checkbox, % "x" xIt " y" yIt+5 " w12 h12"
					Gui, auf: Add, Text, % "x" xIt+20 " y" yIt , %A_LoopReadLine%
			}

		yIt+= nexty
	}

	Gui, auf: Add, Picture, % "x" (LW/2-158) " y" (LH-200)  "w148 h158 BackGroundTrans vaufLogo1 gaufClick"			, %logo1%
	Gui, auf: Add, Picture, % "x" (LW/2-158) " y" (LH-200)  "w148 h148 BackGroundTrans HIDE vaufLogo2 gaufClick", %logo2%
	Gui, auf: Add, Picture, % "x" (LW/2-158) " y" (LH-200)  "w148 h148 BackGroundTrans HIDE vaufLogo3 gaufClick", %logo3%

		Gui, auf: Font, s136 q5 c00FF00 cWhite, Futura Bk Md
	Gui, auf: Add, Text, % "x" (LW/2+10) " y" (LH-230) "BackGroundTrans vaufBuchstabeK gaufClick" , K

		Gui, auf: Font, s1 q5 c00FF00 cWhite, Futura Bk Bt
	Gui, auf: Add, GroupBox, % "x350 y30 w" (LW-700) " h300"

		Gui, auf: Font, s14 q4 cWhite, Futura Bk Bt
	Gui, auf: Add, GroupBox, % "x30 y340 w" (LW-60) " h" (LH-360), Und hier ist Deine Liste

		Gui, auf: Font, s14 q4 cWhite, Futura Bk Bt
	Gui, auf: Add, GroupBox, % "x" (LW/2-168) " y" (LH-210) " w" (128*2+60) " h168"

	Gui, auf: Show, % "x" LX " y" LY " w" LW " h" LH, Auffuellfenster

		OnMessage(0x200, "WM_MMOVE_Auffuellen")
	WinWaitClose, ahk_id %Lhauf%
		OnMessage(0x200, "")

return
;}

aufClick: ;{
	CName:= A_GuiControl
	Gui, Submit, nohide

	If (Instr(CName,"aufLogo")) OR (CName = "aufBuchstabeK") {
			OnMessage(0x200, "")
			guicontrol, Hide, aufLogo1
			guicontrol, Show, aufLogo2
			guicontrol, Hide, aufLogo3
			sleep, 1000
			checkedB:=0, ch:=0

			Loop, %Boxes%
			{
				ControlGet, checkstatus, checked,, Button%A_Index%, ahk_id %Lhauf%
				If (checkstatus=1) {
						checkedB ++
						If (A_Index= 1 or A_Index=Boxes) {
							ch++
								If (ch=2) { ;stopcode found
										Gui auf: Destroy
										return
								}
						}

					}
			}

			If (checkedB = Boxes) or (checkedB = 4) {	;ich habe keine Lust beim Testen alle Checkboxes abzuhaken

						Gui, auf: Destroy
						return

					} else {

						WinSet, AlwaysOnTop, Off, ahk_id %Lhauf%
						MsgBox,1, Nö!, Du sollst alles auffüllen!`nUnd deshalb musst Du auch alles abhaken!
						WinSet, AlwaysOnTop, On, ahk_id %Lhauf%
						OnMessage(0x200, "WM_MMOVE_Auffuellen")

					}


	}
return
;}

}

KofferAuffuellen() {  																										;--die Schwestern denken selten an meinen Hausbesuchskoffer

;{diese GUI ist für eine Auflösung von 1920x1080 am Besten geeignet
	Global aufLogo1,aufLogo2,aufLogo3, aufBuchstabeK, Schwesterchen2, Doc, hSis, KKhauf

	Logo1 = %A_ScriptDir%\assets\LogobuttonNormal.png
	Logo2 = %A_ScriptDir%\assets\LogobuttonClicked.png
	Logo3 = %A_ScriptDir%\assets\LogobuttonHovered.png

	ich = %AddendumDir%\assets\ich.png
	schwester = %AddendumDir%\assets\Schwester2 und der Koffer 300x.png

	SysGet, Mon1, Monitor, 1
	LX:= Mon1Left, LY:=Mon1Top, LW:=Mon1Right, LH:=Mon1Bottom

	Gui, auf: New, +HWNDKhauf Ex0x50400007 -Caption AlwaysOnTop
	Gui, auf: Color, c172842
	Gui, auf: Margin, 0, 0
	Gui, auf: Add, Picture, x150 y100 HWNDhSis2 vSchwesterchen2 gKofferClick, %schwester%
	;Gui, auf: Add, Picture, % "x" (W-340) " y30 vDoc gaufClick", %ich%
	Gui, auf: Font, s72 q4 cWhite, Futura Bk Md
	Gui, auf: Add, Text, % "x450 y70 w" (LW-500) " Center", Liebe Schwester2,
	Gui, auf: Font, s30 q4 cWhite, Futura Bk Md
	Gui, auf: Add, Text, % "x450 y200 w" (LW-500) " Center", heute ist Mittwoch! Und ich möchte Dich`nan den Hausbesuchskoffer errinnern!

	Gui, auf: Font, s14 q4 cWhite, Futura Bk Bt
		static xIt:= 500, nextx:= 500, yIt:=390, nexty:=24, boxes:=0
	Loop, Read, %A_ScriptDir%\DieListe.txt
	{
		If (A_LoopReadLine = "--") {
					xIt += nextx
					yIt := 390
					continue
			}

		If (A_LoopReadLine <> "") {
					boxes ++
					Gui, auf: Add, checkbox, % "x" xIt " y" yIt+5 " w12 h12"
					Gui, auf: Add, Text, % "x" xIt+20 " y" yIt , %A_LoopReadLine%
			}

		yIt+= nexty
	}

	Gui, auf: Add, Picture, % "x" (LW/8-158) " y" (LH-200)  "w148 h158 BackGroundTrans vaufLogo1 gKofferClick"			, %logo1%
	Gui, auf: Add, Picture, % "x" (LW/8-158) " y" (LH-200)  "w148 h148 BackGroundTrans HIDE vaufLogo2 gKofferClick", %logo2%
	Gui, auf: Add, Picture, % "x" (LW/8-158) " y" (LH-200)  "w148 h148 BackGroundTrans HIDE vaufLogo3 gKofferClick", %logo3%

		Gui, auf: Font, s136 q5 c00FF00 cWhite, Futura Bk Md
	Gui, auf: Add, Text, % "x" (LW/8+10) " y" (LH-230) "BackGroundTrans vaufBuchstabeK gaufClick" , K

		Gui, auf: Font, s1 q5 c00FF00 cWhite, Futura Bk Bt
	Gui, auf: Add, GroupBox, % "x450 y30 w" (LW-500) " h300"

		Gui, auf: Font, s14 q4 cWhite, Futura Bk Bt
	Gui, auf: Add, GroupBox, % "x450 y350 w" (LW-500) " h" (LH-360), Und hier ist Deine Liste

	;	Gui, auf: Font, s14 q4 cWhite, Futura Bk Bt
	;Gui, auf: Add, GroupBox, % "x" (W/5-168) " y" H-210 " w" (128*2+60) " h168"

	Gui, auf: Show, % "x" LX " y" LY " w" LW " h" LH, Auffuellfenster

	oldx:= 0, oldy:= 0
	SetTimer, BewegeSchwester2, 200

		OnMessage(0x200, "WM_MMOVE_Auffuellen")
	WinWaitClose, ahk_id %Khauf%
		OnMessage(0x200, "")

	SetTimer, BewegeSchwester2, off

return
;}

KofferClick: ;{
	CName:= A_GuiControl
	Gui, Submit, nohide

	If (Instr(CName,"aufLogo")) OR (CName = "aufBuchstabeK") {
			OnMessage(0x200, "")
			guicontrol, Hide, aufLogo1
			guicontrol, Show, aufLogo2
			guicontrol, Hide, aufLogo3
			sleep, 1000
			checkedB:=0, ch:=0
			Gui, auf: Destroy
			return
		}


return
;}

BewegeSchwester2: ;{

	SetTimer, BewegeSchwester2, off
	Critical
	Random, destx, 0, 30
	Random, desty, 0, 40
	;ToolTip %destx% : %desty%
	RandomBezier( oldx, oldy, destx, destx, hSis2, Khauf, "T2000")
	oldx:=destx, oldy:=desty
	Critical, Off
	SetTimer, BewegeSchwester2, On

return
;}

}

;{ sub of KofferAuffuellen
WM_MMOVE_Auffuellen() {

	LCName:= A_GuiControl

	If (Instr(LCName ,"aufLogo")) OR  (LCName="aufBuchstabeK") {

			guicontrol, Hide, aufLogo1
			guicontrol, Hide, aufLogo2
			guicontrol, Show, aufLogo3

	} else if (LCName = "Schwesterchen1") {

			ToolTip, Schwester`nSchwester1

	} else if (LCName = "Schwesterchen2") {

			ToolTip, Schwester`nSchwester2

	} else if (LCName = "Doc") {

			ToolTip,     Doktor`nSchwesterntreiber

	} else if (LCName = "") {

			guicontrol, Show, aufLogo1
			guicontrol, Hide, aufLogo2
			guicontrol, Hide, aufLogo3
			ToolTip

	}

}
RandomBezier( X0, Y0, Xf, Yf, hControlID, hWinID, O="" ) {
	 SetControlDelay, -1
	 SetBatchlines, -1
			/* 			RandomBezier.ahk
			Copyright (C) 2012,2013 Antonio França

			This script is free software: you can redistribute it and/or modify
			it under the terms of the GNU Affero General Public License as
			published by the Free Software Foundation, either version 3 of the
			License, or (at your option) any later version.

			This script is distributed in the hope that it will be useful,
			but WITHOUT ANY WARRANTY; without even the implied warranty of
			MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
			GNU Affero General Public License for more details.

			You should have received a copy of the GNU Affero General Public License
			along with this script.  If not, see <http://www.gnu.org/licenses/>.


		;========================================================================
		;
		; Function:     RandomBezier
		; Description:  Moves the mouse through a random Bézier path
		; URL (+info):  --------------------
		;
		; Last Update:  30/May/2013 03:00h BRT
		;
		; Created by MasterFocus
		; - https://github.com/MasterFocus
		; - http://masterfocus.ahk4.net
		; - http://autohotkey.com/community/viewtopic.php?f=2&t=88198
		;
		;========================================================================
		*/

    Time := RegExMatch(O,"i)T(\d+)",LM)&&(LM1>0)? LM1: 200
    RO := InStr(O,"RO",0) , RD := InStr(O,"RD",0)
    LN:=!RegExMatch(O,"i)P(\d+)(-(\d+))?",LM)||(LM1<2)? 2: (LM1>19)? 19: LM1
    If ((LM:=(LM3!="")? ((LM3<2)? 2: ((LM3>19)? 19: LM3)): ((LM1=="")? 5: ""))!="")
        Random, LN, 1, 6     ;%N%, %M%
    OfT:=RegExMatch(O,"i)OT(-?\d+)",LM)? LM1: 100, OfB:=RegExMatch(O,"i)OB(-?\d+)",LM)? LM1: 100
    OfL:=RegExMatch(O,"i)OL(-?\d+)",LM)? LM1: 100, OfR:=RegExMatch(O,"i)OR(-?\d+)",LM)? LM1: 100
    Random, LXM, 0, 5
	Random, LYM, 0, 5
    If ( RO )
        X0 += LXM, Y0 += LYM
    If ( RD )
        Xf += LXM, Yf += LYM
    If ( X0 < Xf )
        sX := X0-OfL, bX := Xf+OfR
    Else
        sX := Xf-OfL, bX := X0+OfR
    If ( Y0 < Yf )
        sY := Y0-OfT, bY := Yf+OfB
    Else
        sY := Yf-OfT, bY := Y0+OfB
    Loop, % (--LN)-1 {
        Random, X%A_Index%, %sX%, % bX-2
        Random, Y%A_Index%, %sY%, % bY-2
    }
    X%LN% := Xf, Y%LN% := Yf, LE := ( LI := A_TickCount ) + Time
    While ( A_TickCount < LE ) {
        LU := 1 - (LT := (A_TickCount-LI)/Time)
        Loop, % LN + 1 + (LX := LY := 0) {
            Loop, % Idx := A_Index - (F1 := F2 := F3 := 1)
                F2 *= A_Index, F1 *= A_Index
            Loop, % LD := LN-Idx
                F3 *= A_Index, F1 *= A_Index+Idx
            LM:=(F1/(F2*F3))*((LT+0.000001)**Idx)*((LU-0.000001)**LD), LX+=LM*X%Idx%, LY+=LM*Y%Idx%
        }
        ControlMove, Static1, % LX+70, % LY+70,,, ahk_id %hWinID%
        Sleep, 1
    }
    ControlMove, Static1, % X%LN% + 70, % Y%LN%+70 ,,,, ahk_id %hWinID%					;X%N%, Y%N%
    Return LN+1
}
CatMull_ControlMove( px0, py0, px1, py1, px2, py2, px3, py3, Segments=8, Rel=0, Speed=2 ) {
; Function by [evandevon]. Moves the mouse through 4 points (without control point "gaps"). Inspired by VXe's
;cubic bezier curve function (with some borrowed code).
   MouseGetPos, px0, py0
   If Rel
      px1 += px0, px2 += px0, px3 += px0, py1 += py0, py2 += py0, py3 += py0
   Loop % Segments - 1
   {
	;CatMull Rom Spline - Working
	  Lu := 1 - Lt := A_Index / Segments
	  cmx := Round(0.5*((2*px1) + (-px0+px2)*Lt + (2*px0 - 5*px1 +4*px2 - px3)*Lt**2 + (-px0 + 3*px1 - 3*px2 + px3)*Lt**3) )
	  cmy := Round(0.5*((2*py1) + (-py0+py2)*Lt + (2*py0 - 5*py1 +4*py2 - py3)*Lt**2 + (-py0 + 3*py1 - 3*py2 + py3)*Lt**3) )

	  MouseMove, cmx, cmy, Speed,

   }
   MouseMove, px3, py3, Speed
} ; CatMull_MouseMove( px1, py1, px2, py2, px3, py3, Segments=5, Rel=0, Speed=2 ) -------------------
;}

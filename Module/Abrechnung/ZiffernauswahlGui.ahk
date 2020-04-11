#NoEnv
#SingleInstance
#Persistent

;#include libs\DBase.ahk	

DetectHiddenWindows, On
;---Liste der EventIDs

EventID1:= "EVENT_SYSTEM_SOUND"
EventID2:= "EVENT_SYSTEM_ALERT"
EventID3:= "EVENT_SYSTEM_FOREGROUND"
EventID4:= "EVENT_SYSTEM_MENUSTART"
EventID5:= "EVENT_SYSTEM_MENUEND"
EventID6:= "EVENT_SYSTEM_MENUPOPUPSTART"
EventID7:= "EVENT_SYSTEM_MENUPOPUPEND"
EventID8:= "EVENT_SYSTEM_CAPTURESTART"
EventID9:= "EVENT_SYSTEM_CAPTUREEND"
EventID10:= "EVENT_SYSTEM_MOVESIZESTART"
EventID11:= "EVENT_SYSTEM_MOVESIZEEND"
EventID12:= "EVENT_SYSTEM_CONTEXTHELPSTART"
EventID13:= "EVENT_SYSTEM_CONTEXTHELPEND"
EventID14:= "EVENT_SYSTEM_DRAGDROPSTART"
EventID15:= "EVENT_SYSTEM_DRAGDROPEND"
EventID16:= "EVENT_SYSTEM_DIALOGSTART"
EventID17:= "EVENT_SYSTEM_DIALOGEND"

EventID18:= "EVENT_SYSTEM_SCROLLINGSTART"
EventID19:= "EVENT_SYSTEM_SCROLLINGEND"
EventID20:= "EVENT_SYSTEM_SWITCHSTART"
EventID21:= "EVENT_SYSTEM_SWITCHEND"
EventID22:= "EVENT_SYSTEM_MINIMIZESTART"
EventID23:= "EVENT_SYSTEM_MINIMIZEEND"

EventID32768:= "EVENT_OBJECT_CREATE"
EventID32769:= "EVENT_OBJECT_DESTROY"
EventID0x8002:= "EVENT_OBJECT_SHOW"
EventID0x8003:= "EVENT_OBJECT_HIDE"
EventID0x8004:= "EVENT_OBJECT_REORDER"
; EventID0x??????:= "EVENT_OBJECT_INVOKED"                 ; laut MSDN vorhanden
EventID0x8005:= "EVENT_OBJECT_FOCUS"
EventID0x8006:= "EVENT_OBJECT_SELECTION"      
EventID0x8007:= "EVENT_OBJECT_SELECTIONADD"
EventID0x8008:= "EVENT_OBJECT_SELECTIONREMOVE"
EventID0x8009:= "EVENT_OBJECT_SELECTIONWITHIN"
; EventID0x??????:= "EVENT_OBJECT_TEXTSELECTIONCHANGED"    ; laut MSDN vorhanden
; EventID0x??????:= "EVENT_OBJECT_CONTENTSCROLLED"       ; laut MSDN vorhanden
EventID0x800a:= "EVENT_OBJECT_STATECHANGE"
EventID0x800b:= "EVENT_OBJECT_LOCATIONCHANGE"
EventID0x800c:= "EVENT_OBJECT_NAMECHANGE"
EventID0x800d:= "EVENT_OBJECT_DESCRIPTIONCHANGE"
EventID0x800e:= "EVENT_OBJECT_VALUECHANGE"
EventID0x800f:= "EVENT_OBJECT_PARENTCHANGE" ;----
EventID0x8010:= "EVENT_OBJECT_HELPCHANGE"
EventID0x8011:= "EVENT_OBJECT_DEFACTIONCHANGE"
EventID0x8012:= "EVENT_OBJECT_ACCELERATORCHANGE" 

   zn:=0
   zz:=0
   AnzKopf :=
   AnzZeilen:=

CoordMode, Mouse, Screen
   ZWinEventProc:=0
   ZSetWinEventProc:=0
   lpfnWinEventProc := RegisterCallback("WinEventProc", "",7  )
   
   Process, wait, AlbisCS.exe, 3
   ProcessID = %ErrorLevel%
   hWinEventHook    := SetWinEventHook(32000, 32770,0,lpfnWinEventProc,ProcessID,0,0)   

OnExit, UnregisterHook 

/* #####################################################################################################
   ##############################   						Funktionen												   ################################### 
   #####################################################################################################
*/

LV_ColorInitiate(Gui_Number=1, Control="") ; initiate listview color change procedure 
{ 
  global hw_LV_ColorChange 
  If Control =
    Control =SysListView321
  Gui, %Gui_Number%:+Lastfound 
  Gui_ID := WinExist() 
  ControlGet, hw_LV_ColorChange, HWND,, %Control%, ahk_id %Gui_ID% 
  OnMessage( 0x4E, "WM_NOTIFY" ) 
} 

LV_ColorChange(Index="", TextColor="", BackColor="") ; change specific line's color or reset all lines
{ 
  global 
  If Index = 
    Loop, % LV_GetCount() 
      LV_ColorChange(A_Index) 
  Else
    { 
    Line_Color_%Index%_Text := TextColor 
    Line_Color_%Index%_Back := BackColor 
    WinSet, Redraw,, ahk_id %hw_LV_ColorChange% 
    } 
}

WinEventProc(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime ) {
global zn, zz, AnzZeilen
   zn:=zn +1
   zz:=zz +1
   IF (zz = 51)
   {
   AnzKopf:=Überschrift
   Loop, parse, AnzZeilen, `n
      {
      IF (A_Index < 46)
         Continue
      AnzKopf:=AnzKopf . A_LoopField . "`n"
      }
      AnzZeilen:=AnzKopf
      AnzKopf:=
      zz:=1
   }
   dez:=event
   EventName:=EventID%event%
   FormatTime, OutTime ,T12 , Time
   SetFormat,integer,hex
   hwnd += 0
   idobject += 0
   idchild += 0
   SetFormat,integer,d
   WinGetTitle, Titel, ahk_class #32770 
	
	;If (event = 32768) {
   ;AnzZeilen := AnzZeilen . "#" . zn . "`thwnd: " . hwnd . "    `tidObject: " . idobject . " tidchild: " . idchild . "`tThread: " . dwEventThread . "`t`tTime: " . OutTime . "`tEvent: " . EventName . " [" . event . "]`n"
   ;AnzZeilen := AnzZeilen . "#" . zn . "`thwnd: " . hwnd . " Fenstername: " . Titel . " Klasse: " . Klasse . " Steuerelement: " . Steuerelement . " `tEvent: " . EventName . " [" . event . "]`n"
   ;tooltip,  %AnzZeilen% ,0,0
								}
		
	

SetWinEventHook(eventMin, eventMax, hmodWinEventProc, lpfnWinEventProc, idProcess, idThread, dwFlags) { ; hier kommt er einmal rein????? warum    
;  DllCall("CoInitialize", "uint", 0)  ; notwendig??? keine Änderung wenn nicht aktiv!!!
   return DllCall("SetWinEventHook", "uint", eventMin, "uint", eventMax
                , "uint", hmodWinEventProc, "uint", lpfnWinEventProc
                , "uint", idProcess, "uint", idThread, "uint", dwFlags)   
}


/* ##################################       Bereichsende -Funktionen-    ###############################################
	####################################################################################################
*/

gosub, ZiffernGui
sleeptime:= 100

SetTimer, AlbisControl, 100
LControl & ä::GoSub UnregisterHook

return

AlbisControl:

 If InStr(Titel, "Muster 16") {
	WinClose, Titel
	WinActivate, ALBIS -
	Send, {ALTDOWN}o{ALTUP}{ALTDOWN}{ALTUP}{ALTDOWN}of{ALTUP}
	WinWait, Formulare, 
	IfWinNotActive, Formulare, , WinActivate, Formulare, 
	WinWaitActive, Formulare, 
	Send, rezept
	}

return

Chroniker:


;Lese Programmdatum aus

	WinWait, ALBIS -
	IfWinNotActive, ALBIS -
				WinActivate, ALBIS -
	WinWaitActive, ALBIS -
	
	Send, {INS}{SHIFTDOWN}{TAB}{SHIFTUP}           						;Nutzer wählt vorher eine Zeile aus in der das GVU Datum steht
	Sleep, %sleeptime%	
	ControlGetText, Datum, Edit2, ALBIS -
	StringMid, Monat, Datum, 4, 2
	StringRight, Jahr, Datum, 2
	
	Send, {TAB}																				;ein Feld weiter rücken
	Sleep, %sleeptime%	

return


ZiffernGui:

Gui, AC1:NEW
Gui, AC1:Font, S10 CDefault, Futura Bk Bt
Gui, AC1:+LastFound +AlwaysOnTop
Gui, AC1:Add, Treeview, r10 w330 h890 AltSubmit Checked gAAgTreeView BackgroundAAAAFF

P1 	 := TV_Add("Betreuungsstrukturvertrag OST", 0,  "bold Checked:0")
P1C1 := TV_Add("	93511", P1)
P1C2 := TV_Add("93512", P1)
P1C3 := TV_Add("	93513", P1)
P1C4 := TV_Add("	93513 (nur TKK)", P1)

P2      := TV_Add("Hausärztl.-geriatr. Betreuungskomplex", 0, "bold Option:Vis") 
P2C1  := TV_Add("03360", P2)
P2C2  := TV_Add("03362", P2)
P2C3  := TV_Add("30980", P2)
P2C4  := TV_Add("30981", P2)
P2C5  := TV_Add("30985", P2)
P2C6  := TV_Add("30986", P2)
P2C7  := TV_Add("30988", P2)

P3	  := TV_Add("Rheumastrukturvertrag", 0,"Bold")
P3C1  := TV_Add("93420", P3)
P3C2  := TV_Add("93421", P3)
P3C3  := TV_Add("93422", P3)

P4       := TV_Add("Schriftliche Mitteilungen, Gutachten", 0,"Bold")
P4C1   :=TV_Add("81000", P4)
P4C2   :=TV_Add("81001", P4)
P4C3   :=TV_Add("81002", P4)

P5       := TV_Add("Hautscreening BARMER,HEK,TKK (AOK Nr.3)", 0, "Bold")
P5C1   := TV_Add("94100", P5)
P5C2   := TV_Add("94101", P5)
P5C3	:= TV_Add("01745H", P5)

P6       := TV_Add("Überweisungssteuerung BARMER,TK", 0, "Bold")
P6C1   := TV_Add("93480A", P6)
P6C2   := TV_Add("93480B", P6)

P7       := TV_Add("Überweisungssteuerung Knappschaft", 0, "Bold")
P7C1   := TV_Add("81110B", P7)
P7C2   := TV_Add("81112", P7)

P8       := TV_Add("hausarztzentrierten Versorgung AOK", 0, "Bold")
P8C1   :=TV_Add("95051", P8)
P8C2   :=TV_Add("95052", P8)
P8C3   :=TV_Add("95053", P8)
P8C4   :=TV_Add("95056", P8)


Gui, AC1:Show,x1500 y50 w350 h900, AlbisAutomat Module Ziffern
GuiControl, +Redraw, AAgTreeView

return

AAgTreeView:
	{
	;Gui, AC1:TreeView, TV1

	;if A_GuiEvent <> T
	;	return
	
	TV_GetText(SelectedItemText, A_EventInfo)
	ParentID := A_EventInfo
	MouseGetPos, mx, my
	ToolTip, %SelectedItemText%, %mx%, %my%
}
return


UnregisterHook:
  DllCall("UnhookWinEvent", Uint, hWinEventHook)
  ExitApp 
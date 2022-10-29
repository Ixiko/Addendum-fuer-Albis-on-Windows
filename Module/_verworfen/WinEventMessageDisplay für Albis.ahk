;----------------------------------------------------------------------------------------------------------------------------------------------
; WinEventHook Messages v0.3 by Serenity
; modified version v0.3i by Ixiko 14.09.2019
; http://www.autohotkey.com/forum/viewtopic.php?t=35659
;----------------------------------------------------------------------------------------------------------------------------------------------


;----------------------------------------------------------------------------------------------------------------------------------------------
; 	SCRIPT SETTINGS
;----------------------------------------------------------------------------------------------------------------------------------------------
    #SingleInstance Force
    #Persistent
    SetBatchLines,-1
    ; Process, Priority,, High

;----------------------------------------------------------------------------------------------------------------------------------------------
; VARIABLES
;----------------------------------------------------------------------------------------------------------------------------------------------
    ExcludeScriptMessages	:= 1                                         	; 0 to include
            ExcludeGuiEvents	:= 1                                         	; 0 to include
                                Title	:= "WinEventHook Messages"
    					      Filters	:= ""
    						  Pause	:= 0
                 WM_VSCROLL	:= 0x115
    			   SB_BOTTOM	:= 7
    						 hFilter	:= 0
				       global log	:= ""
				  global EvMsg	:= Object()
							EvMsg	:= {"0x1": "EVENT_SYSTEM_SOUND"
					        				, "0x2" : "EVENT_SYSTEM_ALERT"
					        				, "0x3" : "EVENT_SYSTEM_FOREGROUND"
					        				, "0x4" : "EVENT_SYSTEM_MENUSTART"
					        				, "0x5" : "EVENT_SYSTEM_MENUEND"
					        				, "0x6" : "EVENT_SYSTEM_MENUPOPUPSTART"
					        				, "0x7" : "EVENT_SYSTEM_MENUPOPUPEND"
					        				, "0x8" : "EVENT_SYSTEM_CAPTURESTART"
					        				, "0x9" : "EVENT_SYSTEM_CAPTUREEND"
					        				, "0xA" : "EVENT_SYSTEM_MOVESIZESTART"
					        				, "0xB" : "EVENT_SYSTEM_MOVESIZEEND"
					        				, "0xC" : "EVENT_SYSTEM_CONTEXTHELPSTART"
					        				, "0xD" : "EVENT_SYSTEM_CONTEXTHELPEND"
					        				, "0xE" : "EVENT_SYSTEM_DRAGDROPSTART"
					        				, "0xF" : "EVENT_SYSTEM_DRAGDROPEND"
					        				, "0x10" : "EVENT_SYSTEM_DIALOGSTART"
					        				, "0x11" : "EVENT_SYSTEM_DIALOGEND"
					        				, "0x12" : "EVENT_SYSTEM_SCROLLINGSTART"
					        				, "0x13" : "EVENT_SYSTEM_SCROLLINGEND"
					        				, "0x14" : "EVENT_SYSTEM_SWITCHSTART"
					        				, "0x15" : "EVENT_SYSTEM_SWITCHEND"
					        				, "0x16" : "EVENT_SYSTEM_MINIMIZESTART"
					        				, "0x17" : "EVENT_SYSTEM_MINIMIZEEND"
					        				, "0x20" : "EVENT_SYSTEM_DESKTOPSWITCH"
					        				, "0x8000" : "EVENT_OBJECT_CREATE"
					        				, "0x8001" : "EVENT_OBJECT_DESTROY"
					        				, "0x8002" : "EVENT_OBJECT_SHOW"
					        				, "0x8003" : "EVENT_OBJECT_HIDE"
					        				, "0x8004" : "EVENT_OBJECT_REORDER"
					        				, "0x8005" : "EVENT_OBJECT_FOCUS"
					        				, "0x8006" : "EVENT_OBJECT_SELECTION"
					        				, "0x8007" : "EVENT_OBJECT_SELECTIONADD"
					        				, "0x8008" : "EVENT_OBJECT_SELECTIONREMOVE"
					        				, "0x8009" : "EVENT_OBJECT_SELECTIONWITHIN"
					        				, "0x800A" : "EVENT_OBJECT_STATECHANGE"
					        				, "0x800B" : "EVENT_OBJECT_LOCATIONCHANGE"
					        				, "0x800C" : "EVENT_OBJECT_NAMECHANGE"
					        				, "0x800D" : "EVENT_OBJECT_DESCRIPTIONCHANGE"
					        				, "0x800E" : "EVENT_OBJECT_VALUECHANGE"
					        				, "0x800F" : "EVENT_OBJECT_PARENTCHANGE"
					        				, "0x8010" : "EVENT_OBJECT_HELPCHANGE"
					        				, "0x8011" : "EVENT_OBJECT_DEFACTIONCHANGE"
					        				, "0x8012" : "EVENT_OBJECT_ACCELERATORCHANGE"
					        				, "0x8013" : "EVENT_OBJECT_INVOKED"
					        				, "0x8014" : "EVENT_OBJECT_TEXTSELECTIONCHANGED"
					        				, "0x8015" : "EVENT_OBJECT_CONTENTSCROLLED"
					        				, "0x7FFFFFFF" : "EVENT_MAX"}

;----------------------------------------------------------------------------------------------------------------------------------------------
; BUILD UP GUI
;----------------------------------------------------------------------------------------------------------------------------------------------
    FilterMenu()
    Gui()
    ahk := WinExist()

;----------------------------------------------------------------------------------------------------------------------------------------------
; ASK TO FILTER FOR ALBIS
;----------------------------------------------------------------------------------------------------------------------------------------------
    ;MsgBox, 4, Question, Do you want to filter only messages for`nAlbis on Windows ?
    ;IfMsgBox, Yes
	; {
		WinGet, hFilter, PID, ahk_class OptoAppClass
		hApp := WinExist("ahk_class OptoAppClass")
	; }
	WinSet, AlwaysOnTop, On, % "ahk_id " hEventsGui

;----------------------------------------------------------------------------------------------------------------------------------------------
; INITIALIZE HOOKS
;----------------------------------------------------------------------------------------------------------------------------------------------
	HookProcAdr := RegisterCallback( "HookProc", "F" )
	dwFlags := ( ExcludeScriptMessages = 1 ? 0x1 : 0x0 )
	hWinEventHook1 := SetWinEventHook( 0x00008005, 0x00008005, 0, HookProcAdr, hFilter, 0, dwFlags )
	;hWinEventHook1 := SetWinEventHook( 0x00000001, 0x00008050, 0, HookProcAdr, hFilter, 0, dwFlags )
	;hWinEventHook2 := SetWinEventHook( 0x3, 0x20, 0, HookProcAdr, hFilter, 0, dwFlags )
	hooks = 2

	Return
;----------------------------------------------------------------------------------------------------------------------------------------------
; HOTKEYS
;----------------------------------------------------------------------------------------------------------------------------------------------
	^!-::
	Reload
	return

;----------------------------------------------------------------------------------------------------------------------------------------------
; FUNCTIONS
;----------------------------------------------------------------------------------------------------------------------------------------------
	FilterMenu() {

		   Global FilterList
		   Menu, Filter, Add, Filter &All, FilterAll
		   Menu, Filter, Add, Filter &None, FilterNone
		   Menu, Filter, Add
		   FilterList = SOUND,ALERT,FOREGROUND,MENUSTART,MENUEND,MENUPOPUPSTART,MENUPOPUPEND
		   ,CAPTURESTART,CAPTUREEND,MOVESIZESTART,MOVESIZEEND,CONTEXTHELPSTART
		   ,CONTEXTHELPEND,DRAGDROPSTART,DRAGDROPEND,DIALOGSTART,DIALOGEND,SCROLLINGSTART
		   ,SCROLLINGEND,SWITCHSTART,SWITCHEND,MINIMIZESTART,MINIMIZEEND

			Menu, Extras, Add, Save listview content, SaveContent

		   Loop, Parse, FilterList, `,
		   {
			  If A_Loopfield
				 Menu, Filter, Add, %A_Loopfield%, SetFilter
		   }
		   Menu, FilterMenu, Add, Message &Filter, :Filter
		   Gui, Menu, FilterMenu

	}

	HookProc( hWinEventHook, Event, hWnd, idObject, idChild, dwEventThread, dwmsEventTime ) {

		   Global Pause,WM_VSCROLL,SB_BOTTOM,ahk,Filters
			static hWndo, Evento, idObjecto, idChildo, WinTitleo,ClassNNo

		   SetFormat, Integer, Hex
		   ; Event += 0
		   ; Sleep, 50 ; give a little time for WinGetTitle/WinGetActiveTitle functions, otherwise they return blank
		  ; hWnd += 0, idObject += 0, idChild += 0

		   If Event not in %Filters%
		   {
			  If ( Pause = 0 )
			  {
					classNN:= Control_GetClassNN(GetAncestor(hWnd, 2), hWnd)
					WinTitle:= WinGetTitle(hWnd)
					;If (hWndo = hwnd) && (Evento = Event) && (idObjecto = idObject) && (idChildo = idChild) && (WinTitleo = WinTitle) && (ClassNNo = ClassNN)
					;	return
					LV_Add( "", hWnd, idObject, idChild, WinTitle, ClassNN, Event, EvMsg[(Event "")] )
     				FormatTime, logtime, % A_Now, hh:mm:ss:ms
					log .= logtime "`t|`t" hWnd "`t|`t" idObject "`t|`t" idChild "`t|`t" WinTitle "`t|`t" ClassNN "`t|`t" Event "`t|`t" EvMsg[(Event "")] "`n"
					;hWndo:= hwnd, Evento:= Event, idObjecto:= idObject, idChildo:= idChild, WinTitleo:= WinTitle, ClassNNo:= ClassNN
			  }
			  SendMessage, WM_VSCROLL, SB_BOTTOM, 0, SysListView321, ahk_id %ahk%
		   }
	}

	Gui() {

	   Global
	   Gui, +LastFound +Resize +hWndhEventsGui ; +ToolWindow
	   Gui, Margin, 0, 0
	   Gui, Font, s8, Futura BK Bt
	   Gui, Color,, DEDEDE
	   Gui, Add, ListView, w600 r10 vData +Grid +NoSort +HWNDhLV, Hwnd|idObject|idChild|Title|ClassNN|Event|Message
	   LV_ModifyCol( 1, 60 ), LV_ModifyCol( 2, 40), LV_ModifyCol( 3, 40)
	   LV_ModifyCol( 4, 100 ), LV_ModifyCol( 5, 100 ), LV_ModifyCol( 7, 190 )
	   Gui, Show,, %Title%

	}

	WinGetTitle( hwnd ) {
		; WinGetTitle, wtitle, ahk_id %hwnd%
		VarSetCapacity(sClass,80,0)
		DllCall("GetWindowTextW", "UInt", hWnd, "Str", sClass, "Int", VarSetCapacity(sClass)+1)
		wtitle := sClass
		sClass =
		Return wtitle
	}

	WinGetClass( hwnd ) {
		; WinGetClass, wclass, ahk_id %hwnd%
		VarSetCapacity(sClass,80,0)
		DllCall("GetClassNameW", "UInt", hWnd, "Str", sClass, "Int", VarSetCapacity(sClass)+1)
		wclass := sClass
		sClass =
		Return wclass
	}

	SetWinEventHook(eventMin, eventMax, hmodWinEventProc, lpfnWinEventProc, idProcess, idThread, dwFlags) {
	   DllCall("CoInitialize", Uint, 0)
	   return DllCall("SetWinEventHook"
	   , Uint,eventMin
	   , Uint,eventMax
	   , Uint,hmodWinEventProc
	   , Uint,lpfnWinEventProc
	   , Uint,idProcess
	   , Uint,idThread
	   , Uint,dwFlags)
	}

	NewHook() {

	   Global

	   If HookSelected = 1
	   {
		  LV_GetText( SelHwnd, LV_GetNext(0, "Focused") )
		  WinGet, idProcess, PID, ahk_id %SelHwnd%
		  If idProcess =
		  {
			 Menu, Ctx, ToggleCheck, &Receive events from this process only
			 HookSelected := !HookSelected
			 Return
		  }
	   }
	   Else
	   {
		  idProcess = 0 ; hook all
	   }

	   UnhookWinEvent()
	   HookProcAdr := RegisterCallback( "HookProc", "F" ) ; new hook
	   hWinEventHook1 := SetWinEventHook( 0x3, 0x3, 0, HookProcAdr, idProcess, 0, dwFlags )
	   hWinEventHook2 := SetWinEventHook( 0x00008001, 0x00008001, 0, HookProcAdr, idProcess, 0, dwFlags )
	}

	UnhookWinEvent() {
	   Global
	   Loop %hooks%
	   {
		   DllCall( "UnhookWinEvent", Uint,hWinEventHook%A_Index% )
	   }
	   DllCall( "GlobalFree", UInt,&HookProcAdr ) ; free up allocated memory for RegisterCallback
	}

	LV_AutoColumSizer(hLV, Sizes, Options:="") {                                         	;-- computes and changes the pixel width of the columns across the full width of a listview

	; PARAMETERS:
	; ----------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Sizes   	- 	this example is for a 4 column listview, for a better understanding it is possible to use a different syntax
	;               	Sizes:= "15%, 18%, 60%" or "15, 18, 60" or "15,18,60" or "15|18|60" or "15% 18% 60%"
	;               	It does not matter which characters or strings you use for subdivision, the little RegEx algorithm recognizes the dividers
	;               	REMARK: !avoid specifying the last column width, this size will be computed!
	;                 	    *		*		*		*		*		*		*		*		*		*		*		*		*		*		*
	; ** todo **	there is also an automatic mode which calculates the column width of the listview over the maximum pixel width of the content of the columns
	;                	you have to use Sizes:= "AutoColumnWidth"
	;
	; ** todo ** Options 	-	can be passed to limit the maximum column width to the maximum pixel width of the column contents
	;                  	or to prevent undersizing of columns

	static hHeader, LVP, hLVO, SizesO
	w:= LVP:= []

	If hLVO <> hLV
			hHeader:= LV_EX_GetHeader(hLV), hLVO:= hLV

	If SizesO <> Sizes
	{
			pos := 1
			If !Instr(Sizes, "AutoColumnWidth")
					While pos:= RegExMatch(Sizes, "\d+", num, StrLen(num)+pos)
							LVP[A_Index] := num
			else
				nin:=1

			LVP_Last := 100

			Loop, % LVP.MaxIndex()
			{
					LVP[A_Index]	:= 	"0" . LVP[A_Index]
					LVP[A_Index]	+=	0
					LVP_Last      	-=	LVP[A_Index]
					LVP[A_Index]	:= 	Round(LVP[A_Index]/100, 2)
			}
			LVP.Push(Round((LVP_Last-1)/100, 2))
			SizesO:= Sizes
	}

	ControlGetPos,,, LV_Width,,, % "ahk_id " hLV
	LV_Width -= DllCall("GetScrollPos", "UInt", hLV, "Int", 1)	;subtracts the width of the vertical scrollbar to get the client size of the listview

	Loop, % LVP.MaxIndex()
		DllCall("SendMessage", "uint", hLV, "uint", 4126, "uint", A_Index-1, "int", Floor(LV_Width * LVP[A_Index])) 	;sets the column width
	}

	LV_EX_GetHeader(HLV) {                                                                         	;-- Retrieves the handle of the header control used by the list-view control.
   ; LVM_GETHEADER = 0x101F -> http://msdn.microsoft.com/en-us/library/bb774937(v=vs.85).aspx
   SendMessage, 0x101F, 0, 0, , % "ahk_id " . HLV
   Return ErrorLevel
	}

	LV_EX_GetColumnWidth(HLV, Column) {                                                	;-- gets the width of a column in report or list view.
	   ; LVM_GETCOLUMNWIDTH = 0x101D -> http://msdn.microsoft.com/en-us/library/bb774915(v=vs.85).aspx
	   SendMessage, 0x101D, % (Column - 1), 0, , % "ahk_id " . HLV
	   Return ErrorLevel
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

	GetAncestor(hWnd, Flag := 2) {	                                                                                                	;--
		;1 - Parent , 2 - Root
		Return DllCall("GetAncestor", "Ptr", hWnd, "UInt", Flag)
	}

;----------------------------------------------------------------------------------------------------------------------------------------------
; LABEL
;----------------------------------------------------------------------------------------------------------------------------------------------
	GuiContextMenu: ;{
		Menu, Ctx, Add, &Receive events from this process only, HookMode
		Menu, Ctx, Add, &Clear list, ClearList
		Menu, Ctx, Show
	Return ;}

	HookMode: ;{
		Menu, Ctx, ToggleCheck, &Receive events from this process only
		HookSelected := !HookSelected
		NewHook()
	Return ;}

	ClearList: ;{
		WinSet, AlwaysOnTop, Off, % "ahk_id " hEventsGui
		MsgBox, 4, Question, Should the EventMessage log be saved before emptying the listview?
		WinSet, AlwaysOnTop, On, % "ahk_id " hEventsGui
		LV_Delete()
		log := ""
	return ;}

	GuiSize: ;{
		GuiControl, Move, Data, w%A_GuiWidth% h%A_GuiHeight%
		LV_AutoColumSizer(hLV, "7% 7% 7% 30% 20% 10%")
		SendMessage, WM_VSCROLL, SB_BOTTOM, 0, SysListView321, ahk_id %ahk%
	Return ;}

	GuiClose: ;{
	GuiEscape:
	ExitApp
	Return ;}

	SaveContent: ;{
		Pause := 1
		FileSelectFile, filePath, 24, % A_ScriptDir, Choose path and filename for eventlog, *.csv
		If ErrorLevel = 1
			return
		file:= FileOpen(filepath, "w", "UTF8")
		file.WriteLine("time       `t|`t   Hwnd   `t|`t   idObject   `t|`t   idChild   `t|`t   Title   `t|`t   ClassNN   `t|`t   Event   `t|`t   Message   `t")
		Pause := 0
	return ;}

	SetFilter: ;{
	Menu, Filter, ToggleCheck, %A_ThisMenuItem%
	Loop, Parse, FilterList, `,
	{
	   If ( A_ThisMenuItem = A_Loopfield )
	   {
		  If A_Index in %Filters% ; remove from filter
		  {
			 Filter := A_Index
			 Loop, Parse, Filters, `,
			 {
				If ( A_Loopfield != Filter )
				   NewFilters .= A_Loopfield . ( A_Loopfield != "" ? "`," : "" )
			 }
			 Filters := NewFilters, NewFilters := ""
		  }
		  Else ; add to filter
		  {
			 Filters .= A_Index . ","
		  }
	   }
	}
	Return ;}

	FilterAll: ;{
	Filters =
	Loop, Parse, FilterList, `,
	{
	   Menu, Filter, Check, %A_Loopfield%
	   Filters .= A_Index . ( A_Index != 23 ? "," : "" )
	}
	Return ;}

	FilterNone: ;{
	Loop, Parse, FilterList, `,
	{
	   Menu, Filter, UnCheck, %A_Loopfield%
	   Filters =
	}
	Return ;}

	#IfWinActive WinEventHook Messages ;{
	   C::LV_Delete()

	   P::
	   Pause :=! Pause, WinTitle := ( Pause = 0 ? Title : Title . " (Paused)" )
	   WinSetTitle %WinTitle%
	   Return

	   R::Reload
	   X::ExitApp

	   ^C::
	   Clipboard =
	   Loop, % LV_GetCount("Col")
	   {
		  LV_GetText( lv%A_Index%, LV_GetNext(0, "Focused"), A_Index )
		  Clipboard .= lv%A_Index% . ( A_Index != LV_GetCount("Col") ? "|" : "" )
	   }
	   Return
	#IfWinActive
	;}

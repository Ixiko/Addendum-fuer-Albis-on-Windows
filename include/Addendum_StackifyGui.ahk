; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;     	Addendum_StackifyGui - V0.1 alpha
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;		Verwendung:	   	-	Multi Debug-Consolenklasse für Kontrollausgaben für Addendum
;
; 		Beschreibung:       -	kann als Library eingebunden werden
;                                	- 	aber auch als Thread mit AutohotkeyH
;									-	oder als eigenständiges Skript betrieben werden
;									-	Datenaustausch kann direkt über das Klassenobjekt oder über das von der Klasse erstellte
;										COM-Objekt erfolgen (CLSID: {6B39CAA1-A320-4CB0-8DB4-352AA81E460E})
;
;
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Addendum_StackifyGui started:          	09.12.2020
;       Addendum_StackifyGui last change:    	15.12.2021
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


   debugv := new StackifyGui("Stackify TestGui")

   for name, obj in StackifyGui.GetActiveObjects() {
		nametyp := name SubStr("                        ", -1*(50-StrLen(name))) "`t " ComObjType(obj, "Name")
         list .= nametyp "`n"
		 debugv.Notify(nametyp)
		 Sleep 1500
	}



return
ExitApp

class StackifyGui   {

	  __New(GuiName, Options="Font:s10 q5, Calibri ", Position="Auto", COMIntegration=true) {

			global

			this.stackifyIsMinimized:= false
			this.userWantsExitApp:= false
			this.width:= 850
			this.x:= 1400, This.y:= 100
			this.dbgV := Array()

			Gui, New, -Caption +ToolWindow -DPIScale AlwaysOnTop -MinimizeBox -Border HWNDhwnd Resize

			Gui, %hwnd%: Margin, 0, 0
			Gui, %hwnd%: Color, c5151B7, c172842

			Gui, %hwnd%: Font, s10 cWhite bold
			Gui, %hwnd%: Add, Text, % "x0 y2 w400 BackgroundTrans vStTitle"

			Gui, %hwnd%: Font, s7 cWhite Normal, Consolas
			Gui, %hwnd%: Add, Button, % "x" this.width-40 " y2 gStackifyMin vStackifyMin hwndhStackifyMin", _
			Gui, %hwnd%: Add, Button, % "x" this.width-20 " y2 gStackifyQuit vStackifyQuit hwndhStackifyQuit", X
			;~ fnQuit := Func("Destroy")
			;~ GuiControl, %hwnd%: +g, StackifyQuit,

			Gui, %hwnd%: Font, s10 cWhite
			Gui, %hwnd%: Add, Edit, % "x0 w" this.width " r30 -Border -VScroll vStackify hwndhStackifyEdit"

			Gui, %hwnd%: Show, % "x" this.x " y" this.y " Hide NA", Stackify_%hwnd%

			this.Hwnd	:= hwnd
			this.hEdit 	:= Format("0x{:X}", hStackifyEdit)

  		  ; Enable shadow
			VarSetCapacity(margins, 16, 0)
			NumPut(1, &margins, 0, "Int")
			DllCall("Dwmapi\DwmExtendFrameIntoClientArea", "UPtr", hwnd, "UPtr", &margins)

			GuiControl, %hwnd%: , StTitle, % GuiName "           | hwnd output: "  this.hEdit

			WinSet, ExStyle	, 0x00000000 	, % "ahk_id " this.hEdit
			WinSet, Style   	, +0x100      	, % "ahk_id " this.hEdit

			Gui, %hwnd%: Show, NoActivate

			WinSet, Transparent, off  	, % "ahk_id " this.Hwnd
			WinSet, Transparent, 180	, % "ahk_id " this.Hwnd

			OnMessage(0x0201, "WM_LBUTTONDOWN")

			this.ObjRegisterActive(this, "{6B39CAA1-A320-4CB0-8DB4-352AA81E460E}")
			this.OnExit("Revoke")

	  }

	Hide() {
		Gui, % this.hwnd ": Show", % "Hide NA"
	}

	Destroy() {
		Gui, % this.hwnd ": Destroy"
	}

	Show() {
		Gui, % this.hwnd ": Show", % "NA"
	}

	Notify(msg) {

		; make ShowInfos global in your script to handle output to this gui
		msg := RegExReplace(msg, "[\n\r]+")
		If (StrLen(msg) > 0) {
			Edit_prepend(this.hEdit, msg)
			ControlSend,, {Enter}, % "ahk_id " this.hEdit
		}

	return
	}

	Revoke() {
		 ; This "revokes" the object, preventing any new clients from connecting
		 ; to it, but doesn't disconnect any clients that are already connected.
		 ; In practice, it's quite unnecessary to do this on exit.
		 this.ObjRegisterActive(This, "")
	  return
	  }

	AddDebugView(DbgVName, DbgTitle, Options) {

		 viewfound := false
		 For dbgIndex, obj in This.dbgV {
			   If (obj.name = DbgVName) {
					 viewfound := true
					 break
			   }
		 }

		 If !viewfound {

			 this.dbgV.Push({"name": DbgVName, "title":Tile, "options":Options})

		 }

	  }

	class StackifyObject {
			 ; Simple message-passing example.
			 Message(Data) {
				 ;~ SciTEOutput("received message: " Data)
				 return 42
			 }
			 ; "Worker thread" example.
			 static WorkQueue := []
			 BeginWork(WorkOrder) {
				 this.WorkQueue.Insert(WorkOrder)
				 SetTimer Work, -100
				 return
				 Work:
				 ActiveObject.Work()
				 return
			 }
			 Work() {
				 work := this.WorkQueue.Remove(1)
				 ; Pretend we're working.
				 Sleep 5000
				 ; Tell the boss we're finished.
				 work.Complete(this)
			 }
			 Quit() {
				 MsgBox Quit was called.
				 DetectHiddenWindows On  ; WM_CLOSE=0x10
				 PostMessage 0x10,,,, % "ahk_id " A_ScriptHwnd
				 ; Now return, so the client's call to Quit() succeeds.
			 }
	  }

	ObjRegisterActive(Object, CLSID, Flags:=0) {

	   /*
		  ObjRegisterActive(Object, CLSID, Flags:=0)

			  Registers an object as the active object for a given class ID.
			  Requires AutoHotkey v1.1.17+; may crash earlier versions.

		  Object:
				  Any AutoHotkey object.
		  CLSID:
				  A GUID or ProgID of your own making.
				  Pass an empty string to revoke (unregister) the object.
		  Flags:
				  One of the following values:
					0 (ACTIVEOBJECT_STRONG)
					1 (ACTIVEOBJECT_WEAK)
				  Defaults to 0.

		  Related:
			  http://goo.gl/KJS4Dp - RegisterActiveObject
			  http://goo.gl/no6XAS - ProgID
			  http://goo.gl/obfmDc - CreateGUID()

	  */

		  static cookieJar := {}
		  if (!CLSID) {
			  if (cookie := cookieJar.Remove(Object)) != ""
				  DllCall("oleaut32\RevokeActiveObject", "uint", cookie, "ptr", 0)
			  return
		  }
		  if cookieJar[Object]
			  throw Exception("Object is already registered", -1)
		  VarSetCapacity(_clsid, 16, 0)
		  if (hr := DllCall("ole32\CLSIDFromString", "wstr", CLSID, "ptr", &_clsid)) < 0
			  throw Exception("Invalid CLSID", -1, CLSID)
		  hr := DllCall("oleaut32\RegisterActiveObject", "ptr", &Object, "ptr", &_clsid, "uint", Flags, "uint*", cookie, "uint")
		  if hr < 0
			  throw Exception(format("Error 0x{:x}", hr), -1)
		  cookieJar[Object] := cookie
	  }

	GetActiveObjects(Prefix:="", CaseSensitive:=false) {                                ; GetActiveObjects v1.0 by Lexikos
		; http://ahkscript.org/boards/viewtopic.php?f=6&t=6494
		  objects := {}
		  DllCall("ole32\CoGetMalloc", "uint", 1, "ptr*", malloc) ; malloc: IMalloc
		  DllCall("ole32\CreateBindCtx", "uint", 0, "ptr*", bindCtx) ; bindCtx: IBindCtx
		  DllCall(NumGet(NumGet(bindCtx+0)+8*A_PtrSize), "ptr", bindCtx, "ptr*", rot) ; rot: IRunningObjectTable
		  DllCall(NumGet(NumGet(rot+0)+9*A_PtrSize), "ptr", rot, "ptr*", enum) ; enum: IEnumMoniker
		  while DllCall(NumGet(NumGet(enum+0)+3*A_PtrSize), "ptr", enum, "uint", 1, "ptr*", mon, "ptr", 0) = 0 ; mon: IMoniker
		  {
			  DllCall(NumGet(NumGet(mon+0)+20*A_PtrSize), "ptr", mon, "ptr", bindCtx, "ptr", 0, "ptr*", pname) ; GetDisplayName
			  name := StrGet(pname, "UTF-16")
			  DllCall(NumGet(NumGet(malloc+0)+5*A_PtrSize), "ptr", malloc, "ptr", pname) ; Free
			  if InStr(name, Prefix, CaseSensitive) = 1 {
				  DllCall(NumGet(NumGet(rot+0)+6*A_PtrSize), "ptr", rot, "ptr", mon, "ptr*", punk) ; GetObject
				  ; Wrap the pointer as IDispatch if available, otherwise as IUnknown.
				  if (pdsp := ComObjQuery(punk, "{00020400-0000-0000-C000-000000000046}"))
					  obj := ComObject(9, pdsp, 1), ObjRelease(punk)
				  else
					  obj := ComObject(13, punk, 1)
				  ; Store it in the return array by suffix.
				  objects[SubStr(name, StrLen(Prefix) + 1)] := obj
			  }
			  ObjRelease(mon)
		  }
		  ObjRelease(enum)
		  ObjRelease(rot)
		  ObjRelease(bindCtx)
		  ObjRelease(malloc)
		  return objects
	  }

}




StackifyMin:
	stackifyIsMinimized:= true
return

StackifyQuit:
	userWantsExitApp:= true
return

SfyGuiSize:

	Critical, Off
	Critical
	Gui, Sfy: Default
	GuiControl, Sfy: Move, Stackify	, % "w" (A_GuiWidth) " h" A_GuiHeight-10
	GuiControl, Sfy: Move, StackifyMin	, % "x" (A_GuiWidth-40)
	GuiControl, Sfy: Move, StackifyQuit, % "x" (A_GuiWidth-20)
	Critical, Off

return


WM_LBUTTONDOWN(wP, lP) {
PostMessage, 0xA1, 2
}

Notify(msg) {

	; make ShowInfos global in your script to handle output to this gui

	global hStackifyEdit, hStacker
	static StackifyLine := 0
	static t

	If !ShowInfos
		return

	StackifyLine ++
	str := SubStr("00000" StackifyLine, -4) " " msg "`r`n"

	;If !WinExist("Stackify ahk_class AutohotkeyGUI")
	;~ If !hStacker
		;~ StackifyGui()

	Edit_prepend(hStackifyEdit, str)

	;~ SendMessage, 0x000E	, 0           	, 0           	,, % "ahk_id " hStackify	;WM_GETTEXTLENGTH
	;~ iPos:= ErrorLevel
	;~ SendMessage, 0xB1	, iPos    	, iPos    	,, % "ahk_id " hStackify  	;EM_SETSEL
    ;~ SendMessage, 0xC2	, false    	, % &str    	,, % "ahk_id " hStackify  	;EM_REPLACESEL
	;~ SendMessage, 0x115, 7           	, 0			,, % "ahk_id " hStackify
	;~ ToolTip, % t .= iPos "`n"
return
}

Edit_Prepend(hEdit, Text) { ;www.autohotkey.com/community/viewtopic.php?p=565894#p565894
iPos:= DllCall( "SendMessage", UInt,hEdit, UInt,0x0E	, UInt,0 , UInt,0 ) 	; WM_GetTextLength
DllCall( "SendMessage", UInt,hEdit, UInt,0xB1	, UInt,iPos , UInt,iPos )    	; EM_SETSEL
DllCall( "SendMessage", UInt,hEdit, UInt,0xC2, UInt,0 	, UInt,&Text ) 	; EM_REPLACESEL
iPos:= DllCall( "SendMessage", UInt,hEdit, UInt,0x0E	, UInt,0 , UInt,0 ) 	; WM_GetTextLength
DllCall( "SendMessage", UInt,hEdit, UInt,0xB1	, UInt,iPos	, UInt,iPos )       	; EM_SETSEL
}

Class CtlColors {
   ; ========================================================
   ; Class variables
   ; ========================================================
   ; Registered Controls
   Static Attached := {}
   ; OnMessage Handlers
   Static HandledMessages := {Edit: 0, ListBox: 0, Static: 0}
   ; Message Handler Function
   Static MessageHandler := "CtlColors_OnMessage"
   ; Windows Messages
   Static WM_CTLCOLOR := {Edit: 0x0133, ListBox: 0x134, Static: 0x0138}
   ; HTML Colors (BGR)
   Static HTML := {AQUA: 0xFFFF00, BLACK: 0x000000, BLUE: 0xFF0000, FUCHSIA: 0xFF00FF, GRAY: 0x808080, GREEN: 0x008000
                 , LIME: 0x00FF00, MAROON: 0x000080, NAVY: 0x800000, OLIVE: 0x008080, PURPLE: 0x800080, RED: 0x0000FF
                 , SILVER: 0xC0C0C0, TEAL: 0x808000, WHITE: 0xFFFFFF, YELLOW: 0x00FFFF}
   ; Transparent Brush
   Static NullBrush := DllCall("GetStockObject", "Int", 5, "UPtr")
   ; System Colors
   Static SYSCOLORS := {Edit: "", ListBox: "", Static: ""}
   ; Error message in case of errors
   Static ErrorMsg := ""
   ; Class initialization
   Static InitClass := CtlColors.ClassInit()
   ; ===========================================================
   ; Constructor / Destructor
   ;===========================================================
   __New() { ; You must not instantiate this class!
      If (This.InitClass == "!DONE!") { ; external call after class initialization
         This["!Access_Denied!"] := True
         Return False
      }
   }
   ; ----------------------------------------------------------------------------------------------------------------
   __Delete() {
      If This["!Access_Denied!"]
         Return
      This.Free() ; free GDI resources
   }
   ; ===============================================================
   ; ClassInit       Internal creation of a new instance to ensure that __Delete() will be called.
   ; ===============================================================
   ClassInit() {
      CtlColors := New CtlColors
      Return "!DONE!"
   }
   ; ===============================================================
   ; CheckBkColor    Internal check for parameter BkColor.
   ; ===============================================================
   CheckBkColor(ByRef BkColor, Class) {
      This.ErrorMsg := ""
      If (BkColor != "") && !This.HTML.HasKey(BkColor) && !RegExMatch(BkColor, "^[[:xdigit:]]{6}$") {
         This.ErrorMsg := "Invalid parameter BkColor: " . BkColor
         Return False
      }
      BkColor := BkColor = "" ? This.SYSCOLORS[Class]
              :  This.HTML.HasKey(BkColor) ? This.HTML[BkColor]
              :  "0x" . SubStr(BkColor, 5, 2) . SubStr(BkColor, 3, 2) . SubStr(BkColor, 1, 2)
      Return True
   }
   ; ===============================================================
   ; CheckTxColor    Internal check for parameter TxColor.
   ; ===============================================================
   CheckTxColor(ByRef TxColor) {
      This.ErrorMsg := ""
      If (TxColor != "") && !This.HTML.HasKey(TxColor) && !RegExMatch(TxColor, "i)^[[:xdigit:]]{6}$") {
         This.ErrorMsg := "Invalid parameter TextColor: " . TxColor
         Return False
      }
      TxColor := TxColor = "" ? ""
              :  This.HTML.HasKey(TxColor) ? This.HTML[TxColor]
              :  "0x" . SubStr(TxColor, 5, 2) . SubStr(TxColor, 3, 2) . SubStr(TxColor, 1, 2)
      Return True
   }
   ; ===============================================================
   ; Attach          Registers a control for coloring.
   ; Parameters:     HWND        - HWND of the GUI control
   ;                 BkColor     - HTML color name, 6-digit hexadecimal RGB value, or "" for default color
   ;                 ----------- Optional
   ;                 TxColor     - HTML color name, 6-digit hexadecimal RGB value, or "" for default color
   ; Return values:  On success  - True
   ;                 On failure  - False, CtlColors.ErrorMsg contains additional informations
   ; ===============================================================
   Attach(HWND, BkColor, TxColor := "") {
      ; Names of supported classes
      Static ClassNames := {Button: "", ComboBox: "", Edit: "", ListBox: "", Static: ""}
      ; Button styles
      Static BS_CHECKBOX := 0x2, BS_RADIOBUTTON := 0x8
      ; Editstyles
      Static ES_READONLY := 0x800
      ; Default class background colors
      Static COLOR_3DFACE := 15, COLOR_WINDOW := 5
      ; Initialize default background colors on first call -------------------------------------------------------------
      If (This.SYSCOLORS.Edit = "") {
         This.SYSCOLORS.Static := DllCall("User32.dll\GetSysColor", "Int", COLOR_3DFACE, "UInt")
         This.SYSCOLORS.Edit := DllCall("User32.dll\GetSysColor", "Int", COLOR_WINDOW, "UInt")
         This.SYSCOLORS.ListBox := This.SYSCOLORS.Edit
      }
      This.ErrorMsg := ""
      ; Check colors ---------------------------------------------------------------------------------------------------
      If (BkColor = "") && (TxColor = "") {
         This.ErrorMsg := "Both parameters BkColor and TxColor are empty!"
         Return False
      }
      ; Check HWND -----------------------------------------------------------------------------------------------------
      If !(CtrlHwnd := HWND + 0) || !DllCall("User32.dll\IsWindow", "UPtr", HWND, "UInt") {
         This.ErrorMsg := "Invalid parameter HWND: " . HWND
         Return False
      }
      If This.Attached.HasKey(HWND) {
         This.ErrorMsg := "Control " . HWND . " is already registered!"
         Return False
      }
      Hwnds := [CtrlHwnd]
      ; Check control's class ------------------------------------------------------------------------------------------
      Classes := ""
      WinGetClass, CtrlClass, ahk_id %CtrlHwnd%
      This.ErrorMsg := "Unsupported control class: " . CtrlClass
      If !ClassNames.HasKey(CtrlClass)
         Return False
      ControlGet, CtrlStyle, Style, , , ahk_id %CtrlHwnd%
      If (CtrlClass = "Edit")
         Classes := ["Edit", "Static"]
      Else If (CtrlClass = "Button") {
         IF (CtrlStyle & BS_RADIOBUTTON) || (CtrlStyle & BS_CHECKBOX)
            Classes := ["Static"]
         Else
            Return False
      }
      Else If (CtrlClass = "ComboBox") {
         VarSetCapacity(CBBI, 40 + (A_PtrSize * 3), 0)
         NumPut(40 + (A_PtrSize * 3), CBBI, 0, "UInt")
         DllCall("User32.dll\GetComboBoxInfo", "Ptr", CtrlHwnd, "Ptr", &CBBI)
         Hwnds.Insert(NumGet(CBBI, 40 + (A_PtrSize * 2, "UPtr")) + 0)
         Hwnds.Insert(Numget(CBBI, 40 + A_PtrSize, "UPtr") + 0)
         Classes := ["Edit", "Static", "ListBox"]
      }
      If !IsObject(Classes)
         Classes := [CtrlClass]
      ; Check background color -----------------------------------------------------------------------------------------
      If (BkColor <> "Trans")
         If !This.CheckBkColor(BkColor, Classes[1])
            Return False
      ; Check text color -----------------------------------------------------------------------------------------------
      If !This.CheckTxColor(TxColor)
         Return False
      ; Activate message handling on the first call for a class --------------------------------------------------------
      For I, V In Classes {
         If (This.HandledMessages[V] = 0)
            OnMessage(This.WM_CTLCOLOR[V], This.MessageHandler)
         This.HandledMessages[V] += 1
      }
      ; Store values for HWND ------------------------------------------------------------------------------------------
      If (BkColor = "Trans")
         Brush := This.NullBrush
      Else
         Brush := DllCall("Gdi32.dll\CreateSolidBrush", "UInt", BkColor, "UPtr")
      For I, V In Hwnds
         This.Attached[V] := {Brush: Brush, TxColor: TxColor, BkColor: BkColor, Classes: Classes, Hwnds: Hwnds}
      ; Redraw control -------------------------------------------------------------------------------------------------
      DllCall("User32.dll\InvalidateRect", "Ptr", HWND, "Ptr", 0, "Int", 1)
      This.ErrorMsg := ""
      Return True
   }
   ; ===============================================================
   ; Change          Change control colors.
   ; Parameters:     HWND        - HWND of the GUI control
   ;                 BkColor     - HTML color name, 6-digit hexadecimal RGB value, or "" for default color
   ;                 ----------- Optional
   ;                 TxColor     - HTML color name, 6-digit hexadecimal RGB value, or "" for default color
   ; Return values:  On success  - True
   ;                 On failure  - False, CtlColors.ErrorMsg contains additional informations
   ; Remarks:        If the control isn't registered yet, Add() is called instead internally.
   ; ===============================================================
   Change(HWND, BkColor, TxColor := "") {
      ; Check HWND -----------------------------------------------------------------------------------------------------
      This.ErrorMsg := ""
      HWND += 0
      If !This.Attached.HasKey(HWND)
         Return This.Attach(HWND, BkColor, TxColor)
      CTL := This.Attached[HWND]
      ; Check BkColor --------------------------------------------------------------------------------------------------
      If (BkColor <> "Trans")
         If !This.CheckBkColor(BkColor, CTL.Classes[1])
            Return False
      ; Check TxColor ------------------------------------------------------------------------------------------------
      If !This.CheckTxColor(TxColor)
         Return False
      ; Store Colors ---------------------------------------------------------------------------------------------------
      If (BkColor <> CTL.BkColor) {
         If (CTL.Brush) {
            If (Ctl.Brush <> This.NullBrush)
               DllCall("Gdi32.dll\DeleteObject", "Prt", CTL.Brush)
            This.Attached[HWND].Brush := 0
         }
         If (BkColor = "Trans")
            Brush := This.NullBrush
         Else
            Brush := DllCall("Gdi32.dll\CreateSolidBrush", "UInt", BkColor, "UPtr")
         For I, V In CTL.Hwnds {
            This.Attached[V].Brush := Brush
            This.Attached[V].BkColor := BkColor
         }
      }
      For I, V In Ctl.Hwnds
         This.Attached[V].TxColor := TxColor
      This.ErrorMsg := ""
      DllCall("User32.dll\InvalidateRect", "Ptr", HWND, "Ptr", 0, "Int", 1)
      Return True
   }
   ; ===============================================================
   ; Detach          Stop control coloring.
   ; Parameters:     HWND        - HWND of the GUI control
   ; Return values:  On success  - True
   ;                 On failure  - False, CtlColors.ErrorMsg contains additional informations
   ; ===============================================================
   Detach(HWND) {
      This.ErrorMsg := ""
      HWND += 0
      If This.Attached.HasKey(HWND) {
         CTL := This.Attached[HWND].Clone()
         If (CTL.Brush) && (CTL.Brush <> This.NullBrush)
            DllCall("Gdi32.dll\DeleteObject", "Prt", CTL.Brush)
         For I, V In CTL.Classes {
            If This.HandledMessages[V] > 0 {
               This.HandledMessages[V] -= 1
               If This.HandledMessages[V] = 0
                  OnMessage(This.WM_CTLCOLOR[V], "")
         }  }
         For I, V In CTL.Hwnds
            This.Attached.Remove(V, "")
         DllCall("User32.dll\InvalidateRect", "Ptr", HWND, "Ptr", 0, "Int", 1)
         CTL := ""
         Return True
      }
      This.ErrorMsg := "Control " . HWND . " is not registered!"
      Return False
   }
   ; ===============================================================
   ; Free            Stop coloring for all controls and free resources.
   ; Return values:  Always True.
   ; ===============================================================
   Free() {
      For K, V In This.Attached
         If (V.Brush) && (V.Brush <> This.NullBrush)
            DllCall("Gdi32.dll\DeleteObject", "Ptr", V.Brush)
      For K, V In This.HandledMessages
         If (V > 0) {
            OnMessage(This.WM_CTLCOLOR[K], "")
            This.HandledMessages[K] := 0
         }
      This.Attached := {}
      Return True
   }
   ; ===============================================================
   ; IsAttached      Check if the control is registered for coloring.
   ; Parameters:     HWND        - HWND of the GUI control
   ; Return values:  On success  - True
   ;                 On failure  - False
   ; ===============================================================
   IsAttached(HWND) {
      Return This.Attached.HasKey(HWND)
   }
}

CtlColors_OnMessage(HDC, HWND) {
   Critical
   If CtlColors.IsAttached(HWND) {
      CTL := CtlColors.Attached[HWND]
      If (CTL.TxColor != "")
         DllCall("Gdi32.dll\SetTextColor", "Ptr", HDC, "UInt", CTL.TxColor)
      If (CTL.BkColor = "Trans")
         DllCall("Gdi32.dll\SetBkMode", "Ptr", HDC, "UInt", 1) ; TRANSPARENT = 1
      Else
         DllCall("Gdi32.dll\SetBkColor", "Ptr", HDC, "UInt", CTL.BkColor)
      Return CTL.Brush
   }
}




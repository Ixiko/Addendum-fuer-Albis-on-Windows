﻿
;
;		☠	        	 ☠	        	  ☠	        	  ☠          	  ☠	         	  ☠	        	  ☠	        	  ☠
;	☠	☠	     ☠	 ☠	      ☠	  ☠	      ☠	  ☠	      ☠	  ☠	      ☠	  ☠	     ☠	  ☠	     ☠	  ☠
;		☠	        	 ☠	        	  ☠	        	  ☠	         	  ☠		     	  ☠		    	  ☠		    	  ☠
;		☠	        	 ☠	        	  ☠	        	  ☠	         	  ☠		     	  ☠		    	  ☠		    	  ☠
;		☠	        	 ☠	        	  ☠	        	  ☠	         	  ☠		     	  ☠		    	  ☠		    	  ☠
; 	‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›
;
;        	    		    		   ❗  D O   N O T   R E P L A C E   T H I S   L I B R A R Y  ❗
;			    	￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;	         	        this library is modificated for the needs of Addendum für AlbisOnWindows
;		     		＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
;
; 	‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›‹›
;		☠	        	 ☠	        	  ☠	        	  ☠          	  ☠	         	  ☠	        	  ☠	        	  ☠
;		☠	        	 ☠	        	  ☠	        	  ☠	         	  ☠		     	  ☠		    	  ☠		    	  ☠
;		☠	        	 ☠	        	  ☠	        	  ☠	         	  ☠		     	  ☠		    	  ☠		    	  ☠
;	☠	☠	     ☠	 ☠	      ☠	  ☠	      ☠	  ☠	      ☠	  ☠	      ☠	  ☠	     ☠	  ☠	     ☠	  ☠
;		☠	        	 ☠	        	  ☠	        	  ☠	         	  ☠		     	  ☠		    	  ☠		    	  ☠
;


/*  Neutron.ahk v1.0.1 adm

	Copyright (c) 2020 Philip Taylor (known also as GeekDude, G33kDude)
	https://github.com/G33kDude/Neutron.ahk

	MIT License

	Permission is hereby granted, free of charge, to any person obtaining a copy
	of this software and associated documentation files (the "Software"), to deal
	in the Software without restriction, including without limitation the rights
	to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
	copies of the Software, and to permit persons to whom the Software is
	furnished to do so, subject to the following conditions:

	The above copyright notice and this permission notice shall be included in all
	copies or substantial portions of the Software.

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
	IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
	AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
	LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
	SOFTWARE.

 */

class NeutronWindow {

	static TEMPLATE := ""

	;{ --- Constants ---

	static VERSION := "1.0.1 adm"

	; Windows Messages
	, WM_DESTROY                 	:= 0x02
	, WM_SIZE                        	:= 0x05
	, WM_NCCALCSIZE          	:= 0x83
	, WM_NCHITTEST             	:= 0x84
	, WM_NCLBUTTONDOWN 	:= 0xA1
	, WM_KEYDOWN              	:= 0x100
	, WM_KEYUP                     	:= 0x101
	, WM_SYSKEYDOWN         	:= 0x104
	, WM_SYSKEYUP                 	:= 0x105
	, WM_MOUSEMOVE         	:= 0x200
	, WM_LBUTTONDOWN     	:= 0x201

	; Virtual-Key Codes
	, VK_TAB           	:= 0x09
	, VK_SHIFT         	:= 0x10
	, VK_CONTROL 	:= 0x11
	, VK_MENU       	:= 0x12
	, VK_F5             	:= 0x74

	; Non-client hit test values (WM_NCHITTEST)
	, HT_VALUES := [[13, 12, 14], [10, 1, 11], [16, 15, 17]]

	; Registry keys
	, KEY_FBE := "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION"

	; Undoucmented Accent API constants
	; https://withinrafael.com/2018/02/02/adding-acrylic-blur-to-your-windows-10-apps-redstone-4-desktop-apps/
	, ACCENT_ENABLE_BLURBEHIND 	:= 3
	, WCA_ACCENT_POLICY          	:= 19

	; Other constants
	, EXE_NAME := A_IsCompiled ? A_ScriptName : StrSplit(A_AhkPath, "\").Pop()


	; --- Instance Variables ---

	LISTENERS := [this.WM_DESTROY, this.WM_SIZE, this.WM_NCCALCSIZE
						, this.WM_KEYDOWN, this.WM_KEYUP, this.WM_SYSKEYDOWN, this.WM_SYSKEYUP
						, this.WM_LBUTTONDOWN]

	; Maximum pixel inset for sizing handles to appear
	border_size := 6

	; The window size
	w := 800
	h := 1024

	; Modifier keys as seen by neutron
	MODIFIER_BITMAP := {this.VK_SHIFT: 1<<0, this.VK_CONTROL: 1<<1
									, 	this.VK_MENU: 1<<2}
	modifiers := 0

	; Shortcuts to not pass on to the web control
	disabled_shortcuts :=
	( Join ; ahk
	{
		0: {
			this.VK_F5: true
		},
		this.MODIFIER_BITMAP[this.VK_CONTROL]: {
			GetKeyVK("F"): true,
			GetKeyVK("L"): true,
			GetKeyVK("N"): true,
			GetKeyVK("O"): true,
			GetKeyVK("P"): true
		}
	}
	)
	;}

	;{ --- Properties ---

	; Get the JS DOM object
	doc[]	{
		get
		{
			return this.wb.Document
		}
	}

	; Get the JS Window object
	wnd[]	{
		get
		{
			return this.wb.Document.parentWindow
		}
	}
	;}

	; --- Construction, Destruction, Meta-Functions ---

	__New(html:="", css:="", js:="", title:="Neutron", options:="w600 h400") 	{

		static wb

		; Create necessary circular references
		this.bound := {}
		this.bound._OnMessage := this._OnMessage.Bind(this)

		; Bind message handlers
		for i, message in this.LISTENERS
			OnMessage(message, this.bound._OnMessage)

		; gui size and pos
		this.x := RegExMatch(options, "i)(^|\s)w(?<X>\d+)", gui)  	? guiX : 0
		this.y := RegExMatch(options, "i)(^|\s)w(?<Y>\d+)", gui)  	? guiY : 0
		;~ this.w := RegExMatch(options, "i)(^|\s)w(?<W>\d+)", gui)	? guiW : 800
		;~ this.h := RegExMatch(options, "i)(^|\s)w(?<H>\d+)", gui)  	? guiH : 600

		; Get MinSize if avaible
		If RegExMatch(options, "i)(^|\s)MinSize\d+x\d+", MinSize)
			options := StrReplace(options, MinSize)

		; embed neutron gui
		RegExMatch(options, "i)(^|\s)[\+\-]Parent[0-9a-fx]+", Parent)

		; Gui resize options
		If !RegExMatch(options, "i)(^|\s)[\+\-]Resize", GuiResize)
			GuiResize := "+Resize"

		; Create and save the GUI
		; TODO: Restore previous default GUI
		Gui, New, % "+hWndhWnd -DPIScale " GuiResize " " MinSize " " Parent
		this.hWnd := hWnd

		; Enable shadow
		VarSetCapacity(margins, 16, 0)
		NumPut(1, &margins, 0, "Int")
		DllCall("Dwmapi\DwmExtendFrameIntoClientArea", "UPtr", hWnd, "UPtr", &margins)

		/*	When manually resizing a window, the contents of the window often "lag
			behind" the new window boundaries. Until they catch up, Windows will
			render the border and default window color to fill that area. On most
			windows this will cause no issue, but for borderless windows this can
			cause rendering artifacts such as thin borders or unwanted colors to
			appear in that area until the rest of the window catches up.

			When creating a dark-themed application, these artifacts can cause
			jarringly visible bright areas. This can be mitigated some by changing
			the window settings to cause dark/black artifacts, but it's not a
			generalizable approach, so if I were to do that here it could cause
			issues with light-themed apps.

			Some borderless window libraries, such as rossy's C implementation
			(https://github.com/rossy/borderless-window) hide these artifacts by
			playing with the window transparency settings which make them go away
			but also makes it impossible to show certain colors (in rossy's case,
			Fuchsia/FF00FF).

			Luckly, there's an undocumented Windows API function in user32.dll
			called SetWindowCompositionAttribute, which allows you to change the
			window accenting policies. This tells the DWM compositor how to fill
			in areas that aren't covered by controls. By enabling the "blurbehind"
			accent policy, Windows will render a blurred version of the screen
			contents behind your window in that area, which will not be visually
			jarring regardless of the colors of your application or those behind
			it.

			Because this API is undocumented (and unavailable in Windows versions
			below 10) it's not a one-size-fits-all solution, and could break with
			future system updates. Hopefully a better soultion for the problem
			this hack addresses can be found for future releases of this library.

			https://withinrafael.com/2018/02/02/adding-acrylic-blur-to-your-windows-10-apps-redstone-4-desktop-apps/
			https://github.com/melak47/BorderlessWindow/issues/13#issuecomment-309154142
			http://undoc.airesoft.co.uk/user32.dll/SetWindowCompositionAttribute.php
			https://gist.github.com/riverar/fd6525579d6bbafc6e48
			https://vhanla.codigobit.info/2015/07/enable-windows-10-aero-glass-aka-blur.html
		 */

		Gui, Color, 0, 0
		VarSetCapacity(wcad, A_PtrSize+A_PtrSize+4, 0)
		NumPut(this.WCA_ACCENT_POLICY, &wcad, 0, "Int")
		VarSetCapacity(accent, 16, 0)
		NumPut(this.ACCENT_ENABLE_BLURBEHIND, &accent, 0, "Int")
		NumPut(&accent, &wcad, A_PtrSize, "Ptr")
		NumPut(16, &wcad, A_PtrSize+A_PtrSize, "Int")
		DllCall("SetWindowCompositionAttribute", "UPtr", hWnd, "UPtr", &wcad)

		/* Creating an ActiveX control with a valid URL instantiates a
			WebBrowser, saving its object to the associated variable. The "about"
			URL scheme allows us to start the control on either a blank page, or a
			page with some HTML content pre-loaded by passing HTML after the
			colon: "about:<!DOCTYPE html><body>...</body>"

			Read more about the WebBrowser control here:
			http://msdn.microsoft.com/en-us/library/aa752085

			For backwards compatibility reasons, the WebBrowser control defaults
			to IE7 emulation mode. The standard method of mitigating this is to
			include a compatibility meta tag in the HTML, but this requires
			tampering to the HTML and does not solve all compatibility issues.
			By tweaking the registry before and after creation of the control we
			can opt-out of the browser emulation feature altogether with minimal
			impact on the rest of the system.

			Read more about browser compatibility modes here:
			https://docs.microsoft.com/en-us/archive/blogs/patricka/controlling-webbrowser-control-compatibility
		 */

		RegRead, fbe, % this.KEY_FBE, % this.EXE_NAME
		RegWrite, REG_DWORD, % this.KEY_FBE, % this.EXE_NAME, 11000
		Gui, Add, ActiveX, % "vwb hWndhWB x" this.x " y" this.y " w" this.w " h" this.h, about:blank
		if (fbe = "")
			RegDelete, % this.KEY_FBE, % this.EXE_NAME
		else
			RegWrite, REG_DWORD, % this.KEY_FBE, % this.EXE_NAME, % fbe

		; Save the WebBrowser control to reference later
		this.wb  	:= wb
		this.hWB 	:= hWB

		; Connect the web browser's event stream to a new event handler object
		ComObjConnect(this.wb, new this.WBEvents(this))

		; Compute the HTML template if necessary
		if !(html ~= "i)^<!DOCTYPE")
			html := Format(this.TEMPLATE, css, title, html, js)

		; Write the given content to the page
		this.doc.write(html)
		this.doc.close()

		; Inject the AHK objects into the JS scope
		this.wnd.neutron 	:= this
		this.wnd.ahk     	:= new this.Dispatch(this)

		; Wait for the page to finish loading
		while wb.readyState < 4
			Sleep, 50

		; Subclass the rendered Internet Explorer_Server control to intercept
		; its events, including WM_NCHITTEST and WM_NCLBUTTONDOWN.
		; Read more here: https://forum.juce.com/t/_/27937
		; And in the AutoHotkey documentation for RegisterCallback (Example 2)

		dhw := A_DetectHiddenWindows
		DetectHiddenWindows, On
		ControlGet, hWnd, hWnd,, Internet Explorer_Server1	, % "ahk_id" this.hWnd
		this.hIES    	:= hWnd
		ControlGet, hWnd, hWnd,, Shell DocObject View1  	, % "ahk_id" this.hWnd
		this.hSDOV 	:= hWnd
		DetectHiddenWindows, % dhw

		this.pWndProc    	:= RegisterCallback(this._WindowProc, "", 4, &this)
		this.pWndProcOld := DllCall("SetWindowLong" (A_PtrSize == 8 ? "Ptr" : "")
													, "Ptr", this.hIES   		  	; HWND     hWnd
													, "Int", -4      			      	; int      nIndex (GWLP_WNDPROC)
													, "Ptr", this.pWndProc 	; LONG_PTR dwNewLong
													, "Ptr") 							; LONG_PTR

		; Stop the WebBrowser control from consuming file drag and drop events
		this.wb.RegisterAsDropTarget := False
		DllCall("ole32\RevokeDragDrop", "UPtr", this.hIES)
	}

	; Show an alert for debugging purposes when the class gets garbage collected*
	; (* happens when you empty the object)
	 __Delete() {
	 	;MsgBox, __Delete
		this.Closed := true
	}


	; --- Event Handlers ---

	_OnMessage(wParam, lParam, Msg, hWnd)	{
		if (hWnd == this.hWnd)	{
			; Handle messages for the main window

			if (Msg == this.WM_NCCALCSIZE)	{
				; Size the client area to fill the entire window.
				; See this project for more information:
				; https://github.com/rossy/borderless-window

				; Fill client area when not maximized
				if !DllCall("IsZoomed", "UPtr", hWnd)
					return 0
				; else crop borders to prevent screen overhang

				; Query for the window's border size
				VarSetCapacity(windowinfo, 60, 0)
				NumPut(60, windowinfo, 0, "UInt")
				DllCall("GetWindowInfo", "UPtr", hWnd, "UPtr", &windowinfo)
				cxWindowBorders := NumGet(windowinfo, 48, "Int")
				cyWindowBorders := NumGet(windowinfo, 52, "Int")

				; Inset the client rect by the border size
				NumPut(NumGet(lParam+0, "Int")  	+	cxWindowBorders	, lParam+0	, "Int")
				NumPut(NumGet(lParam+4, "Int")  	+	cyWindowBorders	, lParam+4	, "Int")
				NumPut(NumGet(lParam+8, "Int")  	- 	cxWindowBorders	, lParam+8	, "Int")
				NumPut(NumGet(lParam+12, "Int") 	- 	cyWindowBorders	, lParam+12	, "Int")

				return 0
			}
			else if (Msg == this.WM_SIZE)		{
				; Extract size from LOWORD and HIWORD (preserving sign)
				this.w	:= w	:= lParam<<48>>48
				this.h 	:= h	:= lParam<<32>>48

				DllCall("MoveWindow", "UPtr", this.hWB, "Int", 0, "Int", 0, "Int", w, "Int", h, "UInt", 0)

				return 0
			}
			else if (Msg == this.WM_DESTROY)	{
				; Clean up all our circular references so that the object may be
				; garbage collected.

				for i, message in this.LISTENERS
					OnMessage(message, this.bound._OnMessage, 0)
				ComObjConnect(this.wb)
				this.bound := []
			}
		}
		else if (hWnd == this.hIES || hWnd == this.hSDOV)		{
			; Handle messages for the rendered Internet Explorer_Server

			pressed := (Msg == this.WM_KEYDOWN || Msg == this.WM_SYSKEYDOWN)
			released := (Msg == this.WM_KEYUP || Msg == this.WM_SYSKEYUP)

			if (pressed || released)			{
				; Track modifier states
				if (bit := this.MODIFIER_BITMAP[wParam])
					this.modifiers := (this.modifiers & ~bit) | (pressed * bit)

				; Block disabled key combinations
				if (this.disabled_shortcuts[this.modifiers, wParam])
					return 0

				; When you press tab with the last tabbable item in the
				; document already selected, focus will be taken from the IES
				; control and moved to the SDOV control. The accelerator code
				; from the AutoHotkey installer uses a conditional loop in an
				; attempt to work around this behavior, but as implemented it
				; did not work correctly on my system. Instead, listen for the
				; tab up event on the SDOV and swap it for a tab down before
				; translating it. This should prevent the user from tabbing to
				; the SDOV in most cases, though there may still be some way to
				; tab to it that I am not aware of. A more elegant solution may
				; be to subclass the SDOV like was done for the IES, then
				; forward the WM_SETFOCUS message back to the IES control.
				; However, given the relative complexity of subclassing and the
				; fact that this message substution approach appears to work
				; just as well, we will use the message substitution. Consider
				; implementing the other approach if it turns out that the
				; undesirable behavior continues to manifest under some
				; circumstances.
				Msg := hWnd == this.hSDOV ? this.WM_KEYDOWN : Msg

				; Modified accelerator handling code from AutoHotkey Installer
				Gui +OwnDialogs ; For threadless callbacks which interrupt this.
				pipa := ComObjQuery(this.wb, "{00000117-0000-0000-C000-000000000046}")
				VarSetCapacity(kMsg, 48), NumPut(A_GuiY, NumPut(A_GuiX
				, NumPut(A_EventInfo, NumPut(lParam, NumPut(wParam
				, NumPut(Msg, NumPut(hWnd, kMsg)))), "uint"), "int"), "int")
				r := DllCall(NumGet(NumGet(1*pipa)+5*A_PtrSize), "ptr", pipa, "ptr", &kMsg)
				ObjRelease(pipa)

				if (r == 0) ; S_OK: the message was translated to an accelerator.
					return 0
				return
			}
		}
	}

	_WindowProc(Msg, wParam, lParam)	{
		Critical
		hWnd := this
		this := Object(A_EventInfo)

		if (Msg == this.WM_NCHITTEST)		{
			; Check to see if the cursor is near the window border, which
			; should be treated as the "non-client" drag-to-resize area.
			; https://autohotkey.com/board/topic/23969-/#entry155480

			; Extract coordinates from LOWORD and HIWORD (preserving sign)
			x := lParam<<48>>48, y := lParam<<32>>48

			; Get the window position for comparison
			WinGetPos, wX, wY, wW, wH, % "ahk_id" this.hWnd

			; Calculate positions in the lookup tables
			row := (x < wX + this.BORDER_SIZE) ? 1 : (x >= wX + wW - this.BORDER_SIZE) ? 3 : 2
			col := (y < wY + this.BORDER_SIZE) ? 1 : (y >= wY + wH - this.BORDER_SIZE) ? 3 : 2

			return this.HT_VALUES[col, row]
		}
		else if (Msg == this.WM_NCLBUTTONDOWN)		{
			; Hoist nonclient clicks to main window
			return DllCall("SendMessage", "Ptr", this.hWnd, "UInt", Msg, "UPtr", wParam, "Ptr", lParam, "Ptr")
		}

		; Otherwise (since above didn't return), pass all unhandled events to the original WindowProc.
		Critical, Off
		return DllCall("CallWindowProc"	, "Ptr", this.pWndProcOld 	; WNDPROC lpPrevWndFunc
														, "Ptr", hWnd 			            ; HWND    	hWnd
														, "UInt", Msg             			; UINT    		Msg
														, "UPtr", wParam          		; WPARAM  	wParam
														, "Ptr", lParam           			; LPARAM  	lParam
														, "Ptr") 								; LRESULT
	}


	; --- Instance Methods ---

	; Triggers window dragging. Call this on mouse click down. Best used as your title bar's onmousedown attribute.
	DragTitleBar()	{
		PostMessage, this.WM_NCLBUTTONDOWN, 2, 0,, % "ahk_id" this.hWnd
	}

	; Minimizes the Neutron window. Best used in your title bar's minimize button's onclick attribute.
	Minimize()	{
		Gui, % this.hWnd ":Minimize"
	}

	; Maximize the Neutron window. Best used in your title bar's maximize button's onclick attribute.
	Maximize()	{
		Gui, % this.hWnd (DllCall("IsZoomed", "UPtr", this.hWnd) ? ":Restore" : ":Maximize")
	}

	; Closes the Neutron window. Best used in your title bar's close button's onclick attribute.
	Close()	{
		WinClose, % "ahk_id" this.hWnd
	}

	; Hides the Nuetron window.
	Hide()	{
		Gui, % this.hWnd ":Hide"
	}

	; Destroys the Neutron window.
	Destroy()	{
		; Do this when you would no longer want to
		; re-show the window, as it will free the memory taken up by the GUI and
		; ActiveX control. This method is best used either as your title bar's close
		; button's onclick attribute, or in a custom window close routine.
		Gui, % this.hWnd ":Destroy"
	}

	; Shows a hidden Neutron window.
	Show(options:="")	{

		w	:= RegExMatch(options, "w\s*\K\d+", match)  	? match 	: this.w
		h	:= RegExMatch(options, "h\s*\K\d+", match)  	? match 	: this.h
		m	:= RegExMatch(options, "i)Maximize", MinMax) 	? MinMax : ""

		options := RegExReplace(options, "w\s*\K\d+")
		options := RegExReplace(options, "h\s*\K\d+")


		; AutoHotkey sizes the window incorrectly, trying to account for borders
		; that aren't actually there. Call the function AHK uses to offset and
		; apply the change in reverse to get the actual wanted size.
		VarSetCapacity(rect, 16, 0)
		DllCall("AdjustWindowRectEx"	, "Ptr", &rect                    	; 	LPRECT lpRect
													, "UInt", 0x80CE0000    	; 	DWORD  dwStyle
													, "UInt", 0                      	; 	BOOL   bMenu
													, "UInt", 0                      	; 	DWORD  dwExStyle
													, "UInt")                         	;	BOOL
		w += NumGet(&rect, 0, "Int")-NumGet(&rect, 8, "Int")
		h += NumGet(&rect, 4, "Int")-NumGet(&rect, 12, "Int")

		Gui, % this.hWnd ":Show", % "w" w " h" h " " m	; " " options
	}

	; Loads an HTML file by name (not path).
	Load(fileName)	{

		/*	When running the script uncompiled, looks for the file in the local directory. When running the script
			compiled, looks for the file in the EXE's RCDATA. Files included in your
			compiled EXE by FileInstall are stored in RCDATA whether they get
			extracted or not. An easy way to get your Neutron resources into a
			compiled script, then, is to put FileInstall commands for them right below
			the return at the bottom of your AutoExecute section.

			Parameters:
			fileName - The name of the HTML file to load into the Neutron window.
						Make sure to give just the file name, not the full path.

			Returns: nothing

			Example:

			; AutoExecute Section
			neutron := new NeutronWindow()
			neutron.Load("index.html")
			neutron.Show()
			return
			FileInstall, index.html, index.html
			FileInstall, index.css, index.css
		 */

		; Complete the path based on compiled state
		if A_IsCompiled
			this.wb.Navigate("res://" this.wnd.encodeURIComponent(A_ScriptFullPath) "/10/" fileName)
		else if RegExMatch(fileName, "[A-Z]:\\")
			this.wb.Navigate(fileName)
		else
			this.wb.Navigate(fileName)

		; Navigate to the calculated file URL


		; Wait for the page to finish loading
		while this.wb.readyState < 3
			Sleep, 50

		; Inject the AHK objects into the JS scope
		this.wnd.neutron 	:= this
		this.wnd.ahk     	:= new this.Dispatch(this)

		; Wait for the page to finish loading
		while this.wb.readyState < 4
			Sleep, 50
	}

	; Shorthand method for document.querySelector
	qs(selector)	{
		return this.doc.querySelector(selector)
	}

	; Shorthand method for document.querySelectorAll
	qsa(selector)	{
		return this.doc.querySelectorAll(selector)
	}

	; Passthrough method for the Gui command, targeted at the Neutron Window
	; instance
	Gui(subCommand, value1:="", value2:="", value3:="")	{
		Gui, % this.hWnd ":" subCommand, %value1%, %value2%, %value3%
	}


	; --- Static Methods ---

	; Given an HTML Collection (or other JavaScript array), return an enumerator
	Each(htmlCollection)	{

		/*  that will iterate over its items.

			Parameters:
				htmlCollection - The JavaScript array to be iterated over

			Returns: An Enumerable object

			Example:

			neutron := new NeutronWindow("<body><p>A</p><p>B</p><p>C</p></body>")
			neutron.Show()
			for i, element in neutron.Each(neutron.body.children)
				MsgBox, % i ": " element.innerText
		 */

		return new this.Enumerable(htmlCollection)
	}

	; Given an HTML Form Element, construct a FormData object
	GetFormData(formElement, useIdAsName:=True)	{

		/*  Parameters:
			formElement 	- The HTML Form Element
			useIdAsName 	- When a field's name is blank, use it's ID instead

			Returns: 			- A FormData object

			Example:
			neutron := new NeutronWindow("<form>"
				. "<input type='text' name='field1' value='One'>"
				. "<input type='text' name='field2' value='Two'>"
				. "<input type='text' name='field3' value='Three'>"
				. "</form>")
			neutron.Show()
			formElement	:= neutron.doc.querySelector("form")      	; Grab 1st form on page
			formData  	:= neutron.GetFormData(formElement)  	; Get form data
			MsgBox, % formData.field2 ; Pull a single field
			for name, element in formData ; Iterate all fields
				MsgBox, %name%: %element%
		 */

		formData := new this.FormData()

		for i, field in this.Each(formElement.elements)		{
		  ; Discover the field's name
			name := ""
			try ; fieldset elements error when reading the name field
				name := field.name
			if (name == "" && useIdAsName)
				name := field.id

		  ; Filter against fields which should be omitted
			if (name == "" || field.disabled || field.type ~= "^file|reset|submit|button$")
				continue

		  ; Handle select-multiple variants
			if (field.type == "select-multiple")	{
				for j, option in this.Each(field.options)
					if (option.selected)
						formData.add(name, option.value)
				continue
			}

		  ; Filter against unchecked checkboxes and radios
			if (field.type ~= "^checkbox|radio$" && !field.checked)
				continue

		  ; Return the field values
			formData.add(name, field.value)
		}

		return formData
	}

	; Given a potentially HTML-unsafe string, return an HTML safe string
	; https://stackoverflow.com/a/6234804
	EscapeHTML(unsafe)	{
		unsafe := StrReplace(unsafe, "&", "&amp;")
		unsafe := StrReplace(unsafe, "<", "&lt;")
		unsafe := StrReplace(unsafe, ">", "&gt;")
		unsafe := StrReplace(unsafe, """", "&quot;")
		unsafe := StrReplace(unsafe, "''", "&#039;")
		return unsafe
	}

	; Wrapper for Format that applies EscapeHTML to each value before passing
	; them on. Useful for dynamic HTML generation.
	FormatHTML(formatStr, values*)	{
		for i, value in values
			values[i] := this.EscapeHTML(value)
		return Format(formatStr, values*)
	}


	; --- Nested Classes ---	 FOR INTERNAL USE ONLY.

	; Proxies method calls to AHK function calls, binding a given value to the first parameter of the target function.
	class Dispatch	{

		; Parameters:
		;   parent - The value to bind

		__New(parent)		{
			this.parent := parent
		}

		__Call(params*)		{
			; Make sure the given name is a function
			if !(fn := Func(params[1]))
				throw Exception("Unknown function: " params[1])

			; Make sure enough parameters were given
			if (params.length() < fn.MinParams)
				throw Exception("Too few parameters given to " fn.Name ": " params.length() " [" fn.MinParams " parameters needed]")

			; Make sure too many parameters weren't given
			if (params.length() > fn.MaxParams && !fn.IsVariadic)
				throw Exception("Too many parameters given to " fn.Name ": " params.length() " [a maximum of " fn.MaxParams " parameters is aloud]")

			; Change first parameter from the function name to the neutron instance
			params[1] := this.parent

			; Call the function
			return fn.Call(params*)
		}
	}

	; Handles Web Browser events
	class WBEvents	{

		; https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa768283%28v%3dvs.85%29
		;
		; For internal use only
		;
		; Parameters:
		;   parent - An instance of the Neutron class
		;

		__New(parent)		{
			this.parent := parent
		}

		DocumentComplete(wb)		{
			; Inject the AHK objects into the JS scope
			wb.document.parentWindow.neutron	:= this.parent
			wb.document.parentWindow.ahk    	:= new this.parent.Dispatch(this.parent)
		}
	}

	; Enumerator class that enumerates the items of an HTMLCollection (or other JavaScript array).
	class Enumerable	{

		; Best accessed through the .Each() helper method.
		;
		; Parameters:
		;   htmlCollection - The HTMLCollection to be enumerated.

		i := 0

		__New(htmlCollection)	{
			this.collection := htmlCollection
		}

		_NewEnum()	{
			return this
		}

		Next(ByRef i, ByRef elem)		{
			if (this.i >= this.collection.length)
				return False
			i := this.i
			elem := this.collection.item(this.i++)
			return True
		}
	}

	; A collection similar to an OrderedDict designed for holding form data.
	class FormData	{

		; This collection allows duplicate keys and enumerates key value pairs in
		; the order they were added.
		names 	:= []
		values	:= []

		; Add a field to the FormData structure.
		Add(name, value)		{

			; Parameters:
			;   name - The form field name associated with the value
			;   value - The value of the form field
			;
			; Returns: Nothing

			this.names.Push(name)
			this.values.Push(value)
		}

		; Get an array of all values associated with a name.
		All(name)		{

			; Parameters:
			;   name - The form field name associated with the values
			;
			; Returns: An array of values
			;
			; Example:
			;
			; fd := new NeutronWindow.FormData()
			; fd.Add("foods", "hamburgers")
			; fd.Add("foods", "hotdogs")
			; fd.Add("foods", "pizza")
			; fd.Add("colors", "red")
			; fd.Add("colors", "green")
			; fd.Add("colors", "blue")
			; for i, food in fd.All("foods")
			;     out .= i ": " food "`n"
			; MsgBox, %out%

			values := []
			for i, v in this.names
				if (v == name)
					values.Push(this.values[i])
			return values
		}

		; Meta-function to allow direct access of field values using either dot or bracket notation. Can retrieve the nth
		;  item associated with a given name by passing more than one value in when bracket notation.
		__Get(name, n := 1)		{

			; Example:
			;
			; fd := new NeutronWindow.FormData()
			; fd.Add("foods", "hamburgers")
			; fd.Add("foods", "hotdogs")
			; MsgBox, % fd.foods ; hamburgers
			; MsgBox, % fd["foods", 2] ; hotdogs

			for i, v in this.names
				if (v == name && !--n)
					return this.values[i]
		}

		; Allow iteration in the order fields were added, instead of a normal
		; object's alphanumeric order of iteration.
		_NewEnum()		{

			; Example:
			;
			; fd := new NeutronWindow.FormData()
			; fd.Add("z", "3")
			; fd.Add("y", "2")
			; fd.Add("x", "1")
			; for name, field in fd
			;     out .= name ": " field ","
			; MsgBox, %out% ; z: 3, y: 2, x: 1

			return {"i": 0, "base": this}
		}

		Next(ByRef name, ByRef value)		{
			if (++this.i > this.names.length())
				return False
			name := this.names[this.i]
			value := this.values[this.i]
			return True
		}
	}
}

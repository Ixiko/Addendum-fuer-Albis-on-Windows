; Send Mail to Telegram
; new mail event from Outlook will be used to explore mail details, extract attachments
; and send body and attachment to your prefered Telegram account using Telegram's Bot Api
; scripted with help from just me


	; script options
		#NoEnv
		#Persistent
		SetBatchLines, -1

	; variables
		global MT
		MT     	:= Object()
		MT.ini 	:= A_ScriptDir "\MTgram.ini"
		MT.tmp := A_Temp

		If !FileExist(MT.ini)
			MT := Properties()

		;IniReadExt


	; starts Outllook - Inbox Events
		ComObjError(0)
		OL_App := ComObjActive("Outlook.Application")
		If !IsObject(OL_App) {
		   MsgBox, 16, Fehler!, Bitte vor dem Start des Skripts Outlook öffnen!
		   ExitApp
		}
		ComObjError(1)
		OL_Items := OL_App.GetNamespace("MAPI").GetDefaultFolder(6).Items ; olFolderInbox = 6
		ComObjConnect(OL_Items, "Inbox_")

Return

Inbox_ItemAdd(Item) { ; docs.microsoft.com/en-us/office/vba/api/outlook.items.itemadd

	If (attCount := Item.Attachments.Count) {

		att := attCount "`n"
		att .= "  -Class:`t`t"  	Item.Attachments.Class "`n"
		att .= "  -Parent:`t`t" 	Item.Attachments.Parent "`n"
		Loop, % attCount {
			att .= "  -Name("   	A_Index "):`t" 	Item.Attachments.Item(A_Index).DisplayName "`n"
			att .= "  -Filename("	A_Index "):`t" 	Item.Attachments.Item(A_Index).FileName "`n"
			att .= "  -BlockLevel("	A_Index "):`t" 	Item.Attachments.Item(A_Index).BlockLevel "`n"
			att .= "  -Class("	A_Index "):     `t"      	Item.Attachments.Item(A_Index).Class "`n" ; 5 = olAttachment object
			att .= "  -Size("	A_Index "):     `t"       	Item.Attachments.Item(A_Index).Size " bytes`n"
			att .= "  -Type("	A_Index "):     `t"      	Item.Attachments.Item(A_Index).Type "`n"
		}
		;att .= "`t-Name:`t`t" 	Item.Attachments.Attachment[1].DisplayName

	}
	else
		attCount := "no attachements"

   SciTEOutput(""
			. "New Mail:`n"
            . "From:`t`t"      	Item.SenderEmailAddress "`n"
            . "Subject:`t`t"   	Item.Subject "`n"
            . "Attachments:`t"	att)

}

IniReadExt(SectionOrFullFilePath, Key:="", DefaultValue:="") {                                                  	;-- eigene IniRead funktion für Addendum

	; beim ersten Aufruf der Funktion !nur! Übergabe des ini Pfades mit dem Parameter SectionOrFullFilePath
	; die Funktion behandelt einen geschriebenen Wert der "ja" oder "nein" ist, als Wahrheitswert, also true oder false
	; letzte Änderung: 08.06.2020

		static WorkIni

	; Arbeitsini Datei wird erkannt wenn der übergebene Parameter einen Pfad ist
		If RegExMatch(SectionOrFullFilePath, "^[A-Z]\:.*\\")	{
			If !FileExist(SectionOrFullFilePath)	{
					MsgBox,, Addendum für AlbisOnWindows, % "Die .ini Datei existiert nicht!`n`n" WorkIni "`n`nDas Skript wird jetzt beendet.", 10
					ExitApp
			}
			WorkIni := SectionOrFullFilePath
			return WorkIni
		}

	; Section, Key einlesen
		IniRead, OutPutVar, % WorkIni, % SectionOrFullFilePath, % Key

	; Bearbeiten des Wertes vor Rückgabe
		If InStr(OutPutVar, "ERROR")
			If (StrLen(DefaultValue) > 0) ; Defaultwert vorhanden, dann diesen Schreiben und Zurückgeben
				IniWrite, % (OutPutVar := DefaultValue), % WorkIni, % SectionOrFullFilePath, % key
			else
				return "ERROR"
		else if InStr(OutPutVar, "%AddendumDir%")
				OutPutVar := StrReplace(OutPutVar, "%AddendumDir%", Addendum.AddendumDir)
		else if RegExMatch(OutPutVar, "i)^\s*(ja|nein)\s*$", bool)
				OutPutVar := (bool1= "ja") ? true : false

return Trim(OutPutVar)
}

Properties() {                                                                                                                            ;-- user properties gui

	global

	local	Opt      	:= "Center BackgroundTrans"
		, 	Opt1    	:= "-E0x200 Center cWhite ggHandler"
		, 	BGColor	:= "0293FD"
		,	BGColor1	:= "58B8FE"
		,	FntColor	:= "FFFFFF"
		,	FntColor1	:= "015B9D"
		,	FntOpt1	:= "s15 c015B9D bold"
		,	FntOpt2	:= "s12 cFFFFFF bold underline"
		,	FntOpt3	:= "s13 cFFFFFF normal"
		,	guiW    	:= 770
		,	GPPlus		:= 10
		,	IGPlus		:= 8

	static fcomments:= "; forward, from: test@mail.com, subject contains: support`n; ignore, from: test2@mail.com`n"

	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; GUI +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ ;{

		Gui, ini: new 	, -DPIScale
		Gui, ini: Margin	, 0, 0
		Gui, ini: Color	, % BGColor
		ImageButton.SetGuiColor("0x" . BgColor)


		Gui, ini: Add 	, Progress 	, % "x0 y10 	w" guiW " h40 cFFFFFF vPGIniPath"                           	, 100
		Gui, ini: Font 	, % FntOpt1
		Gui, ini: Add 	, Text       	, % "x10 y12	w" guiW-20 " vGBInipath " Opt                                  	, % "I N I   f i l e p a t h"
		GuiControlGet, c, Pos, GBiniPath
		GuiControl, MoveDraw, PGIniPath, % "h" cH + 4

		Gui, ini: Font 	, s12 cFFFFFF Normal
		Y := cY + cH + IGPlus
		Gui, ini: Add 	, Edit        	, % "x10 y" Y " w700 r1 vIniPath HWNDhIniPath -E0x200 ggHandler"	, % MT.ini
		Gui, ini: Font 	, s10 cFFFFFF bold
		Gui, ini: Add 	, Button    	, % "x+10 h12 vBtnIniPath HWNDhBtnIPath ggHandler Center"      	, % ". . ."
		GuiControl, Focus, PGIniPath

	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

		GuiControlGet, c, Pos, iniPath
		Y := cY + cH + GPPlus
		Gui, ini: Add 	, Progress  	, % "x0 y" Y " w" guiW " h40 cFFFFFF vPGTgram"                     	, 100
		GuiControlGet, c, Pos, PGTgram
		Gui, ini: Font 	, % FntOpt1
		Gui, ini: Add 	, Text        	, % "x10 y" Y+2 " w" guiW-20 " vGBTgram " Opt                     	, % "T e l e g r a m   p r o p e r t i e s"
		GuiControlGet, c, Pos, GBTgram
		GuiControl, MoveDraw, PGTgram, % "h" cH + 4

		Gui, ini: Font 	, % FntOpt2
		Y := cY + cH + IGPlus
		Gui, ini: Add 	, Text         	, % "x" cX       	" y" Y " w180 vTBot1 " Opt                                	, % "B o t N a m e"
		Gui, ini: Add 	, Text         	, % "x" cX+190	" y" Y " w300 vTBot2 " Opt                                	, % "B o t T o k e n"
		Gui, ini: Add 	, Text         	, % "x" cX+500	" y" Y " w250 vTBot3 " Opt                                	, % "C h a t I D"

		Gui, ini: Font 	, % FntOpt3
		GuiControlGet, c, Pos, TBot1
		Y := cY + cH + IGPlus + 2
		Gui, ini: Add 	, Edit         	, % "x" cX " y" Y " w" cw " vBotName HWNDhBotName " Opt1   	, % (MT.BotName 	? MT.BotName	: "- - -")
		GuiControlGet, c, Pos, TBot2
		Gui, ini: Add 	, Edit         	, % "x" cX " y" Y " w" cw " vBotToken HWNDhBotToken " Opt1    	, % (MT.BotToken 	? MT.BotToken	: "- - -")
		GuiControlGet, c, Pos, TBot3
		Gui, ini: Add 	, Edit         	, % "x" cX " y" Y " w" cw " vChatID HWNDhChatID " Opt1           	, % (MT.ChatID  	? MT.ChatID 	: "- - -")

		GuiControlGet, c, Pos, ChatID
		GuiControlGet, d, Pos, GBTgram
		GuiControl, MoveDraw, GBTgram, % "h" cY + cH + 10 - dY

	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

		Y := cY + cH + GPPlus
		Gui, ini: Add 	, Progress  	, % "x0 y" Y " w" guiW " h30 cFFFFFF vPGOutlook"                   	, 100
		GuiControlGet, c, Pos, PGOutlook
		Gui, ini: Font 	, % FntOpt1
		Gui, ini: Add 	, Text        	, % "x10 y" Y+2 " w" guiW-20 " vGBOutlook " Opt                   	, % "O u t l o o k   p r o p e r t i e s"
		GuiControlGet, c, Pos, GBOutlook
		GuiControl, MoveDraw, PGOutllook, % "h" cH

		Gui, ini: Font 	, % FntOpt2
		Y := cY + cH + IGPlus + 2
		Gui, ini: Add 	, Text         	, % "x10 y" Y " w" guiW - 20 " vTFilt1 " Opt                                	, % "E m a i l  -  f i l t e r   p r o p e r t i e s"
		GuiControlGet, c, Pos, TFilt1
		Y := cY + cH + IGPlus
		Gui, ini: Font 	, % FntOpt3
		Gui, ini: Add 	, Edit         	, % "x" cX " y" Y " w" cw " vMFilter HWNDhMFilter r30 -E0x200"    	, % fcomments . MT.MFilter

	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

		GuiControlGet, c, Pos, MFilter
		Y := cY + cH + GPPlus
		Gui, ini: Font 	, s15 cFFFFFF bold
		Gui, ini: Add 	, Button    	, % "x10 y" Y " w100 vBtnSave  	HWNDhSave 	ggHandler Center"                	, % "Save"
		Gui, ini: Add 	, Button    	, % "x60 y" Y " w100 vBtnCancel	HWNDhCancel	ggHandler Center"                	, % "Cancel"
		GuiControlGet, c, Pos, BtnCancel
		GuiControl, MoveDraw, BtnCancel, % "x" guiW - 10 - cw

	; ImageButtons options - [Mode, StartColor, TargetColor, TextColor, Rounded, GuiColor, BorderColor, BorderWidth] ;{
		Btn := Array()
		Btn.save 	:= [ 	[0, 0xFF58B8FE , , 0xFFFFFFFF	, "H", 0xFF0293FD	, 0xFFFFFFFF, 2]     	; normal
							,	[, 0xFF0293FD]                                                                                 	; hot
							,	[0, 0xFF88E8FF	, , 0xFF0293FD	, "H",  0xFF0293FD 	, 0xFFFFFFFF, 4] ]	; pressed
		;	    			, 	[ , , , "White"]                                                                                      	; defaulted text color -> animation
		Btn.Cancel:= [ 	[0, 0xFF58B8FE , , 0xFFFFFFFF	, "H", 0xFF0293FD	, 0xFFFFFFFF, 2]     	; normal
							,	[, 0xFF0293FD]                                                                                 	; hot
							,	[0, 0xFF88E8FF	, , 0xFF0293FD	, "H",  0xFF0293FD 	, 0xFFFFFFFF, 4] ]	; pressed
		;		    		, 	[ , , , "White"]                                                                                      	; defaulted text color -> animation
		Btn.BtnIPath:= [ 	[0, 0xFF58B8FE , , 0xFFFFFFFF	, "H", 0xFF0293FD	, 0xFFFFFFFF, 2] 	; normal
							,	[, 0xFF0293FD]                                                                                 	; hot
							,	[0, 0xFF88E8FF	, , 0xFF0293FD	, "H",  0xFF0293FD 	, 0xFFFFFFFF, 4] ]	; pressed
		;		    		, 	[ , , , "White"]                                                                                      	; defaulted text color -> animation


		ImageButton.Create(hSave     	, Btn.save[1]   	, Btn.save[2]  	, Btn.save[3]  	)
		ImageButton.Create(hCancel 	, Btn.cancel[1]	, Btn.Cancel[2]	, Btn.Cancel[3]	)
		ImageButton.Create(hBtnIPath	, Btn.BtnIPath[1]	, Btn.BtnIPath[2]	, Btn.BtnIPath[3]	)

	;}

	; ControlColors ;{

		CTLCOLORS.Attach(hIniPath  	, BgColor1, FntColor1)
		CTLCOLORS.Attach(hBotName 	, BgColor1, FntColor1)
		CTLCOLORS.Attach(hBotToken 	, BgColor1, FntColor1)
		CTLCOLORS.Attach(hChatID 	, BgColor1, FntColor1)
		CTLCOLORS.Attach(hBtnIniPath	, BgColor1, FntColor1)
		CTLCOLORS.Attach(hMFilter  	, BgColor1, FntColor1)

	;}

		Gui, ini: Show, % "Hide AutoSize", Send mails to Telegram
		Gui, ini: Show, % "h" cY + cH + 10 , Send mails to Telegram

return MT ;}

gHandler: ;{

	SciTEOutput("control: " A_GuiControl ", event: " A_GuiEvent ", info: " A_EventInfo)

	Switch A_GuiControl
	{
		Case "BtnCancel":
			Gui, ini: Destroy
			return
		Case "BtnSave":
			Gui, ini: Submit, NoHide
			props := ParseProperties({	"ini"           	: iniPath
					                    		, 	"BotName"	: StrReplace(BotName, "- - -")
					                    		,	"BotToken" 	: StrReplace(BotToken, "- - -")
					                    		,	"ChatID"    	: StrReplace(ChatID, "- - -")
					                    		,	"MFilter"    	: RTrim(StrReplace(MFilter, fcomments), "`n")
					                    		,	"tmp"         	: MT.tmp})
			If IsObject(props) {
				Gui, ini: Destroy
				return props
			}
	}

return ;}

iniGuiClose:
iniGuiEscape:
ExitApp

}

ParseProperties(props) {                                                                                                            ;-- check properties

	SciTEOutput("inipath: "   	props.ini)
	SciTEOutput("BotName: " 	props.BotName)
	SciTEOutput("BotToken: " 	props.BotToken)
	SciTEOutput("ChatID: "    	props.ChatID)
	SciTEOutput("Mailfilter: "	props.MFilter)
	SciTEOutput("tmp dir: "   	props.tmp)

	If !RegExMatch(props.ini, "[A-Z]\:\\") {
		MsgBox, % "This is not a valid path!"
		return false
	}
	else if !props.BotName {
		MsgBox, % "Your bot must have a name"
		GuiControl, Focus, BotName
		return false
	}
	else if !props.BotToken {
		MsgBox, % "The Telgram BotToken is needed!"
		GuiControl, Focus, BotToken
		return false
	}
	else if !props.ChatID {
		MsgBox, % "The Telgram ChatID is needed for forwarding messages!"
		GuiControl, Focus, ChatID
		return false
	}

return props
}


; Gui classes for changing control colors
Class CTLCOLORS {

   ; =========================================================================================
   ; AHK 1.1 +
   ; =========================================================================================
   ; Function:        Helper object to color controls on WM_CTLCOLOR... notifications.
   ;                    	Supported controls are: Checkbox, ComboBox, DropDownList, Edit, ListBox, Radio, Text.
   ;                    	Checkboxes and Radios accept background colors only due to design.

   ; Namespace:	CTLCOLORS
   ; AHK version:	1.1.11.01
   ; Language:   	English
   ; Version:      	0.9.01.00/2012-04-05/just me
   ;                    	0.9.02.00/2013-06-26/just me  -  fixed to run on Win 7 x64
   ;                    	0.9.03.00/2013-06-27/just me  -  added support for disabled edit controls
   ;
   ; How to use:   	To register a control for coloring call
   ;                       CTLCOLORS.Attach()
   ;                       passing up to three parameters:
   ;                       Hwnd        - Hwnd of the GUI control                                   		(Integer)
   ;                       BkColor     - HTML color name, 6-digit hex value ("RRGGBB")      (String)
   ;                                     		or "" for default color

   ;                       ------------- Optional -------------------------------------------------------------------------
   ;                       TextColor   - HTML color name, 6-digit hex value ("RRGGBB")             (String)
   ;                                     		or "" for default color

   ;                    	If both BkColor and TextColor are "" the control will not be added and the call returns False.
   ;
   ;                    	To change the colors for a registered control call
   ;                       CTLCOLORS.Change()
   ;                    	passing up to three parameters:
   ;                       Hwnd        - see above
   ;                       BkColor     - see above
   ;                       ------------- Optional -------------------------------------------------------------------------
   ;                       TextColor   - see above
   ;                    	Both BkColor and TextColor may be "" to reset them to default colors.
   ;                    	If the control is not registered yet, CTLCOLORS.Attach() is called internally.
   ;
   ;                    	To unregister a control from coloring call
   ;                       CTLCOLORS.Detach()
   ;                    	passing one parameter:
   ;                       Hwnd      - see above
   ;
   ;                    	To stop all coloring and free the resources call
   ;                       CTLCOLORS.Free()
   ;                    	It's a good idea to insert this call into the scripts exit-routine.
   ;
   ;                    	To check if a control is already registered call
   ;                       CTLCOLORS.IsAttached()
   ;                    	passing one parameter:
   ;                       Hwnd      - see above
   ;
   ;                    	To get a control's Hwnd use either the option "HwndOutputVar" with "Gui, Add" or the command
   ;                    	"GuiControlGet" with sub-command "Hwnd".
   ;
   ;                    	Properties/methods/functions declared as PRIVATE must not be set/called by the script!
   ;
   ; Special features:
   ;                    	On the first call for a specific control class the function registers the CTLCOLORS_OnMessage()
   ;                    	function as message handler for WM_CTLCOLOR messages of this class(es).
   ;
   ;                    	Buttons (Checkboxes and Radios) do not make use of the TextColor to draw the text, instead of
   ;                    	that they use it to draw the focus rectangle.
   ;
   ;                    	After displaying the GUI per "Gui, Show" you have to execute "WinSet, Redraw" once.
   ;                    	It's no bad idea to do it using a GuiSize label, because it avoids rare problems when restoring
   ;                    	a minimized window:
   ;                       GuiSize:
   ;                          If (A_EventInfo != 1) {
   ;                             Gui, %A_Gui%:+LastFound
   ;                             WinSet, ReDraw
   ;                          }
   ;                       Return
   ; ================================================================================
   ; This software is provided 'as-is', without any express or implied warranty.
   ; In no event will the authors be held liable for any damages arising from the use of this software.
   ; ================================================================================


   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; PRIVATE Properties and Methods +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; Registered Controls
   Static Attached := {}
   ; OnMessage Handlers
   Static HandledMessages := {Edit: 0, ListBox: 0, Static: 0}
   ; Message Handler Function
   Static MessageHandler := "CTLCOLORS_OnMessage"
   ; Windows Messages
   Static WM_CTLCOLOR := {Edit: 0x0133, ListBox: 0x134, Static: 0x0138}
   ; HTML Colors (BGR)
   Static HTML := {AQUA:    0xFFFF00, BLACK:   0x000000, BLUE:    0xFF0000, FUCHSIA: 0xFF00FF, GRAY:    0x808080
                 , GREEN:   0x008000, LIME:    0x00FF00, MAROON:  0x000080, NAVY:    0x800000, OLIVE:   0x008080
                 , PURPLE:  0x800080, RED:     0x0000FF, SILVER:  0xC0C0C0, TEAL:    0x808000, WHITE:   0xFFFFFF
                 , YELLOW:  0x00FFFF}
   ; System Colors
   Static SYSCOLORS := {Edit: "", ListBox: "", Static: ""}
   Static Initialize := CTLCOLORS.InitClass()
   ; ================================================================================
   ; PRIVATE SUBCLASS CTLCOLORS_Base  - Base class
   ; ================================================================================
   Class CTLCOLORS_Base {
      __New() {   ; This class is a helper object, you must not instantiate it.
         Return False
      }
      __Delete() {
         This.Free()
      }
   }
   ; ==================================================================================
   ; PRIVATE METHOD Init  Class       - Set the base
   ; ==================================================================================
   InitClass() {
      This.Base := This.CTLCOLORS_Base
      Return "DONE"
   }
   ; ==================================================================================
   ; PRIVATE METHOD CheckColors       - Check parameters BkColor and TextColor not to be empty both
   ; ==================================================================================
   CheckColors(BkColor, TextColor) {
      This.ErrorMsg := ""
      If (BkColor = "") && (TextColor = "") {
         This.ErrorMsg := "Both parameters BkColor and TextColor are empty!"
         Return False
      }
      Return True
   }
   ; ==================================================================================
   ; PRIVATE METHOD CheckBkColor      - Check parameter BkColor
   ; ==================================================================================
   CheckBkColor(ByRef BkColor, Class) {
      This.ErrorMsg := ""
      If (BkColor != "") && !This.HTML.HasKey(BkColor) && !RegExMatch(BkColor, "i)^[0-9A-F]{6}$") {
         This.ErrorMsg := "Invalid parameter BkColor: " . BkColor
         Return False
      }
      BkColor := BkColor = "" ? This.SYSCOLORS[Class]
               : This.HTML.HasKey(BkColor) ? This.HTML[BkColor]
               : "0x" . SubStr(BkColor, 5, 2) . SubStr(BkColor, 3, 2) . SubStr(BkColor, 1, 2)
      Return True
   }
   ; ==================================================================================
   ; PRIVATE METHOD CheckTextColor    - Check parameter TextColor
   ; ==================================================================================
   CheckTextColor(ByRef TextColor) {
      This.ErrorMsg := ""
      If (TextColor != "") && !This.HTML.HasKey(TextColor) && !RegExMatch(TextColor, "i)^[\dA-F]{6}$") {
         This.ErrorMsg := "Invalid parameter TextColor: " . TextColor
         Return False
      }
      TextColor := TextColor = "" ? ""
                 : This.HTML.HasKey(TextColor) ? This.HTML[TextColor]
                 : "0x" . SubStr(TextColor, 5, 2) . SubStr(TextColor, 3, 2) . SubStr(TextColor, 1, 2)
      Return True
   }
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; PUBLIC Interface +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; Error message in case of errors
   Static ErrorMsg := ""
   ; ==================================================================================
   ; METHOD Attach      Register control for coloring
   ; Parameters:           	Hwnd        - HWND of the GUI control                                   (Integer)
   ;                       			BkColor     - HTML color name, 6-digit hex value ("RRGGBB")             (String)
   ;                                     				or "" for default color
   ;                       ------------- Optional ----------------------------------------------------------------------
   ;                       			TextColor   - HTML color name, 6-digit hex value ("RRGGBB")             (String)
   ;                                     				or "" for default color
   ; Return values:      	On success  - True
   ;                       			On failure  - False, CTLCOLORS.ErrorMsg contains additional informations
   ; ==================================================================================
   Attach(Hwnd, BkColor, TextColor = "") {
      ; Names of supported classes
      Static ClassNames := {Button: "", ComboBox: "", Edit: "", ListBox: "", Static: ""}
      ; Button styles
      Static BS_CHECKBOX := 0x2
           , BS_RADIOBUTTON := 0x8
      ; Editstyles
      Static ES_READONLY := 0x800
      ; Default class background colors
      Static COLOR_3DFACE := 15
           , COLOR_WINDOW := 5
      ; Initialize default background colors on first call -------------------------------------------------------------
      If (This.SYSCOLORS.Edit = "") {
         This.SYSCOLORS.Static := DllCall("User32.dll\GetSysColor", "Int", COLOR_3DFACE, "UInt")
         This.SYSCOLORS.Edit := DllCall("User32.dll\GetSysColor", "Int", COLOR_WINDOW, "UInt")
         This.SYSCOLORS.ListBox := This.SYSCOLORS.Edit
      }
      ; Check Hwnd -----------------------------------------------------------------------------------------------------
      This.ErrorMsg := ""
      If !(CtrlHwnd := Hwnd + 0)
      Or !DllCall("User32.dll\IsWindow", "UPtr", Hwnd, "UInt") {
         This.ErrorMsg := "Invalid parameter Hwnd: " . Hwnd
         Return False
      }
      If This.Attached.HasKey(Hwnd) {
         This.ErrorMsg := "Control " . Hwnd . " is already registered!"
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
      ; Check colors ---------------------------------------------------------------------------------------------------
      If !This.CheckColors(BkColor, TextColor)
         Return False
      ; Check background color -------------------------------------------------------------------------------------
      If !This.CheckBkColor(BkColor, Classes[1])
         Return False
      ; Check text color -----------------------------------------------------------------------------------------------
      If !This.CheckTextColor(TextColor)
         Return False
      ; Activate message handling on the first call for a class ---------------------------------------------------
      For I, V In Classes {
         If (This.HandledMessages[V] = 0)
            OnMessage(This.WM_CTLCOLOR[V], This.MessageHandler)
         This.HandledMessages[V] += 1
      }
      ; Store values for Hwnd ------------------------------------------------------------------------------------------
      Brush := DllCall("Gdi32.dll\CreateSolidBrush", "UInt", BkColor, "UPtr")
      For I, V In Hwnds
         This.Attached[V] := {Brush: Brush, TextColor: TextColor, BkColor: BkColor, Classes: Classes, Hwnds: Hwnds}
      ; Redraw control -------------------------------------------------------------------------------------------------
      DllCall("User32.dll\InvalidateRect", "Ptr", Hwnd, "Ptr", 0, "Int", 1)
      This.ErrorMsg := ""
      Return True
   }
   ; ======================================================================================================
   ; METHOD Change         Change control colors
   ; Parameters:           Hwnd        - HWND of the GUI control                                   (Integer)
   ;                       BkColor     - HTML color name, 6-digit hex value ("RRGGBB")             (String)
   ;                                     or "" for default color
   ;                       ------------- Optional ----------------------------------------------------------------------
   ;                       TextColor   - HTML color name, 6-digit hex value ("RRGGBB")             (String)
   ;                                     or "" for default color
   ; Return values:        On success  - True
   ;                       On failure  - False, CTLCOLORS.ErrorMsg contains additional informations
   ; Remarks:              If the control isn't registered yet, METHOD Add() is called instead internally.
   ; ======================================================================================================
   Change(Hwnd, BkColor, TextColor = "") {
      ; Check Hwnd -----------------------------------------------------------------------------------------------------
      This.ErrorMsg := ""
      Hwnd += 0
      If !This.Attached.HasKey(Hwnd)
         Return This.Attach(Hwnd, BkColor, TextColor)
      CTL := This.Attached[Hwnd]
      ; Check BkColor --------------------------------------------------------------------------------------------------
      If !This.CheckBkColor(BkColor, CTL.Classes[1])
         Return False
      ; Check TextColor ------------------------------------------------------------------------------------------------
      If !This.CheckTextColor(TextColor)
         Return False
      ; Store Colors ---------------------------------------------------------------------------------------------------
      If (BkColor <> CTL.BkColor) {
         If (CTL.Brush) {
            DllCall("Gdi32.dll\DeleteObject", "Prt", CTL.Brush)
            This.Attached[Hwnd].Brush := 0
         }
         Brush := DllCall("Gdi32.dll\CreateSolidBrush", "UInt", BkColor, "UPtr")
         This.Attached[Hwnd].Brush := Brush
         This.Attached[Hwnd].BkColor := BkColor
      }
      This.Attached[Hwnd].TextColor := TextColor
      This.ErrorMsg := ""
      DllCall("User32.dll\InvalidateRect", "Ptr", Hwnd, "Ptr", 0, "Int", 1)
      Return True
   }
   ; =======================================================
   ; METHOD Detach         Stop control coloring
   ; Parameters:           Hwnd        - HWND of the GUI control                                   (Integer)
   ; Return values:        On success  - True
   ;                       On failure  - False, CTLCOLORS.ErrorMsg contains additional informations
   ; =======================================================
   Detach(Hwnd) {
      This.ErrorMsg := ""
      Hwnd += 0
      If This.Attached.HasKey(Hwnd) {
         CTL := This.Attached[Hwnd].Clone()
         If (CTL.Brush)
            DllCall("Gdi32.dll\DeleteObject", "Prt", CTL.Brush)
         For I, V In CTL.Classes {
            If This.HandledMessages[V] > 0 {
               This.HandledMessages[V] -= 1
               If This.HandledMessages[V] = 0
                  OnMessage(This.WM_CTLCOLOR[V], "")
         }  }
         For I, V In CTL.Hwnds
            This.Attached.Remove(V, "")
         DllCall("User32.dll\InvalidateRect", "Ptr", Hwnd, "Ptr", 0, "Int", 1)
         CTL := ""
         Return True
      }
      This.ErrorMsg := "Control " . Hwnd . " is not registered!"
      Return False
   }
   ; ===========================================================
   ; METHOD Free           Stop coloring for all controls and free resources
   ; Return values:        Always True
   ; ===========================================================
   Free() {
      For K, V In This.Attached
         DllCall("Gdi32.dll\DeleteObject", "Ptr", V.Brush)
      For K, V In This.HandledMessages
         If (V > 0) {
            OnMessage(This.WM_CTLCOLOR[K], "")
            This.HandledMessages[K] := 0
         }
      This.Attached := {}
      Return True
   }
   ; ===========================================================
   ; METHOD IsAttached     Check if the control is registered for coloring
   ; Parameters:           Hwnd        - HWND of the GUI control                                   (Integer)
   ; Return values:        On success  - True
   ;                       On failure  - False
   ; ===========================================================
   IsAttached(Hwnd) {
      Return This.Attached.HasKey(Hwnd)
   }
}
CTLCOLORS_OnMessage(wParam, lParam) {

	; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; PRIVATE Functions ++++++++++++++++++++++++++++++++++++++++++++++++++
	; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	; ============================================================
	; PRIVATE FUNCTION CTLCOLORS_OnMessage
	; This function is destined to handle CTLCOLOR messages. There's no reason to call it manually!
	; ============================================================

   Global CTLCOLORS
   Static SetTextColor := 0, SetBkColor := 0, Counter := 0
   Critical, 50
   If (SetTextColor = 0) {
      HM := DllCall("Kernel32.dll\GetModuleHandle", "Str", "Gdi32.dll", "UPtr")
      SetTextColor := DllCall("Kernel32.dll\GetProcAddress", "Ptr", HM, "AStr", "SetTextColor", "UPtr")
      SetBkColor := DllCall("Kernel32.dll\GetProcAddress", "Ptr", HM, "AStr", "SetBkColor", "UPtr")
   }
   Hwnd := lParam + 0, HDC := wParam + 0
   If CTLCOLORS.IsAttached(Hwnd) {
      CTL := CTLCOLORS.Attached[Hwnd]
      If (CTL.TextColor != "")
         DllCall(SetTextColor, "Ptr", HDC, "UInt", CTL.TextColor)
      DllCall(SetBkColor, "Ptr", HDC, "UInt", CTL.BkColor)
      Return CTL.Brush
   }
}
Class ImageButton {

	; ==============================================================================
	; Namespace:         ImageButton
	; Function:          Create images and assign them to pushbuttons.
	; Tested with:       AHK 1.1.14.03 (A32/U32/U64)
	; Tested on:         Win 7 (x64)
	; Change history:             /2017-02-05/tmplinshi - added DisableFadeEffect(). Thanks to Klark92.
	;                             /2017-01-21/tmplinshi - added support for icon and checkbox/radio buttons
	;                    1.4.00.00/2014-06-07/just me - fixed bug for button caption = "0", "000", etc.
	;                    1.3.00.00/2014-02-28/just me - added support for ARGB colors
	;                    1.2.00.00/2014-02-23/just me - added borders
	;                    1.1.00.00/2013-12-26/just me - added rounded and bicolored buttons
	;                    1.0.00.00/2013-12-21/just me - initial release
	; How to use:
	;     1. Create a push button (e.g. "Gui, Add, Button, vMyButton hwndHwndButton, Caption") using the 'Hwnd' option
	;        to get its HWND.
	;     2. Call ImageButton.Create() passing two parameters:
	;        HWND        -  Button's HWND.
	;        Options*    -  variadic array containing up to 6 option arrays (see below).
	;        ---------------------------------------------------------------------------------------------------------------
	;        The index of each option object determines the corresponding button state on which the bitmap will be shown.
	;        MSDN defines 6 states (http://msdn.microsoft.com/en-us/windows/bb775975):
	;           PBS_NORMAL    = 1
	;	         PBS_HOT       = 2
	;	         PBS_PRESSED   = 3
	;	         PBS_DISABLED  = 4
	;	         PBS_DEFAULTED = 5
	;	         PBS_STYLUSHOT = 6 <- used only on tablet computers (that's false for Windows Vista and 7, see below)
	;        If you don't want the button to be 'animated' on themed GUIs, just pass one option object with index 1.
	;        On Windows Vista and 7 themed bottons are 'animated' using the images of states 5 and 6 after clicked.
	;        ---------------------------------------------------------------------------------------------------------------
	;        Each option array may contain the following values:
	;           Index Value
	;           1     Mode        mandatory:
	;                             0  -  unicolored or bitmap
	;                             1  -  vertical bicolored
	;                             2  -  horizontal bicolored
	;                             3  -  vertical gradient
	;                             4  -  horizontal gradient
	;                             5  -  vertical gradient using StartColor at both borders and TargetColor at the center
	;                             6  -  horizontal gradient using StartColor at both borders and TargetColor at the center
	;                             7  -  'raised' style
	;           2     StartColor  mandatory for Option[1], higher indices will inherit the value of Option[1], if omitted:
	;                             -  ARGB integer value (0xAARRGGBB) or HTML color name ("Red").
	;                             -  Path of an image file or HBITMAP handle for mode 0.
	;           3     TargetColor mandatory for Option[1] if Mode > 0, ignored if Mode = 0. Higher indcices will inherit
	;                             the color of Option[1], if omitted:
	;                             -  ARGB integer value (0xAARRGGBB) or HTML color name ("Red").
	;           4     TextColor   optional, if omitted, the default text color will be used for Option[1], higher indices
	;                             will inherit the color of Option[1]:
	;                             -  ARGB integer value (0xAARRGGBB) or HTML color name ("Red").
	;                                Default: 0xFF000000 (black)
	;           5     Rounded     optional:
	;                             -  Radius of the rounded corners in pixel; the letters 'H' and 'W' may be specified
	;                                also to use the half of the button's height or width respectively.
	;                                Default: 0 - not rounded
	;           6     GuiColor    optional, needed for rounded buttons if you've changed the GUI background color:
	;                             -  RGB integer value (0xRRGGBB) or HTML color name ("Red").
	;                                Default: AHK default GUI background color
	;           7     BorderColor optional, ignored for modes 0 (bitmap) and 7, color of the border:
	;                             -  RGB integer value (0xRRGGBB) or HTML color name ("Red").
	;           8     BorderWidth optional, ignored for modes 0 (bitmap) and 7, width of the border in pixels:
	;                             -  Default: 1
	;        ---------------------------------------------------------------------------------------------------------------
	;        If the the button has a caption it will be drawn above the bitmap.
	; Credits:           THX tic     for GDIP.AHK     : http://www.autohotkey.com/forum/post-198949.html
	;                    THX tkoi    for ILBUTTON.AHK : http://www.autohotkey.com/forum/topic40468.html
	; ==============================================================================
	; This software is provided 'as-is', without any express or implied warranty.
	; In no event will the authors be held liable for any damages arising from the use of this software.
	; ==============================================================================
	; ==============================================================================
	; CLASS ImageButton()
	; ==============================================================================


   ; =================================================================================
   ; PUBLIC PROPERTIES ====================================================================
   ; =================================================================================
   Static DefGuiColor  := ""        ; default GUI color                             (read/write)
   Static DefTxtColor := "Black"    ; default caption color                         (read/write)
   Static LastError := ""           ; will contain the last error message, if any   (readonly)
   ; =================================================================================
   ; PRIVATE PROPERTIES ===================================================================
   ; =================================================================================
   Static BitMaps := []
   Static GDIPDll := 0
   Static GDIPToken := 0
   Static MaxOptions := 8
   ; HTML colors
   Static HTML := {BLACK: 0x000000, GRAY: 0x808080, SILVER: 0xC0C0C0, WHITE: 0xFFFFFF, MAROON: 0x800000
                 , PURPLE: 0x800080, FUCHSIA: 0xFF00FF, RED: 0xFF0000, GREEN: 0x008000, OLIVE: 0x808000
                 , YELLOW: 0xFFFF00, LIME: 0x00FF00, NAVY: 0x000080, TEAL: 0x008080, AQUA: 0x00FFFF, BLUE: 0x0000FF}
   ; Initialize
   Static ClassInit := ImageButton.InitClass()
   ; =================================================================================
   ; PRIVATE METHODS ====================================================================
   ; =================================================================================
   __New(P*) {
      Return False
   }
   ; =================================================================================
   InitClass() {
      ; ----------------------------------------------------------------------------------------------------------------
      ; Get AHK's default GUI background color
      GuiColor := DllCall("User32.dll\GetSysColor", "Int", 15, "UInt") ; COLOR_3DFACE is used by AHK as default
      This.DefGuiColor := ((GuiColor >> 16) & 0xFF) | (GuiColor & 0x00FF00) | ((GuiColor & 0xFF) << 16)
      Return True
   }
   ; =================================================================================
   GdiplusStartup() {
      This.GDIPDll := This.GDIPToken := 0
      If (This.GDIPDll := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "Ptr")) {
         VarSetCapacity(SI, 24, 0)
         Numput(1, SI, 0, "Int")
         If !DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", GDIPToken, "Ptr", &SI, "Ptr", 0)
            This.GDIPToken := GDIPToken
         Else
            This.GdiplusShutdown()
      }
      Return This.GDIPToken
   }
   ; =================================================================================
   GdiplusShutdown() {
      If This.GDIPToken
         DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", This.GDIPToken)
      If This.GDIPDll
         DllCall("Kernel32.dll\FreeLibrary", "Ptr", This.GDIPDll)
      This.GDIPDll := This.GDIPToken := 0
   }
   ; =================================================================================
   FreeBitmaps() {
      For I, HBITMAP In This.BitMaps
         DllCall("Gdi32.dll\DeleteObject", "Ptr", HBITMAP)
      This.BitMaps := []
   }
   ; =================================================================================
   GetARGB(RGB) {
      ARGB := This.HTML.HasKey(RGB) ? This.HTML[RGB] : RGB
      Return (ARGB & 0xFF000000) = 0 ? 0xFF000000 | ARGB : ARGB
   }
   ; =================================================================================
   PathAddRectangle(Path, X, Y, W, H) {
      Return DllCall("Gdiplus.dll\GdipAddPathRectangle", "Ptr", Path, "Float", X, "Float", Y, "Float", W, "Float", H)
   }
   ; =================================================================================
   PathAddRoundedRect(Path, X1, Y1, X2, Y2, R) {
      D := (R * 2), X2 -= D, Y2 -= D
      DllCall("Gdiplus.dll\GdipAddPathArc"
            , "Ptr", Path, "Float", X1, "Float", Y1, "Float", D, "Float", D, "Float", 180, "Float", 90)
      DllCall("Gdiplus.dll\GdipAddPathArc"
            , "Ptr", Path, "Float", X2, "Float", Y1, "Float", D, "Float", D, "Float", 270, "Float", 90)
      DllCall("Gdiplus.dll\GdipAddPathArc"
            , "Ptr", Path, "Float", X2, "Float", Y2, "Float", D, "Float", D, "Float", 0, "Float", 90)
      DllCall("Gdiplus.dll\GdipAddPathArc"
            , "Ptr", Path, "Float", X1, "Float", Y2, "Float", D, "Float", D, "Float", 90, "Float", 90)
      Return DllCall("Gdiplus.dll\GdipClosePathFigure", "Ptr", Path)
   }
   ; =================================================================================
   SetRect(ByRef Rect, X1, Y1, X2, Y2) {
      VarSetCapacity(Rect, 16, 0)
      NumPut(X1, Rect, 0, "Int"), NumPut(Y1, Rect, 4, "Int")
      NumPut(X2, Rect, 8, "Int"), NumPut(Y2, Rect, 12, "Int")
      Return True
   }
   ; =================================================================================
   SetRectF(ByRef Rect, X, Y, W, H) {
      VarSetCapacity(Rect, 16, 0)
      NumPut(X, Rect, 0, "Float"), NumPut(Y, Rect, 4, "Float")
      NumPut(W, Rect, 8, "Float"), NumPut(H, Rect, 12, "Float")
      Return True
   }
   ; =================================================================================
   SetError(Msg) {
      This.FreeBitmaps()
      This.GdiplusShutdown()
      This.LastError := Msg
      Return False
   }
   ; =================================================================================
   ; PUBLIC METHODS =====================================================================
   ; =================================================================================
   Create(HWND, Options*) {
      ; Windows constants
      Static BCM_SETIMAGELIST := 0x1602
           , BS_CHECKBOX := 0x02, BS_RADIOBUTTON := 0x04, BS_GROUPBOX := 0x07, BS_AUTORADIOBUTTON := 0x09
           , BS_LEFT := 0x0100, BS_RIGHT := 0x0200, BS_CENTER := 0x0300, BS_TOP := 0x0400, BS_BOTTOM := 0x0800
           , BS_VCENTER := 0x0C00, BS_BITMAP := 0x0080
           , BUTTON_IMAGELIST_ALIGN_LEFT := 0, BUTTON_IMAGELIST_ALIGN_RIGHT := 1, BUTTON_IMAGELIST_ALIGN_CENTER := 4
           , ILC_COLOR32 := 0x20
           , OBJ_BITMAP := 7
           , RCBUTTONS := BS_CHECKBOX | BS_RADIOBUTTON | BS_AUTORADIOBUTTON
           , SA_LEFT := 0x00, SA_CENTER := 0x01, SA_RIGHT := 0x02
           , WM_GETFONT := 0x31
      ; ----------------------------------------------------------------------------------------------------------------
      This.LastError := ""
      ; ----------------------------------------------------------------------------------------------------------------
      ; Check HWND
      If !DllCall("User32.dll\IsWindow", "Ptr", HWND)
         Return This.SetError("Invalid parameter HWND!")
      ; ----------------------------------------------------------------------------------------------------------------
      ; Check Options
      If !(IsObject(Options)) || (Options.MinIndex() <> 1) ; || (Options.MaxIndex() > This.MaxOptions)
         Return This.SetError("Invalid parameter Options!")
      ; ----------------------------------------------------------------------------------------------------------------
      ; Get and check control's class and styles
      WinGetClass, BtnClass, ahk_id %HWND%
      ControlGet, BtnStyle, Style, , , ahk_id %HWND%
      If (BtnClass != "Button") || ((BtnStyle & 0xF ^ BS_GROUPBOX) = 0)
         Return This.SetError("The control must be a pushbutton!")
      If ((BtnStyle & RCBUTTONS) > 1)
         GuiControl, +0x1000, %HWND% ; BS_PUSHLIKE = 0x1000
      ; ----------------------------------------------------------------------------------------------------------------
      ; Load GdiPlus
      If !This.GdiplusStartup()
         Return This.SetError("GDIPlus could not be started!")
      ; ----------------------------------------------------------------------------------------------------------------
      ; Get the button's font
      GDIPFont := 0
      HFONT := DllCall("User32.dll\SendMessage", "Ptr", HWND, "UInt", WM_GETFONT, "Ptr", 0, "Ptr", 0, "Ptr")
      DC := DllCall("User32.dll\GetDC", "Ptr", HWND, "Ptr")
      DllCall("Gdi32.dll\SelectObject", "Ptr", DC, "Ptr", HFONT)
      DllCall("Gdiplus.dll\GdipCreateFontFromDC", "Ptr", DC, "PtrP", PFONT)
      DllCall("User32.dll\ReleaseDC", "Ptr", HWND, "Ptr", DC)
      If !(PFONT)
         Return This.SetError("Couldn't get button's font!")
      ; ----------------------------------------------------------------------------------------------------------------
      ; Get the button's rectangle
      VarSetCapacity(RECT, 16, 0)
      If !DllCall("User32.dll\GetWindowRect", "Ptr", HWND, "Ptr", &RECT)
         Return This.SetError("Couldn't get button's rectangle!")
      BtnW := NumGet(RECT,  8, "Int") - NumGet(RECT, 0, "Int")
      BtnH := NumGet(RECT, 12, "Int") - NumGet(RECT, 4, "Int")
      ; ----------------------------------------------------------------------------------------------------------------
      ; Get the button's caption
      ControlGetText, BtnCaption, , ahk_id %HWND%
      If (ErrorLevel)
         Return This.SetError("Couldn't get button's caption!")
      ; ----------------------------------------------------------------------------------------------------------------
      ; Create the bitmap(s)
      This.BitMaps := []
      For Index, Option In Options {
         If !IsObject(Option)
            Continue
         BkgColor1 := BkgColor2 := TxtColor := Mode := Rounded := GuiColor := Image := ""
         ; Replace omitted options with the values of Options.1
         Loop, % This.MaxOptions {
            If (Option[A_Index] = "")
               Option[A_Index] := Options.1[A_Index]
         }
         ; -------------------------------------------------------------------------------------------------------------
         ; Check option values
         ; Mode
         Mode := SubStr(Option.1, 1 ,1)
         If !InStr("0123456789", Mode)
            Return This.SetError("Invalid value for Mode in Options[" . Index . "]!")
         ; StartColor & TargetColor
         If (Mode = 0)
         && (FileExist(Option.2) || (DllCall("Gdi32.dll\GetObjectType", "Ptr", Option.2, "UInt") = OBJ_BITMAP))
            Image := Option.2
         Else {
            If !(Option.2 + 0) && !This.HTML.HasKey(Option.2) && !Option.Icon
               Return This.SetError("Invalid value for StartColor in Options[" . Index . "]!")
            BkgColor1 := This.GetARGB(Option.2)
            If (Option.3 = "")
               Option.3 := Option.2
            If !(Option.3 + 0) && !This.HTML.HasKey(Option.3) && !Option.Icon
               Return This.SetError("Invalid value for TargetColor in Options[" . Index . "]!")
            BkgColor2 := This.GetARGB(Option.3)
         }
         ; TextColor
         If (Option.4 = "")
            Option.4 := This.DefTxtColor
         If !(Option.4 + 0) && !This.HTML.HasKey(Option.4)
            Return This.SetError("Invalid value for TxtColor in Options[" . Index . "]!")
         TxtColor := This.GetARGB(Option.4)
         ; Rounded
         Rounded := Option.5
         If (Rounded = "H")
            Rounded := BtnH * 0.5
         If (Rounded = "W")
            Rounded := BtnW * 0.5
         If !(Rounded + 0)
            Rounded := 0
         ; GuiColor
         If (Option.6 = "")
            Option.6 := This.DefGuiColor
         If !(Option.6 + 0) && !This.HTML.HasKey(Option.6)
            Return This.SetError("Invalid value for GuiColor in Options[" . Index . "]!")
         GuiColor := This.GetARGB(Option.6)
         ; BorderColor
         BorderColor := ""
         If (Option.7 <> "") {
            If !(Option.7 + 0) && !This.HTML.HasKey(Option.7)
               Return This.SetError("Invalid value for BorderColor in Options[" . Index . "]!")
            BorderColor := 0xFF000000 | This.GetARGB(Option.7) ; BorderColor must be always opaque
         }
         ; BorderWidth
         BorderWidth := Option.8 ? Option.8 : 1
         ; -------------------------------------------------------------------------------------------------------------
         ; Create a GDI+ bitmap
         DllCall("Gdiplus.dll\GdipCreateBitmapFromScan0", "Int", BtnW, "Int", BtnH, "Int", 0
               , "UInt", 0x26200A, "Ptr", 0, "PtrP", PBITMAP)
         ; Get the pointer to its graphics
         DllCall("Gdiplus.dll\GdipGetImageGraphicsContext", "Ptr", PBITMAP, "PtrP", PGRAPHICS)
         ; Quality settings
         DllCall("Gdiplus.dll\GdipSetSmoothingMode", "Ptr", PGRAPHICS, "UInt", 4)
         DllCall("Gdiplus.dll\GdipSetInterpolationMode", "Ptr", PGRAPHICS, "Int", 7)
         DllCall("Gdiplus.dll\GdipSetCompositingQuality", "Ptr", PGRAPHICS, "UInt", 4)
         DllCall("Gdiplus.dll\GdipSetRenderingOrigin", "Ptr", PGRAPHICS, "Int", 0, "Int", 0)
         DllCall("Gdiplus.dll\GdipSetPixelOffsetMode", "Ptr", PGRAPHICS, "UInt", 4)
         ; Clear the background
         DllCall("Gdiplus.dll\GdipGraphicsClear", "Ptr", PGRAPHICS, "UInt", GuiColor)
         ; Create the image
         If (Image = "") { ; Create a BitMap based on the specified colors
            PathX := PathY := 0, PathW := BtnW, PathH := BtnH
            ; Create a GraphicsPath
            DllCall("Gdiplus.dll\GdipCreatePath", "UInt", 0, "PtrP", PPATH)
            If (Rounded < 1) ; the path is a rectangular rectangle
               This.PathAddRectangle(PPATH, PathX, PathY, PathW, PathH)
            Else ; the path is a rounded rectangle
               This.PathAddRoundedRect(PPATH, PathX, PathY, PathW, PathH, Rounded)
            ; If BorderColor and BorderWidth are specified, 'draw' the border (not for Mode 7)
            If (BorderColor <> "") && (BorderWidth > 0) && (Mode <> 7) {
               ; Create a SolidBrush
               DllCall("Gdiplus.dll\GdipCreateSolidFill", "UInt", BorderColor, "PtrP", PBRUSH)
               ; Fill the path
               DllCall("Gdiplus.dll\GdipFillPath", "Ptr", PGRAPHICS, "Ptr", PBRUSH, "Ptr", PPATH)
               ; Free the brush
               DllCall("Gdiplus.dll\GdipDeleteBrush", "Ptr", PBRUSH)
               ; Reset the path
               DllCall("Gdiplus.dll\GdipResetPath", "Ptr", PPATH)
               ; Add a new 'inner' path
               PathX := PathY := BorderWidth, PathW -= BorderWidth, PathH -= BorderWidth, Rounded -= BorderWidth
               If (Rounded < 1) ; the path is a rectangular rectangle
                  This.PathAddRectangle(PPATH, PathX, PathY, PathW - PathX, PathH - PathY)
               Else ; the path is a rounded rectangle
                  This.PathAddRoundedRect(PPATH, PathX, PathY, PathW, PathH, Rounded)
               ; If a BorderColor has been drawn, BkgColors must be opaque
               BkgColor1 := 0xFF000000 | BkgColor1
               BkgColor2 := 0xFF000000 | BkgColor2
            }
            PathW -= PathX
            PathH -= PathY
            If (Mode = 0) { ; the background is unicolored
               ; Create a SolidBrush
               DllCall("Gdiplus.dll\GdipCreateSolidFill", "UInt", BkgColor1, "PtrP", PBRUSH)
               ; Fill the path
               DllCall("Gdiplus.dll\GdipFillPath", "Ptr", PGRAPHICS, "Ptr", PBRUSH, "Ptr", PPATH)
            }
            Else If (Mode = 1) || (Mode = 2) { ; the background is bicolored
               ; Create a LineGradientBrush
               This.SetRectF(RECTF, PathX, PathY, PathW, PathH)
               DllCall("Gdiplus.dll\GdipCreateLineBrushFromRect", "Ptr", &RECTF
                     , "UInt", BkgColor1, "UInt", BkgColor2, "Int", Mode & 1, "Int", 3, "PtrP", PBRUSH)
               DllCall("Gdiplus.dll\GdipSetLineGammaCorrection", "Ptr", PBRUSH, "Int", 1)
               ; Set up colors and positions
               This.SetRect(COLORS, BkgColor1, BkgColor1, BkgColor2, BkgColor2) ; sorry for function misuse
               This.SetRectF(POSITIONS, 0, 0.5, 0.5, 1) ; sorry for function misuse
               DllCall("Gdiplus.dll\GdipSetLinePresetBlend", "Ptr", PBRUSH
                     , "Ptr", &COLORS, "Ptr", &POSITIONS, "Int", 4)
               ; Fill the path
               DllCall("Gdiplus.dll\GdipFillPath", "Ptr", PGRAPHICS, "Ptr", PBRUSH, "Ptr", PPATH)
            }
            Else If (Mode >= 3) && (Mode <= 6) { ; the background is a gradient
               ; Determine the brush's width/height
               W := Mode = 6 ? PathW / 2 : PathW  ; horizontal
               H := Mode = 5 ? PathH / 2 : PathH  ; vertical
               ; Create a LineGradientBrush
               This.SetRectF(RECTF, PathX, PathY, W, H)
               DllCall("Gdiplus.dll\GdipCreateLineBrushFromRect", "Ptr", &RECTF
                     , "UInt", BkgColor1, "UInt", BkgColor2, "Int", Mode & 1, "Int", 3, "PtrP", PBRUSH)
               DllCall("Gdiplus.dll\GdipSetLineGammaCorrection", "Ptr", PBRUSH, "Int", 1)
               ; Fill the path
               DllCall("Gdiplus.dll\GdipFillPath", "Ptr", PGRAPHICS, "Ptr", PBRUSH, "Ptr", PPATH)
            }
            Else { ; raised mode
               DllCall("Gdiplus.dll\GdipCreatePathGradientFromPath", "Ptr", PPATH, "PtrP", PBRUSH)
               ; Set Gamma Correction
               DllCall("Gdiplus.dll\GdipSetPathGradientGammaCorrection", "Ptr", PBRUSH, "UInt", 1)
               ; Set surround and center colors
               VarSetCapacity(ColorArray, 4, 0)
               NumPut(BkgColor1, ColorArray, 0, "UInt")
               DllCall("Gdiplus.dll\GdipSetPathGradientSurroundColorsWithCount", "Ptr", PBRUSH, "Ptr", &ColorArray
                   , "IntP", 1)
               DllCall("Gdiplus.dll\GdipSetPathGradientCenterColor", "Ptr", PBRUSH, "UInt", BkgColor2)
               ; Set the FocusScales
               FS := (BtnH < BtnW ? BtnH : BtnW) / 3
               XScale := (BtnW - FS) / BtnW
               YScale := (BtnH - FS) / BtnH
               DllCall("Gdiplus.dll\GdipSetPathGradientFocusScales", "Ptr", PBRUSH, "Float", XScale, "Float", YScale)
               ; Fill the path
               DllCall("Gdiplus.dll\GdipFillPath", "Ptr", PGRAPHICS, "Ptr", PBRUSH, "Ptr", PPATH)
            }
            ; Free resources
            DllCall("Gdiplus.dll\GdipDeleteBrush", "Ptr", PBRUSH)
            DllCall("Gdiplus.dll\GdipDeletePath", "Ptr", PPATH)
         } Else { ; Create a bitmap from HBITMAP or file
            If (Image + 0)
               DllCall("Gdiplus.dll\GdipCreateBitmapFromHBITMAP", "Ptr", Image, "Ptr", 0, "PtrP", PBM)
            Else
               DllCall("Gdiplus.dll\GdipCreateBitmapFromFile", "WStr", Image, "PtrP", PBM)
            ; Draw the bitmap
            DllCall("Gdiplus.dll\GdipDrawImageRectI", "Ptr", PGRAPHICS, "Ptr", PBM, "Int", 0, "Int", 0
                  , "Int", BtnW, "Int", BtnH)
            ; Free the bitmap
            DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", PBM)
         }

         if (oIcon := Option.Icon) {
            DllCall("Gdiplus.dll\GdipCreateBitmapFromFile", "WStr", oIcon.file, "PtrP", PBM)

            If !oIcon.w {
               DllCall("Gdiplus.dll\GdipGetImageWidth", "Ptr", PBM, "UInt*", __w)
               oIcon.w := __w
            }
            If !oIcon.h {
               DllCall("Gdiplus.dll\GdipGetImageHeight", "Ptr", PBM, "UInt*", __h)
               oIcon.h := __h
            }

            if !oIcon.HasKey("padding")
               oIcon.padding := 5

            if !oIcon.HasKey("y")
               oIcon.y := (BtnH - oIcon.h)//2

            icon_x := oIcon.HasKey("x") ? oIcon.x : oIcon.padding

            DllCall("Gdiplus.dll\GdipDrawImageRectI", "Ptr", PGRAPHICS, "Ptr", PBM, "Int", icon_x, "Int", oIcon.y
                  , "Int", oIcon.w, "Int", oIcon.h)
            ; Free the bitmap
            DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", PBM)
         }

         ; -------------------------------------------------------------------------------------------------------------
         ; Draw the caption
         If (BtnCaption <> "") {
            ; Create a StringFormat object
            DllCall("Gdiplus.dll\GdipStringFormatGetGenericTypographic", "PtrP", HFORMAT)
            ; Text color
            DllCall("Gdiplus.dll\GdipCreateSolidFill", "UInt", TxtColor, "PtrP", PBRUSH)
            ; Horizontal alignment
            HALIGN := (BtnStyle & BS_CENTER) = BS_CENTER ? SA_CENTER
                    : (BtnStyle & BS_CENTER) = BS_RIGHT  ? SA_RIGHT
                    : (BtnStyle & BS_CENTER) = BS_Left   ? SA_LEFT
                    : SA_CENTER
            DllCall("Gdiplus.dll\GdipSetStringFormatAlign", "Ptr", HFORMAT, "Int", HALIGN)
            ; Vertical alignment
            VALIGN := (BtnStyle & BS_VCENTER) = BS_TOP ? 0
                    : (BtnStyle & BS_VCENTER) = BS_BOTTOM ? 2
                    : 1
            DllCall("Gdiplus.dll\GdipSetStringFormatLineAlign", "Ptr", HFORMAT, "Int", VALIGN)
            ; Set render quality to system default
            DllCall("Gdiplus.dll\GdipSetTextRenderingHint", "Ptr", PGRAPHICS, "Int", 0)
            ; Set the text's rectangle
            VarSetCapacity(RECT, 16, 0)

            If oIcon {
               If !oIcon.x || (HALIGN = SA_CENTER) {
                  NumPut( _left := oIcon.w + oIcon.padding*2, RECT,  0, "Float")
                  NumPut(BtnW - _left                       , RECT,  8, "Float")
                  NumPut(BtnH          , RECT, 12, "Float")
               } Else {
                  NumPut(BtnW, RECT,  8, "Float")
                  NumPut(BtnH, RECT, 12, "Float")
               }
            } Else {
               NumPut(BtnW, RECT,  8, "Float")
               NumPut(BtnH, RECT, 12, "Float")
            }

            ; Draw the text
            DllCall("Gdiplus.dll\GdipDrawString", "Ptr", PGRAPHICS, "WStr", BtnCaption, "Int", -1
                  , "Ptr", PFONT, "Ptr", &RECT, "Ptr", HFORMAT, "Ptr", PBRUSH)
         }
         ; -------------------------------------------------------------------------------------------------------------
         ; Create a HBITMAP handle from the bitmap and add it to the array
         DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", PBITMAP, "PtrP", HBITMAP, "UInt", 0X00FFFFFF)
         This.BitMaps[Index] := HBITMAP
         ; Free resources
         DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", PBITMAP)
         DllCall("Gdiplus.dll\GdipDeleteBrush", "Ptr", PBRUSH)
         DllCall("Gdiplus.dll\GdipDeleteStringFormat", "Ptr", HFORMAT)
         DllCall("Gdiplus.dll\GdipDeleteGraphics", "Ptr", PGRAPHICS)
         ; Add the bitmap to the array
      }
      ; Now free the font object
      DllCall("Gdiplus.dll\GdipDeleteFont", "Ptr", PFONT)
      ; ----------------------------------------------------------------------------------------------------------------
      ; Create the ImageList
      HIL := DllCall("Comctl32.dll\ImageList_Create"
                   , "UInt", BtnW, "UInt", BtnH, "UInt", ILC_COLOR32, "Int", 6, "Int", 0, "Ptr")
      Loop, % (This.BitMaps.MaxIndex() > 1 ? 6 : 1) {
         HBITMAP := This.BitMaps.HasKey(A_Index) ? This.BitMaps[A_Index] : This.BitMaps.1
         DllCall("Comctl32.dll\ImageList_Add", "Ptr", HIL, "Ptr", HBITMAP, "Ptr", 0)
      }
      ; Create a BUTTON_IMAGELIST structure
      VarSetCapacity(BIL, 20 + A_PtrSize, 0)
      NumPut(HIL, BIL, 0, "Ptr")
      Numput(BUTTON_IMAGELIST_ALIGN_CENTER, BIL, A_PtrSize + 16, "UInt")
      ; Hide buttons's caption
      ControlSetText, , , ahk_id %HWND%
      Control, Style, +%BS_BITMAP%, , ahk_id %HWND%
      ; Assign the ImageList to the button
      SendMessage, %BCM_SETIMAGELIST%, 0, 0, , ahk_id %HWND%
      SendMessage, %BCM_SETIMAGELIST%, 0, % &BIL, , ahk_id %HWND%
      ; Free the bitmaps
      This.FreeBitmaps()
      ; ----------------------------------------------------------------------------------------------------------------
      ; All done successfully
      This.GdiplusShutdown()
      Return True
   }
   ; =================================================================================
   ; Set the default GUI color
   SetGuiColor(GuiColor) {
      ; GuiColor     -  RGB integer value (0xRRGGBB) or HTML color name ("Red").
      If !(GuiColor + 0) && !This.HTML.HasKey(GuiColor)
         Return False
      This.DefGuiColor := (This.HTML.HasKey(GuiColor) ? This.HTML[GuiColor] : GuiColor) & 0xFFFFFF
      Return True
   }
   ; =================================================================================
   ; Set the default text color
   SetTxtColor(TxtColor) {
      ; TxtColor     -  RGB integer value (0xRRGGBB) or HTML color name ("Red").
      If !(TxtColor + 0) && !This.HTML.HasKey(TxtColor)
         Return False
      This.DefTxtColor := (This.HTML.HasKey(TxtColor) ? This.HTML[TxtColor] : TxtColor) & 0xFFFFFF
      Return True
   }
   ; =================================================================================
   DisableFadeEffect() {
      ; SPI_GETCLIENTAREAANIMATION = 0x1042
      DllCall("SystemParametersInfo", "UInt", 0x1042, "UInt", 0, "UInt*", isEnabled, "UInt", 0)

      if isEnabled {
         ; SPI_SETCLIENTAREAANIMATION = 0x1043
         DllCall("SystemParametersInfo", "UInt", 0x1043, "UInt", 0, "UInt", 0, "UInt", 0)
         Progress, 10:P100 Hide
         Progress, 10:Off
         DllCall("SystemParametersInfo", "UInt", 0x1043, "UInt", 0, "UInt", 1, "UInt", 0)
      }
   }
}


SciTEOutput(Text:="", Clear=false, LineBreak=true, Exit=false) {

		static SCI_GETLENGTH	:= 2006, SCI_GOTOPOS	:= 2025

		try
			SciObj := ComObjActive("SciTE4AHK.Application")           	;get pointer to active SciTE window
		catch
			return                                                                            	;if not return

		; move Caret to end of output window
		ControlSend, Scintilla2, {LControl Down}{End}{LControl Up} , % "ahk_id " SciObj.SciTEHandle

		If Clear || (StrLen(Text) = 0) {
			SendMessage, SciObj.Message(0x111, 420)                   	;If clear=true -> Clear output window
			return
		}

		SciObj.Output(Text (LineBreak ? "`r`n": ""))                            	;send text to SciTE output pane

		If Exit {
			MsgBox, 36, Exit App?, Exit Application?                         	;If Exit=1 ask if want to exit application
			IfMsgBox,Yes, ExitApp                                                       	;If Msgbox=yes then Exit the appliciation
		}

}
GetHex(hwnd) {
return Format("0x{:X}", hwnd)
}
GetDec(hwnd) {
return Format("{:u}", hwnd)
}
GetWindowSpot(hWnd) {                                                                                                           	;-- like GetWindowInfo, but faster because it only returns position and sizes
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

IFileDialogEvents_new(){
	vtbl := IFileDialogEvents_Vtbl()
	; VarSetCapacity apparently tries to emulate the peculiarities of stack allocation so use GlobalAlloc here
	fde := DllCall("GlobalAlloc", "UInt", 0x0000, "Ptr", A_PtrSize + 4, "Ptr") ; A_PtrSize to store the pointer to the vtable struct + sizeof unsigned int to store this object's refcount
	if (!fde)
		return 0

	NumPut(vtbl, fde+0,, "Ptr") ; place pointer to vtable in beginning of the IFileDialogEvents structure (you know how all this works from the other side)
	NumPut(1, fde+0, A_PtrSize, "UInt") ; Start with a refcount of one (thanks, just me)

	return fde
}
IFileDialogEvents_Vtbl(ByRef vtblSize := 0){
	/* This vtable approach is quite rigid and unflexible in its approach.
		I mean, ideally, you'd want each object to have its own set of methods that are called.
		With this, however, nope - same methods, just a different "this".

		I leave fixing that part up to you. I imagine it involves each object getting its own
		vtable (you could leave all the callback pointers inside the IFileDialogEvents struct after the vtable pointer)
		instead of sharing this one, with the functions to be called for each object determined at creation. Or something like that.
	*/
	static vtable ; This mustn't be freed automatically when it goes out of scope
	if (!VarSetCapacity(vtable)) {
		; Three IUnknown methods that must be implemented, along with the many methods IFileDialogEvents adds on top
		extfuncs := ["QueryInterface", "AddRef", "Release", "OnFileOk", "OnFolderChanging", "OnFolderChange", "OnSelectionChange", "OnShareViolation", "OnTypeChange", "OnOverwrite"]

		; Create IFileDialogEventsVtbl struct
		VarSetCapacity(vtable, extfuncs.Length() * A_PtrSize)

		for i, name in extfuncs
			NumPut(RegisterCallback("IFileDialogEvents_" . name), vtable, (i-1) * A_PtrSize)
	}
	if (IsByRef(vtblSize))
		vtblSize := VarSetCapacity(vtable)
	return &vtable
}
IFileDialogEvents_QueryInterface(this_, riid, ppvObject){ ; Called on a "ComObjQuery"
	static IID_IUnknown, IID_IFileDialogEvents
	if (!VarSetCapacity(IID_IUnknown))
		VarSetCapacity(IID_IUnknown, 16), VarSetCapacity(IID_IFileDialogEvents, 16)
		,DllCall("ole32\CLSIDFromString", "WStr", "{00000000-0000-0000-C000-000000000046}", "Ptr", &IID_IUnknown)
		,DllCall("ole32\CLSIDFromString", "WStr", "{973510db-7d7f-452b-8975-74a85828d354}", "Ptr", &IID_IFileDialogEvents)

	; If someone calls our QI asking for IUnknown or IFileDialogEvents, then respond by:
	if (DllCall("ole32\IsEqualGUID", "Ptr", riid, "Ptr", &IID_IFileDialogEvents) || DllCall("ole32\IsEqualGUID", "Ptr", riid, "Ptr", &IID_IUnknown)) {
		NumPut(this_, ppvObject+0, "Ptr") ; filling in the pointer to a pointer with the address of this object
		IFileDialogEvents_AddRef(this_)
		return 0 ; S_OK
	}

	; Else
	NumPut(0, ppvObject+0, "Ptr") ; no object for the caller
	return 0x80004002 ; E_NOINTERFACE
}
IFileDialogEvents_AddRef(this_){ ; Called on an "ObjAddRef"
	; get and increment our reference count member inside the IFileDialogEvents struct
	NumPut((_refCount := NumGet(this_+0, A_PtrSize, "UInt") + 1), this_+0, A_PtrSize, "UInt")
	return _refCount ; new refcount must be returned
}
IFileDialogEvents_Release(this_) { ; Called on an "ObjRelease"
	_refCount := NumGet(this_+0, A_PtrSize, "UInt") ; read current refcount from IFileDialogEvents struct
	if (_refCount > 0) {
		_refCount -= 1 ; decrease it
		NumPut(_refCount, this_+0, A_PtrSize, "UInt") ; store it
		if (_refCount == 0) ; if it's zero, then
			DllCall("GlobalFree", "Ptr", this_, "Ptr") ; it's time for this object to free itself
	}
	return _refCount ; new refcount must be returned
}
IFileDialogEvents_OnFileOk(this_, pfd){
	return 0x80004001 ; E_NOTIMPL ("[IFileDialogEvents] methods that are not implemented should return E_NOTIMPL.")
}
IFileDialogEvents_OnFolderChanging(this_, pfd, psiFolder){
	return 0x80004001 ; E_NOTIMPL
}
IFileDialogEvents_OnFolderChange(this_, pfd){
	return 0x80004001 ; E_NOTIMPL
}
IFileDialogEvents_OnSelectionChange(this_, pfd){
	if (DllCall(NumGet(NumGet(pfd+0)+14*A_PtrSize), "Ptr", pfd, "Ptr*", psi) >= 0) { ; IFileDialog::GetCurrentSelection
         GetDisplayName := NumGet(NumGet(psi + 0, "UPtr"), A_PtrSize * 5, "UPtr")
         If !DllCall(GetDisplayName, "Ptr", psi, "UInt", 0x80028000, "PtrP", StrPtr) { ; SIGDN_DESKTOPABSOLUTEPARSING
            SelectedFolder := StrGet(StrPtr, "UTF-16"), DllCall("Ole32.dll\CoTaskMemFree", "Ptr", StrPtr)
			ToolTip % SelectedFolder
		 }
		ObjRelease(psi)
	}
	return 0 ; S_OK
}
IFileDialogEvents_OnShareViolation(this_, pfd, psi, pResponse){
	return 0x80004001 ; E_NOTIMPL
}
IFileDialogEvents_OnTypeChange(this_, pfd){
	return 0x80004001 ; E_NOTIMPL
}
IFileDialogEvents_OnOverwrite(this_, pfd, psi, pResponse){
	return 0x80004001 ; E_NOTIMPL
}
SelectFolder(fde:=0, initFolder:="") {

   ; Common Item Dialog -> msdn.microsoft.com/en-us/library/bb776913%28v=vs.85%29.aspx
   ; IFileDialog        -> msdn.microsoft.com/en-us/library/bb775966%28v=vs.85%29.aspx
   ; IShellItem         -> msdn.microsoft.com/en-us/library/bb761140%28v=vs.85%29.aspx

   Static OsVersion 	:= DllCall("GetVersion", "UChar")
   Static Show          	:= A_PtrSize * 3
   Static SetOptions	:= A_PtrSize * 9
   Static GetResult    	:= A_PtrSize * 20
   SelectedFolder    	:= initFolder

   If (OsVersion < 6) { ; IFileDialog requires Win Vista+
      FileSelectFolder, SelectedFolder
      Return SelectedFolder
   }

   If !(FileDialog := ComObjCreate("{DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7}", "{42f85136-db7e-439c-85f1-e4075d135fc8}"))
      Return ""
   VTBL := NumGet(FileDialog + 0, "UPtr")
   DllCall(NumGet(VTBL + SetOptions, "UPtr"), "Ptr", FileDialog, "UInt", 0x00000028, "UInt") ; FOS_NOCHANGEDIR | FOS_PICKFOLDERS

	if (fde) {
		DllCall(NumGet(NumGet(FileDialog+0)+7*A_PtrSize), "Ptr", FileDialog, "Ptr", fde, "UInt*", dwCookie := 0)
	}

	showSucceeded := DllCall(NumGet(VTBL + Show, "UPtr"), "Ptr", FileDialog, "Ptr", 0) >= 0

	if (dwCookie)
		DllCall(NumGet(NumGet(FileDialog+0)+8*A_PtrSize), "Ptr", FileDialog, "UInt", dwCookie)

   If (showSucceeded) {
	   If !DllCall(NumGet(VTBL + GetResult, "UPtr"), "Ptr", FileDialog, "PtrP", ShellItem, "UInt") {
         GetDisplayName := NumGet(NumGet(ShellItem + 0, "UPtr"), A_PtrSize * 5, "UPtr")
         If !DllCall(GetDisplayName, "Ptr", ShellItem, "UInt", 0x80028000, "PtrP", StrPtr) ; SIGDN_DESKTOPABSOLUTEPARSING
            SelectedFolder := StrGet(StrPtr, "UTF-16"), DllCall("Ole32.dll\CoTaskMemFree", "Ptr", StrPtr)
         ObjRelease(ShellItem)
      }
   }

   ObjRelease(FileDialog)

Return SelectedFolder
}



;#include



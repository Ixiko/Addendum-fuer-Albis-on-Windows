/*
  version: 2022.11.05 - Options.CoordMode wurde nicht übergeben

  If you want to add your own style to the built-in style, you can add it directly in btt().
  Advantage:
  *High performance is 2-1000 times the performance of the built-in ToolTip (the larger the text, the greater the performance comparison, 2-5 times the performance in most common application scenarios)
  *High compatibility Compatible with everything built-in ToolTip, including syntax, WhichToolTip, A_CoordModeToolTip, automatic line wrapping, etc.
  *Easy to use, one line of code is ready to use, no need to manually create and release resources
  *Multiple sets of built-in styles can quickly realize custom styles through templates
  *Customizable style Rounded corner size Thin border (color, size) Margin size Text (font, font size, style, color, rendering method) Background color (gradient colors available, support horizontal and vertical)
  *Customizable parameters Follow the mouse distance The text box will never be blocked Attach the target handle Coordinate mode
  *No flicker, no jump
  *Multi-monitor support
  *Zoom display support
  *Follow the mouse without out of bounds
  *Can attach designated targets
  todo:
  ANSI version support.
  Draw shadows
  Text perfectly centered
  There are too many texts to give obvious prompts (such as flashing) when the display is not complete

*/
btt(Text:="", X:="", Y:="", WhichToolTip:="", BuiltInStyleOrStyles:="", BuiltInOptionOrOptions:="") {

  static 	BTT
       , 	Style1 	:= {TextColor:0xffeef8f6
                   , BackgroundColor:0xff1b8dff
                   , FontSize:14}
       , 	Style2 	:= {Border:1
                   , Rounded:8
                   , TextColor:0xfff4f4f4
                   , BackgroundColor:0xaa3e3d45
                   , FontSize:14}
       , 	Style3 	:= {Border:2
                   , Rounded:0
                   , TextColor:0xffF15839
                   , BackgroundColor:0xffFCEDE6
                   , FontSize:14}
       , 	Style4 	:= {Border:5
                   , Rounded:15
                   , BorderColor:0xff503a68
                   , TextColor:0xffF3AE00
                   , BackgroundColorLinearGradientStart:0xff9A83AF
                   , BackgroundColorLinearGradientEnd:0xff7A638F
                   , BackgroundColorLinearGradientDirection:1
                   , FontSize:14
				   , FontRender:5
                   , FontStyle:"Bold Italic"}
       , 	Style5 	:= {Border:0
                   , Rounded:5
                   , TextColor:0xffeeeeee
				   , FontRender:5
                   , BackgroundColorLinearGradientStart:0xff134E5E
                   , BackgroundColorLinearGradientEnd:0xff326f69
                   , BackgroundColorLinearGradientDirection:1}
       , 	Style98:=  {Border:5
                    , Rounded:15
                    , Margin:10
                    , BorderColor:0xffaabbcc
                    , TextColor:0xff112233
                    , BackgroundColor:0xff778899
                    , BackgroundColorLinearGradientStart:0xffF4CFC9
                    , BackgroundColorLinearGradientEnd:0xff8DA5D3
                    , BackgroundColorLinearGradientDirection:1
					, BackgroundColorLinearGradientAngle:135
                    , Font:"Consolas"
                    , FontSize:14
                    , FontRender:5
                    , FontStyle:"Normal"}
       , 	Style99:=  {Border:20
                    , Rounded:30
                    , Margin:30
                    , BorderColor:0xffaabbcc
                    , TextColor:0xff112233
                    , BackgroundColor:0xff778899
                    , BackgroundColorLinearGradientStart:0xffF4CFC9
                    , BackgroundColorLinearGradientEnd:0xff8DA5D3
                    , BackgroundColorLinearGradientDirection:1
                    , Font:"Font Name"
                    , FontSize:12
                    , FontRender:5
                    , FontStyle:"Regular Bold Italic BoldItalic Underline Strikeout"}
       , 	Option99 := {TargetHWND:""
                    , CoordMode:"Screen"
                    , MouseNeverCoverToolTip:""
                    , DistanceBetweenMouseXAndToolTip:""
                    , DistanceBetweenMouseYAndToolTip:""}

  if (BTT="")
    BTT:= new BeautifulToolTip()

  BTT.ToolTip(Text, X, Y, WhichToolTip
	            , 	%BuiltInStyleOrStyles%=""       	? BuiltInStyleOrStyles      	: %BuiltInStyleOrStyles%
	            , 	%BuiltInOptionOrOptions%=""	? BuiltInOptionOrOptions 	: %BuiltInOptionOrOptions%)
}

Class BeautifulToolTip{

  static MouseNeverCoverToolTip:=1
		, DistanceBetweenMouseXAndToolTip:=16
		, DistanceBetweenMouseYAndToolTip:=16
		, DebugMode:=0

  __New()  {

    if (!this.pToken)    {

      SavedBatchLines := A_BatchLines
      SetBatchLines, -1
      this.pToken := Gdip_Startup()
      if (!this.pToken)      {
        MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
        ExitApp
      }

      this.Monitors := MDMF_Enum()
      SysGet, VirtualWidth, 78
      SysGet, VirtualHeight, 79
      this.DIBWidth  := VirtualWidth
      this.DIBHeight := VirtualHeight
      this.ToolTipFontName := Fnt_GetTooltipFontName()

      loop 20      {

        Gui, _BTT%A_Index%: +E0x80000 -Caption +ToolWindow +LastFound +AlwaysOnTop +Hwnd_hBTT%A_Index%
        Gui, _BTT%A_Index%: Show, NA
        this["hBTT" A_Index] 	:= _hBTT%A_Index%
        this["hbm" A_Index]  	:= CreateDIBSection(this.DIBWidth, this.DIBHeight)
        this["hdc" A_Index]   	:= CreateCompatibleDC()
        this["obm" A_Index]  	:= SelectObject(this["hdc" A_Index], this["hbm" A_Index])
        this["G" A_Index]      	:= Gdip_GraphicsFromHDC(this["hdc" A_Index])
        Gdip_SetSmoothingMode(this["G" A_Index], 4)
        Gdip_SetPixelOffsetMode(this["G" A_Index], 2)

      }

      SetBatchLines, % SavedBatchLines

    }
    else
      return

  }

  __Delete()  {

		loop, 20    {
			Gdip_DeleteGraphics(this["G" A_Index])
		  , SelectObject(this["hdc" A_Index], this["obm" A_Index])
		  , DeleteObject(this["hbm" A_Index])
		  , DeleteDC(this["hdc" A_Index])
		}
		Gdip_Shutdown(this.pToken)

  }

  ToolTip(Text:="", X:="", Y:="", WhichToolTip:="", Styles:="", Options:="")  {

    WhichToolTip:=(WhichToolTip="") ? 1 : Range(WhichToolTip, 1, 20)
    O := this._CheckStylesAndOptions(Styles, Options)
    FirstCallOrNeedToUpdate := (Text != this["SavedText" WhichToolTip]
                           || O.Checksum != this["SavedOptions" WhichToolTip])



    if (Text="")    {
      Gdip_GraphicsClear(this["G" WhichToolTip])
      UpdateLayeredWindow(this["hBTT" WhichToolTip], this["hdc" WhichToolTip])
        this["SavedText" WhichToolTip]       	:= ""
      , this["SavedOptions" WhichToolTip]    	:= ""
      , this["SavedX" WhichToolTip]           	:= ""
      , this["SavedY" WhichToolTip]           	:= ""
      , this["SavedW" WhichToolTip]          	:= ""
      , this["SavedH" WhichToolTip]           	:= ""
      , this["SavedCoordMode" WhichToolTip]  := ""
      , this["SavedTargetHWND" WhichToolTip] := ""
    }
    else if (FirstCallOrNeedToUpdate) {

      SavedBatchLines:=A_BatchLines
      SetBatchLines, -1

      TargetSize := this._CalculateDisplayPosition(X, Y, "", "", O, GetTargetSize := 1)

      MaxTextWidth := TargetSize.W - O.Margin*2 - O.Border*2
      MaxTextHeight := (TargetSize.H*90)//100 - O.Margin*2 - O.Border*2
      O.Width      	:= MaxTextWidth, O.Height := MaxTextHeight
      TextArea     	:= StrSplit(Gdip_TextToGraphics2(this["G" WhichToolTip], Text, O, Measure := 1), "|")
      TextWidth      	:= Min(Ceil(TextArea[3]), MaxTextWidth)
      TextHeight     	:= Min(Ceil(TextArea[4]), MaxTextHeight)
      RectWidth      	:= TextWidth+O.Margin*2
      RectHeight     	:= TextHeight+O.Margin*2
      RectWithBorderWidth := RectWidth+O.Border*2
      RectWithBorderHeight := RectHeight+O.Border*2
      R := (O.Rounded>Min(RectWidth, RectHeight)//2) ? Min(RectWidth, RectHeight)//2 : O.Rounded
      Gdip_GraphicsClear(this["G" WhichToolTip])
      pBrushBorder := Gdip_BrushCreateSolid(O.BorderColor)

      if (O.BGCLGD and O.BGCLGS and O.BGCLGE)  {
        Left:=O.Border, Top:=O.Border, Right:=Left+RectWidth, Bottom:=Top+RectHeight
        switch, O.BGCLGD  {
          case, "1": pBrushBackground:=Gdip_CreateLinearGrBrush(Left, Top, Right, Top,    O.BGCLGS, O.BGCLGE)
          case, "2": pBrushBackground:=Gdip_CreateLinearGrBrush(Left, Top, Right, Bottom, O.BGCLGS, O.BGCLGE)
          case, "3": pBrushBackground:=Gdip_CreateLinearGrBrush(Left, Top, Left,  Bottom, O.BGCLGS, O.BGCLGE)
        }
      }
      else
        pBrushBackground := Gdip_BrushCreateSolid(O.BackgroundColor)

      if (O.Border>0)
        switch, R        {
          case, "0": Gdip_FillRectangle(this["G" WhichToolTip]
          , pBrushBorder, 0, 0, RectWithBorderWidth, RectWithBorderHeight)
          Default  : Gdip_FillRoundedRectanglePath(this["G" WhichToolTip]
          , pBrushBorder, 0, 0, RectWithBorderWidth, RectWithBorderHeight, R)
        }
		switch, R      {
        case, "0": Gdip_FillRectangle(this["G" WhichToolTip]
        , pBrushBackground, O.Border, O.Border, RectWidth, RectHeight)
        Default  : Gdip_FillRoundedRectanglePath(this["G" WhichToolTip]
        , pBrushBackground, O.Border, O.Border, RectWidth, RectHeight
        , (R>O.Border) ? R-O.Border : R)
      }

	  Gdip_DeleteBrush(pBrushBorder)
      Gdip_DeleteBrush(pBrushBackground)
      O.X:=O.Border+O.Margin
	  O.Y:=O.Border+O.Margin
	  O.Width:=TextWidth
	  O.Height:=TextHeight

      TempText := (TextArea[5]<StrLen(Text)) ? TextArea[5]>4 ? SubStr(Text, 1 ,TextArea[5]-4) "…………" : SubStr(Text, 1 ,1) "…………" : Text

	  Gdip_TextToGraphics2(this["G" WhichToolTip], TempText, O)

	 if (this.DebugMode)      {
        pBrush := Gdip_BrushCreateSolid(0x20ff0000)
        Gdip_FillRectangle(this["G" WhichToolTip], pBrush, O.Border+O.Margin, O.Border+O.Margin, TextWidth, TextHeight)
        Gdip_DeleteBrush(pBrush)
      }

      this._CalculateDisplayPosition(X, Y, RectWithBorderWidth, RectWithBorderHeight, O)
      UpdateLayeredWindow(this["hBTT" WhichToolTip], this["hdc" WhichToolTip], X, Y, RectWithBorderWidth, RectWithBorderHeight)
	  this["SavedText" WhichToolTip]                	:= Text
      this["SavedOptions" WhichToolTip]          	:= O.Checksum
      this["SavedX" WhichToolTip]                   	:= X
      this["SavedY" WhichToolTip]                    	:= Y
      this["SavedW" WhichToolTip]                  	:= RectWithBorderWidth
      this["SavedH" WhichToolTip]                   	:= RectWithBorderHeight
      this["SavedCoordMode" WhichToolTip]  	:= O.CoordMode
      this["SavedTargetHWND" WhichToolTip]	:= O.TargetHWND

      SetBatchLines, %SavedBatchLines%

    }
    else if ((X="" || Y="") || O.CoordMode!="Screen" || O.CoordMode!=this.SavedCoordMode || O.TargetHWND!=this.SavedTargetHWND)    {

      this._CalculateDisplayPosition(X, Y, this["SavedW" WhichToolTip], this["SavedH" WhichToolTip], O)
      if (X!=this["SavedX" WhichToolTip] || Y!=this["SavedY" WhichToolTip])      {
        UpdateLayeredWindow(this["hBTT" WhichToolTip], this["hdc" WhichToolTip], X, Y, this["SavedW" WhichToolTip], this["SavedH" WhichToolTip])
          this["SavedX" WhichToolTip]                    	:= X
        , this["SavedY" WhichToolTip]                  	:= Y
        , this["SavedCoordMode" WhichToolTip]  	:= O.CoordMode
        , this["SavedTargetHWND" WhichToolTip] 	:= O.TargetHWND
      }
    }

  }

  _CheckStylesAndOptions(Styles, Options)  {

      O                 := IsObject(Styles)     ? Styles.Clone()       : Object()
    , O.Border          := O.Border=""          ? 1                    : Range(O.Border, 0, 20)
    , O.Rounded         := O.Rounded=""         ? 3                    : Range(O.Rounded, 0, 30)
    , O.Margin          := O.Margin=""          ? 5                    : Range(O.Margin, 0, 30)
    , O.TextColor       := O.TextColor=""       ? 0xff575757           : O.TextColor
    , O.BackgroundColor := O.BackgroundColor="" ? 0xffffffff           : O.BackgroundColor
    , O.Font            := O.Font=""            ? this.ToolTipFontName : O.Font
    , O.FontSize        := O.FontSize=""        ? 12                   : O.FontSize
    , O.FontRender      := O.FontRender=""      ? 5                    : Range(O.FontRender, 0, 5)
    , BlendedColor      := ((O.BackgroundColor>>24)<<24) + (O.TextColor&0xffffff)
    , O.BorderColor     := O.BorderColor=""      ? BlendedColor         : O.BorderColor
    , O.BGCLGS          := O.BackgroundColorLinearGradientStart
    , O.BGCLGE          := O.BackgroundColorLinearGradientEnd
    , O.BGCLGD          := O.BackgroundColorLinearGradientDirection
    , O.BGCLGD          := O.BGCLGD=""           ? ""                   : Range(O.BGCLGD, 1, 3)
    , O.FontStyle       := O.FontStyle
    , O.TargetHWND      := Options.TargetHWND="" ? WinExist("A")        : Options.TargetHWND+0
    , O.CoordMode       := Options.CoordMode=""  ? A_CoordModeToolTip   : Options.CoordMode
    , temp1:=Options.MouseNeverCoverToolTip
    , temp2:=Options.DistanceBetweenMouseXAndToolTip
    , temp3:=Options.DistanceBetweenMouseYAndToolTip
    , this.MouseNeverCoverToolTip          := temp1="" ? this.MouseNeverCoverToolTip          : temp1
    , this.DistanceBetweenMouseXAndToolTip := temp2="" ? this.DistanceBetweenMouseXAndToolTip : temp2
    , this.DistanceBetweenMouseYAndToolTip := temp3="" ? this.DistanceBetweenMouseYAndToolTip : temp3
    , O.Checksum        := Format("{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13}"
                                , O.Border, O.Rounded, O.Margin, O.BorderColor, O.TextColor
                                , O.BackgroundColor, O.BGCLGS, O.BGCLGE, O.BGCLGD
                                , O.Font, O.FontSize, O.FontRender, O.FontStyle)
    return O
  }

  _CalculateDisplayPosition(ByRef X, ByRef Y, W, H, Options, GetTargetSize:=0)  {

    VarSetCapacity(Point, 8, 0)
	DllCall("GetCursorPos", "Ptr", &Point)
    MouseX := NumGet(Point, 0, "Int")
	MouseY := NumGet(Point, 4, "Int")

    if (X="" && Y="")    {
        DisplayX     := MouseX
      , DisplayY     := MouseY
      , hMonitor     := MDMF_FromPoint(DisplayX, DisplayY, MONITOR_DEFAULTTONEAREST:=2)
      , TargetLeft   := this.Monitors[hMonitor].Left
      , TargetTop    := this.Monitors[hMonitor].Top
      , TargetRight  := this.Monitors[hMonitor].Right
      , TargetBottom := this.Monitors[hMonitor].Bottom
      , TargetWidth  := TargetRight-TargetLeft
      , TargetHeight := TargetBottom-TargetTop
    }
    else if (Options.CoordMode  = "Window" || Options.CoordMode  = "Relative")    {
        WinGetPos, WinX, WinY, WinW, WinH, % "ahk_id " Options.TargetHWND
        XInScreen    := WinX+X
        YInScreen    := WinY+Y
        TargetLeft   := WinX
        TargetTop    := WinY
        TargetWidth  := WinW
        TargetHeight := WinH
        TargetRight  := TargetLeft+TargetWidth
        TargetBottom := TargetTop+TargetHeight
        DisplayX     := (X="") ? MouseX : XInScreen
        DisplayY     := (Y="") ? MouseY : YInScreen
    }
    else if (Options.CoordMode  = "Client")    {
        VarSetCapacity(ClientArea, 16, 0)
        DllCall("GetClientRect", "Ptr", Options.TargetHWND, "Ptr", &ClientArea)
        DllCall("ClientToScreen", "Ptr", Options.TargetHWND, "Ptr", &ClientArea)
        ClientX      := NumGet(ClientArea, 0, "Int")
        ClientY      := NumGet(ClientArea, 4, "Int")
        ClientW      := NumGet(ClientArea, 8, "Int")
        ClientH      := NumGet(ClientArea, 12, "Int")
        XInScreen    := ClientX+X
        YInScreen    := ClientY+Y
        TargetLeft   := ClientX
        TargetTop    := ClientY
        TargetWidth  := ClientW
        TargetHeight := ClientH
        TargetRight  := TargetLeft+TargetWidth
        TargetBottom := TargetTop+TargetHeight
        DisplayX     := (X="") ? MouseX : XInScreen
        DisplayY     := (Y="") ? MouseY : YInScreen
    }
    else if (Options.CoordMode  = "Screen") {
    ;~ else {
        DisplayX    	:= (X="") ? MouseX : X
        DisplayY     	:= (Y="") ? MouseY : Y
        hMonitor     	:= MDMF_FromPoint(DisplayX, DisplayY, MONITOR_DEFAULTTONEAREST:=2)
        TargetLeft   	:= this.Monitors[hMonitor].Left
        TargetTop    	:= this.Monitors[hMonitor].Top
        TargetRight  	:= this.Monitors[hMonitor].Right
        TargetBottom:= this.Monitors[hMonitor].Bottom
        TargetWidth  := TargetRight-TargetLeft
        TargetHeight := TargetBottom-TargetTop
    }


    if (GetTargetSize=1)    {
        TargetSize   := []
        TargetSize.X := TargetLeft
        TargetSize.Y := TargetTop
        TargetSize.W := Min(TargetWidth, this.DIBWidth)
        TargetSize.H := Min(TargetHeight, this.DIBHeight)
      return TargetSize
    }

      DPIScale := A_ScreenDPI/96

      DisplayX := (X="") ? DisplayX+this.DistanceBetweenMouseXAndToolTip*DPIScale : DisplayX
      DisplayY := (Y="") ? DisplayY+this.DistanceBetweenMouseYAndToolTip*DPIScale : DisplayY
      DisplayX := (DisplayX+W>=TargetRight)  ? TargetRight-W  : DisplayX
      DisplayY := (DisplayY+H>=TargetBottom) ? TargetBottom-H : DisplayY
      DisplayX := (DisplayX<TargetLeft) ? TargetLeft : DisplayX
      DisplayY := (DisplayY<TargetTop)  ? TargetTop  : DisplayY

    if  (this.MouseNeverCoverToolTip=1 && (X="" || Y="") && MouseX>=DisplayX && MouseY>=DisplayY && MouseX<=DisplayX+W && MouseY<=DisplayY+H)    {
      DisplayY := MouseY-H-16>=TargetTop ? MouseY-H-16 : MouseY+H+16<=TargetBottom ? MouseY+16 : DisplayY
    }

    X := DisplayX , Y := DisplayY
  }

}

NonNull(Value1, Value2){
  return, Value1="" ? Value2 : Value1
}

Range(Value, MinValue, MaxValue){
  return, Max(Min(Value, MaxValue), MinValue)
}

Gdip_TextToGraphics2(pGraphics, Text, Options, Measure:=0){
	static Styles := "Regular|Bold|Italic|BoldItalic|Underline|Strikeout"
	Style := 0
	For eachStyle, valStyle in StrSplit(Styles, "|")	{
		if InStr(Options.FontStyle, valStyle)
			Style |= (valStyle != "StrikeOut") ? (A_Index-1) : 8
	}
	If !(hFontFamily := Gdip_FontFamilyCreate(Options.Font))
		hFontFamily := Gdip_FontFamilyCreateGeneric(1)
	hFont := Gdip_FontCreate(hFontFamily, Options.FontSize, Style, Unit:=0)
	hStringFormat := Gdip_StringFormatCreate(0x00002020)
	if !hStringFormat
		hStringFormat := Gdip_StringFormatGetGeneric(0)
	pBrush := Gdip_BrushCreateSolid(Options.TextColor)
	if !(hFontFamily && hFont && hStringFormat && pBrush && pGraphics)	{
		E := !pGraphics ? -2 : !hFontFamily ? -3 : !hFont ? -4 : !hStringFormat ? -5 : !pBrush ? -6 : 0
		if pBrush
			Gdip_DeleteBrush(pBrush)
		if hStringFormat
			Gdip_DeleteStringFormat(hStringFormat)
		if hFont
			Gdip_DeleteFont(hFont)
		if hFontFamily
			Gdip_DeleteFontFamily(hFontFamily)
		return E
	}
	Gdip_SetStringFormatAlign(hStringFormat, Align:=0)
	Gdip_SetTextRenderingHint(pGraphics, Options.FontRender)
	CreateRectF(RC
            , Options.X="" ? 0 : Options.X
            , Options.Y="" ? 0 : Options.Y
            , Options.Width, Options.Height)
	returnRC := Gdip_MeasureString(pGraphics, Text, hFont, hStringFormat, RC)
	if !Measure
		_E := Gdip_DrawString(pGraphics, Text, hFont, hStringFormat, pBrush, RC)
	Gdip_DeleteBrush(pBrush)
	Gdip_DeleteFont(hFont)
	Gdip_DeleteStringFormat(hStringFormat)
	Gdip_DeleteFontFamily(hFontFamily)
	return _E ? _E : returnRC
}

Fnt_GetTooltipFontName(){
	static LF_FACESIZE:=32
	return StrGet(Fnt_GetNonClientMetrics()+(A_IsUnicode ? 316:220)+28,LF_FACESIZE)
}
Fnt_GetNonClientMetrics(){
	static Dummy15105062,SPI_GETNONCLIENTMETRICS:=0x29,NONCLIENTMETRICS
	cbSize:=A_IsUnicode ? 500:340
	if (((GV:=DllCall("GetVersion"))&0xFF . "." . GV>>8&0xFF)>=6.0)
		cbSize+=4
	VarSetCapacity(NONCLIENTMETRICS,cbSize,0)
	NumPut(cbSize,NONCLIENTMETRICS,0,"UInt")
	if !DllCall("SystemParametersInfo", "UInt",SPI_GETNONCLIENTMETRICS, "UInt",cbSize, "Ptr",&NONCLIENTMETRICS, "UInt",0)
		return false
	return &NONCLIENTMETRICS
}

MDMF_Enum(HMON := "") {
   Static CallbackFunc := Func(A_AhkVersion < "2" ? "RegisterCallback" : "CallbackCreate")
   Static EnumProc := CallbackFunc.Call("MDMF_EnumProc")
   Static Obj := (A_AhkVersion < "2") ? "Object" : "Map"
   Static Monitors := {}
   If (HMON = "")   {
      Monitors := %Obj%("TotalCount", 0)
      If !DllCall("User32.dll\EnumDisplayMonitors", "Ptr", 0, "Ptr", 0, "Ptr", EnumProc, "Ptr", &Monitors, "Int")
         Return False
   }
   Return (HMON = "") ? Monitors : Monitors.HasKey(HMON) ? Monitors[HMON] : False
}
MDMF_EnumProc(HMON, HDC, PRECT, ObjectAddr) {
   Monitors := Object(ObjectAddr)
   Monitors[HMON] := MDMF_GetInfo(HMON)
   Monitors["TotalCount"]++
   If (Monitors[HMON].Primary)
      Monitors["Primary"] := HMON
   Return True
}
MDMF_FromHWND(HWND, Flag := 0) {
   Return DllCall("User32.dll\MonitorFromWindow", "Ptr", HWND, "UInt", Flag, "Ptr")
}
MDMF_FromPoint(ByRef X := "", ByRef Y := "", Flag := 0) {
   If (X = "") || (Y = "") {
      VarSetCapacity(PT, 8, 0)
      DllCall("User32.dll\GetCursorPos", "Ptr", &PT, "Int")
      If (X = "")
         X := NumGet(PT, 0, "Int")
      If (Y = "")
         Y := NumGet(PT, 4, "Int")
   }
   Return DllCall("User32.dll\MonitorFromPoint", "Int64", (X & 0xFFFFFFFF) | (Y << 32), "UInt", Flag, "Ptr")
}
MDMF_FromRect(X, Y, W, H, Flag := 0) {
   VarSetCapacity(RC, 16, 0)
   NumPut(X, RC, 0, "Int"), NumPut(Y, RC, 4, "Int"), NumPut(X + W, RC, 8, "Int"), NumPut(Y + H, RC, 12, "Int")
   Return DllCall("User32.dll\MonitorFromRect", "Ptr", &RC, "UInt", Flag, "Ptr")
}
MDMF_GetInfo(HMON) {
   NumPut(VarSetCapacity(MIEX, 40 + (32 << !!A_IsUnicode)), MIEX, 0, "UInt")
   If DllCall("User32.dll\GetMonitorInfo", "Ptr", HMON, "Ptr", &MIEX, "Int")
      Return {Name:      (Name := StrGet(&MIEX + 40, 32))
            , Num:       RegExReplace(Name, ".*(\d+)$", "$1")
            , Left:      NumGet(MIEX, 4, "Int")
            , Top:       NumGet(MIEX, 8, "Int")
            , Right:     NumGet(MIEX, 12, "Int")
            , Bottom:    NumGet(MIEX, 16, "Int")
            , WALeft:    NumGet(MIEX, 20, "Int")
            , WATop:     NumGet(MIEX, 24, "Int")
            , WARight:   NumGet(MIEX, 28, "Int")
            , WABottom:  NumGet(MIEX, 32, "Int")
            , Primary:   NumGet(MIEX, 36, "UInt")}
   Return False
}
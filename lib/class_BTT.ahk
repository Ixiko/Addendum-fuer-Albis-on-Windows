/*
  version: 2021.02.20
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
btt(Text:="", X:="", Y:="", WhichToolTip:="", BulitInStyleOrStyles:="", BulitInOptionOrOptions:="") {
  static BTT
       , Style1 := {TextColor:0xffeef8f6
                   , BackgroundColor:0xff1b8dff
                   , FontSize:14}

       , Style2 := {Border:1
                   , Rounded:8
                   , TextColor:0xfff4f4f4
                   , BackgroundColor:0xaa3e3d45
                   , FontSize:14}

       , Style3  := {Border:2
                   , Rounded:0
                   , TextColor:0xffF15839
                   , BackgroundColor:0xffFCEDE6
                   , FontSize:14}

       , Style4  := {Border:5
                   , Rounded:15
                   , BorderColor:0xff503a68
                   , TextColor:0xffF3AE00
                   , BackgroundColorLinearGradientStart:0xff9A83AF
                   , BackgroundColorLinearGradientEnd:0xff7A638F
                   , BackgroundColorLinearGradientDirection:1
                   ;, BackgroundColor:0xff7A638F
                   , FontSize:16
				   , FontRender:5
                   , FontStyle:"Bold Italic"}

       , Style5  := {Border:0
                   , Rounded:5
                   , TextColor:0xffeeeeee
				   , FontRender:5
                   , BackgroundColorLinearGradientStart:0xff134E5E
                   , BackgroundColorLinearGradientEnd:0xff326f69
                   , BackgroundColorLinearGradientDirection:1}

       ; You can customize your own style.
       ; All supported parameters are listed below. All parameters can be omitted.
       ; Please share your custom style and include a screenshot. It will help a lot of people.
       ; Attention:
       ; Color => ARGB => Alpha Red Green Blue => 0x ff aa bb cc => 0xffaabbcc
       , Style99 :=  {Border:20
                    , Rounded:30
                    , Margin:30
                    , BorderColor:0xffaabbcc                         ; ARGB
                    , TextColor:0xff112233                           ; ARGB
                    , BackgroundColor:0xff778899                     ; ARGB
                    , BackgroundColorLinearGradientStart:0xffF4CFC9  ; ARGB
                    , BackgroundColorLinearGradientEnd:0xff8DA5D3    ; ARGB
                    , BackgroundColorLinearGradientDirection:1       ; 1 = Horizontal   2 = Oblique   3 = Vertical
                    , Font:"Font Name"                               ; If omitted, ToolTip's Font will be used.
                    , FontSize:12
                    , FontRender:5                                   ; 0-5 (recommended value is 5)
                    , FontStyle:"Regular Bold Italic BoldItalic Underline Strikeout"}

        , Option99 := {TargetHWND:""                                 ; If omitted, active window will be used.
                    , CoordMode:"Screen"                             ; If omitted, A_CoordModeToolTip will be used.
                    , MouseNeverCoverToolTip:""                      ; If omitted, 1 will be used.
                    , DistanceBetweenMouseXAndToolTip:""             ; If omitted, 16 will be used. This value can be negative.
                    , DistanceBetweenMouseYAndToolTip:""}            ; If omitted, 16 will be used. This value can be negative.

  ; Initializing BTT directly in static will report an error, so you can only write like this
  if (BTT="")
    BTT:= new BeautifulToolTip()

  BTT.ToolTip(Text, X, Y, WhichToolTip
            ; If Style is the name of a built-in preset, use the value of the corresponding built-in preset, otherwise use the value of Styles itself. Options are the same.
            , %BulitInStyleOrStyles%=""       	? BulitInStyleOrStyles      	: %BulitInStyleOrStyles%
            , %BulitInOptionOrOptions%=""	? BulitInOptionOrOptions 	: %BulitInOptionOrOptions%)
}

Class BeautifulToolTip{
  ; The following are static variables in the class. The number 1 at the end indicates that there are at most 1-20 such variants.
  ; For example, _BTT1 has 20 similar variables _BTT1 ... _BTT20.
  ; pToken, Monitors, ToolTipFontName, DIBWidth, DIBHeight
  ; hBTT1（GUI句柄）, hbm1, hdc1, obm1, G1
  ; SavedText1, SavedOptions1, SavedX1, SavedY1, SavedW1, SavedH1, SavedCoordMode1, SavedTargetHWND1
  static MouseNeverCoverToolTip:=1, DistanceBetweenMouseXAndToolTip:=16, DistanceBetweenMouseYAndToolTip:=16, DebugMode:=0

  __New()  {

    if (!this.pToken)    {
      ; 加速
      SavedBatchLines:=A_BatchLines
      SetBatchLines, -1

      ; Start gdi+
      this.pToken := Gdip_Startup()
      if (!this.pToken)      {
        MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
        ExitApp
      }

      ; 多显示器支持
      this.Monitors := MDMF_Enum()

      ; 获取整个桌面的分辨率，即使跨显示器
      SysGet, VirtualWidth, 78
      SysGet, VirtualHeight, 79
      this.DIBWidth  := VirtualWidth
      this.DIBHeight := VirtualHeight

      ; 获取 ToolTip 的默认字体
      this.ToolTipFontName := Fnt_GetTooltipFontName()

      ; create 20 guis for gdi+
      ; 最多20个 ToolTip ，与原版对应。
      loop, 20      {
        ; _BTT1（GUI 名称） 与 _hBTT1（GUI 句柄） 都是临时变量，后者被储存了。
        Gui, _BTT%A_Index%: +E0x80000 -Caption +ToolWindow +LastFound +AlwaysOnTop +Hwnd_hBTT%A_Index%
        Gui, _BTT%A_Index%: Show, NA

          this["hBTT" A_Index] 	:= _hBTT%A_Index%
        , this["hbm" A_Index]  	:= CreateDIBSection(this.DIBWidth, this.DIBHeight)
        , this["hdc" A_Index]   	:= CreateCompatibleDC()
        , this["obm" A_Index]  	:= SelectObject(this["hdc" A_Index], this["hbm" A_Index])
        , this["G" A_Index]      	:= Gdip_GraphicsFromHDC(this["hdc" A_Index])
        , Gdip_SetSmoothingMode(this["G" A_Index], 4)
        , Gdip_SetPixelOffsetMode(this["G" A_Index], 2)		; 此参数是画出完美圆角矩形的关键
      }
      SetBatchLines, %SavedBatchLines%
    }
    else
      return
  }

  ; new The variables obtained later will automatically jump here to run after being destroyed, so it is very suitable for automatic resource recovery.
  __Delete()  {
    loop, 20    {
        Gdip_DeleteGraphics(this["G" A_Index])
      , SelectObject(this["hdc" A_Index], this["obm" A_Index])
      , DeleteObject(this["hbm" A_Index])
      , DeleteDC(this["hdc" A_Index])
    }
    Gdip_Shutdown(this.pToken)
  }

  ; 参数默认全部为空只是为了让清空 ToolTip 的语法简洁而已。
  ToolTip(Text:="", X:="", Y:="", WhichToolTip:="", Styles:="", Options:="")  {
    ; 给出 WhichToolTip 的默认值1，并限制 WhichToolTip 的范围为 1-20
    WhichToolTip:=(WhichToolTip="") ? 1 : Range(WhichToolTip, 1, 20)
    ; 检查并解析 Options 。无论不传参、部分传参、完全传参，此函数均能正确返回所需参数。
    O:=this._CheckStylesAndOptions(Styles, Options)

    ; 判断显示内容是否发生变化。由于前面给了 Options 一个默认值，所以首次进来时下面的 O.Options!=SavedOptions 是必然成立的。
    FirstCallOrNeedToUpdate:=(Text       != this["SavedText" WhichToolTip]
                           || O.Checksum != this["SavedOptions" WhichToolTip])

    if (Text="")    {
      ; 清空 ToolTip
      Gdip_GraphicsClear(this["G" WhichToolTip])
      UpdateLayeredWindow(this["hBTT" WhichToolTip], this["hdc" WhichToolTip])
      ; 清空变量
        this["SavedText" WhichToolTip]       	:= ""
      , this["SavedOptions" WhichToolTip]    	:= ""
      , this["SavedX" WhichToolTip]           	:= ""
      , this["SavedY" WhichToolTip]           	:= ""
      , this["SavedW" WhichToolTip]          	:= ""
      , this["SavedH" WhichToolTip]           	:= ""
      , this["SavedCoordMode" WhichToolTip]  := ""
      , this["SavedTargetHWND" WhichToolTip] := ""
    }
    else if (FirstCallOrNeedToUpdate) {		; First Call or NeedToUpdate

      ; 加速
      SavedBatchLines:=A_BatchLines
      SetBatchLines, -1

      ; Obtain the target size, which is used to add width and height restrictions when calculating the text size.
	  ; Otherwise, when the target is a screen, the ultra-wide and ultra-high size may be calculated, resulting in failure to display.
        TargetSize := this._CalculateDisplayPosition(X, Y, "", "", O, GetTargetSize := 1)
      ; Make the text area + margin + thin border not exceed the target width.
      , MaxTextWidth := TargetSize.W - O.Margin*2 - O.Border*2
      ; 使得 文字区域+边距+细边框 不会超过目标高度的90%。
      ; 之所以高度限制90%是因为两个原因，1是留出一些上下的空白，避免占满全屏，鼠标点不了其它地方，难以退出。
      ; 2是在计算文字区域时，即使已经给出了宽高度限制，且因为自动换行的原因，宽度的返回值通常在范围内，但高度的返回值偶尔还是会超过1行，所以提前留个余量。
      , MaxTextHeight := (TargetSize.H*90)//100 - O.Margin*2 - O.Border*2
      ; 为 Gdip_TextToGraphics2 计算区域提供高宽限制。
      , O.Width := MaxTextWidth, O.Height := MaxTextHeight
      ; Calculate the text display area TextArea = x|y|width|height|chars|lines
      , TextArea := StrSplit(Gdip_TextToGraphics2(this["G" WhichToolTip], Text, O, Measure := 1), "|")
		; This must be rounded up.
       ; The reason for the upward direction is, for example, if 1.2 is rounded to 1, then the rightmost character may not be displayed completely.
       ; The reason for rounding is that you cannot draw a perfect rounded rectangle without rounding.
       ; When the AutoTrim option is used, that is, the out-of-range text is automatically displayed as "...", the width and height values returned at this time do not include "...".
       ; So after adding the width and height of "...", it may still exceed the limit. It needs to be checked again and restricted.
       ; Once the width and height exceed the limit (the size created in CreateDIBSection()), it will cause UpdateLayeredWindow() to fail to draw the image.
      , TextWidth  	:= Min(Ceil(TextArea[3]), MaxTextWidth)
      , TextHeight 	:= Min(Ceil(TextArea[4]), MaxTextHeight)
      , RectWidth  	:= TextWidth+O.Margin*2                                   ; 文本+边距。
      , RectHeight 	:= TextHeight+O.Margin*2
      , RectWithBorderWidth := RectWidth+O.Border*2                         ; 文本+边距+细边框。
      , RectWithBorderHeight := RectHeight+O.Border*2
      ; 圆角超过矩形宽或高的一半时，会画出畸形的圆，所以这里验证并限制一下。
      , R := (O.Rounded>Min(RectWidth, RectHeight)//2) ? Min(RectWidth, RectHeight)//2 : O.Rounded

      ; 画之前务必清空画布，否则会出现异常。
      Gdip_GraphicsClear(this["G" WhichToolTip])

      ; 准备画刷
      pBrushBorder := Gdip_BrushCreateSolid(O.BorderColor)                    ; 纯色画刷 画细边框
      if (O.BGCLGD and O.BGCLGS and O.BGCLGE)  {                               ; 渐变画刷 画细边框

        Left:=O.Border, Top:=O.Border, Right:=Left+RectWidth, Bottom:=Top+RectHeight
        switch, O.BGCLGD  {
          ; O.BGCLGD 1是横向渐变，2是左上角到右下角斜向渐变，3是垂直渐变。
          case, "1": pBrushBackground:=Gdip_CreateLinearGrBrush(Left, Top, Right, Top,    O.BGCLGS, O.BGCLGE)
          case, "2": pBrushBackground:=Gdip_CreateLinearGrBrush(Left, Top, Right, Bottom, O.BGCLGS, O.BGCLGE)
          case, "3": pBrushBackground:=Gdip_CreateLinearGrBrush(Left, Top, Left,  Bottom, O.BGCLGS, O.BGCLGE)
        }
      }
      else
        pBrushBackground := Gdip_BrushCreateSolid(O.BackgroundColor)          ; 纯色画刷 画文本框

      if (O.Border>0)
        switch, R        {
          ; 圆角为0则使用矩形画。不单独处理，会画出显示异常的图案。
          case, "0": Gdip_FillRectangle(this["G" WhichToolTip]                ; 矩形细边框
          , pBrushBorder, 0, 0, RectWithBorderWidth, RectWithBorderHeight)
          Default  : Gdip_FillRoundedRectanglePath(this["G" WhichToolTip]     ; 圆角细边框
          , pBrushBorder, 0, 0, RectWithBorderWidth, RectWithBorderHeight, R)
        }

		switch, R      {
        case, "0": Gdip_FillRectangle(this["G" WhichToolTip]                  ; 矩形文本框
        , pBrushBackground, O.Border, O.Border, RectWidth, RectHeight)
        Default  : Gdip_FillRoundedRectanglePath(this["G" WhichToolTip]       ; 圆角文本框
        , pBrushBackground, O.Border, O.Border, RectWidth, RectHeight
        , (R>O.Border) ? R-O.Border : R)                                      ; 确保内外圆弧看起来同心
      }

      ; 清理画刷
      Gdip_DeleteBrush(pBrushBorder)
      Gdip_DeleteBrush(pBrushBackground)

      ; 计算居中显示坐标。由于 TextArea 返回的文字范围右边有很多空白，所以这里的居中坐标并不精确。
      O.X:=O.Border+O.Margin, O.Y:=O.Border+O.Margin, O.Width:=TextWidth, O.Height:=TextHeight

      ; If the display area is too small and the text cannot be displayed completely, replace the last 4 characters of the text to be displayed with 4 ellipsis, indicating that the display is not complete.
       ; Although there is a GdipSetStringFormatTrimming function to set the ellipsis at the end, it occasionally needs to add extra width to display it.
       ; Because it is difficult to judge whether the extra width needs to be added, and how much it needs to be added, etc., so directly use this method to realize the display of the ellipsis.
       ; The reason why we choose to replace the last 4 characters is because we generally replace the last 2 characters to ensure that at least one ellipsis is displayed.
       ; In order to deal with unexpected situations and make the ellipsis more obvious, I chose to replace the last four.
       ; The original Text needs to be used for comparison before display, so the original value cannot be changed, and TempText must be used.
      if (TextArea[5]<StrLen(Text))
        TempText := TextArea[5]>4 ? SubStr(Text, 1 ,TextArea[5]-4) "…………" : SubStr(Text, 1 ,1) "…………"
      else
        TempText := Text

      ; 写字到框上。这个函数使用 O 中的 X,Y 去调整文字的位置。
      Gdip_TextToGraphics2(this["G" WhichToolTip], TempText, O)

      ; 调试用，可显示计算得到的文字范围。
      if (this.DebugMode)      {
        pBrush := Gdip_BrushCreateSolid(0x20ff0000)
        Gdip_FillRectangle(this["G" WhichToolTip], pBrush, O.Border+O.Margin, O.Border+O.Margin, TextWidth, TextHeight)
        Gdip_DeleteBrush(pBrush)
      }

      ; 返回文本框不超出目标范围（比如屏幕范围）的最佳坐标。
      this._CalculateDisplayPosition(X, Y, RectWithBorderWidth, RectWithBorderHeight, O)

      ; 显示
      UpdateLayeredWindow(this["hBTT" WhichToolTip], this["hdc" WhichToolTip], X, Y, RectWithBorderWidth, RectWithBorderHeight)

      ; 保存参数值，以便之后比对参数值是否改变
        this["SavedText" WhichToolTip]            	:= Text
      , this["SavedOptions" WhichToolTip]      	:= O.Checksum
      , this["SavedX" WhichToolTip]                	:= X                             	; 这里的 X,Y 是经过 _CalculateDisplayPosition() 计算后的新 X,Y
      , this["SavedY" WhichToolTip]                	:= Y
      , this["SavedW" WhichToolTip]              	:= RectWithBorderWidth
      , this["SavedH" WhichToolTip]                	:= RectWithBorderHeight
      , this["SavedCoordMode" WhichToolTip]  	:= O.CoordMode
      , this["SavedTargetHWND" WhichToolTip]	:= O.TargetHWND

      SetBatchLines, %SavedBatchLines%
    }
    ; x,y 任意一个跟随鼠标 或 使用窗口或客户区模式（窗口大小可能发生改变）
    ; 或 坐标模式发生变化 或 目标窗口发生变化 这4种情况可能需要移动位置，需要进行坐标计算。
    else if ((X="" || Y="") || O.CoordMode!="Screen" || O.CoordMode!=this.SavedCoordMode || O.TargetHWND!=this.SavedTargetHWND)    {
      ; 返回文本框不超出目标范围（比如屏幕范围）的最佳坐标。
      this._CalculateDisplayPosition(X, Y, this["SavedW" WhichToolTip], this["SavedH" WhichToolTip], O)
      ; 判断文本框显示位置是否发生改变
      if (X!=this["SavedX" WhichToolTip] || Y!=this["SavedY" WhichToolTip])      {
        ; 显示
        UpdateLayeredWindow(this["hBTT" WhichToolTip], this["hdc" WhichToolTip], X, Y, this["SavedW" WhichToolTip], this["SavedH" WhichToolTip])

        ; 保存新的位置
          this["SavedX" WhichToolTip]                    	:= X
        , this["SavedY" WhichToolTip]                  	:= Y
        , this["SavedCoordMode" WhichToolTip]  	:= O.CoordMode
        , this["SavedTargetHWND" WhichToolTip] 	:= O.TargetHWND
      }
    }
  }

  ; 此函数确保传入空值或者错误值均可返回正确值。
  _CheckStylesAndOptions(Styles, Options)  {
      O                 := IsObject(Styles)     ? Styles.Clone()       : Object()
    , O.Border          := O.Border=""          ? 1                    : Range(O.Border, 0, 20)        ; 细边框  	默认1 0-20
    , O.Rounded         := O.Rounded=""         ? 3                    : Range(O.Rounded, 0, 30)       ; 圆角    	默认3 0-30
    , O.Margin          := O.Margin=""          ? 5                    : Range(O.Margin, 0, 30)        ; 边距    	默认5 0-30
    , O.TextColor       := O.TextColor=""       ? 0xff575757           : O.TextColor                   ; 文本色  	默认0xff575757
    , O.BackgroundColor := O.BackgroundColor="" ? 0xffffffff           : O.BackgroundColor             ; 背景色  	默认0xffffffff
    , O.Font            := O.Font=""            ? this.ToolTipFontName : O.Font                        ; 字体    	默认与 ToolTip 一致
    , O.FontSize        := O.FontSize=""        ? 12                   : O.FontSize                    ; 字号   		默认12
    , O.FontRender      := O.FontRender=""      ? 5                    : Range(O.FontRender, 0, 5)     ; 渲染模式		默认5 0-5

    ; a:=0xaabbccdd 下面是运算规则
    ; a>>16    = 0xaabb
    ; a>>24    = 0xaa
    ; a&0xffff = 0xccdd
    ; a&0xff   = 0xdd
    ; 0x88<<16 = 0x880000
    ; 0x880000+0xbbcc = 0x88bbcc
    , BlendedColor      := ((O.BackgroundColor>>24)<<24) + (O.TextColor&0xffffff)                      ; 混合色 背景色的透明度与文本色混合
    , O.BorderColor     := O.BorderColor=""      ? BlendedColor         : O.BorderColor                ; 细边框色  		默认混合色

    ; 名字太长，建个缩写副本。
    , O.BGCLGS          := O.BackgroundColorLinearGradientStart                                        ; 背景渐变色		默认无
    , O.BGCLGE          := O.BackgroundColorLinearGradientEnd                                          ; 背景渐变色		默认无
    , O.BGCLGD          := O.BackgroundColorLinearGradientDirection
    , O.BGCLGD          := O.BGCLGD=""           ? ""                   : Range(O.BGCLGD, 1, 3)        ; 背景渐变方向	默认无 1-3
    , O.FontStyle       := O.FontStyle                                                                 ; 字体样式    	默认无

    , O.TargetHWND      := Options.TargetHWND="" ? WinExist("A")        : Options.TargetHWND+0         ; 目标句柄    	默认活动窗口
    , O.CoordMode       := Options.CoordMode=""  ? A_CoordModeToolTip   : Options.CoordMode            ; 坐标模式    	默认与 ToolTip 一致

    ; 名字太长，一行写不下，故缩短。
    , temp1:=Options.MouseNeverCoverToolTip
    , temp2:=Options.DistanceBetweenMouseXAndToolTip
    , temp3:=Options.DistanceBetweenMouseYAndToolTip
    , this.MouseNeverCoverToolTip          := temp1="" ? this.MouseNeverCoverToolTip          : temp1
    , this.DistanceBetweenMouseXAndToolTip := temp2="" ? this.DistanceBetweenMouseXAndToolTip : temp2
    , this.DistanceBetweenMouseYAndToolTip := temp3="" ? this.DistanceBetweenMouseYAndToolTip : temp3

    ; 难以比对两个对象是否一致，所以造一个变量比对。
    ; 这里的校验因素，必须是那些改变后会使画面内容也产生变化的因素。
    ; 所以没有 TargetHWND 和 CoordMode ，因为这两个因素只影响位置。
    , O.Checksum        := Format("{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13}"
                                , O.Border, O.Rounded, O.Margin, O.BorderColor, O.TextColor
                                , O.BackgroundColor, O.BGCLGS, O.BGCLGE, O.BGCLGD
                                , O.Font, O.FontSize, O.FontRender, O.FontStyle)
    return, O
  }

  ; 此函数确保文本框显示位置不会超出目标范围。
  ; 使用 ByRef X, ByRef Y 返回不超限的位置。
  _CalculateDisplayPosition(ByRef X, ByRef Y, W, H, Options, GetTargetSize:=0)  {
      VarSetCapacity(Point, 8, 0)
    ; 获取鼠标位置
    , DllCall("GetCursorPos", "Ptr", &Point)
    , MouseX := NumGet(Point, 0, "Int"), MouseY := NumGet(Point, 4, "Int")

    ; x,y 即 ToolTip 显示的位置。
    ; x,y 同时为空表明完全跟随鼠标。
    ; x,y 单个为空表明只跟随鼠标横向或纵向移动。
    ; x,y 都有值，则说明被钉在屏幕或窗口或客户区的某个位置。
    ; MouseX,MouseY 是鼠标的屏幕坐标。
    ; DisplayX,DisplayY 是 x,y 经过转换后的屏幕坐标。
    ; 以下过程 x,y 不发生变化， DisplayX,DisplayY 储存转换好的屏幕坐标。
    ; 不要尝试合并分支 (X="" and Y="") 与 (A_CoordModeToolTip = "Screen")。
    ; 因为存在把坐标模式设为 Window 或 Client 但又同时不给出 x,y 的情况！！！！！！
    if (X="" and Y="")    {	; 没有给出 x,y 则使用鼠标坐标
        DisplayX     := MouseX
      , DisplayY     := MouseY
      ; 根据坐标判断在第几个屏幕里，并获得对应屏幕边界。
      ; 使用 MONITOR_DEFAULTTONEAREST 设置，可以在给出的点不在任何显示器内时，返回距离最近的显示器。
      ; 这样可以修正使用 1920,1080 这种错误的坐标，导致返回空值，导致画图失败的问题。
      ; 为什么 1920,1080 是错误的呢？因为 1920 是宽度，而坐标起点是0，所以最右边坐标值是 1919，最下面是 1079。
      , hMonitor     := MDMF_FromPoint(DisplayX, DisplayY, MONITOR_DEFAULTTONEAREST:=2)
      , TargetLeft   := this.Monitors[hMonitor].Left
      , TargetTop    := this.Monitors[hMonitor].Top
      , TargetRight  := this.Monitors[hMonitor].Right
      , TargetBottom := this.Monitors[hMonitor].Bottom
      , TargetWidth  := TargetRight-TargetLeft
      , TargetHeight := TargetBottom-TargetTop
    }
    ; 已给出 x和y 或x 或y，都会走到下面3个分支去。
    else if (Options.CoordMode  = "Window" || Options.CoordMode  = "Relative")    {	; 已给出 x或y 且使用窗口坐标
        WinGetPos, WinX, WinY, WinW, WinH, % "ahk_id " Options.TargetHWND

        XInScreen    := WinX+X
      , YInScreen    := WinY+Y
      , TargetLeft   := WinX
      , TargetTop    := WinY
      , TargetWidth  := WinW
      , TargetHeight := WinH
      , TargetRight  := TargetLeft+TargetWidth
      , TargetBottom := TargetTop+TargetHeight
      , DisplayX     := (X="") ? MouseX : XInScreen
      , DisplayY     := (Y="") ? MouseY : YInScreen
    }
    else if (Options.CoordMode  = "Client")    {	; 已给出 x或y 且使用客户区坐标
        VarSetCapacity(ClientArea, 16, 0)
      , DllCall("GetClientRect", "Ptr", Options.TargetHWND, "Ptr", &ClientArea)
      , DllCall("ClientToScreen", "Ptr", Options.TargetHWND, "Ptr", &ClientArea)
      , ClientX      := NumGet(ClientArea, 0, "Int")
      , ClientY      := NumGet(ClientArea, 4, "Int")
      , ClientW      := NumGet(ClientArea, 8, "Int")
      , ClientH      := NumGet(ClientArea, 12, "Int")

        XInScreen    := ClientX+X
      , YInScreen    := ClientY+Y
      , TargetLeft   := ClientX
      , TargetTop    := ClientY
      , TargetWidth  := ClientW
      , TargetHeight := ClientH
      , TargetRight  := TargetLeft+TargetWidth
      , TargetBottom := TargetTop+TargetHeight
      , DisplayX     := (X="") ? MouseX : XInScreen
      , DisplayY     := (Y="") ? MouseY : YInScreen
    }
    else { ; 这里必然 A_CoordModeToolTip = "Screen"
    	; 已给出 x或y 且使用屏幕坐标
        DisplayX     := (X="") ? MouseX : X
      , DisplayY     := (Y="") ? MouseY : Y
      ; 根据坐标判断在第几个屏幕里，并获得对应屏幕边界。
      ; 使用 MONITOR_DEFAULTTONEAREST 设置，可以在给出的点不在任何显示器内时，返回距离最近的显示器。
      ; 这样可以修正使用 1920,1080 这种错误的坐标，导致返回空值，导致画图失败的问题。
      ; 为什么 1920,1080 是错误的呢？因为 1920 是宽度，而坐标起点是0，所以最右边坐标值是 1919，最下面是 1079。
      , hMonitor     := MDMF_FromPoint(DisplayX, DisplayY, MONITOR_DEFAULTTONEAREST:=2)
      , TargetLeft   := this.Monitors[hMonitor].Left
      , TargetTop    := this.Monitors[hMonitor].Top
      , TargetRight  := this.Monitors[hMonitor].Right
      , TargetBottom := this.Monitors[hMonitor].Bottom
      , TargetWidth  := TargetRight-TargetLeft
      , TargetHeight := TargetBottom-TargetTop
    }

    if (GetTargetSize=1)    {
        TargetSize   := []
      , TargetSize.X := TargetLeft
      , TargetSize.Y := TargetTop
      ; 一个窗口，有各种各样的方式可以让自己的高宽超过屏幕高宽。
      ; 例如最大化的时候，看起来刚好填满了屏幕，应该是1920*1080，但实际获取会发现是1936*1096。
      ; 还可以通过拖动窗口边缘调整大小的方式，让它变1924*1084。
      ; 还可以直接在创建窗口的时候，指定一个数值，例如3000*3000。
      ; 由于设计的时候， DIB 最大就是多个屏幕大小的总和。
      ; 当造出一个超过屏幕大小总和的窗口，又使用了 A_CoordModeToolTip = "Window" 之类的参数，同时待显示文本单行又超级长。
      ; 此时 (显示宽高 = 窗口宽高) > DIB宽高，会导致 UpdateLayeredWindow() 显示失败。
      ; 所以这里做一下限制。
      , TargetSize.W := Min(TargetWidth, this.DIBWidth)
      , TargetSize.H := Min(TargetHeight, this.DIBHeight)
      return, TargetSize
    }

      DPIScale := A_ScreenDPI/96
    ; 为跟随鼠标显示的文本框增加一个距离，避免鼠标和文本框挤一起发生遮挡。
    ; 因为前面需要用到原始的 DisplayX 和 DisplayY 进行计算，所以在这里才增加距离。
    , DisplayX := (X="") ? DisplayX+this.DistanceBetweenMouseXAndToolTip*DPIScale : DisplayX
    , DisplayY := (Y="") ? DisplayY+this.DistanceBetweenMouseYAndToolTip*DPIScale : DisplayY

    ; 处理目标边缘（右和下）的情况，让文本框可以贴边显示，不会超出目标外。
    , DisplayX := (DisplayX+W>=TargetRight)  ? TargetRight-W  : DisplayX
    , DisplayY := (DisplayY+H>=TargetBottom) ? TargetBottom-H : DisplayY
    ; 处理目标边缘（左和上）的情况，让文本框可以贴边显示，不会超出目标外。
    ; 不建议合并代码，理解会变得困难。
    , DisplayX := (DisplayX<TargetLeft) ? TargetLeft : DisplayX
    , DisplayY := (DisplayY<TargetTop)  ? TargetTop  : DisplayY

    ; 处理鼠标遮挡文本框的情况（即鼠标跑到文本框坐标范围内了）。这里要做测试的话，需要测试5种情况。
    ; X跟随 Y跟随。
    ; X跟随 Y固定。0和1919都要测
    ; Y跟随 X固定。0和1079都要测
    ; X固定 Y固定。此种情况文本框可被鼠标遮挡，无需测试。
    if  (this.MouseNeverCoverToolTip=1
    and (X="" || Y="")
    and MouseX>=DisplayX and MouseY>=DisplayY and MouseX<=DisplayX+W and MouseY<=DisplayY+H)    {
      ; MouseY-H-16 是往上弹，应对在左下角和右下角的情况。
      ; MouseY+H+16 是往下弹，应对在右上角和左上角的情况。
      ; 这里不要去用 Abs(this.DistanceBetweenMouseYAndToolTip) 替代 16。因为当前者很大时，显示效果不好。
      ; 优先往上弹，如果不超限，则上弹。如果超限则往下弹，下弹超限，则不弹。
      DisplayY := MouseY-H-16>=TargetTop ? MouseY-H-16 : MouseY+H+16<=TargetBottom ? MouseY+16 : DisplayY
    }

    ; 使用 ByRef 变量特性返回计算得到的 X和Y
    X := DisplayX , Y := DisplayY
  }

}

NonNull(Value1, Value2){
  return, Value1="" ? Value2 : Value1
}

Range(Value, MinValue, MaxValue){
  ; 三元的写法不会更快
  return, Max(Min(Value, MaxValue), MinValue)
}

Gdip_TextToGraphics2(pGraphics, Text, Options, Measure:=0){
	static Styles := "Regular|Bold|Italic|BoldItalic|Underline|Strikeout"


	; 设置字体样式
	Style := 0
	For eachStyle, valStyle in StrSplit(Styles, "|")	{
		if InStr(Options.FontStyle, valStyle)
			Style |= (valStyle != "StrikeOut") ? (A_Index-1) : 8
	}

	; 加载字体
	If !(hFontFamily := Gdip_FontFamilyCreate(Options.Font))
		hFontFamily := Gdip_FontFamilyCreateGeneric(1)
	hFont := Gdip_FontCreate(hFontFamily, Options.FontSize, Style, Unit:=0)

	; Set text formatting style, LineLimit = 0x00002000 only display complete lines.
	; For example, the last line, because of the limited layout height, can only display half of it, so it will not be displayed at all.
	hStringFormat := Gdip_StringFormatCreate(0x00002020)
	if !hStringFormat
		hStringFormat := Gdip_StringFormatGetGeneric(0)

  ; 设置文字颜色
	pBrush := Gdip_BrushCreateSolid(Options.TextColor)

	; 检查参数是否齐全
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

	Gdip_SetStringFormatAlign(hStringFormat, Align:=0)                         ; 设置左对齐
	Gdip_SetTextRenderingHint(pGraphics, Options.FontRender)                   ; 设置渲染模式
	CreateRectF(RC                                                             ; x,y 需要至少为0
            , Options.X="" ? 0 : Options.X
            , Options.Y="" ? 0 : Options.Y
            , Options.Width, Options.Height)                                 ; 宽高可以为空
	returnRC := Gdip_MeasureString(pGraphics, Text, hFont, hStringFormat, RC)  ; 计算大小

	if !Measure
		_E := Gdip_DrawString(pGraphics, Text, hFont, hStringFormat, pBrush, RC)

	Gdip_DeleteBrush(pBrush)
	Gdip_DeleteFont(hFont)
	Gdip_DeleteStringFormat(hStringFormat)
	Gdip_DeleteFontFamily(hFontFamily)
	return _E ? _E : returnRC
}

Fnt_GetTooltipFontName(){
	static LF_FACESIZE:=32  ;-- In TCHARS
	return StrGet(Fnt_GetNonClientMetrics()+(A_IsUnicode ? 316:220)+28,LF_FACESIZE)
}
Fnt_GetNonClientMetrics(){
	static Dummy15105062
		,SPI_GETNONCLIENTMETRICS:=0x29
		,NONCLIENTMETRICS

	;-- Set the size of NONCLIENTMETRICS structure
	cbSize:=A_IsUnicode ? 500:340
	if (((GV:=DllCall("GetVersion"))&0xFF . "." . GV>>8&0xFF)>=6.0)  ;-- Vista+
		cbSize+=4

	;-- Create and initialize NONCLIENTMETRICS structure
	VarSetCapacity(NONCLIENTMETRICS,cbSize,0)
	NumPut(cbSize,NONCLIENTMETRICS,0,"UInt")

	;-- Get nonclient metrics parameter
	if !DllCall("SystemParametersInfo"
		,"UInt",SPI_GETNONCLIENTMETRICS
		,"UInt",cbSize
		,"Ptr",&NONCLIENTMETRICS
		,"UInt",0)
		return false

	;-- Return to sender
	return &NONCLIENTMETRICS
}

MDMF_Enum(HMON := "") {
   Static CallbackFunc := Func(A_AhkVersion < "2" ? "RegisterCallback" : "CallbackCreate")
   Static EnumProc := CallbackFunc.Call("MDMF_EnumProc")
   Static Obj := (A_AhkVersion < "2") ? "Object" : "Map"
   Static Monitors := {}
   If (HMON = "") ; new enumeration
   {
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
      Return {Name:      (Name := StrGet(&MIEX + 40, 32))  ; CCHDEVICENAME = 32
            , Num:       RegExReplace(Name, ".*(\d+)$", "$1")
            , Left:      NumGet(MIEX, 4, "Int")    ; display rectangle
            , Top:       NumGet(MIEX, 8, "Int")    ; "
            , Right:     NumGet(MIEX, 12, "Int")   ; "
            , Bottom:    NumGet(MIEX, 16, "Int")   ; "
            , WALeft:    NumGet(MIEX, 20, "Int")   ; work area
            , WATop:     NumGet(MIEX, 24, "Int")   ; "
            , WARight:   NumGet(MIEX, 28, "Int")   ; "
            , WABottom:  NumGet(MIEX, 32, "Int")   ; "
            , Primary:   NumGet(MIEX, 36, "UInt")} ; contains a non-zero value for the primary monitor.
   Return False
}

LoadBar_Gui(show:=1, opt:="") {

	global load_BarGUI, LBar, hLoad_BarWin

	If !IsObject(opt)
		opt:={title1:"LoadBar", title2:"Loadbar", "w":320,	"h":36}
	If !IsObject(opt.col)
		opt.col := ["0x4D4D4D","0xFFFFFF","0xEFEFEF"]

	Gui, load_BarGUI: -Border -Caption +ToolWindow HWNDhLoad_BarWin
	Gui, load_BarGUI: Color, % opt.col.1, % opt.col.2

   ;                                                                                                      sDesc, FontColorDesc, FontColor, BG
	;~ LBar := new LoaderBar("load_BarGUI", 3, 3, opt.w, opt.h, opt.title1, opt.title2, "FFFFFF", opt.col.3, "", "B2B2B2|FF871A|6D6D6D", "0x66A3E2|0x4B79AF|0x385D87")
	LBar := new LoaderBar("load_BarGUI", 3, 3, opt.w, opt.h, opt.title1, opt.title2, opt.col.1, opt.col.2)

	Gui, load_BarGUI: Show, % "w" LBar.Width+2*LBar.X " h" LBar.Height+2*LBar.Y, % opt.title2
	LBar.hWin := hLoad_BarWin

return LBar
}

LoadBar_Callback(param*) {

	; param.1 = index, 2= maxIndex, 3= len, 4= matchcount
	global LBar
	static lastfProgress, msg, lastmsg

	If (param.1 ~= "\D")
		msg := param.RemoveAt(1)

	LBar.Set(Floor(param.1*100/param.2), msg~="^\d" ? msg : msg " / " param.4)

}

class LoaderBar {

	__New(GUI_ID:="Default",x:=0,y:=0,w:=280,h:=28
							, Title:=""
							, ShowDesc:=0
							, FontColorDesc:="FEFEFE"
							, FontColor:="EFEFEF"
							, BG:="2B2B2B|2F2F2F|323232"
							,FG:="66A3E2|4B79AF|385D87") {
		SetWinDelay,0
		SetBatchLines,-1
		if (StrLen(A_Gui))
			_GUI_ID:=A_Gui
		else
			_GUI_ID:=1
		if ( (GUI_ID="Default") || !StrLen(GUI_ID) || GUI_ID==0 )
			GUI_ID:=_GUI_ID

		this.GUI_ID := GUI_ID
		Gui, %GUI_ID%:Default

		this.BG     	:= StrSplit(BG,"|")
		this.BG.W  	:= w
		this.BG.H  	:= h
		this.Width 	:=w
		this.Height	:=h
		this.FG       	:= StrSplit(FG,"|")
		this.FG.W  	:= this.BG.W - 2
		this.FG.H   	:= (fg_h:=(this.BG.H - 2))
		this.Percent 	:= 0
		this.X        	:= x
		this.Y        	:= y
		fg_x            	:= this.X + 1
		fg_y            	:= this.Y + 1
		this.FontColor := FontColor
		this.FontColor2 := FontColor2 := "FFFFFF"
		this.ShowDesc := ShowDesc
		this.FontColorDesc := FontColorDesc

		;DescBGColor:="4D4D4D"
		DescBGColor:="Black"
		this.DescBGColor := DescBGColor


		Gui,Font, Bold s10
		Gui, Add, Text, % "x" x " y" y " w" w " h" h " BackgroundTrans 0xE hwndhLoaderBarTitle cFFFFFF", % Title
		this.hLoaderBarTitle := hLoaderBarTitle


		Gui,Font, Normal s8

		rcol := []
		Loop 3 {
			Random, rColor, 0x222222, 0xDDDDDD
			rcol.Push(rColor)
		}

		Gui, Add, Picture, x%x% y%y% w%w% h%h% 0xE hwndhLoaderBarBG, % "HBITMAP:" this.CreateDIB(rcol.1 "|" rcol.1 "|" rcol.1 "|" rcol.2 "|" rcol.2 "|" rcol.2 "|" rcol.3 "|" rcol.3 "|" rcol.3, 3, 3, w, h, 1,0)
			this.hLoaderBarBG := hLoaderBarBG

		Gui, Add, Picture, x%fg_x% y%fg_y% w0 h%fg_h% 0xE hwndhLoaderBarFG, % "HBITMAP:" this.CreateDIB("FFBE84|FFBE84|FFBE84|FF891E|FF891E|FF891E|C15B00|C15B00|C15B00", 3, 3, 0, h, 1, 1 )
			this.hLoaderBarFG := hLoaderBarFG

		Gui,Font, Normal s10, Consolas
		Gui, Add, Text, x%x% y%y% w%w% h%h% 0x200 border center BackgroundTrans hwndhLoaderNumber c%FontColor%, % "[ 0 % ]"
			this.hLoaderNumber := hLoaderNumber

		Gui,Font, Normal s8, Arial
		if (this.ShowDesc) {
			Gui, Add, Text, xp y+2 w%w% h16 0x200 Center border BackgroundTrans hwndhLoaderDesc c%FontColor2%, this.ShowDesc
			this.hLoaderDesc := hLoaderDesc
			this.Height:=h+18
		}

		Gui,Font

		Gui, %_GUI_ID%:Default
	}

	Set(p,w:="Loading...") {
		if (StrLen(A_Gui))
			_GUI_ID:=A_Gui
		else
			_GUI_ID:=1
		GUI_ID := this.GUI_ID

		Gui, %GUI_ID%:Default
		GuiControlGet, LoaderBarBG, Pos, % this.hLoaderBarBG

		this.Percent:=(p>=100) ? p:=100 : p

		PercentNum	:= Round(this.Percent,0)
		PercentBar	:= Floor((this.Percent/100)*this.FG.W)
		hLoaderBarTitle		:= this.hLoaderBarTitle
		hLoaderBarFG  	:= this.hLoaderBarFG
		hLoaderBarBG  	:= this.hLoaderBarBG
		hLoaderNumber 	:= this.hLoaderNumber

		GuiControl,Move	,% hLoaderBarFG  	, % "w" PercentBar
		GuiControl,       	,% hLoaderBarFG   	, % "HBITMAP: " this.CreateDIB("FFBE84|FFBE84|FFBE84|FF891E|FF891E|FF891E|994800|994800|994800", 3, 3, PercentBar, this.BG.H, 1, 0 )
		GuiControl,       	,% hLoaderNumber 	, % "[ " PercentNum "% ]"

		if (this.ShowDesc) {
			hLoaderDesc := this.hLoaderDesc
			GuiControl,,%hLoaderDesc%, %w%
		}
		Gui, %_GUI_ID%:Default
	}

	ApplyGradient( Hwnd, LT := "101010", MB := "0000AA", RB := "00FF00", Vertical := 1 ) {
		Static STM_SETIMAGE := 0x172
		ControlGetPos,,, W, H,, ahk_id %Hwnd%
		PixelData := Vertical ? LT "|" LT "|" LT "|" MB "|" MB "|" MB "|" RB "|" RB "|" RB : LT "|" MB "|" RB "|" LT "|" MB "|" RB "|" LT "|" MB "|" RB
		PixelData := "
			( LTrim Join|
			  CC0000|CC0000|CC0000|CC0000|CC0000|CC0000|CC0000|CC0000|CC0000
			  AA0000|AA0000|AA0000|AA0000|AA0000|AA0000|AA0000|AA0000|AA0000
			  990000|990000|990000|990000|990000|990000|990000|990000|990000
			  880000|880000|880000|880000|880000|880000|880000|880000|880000
			  770000|770000|770000|770000|770000|770000|770000|770000|770000
			  660000|660000|660000|660000|660000|660000|660000|660000|660000
			  550000|550000|550000|550000|550000|550000|550000|550000|550000
			  440000|440000|440000|440000|440000|440000|440000|440000|440000
			  330000|330000|330000|330000|330000|330000|330000|330000|330000
			)"
		hBitmap := this.CreateDIB( PixelData, 3, 3, W, H, True )
		oBitmap := DllCall( "SendMessage", "Ptr",Hwnd, "UInt",STM_SETIMAGE, "Ptr",0, "Ptr",hBitmap )
		Return hBitmap, DllCall( "DeleteObject", "Ptr",oBitmap )
	}

	CreateDIB(PixelData, W, H, ResizeW:=0, ResizeH:=0, Gradient:=0, DIB:=1) {
		Static OSV    ; CreateDIB v0.90, by SKAN on CT41/D345 @ tiny.cc/createdib
		Local
		  If !VarSetCapacity(OSV) {
			 FileGetVersion, OSV, user32.dll
			 OSV := Format("{1:}.{2:}", StrSplit(OSV,".")*)
			}
		  LR_1   :=  0x2000|0x8|0x4              	; LR_CREATEDIBSECTION | LR_COPYDELETEORG | LR_COPYRETURNORG
		  LR_2   :=         0x8|0x4                  	; LR_COPYDELETEORG | LR_COPYRETURNORG
		  Flags  :=  ( OSV>6.3 ? (Gradient ? LR_2 : LR_1) : (Gradient ? LR_1 : LR_2) )
		  WB     :=  Ceil((W*3)/2)*2,    VarSetCapacity(BMBITS, WB*H, 0),      P := &BMBITS,  PE := P+(WB*H)

		  Loop, Parse, PixelData, |
			P := P<PE ? Numput("0x" . A_LoopField, P+0, "UInt")-(W & 1 && Mod(A_Index*3, W*3)=0 ? 0 : 1) : P

		  hBM := DllCall("CreateBitmap", "Int",W, "Int",H, "UInt",1, "UInt",24, "Ptr",0, "Ptr")
		  hBM := DllCall("CopyImage", "Ptr",hBM, "UInt",0, "Int",0, "Int",0, "UInt",LR_1, "Ptr")
				 DllCall("SetBitmapBits", "Ptr",hBM, "UInt",WB*H, "Ptr",&BMBITS)
		  hBM := DllCall("CopyImage", "Ptr",hBM, "Int",0, "Int",0, "Int",0, "Int",Flags, "Ptr")
		  hBM := DllCall("CopyImage", "Ptr",hBM, "Int",0, "Int",ResizeW, "Int",ResizeH, "UInt",Flags, "Ptr")
		  hBM := DllCall("CopyImage", "Ptr",hBM, "Int",0, "Int",0, "Int",0, "UInt",LR_2, "Ptr")
		Return DllCall("CopyImage", "Ptr",hBM, "Int",0, "Int",0, "Int",0, "UInt",DIB ? LR_1 : LR_2, "Ptr")
		}
}

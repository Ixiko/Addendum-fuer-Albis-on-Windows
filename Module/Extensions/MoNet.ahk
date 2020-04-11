;---------------------------------------------------------------------------------------------------------------------
;                                           MODUL: MoNet - MOnitor for your NETwork
;---------------------------------------------------------------------------------------------------------------------
    Monet_Version:= "0.4"
   ModulShort:= "M2"
;
;  for Addendum für AlbisOnWindows
;  written by Ixiko -this version is from 09.03.2018
;  please report errors: Ixiko@mailbox.org
;  use subject: "Addendum" so that you don't end up in the spam folder
;  GNU Lizenz - can be found in main directory  - 2017
;---------------------------------------------------------------------------------------------------------------------

;#TODO: MouseHover after reduced gui then ping one time


;{Script defaults
#NoEnv
#SingleInstance Force
SetWorkingDir %A_ScriptDir%
SetBatchLines, -1
SetBatchLines, 100000ms
ListLines, Off
DetectHiddenWindows, On
CoordMode, Pixel, Window
CoordMode, ToolTip, Window
CoordMode, Mouse, Window
FileEncoding, UTF-8

hIBitmap:= Create_Monet_ico(true)
Menu Tray, Icon, hIcon:  %hIBitmap%
;}

OnExit, Monet1GuiClose

;{includes
#include %A_ScriptDir%\..\..\include\Gdip_all.ahk
#include %A_ScriptDir%\..\..\include\AddendumFunctions.ahk
#include %A_ScriptDir%\..\..\include\Socket.ahk
;}

;{Variables and ini.read
;RegRead, AddendumDir, HKEY_LOCAL_MACHINE, SOFTWARE\Addendum für AlbisOnWindows, ApplicationDir
Global AddendumDir, hwndMonet1, yc                                                                                                        ;hwndMonet1 global für WM_MouseMove, yc ist die berechnete Guihöhe
AddendumDir := RegReadUniCode64("HKEY_LOCAL_MACHINE", "SOFTWARE\Addendum für AlbisOnWindows", "ApplicationDir")
MoMin:=0
CompNow:= A_ComputerName
NetDevice:= Object()
Result:= Object()

IniRead:
    IniRead, Computer, %AddendumDir%\Addendum.ini, Computer , Computer
    IniRead, StnFont, %AddendumDir%\Addendum.ini, %CompNow%, Monet_StandardFont, Futura Bk Bt
    IniRead, ColRow1, %AddendumDir%\Addendum.ini, %CompNow%, Monet_ColRow1, 234066
    IniRead, ColRow2, %AddendumDir%\Addendum.ini, %CompNow%, Monet_ColRow2, 0D1E38
    IniRead, MonetPos, %AddendumDir%\Addendum.ini, %CompNow%, Monet_PositionXY, 950|150

    Loop, Parse, Computer, `|
    {
        CompField:= A_LoopField
        IniRead, ip, %AddendumDir%\Addendum.ini, %CompField%, Ip
        IniRead, mac, %AddendumDir%\Addendum.ini, %CompField%, Mac
        IniRead, LastOn, %AddendumDir%\Addendum.ini, %CompField%, LastOnline, ....
        IniRead, PadStatus, %AddendumDir%\Addendum.ini, %CompField%, PadLock, 0                 ;0 means is unlocked
        NetDevice[A_Index]:= {"CompName" : CompField, "Ip" : ip, "Mac" : mac, "LastOnline" : LastOn, "Padlock" : PadStatus}
        CMax:= A_Index
    }

    LastOn:="", mac:="", ip:=""

    ;this is for the gui
        wMonet1 = 340
        hMonet1 = 400
        SysGet, CapSizeX, 30                ;SM_CXSIZE - Width
        SysGet, CapSizeY, 31                ;SM_CYSIZE - and height of a button in a window's caption or title bar, in pixels

    ;ping not to often
        pingtime:= 5*60*1000              ;5 minutes between two ping calls as a regular interval - below there is a function that calls the ping label at gui opening at mousehover
;}

;{MONET - MOnitor for your NETwork ----- Gui

        Gui Monet1: New, +HWNDhwndMonet1 -SysMenu -Caption +OwnDialogs +Resize ;+E0x00000100
        Gui Monet1: Font, s6 q5 cBlack, Futura Bk Md
        Gui Monet1: Color, c3F627F                  %ColRow%
        Gui Monet1: Margin, 0, 0
        Gui Monet1: Font, s10 q5 c00FF00, Futura Bk Md
        Gui Monet1: Add, Text, x6 y2 h14 BackgroundTrans HWNDhTitle , % "Praxisnetzwerk Monitor V" Monet_Version
        GuiControlGet, tp, Pos, Static1 ;ahk_id %hTitle%
        tpW += 108
        tpW += 42
            Gui Monet1: Font, s6 q5 cBlack, Futura Bk Md
        Gui Monet1: Add, Button, x%tpW% y3 w15 h15 0x8000 gMClick vBut3, X
            Gui Monet1: Font, s7 q5 cBlack, Futura Bk Md
        Gui Monet1: Add, Button, x205 y3 h15 0x8000 gMClick vBut1, Shutdown all clients

; build up Rows
    yc:= 20, s = 1          ;Titlebar height should be 20pixel

Loop, %CMax%
{
            If (s = 3) {              ;changes between the two defined colors
                s = 1
                }

            Gui, Monet1: Add, Progress, % "x0 y" yc " w" wMonet1 " h1"
            Gui, Monet1: Add, Picture   , % "x6    y" yc+1 " w48 h48 BackgroundTrans HWNDhComp" A_Index , assets\clientoffline.png

        ;PADLOCK

            If (NetDevice[A_Index].Padlock = 0) {
                        Gui, Monet1: Add, Picture   , % "x174 y" yc+4 " w24 h24 BackgroundTrans gMClick HWNDhCha" A_Index " vCha" A_Index, assets\changeS.png
            } else {
                        Gui, Monet1: Add, Picture   , % "x174 y" yc+4 " w24 h24 BackgroundTrans gMClick HWNDhCha" A_Index " vCha" A_Index, assets\nochange.png
            }

        ;COMMAND BUTTONS
                Gui, Monet1: Font, s6 q6 cWhite, Futura Bk Bt
            Gui, Monet1: Add, Button      , % "x204 y" yc+4 " w60 h18 0x8000 Center BackgroundTrans gMClick HWNDhWOL" A_Index " vWoL" A_Index, WAKE on LAN
                Gui, Monet1: Font, s8 q6 cWhite, Futura Bk Bt
           Gui, Monet1: Add, Button      , % "x204 y" yc+27 " w60 h18 0x8000 Center BackgroundTrans Disabled gMClick HWNDhSht" A_Index " vSht" A_Index, Shutdown
           Gui, Monet1: Add, Button      , % "x270 y" yc+4 " w60 h18 0x8000 Center BackgroundTrans Disabled gMClick HWNDhMes" A_Index " vMes" A_Index, Message
           Gui, Monet1: Add, Button      , % "x270 y" yc+27 " w60 h18 0x8000 Center BackgroundTrans Disabled gMClick HWNDhCmm" A_Index " vCmm" A_Index, Command

         ;STATUS INFORMATIONS
                    yc += 2
                Gui, Monet1: Font, s10 q6 cWhite, Futura Bk MD
            Gui, Monet1: Add, Text      , % "x54   y" yc " h12 BackgroundTrans HWNDhTextb" A_Index " vTextb" A_Index, % NetDevice[A_Index].CompName
                Gui, Monet1: Font, s8 q6 cWhite, Futura Bk Bt
                    yc += 17
            Gui, Monet1: Add, Text      , % "x54   y" yc " w100 h12 BackgroundTrans HWNDhTextc" A_Index " vTextc" A_Index, % "last online: " NetDevice[A_Index].LastOnline
                    yc += 13
            Gui, Monet1: Add, Text      , % "x54   y" yc " h12 HWNDhTextd" A_Index " vTextd" A_Index, % "Ip_adress: " NetDevice[A_Index].Ip
                    yc += 20, s ++

}

;locked clients must have a locked Shutdown button
Loop %CMax% {

    padStatus:= NetDevice[A_Index].Padlock
    if (padstatus = 1) {
            GuiControl, Monet1:Disable, Sht%A_Index%
        }

}

    padStatus:= ""
;dont show yet, let's arrange the windows first
CoordMode, Pixel, Screen
MonetPos:= StrSplit(MonetPos, "|")
MonetX:= MonetPos[1]
MonetY:= MonetPos[2]
Gui Monet1: Show, % "x" MonetX " y" MonetY " w" wMonet1 " h" yc, Monet
Gui Monet1: +MinSize%wMonet1%x%yc% +MaxSize%wMonet1%x1000
WinSet, AlwaysOnTop , On, ahk_id %hwndMonet1%
WinSet, Redraw,, ahk_id %hwndMonet1%
CoordMode, Pixel, Window
;}

gosub, pinger   ;for thirst ping

OnMessage(0x201,"WM_LBUTTONDOWN")
;OnMessage(0x200,"WM_MouseMove")
;OnMessage(0x4a, "Receive_WM_COPYDATA")

SetTimer, pinger, %pingtime%
SetTimer, CheckGui, 50


Return


;{Monet - Gui commands

MClick: ;{

    CName:= A_GuiControl

    If (CName = "But1") {
                Reload
    } else  If (CName = "But3") {
                ExitApp
    } else if Instr(CName, "Cha") {                             ;changing padlock status
        StringRight, padnum, CName, StrLen(CName)-3
        GuiControlGet, chaPos, Monet1:Pos, cha%padnum%
        chaPosX -=10, chaPosy -= 12

                If (NetDevice[padnum].Padlock = 0) {

                                GuiControl, Monet1: ,Cha%padnum%, assets\nochange.png
                                GuiControl, Monet1:Disable, Sht%padnum%
                                NetDevice[padnum].Padlock:= 1
                                ToolTip, Padlock off, %chaPosX%, %chaPosY%, 9
                                    SetTimer, TToff, -2000
                } else {
                                GuiControl, Monet1: , Cha%padnum%, assets\changeS.png
                                GuiControl, Monet1:Enable, Sht%padnum%
                                NetDevice[padnum].Padlock:= 0
                                ToolTip, Padlock on, %chaPosX%, %chaPosY%, 9
                                    SetTimer, TToff, -2000
                }

    } else if Instr(CName, "WOL") {
        StringRight, padnum, CName, StrLen(CName)-3
        GuiControlGet, chaPos, Monet1:Pos, cha%padnum%
        chaPosX -=15, chaPosy -= 24
        mac:= NetDevice[padnum].mac
        WakeOnLAN(mac)
        ToolTip, WakeOnLan command is send`nyou have to wait , %chaPosX%, %chaPosY%, 9
            SetTimer, TToff, -4000
    }


return
;}

Monet1GuiEscape:
Monet1GuiClose: ;{


    Loop, %CMax%
    {
        LastOn:= NetDevice[A_Index].LastOnline
        CompField:= NetDevice[A_Index].CompName
        PadStatus:= NetDevice[A_Index].Padlock
        If (CompField != "") {
                IniWrite, %lastOn%, %AddendumDir%\Addendum.ini, %CompField%, LastOnline
                IniWrite, %PadStatus%, %AddendumDir%\Addendum.ini, %CompField%, PadLock
            }
    }

     WinGetPos, wx, wy,,, ahk_id %hwndMonet1%
      If !ErrorLevel
            IniWrite, %wx%|%wy%, %AddendumDir%\Addendum.ini, %CompNow%, Monet_PositionXY

    DllCall("FreeConsole")
    Process Close, %pid%

ExitApp
;}
;}

TToff:
;{
        ToolTip,,,,9
return
;}

Pinger:
;{

    jo:=""

    Loop, %CMax%
    {

        ip:= NetDevice[A_Index].Ip
        RTT:= Ping4( ip, Result, 1024 )
        jo:= ErrorLevel

        CompField:= % hComp%A_Index%

        if (jo = 0) {   ;button change for client alive dependend command buttons
            GuiControl,Monet1:, %CompField%, assets\clientonline.png
                LastOn1 = % A_DD "`." A_MM "`." A_Year "`, " A_Hour ":" A_Min
                NetDevice[A_Index].LastOnline:= LastOn1
                GuiControl,Monet1:,Textc%A_Index%,                                      ;changing of text looks like overwriting text
                GuiControl,Monet1:,Textc%A_Index%, % "client is online"
            ;changes status of command buttons
                GuiControl, Monet1:Disable, WOL%A_Index%
                if (NetDevice[A_Index].Padlock = 1) {
                        GuiControl, Monet1:Disable, Sht%A_Index%
                } else {
                        GuiControl, Monet1:Enable, Sht%A_Index%
                }
                GuiControl, Monet1:Enable, Mes%A_Index%
                GuiControl, Monet1:Enable, Cmm%A_Index%
        } else {
                GuiControl,Monet1:, %CompField%, assets\clientoffline.png
                GuiControl,Monet1:,Textc%A_Index%,                                          ;changing of text looks like overwriting text
                laston:= NetDevice[A_Index].LastOnline
                GuiControl,Monet1:,Textc%A_Index%, % "last online: " laston
            ;changes status of command buttons
                GuiControl, Monet1:Disable, Sht%A_Index%
                GuiControl, Monet1:Disable, Mes%A_Index%
                GuiControl, Monet1:Disable, Cmm%A_Index%
                GuiControl, Monet1:Enable, WOL%A_Index%
        }

            sleep, 500
        jo:=""

    }

return
;}

CheckGui: ;{                                                                                is Mouse hovering the gui it will be maximized, if not it will be small

    MouseGetPos,,, oWin
    ;ToolTip % oWin "`n" oWin_old  "`n" HwndMonet1
        if Instr(oWin, HwndMonet1) And !MoMin {
               WinMove, Monet,,,,, yc
               MoMin:=1
        } else Instr(oWin, HwndMonet1) And MoMin {
                WinMove, Monet,,,,, 30
                MoMin:=0
        }


return
;}

;{ Functions
Ping4(Addr, ByRef Result := "", Timeout := 1024) {

   ; ===================================================================
   ; Function:       IPv4 ping with name resolution, based upon 'SimplePing' by Uberi ->
   ;                 http://www.autohotkey.com/board/topic/87742-simpleping-successor-of-ping/
   ; Parameters:     Addr     -  IPv4 address or host / domain name.
   ;                 ----------  Optional:
   ;                 Result   -  Object to receive the result in three keys:
   ;                             -  InAddr - Original value passed in parameter Addr.
   ;                             -  IPAddr - The replying IPv4 address.
   ;                             -  RTTime - The round trip time, in milliseconds.
   ;                 Timeout  -  The time, in milliseconds, to wait for replies.
   ; Return values:  On success: The round trip time, in milliseconds.
   ;                 On failure: "", ErrorLevel contains additional informations.
   ; Tested with:    AHK 1.1.22.03
   ; Tested on:      Win 8.1 x64
   ; Authors:        Uberi / just me
   ; Change log:     1.0.01.00/2015-07-16/just me - fixed bug on Win 8
   ;                 1.0.00.00/2013-11-06/just me - initial release
   ; MSDN:           Winsock Functions   -> http://msdn.microsoft.com/en-us/library/ms741394(v=vs.85).aspx
   ;                 IP Helper Functions -> hhttp://msdn.microsoft.com/en-us/library/aa366071(v=vs.85).aspx
   ; ===================================================================
   ; ICMP status codes -> http://msdn.microsoft.com/en-us/library/aa366053(v=vs.85).aspx
   ; WSA error codes   -> http://msdn.microsoft.com/en-us/library/ms740668(v=vs.85).aspx
   Static WSADATAsize := (2 * 2) + 257 + 129 + (2 * 2) + (A_PtrSize - 2) + A_PtrSize
   OrgAddr := Addr
   Result := ""
   ; -------------------------------------------------------------------------------------------------------------------
   ; Initiate the use of the Winsock 2 DLL
   VarSetCapacity(WSADATA, WSADATAsize, 0)
   If (Err := DllCall("Ws2_32.dll\WSAStartup", "UShort", 0x0202, "Ptr", &WSADATA, "Int")) {
      ErrorLevel := "WSAStartup failed with error " . Err
      Return ""
   }
   If !RegExMatch(Addr, "^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$") { ; Addr contains a name
      If !(HOSTENT := DllCall("Ws2_32.dll\gethostbyname", "AStr", Addr, "UPtr")) {
         DllCall("Ws2_32.dll\WSACleanup") ; Terminate the use of the Winsock 2 DLL
         ErrorLevel := "gethostbyname failed with error " . DllCall("Ws2_32.dll\WSAGetLastError", "Int")
         Return ""
      }
      PAddrList := NumGet(HOSTENT + 0, (2 * A_PtrSize) + 4 + (A_PtrSize - 4), "UPtr")
      PIPAddr   := NumGet(PAddrList + 0, 0, "UPtr")
      Addr := StrGet(DllCall("Ws2_32.dll\inet_ntoa", "UInt", NumGet(PIPAddr + 0, 0, "UInt"), "UPtr"), "CP0")
   }
   INADDR := DllCall("Ws2_32.dll\inet_addr", "AStr", Addr, "UInt") ; convert address to 32-bit UInt
   If (INADDR = 0xFFFFFFFF) {
      ErrorLevel := "inet_addr failed for address " . Addr
      Return ""
   }
   ; Terminate the use of the Winsock 2 DLL
   DllCall("Ws2_32.dll\WSACleanup")
   ; -------------------------------------------------------------------------------------------------------------------
   HMOD := DllCall("LoadLibrary", "Str", "Iphlpapi.dll", "UPtr")
   Err := ""
   If (HPORT := DllCall("Iphlpapi.dll\IcmpCreateFile", "UPtr")) { ; open a port
      REPLYsize := 32 + 8
      VarSetCapacity(REPLY, REPLYsize, 0)
      If DllCall("Iphlpapi.dll\IcmpSendEcho", "Ptr", HPORT, "UInt", INADDR, "Ptr", 0, "UShort", 0
                                            , "Ptr", 0, "Ptr", &REPLY, "UInt", REPLYsize, "UInt", Timeout, "UInt") {
         Result := {}
         Result.InAddr := OrgAddr
         Result.IPAddr := StrGet(DllCall("Ws2_32.dll\inet_ntoa", "UInt", NumGet(Reply, 0, "UInt"), "UPtr"), "CP0")
         Result.RTTime := NumGet(Reply, 8, "UInt")
      }
      Else
         Err := "IcmpSendEcho failed with error " . A_LastError
      DllCall("Iphlpapi.dll\IcmpCloseHandle", "Ptr", HPORT)
   }
   Else
      Err := "IcmpCreateFile failed to open a port!"
   DllCall("FreeLibrary", "Ptr", HMOD)
   ; -------------------------------------------------------------------------------------------------------------------
   If (Err) {
      ErrorLevel := Err
      Return ""
   }
   ErrorLevel := 0
   Return Result.RTTime
}

WakeOnLAN(mac) {
    magicPacket_HexString := GenerateMagicPacketHex(mac)
    size := CreateBinary(magicPacket_HexString, magicPacket)
    UdpOut := new SocketUDP()
    UdpOut.connect("addr_broadcast", 9)
    UdpOut.enableBroadcast()
    UdpOut.send(&magicPacket, size)
}

GenerateMagicPacketHex(mac) {
    magicPacket_HexString := "0xFFFFFFFFFFFF" ;64bit String
    Loop, 16
        magicPacket_HexString .= mac
    Return magicPacket_HexString
}

CreateBinary(hexString, ByRef var) { ;Credits to RHCP!
    sizeBytes := StrLen(hexString)//2
    VarSetCapacity(var, sizeBytes)
    Loop, % sizeBytes
        NumPut("0x" SubStr(hexString, A_Index * 2 - 1, 2), var, A_Index - 1, "UChar")
    Return sizeBytes
}

WM_LBUTTONDOWN(wParam,lParam,msg,hwnd) {
	Global
    Critical
			PostMessage, 0xA1, 2 ; WM_NCLBUTTONDOWN

}

WM_MouseMove() {

    static oWin_old

    MouseGetPos,,, oWin
    If (oWin_old<>oWin) {
        oWin_old:= oWin
        if (oWin=HwndMonet1) {
            WinMove, Monet,,,,, yc
        } else {
            WinMove, Monet,,,,, 30
        }
    }

}

WinSetParent(Hwnd, HParent=0, bFixStyle=false) {

	static WS_POPUP=0x80000000, WS_CHILD=0x40000000, WM_CHANGEUISTATE=0x127, UIS_INITIALIZE=3

	if (bFixStyle) {
		s1 := HParent ? "+" : "-", s2 := HParent ? "-" : "+"
		WinSet, Style, %s1%%WS_CHILD%, ahk_id %Hwnd%
		WinSet, Style, %s2%%WS_POPUP%, ahk_id %Hwnd%
	}
	r := DllCall("SetParent", "Ptr", Hwnd, "uint", HParent, "Uint")
	ifEqual, r, 0, return 0
	SendMessage, WM_CHANGEUISTATE, UIS_INITIALIZE,,,ahk_id %HParent%
	return r
}

Create_Monet_ico(NewHandle := False) {
Static hBitmap := 0
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "AAABAAEAMDAAAAEAGACoHAAAFgAAACgAAAAwAAAAYAAAAAEAGAAAAAAAAAAAAGAAAABgAAAAAAAAAAAAAAAAAAAAAAD/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo8AAAAAAAAAAAD/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/1lJX/oo//oo//oo//oo8AAAD/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/9oY7/oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/9oY7/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdSMyHRhHLij3z0m4jZiXe/eWZCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdaNybXiHXjj33WiHW3c2FCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBfMgW77n4z7n4zplIHmkn/7n4z/oo+8dmRKLRxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfEfGn7n4zmkn/WiHXnk4D6n4zCe2hCKBdC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KBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdGKhr/oo//oo//oo//oo9FKhlCKBdCKBf/oo//oo//oo//oo//oo//oo//oo/9oY7/oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo8AAAD/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo8AAAAAAAAAAAD/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo8AAAAAAADAAAAAAAMAAIAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAAAAQAAwAAAAAADAAA="
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
VarSetCapacity(Dec, DecLen, 0)
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
DllCall("Gdiplus.dll\GdipCreateHICONFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
Return hBitmap
}
;}




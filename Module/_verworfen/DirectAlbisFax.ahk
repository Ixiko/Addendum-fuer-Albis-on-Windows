
;#######################################################
;##																																		##
;##					ADDENDUM für AlbisOnWindows - Extensionmodul - Direct Albisfax 				##
;##												Scriptversion siehe unten													##
;##														by Ixiko 2018															##
;##																																		##
;#######################################################

												  DirectAlbisFax:= "V0.1" 			
												  ;Datum 26.01.2018

/* Funktionsbeschreibung

DirectAlbisFax - ein Modul um direct nach Eingabe eines Hotstrings (definiert im Praxomat_st.ahk Script)
nach Eingabe von /fax in der Patientenakte öffnet sich ein Fenster mit der Auswahl von Faxnummern
und ein Eingabefeld zur Eingabe von Text. Geplant ist eine Drag und Drop oder zumindestens eine
FileSelect-Fenster um bis zu 3 Dateien BMP, JPEG; PDF oder XML anzuhängen. Gefaxt wird dann über die
web API von fax.de - kein lästiges Ausdrucken mehr oder umständliches Drucken mit Umweg über den
virtuellen Druckertreiber. Alternativ Aufruf der Fax GUI über Hotkey oder meiner noch nicht fertigen 
Menueinbindung direkt in das Albisprogramm.
Der eingeschriebene Text wird in die Akte geschrieben samt Rufnummer zu Dokumentationszwecken oder
zur späteren Verwendung.

Thank you, that you allow me for stealing your GUI - dear Jack Dunning (http://www.computoredge.com/AutoHotkey/)

*/
;{ <><><><><><><><><><>Ablaufeinstellungen des Scriptes<><><><><><><>
	;#NoTrayIcon
	#NoEnv
	#SingleInstance, Force
	#Persistent
	SetTitleMatchMode, 2
	DetectHiddenWindows, On
	DetectHiddenText, On
	SetBatchLines -1            ; Script nicht ausbremsem (Default: "sleep, 10" alle 10 ms) 
	SetControlDelay, -1         ; Wartezeit beim Zugriff auf langsame Controls abschalten 
	SetWinDelay, -1             ; Verzögerung bei allen Fensteroperationen abschalten 
	CoordMode, Tooltip, Screen
	CoordMode, Mouse, Window
	CoordMode, Pixel, Screen
	CoordMode, Menu, Screen
	
;}
RegRead, AddendumDir, HKEY_LOCAL_MACHINE, SOFTWARE\Addendum für AlbisOnWindows, ApplicationDir
IniFile = %AddendumDir%\Praxomat.ini

IniRead, F1Sz, %IniFile%, DCCo, F1Sz
IniRead, F2Sz, %IniFile%, DCCo, F2Sz
IniRead, F3Sz, %IniFile%, DCCo, F3Sz
IniRead, Font, %IniFile%, DAFx, Font
IniRead, FColor, %IniFile%, DAFx, FColor
IniRead, BgColor, %IniFile%, DAFx, BgColor
HexColor:= rgb2Hex(BgColor)

Gui +AlwaysOnTop +Resize
Gui Color, %HexColor%
Gui Font, s%F3Sz% q5 cBlack, %Font%
Gui Margin, 10,10
Gui, Add, ListView, sort -ReadOnly r10 w1000 vMyListView gMyListview, Nachname | Vorname | E-mail | Telephone | Fax | Straße | Stadt | Land | PLZ 
Gui, Add, Edit, Section w100 xs vAddr2, Vorname
Gui, Add, Edit, w150 ys vAddr1 -Background, Nachname
Gui, Add, Edit, w200 ys vAddr3, Fax
Gui, Add, Edit, w100 ys vAddr4, Telefon
Gui, Add, Edit, w50 ys Limit5 number vAddr6, PLZ
Gui, Add, Edit, w50 ys vAddr7, Stadt
Gui, Add, Edit, w100 ys vAddr8, PLZ
Gui, Add, Button, ys gAddItem, zur Liste hinzufügen
Gui Font, s%F1Sz% q5 cWhite, %Font%
Gui, Add, GroupBox, xm ys+25 Section w520 h190 Center, Hier zu sendenden Text eingeben
Gui Font, s%F3Sz% q5 cBlack, %Font%
Gui, Add, Edit, xp+10 yp+25 Section w500 r10 vText1 
Gui, Add, Button, x+M gAddItem, Anhang hinzufügen
SizeCols()

Menu, MycontextMenu, Add, Adresse einfügen, InsertItem
Menu, MyContextMenu, Add, Ändern, EditItem
Menu, MyContextMenu, Add, Löschen, DeleteItem
Menu, Tray, Add, Show Address Book, ZeigeAdressbuch


ZeigeAdressbuch:
  Gui, Show, AutoSize, Address Book
Return

MyListView:
  If A_GuiEvent = e          ;Finished editing first field of a row
    LV_ModifyCol(1,"Sort")
  If A_GuiEvent = ColClick   ;Clicked column header
    {
      If A_EventInfo = 10    ;Number of column header clicked
        LV_ModifyCol(9,"SortDesc")
    }
    UpdateFile()
Return

GuiContextMenu:  ; Launched in response to a right-click or press of the Apps key.
if A_GuiControl <> MyListView  ; Display the menu only for clicks inside the ListView.
    Return
  LV_GetText(ColText, A_EventInfo,1)    ;Gather column data in string EditText
  EditText := ColText
  Loop 9
    {
      LV_GetText(ColText, A_EventInfo, A_Index+1)
      EditText := EditText . "|" . ColText
    }
Menu, MyContextMenu, Show, %A_GuiX%, %A_GuiY%
return

InsertItem:        ;Adds formatted name and address to any document
    StringSplit, RowData, EditText , |
    SendInput, !{Escape}
    SendInput, %RowData2% %RowData1%`n%RowData5%`n%RowData6%, %RowData7% %RowData8%`n`n
Return

EditItem:          ;Move row from ListView columns into edit fields
  SelectedRow := LV_GetNext()
  StringSplit, RowData, EditText , |
  Loop, 9
    {
      GuiControl, ,Addr%A_Index%, % Rowdata%A_Index%
    }
  GuiControl, ,Button1, Update
Return

DeleteItem:         ;Deletes selected rows
MsgBox, 4100, Adresse löschen?, Adresse löschen? Wähle Ja oder Nein?
IfMsgBox No    ;Don't delete
  Return
RowNumber = 0  ; This causes the first iteration to start the search at the top.
Loop
{
    ; Since deleting a row reduces the RowNumber of all other rows beneath it,
    ; subtract 1 so that the search includes the same row number that was previously
    ; found (in case adjacent rows are selected):
    RowNumber := LV_GetNext(RowNumber - 1)
    if not RowNumber  ; The above returned zero, so there are no more selected rows.
        break
    LV_Delete(RowNumber)  ; Clear the row from the ListView.
}
UpdateFile()
Return

AddItem:                 ;Add new or update ListView row
  Gui, Submit, NoHide
  If Addr9 =             ; Setup Column 10 date format
    Addr10 := ""
  Else
    FormatTime, Addr10, %Addr9%, dddd MMMM d, yyyy


If SelectedRow = 0
  {
    LV_Add("", Trim(Addr1),Trim(Addr2),Trim(Addr3),Trim(Addr4)
     ,Trim(Addr5),Trim(Addr6),Trim(Addr7),Trim(Addr8), Trim(Addr9), Addr10)
  }
else
  {
    LV_Modify(SelectedRow,"", Addr1,Addr2,Addr3,Addr4
        ,Addr5,Addr6,Addr7,Addr8,Addr9,Addr10)
    LV_ModifyCol(1,"Sort")
    SelectedRow := 0
    GuiControl, ,Button1, Add to list
  }

SizeCols()
UpdateFile()
Return

GuiEscape:
GuiClose:
UpdateFile:  ;When exiting
  DetectHiddenWindows On
  UpdateFile()
  ExitApp
Return



GuiSize:  ; Widen or narrow the ListView in response to the user's resizing of the window.
if A_EventInfo = 1  ; The window has been minimized.  No action needed.
    return
; Otherwise, the window has been resized or maximized. Resize the ListView to match.
GuiControl, Move, MyListView, % "W" . (A_GuiWidth - 20)

Return

SizeCols() { ;resize all columns, hide Column 9, Column 10 NoSort
  
     Loop, 8
       {
          LV_ModifyCol(A_Index,"AutoHdr")
       }
       LV_ModifyCol(9, 0)
       LV_ModifyCol(10,"AutoHdr")
       LV_ModifyCol(10,"NoSort")
  }

#IfWinActive, Address Book ; CTRL+A to Select All
^a::LV_Modify(0,"Select")
#IfWinActive

UpdateFile() {   ;Saves data to file when ListView updated
  
    FileDelete, AddressBook.txt
    WinGet, Min, MinMax, Address Book
    If Min = -1
      WinRestore, Address Book
    WinGetPos, X, Y, Width, Height, Address Book
    Width -= 16
    Height -= 38
    FileAppend, x%x% y%y% w%Width% h%Height% `n, AddressBook.txt
    Loop % LV_GetCount()
     {
         RowNum := A_Index
         Loop, 8
           {
             LV_GetText(Text, RowNum, A_Index)
             TrimText := Trim(Text)
             FileAppend, "%TrimText%"`,, AddressBook.txt
           }
             LV_GetText(Text, RowNum, 9)
             TrimText := Trim(Text)
             FileAppend, "%TrimText%"`n, AddressBook.txt
     }
   }
   
rgb2Hex(s, d = "") {

	StringSplit, s, s, % d = "" ? "," : d

	SetFormat, Integer, % (f := A_FormatInteger) = "D" ? "H" : f

	h := s1 + 0 . s2 + 0 . s3 + 0

	SetFormat, Integer, %f%

	Return, "0x" . RegExReplace(RegExReplace(h, "0x(.)(?=$|0x)", "0$1"), "0x")

}
 
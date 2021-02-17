;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Author:         Jack Dunning, ceeditor@computoredge.com
;
; Script Function: Basic Address Book
;	Second version includes Age calculation
;

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

OnExit, UpdateFile

SelectedRow := 0

Gui +AlwaysOnTop +Resize
Gui, Add, ListView, sort -ReadOnly r10 w800 vMyListView gMyListview, Last Name
     | First Name | E-mail | Telephone | Street | City | State | Zip | | Birthday
Gui, Add, Button, section gAddItem, Add to List
Gui, Add, Text, w75 section , First Name
Gui, Add, Edit, w100 ys vAddr2, First Name
Gui, Add, Text, ys, Last Name
Gui, Add, Edit, w150 ys vAddr1, Last Name
Gui, Add, Text, w75 xm section, E-mail Address
Gui, Add, Edit, w200 ys vAddr3, E-mail Address
Gui, Add, Text, ys, Telephone
Gui, Add, Edit, w100 ys vAddr4, Telephone
Gui, Add, Text, w75 xm section, Street
Gui, Add, Edit, w200 ys vAddr5, Street
Gui, Add, Text, w75 xm  section, City
Gui, Add, Edit, w150 ys vAddr6, City
Gui, Add, Text, ys, State
Gui, Add, Edit, w50 ys vAddr7, State
Gui, Add, Text, ys, Zipcode
Gui, Add, Edit, w100 ys vAddr8, Zipcode
Gui, Add, Text, w75 xm section, Birthday
Gui, Add, DateTime, ys vAddr9 ChooseNone, LongDate


IfExist, AddressBook.txt ;Add data from AddressBook.txt to ListView
{
 FileCopy, AddressBook.txt, AddressBook%A_Now%.txt  ;incremental backup
 Loop, Read, AddressBook.txt
  {
  If (A_index = 1 and SubStr(A_LoopReadLine, 1, 1) = "x")
     {
       WinPos := A_LoopReadLine
       Continue
     }
  Else
     {
     Loop, Parse, A_LoopReadLine , CSV
       {
          RowData%A_Index% := A_LoopField
       }
       If RowData9 =           ; Setup Column 10 date format
        RowData10 := ""
       Else
        FormatTime, RowData10, %RowData9%, dddd MMMM d, yyyy

       LV_Add("", RowData1,RowData2,RowData3,RowData4
             ,RowData5,RowData6,RowData7,RowData8,RowData9,RowData10)
     }
  }
}


SizeCols()

Menu, MycontextMenu, Add, Insert Address, InsertItem
Menu, MyContextMenu, Add, Send E-mail, EmailName
Menu, MyContextMenu, Add, How old?, Age
Menu, MyContextMenu, Add, Edit, EditItem
Menu, MyContextMenu, Add, Delete, DeleteItem
Menu, Tray, Add, Show Address Book, ShowAddressBook

IfExist, AddressBook.txt
  {
     Gui, Show, %WinPos% , Address Book
  }
Else
  {
     WinGetPos,X1,Y1,W1,H1,Program Manager
     X2 := W1-800
     Gui, Show, x%x2% y50 , Address Book
  }


Hotkey, ^!a, ShowAddressBook
Return

ShowAddressBook:
  Gui, Show,, Address Book
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

EmailName:          ;Opens new e-mail window with To: address
   StringSplit, RowData, EditText , |
   Run, mailto:%RowData2% %RowData1%<%RowData3%>?subject=This is the subject line
                   &body=This is the message body's text.
Return

Age:              ;Calculates age
   StringSplit, RowData, EditText , |
   HowOld(RowData9,A_Now)
   MsgBox, %RowData2% %RowData1% Born: %RowData10%`n%Years% Years %Months% Months %Days% Days
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
MsgBox, 4100, Delete Address?, Delete Address? Click Yes or No?
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


UpdateFile()    ;Saves data to file when ListView updated
  {
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

SizeCols()  ;resize all columns, hide Column 9, Column 10 NoSort
  {
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

HowOld(FromDay,ToDay)   ;Age calculation function
  {
    FromDay := substr(FromDay,1,8)
    ToDay := Substr(ToDay,1,8)
    Global Years,Months,Days

; If born on February 29

    If SubStr(FromDay,5,4) = 0229 and Mod(SubStr(ToDay,1,4), 4) != 0 and SubStr(ToDay,5,4) = 0228
	PlusOne = 1

 ThisMonth := SubStr(ToDay,1,6)

; Set ThisMonthLength equal to next month

    ThisMonthLength := % SubStr(ToDay,5,2) = "12" ? SubStr(ToDay,1,4)+1 . "01"
           : SubStr(ToDay,1,4) . Substr("0" . SubStr(ToDay,5,2)+1,-1)

; Days in this month saved in ThisMonthLength

    EnvSub, ThisMonthLength, %ThisMonth%, d

; Set ThisMonthday to FromDay or  (if FromDay higher) last day of this month

    If SubStr(FromDay,7,2) > ThisMonthLength
        ThisMonthDay :=  ThisMonth . ThisMonthLength
    Else
        ThisMonthDay :=  ThisMonth . SubStr(FromDay,7,2)

; Calculate last month's length

    LastMonthLength := % SubStr(ToDay,5,2) = "01" ? SubStr(ToDay,1,4)-1 . "12"
           : SubStr(ToDay,1,4) . Substr("0" . SubStr(ToDay,5,2)-1,-1)
    LastMonth := LastMonthLength

; Days in last month saved in LastMonthLength

    EnvSub, LastMonthLength, %ThisMonth% ,d
    LastMonthLength := LastMonthLength*(-1)



; Set LastMonthday to FromDay or (if FromDay higher) last day of last month

    If SubStr(FromDay,7,2) > LastMonthLength
        LastMonthDay :=  LastMonth . LastMonthLength
    Else
        LastMonthDay :=  LastMonth . SubStr(FromDay,7,2)


; Calculate years

    Years  := % SubStr(ToDay,5,4) - SubStr(FromDay,5,4) < 0 ? SubStr(ToDay,1,4)-SubStr(FromDay,1,4)-1
            : SubStr(ToDay,1,4)-SubStr(FromDay,1,4)

; Calculate months


    Months := % SubStr(ToDay,5,2)-SubStr(FromDay,5,2) < 0 ? SubStr(ToDay,5,2)-SubStr(FromDay,5,2)+12
            : SubStr(ToDay,5,2)-SubStr(FromDay,5,2)

    Months := % SubStr(ToDay,7,2) - SubStr(ThisMonthDay,7,2) < 0 ? Months -1 : Months
    Months := % Months = -1 ? 11 : Months

; Calculate days


    TodayDate := SubStr(ToDay,1,8)          ; Remove any time portion of stamp
    EnvSub, ThisMonthDay,ToDayDate , d
    EnvSub, LastMonthDay,ToDayDate , d

    Days  := % ThisMonthDay <= 0 ? -1*ThisMonthDay : -1*LastMonthDay

; If February 28

    Years := % plusone = 1 ? Years +1 : Years
    days := % plusone = 1 ? 0 : days

    If (TodayDate <= FromDay)
	Years := 0, Months := 0,Days := 0

;   MsgBox, %Years% Years, %Months% Months, %Days% Days


  }
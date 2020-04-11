OnMessage(0x4a, "Receive_WM_COPYDATA")  ; 0x4a is WM_COPYDATA


Receive_WM_COPYDATA(wParam, lParam) {

    StringAddress := NumGet(lParam + 8)  ; lParam+8 is the address of CopyDataStruct's lpData member.
    StringLength := DllCall("lstrlen", UInt, StringAddress)
    if StringLength <= 0
        ToolTip %A_ScriptName%`nA blank string was received or there was an error.
    else
    {
        VarSetCapacity(CopyOfData, StringLength)
        DllCall("lstrcpy", "str", CopyOfData, "uint", StringAddress)  ; Copy the string out of the structure.
        ; Show it with ToolTip vs. MsgBox so we can return in a timely fashion:
        ToolTip %A_ScriptName%`nReceived the following string:`n%CopyOfData%
    }
    return true  ; Returning 1 (true) is the traditional way to acknowledge this message.
}

; Save the following script as "Sender.ahk" then launch it.  After that, press the Win+Space hotkey.
;TargetScriptTitle = Receiver.ahk ahk_class AutoHotkey
;result := Send_WM_COPYDATA(StringToSend, TargetScriptTitle)
;if result = FAIL
;    MsgBox SendMessage failed. Does the following WinTitle exist?:`n%TargetScriptTitle%
;else if result = 0
;    MsgBox Message sent but the target window responded with 0, which may mean it ignored it.
;return

Send_WM_COPYDATA(ByRef StringToSend, ByRef TargetScriptTitle) { ; ByRef saves a little memory in this case.
; This function sends the specified string to the specified window and returns the reply.
; The reply is 1 if the target window processed the message, or 0 if it ignored it.
    VarSetCapacity(CopyDataStruct, 12, 0)  ; Set up the structure's memory area.
    ; First set the structure's cbData member to the size of the string, including its zero terminator:
    NumPut(StrLen(StringToSend) + 1, CopyDataStruct, 4)  ; OS requires that this be done.
    NumPut(&StringToSend, CopyDataStruct, 8)  ; Set lpData to point to the string itself.
    Prev_DetectHiddenWindows := A_DetectHiddenWindows
    Prev_TitleMatchMode := A_TitleMatchMode
    DetectHiddenWindows On
    SetTitleMatchMode 2
    SendMessage, 0x4a, 0, &CopyDataStruct,, %TargetScriptTitle%  ; 0x4a is WM_COPYDATA. Must use Send not Post.
    DetectHiddenWindows %Prev_DetectHiddenWindows%  ; Restore original setting for the caller.
    SetTitleMatchMode %Prev_TitleMatchMode%         ; Same.
    return ErrorLevel  ; Return SendMessage's reply back to our caller.
}
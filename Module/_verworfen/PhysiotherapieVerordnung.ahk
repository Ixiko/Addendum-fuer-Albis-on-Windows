; This script was created using Pulover's Macro Creator
; www.macrocreator.com

#NoEnv
SetWorkingDir %A_ScriptDir%
CoordMode, Mouse, Window
SendMode Input
#SingleInstance Force
SetTitleMatchMode 2
#WinActivateForce
;#NoTrayIcon
SetControlDelay 1
SetWinDelay 0
SetKeyDelay -1
SetMouseDelay -1
SetBatchLines -1
#Persistent


#F5::
Macro1:
WinActivate, ALBIS ahk_class OptoAppClass
;WinWaitActive, ALBIS -* ahk_class OptoAppClass
Sleep, 50
SendRaw, {fhv13}
Send, {Tab}
;ControlSend, Edit3, {f}, ALBIS -* ahk_class OptoAppClass
;ControlSend, Edit3, {h}, ALBIS -* ahk_class OptoAppClass
;ControlSend, Edit3, {v}, ALBIS -* ahk_class OptoAppClass
;ControlSend, Edit3, {Tab}, ALBIS -* ahk_class OptoAppClass
WinWait, ahk_class #32770 ahk_exe albisCS.exe
WinActivate, ahk_class #32770 ahk_exe albisCS.exe
Send, {Tab}
Send, {Tab}
SendRaw, {bb}
Send, {Tab}
Send, {LShift} & {F3}
Send, {LShift Up}
Send, {Tab}
Send, {Tab}
Send, {Tab}
Send, {Tab}
SendRaw, {Myo}
Send, {Tab}
Click, 38, 669 Left, Down
Click, 38, 669 Left, Up
ControlSend, Edit22, {F3}, ahk_class #32770 ahk_exe albisCS.exe*
WinActivate, CGM HEILMITTELKATALOG ahk_class Qt5QWindowIcon
Click, 113, 613 Left, Down
Click, 113, 613 Left, Up
Click, 138, 617 Left, Down
Click, 138, 617 Left, Up
Click, 108, 614 Left, Down
Click, 108, 614 Left, Up
Click, 198, 783 Left, Down
Click, 198, 783 Left, Up
Click, 865, 846 Left, Down
Click, 865, 846 Left, Up
Click, 38, 456 Left, Down
Click, 38, 456 Left, Up
Click, 79, 474 Left, Down
Click, 79, 474 Left, Up
Click, 1006, 700 Left, Down
Click, 1006, 700 Left, Up
Click, 141, 720 Left, Down
Click, 141, 720 Left, Up
Click, 44, 722 Left, Down
Click, 44, 722 Left, Up
Click, 109, 692 Left, Down
Click, 109, 692 Left, Up
Click, 107, 744 Left, Down
Click, 107, 744 Left, Up
Click, 1004, 747 Left, Down
Click, 1004, 747 Left, Up
Click, 40, 766 Left, Down
Click, 40, 766 Left, Up
Click, 354, 672 Left, Down
Click, 354, 672 Left, Up
Click, 877, 841 Left, Down
Click, 877, 841 Left, Up
Click, 37, 328 Left, Down
Click, 37, 328 Left, Up
Click, 81, 332 Left, Down
Click, 81, 332 Left, Up
Click, 69, 412 Left, Down
Click, 69, 412 Left, Up
Click, 35, 694 Left, Down
Click, 35, 694 Left, Up
Click, 80, 706 Left, Down
Click, 80, 706 Left, Up
Click, 69, 784 Left, Down
Click, 69, 784 Left, Up
Click, 848, 838 Left, Down
Click, 848, 838 Left, Up
WinActivate, ahk_class #32770 ahk_exe albisCS.exe*
Click, 96, 1004 Left, Down
Click, 96, 1004 Left, Up
Return


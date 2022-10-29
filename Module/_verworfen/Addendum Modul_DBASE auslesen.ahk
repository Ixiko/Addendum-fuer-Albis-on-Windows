
#NoEnv

global AddendumDir
AddendumDir := RegReadUniCode64("HKEY_LOCAL_MACHINE", "SOFTWARE\Addendum für AlbisOnWindows", "ApplicationDir")
hPatBase:= DBase_OpenDBF("M:\AlbisWin\_DB\patient.dbf")

InputBox, Benutzereingabe, Lassen Sie uns was suchen, Geben Sie mal einen NAmen ein, , 640, 480
if ErrorLevel
{
    MsgBox, Sie haben CANCEL gedrückt.
} else {
     FilledArray = DllStructCreate("DWORD[100]")
     recCnt = Search(hPatBase, Benutzereingabe ."*", -1, DllStructGetPtr(FilledArray), DllStructGetSize(FilledArray))

	MsgBox, 1, Ich habe folgendes gefunden., %FilledArray%
}



DBase_CloseDBF(hPatBase)


#include %A_ScriptDir%\..\..\include\dbase.ahk
#include %A_ScriptDir%\..\..\include\AddendumFunctions.ahk
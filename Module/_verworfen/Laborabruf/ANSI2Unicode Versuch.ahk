;ANSI to Unicode Versuch
SetBatchLines, -1







;FileRead, LDTDaten, %A_ScriptDir%\Labordaten\15.04.2018-X01CLEST.LDT
FileRead, LDTDaten, % "M:\Praxis\Skripte\Skripte Neu\FindName\extracts\1000\00009.txt"
;MsgBox, % StrLen(LDTDaten)
;CP-850 sieht bisher am besten aus
Loop, Parse, LDTDaten, `n, `r
{

	zeile:= Ansi2UTF8(A_LoopField) . "`n"
	FileAppend, %zeile%, %A_ScriptDir%\LDTUTF8text.txt
   ToolTip, % A_LoopField "`n" zeile
   Sleep 1500

}


ExitApp

Ansi2Oem(sString) {
   Ansi2Unicode(sString, wString, 0)
   Unicode2Ansi(wString, zString, 1)
   Return zString
}

Oem2Ansi(zString) {
   Ansi2Unicode(zString, wString, 1)
   Unicode2Ansi(wString, sString, 0)
   Return sString
}

Ansi2UTF8(sString) {
   Ansi2Unicode(sString, wString, 850)
   Unicode2Ansi(wString, zString, 65001)
   Return zString
}

UTF82Ansi(zString) {
   Ansi2Unicode(zString, wString, 65001)
   Unicode2Ansi(wString, sString, 0)
   Return sString
}

Ansi2Unicode(ByRef sString, ByRef wString, CP = 0) {
     nSize := DllCall("MultiByteToWideChar"
      , "Uint", CP
      , "Uint", 0
      , "Uint", &sString
      , "int",  -1
      , "Uint", 0
      , "int",  0)

   VarSetCapacity(wString, nSize * 2)

   DllCall("MultiByteToWideChar"
      , "Uint", CP
      , "Uint", 0
      , "Uint", &sString
      , "int",  -1
      , "Uint", &wString
      , "int",  nSize)
}

Unicode2Ansi(ByRef wString, ByRef sString, CP = 0) {
     nSize := DllCall("WideCharToMultiByte"
      , "Uint", CP
      , "Uint", 0
      , "Uint", &wString
      , "int",  -1
      , "Uint", 0
      , "int",  0
      , "Uint", 0
      , "Uint", 0)

   VarSetCapacity(sString, nSize)

   DllCall("WideCharToMultiByte"
      , "Uint", CP
      , "Uint", 0
      , "Uint", &wString
      , "int",  -1
      , "str",  sString
      , "int",  nSize
      , "Uint", 0
      , "Uint", 0)
}
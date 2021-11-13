ComObjError(true) ; Zeige uns COM Fehler an

xl := ComObjActive("Excel.Application") ;// startet eine Excel Instanz
xl.Visible := 1

WinActivate, ahk_exe Excel.exe
;sheet  := excel.ActiveWorkbook  ;// aktuell geöffnete Datei

i=1

Loop	{
			i++
			Cell:= "D" . i
			CellValue := xl.Range(Cell).Value
				if CellValue=""
						break
			;StringSplit, split, Cellvalue, %A_Space%
			;pos:=RegExMatch(CellValue, "\d{1,3}\x2F\d{2}", RNr)
			;pos:=RegExMatch(CellValue, "b[\d]{1,3}\x2Fb[\d]{2}", RNr)
			pos:=RegExMatch(CellValue, "\b[\d]{1,3}\x2F\b[\d]{2}", RNr)
				If Rnr=""
					pos:=RegExMatch(CellValue, "(?!Nr. )\d{5}(?!\d)(?!\s)", RNr)
			;ListVars
				If (Rnr<>"") {
				Cellnew:= "C" . i
				xl.Range(CellNew).Value := RNr
			}

}


#^e::ExitApp

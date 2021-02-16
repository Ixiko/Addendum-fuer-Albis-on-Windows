; * Client Label für Addendum.ahk - Editieren Sie hier für die Verwendung für ihre PC's im Netzwerk

; Client - Labels haben den Namenszusatz des jeweiligen Netzwerkclients auf dem sie laufen sollen

;*********************************************************************************************************
; 			PC Sprechzimmer 1 WS | PC Sprechzimmer 1 WS | PC Sprechzimmer 1 WS
;*********************************************************************************************************;{
nopopup_SP1WS:	;{
	noTrayOrphansWin10()
	;HELP Taskbar für Praxomat
	IfWinActive, ahk_class OptoAppClass
		AlbisHotKeyHilfe(Addendum.Help, "")
return ;}
seldom_SP1WS: ;{
	If ( TimeDiff("000000", "now", "m")=0 )
			gosub SkriptReload
return ;}
;}

;*********************************************************************************************************
;	 		PC Sprechzimmer 2 | PC Sprechzimmer 2 | PC Sprechzimmer 2 | PC Sprechzimmer 2
;*********************************************************************************************************;{
nopopup_SP2:	;{
	noTrayOrphansWin10()
	IfWinActive, ahk_class OptoAppClass
			AlbisHotKeyHilfe(Addendum.Help, "")
return ;}
seldom_SP2: ;{
	SetTimer, seldom_SP2, Off
return ;}
;}

;*********************************************************************************************************
;         PC Anmeldung  |  PC Anmeldung  |  PC Anmeldung  |  PC Anmeldung  |  PC Anmeldung
;*********************************************************************************************************;{
nopopup_AnmeldungPC: ;{
	noTrayOrphansWin10()
	IfWinActive, ahk_class OptoAppClass
		AlbisHotKeyHilfe(Addendum.Help, "")
return ;}
seldom_AnmeldungPC: ;{
	If (TimeDiff("000000", "now", "m") = 0)
				gosub SkriptReload
return ;}


;}

;*********************************************************************************************************
;         PC Labor  |  PC Labor  |  PC Labor  |  PC Labor  |  PC Labor  |  PC Labor  |  PC Labor
;*********************************************************************************************************;{
nopopup_LABOR: ;{
	noTrayOrphansWin10()
	IfWinActive, ahk_class OptoAppClass
		AlbisHotKeyHilfe(Addendum.Help, "")
return
;}
seldom_Labor: ;{
	SetTimer, seldom_Labor, Off
return
;}
;}

;*********************************************************************************************************
; 		EKG/Lufu PC | EKG/Lufu PC | EKG/Lufu PC | EKG/Lufu PC | EKG/Lufu PC | EKG/Lufu PC
;*********************************************************************************************************;{
nopopup_EKGPC: ;{
	noTrayOrphansWin10()
	IfWinActive, ahk_class OptoAppClass
		AlbisHotKeyHilfe(Addendum.Help, "")
return
;}
seldom_EKGPC:	;{
	SetTimer, seldom_EKGPC, Off
return ;}
;}

TimeDiff(time1, time2="now", output="m") {                                 ; errechne die Zeitdifferenz zwischen Zeit1 und Zeit2, Ausgabe in Minuten (output="m") oder Sekunden (output="s")

	ListLines, Off

	if Instr(time2,"now") {
			time2:= A_Now
			FormatTime, time2,, HHmmss
	}

	if Instr(time1,"000000")
			time1:= "240000"

	if Instr(output,"m")
			timediff:= ( SubStr(time1,1,2)*60 + SubStr(time1,3,2) ) - ( SubStr(time2,1,2)*60 + SubStr(time2,3,2) )
	else if Instr(output,"s")
			timediff:= ( SubStr(time1,1,2)*3600 + SubStr(time1,3,2) + SubStr(time1,5,2) ) - ( SubStr(time2,1,2)*3600 + SubStr(time2,3,2)*60 + SubStr(time2,5,2) )


return timediff
}



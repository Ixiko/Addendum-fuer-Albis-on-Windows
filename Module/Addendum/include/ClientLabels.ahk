; * Client Label für Addendum.ahk - Editieren Sie hier für die Verwendung für ihre PC's im Netzwerk

; Client - Labels haben den Namenszusatz des jeweiligen Netzwerkclients auf dem sie laufen sollen

;*********************************************************************************************************
; 			PC Sprechzimmer 1 WS | PC Sprechzimmer 1 WS | PC Sprechzimmer 1 WS
;*********************************************************************************************************;{
nopopup_SP1WS:	;{
	noTrayOrphansWin10()
	;HELP Taskbar für Praxomat
	IfWinActive, ahk_class OptoAppClass
		AlbisHotKeyHilfe(AddendumHelp, "")
return
;}

seldom_SP1WS: ;{

	If ( TimeDiff("000000", "now", "m")=0 ) {
						gosub SkriptReload
		}

return
;}


;}

;*********************************************************************************************************
;	 		PC Sprechzimmer 2 | PC Sprechzimmer 2 | PC Sprechzimmer 2 | PC Sprechzimmer 2
;*********************************************************************************************************;{
nopopup_SP2:	;{
	noTrayOrphansWin10()
	;HELP Taskbar für Praxomat
	IfWinActive, ahk_id %AlbisWinID%
			AlbisHotKeyHilfe(AddendumHelp, PraxomatHelp)
return
;}
seldom_SP2: ;{
	SetTimer, seldom_SP2, Off
return
;}
;}

;*********************************************************************************************************
;         PC Anmeldung  |  PC Anmeldung  |  PC Anmeldung  |  PC Anmeldung  |  PC Anmeldung
;*********************************************************************************************************;{
nopopup_AnmeldungPC: ;{
	noTrayOrphansWin10()
	IfWinActive, ahk_id %AlbisWinID%
		AlbisHotKeyHilfe(AddendumHelp, "")
return

;}
seldom_AnmeldungPC: ;{
	SetTimer, seldom_AnmeldungPC, Off
return
;}
Wartezeit_AnmeldungPC: ;{



return
;}

;}

;*********************************************************************************************************
;         PC Labor  |  PC Labor  |  PC Labor  |  PC Labor  |  PC Labor  |  PC Labor  |  PC Labor
;*********************************************************************************************************;{
nopopup_LABOR: ;{
	noTrayOrphansWin10()
	IfWinActive, ahk_id %AlbisWinID%
		AlbisHotKeyHilfe(AddendumHelp, "")
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
	IfWinActive, ahk_id %AlbisWinID%
		AlbisHotKeyHilfe(AddendumHelp, "")
return
;}

seldom_EKGPC:
SetTimer, seldom_EKGPC, Off
return
;}


/*
	;wenn Wochentag=Mittwoch und es zwischen 8-12 Uhr ist, wird stündlich die Errinnerung gezeigt
	If ( (A_WDay = Mittwoch) && (A_HOUR >7) && (A_HOUR <13) && (!A_HOUR == KofferAuffuellFlag) ) {

				KofferAuffuellflag:= A_HOUR
				;die letzte Anzeigestunde wird in der ini gespeichert
				IniWrite, %KofferAuffuellflag%, %AddendumDir%\Addendum.ini, FUNKTION, KofferAuffuellflag
				KofferAuffuellen()

	} else if ( (!KofferAuffuellflag == "NA") && (A_WDAY <> 4) ) {

				KofferAuffuellflag:= "NA"
				IniWrite, %KofferAuffuellflag%, %AddendumDir%\Addendum.ini, FUNKTION, KofferAuffuellflag

	}
*/

/*
		;blendet zum Cave Fenster Informationen ein
		If (CaveVonID:=WinExist("Cave! von"))  {
					If (CTTExist=0)
							AlbisCaveVonToolTip(compname, CaveVonID)
		} else {
				If (CTTExist=1) {
							CTTExist=0
							ToolTip,,,, 10
				}
		}
*/
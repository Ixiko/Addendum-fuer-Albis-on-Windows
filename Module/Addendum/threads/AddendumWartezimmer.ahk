; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                 	Addendum Wartezimmer
;
;      Funktion:           	meldet Veränderungen in den Wartezimmern ohne das Wartezimmer anzeigen zu müssen
;									zeigt bisherige Wartezeiten, An- und Abmeldungen und die Zahl der aktuell wartenden Patienten an
;
;		Hinweis:
;
;		Abhängigkeiten:	siehe includes
;
;	                    			Addendum für Albis on Windows
;                        			by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;									begin: 29.03.2021,	last change 29.03.2021
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

; Einstellungen                                                                         	;{
	#NoEnv
	#Persistent
	#KeyHistory, Off

	SetBatchLines                  	, -1
	ListLines                        	, Off
	SetWinDelay                    	, -1
	SetControlDelay            	, -1
	AutoTrim                       	, On
	FileEncoding                 	, UTF-8
;}

; Variablen Albis Datenbankpfad / Addendum Verzeichnis      	;{

	global Props, PatDB, adm

  ; Pfad zu Albis und Addendum ermitteln
	RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
	RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)

  ; Addendum Objekt
	adm                 	:= Object()
	adm.Dir            	:= AddendumDir
	adm.Ini              	:= AddendumDir "\Addendum.ini"
	adm.DBPath      	:= AddendumDir "\logs'n'data\_DB"
	;adm.AlbisPath    	:= AlbisPath
	adm.AlbisPath    	:= "E:"
	adm.AlbisDBPath 	:= adm.AlbisPath "\DB"
	adm.compname	:= StrReplace(A_ComputerName, "-")                                                                 	; der Name des Computer auf dem das Skript läuft

;}

	Inspektor := new Wartezimmer(adm.Ini, "CImpfung")
	Inspektor.StartTimer(10)

return

Esc::ExitApp
^!+::Reload

class Wartezimmer {

		__New(admIniPath, wzName, options:="") {

			; Addendum.ini erweitern (nur erster Start)
				Section := "Wartezimmer_"

			; Einstellungen laden
				this.compname	:= StrReplace(A_ComputerName, "-")
				this.wzName 	:= wzName
				this.workini     	:= IniReadExt(admIniPath)
				this.interval    	:= IniReadExt(adm.compname, Section "Timer", 60)
				this.winsize    	:= IniReadExt(adm.compname, Section "Fensterposition", "Snap Albis w200 h200")


		}


	;----------------------------------------------------------------------------------------------------------------------------------------------
	; timer methods
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		StartTimer(timerB:="") {	                                                                            	;

		; starting the handler function
			this.timer := ObjBindMethod(this, "GetUpdates")
			timer := this.timer
			SetTimer, % timer, % (timerB ? timerB*1000 : this.interval*1000)
			PraxTT("Timer gestartet...", "3 1")

		; get first updates before starting the timer
			this.GetUpdates()

		}

		StopTimer() {	                                                                            	; stop

			timer := this.timer
			SetTimer % timer, Off
			RestoreTimer := this.RestoreTimer
			SetTimer, % RestoreTimer, Delete

		}

		SetUpdatesInterval(new_updateInterval) {                                     	; set update interval

				If (StrLen(new_updateInterval) > 0) {
					this.interval := new_updateInterval
					timer := this.timer
					SetTimer, % timer, % (this.interval * 1000)
				}

		}

		;}

		GetUpdates() {                                                                           	;

			nHour	:= A_Hour
			nMin 	:= A_Min
			Today	:= A_DD "." A_MM "." A_YYYY
			tmp  	:= AlbisWZListe(this.wzName)
			wz	 	:= this.ParseTable(tmp)
			anwesende := cnt := TimeSum := 0
			maxWT := 0


			For idx, Pat in wz {

				If (StrLen(Pat.Anwesenheit) > 0)
					wzStats[Pat.Anwesenheit] += 1

				if (today = Pat.Datum) && RegExMatch(Pat.Anwesenheit, "(Angemeldet|In_Behandlung)") && RegExMatch(Pat.Wartezeit, "\d+\:\d+"){
					anwesende ++
					wtime := StrSplit(Pat.Wartezeit, ":")
					wmin := wTime.1*60 + wtime.2
					maxWT := wmin > maxWT ? wmin : maxWT
					TimeSum += wmin
				}
			}

			atoffice := wzStats.AH_Anwesend + wz.Stats.AH_In_Behandlung
			wMinutes := TimeSum / atoffice
			wTime := "⌀ " Floor(wMinutes/60) ":" SubStr("0" (wMinutes-Floor(wMinutes/60)), -1) " (h:m)"


			cnt ++
			t := "wartend: " anwesende "`n-------------------------"
			t.= "`n" wTime "`nmax.WZ: " maxWT "`nZähler: " cnt
			ToolTip, % t, 1400, 1, 17


		}

		GetUpdates_O() {                                                                           	;

			wzStats := Object()
			wzStats	:= {	"AH_Ohne_Status":0  	, "AH_Abwesend":0   	, "AH_Angemeldet":0
							, 	"AH_In_Behandlung":0	, "AH_Erreichbar":0	, 	"Gerufen":0}

			nHour	:= A_Hour
			nMin 	:= A_Min
			Today	:= A_DD "." A_MM "." A_YYYY
			tmp  	:= AlbisWZListe(this.wzName)
			wz	 	:= this.ParseTable(tmp)
			cnt := TimeSum := 0
			maxWT := Object()

			For idx, Pat in wz {

				If (StrLen(Pat.Anwesenheit) > 0)
					wzStats["AH_" . Pat.Anwesenheit] += 1

				if (today = Pat.Datum) && RegExMatch(Pat.Anwesenheit, "(Anwesend|In_Behandlung)") {
					wtime := StrSplit(Pat.Wartezeit, ":")
					TimeSum += wTime.1*60 + wtime.2
					SciTEOutput(Pat.Wartezeit)
				}
			}

			atoffice := wzStats.AH_Anwesend + wz.Stats.AH_In_Behandlung
			wMinutes := TimeSum / atoffice
			wTime := "⌀ " Floor(wMinutes/60) ":" Mod(wMinutes, 60) " (h:m)"

			For key, val in wzStats {
				key := StrReplace(key, "AH_")
				key := StrReplace(key, "_", " ")
				t .= key ":" SubStr("       ", 1, 13-StrLen(key)) "`t" val "`n"
			}

			cnt ++
			t.= "`n" wTime "`nMinutes: " TimeSum "`nZähler: " cnt
			ToolTip, % t, 1400, 1, 17


		}

		ParseTable(obj) {                                                                     		; parsed die Tabellendaten noch einmal

			wz := Array()
			For idx, line in obj {

				wz.Push(Object())
				For key, val in line {

					If RegExMatch(key, "i)Name") {
						RegExMatch(val, "^(?<Nn>.*)?,\s*(?<Vn>.*)?\((?<ID>\d+),\s*(?<Gt>\w+)", Pat_)
						wz[idx].ID	:= Pat_ID
						wz[idx].Nn	:= Pat_Nn
						wz[idx].Vn	:= Pat_Vn
						wz[idx].Gt	:= Pat_Gt

					}
					else If RegExMatch(key, "i)Geburtsdatum") {

						RegExMatch(val, "^(?<Gd>\d\d\.\d\d.\d\d\d\d)", Pat_)
						wz[idx].Gd := Pat_Gd

					}
					else If RegExMatch(key, "i)Anwesenheit") {

						RegExMatch(val, "^(?<AH>\pL+)?(\s+\(|$)", Pat_)
						val := StrReplace(val, " ", "_")
						wz[idx].Anwesenheit := Pat_AH

					} else
						wz[idx][key] := val

				}
			}

		return wz
		}

}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Includes
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
#Include %A_ScriptDir%\..\..\..\include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_DATUM.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_PDFHelper.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_PraxTT.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\..\include\Addendum_Window.ahk

#Include %A_ScriptDir%\..\..\..\lib\acc.ahk
#Include %A_ScriptDir%\..\..\..\lib\class_JSON.ahk
#Include %A_ScriptDir%\..\..\..\lib\class_Neutron.ahk
#include %A_ScriptDir%\..\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\..\lib\ini.ahk
#include %A_ScriptDir%\..\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\..\lib\Sift.ahk
;}
/*
	IniReadExt(admIniPath)
				val := IniReadExt(Section)
				SciTEOutput(val)
				If !(StrLen(val) = 0 || InStr(val, "ERROR"))
					If FilePathCreate(adm.Dir "\_backup") {
						If IniInsertSection(admIniPath, "B--", "Wartezimmer", adm.Dir "\_backup")
							PraxTT("neue Ini-Sektion angelegt", "5 1")
						else
							PraxTT("Fehler beim erstmaligen Anlegen neuer Ini-Einträge", "5 1")
				}

*/

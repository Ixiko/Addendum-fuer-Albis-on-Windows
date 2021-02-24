; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                                                     ALBIS REANIMATOR
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;		Albis läßt sich manchmal nicht stabil betreiben,   z.B. weil die Verbindung zur Ifap - Medikamentendatenbank  verloren gegangen ist.
;		Dann beendet man per Hand zumeist Albis  und muss es anschliessend neu starten.   Wenn sich der Neustart  hinzieht oder Albis im
;		Startvorgang hängen bleibt, greift man über den Taskmanager ein und beendet dort die Prozesse.
;		Denn häufig werden Programme wie javaw.exe , Ifap.exe , bubblehostmanagersystem.exe  oder  Minervahostsystem.exe,  nach  dem
; 		nach dem Beenden von Albis weiterhin ausgeführt werden.  Derzeit sehe  ich viele  Probleme mit der  Ifap.exe die  außerdem für ihre
;		Ausführung Java benötigt. So kommt es eben vor das die javaw.exe nach Beendigung oder Absturz von Ifap nicht mit beendet wurde.
;		So sehe ich schon bis zu 4 Javaprozesse für ein laufendes Ifap im Taskmanager gesehen.
;
; 		Und hier setzt der Reanimator an.   Er beendet zunächst alle Prozesse die mit Albis ausgeführt wurden und ruft erst danach Albis auf.
;		Es kann leider nicht die Verbindung zwischen Ifap und Albis ohne Neustart herstellen. Auf die Beseitigung dieses Problems von Seiten
;		der Compugroup warte ich schon seit Jahren.
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;      	Funktion: 		 	    ▹	schließt alle laufenden Albisprozesse inklusive der zugehörigen von Albis gestarteten Hintergrundprozesse wie:
;                       		    	albis.exe, albisCS.exe, javaw.exe, wkflsr.exe, bubblemanagerhostsystem.exe und MinervaHostsystem.exe
;
;									▹	das lokale sowie das Albis Stammverzeichnis, ebenso wie die auszuführende Albis-Programmvariante werden
;										aus der Windows-Registry ausgelesen. Sollte dies nicht funktionieren, können alle 3 Werte auch optional der
;										Funktion übergeben werden. Die Funktionsparameter sind dominant! Dies bedeutet das die aus der Registry
;										erhaltenen Werte dann nicht für den Albisaufruf verwendet werden. Da die Funktionsparameter optional sind
;										kann man auch eine Kombination aus Registry-Werten und Funktionsparametern benutzen.
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;      	Basisskript:    		 	keines
;		Abhängigkeiten: 	 	siehe #includes
;
;	                    				Addendum für Albis on Windows
;                        				by Ixiko started in September 2017 - last change 21.02.2021 - this file runs under Lexiko's GNU Licence
;										proof-of-concept version
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

AlbisReanimator()
return

AlbisReanimator(AlbisMainPath:="", AlbisLocalPath:="", AlbisExe:="") {

		closed := "Albisprozesse                           `tgeschlossen`n"
		AlbisPIDs := AlbisGetRunningPIDs()
		For procNr, proc in AlbisPIDs {
			btt("Process " proc.Name " wird geschlossen",,,, "Style4")
			Process, Close    	, % proc.PID
			geschlossen := (ErrorLevel ? "ja":"nein")
			len := 27-StrLen(proc.Name)
			len := len > 5 ? len*2 : len
			closed .= proc.Name SubStr("                                       ",1, len) "`t" geschlossen "`n"
		}

		MsgBox, 1028	, % "Addendum Albis-Reanimator"
								, %  closed "`n"
								. 	  "Albis jetzt neu starten?"
		btt()
		IfMsgBox, No
			return

		albis := GetAlbisPaths()
		If StrLen(AlbisMainPath) > 0
			albis.MainPath := AlbisMainPath
		If StrLen(AlbisLocalPath) > 0
			albis.MainPath := AlbisLocalPath
		If StrLen(AlbisExe) > 0
			albis.Exe := AlbisExe

		Run, %  albis.MainPath "\" albis.Exe, % albis.LocalPath
		WinWait, % "ahk_class OptoAppClass"

}

AlbisGetRunningPIDs() {

	PIDs := 0
	AlbisPIDs := []
	static AlbisProcs := [{"exe": "Ifap"                                      	, "cmdline":""}
								, {"exe": "javaw"                                 	, "cmdline":"CG\Java"}
								, {"exe": "albis"                                      	, "cmdline":""}
								, {"exe": "wkflsr"                                  	, "cmdline":""}
								, {"exe": "bubblemanagerhostsystem"  	, "cmdline":""}
								, {"exe": "MinervaHostsystem"             	, "cmdline":""}]

	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
		for procNr, proc in AlbisProcs
			 If InStr(process.Name, proc.exe) && InStr(Process.CommandLine, proc.cmdLine)
				AlbisPIDs.Push({"name":process.Name, "PID" : process.ProcessID})

return AlbisPIDs
}

WMIEnumProcessID(ProcessSearched, cmdLineSearched:="") {

	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
	   If InStr(process.Name, ProcessSearched) && InStr(Process.CommandLine, cmdLineSearched)
			return Process.ProcessID

return 0
}

GetAlbisPaths() {

	If (A_PtrSize = 8)
		SetRegView	, 64
	else
		SetRegView	, 32

	RegRead, MainPath	, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-MainPath
	RegRead, LocalPath 	, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-LocalPath
	RegRead, Exe        	, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-Exe

	If (StrLen(MainPath) = 0)
		throw " Registry Eintrag des Albis 'MainPath' konnte nicht gelesen werden"
	If (StrLen(LocalPath) = 0)
		throw " Registry Eintrag des Albis 'LocalPath' konnte nicht gelesen werden"
	If (StrLen(Exe) = 0)
		throw " Registry Eintrag der Albis-'Exe' konnte nicht gelesen werden"

return {"MainPath":MainPath, "LocalPath":LocalPath, "Exe":Exe}
}

#include %A_ScriptDir%\..\..\lib\class_BTT.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk













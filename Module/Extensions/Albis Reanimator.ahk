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
;
;		Anmerkung:			▹	lassen sich die Prozesse nicht wie gewünscht mit diesem Skript schließen, dann fehlen dem Skript zumeist die
;										Rechte. Versuchen Sie das Skript mit Administrator-Berechtigung auszuführen oder beenden Sie die
;										verbliebenen Prozesse über den Windows Taskmanager.
;										siehe: https://www.autohotkey.com/boards/viewtopic.php?t=56155
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;      	Basisskript:    		 	keines, nicht als Funktionsbibliothek gedacht
;		Abhängigkeiten: 	 	siehe #includes
;
;	                    				Addendum für Albis on Windows
;                        				by Ixiko started in September 2017 - last change 28.02.2021 - this file runs under Lexiko's GNU Licence
;										proof-of-concept version
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	#NoEnv
	SetBatchLines, -1

	global adm

	RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)
	adm := AddendumBaseProperties(AddendumDir)

	; auf false setzen wenn man keine Logdatei führen will
	adm.ReanimatorLog := true

	Menu, Tray, Icon, % adm.Dir "\assets\ModulIcons\AlbisReanimator.ico"

	AlbisReanimator()

ExitApp

AlbisReanimator(AlbisMainPath:="", AlbisLocalPath:="", AlbisExe:="") {                    	;-- beendet alle Prozesse und startet auf Wunsch Albis neu

		AlbisProzesse := Object()

	; Prozesse finden und schliessen
		procView := ""
		For procNr, proc in AlbisGetRunningPIDs() {

			; Information
				btt(procView . "Process " proc.Name "[" proc.PID "] wird beendet", A_ScreenWidth-500, 20,, "Style4")

			; Versuch per Process WaitClose
				Process, WaitClose, % proc.PID, 5
				Process, Exist, % proc.PID
				ProcExist	:= ErrorLevel

			; Versuch per cmd.exe
				If ProcExist {
					Run, % comspec " /k taskkill /F /IM " proc.Name " && exit ", Hide
					while ProcExist && (A_Index <= 50) { ; max 10 Sekunden
						Sleep 200
						Process, Exist, % proc.PID
						ProcExist	:= ErrorLevel
					}
				}

			; Auflisten
				If !AlbisProzesse.haskey(proc.Name)
					AlbisProzesse[proc.Name] := {"count": 1, "closed": (ProcExist ? 0 : 1)}
				else  {
					AlbisProzesse[proc.Name].count 	+= 1
					AlbisProzesse[proc.Name].closed	+= (ProcExist ? 0 : 1)
				}

			; Fortschritt anzeigen
				procView .=  "Process " proc.Name "[" proc.PID "]" (ProcExist ? " konnte nicht beendet werden" : " wurde beendet") "`n"
				btt(procView, A_ScreenWidth-500, 20,, "Style4")

		}

	; Text-Ausgaben vorbereiten
		If (AlbisProzesse.Count() > 0) {
			closed	:= "Albisprozesse                   `tgefunden/beendet`n"
			logText	:= A_YYYY "-" A_MM "-" A_DD " " A_Hour ":" A_Min ":" A_Min "`n"
			For procName, data in AlbisProzesse {
				len := 27 - StrLen(proc.Name)
				len := len > 4 ? len*2 : len
				closed 	.= procName SubStr("                                       ",1, len) "`t" data.count "/" data.closed "`n"
				logText .= procName " [" data.count "/" data.closed "], "
			}
		}
		else
			closed := "Keine laufenden Albisprozesse gefunden."

	; Logdatei Verzeichnis anlegen, vorbereiten und schreiben
		If adm.ReanimatorLog && (AlbisProzesse.Count() > 0) {

			If !InStr(FileExist(adm.DBPath "\sonstiges"), "D")
				FileCreateDir, % adm.DBPath "\sonstiges"
			If !FileExist(adm.DBPath "\sonstiges\Albis Reanimator Log.txt") {
				FileAppend, %	"----------------------------------------------------------------------------------------------------------------------------------------`n"
								.		"Logdatei für das Dokumentieren von laufenden Albis- und zugehörigen Prozessen durch AlbisReanimator.ahk`n"
								.	  	"aufgelistet werden ein Zeitstempel und in der Zeile darunter die Namen der geschlossenen Prozesse.`n"
								.	  	"Die erste Zahl in der eckigen Klammer ist die Anzahl der gefunden Prozesse mit gleichem Namen, die zweite`n"
								.	  	"Zahl sind die entsprechenden erfolgreichen Prozeßbeendigungen`n"
								.	  	"----------------------------------------------------------------------------------------------------------------------------------------`n"
								, % adm.DBPath "\sonstiges\Albis_Reanimator_Log.txt", UTF-8
			}

			FileAppend, % RTrim(logText, ", ") "`n", % adm.DBPath "\sonstiges\Albis_Reanimator_Log.txt", UTF-8

		}

	; Albis neu starten
		MsgBox, 1028	, % "Addendum Albis-Reanimator", % closed "`n`n" (AlbisProzesse.Count() > 0 ? "Albis erneut starten?" : "Albis starten?")
		IfMsgBox, Yes
		{

			albis := GetAlbisPaths()
			albis.MainPath 			:= StrLen(AlbisMainPath) > 0	? AlbisMainPath	: albis.MainPath
			albis.AlbisLocalPath 	:= StrLen(AlbisLocalPath) > 0 	? AlbisLocalPath	: albis.AlbisLocalPath
			albis.Exe                	:= StrLen(AlbisExe) > 0         	? AlbisExe         	: albis.Exe

			Run, %  albis.MainPath "\" albis.Exe, % albis.LocalPath
			WinWait, % "ahk_class OptoAppClass"

		}

	; Beautiful ToolTip schliessen
		btt()

return
}

AlbisGetRunningPIDs() {                                                                                        	;-- sammelt die ProcessIDs von Albis und abhängigen Prozessen ein

	PIDs := 0
	AlbisPIDs := []
	static AlbisProcs := [{"exe": "Ifap"                                    	, "cmdline":""}
								, {"exe": "ipc"                                     	, "cmdline":""}
								, {"exe": "javaw"                                 	, "cmdline":"CG\Java"}
								, {"exe": "albis"                                   	, "cmdline":""}
								, {"exe": "wkflsr"                                 	, "cmdline":""}
								, {"exe": "bubblemanagerhostsystem"  	, "cmdline":""}
								, {"exe": "MinervaHostsystem"             	, "cmdline":""}]

	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
		for procNr, proc in AlbisProcs
			 If InStr(process.Name, proc.exe) && InStr(Process.CommandLine, proc.cmdLine)
				AlbisPIDs.Push({"name":process.Name, "PID" : process.ProcessID})

return AlbisPIDs
}

AddendumBaseProperties(AddendumDir) {                                                               	;-- Basiseinstellungen für jedes Addendum-Modul

	; Hauptverzeichnis und ini Datei
		props             	:= Object()
		props.Dir         	:= AddendumDir
		props.Ini          	:= props.Dir "\Addendum.ini"

	; Log Dateipfad und Datenbankverzeichnis
		workini := IniReadExt(props.Ini)
		props.LogPath	:= IniReadExt("Addendum", "AddendumLogPath"	,, true)
		props.DBPath  	:= IniReadExt("Addendum", "AddendumDBPath" 	,, true)

	; Log Dateipfad prüfen und evtl. anlegen
		If !RegExMatch(props.LogPath, "i)[A-Z]\:\\")
			props.LogPath := props.Dir "\logs'n'data"
		If !InStr(FileExist(props.LogPath), "D")
			FileCreateDir, % props.LogPath

	; Datenbankverzeichnis prüfen und evtl. anlegen
		If !RegExMatch(props.DBPath	, "i)[A-Z]\:\\")
			props.DBPath := props.Dir "\logs'n'data\_DB"
		If !InStr(FileExist(props.DBPath), "D")
			FileCreateDir, % props.DBPath

return props
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

; ab hier Funktionen für ini-Read und AddendumBaseProperties
IniReadExt(SectionOrFullFilePath, Key:="", DefaultValue:="", convert:=true) {            	;-- eigene IniRead funktion für Addendum

	; beim ersten Aufruf der Funktion !nur! Übergabe des ini Pfades mit dem Parameter SectionOrFullFilePath
	; die Funktion behandelt einen geschriebenen Wert der "ja" oder "nein" ist, als Wahrheitswert, also true oder false
	; UTF-16 in UTF-8 Zeichen-Konvertierung
	; Der Pfad in Addendum.Dir wird einer anderen Variable übergeben. Brauche dann nicht immer ein globales Addendum-Objekt
	; letzte Änderung: 31.01.2021

		static admDir
		static WorkIni

	; Arbeitsini Datei wird erkannt wenn der übergebene Parameter einen Pfad ist
		If RegExMatch(SectionOrFullFilePath, "^[A-Z]\:.*\\")	{
			If !FileExist(SectionOrFullFilePath)	{
				MsgBox,, % "Addendum für AlbisOnWindows", % "Die .ini Datei existiert nicht!`n`n" WorkIni "`n`nDas Skript wird jetzt beendet.", 10
				ExitApp
			}
			WorkIni := SectionOrFullFilePath
			If RegExMatch(WorkIni, "[A-Z]\:.*?AlbisOnWindows", rxDir)
				admDir := rxDir
			else
				admDir := Addendum.Dir
			return WorkIni
		}

	; Workini ist nicht definiert worden, dann muss das komplette Skript abgebrochen werden
		If !WorkIni {
			MsgBox,, Addendum für AlbisOnWindows, %	"Bei Aufruf von IniReadExt muss als erstes`n"
																			. 	"der Pfad zur ini Datei übergeben werden.`n"
																			.	"Das Skript wird jetzt beendet.", 10
			ExitApp
		}

	; Section, Key einlesen, ini Encoding in UTF.8 umwandeln
		IniRead, OutPutVar, % WorkIni, % SectionOrFullFilePath, % Key
		If convert
			OutPutVar := StrUtf8BytesToText(OutPutVar)

	; Bearbeiten des Wertes vor Rückgabe
		If InStr(OutPutVar, "ERROR")
			If (StrLen(DefaultValue) > 0) { ; Defaultwert vorhanden, dann diesen Schreiben und Zurückgeben
				OutPutVar := DefaultValue
				IniWrite, % DefaultValue, % WorkIni, % SectionOrFullFilePath, % key
				If ErrorLevel
					TrayTip, % A_ScriptName, % "Der Defaultwert <" DefaultValue "> konnte geschrieben werden.`n`n[" WorkIni "]", 2
			}
			else return "ERROR"
		else if InStr(OutPutVar, "%AddendumDir%")
				return StrReplace(OutPutVar, "%AddendumDir%", admDir)
		else if RegExMatch(OutPutVar, "%\.exe$") && !RegExMatch(OutPutVar, "i)[A-Z]\:\\")
				return GetAppImagePath(OutPutVar)
		else if RegExMatch(OutPutVar, "i)^\s*(ja|nein)\s*$", bool)
				return (bool1= "ja") ? true : false

return Trim(OutPutVar)
}

StrUtf8BytesToText(vUtf8) {                                                                                       	;-- Umwandeln von Text aus .ini Dateien in UTF-8
	if A_IsUnicode 	{
		VarSetCapacity(vUtf8X, StrPut(vUtf8, "CP0"))
		StrPut(vUtf8, &vUtf8X, "CP0")
		return StrGet(&vUtf8X, "UTF-8")
	} else
		return StrGet(&vUtf8, "UTF-8")
}

GetAppImagePath(appname) {                                                                                 	;-- Installationspfad eines Programmes

	headers:= {	"DISPLAYNAME"                  	: 1
					,	"VERSION"                         	: 2
					, 	"PUBLISHER"             	         	: 3
					, 	"PRODUCTID"                    	: 4
					, 	"REGISTEREDOWNER"        	: 5
					, 	"REGISTEREDCOMPANY"    	: 6
					, 	"LANGUAGE"                     	: 7
					, 	"SUPPORTURL"                    	: 8
					, 	"SUPPORTTELEPHONE"       	: 9
					, 	"HELPLINK"                        	: 10
					, 	"INSTALLLOCATION"          	: 11
					, 	"INSTALLSOURCE"             	: 12
					, 	"INSTALLDATE"                  	: 13
					, 	"CONTACT"                        	: 14
					, 	"COMMENTS"                    	: 15
					, 	"IMAGE"                            	: 16
					, 	"UPDATEINFOURL"            	: 17}

   appImages := GetAppsInfo({mask: "IMAGE", offset: A_PtrSize*(headers["IMAGE"] - 1) })
   Loop, Parse, appImages, "`n"
	If Instr(A_loopField, appname)
		return A_loopField

return ""
}

GetAppsInfo(infoType) {

	static CLSID_EnumInstalledApps := "{0B124F8F-91F0-11D1-B8B5-006008059382}"
        , IID_IEnumInstalledApps     	:= "{1BC752E1-9046-11D1-B8B3-006008059382}"

        , DISPLAYNAME            	:= 0x00000001
        , VERSION                    	:= 0x00000002
        , PUBLISHER                  	:= 0x00000004
        , PRODUCTID                	:= 0x00000008
        , REGISTEREDOWNER    	:= 0x00000010
        , REGISTEREDCOMPANY	:= 0x00000020
        , LANGUAGE                	:= 0x00000040
        , SUPPORTURL               	:= 0x00000080
        , SUPPORTTELEPHONE  	:= 0x00000100
        , HELPLINK                     	:= 0x00000200
        , INSTALLLOCATION     	:= 0x00000400
        , INSTALLSOURCE         	:= 0x00000800
        , INSTALLDATE              	:= 0x00001000
        , CONTACT                  	:= 0x00004000
        , COMMENTS               	:= 0x00008000
        , IMAGE                        	:= 0x00020000
        , READMEURL                	:= 0x00040000
        , UPDATEINFOURL        	:= 0x00080000

   pEIA := ComObjCreate(CLSID_EnumInstalledApps, IID_IEnumInstalledApps)

   while DllCall(NumGet(NumGet(pEIA+0) + A_PtrSize*3), Ptr, pEIA, PtrP, pINA) = 0  {
      VarSetCapacity(APPINFODATA, size := 4*2 + A_PtrSize*18, 0)
      NumPut(size, APPINFODATA)
      mask := infoType.mask
      NumPut(%mask%, APPINFODATA, 4)

      DllCall(NumGet(NumGet(pINA+0) + A_PtrSize*3), Ptr, pINA, Ptr, &APPINFODATA)
      ObjRelease(pINA)
      if !(pData := NumGet(APPINFODATA, 8 + infoType.offset))
         continue
      res .= StrGet(pData, "UTF-16") . "`n"
      DllCall("Ole32\CoTaskMemFree", Ptr, pData)  ; not sure, whether it's needed
   }
   Return res
}

#include %A_ScriptDir%\..\..\lib\class_BTT.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk













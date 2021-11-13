;###############################################################
;------------------------------------------------------------------------------------------------------------------------------------
;----------------------------------------------- Addendum für AlbisOnWindows ---------------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------
;----------------------------------------------- 	Modul:     	ScanPool            	----------------------------------------------
;----------------------------------------------- 	Skript:      	ScanPoolServer  	----------------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------
;{ 					überwacht den Befundordner und erstellt automatisch eine aktuelle Indexdatei

	; Dieses Skript ist Teil des ScanPool Modules.

	; Zweck dieses Skriptes ist es unnötige Lese und Schreibzugriffe auf die Festplatte(n) zu vermeiden, welche
	; den Befund oder PDF-Datei Ordner enthalten. Dazu überwacht es selbstständig den in der Addendum.ini
	; hinterlegten Befundordner und indiziert aufgetretene Veränderungen.
	; Neue PDF-Dokumente werden erfasst, Seitenzahl, Größe und der enthaltene Text werden indiziert.
	; Gelöschte oder in andere Verzeichnisse kopierte Dateien werden aus dem Index entfernt.

	; Das Skript kann auf irgend einem Computer im Hintergrund laufen, welcher üblicherweise
	; während des Praxisbetriebes in Benutzung ist. Dabei muss er nicht dem Rechner der scannt und auch nicht
	; auf dem Computer sein welcher den Befundordner bereitstellt. Das Skript kann ohne größere Eingriffe
	; auch die Daten von Netzwerkfestplatten (NTFS-Format) von überall auslesen. Es ist auch nicht notwendig
	; dieses Skript auf dem Server laufen zu lassen. Wichtig ist nur das sicher gestellt wird, das das Skript
	; läuft sobald neue Befunde dem Ordner hinzugefügt werden sollen.

;}
;------------------------------------------------------------------------------------------------------------------------------------
Version = V0.01
;------------------------------------------------------------------------------------------------------------------------------------
;------------------------------------- written by Ixiko -this version is from 11.12.2018 ------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------
;------------------------------- please report errors and suggestions to me: Ixiko@mailbox.org ------------------------
;---------------------------- use subject: "Addendum" so that you don't end up in the spam folder ---------------------
;------------------------------------------------------------------------------------------------------------------------------------
;--------------------------------- GNU Lizenz - can be found in main directory  - 2017 ----------------------------------
;------------------------------------------------------------------------------------------------------------------------------------
;###############################################################

/* 					Skript Versionsgeschichte

	-------------------------------------------------------------------------------------------------------------------------------------------------------------
	20.10.2018			geplante Veränderungen für Addendum für AlbisOnWindows:
								Der Grund zur Erstellung des ScanPool_IndexingServer.ahk Skriptes war,
								- eine Beschleunigung der Anzeige aller PDF Dokumente im ScanPool (max. 1000ms Verzögerung) zu erreichen -
								- Lese- und insbesondere viele Schreibzugriffe auf den BefundOrdner zu vermeiden -
								- auch anderen Skripten einen Zugriff auf die Daten zu ermöglichen -
								. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
								1. Indexing Skript - fungiert als kleiner Datenbankserver, ScanPool.ahk holt sich über das LAN den durch
									dieses Skript erstellten Index in Form eines Objektes
								2. ScanPool.ahk - kann Befehle an dieses Skript senden, z.B. den Index zu erneuern falls die WatchFolder-
									Funktion nicht funktioniert
								3. im ScanPool.ahk Skript - gibt es eine Fallback Funktion falls das Indexskript mal nicht laufen sollte,
									die Fallback Funktion erstellt dann vorrübergehend den Index selbst
								4. andere Skripte sollen auch mit dem Indexskript kommunizieren können, z.B. könnte Addendum.ahk
									einen kleinen Hinweis zu jeder geöffneten Patientenakte das neue Befunde vorhanden sind anzeigen.
	-------------------------------------------------------------------------------------------------------------------------------------------------------------

*/


;{ 1. Scripteinstellungen / Includes

	#NoEnv
	#SingleInstance force
	#Persistent
	#MaxMem 4095	; INCREASE MAXIMUM MEMORY ALLOWED FOR EACH VARIABLE - NECESSARY FOR THE INDEXING VARIABLE
	;#NoTrayIcon

	SendMode Input
	SetWorkingDir %A_ScriptDir%
	SetTitleMatchMode, 2
	SetControlDelay -1
	SetWinDelay, -1
	SetBatchLines -1
	CoordMode, Mouse, Screen
	CoordMode, Pixel, Screen
	CoordMode, Menu, Screen
	CoordMode, Caret, Screen
	CoordMode, Tooltip, Screen

	FileEncoding, UTF-8

	OnExit, OhNoNotYet

;}

;{ 2. Variblen Setup / Registry auslesen

	;{ a) ---------------------------------------------------------------- globale Variablen -----------------------------------------------------------------------


		;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		; allgemeine globale Variablen
		;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			global AddendumDir		:= FileOpen("C:\albiswin\AddendumDir-DO_NOT_DELETE","r").Read()
			global AlbisWinID			:= AlbisWinID()
			global BefundOrdner				     					;BefundOrdner - der Ordner in dem sich die ungelesenen PdfDateien befinden
			global files													;files-Object - enthält sämtliche benötigte Daten
			global ExecPID												;Consolen PID
			global pdfError												;zählt Lesefehler von PDF Dateien
			global RamDrv												;Laufwerksbezeichnung für die RamDisk
			global PageSum											;die Gesamtzahl an Pdf Seiten
			global FileCount											;Anzahl der Pdf Dokumente im Ordner
			global filesMaxIndex										;alternativer Zähler der Pdf-Dateiliste
			global PdfIndexFile										;kompletter Pfad zum Pdf-Index
			global oPdf:= []											;Pdf-Array für die enthaltenen Pdf-Dokumente im Befundordner
			global Pdf:= Object()									;Pdf-Array für die enthaltenen Pdf-Dokumente im Befundordner
		;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		; globale Variablen für die Addendum Patientendatenbank
		;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			global oPat:= Object()						           	;Patientendatenbank Objekt für Addendum
			global AddendumDBPath			                	;der Pfad zur Datenbank sollte immer global angelegt sein

		;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		; Auslesen des Addendum Stammverzeichnisses aus der Registry
		;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			global logFile:=  % AddendumDir "\logs'n'data\AddendumLog.txt"
	;}

	;{ b) ------------------------------------------------- Einstellungen aus der Addendum.ini einlesen ------------------------------------------------------

		;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		; Pfad des Albis Befund-Ordner ermitteln - ACHTUNG: die Datenbank wird von Addendum.ahk zuerst eingerichtet!
		;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

			IniRead, BefundOrdner, % AddendumDir "\Addendum.ini", ScanPool, BefundOrdner							; BefundOrdner = Scan-Ordner für neue Befundzugänge
			If InStr(BefundOrdner, "Error") {
				;gosub CreateBefundOrdnerDialog
				Hinweis=
				(LTrim
					A C H T U N G


					In der Einstellungsdatei `"Addendum.ini`" ist noch kein Eintrag zum
					Albis Befund/Importordner vorhanden. Bevor dieses Skript ausgeführt werden kann,
					muss dieses Verzeichnis dort hinterlegt werden.

					Wählen Sie bitte den Ordner aus, welcher PDF-Dokumente enthält.

				)

				If FileExist("C:\AlbisWin.loc\local.ini")
						localini:="C:\AlbisWin.loc\local.ini"
				else if FileExist("C:\AlbisWin\local.ini")
						localini:="C:\AlbisWin\local.ini"

				befunde:= Object()
				IniRead, MainPath, % localini, Pfade, Main
				If InStr(MainPath, "Error")
					defaultDrive:="C:\"
				else
					SplitPath, MainPath,,,,, defaultDrive

				Loop, 8 {
						IniRead, dir	, % localini, bvl, % "VP600" A_Index
						If InStr(dir, "Error")
								break
						else
						{
								If (InStr(dir,":") and !befunde.Haskey(dir))
											befunde[(dir)]:= prog
						}
				}

				befIdx:= 0, faktor:= 1

				IniRead, Font			, % AddendumDir "\Addendum.ini", Addendum, StandardFont
				IniRead, boldFont	, % AddendumDir "\Addendum.ini", Addendum, StandardBoldFont
				IniRead, FontSize	, % AddendumDir "\Addendum.ini", Addendum, StandardFontSize
				IniRead, FontColor, % AddendumDir "\Addendum.ini", Addendum, DefaultFntColor
				IniRead, BgColor	, % AddendumDir "\Addendum.ini", Addendum, DefaultBgColor
				IniRead, BgColor1	, % AddendumDir "\Addendum.ini", Addendum, DefaultBgColor2
				IniRead, BgColor2	, % AddendumDir "\Addendum.ini", Addendum, DefaultBgColor3

				Gui, local: new, -DPIScale HWNDhlocalGui
				Gui, local: Color, % "c" BgColor
				Gui, local: Add, Progress, % "x10 y10 w400 h60 c" BgColor2 " Background" BgColor2 " vPrgr1 HWNDhPrgr1", 100
				WinSet, ExStyle, -0x00020000, ahk_id %hPrgr1%
				Loop, Parse, Hinweis, `n, `r
				{
						If (A_Index = 1)
						{
							Gui, local: Font, % "s" FontSize+(fntp:=30) " c" BgColor " Bold q5", % boldFont
							Gui, local: Add, Text, % "xm y15 BackgroundTrans vZeile" A_Index, % A_LoopField
						}
						else
						{
							Gui, local: Font, % "s" FontSize+(fntp:=6) " c" FontColor " Normal q5", % Font
							If (A_LoopField ="")
									faktor+=1
							else
							{
									Gui, local: Add, Text, % "xm yp+" (FontSize+fntp)*1.75*faktor " vZeile" A_Index, % A_LoopField
									faktor:= 1
							}
						}
						TextZeilen:= A_Index
				}

				faktor:= "+" (FontSize+fntp)*1.75*faktor*0.75
				Gui, local: Font, % "s" FontSize+(fntp:=2) " c" FontColor " Normal q5", % Font

				For key in befunde
				{
						befIdx++
						Gui, local: Add, Button, % "xm yp" faktor " vBefBtn" befIdx " gBefundOrdnerAuswahl", % key
						;faktor:=""
				}

				befIdx++
				Gui, local: Add, Button, % "xm yp" faktor " vBefBtn" befIdx " gBefundOrdnerAuswahl", % "Ordner per Dialogfenster auswählen..."

				Gui, local: Show, Hide, mögliche Befundordner

				win:= Object()
				win:= GetWindowInfo(hlocalGui)
				GuiControlGet, size, local: Pos, Zeile1
				GuiControl, local: MoveDraw, Zeile1, % "x" win.WindowW//2-sizeW//2
				GuiControl, local: MoveDraw, prgr1, % "w" win.ClientW-20 " h" sizeH+sizeX
				Loop, % befIdx
				{
						GuiControlGet, size, local: Pos, % "BefBtn" A_Index
						GuiControl, local: MoveDraw, % "BefBtn" A_Index, % "x" win.WindowW//2-sizeW//2
				}

				Gui, local: Show, % "h" win.WindowH+5, mögliche Befundordner

				WinWaitClose, mögliche Befundordner
				Sleep, 1000
				;MsgBox,, Addendum für AlbisOnWindows, % Hinweis, 20
				;Run, notepad.exe %AddendumDir%\Addendum.ini
				;ExitApp

			}

			PdfIndexFile:= BefundOrdner "\ScanPool-Index.bin"

		;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		; Pfad der Addendum Datenbank ermitteln - ACHTUNG: die Datenbank wird von Addendum.ahk zuerst eingerichtet!
		;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			IniRead, AddendumDBPath, %AddendumDir%\Addendum.ini, Addendum, AddendumDBPath
			If InStr(AddendumDBPath, "Error") {
				Hinweis=
				(LTrim
					-----------------------------------------------------------------------------------------
					|---------------------------------< ACHTUNG! >---------------------------------|
					-----------------------------------------------------------------------------------------

					  In der Einstellungsdatei `"Addendum.ini`" ist noch kein Eintrag zum
					  Addendum Datenbank Verzeichnis vorhanden.

					  Bevor dieses Skript ausgeführt werden kann muss dieses Verzeichnis
					  erstellt und eine Datei mit dem Namen -> Patienten.txt <- vorhanden sein.

					  Starten Sie auf einem Computer das Skript `"Addendum.ahk`".
					  Dieses finden Sie im Stammverzeichnis der Skriptsammlung.
					  Anschließend rufen Sie dieses Skript erneut auf.

					  Das Skript wird in wenigen Sekunden automatisch beendet!
				)
				MsgBox,, Addendum für AlbisOnWindows, % Hinweis, 20
				ExitApp

			}



	;}

	;{ c) ------------------------------------------------------------ Variablen deklarieren -----------------------------------------------------------------------

	;}

	;{ d) ---------------------------------------------------- Einlesen der Addendum Datenbank --------------------------------------------------------------
			If FileExist(AddendumDBPath . "\Patienten.txt") {
			;Einlesen der Datenbank als Textliste, Sortieren aufsteigend nach PatID, Aussortieren doppelter Einträge (später neue Einträge unter den Skripten kommunizieren?)
				FileRead, PatDBtmp, %AddendumDBPath%\Patienten.txt
			;aus der Liste ein Objekt generieren
				Loop, Parse, PatDBtmp, `n, `r
				{
						lSplit:= StrSplit(A_LoopField, "`;", A_Space)
						PatID				:= lSplit[1]
						Nachname		:= lSplit[2]
						Vorname			:= lSplit[3]
						Geschlecht		:= lSplit[4]
						Geburtsdatum	:= lSplit[5]
						Krankenkasse	:= lSplit[6]
						oPat[PatID]:= {"Nn": Nachname, "Vn": Vorname, "Gt": Geschlecht, "Gd": Geburtsdatum, "Kk": Krankenkasse}
				}
			;Freigeben von Resourcen
				VarSetCapacity(PatDBtmp, 0)
				Loop, 6 {
				VarSetCapacity(lSplit[A_Index], 0)
				}
				VarSetCapacity(PatID, 0)
				VarSetCapacity(Nachname, 0)
				VarSetCapacity(Vorname, 0)
				VarSetCapacity(Geschlecht, 0)
				VarSetCapacity(Geburtsdatum, 0)
				VarSetCapacity(Krankenkasse, 0)
		}


	;}

	;{ e) ----------------------------------------- Einlesen oder neu Erstellen der ScanPool-Index.bin -------------------------------------------------------

	/*
			FileDelete, % PdfIndexFile
			If !FileExist(PdfIndexFile)
			{
						SciTEOutput("Beginne mit dem indizieren.", 1, 1, 0)
						gosub newPdfIndex
			}
			else
			{
						oPdf:= ObjLoad(PdfIndexFile)
						If IsObject(oPdf) {
							FileAppend, % A_Now " : [ScanPool - Serverskript] - Die ScanPool PdfIndex Datei wurde nach Start des Skriptes erfolgreich geladen.`n", % logFile
						} else {
							FileAppend, % A_Now " : [ScanPool - Serverskript] - Die ScanPool PdfIndex Datei konnte nach Start des Skriptes nicht geladen werden. Es wird ein neuer Index erzeugt.`n", % logFile
							gosub newPdfIndex
						}
						gosub newPdfIndex
						;checkPdfIndex(BefundDir, oPdf)
						ObjTree(oPdf, "PdfObject", "+ReadOnly +Resize,GuiShow=w1400 h600",-1)
			}
	*/
	;}




;}

		If !FileExist("pdfcont.txt")
		{
			File:= FileOpen("pdfcont.txt", "w", "UTF-8")
			PdfTxt:= PdfToText("M:\Befunde\2019_01_14_15_37_55.pdf", 1, "Latin1")
			File.Write(PdfTxt)
			File.Close()
		}
		else
		{
				FileRead, PdfTxt, pdfcont.txt
		}

		RegExMatch(PdfTxt, "im)(Name).*(Geburtsdatum).*$", match)
		MsgBox, % match

	/*
		Loop, Parse, PdfTxt, `n, `r
		{
				If RegExMatch(A_LoopField)
		}

	*/

ExitApp
return


;{ Funktionen

BefundOrdnerAuswahl: ;{

	;Gui, local: Submit
	nmb:= StrReplace(A_GuiControl, "BefBtn")
	if (nmb=befIdx)
		FileSelectFolder, BefundOrdner, % defaultDrive
	else
		GuiControlGet, BefundOrdner,, % A_GuiControl

	IniWrite, % RTrim(BefundOrdner, "\"), % AddendumDir "\Addendum.ini", ScanPool, BefundOrdner

	Gui, local: Destroy

return
;}

newPdfIndex:  ;{
		PageSum:= CreateIndex(BefundOrdner, oPdf)
		sz:= ObjDump(PdfIndexFile, oPdf)
		SciTEOutput("neuer Index ist erstellt.", 0, 1, 0)
		FileAppend, % A_Now " : [ScanPool - Serverskript] - Der BefundOrdner ist neu indiziert worden. Es sind " oPdf[1] " Pdf-Dateien im Ordner vorhanden.`n", % logFile
return
;}

PdfRename(PdffileName) {



}

CreateIndex(Directory, ByRef oPdf) {

		; Directory - Befundordnerpfad
		; oPdf 		- key-value object das als key den Dateinamen und als Werte Dateigröße und Seitenanzahl enthalten soll
		;					oPdf kann schon Informationen enthalten, z.B. weil die Indexdatei eingelesen wurde und nach einem
		;					Computerneustart überprüft werden muss welche Dateien vorhanden sind und welche nicht mehr

		PageSum:=0									;temp. Variable die die Seitenzahl enthalten wird
		IndexedFiles:=""
		oIndex:= []
		PdfInfo:= Object()

		Loop, %Directory%\*.pdf, 0, 0
		{
					If (A_LoopFileExt = "")								 ; Skip any file without a file extension
						continue
					if A_LoopFileAttrib contains H,S				 	; Skip any file that is either H (Hidden), R (Read-only), or S (System). Note: No spaces in "H,R,S".
						continue
					if A_LoopFileExt contains PDF
						IndexedFiles.= A_LoopFileName . "`n"
		}

		Sort, IndexedFiles

		;setzt den Status für pdf-Dateien auf ungeprüft, werden im Anschluss darauf geprüft ob die Dateien noch vorhanden sind
		For filename in oPdf
			oPdf[(filename)].ckd:= 0

		Loop, Parse, IndexedFiles, `n
		{
					filename:= A_LoopField
					PdfInfo:= PdfInfo2(Directory "\" filename)
					SciTEOutput(filename ", Pages: " (PdfInfo.Pages) ", Tagged: " (PdfInfo.Tagged) ", CreationDate: " (PdfInfo.CreationDate) ", Form: " (PdfInfo.Form) , 0, 1, 0)
					oPdf[(filename)]:={"size": GetKb(StrReplace(PdfInfo.Filesize, " bytes")), "pages": (PdfInfo.Pages), "tagged": (PdfInfo.Tagged), "cdate":(PdfInfo.CreationDate), "Form":(PdfInfo.Form), "ckd": 1}
					PageSum += PdfInfo.Pages
		}

		For filename in oPdf
			If (oPdf[(filename)].ckd = 0)
					oPdf.Delete(filename)

		oPdf["PageSum"]:= PageSum

return PageSum
}

GetKb(bytes) {
return Format("{1:0.1f}", bytes/1024)
}

WatchDirectory(p*) {

      ; Structures
	static FILE_NOTIFY_INFORMATION:="DWORD NextEntryOffset,DWORD Action,DWORD FileNameLength,WCHAR FileName[1]"
	static OVERLAPPED:="ULONG_PTR Internal,ULONG_PTR InternalHigh,{{DWORD offset,DWORD offsetHigh},PVOID Pointer},HANDLE hEvent"

      ; Variables
	static running, sizeof_FNI=65536, WatchDirectory:=RegisterCallback("WatchDirectory","F",0,0)                    ;nReadLen:=VarSetCapacity(nReadLen,8),
	static timer, ReportToFunction, LP, nReadLen:=VarSetCapacity(LP,(260)*(A_PtrSize/2),0)
	static @:=Object(), reconnect:=Object(), #:=Object(), DirEvents, StringToRegEx="\\\|.\.|+\+|[\[|{\{|(\(|)\)|^\^|$\$|?\.?|*.*"

      ; ReadDirectoryChanges related
	static FILE_NOTIFY_CHANGE_FILE_NAME=0x1, FILE_NOTIFY_CHANGE_DIR_NAME=0x2, FILE_NOTIFY_CHANGE_ATTRIBUTES=0x4
			, FILE_NOTIFY_CHANGE_SIZE=0x8, FILE_NOTIFY_CHANGE_LAST_WRITE=0x10, FILE_NOTIFY_CHANGE_CREATION=0x40
			, FILE_NOTIFY_CHANGE_SECURITY=0x100
	static FILE_ACTION_ADDED=1, FILE_ACTION_REMOVED=2,FILE_ACTION_MODIFIED=3
			, FILE_ACTION_RENAMED_OLD_NAME=4, FILE_ACTION_RENAMED_NEW_NAME=5
	static OPEN_EXISTING=3, FILE_FLAG_BACKUP_SEMANTICS=0x2000000, FILE_FLAG_OVERLAPPED=0x40000000
			, FILE_SHARE_DELETE=4, FILE_SHARE_WRITE=2, FILE_SHARE_READ=1, FILE_LIST_DIRECTORY=1

	If p.MaxIndex(){
		If (p.MaxIndex()=1 && p.1=""){
			for i,folder in #
				DllCall("CloseHandle","Uint",@[folder].hD),DllCall("CloseHandle","Uint",@[folder].O.hEvent)
				,@.Remove(folder)
			#:=Object()
			DirEvents:=Struct("HANDLE[1000]")
			DllCall("KillTimer","Uint",0,"Uint",timer)
			timer=
			Return 0
		} else {
			if p.2
				ReportToFunction:=p.2
			If !IsFunc(ReportToFunction)
				Return -1 ;DllCall("MessageBox","Uint",0,"Str","Function " ReportToFunction " does not exist","Str","Error Missing Function","UInt",0)
			RegExMatch(p.1,"^([^/\*\?<>\|""]+)(\*)?(\|.+)?$",dir)
			if (SubStr(dir1,0)="\")
				StringTrimRight,dir1,dir1,1
			StringTrimLeft,dir3,dir3,1
			If (p.MaxIndex()=2 && p.2=""){
				for i,folder in #
					If (dir1=SubStr(folder,1,StrLen(folder)-1))
						Return 0 ,DirEvents[i]:=DirEvents[#.MaxIndex()],DirEvents[#.MaxIndex()]:=0
									@.Remove(folder),#[i]:=#[#.MaxIndex()],#.Remove(i)
				Return 0
			}
		}
		if !InStr(FileExist(dir1),"D")
			Return -1 ;DllCall("MessageBox","Uint",0,"Str","Folder " dir1 " does not exist","Str","Error Missing File","UInt",0)
		for i,folder in #
		{
			If (dir1=SubStr(folder,1,StrLen(folder)-1) || (InStr(dir1,folder) && @[folder].sD))
					Return 0
			else if (InStr(SubStr(folder,1,StrLen(folder)-1),dir1 "\") && dir2){ ;replace watch
				DllCall("CloseHandle","Uint",@[folder].hD),DllCall("CloseHandle","Uint",@[folder].O.hEvent),reset:=i
			}
		}
		LP:=SubStr(LP,1,DllCall("GetLongPathName","Str",dir1,"Uint",&LP,"Uint",VarSetCapacity(LP))) "\"
		If !(reset && @[reset]:=LP)
			#.Insert(LP)
		@[LP,"dir"]:=LP
		@[LP].hD:=DllCall("CreateFile","Str",StrLen(LP)=3?SubStr(LP,1,2):LP,"UInt",0x1,"UInt",0x1|0x2|0x4
						,"UInt",0,"UInt",0x3,"UInt",0x2000000|0x40000000,"UInt",0)
		@[LP].sD:=(dir2=""?0:1)

		Loop,Parse,StringToRegEx,|
			StringReplace,dir3,dir3,% SubStr(A_LoopField,1,1),% SubStr(A_LoopField,2),A
		StringReplace,dir3,dir3,%A_Space%,\s,A
		Loop,Parse,dir3,|
		{
			If A_Index=1
				dir3=
			pre:=(SubStr(A_LoopField,1,2)="\\"?2:0)
			succ:=(SubStr(A_LoopField,-1)="\\"?2:0)
			dir3.=(dir3?"|":"") (pre?"\\\K":"")
					. SubStr(A_LoopField,1+pre,StrLen(A_LoopField)-pre-succ)
					. ((!succ && !InStr(SubStr(A_LoopField,1+pre,StrLen(A_LoopField)-pre-succ),"\"))?"[^\\]*$":"") (succ?"$":"")
		}
		@[LP].FLT:="i)" dir3
		@[LP].FUNC:=ReportToFunction
		@[LP].CNG:=(p.3?p.3:(0x1|0x2|0x4|0x8|0x10|0x40|0x100))
		If !reset {
			@[LP].SetCapacity("pFNI",sizeof_FNI)
			@[LP].FNI:=Struct(FILE_NOTIFY_INFORMATION,@[LP].GetAddress("pFNI"))
			@[LP].O:=Struct(OVERLAPPED)
		}
		@[LP].O.hEvent:=DllCall("CreateEvent","Uint",0,"Int",1,"Int",0,"UInt",0)
		If (!DirEvents)
			DirEvents:=Struct("HANDLE[1000]")
		DirEvents[reset?reset:#.MaxIndex()]:=@[LP].O.hEvent
		DllCall("ReadDirectoryChangesW","UInt",@[LP].hD,"UInt",@[LP].FNI[],"UInt",sizeof_FNI
					,"Int",@[LP].sD,"UInt",@[LP].CNG,"UInt",0,"UInt",@[LP].O[],"UInt",0)
		Return timer:=DllCall("SetTimer","Uint",0,"UInt",timer,"Uint",50,"UInt",WatchDirectory)
	} else {
		Sleep, 0
		for LP in reconnect
		{
			If (FileExist(@[LP].dir) && reconnect.Remove(LP)){
				DllCall("CloseHandle","Uint",@[LP].hD)
				@[LP].hD:=DllCall("CreateFile","Str",StrLen(@[LP].dir)=3?SubStr(@[LP].dir,1,2):@[LP].dir,"UInt",0x1,"UInt",0x1|0x2|0x4
						,"UInt",0,"UInt",0x3,"UInt",0x2000000|0x40000000,"UInt",0)
				DllCall("ResetEvent","UInt",@[LP].O.hEvent)
				DllCall("ReadDirectoryChangesW","UInt",@[LP].hD,"UInt",@[LP].FNI[],"UInt",sizeof_FNI
					,"Int",@[LP].sD,"UInt",@[LP].CNG,"UInt",0,"UInt",@[LP].O[],"UInt",0)
			}
		}
		if !( (r:=DllCall("MsgWaitForMultipleObjectsEx","UInt",#.MaxIndex()
					,"UInt",DirEvents[],"UInt",0,"UInt",0x4FF,"UInt",6))>=0
					&& r<#.MaxIndex() ){
			return
		}
		DllCall("KillTimer", UInt,0, UInt,timer)
		LP:=#[r+1],DllCall("GetOverlappedResult","UInt",@[LP].hD,"UInt",@[LP].O[],"UIntP",nReadLen,"Int",1)
		If (A_LastError=64){ ; ERROR_NETNAME_DELETED - The specified network name is no longer available.
			If !FileExist(@[LP].dir) ; If folder does not exist add to reconnect routine
				reconnect.Insert(LP,LP)
		} else
			Loop {
				FNI:=A_Index>1?Struct(FILE_NOTIFY_INFORMATION,FNI[]+FNI.NextEntryOffset):Struct(FILE_NOTIFY_INFORMATION,@[LP].FNI[])
				If (FNI.Action < 0x6){
					FileName:=@[LP].dir StrGet(FNI.FileName[""],FNI.FileNameLength/2,"UTF-16")
					If (FNI.Action=FILE_ACTION_RENAMED_OLD_NAME)
							FileFromOptional:=FileName
					If (@[LP].FLT="" || RegExMatch(FileName,@[LP].FLT) || FileFrom)
						If (FNI.Action=FILE_ACTION_ADDED){
							FileTo:=FileName
						} else If (FNI.Action=FILE_ACTION_REMOVED){
							FileFrom:=FileName
						} else If (FNI.Action=FILE_ACTION_MODIFIED){
							FileFrom:=FileTo:=FileName
						} else If (FNI.Action=FILE_ACTION_RENAMED_OLD_NAME){
							FileFrom:=FileName
						} else If (FNI.Action=FILE_ACTION_RENAMED_NEW_NAME){
							FileTo:=FileName
						}
          If (FNI.Action != 4 && (FileTo . FileFrom) !="")
						@[LP].Func(FileFrom=""?FileFromOptional:FileFrom,FileTo)
				}
			} Until (!FNI.NextEntryOffset || ((FNI[]+FNI.NextEntryOffset) > (@[LP].FNI[]+sizeof_FNI-12)))
		DllCall("ResetEvent","UInt",@[LP].O.hEvent)
		DllCall("ReadDirectoryChangesW","UInt",@[LP].hD,"UInt",@[LP].FNI[],"UInt",sizeof_FNI
					,"Int",@[LP].sD,"UInt",@[LP].CNG,"UInt",0,"UInt",@[LP].O[],"UInt",0)
		timer:=DllCall("SetTimer","Uint",0,"UInt",timer,"Uint",50,"UInt",WatchDirectory)
		Return
	}
	Return
}

PraxTTOff:
return
;}



OhNoNotYet:

ExitApp

#Include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Menu.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Misc.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PdfHelper.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Window.ahk
#Include %A_ScriptDir%\..\..\include\Gui\PraxTT.ahk

#Include %A_ScriptDir%\..\..\..\include\ini.ahk
;#Include %A_ScriptDir%\..\..\..\include\Struct.ahk
#Include %A_ScriptDir%\..\..\..\include\Socket.ahk
#Include %A_ScriptDir%\..\..\..\include\ObjDump.ahk
#Include %A_ScriptDir%\..\..\..\include\ObjTree.ahk

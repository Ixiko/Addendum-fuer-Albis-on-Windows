;------------------------------------------------- Modul Dicom2Albis V1.01 ---------------------------------------------
;-------------------------------------------- Addendum  für  AlbisOnWindows ----------------------------------------
;------------------------------------ written by Ixiko -this version is from 11.12.2018 ------------------------------
;--------------------------- please report errors and suggestions to me: Ixiko@mailbox.org -------------------
;------------------------ use subject: "Addendum" so that you don't end up in the spam folder ---------------
;--------------------------------- GNU Lizenz - can be found in main directory  - 2017 --------------------------
;----------------------------------------------------------------------------------------------------------------------------------

/*																				##### DESCRIPTION	#####

GERMAN

	dieses Modul öffnet eine DICOM CD, liest die Inhalte der DICOMDIR Datei
	konvertiert im Anschluß die DICOM Bilddateien nach jpeg und integriert diese in die Patientenakte
	ich benötige die Dateien nicht im DICOM Format, es reicht komprimiert als jpeg und ich habe dadurch geringere Datenmenge
	Grund: ich möchte das integrierte Praxisarchiv abschaffen, das finden der Bilder geht mir so schneller

ENGLISH

								############# ATTENTION - USING PURPOSES #########################

YOU need the following files to use my script:1.dcm2xml.exe (dcmdata.dll, oflog.dll, ofstd.dll) und 2.dicom2.exe
and a program for converting continuous shooting (MRI, CT) into avi - I recommend MicroDicom (previously free)
1. Download: http://dcmtk.org/dcmtk.php.en 	or 	Official DCMTK Github Mirror http://git.dcmtk.org
2. Download: http://dicom2.barre.nom.fr
3. Donwload:

	this module opens a DICOM CD, reads the contents of the DICOMDIR file
	convert the DICOM image files to jpeg and integrate them into the patient file
	I do not need the files in DICOM format (to large and i need no diagnostic monitor) ,  a compressed format is enough for me completely
	Reason: I want to abolish the integrated archive software, the finding of the pictures will be faster without starting an extra program;
																																																							 */

/*																				##### TO-DO LIST #####

~ DICOM CD with mixed content - X-ray and MRI e.g. probably will not be read or just one of them
~ sometimes problems with some DICOM CD's from one radiology - can't read or can't convert - converting works but in the end theres no file
~ Datum der Aufnahme muss in die Akte

																																																							 */

;dies ist die interne Modulbezeichnung - this is the internal module name
ModulShort:= "A1", D2A_Version:= "V1.00" ;kleinste Veränderungen gemacht, Fehler ist noch nicht behoben

;{	1. Scriptablaufeinstellungen (#NoEnV, #Persistant ......) | Script settings
#NoEnv
SetTitleMatchMode, 2
DetectHiddenWindows, On
DetectHiddenText, On
SetBatchLines -1            ; Script nicht ausbremsem (Default: "sleep, 10" alle 10 ms)
SetControlDelay, -1         ; Wartezeit beim Zugriff auf langsame Controls abschalten
SetWinDelay, -1             ; Verzögerung bei allen Fensteroperationen abschalten
OnExit, CloseVerlauf

hIBitmap:= Create_Dicom2Albis_ico(true)
Menu Tray, Icon, hIcon:  %hIBitmap%
Menu, Tray, NoStandard
Menu, Tray, Add, Variablen zeigen, ShowSkriptVars
Menu, Tray, Add, neu starten, SkriptReload
Menu, Tray, Add, Beenden, CloseVerlauf
;}

;{	2. Includebereich und OnExit-Anweisung (#include libs/***.ahk) | Include area and OnExit statement

;}

;{	3. ein paar Variablen werden definiert oder aus der Ini geladen | a few variables are defined or loaded from the Ini

Indipendent = 0								; dies ist die Variable für die 2 Skriptmodi -
														; -	abhängig von Addendum für AlbisOnWindows oder
														; -	standalone Version nur zum DICOM Format umwandeln
CompName:=A_ComputerName
D2Aini = %A_ScriptDir%\Dicom2Albis.ini
SysGet, Mon1, MonitorWorkArea
FileEncoding, UTF-8
decission = 0									;Entscheidungsflag
param = %1%

;{ ;1													; einlesen der .ini Daten
		global AddendumDir
		AddendumDir := RegReadUniCode64("HKEY_LOCAL_MACHINE", "SOFTWARE\Addendum für AlbisOnWindows", "ApplicationDir")
		AlbisWinID := AlbisWinID()
		If !AlbisWinID
			AlbisIsClosed:=1

		scriptPID := DllCall("GetCurrentProcessId")

		;{  -------------------------------------------					Ini - Read 				-------------------------------------------
	IniRead:
	If FileExist(AddendumDir . "\Addendum.ini") {
			IniRead, Laufwerk, %AddendumDir%\Addendum.ini, %CompName%, DicomCD
			IniRead, Module, %AddendumDir%\Addendum.ini, %CompName%, Module
			IniRead, ABefDir, %AddendumDir%\Addendum.ini, ALBIS, AlbisBefundDir
			IniRead, Comp, %AddendumDir%\Addendum.ini, Computer, %CompName%
			IniRead, Font, %AddendumDir%\Addendum.ini, Dicom2Albis, Font
			IniRead, Options, %AddendumDir%\Addendum.ini, Dicom2Albis, Options
			IniRead, GSize, %AddendumDir%\Addendum.ini, Dicom2Albis, Gui
			IniRead, DICOM2AVI_exe, %AddendumDir%\Addendum.ini, Dicom2Albis, DICOM2AVI_exe, C:\Program Files\MicroDicom\mDicom.exe
			IniRead, AlbisKey4, %AddendumDir%\Addendum.ini, Albis, AlbisCSExe_Regkey_4

								;Defaultwerteinstellung für die GUI wenn nötig
								If (GSize = "Error") or (GSize = ",,,") {									;Zuordnen der Fensterposition
										ACw:= 450																;Breite des Verlauf Gui
										ACx:= Mon1Right - ACw - 80									;x Position
										ACy:= 100																;y Position
										ACh:= Mon1Bottom - ACy - 150								;Höhe der Gui
								} else If Instr(GSize,"`,") {
									StringSplit, GSize, GSize, `,
										ACx:=GSize1, ACy:=GSize2, ACw:=GSize3, ACh:=GSize4
											}
			}	else {
					MsgBox,4, Finde ini-Datei nicht., Das ini-File %AddendumDir%\Addendum.ini wurde nicht gefunden`noder konnte nicht geöffnet werden. Bitte wähle den Ordner manuell aus.
					FileSelectFolder, AddendumDir
					goto, IniRead
					}

			;} ---------------------------------------------------- End Ini Read --------------------------------------------------------------


			If Not Instr(Module, ModulShort) and Not Instr(Module,"All") {															;ermitteln ob die Software auf dem Rechner laufen darf
					MsgBox, 0x1000, Dicom2Albis, Dieses Modul hat keine Freigabe für diesen Computer.`nDas Dicom2Albis-Modul wird deshalb beendet!, 10
							ExitApp
											}

				Loop, Parse, Comp, `|																												; ermitteln von User und Pass mittels Parsen
					{
						Login%A_Index% := A_LoopField
					}

			IncludeDir:= AddendumDir . "\include"																							;includedir must be different, depending from run mode

			If (param<>"") {																																;zugewiesenes automatisches Laufwerk mit der DICOM CD nutzen
				MsgBox, 4, Dicom2Albis, Das Skript wurde automatisch gestartet!`nZum Abbrechen des weiteren Vorgangs drücken Sie bitte auf 'Nein', 10
				IfMsgBox, No
						ExitApp
				Laufwerk:= param
					}

			;the RegKey Path must be splitted for use with RegRead
			HKeyPos:= Instr(Albiskey4, "`\",false, 1)-1 , keyPos:= Instr(Albiskey4, "`\",false, -1)+1
			HKEY:= SubStr(AlbisKey4, 1, HKeyPos)
			Keypath:=SubStr(AlbisKey4, HKeyPos+2, StrLen(Albiskey4)-HKeyPos-keypos-3)
			key:= SubStr(AlbisKey4, keyPos)
			RegRead, AlbisDir, %HKEY%, %Keypath%, %key%

			If (AlbisDir="") {																															;if no albis is installed, the script will still convert the files

			} else if (AlbisExist() = 0 ) and (Indipendent = 0) {																			; if Dicom2Albis is an Addendum module Albis has to be started
						Result:=AlbisNeuStart(CompName, Login1, Login2, "Dicom2Albis", AddendumDir)
								if (Result = 0) {
									MsgBox, Dicom2Albis, Albis konnte nicht gestartet werden.`nDicom2Albis wird jetzt beendet.
										ExitApp
													}
						}
;}  ;1

;{										3.a. Deklaration weiterer Variablen	- ACx ...


;DICOM2ALBIS Logo
Gdiw = 500																		;Breite des DICOM2ALBIS Logo
Gdih = 270																		;Höhe
Gdix:= Floor((Mon1Right - Gdiw)/2)									;x Position
Gdiy:= Floor((Mon1Bottom - Gdih)/1.8)							;y Position

dicomindex:= 0																;Variable für den DICOM Bilder Index
imager:=""																		;Sammel-Variable für alle Bilderpfade
picsconverted = 0																;falls das Skript vom User abgebrochen wird, kann nachgefragt werden
dicomfile = %Laufwerk%\DICOMDIR									;never touch this line
drv = %Laufwerk%
;}

;{										3.b. Liste von Untersuchungen die das Script in jpeg umwandeln kann - siehe Modalitys-Processable
PicModList:= {}, MovModList:= {}
;Kürzelliste für Untersuchungen die in Bilddateien umgewandelt werden können
PicModList:={AR:"Autorefraction", BDUS:"Bone Densitometry - ultrasound", BI:"Biomagnetic imaging", BMD:"Bone Densitometry - X-Ray", CR:"Computed Radiography", DG:"Diaphanography", DX:"Digitales Röntgen", ES:"Endoscopy", GM:"General Microscopy", HD:"Hemodynamic Waveform", IO:"Intra-Oral Radiography", IVOCT:"Intravascular Optical Coherence Tomography", IVUS:"Intravascular Ultrasound", KER:"Keratometry", LEN:"Lensometry", LS:"Laser surface scan", MG:"Mammography", PX:"Panoramic X-Ray", RG:"Radiographic imaging - conventional film`/screen", RF:"Radio Fluoroscopy", RTIMAGE:"Radiotherapy Image", SM:"Slide Microscopy", TG:"Thermography", US:"Ultrasound", OP:"Ophthalmic Photography", XA:"X-Ray Angiography", XC:"External-camera Photography"}
;Kürzelliste für Untersuchungen die in Filme umgewandelt werden sollten
MovModList:={CT:"Computertomographie", MR:"Magnetresonanztomographie MRT", NM:"Nuklearmedizin", OCT:"Optical Coherence Tomography (non-Ophthalmic)", OPT:"Ophthalmic Tomography", PT:"	Positron emission tomography (PET)"}
;}

file:= A_ScriptDir . "\Dicom2Albis500.png"

If !FileExist(file) {				;prüfen ob das Logo vorhanden ist
	MsgBox ,0x1000 ,Das Dicom2Albis-Logobild fehlt., Dicom2Albis wird beendet. Sie können das fehlende`nBild durch ein anderes Bild mit der Größe 500x270px im .png Format ersetzen.
	ExitApp
		}

;}

;{ 4. wenn keine CD vorhanden ist abbrechen oder warten bis die CD bereit ist | if there is no CD, stop or wait until the CD is ready

CDWait:
If !FileExist(dicomfile) and param="" {								;prüft ob eine CD mit DICOM-Dateien eingelegt ist
			MsgBox, 0x40035, Albis2Dicom, Entweder ist keine Dicom CD eingelegt oder`ndas Laufwerk ist nicht bereit.`nDrücken Sie auf `'Wiederholen`' wenn Sie soweit sind.
			IfMsgBox, Retry
				goto, CDWait
			IfMsgBox, Cancel
				ExitApp
	}
;}

;{ 5. Anzeigen des Logobildes - Hwnd Variable(AC2ID) und hdc Variabel(GGhdc) | Display the logo image
GdiGui:=PIC_GUI(20, file, Gdix, Gdiy, Gdiw, Gdih)
AC2ID:= GdiGui[1]
GGhdc:= GdiGui[2]

;}

;{ 6. Gosub-Routine zur Verlaufs-Gui und Start des Programmablaufes | gosub routine for history gui and start of the program
gosub, VerlaufGui

gosub, CopyDicom
;}

;{ 7. HOTKEY BEREICH ------------------------------------------------------------------------------------------
^+ü::
		ListVars
		return

^p::Pause, Toggle

$ESC::ExitApp

;}

Return
;________________________________ENDE AUTOEXECUTE__________________________________________________

CopyDicom:
; 8. enthält sämtliche Routinen zum Auslesen der DicomCD, Erstellen der Dateinamen, Umwandeln der Dicombilddateien
;{ Contains all the routines for reading the DicomCD, creating the filenames, converting the Dicom image files
											;und Funktionen zum Einsortieren in die Patientenakte, + Animationen  ## Serien werden hier nicht in avi oder wmv umgewandelt
											;and functions for sorting into the patient file, Animations ## series are not converted into avi or wmv here

;{													 		8a. wandelt das DicomDIR File in eine lesbare Textdatei um | converts the DicomDIR file into a readable text file
	LV_Add("", "Öffne die DICOMDIR Datei im CD-Laufwerk " . Laufwerk)
	Sleep, 1000

	DICOMtemp = %A_Temp%\DICOM.txt

	cmdline = %dicomfile% --convert-to-utf8 %A_Temp%\DICOM.txt

	If !FileExist(A_Temp . "\DICOM.txt") {																					;converts the DICOMDIR file in the CD root directory into a readable format
			RunWait, %IncludeDir%\dcm2xml.exe %cmdline%, , Min
				sleep, 500
				If !FileExist(A_Temp . "\DICOM.txt") {																		;Aborting the program if creating the DICOM.txt in the Temp folder fails
							MsgBox,1, Das Schreiben der DICOM.txt in den %A_Temp%-Ordner ist fehlgeschlagen.`nDas Programm muß daher beendet werden.
									ExitApp
										}
		}

;}

;{													 		8b. die Tags aus der aufbereiteten DICOMDIR werden ausgelesen und Variablen zugeordnet
	FileRead, html, %A_Temp%\DICOM.txt																				;liest die umgewandelte Datei in eine Variable
	dicom_Modality:=Strx(html, "Modality",1,10, "</element>",1,10)
	LV_Add("", "Untersuchungskürzel: '" . dicom_Modality . "' " . PicModlist[dicom_Modality] )

	dicom_PatName:= StrX(html, "PatientName",1,13,"</element>",1,10)
	dicom_Geschlecht:=StrX(html, "PatientSex",1,12,"</element>",1,10)
	dicom_GebDatum:= StrX(html, "PatientBirthDate",1,18,"</element>",1,10)
	dicom_UDatum:= StrX(html, "StudyDate",1,11,"</element>",1,10)
	dicom_Nummer:= StrX(html, "AccessionNumber",1,17,"</element>",1,10)
	dicom_institut:= StrX(html, "InstitutionName",1,17,"</element>",1,10)
	dicom_instAddr:= StrX(html, "InstitutionAddress",1,20,"</element>",1,10)
	SZpos:= 1																															;nachsehen ob es mehrere von den Sonderzeichen gibt
	dicom_PatName:= Trim(dicom_PatName, "`^")
	dicom_PatName:= Trim(dicom_PatName)
	StringReplace, dicom_PatName, dicom_PatName, `^, `,%A_Space%
	dicom_GebDatum:= SubStr(dicom_GebDatum, 7, 2) . "`." . SubStr(dicom_GebDatum, 5, 2) . "`." . SubStr(dicom_GebDatum, 1, 4)
	dicom_UDatum:= SubStr(dicom_UDatum, 7, 2) . "`." . SubStr(dicom_UDatum, 5, 2) . "`." . SubStr(dicom_UDatum, 1, 4)

;}

;{ 												 	 	8c. Ausgabe der Tags in der Verlaufsgui damit man sieht das etwas passiert

ImgText = 																																;Text der nachher auf den umgewandelten Bildern erscheinen soll
(
         Patientenname: %dicom_PatName%
               Geschlecht: %dicom_Geschlecht%
          Geburtsdatum: %dicom_GebDatum%
Untersuchungsdatum: %dicom_UDatum%
      Acession number: %dicom_Nummer%
                      Institut: %dicom_institut%
         Institutsadresse: %dicom_instAddr%
)

		LV_Add("", "Patientenname: " . dicom_PatName)
		LV_Add("", "Geschlecht: " . dicom_Geschlecht)
		LV_Add("", "Geburtsdatum: " . dicom_GebDatum)
		LV_Add("", "Untersuchungsdatum: " . dicom_UDatum)
		LV_Add("", "Acession number: " . dicom_Nummer)
		LV_Add("", "Institut: " . dicom_institut)
		LV_Add("", "Institutsadresse: " . dicom_instAddr)
		LV_Add("", "erstelle eine Liste aller DICOM - Bilder...")

;}

;{															8d. Entscheidung wird getroffen, ob die Bilder über die interne Programmfunktion umgewandelt werden können

	if PicModList[dicom_Modality]="" {
		MsgBox, 0x1000, Dicom2Albis, Die Umwandlung der auf dieser CD enthaltenen DICOM-Daten`nwird zur Zeit noch nicht unterstützt.`nDas Dicom2Albis-Modul wird beendet! Es wird dafür das Microdicom-Tool gestartet, 15
		FileDelete, %A_Temp%\DICOM.txt
			gosub, kleine_Animation
				decission= 1
			gosub, CloseVerlauf
		Run, %DICOM2AVI_exe% %dicomfile%
			If InStr( DICOM2AVI_exe, "mDicom") {
			gosub, DICOM2AVI
				}
		ExitApp
	}
;}

;{ 													    8e. einlesen der DICOM Bilddateien in einen Array
Loop, Parse, html, `n
{

		dline:= A_LoopField
		ttline:= SubStr(dline, 1, 50)
		ToolTip, %ttline%, %dgx%+100, %dgy%+50, 1

		If InStr(dline, "StudyDescription") {
			pos1:=  Instr(dline, "StudyDescription") + 18
			pos2:=  Instr(dline, "</element>") ;-11
			StudyDescription:= Substr(dline, pos1, pos2 - pos1)
		} else If InStr(dline, "SeriesDescription") {
			pos1:=  Instr(dline, "SeriesDescription") + 19
			pos2:=  Instr(dline, "</element>") ;-11
			SeriesDescription:= Substr(dline, pos1, pos2 - pos1)
		} else If InStr(dline, "ReferencedFileID") {
			imgpos:=  Instr(dline, "ReferencedFileID") + 18
			imglength:=  Instr(dline, "</element>") ;-11
			dicomindex += 1
			dicom_img%dicomindex%:= Substr(dline, imgpos, imglength-imgpos)
			dicom_Untersuchung%dicomindex%:= StudyDescription . "-" . SeriesDescription
			LV_Add("",  "Studie: " . dicom_Untersuchung%dicomindex%)
			LV_Add("",  "DICOM" . dicomindex . ": " . dicom_img%dicomindex%)
		}

}


FileDelete, %A_Temp%\DICOM.txt
ToolTip,,,, 1
;}

;{ 														8f. starten der GDIP Engine von Windows
If !pToken := Gdip_Startup() {
	MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
	ExitApp
}
;}

;{															8g. jetzt erfolgt Umwandeln und Beschriften der Bilder und kopieren in das Befundverzeichnis und y-Achsen Animation des Logo
AI:=""
Loop, %dicomindex%
{
	AI:= SubStr("00000" . A_Index, -StrLen(dicomindex))
	LV_Add("", "konvertiere Datei " . A_Index)
	file:= dicom_img%A_Index%																											;auslesen einer Pfadzeile der Bilder
	untersuchung:= dicom_Untersuchung%A_Index%																			;Vorbereitungen für den Dateinamen
	cmdline:= "--to`=" . ABefDir . "`\ `-p " . Laufwerk . "`\" . file															;Command line für die dicom2.exe zum Umwandeln der DICOM Bilder ins PNG-Format
	RunWait, %IncludeDir%\dicom2.exe %cmdline%, , NoHide																	;Bild für Bild wird umgewandelt

	StringRight, DicomImgName, file, StrLen(file) - Instr(file, "\", false, 0)												;es werden nur Teile PNG-Namens benötigt
		PngImg:= ABefDir . "`\" . DicomImgName . "`.png"																	;die temporärer PNG-Datei
		JpgImg:= ABefDir . "`\" . AI . ". `[ " . untersuchung . " `] " . dicom_PatName . "`.jpg"					;der Name des späteren JPEG File

	ImageTextOverlay(PngImg, JpgImg, ImgText, Options, Font)															;jetzt werden über die temporäre Datei die Bilddaten gelegt und das Ergebnis wird als JPEG-Datei gespeichert
	dicom_img%A_Index%:=AI . ". `[ " . untersuchung . " `] " . dicom_PatName . "`.jpg"
	picsconverted:= A_Index
	FileDelete, %PngImg%
}

;FileDelete, 00000000.png

kleine_Animation:
y:= Gdiy
maxMove:=  A_ScreenHeight - Gdih - y - 10
MovingTime:= 1200   ;ms
SleepTime:= maxMove//MovingTime
Loop, %maxMove%
{
	y +=1
	WinMove, ahk_id %AC2ID%,,, %y%
	Sleep, %SleepTime%
}


;}

;{															8h.  und fertig mit dem Umwandeln - auf gehts`s ans Unterbringen der Bilder in die Akte

BefundeEintragen:

	AlbisAkteOeffnen(dicom_PatName)

	WinActivate, ahk_exe albisCS.exe
	MsgBox, 4, Nachfrage, Ist die richtige Akte geöffnet worden?
	IfMsgBox, Yes
		Result = 0
	IfMsgBox, No
		Result = 1

	If (Result = 1) {
				MsgBox, 4, neuer Versuch?, Soll ich es nochmal versuchen? Wenn nicht dann öffne die Akte bitte jetzt von Hand`nDu kannst auch auf "Nein" drücken.
						IfMsgBox, Yes
							Result:= AlbisAkteOeffnen(dicom_PatName)
						IfMsgBox, No
							ExitApp
														}

	If (Result = 1) 		;falls erneut die falsche Akte geöffnet ist , springt er zurück zum Anfang der Subroutine, diese Abfrage kann nur noch 1 sein wenn sie von der MSgBox innerhalb der vorherigen Abfrage stammt
		goto, BefundeEintragen

	LV_Add("", "Anzahl der Bilder: " . dicomindex)
	Loop, %dicomindex%
	{
			ImgPath:= dicom_img%A_Index%
			Result:= AlbisOeffneGrafischerBefund()						;diese Funktion öffnet das Fenster GrafischerBefund ausschließlich für jpg Dateien!!
				sleep, 250
			AlbisUebertrageGrafischenBefund(ImgPath)
				sleep, 500
			If WinExist("ahk_class IrfanView") {
				WinClose, ahk_class IrfanView
			}

			WinSet, AlwaysOnTop, On, ahk_exe albisCS.exe
			WinActivate, ahk_class OptoAppClass
				sleep, 500
			SendInput, {ESC}
				sleep, 150
			SendInput, {ESC}
				sleep, 150
			WinSet, AlwaysOnTop, Off, ahk_exe albisCS.exe
			LV_Add("","Bild`: " . A_Index . "`/" . dicomindex . " einsortiert`.")
				dicom_img%A_Index% =
				picsconverted -=1
	}

MsgBox , 0x40030 ,Albis2Dicom, Fertig mit Einsortieren!

;}

ExitApp
;}

;{ 9. SplashText - Kontrollanzeige bei ALBIS Neustart
SplashTextAus:
	SplashTextOff
	ToolTip, , , , %TTnum%
	SetTimer, SplashTextAus, Off

return

;}

VerlaufGui:
;{ 10. Gui für Skriptverlauf

	Gui, AC1:NEW
	Gui, AC1:Font, S10 CDefault, Futura Bk Bt
	minw:= 350
	maxw:= Floor(Mon1Right/4)
	maxh:= Mon1Bottom - 30
	LVw:= ACw-30
	LVh:= ACh-50

	Gui, AC1:Margin,5,5
	Gui, AC1:+LastFound +AlwaysOnTop +Resize +HwndhAC1	MinSize%minw%x600 MaxSize%maxw%x%maxh%
	Gui, AC1:Add, Listview, xm ym w%LVw% r40 vDasLV BackgroundAAAAFF Grid NoSort , Dicom2Albis Fortschritt... 					;h%LVh%

	Gui, AC1:Add, Button, x0 y0 gAbfrage vSPause , Skript pausieren
	Gui, AC1:Add, Button, x180 y300 gAbfrage vSAbbruch, Skript abbrechen
		GuiControlGet, Abb, Pos, SAbbruch

	Gui, AC1:Show, x%ACx% y%ACy% , Dicom2Albis - Addendum für AlbisOnWindows			;w%ACw% h%ACh%
	position:= ACx . "," . ACy . "," . ACw . "," . ACh
	AC1h:= WinExist("A")

	LV_ModifyCol(0,0)
	LV_ModifyCol(1,450)
	Gui, AC1: Default

return

;{					Abfrage
Abfrage:

	ctrl:= A_GuiControl
	if ctrl = SPause
			Pause, Toggle
	if ctrl = SAbbruch
			ExitApp

return
;}

;{					GuiSize
AC1GuiSize:
if ErrorLevel = 1  ; The window has been minimized.  No action needed.
    return
; Otherwise, the window has been resized or maximized. Resize the Edit control to match.
Critical
WinGetPos, wax, way, waw, wah, ahk_id %AC1h%
	position:= wax . "," . way . "," . waw . "," . wah

hNew:= A_GuiHeight-55
wNew:= A_GuiWidth-10

xAbb:= A_GuiWidth - AbbW - 20
yAbb:= hNew+20

GuiControl, Move, DasLV, w%wNew% h%hNew%
GuiControl, Move, SPause, x20 y%yAbb%
GuiControl, Move, SAbbruch, x%xAbb% y%yAbb%
;GuiControl,  BackgroundAAAAFF Grid, MeineListView
WinSet, ReDraw,, ahk_id %AC1h%
Critical, Off
return
;}

AC1GuiClose:

;}

;{ 11.Abschluß Animation - Resource GDI(AC2ID) freigeben und Dicom2Albis Beenden (ONEXIT ROUTINE)
CloseVerlauf:
	If (decission=2)
			ExitApp

	;Verlaufs Gui beenden
	WinGetPos, wax, way, waw, wah, ahk_id %AC1h%
	position:= wax . "," . way . "," . waw . "," . wah
	if (wax<>"" && way<>"" && waw<>"" && wah<>"") {
		If (wax >-1 and way >-1)
			IniWrite, %position%, %AddendumDir%\Addendum.ini, Dicom2Albis, Gui
									}
	If (picsconverted>0) {																							;wenn schon Bilder erstellt nachfragen, ob diese gelöscht werden sollen
		Function_Eject(Drv)
		MsgBox, 4, Dicom2Albis, Möchten Sie die umgewandelten Bilder behalten?
		IfMsgBox, No
		{
			Loop, % dicomindex
				FileDelete, % dicom_img%A_Index%

		}
	}

	x:= Gdix
	maxMove:=  Mon1Right - x - 50
	MovingTime:= 2200   ;ms
	TransStep:= 255/maxMove
	SleepTime:= maxMove//MovingTime
	trans:= 255

	Loop, % maxMove	{
		x +=1
		UpdateLayeredWindow(AC2ID, GGhdc, x, y , GDIw, GDIh, trans)
		Sleep, % SleepTime
		trans:=  trans - TransStep < 0 ? trans =0 : trans - TransStep
	}

	Gui, AC1:Hide
	WinHide, % "ahk_id " AC2ID
		sleep, 500
	Gdip_Shutdown(AC2ID)
	Gui, AC1:Destroy

	If (decission = 1) {																								 ;return um zur AVI Serienumwandlung zurückzukehren
		decission := 2
		return
	}

	Function_Eject(drv)

ExitApp
;}


;####### THE END ** THE END ** THE END ** THE END ** THE END ** THE END ** THE END ** THE END ** THE END ** THE END ** THE END ** THE END ** THE END #############

;{ 12.Umwandeln von Serienbildern in ein Videoformat mittels mDicom.exe -
DICOM2AVI:
;											dont change position of this label to above

	Loop {

		WinGetActiveTitle, aTitle
		If Instr(aTitle, "MicroDicom") {
			hMicroDicom:=WinExist("A")
			break
		}
		sleep, 300
		If (A_Index> 50) {
			MsgBox, 0x1000, Dicom2Albis, Das MicoDicom Tool konnte nicht gestartet werden.`nDicom2Albis wird jetzt beendet, 10
			ExitApp
		}
	}

	sleep, 5000
	;WinMenuSelectItem, ahk_class Afx:ToolBar:40000000:8:10005:1010,, 1&, 1&

	SendInput, {LControl Down}v{Alt Down}
	SendInput, {LControl Up}{Alt Up}

	WinWait        	, Export to video ahk_class #32770, 6
	WinWaitActive	, Export to video ahk_class #32770, 6

	hExportVideo := WinExist("Export to video ahk_class #32770")

	ControlSetText, Edit1, M:\Befunde\Filme, % "ahk_id " hExportVideo
	ControlSetText, Edit2, % dicom_Modality "_" dicom_PatName "_" dicom_GebDatum "_" dicom_UDatum, % "ahk_id " hExportVideo
	Control, ChooseString, Current patient, ComboBox1, % "ahk_id " hExportVideo
	Control, Check,, Button4,   	, % "ahk_id " hExportVideo
	Control, Check,, Button7,   	, % "ahk_id " hExportVideo
	ControlSetText, Edit3, 25   	, % "ahk_id " hExportVideo
	ControlSetText, Edit4, 100 	, % "ahk_id " hExportVideo
	Control, Check,, Button12, 	, % "ahk_id " hExportVideo
	Control, Check,, Button19, 	, % "ahk_id " hExportVideo
	ControlClick, Button20      	, % "ahk_id " hExportVideo

	WinWaitClose,  % "ahk_id " hExportVideo
		WinClose, MicroDicom - free DICOM

return
;}

;{ 13. Funktionen | functions
StrX( H,  BS="",BO=0,BT=1, ES="",EO=0, ET=1,  ByRef N="" ) {
 ;    | by Skan | 19-Nov-2009  Auto-Parser for XML / HTML by SKAN
/*
1 ) H = HayStack. The "Source Text"
2 ) BS = BeginStr. Pass a String that will result at the left extreme of Resultant String
3 ) BO = BeginOffset.
		Number of Characters to omit from the left extreme of "Source Text" while searching for BeginStr
		Pass a 0 to search in reverse ( from right-to-left ) in "Source Text"
		If you intend to call StrX() from a Loop, pass the same variable used as 8th Parameter, which will simplify the parsing process.
4 ) BT = BeginTrim.
		Number of characters to trim on the left extreme of Resultant String
		Pass the String length of BeginStr if you want to omit it from Resultant String
		Pass a Negative value if you want to expand the left extreme of Resultant String
5 ) ES = EndStr. Pass a String that will result at the right extreme of Resultant String
6 ) EO = EndOffset.
		Can be only True or False.
		If False, EndStr will be searched from the end of Source Text.
		If True, search will be conducted from the search result offset of BeginStr or from offset 1 whichever is applicable.
7 ) ET = EndTrim.
		Number of characters to trim on the right extreme of Resultant String
		Pass the String length of EndStr if you want to omit it from Resultant String
		Pass a Negative value if you want to expand the right extreme of Resultant String
8 ) NextOffset : A name of ByRef Variable that will be updated by StrX() with the current offset, You may pass the same variable as Parameter 3, to simplify data parsing in a loop[/list]
*/



Return SubStr(H,P:=(((Z:=StrLen(ES))+(X:=StrLen(H))+StrLen(BS)-Z-X)?((T:=InStr(H,BS,0,((BO
 <0)?(1):(BO))))?(T+BT):(X+1)):(1)),(N:=P+((Z)?((T:=InStr(H,ES,0,((EO)?(P+1):(0))))?(T-P+Z
 +(0-ET)):(X+P)):(X)))-P) ; v1.0-196c 21-Nov-2009 www.autohotkey.com/forum/topic51354.html
}

ImageTextOverlay(imgInput, imgOutput, Text, Options, Font) {

		;Anzahl der Zeilen und die längste Anzahl der Zeichen merken , im Anschluß soll dann aus Schriftgröße und Zeilenabstand die Größe des Hintergrundrechteckes errechnet werden
		;hier habe ich noch nichts gemacht
		zn = 0 , zl = 0
		Loop, Parse, Text, `n
		{
				zn ++
				sl:= StrLen(A_LoopField)
				If sl>zl
					zl:=sl
								}
		;brauche noch ein paar Informationen über die Bildgröße bevor ich die Schrift über das Bild legen kann
		TextFont := Font, Gdip_FontFamilyCreate(TextFont)
		pBitmap := Gdip_CreateBitmapFromFile(imgInput)
		iHeight:= Gdip_GetImageHeight(pBitmap)
		iWidth:= Gdip_GetImageWidth(pBitmap)


		Fontsize:= Round(iHeight*32/4000)												;Schriftgröße 32 ist bei 4000x4000Pixel noch gut lesbar - einfache Verhältnisrechnung
		Options:=Options . " s" . Fontsize

		RegExMatch(Options, "i)R(\d)", Rendering)
		RegExMatch(Options, "i)S(\d+)(p*)", Size)
		RegExMatch(Options, "i)X([\-\d\.]+)(p*)", xpos)
		RegExMatch(Options, "i)Y([\-\d\.]+)(p*)", ypos)

		rwidth:= Size * zl + 10
		rheight:= Rendering * zn + 10
		rxpos:= xpos - 5
		rypos:= ypos - 5


		G := Gdip_GraphicsFromImage(pBitmap)
		Gdip_SetCompositingMode(G, 4)
		pBrush := Gdip_BrushCreateSolid(0xee000000)
		;Gdip_FillRectangle(G, pBrush, rxpos, rypos, rWidth, rHeight)
		Gdip_DeleteBrush(pBrush)
		Gdip_SetCompositingMode(G, 0)
		Gdip_TextToGraphics(G, Text, Options, TextFont)
		Gdip_SaveBitmapToFile(pBitmap,imgOutput, 90)
		Gdip_DeleteGraphics(G), Gdip_DisposeImage(pBitmap)


}

FileGetDetail(FilePath, Index) {                                                           	; Bestimmte Dateieigenschaft per Index abrufen
   Static MaxDetails := 350
   SplitPath, FilePath, FileName , FileDir
   If (FileDir = "")
      FileDir := A_WorkingDir
   Shell := ComObjCreate("Shell.Application")
   Folder := Shell.NameSpace(FileDir)
   Item := Folder.ParseName(FileName)
   Return Folder.GetDetailsOf(Item, Index)
}

FileGetDetails(FilePath) {                                                                   	; Array der konkreten Dateieigenschaften erstellen
   Static MaxDetails := 350
   Shell := ComObjCreate("Shell.Application")
   Details := []
   SplitPath, FilePath, FileName , FileDir
   If (FileDir = "")
      FileDir := A_WorkingDir
   Folder := Shell.NameSpace(FileDir)
   Item := Folder.ParseName(FileName)
   Loop, %MaxDetails% {
      If (Value := Folder.GetDetailsOf(Item, A_Index - 1))
         Details[A_Index - 1] := [Folder.GetDetailsOf(0, A_Index - 1), Value]
   }
   Return Details
}

GetDetails() {                                                                                    	; Array der möglichen Dateieigenschaften erstellen
   Static MaxDetails := 350
   Shell := ComObjCreate("Shell.Application")
   Details := []
   Folder := Shell.NameSpace(A_ScriptDir)
   Loop, %MaxDetails% {
      If (Value := Folder.GetDetailsOf(0, A_Index - 1)) {
         Details[A_Index - 1] := Value
         Details[Value] := A_Index - 1
      }
   }
   Return Details
}

Function_Eject(Drive) {
	Try 	{

		hVolume := DllCall("CreateFile"
		    , Str, "\\.\" . Drive
		    , UInt, 0x80000000 | 0x40000000  	; GENERIC_READ | GENERIC_WRITE
		    , UInt, 0x1 | 0x2 	                           	; FILE_SHARE_READ | FILE_SHARE_WRITE
		    , UInt, 0
		    , UInt, 0x3                                     	; OPEN_EXISTING
		    , UInt, 0, UInt, 0)

		if (hVolume <> -1) 		{
		    DllCall("DeviceIoControl"
		        , UInt, hVolume
		        , UInt, 0x2D4808   ; IOCTL_STORAGE_EJECT_MEDIA
		        , UInt, 0, UInt, 0, UInt, 0, UInt, 0
		        , UIntP, dwBytesReturned  ; Unused.
		        , UInt, 0)
		    DllCall("CloseHandle", UInt, hVolume)
		}
	Return 1
	} Catch 	{

		Return 0
	}
}

GetToolbarItems(hToolbar) {
    WinGet PID, PID, ahk_id %hToolbar%

    If !(hProc := DllCall("OpenProcess", "UInt", 0x438, "Int", False, "UInt", PID, "Ptr")) {
        Return
    }

    If (A_Is64bitOS) {
        Try DllCall("IsWow64Process", "Ptr", hProc, "int*", Is32bit := true)
    } Else {
        Is32bit := True
    }

    RPtrSize := Is32bit ? 4 : 8
    TBBUTTON_SIZE := 8 + (RPtrSize * 3)

    SendMessage 0x418, 0, 0,, ahk_id %hToolbar% ; TB_BUTTONCOUNT
    ButtonCount := ErrorLevel

    IDs := [] ; Command IDs
    Loop %ButtonCount% {
        Address := DllCall("VirtualAllocEx", "Ptr", hProc, "Ptr", 0, "uPtr", TBBUTTON_SIZE, "UInt", 0x1000, "UInt", 4, "Ptr")

        SendMessage 0x417, % A_Index - 1, Address,, ahk_id %hToolbar% ; TB_GETBUTTON
        If (ErrorLevel == 1) {
            VarSetCapacity(TBBUTTON, TBBUTTON_SIZE, 0)
            DllCall("ReadProcessMemory", "Ptr", hProc, "Ptr", Address, "Ptr", &TBBUTTON, "uPtr", TBBUTTON_SIZE, "Ptr", 0)
            IDs.Push(NumGet(&TBBUTTON, 4, "Int"))
        }

        DllCall("VirtualFreeEx", "Ptr", hProc, "Ptr", Address, "UPtr", 0, "UInt", 0x8000) ; MEM_RELEASE
    }

    ToolbarItems := []
    Loop % IDs.Length() {
        ButtonID := IDs[A_Index]
        ;SendMessage 0x44B, %ButtonID% , 0,, ahk_id %hToolbar% ; TB_GETBUTTONTEXTW
        ;BufferSize := ErrorLevel * 2
        BufferSize := 128

        Address := DllCall("VirtualAllocEx", "Ptr", hProc, "Ptr", 0, "uPtr", BufferSize, "UInt", 0x1000, "UInt", 4, "Ptr")

        SendMessage 0x44B, %ButtonID%, Address,, ahk_id %hToolbar% ; TB_GETBUTTONTEXTW

        VarSetCapacity(Buffer, BufferSize, 0)
        DllCall("ReadProcessMemory", "Ptr", hProc, "Ptr", Address, "Ptr", &Buffer, "uPtr", BufferSize, "Ptr", 0)

        ToolbarItems.Push({"ID": IDs[A_Index], "String": Buffer})

        DllCall("VirtualFreeEx", "Ptr", hProc, "Ptr", Address, "UPtr", 0, "UInt", 0x8000) ; MEM_RELEASE
    }

    DllCall("CloseHandle", "Ptr", hProc)

    Return ToolbarItems
}

Create_Dicom2Albis_ico(NewHandle := False) {
Static hBitmap := 0
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 4508 << !!A_IsUnicode)
B64 := "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"
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
VarSetCapacity(Dec, DecLen, 0)
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
DllCall("Gdiplus.dll\GdipCreateHICONFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
Return hBitmap
}

ShowSkriptVars: ;{
	ListVars
return
;}

SkriptReload: ;{

	Script:= A_ScriptDir "\" A_ScriptName
	scriptPID := DllCall("GetCurrentProcessId")
	run,  Autohotkey.exe /f "%AddendumDir%\include\RestartScript.ahk" "%Script%" "2" "%A_ScriptHwnd%" "%scriptPID%"

return
;}
;}
PraxTTOff:
return


#Include %A_ScriptDir%\..\..\..\include\gdip_all.ahk
#Include %A_ScriptDir%\..\..\..\include\AddendumFunctions.ahk
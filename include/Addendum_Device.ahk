;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                          	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                          	by Ixiko started in September 2017 - last change 22.07.2022 - this file runs under Lexiko's GNU Licence
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

;	2 Funktionen für die automatische Erkennung einer eingelegten DICOM Daten CD
WM_DEVICECHANGE(wParam, lParam) {                                              	;-- erkennt das Einlegen einer CD

	 /*
		When wParam is DBT_DEVICEARRIVAL lParam will be a pointer to a structure identifying the
		device inserted. The structure consists of an event-independent header,followed by event
		-dependent members that describe the device. To use this structure,  treat the structure
		as a DEV_BROADCAST_HDR structure, then check its dbch_devicetype member to determine the
		device type.
	 */

	global Drv, DriveNotification, DICOMDevice

	static DBT_DEVICEARRIVAL   	:= 0x8000 	; http://msdn2.microsoft.com/en-us/library/aa363205.aspx
	static DBT_DEVTYP_VOLUME	:= 0x2      	; http://msdn2.microsoft.com/en-us/library/aa363246.aspx

	dbch_devicetype := NumGet(lParam+4)  	; dbch_devicetype is member 2 of DEV_BROADCAST_HDR

	; Confirmed lParam is a pointer to DEV_BROADCAST_VOLUME and should retrieve Member 4 which is dbcv_unitmask
		If ( wParam=DBT_DEVICEARRIVAL && dbch_devicetype=DBT_DEVTYP_VOLUME ) {

		 ; The logical unit mask identifying one or more logical units. Each bit in the mask corres
		 ; ponds to one logical drive.Bit 0 represents drive A, Bit 1 represents drive B, and so on
			dbcv_unitmask := NumGet(lParam+12 )
			Loop 32                                                                         	; Scan Bits from LSB to MSB
				If ((dbcv_unitmask >> (A_Index-1) & 1) = 1) {           	; If Bit is "ON"
					Drv := Chr(64+A_Index)                                         	; Set Drive letter
					func_DriveData := Func("DriveData").Bind(Drv)
					If !DICOMDevice[Drv].CDDetected                      		; no execution in case CD Drive is allready detected
						SetTimer, % func_DriveData, -1
					Break
				}
		}

Return TRUE
}

DriveData(Drv, FileSystemFilter:="FAT,FAT32,NTFS", opt:="") {                                 		;-- identifiziert das Laufwerk und die Art des Mediums

	; letzte Änderung 21.07.2022
	; nur für CD-Laufwerkserkennung und auch nur für DICOM-CD's (FileSystemFilter schließt bestimmte Formate aus)

		global DICOMDevice

		 If !IsObject(DICOMDevice)
			DICOMDevice := Object()

	; Informationen abgreifen
		DriveGet, Type      	, Type  	   	, % device
		DriveGet, Label     	, Label      	, % device
		DriveGet, Filesystem	, Filesystem	, % device
		DriveGet, Serial     	, Serial     	, % device
		DriveGet, Status     	, Status     	, % device
		DriveGet, StatusCD   	, StatusCD  	, % device

	; Abbruch wenn Festplatte erkannt, zB ein Netzwerklaufwerk
		if Filesystem contains %FileSystemFilter%
			return

	; schaut nach ob auf dem neu angemeldeten Medium die Datei DICOMDIR enthalten ist
		device  	:= Drv ":"
		If !FileExist(dicomfile	:= device "\DICOMDIR") {
			PraxTT("Abbruch`nDie eingelegte CD enthält keine DICOM Daten.", "2 1")
			return
		}

	; Daten des Laufwerks in globalem Objekt speichern
		DICOMDevice[Drv] := {"Type":Type, "Label":Label, "FileSystem":FileSystem, "Serial":Serial, "Status":Status, "StatusCD":StatusCD, "CDDetected":true}

	; Hinweis ausgeben
		PraxTT((t := "01: neue CD in Laufwerk " device ", Type: " Type ", FileSystem: " FileSystem ", Serial: " Serial "`n Status: " Status), "4 1")
		;~ Sleep 6000

		cmdline 	:= device "\DICOMDIR --convert-to-utf8 " A_Temp "\DICOM.txt"
		;~ SciTEOutput( q Addendum.Dir "\include\cmdline\dcm2xml.exe" q " "  cmdline  "`n" dicomfile )

	; Konvertierung starten
		PraxTT(t := "`n02: Namen, Art der Untersuchung und Datum ermitteln ", "8 1")
		RunWait, % q Addendum.Dir "\include\cmdline\dcm2xml.exe" q " " cmdline  , , Min

	; Abbruch des Programmes wenn das Erstellen der DICOM.txt im Temp-Ordner fehlschlägt
		If !FileExist(A_Temp "\DICOM.txt") {
			PraxTT(	"`n03: Das Erstellen der DICOM.txt Datei`n"
					. 	"im " A_Temp "-Ordner ist fehlgeschlagen.`n"
					. 	"Die automatische Umwandlungsroutine für DICOM-CD Inhalte`n"
					. 	"wird nicht ausgeführt.", "2 1")
			return ""
        }

		If !IsObject(dicomTags := DICOMtxt2obj(A_Temp "\DICOM.txt")) {
			PraxTT(t.="`n03: Datenermittlung aus eingelegter Dicom-CD im Laufwerk " device " ist fehlgeschlagen!" , "5 2")
			return
		}

		PraxTT(t.="`n03: Datenermittlung aus eingelegter Dicom-CD im Laufwerk " device " ist durchgeführt!" , "2 2")
		Sleep 2000

		SciTEOutput(cJSON.Dump(dicomTags, 1))
		DICOMFilesCopy(dicomTags, Drv, Addendum.Dicom.Dir)

return dicomTags
}

; --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---

DICOMFilesCopy(dcTags, Drv, basePath, FPattern:="PatientName") {  	;-- kopiert alle Dateiordner in einen Basispfad (basepath)

	; überschreibt keine vorhandenen Verzeichnisse aber wird doppelte Kopien anlegen, da keine Existenzprüfung gemacht wird
	; das Vorgehen ist nicht auf Konsistenz mit den Vorgaben eines DICOM-Servers abgeglichen oder vergleichbar

	static CandyText, CandyProgress, CandyText2

	dbg := true
	dcFolders  	:= []
	FPattern 		.= " PatientBirthDate"					; Mustermann, Max 19800120

  ; String für das Patientenverzeichnis erstellen
	For dcTag, dcValue in dcTags
		FPattern := StrReplace(FPattern, dcTag, dcTags[dcTag])

  ; Verzeichnisse der CD ermitteln
	Loop, % Drv ":\DICOM\*.*", 1, 0
		If InStr(FileExist(A_LoopFileFullPath), "D")
			dcFolders.Push(StrReplace(A_LoopFileFullPath, Drv ":\DICOM\"))

  ; bereits vorhandene Verzeichnisse im Kopieverzeichnis ermitteln
	If InStr(FileExist(basePath "\" FPattern), "D") {
		baseFolders 	:= []
		Loop % basePath "\" FPattern "\*.*" , 1, 0
			If InStr(FileExist(A_LoopFileFullPath), "D")
				baseFolders.Push(StrReplace(A_LoopFileFullPath, basePath "\" FPattern "\"))
	} else
		FileCreateDir % basePath "\" FPattern

	If dbg {
		SciTEOutput("dcFolders:")
		SciTEOutput(dcFolders.Count() ? cJSON.Dump(dcFolders, 1) : "leer")
		SciTEOutput("baseFolders:")
		SciTEOutput(baseFolders.Count() ? cJSON.Dump(baseFolders, 1) : "leer")
	}

  ; vergleicht die Namen der Verzeichnisse auf der CD und im Kopieverzeichnis (inkrementiert Verzeichnisse um ein Überschreiben zu verhindern )
	dcCopyData := [], maxFiles :=0, subschanged := false
	For dcIndex, dcsubfolder in dcFolders {                    	; Verzeichnisse der Dicom-CD

	  ; verglichen wird hier
		subnr := 0, subchanged := false
		For subindex, subfolder in baseFolders  {                        ; Verzeichnisse im Kopie-Basispfad
			If (subfolder=dcsubfolder) {
				Loop {
					subfolder := subfolder (subnr++)
					If (subfolder != dcsubfolder) {
						subchanged := true
						break
					}
				}
			}
			if subchanged
				break
		}

		subschanged := subchanged && !subschanged ? true : subschanged

	  ; dcCopyData enthält alle notwendigen Daten die für das Kopien notwendig sind
		tmpFiles := Array()
		Loop, % Drv ":\DICOM\" dcsubfolder "\*.*", 0, 0
			If !InStr(FileExist(A_LoopFileFullPath), "D") 		; ist eine Datei und kein weiteres Verzeichnis
				tmpFiles.Push(A_LoopFileName)
		maxFiles += tmpFiles.Count()

		copyToPath := subschanged ? subfolder : dcsubfolder
		dcCopyData.Push({ "from"	: Drv ":\DICOM\" dcsubfolder
								, 	  "files"	: tmpFiles
								, 	  "to"    	: basePath "\" FPattern "\" StrReplace(copyToPath, "DICOM\")})

	}

  ; Hinweis das eventuell eine identische Kopie angelegt wird gerade
	If subschanged {
		MsgBox, 4164, % StrReplace(A_ScriptName, ".ahk"), % "Sind wahrscheinlich im Begriff eine identische Kopie dieser DICOM CD anzulegen`n"
																						.	 "Wollen Sie den Inhalt vergleichen, bestätigen Sie mit Ja.`n"
																						.	 "Nach Bestätigung wird ein Explorer Fenster geöffnet`n"
		IfMsgBox, Yes
		{
			Run, % "explorer.exe " q basePath "\" FPattern q
			MsgBox, 4164, % StrReplace(A_ScriptName, ".ahk"), % "Alles überprüft? Wollen Sie eine Kopie anlegen?`n"
			IfMsgBox, No
				return
		}
	}

	If dbg
		SciTEOutput(dcCopyData.Count() ? cJSON.Dump(dcCopyData, 1) : "leer")

  ; Progress-Gui konstruieren - leider nur gelber Hintergrund möglich ;{
	monNr := GetMonitorIndexFromWindow(AlbisWinID() ? AlbisWinID : 0x0)
	dims := ScreenDims(monNr)
	Gui, dcc: new, -Caption -DPIScale +HWNDhdcc +AlwaysOnTop +ToolWindow +Border +E0x02080000						;doublebuffered GUI
	Gui, dcc: -DPIScale
	Gui, dcc: Color, FFCC00			                                                          					;GUI windowcolor (see edge of left & right part)
	Gui, dcc: Margin, 3 , 3
	Gui, dcc: Font, s12 q5, % Addendum.Default, Font
	Gui, dcc: Add,   Text		 , xm  ym w518 h18 vCandyText  cFFFFFF BackgroundTrans Center
	Gui, dcc: Add,   Pic		 , xm  ym w9  h18  BackgroundTrans Icon1, % Addendum.Dir "\assets\cpIcons.icl"		;left part
	Gui, dcc: Add,   Pic		 , x+  ym w500 h18 BackgroundTrans Icon2, % Addendum.Dir "\assets\cpIcons.icl"		;middle part
	Gui, dcc: Add,   Pic		 , x+  ym w9  h18  BackgroundTrans Icon3, % Addendum.Dir "\assets\cpIcons.icl"		;right part
	Gui, dcc: Add,   Progress,  xm ym w518 h18 vCandyProgress 					;-Smooth = deactivates colorizing!
	Gui, dcc: Add,   Text		 , xm  y+2 w518 h18  vCandyText2 cBlack BackgroundTrans Center, Kopiere Dicom-Dateien
	Gui, dcc: Show,  % "Hide NoActivate", CandyProgress

	dcc := GetWindowSpot(hdcc)
	Gui, dcc: Show,  % "xCenter y" dims.H-TaskbarHeight(monNr)-dcc.H-2  " NoActivate"
	;}

  ; jetzt wird kopiert
	filenr := 0
	If dcCopyData.Count() {

		For dccindex, dc in dcCopyData {

		  ; Verzeichnisse werden manuell angelegt, da die Dateien einzeln kopiert werden
			If !InStr(FileExist(dc.to), "D")
				FileCreateDir, % dc.to

			For each, file in dc.files {

			  ; Fortschrittsanzeige
				filenr += 1
				value := Floor(filenr/maxFiles*100)
				color := (value >= 33) && (value < 66) ? "FF00FF" : (value >= 66) ? "00A0FF" : "000000"
				GuiControl, +c%color%, CandyProgress			;set progress color
				GuiControl, , CandyProgress, % abs(value)		;set progress value
				GuiControl, , CandyText, % abs(value) "%"	    	;set progress text %
				GuiControl, , CandyText2, % "Kopiere Dicom-Datei [Verz.Nr: " dccindex ", Datei Nr: "  each "/" dc.files.Count() "]: " dc.from "\" file " ⇨ " dc.to "\" file     	;set progress text %

			  ; Datei kopieren
				If !FileExist(dc.from "\" file)
					continue
				FileCopy, % dc.from "\" file, % dc.to "\" file
				If ((lasterror := GetLastError())>1)
					FileAppend, % FPattern ": " dc.from "\" file ", [Id: " lasterror "] " FormatMessage(lasterror), % Addendum.DBPath "\logs\DicomCD_Kopierfehler.txt", UTF-8

			 }

		}

	}


}

DICOMtxt2obj(dicomtxtPath) {                                                             	;-- Untersuchungsdaten aus konvertierter DICOMDIR auslesen

	; die Funktion ist nur für die Ermittlung der unten in rxDCTags hinterlegten Werte ausgelegt
	; die Rückgabe erfolgt als Objekt

	static rxDCStd1 	:= Chr(0x22) ">(?<Nachname>.*?)\^(?<Vorname>.*?)\<\/"
	static rxDCStd2 	:= Chr(0x22) "\>(?<value>.*?)\<\/"

	static rxDCTags := [	"PatientName", "PatientBirthDate", "PatientID", "PatientSex", "PatientSize", "PatientAge", "PatientWeight"
								,	"PatientMotherBirthName", "Modality", "ReferringPhysicianName", "DirectoryRecordType", "StudyDescription", "SeriesDescription"
								, 	"StudyDate", "SeriesDate", "AcquisitionDate", "InstitutionName", "ContentDate", "StudyTime", "SeriesTime", "AcquisitionTime"
								, 	"ContentTime", "AccessionNumber", "SliceThickness", "ScanOptions", "SequenceVariant", "ScanningSequence", "PatientPosition"
								, 	"StudyID", "StudyInstanceUID" , "SeriesInstanceUID", "SeriesNumber", "CodeValue", "CodingSchemeDesignator", "CodeMeaning"
								,	"ReferencedFileID"]

	static PicModList := {	AR:"Autorefraction", BDUS:"Bone Densitometry - ultrasound", BI:"Biomagnetic imaging", BMD:"Bone Densitometry - X-Ray"
								, 	CR:"Computed Radiography", DG:"Diaphanography", DX:"Digitales Röntgen", ES:"Endoscopy", GM:"General Microscopy"
								, 	HD:"Hemodynamic Waveform", IO:"Intra-Oral Radiography", IVOCT:"Intravascular Optical Coherence Tomography"
								,	IVUS:"Intravascular Ultrasound", KER:"Keratometry", LEN:"Lensometry", LS:"Laser surface scan", MG:"Mammography"
								, 	PX:"Panoramic X-Ray", RG:"Radiographic imaging - conventional film`/screen", RF:"Radio Fluoroscopy"
								, 	RTIMAGE:"Radiotherapy Image", SM:"Slide Microscopy", TG:"Thermography", US:"Ultrasound", OP:"Ophthalmic Photography"
								, 	XA:"X-Ray Angiography", XC:"External-camera Photography"}

	static SeqModList	:= {CT:"Computertomographie", MR:"Magnetresonanztomographie MRT", NM:"Nuklearmedizin"
									, OCT:"Optical Coherence Tomography (non-Ophthalmic)", OPT:"Ophthalmic Tomography"
									, PT:"Positron emission tomography (PET)"}

	If !FileExist(dicomtxtPath) {
		PraxTT("Die konvertierte Datei DICOM.xt ist nicht vorhanden.`n" dicomtxtPath, "3 1")
		return
	}

	dicomTags 	:= Object()
	dicomtxt     	:= FileOpen(dicomtxtPath, "r", "UTF-8").Read()
	For rxIndex, rxTag in rxDCTags
		If RegExMatch(dicomtxt, "i)" rxTag . (rxTag = "PatientName" ? rxDCStd1 : rxDCStd2), dc)
			dicomTags[rxTag] := rxTag = "PatientName"=1 ? dcNachname ", " dcVorname : dcValue

return dicomTags
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

GetLastError() {
	; ================================================
	; Function......: GetLastError
	; DLL...........: Kernel32.dll
	; Library.......: Kernel32.lib
	; U/ANSI........:
	; Author........: jNizM
	; Modified......:
	; Links.........: https://msdn.microsoft.com/en-us/library/ms679360.aspx
	;                 https://msdn.microsoft.com/en-us/library/windows/desktop/ms679360.aspx
	; ===============================================
    return DllCall("kernel32.dll\GetLastError")
}

FormatMessage(MessageId){
	; ===============================================
	; Function......: FormatMessage
	; DLL...........: Kernel32.dll
	; Library.......: Kernel32.lib
	; U/ANSI........: FormatMessageW (Unicode) and FormatMessageA (ANSI)
	; Author........: jNizM
	; Modified......:
	; Links.........: https://msdn.microsoft.com/en-us/library/ms679351.aspx
	;                 https://msdn.microsoft.com/en-us/library/windows/desktop/ms679351.aspx
	; ==============================================
    static size := 2024
	VarSetCapacity(buf, size)
    if !(DllCall("kernel32.dll\FormatMessage", "UInt", 0x1000, "Ptr", 0, "UInt", MessageId, "UInt", 0x0800, "Ptr", &buf, "UInt", size, "UInt*", 0))
        return DllCall("kernel32.dll\GetLastError")
    return StrGet(&buf)
}




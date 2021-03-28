;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                            	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                                            	by Ixiko started in September 2017 - last change 22.03.2021 - this file runs under Lexiko's GNU Licence
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

;	2 Funktionen für die automatische Erkennung einer eingelegten DICOM Daten CD
WM_DEVICECHANGE( wParam, lParam) {                                                        	;--erkennt ob ein einlegen einer CD statt gefunden hat und gibt auch das Laufwerk aus - global drv

	 /*
		When wParam is DBT_DEVICEARRIVAL lParam will be a pointer to a structure identifying the
		device inserted. The structure consists of an event-independent header,followed by event
		-dependent members that describe the device. To use this structure,  treat the structure
		as a DEV_BROADCAST_HDR structure, then check its dbch_devicetype member to determine the
		device type.
	 */

	global Drv, DriveNotification

	static DBT_DEVICEARRIVAL   	:= 0x8000 	; http://msdn2.microsoft.com/en-us/library/aa363205.aspx
	static DBT_DEVTYP_VOLUME	:= 0x2      	; http://msdn2.microsoft.com/en-us/library/aa363246.aspx

	dbch_devicetype := NumGet(lParam+4)  	; dbch_devicetype is member 2 of DEV_BROADCAST_HDR

	; Confirmed lParam is a pointer to DEV_BROADCAST_VOLUME and should retrieve Member 4 which is dbcv_unitmask
		If ( wParam = DBT_DEVICEARRIVAL AND dbch_devicetype = DBT_DEVTYP_VOLUME ) {

		 ; The logical unit mask identifying one or more logical units. Each bit in the mask corres
		 ; ponds to one logical drive.Bit 0 represents drive A, Bit 1 represents drive B, and so on
			dbcv_unitmask := NumGet(lParam+12 )
			Loop 32                                                                         	; Scan Bits from LSB to MSB
				If ((dbcv_unitmask >> (A_Index-1) & 1) = 1) {           	; If Bit is "ON"
					Drv := Chr(64+A_Index)                                             	; Set Drive letter
					Break
				}
		; special operation for Addendum für AlbisOnWindows
			If IsObject(Addendum)
				Addendum.DriveNotification := DriveData(Drv)
			else
				DriveNotification := DriveData(Drv)
		}

Return TRUE
}

DriveData(Drv) {                                                                                          		;--was wurde eingelegt. Identifiziert das Laufwerk und die Art des Mediums

	; letzte Änderung 22.03.2021 - Skriptstil modernisiert

		global DICOMDevice
		global ModalityResult

		device    	:= Drv ":"
		cmdline 	:= % device "\DICOMDIR --convert-to-utf8 " A_Temp "\DICOM.txt"
		dicomfile	:= % device "\DICOMDIR"

		DriveGet, Type      	, Type  	    	, % device
		DriveGet, Label     	, Label      	, % device
		DriveGet, Filesystem	, Filesystem	, % device
		DriveGet, Serial     	, Serial      	, % device

	; Abbruch wenn Festplatte erkannt, zB ein Netzwerklaufwerk
		if Filesystem contains FAT,FAT32,NTFS
			return

	; auf eingelegte CD hinweisen
		DriveNotificationTitle := "neues Medium im Laufwerk [" Drv "]"
		DriveNotificationText := "Type: " Type "`nLabel: " Label "`nFilesystem: " Filesystem "`nSerial: " Serial
		PraxTT(	DriveNotificationTitle "`n"
				. 	DriveNotificationText "`n"
				.  "Versuche den Inhalt der CD zu identifizieren.", "3 3")

	 ; wartet hier so lange bis vom Betriebssystem kein Zugriff mehr auf das Laufwerk erfolgt.
		status := ""
		while (status <> "stopped") && (A_Index <= 90) {
			DriveGet, status, StatusCD , % device
			PraxTT(DriveNotificationTitle "`n"
					. "Versuche den Inhalt der CD zu identifizieren. `n"
					. 	"Warte auf Dateisystemzugriff seit: " Round(A_Index*300/1000, 1) "s`n"
					. 	"Der Status des Lauwerkes ist: " status , "3 3")
			If (status = "open") {
				PraxTT(DriveNotificationTitle "`nDas Laufwerk ist geöffnet. Das Einlesen der Daten wird abgebrochen!", "6 2")
				return 0
			} else if (status = "stopped") {
				break
			}
			sleep 300
		}

	; Dateizugriff nicht erhalten
		if (status <> "stopped") {
			PraxTT(DriveNotificationTitle "`nZugriff auf den CD Inhalt nicht möglich`n", "3 3")
			return 0
		}

	; schaut nach ob auf dem neu angemeldeten Medium die Datei DICOMDIR enthalten ist
		If !FileExist(dicomfile) {
			PraxTT("Abbruch`nDie eingelegte CD enthält keine DICOM Daten.", "2 1")
			return 0
		}

	; Konvertierung starten
		PraxTT("Konvertierung`nKonvertiere die DICOMDIR-Datei nach " A_Temp, "2 10"
		RunWait, % Addendum.Dir "\lib\dcm2xml.exe " q cmdline q , , Min

	; Abbruch des Programmes wenn das Erstellen der DICOM.txt im Temp-Ordner fehlschlägt
		If !FileExist(A_Temp "\DICOM.txt") {
			PraxTT(	"Das Erstellen der DICOM.txt Datei`n"
					. 	"im " A_Temp "-Ordner ist fehlgeschlagen.`n"
					. 	"Die automatische Umwandlungsroutine für DICOM-CD Inhalte`n"
					. 	"wird nicht weiter ausgeführt.", "2 1")
			return 0
        }
		PraxTT("Konvertierung`nKonvertierung nach " A_Temp " beendet.", "2 2")

	; umgewandelte Datei lesen und die Art der Untersuchung auslesen
		For lineNR, line in StrSplit(FileOpen(A_Temp "\DICOM.txt", "r", "UTF-8").Read(), "`n", "`r")
			If InStr(line, "Modality") {
				ModalityResult := line
				break
			}

	; Einstellung für RegRead je nachdem welche AutohotkeyExe gewählt wurde 32bit oder 64bit, sonst kann die Registry nicht gelesen werden
		SetRegView, % (A_PtrSize = 8 ? 64 : 32

	; Skript ausführen
		Run, % "Autohotkey.exe " q "/f " q Addendum.Dir "\Module\Albis_Funktionen\DICOM2Albis\Dicom2Albis.ahk " q " " q device q

return ModalityResult
}

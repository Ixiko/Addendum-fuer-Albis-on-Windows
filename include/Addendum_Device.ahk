;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                            	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                                            	by Ixiko started in September 2017 - last change 07.10.2019 - this file runs under Lexiko's GNU Licence
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

;	2 Funktionen für die automatische Erkennung einer eingelegten DICOM Daten CD
WM_DEVICECHANGE( wParam, lParam) {                                                        	;--erkennt ob ein einlegen einer CD statt gefunden hat und gibt auch das Laufwerk aus - global drv
 Global Drv
 global DriveNotification
 Static DBT_DEVICEARRIVAL := 0x8000 ; http://msdn2.microsoft.com/en-us/library/aa363205.aspx
 Static DBT_DEVTYP_VOLUME := 0x2    ; http://msdn2.microsoft.com/en-us/library/aa363246.aspx

 /*
    When wParam is DBT_DEVICEARRIVAL lParam will be a pointer to a structure identifying the
    device inserted. The structure consists of an event-independent header,followed by event
    -dependent members that describe the device. To use this structure,  treat the structure
    as a DEV_BROADCAST_HDR structure, then check its dbch_devicetype member to determine the
    device type.
 */

 dbch_devicetype := NumGet(lParam+4) ; dbch_devicetype is member 2 of DEV_BROADCAST_HDR

 If ( wParam = DBT_DEVICEARRIVAL AND dbch_devicetype = DBT_DEVTYP_VOLUME )
 {

 ; Confirmed lParam is a pointer to DEV_BROADCAST_VOLUME and should retrieve Member 4
 ; which is dbcv_unitmask

   dbcv_unitmask := NumGet(lParam+12 )

 ; The logical unit mask identifying one or more logical units. Each bit in the mask corres
 ; ponds to one logical drive.Bit 0 represents drive A, Bit 1 represents drive B, and so on

   Loop 32                                           ; Scan Bits from LSB to MSB
     If ( ( dbcv_unitmask >> (A_Index-1) & 1) = 1 )  ; If Bit is "ON"
      {
        Drv := Chr(64+A_Index)                       ; Set Drive letter
        Break
      }
   DriveNotification:=DriveData(Drv)
 }
Return TRUE
}

DriveData(Drv) {                                                                                          		;--was wurde eingelegt. Identifiziert das Laufwerk und die Art des Mediums

	global DICOMDevice
	;AddendumDir sollte im aufrufenden Skript global gemacht sein
	global ModalityResult

	DriveGet, Type  , Type  , %Drv%:
	DriveGet, Label , Label , %Drv%:
	DriveGet, Filesystem, Filesystem, %Drv%:

		;abbrechen des Vorganges wenn es eine Festplatte seien sollte, zB ein Netzwerklaufwerk
		if Filesystem contains FAT,FAT32,NTFS
					return

	DriveGet, Serial, Serial, %Drv%:
	DriveNotificationTitle = neues Medium in Drive : %Drv%:
	DriveNotificationText = Type: %Type%`nLabel: %Label%`nFilesystem: %Filesystem%`nSerial: %Serial%
	TrayTip, %DriveNotificationTitle%, %DriveNotificationText%, 5, 5
		sleep, 5000
	TrayTip, %DriveNotificationTitle%, Versuche den Inhalt der CD zu identifizieren., 5, 25
		sleep, 3000


	cmdline = %Drv%:\DICOMDIR --convert-to-utf8 %A_Temp%\DICOM.txt
	dicomfile = %Drv%:\DICOMDIR
	localIndex:= 0

	Loop {                                                                                    ;wartet hier so lange bis vom Betriebssystem kein Zugriff mehr auf das Laufwerk erfolgt.
		DriveGet, status, StatusCD , %Drv%:
		sleep,1000
		localIndex++
			If (status = "open") {
						TrayTip, %DriveNotificationTitle%, Das Laufwerk ist offen. Das Einlesen der Daten wird abgebrochen, 6, 2
						return
				}
			if localIndex>15
				break

		TrayTip, %DriveNotificationTitle%, Versuche den Inhalt der CD zu identifizieren. `nWarte auf Zugriff auf das Dateisystem seit %localIndex% Sekunde.`nDer Status des Lauwerkes ist: %status%, 5, 3
	} Until status = "stopped"

	DICOMDevice:= Drv . ":"
	If !FileExist(dicomfile) {					                                    	;schaut nach ob auf dem neu angemeldeten Medium die Datei DICOMDIR enthalten ist
			TrayTip, Abbruch, Die eingelegte CD enthält keine DICOM Daten., 300, 300, 6
					return
				}

	TrayTip, Konvertierung, Konvertiere die DICOMDIR-Datei nach %A_Temp%, 300, 300, 6
	RunWait, %AddendumDir%\lib\dcm2xml.exe %cmdline%, , Min
			If !FileExist(A_Temp . "\DICOM.txt") {                            ;Abbruch des Programmes wenn das Erstellen der DICOM.txt im Temp-Ordner fehlschlägt
					MsgBox,1, Das Erstellen der DICOM.txt Datei in den %A_Temp%-Ordner ist fehlgeschlagen.`nDie automatische Umwandlungsroutine für DICOM-CD Inhalte wird nicht weiter ausgeführt.
					return
                            		}
	TrayTip, Konvertierung, Konvertierung nach %A_Temp% beendet., 300, 300, 6

	line:=""
	Loop, Read, %A_Temp%\DICOM.txt
	{
		line:= A_LoopReadLine
		If InStr(line, "Modality") {
			ModalityResult:=line
			break
		}
	}

	if (A_PtrSize = 8) { 						;Einstellung für RegRead je nachdem welche AutohotkeyExe gewählt wurde 32bit oder 64bit, sonst kann die Registry nicht gelesen werden
					SetRegView, 64
				} else if (A_PtrSize = 4) {
						SetRegView, 32
					}

	Run, Autohotkey.exe /f "%AddendumDir%\Module\Albis_Funktionen\DICOM2Albis\Dicom2Albis.ahk " %DICOMDevice%

 return ModalityResult
}

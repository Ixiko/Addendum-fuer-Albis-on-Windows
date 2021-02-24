; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                          	** FAXEINGANG SPEICHERN **
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;
;		Beschreibung:
;
; 		Inhalt:
;       Abhängigkeiten:
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Outlook_Anhang_Speichern.ahk Letzte Änderung:    	07.02.2021
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#NoEnv
SetBatchLines, -1

global compname := StrReplace(A_ComputerName, "-")                    	; der Name des Computer auf dem das Skript läuft
global adm := Object()
adm.scriptname := RegExReplace(A_ScriptName, "\.ahk$")

; -------------------------------------------------------------------------------------------------------------------------------------------------------
  ; Pfade und Einstellungen
  ; -------------------------------------------------------------------------------------------------------------------------------------------------------;{
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Pfad zu Addendum und den Einstellungen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)
		adm.ini	:= AddendumDir "\Addendum.ini"
		If !FileExist(adm.ini) {
			MsgBox, 1024, % adm.scriptname, % "Die Einstellungsdatei ist nicht vorhanden!`n[" adm.ini "]"
			ExitApp
		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Backup Pfad für empfangene Labordateien einlesen und bei Bedarf anlegen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		workini	:= IniReadExt(adm.ini)
		adm.DBPath := IniReadExt("Addendum", "AddendumDBPath")           	; Datenbankverzeichnis
		If InStr(adm.DBPath, "ERROR") {
			MsgBox, 1024, % adm.scriptname, % "In der Addendum.ini ist kein Pfad für`ndie Addendum Datendateien hinterlegt!`n"
			ExitApp
		}
		If !FilePathCreate(adm.DBPath "\sonstiges") {
			MsgBox, 1024, % adm.scriptname, % "Datenbankpfad konnte nicht angelegt werden!"
			ExitApp
		}
		adm.olFolder            	:= IniReadExt(compname, "Outlook_Folder", "Posteingang")
		adm.olSubjectIn         	:= IniReadExt(compname, "Outlook_EMailSubjectFilter_In", "i)Fax\s")
		adm.olSubjectOut     	:= IniReadExt(compname, "Outlook_EMailSubjectFilter_Out", "########")
		adm.olSubjectData    	:= IniReadExt(compname, "Outlook_EMailSubject_Daten", "i)von\s*(?<Absender>[\w\s]+)*\s\(*(?<Nummer>\d\d+)")
		adm.olAttachmentIn  	:= IniReadExt(compname, "Outlook_AttachmentFilter_In", ".*")
		adm.olAttachmentOut	:= IniReadExt(compname, "Outlook_AttachmentFilter_Out", "########")
		adm.BefundOrdner   	:= IniReadExt("ScanPool", "BefundOrdner")


;}

	If InStr(compname, "SP1")
		adm.BefundOrdner := "C:\tmp\outlook"

	ExtractAttachmentOutlook(adm.olFolder, adm.olSubjectIn, adm.olSubjectData, adm.olAttachmentIn, adm.BefundOrdner )
	;ExtractAttachmentOutlook(adm.olFolder, ".*", "i)\s*(?<Absender>.*)", adm.olAttachmentIn, adm.BefundOrdner )

ExitApp

class Outlook {

		__New(DBPath, MainFolder:="") {

				this.DBPath := DBPath

		}

		Connect() {

			try this.ol := ComObjActive("Outlook.Application")
			catch
				this.ol := ComObjCreate("Outlook.Application")

		}

		ExtractAttachment(folderName, SJct_In, SJct_Out, AMFilterIn, AMFilterOut, SaveTo) {

				EMailsWithAttachment := unread := AttachmentCounter := 0
				IF StrLen(SJct_In) = 0
					SJct_In := "i)Fax\s"
				IF StrLen(AMFilterIn) = 0
					AMFilterIn := "i)von\s*(?<Absender>[\w\s]+)*\s\(*(?<Nummer>\d\d+)"

				if !IsObject(this.ol)
					this.Connect()

			olNameSpace := this.ol.GetNameSpace("MAPI")
			olFolder := olNameSpace.Folders(1).Folders(folderName)
			Loop % olFolder.items.count 	{

				email := olFolder.items.Item(A_Index)
				If email.unread
					unread ++

				if email.unread && RegExMatch(email.Subject, SJct_In) && !RegExMatch(email.Subject, SJct_Out) 	{

					; EMail als gelesen markieren
						email.unread := 0
						EMailsWithAttachment ++

					; Posteingang protokollieren
						If EMailsWithAttachment = 1
							FileAppend, % this.datestamp() , % this.DBPath "\outlook.txt", UTF-8

					; FaxAbsender und Faxnummer
						RegExMatch(email.Subject, "i)Fax\svon\s*(?<Absender>[\w\s]+)*\s\(*(?<Nummer>\d\d+)", Fax)
						MBoxMsg .= "Absender:`t"  	FaxAbsender "`n"
						MBoxMsg .= "Nummer: `t" 	FaxNummer "`n"

					; Posteingang protokollieren 2
						FileAppend, % "`t[" FaxAbsender " | " FaxNummer "]", % this.DBPath "\outlook.txt", UTF-8

					; Anhänge entpacken
						Attachments := email.Attachments
						Loop % Attachments.Count		{
							thisattachment := Attachments.Item(A_Index)
							AttachmentCounter ++
							if RegExMatch(thisattachment.DisplayName, AMFilterIn ) 	{
								Fullpath := PathToSaveTo "\" thisattachment.DisplayName
								thisattachment.SaveAsFile(FullPath)
							}

					}
				}
			}

			; Posteingang protokollieren 3
				If EMailsWithAttachment > 0
					FileAppend, % "`n" , % this.DBPath "\outlook.txt", UTF-8

				MsgBox, % "           " A_DDDD ", " A_DD "." A_MM "." A_YYYY " " A_Hour ":" A_Min "`n"   ;{
							.	 "--------------------------------------------------`n"
							. 	 (AttachmentCounter > 0 ? 	unread " ungelesene EMail" (EMailsWithAttachment > 1 ? "s" : "") "`n" : "")
							. 	 (AttachmentCounter > 0 ? 	EMailsWithAttachment " EMail" (EMailsWithAttachment > 1 ? "s" : "") " hatten Anhänge`n" : "")
							.	 (AttachmentCounter > 0 ? 	AttachmentCounter " " (AttachmentCounter > 1 ? "Anhänge wurden" : "Anhang wurde") " gespeichert`n" : "")
							.	 (AttachmentCounter = 0 ?	"              Heute gab es kein Fax!`n" : "")
							.	 "--------------------------------------------------"  (AttachmentCounter = 0 ?	"" : "`n")
							.	 (AttachmentCounter > 0 ? 	"                  Liste der Absender               `n" : "")
							.	 (AttachmentCounter > 0 ? 	"--------------------------------------------------`n" : "")
							.	 (AttachmentCounter > 0 ? 	Trim(MBoxMsg, "`n") "`n" : "")
							.	 (AttachmentCounter > 0 ? 	"--------------------------------------------------`n" : "")
							.	 (AttachmentCounter > 0 ? "                        " PathToSaveTo "`n" : "")
							.	 (AttachmentCounter > 0 ? 	"--------------------------------------------------" : "")
				;}

		}


		datestamp(nr:=1) {                                                                                                	;-- für Protokolle
			If (nr = 1)
				return (A_DD "." A_MM "." A_YYYY ", " A_Hour ":" A_Min ":" A_Min)
			else
				return (A_Hour ":" A_Min ":" A_Sec)
		}

}

ExtractAttachmentOutlook(folderName, EmailSubject, EmailSubjectData, attachmentname, PathToSaveTo) {

	EMailsWithAttachment := unread := AttachmentCounter := 0

	try ol := ComObjActive("Outlook.Application")
	catch
		try
     		this.ol := ComObjCreate("Outlook.Application")
				catch
					throw Exception("Kann nicht mit Outlook verbinden.")

	olNameSpace := ol.GetNameSpace("MAPI")
	olFolder := olNameSpace.Folders(1).Folders("Posteingang")
	MaxMails := olFolder.items.count
	Loop % olFolder.items.count 	{

		email := olFolder.items.Item(A_Index)
		If email.unread
			unread ++

		if email.unread && RegExMatch(email.Subject, EmailSubject)	{

			email.unread := 0
			Attachments := email.Attachments
			EMailsWithAttachment ++

			RegExMatch(email.Subject, EmailSubjectData, Fax)
			t .= "Absender:`t"  	FaxAbsender "`n"
			t .= "Nummer: `t" 	FaxNummer "`n"

			Loop % Attachments.Count		{
				thisattachment := Attachments.Item(A_Index)
				AttachmentCounter ++
				if RegExMatch(thisattachment.DisplayName , attachmentname ) 	{
					Fullpath := PathToSaveTo "\" thisattachment.DisplayName
					thisattachment.SaveAsFile(FullPath)
				}

			}
		}
	}

	MsgBox, % "           " A_DDDD ", " A_DD "." A_MM "." A_YYYY " " A_Hour ":" A_Min "`n"
				.	 "--------------------------------------------------`n"
				.	 MaxMails " Mails enthält der Ordner`n"
				. 	 (EMailsWithAttachment > 0 ? 	unread " ungelesene EMail" (EMailsWithAttachment > 1 ? "s" : "") "`n" : "")
				. 	 (EMailsWithAttachment > 0 ? 	EMailsWithAttachment " EMail" (EMailsWithAttachment > 1 ? "s" : "") " hatten Anhänge`n" : "")
				.	 (EMailsWithAttachment > 0 ? 	AttachmentCounter " " (EMailsWithAttachment > 1 ? "Anhänge wurden" : "Anhang wurde") " gespeichert`n" : "")
				.	 (EMailsWithAttachment = 0 ?	"Keine passenden Mails mit Anhang gefunden!`n" : "")
				.	 "--------------------------------------------------"  (EMailsWithAttachment = 0 ?	"" : "`n")
				.	 (EMailsWithAttachment > 0 ? 	"                  Liste der Absender               `n" : "")
				.	 (EMailsWithAttachment > 0 ? 	"--------------------------------------------------`n" : "")
				.	 (EMailsWithAttachment > 0 ? 	Trim(t, "`n") "`n" : "")
				.	 (EMailsWithAttachment > 0 ? 	"--------------------------------------------------`n" : "")
				.	 (AttachmentCounter > 0 ? "                        " PathToSaveTo "`n" : "")
				.	 (EMailsWithAttachment > 0 ? 	"--------------------------------------------------" : "")


}


IniReadExt(SectionOrFullFilePath, Key:="", DefaultValue:="", convert:=true) {                                          	;-- eigene IniRead funktion für Addendum

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
					MsgBox, % "Der Defaultwert <" DefaultValue "> konnte geschrieben werden.`n`n[" WorkIni "]"
			}
			else return "ERROR"
		else if InStr(OutPutVar, "%AddendumDir%")
				return StrReplace(OutPutVar, "%AddendumDir%", admDir)
		else if RegExMatch(OutPutVar, "i)^\s*(ja|nein)\s*$", bool)
				return (bool1= "ja") ? true : false

return Trim(OutPutVar)
}

IniAppend(value, filename, section, key) {                                                                                 	;-- vorhandenen Werten weitere Werte hinzufügen

	IniRead, schreib, % filename, % section, % key
	If Instr(schreib, "Error")
		schreib:=""
	IniWrite, % schreib . value, % filename, % section, % key	;, UTF-8

}

StrUtf8BytesToText(vUtf8) {                                                                                                       	;-- Umwandeln von Text aus .ini Dateien in UTF-8
	if A_IsUnicode 	{
		VarSetCapacity(vUtf8X, StrPut(vUtf8, "CP0"))
		StrPut(vUtf8, &vUtf8X, "CP0")
		return StrGet(&vUtf8X, "UTF-8")
	} else
		return StrGet(&vUtf8, "UTF-8")
}


FilePathCreate(path) {                                                                                                              	;-- erstellt einen Dateipfad falls dieser noch nicht existiert

	If !FilePathExist(path) {
		FileCreateDir, % path
		If ErrorLevel
			return 0
		else
			return 1
	}

return 1
}

FilePathExist(path) {                                                                                                                 	;-- prüft ob ein Dateipfad vorhanden ist
	If !InStr(FileExist(path "\"), "D")
		return 0
return 1
}

isFullFilePath(path) {                                                                                                               	;-- prüft Pfadstring auf die Angabe eines Laufwerkes
	If RegExMatch(path, "[A-Z]\:\\")
		return 1
return 0
}

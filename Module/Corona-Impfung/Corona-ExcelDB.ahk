; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;                                    Addendum Assistent für hausärztliche Corona Impfplanung
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;      Funktion: 				 ⚬	Impfplanung über Excel
;                                	 ⚬
;                                	 ⚬
;                                	 ⚬
;                                	 ⚬
;
;		Hinweis:
;
;		Abhängigkeiten:		siehe includes
;
;      	begonnen:       	    	21.05.2021
; 		letzte Änderung:	 	15.06.2021
;
;	  	Addendum für Albis on Windows by Ixiko started in September 2017
;      	- this file runs under Lexiko's GNU Licence
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣

	;{  Skripteinstellungen
		#NoEnv
		#Persistent
		#SingleInstance               	, Force
		#KeyHistory                  	, Off
		SetBatchLines                	, -1
		FileEncoding                 	, UTF-8
		DetectHiddenText          	, Off
		DetectHiddenWindows   	, Off
		AutoTrim                          	, On
		;ListLines                           	, Off
		SetTitleMatchMode        	, 2	              	;Fast is default
		SetTitleMatchMode        	, Fast        		;Fast is defaultListLines                        	, Off
		OnExit("DasEnde")

		; Client Namen feststellen
			global compname := StrReplace(A_ComputerName, "-")                    	; der Name des Computer auf dem das Skript läuft

		 ;Tray Icon erstellen
			If (hIconImpfen := Create_Impfen_ico())
				Menu, Tray, Icon, % "hIcon: " hIconImpfen
			else If (hIconImpfen := Create_Impfen_ico())
				Menu, Tray, Icon, % "hIcon: " hIconImpfen
            else
               Menu, Tray, Icon, % A_ScriptDir "\..\..\assets\ModulIcons\Impfen.ico"

			Menu, Tray, Add, Liste korrigieren, Impfliste_Korrektur
			Menu, Tray, Add, Warteliste prüfen, Warteliste_Korrektur
			Menu, Tray, Add, Patient suchen, Patientsuchen


		; startet die Windows Gdip Funktion
		If !(pToken:=Gdip_Startup()) {
			MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
			ExitApp
		}
	;}

	;{  Variablen
		global q                     	:= Chr(0x22)                           	; ist dieses Zeichen -> "
		global adm               	:= Object()
		global AddendumDir
		global Mdi                	:= Object()
		global XL
		global Pat
		global Impflinge
		global Statistic
		global XLBar, percent, lBEdit, lBHEdit, HeaderPos

		RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)
		workini := IniReadExt(AddendumDir "\Addendum.ini")
		adm.Default := {}
		adm.Default.Font              	:= IniReadExt("Addendum"	, "StandardFont"        	)
		adm.Default.BoldFont          	:= IniReadExt("Addendum"	, "StandardBoldFont"    	)
		adm.Default.FontSize           	:= IniReadExt("Addendum"	, "StandardFontSize"    	)
		adm.Default.FntColor        	:= IniReadExt("Addendum"	, "DefaultFntColor"     	)
		adm.Default.BGColor        	:= IniReadExt("Addendum"	, "DefaultBGColor"     	)
		adm.Default.BGColor1      	:= IniReadExt("Addendum"	, "DefaultBGColor1"   	)
		adm.Default.BGColor2      	:= IniReadExt("Addendum"	, "DefaultBGColor2"   	)
		adm.Default.BGColor3      	:= IniReadExt("Addendum"	, "DefaultBGColor3"   	)

		adm.Dir                   	:= AddendumDir
		adm.BefundOrdner	:= "M:\Befunde"
		adm.AddendumDir	:= AddendumDir
		adm.DBPath 	    		:= AddendumDir "\logs'n'data\_DB"
		adm.AlbisDBPath     	:= GetAlbisPath() "\db"

		adm.xlsmPath := IniReadExt("Addendum", "CoronaImpfXLSMPfad")
		If InStr(adm.xlsmPath, "ERROR") || !adm.xlsmPath {
			SciTEOutput("Kein Pfad zur Exceldatei gefunden: " adm.xlsmPath "`n" workini)
			ExitApp
		}
		xlsmPath := adm.xlsmPath
		SplitPath, xlsmPath, xlfilename, xlfileDir, xlfileExt, xlfilenameNoExt
		adm.xlfilename     	:= xlfilename
		adm.xlfileDir          	:= xlfileDir
		adm.xlfileExt           	:= xlfileExt
		adm.xlfilenameNoExt	:= xlfilenameNoExt

		adm.ExcelServer 	:= IniReadExt("Addendum", "CoronaImpfServer")
		If InStr(adm.ExcelServer, "Error")
			adm.ExcelServer	:= ""
		adm.ExcelClients	:= IniReadExt("Addendum", "CoronaImpfClients")
		If InStr(adm.ExcelClients, "Error")
			adm.ExcelClients	:= ""

		colN     	:= [{"start":"C", "last":"F", "name":"D", "vorname":"E"}
							, {"start":"L", "last":"O", "name":"M", "vorname":"N"}
							, {"start":"U", "last":"X", "name":"V", "vorname":"W"}]
		pseudo  	:= 1
		today     	:= A_YYYY A_MM A_DD
		xlYear   	:= 2021

		Impflinge	:= Object()
		Statistic  	:= {"wait1":0, "ready1":0, "wait2":0, "ready2":0, "week":{}}


;}

	;{ Excel über LAN automatisieren
		If  InStr(adm.ExcelServer, compname) || InStr(adm.ExcelClients, compname) {
			global TCPServer
			TCPServer := new SocketTCP()
			TCPServer.bind("addr_any", 12345) ; adm.LAN.admServer
			TCPServer.listen()
			TCPServer.onAccept := Func("OnAccept")
		}
		else {
			MsgBox, Dieser PC ist nicht als Client/Server zugelassen!`nDas Skript wird beendet.
			ExitApp
		}

	  ; Server oder Client-IPs ermitteln
		If !InStr(adm.ExcelServer, compname) {
			adm.xlServerIP := IniReadExt(adm.ExcelServer, "IP")
			If InStr(adm.xlServerIP, "Error") {
				SciTEOutput(A_ScriptName ": Server IP nicht gefunden")
				adm.xlServerIP	:= ""
			}
			else {
				SciTEOutput("trying to connect to server...")
				SendText(adm.xlServerIP, "Client [" compname "]")
			}

		} else {
			adm.xlClientIP := {}
			For ClientNR, Client in StrSplit(adm.ExcelClients, "|") {
					ClientIP := IniReadExt(Client, "IP")
					If !StrLen(ClientIP) > 0 && !InStr(ClientIP, "ERROR")
						adm.xlClientIP[Client] := ClientIP
			}
		}


	;}

	;{ Excel starten, Patientendatenbank einlesen
		SciTEOutput(" -----`n")
		XL  := new Excel(adm.xlsmPath)
		ComObjConnect(XL.xlSheet, Worksheet_Events)

		If !IsObject(Pat)
			Pat := new PatDBF()

	;}

	;{  Interskriptkommunikation

		; zum Steuern des Skriptes durch ein anderes Skript
			Gui, AMsg: New	, +HWNDhMsgGui +ToolWindow
			Gui, AMsg: Add	, Edit, xm ym w100 h100 HWNDhEditMsgGui
			Gui, AMsg: Show	, AutoSize NA Hide, % "Addendum " StrReplace(A_ScriptName, ".ahk") " Gui"
			OnMessage(0x4A	, "Receive_WM_COPYDATA")

	;}

	;{ Shellhook wird gestartet
		DllCall("RegisterShellHookWindow", UInt, hMsgGui)
		MsgNum := DllCall("RegisterWindowMessage", Str, "SHELLHOOK")
		OnMessage(MsgNum, "ShellHookProc")
	;}


;~ Func_IFLko 	:= Func("IF_ActiveLko")
;~ Func_CNR	:= Func("Chargennummern")
;~ Hotkey, IF	, % Func_IFLko
;~ Hotkey, NumpadAdd	, % Func_CNR
;~ Hotkey, ^LButton		, % Func_CNR
;~ Hotkey, IF

return
IF_ActiveLko() {

	global KK

	If WinActive("ahk_class OptoAppClass")  {

		kuerzel := AlbisGetActiveControl("contraction")
		If (RegExMatch(kuerzel, "i)^lk\w") || StrLen(kuerzel) = 0 || InStr(AlbisGetActiveWindowType(), "Karteikarte"))
			return true

	}

return false
}

Warteliste_Korrektur:               	;{

		XLBPruefung := true
		lostIDsFile := adm.DBPath "\sonstiges\COVID-19\COVID19-LostIDs.json"
		If FileExist(lostIDsFile) {
			FileGetTime, ftime, % lostIDsFile
			ftime := ConvertDBASEDate(ftime)
			MsgBox, 0x1024, % A_ScriptName, % "Die Backupdateien wurden zuletzt am: " ftime " geprüft.`nJetzt erneut prüfen?"
			IfMsgBox, No
			{
				XLBPruefung := false
				lostIDs := JSONData.Load(lostIDsFile, "", "UTF-8")
			}
		}

	; Impflinge Objekt leeren ohne löschen
		Loop, % Impflinge.Count()
			Impflinge.Pop()

	; Impflinge einsammeln
		gosub Impfliste_Korrektur

	; Namen aus den Wartelisten aller Backupdateien beziehen
		If XLBPruefung {

			; Hinweis anzeigen
				Edit_Append(lBHEdit, "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -")
				Edit_Append(lBHEdit, "             durchsuche Excel Backupdateien nach Wartelistenpatienten ")
				Edit_Append(lBHEdit, "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -`n")

				;~ csv := FileOpen(filepattern, "r", "UTF-8").Read()
				;~ For lineNr, line in csv
				firstfile := true
				lostIDs := {}
				filepattern:= "M:\Praxis\Corona-Impforganisation\Backup\*.xl*"
				Loop, Files, % "M:\Praxis\Corona Impforganisation\Backup\*.*"
				{

					If !RegExMatch(A_LoopFileName, "\.xlsx|m$")
						continue

					If !firstfile
						Edit_Append(lBHEdit, "> > > > > > > > > > > > > > > > > > > > > > > >")
					Edit_Append(lBHEdit, "[" SubStr("00" A_Index, -1) "] filename: " A_LoopFileName)

					firstfile 		:= false
					workXL  	:= new Excel(A_LoopFileFullPath)
										 workXL.Speedup(true)
					workRows 	:= workXL.UsedRows()
					IDs   		:= workXL.CopyToArr(colN[3].start      	"3:" colN[3].start . workRows)
					names 		:= workXL.CopyToArr(colN[3].name   	"3:" colN[3].name . workRows)
					vnames		:= workXL.CopyToArr(colN[3].vorname 	"3:" colN[3].vorname . workRows)

					for row, value in IDs {
						thisID := Trim(value)
						If RegExMatch(thisID, "^(\d{5})|(#\d+)$")  ;&& RegExMatch(names[row], "i)[\pL\s\-]")
							If !lostIDs.HasKey(thisID) {
								lostIDs[thisID] := names[row] ", " vnames[row]
								Edit_Append(lBHEdit, SubStr("000" lostIDs.Count(), -2) "[" lostID "] " names[row] ", " vnames[row])
							}
					}

					workXL.Speedup(false)
					workXL.CloseActiveWorkbook(false)
					workXL.Disconnect()
				}

				Edit_Append(lBHEdit, " >> Excel files parsed")
				JSONData.Save(adm.DBPath "\sonstiges\COVID-19\COVID19-LostIDs.json", lostIDs	, true,, 1, "UTF-8")

		}

		; Hinweis anzeigen
			Edit_Append(lBHEdit, "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -")
			Edit_Append(lBHEdit, "             verlorene Impflinge erneut eintragen ")
			Edit_Append(lBHEdit, "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -`n")

	; lostIDs mit vorhandenen Impflingen vergleichen
		For lostID, lostPatName in lostIDs {

			If !Impflinge.HasKey(lostID) {

				If RegExMatch(lostID, "#")
					For ImpflingID, Impfling in Impflinge
						If ((Impfling.Name ", " Impfling.Vorname) = lostPatName)
							continue

				MsgBox, 0x1024, % A_ScriptName, % "Soll [" lostID "] " lostPatName " wieder in die Warteliste aufgenommen werden?"
				IfMsgBox, Yes
				{
					InsertRow := PatientEintragen(lostID)
					Edit_Append(lBHEdit, " [" lostID "] " lostPatName " wieder aufgenommen")
				}

			}

		}

return
;}

DatenUmwandeln:                 	;{




return ;}

Patientsuchen:                        	;{

return ;}

Impfliste_Korrektur:               	;{

	; Impflinge Objekt leeren ohne löschen
		Loop, % Impflinge.Count()
			Impflinge.Pop()

		Rows         	:= XL.UsedRows()
		SheetName 	:=XL.xlSheetName
		;SciTEOutput("maxRows: " XL.UsedRows() ", " SheetName)                              ; xlUp = -4162)

	; Fortschrittsanzeige
		If !WinExist("Corona-Impfplaner ahk_class AutoHotkeyGui")  || !IsObject(XLBar)
			XLBarGui()
		XLBar.Set(0, "sammle Daten [" SubStr("00000", -1*(StrLen(Rows)-1)) "/" Rows "]")

	; Hinweis anzeigen
		Edit_Append(lBHEdit, " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -")
		Edit_Append(lBHEdit, "aktive Tabelle: " SheetName)
		Edit_Append(lBHEdit, " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -")

	; Korrekturen, Daten sammeln, fehlende Patientendaten ergänzen
		startTime := A_TickCount
		WLPat := WLZIPat := EIPat := ZIPat := 0
		XL.Speedup(true)
		Loop % Rows {

			xlRow       	:= A_Index
			percent 	:= Round((xlRow*100)/Rows, 1)
			RowVon 	:= SubStr("00000" xlRow, -1*(StrLen(Rows)-1)) "/" Rows

			XLBar.Set(Floor(percent), "[ Zeile: " RowVon " | Patienten: " Impflinge.Count() . " | Erstimpfung: " EIPat ", wartend: " WLPat " | Zweitmpfung: " ZIPat ", wartend:" WLZIPat "] ")

			For colidx, col in colN {

					; -----------------------------------------------------------------------------------------------------------------
					; Zellen identifizieren
						ColorIdx := XL.ColorIndex(col.start . xlRow)
						If !RegExMatch(ColorIdx, "(17|19|24|40)")
							continue

					; -----------------------------------------------------------------------------------------------------------------
					; Datum der Impfung
					If (ColorIdx = 24) {
						IDatum := XL.CopyToVar("D"	xlRow)
						 if RegExMatch(IDatum, "\d\d\.\d\d\.\d*\d*\d*\d*") {
							ImpfTag := IDatum
							continue
						}
					}

					; -----------------------------------------------------------------------------------------------------------------
					; Kalenderwoche
					If (ColorIdx = 17) {
						KWText := XL.CopyToVar("B"	xlRow)
						If RegExMatch(KWText, "i)KW") {
							WeekNr   	:= xl.CopyToVar("A"	xlRow)
							KWeek  	:= "#" SubStr("0" WeekNr, -1)
							If !IsObject(Statistic.week[KWeek])
								Statistic.week[KWeek] := {	"I1":0, "I1IS1":0, "I1IS2":0, "I1min":200, "I1max":0
																	, 	"I2":0, "I2IS1":0, "I12S2":0, "I2min":200, "I2max":0
																	, "age":0, "agemin":200, "agemax":0}
							xlYear += (KWeek < KWeekLast) ? 1 : 0
							KWeekLast := KWeek
							continue
						}
					}

					; -----------------------------------------------------------------------------------------------------------------
					; Patientendaten
					If RegExMatch(ColorIdx, "(19|40)") {

						CellText	:= xl.CopyToArr(col.start . xlRow ":" col.last . xlRow)
						If !RegExMatch(CellText[2] . CellText[3], "i)[\pL\s\-]+") ; kein Name dann weiter
							continue
						dispRow 	:= SubStr("0000 " xlRow, -1*(StrLen(Rows)-1))
						Name		:= Trim(CellText[2])
						Vorname	:= Trim(CellText[3])
						PatID    	:= Trim(CellText[1])
						noPatID	:= RegExMatch(PatID, "^\d+$") ? false : true
						PatID    	:= noPatID ? Pat.StringSimilarityID(Name, Vorname) : PatID
						ALTER    	:= RegExMatch(Trim(CellText[4]), "^\?$") ? "" : CellText[4]

					; -----------------------------------------------------------------------------------------------------------------
					; PatID ergänzen falls diese fehlt
						age_added := ID_added := false
						If noPatID && PatID {
							ID_added := true
							xl.SetCell(col.start      	. xlRow, PatID)
							xl.SetCell(col.name    	. xlRow, Pat.Get(PatID, "NAME"))
							xl.SetCell(col.vorname 	. xlRow, Pat.Get(PatID, "VORNAME"))
						}

					; -----------------------------------------------------------------------------------------------------------------
					; fehlende Patientendaten ergänzen
						If !ALTER && PatID && Name && Vorname  {
							age_added := true
							xl.SetCell(col.last . xlRow, (ALTER := HowLong(Pat.Get(PatID, "GEBURT"), A_YYYY . A_MM . A_DD).years))
						}

					; -----------------------------------------------------------------------------------------------------------------
					; Ergänzungen ausgeben
						If (ID_added || age_added) {
							Edit_Append(lBHEdit,
											. dispRow  " | "
											. 	(age_added                    	? "Alter "	: "" )
											. 	(ID_added && age_added	? "und " 	: "" )
											. 	(ID_added                       	? "ID"    	: "" )
											. 	" ergänzt: " PatID ", "  Name ", " Vorname ", " ALTER)
						}

					; -----------------------------------------------------------------------------------------------------------------
					; Pseudo-ID für unbekannte Impflinge
						else If !PatID && Name && Vorname  {
							PatID := "#" (pseudo++)
							For ImpfPatID, Data in Impflinge
								If (Data.Name = Name && Data.Vorname = Vorname) {
									PatID := ImpfPatID
									break
								}
							xl.SetCell(col.start 	. xlRow, PatID)
							xl.SetCell(col.last 	. xlRow, "?")
							xl.Font(col.start 		. xlRow, "jR jvC")
							xl.Font(col.last 		. xlRow, "jC jvC")
							Edit_Append(lBHEdit, dispRow " | Patient unbekannt: " PatID ", "  Name ", " Vorname ", " ALTER)
						}

					; -----------------------------------------------------------------------------------------------------------------
					; Impflinge inventarisieren
						If RegExMatch(PatID, "\#*\d+") {

								If !Impflinge.haskey(PatID)
									Impflinge[PatID] := {	"WL"           	: []                 	; Warteliste
																,	"I1"           	: []                  	; 1.Impfung
																, 	"I2"           	: []                  	; 2.Impfung
																, 	"Name"        	: (!RegExMatch(PatID, "\#") ? Pat.Get(PatID, "NAME")                                       	: Name     	)
																, 	"Vorname"    	: (!RegExMatch(PatID, "\#") ? Pat.Get(PatID, "VORNAME")                                	: Vorname 	)
																,	"Geburt"     	: (!RegExMatch(PatID, "\#") ? ConvertDBASEDate(Pat.Get(PatID, "GEBURT"))      	: ""            	)
																, 	"Alter"        	: (!ALTER ? HowLong(Pat.Get(PatID, "GEBURT"), A_YYYY . A_MM . A_DD).years	: ALTER     	)}

								If (colidx = 3)            	{

									comment  	:= xl.CopyToVar("Y" xlRow)
									Impflinge[PatID].WL.Push({"KW"    	: KWeek
																		, 	"cell"    	: xlRow
																		, 	"cmt" 	: comment})
									WLPat ++

								}
								else                          	{

									comment  	:= xl.CopyToVar((colidx=1 ? "G"	: "P"	) . xlRow)
									location     	:= xl.CopyToVar((colidx=1 ? "H"	: "Q"	) . xlRow)
									vaccination 	:= xl.CopyToVar((colidx=1 ? "I"		: "R"	) . xlRow)
									RegExMatch(vaccination, "\s*(?<accine>[A-Z]+)\s+(?<ch>[A-Z\d]+)", v)

									Impflinge[PatID]["I" colidx].Push({"day"  	: ImpfTag
																				, 	"cmt"   	: comment
																				, 	"loc"    	: location
																				, 	"vce"    	: vaccine
																				, 	"ch"     	: vch
																				, 	"cell"   	: xlRow})

									VacReady := VacReady(ImpfTag)
									GuiControl, IMPF:, lBInfo2, % "Impfdatum: " ImpfTag ", verstrichene Zeit: " Round((A_TickCount-startTime)/1000, 1) "s"

									; Zählung
									If (colidx = 1) {

										If VacReady
											EIPat ++
										else
											WLPat ++

									} else if (colidx = 2) {

										If VacReady
											ZIPat ++
										else
											WLZIPat ++

									}

								}
						}

					}

			}

		}

		XL.Speedup(false)

	;{  Statistiken

		Statistic.Patienten                   	:= Impflinge.Count()
		Statistic.Erstimpfung_fertig      	:= EIPat
		Statistic.Erstimpfung_wartend  	:= WLPat
		Statistic.Zweitmpfung_fertig    	:= ZIPat
		Statistic.ZweitImpfung_wartend 	:= WLZIPat

	; fehlende Zweittermine suchen, doppelte Einträge löschen, Impfabstände kontrollieren
	; Statistik erstellen für Infoseite

		today := A_YYYY A_MM A_DD
		vaccination_gaps := {"AZ":12, "CMY":6}
		Edit_Append(lBHEdit, " ")
		Edit_Append(lBHEdit, " - - - - - - - - - - - - - - - - - - - - - - - - - - - - -")
		Edit_Append(lBHEdit, " - - -                Fehlerliste               - - -")
		Edit_Append(lBHEdit, " - - - - - - - - - - - - - - - - - - - - - - - - - - - - -")

		For PatID, data in Impflinge {

				name    	:= Data.name,
				vorname 	:= Data.Vorname
				fullid 		:= PatID ", " Pat.Get(PatID, "NAME") "," Pat.Get(PatID, "VORNAME")

			; doppelt/mehrfach in der Warteliste
				If (data.WL.Count() > 0) {

					Loop % data.WL.Count()-1 {       ; löscht den Inhalt der 5 Zellen
						WL := Impflinge[PatID].WL.Pop()
						KW := WL.KW
						;xl.SetCell(colN[3].start . WL.cell ":Y" WL.cell, "")
						Edit_Append(lBHEdit, "doppelt in Warteliste: " fullid " | Zeile:  " WL.cell )
					}

				}

			; noch in der Warteliste
				If (data.WL.Count() > 0) && (data.I1.Count() > 0 || data.I2.Count() > 0) {

					Loop % data.WL.Count() {
						WL := Impflinge[PatID].WL.Pop()
						;xl.SetCell(colN[3].start . WL.cell ":Y" WL.cell, "")
						Edit_Append(lBHEdit, "geimpft u. in Warteliste: " fullid " | Zeile:  " xlRow )
					}

				}

			; 1.Termin wurde mehrfach vergeben
				If (data.I1.Count() > 1) {

					vcnDays := ""
					For idx, vcn in data.I1
						vcnDays .= vcn.day " [" vcn.cell "]" ", "

					If vcnDays
						Edit_Append(lBHEdit, "1.Termin mehrfach vergeben: " fullid " | Info:  " RTrim(vcnDays, ", "))

				}

			; 2.Termin wurde mehrfach vergeben
				If (data.I2.Count() > 1) {

					vcnDays := ""
					For idx, vcn in data.I2
						vcnDays .= vcn.day " [" vcn.cell "]" ", "

					If vcnDays
						Edit_Append(lBHEdit, "2.Termin mehrfach vergeben: " fullid " | Info:  " RTrim(vcnDays, ", "))

				}

			; ohne 2.Termin eingetragen ConvertToDBASEDate(Date)
				If (data.I1.Count() > 0 && data.I2.Count() = 0) {

					vcnDays := ""
					For idx, vcn in data.I1
						If VacReady(vcn.day)
							vcnDays .= vcn.day " [" vcn.cell "]" ", "

					;~ If vcnDays
						;~ Edit_Append(lBHEdit, "ohne 2.Termin: " fullid " | Info:  " RTrim(vcnDays, ", "))
				}

			; Impfabstand nicht passend
				If (data.I1.Count() > 0 && data.I2.Count() > 0) {

					For idx, vcn in data.I1 {

						;SciTEOutput(vcn.day ", " Data.I2[idx].day)
						vacDay1	:= ConvertToDBASEDate(vcn.day)
						vacDay2	:= ConvertToDBASEDate(Data.I2[idx].day)
						If vacDay1 && vacDay2 {
							weekgap := Floor(DaysBetween(vacDay1, vacDay2)/7)
							;~ If (weekgap < vaccination_gaps[vcn.vce])
								;~ Edit_Append(lBHEdit, "Impfabstand zu gering: " fullid
															;~ . 	" | Info:  [Soll: " vce " " vaccination_gaps[vcn.vce] "W, Ist: " weekgap "W] "
															;~ .  	"| 1. " vcn.day " , 2. " Data.I2[idx].day " [" vcn.cell "]")

						}

					}

				}

			; Statistik
				If (data.I1.Count() > 0|| data.I2.Count() > 0) {

					If Data.I1[1].day {

						vacDay         	:= ConvertToDBASEDate(Data.I1[1].day)
						KWeek           	:= "#" SubStr("0" WeekOfYear(vacday), -1)

						I1min             	:= Statistic.week[KWeek].I1min
						I1max             	:= Statistic.week[KWeek].I1max
						agemin         	:= Statistic.week[KWeek].agemin
						agemax         	:= Statistic.week[KWeek].agemax

						Statistic.ready1	:= Statistic.week[KWeek].I1 += (vacDay <= today) ? 1 : 0
						Statistic.wait1 	+= (vacDay > 	today) ? 1 : 0

						Statistic.week[KWeek].age      	+=	Data.I1[1].ALTER
						Statistic.week[KWeek].agemin  	:= 	(agemin 	> 	Data.ALTER) ? Data.ALTER : agemin
						Statistic.week[KWeek].agemax 	:= 	(agemax 	< 	Data.ALTER) ? Data.ALTER : agemax
						Statistic.week[KWeek].I1min   	:= 	(I1min    	> 	Data.ALTER) ? Data.ALTER : I1min
						Statistic.week[KWeek].I1max  	:= 	(I1max    	< 	Data.ALTER) ? Data.ALTER : I1max

					}

					If Data.I2[1].day {

						vacDay         	:= ConvertToDBASEDate(Data.I2[1].day)
						KWeek           	:= "#" SubStr("0" WeekOfYear(vacday), -1)

						I2min             	:= Statistic.week[KWeek].I2min
						I2max             	:= Statistic.week[KWeek].I2max
						agemin         	:= Statistic.week[KWeek].agemin
						agemax         	:= Statistic.week[KWeek].agemax

						Statistic.ready2	:= Statistic.week[KWeek].I2 += (vacDay <= 	today) ? 1 : 0
						Statistic.wait2 	+= (vacDay > today) ? 1 : 0

						Statistic.week[KWeek].age  		+= Data.I2[1].ALTER
						Statistic.week[KWeek].agemin 	:= 	(agemin 	> 	Data.ALTER) ? Data.ALTER : agemin
						Statistic.week[KWeek].agemax 	:= 	(agemax	< 	Data.ALTER) ? Data.ALTER : agemax
						Statistic.week[KWeek].I2min   	:= 	(I2min     	> 	Data.ALTER) ? Data.ALTER : I2min
						Statistic.week[KWeek].I2max  	:= 	(I2max    	< 	Data.ALTER) ? Data.ALTER : I2max

				}

			}

		}

		; weitere Statistiken
		/*

			ist mit Biontech geimpft, 6 Wochen Abstand,keinerlei Nebenwirkungen
			beide Corona Impfungen schon im März bekommen , keine Probleme

		*/

	;}

		XL.Speedup(false)

		Edit_Append(lBHEdit, " ")
		Edit_Append(lBHEdit, " - - - - - - - - - - - - - - - - - - - - - - - - - - - - -")
		Edit_Append(lBHEdit, " - - -       Auswertung beendet       - - -")
		Edit_Append(lBHEdit, " - - - - - - - - - - - - - - - - - - - - - - - - - - - - -")
		Edit_Append(lBHEdit, " Zeit für Ausführung: "       	Round((A_TickCount-startTime)/1000, 1) "s" )
		Edit_Append(lBHEdit, " Rows: "                             	Rows)
		Edit_Append(lBHEdit, " Größe Statistic: "             	Statistic.Count())
		Edit_Append(lBHEdit, " Anzahl der Impflinge: "    	Impflinge.Count())

		JSONData.Save(adm.DBPath "\sonstiges\COVID-19\COVID19-Impflinge.json"   	, Impflinge, true,, 1, "UTF-8")
		JSONData.Save(adm.DBPath "\sonstiges\COVID-19\COVID19-Impfstatistik.json"	, Statistic	, true,, 1, "UTF-8")
		;JSONData.Save(adm.DBPath "\PatData\PatientDBF.json"                                     	, Pat      	, true,, 1, "UTF-8")

	; speedup aus : 832 Zeilen in  Sekunden

return  ;}


; --- Skriptfunktionen                	;{

OnAccept()                                        	{

	newTCP := TCPServer.accept()
	newTCP.SendText("Successful Connection!")
	recv := newTCP.recvText()
	SciTEOutput("received: " newTCP.recvText())

}

SendText(to, Text) {

	global DC_SERV
	global DC_CLI

	x := new SocketTCP()
	Connected := x.Connect([to, 12345])
	x.SendText(Text)
	x.Disconnect()

	If Connected
		SciTEOutput("connected to: " to " (" Text ")")
	else
		SciTEOutput("can't connect to: " to)

}

VacReady(VaccinationDay)                 	{                         	; Impfung abgeschlossen oder Termin zur Impfung
	static today := A_YYYY A_MM A_DD ;"000000"
	If (StrLen(VaccinationDay) = 0)
		return false
	vacDay := ConvertToDBASEDate(VaccinationDay)
	;SciTEOutput(vacDay " = " today)
return vacDay <= today ? true : false
}

Patientsuchen(PatStr)                          	{

		global PatIDCalled, call, TakeMeMsg, MdiPos

	; Name oder PatID übergeben
		If RegExMatch(PatStr, "^\d+$") {
			PatID := PatStr
		} else If (!PatStr || RegExMatch(PatStr, "^\s*#")) {
			return
		} else {
			DbID := Pat.StringSimilarityID(StrSplit(PatStr, ",").1, StrSplit(PatStr, ",").2)
			PatID := IsObject(DbID) ? DbID[1] : DbID
		}

	; Zelldaten auslesen ;{
		IDCol    	:= []
		cellfound	:= []
		ImpfTage	:= []
		Rows     	:= XL.UsedRows()
		IRow      	:= 0
		vacdate 	:= ""

							 XL.Speedup(true)
		ImpfCol	:= XL.CopyToArr("D3:D" 	Rows)
		IDCol[1] 	:= XL.CopyToArr("C3:C" 	Rows)
		IDCol[2] 	:= XL.CopyToArr("L3:L"  	Rows)
		IDCol[3] 	:= XL.CopyToArr("U3:U" 	Rows)
							 XL.Speedup(false)

		For impfRow, CellValue in ImpfCol {

			If RegExMatch(CellValue, "\d+\.\d+\.\d+") {
				vacdate := CellValue
				If (DayOfWeek(vacdate) = "Di")
					ImpfDienstag := CellValue
				else (DayOfWeek(vacdate) = "Mi")
					ImpfMittwoch := CellValue
			}
			else if RegExMatch(CellValue, "i)Dienstag.*Mittwoch")
				vacdate := ImpfDienstag "/" ImpfMittwoch

			ImpfTage[impfRow] := vacdate

		}

		For ColIndex, CellValues in IDCol {

			For rowIndex, value in CellValues
				If (value = PatStr) {
					cellfound[ColIndex] := rowIndex
					break
				}

		}

	;}

	; wo gefunden
		msg :=""
		For colidx, rowNr in cellfound {

			If (colidx = 1)
				msg := "1.Impfung: " ImpfTage[rowNr] "(" DayOfWeek(ImpfTage[rowNr]) ")"
			else if (colidx = 2)
				msg .= "`n2.Impfung: " ImpfTage[rowNr] "(" DayOfWeek(ImpfTage[rowNr]) ")"
			else if (colidx = 3)
				msg .= "`nWarteliste: " ImpfTage[rowNr] "(" DayOfWeek(ImpfTage[rowNr]) ")"

		}

	; Hinweis im Albisfenster einblenden ;{
		MdiPos	:= GetWindowSpot(AlbisMDIClientHandle())

		pre := Pat.Get(PatID, "Name") ", " Pat.Get(PatID, "VORNAME") ", " ConvertDBASEDate(Pat.Get(PatID, "GEBURT")) "`n"
		msg := pre . (cellfound.Count() > 0 ? LTrim(msg, "`n") : "ist nicht im Impfplaner.`nKlick mich um Patient zu übernehmen!")
		TakeMeMsg := pre

		If IsObject(call) {                      ; nur ein Fenster erlauben
			call.DestroyWindow()
			call := ""
		}
		call := TextRender(msg, "time:" (cellfound.Count() > 0 ? 10000 : 30000) " color:OrangeRed radius:10 margin:(5px 5px)"
								, 	"color:White s20 font:(Futura Md Bt)")

		Win := GetWindowSpot(call.hwnd)
		SetWindowPos(call.hwnd, MdiPos.X+MdiPos.CW-Win.W-20, MdiPos.Y+MdiPos.H-Win.H, Win.W, Win.H)
		PatIDCalled := cellfound.Count() = 0 ? PatID : 0

		;}

		;XL.xlsheets.Controls("PatName")
}

PatientEintragen(IDParam:="")              	{

	; ColorIndex 19=1.Impfung, 40=2.Impfung, 17=Wochenanfang, 24=Impftag, 16/15=gesperrter Impfeintrag, 48=gesperrter Impftag,

		global PatIDCalled, xlWeeks, TakeMeMsg, call, callS, MdiPos
		static WLPlace

		If !IsObject(XL) {
			MsgBox, Es besteht keine Verbindung zu Excel
			return
		}

		If RegExMatch(IDParam, "^\d+$") {
			SciTEOutput("IDParam: " IDParam)
			PatIDCalled := IDParam
		}

	; indiziert einmalig die Tabelle
		Rows := XL.UsedRows()
		If !IsObject(xlWeeks) {

			xlWeeks 	:= Object()
			WeekNr 	:= XL.CopyToArr("A3:A" 	Rows)
			KWString 	:= XL.CopyToArr("B3:B" 	Rows)
			Loop % Rows {
				If RegExMatch(KWString[A_Index], "i)\s*KW\s*$")
					xlWeeks[WeekNr[A_Index]] := A_Index
			}

		}

	; Pat. ist eingetragen
		If !PatIDCalled {
			errorText := TextRender("Fehler beim Eintragen des Patienten.`nKeine ID Nummer!"
						    			, "time:4000 color:Red radius:10 margin:(5px 5px)", 	"color:Yellow s" 12 " font:(Futura Md Bt)")
			Return
		}

	; Nachfragen für welchen Impfstoff eingetragen werden soll
		ImpfstoffErfragen:
		InputBox, Impfstoff, Corona Impfliste, % "Welchen Impfstoff soll der Patient bekommen?`n- gib eine 1 ein für Biontech`n- eine 2 für AstraZeneca",,400,200,,,,30
		If (ErrorLevel = 1)
			return
		If !RegExMatch(Impfstoff, "^(1|2)$")
			goto ImpfstoffErfragen

		If (Impfstoff = 1) {
			xlColor:=19
			Impfstoff:="BionTech"
		}
		else If (Impfstoff = 2) {
			xlColor:=40
			Impfstoff:="AstraZeneca"
		}

	; Übernahmehinweis ;{
		If IsObject(call)                       ; nur ein Fenster erlauben
			call.DestroyWindow()
		If IsObject(callS)                       ; nur ein Fenster erlauben
			callSOld := callS

		callS := TextRender("Suche einen Platz in der Impfliste für eine `n Impfung mit " Impfstoff " für`n"  RTrim(TakeMeMsg, "`n")
								,  "time:30000 color:LightGreen radius:10 margin:(5px 5px)"
								, 	"color:Black s20 font:(Futura Md Bt)")

		Win     	:= GetWindowSpot(callS.hwnd)
		MdiPos	:= GetWindowSpot(AlbisMDIClientHandle())
		SetWindowPos(callS.hwnd, MdiPos.X+MdiPos.CW-Win.W-20, MdiPos.Y+MdiPos.H-Win.H, Win.W, Win.H)

		If IsObject(callSOld)   {          ; nur ein Fenster erlauben
			callSOld.DestroyWindow()
			callSOld := ""
		}
	;}

		If !adm.ExcelServer {
			FileAppend, % A_YYYY A_MM A_DD ":" A_Hour A_Min A_Sec "|"
								. PatIDCalled "|"
								. Pat.Get(PatIDCalled, "NAME") "|"
								. Pat.Get(PatIDCalled, "VORNAME") "|"
								. Impfstoff "`n" , % adm.DBPath "\sonstiges\COVID19-neue Impflinge.txt"
		}

	; Start der Suche mit der aktuellen Kalenderwoche
		xlRow	:= WLPlace > 0 ? WLPlace : xlWeeks[WeekOfYear(A_YYYY A_MM A_DD)]

	; zwei Bereiche erfassen
		XL.Speedup(true)
		EIDs           	:= XL.CopyToArr("C1:C" 	Rows)
		ENames    	:= XL.CopyToArr("D1:D" 	Rows)
		WLIDs       	:= XL.CopyToArr("U1:U" 	Rows)
		WLNames 	:= XL.CopyToArr("V1:V" 	Rows)

	; erfasst nur die Zellen mit ColorIndex 19 (1.Impfung und Warteliste). Sucht dabei den letzten eingetragenen Patienten in der Warteliste)
		while (xlRow <= Rows) {

			If RegExMatch(XL.xlsheet.Range("C" xlRow).Interior.ColorIndex, xlColor) && RegExMatch(XL.xlsheet.Range("U" xlRow).Interior.ColorIndex, xlColor) {

				If !EIDs[xlRow] && !ENames[xlRow] && !WLIDs[xlRow] && !WLNames[xlRow] {

					WLPlace := xlRow + 1

					SciTEOutput("Schreibe nach Excel: " PatIDCalled )

					XL.WinActivate()
					XL.Activate()
					XL.ScrollRow(xlRow-8)
					XL.SetCell("U" 	xlRow, PatIDCalled)
					XL.SetCell("V" 	xlRow, Pat.Get(PatIDCalled, "NAME"))
					XL.SetCell("W" 	xlRow, Pat.Get(PatIDCalled, "VORNAME"))
					XL.SetCell("X" 	xlRow, HowLong(Pat.Get(PatIDCalled, "GEBURT"), A_YYYY . A_MM . A_DD).years)
					XL.SelectRange("Y" xlRow)

					If IsObject(call)                   ; nur ein Fenster erlauben
						callOld := call

					xlPos	:= XL.GetSheetPos()
					if (xlPos.W <= 1990) {
						trFSize	:= 11
						xminus	:= 300
					} else {
						trFSize := 22
						xminus := 400
					}

					call := TextRender(TakeMeMsg "Platz in Zeile: " xlRow " gefunden.`nVergiss nicht einen Kommentar zu hinterlassen! "
												, "time:20000 color:Yellow radius:10 margin:(5px 5px)"
												, 	"color:Black s" trFSize " font:(Futura Md Bt)")

					Win    	:= GetWindowSpot(call.hwnd)
					SetWindowPos(call.hwnd, xlPos.X + xlPos.CW - Win.W - xminus , xlPos.Y + 20, Win.W, Win.H)

					If IsObject(callS)            	{          ; nur ein Fenster erlauben
						callS.DestroyWindow()
						callS := ""
					}
					If IsObject(callOld)           	{          ; nur ein Fenster erlauben
						callOld.DestroyWindow()
						callOld := ""
					}

					break

				}

			}

			xlRow ++

		}

		XL.Speedup(false)

return xlRow
}

ShellHookProc(lParam, wParam)        	{                       	; Startet oder entfernt einen WinEventHook bei Erscheinen oder Schließen des Albisprogramms

		global call, callS
		static XLWasClosedBefore := false

		Critical	;, 50

	; return on empty callback parameters ;{
		If (wParam = 0)
			return 0
		If (StrLen(class := WinGetClass(wparam)) = 0)
			class := GetClassName(wParam)
		If (StrLen(class . (Title := WinGetTitle(wparam))) = 0)
			return 0
	;}

	;
		If   (lParam = 1) {

			If InStr(class, "XLMAIN") && WinExist(adm.xlfilename " ahk_class XLMAIN") {

				;SHookHwnd := Format("0x{:x}", wparam)
				XL := new Excel(adm.xlsmPath)
				ComObjConnect(XL.xlSheet, Worksheet_Events)

			}

		}
		else if (lParam = 2) && IsObject(XL) && !WinExist(adm.xlfilename " ahk_class XLMAIN")	{	                                               	;

			SciTEOutput("Excel Tabelle wurde geschlossen!")

			; Angezeigtes schliessen
				Worksheet_Events.DestroyTextRender()
				If IsObject(call)
					call.DestroyWindow()
				If IsObject(callS)
					callS.DestroyWindow()

			; Event-Verbindung beenden
				ComObjConnect(XL.xlSheet)

			; Excel-COM-Objekt Verbindung schliessen und Objekt leeren
				XL.Disconnect()
				XL := ""

			return

		}


return
}

CheckFileChanges(fullfilepath)              	{

return
}

DasEnde(ExitReason, ExitCode)             	{

	XL.Speedup(false)
	XL.Disconnect()

	OnExit("DasEnde", 0)

ExitApp
}

Create_Impfen_ico()                             	{

VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACHh4fT09P////////////Pz88AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACHh4fT09P////////////Pz88AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADU1NT////////////Nzc0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACHh4fm5ub////////Nzc0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACOjo7j4+PLy8sAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB///////8AAD///////wAAn///////AACP//////8AAMP//////wAA4P//////AADwP/////8AAPgOH////wAA+AQP////AAD8AAf///8AAPwAA////wAA/AAB////AAD+AAD///8AAP8AAH///wAA/wAAP///AAD+AAAf//8AAPwAAA///wAA/AAAB///AAD8AAAD//8AAP4AAAH//wAA/wAAAP//AAD/gAAAf/8AAP/AAAA//wAA/+AAAB//AAD/8AAAD/8AAP/4AAAH/wAA//wAAAP/AAD//gAAAf8AAP//AAAA/wAA//+AAAD/AAD//8AAAP8AAP//4AAB/wAA///wAAP/AAD///gAB/8AAP///AAH/wAA///+AAP/AAD///8AAfAAAP///4BA4AAA////wOBAAAD////h8AEAAP/////4AwAA//////wHAAD//////A8AAP/////4HwAA//////A/AAD/////8H8AAP/////g/wAA//////H/AAA="
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return -1
VarSetCapacity(Dec, DecLen, 0)
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return -2
hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
DllCall("gdiplus\GdipCreateHICONFromBitmap", A_PtrSize ? "UPtr" : "UInt", pBitmap, A_PtrSize ? "UPtr*" : "uint*", hIcon)
DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
Return hIcon
}

;}

; --- EXCEL Funktionen              	;{

class Worksheet_Events                       	{

		__Call(Event, Args*) {

			SciTEOutput("Event ist Object: " (IsObject(Event) ? 1 : 0 ) ", " this[Event])
			If !IsFunc(this[Event]) {
				params :=""
				For key, obj in Args
					params .= key " | "
				MouseGetPos, mx, my
				ToolTip, % params , % mx, % my+20, 15

			}

		}

		SelectionChange(cell) {

				static colRange 	:= ["ABCDEFGHI", "JKLMNOPQ", "TUVWXY"]
				static colRead 	:= [["C", "F"], ["L", "O"], ["U", "W"]]
				static Telefon    	:= ["TELEFON", "TELEFON2", "FAX"]
				static mobil   	:= "
											(LTrim Join
											01511|01512|01514|01515|01516|01517|0160|0170|0171|0175|
											01520|01522|01523|01525|0162|0172|0173|0174|
											01570|01573|01575|01577|01578|0163|
											0177|0178|01590|0176|0179
											)"

			; aktive Spalte und Zeile ermitteln    	;{
				RegExMatch(cell.address[0,0], "(?<Col>[A-Z]+)(?<Row>\d+)", active)
				If (StrLen(activeCol) > 1) || (activeRow <=2) {
					return
                 }
				found := false
				For colidx, charRange in colRange {
					If InStr(charRange, activeCol) {
						found := true
						break
					}
				}
				If !found {
					XL.SetCell("$E$2:$Q$2"	, "")
					return
				}
			;}

			; Daten zusammen stellen              	;{
				values    	:= XL.CopyToArr(colRead[colidx][1] . activeRow ":" colRead[colidx][2] . activeRow)
				PatID     	:= values.1
				Name     	:= values.2
				Vorname 	:= values.3
				Alter     	:= values.4
				If !RegExMatch(PatID, "^\d+$") || !RegExMatch(name . Vorname, "^[A-ZÄÖÜ][\p{L}\-\s]+$") {
					XL.SetCell("$E$2:$Q$2"	, "")
					return
				}
			;}

			; Telefonnummern
				Tele := []
				For TeIidx, TelName in Telefon {

					TelNr := Pat.Get(PatID, TelName)
					If RegExMatch(TelNr, "^[\s0-9\\\-\+]+$") {

						TelNr := RegExReplace(Trim(TelNr), "[\s\\\-\/\(\)]")
						If (StrLen(TelNr) >= 10) {

							RegExMatch(TelNr, "(?<Pre>" mobil ")(?<Nr>\d+)", Phone)
							If StrLen(PhoneNr) = 8
								PhoneNr := SubStr(PhoneNr, 1, 4) " " SubStr(PhoneNr, 5, 2) " " SubStr(PhoneNr, 7, 2)
							else if StrLen(PhoneNr) = 7
								PhoneNr := SubStr(PhoneNr, 1, 3) " " SubStr(PhoneNr, 4, 2) " " SubStr(PhoneNr, 6, 2)

						} else {

							PhonePre	:= !RegExMatch(TelNr, "^0") ? "033056" : ""
							PhoneNr	:= TelNr

						}
						Tele.push({"pre": PhonePre, "nr": PhoneNr})
					}
				}

			; Telefonnummern anzeigen
				info_PatName := "[" PatID "] " Name ", " Vorname ", " ConvertDBASEDate(Pat.Get(PatID, "GEBURT"))
				XL.SetCell("$E$2", info_PatName)
				XL.SetCell("$H$2", (Tele[1].pre ? "(" Tele[1].pre ") " : "") Tele[1].nr)
				XL.SetCell("$K$2",	(Tele[2].pre ? "(" Tele[2].pre ") " : "") Tele[2].nr)
				XL.SetCell("$N$2", (Tele[3].pre ? "(" Tele[3].pre ") " : "") Tele[3].nr)
				XL.SetCell("$Q$2", info_Mail)

		}

		DestroyTextRender() {
			If IsObject(this.tr)                   ; nur ein Fenster erlauben
				this.tr.DestroyWindow()
		}

}

class Application_Events                       	{


		__Call(Event, Args*) {

			SciTEOutput("Event: " Event ", " this[Event])
			If !IsFunc(this[Event]) {
				params :=""
				For key, obj in Args
					params .= key " | "
				MouseGetPos, mx, my
				ToolTip, % params , % mx, % my+20, 15

			}

		}
}

class Excel                                        		{

		/* weiterer verwendbarer Code

				; https://autohotkey.com/board/topic/69033-basic-ahk-l-com-tutorial-for-excel/
					- xlSheet.Range("C:C").PasteSpecial(-4163)
					- xlSheet.CutCopyMode := False
					- xlSheet.Range("A:A").Copy
					- xlSheet.Range("C" . A_Index).Interior.ColorIndex := 3
					- ComObjType(xlSheet, "Name")
					- Xl.Sheets(1).Move(ComObjMissing(), Xl.Sheets("Sheet3")) ;// .Move(Before, After)
					- objExcel.Worksheets(1).Cells(1, "A").End(-4121).Select                                                 ; (-4121) which stands for down
					- 'Workbooks("BOOK1.XLS").Close SaveChanges:=False' 	= 	X1.Workbooks("BookName.xls").Close(False)
					- Xl.Visible := True
					- WorkBook Events: 	oWorkbook := oExcel.ActiveWorkbook
													ComObjConnect(oWorkbook, Workbook_Events)
													BeforeClose(Cancel) {
														msgbox % "closing"
													}


		 */

		__New(Path:="", sheet:="ActiveSheet", objerror:=1, Visible:=1, DisplayAlerts:=1) {


			ComObjError(objerror)
			this.objerror := objerror

			SplitPath, Path, OutFileName, OutDir, OutExtension, OutNameNoExt
			this.Path    	:= Path
			this.xlName 	:= OutFileName
			Found        	:= false

			Try{
				this.xl 	:= ComObjActive("Excel.Application")
				for xlWB in this.xl.Workbooks {
					if (xlWB.Name = OutFileName) {
						Found := true
						break
					}
				}
			}
			Catch
				this.xl := ComObjCreate("Excel.Application")

			if !Found
				xlWB := this.xl.Workbooks.Open(this.Path)

			this.xl.Visible         	:= Visible
			this.xlApplication     	:= this.xl.Application
			;~ this.xl.Sheets(sheet).Select
			this.xlSheet             	:= this.xl.Application.ActiveSheet
			this.xlWorkbook     	:= IsObject(xlWB) ? xlWB : this.xl.Application.ActiveWorkbook
			this.xlSheetName	 	:= this.xlSheet.name

			this.xl.Application.DisplayAlerts := DisplayAlerts ;turn off warnings

		}

		Add(filePath:="")                                                                         	{	; Create a new blank workbook.
			; using filePath for template
			this.xl.Workbooks.Add()
		}

		BgColor(XLRange, BgColor="19")                                              	{
		  this.xlsheet.Range(XLRange).Interior.ColorIndex := BgColor
		}

		ColorIndex(XLRange)                                                                 	{
			return this.xlsheet.Range(XLRange).Interior.ColorIndex
		}

		Font(XLRange, Options="B", Font="")                                          	{

			/*
				Excel_Font(_ID [,_start, _end, Options, Font])

				  Parameters:
					_ID - The specified Excel Object.
					_start - Default is A1. The starting cell.
					_end - Default is "" (Blank). If blank, it defaults to _start. Otherwise specify the ending cell.
					_Options - Default is B. See below for more options.
				Options:
				  On or On:   | (Default) If specified, B, I, U, S, Sub, Sup, or Strike Will be applied to the cell(s).
				  Off or Off: | If specified, B, I, U, S, Sub, Sup, or Strike Will be removed from the cell(s).
				  Sub         | Subscript.
				  Sup         | Superscript.
				  Strike      | Strikethrough.
				  B           | Bold.
				  I           | Italic.
				  U#          | Underline. If no number is given, it uses a standard solid underline. View chart below for more.
				  S#          | Size. Specify a font size number. Example: S12.
				  J#          | Justify or Align the Cells. if no value is given, Defaults to Left. View chart below for more.
				  C#          | Text color. Give a valid Excel ColorIndex value.
				  BG#         | Background color. Give a valid Excel ColorIndex value. View chart below for more.
				Background options:
				  bg# | Any valid Excel Color Index number
				  bgA | xlBackgroundAutomatic   | -4105 | Excel controls the background.
				  bgO | xlBackgroundOpaque      | 3     | Opaque background.
				  bgT | xlBackgroundTransparent | 2     | Transparent background.
				Underline styles:
				  u  | xlUnderlineStyleSingle           | 2     | Single underlining.
				  u2 | xlUnderlineStyleDoubleAccounting | 5     | Two thin underlines placed close together.
				  u3 | xlUnderlineStyleDouble           | -4119 | Double thick underline.
				  u4 | xlUnderlineStyleNone             | -4142 | No underlining.
				Align Options:
				  j   | xlHAlignLeft                  | -4131 | Align the text to the left.
				  jl  | xlHAlignLeft                  | -4131 | Align the text to the left.
				  jR  | xlHAlignRight                 | -4152 | Align the text to the right.
				  jC  | xlHAlignCenterAcrossSelection | 7     | Align the text to the center of the cell.
				  jJ  | xlHAlignJustify               | -4130 | Align the text to the justified.
				  jD  | xlHAlignDistributed           | -4117 | Distribute the text evenly thoughout the cell.
				  jG  | xlHAlignGeneral               | 1     | Standard alignment.
				  jF  | xlHAlignFill                  | 5     | Repeat the text in the cell until filled.
				  jA  | xlHAlignCenter                | -4108 | Align the text to the center of the cell.
				  jvT | xlVAlignTop                   | -4160 | Center the text on the top of the cell.
				  jvC | xlVAlignCenter                | -4108 | Center the text in the center of the cell.
				  jvB | xlVAlignBottom                | -4107 | Center the text on the bottom of the cell.
				  jvD | xlVAlignDistributed           | -4117 | Center the text evenly (vertically) throughout the cell.
				  jvJ | xlVAlignJustify               | -4130 | Center the text evenly (vertically) throughout the cell.
				  Samples:
					; Make the text Bold and underlined
					Excel_Font(xls, "A1", "B2", "B U")
					; Remove Italics from the range.
					Excel_Font(xls, "A1", "B2", "Off: I")
					; Add Italics and remove any Superscript from the range.
					Excel_Font(xls, "A1", "B2", "Off: Sub On: I")
			*/

		  static _Table:=Object(	 "u", 2, "u2", 5, "u3", -4119, "u4", -4142            	; Underline
					                    	,"bgA", -4105, "bgO", 3, "bgT", 2                     	; Backgrounds
					                    	,"j", -4131, "jl", -4131, "jR", -4152, "jA", -4108   	; Horizontal Alignment
					                    	,"jJ", -4130, "jD", -4117, "jG", 1, "jF", 5, "jC", 7  	; Horizontal Alignment
					                    	,"jvT", -4160, "jvC", -4108, "jvB", -4017             	; Vertical Alignment
					                    	, "jvD", -4117, "jvJ", -4130)                            	; Vertical Alignment

		  Mode       	:= 1
		  FontObject	:= this.xlsheet.Range(XLRange).Font

		  loop, parse, Options, %A_Space% %A_Tab%
		  {
			; rename A_LoopField as ALF just to reduce the length of the code a bit.
			ALF:=A_LoopField

			If (ALF=="")
			  Continue

			If (ALF="On" || ALF="On:")
			  Mode:=1
			If (ALF="Off" || ALF="Off:")
			  Mode:=0

			letter:=SubStr(ALF, 1, 1)
			If (ALF="Sub")
				FontObject.SubScript := Mode
			Else If (ALF="Sup")
				FontObject.Superscript := Mode
			Else If (ALF="Strike")
				FontObject.StrikeThrough := Mode
			Else If (SubStr(ALF, 1, 2)="BG" && (SubStr(ALF, 3, 1)="T"	 || SubStr(ALF, 3, 1)="O" || SubStr(ALF, 3, 1)="A"))
				this.xlsheet.Range(XLRange).Interior.ColorIndex :=_Table[ALF]
			Else If (SubStr(ALF, 1, 2)="BG")
				this.xlsheet.Range(XLRange).Interior.ColorIndex := SubStr(ALF, 3, StrLen(ALF))
			Else If (letter="B")
				FontObject.Bold := Mode
			Else If (letter="I")
				FontObject.Italic := Mode
			Else If (letter="U")
				FontObject.Underline := _Table[ALF]
			Else If (letter="S")
				FontObject.Size:=SubStr(ALF, 2, StrLen(ALF)-1)
			Else If (SubStr(ALF, 1, 2)="jv")
				this.xlsheet.Range(XLRange).VerticalAlignment := _Table[ALF]
			Else If (letter="J")
				this.xlsheet.Range(XLRange).HorizontalAlignment := _Table[ALF]
			Else If (letter="C")
				FontObject.ColorIndex := SubStr(ALF, 2, StrLen(ALF)-1)
			Else
				Continue

		  }

		  If (Font!="")
			FontObject.FontStyle := Font

		}

		SelectRange(XLRange)                                                               		{
			this.xlsheet.Range(XLRange).Select
		}

		SetCell(XLRange, Value)                                                             	{
			this.xlsheet.Range(XLRange) := Value
		}

		CopyToVar(XLRange:="", Delim:="|")                                         	{	; pipe is default

			For Cell in this.xl.Application.ActiveSheet.Range(XLRange)
				Data 	.= Trim(RegExReplace(Cell.Value, "\.[0]+$")) . Delim

				;Data 	.= Trim(Cell.Text) . Delim
			;SciTEOutput(XLRange " = " RTrim(Values, "|") " / "  RTrim(Data, "|"))

		return RTrim(Data, "|")
		}

		CopyToArr(XLRange:="")                                                            	{

			Data := []
			For Cell in this.xl.Application.ActiveSheet.Range(XLRange)
				Data.Push(Trim(RegExReplace(Cell.Value, "\.[0]+$")))

		return Data
		}

		CopyToObject(XLRange:="", BlankValues:=0)                             	{

			Data:=[]
			For Cell in this.xl.Application.ActiveSheet.Range(XLRange)
				if (Cell.Text || Blank)
					Data[Cell.Address(0, 0)]:=Cell.Text

		return Data
		}

		LastColumnRow(Col)                                                                  	{
		return this.xlSheet.Cells(this.xlSheet.Rows.Count, this.StringToNumber(Col)).End(-4162).Row
		}

		UsedRows()                                                                                	{
		;Return this.xlsheet.UsedRange.Rows.Count
		If !(lastrow := this.xlSheet.UsedRange.Rows.Count)
			lastRow := this.xl.Application.Range("A" this.xl.Application.Rows.Count).End(-4162).Row
		Return lastRow
		}

		FindAndReplace(search, replace)                                                	{
			this.xl.Selection.Find.Execute( search, 0, 0, 0, 0, 0, 1, 1, 0, replace, 2)
		}


	; Workbook
		Activate()                                                                                     	{
			this.xlWorkbook.Activate
		}

		OpenWorkbook(filePath)                                                            	{
			this.ActiveWorkbook	:= this.xl.Workbooks.Open(filePath)
			this.ActiveFullName	:= this.xl.ActiveWorkbook.FullName
		}

		CloseActiveWorkbook(SaveChanges:=false)                                  	{
			this.xl.ActiveWorkbook.Close(SaveChanges ? 1 : 0)
		}

		CloseWorkbook(SaveChanges:=false)                                        	{
			this.xl.Workbooks.Close(SaveChanges ? 1 : 0)
		}

		goto(cell)                                                                                   	{ 	; scroll a row into view
			this.xlApplication.goto(this.xlApplication.Range(cell), 1)
		}

		ListWorkbooks()                                                                         	{
			wbObj:=[]
			i:=1
			for name, obj in this.GetActiveObjects()
				if (ComobjType(obj, "Name") = "_Workbook"){
					splitpath, name, oFN
					wbObj.Push(oFN)
				}
		return wbObj
		}

		PrintOut(printer)                                                                        	{

			this.xl.ActivePrinter := printer

			PageSetup := this.xl.Selection.PageSetup
			PageSetup.Orientation 									:= wdOrientPortrait	    	:=	0
			PageSetup.FirstPageTray 								:= wdPrinterDefaultBin  	:=	0
			PageSetup.OtherPagesTray 							:= wdPrinterDefaultBin  	:=	0
			PageSetup.SectionStart       							:= wdSectionContinuous 	:=	0
			PageSetup.OddAndEvenPagesHeaderFooter 	:= False
			PageSetup.DifferentFirstPageHeaderFooter 	:= False
			PageSetup.VerticalAlignment 							:= wdAlignVerticalTop		:=	0
			PageSetup.SuppressEndnotes 						:= False
			PageSetup.MirrorMargins 								:= False
			PageSetup.TwoPagesOnOne 			     			:= False
			PageSetup.BookFoldPrinting 							:= False
			PageSetup.BookFoldRevPrinting 						:= False
			PageSetup.BookFoldPrintingSheets 				:= 1
			Pagesetup.GutterPos 										:= wdGutterPosLeft			:= 	0

			; expression.PrintOut(Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName
			;	    					, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight)

			Background			    	:= 0
			Append				     		:= 0
			Range							:= ""
			OutputFileName			:= ""
			From							:= 0
			To								:= 0
			Item					     		:= ""
			Copies				     		:= 1
			Pages							:= 0
			PageType				    	:= wdPrintAllPages := 0
			PrintToFile				    	:= 0
			Collate					    	:= 1
			FileName				    	:= "doku"
			ActivePrinterMacGX   	:= ""
			ManualDuplexPrint		:= 1
			PrintZoomColumn 		:= 0
			PrintZoomRow				:= 0
			PrintZoomPaperWidth	:= 0
			PrintZoomPaperHeight	:= 0

			this.xl.PrintOut(Background,,,,,,,Copies)

			/*
			;~ this.xl.PrintOut(Background
										;~ , Append
										;~ , Range
										;~ , OutputFileName
										;~ , From
										;~ , To
										;~ , Item
										;~ , Copies
										;~ , Pages
										;~ , PageType
										;~ , PrintToFile
										;~ , Collate
										;~ , FileName
										;~ , ActivePrinterMacGX
										;~ , ManualDuplexPrint
										;~ , PrintZoomColumn
										;~ , PrintZoomRow
										;~ , PrintZoomPaperWidth
										;~ , PrintZoomPaperHeight)

				 */

		}

		Save(New, Path)                                                                        	{

			SplitPath, Path,,, pathext
			if (pathext = "xlsm")
				xtcode := 52
			if (pathext = "xls")
				xtcode := 56
			if (pathext = "xlsx")
				xtcode := 51

			if !New
				res := this.xl.activeworkbook.saved := True
			else
				res := this.xl.activeworkbook.SaveAs(Path, xtcode)

		return res
		}

		ScreenUpdate()                                                                          	{	; toggle update
			this.xl.Application.ScreenUpdating := ! this.xl.Application.ScreenUpdating
		}

		ScrollRow(row)                                                                           	{	 ; scroll a row into view

			this.xlApplication.ActiveWindow.ScrollRow := row

		}

		SelectSheet(SheetName)                                                             	{
		this.xl.Sheets(SheetName).Select
		}

		Reference()                                                                                	{	; will pop up with a message box showing what pointer is referencing
			MsgBox % ComObjType(this.xl, "Name")
		}

		GetActiveObjects(Prefix:="", CaseSensitive:=false)                      	{
			objects := {}
			DllCall("ole32\CoGetMalloc", "uint", 1, "ptr*", malloc) ; malloc: IMalloc
			DllCall("ole32\CreateBindCtx", "uint", 0, "ptr*", bindCtx) ; bindCtx: IBindCtx
			DllCall(NumGet(NumGet(bindCtx+0)+8*A_PtrSize), "ptr", bindCtx, "ptr*", rot) ; rot: IRunningObjectTable
			DllCall(NumGet(NumGet(rot+0)+9*A_PtrSize), "ptr", rot, "ptr*", enum) ; enum: IEnumMoniker
			while (DllCall(NumGet(NumGet(enum+0)+3*A_PtrSize), "ptr", enum, "uint", 1, "ptr*", mon, "ptr", 0) = 0) { ; mon: IMoniker
				DllCall(NumGet(NumGet(mon+0)+20*A_PtrSize), "ptr", mon, "ptr", bindCtx, "ptr", 0, "ptr*", pname) ; GetDisplayName
				name := StrGet(pname, "UTF-16")
				DllCall(NumGet(NumGet(malloc+0)+5*A_PtrSize), "ptr", malloc, "ptr", pname) ; Free
				if InStr(name, Prefix, CaseSensitive) = 1 {
					DllCall(NumGet(NumGet(rot+0)+6*A_PtrSize), "ptr", rot, "ptr", mon, "ptr*", punk) ; GetObject
					; Wrap the pointer as IDispatch if available, otherwise as IUnknown.
					if (pdsp := ComObjQuery(punk, "{00020400-0000-0000-C000-000000000046}"))
						obj := ComObject(9, pdsp, 1), ObjRelease(punk)
					else
						obj := ComObject(13, punk, 1)
					; Store it in the return array by suffix.
					objects[SubStr(name, StrLen(Prefix) + 1)] := obj
				}
				ObjRelease(mon)
			}
			ObjRelease(enum)
			ObjRelease(rot)
			ObjRelease(bindCtx)
			ObjRelease(malloc)
		return objects
		}

	; Application
		Quit()                                                                                         	{
		this.xl.Application.Quit
		}

	; others
		Disconnect()                                                                              	{
			ObjRelease(this.xl) ; := ""
		}

		Speedup(i)                                                                                	{

			If !IsObject(this.xl)
				return

			if  (i) 	{
				this.xl.Displayalerts    	:= 0
				this.xl.EnableEvents    	:= 0
				this.xl.ScreenUpdating	:= 0
				this.xl.Calculation      	:= -4135
			}
			else	{
				this.xl.Displayalerts   	:= 1
				this.xl.EnableEvents  	:= 1
				this.xl.ScreenUpdating 	:= 1
				this.xl.Calculation     	:= -4105
			}

		}

		WinActivate()                                                                                 {

			WinActivate, % this.xlName " ahk_class XLMAIN"

		}

		GetSheetPos()                                                                            	{

			xlHwnd	:= WinExist(this.xlSheetName  " ahk_class XLMAIN")
			ControlGet, xlDeskHwnd, Hwnd,, XLDESK1, % "ahk_id " xlHwnd

		return GetWindowSpot(xlDeskHwnd)
		}

		GetScreenPos()                                                                          	{

			; https://www.autohotkey.com/boards/viewtopic.php?t=37053
				this.ScrollRow(1)
				this.SelectRange("A1")

				Zoom	:= Floor(this.xl.ActiveWindow.Zoom)
				Factor	:= (Zoom / 100) * (A_ScreenDPI / 72)
				WinX 	:= this.xl.ActiveWindow.PointsToScreenPixelsX(0)
				WinY 	:= this.xl.ActiveWindow.PointsToScreenPixelsY(0)

			return {	"Factor"     	: Factor
					,	"Zoom"        	: Zoom
					, 	"WinX"         	: Floor(WinX)
					, 	"WinY"      	: Floor(WinY)
					,	"UsableW" 	: Floor(this.xl.ActiveWindow.UsableWidth)
					,	"UsableH" 	: Floor(this.xl.ActiveWindow.UsableHeight)
					, 	"SelX"           	: WinX + Floor(this.xl.Selection.Left   	* Factor)
					, 	"SelY"          	: WinY + Floor(this.xl.Selection.Top   	* Factor)
					, 	"SelW"         	: Floor(this.xl.Selection.Width         	* Factor)
					, 	"SelH"          	: Floor(this.xl.Selection.Height        	* Factor)}
		}

		GetZoom()                                                                                	{
			return this.xlApplication.ActiveWindow.Zoom
		}

		ColToChar(Index)                                                                       	{	; Converting Columns to Numeric for Excel
		return Index<=26?(Chr(64+index)):Index>26?Chr((index-1)/26+64) . Chr(mod((index - 1),26)+65):""
		}

		StringToNumber(Column)                                                          	{
			; borrowed from automator.com
			StringUpper, Column, Column
			Index := 0
			Loop, Parse, Column  ;loop for each character
			{ascii := asc(A_LoopField)
				if (ascii >= 65 && ascii <= 90)
				index := index * 26 + ascii - 65 + 1    ;Base = 26 (26 letters)
			else { return
			} }
		return, index+0 ;Adding zero here is needed to ensure you're returning an Integer, not a String
		}

}

;}

;--- Interskriptkommunikation 	;{
MessageWorker(InComing) {                                                                                    	;-- verarbeitet die eingegangen Nachrichten

		global recv

		recv := {	"msg"		: (StrSplit(InComing, "|").1)
					, 	"opt"     	: (StrSplit(InComing, "|").2)
					, 	"fromID"	: (StrSplit(InComing, "|").3)}


	; Kommunikation vom aufrufenden Skript - sendet z.B. die PatNR
		if InStr(recv.msg, "NewCase") {
			PatID := recv.opt
			fn_PatSuche := Func("Patientsuchen").Bind(PatID)
			SetTimer, % fn_PatSuche, -100
		}

return
}


;}

;--- Debug Gui                      	;{
XLBarGui(show:=1, opt:="")              	{

	global IMPF, hIMPF, lBEdit, lBExit, lBYPos, lBInfo2, XLBar

	If !IsObject(opt)
		opt:={"col": ["FFF2CC","8EA9DB","F8CBAD"], "W":680,	"H":18}

	xlHwnd := WinExist("PatListe-COVID-Impfungen")
	xlPos:= GetWindowSpot(xlHwnd)

	Gui, IMPF: +ToolWindow +AlwaysOnTop +Resize HWNDhIMPF
	Gui, IMPF: Color, % opt.col.2, % opt.col.3
	Gui, IMPF: Margin, 0, 0

	XLBar := new LoaderBar("IMPF", 3, 3, opt.W, opt.H, "COVID-19-Excel-Helfer", 1, opt.col.1)
	wW :=XLBar.Width + 2*XLBar.X

	Gui, IMPF: Color, % opt.col.2, % "c000000"
	Gui, IMPF: Font, % "cWhite"
	Gui, IMPF: Add, Button	, % "x" XLBar.X " y+" XLBar.Y " vlBExit gIMPFGuiClose", % "Beenden/Abbrechen"

	Gui, IMPF: Font, % "cDarkBlue"
	Gui, IMPF: Add, Text  	, % "x+20 yp+5"  " vlBInfo2 BackgroundTrans ", % "                                                                                           "

	Gui, IMPF: Font, % "cYellow"
	Gui, IMPF: Add, Edit  	, % "x" XLBar.X " y+10 w" wW-2*XLBar.X " r50 vlBEdit hwndlbHEdit"
	SendMessage, 0x00D3, 3, 5,, % "ahk_id " lBHEdit ; SetMargin

	WinSet, ExStyle, 0x0, % "ahk_id " lbHEdit
	cp := GetWindowSpot(lBHEdit)
	lBYPos := cp.Y
	wH :=XLBar.Height + cp.H + 2*cp.BW + 3*XLBar.Y

	Gui, IMPF: Show, % "x" xlPos.X+xlPos.W-wW-38 " y" xlPos.Y+xlPos.H-wH-90  " w" wW " h" wH " NA", % "Corona-Impfplaner"  ; " h" wH
	XLBar.hWin := hIMPF

return

IMPFGuiClose:
IMPFGuiEscape: 	;{
IMPFGuiReload:
	Gui, IMPF: Destroy
return ;}

IMPFGuiSize:    	;{

	Critical, Off	; erst Critical Off soll Critical On dann schneller machen oder zuverlässiger machen (hab vergessen woher ich das habe)
	if A_EventInfo = 1
		return
	Critical
	QSW := A_GuiWidth, QSH:= A_GuiHeight
	GuiControl, IMPF: MoveDraw	, lBEdit, % "w" QSW-2*XLBar.X " h" QSH-lBYPos-1*XLBar.Y+30
	GuiControl, IMPF: MoveDraw	, % XLBar.hLoaderBarBG   	, % "w" QSW-2*XLBar.X
	GuiControl, IMPF: MoveDraw	, % XLBar.hLoaderBarFG     	, % "w" QSW-2*XLBar.X
	GuiControl, IMPF: MoveDraw	, % XLBar.hLoaderNumber 	, % "w" QSW-2*XLBar.X
	GuiControl, IMPF: MoveDraw	, % XLBar.hLoaderDesc      	, % "w" QSW-2*XLBar.X
	GuiControl, IMPF: MoveDraw	, % XLBar.hLoaderBarTitle   	, % "w" QSW-2*XLBar.X

	WinSet, Redraw,, % "ahk_id " hIMPF

return ;}
}

class LoaderBar                                 	{

	__New(GuiID:="Default",x:=0,y:=0,w:=280,h:=28,Title:="",ShowDesc:=0,FontColorDesc:="2B2B2B",FontColor:="EFEFEF",BG:="2B2B2B|2F2F2F|323232",FG:="66A3E2|4B79AF|385D87") {

		SetWinDelay  	, 0
		SetBatchLines	, -1

		_GuiID	:= StrLen(A_Gui) ? A_Gui : 1

		if ( (GuiID="Default") || !StrLen(GuiID) || GuiID==0 )
			GuiID :=_GuiID

		Gui, % GuiID ":Default"

		this.GuiID     	:= GuiID
		this.BG         	:= StrSplit(BG,"|")
		this.BG.W     	:= w
		this.BG.H      	:= h
		this.Width     	:=w
		this.Height    	:=h
		this.FG          	:= StrSplit(FG,"|")
		this.FG.W     	:= this.BG.W - 2
		this.FG.H      	:= (fg_h:=(this.BG.H - 2))
		this.Percen     	:= 0
		this.X            	:= x
		this.Y            	:= y
		fg_x                	:= this.X + 1
		fg_y                	:= this.Y + 1
		this.FontColor 	:= FontColor
		this.ShowDesc 	:= ShowDesc

		;DescBGColor:="4D4D4D"
		DescBGColor:="Black"
		this.DescBGColor := DescBGColor

		this.FontColorDesc := FontColorDesc

		Gui, Font,s10
		Gui, Add, Text, % "x" x " y" y " w" w " h" h " BackgroundTrans 0xE hwndhLoaderBarTitle", % Title
		this.hLoaderBarTitle := hLoaderBarTitle

		Gui, Font,s8
		Gui, Add, Text, % "x" x " y+1 w" w " h" h " 0xE hwndhLoaderBarBG"
		this.ApplyGradient(this.hLoaderBarBG	:= hLoaderBarBG,this.BG.1, this.BG.2, this.BG.3,1)

		Gui, Add, Text, % "x" fg_x " y" fg_y " w0 h" fg_h " 0xE hwndhLoaderBarFG"
		this.ApplyGradient(this.hLoaderBarFG   	:= hLoaderBarFG,this.FG.1, this.FG.2, this.FG.3,1)

		Gui, Add, Text, % "x" x " y" y " w" w " h" h " 0x200 border center BackgroundTrans hwndhLoaderNumber c" FontColor, % "[ 0 % ]"
			this.hLoaderNumber := hLoaderNumber

		if (this.ShowDesc) {
			Gui, Add, Text, % "xp y+2 w" w " h16 0x200 Center border BackgroundTrans hwndhLoaderDesc c" FontColorDesc, Loading...
			this.hLoaderDesc := hLoaderDesc
			this.Height:=h+18
		}

		Gui, Font

		Gui, % _GuiID ":Default"
	}

	Set(p,w:="Loading...") {

		_GuiID	:= StrLen(A_Gui) ? A_Gui : 1
		GuiID 	:= this.GuiID

		Gui, % GuiID ": Default"
		GuiControlGet, LoaderBarBG, Pos, % this.hLoaderBarBG

		this.BG.W 	:= LoaderBarBGW
		this.FG.W  	:= LoaderBarBGW - 2
		this.Percent	:=(p>=100) ? p := 100 : p

		PercentNum	:= Round(this.Percent,0)
		PercentBar	:= Floor((this.Percent/100)*(this.FG.W))

		hLoaderBarTitle		:= this.hLoaderBarTitle
		hLoaderBarFG  	:= this.hLoaderBarFG
		hLoaderNumber 	:= this.hLoaderNumber

		GuiControl, Move	,% hLoaderBarFG  	, % "w" PercentBar
		GuiControl,       	,% hLoaderNumber 	, % "[" PercentNum "%]"

		if (this.ShowDesc) {
			hLoaderDesc := this.hLoaderDesc
			GuiControl,,% hLoaderDesc, %w%
		}

		Gui, % _GuiID ":Default"
	}

	ApplyGradient( Hwnd, LT := "101010", MB := "0000AA", RB := "00FF00", Vertical := 1 ) {
		Static STM_SETIMAGE := 0x172
		ControlGetPos,,, W, H,, % "ahk_id " Hwnd
		PixelData := Vertical ? LT "|" LT "|" LT "|" MB "|" MB "|" MB "|" RB "|" RB "|" RB : LT "|" MB "|" RB "|" LT "|" MB "|" RB "|" LT "|" MB "|" RB
		hBitmap := this.CreateDIB( PixelData, 3, 3, W, H, True )
		oBitmap := DllCall( "SendMessage", "Ptr",Hwnd, "UInt",STM_SETIMAGE, "Ptr",0, "Ptr",hBitmap )
		Return hBitmap, DllCall( "DeleteObject", "Ptr",oBitmap )
	}

	CreateDIB( PixelData, W, H, ResizeW := 0, ResizeH := 0, Gradient := 1  ) {
		; http://ahkscript.org/boards/viewtopic.php?t=3203                  SKAN, CD: 01-Apr-2014 MD: 05-May-2014
		Static LR_Flag1 := 0x2008 ; LR_CREATEDIBSECTION := 0x2000 | LR_COPYDELETEORG := 8
			,  LR_Flag2 := 0x200C ; LR_CREATEDIBSECTION := 0x2000 | LR_COPYDELETEORG := 8 | LR_COPYRETURNORG := 4
			,  LR_Flag3 := 0x0008 ; LR_COPYDELETEORG := 8
		WB := Ceil( ( W * 3 ) / 2 ) * 2,  VarSetCapacity( BMBITS, WB * H + 1, 0 ),  P := &BMBITS
		Loop, Parse, PixelData, |
		P := Numput( "0x" A_LoopField, P+0, 0, "UInt" ) - ( W & 1 and Mod( A_Index * 3, W * 3 ) = 0 ? 0 : 1 )
		hBM := DllCall( "CreateBitmap", "Int",W, "Int",H, "UInt",1, "UInt",24, "Ptr",0, "Ptr" )
		hBM := DllCall( "CopyImage", "Ptr",hBM, "UInt",0, "Int",0, "Int",0, "UInt",LR_Flag1, "Ptr" )
		DllCall( "SetBitmapBits", "Ptr",hBM, "UInt",WB * H, "Ptr",&BMBITS )
		If not ( Gradient + 0 )
			hBM := DllCall( "CopyImage", "Ptr",hBM, "UInt",0, "Int",0, "Int",0, "UInt",LR_Flag3, "Ptr" )
		Return DllCall( "CopyImage", "Ptr",hBM, "Int",0, "Int",ResizeW, "Int",ResizeH, "Int",LR_Flag2, "UPtr" )
	}
}

Edit_Append(hEdit, Txt)                      	{ ; Modified version by SKAN
	Local        ; Original by TheGood on 09-Apr-2010 @ autohotkey.com/board/topic/52441-/?p=328342
	L :=	DllCall("SendMessage", "Ptr",hEdit, "UInt",0x0E, "Ptr",0 , "Ptr",0)     	; WM_GETTEXTLENGTH
	    	DllCall("SendMessage", "Ptr",hEdit, "UInt",0xB1, "Ptr",L , "Ptr",L)       	; EM_SETSEL
	    	DllCall("SendMessage", "Ptr",hEdit, "UInt",0xC2, "Ptr",0, "Str",Txt )   	; EM_REPLACESEL
	ControlSend,, {Enter}, % "ahk_id " hEdit
}

;}

;--- Hotkeys                           	;{
^!e::
	XL.Disconnect()
	XL := ""
	Reload
return

#IfWinActive, ahk_class XLMAIN
^s::
	res1 := XL.Save(false, adm.xlsmPath)
	res2 := XL.Save(true, adm.xlfileDir "\" adm.xlfilenameNoExt "-1." adm.xlfileExt)
	msg := "Exceldatei gespeichert. (" res1 ")`nBackup erstellt (" res2 ")"
	call := TextRender(msg, "time:8000 color:DarkGreen radius:10 margin:(5px 5px)",	"color:White s18 font:(Futura Md Bt)")
return
#IfWinActive

#IfWinActive, scanne Tabelle... ahk_class AutoHotkeyGui
ESC::
	XL.Speedup(false)
	XL.Disconnect()
ExitApp
#IfWinActive
;}

;--- TextRender Bibliothek         	;{

TextRender(text:="", background_style:="", text_style:="") {
   ; Script:    TextRender.ahk
   ; License:   MIT License
   ; Author:    Edison Hua (iseahound)
   ; Date:      2021-05-22
   ; Version:   v1.06
   return (new TextRender).Render(text, background_style, text_style)
}

; TextRender() - Display custom text on screen.
class TextRender {

   static windows := {}

   __New(title := "", WindowStyle := "", WindowExStyle := "", hwndParent := 0) {

      this.gdiplusStartup()

      ; Set a DPI awareness context for CreateWindow().
      dpi := DllCall("SetThreadDpiAwarenessContext", "ptr", -4, "ptr")

      ; Create and show the window.
      this.hwnd := this.CreateWindow(title, WindowStyle, WindowExStyle, hwndParent)
      DllCall("ShowWindow", "ptr", this.hwnd, "int", 4) ; SW_SHOWNOACTIVATE

      ; Restore old DPI awareness context.
      DllCall("SetThreadDpiAwarenessContext", "ptr", dpi, "ptr")

      ; Store a reference to this object accessed by the window handle.
      ; When processing window messages the hwnd can be used to retrieve "this".
      TextRender.windows[this.hwnd] := this
      ObjRelease(&this) ; Allow __Delete() to be called. RefCount - 1.

      ; Fail UpdateMemory() check to access LoadMemory().
      this.BitmapWidth := this.BitmapHeight := 0

      ; Bypass FreeMemory() check before LoadMemory().
      this.gfx := this.obm := this.hbm := this.hdc := ""

      ; Saves repeated calls of Draw().
      this.layers := {}

      ; Initalize default events.
      this.events := {}
	  this.UserCallback := Func("PatientEintragen")
	  this.OnEvent("LeftMouseDown", this.UserCallback)
      this.OnEvent("MiddleMouseDown", this.EventShowCoordinates)
      this.OnEvent("RightMouseDown", this.EventCopyText)

      ; Prevents an unnecessary call of Flush().
      this.drawing := true

      return this
   }

   __Delete() {
      ; FreeMemory() is called by DestroyWindow().
      this.DestroyWindow()
      this.gdiplusShutdown()

      ; Re-add the reference to avoid calling __Delete() twice.
      ObjAddRef(&this)
      ; An unmanaged reference to "this" should be deleted manually.
      TextRender.windows[this.hwnd] := ""
   }

   Render(terms*) {
      ; Check the terms to avoid drawing a default square.
      if (terms.1 != "" || terms.2 != "" || terms.3 != "") {
         this.Draw(terms*)
      }

      ; Allow Render() to commit only when previous calls to Draw() have occurred.
      if (this.layers.length() > 0) {
         ; Define the smaller of canvas and bitmap coordinates.
         this.WindowLeft   := (this.BitmapLeft   > this.x)  ? this.BitmapLeft   : this.x
         this.WindowTop    := (this.BitmapTop    > this.y)  ? this.BitmapTop    : this.y
         this.WindowRight  := (this.BitmapRight  < this.x2) ? this.BitmapRight  : this.x2
         this.WindowBottom := (this.BitmapBottom < this.y2) ? this.BitmapBottom : this.y2
         this.WindowWidth  := this.WindowRight - this.WindowLeft
         this.WindowHeight := this.WindowBottom - this.WindowTop

         ; Reminder: Only the visible screen area will be rendered. Clipping will occur.
         this.UpdateLayeredWindow(this.WindowLeft, this.WindowTop, this.WindowWidth, this.WindowHeight)

         ; Ensure that Flush() will be called at the start of a new drawing.
         ; This approach keeps this.layers and the underlying graphics intact,
         ; so that calls to Save() and Screenshot() will not encounter a blank canvas.
         this.drawing := false

         ; Create a timer that eventually clears the canvas.
         if (this.t > 0) {
            ; Create a reference to the object held by a timer.
            blank := ObjBindMethod(this, "blank", this.status) ; Calls Blank()
            SetTimer % blank, % -this.t ; Calls __Delete.
         }
      }

      return this
   }

   RenderOnScreen(terms*) {
      ; Check the terms to avoid drawing a default square.
      if (terms.1 != "" || terms.2 != "" || terms.3 != "") {
         this.Draw(terms*)
      }

      ; Allow Render() to commit when previous Draw() has happened.
      if (this.layers.length() > 0) {
         ; Use the default rendering when the canvas coordinates fall within the bitmap area.
         if this.InBounds()
            return this.Render(terms*)

         ; Render objects that reside off screen.
         ; Create a new bitmap using the width and height of the canvas object.
         hdc := DllCall("CreateCompatibleDC", "ptr", 0, "ptr")
         VarSetCapacity(bi, 40, 0)              ; sizeof(bi) = 40
            NumPut(       40, bi,  0,   "uint") ; Size
            NumPut(   this.w, bi,  4,   "uint") ; Width
            NumPut(  -this.h, bi,  8,    "int") ; Height - Negative so (0, 0) is top-left.
            NumPut(        1, bi, 12, "ushort") ; Planes
            NumPut(       32, bi, 14, "ushort") ; BitCount / BitsPerPixel
         hbm := DllCall("CreateDIBSection", "ptr", hdc, "ptr", &bi, "uint", 0, "ptr*", pBits:=0, "ptr", 0, "uint", 0, "ptr")
         obm := DllCall("SelectObject", "ptr", hdc, "ptr", hbm, "ptr")
         gfx := DllCall("gdiplus\GdipCreateFromHDC", "ptr", hdc , "ptr*", gfx:=0, "int") ? false : gfx

         ; Set the origin to this.x and this.y
         DllCall("gdiplus\GdipTranslateWorldTransform", "ptr", gfx, "float", -this.x, "float", -this.y, "int", 0)

         ; Redraw on the canvas.
         for i, layer in this.layers
            this.DrawOnGraphics(gfx, layer[1], layer[2], layer[3], this.BitmapWidth, this.BitmapHeight)

         ; Show the objects on screen.
         ; This suffers from a windows limitation in that windows will appear in places that do not match the intended coordinates.
         ; Therefore this is not the default rendering approach as style commands are not respected.
         DllCall("UpdateLayeredWindow"
                  ,    "ptr", this.hwnd                ; hWnd
                  ,    "ptr", 0                        ; hdcDst
                  ,"uint64*", this.x | this.y << 32    ; *pptDst
                  ,"uint64*", this.w | this.h << 32    ; *psize
                  ,    "ptr", hdc                      ; hdcSrc
                  ,"uint64*", 0                        ; *pptSrc
                  ,   "uint", 0                        ; crKey
                  ,  "uint*", 0xFF << 16 | 0x01 << 24  ; *pblend
                  ,   "uint", 2)                       ; dwFlags

         ; Adjust location
         DllCall("SetWindowPos", "ptr", this.hwnd, "ptr", 0, "int", this.x, "int", this.y, "int", 0, "int", 0
            , "uint", 0x400 | 0x10 | 0x4 | 0x1) ; SWP_NOSENDCHANGING | SWP_NOACTIVATE | SWP_NOZORDER | SWP_NOSIZE

         ; Cleanup
         DllCall("gdiplus\GdipDeleteGraphics", "ptr", gfx)
         DllCall("SelectObject", "ptr", hdc, "ptr", obm)
         DllCall("DeleteObject", "ptr", hbm)
         DllCall("DeleteDC",     "ptr", hdc)

         ; Set Coordinates
         WinGetPos x, y, w, h, % "ahk_id " this.hwnd
         this.WindowLeft := x, this.WindowTop := y, this.WindowWidth := w, this.WindowHeight := h
         this.WindowRight  := this.WindowLeft + this.WindowWidth
         this.WindowBottom := this.WindowTop + this.WindowHeight
      }

      ; Create a timer that eventually clears the canvas.
      if (this.t > 0) {
         ; Create a reference to the object held by a timer.
         blank := ObjBindMethod(this, "blank", this.status) ; Calls Blank()
         SetTimer % blank, % -this.t ; Calls __Delete.
      }

      ; Ensure that Flush() will be called at the start of a new drawing.
      ; This approach keeps this.layers and the underlying graphics intact,
      ; so that calls to Save() and Screenshot() will not encounter a blank canvas.
      this.drawing := false
      return this
   }

   Fade(fade_in := 250, fade_out := 250, status := "") {
      if (fade_in > 0) {
         ; Render: Off-Screen areas are not rendered. Clip objects that reside off screen.
         this.WindowLeft   := (this.BitmapLeft   > this.x)  ? this.BitmapLeft   : this.x
         this.WindowTop    := (this.BitmapTop    > this.y)  ? this.BitmapTop    : this.y
         this.WindowRight  := (this.BitmapRight  < this.x2) ? this.BitmapRight  : this.x2
         this.WindowBottom := (this.BitmapBottom < this.y2) ? this.BitmapBottom : this.y2
         this.WindowWidth  := this.WindowRight - this.WindowLeft
         this.WindowHeight := this.WindowBottom - this.WindowTop

         duration := 0
         current := -1
         ;count := 0


         DllCall("QueryPerformanceFrequency", "int64*", frequency:=0)
         DllCall("QueryPerformanceCounter", "int64*", start:=0)
         while (duration < fade_in) {
            alpha := Ceil(duration/fade_in * 255)
            if (alpha != current) {
               ;if (count != alpha)
               ;   FileAppend % count ", " alpha "`n", log.txt
               ;count++
               this.UpdateLayeredWindow(this.WindowLeft, this.WindowTop, this.WindowWidth, this.WindowHeight, alpha)
               current := alpha
            }
            DllCall("QueryPerformanceCounter", "int64*", now:=0)
            duration := (now - start)/frequency * 1000
         }

         if (alpha != 255)
            this.UpdateLayeredWindow(this.WindowLeft, this.WindowTop, this.WindowWidth, this.WindowHeight)

         ; Create a timer that eventually clears the canvas.
         if (this.t > 0) {
            ; Create a reference to the object held by a timer.
            fade := ObjBindMethod(this, "fade", 0, fade_out, this.status) ; Calls Fade() with no fade_in.
            SetTimer % fade, % -this.t ; Calls __Delete.
         }

         ; Ensure that Flush() will be called at the start of a new drawing.
         ; This approach keeps this.layers and the underlying graphics intact,
         ; so that calls to Save() and Screenshot() will not encounter a blank canvas.
         this.drawing := false
         return this
      }

      ; Check to see if the state of the canvas has changed before clearing and updating.
      if (fade_out > 0 && this.status = status) {
         duration := 0
         current := -1
         ;count := 0
         DllCall("QueryPerformanceFrequency", "int64*", frequency:=0)
         DllCall("QueryPerformanceCounter", "int64*", start:=0)
         while (duration < fade_out) {
            alpha := 255 - Ceil(duration/fade_out * 255)
            if (alpha != current) {
               ;if (count != alpha)
               ;   FileAppend % count ", " alpha "`n", log.txt
               ;count++
               this.UpdateLayeredWindow(this.WindowLeft, this.WindowTop, this.WindowWidth, this.WindowHeight, alpha)
               current := alpha
            }
            DllCall("QueryPerformanceCounter", "int64*", now:=0)
            duration := (now - start)/frequency * 1000
         }
         this.UpdateLayeredWindow(this.WindowLeft, this.WindowTop, this.WindowWidth, this.WindowHeight, 0)
         return this
      }
   }

   Blank(status) {
      ; Check to see if the state of the canvas has changed before clearing and updating.
      if (this.status = status) {
         this.UpdateLayeredWindow(this.WindowLeft, this.WindowTop, this.WindowWidth, this.WindowHeight, 0)
      }
   }

   Draw(data := "", styles*) {

      ; If the drawing flag is false then a render to screen operation has occurred.
      if (this.drawing = false)
         this.Flush() ; Clear the internal canvas.

      this.UpdateMemory()

      if (styles[1] = "" && styles[2] = "")
         styles := this.styles
      this.data := data
      this.styles := styles
      this.layers.push([data, styles*])

      ; Drawing
      obj := this.DrawOnGraphics(this.gfx, data, styles[1], styles[2], A_ScreenWidth, A_ScreenHeight)

      ; Create a unique signature for each call to Draw().
      this.CanvasChanged()

      ; Set canvas coordinates.
      this.t  := (this.t  == "") ? obj.t  : (this.t  > obj.t)  ? this.t  : obj.t
      this.x  := (this.x  == "") ? obj.x  : (this.x  < obj.x)  ? this.x  : obj.x
      this.y  := (this.y  == "") ? obj.y  : (this.y  < obj.y)  ? this.y  : obj.y
      this.x2 := (this.x2 == "") ? obj.x2 : (this.x2 > obj.x2) ? this.x2 : obj.x2
      this.y2 := (this.y2 == "") ? obj.y2 : (this.y2 > obj.y2) ? this.y2 : obj.y2
      this.w  := this.x2 - this.x
      this.h  := this.y2 - this.y
      this.chars := obj.chars
      this.words := obj.words
      this.lines := obj.lines

      return this
   }

   Flush() {
      DllCall("gdiplus\GdipSetClipRect", "ptr", this.gfx, "float", this.x, "float", this.y, "float", this.w, "float", this.h, "int", 0)
      DllCall("gdiplus\GdipGraphicsClear", "ptr", this.gfx, "uint", 0x00FFFFFF)
      DllCall("gdiplus\GdipResetClip", "ptr", this.gfx)
      this.CanvasChanged()

      this.t := this.x := this.y := this.x2 := this.y2 := this.w := this.h := ""
      this.layers := {}
      this.drawing := true
      return this
   }

   Clear() {
      this.Flush()
      this.UpdateLayeredWindow(this.WindowLeft, this.WindowTop, this.WindowWidth, this.WindowHeight, 0)
      return this
   }

   Sleep(milliseconds := 0) {
      this.Clear()
      if (milliseconds)
         Sleep % milliseconds
      return this
   }

   Counter() { ; Returns a number in units of milliseconds.
      static freq := DllCall("QueryPerformanceFrequency", "int64*", freq:=0, "int") ? freq*1000 : false
      return DllCall("QueryPerformanceCounter", "int64*", counter:=0, "int") ? counter/freq : false
   }

   CanvasChanged() {
      Random rand, -2147483648, 2147483647
      this.status := rand
      if callback := this.events["CanvasChange"]
         return %callback%(this) ; Callbacks have a reference to "this".
   }

   DrawOnGraphics(gfx, text := "", style1 := "", style2 := "", CanvasWidth := "", CanvasHeight := "") {
      ; Get Graphics Width and Height.
      CanvasWidth := (CanvasWidth != "") ? CanvasWidth : NumGet(gfx + 20 + A_PtrSize, "uint")
      CanvasHeight := (CanvasHeight != "") ? CanvasHeight : NumGet(gfx + 24 + A_PtrSize, "uint")

      ; Remove excess whitespace for proper RegEx detection.
      style1 := !IsObject(style1) ? RegExReplace(style1, "\s+", " ") : style1
      style2 := !IsObject(style2) ? RegExReplace(style2, "\s+", " ") : style2

      ; RegEx help? https://regex101.com/r/xLzZzO/2
      static q1 := "(?i)^.*?\b(?<!:|:\s)\b"
      static q2 := "(?!(?>\([^()]*\)|[^()]*)*\))(:\s*)?\(?(?<value>(?<=\()([\\\/\s:#%_a-z\-\.\d]+|\([\\\/\s:#%_a-z\-\.\d]*\))*(?=\))|[#%_a-z\-\.\d]+).*$"

      ; Extract styles to variables.
      if IsObject(style1) {
         _t  := (style1.time != "")     ? style1.time     : style1.t
         _a  := (style1.anchor != "")   ? style1.anchor   : style1.a
         _x  := (style1.left != "")     ? style1.left     : style1.x
         _y  := (style1.top != "")      ? style1.top      : style1.y
         _w  := (style1.width != "")    ? style1.width    : style1.w
         _h  := (style1.height != "")   ? style1.height   : style1.h
         _r  := (style1.radius != "")   ? style1.radius   : style1.r
         _c  := (style1.color != "")    ? style1.color    : style1.c
         _m  := (style1.margin != "")   ? style1.margin   : style1.m
         _q  := (style1.quality != "")  ? style1.quality  : (style1.q) ? style1.q : style1.SmoothingMode
      } else {
         _t  := ((___ := RegExReplace(style1, q1    "(t(ime)?)"          q2, "${value}")) != style1) ? ___ : ""
         _a  := ((___ := RegExReplace(style1, q1    "(a(nchor)?)"        q2, "${value}")) != style1) ? ___ : ""
         _x  := ((___ := RegExReplace(style1, q1    "(x|left)"           q2, "${value}")) != style1) ? ___ : ""
         _y  := ((___ := RegExReplace(style1, q1    "(y|top)"            q2, "${value}")) != style1) ? ___ : ""
         _w  := ((___ := RegExReplace(style1, q1    "(w(idth)?)"         q2, "${value}")) != style1) ? ___ : ""
         _h  := ((___ := RegExReplace(style1, q1    "(h(eight)?)"        q2, "${value}")) != style1) ? ___ : ""
         _r  := ((___ := RegExReplace(style1, q1    "(r(adius)?)"        q2, "${value}")) != style1) ? ___ : ""
         _c  := ((___ := RegExReplace(style1, q1    "(c(olor)?)"         q2, "${value}")) != style1) ? ___ : ""
         _m  := ((___ := RegExReplace(style1, q1    "(m(argin)?)"        q2, "${value}")) != style1) ? ___ : ""
         _q  := ((___ := RegExReplace(style1, q1    "(q(uality)?)"       q2, "${value}")) != style1) ? ___ : ""
      }

      if IsObject(style2) {
         t  := (style2.time != "")        ? style2.time        : style2.t
         a  := (style2.anchor != "")      ? style2.anchor      : style2.a
         x  := (style2.left != "")        ? style2.left        : style2.x
         y  := (style2.top != "")         ? style2.top         : style2.y
         w  := (style2.width != "")       ? style2.width       : style2.w
         h  := (style2.height != "")      ? style2.height      : style2.h
         m  := (style2.margin != "")      ? style2.margin      : style2.m
         f  := (style2.font != "")        ? style2.font        : style2.f
         s  := (style2.size != "")        ? style2.size        : style2.s
         c  := (style2.color != "")       ? style2.color       : style2.c
         b  := (style2.bold != "")        ? style2.bold        : style2.b
         i  := (style2.italic != "")      ? style2.italic      : style2.i
         u  := (style2.underline != "")   ? style2.underline   : style2.u
         j  := (style2.justify != "")     ? style2.justify     : style2.j
         v  := (style2.vertical != "")    ? style2.vertical    : style2.v
         n  := (style2.noWrap != "")      ? style2.noWrap      : style2.n
         z  := (style2.condensed != "")   ? style2.condensed   : style2.z
         d  := (style2.dropShadow != "")  ? style2.dropShadow  : style2.d
         o  := (style2.outline != "")     ? style2.outline     : style2.o
         q  := (style2.quality != "")     ? style2.quality     : (style2.q) ? style2.q : style2.TextRenderingHint
      } else {
         t  := ((___ := RegExReplace(style2, q1    "(t(ime)?)"          q2, "${value}")) != style2) ? ___ : ""
         a  := ((___ := RegExReplace(style2, q1    "(a(nchor)?)"        q2, "${value}")) != style2) ? ___ : ""
         x  := ((___ := RegExReplace(style2, q1    "(x|left)"           q2, "${value}")) != style2) ? ___ : ""
         y  := ((___ := RegExReplace(style2, q1    "(y|top)"            q2, "${value}")) != style2) ? ___ : ""
         w  := ((___ := RegExReplace(style2, q1    "(w(idth)?)"         q2, "${value}")) != style2) ? ___ : ""
         h  := ((___ := RegExReplace(style2, q1    "(h(eight)?)"        q2, "${value}")) != style2) ? ___ : ""
         m  := ((___ := RegExReplace(style2, q1    "(m(argin)?)"        q2, "${value}")) != style2) ? ___ : ""
         f  := ((___ := RegExReplace(style2, q1    "(f(ont)?)"          q2, "${value}")) != style2) ? ___ : ""
         s  := ((___ := RegExReplace(style2, q1    "(s(ize)?)"          q2, "${value}")) != style2) ? ___ : ""
         c  := ((___ := RegExReplace(style2, q1    "(c(olor)?)"         q2, "${value}")) != style2) ? ___ : ""
         b  := ((___ := RegExReplace(style2, q1    "(b(old)?)"          q2, "${value}")) != style2) ? ___ : ""
         i  := ((___ := RegExReplace(style2, q1    "(i(talic)?)"        q2, "${value}")) != style2) ? ___ : ""
         u  := ((___ := RegExReplace(style2, q1    "(u(nderline)?)"     q2, "${value}")) != style2) ? ___ : ""
         j  := ((___ := RegExReplace(style2, q1    "(j(ustify)?)"       q2, "${value}")) != style2) ? ___ : ""
         v  := ((___ := RegExReplace(style2, q1    "(v(ertical)?)"      q2, "${value}")) != style2) ? ___ : ""
         n  := ((___ := RegExReplace(style2, q1    "(n(oWrap)?)"        q2, "${value}")) != style2) ? ___ : ""
         z  := ((___ := RegExReplace(style2, q1    "(z|condensed)"      q2, "${value}")) != style2) ? ___ : ""
         d  := ((___ := RegExReplace(style2, q1    "(d(ropShadow)?)"    q2, "${value}")) != style2) ? ___ : ""
         o  := ((___ := RegExReplace(style2, q1    "(o(utline)?)"       q2, "${value}")) != style2) ? ___ : ""
         q  := ((___ := RegExReplace(style2, q1    "(q(uality)?)"       q2, "${value}")) != style2) ? ___ : ""
      }

      ; Define color.
      _c := this.parse.color(_c, 0xDD212121) ; Default color for background is transparent gray.
      SourceCopy := (c ~= "i)(delete|eraser?|overwrite|sourceCopy)") ? 1 : 0 ; Eraser brush for text.
      if (!c) ; Default text color changes between white and black.
         c := (this.parse.grayscale(_c) < 128) ? 0xFFFFFFFF : 0xFF000000
      c  := (SourceCopy) ? 0x00000000 : this.parse.color( c)

      ; Default SmoothingMode is 5 for outlines and rounded corners. To disable use 0. See Draw 1, 2, 3.
      _q := (_q >= 0 && _q <= 5) ? _q : 5 ; SmoothingModeAntiAlias8x8

      ; Default TextRenderingHint is Cleartype on a opaque background and Anti-Alias on a transparent background.
      if (q < 0 || q > 5)
         q := (_c & 0xFF000000 = 0xFF000000) ? 5 : 4 ; TextRenderingHintClearTypeGridFit = 5, TextRenderingHintAntialias = 4

      ; Save original Graphics settings.
      DllCall("gdiplus\GdipSaveGraphics", "ptr", gfx, "ptr*", pState:=0)

      ; Use pixels as the defualt unit when rendering.
      DllCall("gdiplus\GdipSetPageUnit", "ptr", gfx, "int", 2) ; A unit is 1 pixel.

      ; Set Graphics settings.
      DllCall("gdiplus\GdipSetPixelOffsetMode",    "ptr", gfx, "int", 4) ; PixelOffsetModeHalf
      ;DllCall("gdiplus\GdipSetCompositingMode",    "ptr", gfx, "int", 1) ; CompositingModeSourceCopy
      DllCall("gdiplus\GdipSetCompositingQuality", "ptr", gfx, "int", 4) ; CompositingQualityGammaCorrected
      DllCall("gdiplus\GdipSetSmoothingMode",      "ptr", gfx, "int", _q)
      DllCall("gdiplus\GdipSetInterpolationMode",  "ptr", gfx, "int", 7) ; HighQualityBicubic
      DllCall("gdiplus\GdipSetTextRenderingHint",  "ptr", gfx, "int", q)

      ; These are the type checkers.
      static valid := "(?i)^\s*(\-?(?:(?:\d+(?:\.\d*)?)|(?:\.\d+)))\s*(%|pt|px|vh|vmin|vw)?\s*$"
      static valid_positive := "(?i)^\s*((?:(?:\d+(?:\.\d*)?)|(?:\.\d+)))\s*(%|pt|px|vh|vmin|vw)?\s*$"

      ; Define viewport width and height. This is the visible canvas area.
      vw := 0.01 * CanvasWidth         ; 1% of viewport width.
      vh := 0.01 * CanvasHeight        ; 1% of viewport height.
      vmin := (vw < vh) ? vw : vh      ; 1vw or 1vh, whichever is smaller.
      vr := CanvasWidth / CanvasHeight ; Aspect ratio of the viewport.

      ; Get background width and height.
      _w := (_w ~= valid_positive) ? RegExReplace(_w, "\s", "") : ""
      _w := (_w ~= "i)(pt|px)$") ? SubStr(_w, 1, -2) : _w
      _w := (_w ~= "i)(%|vw)$") ? RegExReplace(_w, "i)(%|vw)$", "") * vw : _w
      _w := (_w ~= "i)vh$") ? RegExReplace(_w, "i)vh$", "") * vh : _w
      _w := (_w ~= "i)vmin$") ? RegExReplace(_w, "i)vmin$", "") * vmin : _w

      _h := (_h ~= valid_positive) ? RegExReplace(_h, "\s", "") : ""
      _h := (_h ~= "i)(pt|px)$") ? SubStr(_h, 1, -2) : _h
      _h := (_h ~= "i)vw$") ? RegExReplace(_h, "i)vw$", "") * vw : _h
      _h := (_h ~= "i)(%|vh)$") ? RegExReplace(_h, "i)(%|vh)$", "") * vh : _h
      _h := (_h ~= "i)vmin$") ? RegExReplace(_h, "i)vmin$", "") * vmin : _h

      ; Get Font size.
      s  := (s ~= valid_positive) ? RegExReplace(s, "\s", "") : "2.23vh"          ; Default font size is 2.23vh.
      s  := (s ~= "i)(pt|px)$") ? SubStr(s, 1, -2) : s                            ; Strip spaces, px, and pt.
      s  := (s ~= "i)vh$") ? RegExReplace(s, "i)vh$", "") * vh : s                ; Relative to viewport height.
      s  := (s ~= "i)vw$") ? RegExReplace(s, "i)vw$", "") * vw : s                ; Relative to viewport width.
      s  := (s ~= "i)(%|vmin)$") ? RegExReplace(s, "i)(%|vmin)$", "") * vmin : s  ; Relative to viewport minimum.

      ; Get Bold, Italic, Underline, NoWrap, and Justification of text.
      style := (b) ? 1 : 0         ; bold
      style += (i) ? 2 : 0         ; italic
      style += (u) ? 4 : 0         ; underline
      ; style += (strikeout) ? 8 : 0 ; strikeout, not implemented.
      n  := (n) ? 0x4000 | 0x1000 : 0x4000 ; Defaults to text wrapping.

      ; Define text justification. Default text justification to center.
      j  := (j ~= "i)(near|left)") ? 0 : (j ~= "i)cent(er|re)") ? 1 : (j ~= "i)(far|right)") ? 2 : j
      j  := (j ~= "^[0-2]$") ? j : 1

      ; Define vertical alignment. Default vertical alignment to top.
      v  := (v ~= "i)(near|top)") ? 0 : (v ~= "i)cent(er|re)") ? 1 : (v ~= "i)(far|bottom)") ? 2 : j
      v  := (v ~= "^[0-2]$") ? v : 0

      ; Later when text x and w are finalized and it is found that x + width exceeds the screen,
      ; then the _redrawBecauseOfCondensedFont flag is set to true.
      static _redrawBecauseOfCondensedFont := false
      if (_redrawBecauseOfCondensedFont == true)
         f:=z, z:=0, _redrawBecauseOfCondensedFont := false

      ; Specifies whether to load an external font file, or to use an font already installed on the system.
      if (f ~= "(ttf|otf)$") {
         ; Temporarily load a font from file. This does not install the font.
         DllCall("gdiplus\GdipNewPrivateFontCollection", "ptr*", hCollection:=0)
         DllCall("gdiplus\GdipPrivateAddFontFile", "ptr", hCollection, "wstr", f)

         ; A collection of fonts can hold more than just 1 font. Since only 1 font will be needed, a single pointer suffices.
         DllCall("gdiplus\GdipGetFontCollectionFamilyList", "ptr", hCollection, "int", 1, "ptr*", pFontFamily:=0, "int*", found:=0)

         ; Normally, pFontFamily is an array of pointers. For a single pointer, no special requirements are needed.
         VarSetCapacity(FontName, 256)
         DllCall("gdiplus\GdipGetFamilyName", "ptr", pFontFamily, "str", FontName, "ushort", 1033) ; en-US

         ; Create a font family. For ANSI compatibility, use str as the output type and StrGet to pass wide chars.
         DllCall("gdiplus\GdipCreateFontFamilyFromName", "wstr", StrGet(&FontName, "UTF-16"), "ptr", hCollection, "ptr*", hFamily:=0)

         ; Delete the private font collection. It is strange a pointer reference is used.
         DllCall("gdiplus\GdipDeletePrivateFontCollection", "ptr*", hCollection)
      } else {
         ; Create Font. Defaults to Segoe UI or Tahoma on older systems.
         if DllCall("gdiplus\GdipCreateFontFamilyFromName", "wstr",          f, "uint", 0, "ptr*", hFamily:=0)
         if DllCall("gdiplus\GdipCreateFontFamilyFromName", "wstr", "Segoe UI", "uint", 0, "ptr*", hFamily:=0)
            DllCall("gdiplus\GdipCreateFontFamilyFromName", "wstr",   "Tahoma", "uint", 0, "ptr*", hFamily:=0)
      }

      DllCall("gdiplus\GdipCreateFont", "ptr", hFamily, "float", s, "int", style, "int", 0, "ptr*", hFont:=0)
      DllCall("gdiplus\GdipCreateStringFormat", "int", n, "int", 0, "ptr*", hFormat:=0)
      DllCall("gdiplus\GdipSetStringFormatAlign", "ptr", hFormat, "int", j) ; Left = 0, Center = 1, Right = 2
      DllCall("gdiplus\GdipSetStringFormatLineAlign", "ptr", hFormat, "int", v) ; Top = 0, Center = 1, Bottom = 2

      ; Use the declared width and height of the text box if given.
      VarSetCapacity(RectF, 16, 0)                         ; sizeof(RectF) = 16
         (_w != "") ? NumPut(_w, RectF,  8,  "float") : "" ; Width
         (_h != "") ? NumPut(_h, RectF, 12,  "float") : "" ; Height

      ; Otherwise simulate the drawing...
      DllCall("gdiplus\GdipMeasureString"
               ,    "ptr", gfx
               ,   "wstr", text
               ,    "int", -1                 ; string length is null terminated.
               ,    "ptr", hFont
               ,    "ptr", &RectF             ; (in) layout RectF that bounds the string.
               ,    "ptr", hFormat
               ,    "ptr", &RectF             ; (out) simulated RectF that bounds the string.
               ,  "uint*", chars:=0
               ,  "uint*", lines:=0)

      ; Extract the simulated width and height of the text string's bounding box...
      width := NumGet(RectF, 8, "float")
      height := NumGet(RectF, 12, "float")
      minimum := (width < height) ? width : height
      aspect := (height != 0) ? width / height : 0

      ; And use those values for the background width and height.
      (_w == "") ? _w := width : ""
      (_h == "") ? _h := height : ""


      ; Get background anchor. This is where the origin of the image is located.
      _a := RegExReplace(_a, "\s", "")
      _a := (_a ~= "i)top" && _a ~= "i)left") ? 1 : (_a ~= "i)top" && _a ~= "i)cent(er|re)") ? 2
         : (_a ~= "i)top" && _a ~= "i)right") ? 3 : (_a ~= "i)cent(er|re)" && _a ~= "i)left") ? 4
         : (_a ~= "i)cent(er|re)" && _a ~= "i)right") ? 6 : (_a ~= "i)bottom" && _a ~= "i)left") ? 7
         : (_a ~= "i)bottom" && _a ~= "i)cent(er|re)") ? 8 : (_a ~= "i)bottom" && _a ~= "i)right") ? 9
         : (_a ~= "i)top") ? 2 : (_a ~= "i)left") ? 4 : (_a ~= "i)right") ? 6 : (_a ~= "i)bottom") ? 8
         : (_a ~= "i)cent(er|re)") ? 5 : (_a ~= "^[1-9]$") ? _a : 1 ; Default anchor is top-left.

      ; The anchor can be implied from _x and _y (left, center, right, top, bottom).
      _a := (_x ~= "i)left") ? 1+(((_a-1)//3)*3) : (_x ~= "i)cent(er|re)") ? 2+(((_a-1)//3)*3) : (_x ~= "i)right") ? 3+(((_a-1)//3)*3) : _a
      _a := (_y ~= "i)top") ? 1+(mod(_a-1,3)) : (_y ~= "i)cent(er|re)") ? 4+(mod(_a-1,3)) : (_y ~= "i)bottom") ? 7+(mod(_a-1,3)) : _a

      ; Convert English words to numbers. Don't mess with these values any further.
      _x := (_x ~= "i)left") ? 0 : (_x ~= "i)cent(er|re)") ? 0.5*CanvasWidth : (_x ~= "i)right") ? CanvasWidth : _x
      _y := (_y ~= "i)top") ? 0 : (_y ~= "i)cent(er|re)") ? 0.5*CanvasHeight : (_y ~= "i)bottom") ? CanvasHeight : _y

      ; Get _x and _y.
      _x := (_x ~= valid) ? RegExReplace(_x, "\s", "") : ""
      _x := (_x ~= "i)(pt|px)$") ? SubStr(_x, 1, -2) : _x
      _x := (_x ~= "i)(%|vw)$") ? RegExReplace(_x, "i)(%|vw)$", "") * vw : _x
      _x := (_x ~= "i)vh$") ? RegExReplace(_x, "i)vh$", "") * vh : _x
      _x := (_x ~= "i)vmin$") ? RegExReplace(_x, "i)vmin$", "") * vmin : _x

      _y := (_y ~= valid) ? RegExReplace(_y, "\s", "") : ""
      _y := (_y ~= "i)(pt|px)$") ? SubStr(_y, 1, -2) : _y
      _y := (_y ~= "i)vw$") ? RegExReplace(_y, "i)vw$", "") * vw : _y
      _y := (_y ~= "i)(%|vh)$") ? RegExReplace(_y, "i)(%|vh)$", "") * vh : _y
      _y := (_y ~= "i)vmin$") ? RegExReplace(_y, "i)vmin$", "") * vmin : _y

      ; Default x and y to center of the canvas. Default anchor to horizontal center and vertical center.
      if (_x == "")
         _x := 0.5*CanvasWidth, _a := 2+(((_a-1)//3)*3)
      if (_y == "")
         _y := 0.5*CanvasHeight, _a := 4+(mod(_a-1,3))

      ; Now let's modify the _x and _y values with the _anchor, so that the image has a new point of origin.
      ; We need our calculated _width and _height for this!
      _x -= (mod(_a-1,3) == 0) ? 0 : (mod(_a-1,3) == 1) ? _w/2 : (mod(_a-1,3) == 2) ? _w : 0
      _y -= (((_a-1)//3) == 0) ? 0 : (((_a-1)//3) == 1) ? _h/2 : (((_a-1)//3) == 2) ? _h : 0

      ; Prevent half-pixel rendering and keep image sharp.
      _w := Round(_x + _w) - Round(_x) ; Use real x2 coordinate to determine width.
      _h := Round(_y + _h) - Round(_y) ; Use real y2 coordinate to determine height.
      _x := Round(_x)                  ; NOTE: simple Floor(w) or Round(w) will NOT work.
      _y := Round(_y)                  ; The float values need to be added up and then rounded!

      ; Get the text width and text height.
      w  := ( w ~= valid_positive) ? RegExReplace( w, "\s", "") : width ; Default is simulated text width.
      w  := ( w ~= "i)(pt|px)$") ? SubStr( w, 1, -2) :  w
      w  := ( w ~= "i)vw$") ? RegExReplace( w, "i)vw$", "") * vw :  w
      w  := ( w ~= "i)vh$") ? RegExReplace( w, "i)vh$", "") * vh :  w
      w  := ( w ~= "i)vmin$") ? RegExReplace( w, "i)vmin$", "") * vmin :  w
      w  := ( w ~= "%$") ? RegExReplace( w, "%$", "") * 0.01 * _w :  w

      h  := ( h ~= valid_positive) ? RegExReplace( h, "\s", "") : height ; Default is simulated text height.
      h  := ( h ~= "i)(pt|px)$") ? SubStr( h, 1, -2) :  h
      h  := ( h ~= "i)vw$") ? RegExReplace( h, "i)vw$", "") * vw :  h
      h  := ( h ~= "i)vh$") ? RegExReplace( h, "i)vh$", "") * vh :  h
      h  := ( h ~= "i)vmin$") ? RegExReplace( h, "i)vmin$", "") * vmin :  h
      h  := ( h ~= "%$") ? RegExReplace( h, "%$", "") * 0.01 * _h :  h

      ; Manually justify because text width and height may be set above.
      ; If text justification is set but x is not, align the justified text relative to the center
      ; or right of the backgound, after taking into account the text width.
      if (x == "")
         x  := (j = 1) ? _x + (_w/2) - (w/2) : (j = 2) ? _x + _w - w : x
      if (y == "")
         y  := (v = 1) ? _y + (_h/2) - (h/2) : (v = 2) ? _y + _h - h : y

      ; Get anchor.
      a  := RegExReplace( a, "\s", "")
      a  := (a ~= "i)top" && a ~= "i)left") ? 1 : (a ~= "i)top" && a ~= "i)cent(er|re)") ? 2
         : (a ~= "i)top" && a ~= "i)right") ? 3 : (a ~= "i)cent(er|re)" && a ~= "i)left") ? 4
         : (a ~= "i)cent(er|re)" && a ~= "i)right") ? 6 : (a ~= "i)bottom" && a ~= "i)left") ? 7
         : (a ~= "i)bottom" && a ~= "i)cent(er|re)") ? 8 : (a ~= "i)bottom" && a ~= "i)right") ? 9
         : (a ~= "i)top") ? 2 : (a ~= "i)left") ? 4 : (a ~= "i)right") ? 6 : (a ~= "i)bottom") ? 8
         : (a ~= "i)cent(er|re)") ? 5 : (a ~= "^[1-9]$") ? a : 1 ; Default anchor is top-left.

      ; Text x and text y can be specified as locations (left, center, right, top, bottom).
      ; These location words in text x and text y take precedence over the values in the text anchor.
      a  := ( x ~= "i)left") ? 1+((( a-1)//3)*3) : ( x ~= "i)cent(er|re)") ? 2+((( a-1)//3)*3) : ( x ~= "i)right") ? 3+((( a-1)//3)*3) :  a
      a  := ( y ~= "i)top") ? 1+(mod( a-1,3)) : ( y ~= "i)cent(er|re)") ? 4+(mod( a-1,3)) : ( y ~= "i)bottom") ? 7+(mod( a-1,3)) :  a

      ; Convert English words to numbers. Don't mess with these values any further.
      ; Also, these values are relative to the background.
      x  := ( x ~= "i)left") ? _x : (x ~= "i)cent(er|re)") ? _x + 0.5*_w : (x ~= "i)right") ? _x + _w : x
      y  := ( y ~= "i)top") ? _y : (y ~= "i)cent(er|re)") ? _y + 0.5*_h : (y ~= "i)bottom") ? _y + _h : y

      ; Default text x is background x.
      x  := ( x ~= valid) ? RegExReplace( x, "\s", "") : _x
      x  := ( x ~= "i)(pt|px)$") ? SubStr( x, 1, -2) :  x
      x  := ( x ~= "i)vw$") ? RegExReplace( x, "i)vw$", "") * vw :  x
      x  := ( x ~= "i)vh$") ? RegExReplace( x, "i)vh$", "") * vh :  x
      x  := ( x ~= "i)vmin$") ? RegExReplace( x, "i)vmin$", "") * vmin :  x
      x  := ( x ~= "%$") ? RegExReplace( x, "%$", "") * 0.01 * _w :  x

      ; Default text y is background y.
      y  := ( y ~= valid) ? RegExReplace( y, "\s", "") : _y
      y  := ( y ~= "i)(pt|px)$") ? SubStr( y, 1, -2) :  y
      y  := ( y ~= "i)vw$") ? RegExReplace( y, "i)vw$", "") * vw :  y
      y  := ( y ~= "i)vh$") ? RegExReplace( y, "i)vh$", "") * vh :  y
      y  := ( y ~= "i)vmin$") ? RegExReplace( y, "i)vmin$", "") * vmin :  y
      y  := ( y ~= "%$") ? RegExReplace( y, "%$", "") * 0.01 * _h :  y

      ; Modify text x and text y values with the anchor, so that the text has a new point of origin.
      ; The text anchor is relative to the text width and height before margin/padding.
      ; This is NOT relative to the background width and height.
      x  -= (mod(a-1,3) == 0) ? 0 : (mod(a-1,3) == 1) ? w/2 : (mod(a-1,3) == 2) ? w : 0
      y  -= (((a-1)//3) == 0) ? 0 : (((a-1)//3) == 1) ? h/2 : (((a-1)//3) == 2) ? h : 0

      ; Get margin. Default margin is 1vmin.
      m  := this.parse.margin_and_padding( m, vw, vh)
      _m := this.parse.margin_and_padding(_m, vw, vh, (m.void && _w > 0 && _h > 0) ? "1vmin" : "")

      ; Modify _x, _y, _w, _h with margin and padding, increasing the size of the background.
      _w += Round(_m.2) + Round(_m.4) + Round(m.2) + Round(m.4)
      _h += Round(_m.1) + Round(_m.3) + Round(m.1) + Round(m.3)
      _x -= Round(_m.4)
      _y -= Round(_m.1)

      ; If margin/padding are defined in the text parameter, shift the position of the text.
      x  += Round(m.4)
      y  += Round(m.1)

      ; Re-run: Condense Text using a Condensed Font if simulated text width exceeds screen width.
      if (z) {
         if (width + x > CanvasWidth) {
            _redrawBecauseOfCondensedFont := true
            return this.DrawOnGraphics(gfx, text, style1, style2, CanvasWidth, CanvasHeight)
         }
      }

      ; Define the smaller of the backgound width or height.
      _min := (_w > _h) ? _h : _w

      ; Define the maximum roundness of the background bubble.
      _rmax := _min / 2

      ; Define radius of rounded corners. The default radius is 0, or square corners.
      _r := (_r ~= "i)max") ? _rmax : _r
      _r := (_r ~= valid_positive) ? RegExReplace(_r, "\s", "") : 0
      _r := (_r ~= "i)(pt|px)$") ? SubStr(_r, 1, -2) : _r
      _r := (_r ~= "i)vw$") ? RegExReplace(_r, "i)vw$", "") * vw : _r
      _r := (_r ~= "i)vh$") ? RegExReplace(_r, "i)vh$", "") * vh : _r
      _r := (_r ~= "i)vmin$") ? RegExReplace(_r, "i)vmin$", "") * vmin : _r
      _r := (_r ~= "%$") ? RegExReplace(_r, "%$", "") * 0.01 * _min : _r ; percentage of minimum
      _r := (_r > _rmax) ? _rmax : _r ; Exceeding _rmax will create a candy wrapper effect.

      ; Define outline and dropShadow.
      o := this.parse.outline(o, vw, vh, s, c)
      d := this.parse.dropShadow(d, vw, vh, width, height, s)


      ; Draw 1 - Background
      if (_w && _h && (_c & 0xFF000000)) {
         ; Create background solid brush.
         DllCall("gdiplus\GdipCreateSolidFill", "uint", _c, "ptr*", pBrush:=0)

         ; Fill a rectangle with a solid brush. Draw sharp rectangular edges.
         if (_r == 0) {
            DllCall("gdiplus\GdipSetSmoothingMode", "ptr", gfx, "int", 0) ; SmoothingModeNoAntiAlias
            DllCall("gdiplus\GdipFillRectangle", "ptr", gfx, "ptr", pBrush, "float", _x, "float", _y, "float", _w, "float", _h) ; DRAWING!
            DllCall("gdiplus\GdipSetSmoothingMode", "ptr", gfx, "int", _q)
         }

         ; Fill a rounded rectangle with a solid brush.
         else {
            _r2 := (_r * 2) ; Calculate diameter
            DllCall("gdiplus\GdipCreatePath", "uint", 0, "ptr*", pPath:=0)
            DllCall("gdiplus\GdipAddPathArc", "ptr", pPath, "float", _x           , "float", _y           , "float", _r2, "float", _r2, "float", 180, "float", 90)
            DllCall("gdiplus\GdipAddPathArc", "ptr", pPath, "float", _x + _w - _r2, "float", _y           , "float", _r2, "float", _r2, "float", 270, "float", 90)
            DllCall("gdiplus\GdipAddPathArc", "ptr", pPath, "float", _x + _w - _r2, "float", _y + _h - _r2, "float", _r2, "float", _r2, "float",   0, "float", 90)
            DllCall("gdiplus\GdipAddPathArc", "ptr", pPath, "float", _x           , "float", _y + _h - _r2, "float", _r2, "float", _r2, "float",  90, "float", 90)
            DllCall("gdiplus\GdipClosePathFigure", "ptr", pPath) ; Connect existing arc segments into a rounded rectangle.
            DllCall("gdiplus\GdipFillPath", "ptr", gfx, "ptr", pBrush, "ptr", pPath) ; DRAWING!
            DllCall("gdiplus\GdipDeletePath", "ptr", pPath)
         }

         ; Delete background solid brush.
         DllCall("gdiplus\GdipDeleteBrush", "ptr", pBrush)
      }


      ; Draw 2 - DropShadow
      if (!d.void) {
         offset2 := d.3 + d.6 + Ceil(0.5*o.1)

         ; If blur is present, a second canvas must be seperately processed to apply the Gaussian Blur effect.
         if (true) {
            ;DropShadow := Gdip_CreateBitmap(w + 2*offset2, h + 2*offset2)
            ;DropShadow := Gdip_CreateBitmap(A_ScreenWidth, A_ScreenHeight, 0xE200B)
            DllCall("gdiplus\GdipCreateBitmapFromScan0", "int", A_ScreenWidth, "int", A_ScreenHeight
               , "uint", 0, "uint", 0xE200B, "ptr", 0, "ptr*", DropShadow:=0)
            DllCall("gdiplus\GdipGetImageGraphicsContext", "ptr", DropShadow, "ptr*", DropShadowG:=0)
            DllCall("gdiplus\GdipSetSmoothingMode", "ptr", DropShadowG, "int", 0) ; SmoothingModeNoAntiAlias
            DllCall("gdiplus\GdipSetTextRenderingHint", "ptr", DropShadowG, "int", 1) ; TextRenderingHintSingleBitPerPixelGridFit
            DllCall("gdiplus\GdipGraphicsClear", "ptr", gfx, "uint", d.4 & 0xFFFFFF)
            VarSetCapacity(RectF, 16, 0)          ; sizeof(RectF) = 16
               NumPut(d.1+x, RectF,  0,  "float") ; Left
               NumPut(d.2+y, RectF,  4,  "float") ; Top
               NumPut(    w, RectF,  8,  "float") ; Width
               NumPut(    h, RectF, 12,  "float") ; Height

            ;CreateRectF(RC, offset2, offset2, w + 2*offset2, h + 2*offset2)
         } else {
            ;CreateRectF(RC, x + d.1, y + d.2, w, h)
            VarSetCapacity(RectF, 16, 0)          ; sizeof(RectF) = 16
               NumPut(d.1+x, RectF,  0,  "float") ; Left
               NumPut(d.2+y, RectF,  4,  "float") ; Top
               NumPut(    w, RectF,  8,  "float") ; Width
               NumPut(    h, RectF, 12,  "float") ; Height
            DropShadowG := gfx
         }

         ; Use Gdip_DrawString if and only if there is a horizontal/vertical offset.
         if (o.void && d.6 == 0)
         {
            ; Use shadow solid brush.
            DllCall("gdiplus\GdipCreateSolidFill", "uint", d.4, "ptr*", pBrush:=0)
            DllCall("gdiplus\GdipDrawString"
                     ,    "ptr", DropShadowG
                     ,   "wstr", text
                     ,    "int", -1
                     ,    "ptr", hFont
                     ,    "ptr", &RectF
                     ,    "ptr", hFormat
                     ,    "ptr", pBrush)
            DllCall("gdiplus\GdipDeleteBrush", "ptr", pBrush)
         }
         else ; Otherwise, use the below code if blur, size, and opacity are set.
         {
            ; Draw the outer edge of the text string.
            DllCall("gdiplus\GdipCreatePath", "int",1, "ptr*",pPath:=0)
            DllCall("gdiplus\GdipAddPathString"
                     ,    "ptr", pPath
                     ,   "wstr", text
                     ,    "int", -1
                     ,    "ptr", hFamily
                     ,    "int", style
                     ,  "float", s
                     ,    "ptr", &RectF
                     ,    "ptr", hFormat)
            DllCall("gdiplus\GdipCreatePen1", "uint", d.4, "float", 2*d.6 + o.1, "int", 2, "ptr*", pPen:=0)
            DllCall("gdiplus\GdipSetPenLineJoin", "ptr", pPen, "uint", 2) ; LineJoinTypeRound
            DllCall("gdiplus\GdipDrawPath", "ptr", DropShadowG, "ptr", pPen, "ptr", pPath)
            DllCall("gdiplus\GdipDeletePen", "ptr", pPen)

            ; Fill in the outline. Turn off antialiasing and alpha blending so the gaps are 100% filled.
            DllCall("gdiplus\GdipCreateSolidFill", "uint", d.4, "ptr*", pBrush:=0)
            DllCall("gdiplus\GdipSetCompositingMode", "ptr", DropShadowG, "int", 1) ; CompositingModeSourceCopy
            DllCall("gdiplus\GdipSetSmoothingMode", "ptr", DropShadowG, "int", 0) ; SmoothingModeNoAntiAlias
            DllCall("gdiplus\GdipFillPath", "ptr", DropShadowG, "ptr", pBrush, "ptr", pPath) ; DRAWING!
            DllCall("gdiplus\GdipDeleteBrush", "ptr", pBrush)
            DllCall("gdiplus\GdipDeletePath", "ptr", pPath)
            DllCall("gdiplus\GdipSetCompositingMode", "ptr", DropShadowG, "int", 0) ; CompositingModeSourceOver
            DllCall("gdiplus\GdipSetSmoothingMode", "ptr", DropShadowG, "int", _q)
         }

         if (true) {
            DllCall("gdiplus\GdipDeleteGraphics", "ptr", DropShadowG)
            this.filter.GaussianBlur(DropShadow, d.3, d.5)
            DllCall("gdiplus\GdipSetInterpolationMode", "ptr", gfx, "int", 5) ; NearestNeighbor
            DllCall("gdiplus\GdipSetSmoothingMode", "ptr", gfx, "int", 0) ; SmoothingModeNoAntiAlias
            ;Gdip_DrawImage(gfx, DropShadow, x + d.1 - offset2, y + d.2 - offset2, w + 2*offset2, h + 2*offset2) ; DRAWING!
            ;Gdip_DrawImage(gfx, DropShadow, 0, 0, A_Screenwidth, A_ScreenHeight) ; DRAWING!
            DllCall("gdiplus\GdipDrawImageRectRectI" ; DRAWING!
                     ,    "ptr", gfx
                     ,    "ptr", DropShadow
                     ,    "int", 0, "int", 0, "int", A_Screenwidth, "int", A_Screenwidth ; destination rectangle
                     ,    "int", 0, "int", 0, "int", A_Screenwidth, "int", A_Screenwidth ; source rectangle
                     ,    "int", 2  ; UnitTypePixel
                     ,    "ptr", 0  ; imageAttributes
                     ,    "ptr", 0  ; callback
                     ,    "ptr", 0) ; callbackData
            DllCall("gdiplus\GdipSetSmoothingMode", "ptr", gfx, "int", _q)
            DllCall("gdiplus\GdipDisposeImage", "ptr", DropShadow)
         }
      }


      ; Draw 3 - Outline
      if (!o.void) {
         ; Convert our text to a path.
         VarSetCapacity(RectF, 16, 0)          ; sizeof(RectF) = 16
            NumPut(    x, RectF,  0,  "float") ; Left
            NumPut(    y, RectF,  4,  "float") ; Top
            NumPut(    w, RectF,  8,  "float") ; Width
            NumPut(    h, RectF, 12,  "float") ; Height
         DllCall("gdiplus\GdipCreatePath", "int", 1, "ptr*", pPath:=0)
         DllCall("gdiplus\GdipAddPathString"
                  ,    "ptr", pPath
                  ,   "wstr", text
                  ,    "int", -1
                  ,    "ptr", hFamily
                  ,    "int", style
                  ,  "float", s
                  ,    "ptr", &RectF
                  ,    "ptr", hFormat)

         ; Create a glow effect around the edges.
         if (o.3) {
            DllCall("gdiplus\GdipSetClipPath", "ptr", gfx, "ptr", pPath, "int", 3) ; Exclude original text region from being drawn on.
            ARGB := Format("0x{:02X}",((o.4 & 0xFF000000) >> 24)/o.3) . Format("{:06X}",(o.4 & 0x00FFFFFF))
            DllCall("gdiplus\GdipCreatePen1", "uint", ARGB, "float", 1, "int", 2, "ptr*", pPenGlow:=0) ; UnitTypePixel = 2
            DllCall("gdiplus\GdipSetPenLineJoin", "ptr",pPenGlow, "uint",2) ; LineJoinTypeRound

            Loop % o.3
            {
               DllCall("gdiplus\GdipSetPenWidth", "ptr", pPenGlow, "float", o.1 + 2*A_Index)
               DllCall("gdiplus\GdipDrawPath", "ptr", gfx, "ptr", pPenGlow, "ptr", pPath) ; DRAWING!
            }
            DllCall("gdiplus\GdipDeletePen", "ptr", pPenGlow)
            DllCall("gdiplus\GdipResetClip", "ptr", gfx)
         }

         ; Draw outline text.
         if (o.1) {
            DllCall("gdiplus\GdipCreatePen1", "uint", o.2, "float", o.1, "int", 2, "ptr*", pPen:=0) ; UnitTypePixel = 2
            DllCall("gdiplus\GdipSetPenLineJoin", "ptr", pPen, "uint", 2) ; LineJoinTypeRound
            DllCall("gdiplus\GdipDrawPath", "ptr", gfx, "ptr", pPen, "ptr", pPath) ; DRAWING!
            DllCall("gdiplus\GdipDeletePen", "ptr", pPen)
         }

         ; Fill outline text.
         DllCall("gdiplus\GdipCreateSolidFill", "uint", c, "ptr*", pBrush:=0)
         DllCall("gdiplus\GdipSetCompositingMode", "ptr", gfx, "int", SourceCopy)
         DllCall("gdiplus\GdipFillPath", "ptr", gfx, "ptr", pBrush, "ptr", pPath) ; DRAWING!
         DllCall("gdiplus\GdipSetCompositingMode", "ptr", gfx, "int", 0) ; CompositingModeSourceOver
         DllCall("gdiplus\GdipDeleteBrush", "ptr", pBrush)
         DllCall("gdiplus\GdipDeletePath", "ptr", pPath)
      }


      ; Draw 4 - Text
      if (text != "" && o.void) {
         DllCall("gdiplus\GdipSetCompositingMode", "ptr", gfx, "int", SourceCopy)

         VarSetCapacity(RectF, 16, 0)          ; sizeof(RectF) = 16
            NumPut(    x, RectF,  0,  "float") ; Left
            NumPut(    y, RectF,  4,  "float") ; Top
            NumPut(    w, RectF,  8,  "float") ; Width
            NumPut(    h, RectF, 12,  "float") ; Height

         DllCall("gdiplus\GdipMeasureString"
                  ,    "ptr", gfx
                  ,   "wstr", text
                  ,    "int", -1                 ; string length.
                  ,    "ptr", hFont
                  ,    "ptr", &RectF             ; (in) layout RectF that bounds the string.
                  ,    "ptr", hFormat
                  ,    "ptr", &RectF             ; (out) simulated RectF that bounds the string.
                  ,  "uint*", chars:=0
                  ,  "uint*", lines:=0)

         DllCall("gdiplus\GdipCreateSolidFill", "uint", c, "ptr*", pBrush:=0)
         DllCall("gdiplus\GdipDrawString"
                  ,    "ptr", gfx
                  ,   "wstr", text
                  ,    "int", -1
                  ,    "ptr", hFont
                  ,    "ptr", &RectF
                  ,    "ptr", hFormat
                  ,    "ptr", pBrush)
         DllCall("gdiplus\GdipDeleteBrush", "ptr", pBrush)

         x := NumGet(RectF,  0, "float")
         y := NumGet(RectF,  4, "float")
         w := NumGet(RectF,  8, "float")
         h := NumGet(RectF, 12, "float")
      }


      ; Cleanup.
      DllCall("gdiplus\GdipDeleteStringFormat", "ptr", hFormat)
      DllCall("gdiplus\GdipDeleteFont", "ptr", hFont)
      DllCall("gdiplus\GdipDeleteFontFamily", "ptr", hFamily)

      ; Restore original Graphics settings.
      DllCall("gdiplus\GdipRestoreGraphics", "ptr", gfx, "ptr", pState)

      ; Calulate the number of words.
      ; First, use the number of chars displayed by GdipMeasureString to truncate "text".
      ; Then count the number of words, as defined by Unicode Code Points, i.e. all languages.
      RegExReplace(SubStr(text, 1, chars), "(*UCP)\b\w+\b", "", words)

      ; Calculate time.
      t  := (_t) ? _t : t
      if (t = "fast") ; To be used when the user has seen the text before; to linger on screen momentarily.
         t := 1250 + 8*chars ; Every character adds 8 milliseconds.
      if (t = "auto") {
         ; The average human reaction time is 250 ms. For when text suddenly appears on screen.
         ; Using 200 words/minute, divide 60,000 ms by 200 words to get 300 ms per word.
         t := 250 + 300*words
      }

      ; Extract the time variable and save it for a later when we Render() everything.
      static times := "(?i)^\s*(\d+)\s*(ms|mil(li(second)?)?|s(ec(ond)?)?|m(in(ute)?)?|h(our)?|d(ay)?)?s?\s*$"
      t  := ( t ~= times) ? RegExReplace( t, "\s", "") : 0 ; Default time is zero.
      t  := ((___ := RegExReplace( t, "i)(\d+)(ms|mil(li(second)?)?)s?$", "$1")) !=  t) ? ___ *        1 : t
      t  := ((___ := RegExReplace( t, "i)(\d+)s(ec(ond)?)?s?$"          , "$1")) !=  t) ? ___ *     1000 : t
      t  := ((___ := RegExReplace( t, "i)(\d+)m(in(ute)?)?s?$"          , "$1")) !=  t) ? ___ *    60000 : t
      t  := ((___ := RegExReplace( t, "i)(\d+)h(our)?s?$"               , "$1")) !=  t) ? ___ *  3600000 : t
      t  := ((___ := RegExReplace( t, "i)(\d+)d(ay)?s?$"                , "$1")) !=  t) ? ___ * 86400000 : t

      ; Define canvas coordinates.
      t_bound :=  t                              ; string/background boundary.
      x_bound := (_c & 0xFF000000) ? _x : x
      y_bound := (_c & 0xFF000000) ? _y : y
      w_bound := (_c & 0xFF000000) ? _w : w
      h_bound := (_c & 0xFF000000) ? _h : h

      o_bound := Ceil(0.5 * o.1 + o.3)                     ; outline boundary.
      x_bound := (x - o_bound < x_bound)        ? x - o_bound        : x_bound
      y_bound := (y - o_bound < y_bound)        ? y - o_bound        : y_bound
      w_bound := (w + 2 * o_bound > w_bound)    ? w + 2 * o_bound    : w_bound
      h_bound := (h + 2 * o_bound > h_bound)    ? h + 2 * o_bound    : h_bound
      ; Tooltip % x_bound ", " y_bound ", " w_bound ", " h_bound
      d_bound := Ceil(0.5 * o.1 + d.3 + d.6)            ; dropShadow boundary.
      x_bound := (x + d.1 - d_bound < x_bound)  ? x + d.1 - d_bound  : x_bound
      y_bound := (y + d.2 - d_bound < y_bound)  ? y + d.2 - d_bound  : y_bound
      w_bound := (w + 2 * d_bound > w_bound)    ? w + 2 * d_bound    : w_bound
      h_bound := (h + 2 * d_bound > h_bound)    ? h + 2 * d_bound    : h_bound

      return {t: t_bound
            , x: x_bound, y: y_bound
            , w: w_bound, h: h_bound
            , x2: x_bound + w_bound, y2: y_bound + h_bound
            , chars: chars
            , words: words
            , lines: lines}
   }

   DrawOnBitmap(pBitmap, text := "", style1 := "", style2 := "") {
      DllCall("gdiplus\GdipGetImageGraphicsContext", "ptr", pBitmap, "ptr*", gfx:=0)
      obj := this.DrawOnGraphics(gfx, text, style1, style2)
      DllCall("gdiplus\GdipDeleteGraphics", "ptr", gfx)
      return obj
   }

   DrawOnHDC(hdc, text := "", style1 := "", style2 := "") {
      DllCall("gdiplus\GdipCreateFromHDC", "ptr", hdc, "ptr*", gfx:=0)
      obj := this.DrawOnGraphics(gfx, text, style1, style2)
      DllCall("gdiplus\GdipDeleteGraphics", "ptr", gfx)
      return obj
   }

   class parse {

      color(c, default := 0xDD424242) {
         static xARGB := "^0x([0-9A-Fa-f]{8})$"
         static xRGB  := "^0x([0-9A-Fa-f]{6})$"
         static ARGB  :=   "^([0-9A-Fa-f]{8})$"
         static RGB   :=   "^([0-9A-Fa-f]{6})$"

         if ObjGetCapacity([c], 1) {
            c  := (c ~= "^#") ? SubStr(c, 2) : c
            c  := ((___ := this.colormap(c)) != "") ? ___ : c
            c  := (c ~= xRGB) ? "0xFF" RegExReplace(c, xRGB, "$1") : (c ~= ARGB) ? "0x" c : (c ~= RGB) ? "0xFF" c : c
            c  := (c ~= xARGB) ? c : default
         }

         return (c != "") ? c : default
      }

      colormap(c) {
         if (c = "random") ; 93% opacity + random RGB.
            return "0xEE" SubStr(ComObjCreate("Scriptlet.TypeLib").GUID, 2, 6)

         if (c = "random2") ; Solid opacity + random RGB.
            return "0xFF" SubStr(ComObjCreate("Scriptlet.TypeLib").GUID, 2, 6)

         if (c = "random3") ; Fully random opacity and RGB.
            return SubStr(ComObjCreate("Scriptlet.TypeLib").GUID, 2, 8)

         static colors1 :=
         ( LTrim Join
         {
            "Clear"                 : "0x00000000",
            "None"                  : "0x00000000",
            "Off"                   : "0x00000000",
            "Transparent"           : "0x00000000",
            "AliceBlue"             : "0xFFF0F8FF",
            "AntiqueWhite"          : "0xFFFAEBD7",
            "Aqua"                  : "0xFF00FFFF",
            "Aquamarine"            : "0xFF7FFFD4",
            "Azure"                 : "0xFFF0FFFF",
            "Beige"                 : "0xFFF5F5DC",
            "Bisque"                : "0xFFFFE4C4",
            "Black"                 : "0xFF000000",
            "BlanchedAlmond"        : "0xFFFFEBCD",
            "Blue"                  : "0xFF0000FF",
            "BlueViolet"            : "0xFF8A2BE2",
            "Brown"                 : "0xFFA52A2A",
            "BurlyWood"             : "0xFFDEB887",
            "CadetBlue"             : "0xFF5F9EA0",
            "Chartreuse"            : "0xFF7FFF00",
            "Chocolate"             : "0xFFD2691E",
            "Coral"                 : "0xFFFF7F50",
            "CornflowerBlue"        : "0xFF6495ED",
            "Cornsilk"              : "0xFFFFF8DC",
            "Crimson"               : "0xFFDC143C",
            "Cyan"                  : "0xFF00FFFF",
            "DarkBlue"              : "0xFF00008B",
            "DarkCyan"              : "0xFF008B8B",
            "DarkGoldenRod"         : "0xFFB8860B",
            "DarkGray"              : "0xFFA9A9A9",
            "DarkGrey"              : "0xFFA9A9A9",
            "DarkGreen"             : "0xFF006400",
            "DarkKhaki"             : "0xFFBDB76B",
            "DarkMagenta"           : "0xFF8B008B",
            "DarkOliveGreen"        : "0xFF556B2F",
            "DarkOrange"            : "0xFFFF8C00",
            "DarkOrchid"            : "0xFF9932CC",
            "DarkRed"               : "0xFF8B0000",
            "DarkSalmon"            : "0xFFE9967A",
            "DarkSeaGreen"          : "0xFF8FBC8F",
            "DarkSlateBlue"         : "0xFF483D8B",
            "DarkSlateGray"         : "0xFF2F4F4F",
            "DarkSlateGrey"         : "0xFF2F4F4F",
            "DarkTurquoise"         : "0xFF00CED1",
            "DarkViolet"            : "0xFF9400D3",
            "DeepPink"              : "0xFFFF1493",
            "DeepSkyBlue"           : "0xFF00BFFF",
            "DimGray"               : "0xFF696969",
            "DimGrey"               : "0xFF696969",
            "DodgerBlue"            : "0xFF1E90FF",
            "FireBrick"             : "0xFFB22222",
            "FloralWhite"           : "0xFFFFFAF0",
            "ForestGreen"           : "0xFF228B22",
            "Fuchsia"               : "0xFFFF00FF",
            "Gainsboro"             : "0xFFDCDCDC",
            "GhostWhite"            : "0xFFF8F8FF",
            "Gold"                  : "0xFFFFD700",
            "GoldenRod"             : "0xFFDAA520",
            "Gray"                  : "0xFF808080",
            "Grey"                  : "0xFF808080",
            "Green"                 : "0xFF008000",
            "GreenYellow"           : "0xFFADFF2F",
            "HoneyDew"              : "0xFFF0FFF0",
            "HotPink"               : "0xFFFF69B4",
            "IndianRed"             : "0xFFCD5C5C",
            "Indigo"                : "0xFF4B0082",
            "Ivory"                 : "0xFFFFFFF0",
            "Khaki"                 : "0xFFF0E68C",
            "Lavender"              : "0xFFE6E6FA",
            "LavenderBlush"         : "0xFFFFF0F5",
            "LawnGreen"             : "0xFF7CFC00",
            "LemonChiffon"          : "0xFFFFFACD",
            "LightBlue"             : "0xFFADD8E6",
            "LightCoral"            : "0xFFF08080",
            "LightCyan"             : "0xFFE0FFFF",
            "LightGoldenRodYellow"  : "0xFFFAFAD2",
            "LightGray"             : "0xFFD3D3D3",
            "LightGrey"             : "0xFFD3D3D3",
            "LightGreen"            : "0xFF90EE90",
            "LightPink"             : "0xFFFFB6C1",
            "LightSalmon"           : "0xFFFFA07A",
            "LightSeaGreen"         : "0xFF20B2AA",
            "LightSkyBlue"          : "0xFF87CEFA",
            "LightSlateGray"        : "0xFF778899",
            "LightSlateGrey"        : "0xFF778899",
            "LightSteelBlue"        : "0xFFB0C4DE",
            "LightYellow"           : "0xFFFFFFE0",
            "Lime"                  : "0xFF00FF00",
            "LimeGreen"             : "0xFF32CD32",
            "Linen"                 : "0xFFFAF0E6"
         }
         )
         static colors2 :=
         ( LTrim Join
         {
            "Magenta"               : "0xFFFF00FF",
            "Maroon"                : "0xFF800000",
            "MediumAquaMarine"      : "0xFF66CDAA",
            "MediumBlue"            : "0xFF0000CD",
            "MediumOrchid"          : "0xFFBA55D3",
            "MediumPurple"          : "0xFF9370DB",
            "MediumSeaGreen"        : "0xFF3CB371",
            "MediumSlateBlue"       : "0xFF7B68EE",
            "MediumSpringGreen"     : "0xFF00FA9A",
            "MediumTurquoise"       : "0xFF48D1CC",
            "MediumVioletRed"       : "0xFFC71585",
            "MidnightBlue"          : "0xFF191970",
            "MintCream"             : "0xFFF5FFFA",
            "MistyRose"             : "0xFFFFE4E1",
            "Moccasin"              : "0xFFFFE4B5",
            "NavajoWhite"           : "0xFFFFDEAD",
            "Navy"                  : "0xFF000080",
            "OldLace"               : "0xFFFDF5E6",
            "Olive"                 : "0xFF808000",
            "OliveDrab"             : "0xFF6B8E23",
            "Orange"                : "0xFFFFA500",
            "OrangeRed"             : "0xFFFF4500",
            "Orchid"                : "0xFFDA70D6",
            "PaleGoldenRod"         : "0xFFEEE8AA",
            "PaleGreen"             : "0xFF98FB98",
            "PaleTurquoise"         : "0xFFAFEEEE",
            "PaleVioletRed"         : "0xFFDB7093",
            "PapayaWhip"            : "0xFFFFEFD5",
            "PeachPuff"             : "0xFFFFDAB9",
            "Peru"                  : "0xFFCD853F",
            "Pink"                  : "0xFFFFC0CB",
            "Plum"                  : "0xFFDDA0DD",
            "PowderBlue"            : "0xFFB0E0E6",
            "Purple"                : "0xFF800080",
            "RebeccaPurple"         : "0xFF663399",
            "Red"                   : "0xFFFF0000",
            "RosyBrown"             : "0xFFBC8F8F",
            "RoyalBlue"             : "0xFF4169E1",
            "SaddleBrown"           : "0xFF8B4513",
            "Salmon"                : "0xFFFA8072",
            "SandyBrown"            : "0xFFF4A460",
            "SeaGreen"              : "0xFF2E8B57",
            "SeaShell"              : "0xFFFFF5EE",
            "Sienna"                : "0xFFA0522D",
            "Silver"                : "0xFFC0C0C0",
            "SkyBlue"               : "0xFF87CEEB",
            "SlateBlue"             : "0xFF6A5ACD",
            "SlateGray"             : "0xFF708090",
            "SlateGrey"             : "0xFF708090",
            "Snow"                  : "0xFFFFFAFA",
            "SpringGreen"           : "0xFF00FF7F",
            "SteelBlue"             : "0xFF4682B4",
            "Tan"                   : "0xFFD2B48C",
            "Teal"                  : "0xFF008080",
            "Thistle"               : "0xFFD8BFD8",
            "Tomato"                : "0xFFFF6347",
            "Turquoise"             : "0xFF40E0D0",
            "Violet"                : "0xFFEE82EE",
            "Wheat"                 : "0xFFF5DEB3",
            "White"                 : "0xFFFFFFFF",
            "WhiteSmoke"            : "0xFFF5F5F5",
            "Yellow"                : "0xFFFFFF00",
            "YellowGreen"           : "0xFF9ACD32"
         }
         )
         return colors1.HasKey(c) ? colors1[c] : colors2[c]
      }

      dropShadow(d, vw, vh, width, height, font_size) {
         static q1 := "(?i)^.*?\b(?<!:|:\s)\b"
         static q2 := "(?!(?>\([^()]*\)|[^()]*)*\))(:\s*)?\(?(?<value>(?<=\()([\\\/\s:#%_a-z\-\.\d]+|\([\\\/\s:#%_a-z\-\.\d]*\))*(?=\))|[#%_a-z\-\.\d]+).*$"
         static valid := "(?i)^\s*(\-?(?:(?:\d+(?:\.\d*)?)|(?:\.\d+)))\s*(%|pt|px|vh|vmin|vw)?\s*$"
         vmin := (vw < vh) ? vw : vh

         if IsObject(d) {
            d.1 := (d.horizontal != "") ? d.horizontal : (d.h != "") ? d.h : d.1
            d.2 := (d.vertical   != "") ? d.vertical   : (d.v != "") ? d.h : d.2
            d.3 := (d.blur       != "") ? d.blur       : (d.b != "") ? d.h : d.3
            d.4 := (d.color      != "") ? d.color      : (d.c != "") ? d.h : d.4
            d.5 := (d.opacity    != "") ? d.opacity    : (d.o != "") ? d.h : d.5
            d.6 := (d.size       != "") ? d.size       : (d.s != "") ? d.h : d.6
         } else if (d != "") {
            _ := RegExReplace(d, ":\s+", ":")
            _ := RegExReplace(_, "\s+", " ")
            _ := StrSplit(_, " ")
            _.1 := ((___ := RegExReplace(d, q1    "(h(orizontal)?)"    q2, "${value}")) != d) ? ___ : _.1
            _.2 := ((___ := RegExReplace(d, q1    "(v(ertical)?)"      q2, "${value}")) != d) ? ___ : _.2
            _.3 := ((___ := RegExReplace(d, q1    "(b(lur)?)"          q2, "${value}")) != d) ? ___ : _.3
            _.4 := ((___ := RegExReplace(d, q1    "(c(olor)?)"         q2, "${value}")) != d) ? ___ : _.4
            _.5 := ((___ := RegExReplace(d, q1    "(o(pacity)?)"       q2, "${value}")) != d) ? ___ : _.5
            _.6 := ((___ := RegExReplace(d, q1    "(s(ize)?)"          q2, "${value}")) != d) ? ___ : _.6
            d := _
         }
         else return {"void":true, 1:0, 2:0, 3:0, 4:0, 5:0, 6:0}

         for key, value in d {
            if (key = 4) ; Don't mess with color data.
               continue
            d[key] := (d[key] ~= valid) ? RegExReplace(d[key], "\s", "") : 0 ; Default for everything is 0.
            d[key] := (d[key] ~= "i)(pt|px)$") ? SubStr(d[key], 1, -2) : d[key]
            d[key] := (d[key] ~= "i)vw$") ? RegExReplace(d[key], "i)vw$", "") * vw : d[key]
            d[key] := (d[key] ~= "i)vh$") ? RegExReplace(d[key], "i)vh$", "") * vh : d[key]
            d[key] := (d[key] ~= "i)vmin$") ? RegExReplace(d[key], "i)vmin$", "") * vmin : d[key]
         }

         d.1 := (d.1 ~= "%$") ? SubStr(d.1, 1, -1) * 0.01 * width : d.1
         d.2 := (d.2 ~= "%$") ? SubStr(d.2, 1, -1) * 0.01 * height : d.2
         d.3 := (d.3 ~= "%$") ? SubStr(d.3, 1, -1) * 0.01 * font_size : d.3
         d.4 := this.color(d.4, 0xFFFF0000) ; Default color is red.
         d.5 := (d.5 ~= "%$") ? SubStr(d.5, 1, -1) / 100 : d.5
         d.5 := (d.5 <= 0 || d.5 > 1) ? 1 : d.5 ; Range Opacity is a float from 0-1.
         d.6 := (d.6 ~= "%$") ? SubStr(d.6, 1, -1) * 0.01 * font_size : d.6
         return d
      }

      grayscale(sRGB) {
         static rY := 0.212655
         static gY := 0.715158
         static bY := 0.072187

         c1 := 255 & ( sRGB >> 16 )
         c2 := 255 & ( sRGB >> 8 )
         c3 := 255 & ( sRGB )

         loop 3 {
            c%A_Index% := c%A_Index% / 255
            c%A_Index% := (c%A_Index% <= 0.04045) ? c%A_Index%/12.92 : ((c%A_Index%+0.055)/(1.055))**2.4
         }

         v := rY*c1 + gY*c2 + bY*c3
         v := (v <= 0.0031308) ? v * 12.92 : 1.055*(v**(1.0/2.4))-0.055
         return Round(v*255)
      }

      margin_and_padding(m, vw, vh, default := "") {
         static q1 := "(?i)^.*?\b(?<!:|:\s)\b"
         static q2 := "(?!(?>\([^()]*\)|[^()]*)*\))(:\s*)?\(?(?<value>(?<=\()([\\\/\s:#%_a-z\-\.\d]+|\([\\\/\s:#%_a-z\-\.\d]*\))*(?=\))|[#%_a-z\-\.\d]+).*$"
         static valid := "(?i)^\s*(\-?(?:(?:\d+(?:\.\d*)?)|(?:\.\d+)))\s*(%|pt|px|vh|vmin|vw)?\s*$"
         vmin := (vw < vh) ? vw : vh

         if IsObject(m) {
            m.1 := (m.top    != "") ? m.top    : (m.t != "") ? m.t : m.1
            m.2 := (m.right  != "") ? m.right  : (m.r != "") ? m.r : m.2
            m.3 := (m.bottom != "") ? m.bottom : (m.b != "") ? m.b : m.3
            m.4 := (m.left   != "") ? m.left   : (m.l != "") ? m.l : m.4
         } else if (m != "") {
            _ := RegExReplace(m, ":\s+", ":")
            _ := RegExReplace(_, "\s+", " ")
            _ := StrSplit(_, " ")
            _.1 := ((___ := RegExReplace(m, q1    "(t(op)?)"           q2, "${value}")) != m) ? ___ : _.1
            _.2 := ((___ := RegExReplace(m, q1    "(r(ight)?)"         q2, "${value}")) != m) ? ___ : _.2
            _.3 := ((___ := RegExReplace(m, q1    "(b(ottom)?)"        q2, "${value}")) != m) ? ___ : _.3
            _.4 := ((___ := RegExReplace(m, q1    "(l(eft)?)"          q2, "${value}")) != m) ? ___ : _.4
            m := _
         } else if (default != "")
            m := {1:default, 2:default, 3:default, 4:default}
         else return {"void":true, 1:0, 2:0, 3:0, 4:0}

         ; Follow CSS guidelines for margin!
         if (m.2 == "" && m.3 == "" && m.4 == "")
            m.4 := m.3 := m.2 := m.1, exception := true
         if (m.3 == "" && m.4 == "")
            m.4 := m.2, m.3 := m.1
         if (m.4 == "")
            m.4 := m.2

         for key, value in m {
            m[key] := (m[key] ~= valid) ? RegExReplace(m[key], "\s", "") : default
            m[key] := (m[key] ~= "i)(pt|px)$") ? SubStr(m[key], 1, -2) : m[key]
            m[key] := (m[key] ~= "i)vw$") ? RegExReplace(m[key], "i)vw$", "") * vw : m[key]
            m[key] := (m[key] ~= "i)vh$") ? RegExReplace(m[key], "i)vh$", "") * vh : m[key]
            m[key] := (m[key] ~= "i)vmin$") ? RegExReplace(m[key], "i)vmin$", "") * vmin : m[key]
         }
         m.1 := (m.1 ~= "%$") ? SubStr(m.1, 1, -1) * vh : m.1
         m.2 := (m.2 ~= "%$") ? SubStr(m.2, 1, -1) * (exception ? vh : vw) : m.2
         m.3 := (m.3 ~= "%$") ? SubStr(m.3, 1, -1) * vh : m.3
         m.4 := (m.4 ~= "%$") ? SubStr(m.4, 1, -1) * (exception ? vh : vw) : m.4
         return m
      }

      outline(o, vw, vh, font_size, font_color) {
         static q1 := "(?i)^.*?\b(?<!:|:\s)\b"
         static q2 := "(?!(?>\([^()]*\)|[^()]*)*\))(:\s*)?\(?(?<value>(?<=\()([\\\/\s:#%_a-z\-\.\d]+|\([\\\/\s:#%_a-z\-\.\d]*\))*(?=\))|[#%_a-z\-\.\d]+).*$"
         static valid_positive := "(?i)^\s*((?:(?:\d+(?:\.\d*)?)|(?:\.\d+)))\s*(%|pt|px|vh|vmin|vw)?\s*$"
         vmin := (vw < vh) ? vw : vh

         if IsObject(o) {
            o.1 := (o.stroke != "") ? o.stroke : (o.s != "") ? o.s : o.1
            o.2 := (o.color  != "") ? o.color  : (o.c != "") ? o.c : o.2
            o.3 := (o.glow   != "") ? o.glow   : (o.g != "") ? o.g : o.3
            o.4 := (o.tint   != "") ? o.tint   : (o.t != "") ? o.t : o.4
         } else if (o != "") {
            _ := RegExReplace(o, ":\s+", ":")
            _ := RegExReplace(_, "\s+", " ")
            _ := StrSplit(_, " ")
            _.1 := ((___ := RegExReplace(o, q1    "(s(troke)?)"        q2, "${value}")) != o) ? ___ : _.1
            _.2 := ((___ := RegExReplace(o, q1    "(c(olor)?)"         q2, "${value}")) != o) ? ___ : _.2
            _.3 := ((___ := RegExReplace(o, q1    "(g(low)?)"          q2, "${value}")) != o) ? ___ : _.3
            _.4 := ((___ := RegExReplace(o, q1    "(t(int)?)"          q2, "${value}")) != o) ? ___ : _.4
            o := _
         }
         else return {"void":true, 1:0, 2:0, 3:0, 4:0}

         for key, value in o {
            if (key = 2) || (key = 4) ; Don't mess with color data.
               continue
            o[key] := (o[key] ~= valid_positive) ? RegExReplace(o[key], "\s", "") : 0 ; Default for everything is 0.
            o[key] := (o[key] ~= "i)(pt|px)$") ? SubStr(o[key], 1, -2) : o[key]
            o[key] := (o[key] ~= "i)vw$") ? RegExReplace(o[key], "i)vw$", "") * vw : o[key]
            o[key] := (o[key] ~= "i)vh$") ? RegExReplace(o[key], "i)vh$", "") * vh : o[key]
            o[key] := (o[key] ~= "i)vmin$") ? RegExReplace(o[key], "i)vmin$", "") * vmin : o[key]
         }

         o.1 := (o.1 ~= "%$") ? SubStr(o.1, 1, -1) * 0.01 * font_size : o.1
         o.2 := this.color(o.2, font_color) ; Default color is the text font color.
         o.3 := (o.3 ~= "%$") ? SubStr(o.3, 1, -1) * 0.01 * font_size : o.3
         o.4 := this.color(o.4, o.2) ; Default color is outline color.
         return o
      }
   }

   class filter {

      GaussianBlur(pBitmap, radius, opacity := 1) {
         static code := (A_PtrSize = 4)
            ? "
            ( LTrim                                    ; 32-bit machine code
            VYnlV1ZTg+xci0Uci30c2UUgx0WsAwAAAI1EAAGJRdiLRRAPr0UYicOJRdSLRRwP
            r/sPr0UYiX2ki30UiUWoi0UQjVf/i30YSA+vRRgDRQgPr9ONPL0SAAAAiUWci0Uc
            iX3Eg2XE8ECJVbCJRcCLRcSJZbToAAAAACnEi0XEiWXk6AAAAAApxItFxIllzOgA
            AAAAKcSLRaiJZcjHRdwAAAAAx0W8AAAAAIlF0ItFvDtFFA+NcAEAAItV3DHAi12c
            i3XQiVXgAdOLfQiLVdw7RRiNDDp9IQ+2FAGLTcyLfciJFIEPtgwDD69VwIkMh4tN
            5IkUgUDr0THSO1UcfBKLXdwDXQzHRbgAAAAAK13Q6yAxwDtFGH0Ni33kD7YcAQEc
            h0Dr7kIDTRjrz/9FuAN1GItF3CtF0AHwiceLRbg7RRx/LDHJO00YfeGLRQiLfcwB
            8A+2BAgrBI+LfeQDBI+ZiQSPjTwz933YiAQPQevWi0UIK0Xci03AAfCJRbiLXRCJ
            /itdHCt13AN14DnZfAgDdQwrdeDrSot1DDHbK3XcAf4DdeA7XRh9KItV4ItFuAHQ
            A1UID7YEGA+2FBop0ItV5AMEmokEmpn3fdiIBB5D69OLRRhBAUXg66OLRRhDAUXg
            O10QfTIxyTtNGH3ti33Ii0XgA0UID7YUCIsEjynQi1XkAwSKiQSKi1XgjTwWmfd9
            2IgED0Hr0ItF1P9FvAFF3AFF0OmE/v//i0Wkx0XcAAAAAMdFvAAAAACJRdCLRbAD
            RQyJRaCLRbw7RRAPjXABAACLTdwxwItdoIt10IlN4AHLi30Mi1XcO0UYjQw6fSEP
            thQBi33MD7YMA4kUh4t9yA+vVcCJDIeLTeSJFIFA69Ex0jtVHHwSi13cA10Ix0W4
            AAAAACtd0OsgMcA7RRh9DYt95A+2HAEBHIdA6+5CA03U68//RbgDddSLRdwrRdAB
            8InHi0W4O0UcfywxyTtNGH3hi0UMi33MAfAPtgQIKwSPi33kAwSPmYkEj408M/d9
            2IgED0Hr1otFDCtF3ItNwAHwiUW4i10Uif4rXRwrddwDdeA52XwIA3UIK3Xg60qL
            dQgx2yt13AH+A3XgO10YfSiLVeCLRbgB0ANVDA+2BBgPthQaKdCLVeQDBJqJBJqZ
            933YiAQeQ+vTi0XUQQFF4Ouji0XUQwFF4DtdFH0yMck7TRh97Yt9yItF4ANFDA+2
            FAiLBI+LfeQp0ItV4AMEj4kEj408Fpn3fdiIBA9B69CLRRj/RbwBRdwBRdDphP7/
            //9NrItltA+Fofz//9no3+l2PzHJMds7XRR9OotFGIt9CA+vwY1EBwMx/zt9EH0c
            D7Yw2cBHVtoMJFrZXeTzDyx15InyiBADRRjr30MDTRDrxd3Y6wLd2I1l9DHAW15f
            XcM=
            )" : "
            ( LTrim                                    ; 64-bit machine code
            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
            )"

         ; Get width and height.
         DllCall("gdiplus\GdipGetImageWidth", "ptr", pBitmap, "uint*", width:=0)
         DllCall("gdiplus\GdipGetImageHeight", "ptr", pBitmap, "uint*", height:=0)

         ; Create a buffer of raw 32-bit ARGB pixel data.
         VarSetCapacity(Rect, 16, 0)            ; sizeof(Rect) = 16
            NumPut(  width, Rect,  8,   "uint") ; Width
            NumPut( height, Rect, 12,   "uint") ; Height
         VarSetCapacity(BitmapData, 16+2*A_PtrSize, 0) ; sizeof(BitmapData) = 24, 32
         DllCall("gdiplus\GdipBitmapLockBits", "ptr", pBitmap, "ptr", &Rect, "uint", 3, "int", 0x26200A, "ptr", &BitmapData)

         ; Get the Scan0 of the pixel data. Create a working buffer of the exact same size.
         stride := NumGet(BitmapData,  8, "int")
         Scan01 := NumGet(BitmapData, 16, "ptr")
         Scan02 := DllCall("GlobalAlloc", "uint", 0x40, "uptr", stride * height, "ptr")

         ; Call machine code function.
         DllCall("crypt32\CryptStringToBinary", "str", code, "uint", 0, "uint", 0x1, "ptr", 0, "uint*", size:=0, "ptr", 0, "ptr", 0)
         p := DllCall("GlobalAlloc", "uint", 0, "uptr", size, "ptr")
         DllCall("VirtualProtect", "ptr", p, "ptr", size, "uint", 0x40, "uint*", op) ; Allow execution from memory.
         DllCall("crypt32\CryptStringToBinary", "str", code, "uint", 0, "uint", 0x1, "ptr", p, "uint*", size, "ptr", 0, "ptr", 0)
         e := DllCall(p, "ptr", Scan01, "ptr", Scan02, "uint", width, "uint", height, "uint", 4, "uint", radius, "float", opacity)
         DllCall("GlobalFree", "ptr", p)

         ; Free resources.
         DllCall("gdiplus\GdipBitmapUnlockBits", "ptr", pBitmap, "ptr", &BitmapData)
         DllCall("GlobalFree", "ptr", Scan02)

         return e
      }
   }

   OnEvent(event, callback := "") {
      this.events[event] := callback
      return this
   }

   WindowProc(uMsg, wParam, lParam) {
      ; Because the first parameter of an object is "this",
      ; the callback function will overwrite that parameter as hwnd.
      hwnd := this

      ; A dictionary of "this" objects is stored as hwnd:this.
      this := TextRender.windows[hwnd]

      ; WM_DESTROY calls FreeMemory().
      if (uMsg = 0x2)
         return this.DestroyWindow()

      ; WM_DISPLAYCHANGE calls UpdateMemory().
      if (uMsg = 0x7E) {
         for i, layer in this.layers
            this.Draw(layer[1], layer[2], layer[3])
         return this.RenderOnScreen()
      }

      ; Match window messages to Rainmeter event names.
      ; https://docs.rainmeter.net/manual/mouse-actions/
      static dict :=
      ( LTrim Join
      {
         WM_LBUTTONDOWN := 0x0201    : "LeftMouseDown",
         WM_LBUTTONUP := 0x0202      : "LeftMouseUp",
         WM_LBUTTONDBLCLK := 0x0203  : "LeftMouseDoubleClick",
         WM_RBUTTONDOWN := 0x0204    : "RightMouseDown",
         WM_RBUTTONUP := 0x0205      : "RightMouseUp",
         WM_RBUTTONDBLCLK := 0x0206  : "RightMouseDoubleClick",
         WM_MBUTTONDOWN := 0x0207    : "MiddleMouseDown",
         WM_MBUTTONUP := 0x0208      : "MiddleMouseUp",
         WM_MBUTTONDBLCLK := 0x0209  : "MiddleMouseDoubleClick",
         WM_MOUSEHOVER := 0x02A1     : "MouseOver",
         WM_MOUSELEAVE := 0x02A3     : "MouseLeave"
      }
      )

      ; Process windows messages by invoking the associated callback.
      for message, event in dict
         if (uMsg = message)
            if callback := this.events[event]
               return %callback%(this) ; Callbacks have a reference to "this".

      ; Default processing of window messages.
      return DllCall("DefWindowProc", "ptr", hwnd, "uint", uMsg, "uptr", wParam, "ptr", lParam, "ptr")
   }

   EventMoveWindow() {
      ; Allows the user to drag to reposition the window.
      PostMessage 0xA1, 2,,, % "ahk_id" this.hwnd
   }

   EventShowCoordinates() {
      ; Shows a bubble displaying the current window coordinates.
      if !this.friend1 {
         this.friend1 := new TextRender(,,, this.hwnd)
         this.friend1.OnEvent("MiddleMouseDown", "")
      }
      CoordMode Mouse
      MouseGetPos _x, _y
      WinGetPos x, y, w, h, % "ahk_id " this.hwnd
      this.friend1.Render(Format("x:{:5} w:{:5}`r`ny:{:5} h:{:5}", x, w, y, h)
         , "t:7000 r:0.5vmin x" _x+20 " y" _y+20
         , "s:1.5vmin f:(Consolas) o:(0.5) m:0.5vmin j:right")
      WinSet AlwaysOnTop, On, % "ahk_id" this.friend1.hwnd
   }

   EventCopyText() {
      ; Copies the rendered text to clipboard.
      if !this.friend2 {
         this.friend2 := new TextRender(,,, this.hwnd)
         this.friend2.OnEvent("MiddleMouseDown", "")
         this.friend2.OnEvent("RightMouseDown", "")
      }
      clipboard := this.data
      this.friend2.Render("Saved text to clipboard.", "t:1250 c:#F9E486 y:75vh r:10%")
      WinSet AlwaysOnTop, On, % "ahk_id" this.friend2.hwnd
   }

   RegisterClass(vWinClass) {
      static atom := 0

      ; Return the atom to the class if present.
      if (atom)
         return atom

      ; Otherwise register the class name.
      pWndProc := RegisterCallback(this.WindowProc, "Fast",, &this)
      hCursor := DllCall("LoadCursor", "ptr", 0, "ptr", 32512, "ptr") ; IDC_ARROW

      ; struct tagWNDCLASSEXA - https://docs.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-wndclassexa
      ; struct tagWNDCLASSEXW - https://docs.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-wndclassexw
      _ := (A_PtrSize = 4)
      VarSetCapacity(wc, size := _ ? 48:80, 0)        ; sizeof(WNDCLASSEX) = 48, 80
         NumPut(       size, wc,         0,   "uint") ; cbSize
         NumPut(        0x8, wc,         4,   "uint") ; style = CS_DBLCLKS
         NumPut(   pWndProc, wc,         8,    "ptr") ; lpfnWndProc
         NumPut(          0, wc, _ ? 12:16,    "int") ; cbClsExtra
         NumPut(          0, wc, _ ? 16:20,    "int") ; cbWndExtra
         NumPut(          0, wc, _ ? 20:24,    "ptr") ; hInstance
         NumPut(          0, wc, _ ? 24:32,    "ptr") ; hIcon
         NumPut(    hCursor, wc, _ ? 28:40,    "ptr") ; hCursor
         NumPut(         16, wc, _ ? 32:48,    "ptr") ; hbrBackground
         NumPut(          0, wc, _ ? 36:56,    "ptr") ; lpszMenuName
         NumPut( &vWinClass, wc, _ ? 40:64,    "ptr") ; lpszClassName
         NumPut(          0, wc, _ ? 44:72,    "ptr") ; hIconSm

      ; Registers a window class for subsequent use in calls to the CreateWindow or CreateWindowEx function.
      return atom := DllCall("RegisterClassEx", "ptr", &wc, "ushort")
   }

   UnregisterClass(vWinClass) {
      return DllCall("UnregisterClass", "str", vWinClass, "ptr", 0, "int")
   }

   CreateWindow(title := "", WindowStyle := "", WindowExStyle := "", hwndParent := 0) {
      ; Window Styles - https://docs.microsoft.com/en-us/windows/win32/winmsg/window-styles
      WS_POPUP                  := 0x80000000

      ; Extended Window Styles - https://docs.microsoft.com/en-us/windows/win32/winmsg/extended-window-styles
      WS_EX_TOPMOST                 	:=        0x8
      WS_EX_TOOLWINDOW        	:=       0x80
      WS_EX_LAYERED                 	:=    0x80000
      WS_EX_NOACTIVATE          	:=  0x8000000

      if (WindowStyle = "")
         WindowStyle := WS_POPUP ; start off hidden with WS_VISIBLE off

      if (WindowExStyle = "")
         WindowExStyle := WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_LAYERED

      return DllCall("CreateWindowEx"
               ,   "uint", WindowExStyle                     ; dwExStyle
               , "ushort", this.RegisterClass("TextRender")  ; lpClassName
               ,    "str", title                             ; lpWindowName
               ,   "uint", WindowStyle                       ; dwStyle
               ,    "int", 0                                 ; X
               ,    "int", 0                                 ; Y
               ,    "int", 0                                 ; nWidth
               ,    "int", 0                                 ; nHeight
               ,    "ptr", hwndParent                        ; hWndParent
               ,    "ptr", 0                                 ; hMenu
               ,    "ptr", 0                                 ; hInstance
               ,    "ptr", 0                                 ; lpParam
               ,    "ptr")
   }

   ; Duality #2 - Destroys a window.
   DestroyWindow() {
      if (!this.hwnd)
         return this

      this.FreeMemory()
      DllCall("DestroyWindow", "ptr", this.hwnd)
      this.hwnd := ""
      return this
   }

   ; Duality #3 - Allocates the memory buffer.
   LoadMemory() {
      width := this.BitmapWidth, height := this.BitmapHeight

      ; struct BITMAPINFOHEADER - https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapinfoheader
      hdc := DllCall("CreateCompatibleDC", "ptr", 0, "ptr")
      VarSetCapacity(bi, 40, 0)              ; sizeof(bi) = 40
         NumPut(       40, bi,  0,   "uint") ; Size
         NumPut(    width, bi,  4,   "uint") ; Width
         NumPut(  -height, bi,  8,    "int") ; Height - Negative so (0, 0) is top-left.
         NumPut(        1, bi, 12, "ushort") ; Planes
         NumPut(       32, bi, 14, "ushort") ; BitCount / BitsPerPixel
      hbm := DllCall("CreateDIBSection", "ptr", hdc, "ptr", &bi, "uint", 0, "ptr*", pBits:=0, "ptr", 0, "uint", 0, "ptr")
      obm := DllCall("SelectObject", "ptr", hdc, "ptr", hbm, "ptr")
      gfx := DllCall("gdiplus\GdipCreateFromHDC", "ptr", hdc , "ptr*", gfx:=0, "int") ? false : gfx
      DllCall("gdiplus\GdipTranslateWorldTransform", "ptr", gfx, "float", -this.BitmapLeft, "float", -this.BitmapTop, "int", 0)

      this.hdc := hdc
      this.hbm := hbm
      this.obm := obm
      this.gfx := gfx
      this.BitmapBits := pBits
      this.BitmapStride := 4 * this.BitmapWidth
      this.BitmapSize := this.BitmapStride * this.BitmapHeight

      return this
   }

   ; Duality #3 - Frees the memory buffer.
   FreeMemory() {
      if (!this.hdc)
         return this

      DllCall("gdiplus\GdipDeleteGraphics", "ptr", this.gfx)
      DllCall("SelectObject", "ptr", this.hdc, "ptr", this.obm)
      DllCall("DeleteObject", "ptr", this.hbm)
      DllCall("DeleteDC",     "ptr", this.hdc)
      this.gfx := this.obm := this.hbm := this.hdc := ""
      return this
   }

   UpdateMemory() {
      ; Get true virtual screen coordinates.
      dpi := DllCall("SetThreadDpiAwarenessContext", "ptr", -4, "ptr")
      sx := DllCall("GetSystemMetrics", "int", 76, "int")
      sy := DllCall("GetSystemMetrics", "int", 77, "int")
      sw := DllCall("GetSystemMetrics", "int", 78, "int")
      sh := DllCall("GetSystemMetrics", "int", 79, "int")
      DllCall("SetThreadDpiAwarenessContext", "ptr", dpi, "ptr")

      if (sw = this.BitmapWidth && sh = this.BitmapHeight)
         return this

      this.BitmapLeft := sx
      this.BitmapTop := sy
      this.BitmapRight := sx + sw
      this.BitmapBottom := sy + sh
      this.BitmapWidth := sw
      this.BitmapHeight := sh
      this.FreeMemory()
      this.LoadMemory()

      return this
   }

   DebugMemory() {
      x := this.WindowLeft
      y := this.WindowTop
      w := Round(this.WindowWidth)
      h := Round(this.WindowHeight)

      ; Allocate buffer.
      VarSetCapacity(buffer, 4 * w * h, 0)

      ; Create a Bitmap with 32-bit pre-multiplied ARGB. (Owned by this object!)
      DllCall("gdiplus\GdipCreateBitmapFromScan0", "int", this.BitmapWidth, "int", this.BitmapHeight
         , "uint", this.BitmapStride, "uint", 0xE200B, "ptr", this.BitmapBits, "ptr*", pBitmap:=0)

      ; Specify that only a cropped bitmap portion will be copied.
      VarSetCapacity(Rect, 16, 0)            ; sizeof(Rect) = 16
         NumPut(      x, Rect,  0,    "int") ; X
         NumPut(      y, Rect,  4,    "int") ; Y
         NumPut(      w, Rect,  8,   "uint") ; Width
         NumPut(      h, Rect, 12,   "uint") ; Height
      VarSetCapacity(BitmapData, 16+2*A_PtrSize, 0)   ; sizeof(BitmapData) = 24, 32
         NumPut(       4*w, BitmapData,  8,    "int") ; Stride
         NumPut(   &buffer, BitmapData, 16,    "ptr") ; Scan0

      ; Convert pARGB to ARGB using a writable buffer created by LockBits.
      DllCall("gdiplus\GdipBitmapLockBits"
               ,    "ptr", pBitmap
               ,    "ptr", &Rect
               ,   "uint", 5            ; ImageLockMode.UserInputBuffer | ImageLockMode.ReadOnly
               ,    "int", 0x26200A     ; Format32bppArgb
               ,    "ptr", &BitmapData) ; Contains the buffer.
      DllCall("gdiplus\GdipBitmapUnlockBits", "ptr", pBitmap, "ptr", &BitmapData)

      ; Release reference to this.BitmapBits.
      DllCall("gdiplus\GdipDisposeImage", "ptr", pBitmap)

      ; Draw an enlarged pixel grid layout with printed color hexes.
      loop % h {
      _h := A_Index-1
         if _h*70 > A_ScreenHeight * 3
            break
      loop % w {
      _w := A_Index-1
         if _w*70 > A_ScreenWidth * 2
            continue
         formula := _h*w + _w
         pixel := Format("{:08X}", NumGet(buffer, 4*formula, "uint"))
         text := RegExReplace(pixel, "(.{4})(.{4})", "$1`r`n$2")
         this.Draw(text, "x" _w*70 " y"  70*(_h) " w70 h70 m0 c" pixel, "s:24pt v:center")
      }
      }

      ; Calling RenderOnScreen() is rather slow as every redraw happens again.
      this.RenderOnScreen()

      ; Note that this is a slow function in general. I'm not entirely sure how it can be sped up.
      return this
   }

   Hash() {
      return Format("{:08x}", DllCall("ntdll\RtlComputeCrc32", "uint", 0, "ptr", this.BitmapBits, "uptr", this.BitmapSize, "uint"))
   }

   CopyToBuffer() {
      ; Allocate buffer.
      buffer := DllCall("GlobalAlloc", "uint", 0, "uptr", 4 * this.w * this.h, "ptr")

      ; Create a Bitmap with 32-bit pre-multiplied ARGB. (Owned by this object!)
      DllCall("gdiplus\GdipCreateBitmapFromScan0", "int", this.BitmapWidth, "int", this.BitmapHeight
         , "uint", this.BitmapStride, "uint", 0xE200B, "ptr", this.BitmapBits, "ptr*", pBitmap:=0)

      ; Crop the bitmap.
      VarSetCapacity(Rect, 16, 0)            ; sizeof(Rect) = 16
         NumPut( this.x, Rect,  0,    "int") ; X
         NumPut( this.y, Rect,  4,    "int") ; Y
         NumPut( this.w, Rect,  8,   "uint") ; Width
         NumPut( this.h, Rect, 12,   "uint") ; Height
      VarSetCapacity(BitmapData, 16+2*A_PtrSize, 0)   ; sizeof(BitmapData) = 24, 32
         NumPut(  4*this.w, BitmapData,  8,    "int") ; Stride
         NumPut(    buffer, BitmapData, 16,    "ptr") ; Scan0

      ; Use LockBits to create a writable buffer that converts pARGB to ARGB.
      DllCall("gdiplus\GdipBitmapLockBits"
               ,    "ptr", pBitmap
               ,    "ptr", &Rect
               ,   "uint", 5            ; ImageLockMode.UserInputBuffer | ImageLockMode.ReadOnly
               ,    "int", 0x26200A     ; Format32bppArgb
               ,    "ptr", &BitmapData) ; Contains the buffer.
      DllCall("gdiplus\GdipBitmapUnlockBits", "ptr", pBitmap, "ptr", &BitmapData)

      DllCall("gdiplus\GdipDisposeImage", "ptr", pBitmap)

      return buffer
   }

   CopyToHBitmap() {
      ; struct BITMAPINFOHEADER - https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapinfoheader
      hdc := DllCall("CreateCompatibleDC", "ptr", 0, "ptr")
      VarSetCapacity(bi, 40, 0)              ; sizeof(bi) = 40
         NumPut(       40, bi,  0,   "uint") ; Size
         NumPut(   this.w, bi,  4,   "uint") ; Width
         NumPut(  -this.h, bi,  8,    "int") ; Height - Negative so (0, 0) is top-left.
         NumPut(        1, bi, 12, "ushort") ; Planes
         NumPut(       32, bi, 14, "ushort") ; BitCount / BitsPerPixel
      hbm := DllCall("CreateDIBSection", "ptr", hdc, "ptr", &bi, "uint", 0, "ptr*", pBits:=0, "ptr", 0, "uint", 0, "ptr")
      obm := DllCall("SelectObject", "ptr", hdc, "ptr", hbm, "ptr")

      DllCall("gdi32\BitBlt"
               , "ptr", hdc, "int", 0, "int", 0, "int", this.w, "int", this.h
               , "ptr", this.hdc, "int", this.x, "int", this.y, "uint", 0x00CC0020) ; SRCCOPY

      DllCall("SelectObject", "ptr", hdc, "ptr", obm)
      DllCall("DeleteDC",     "ptr", hdc)

      return hbm
   }

   RenderToHBitmap() {
      ; struct BITMAPINFOHEADER - https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapinfoheader
      hdc := DllCall("CreateCompatibleDC", "ptr", 0, "ptr")
      VarSetCapacity(bi, 40, 0)              ; sizeof(bi) = 40
         NumPut(       40, bi,  0,   "uint") ; Size
         NumPut(   this.w, bi,  4,   "uint") ; Width
         NumPut(  -this.h, bi,  8,    "int") ; Height - Negative so (0, 0) is top-left.
         NumPut(        1, bi, 12, "ushort") ; Planes
         NumPut(       32, bi, 14, "ushort") ; BitCount / BitsPerPixel
      hbm := DllCall("CreateDIBSection", "ptr", hdc, "ptr", &bi, "uint", 0, "ptr*", pBits:=0, "ptr", 0, "uint", 0, "ptr")
      obm := DllCall("SelectObject", "ptr", hdc, "ptr", hbm, "ptr")
      gfx := DllCall("gdiplus\GdipCreateFromHDC", "ptr", hdc , "ptr*", gfx:=0, "int") ? false : gfx

      ; Set the origin to this.x and this.y
      DllCall("gdiplus\GdipTranslateWorldTransform", "ptr", gfx, "float", -this.x, "float", -this.y, "int", 0)

      for i, layer in this.layers
         this.DrawOnGraphics(gfx, layer[1], layer[2], layer[3], this.BitmapWidth, this.BitmapHeight)

      DllCall("gdiplus\GdipDeleteGraphics", "ptr", gfx)
      DllCall("SelectObject", "ptr", hdc, "ptr", obm)
      DllCall("DeleteDC",     "ptr", hdc)

      return hbm
   }

   CopyToBitmap() {
      ; Create a Bitmap with 32-bit pre-multiplied ARGB. (Owned by this object!)
      DllCall("gdiplus\GdipCreateBitmapFromScan0", "int", this.BitmapWidth, "int", this.BitmapHeight
         , "uint", this.BitmapStride, "uint", 0xE200B, "ptr", this.BitmapBits, "ptr*", pBitmap:=0)

      ; Crop to fit and convert to 32-bit ARGB. (Managed impartially by GDI+)
      DllCall("gdiplus\GdipCloneBitmapAreaI", "int", this.x, "int", this.y, "int", this.w, "int", this.h
         , "uint", 0x26200A, "ptr", pBitmap, "ptr*", pBitmapCrop:=0)

      DllCall("gdiplus\GdipDisposeImage", "ptr", pBitmap)

      return pBitmapCrop
   }

   RenderToBitmap() {
      DllCall("gdiplus\GdipCreateBitmapFromScan0", "int", this.w, "int", this.h
         , "uint", 0, "uint", 0x26200A, "ptr", 0, "ptr*", pBitmap:=0)
      DllCall("gdiplus\GdipGetImageGraphicsContext", "ptr", pBitmap, "ptr*", gfx:=0)
      DllCall("gdiplus\GdipTranslateWorldTransform", "ptr", gfx, "float", -this.x, "float", -this.y, "int", 0)
      for i, layer in this.layers
         this.DrawOnGraphics(gfx, layer[1], layer[2], layer[3], this.BitmapWidth, this.BitmapHeight)
      DllCall("gdiplus\GdipDeleteGraphics", "ptr", gfx)
      return pBitmap
   }

   Save(filename := "", quality := "") {
      pBitmap := this.InBounds() ? this.CopyToBitmap() : this.RenderToBitmap()
      filepath := this.SaveImageToFile(pBitmap, filename, quality)
      DllCall("gdiplus\GdipDisposeImage", "ptr", pBitmap)
      return filepath
   }

   Screenshot(filename := "", quality := "") {
      pBitmap := this.GetImageFromScreen([this.x, this.y, this.w, this.h])
      filepath := this.SaveImageToFile(pBitmap, filename, quality)
      DllCall("gdiplus\GdipDisposeImage", "ptr", pBitmap)
      return filepath
   }

   SaveImageToFile(pBitmap, filepath := "", quality := "") {
      ; Thanks tic - https://www.autohotkey.com/boards/viewtopic.php?t=6517

      ; Remove whitespace. Seperate the filepath. Adjust for directories.
      filepath := Trim(filepath)
      SplitPath filepath,, directory, extension, filename
      if InStr(FileExist(filepath), "D")
         directory .= "\" filename, filename := ""
      if (directory != "" && !InStr(FileExist(directory), "D"))
         FileCreateDir % directory
      directory := (directory != "") ? directory : "."

      ; Validate filepath, defaulting to PNG. https://stackoverflow.com/a/6804755
      if !(extension ~= "^(?i:bmp|dib|rle|jpg|jpeg|jpe|jfif|gif|tif|tiff|png)$") {
         if (extension != "")
            filename .= "." extension
         extension := "png"
      }
      filename := RegExReplace(filename, "S)(?i:^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$|[<>:|?*\x00-\x1F\x22\/\\])")
      if (filename == "")
         FormatTime, filename,, % "yyyy-MM-dd HH꞉mm꞉ss"
      filepath := directory "\" filename "." extension

      ; Fill a buffer with the available encoders.
      DllCall("gdiplus\GdipGetImageEncodersSize", "uint*", count:=0, "uint*", size:=0)
      VarSetCapacity(ci, size)
      DllCall("gdiplus\GdipGetImageEncoders", "uint", count, "uint", size, "ptr", &ci)
      if !(count && size)
         throw Exception("Could not get a list of image codec encoders on this system.")

      ; Search for an encoder with a matching extension.
      Loop % count
         EncoderExtensions := StrGet(NumGet(ci, (idx:=(48+7*A_PtrSize)*(A_Index-1))+32+3*A_PtrSize, "uptr"), "UTF-16")
      until InStr(EncoderExtensions, "*." extension)

      ; Get the pointer to the index/offset of the matching encoder.
      if !(pCodec := &ci + idx)
         throw Exception("Could not find a matching encoder for the specified file format.")

      ; JPEG is a lossy image format that requires a quality value from 0-100. Default quality is 75.
      if (extension ~= "^(?i:jpg|jpeg|jpe|jfif)$"
      && 0 <= quality && quality <= 100 && quality != 75) {
         DllCall("gdiplus\GdipGetEncoderParameterListSize", "ptr", pBitmap, "ptr", pCodec, "uint*", size:=0)
         VarSetCapacity(EncoderParameters, size, 0)
         DllCall("gdiplus\GdipGetEncoderParameterList", "ptr", pBitmap, "ptr", pCodec, "uint", size, "ptr", &EncoderParameters)

         ; Search for an encoder parameter with 1 value of type 6.
         Loop % NumGet(EncoderParameters, "uint")
            elem := (24+A_PtrSize)*(A_Index-1) + A_PtrSize
         until (NumGet(EncoderParameters, elem+16, "uint") = 1) && (NumGet(EncoderParameters, elem+20, "uint") = 6)

         ; struct EncoderParameter - http://www.jose.it-berater.org/gdiplus/reference/structures/encoderparameter.htm
         ep := &EncoderParameters + elem - A_PtrSize                     ; sizeof(EncoderParameter) = 28, 32
            , NumPut(      1, ep+0,            0,   "uptr")              ; Must be 1.
            , NumPut(      4, ep+0, 20+A_PtrSize,   "uint")              ; Type
            , NumPut(quality, NumGet(ep+24+A_PtrSize, "uptr"), "uint")   ; Value (pointer)
      }

      ; Write the file to disk using the specified encoder and encoding parameters.
      Loop 6 ; Try this 6 times.
         if (A_Index > 1)
            Sleep % (2**(A_Index-2) * 30)
      until (result := !DllCall("gdiplus\GdipSaveImageToFile", "ptr", pBitmap, "wstr", filepath, "ptr", pCodec, "uint", (ep) ? ep : 0))
      if !(result)
         throw Exception("Could not save file to disk.")

      return filepath
   }

   GetImageFromScreen(image) {
      ; Thanks tic - https://www.autohotkey.com/boards/viewtopic.php?t=6517

      ; struct BITMAPINFOHEADER - https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapinfoheader
      hdc := DllCall("CreateCompatibleDC", "ptr", 0, "ptr")
      VarSetCapacity(bi, 40, 0)                ; sizeof(bi) = 40
         , NumPut(       40, bi,  0,   "uint") ; Size
         , NumPut( image[3], bi,  4,   "uint") ; Width
         , NumPut(-image[4], bi,  8,    "int") ; Height - Negative so (0, 0) is top-left.
         , NumPut(        1, bi, 12, "ushort") ; Planes
         , NumPut(       32, bi, 14, "ushort") ; BitCount / BitsPerPixel
      hbm := DllCall("CreateDIBSection", "ptr", hdc, "ptr", &bi, "uint", 0, "ptr*", pBits:=0, "ptr", 0, "uint", 0, "ptr")
      obm := DllCall("SelectObject", "ptr", hdc, "ptr", hbm, "ptr")

      ; Retrieve the device context for the screen.
      sdc := DllCall("GetDC", "ptr", 0, "ptr")

      ; Copies a portion of the screen to a new device context.
      DllCall("gdi32\BitBlt"
               , "ptr", hdc, "int", 0, "int", 0, "int", image[3], "int", image[4]
               , "ptr", sdc, "int", image[1], "int", image[2], "uint", 0x00CC0020 | 0x40000000) ; SRCCOPY | CAPTUREBLT

      ; Release the device context to the screen.
      DllCall("ReleaseDC", "ptr", 0, "ptr", sdc)

      ; Convert the hBitmap to a Bitmap using a built in function as there is no transparency.
      DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "ptr", hbm, "ptr", 0, "ptr*", pBitmap:=0)

      ; Cleanup the hBitmap and device contexts.
      DllCall("SelectObject", "ptr", hdc, "ptr", obm)
      DllCall("DeleteObject", "ptr", hbm)
      DllCall("DeleteDC",     "ptr", hdc)

      return pBitmap
   }

   UpdateLayeredWindow(x, y, w, h, alpha := 255) {
      return DllCall("UpdateLayeredWindow"
               ,    "ptr", this.hwnd                ; hWnd
               ,    "ptr", 0                        ; hdcDst
               ,"uint64*", x | y << 32              ; *pptDst
               ,"uint64*", w | h << 32              ; *psize
               ,    "ptr", this.hdc                 ; hdcSrc
               ,"uint64*", x - this.BitmapLeft
                        |  y - this.BitmapTop << 32 ; *pptSrc
               ,   "uint", 0                        ; crKey
               ,  "uint*", alpha << 16 | 0x01 << 24 ; *pblend
               ,   "uint", 2                        ; dwFlags
               ,    "int")                          ; Success = 1
   }

   InBounds() { ; Check if canvas coordinates are inside bitmap coordinates.
      return this.x >= this.BitmapLeft
         and this.y >= this.BitmapTop
         and this.x2 <= this.BitmapRight
         and this.y2 <= this.BitmapBottom
   }

   Bounds(default := "") {
      return (this.x2 > this.x && this.y2 > this.y) ? [this.x, this.y, this.x2, this.y2] : default
   }

   Rect(default := "") {
      return (this.x2 > this.x && this.y2 > this.y) ? [this.x, this.y, this.w, this.h] : default
   }

   ; All references to gdiplus and pToken must be absolute!
   static gdiplus := 0, pToken := 0

   gdiplusStartup() {
      TextRender.gdiplus++

      ; Startup gdiplus when counter goes from 0 -> 1.
      if (TextRender.gdiplus == 1) {

         ; Startup gdiplus.
         DllCall("LoadLibrary", "str", "gdiplus")
         VarSetCapacity(si, A_PtrSize = 4 ? 16:24, 0) ; sizeof(GdiplusStartupInput) = 16, 24
            , NumPut(0x1, si, "uint")
         DllCall("gdiplus\GdiplusStartup", "ptr*", pToken:=0, "ptr", &si, "ptr", 0)

         TextRender.pToken := pToken
      }
   }

   gdiplusShutdown(cotype := "", pBitmap := "") {
      TextRender.gdiplus--

      ; When a buffer object is deleted a bitmap is sent here for disposal.
      if (cotype == "smart_pointer")
         if DllCall("gdiplus\GdipDisposeImage", "ptr", pBitmap)
            throw Exception("The bitmap of this buffer object has already been deleted.")

      ; Check for unpaired calls of gdiplusShutdown.
      ; if (TextRender.gdiplus < 0)
      ;    throw Exception("Missing TextRender.gdiplusStartup().")

      ; Shutdown gdiplus when counter goes from 1 -> 0.
      if (TextRender.gdiplus == 0) {
         pToken := TextRender.pToken

         ; Shutdown gdiplus.
         DllCall("gdiplus\GdiplusShutdown", "ptr", pToken)
         DllCall("FreeLibrary", "ptr", DllCall("GetModuleHandle", "str", "gdiplus", "ptr"))

         ; Exit if GDI+ is still loaded. GdiplusNotInitialized = 18
         if (18 != DllCall("gdiplus\GdipCreateImageAttributes", "ptr*", ImageAttr:=0)) {
            DllCall("gdiplus\GdipDisposeImageAttributes", "ptr", ImageAttr)
            return
         }

         ; Otherwise GDI+ has been truly unloaded from the script and objects are out of scope.
         if (cotype = "bitmap")
            throw Exception("Out of scope error. `n`nIf you wish to handle raw pointers to GDI+ bitmaps, add the line"
               . "`n`n`t`t" this.__class ".gdiplusStartup()`n`nor 'pToken := Gdip_Startup()' to the top of your script."
               . "`nAlternatively, use 'obj := ImagePutBuffer()' with 'obj.pBitmap'."
               . "`nYou can copy this message by pressing Ctrl + C.")
      }
   }
} ; End of TextRender class.

TextRenderDesktop(text:="", background_style:="", text_style:="") {
   static WS_CHILD := (A_OSVersion = "WIN_7") ? 0x80000000 : 0x40000000 ; Fallback to WS_POPUP for Win 7.
   static WS_EX_LAYERED := 0x80000

   ; Used to show the desktop creations immediately.
   ; Post-Creator's Update Windows 10. WM_SPAWN_WORKER = 0x052C
   DllCall("SendMessage", "ptr", WinExist("ahk_class Progman"), "uint", 0x052C, "ptr", 0xD, "ptr", 0)
   DllCall("SendMessage", "ptr", WinExist("ahk_class Progman"), "uint", 0x052C, "ptr", 0xD, "ptr", 1)

   hwndParent := WinExist("ahk_class Progman")
   return (new TextRender(, WS_CHILD, WS_EX_LAYERED, hwndParent)).Render(text, background_style, text_style)
}

TextRenderWallpaper(text:="", background_style:="", text_style:="") {
   static WS_CHILD := (A_OSVersion = "WIN_7") ? 0x80000000 : 0x40000000 ; Fallback to WS_POPUP for Win 7.
   static WS_EX_LAYERED := 0x80000

   ; Post-Creator's Update Windows 10. WM_SPAWN_WORKER = 0x052C
   DllCall("SendMessage", "ptr", WinExist("ahk_class Progman"), "uint", 0x052C, "ptr", 0xD, "ptr", 0)
   DllCall("SendMessage", "ptr", WinExist("ahk_class Progman"), "uint", 0x052C, "ptr", 0xD, "ptr", 1)

   ; Find a child window of class SHELLDLL_DefView.
   WinGet windows, List, ahk_class WorkerW
   Loop % windows
      hwnd := windows%A_Index%
   until DllCall("FindWindowEx", "ptr", hwnd, "ptr", 0, "str", "SHELLDLL_DefView", "ptr", 0)

   ; Find a child window of the desktop after the previous window of class WorkerW.
   if !(WorkerW := DllCall("FindWindowEx", "ptr", 0, "ptr", hwnd, "str", "WorkerW", "ptr", 0, "ptr"))
      throw Exception("Could not locate hidden window behind desktop icons.")

   return (new TextRender(, WS_CHILD, WS_EX_LAYERED, WorkerW)).Render(text, background_style, text_style)
}

ImageRender(image:="", style:="", polygons:="") {
   return (new ImageRender).Render(image, style, polygons)
}

class ImageRender extends TextRender {

   DrawOnGraphics(gfx, pBitmap := "", style := "", polygons := "", CanvasWidth := "", CanvasHeight := "") {
      ; Get Graphics Width and Height.
      CanvasWidth := (CanvasWidth != "") ? CanvasWidth : NumGet(gfx + 20 + A_PtrSize, "uint")
      CanvasHeight := (CanvasHeight != "") ? CanvasHeight : NumGet(gfx + 24 + A_PtrSize, "uint")

      ; Remove excess whitespace for proper RegEx detection.
      style := !IsObject(style) ? RegExReplace(style, "\s+", " ") : style

      ; RegEx help? https://regex101.com/r/xLzZzO/2
      static q1 := "(?i)^.*?\b(?<!:|:\s)\b"
      static q2 := "(?!(?>\([^()]*\)|[^()]*)*\))(:\s*)?\(?(?<value>(?<=\()([\s:#%_a-z\-\.\d]+|\([\s:#%_a-z\-\.\d]*\))*(?=\))|[#%_a-z\-\.\d]+).*$"

      ; Extract styles to variables.
      if IsObject(style) {
         t  := (style.time != "")        ? style.time        : style.t
         a  := (style.anchor != "")      ? style.anchor      : style.a
         x  := (style.left != "")        ? style.left        : style.x
         y  := (style.top != "")         ? style.top         : style.y
         w  := (style.width != "")       ? style.width       : style.w
         h  := (style.height != "")      ? style.height      : style.h
         m  := (style.margin != "")      ? style.margin      : style.m
         s  := (style.scale != "")       ? style.scale       : style.s
         c  := (style.color != "")       ? style.color       : style.c
         q  := (style.quality != "")     ? style.quality     : (style.q) ? style.q : style.InterpolationMode
      } else {
         t  := ((___ := RegExReplace(style, q1    "(t(ime)?)"          q2, "${value}")) != style) ? ___ : ""
         a  := ((___ := RegExReplace(style, q1    "(a(nchor)?)"        q2, "${value}")) != style) ? ___ : ""
         x  := ((___ := RegExReplace(style, q1    "(x|left)"           q2, "${value}")) != style) ? ___ : ""
         y  := ((___ := RegExReplace(style, q1    "(y|top)"            q2, "${value}")) != style) ? ___ : ""
         w  := ((___ := RegExReplace(style, q1    "(w(idth)?)"         q2, "${value}")) != style) ? ___ : ""
         h  := ((___ := RegExReplace(style, q1    "(h(eight)?)"        q2, "${value}")) != style) ? ___ : ""
         m  := ((___ := RegExReplace(style, q1    "(m(argin)?)"        q2, "${value}")) != style) ? ___ : ""
         s  := ((___ := RegExReplace(style, q1    "(s(cale)?)"         q2, "${value}")) != style) ? ___ : ""
         c  := ((___ := RegExReplace(style, q1    "(c(olor)?)"         q2, "${value}")) != style) ? ___ : ""
         q  := ((___ := RegExReplace(style, q1    "(q(uality)?)"       q2, "${value}")) != style) ? ___ : ""
      }

      ; Extract the time variable and save it for a later when we Render() everything.
      static times := "(?i)^\s*(\d+)\s*(ms|mil(li(second)?)?|s(ec(ond)?)?|m(in(ute)?)?|h(our)?|d(ay)?)?s?\s*$"
      t  := ( t ~= times) ? RegExReplace( t, "\s", "") : 0 ; Default time is zero.
      t  := ((___ := RegExReplace( t, "i)(\d+)(ms|mil(li(second)?)?)s?$", "$1")) !=  t) ? ___ *        1 : t
      t  := ((___ := RegExReplace( t, "i)(\d+)s(ec(ond)?)?s?$"          , "$1")) !=  t) ? ___ *     1000 : t
      t  := ((___ := RegExReplace( t, "i)(\d+)m(in(ute)?)?s?$"          , "$1")) !=  t) ? ___ *    60000 : t
      t  := ((___ := RegExReplace( t, "i)(\d+)h(our)?s?$"               , "$1")) !=  t) ? ___ *  3600000 : t
      t  := ((___ := RegExReplace( t, "i)(\d+)d(ay)?s?$"                , "$1")) !=  t) ? ___ * 86400000 : t

      ; These are the type checkers.
      static valid := "(?i)^\s*(\-?(?:(?:\d+(?:\.\d*)?)|(?:\.\d+)))\s*(%|pt|px|vh|vmin|vw)?\s*$"
      static valid_positive := "(?i)^\s*((?:(?:\d+(?:\.\d*)?)|(?:\.\d+)))\s*(%|pt|px|vh|vmin|vw)?\s*$"

      ; Define viewport width and height. This is the visible screen area.
      vw := 0.01 * CanvasWidth         ; 1% of viewport width.
      vh := 0.01 * CanvasHeight        ; 1% of viewport height.
      vmin := (vw < vh) ? vw : vh      ; 1vw or 1vh, whichever is smaller.
      vr := CanvasWidth / CanvasHeight ; Aspect ratio of the viewport.

      ; Get original image width and height.
      DllCall("gdiplus\GdipGetImageWidth", "ptr", pBitmap, "uint*", width:=0)
      DllCall("gdiplus\GdipGetImageHeight", "ptr", pBitmap, "uint*", height:=0)
      minimum := (width < height) ? width : height
      aspect := width / height

      ; Get width and height.
      w  := ( w ~= valid_positive) ? RegExReplace( w, "\s", "") : ""
      w  := ( w ~= "i)(pt|px)$") ? SubStr( w, 1, -2) :  w
      w  := ( w ~= "i)vw$") ? RegExReplace( w, "i)vw$", "") * vw :  w
      w  := ( w ~= "i)vh$") ? RegExReplace( w, "i)vh$", "") * vh :  w
      w  := ( w ~= "i)vmin$") ? RegExReplace( w, "i)vmin$", "") * vmin :  w
      w  := ( w ~= "%$") ? RegExReplace( w, "%$", "") * 0.01 * width :  w

      h  := ( h ~= valid_positive) ? RegExReplace( h, "\s", "") : ""
      h  := ( h ~= "i)(pt|px)$") ? SubStr( h, 1, -2) :  h
      h  := ( h ~= "i)vw$") ? RegExReplace( h, "i)vw$", "") * vw :  h
      h  := ( h ~= "i)vh$") ? RegExReplace( h, "i)vh$", "") * vh :  h
      h  := ( h ~= "i)vmin$") ? RegExReplace( h, "i)vmin$", "") * vmin :  h
      h  := ( h ~= "%$") ? RegExReplace( h, "%$", "") * 0.01 * height :  h

      ; Default width and height.
      if (w == "" && h == "")
         w := width, h := height, wh_unset := true
      if (w == "")
         w := h * aspect
      if (h == "")
         h := w / aspect

      ; If scale is "fill" scale the image until there are no empty spaces but two sides of the image are cut off.
      ; If scale is "fit" scale the image so that the greatest edge will fit with empty borders along one edge.
      ; If scale is "harmonic" automatically downscale by the harmonic series. Ex: 50%, 33%, 25%, 20%...
      if (s = "auto" || s = "fill" || s = "fit" || s = "harmonic" || s = "limit") {
         if (wh_unset == true)
            w := CanvasWidth, h := CanvasHeight
         s := (s = "auto" || s = "limit")
            ? ((aspect > w / h) ? ((width > w) ? w / width : 1) : ((height > h) ? h / height : 1)) : s
         s := (s = "fill") ? ((aspect < w / h) ? w / width : h / height) : s
         s := (s = "fit") ? ((aspect > w / h) ? w / width : h / height) : s
         s := (s = "harmonic") ? ((aspect > w / h) ? 1 / (width // w + 1) : 1 / (height // h + 1)) : s
         w := width  ; width and height given were maximum values, not actual values.
         h := height ; Therefore restore the width and height to the image width and height.
      }

      s  := ( s ~= valid) ? RegExReplace( s, "\s", "") : ""
      s  := ( s ~= "i)(pt|px)$") ? SubStr( s, 1, -2) :  s
      s  := ( s ~= "i)vw$") ? RegExReplace( s, "i)vw$", "") * vw / width :  s
      s  := ( s ~= "i)vh$") ? RegExReplace( s, "i)vh$", "") * vh / height:  s
      s  := ( s ~= "i)vmin$") ? RegExReplace( s, "i)vmin$", "") * vmin / minimum :  s
      s  := ( s ~= "%$") ? RegExReplace( s, "%$", "") * 0.01 :  s

      ; If scale is negative automatically scale by a geometric series constant.
      ; Example: If scale is -0.5, then downscale by 50%, 25%, 12.5%, 6.25%...
      ; What the equation is asking is how many powers of -1/s can we fit in width/w?
      ; Then we use floor division and add 1 to ensure that we never exceed the bounds.
      ; While this is only designed to handle negative scales from 0 to -1.0,
      ; it works for negative numbers higher than -1.0. In this case, the 0 to -1 bounded
      ; are the left adjoint, meaning they never surpass the w and h. Higher negative Numbers
      ; are the right adjoint, meaning they never surpass w*-s and h*-s. Weird, huh.
      ; To clarify: Left adjoint: w*-s to w, h*-s to h. Right adjoint: w to w*-s, h to h*-s
      ; LaTex: \frac{1}{\frac{-1}{s}^{Floor(\frac{log(x)}{log(\frac{-1}{s})}) + 1}}
      ; Vertical asymptote at s := -1, which resolves to the empty string "".
      if (s < 0 && s != "") {
         if (wh_unset == true)
            w := CanvasWidth, h := CanvasHeight
         s := (s < 0) ? ((aspect > w / h)
            ? (-s) ** ((log(width/w) // log(-1/s)) + 1) : (-s) ** ((log(height/h) // log(-1/s)) + 1)) : s
         w := width  ; width and height given were maximum values, not actual values.
         h := height ; Therefore restore the width and height to the image width and height.
      }

      ; Default scale.
      if (s == "") {
         s := (x == "" && y == "" && wh_unset == true)         ; shrink image if x,y,w,h,s are all unset.
            ? ((aspect > vr)                                   ; determine whether width or height exceeds screen.
               ? ((width > CanvasWidth) ? CanvasWidth / width : 1)       ; scale will downscale image by its width.
               : ((height > CanvasHeight) ? CanvasHeight / height : 1))  ; scale will downscale image by its height.
            : 1                                                ; Default scale is 1.00.
      }

      ; Scale width and height.
      w  := w * s
      h  := h * s

      ; Get anchor. This is where the origin of the image is located.
      a  := RegExReplace( a, "\s", "")
      a  := (a ~= "i)top" && a ~= "i)left") ? 1 : (a ~= "i)top" && a ~= "i)cent(er|re)") ? 2
         : (a ~= "i)top" && a ~= "i)right") ? 3 : (a ~= "i)cent(er|re)" && a ~= "i)left") ? 4
         : (a ~= "i)cent(er|re)" && a ~= "i)right") ? 6 : (a ~= "i)bottom" && a ~= "i)left") ? 7
         : (a ~= "i)bottom" && a ~= "i)cent(er|re)") ? 8 : (a ~= "i)bottom" && a ~= "i)right") ? 9
         : (a ~= "i)top") ? 2 : (a ~= "i)left") ? 4 : (a ~= "i)right") ? 6 : (a ~= "i)bottom") ? 8
         : (a ~= "i)cent(er|re)") ? 5 : (a ~= "^[1-9]$") ? a : 1 ; Default anchor is top-left.

      ; The anchor can be implied and overwritten by x and y (left, center, right, top, bottom).
      a  := ( x ~= "i)left") ? 1+((( a-1)//3)*3) : ( x ~= "i)cent(er|re)") ? 2+((( a-1)//3)*3) : ( x ~= "i)right") ? 3+((( a-1)//3)*3) :  a
      a  := ( y ~= "i)top") ? 1+(mod( a-1,3)) : ( y ~= "i)cent(er|re)") ? 4+(mod( a-1,3)) : ( y ~= "i)bottom") ? 7+(mod( a-1,3)) :  a

      ; Convert English words to numbers. Don't mess with these values any further.
      x  := ( x ~= "i)left") ? 0 : (x ~= "i)cent(er|re)") ? 0.5*CanvasWidth : (x ~= "i)right") ? CanvasWidth : x
      y  := ( y ~= "i)top") ? 0 : (y ~= "i)cent(er|re)") ? 0.5*CanvasHeight : (y ~= "i)bottom") ? CanvasHeight : y

      ; Get x and y.
      x  := ( x ~= valid) ? RegExReplace( x, "\s", "") : ""
      x  := ( x ~= "i)(pt|px)$") ? SubStr( x, 1, -2) :  x
      x  := ( x ~= "i)(%|vw)$") ? RegExReplace( x, "i)(%|vw)$", "") * vw :  x
      x  := ( x ~= "i)vh$") ? RegExReplace( x, "i)vh$", "") * vh :  x
      x  := ( x ~= "i)vmin$") ? RegExReplace( x, "i)vmin$", "") * vmin :  x

      y  := ( y ~= valid) ? RegExReplace( y, "\s", "") : ""
      y  := ( y ~= "i)(pt|px)$") ? SubStr( y, 1, -2) :  y
      y  := ( y ~= "i)vw$") ? RegExReplace( y, "i)vw$", "") * vw :  y
      y  := ( y ~= "i)(%|vh)$") ? RegExReplace( y, "i)(%|vh)$", "") * vh :  y
      y  := ( y ~= "i)vmin$") ? RegExReplace( y, "i)vmin$", "") * vmin :  y

      ; Default x and y.
      if (x == "")
         x := 0.5*CanvasWidth, a := 2+((( a-1)//3)*3)
      if (y == "")
         y := 0.5*CanvasHeight, a := 4+(mod( a-1,3))

      ; Modify x and y values with the anchor, so that the image has a new point of origin.
      x  -= (mod(a-1,3) == 0) ? 0 : (mod(a-1,3) == 1) ? w/2 : (mod(a-1,3) == 2) ? w : 0
      y  -= (((a-1)//3) == 0) ? 0 : (((a-1)//3) == 1) ? h/2 : (((a-1)//3) == 2) ? h : 0

      ; Prevent half-pixel rendering and keep image sharp.
      w  := Round(x + w) - Round(x)    ; Use real x2 coordinate to determine width.
      h  := Round(y + h) - Round(y)    ; Use real y2 coordinate to determine height.
      x  := Round(x)                   ; NOTE: simple Floor(w) or Round(w) will NOT work.
      y  := Round(y)                   ; The float values need to be added up and then rounded!

      ; Get margin.
      m  := this.parse.margin_and_padding(m, vw, vh)

      ; Calculate border using margin.
      _w := w + Round(m.2) + Round(m.4)
      _h := h + Round(m.1) + Round(m.3)
      _x := x - Round(m.4)
      _y := y - Round(m.1)

      ; Save original Graphics settings.
      DllCall("gdiplus\GdipSaveGraphics", "ptr", gfx, "ptr*", pState:=0)

      ; Set some general Graphics settings.
      DllCall("gdiplus\GdipSetPixelOffsetMode",    "ptr",gfx, "int",2) ; Half pixel offset.
      DllCall("gdiplus\GdipSetCompositingMode",    "ptr",gfx, "int",1) ; Overwrite/SourceCopy.
      DllCall("gdiplus\GdipSetCompositingQuality", "ptr",gfx, "int",0) ; AssumeLinear
      DllCall("gdiplus\GdipSetSmoothingMode",      "ptr",gfx, "int",0) ; No anti-alias.
      DllCall("gdiplus\GdipSetInterpolationMode",  "ptr",gfx, "int",7) ; HighQualityBicubic

      ; Begin drawing the image onto the canvas.
      if (pBitmap != "") {

         ; Draw background if color or margin is set.
         if (c != "" || !m.void) {
            c := this.parse.color(c, 0xDD212121) ; Default color is transparent gray.
            if (c & 0xFF000000) {
               DllCall("gdiplus\GdipSetSmoothingMode", "ptr", gfx, "int", 0) ; SmoothingModeNoAntiAlias
               DllCall("gdiplus\GdipCreateSolidFill", "uint", c, "ptr*", pBrush:=0)
               DllCall("gdiplus\GdipFillRectangle", "ptr", gfx, "ptr", pBrush, "float", _x, "float", _y, "float", _w, "float", _h) ; DRAWING!
               DllCall("gdiplus\GdipDeleteBrush", "ptr", pBrush)
            }
         }

         ; Draw image using GDI.
         if (q = 0 || w == width && h == height) {
            ; Get a read-only device context associated with the Graphics object.
            DllCall("gdiplus\GdipGetDC", "ptr", gfx, "ptr*", ddc:=0)

            ; Allocate a top-down device independent bitmap (hbm) by inputting a negative height.
            ; Outputs a pointer to the pixel data. Select the new handle to a bitmap onto the cloned
            ; compatible device context. The old bitmap (obm) is a monochrome 1x1 default bitmap that
            ; will be reselected onto the device context (hdc) before deletion.
            ; struct BITMAPINFOHEADER - https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapinfoheader
            hdc := DllCall("CreateCompatibleDC", "ptr", 0, "ptr")
            VarSetCapacity(bi, 40, 0)              ; sizeof(bi) = 40
               NumPut(       40, bi,  0,   "uint") ; Size
               NumPut(    width, bi,  4,   "uint") ; Width
               NumPut(  -height, bi,  8,    "int") ; Height - Negative so (0, 0) is top-left.
               NumPut(        1, bi, 12, "ushort") ; Planes
               NumPut(       32, bi, 14, "ushort") ; BitCount / BitsPerPixel
            hbm := DllCall("CreateDIBSection", "ptr", hdc, "ptr", &bi, "uint", 0, "ptr*", pBits:=0, "ptr", 0, "uint", 0, "ptr")
            obm := DllCall("SelectObject", "ptr", hdc, "ptr", hbm, "ptr")

            ; The following routine is 4ms faster than hbm := Gdip_CreateHBITMAPFromBitmap(pBitmap).
            ; In the below code we do something really interesting to save a call of memcpy().
            ; When calling LockBits the third argument is set to 0x4 (ImageLockModeUserInputBuf).
            ; This means that we can use the pointer to the bits from our memory bitmap (DIB)
            ; as the Scan0 of the LockBits output. While this is not a speed up, this saves memory
            ; because we are (1) allocating a DIB, (2) getting a pBitmap, (3) using a LockBits buffer.
            ; Instead of LockBits creating a new buffer, we can use the allocated buffer from (1).
            ; The bottleneck in the code is LockBits(), which takes over 20 ms for a 1920 x 1080 image.
            ; https://stackoverflow.com/questions/6782489/create-bitmap-from-a-byte-array-of-pixel-data
            ; https://stackoverflow.com/questions/17030264/read-and-write-directly-to-unlocked-bitmap-unmanaged-memory-scan0
            VarSetCapacity(Rect, 16, 0)              ; sizeof(Rect) = 16
               NumPut(    width, Rect,  8,   "uint") ; Width
               NumPut(   height, Rect, 12,   "uint") ; Height
            VarSetCapacity(BitmapData, 16+2*A_PtrSize, 0)     ; sizeof(BitmapData) = 24, 32
               NumPut(   4 * width, BitmapData,  8,    "int") ; Stride
               NumPut(       pBits, BitmapData, 16,    "ptr") ; Scan0
            DllCall("gdiplus\GdipBitmapLockBits"
                     ,    "ptr", pBitmap
                     ,    "ptr", &Rect
                     ,   "uint", 5            ; ImageLockMode.UserInputBuffer | ImageLockMode.ReadOnly
                     ,    "int", 0xE200B      ; Format32bppPArgb
                     ,    "ptr", &BitmapData)
            DllCall("gdiplus\GdipBitmapUnlockBits", "ptr", pBitmap, "ptr", &BitmapData)

            ; A good question to ask is why don't we use the pBits already associated with the graphics hdc?
            ; One, if a Graphics object associated to a pBitmap via Gdip_GraphicsFromImage() is passed,
            ; there would be no underlying device independent bitmap, and thus no pBits at all!
            ; Two, since the size of the allocated DIB is not the same size as the underlying DIB,
            ; injection using x,y,w,h coordinates is required, and BitBlt supports this.
            ; Note: The Rect in LockBits is crops the image source and does not affect the destination.

            (c != "" || !m.void) ; Check if color or margin is set to invoke AlphaBlend, otherwise BitBlt.

            ; AlphaBlend() does not overwrite the underlying pixels.
            ? DllCall("msimg32\AlphaBlend"
                     , "ptr", ddc, "int", x, "int", y, "int", w    , "int", h
                     , "ptr", hdc, "int", 0, "int", 0, "int", width, "int", height
                     , "uint", 0xFF << 16 | 0x01 << 24) ; BlendFunction

            ; BitBlt() is the fastest operation for copying pixels.
            : DllCall("gdi32\StretchBlt"
                     , "ptr", ddc, "int", x, "int", y, "int", w    , "int", h
                     , "ptr", hdc, "int", 0, "int", 0, "int", width, "int", height
                     , "uint", 0x00CC0020) ; SRCCOPY

            DllCall("SelectObject", "ptr", hdc, "ptr", obm)
            DllCall("DeleteObject", "ptr", hbm)
            DllCall("DeleteDC",     "ptr", hdc)

            DllCall("gdiplus\GdipReleaseDC", "ptr", gfx, "ptr", ddc)
         }

         ; Draw image scaled to a new width and height.
         else {
            ; Set InterpolationMode.
            q := (q >= 0 && q <= 7) ? q : 7    ; HighQualityBicubic

            DllCall("gdiplus\GdipSetPixelOffsetMode",    "ptr", gfx, "int", 2) ; Half pixel offset.
            DllCall("gdiplus\GdipSetCompositingMode",    "ptr", gfx, "int", 1) ; Overwrite/SourceCopy.
            DllCall("gdiplus\GdipSetSmoothingMode",      "ptr", gfx, "int", 0) ; No anti-alias.
            DllCall("gdiplus\GdipSetInterpolationMode",  "ptr", gfx, "int", q)
            DllCall("gdiplus\GdipSetCompositingQuality", "ptr", gfx, "int", 0) ; AssumeLinear

            ; Draw image with proper edges and scaling.
            DllCall("gdiplus\GdipCreateImageAttributes", "ptr*", ImageAttr)
            DllCall("gdiplus\GdipSetImageAttributesWrapMode", "ptr", ImageAttr, "int", 3) ; WrapModeTileFlipXY
            DllCall("gdiplus\GdipDrawImageRectRectI"
                     ,    "ptr", gfx
                     ,    "ptr", pBitmap
                     ,    "int", x, "int", y, "int", w    , "int", h      ; destination rectangle
                     ,    "int", 0, "int", 0, "int", width, "int", height ; source rectangle
                     ,    "int", 2                                        ; UnitTypePixel
                     ,    "ptr", ImageAttr                                ; imageAttributes
                     ,    "ptr", 0                                        ; callback
                     ,    "ptr", 0)                                       ; callbackData
            DllCall("gdiplus\GdipDisposeImageAttributes", "ptr", ImageAttr)
         }
      }

      ; Begin drawing the polygons onto the canvas.
      if (polygons != "") {
         DllCall("gdiplus\GdipSetPixelOffsetMode",   "ptr",gfx, "int",0) ; No pixel offset.
         DllCall("gdiplus\GdipSetCompositingMode",   "ptr",gfx, "int",1) ; Overwrite/SourceCopy.
         DllCall("gdiplus\GdipSetSmoothingMode",     "ptr",gfx, "int",2) ; Use anti-alias.

         DllCall("gdiplus\GdipCreatePen1", "uint", 0xFFFF0000, "float", 1, "int", 2, "ptr*", pPen:=0)

         for i, polygon in polygons {
            DllCall("gdiplus\GdipCreatePath", "int",1, "ptr*",pPath)
            VarSetCapacity(pointf, 8*polygons[i].polygon.maxIndex(), 0)
            for j, point in polygons[i].polygon {
               NumPut(point.x*s + x, pointf, 8*(A_Index-1) + 0, "float")
               NumPut(point.y*s + y, pointf, 8*(A_Index-1) + 4, "float")
            }
            DllCall("gdiplus\GdipAddPathPolygon", "ptr",pPath, "ptr",&pointf, "uint",polygons[i].polygon.maxIndex())
            DllCall("gdiplus\GdipDrawPath", "ptr",gfx, "ptr",pPen, "ptr",pPath) ; DRAWING!
         }

         DllCall("gdiplus\GdipDeletePen", "ptr", pPen)
      }

      ; Restore original Graphics settings.
      DllCall("gdiplus\GdipRestoreGraphics", "ptr", gfx, "ptr", pState)

      ; Define bounds.
      t_bound :=  t
      x_bound := _x
      y_bound := _y
      w_bound := _w
      h_bound := _h

      return {t: t_bound
            , x: x_bound, y: y_bound
            , w: w_bound, h: h_bound
            , x2: x_bound + w_bound, y2: y_bound + h_bound}
   }
} ; End of ImageRender class.

;}

/*
SciTEOutput("Factor: " xlPos.Factor
							. 	"`nZoom: "  xlPos.Zoom
							. 	"`nWinX: "  xlPos.WinX
							.	"`nWinY: " xlPos.WinY
							.	"`nUsaW: " xlPos.UsableW
							.	"`nUsaH: " xlPos.UsableH
							.	"`nSelX: " xlPos.SelX
							.	"`nSelY: " xlPos.SelY
							.	"`nSelW: " xlPos.SelW
							.	"`nSelH: " xlPos.SelH)


					;~ If IsObject(this.tr)                       ; nur ein Fenster erlauben
						;~ trOld := this.tr

					;~ this.tr   	:= TextRender(RTrim(Tele, "`n")
													;~ , "time:30000 color:Green radius:10 margin:(5px 5px)"
													;~ , "color:White s18 font:(Futura Bk Bt) outline:(stroke:1px color:Green)")
					;~ Win    	:= GetWindowSpot(this.tr.hwnd)
					;~ xlHwnd	:= WinExist(XL.xlName  " ahk_class XLMAIN")
					;~ ControlGet, xlDeskHwnd, Hwnd,, XLDESK1, % "ahk_id " xlHwnd
					;~ xlPos		:= GetWindowSpot(xlDeskHwnd)

					;~ SetWindowPos(this.tr.hwnd, xlPos.X + Floor(xlPos.W/2-Win.W/2) , xlPos.Y+20, Win.W, Win.H)
					;SetWindowPos(this.tr.hwnd, HeaderPos.WinX + Floor(HeaderPos.UsaW/2-Win.W/2) , HeaderPos.SelY, Win.W, Win.H)

					;~ If IsObject(trOld)  {                     ; nur ein Fenster erlauben
						;~ trOld.DestroyWindow()
						;~ trOld := ""
					;~ }

 */


;{ INCLUDES
#Include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Datum.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PdfHelper.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Window.ahk
#include %A_ScriptDir%\..\..\include\Addendum_PraxTT.ahk

#Include %A_ScriptDir%\..\..\lib\ACC.ahk
#Include %A_ScriptDir%\..\..\lib\Gdip_All.ahk
#Include %A_ScriptDir%\..\..\lib\class_JSON.ahk
#Include %A_ScriptDir%\..\..\lib\class_socket.ahk
#Include %A_ScriptDir%\..\..\lib\SciteOutput.ahk
#Include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\lib\Sift.ahk

;}





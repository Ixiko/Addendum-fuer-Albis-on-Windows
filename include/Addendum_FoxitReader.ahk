;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	 RPA (robotic prozess automation) FUNKTIONEN für
;                                  FOXIT PDF READER V9.1 als FUNKTIONSBIBLIOTHEK FÜR DIE NUTZUNG IN SCANPOOL.AHK
;                                                                                  	------------------------
;                                                  	FÜR DAS AIS-ADDON: "ADDENDUM FÜR ALBIS ON WINDOWS"
;                                                                                  	------------------------
;    		BY IXIKO STARTED IN SEPTEMBER 2017 - LAST CHANGE 12.07.2020 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

;----------------------------------------------------------------------------------------------------------------------------------------------
; FOXIT READER
;----------------------------------------------------------------------------------------------------------------------------------------------

FoxitInvoke(command, FoxitID := "") {		                                                           	;-- wm_command wrapper for FoxitReader Version:  9.1

		/* DESCRIPTION of FUNCTION:  FoxitInvoke() by Ixiko (version 11.07.2020)

		---------------------------------------------------------------------------------------------------
												a WM_command wrapper for FoxitReader V9.1 by Ixiko
																		...........................................................
													 Remark: maybe not all commands are listed at now!
		---------------------------------------------------------------------------------------------------
				by use  of a valid FoxitID, this function will post your command to FoxitReader
			                                             otherwise this function returns the command code
																		...........................................................
			Remark: You have to control the success of the postmessage command yourself!
		---------------------------------------------------------------------------------------------------
						I intentionally use a text first and then convert it to a -Key: Value- object,
                                                         so you can swap out the object to a file if needed
		---------------------------------------------------------------------------------------------------
		      EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES

		FoxitInvoke("Show_FullPage")                       FoxitInvoke("Place_Signature", FoxitID)
		.................................................          ...............................................................
		this one only returns the Foxit                  sends the command "Place_Signature" to
        command-code                                      your specified FoxitReader process using
																	 parameter 2 (FoxitID) as window handle.
																	        command-code will be returned too
		---------------------------------------------------------------------------------------------------

	*/

	static FoxitCommands
	If !IsObject(FoxitCommands) {

			FoxitCommands := {	"SaveAs":                                                       	1299
											,	"Close":                                                         	57602
	                                		,	"Hand":                                                         	1348        	; Home - Tools
	                                		,	"Select_Text":                                                 	46178      	; Home - Tools
	                                		,	"Select_Annotation":                                       	46017      	; Home - Tools
	                                		,	"Snapshot":                                                    	46069      	; Home - Tools
	                                		,	"Clipboard_SelectAll":                                    	57642      	; Home - Tools
	                                		,	"Clipboard_Copy":                                         	57634      	; Home - Tools
	                                		,	"Clipboard_Paste":                                         	57637      	; Home - Tools
	                                		,	"Actual_Size":                                                 	1332        	; Home - View
	                                		,	"Fit_Page":                                                     	1343        	; Home - View
	                                		,	"Fit_Width":                                                   	1345        	; Home - View
	                                		,	"Reflow":                                                        	32818      	; Home - View
	                                		,	"Zoom_Field":                                                	1363        	; Home - View
	                                		,	"Zoom_Plus":                                                 	1360        	; Home - View
	                                		,	"Zoom_Minus":                                              	1362        	; Home - View
	                                		,	"Rotate_Left":                                                 	1340        	; Home - View
	                                		,	"Rotate_Right":                                               	1337        	; Home - View
	                                		,	"Highlight":                                                    	46130      	; Home - Comment
	                                		,	"Typewriter":                                                  	46096      	; Home - Comment, Comment - TypeWriter
	                                		,	"Open_From_File":                                        	46140      	; Home - Create
	                                		,	"Open_Blank":                                               	46141      	; Home - Create
	                                		,	"Open_From_Scanner":                                  	46165      	; Home - Create
	                                		,	"Open_From_Clipboard":                               	46142      	; Home - Create - new pdf from clipboard
	                                		,	"PDF_Sign":                                                   	46157      	;Home - Protect
	                                		,	"Create_Link":                                                	46080      	; Home - Links
	                                		,	"Create_Bookmark":                                       	46070      	; Home - Links
	                                		,	"File_Attachment":                                          	46094      	; Home - Insert
	                                		,	"Image_Annotation":                                      	46081      	; Home - Insert
	                                		,	"Audio_and_Video":                                       	46082      	; Home - Insert
	                                		,	"Comments_Import":                                      	46083      	; Comments
	                                		,	"Highlight":                                                    	46130      	; Comments - Text Markup
	                                		,	"Squiggly_Underline":                                     	46131      	; Comments - Text Markup
	                                		,	"Underline":                                                   	46132      	; Comments - Text Markup
	                                		,	"Strikeout":                                                     	46133      	; Comments - Text Markup
	                                		,	"Replace_Text":                                              	46134      	; Comments - Text Markup
	                                		,	"Insert_Text":                                                  	46135      	; Comments - Text Markup
	                                		,	"Note":                                                          	46137      	; Comments - Pin
	                                    	,	"File":                                                            	46095      	; Comments - Pin
	                                    	,	"Callout":                                                       	46097      	; Comments - Typewriter
	                                    	,	"Textbox":                                                      	46098      	; Comments - Typewriter
	                                    	,	"Rectangle":                                                   	46101      	; Comments - Drawing
	                                    	,	"Oval":                                                          	46102      	; Comments - Drawing
	                                    	,	"Polygon":                                                      	46103      	; Comments - Drawing
	                                    	,	"Cloud":                                                        	46104      	; Comments - Drawing
	                                    	,	"Arrow":                                                         	46105      	; Comments - Drawing
	                                    	,	"Line":                                                           	46106      	; Comments - Drawing
	                                    	,	"Polyline":                                                      	46107      	; Comments - Drawing
	                                    	,	"Pencil":                                                         	46108      	; Comments - Drawing
	                                    	,	"Eraser":                                                        	46109      	; Comments - Drawing
	                                    	,	"Area_Highligt":                                             	46136      	; Comments - Drawing
	                                    	,	"Distance":                                                     	46110      	; Comments - Measure
	                                    	,	"Perimeter":                                                   	46111      	; Comments - Measure
	                                    	,	"Area":                                                          	46112      	; Comments - Measure
	                                    	,	"Stamp":                                                        	46149      	; Comments - Stamps , opens only the dialog
	                                    	,	"Create_custom_stamp":                                 	46151      	; Comments - Stamps
	                                    	,	"Create_custom_dynamic_stop":                     	46152      	; Comments - Stamps
	                                    	,	"Summarize_Comments":                                	46188      	; Comments - Manage Comments
	                                    	,	"Import":                                                        	46083      	; Comments - Manage Comments
	                                    	,	"Export_All_Comments":                                  	46086      	; Comments - Manage Comments
	                                    	,	"Export_Highlighted_Texts":                            	46087      	; Comments - Manage Comments
	                                    	,	"FDF_via_Email":                                            	46084      	; Comments - Manage Comments
	                                    	,	"Comments":                                                 	46088      	; Comments - Manage Comments
	                                    	,	"Comments_Show_All":                                   	46089      	; Comments - Manage Comments
	                                    	,	"Comments_Hide_All":                                   	46090      	; Comments - Manage Comments
	                                    	,	"Popup_Notes":                                               	46091      	; Comments - Manage Comments
	                                    	,	"Popup_Notes_Open_All":                                	46092      	; Comments - Manage Comments
	                                    	,	"Popup_Notes_Close_All":                               	46093 }     	; Comments - Manage Comments

			 FoxitCommands .= {	"firstPage":                                                      	1286        	; View - Go To
	                                		,	"lastPage":                                                      	1288        	; View - Go To
                                        	,	"nextPage":                                                     	1289        	; View - Go To
                                        	,	"previousPage":                                               	1290        	; View - Go To
                                        	,	"previousView":                                               	1335        	; View - Go To
                                        	,	"nextView":                                                     	1346        	; View - Go To
                                        	,	"ReadMode":                                                 	1351        	; View - Document Views
                                        	,	"ReverseView":                                               	1353        	; View - Document Views
                                        	,	"TextViewer":                                                  	46180      	; View - Document Views
                                        	,	"Reflow":                                                        	32818      	; View - Document Views
                                        	,	"turnPage_left":                                              	1340        	; View - Page Display
                                        	,	"turnPage_right":                                            	1337        	; View - Page Display
                                        	,	"SinglePage":                                                 	1357        	; View - Page Display
                                        	,	"Continuous":                                                	1338        	; View - Page Display
                                        	,	"Facing":                                                       	1356        	; View - Page Display - two pages side by side
                                        	,	"Continuous_Facing":                                     	1339        	; View - Page Display - two pages side by side with scrolling enabled
                                        	,	"Separate_CoverPage":                                  	1341        	; View - Page Display
                                        	,	"Horizontally_Split":                                        	1364        	; View - Page Display
                                        	,	"Vertically_Split":                                            	1365        	; View - Page Display
                                        	,	"Spreadsheet_Split":                                       	1368        	; View - Page Display
                                        	,	"Guides":                                                       	1354        	; View - Page Display
                                        	,	"Rulers":                                                        	1355        	; View - Page Display
                                        	,	"LineWeights":                                               	1350        	; View - Page Display
                                        	,	"AutoScroll":                                                  	1334        	; View - Assistant
                                        	,	"Marquee":                                                    	1361        	; View - Assistant
                                        	,	"Loupe":                                                        	46138      	; View - Assistant
                                        	,	"Magnifier":                                                   	46139      	; View - Assistant
                                        	,	"Read_Activate":                                             	46198      	; View - Read
                                        	,	"Read_CurrentPage":                                      	46199      	; View - Read
                                        	,	"Read_from_CurrentPage":                             	46200      	; View - Read
                                        	,	"Read_Stop":                                                  	46201      	; View - Read
                                        	,	"Read_Pause":                                               	46206      	; View - Read
	                                		,	"Navigation_Panels":                                      	46010      	; View - View Setting
	                                		,	"Navigation_Bookmark":                                	45401      	; View - View Setting
	                                		,	"Navigation_Pages":                                      	45402      	; View - View Setting
	                                		,	"Navigation_Layers":                                      	45403      	; View - View Setting
	                                		,	"Navigation_Comments":                               	45404      	; View - View Setting
	                                		,	"Navigation_Appends":                                  	45405      	; View - View Setting
	                                		,	"Navigation_Security":                                    	45406      	; View - View Setting
	                                		,	"Navigation_Signatures":                                	45408      	; View - View Setting
	                                		,	"Navigation_WinOff":                                    	1318        	; View - View Setting
	                                		,	"Navigation_ResetAllWins":                             	1316        	; View - View Setting
	                                		,	"Status_Bar":                                                  	46008        	; View - View Setting
	                                		,	"Status_Show":                                               	1358        	; View - View Setting
	                                		,	"Status_Auto_Hide":                                       	1333        	; View - View Setting
	                                		,	"Status_Hide":                                                	1349        	;View - View Setting
	                                		,	"WordCount":                                                	46179      	;View - Review
	                                		,	"Form_to_sheet":                                            	46072      	;Form - Form Data
	                                		,	"Combine_Forms_to_a_sheet":                        	46074      	;Form - Form Data
	                                		,	"DocuSign":                                                   	46189      	;Protect
	                                		,	"Login_to_DocuSign":                                     	46190      	;Protect
	                                		,	"Sign_with_DocuSign":                                   	46191      	;Protect
	                                		,	"Send_via_DocuSign":                                    	46192      	;Protect
	                                		,	"Sign_and_Certify":                                        	46181      	;Protect
	                                		,	"-----_-------------":                                         	46182      	;Protect
	                                		,	"Place_Signature":                                          	46183      	;Protect
	                                		,	"Validate":                                                     	46185      	;Protect
	                                		,	"Time_Stamp_Document":                              	46184      	;Protect
	                                		,	"Digital_IDs":                                                 	46186      	;Protect
	                                		,	"Trusted_Certificates":                                     	46187      	;Protect
	                                		,	"Email":                                                         	1296        	;Share - Send To - same like Email current tab
	                                		,	"Email_All_Open_Tabs":                                 	46012      	;Share - Send To
	                                		,	"Tracker":                                                      	46207      	;Share - Tracker
	                                		,	"User_Manual":                                              	1277        	;Help - Help
	                                		,	"Help_Center":                                               	558          	;Help - Help
	                                		,	"Command_Line_Help":                                 	32768      	;Help - Help
	                                		,	"Post_Your_Idea":                                           	1279        	;Help - Help
	                                		,	"Check_for_Updates":                                    	46209      	;Help - Product
	                                		,	"Install_Update":                                            	46210      	;Help - Product
	                                		,	"Set_to_Default_Reader":                                	32770      	;Help - Product
	                                		,	"Foxit_Plug-Ins":                                             	1312        	;Help - Product
	                                		,	"About_Foxit_Reader":                                    	57664      	;Help - Product
	                                		,	"Register":                                                      	1280        	;Help - Register
	                                		,	"Open_from_Foxit_Drive":                              	1024        	;Extras - maybe this is not correct!
	                                		,	"Add_to_Foxit_Drive":                                     	1025        	;Extras - maybe this is not correct!
	                                		,	"Delete_from_Foxit_Drive":                             	1026        	;Extras - maybe this is not correct!
	                                		,	"Options":                                                     	243          	;the following one's are to set directly any options
	                                		,	"Use_single-key_accelerators_to_access_tools":  	128          	;Options/General
	                                		,	"Use_fixed_resolution_for_snapshots":             	126          	;Options/General
	                                		,	"Create_links_from_URLs":                              	133          	;Options/General
	                                		,	"Minimize_to_system_tray":                             	138          	;Options/General
	                                		,	"Screen_word-capturing":                               	127          	;Options/General
	                                		,	"Make_Hand_Tool_select_text":                       	129          	;Options/General
	                                		,	"Double-click_to_close_a_tab":                       	91            	;Options/General
	                                		,	"Auto-hide_status_bar":                                  	162          	;Options/General
	                                		,	"Show_scroll_lock_button":                             	89            	;Options/General
	                                		,	"Automatically_expand_notification_message":	1725        	;Options/General - only 1 can be set from these 3
	                                		,	"Dont_automatically_expand_notification":      	1726        	;Options/General - only 1 can be set from these 3
	                                		,	"Dont_show_notification_messages_again":     	1727        	;Options/General - only 1 can be set from these 3
	                                		,	"Collect_data_to_improve_user_experience":   	111          	;Options/General
	                                		,	"Disable_all_features_which_require_internet":	562          	;Options/General
	                                		,	"Show_Start_Page":                                        	160          	;Options/General
	                                		,	"Change_Skin":                                             	46004
	                                		,	"Filter_Options":                                            	46167      	;the following are searchfilter options
	                                		,	"Whole_words_only":                                     	46168      	;searchfilter option
	                                		,	"Case-Sensitive":                                            	46169      	;searchfilter option
	                                		,	"Include_Bookmarks":                                    	46170      	;searchfilter option
	                                		,	"Include_Comments":                                     	46171      	;searchfilter option
	                                		,	"Include_Form_Data":                                    	46172      	;searchfilter option
	                                		,	"Highlight_All_Text":                                       	46173      	;searchfilter option
	                                		,	"Filter_Properties":                                          	46174      	;searchfilter option
	                                		,	"Print":                                                           	57607
	                                		,	"Properties":                                                   	1302        	;opens the PDF file properties dialog
	                                		,	"Mouse_Mode":                                             	1311
	                                		,	"Touch_Mode":                                              	1174
	                                		,	"predifined_Text":                                           	46099
	                                		,	"set_predefined_Text":                                    	46100
	                                		,	"Create_Signature":                                        	26885      	;Signature
	                                		,	"Draw_Signature":                                          	26902      	;Signature
	                                		,	"Import_Signature":                                        	26886      	;Signature
	                                		,	"Paste_Signature":                                          	26884      	;Signature
	                                		,	"Type_Signature":                                           	27005      	;Signature
	                                		,	"Pdf_Sign_Close":                                          	46164}    	;Pdf-Sign

	}

	If FoxitID
		PostMessage, 0x111, % FoxitCommands[command],,, % "ahk_id " FoxitID
	else
		return FoxitCommands[command]
}

FoxitReader_CloseAllPatientPDF() {	                                                                    	;-- schließt nur die FoxitReader Fenster welche im ScanPool Ordner als Datei vorliegen

		MsgBox, 4, Addendum für AlbisOnWindows, Sollen noch alle Fenster im %PDFReader% mit`nnoch geöffneten Patientenbefunden geschlossen werden?
		IfMsgBox, Yes
		{
			WinGet, WinList, List, ahk_class classFoxitReader
			Loop %WinList%
			{
				WinGetTitle, Fxit%A_Index%, % "ahk_id " WinList%A_Index%
				wl:= Trim(SubStr(Fxit%A_Index%, 1, StrLen(Fxit%A_Index%)-15))
				;If LV_Find(hLV1, wl, 0)
				WinClose, % "ahk_id " WinList%A_Index%
			}

		}

}

FoxitReader_DokumentSignieren(docTitle:="", BefundOrdner:="") {	   						;-- FoxitReader - Automatisierung des Signiervorganges - 1Click Automatisierung!

		/*					DESCRIPTION / BESCHREIBUNG

			Im FoxitReader sind Tastaturkürzel für die Funktion 'Signatur platzieren' anzulegen (ich habe im Skript und im FoxitReader dafür Strg+Shift+Alt+F9 eingestellt).
			Alle weiteren Aufrufe nutzen die schon eingestellten Foxit-Kürzel. Wenn man mehrere PDF Dateien nacheinander signieren möchte, muss das Skript in der derzeitigen
			Variante jeweils neu gestartet werden und im FoxitReader sollte die Einstellung auf mehrere Instanzen eingestellt sein. Ich habe es leider nicht hinbekommen das
			TabControl des Foxit Readers zuverlässig auslesen zu können.

			-----------------------------------------------------------------------------
			Library dependancies:
			-----------------------------------------------------------------------------
				1. AddendumFunctions.ahk
				2. FindText().ahk von feiyue
				3. ScanPool_PdfHelper.ahk
			----------------------------------------------------------------------------
			1. und 2. sind zu finden im Addendum include Ordner, 	3. ScanPool-Skriptverzeichnis \lib

		*/

		;letzte Veränderung: 25.08.2018
		;25.08.2018: 	Kennwortverschlüsselung rund erneuert, Ereignisbehandlung :Signaturfenster schließT sich nicht - programmiert, Routine prüft das im FoxitReader geöffnete Dokument
		;						ob es schon signiert wurde

		static Ort, Grund, SignierenAls, DokumentNachderSignierungSperren, Darstellungstyp, Autokennwort, Genu

		;{ 01. Einstellungen für das Fenster 'Dokument signieren'

			If !Ort
					IniRead, Ort													 , %AddendumDir%\Addendum.ini, ScanPool, Ort
			If !Grund
					IniRead, Grund												 , %AddendumDir%\Addendum.ini, ScanPool, Grund
			If !SignierenAls
					IniRead, SignierenAls										 , %AddendumDir%\Addendum.ini, ScanPool, SignierenAls
			If !DokumentNachDerSignierungSperren
					IniRead, DokumentNachDerSignierungSperren, %AddendumDir%\Addendum.ini, ScanPool, DokumentNachDerSignierungSperren
			If !Darstellungstyp
					IniRead, Darstellungstyp				     				 , %AddendumDir%\Addendum.ini, ScanPool, Darstellungstyp
			If !PasswortOn
					IniRead, PasswordOn    					    			 , %AddendumDir%\Addendum.ini, ScanPool, Passwort_benutzen, 0

			;noch nicht implementiert aufgrund Sicherheitsbedenken (AHK ist zu unsicher was das Abspeichern von Passwörten betrifft)
			IniRead, Autokennwort, %AddendumDir%\Addendum.ini, ScanPool, Autokennwort, 0

			If !Ort or !Grund or !SignierenAls or !DokumentNachDerSignierungSperren or !Darstellungstyp {
						FoxitReader_SignatureProcess()
						return 0
			}

		;}
				P+= (PAdd)
		;{ 02. Auslesen der Fensterposition des Albis Fenster, erstellen zweier Objecte
			AWI			 	:= Object()				                     	;AlbisWindowInfo = AWI.WindowX
			CWI			 	:= Object()				                     	;ChildWindowInfo or WindowOfInterest
			AlbisWinID	:= AlbisWinID()
			AWI			 	:= GetWindowInfo(AlbisWinID)
			BO			 	:= GetWindowInfo(hBO)
			TTx			 	:= BO.WindowX+10
			TTy			 	:= BO.WindowY+38
			PraxTT("Suche aktuelles FoxitReader Fenster!", "12 2")
		;}
				P+= (PAdd)
		;{ 03. Ermitteln der aktuellen ID des obersten FoxitFenster mit Ermittlung des entsprechenden Fenstertitel oder wenn ein Titel übermittelt wird, wird dieses Fenster aufgerufen und auf eine vorhandene Signatur geprüft

			If (docTitle = "")
        	{
					WinGetTitle, docTitle, % "ahk_id " FoxitID:=WinExist("ahk_class classFoxitReader")
					docTitle				:= Trim( SubStr(docTitle, 1, StrLen(docTitle)-15) )
					docTitleFullPath	:= BefundOrdner "\" docTitle
			}
			else
			{
					FoxitID:=WinExist(docTitle)
					If !(FoxitID)
					{
								FoxitWindow:
								msgTextPre	:=  "Das FoxitReader Fenster mit dem Titel: `n'"
								msgTextPost	:= "`nkonnte nicht identifiziert werden.`n`nBitte holen Sie das Fenster mit diesem Titel nach vorne. Klicken Sie DANN`n"
								msgTextPost	.= "erst auf Ok! und innerhalb der nächsten 3 Sekunden wieder`nin das FoxitReader Fenster `(aktivieren!`)"
								MsgBox,, Addendum für AlbisOnWindows - ScanPool - Info, % msgTextPre docTitle msgTextPost
								Loop, 30
								{
											ActwID:= WinExist("A")
											WinGetTitle, wt, % "ahk_id " ActwID
											If Instr(wt, docTitle) {
												 FoxitID:= ActwID
												 PraxTT("Danke! Das richtige FoxitReader Fenster konnte jetzt erfasst werden.`nDieser Dialog schließt sich automatisch in 2 Sekunden`.", "2 2")
												 break
											}
											sleep 100
											If (A_Index>29)
												goto FoxitWindow
								}
						}
				docTitleFullPath:= BefundOrdner "\" docTitle
			}

			;Ermitteln des Childhandles (Docwindow) im Foxitreader (Bildschirmposition des Dokuments)
				hDocWnd:= FindChildWindow({ID:FoxitID}, {Class:"FoxitDocWnd1"}, "On")
			;Schließen aller PopUp-Fenster des FoxitReaders damit diese den Vorgang nicht behindern können
				AlbisCloseLastActivePopups(FoxitID)						;eine universelle Funktion - schließt alle noch offenen PopUpFenster - sollte auch bei anderen Programmen gehen
				sleep, 200
			;FoxitReader mit einem Trick dazu überreden uns mitzuteilen das das Dokument schon signiert wurde
				PraxTT("Teste ob das Dokument signiert ist.`nVersuche deshalb das Fenster 'Dateianhang' zu öffnen.`nÖffnet es sich nicht, ist das Dokument schon signiert.", "12 2")
				;result:= FrCmd("NewAttachment", "AfxWnd100su4", FoxitID)
				result:= FoxitInvoke("NewAttachment", FoxitID)
				WinWait, Dateianhang ahk_class #32770,, 3
				If ErrorLevel
            	{
						PraxTT("Dieses Dokument wurde schon signiert.`nDamit kann als nächstes der Import in die`nPatientenakte gestartet werden..", "5 2")
						sleep, 2000
						P+= 9*(PAdd/5/11)							;überspringe 9 Schritte
						goto CloseFoxitReader
				}
			;Schließen des 'Dateianhang' Dialoges
				WhatFoxitID:= WinExist("Dateianhang ahk_class #32770")
				If (FoxitID = GetParent(WhatFoxitID)) {
						PraxTT("Fahre fort mit dem Signieren`ndes Dokumentes.", "3 2")
				}
				VerifiedClick("Button2", "", "", WhatFoxitID)

		;}
				P+= (PAdd)
		;{ 04. Aktivieren des Fenster und Vorbereitung der Signierung
			PraxTT("Aktivieren und Vorbereitung der Signierung!", "12 2")

			;eigentlich für Albis programmiert sollte diese Funktion auch mit dem FoxitReader funktionieren
				;AlbisIsBlocked(FoxitID, 2)
					WinMaximize	, % "ahk_id " FoxitID
					WinActivate		, % "ahk_id " FoxitID
					WinWaitActive	, % "ahk_id " FoxitID,, 4

			;AfxWnd100su4 ist das ClassNN für das PDFFrame - hier wird das Handle benötigt (gebraucht wird dieses Hwnd nicht für das Skript, lasse die Zeile hier stehen, für eventuell späteren Gebrauch)
				;ControlGet, hPdfFrame, Hwnd,, AfxWnd100su4 , ahk_id %FoxitID%
				ControlClick,, % "ahk_id " hDocWnd

		;}
				P+= (PAdd)
		;{ 05. ganze Seite zeigen, zur ersten Seite blättern und Aufruf von PlaceSignature. das alles per SendMessage. Kein Senden eines Tastaturkürzel notwendig!

				PraxTT("Text selektieren`nSeite einpassen`nErste Seite", "12 2")

			;FoxitReader vorbereiten für das Platzieren der Signatur
				result:= FoxitInvoke("Select_Text", FoxitID)
				sleep, 150
				result:= FoxitInvoke("Fit_Page"	, FoxitID)
				sleep, 150
				result:= FoxitInvoke("FirstPage"	, FoxitID)
				sleep, 150

				PraxTT("sende Befehl: 'Signatur platzieren'", "12 2")
				WinGetPos	, Fwx		, Fwy		, Fww, Fwh, "ahk_id " hDocWnd
				MouseMove	, % Fwx	, % Fwy
				MouseClick

				result:= FoxitInvoke("Place_Signature", FoxitID)
				sleep, 150
		;}
				P+= (PAdd)
		;{ 06. sucht nach der linken oberen Ecke der Pdf Seite der 1.Seite, um dort im Anschluss die Signatur zu erstellen

			PraxTT(" suche nach dem Signierbereich des Dokumentes.", "12 2")
				tryCount:=0
			;!NICHT ÄNDERN! dieser String wird von feiyus FindText() Funktion benötigt, er entspricht der linken oberen Ecke des PDF-Frame (AfxWnd100su4)
				TopLeft:="|<>*210$71.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzw0000003zzzzs0000007zzzzk000000DzzzzU000000Tzzzz0000000zzzzy0000001zzzzw0000003zzzzs0000007zzzzk000000DzzzzU000000Tzzzz0000000zzzzy0000001zzzzw0000003zzzzs0000007zzzzk000000DzzzzU000000Tzzzz0000000zzzzy0000001"

			FindSignatureArea:
			basetolerance:=0
			;sucht hier im Prinzip nach einem Bild (entspricht der linken oberen Ecke des PDF Preview Bereiches) und nicht nach einem Text.
				if (ok:=FindText(Fwx, Fwy, Fww, Fwh, basetolerance, 0, TopLeft)) {
					X:=ok.1.1, Y:=ok.1.2, W:=ok.1.3, H:=ok.1.4, Comment:=ok.1.5, X+=W//2, Y+=H//2

					MouseClickDrag, Left, % x,  %  y, % x+100, % y+50, 0
				} else {
							sleep, 100
						tryCount++
						If (tryCount < 20) {
								basetolerance += 0.1
								goto FindSignatureArea
						}
				}
		;}
				P+= (PAdd)
		;{ 07. Wartet auf das Signierfenster und ermittelt das Handle und aktiviert das Fenster dann

			PraxTT("Warte auf das Signierfenster", "13 2")
			checkDocSigWin:
			WinWait, Dokument signieren ahk_class #32770,, 15
			If !ErrorLevel
			{
					hDokSig:= WinExist("Dokument signieren ahk_class #32770")
					WinActivate		, % "ahk_id " hDokSig
					WinWaitActive	, % "ahk_id " hDokSig,, 5
			}
				else
			{
					MsgBox, 4144, Addendum für AlbisOnWindows - ScanPool - Info, Scheinbar hat das Platzieren der Signatur nicht funktioniert`nBitte führen Sie das Platzieren manuell durch und drücken im Anschluß auf Ok.
					goto checkDocSigWin
			}

		;}
				P+= (PAdd)
		;{ 08. verschiebt das Signierfenster in Richtung des Albis Fenster - mittelt die beiden in der Position aus
			PraxTT("Signierfenster gefunden.`nVerschiebe es in Richtung des Albisfensters.", "13 0")
			WinMove		, ahk_id %hPraxTT%,, % BOx, % BOy
			WinActivate	, ahk_id %FoxitID%
			CWI:= GetWindowInfo(hDokSig)
			WinMove		, % "ahk_id " hDokSig,, % AWI.WindowX + (AWI.WindowW-CWI.WindowW)//2, % AWI.WindowY + (AWI.WindowH-CWI.WindowH)//2
			sleep, 200
		;}
				P+= (PAdd)
		;{ 09. Felder im Signierfenster auf die in der INI festgelegten Werte einstellen
				PraxTT(" Fülle des Signierfenster mit`nWerten aus der Addendum.ini.", "13 2")

				WinActivate, % "ahk_id " hDokSig
				WinWaitActive, % "ahk_id " hDokSig,, 5

			;Signieren als:
				ControlFocus	, ComboBox1													, % "ahk_id " hDokSig
				Control				, ChooseString	, %SignierenAls%, ComboBox1	, % "ahk_id " hDokSig
				sleep, 100
			;Ort:
				ControlFocus	, Edit2																, % "ahk_id " hDokSig
				ControlSetText	, Edit2,																, % "ahk_id " hDokSig
				ControlSetText	, Edit2, %Ort%													, % "ahk_id " hDokSig
				ControlSend		, Edit2, {Tab}													, % "ahk_id " hDokSig
				sleep, 100
			;Grund:
				ControlFocus	, Edit3																, % "ahk_id " hDokSig
				ControlSend		, Edit3,{Delete}													, % "ahk_id " hDokSig
				ControlSetText	, Edit3, %Grund%												, % "ahk_id " hDokSig
				ControlSend		, Edit3, {Tab}													, % "ahk_id " hDokSig
				sleep, 100
			;Dokument nach der Signierung sperren
				If DokumentNachDerSignierungSperren
						VerifiedCheck("Button4","","", hDokSig)
				sleep, 100
			;Darstellungstyp:
				ControlFocus , ComboBox4														, % "ahk_id " hDokSig
				Control			, ChooseString	, %Darstellungstyp%, ComboBox4	, % "ahk_id " hDokSig
				sleep, 100

			;~ if !Genu
			;~ {
				;~ InputBox, Genu, Kennwort benötigt, Geben Sie Ihr Kennwort für das Signieren ein, HIDE, 300, 140
				;~ Genu:= Encode(Genu, "ahk_class OptoAppClass ahk_id " . hBO, 1 )
				;~ varum()
			;~ }
					sleep, 100

			;~ ControlSetText, Edit1, % Decode(Genu, "ahk_class OptoAppClass ahk_id " . hBO, 1), % "ahk_id " hDokSig
			;~ varum()

		;}
				P+= (PAdd)
		;{ 10. Signierfenster schließen und die folgenden Speicherdialage ebenso automatisch abschließen
				PraxTT("Schließe das Signierfenster und`ndie folgenden zugehörigen Dialoge", "13 2")
				SignierFensterCheck:
				VerifiedClick("Button5","","", hDokSig)
				sleep, 200
				If WinExist("ahk_id " . hDokSig) {
						MsgBox,,Addendum für AlbisOnWindows - ScanPool - Info, Das Eintragen des Kennwortes hat nicht funktioniert.`nBitte tragen Sie es bitte manuell ein!`nDrücken Sie danach bitte erst auf Ok.
						goto SignierFensterCheck
			}
			;FoxitReader braucht einen Click ins Fenster um den Speichern unter Dialog zuöffnen
				WinActivate, ahk_id %FoxitID%
				WinGetPos, Fwx, Fwy,,, ahk_id %FoxitID%
				ControlGetPos, Fcx, Fcy,,,, ahk_id %hPdfFrame%
				MouseClick, Left, % (Fwx + Fcx + 100) , % (Fwy + Fcy + 50)
					P+= (PAdd)
				PraxTT("Warte und bestätige den Dialog`nSpeichern unter`.`.`.", "13 2")
				WinActivate, ahk_id %FoxitID%
				While !WinExist("Speichern unter ahk_class #32770 ahk_exe FoxitReader.exe")
				{
						status:= CheckWindowStatus(FoxitID, 100)
						If (status > 0)
							MsgBox, 4144, Addendum für AlbisOnWindows - ScanPool - Info, Das FoxitReader Fenster hat Probleme`nden Speichern unter... Dialog zu öffnen.`nDrücken Sie OK sobald der FoxitReader bereit ist.
						If go=1
							break
						sleep, 300
				}

				While WinExist("Speichern unter ahk_class #32770 ahk_exe FoxitReader.exe") {
						VerifiedClick("Button3", "Speichern unter ahk_class #32770 ahk_exe FoxitReader.exe","", "")
						sleep, 100
						If (A_Index>15) {
							MsgBox,, Addendum für AlbisOnWindows, Scheinbar bekomme ich das `"Speichern unter`" Fenster`n vom FoxitReader bekomme ich nicht geschlossen.`nBitte schließen Sie es manuell und drücken dann auf 'OK'
							break
						}
				}
		;}
				P+= (PAdd)
		;{ 12. FoxitReader Fenster schließen
			CloseFoxitReader:
			PraxTT("Schließe den FoxitReader mit dem Fenstertitel`n" . DocTitel . "`.", "13 0")
				sleep, 100
			result:= Win32_SendMessage(FoxitID)
				If !(result) {
					FoxitID:=FindWindow(docTitle, "", "", "", "", "on", "on")
					WinKill, ahk_id %FoxitID%
				}

		;}
				P+= (PAdd)
			PraxTT("Der Signiervorgang für dieses Dokument ist beendet!", "2 2")
			sleep, 200

		return 1
	}

FoxitReader_SignaturSetzen(ReaderID, PDFReaderWinClass) {									;-- ruft Signatur setzen auf und zeichnet eine Signatur in die linke obere Ecke des Dokumentes

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; Ermitteln des Childhandles (Docwindow) im Foxitreader (Bildschirmposition des Dokuments)
		;----------------------------------------------------------------------------------------------------------------------------------------------
			hDocWnd		:= FindChildWindow({ID:ReaderID}, {Class:"FoxitDocWnd1"}, "On")
			hDocParent	:= FindChildWindow({ID:ReaderID}, {Class:"AfxFrameOrView140su1"}, "On")

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; FoxitReader vorbereiten für das Platzieren der Signatur
		;----------------------------------------------------------------------------------------------------------------------------------------------
			WinActivate		, % "ahk_id " ReaderID
			WinWaitActive	, % "ahk_id " ReaderID,,2

			result2:= FoxitInvoke("SinglePage", ReaderID)
			result2:= FoxitInvoke("SinglePage", ReaderID)
			Sleep, 250
			result3:= FoxitInvoke("Fit_Page", ReaderID)
			sleep, 250
			result4:= FoxitInvoke("firstPage", ReaderID)
			sleep, 150

			PraxTT("sende Befehl: 'Signatur platzieren'", "12 2")
			WinGetPos	, Fwx	, Fwy	, Fww, Fwh , % "ahk_id " hDocWnd
			MouseMove	, % Fwx + 50, % Fwy + 50, 0
			MouseClick	, Left

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; Signatur setzen Menupunkt aufrufen
		;----------------------------------------------------------------------------------------------------------------------------------------------
			result5:= FoxitInvoke("Place_Signature", ReaderID)
			sleep, 250

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; sucht nach der linken oberen Ecke der Pdf Seite der 1.Seite, um dort im Anschluss die Signatur zu erstellen
		;----------------------------------------------------------------------------------------------------------------------------------------------
			PraxTT(" suche nach dem Signierbereich des Dokumentes.", "12 2")
				tryCount:=0, basetolerance := 0
			;!NICHT ÄNDERN! dieser String wird für 'feiyus' FindText() Funktion benötigt, er entspricht der linken oberen Ecke des PDF-Frame (AfxWnd100su4) ;{
				TopLeft:="|<>*210$71.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzw0000003zzzzs0000007zzzzk000000DzzzzU000000Tzzzz0000000zzzzy0000001zzzzw0000003zzzzs0000007zzzzk000000DzzzzU000000Tzzzz0000000zzzzy0000001zzzzw0000003zzzzs0000007zzzzk000000DzzzzU000000Tzzzz0000000zzzzy0000001"
			;}
			FindSignatureRange:
			;sucht hier im Prinzip nach einem Bild (entspricht der linken oberen Ecke des PDF Preview Bereiches) und nicht nach einem Text.
				if (ok:=FindText(Fwx, Fwy, Fww, Fwh, basetolerance, 0, TopLeft))
				{
					PraxTT("Signierbereich gefunden.", "4 2")
					X:=ok.1.1, Y:=ok.1.2, W:=ok.1.3, H:=ok.1.4, Comment:=ok.1.5, X+=W//2, Y+=H//2
					MouseClickDrag, Left, % x, % y, % x + 50, % y + 50, 0
				}
				else
				{
						sleep, 100
						tryCount++, basetolerance += 0.1
						If (tryCount < 20)
							goto FindSignatureRange
						else
							return 0
				}

return 1
}

FoxitReader_SignatureProcess() {                                                                        	;-- ## unfertige ## Funktion wird nur einmalig ausgeführt (ScanPool Erststart)

			mtext=
			(LTrim
					WICHTIG!       WICHTIG!        WICHTIG!		WICHTIG!        WICHTIG!         WICHTIG!

			Sie haben in den Einstellungen noch keine oder unvollständig Ihre Daten für das FoxitReader Fenster
			'Dokument signieren' hinterlegt. Dieses Skript ist in der Lage diese Daten aus dem Fenster auszulesen!

			Eine Automatisierung der Signierung benötigt diese notwendigen Angaben. Sie werden jetzt nicht nach ihrem
			im FoxitReader zum Signieren hinterlegten Paßwort gefragt! Das Skript wird das Paßwort auch später nicht
			auf der Festplatte speichern, solange es keine sichere Möglichkeit der Verschlüsselung gibt.

			Starten Sie jetzt bitte den FoxitReader. Die Daten für dieses Fenster erhalten Sie nur, wenn Sie eine bisher
			unsignierte PDF-Datei öffnen und über das Menu Schützen den Dialog 'Signieren und Zertifizieren' auswählen.
			Wenn Sie auf ja drücken öffnet Ihnen das Skript eine sogenannte 'Lorem ipsum' pdf Datei damit Sie nicht
			nach einer PDF Datei suchen müssen. Lesen Sie bitte aber zunächst weiter!

			Bevor Sie jetzt allerdings fortfahren sollten Sie überprüfen, ob Sie schon eine Signatur angelegt haben.
			Diese läßt sich im Einstellungsfenster unter Auswahl im linken Bereich Signatur und dann im rechten Fenster-
			bereich über 'Neu...' oder 'Bearbeiten...' entsprechend erstellen oder bearbeiten. Im Anschluss starten Sie den
			Dialog 'Signieren und Zertifizieren'. Es öffnet sich das Fenster 'Dokument signieren'. Stellen Sie in diesem
			Dialog ihre gewünschte Signatur unter 'Signieren als:' ein. Schreiben Sie Ihren gewünschten Text in alle Felder,
			aber lassen Sie das Feld 'Kennwort:' leer. Bitte überzeugen Sie sich das in keinem anderen FoxitReader Prozess
			das Fenster 'Dokument signieren' ebenfalls geöffnet ist. Wenn kein weiteres Fenster geöffnet ist drücken Sie
			jetzt auf 'Ja' in diesem Hinweisfenster und das Skript übernimmt Ihre Einstellungen und speichert diese in die
			Addendum.ini Datei im Addendum Hauptordner.
			Bitte achten Sie auch darauf die Checkbox für 'Dokument nach der Signierung sperren' zu setzen oder eben nicht.

			Das Skript wird in Zukunft mit diesen Einstellungen arbeiten und Sie brauchen diese nie wieder zu kontrollieren
			oder neu setzen. Wollen Sie eine Änderung ihrer Einträge in der Addendum.ini durchführen, können Sie diese
			manuell mit einem Texteditor wie dem Microsoft Editor oder über die ScanPool Gui über den Button Einstellungen
			neu einlesen.
			)

			MsgBox, 4, Addendum für AlbisOnWindows - Dokument Signieren, %mtext%
			If MsgBox, Yes
			{
				hDokSig:= WinExist("Dokument signieren ahk_class #32770")
				If !hDokSig
						hDokSig:= WinExist("Sign document ahk_class #32770")
				;hier Programmierung einer ordentlich geführten Eingabeüberprüfung z.B. wenn User vergessen hat bestimmte Felder auszufüllen, sollte es eine Rückfrage geben
				;Prinzip: "MAXIMALE AUTOMATISIERUNG" oder "ONE CLICK SOFTWARE" siehe mtext
			}

return
}

FoxitReader_ConfirmSaveAs(hHook1) {																	;-- zum Schliessen des Dialogfenster "Speichern unter bestätigen"

		hHook	:= hHook1
		dialog	:= Object()
		WinGet, cNames	, ControlList			, % "ahk_id " hHook
		WinGet, chwnd		, ControlListHwnd	, % "ahk_id " hHook
		dialog:= KeyValueObjectFromLists(cNames, cHwnd)
		ControlClick,, % "ahk_id " dialog["Button1"]
		hHook 	:= 0
		PraxTT("", "off 2")

return
}

FoxitReader_CloseSaveAs(hHook2) {																		;-- zum Schliessen des Dialogfenster "Speichern unter" (z.B. als Hookhandler)

	; Funktion liest vor dem Schließen den Pfad und Dateinamen aus und gibt diesen zurück

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; a) Initialisierung
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		If foxitHook
			 return
		foxitHook  	:= GetHex(hHook2)
		HookClass	:= GetClassName(foxitHook)
		while !InStr(HookClass, "#32770")	{

				foxitHook  	:= GetHex(GetParent(foxitHook))
				HookClass	:= GetClassName(foxitHook)

				If InStr(HookClass, "#32770")
					break

				Sleep, 100
				If (A_Index > 20)
					return ; Abbruch - Daten konnten nicht ermittelt werden

		}

		PHook    		:= GetHex(GetParent(foxitHook))
		PTitle	       	:= WinGetTitle(PHook)
		PClass   		:= GetClassName(PHook)
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; b) PDF-Dateinamen aus dem FoxitReader Fenstertitel generieren
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		fname:= RegExReplace(PTitle, "i)(?<=\.pdf).*", "")	;Dateiname aus dem FoxitReader Fenster
		while InStr(fname, "Speichern unter") || (fname = "" )
		{
				If InStr(HookClass, "#32770")
					If fname:= RegExReplace(PTitle , "i)(?<=\.pdf).*", "")
						break

				If A_Index > 10
					 break
				Sleep, 100
		}
		headway := "1. neues PdfReader Fenster wurde erkannt:`nPTitle: " PTitle "`nPClass: " PClass "`nHClass: " HookClass
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; c) Auslesen von Dateiordner und PDF-Dateinamen aus den Edit-Controls auslesen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		;efname := Controls("Edit1", "GetText", foxitHook)
		;				 Controls("ToolBarWindow324", "Click use ControlClick Left", foxitHook)
		; If !fDir	:= Controls("Edit2", "GetText", foxitHook) {
			;			 Controls("ToolBarWindow324", "Click use MouseClick Left", foxitHook)
		;	fDir	:= Controls("Edit2", "GetText", foxitHook)
		; }
		; headway .= "`n2. Informationen wurde ausgelesen`n- " fDir "\" fname " -"
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; d) Vergleich Dateinamen aus Fenstertitel und Edit1, bei Unterschied wird Edit1 mit Fensterdateinamen ersetzt
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		;If !InStr(efname, fname)
		;		VerifiedSetText("", fname, "ahk_id " Controls("Edit1"	, "hwnd", foxitHook) , 200)
		;headway .= "`n3. Einstellungen wurden verglichen. "
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; e) Kontrolle auf den korrekten Speicherordner (unterscheidet anhand des Namensmuster!)
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		;If !RegExMatch(fname, "\d\dA\w\.pdf") and !InStr(fDir, "albiswin")
		; {
		;			propWasSet:=1
		;			If !Controls("Edit2", "GetText", foxitHook)			;falls Edit2 nicht mehr freigeschaltet ist würde ein leerer String zurückgegeben
		;				 Controls("ToolBarWindow324", "Click use MouseClick Left"	, foxitHook)
		;			Controls("Edit2", "SetText " BefundOrdner	, foxitHook)
		;			Controls("Edit2", "Send {Enter}"				, foxitHook)
		; }
		;headway .= "`n4. " ( propWasSet = 1 ? "Pfade mussten angepasst werden." : "alle Speicherpfade sind korrekt. " )
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; f) Speichern Button wird erst gesucht (kann Button2 oder Button3 sein), anschließend wird er gedrückt
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		ControlClass:= Controls("", "ControlFind, &Speichern, Button", foxitHook)				;Deutsch
		If !ControlClass
				ControlClass:= Controls("", "ControlFind, &Save, Button", foxitHook)				;Englisch
		If ControlClass {
				res:= Controls(ControlClass, "SetFocus", foxitHook)
				res:= Controls(ControlClass, "Click use ControlClick Left", foxitHook)            		;ControlClass sollte jetzt Button2 oder Button3 als Text enthalten
				headway .= "`n5. Schaltfäche - Speichern (" ControlClass ") - wurde gedrückt. "
				WinWait, Speichern unter bestätigen ahk_class #32770,, 5
		} else {
				PraxTT("Die 'Speichern' Schaltfläche konnte`nnicht identifiziert werden!`nDer 'Speichern unter' Dialog muss`nvon Hand geschlossen werden.`n`nScanPool wartet bis der Dialog`ngeschlossen wurde!", "0 2")
		}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; f) Löschen des Controls-Objekts, automat. Schließen eines gerade signierten Pdf-Befundes
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		headway .= "`n6. Warte auf das Schließen des Dialoges`nim Anschluß wird das Dokumentfenster geschlossen."
		;PraxTT(headway, "8 2")

	;geöffnetes Pdf-Dokument nach dem Signieren schliessen
		while, WinExist("Speichern unter ahk_class #32770")
		{
				if (hCSaveAs:= WinExist("Speichern unter bestätigen ahk_class #32770"))
							FoxitReader_ConfirmSaveAs(hCSaveAs)

				Sleep, 40
				If A_Index > 60
						break
		}

		Sleep, 1000

	; dies ermöglicht das Schließen des FoxitReader-Fensters noch abbrechen zu können im Ausnahmefall
		If SPData.AutoCloseReader
				PdfReader_Close(PHook, PDFReaderWinClass)   ;-> ein CloseStack entwerfen??

	; alles zurücksetzen
		fError    	:= Controls("", "Reset", foxitHook)
		foxitHook 	:= 0
	;}

		;PdfData.AlbisPdfPath := fDir "\" fname

return
}

FoxitReader_ExceptionDialog(hHook3) {                                                                	;-- zum Schließen eines selten vorkommenden Dialogfensters
	; verhindert das der FoxitReader geschlossen wird, da jetzt ein manueller Eingriff notwendig ist.
		SPData.AutoCloseReader:= 0
return VerifiedClick("Button1", "", "", hHook3)
}

FoxitReader_SignDoc(hDokSig, FoxitTitle, FoxitText:="") {					    					;-- Winhook-Handler zum Bearbeiten des Dokument signieren Dialoges

	PraxTT("Das Fenster 'Dokument signieren' wird gerade bearbeitet....!'", "0 2")

	If !Running
		Running:= "automatische Signatur"

	;{ Auslesen der Fensterposition des Albis Fenster, erstellen zweier Objecte

			AWI				:= Object()					;AlbisWindowInfo = AWI.WindowX
			CWI				:= Object()					;ChildWindowInfo or WindowOfInterest
			AlbisWinID	:= AlbisWinID()
			AWI				:= GetWindowSpot(AlbisWinID)
			CWI          	:= GetWindowSpot(hDokSig)

		  ; Verschieben des Dokument signieren Fensters, es wird mittig über Albis abgelegt
			WinMove, % "ahk_id " hDokSig,, % AWI.X + (AWI.W-CWI.W)//2, % AWI.Y + (AWI.H-CWI.H)//2
			sleep, 200
	;}

	;{ Felder im Signierfenster auf die in der INI festgelegten Werte einstellen

				WinActivate		, % "ahk_id " hDokSig
				WinWaitActive	, % "ahk_id " hDokSig,, 5

			;Signieren als -------------------------------------------------------------------------------------------------------
				ControlFocus                                 	 , ComboBox1, % "ahk_id " hDokSig
				Control, ChooseString, % Addendum.SignierenAls , ComboBox1, % "ahk_id " hDokSig
				sleep, 100

			;Darstellungstyp ----------------------------------------------------------------------------------------------------
				ControlFocus                                 	 , ComboBox4, % "ahk_id " hDokSig

				;prüft das Feld Signaturvorschau auf die in der ini hinterlegte Signatur
				ControlGet, entryNr, FindString	, % Addendum.Darstellungstyp, ComboBox4, % "ahk_id " hDokSig
				If !entryNr
						MsgBox, 4144, Addendum für AlbisOnWindows - ScanPool, % Hinweis5
				else
						Control, ChooseString    	, % Addendum.Darstellungstyp, ComboBox4, % "ahk_id " hDokSig
				Sleep, 100

			;Ort: -----------------------------------------------------------------------------------------------------------------
				ControlFocus 	, Edit2			                                    		, % "ahk_id " hDokSig
				ControlSetText	, Edit2,  % JEE_StrUtf8BytesToText(Addendum.Ort)    	, % "ahk_id " hDokSig
				ControlSend		, Edit2,  {Tab}	                                    	, % "ahk_id " hDokSig
				sleep, 100

			;Grund: -------------------------------------------------------------------------------------------------------------
				ControlFocus 	, Edit3		                                    			, % "ahk_id " hDokSig
				ControlSetText	, Edit3, % JEE_StrUtf8BytesToText(Addendum.Grund)	, % "ahk_id " hDokSig
				ControlSend		, Edit3, {Tab}	                                    	, % "ahk_id " hDokSig
				sleep, 100

			;nach der Signierung sperren: -----------------------------------------------------------------------------------
				If Addendum.DokumentNachDerSignierungSperren
						VerifiedCheck("Button4","","", hDokSig)
				sleep, 100


			;~ if (!Addendum.Genu && PasswordOn)
			;~ {
				;~ InputBox, Genu, Kennwort benötigt, Geben Sie Ihr Kennwort für das Signieren ein, HIDE, 300, 140
				;~ Genu:= Encode(Genu, "ahk_class OptoAppClass ahk_id " . hBO, 1 )
				;~ varum()
			;~ }
					;~ sleep, 100

			;~ ControlSetText, Edit1, % Decode(Genu, "ahk_class OptoAppClass ahk_id " . hBO, 1), % "ahk_id " hDokSig
			;~ varum()

	;}

	;{ Signierfenster schließen und die folgenden Speicherdialage ebenso automatisch abschließen

		;Signaturzähler erhöhen
			SPData.Signatures ++
			IniWrite, % SPData.Signatures, % AddendumDir "\Addendum.ini", % "ScanPool", % "Signatures"

		;jetzt das Signaturfenster schliessen
			while, WinExist("ahk_id " . hDokSig)
			{
				VerifiedClick("Button5","","", hDokSig)
						sleep, 30
				If A_Index>20
					MsgBox,,Addendum für AlbisOnWindows - ScanPool - Info, Das Eintragen des Kennwortes hat nicht funktioniert.`nBitte tragen Sie es bitte manuell ein!`nDrücken Sie danach bitte erst auf Ok.
			}

	;}

		PraxTT("'", "off 2")

		If InStr(Running, "automatische Signatur")
		{
				Running := ""
				Sleep 500
		}

return
}

FoxitReader_GetPDFPath() {                                                                                  	;-- den aktuellen Dokumentenpfad im 'Speichern unter' Dialog auslesen

	foxitSaveAs := "Speichern unter ahk_class #32770 ahk_exe FoxitReader.exe"
	If WinExist(foxitSaveAs) {

		WinGetText, allText, % foxitSaveAs
		RegExMatch(allText, "(?<Name>[\w+\s_\-\,]+\.pdf)\n.*Adresse\:\s*(?<Path>[A-Z]\:[\\\w\s_\-]+)\n", File)
		return FilePath "\" FileName

	}

return ""
}

PdfReader_Close(ReaderTitleOrID, PDFReaderWinClass, CloseTabOnly:=0) {			;-- schließen eines FoxitReader-Tabs oder Beenden einer PDF-Reader Instanz

		headway := "Schließe das PdfReader-Fenster`nmit dem Titel:`n" ReaderTitleOrID
		PraxTT(headway, "0 2")

		If RegExMatch(ReaderTitleOrID, "^0x[\w]+$") or RegExMatch(ReaderTitleOrID, "^\d+$")
			ReaderID:= ReaderTitleOrID
		else
			WinGet, ReaderID, ID, % ReaderTitleOrID

		If CloseTab Only and InStr(PDFReaderWinClass, "FoxitReader")
		{
				WinActivate	 	, % "ahk_id " ReaderID
				WinWaitActive	, % "ahk_id " ReaderID,, 5
				FoxitInvoke("Close", ReaderID)
				Sleep, 120
				return
		}

		WinClose, % "ahk_id " ReaderID
		WinWaitClose, % "ahk_id " ReaderID,, 2
		If ErrorLevel {
			SendMessage 0x112, 0xF060,,, % "ahk_id " ReaderID 			; WMSysCommand + SC_Close
			WinWaitClose, % "ahk_id " ReaderID,, 2
		If ErrorLevel {
			SendMessage 0x10, 0,,, % "ahk_id " ReaderID               			; WM_Close
			WinWaitClose, % "ahk_id " ReaderID,, 2
		If ErrorLevel {
			SendMessage 0x2, 0,,, % "ahk_id " ReaderID 	                		; WM_Destroy
			WinWaitClose, % "ahk_id " ReaderID,, 2
		If ErrorLevel
			Process, Close, % "ahk_id " ReaderID
		}
		}
		}

	;Infofenster ausschalten
		PraxTT("", "off 2")

	;ein PdfReader-Fenster nach vorne holen, wenn noch einer da ist
		If WinExist("ahk_class " PDFReaderWinClass)
				WinActivate, % "ahk_class " PDFReaderWinClass

return ErrorLevel
}

JEE_StrUtf8BytesToText(vUtf8) {
	if A_IsUnicode 	{
		VarSetCapacity(vUtf8X, StrPut(vUtf8, "CP0"))
		StrPut(vUtf8, &vUtf8X, "CP0")
		return StrGet(&vUtf8X, "UTF-8")
	} else
		return StrGet(&vUtf8, "UTF-8")
}

;----------------------------------------------------------------------------------------------------------------------------------------------
; SUMATRA READER
;----------------------------------------------------------------------------------------------------------------------------------------------

SumatraInvoke(command, SumatraID="") {                                                           	;-- wm_command wrapper for Sumatra PDF Version: 3.1

	/* DESCRIPTION of FUNCTION:  SumatraInvoke() by Ixiko (version 12.07.2020)

		---------------------------------------------------------------------------------------------------
										a WM_command wrapper for Sumatra Pdf-Reader V* by Ixiko
																		...........................................................
													 Remark: maybe not all commands are listed at now!
		---------------------------------------------------------------------------------------------------
		        by use  of a valid SumatraID, this function will post your command to Sumatra
			                                             otherwise this function returns the command code
																		...........................................................
			Remark: You have to control the success of the postmessage command yourself!
		---------------------------------------------------------------------------------------------------

		---------------------------------------------------------------------------------------------------
		      EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES

		SumatraInvoke("Show_FullPage")        SumatraInvoke("Place_Signature", SumatraID)
		.................................................       ...................................................................
		this one only returns the Sumatra              sends the command "Place_Signature" to
        command-code                                             your specified Sumatra process using
																 parameter 2 (SumatraID) as window handle.
															 		          command-code will be returned too
		---------------------------------------------------------------------------------------------------

	*/

	static SumatraCommands
	If !IsObject(SumatraCommands) {

		SumatraCommands := { 	"Open":                                 	400  	; File
											,	"Close":                                 	401  	; File
											,	"SaveAs":                               	402  	; File
											,	"Rename":                              	580  	; File
											,	"Print":                                   	403  	; File
											,	"SendMail":                           	408  	; File
											,	"Properties":                           	409  	; File
											,	"OpenLast1":                         	510  	; File
											,	"OpenLast2":                         	511  	; File
											,	"OpenLast3":                         	512  	; File
											,	"Exit":                                    	405  	; File
											,	"SinglePage":                         	410  	; View
											,	"DoublePage":                       	411  	; View
											,	"BookView":                           	412  	; View
											,	"ShowPagesContinuously":     	413  	; View
											,	"TurnCounterclockwise":         	415  	; View
											,	"TurnClockwise":                    	416  	; View
											,	"Presentation":                        	418  	; View
											,	"Fullscreen":                          	421  	; View
											,	"Bookmark":                          	000  	; View - do not use! empty call!
											,	"ShowToolbar":                      	419  	; View
											,	"SelectAll":                             	422  	; View
											,	"CopyAll":                              	420  	; View
											,	"NextPage":                           	430  	; GoTo
											,	"PreviousPage":                      	431  	; GoTo
											,	"FirstPage":                            	432  	; GoTo
											,	"LastPage":                            	433  	; GoTo
											,	"GotoPage":                          	434  	; GoTo
											,	"Back":                                  	558  	; GoTo
											,	"Forward":                             	559  	; GoTo
											,	"Find":                                   	435  	; GoTo
											,	"FitPage":                              	440  	; Zoom
											,	"ActualSize":                          	441  	; Zoom
											,	"FitWidth":                             	442  	; Zoom
											,	"FitContent":                          	456  	; Zoom
											,	"CustomZoom":                     	457  	; Zoom
											,	"Zoom6400":                        	443  	; Zoom
											,	"Zoom3200":                        	444  	; Zoom
											,	"Zoom1600":                        	445  	; Zoom
											,	"Zoom800":                          	446  	; Zoom
											,	"Zoom400":                          	447  	; Zoom
											,	"Zoom200":                          	448  	; Zoom
											,	"Zoom150":                          	449  	; Zoom
											,	"Zoom125":                          	450  	; Zoom
											,	"Zoom100":                          	451  	; Zoom
											,	"Zoom50":                            	452  	; Zoom
											,	"Zoom25":                            	453  	; Zoom
											,	"Zoom12.5":                          	454  	; Zoom
											,	"Zoom8.33":                          	455  	; Zoom
											,	"AddPageToFavorites":           	560  	; Favorites
											,	"RemovePageFromFavorites": 	561  	; Favorites
											,	"ShowFavorites":                    	562  	; Favorites
											,	"CloseFavorites":                    	1106  	; Favorites
											,	"CurrentFileFavorite1":           	600  	; Favorites
											,	"CurrentFileFavorite2":           	601  	; Favorites -> I think this will be increased with every page added to favorites
											,	"ChangeLanguage":               	553  	; Settings
											,	"Options":                             	552  	; Settings
											,	"AdvancedOptions":               	597  	; Settings
											,	"VisitWebsite":                        	550  	; Help
											,	"Manual":                              	555  	; Help
											,	"CheckForUpdates":               	554  	; Help
											,	"About":                                	551}  	; Help

	}

	If !SumatraCommands.HasKey(command)
		return 0

	If SumatraID
		PostMessage, 0x111, % SumatraCommands[command],,, % "ahk_id " SumatraID
	else
		return SumatraCommands[command]

}





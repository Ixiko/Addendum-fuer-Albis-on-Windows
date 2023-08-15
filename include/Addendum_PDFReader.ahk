; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;     	Addendum_PDFReader
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;		Verwendung:	    	-	RPA Funktionen PDF Anzeigeprogramme für
;	                                   	1. Sumatra PDF V3.1,3.2 und 3.3
;	                                  	2. FoxitReader ab V9+
;
;
;		Abhängigkeiten:	-	Addendum_Internal, Addendum_Controls, Addendum_Window
;
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Addendum_StackifyGui last change:    	13.09.2022
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
return

; -------------------------------------------------------------
; SUMATRA/FOXIT Universell
; -------------------------------------------------------------
PdfPrint(Printer, PDFViewerClass, PDFViewerHwnd) {                      	;-- Pdf-Datei ausdrucken

	; Funktion bedient den Druckvorgang auf einem Hardwaredrucker
	; auch für virtuelle Druckertreiber z.B. Fax, PDF-Druckertreiber geeignet

		If InStr(PDFViewerClass, "Foxit") {                ; FoxitReader

			PdfPages	:= FoxitReader_GetPages(PDFViewerHwnd)                                   	; Seitenzahl des Dokumentes ermitteln
			PrintResult := FoxitReader_ToPrint(PDFViewerHwnd, Printer)                           	; Dokument drucken
			If (PrintResult.ItemNr = 0) {                                                                          	; Fehlerausgabe wenn Drucker nicht vorhanden ist
				PraxTT("Drucker: " Printer "`nkonnte nicht gefunden werden.`n"
				      	.  "Bitte die Druckereinstellungen in der Addendum.ini prüfen.", "8 3")
				return                                                                                                    	; Druck-Dialog wird für manuelle Auswahl offen gelassen
			}
			FoxitInvoke("Close", PDFViewerHwnd)                                                           	; dieses Dokument schließen

		}
		else If InStr(PDFViewerClass, "Sumatra") {	; Sumatra PdfReader

			PdfPages	:= Sumatra_GetPages(PDFViewerHwnd)                                        	; Seitenzahl des Dokumentes auslesen
			PrintResult	:= Sumatra_ToPrint(PDFViewerHwnd, Printer)                              	; Dokument drucken
			If (PrintResult.ItemNr = 0) {                                                                           	; Fehlerausgabe wenn Drucker nicht vorhanden ist
				PraxTT("Drucker: " Printer "`nkonnte nicht gefunden werden.`n"
				      	.  "Bitte die Druckereinstellungen in der Addendum.ini prüfen.", "8 3")
				return                                                                                                    	; Druck-Dialog wird für manuelle Auswahl offen gelassen
			}
			SumatraInvoke("Close", PDFViewerHwnd)                                                    	; dieses Dokument schließen

		}

return PdfPages
}

PdfSaveAs(PdfFullFilePath,PDFViewerClass, PDFViewerHwnd) {       	;-- Callback Funktion für Befundexporte (_scan)

		static foxitSaveAs    	:= "Speichern unter ahk_class #32770 ahk_exe FoxitReader.exe"
		static sumatraSaveAs	:= "Speichern unter ahk_class #32770 ahk_exe SumatraPDF.exe"
		;static rxNotAllowed   	:= "[" Chr(0x20) "~#%&<>\?\/|\{\}\|\*]"
		static rxNotAllowed   	:= "[~#%&\<\>\?\/|\{\}\|\*]"

	; illegale Zeichen entfernen
		PdfFullFilePath := RegExReplace(PdfFullFilePath, rxNotAllowed)

	; verschiedene RPA-Routinen je nach benutztem PDFReader (im Moment FoxitReader und
		If InStr(PDFViewerClass, "Foxit") {                                 	; FoxitReader

			FoxitInvoke("SaveAs", PDFViewerHwnd)                    	; 'Speichern unter' - Dialog
			WinWait, % foxitSaveAs,, 6                                       	; wartet 6 Sekunden auf das Dialogfenster
			If !(hSaveAs := GetHex(WinExist(foxitSaveAs)))        	; 'Speichern unter' - Dialog handle
				return
			VerifiedSetText("Edit1", PdfFullFilePath, hSaveAs)      	; Speicherpfad eingeben
			VerifiedClick("Button3", hSaveAs,,, true)                   	; Speichern Button drücken
			FoxitInvoke("Close", PDFViewerHwnd)                        	; dieses Dokument schließen

		}
		else If InStr(PDFViewerClass, "Sumatra") {                   	; Sumatra PdfReader

			SumatraInvoke("SaveAs", PDFViewerHwnd)                	; 'Speichern unter' - Dialog
			WinWait, % sumatraSaveAs,, 6                                  	; wartet 6 Sekunden auf das Dialogfenster
			If !(hSaveAs := GetHex(WinExist(sumatraSaveAs)))     	; 'Speichern unter' - Dialog handle
				return
			VerifiedSetText("Edit1", PdfFullFilePath, hSaveAs)      	; Speicherpfad eingeben
			VerifiedClick("Button2", hSaveAs,,, true)                  	; Speichern Button drücken
			SumatraInvoke("Close", PDFViewerHwnd)                  	; dieses Dokument schließen

		}

	; Wartet bis Datei fertig gespeichert wurde
		while !(FExists := FileExist(PdfFullFilePath)) {
			Sleep, 200
			If (A_Index > 40)
				break
		}

return FExists
}


; -------------------------------------------------------------
; SUMATRA PDF
; -------------------------------------------------------------
SumatraInvoke(command, hSumatra="", ShowExceptions:=true) {  	;-- wm_command wrapper for SumatraPDF V3.1-3.3

	/* DESCRIPTION of FUNCTION:  SumatraInvoke()

		                                                                               by Ixiko (version 13.09.2022)
		..............................................................................................................................
		a wm_command wrapper for SumatraPdf V3.1-V3.4
		..............................................................................................................................
		Remark:
		- SumatraPDF has changed all wm_command codes from V3.1 to V3.4
		- the script tries to automatically recognize the version of the addressed SumatraPDF
		  process in order to send the correct menu commands
		- maybe not all commands are listed!
		-----------------------------------------------------------------------------------------------------
		Parameters:
		- command:        	the names are borrowed from menu or toolbar names. However,
					        		no whitespaces or hyphens are used, only letters
		- hSumatra:       	by use of a valid handle, this function will post your command to
					        		Sumatra. Otherwise by use of a version string ("3.X")
				        			this function returns the wm_command code.
		- ShowExceptions:	you can suppress the throwing of exceptions by this function
		...........................................................
		Result:
		- You have to control the success of postmessage command yourself!
		---------------------------------------------------------------------------------------------------

		---------------------------------------------------------------------------------------------------
		      EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES

		SumatraInvoke("ShowToolbar", "3.2")      SumatraInvoke("DoublePage", hSumatra)
		.................................................       ...................................................................
		this one only returns the Sumatra                    sends the command "DoublePage" to
        command-code                                             your specified Sumatra process using
																   parameter 2 (hSumatra) as window handle.
															 		          command-code will be returned too
		---------------------------------------------------------------------------------------------------

	*/

		static	SumatraCmds

	; Sumatra wm_commands
		If !IsObject(SumatraCmds) {
			SumatraCmds := Object()
			SumatraCmds["3.1"] := { 	"NewWindow":                      	0      	; not available in this version -dummy command
												, 	"Open":                                 	400  	; File
												,	"Close":                                 	401  	; File
												,	"ShowInFolder":                        	0      	; not available in this version -dummy command
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
												,	"MangaMode":                       	0      	; not available in this version -dummy command
												,	"TurnCounterclockwise":         	415  	; View
												,	"RotateLeft":                           	415  	; View
												,	"TurnClockwise":                    	416  	; View
												,	"RotateRight":                          	416  	; View
												,	"Presentation":                        	418  	; View
												,	"Fullscreen":                           	421  	; View
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
												,	"FitASinglePage":                    	[410, 440] ; to hat V3.2 function (not tested)
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
			SumatraCmds["3.2"] := { 	"NewWindow":                      	450  	; File
												, 	"Open":                                 	400  	; File
												,	"Close":                                 	404  	; File
												,	"ShowInFolder":                        	410  	; File
												,	"SaveAs":                               	406  	; File
												,	"Rename":                              	610  	; File
												,	"Print":                                   	408  	; File
												,	"SendMail":                           	418  	; File
												,	"Properties":                           	420  	; File
												,	"OpenLast1":                         	570  	; File
												,	"OpenLast2":                         	571  	; File
												,	"OpenLast3":                         	572  	; File
												,	"Exit":                                    	412  	; File
												,	"SinglePage":                         	422  	; View
												,	"DoublePage":                       	423  	; View
												,	"BookView":                           	424  	; View
												,	"ShowPagesContinuously":     	425  	; View
												,	"MangaMode":                       	426  	; View
												,	"RotateLeft":                              	432  	; View
												,	"RotateRight":                           	434  	; View
												,	"Presentation":                        	438  	; View
												,	"Fullscreen":                          	444  	; View
												,	"Bookmark":                          	000  	; View - do not use! empty call!
												,	"ShowToolbar":                      	440  	; View
												,	"SelectAll":                             	446  	; View
												,	"CopyAll":                              	442  	; View
												,	"NextPage":                           	460  	; GoTo
												,	"PreviousPage":                      	462  	; GoTo
												,	"FirstPage":                            	464  	; GoTo
												,	"LastPage":                            	466  	; GoTo
												,	"GotoPage":                          	468  	; GoTo
												,	"Back":                                  	596  	; GoTo
												,	"Forward":                             	598  	; GoTo
												,	"Find":                                   	470  	; GoTo
												,	"FindNext":                             	472  	; Toolbar
												,	"FindPrevious":                         	474  	; Toolbar
												,	"MatchCase":                          	476  	; Toolbar
												,	"FitWithContinuously":               	3026	; Toolbar
												,	"FitASinglePage":                    	3027	; Toolbar
												,	"ZoomIn":                               	3012	; Toolbar
												,	"ZoomOut":                            	3013	; Toolbar
												,	"FitPage":                              	480  	; Zoom
												,	"ActualSize":                          	481  	; Zoom
												,	"FitWidth":                             	482  	; Zoom
												,	"FitContent":                          	496  	; Zoom
												,	"CustomZoom":                     	497  	; Zoom
												,	"Zoom6400":                        	483  	; Zoom
												,	"Zoom3200":                        	484  	; Zoom
												,	"Zoom1600":                        	485  	; Zoom
												,	"Zoom800":                          	486  	; Zoom
												,	"Zoom400":                          	487  	; Zoom
												,	"Zoom200":                          	488  	; Zoom
												,	"Zoom150":                          	489  	; Zoom
												,	"Zoom125":                          	490  	; Zoom
												,	"Zoom100":                          	491  	; Zoom
												,	"Zoom50":                            	492  	; Zoom
												,	"Zoom25":                            	493  	; Zoom
												,	"Zoom12.5":                          	494  	; Zoom
												,	"Zoom8.33":                          	495  	; Zoom
												,	"AddPageToFavorites":           	600  	; Favorites
												,	"RemovePageFromFavorites": 	602  	; Favorites
												,	"ShowCloseFavorites":               	604  	; Favorites
												,	"CurrentFileFavorite1":           	700  	; Favorites
												,	"CurrentFileFavorite2":           	701  	; Favorites -> I think this will be increased with every page added to favorites
												,	"ChangeLanguage":               	588  	; Settings
												,	"Options":                             	586  	; Settings
												,	"AdvancedOptions":               	632  	; Settings
												,	"VisitWebsite":                        	582  	; Help
												,	"Manual":                              	592  	; Help
												,	"CheckForUpdates":               	590  	; Help
												,	"About":                                	584  	; Help
												,	"HighlightLinks":                       	616  	; Debug
												,	"ToggleEBookUI":                     	624  	; Debug
												,	"MuiDebugPaint":                     	626  	; Debug
												,	"MuiDebugPaint":                     	626  	; Debug
												,	"AnnotationFromSelection":      	628  	; Debug
												,	"DownloadSymbols":                 	630}  	; Debug
         	SumatraCmds["3.3"] := { 	"NewWindow":                      	230  	; File
								        		, 	"Open":                                 	220  	; File
								        		,	"Close":                                 	204  	; File
								        		,	"ShowInFolder":                        	207  	; File
								        		,	"SaveAs":                               	205  	; File
								        		,	"Rename":                              	208  	; File
								        		,	"Print":                                   	206  	; File
								        		,	"SendMail":                           	212  	; File
								        		,	"Properties":                           	213  	; File
								        		,	"OpenLast1":                         	362  	; File
								        		,	"OpenLast2":                         	363  	; File
								        		,	"OpenLast3":                         	364  	; File
								        		,	"Exit":                                    	209  	; File
								        		,	"SinglePage":                         	215  	; View
								        		,	"DoublePage":                       	216  	; View
								        		,	"BookView":                           	217  	; View
								        		,	"ShowPagesContinuously":     	218  	; View
								        		,	"MangaMode":                       	219  	; View
								        		,	"RotateLeft":                              	220  	; View
								        		,	"RotateRight":                           	221  	; View
								        		,	"Presentation":                        	224  	; View
								        		,	"Fullscreen":                          	223  	; View
								        		,	"Bookmark":                          	000  	; View - do not use! empty call!
								        		,	"ShowToolbar":                      	225  	; View
								        		,	"SelectAll":                             	446  	; View
								        		,	"CopyAll":                              	442  	; View
								        		,	"NextPage":                           	235  	; GoTo
								        		,	"PreviousPage":                      	236  	; GoTo
								        		,	"FirstPage":                            	464  	; GoTo
								        		,	"LastPage":                            	466  	; GoTo
								        		,	"GotoPage":                          	468  	; GoTo
								        		,	"Back":                                  	596  	; GoTo
								        		,	"Forward":                             	598  	; GoTo
								        		,	"Find":                                   	245  	; GoTo
								        		,	"FindNext":                             	243  	; Toolbar
								        		,	"FindPrevious":                         	244  	; Toolbar
								        		,	"MatchCase":                          	476  	; Toolbar
								        		,	"FitWithContinuously":               	3026	; Toolbar
								        		,	"FitASinglePage":                    	3027	; Toolbar
								        		,	"ZoomIn":                               	269  	; Toolbar
								        		,	"ZoomOut":                            	270   	; Toolbar
								        		,	"FitPage":                              	251  	; Zoom
								        		,	"ActualSize":                          	252  	; Zoom
								        		,	"FitWidth":                             	271  	; Zoom
								        		,	"FitContent":                          	272  	; Zoom
								        		,	"CustomZoom":                     	268  	; Zoom
								        		,	"Zoom6400":                        	254  	; Zoom
								        		,	"Zoom3200":                        	255  	; Zoom
								        		,	"Zoom1600":                        	256  	; Zoom
								        		,	"Zoom800":                          	257  	; Zoom
								        		,	"Zoom400":                          	258  	; Zoom
								        		,	"Zoom200":                          	259  	; Zoom
								        		,	"Zoom150":                          	260  	; Zoom
								        		,	"Zoom125":                          	261  	; Zoom
								        		,	"Zoom100":                          	262  	; Zoom
								        		,	"Zoom50":                            	263  	; Zoom
								        		,	"Zoom25":                            	264  	; Zoom
								        		,	"Zoom12.5":                          	265  	; Zoom
								        		,	"Zoom8.33":                          	266  	; Zoom
								        		,	"AddPageToFavorites":           	600  	; Favorites
								        		,	"RemovePageFromFavorites": 	602  	; Favorites
								        		,	"ShowCloseFavorites":               	604  	; Favorites
								        		,	"CurrentFileFavorite1":           	700  	; Favorites
								        		,	"CurrentFileFavorite2":           	701  	; Favorites -> I think this will be increased with every page added to favorites
								        		,	"ChangeLanguage":               	588  	; Settings
								        		,	"Options":                             	586  	; Settings
								        		,	"AdvancedOptions":               	632  	; Settings
								        		,	"VisitWebsite":                        	582  	; Help
								        		,	"Manual":                              	592  	; Help
								        		,	"CheckForUpdates":               	590  	; Help
								        		,	"About":                                	584  	; Help
								        		,	"HighlightLinks":                       	616  	; Debug
								        		,	"ToggleEBookUI":                     	624  	; Debug
								        		,	"MuiDebugPaint":                     	626  	; Debug
								        		,	"AnnotationFromSelection":      	628  	; Debug
								        		,	"DownloadSymbols":                 	630}  	; Debug
			SumatraCmds["3.4"] := { 	"NewWindow":                      	234  	; File !
												, 	"Open":                                 	201  	; File !
												,	"Close":                                 	203  	; File !
												,	"ShowInFolder":                        	206  	; File !
												,	"SaveAs":                               	204  	; File !
												,	"Rename":                              	207  	; File !
												,	"Print":                                   	205  	; File !
												,	"SendMail":                           	211  	; File !
												,	"Properties":                           	212  	; File !
												,	"CommandPalette":                 	346  	; File !
												,	"OpenLast1":                         	416  	; File !
												,	"OpenLast2":                         	417  	; File !
												,	"OpenLast3":                         	418  	; File !
												,	"Exit":                                    	208  	; File !

												,	"SinglePage":                         	213  	; View !
												,	"DoublePage":                       	214  	; View !
												,	"BookView":                           	215  	; View !
												,	"ShowPagesContinuously":     	216  	; View !
												,	"RotateLeft":                              	218  	; View !
												,	"RotateRight":                           	219  	; View !
												,	"Presentation":                        	224  	; View !
												,	"Fullscreen":                          	221  	; View !
												,	"ShowBookmarks":                    	220  	; View !
												,	"ShowToolbar":                      	225  	; View !
												,	"ShowScrollbars":                      	226  	; View !
												,	"SelectAll":                             	233  	; View ?
												,	"CopyAll":                              	442  	; View ?

												,	"NextPage":                           	249  	; GoTo !
												,	"PreviousPage":                      	250  	; GoTo !
												,	"FirstPage":                            	251  	; GoTo !
												,	"LastPage":                            	252  	; GoTo !
												,	"Page...":                                	253  	; GoTo !
												,	"Back":                                  	341  	; GoTo !
												,	"Forward":                             	342  	; GoTo !
												,	"Find":                                   	254  	; GoTo !

												,	"FindNext":                             	255  	; Toolbar !
												,	"FindPrevious":                         	256  	; Toolbar !
												,	"MatchCase":                          	259  	; Toolbar !
												,	"FitWithContinuously":               	284	  	; Toolbar !
												,	"FitASinglePage":                    	285   	; Toolbar !
												,	"ZoomIn":                               	282  	; Toolbar !
												,	"ZoomOut":                            	283   	; Toolbar !

												,	"FitPage":                              	264  	; Zoom !
												,	"ActualSize":                          	265  	; Zoom !
												,	"FitWidth":                             	266  	; Zoom !
												,	"FitContent":                          	280  	; Zoom !
												,	"CustomZoom":                     	281  	; Zoom !
												,	"Zoom6400":                        	267  	; Zoom !
												,	"Zoom3200":                        	268  	; Zoom !
												,	"Zoom1600":                        	269  	; Zoom !
												,	"Zoom800":                          	270  	; Zoom !
												,	"Zoom400":                          	271  	; Zoom !
												,	"Zoom200":                          	272  	; Zoom !
												,	"Zoom150":                          	273  	; Zoom !
												,	"Zoom125":                          	274  	; Zoom !
												,	"Zoom100":                          	275  	; Zoom !
												,	"Zoom50":                            	276  	; Zoom !
												,	"Zoom25":                            	277  	; Zoom !
												,	"Zoom12.5":                          	278  	; Zoom !
												,	"Zoom8.33":                          	279  	; Zoom !

												,	"TranslateWithGoogle":             	229  	; Selection !
												,	"TranslateWithDeeplL":             	230  	; Selection !
												,	"SearchWithGoogle":             	231  	; Selection !
												,	"SearchWithGoogle":             	232  	; Selection !
												,	"SearchAll":                           	233  	; Selection !
												,	"CopyToClipboard":                	228  	; Selection !

												,	"AddPageToFavorites":           	311  	; Favorites !
												,	"RemovePageFromFavorites": 	312  	; Favorites !
												,	"ShowFavorites":                     	313  	; Favorites !
												,	"CurrentFileFavorite1":           	449  	; Favorites !
												,	"CurrentFileFavorite2":           	450  	; Favorites ! -> I think this will be increased with every page added to favorites

												,	"ChangeLanguage":               	588  	; Settings !
												,	"Options":                             	303  	; Settings !
												,	"AdvancedOptions":               	304  	; Settings !

												,	"VisitWebsite":                        	308  	; Help !
												,	"Manual":                              	307  	; Help !
												,	"CheckForUpdates":               	306  	; Help !
												,	"About":                                	309  	; Help !

												,	"HighlightLinks":                       	616  	; Debug
												,	"ToggleEBookUI":                     	624  	; Debug
												,	"MuiDebugPaint":                     	626  	; Debug
												,	"AnnotationFromSelection":      	628  	; Debug
												,	"DownloadSymbols":                 	630}  	; Debug

		}

	; ---------------------------------------------------------------------------------------------------------------------
	; try to determine the version of the running SumatraPDF process from the passed window handle
	; ---------------------------------------------------------------------------------------------------------------------
	; parts of following code are taken from WinSpy
		WinGetClass, class, % "ahk_id " hSumatra
		If InStr(class, "SUMATRA_PDF_FRAME") {

		; get version of running sumatra.exe
			WinGet SumatraPID, PID, % "ahk_id " hSumatra
			Enum := ComObjGet("winmgmts:").ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId=" SumatraPID)._NewEnum
			If (Enum[Process])
				FileGetVersion ProgVer, % Process.ExecutablePath
			RegExMatch(ProgVer, "\d\.\d", version)

		; prevent failures
			If (SumatraCmds[version][command] = 0)
				return "" 														         			; return on dummy command
			else If !SumatraCmds[version].haskey(command) {
				If ShowExceptions
					throw "Parameter #1 [" command "] unknown in SumatraPDF version " version
			}

		; execute stream-like commands
			If IsObject(SumatraCmds[version][command]) {
				For cmdindex, cmd in SumatraCmds[version][command] {
					SendMessage, 0x111, % cmd,,, % "ahk_id " hSumatra		; ## maybe I will change it  back to PostMessage, if this takes to long
					EL .= (EL ? "|" : "") ErrorLevel
					;~ If (cmdindex < wmcmds.Count())
						;~ Sleep, 300                                                               	; I think a delay is necessary here
				}
			}
		; execute single command
			else {
				SendMessage, 0x111, % SumatraCmds[version][command],,, % "ahk_id " hSumatra
				return ErrorLevel
			}

		}
	; returns wm_command code for this Sumatra Programmversion
		else {

			If RegExMatch(hSumatra, "\d\.\d", version) {
				If (SumatraCmds[version][command] = 0) {
					return "" 															    		; return on dummy command
				} else If !SumatraCmds[version].haskey(command) {
					If ShowExceptions
						throw "Parameter #1 [" command "] unknown in SumatraPDF version " version
				} else
					return SumatraCmds[version][command]
			}
			else
				If ShowExceptions
					throw 	"Parameter #2 invalid!`n"
							. 	"The passed handle of Sumatra PDF was neither a correct`n"
							. 	"window handle nor a valid string for a program version."

		}

}

SumatraDDE(hSumatra, cmd, params*) {                                  		;-- Befehle an Sumatra per DDE schicken

	/*  DESCRIPTION

		https://github.com/sumatrapdfreader/sumatrapdf/issues/1398
		https://gist.github.com/nod5/4d172a31a3740b147d3621e7ed9934aa
		functions Send_WM_COPYDATA() and RECEIVE_WM_COPYDATA() are required
		Required data to tell SumatraPDF to interpret lpData as DDE command text, always 0x44646557

		SumatraPDF DDE command unicode text, https://www.sumatrapdfreader.org/docs/DDE-Commands.html

		DDE Commands
		Sumatra can be controlled in a limited way from other software by sending DDE commands. They are mostly
		used to use SumatraPDF as a preview tool from e.g. LaTeX editors that generate PDF files.

		Format of DDE comands
			Single DDE command:   	[Command(parameter1, parameter2, ..., )]
			Multiple DDE commands: 	[Command1(parameter1, parameter2, ..., )][Command2(...)][...]

		List of DDE commands:
        	[Open file]
			format:     	[Open("<pdffilepath>"[,<newwindow>,<focus>,<forcerefresh>])]
			arguments:	if newwindow is 1 then a new window is created even if the file is already open
								if focus is 1 then the focus is set to the window
								if forcerefresh is 1 the command forces the refresh of the file window if already open
								(useful for files opened over network that don't get file-change notifications)".
			example:   	[Open("c:\file.pdf", 1, 1, 0)]

			[Forward-Search]
			format: [ForwardSearch(["<pdffilepath>",]"<sourcefilepath>",<line>,<column>[,<newwindow>,<setfocus>])]
			arguments:
			pdffilepath:     	path to the PDF document (if this path is omitted and the document isn't already open,
	                    			SumatraPDF won't open it for you)
			column:         	this parameter is for future use (just always pass 0)
			newwindow:  	1 to open the document in a new window (even if the file is already opened)
			focus:             	1 to set focus to SumatraPDF's window.
			examples:     	[ForwardSearch("c:\file.pdf","c:\folder\source.tex",298,0)]
                                   	[ForwardSearch("c:\folder\source.tex",298,0,0,1)]

           	[GotoNamedDest]
           	format:         	[GotoNamedDest("<pdffilepath>","<destination name>")]
           	example:       	[GotoNamedDest("c:\file.pdf", "chapter.1")]
           	note:             	the pdf file must be already opened

           	[Go to page]
           	format:         	[GotoPage("<pdffilepath>",<page number>)]
           	example:       	[GotoPage("c:\file.pdf", 37)]
           	note:             	the pdf file must be already opened.

        	[SetView]
   			format: 			[SetView("<pdffilepath>","<view mode>",<zoom level>[,<scrollX>,<scrollY>])]
   			arguments:
   			view mode: 		"single page"
   									"facing"
    								"book view"
    								"continuous"
    								"continuous facing"
    								"continuous book view"
   			zoom level : 		either a zoom factor between 8 and 6400 (in percent) or one
	                            	of -1 (Fit Page), -2 (Fit Width) or -3 (Fit Content)
   			scrollX, scrollY: 	PDF document (user) coordinates of the point to be visible in the top-left of the window
   			example: 			[SetView("c:\file.pdf","continuous",-3)]
   			note: 				the pdf file must already be opened
    		Example:			[SetView("c:\file.pdf","continuous",-3)]

			made by Ixiko 	Feb. 2021

	 */

	static dwData := 0x44646557

    														; p1=filepath, p2=newwindow, p3=1 for focus, p4=1 for force refresh
	static lpData := { 	"OpenFile"         	: ("[Open(""_p1"",_p2,_p3,_p4)]")

	    													; [p1=filepath,] p2=sourcefilepath, p3=line, p4=column[, p5=1 for new window, p6=1 to set focus]
		    				,	"ForwardSearch" 	: ("[ForwardSearch(""_p1"",""_p2"",_p3,_p4,_p5,_p6)]")

			    											; p1=filepath, p2=destination name
		    				,	"GotoNamedDest"	: ("[GotoNamedDest(""_p1"",""_p2"")]")

		    												; p1=filepath, p2=PageNr
		    				,	"GotoPage"        	: ("[GotoPage(""_p1"",_p2)]")

		    												; p1=filepath, p2=view mode, p3=zoom level [, p4=scrollX, p5=scrollY>]
		    				,	"SetView"           	: ("[SetView(""_p1"",""_p2"",_p3,_p4,_p5)]")}

	lpDataSend := lpData[cmd]

	For index, param in params
		lpDataSend := StrReplace(lpDataSend, "_p" index, param)
	;~ SciTEOutput("lpdata: " lpDataSend)

	lpDataSend := RegExReplace(lpDataSend, ",""*_p\d""*")
	;~ SciTEOutput("lpdata: " lpDataSend)

return Send_WM_COPYDATA_EX(hSumatra, dwData, lpDataSend)
}

Sumatra_Close(hSumatra) {                                                       	;-- beendet einen Sumatra Prozeß
	WinGet, SumatraPID, PID, % "ahk_id " hSumatra
	Process, Close, % SumatraPID
	If ErrorLevel
		return 1
	else if !ErrorLevel && hSumatra {
		WinGetClass, wclass, % "ahk_id " hSumatra
		If InStr(wclass, "SUMATRA_PDF_FRAME") {
			SumatraInvoke("Exit", hSumatra, false)
			WinGetClass, wclass, % "ahk_id " hSumatra
			return InStr(wclass, "SUMATRA_PDF_FRAME") ? 0 : 1
		}
	}
return 0
}

Sumatra_Embed(hparent, x, y, w, h) {                                        	;-- Einbetten von Sumatra PDF READER in eine AutoHotkey Gui

	/* Hinweise

		Funktion für das Einbetten eines Sumatra PDF READER Fensters in eine AutoHotkey Gui .

			-	Breite und Höhe aller Steuerelemente bis auf den des Renderbereiches werden auf 0 verkleinert
			-	Nutzer können somit prinzipiell nicht mehr mit Sumatra PDF interagieren
			-	Interaktion ist weiterhin über cmdline Aufrufe und DDE-Kommandos möglich
			-	damit das eingebette Fenster nach den Veränderungen neu gezeichnet wird, muss zuvor das
				Elternfenster angezeigt werden. Danach muss das Sumatra PDF Fenster einmalig in der
				Höhe verändert werden um ein Neuzeichnen und dann Sichtbarkeit zu erreichen. Dafür wird
				die zugehörige 2.Funktion (Sumatra_EmbedFinalize()) ausgeführt.
				Rückgabeparameter ist ein Objekt mit dem hwnd des Sumatra PDF Fensters, der Prozeß-ID (PID)
				und dem Handle des Renderbereiches.
				Beispiel:

				smtra := Sumatra_Embed(hparent, smtraX, smtraY, smtraW, smtraH)
				Gui, Show, % "w" smtraX+smtraW+2*GuiMarginX, embedded Sumatra PDF gui
				Sumatra_EmbedFinalize(hparent, smtra.hwnd)

				Achtung: 	-	getestet nur mit Sumatra PDF Version 3.3
								-	es wird eine installierte Sumatra PDF Version benötigt
								-	Skript legt eine Arbeitskopie der Sumatra Einstellungen in einem lokalen Unter-
									verzeichnis (/resources) an, damit die durch das Skript gemachten Änderungen
									an den Einstellungen nur für das eingebundene Sumatra Fenster gelten

	*/

		static 	SmtraExist := false, SmtraInit := true
		static 	SumatraCMD
		static 	SmtraClass := "ahk_class SUMATRA_PDF_FRAME"

		If SmtraInit {
			SmtraInit := false
			SumatraCMD := GetAppImagePath("SumatraPDF")
			If (!SumatraCMD) || !FileExist(SumatraCMD)
				return "Sumatra_fails"
		}

		lDetectHiddenWin	:= A_DetectHiddenWindows
		lDetectHiddenText	:= A_DetectHiddenText
		DetectHiddenWindows	, On
		DetectHiddenText      	, On

	; Sumatra PDF per Console starten
		AppDataPath := A_AppData "\SumatraPDF\SumatraPDF-settings.txt"
		If !FileExist(A_ScriptDir "\resources\SumatraPDF-settings.txt") {
			If !InStr(FileExist(A_ScriptDir "\resources"), "D")
				FileCreateDir, % A_ScriptDir "\resources"
			If InStr(FileExist(A_ScriptDir "\resources"), "D")
				FileCopy, % AppDataPath, A_ScriptDir "\resources\SumatraPDF-settings.txt"
		}
		Run, % q SumatraCMD q " -appdata " q A_ScriptDir "\resources" q ,, Hide UseErrorLevel, PIDSumatra   ; -new-window
		WinWait, % SmtraClass,, 12

	; Sumatra Fenster Handle und erste Größenanpassung
		hSumatra := WinExist(SmtraClass)
		WinMove, % "ahk_id " hSumatra,,,, % w, % h

	; Sumatra Fenster einbinden
		DllCall("SetParent", "uint", hSumatra, "uint", hparent)
		WinMove, % "ahk_id " hSumatra,, % x, % y, % w, % h
		WinActivate, % "ahk_id " hSumatra

	; zunächst rahmenlosen Stil einrichten
		WinSet, Style		, 0x56000000, % "ahk_id " hSumatra
		WinSet, ExStyle	, 0x00000000, % "ahk_id " hSumatra

	; Toolbar ausblenden
		Sumatra_ShowToolbar(hSumatra, false)

	; Ansichtseinstellungen
		SumatraInvoke("SinglePage"  	, hSumatra, false)
		Sleep 100
		SumatraInvoke("FitPage"       	, hSumatra, false)
		Sleep 100

	; Window ClientEdge
		WinSet, ExStyle	, 0x00020000, % "ahk_id " hSumatra

	; weitere ID's für die Rückgabe
		WinGet     	,	SumatraPID, PID, % "ahk_id " hSumatra
		ControlGet	,	hCanvas, HWND,, SUMATRA_PDF_CANVAS1, % "ahk_id " hSumatra

		DetectHiddenWindows	, % lDetectHiddenWin
		DetectHiddenText      	, % lDetectHiddenText

return {"PID":SumatraPID, "Path":SumatraCMD, "hwnd":hSumatra, "hCanvas":hCanvas}
}

Sumatra_EmbedFinalize(hparent, hSumatra) {                             	;-- letzte Anpassung des eingebundenen Sumatra PDF Fensters

	; Modi ändern
		lDetectHiddenWin	:= A_DetectHiddenWindows
		lDetectHiddenText	:= A_DetectHiddenText
		DetectHiddenWindows	, Off
		DetectHiddenText      	, Off

		WinActivate, % "ahk_id " hSumatra
		ControlGet, hCanvas, hwnd,, SUMATRA_PDF_CANVAS1                            	, % "ahk_id " hSumatra

	; eingebundes Sumatra Fenster vergrößern
		ControlGetPos,smtrX, smtrY, smtrW, smtrH, SUMATRA_PDF_CANVAS1        	, % "ahk_id " hSumatra
		ControlMove, SUMATRA_PDF_FRAME1	,,,, % smtrH+29                            	, % "ahk_id " hSumatra

	; Menu- und Sysmenu Button verstecken
		ControlGet, hcuca	, hwnd,, CustomCaption1, % "ahk_id " hSumatra
		WinHide, % "ahk_id " hcuca
		WinGet, cuca, ControlListHwnd, % "ahk_id " hcuca
		For each, hwnd in StrSplit(cuca, "`n") {
			WinGet, Style, Style, % "ahk_id " hwnd
			if (Style & 0x10000000)                              ; nur wenn sichtbar
				WinMove, % "ahk_id " hwnd,,,, 0, 0
		}

	; Canvas-Steuerelement verschieben
		ControlGetPos	, cnvsY, cnvsY, cnvsW, cnvsH, SUMATRA_PDF_CANVAS1       	, % "ahk_id " hSumatra
		ControlMove 	, SUMATRA_PDF_CANVAS1, 0, 0, % cnvsW+2, % cnvsH+56	, % "ahk_id " hSumatra
		WinMove 	, % "ahk_id " hCanvas,, 0, 1
		ControlGetPos	, cnvsY, cnvsY, cnvsW, cnvsH, SUMATRA_PDF_CANVAS1       	, % "ahk_id " hSumatra

	; Modi wiederherstellen
		DetectHiddenWindows	, % lDetectHiddenWin
		DetectHiddenText      	, % lDetectHiddenText

return {"hwnd":hCanvas, "X":cnvsX, "Y":cnvsY, "W":cnvsW, "H":cnvsH}
}

Sumatra_GetActiveTab(hSumatra, cmd:="name") {                     	;-- ermittelt den Namen der aktuell angezeigten Datei
	ControlGet, hTab, Hwnd,, SysTabControl321, % "ahk_id " hSumatra
	TabIndex  	:= TabCtrl_GetCurSel(hTab)
	If RegExMatch(cmd, "i)(name|tabs)")
		tabs := ControlGetTabs(hTab)
return cmd="TabIndex" ? TabIndex : cmd="name" ? tabs[TabIndex] : tabs
}

Sumatra_GetPages(hSumatra="") {                                             	;-- aktuelle und maximale Seiten des aktuellen Dokumentes ermitteln

	If !hSumatra
		hSumatra := WinExist("ahk_class SUMATRA_PDF_FRAME")

	ControlGetText, PageDisp, Edit3 	, % "ahk_id " hSumatra
	ControlGetText, PageMax, Static3	, % "ahk_id " hSumatra
	RegExMatch(PageMax, "\s*(?<Max>\d+)", Page)

return {"disp":PageDisp, "max":PageMax}
}

Sumatra_PrintPDF(filepath:="", Printer:="")  {                              	;-- Pdf per commandline drucken

	global SumatraCMD, SumatraExist

	If !SumatraCMD {
		SumatraCMD := GetAppImagePath("SumatraPDF")
		If SumatraCMD && FileExist(SumatraCMD)
			SumatraExist := true
	}

	If SumatraCMD {

		smtracmdline	:= "-print-to " q Printer q " -print-settings " q "fit,duplex" q " -exit-when-done" ; Dateiname
		stdoutCMD    	:= SumatraCMD " " smtracmdline " " q . filepath . q
		;~ SciTEOutput(A_ThisFunc "`n" stdoutCMD)
		If (stdout := StdoutToVar(stdoutCMD))
			SciTEOutput(stdout)

	}

return stdout ? stdout : 1
}

Sumatra_Show(filepath:="", winParams:="", hparent:=0) {           	;-- SumatraPDF starten und Dokument anzeigen

	; hparent benutzen wenn das Sumatra PDF Fenster in eine Gui eingebettet ist
	; letzte Änderung: 19.02.2022

	static SumatraCMD
	;~ static cmdl1 := q "-new-window -view " q "single page" q " -zoom " q "fit page" q
	;~ static cmdl2 := q "-view " q "single page" q " -zoom " q " fit page" q
	static cmdl1 := q "-new-window -view single page -zoom fit page" q
	static cmdl2 := q "-view single page -zoom fit page" q

	If hparent
		ControlGet, hSumatra, hwnd,, SUMATRA_PDF_FRAME1, % "ahk_id " hparent
	else
		hSumatra := WinExist("ahk_class SUMATRA_PDF_FRAME")

  ; Sumatra PDF starten
	If !hSumatra {

		SumatraCMD	:= !SumatraCMD ? GetAppImagePath("SumatraPDF") : SumatraCMD
		cmdoptions   	:= (winParams ? winParams : cmdl1)
		cmdline         	:= filepath && FileExist(filepath) ? cmdoptions " " q filePath q : ""
		Run, % q SumatraCMD q " " cmdline,,  UseErrorLevel, PIDSumatra  ; Hide

		WinWait, % "ahk_class SUMATRA_PDF_FRAME",, 6
		hSumatra := WinExist("ahk_class SUMATRA_PDF_FRAME")
		WinGet, SumatraPID, PID, % "ahk_id " hSumatra
		WinWaitActive, % "ahk_class SUMATRA_PDF_FRAME",, 6

	}
  ; Dateipfad an Sumatra PDF übergeben
	else {

		If filepath && FileExist(filepath) {

			WinGet, SumatraPID, PID, % "ahk_id " hSumatra
			smtraTab := smtraTabO := Sumatra_GetActiveTab(hSumatra)

		  ; Datei per DDE Befehl laden lassen
			SumatraDDE(hSumatra, "OpenFile", filepath, 0, 0, 1)

		  ; Änderung des Fenstertitel abwarten
			while (smtraTab = smtraTabO) {
				smtraTab := Sumatra_GetActiveTab(hSumatra)
				If (A_Index > 200)     ; ~4s
					return 0
				Sleep 20
			}

		}

	}

	ControlGet, hSumatraCnvs, HWND,, SUMATRA_PDF_CANVAS1, % "ahk_id " hSumatra
	s	:= GetWindowSpot(hSumatra)
	c	:= GetWindowSpot(hSumatraCnvs)

  ; Einrichten der Ansicht
	SumatraDDE(hSumatra, "SetView", filepath, "single page", "-1") 	; -1 = fit page
	;~ SumatraInvoke("ShowToolbar", hSumatra)


return {"viewer": "SumatraPDF", "AR":c.CW/c.CH, "hwnd":hSumatra, "ID":hSumatra, "PID":SumatraPID, "x":s.X, "y":s.Y, "w":s.W, "h":s.H}
}

Sumatra_ShowToolbar(hSumatra, showTb:=true) {                      	;-- die Toolbar ein- oder ausblenden

	ControlGet, hReBar, hwnd,, ReBarWindow321, % "ahk_id " hSumatra
	WinGet, Style, Style, % "ahk_id " hReBar
	TbExist := (Style & 0x10000000) ? 1 : 0
	If !(showTb & TbExist) {
		SumatraInvoke("ShowToolbar", hSumatra, false)
		UserTbState := false
		while (!UserTbState && A_Index < 10) {                                         	; ~100ms
			If (A_Index > 1)
				Sleep 10
			ControlGet, hReBar, hwnd,, ReBarWindow321, % "ahk_id " hSumatra
			WinGet, Style, Style, % "ahk_id " hReBar
			TbExist := (Style & 0x10000000) ? 1 : 0
			UserTbState := showTb & TbExist
		}
		If !UserTbState
			SumatraInvoke("ShowToolbar", hSumatra, false)
	}

return UserTbState
}

Sumatra_ToPrint(SumatraID="", Printer="") {                                 	;-- Druck Dialoghandler - Ausdruck auf übergebenen Drucker

		; druckt das aktuell angezeigte Dokument
		; abhängige Biblitheken: LV_ExtListView.ahk

		static sumatraprint	:= "i)[(Print)|(Drucken)]+ ahk_class i)#32770 ahk_exe i)SumatraPDF.exe"

		rxPrinter:= StrReplace(Trim(Printer), " ", "\s")
		rxPrinter:= StrReplace(rxPrinter, "(", "\(")
		rxPrinter:= StrReplace(rxPrinter, ")", "\)")

		OldMatchMode := A_TitleMatchMode
		SetTitleMatchMode, RegEx                                                              	; RegEx Fenstervergleichsmodus einstellen

		SumatraInvoke("Print", SumatraID)                                                  	; Druckdialog wird aufgerufen
		WinWait, % sumatraprint,, 6                                                             	; wartet 6 Sekunden auf das Dialogfenster
		hSumatraPrint := GetHex(WinExist(sumatraprint))                               	; 'Drucken' - Dialog handle
		ControlGet, hLV, Hwnd,, SysListview321, % "ahk_id " hSumatraPrint    	; Handle der Druckerliste (Listview) ermittlen
		sleep 200                                                                                       	; Pause um Fensteraufbau abzuwarten
		ControlGet, Items	, List  , Col1 	,, % "ahk_id " hLV                             	; Auslesen der vorhandenen Drucker
		ItemNr := 0                                                                                    	; ItemNr auf 0 setzen
		Loop, Parse, Items, `n                                                                    	; Listview Position des Standarddrucker suchen
			If RegExMatch(A_LoopField, "i)^" rxPrinter) {                                	; Standarddrucker gefunden
				ItemNr := A_Index                                                                	; nullbasierende Zählung in Listview Steuerelementen
				break
			}
		If ItemNr {                                                                                    	; Drucker in der externen Listview auswählen
			objLV := ExtListView_Initialize(sumatraprint)                                	; Initialisieren eines externen Speicherzugriff auf den Sumatra-Prozeß
			ControlFocus,, % "ahk_id " objLV.hlv                                            	; Druckerauswahl fokussieren
			err	 := ExtListView_ToggleSelection(objLV, 1, ItemNr - 1)            	; gefundenes Listview-Element (Drucker) fokussieren und selektieren
			ExtListView_DeInitialize(objLV)                                                     	; externer Speicherzugriff muss freigegeben werden
			Sleep 200
			err	:= VerifiedClick("Button13", hSumatraPrint)                           	; 'Drucken' - Button wählen
			WinWaitClose, % "ahk_id " hSumatraPrint,, 3                              	; wartet max. 3 Sek. bis der Dialog geschlossen wurde
		}

		SetTitleMatchMode, % OldMatchMode                                            	; TitleMatchMode zurückstellen

return {"DialogID":hSumatraPrint, "ItemNr":ItemNr}                                 	; für Erfolgskontrolle und eventuelle weitere Abarbeitungen
}

Send_WM_COPYDATA_EX(hWin, dwData, lpData) 	{                     	;-- für die Kommunikation zwischen Sumatra und Autohotkey Skripten

	VarSetCapacity(COPYDATASTRUCT, 3*A_PtrSize, 0)
    cbData := (StrLen(lpData) + 1) * (A_IsUnicode ? 2 : 1)
    NumPut(dwData	, COPYDATASTRUCT, 0*A_PtrSize)
    NumPut(cbData 	, COPYDATASTRUCT, 1*A_PtrSize)
    NumPut(&lpData	, COPYDATASTRUCT, 2*A_PtrSize)
	SendMessage, 0x4a, 0, &COPYDATASTRUCT,, % "ahk_id " hWin ; 0x4a WM_COPYDATA

return ErrorLevel ;== "FAIL" ? false : true
}


; -------------------------------------------------------------
; FOXITREADER
; -------------------------------------------------------------
FoxitInvoke(command, FoxitID="") {		                                      	;-- wm_command wrapper for FoxitReader Version:  9.1

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

		FoxitCommands   := {	"SaveAs":                                                       	1299
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
		FoxitCommands1 := { 	"firstPage":                                                      	1286        	; View - Go To
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

		For key, val in FoxitCommands1
			FoxitCommands[key] := val

	}

	If FoxitID
		PostMessage, 0x111, % FoxitCommands[command],,, % "ahk_id " FoxitID
	else
		return FoxitCommands[command]
}

FoxitReader_GetPages(FoxitID="") {                                              	;-- aktuelle und maximale Seiten des aktuellen Dokumentes ermitteln

	; letzte Änderung 18.10.2020

	; nachsehen ob korrekte FoxitID übergeben wurde
		While (!FoxitID || !WinExist("ahk_id " FoxitID)) {
			If (FoxitID := WinExist("ahk_class classFoxitReader"))
				break
			else if (A_Index > 20)
				return {"disp":1, "max":1}
			Sleep 50
		}

	; Handle der Statusbar ermitteln
		WinGet, hCtrl, ControlList, % "ahk_id " FoxitID
		Loop, Parse, hCtrl, `n
			If InStr(A_LoopField, "BCGPRibbonStatusBar") {
				ControlGet, StatusbarHwnd, Hwnd,, % A_LoopField, % "ahk_id " FoxitID
				break
			}

	; Text der Steuerelemente nach Seitenanzeige durchsuchen
		WinGet, hCtrl, ControlList, % "ahk_id " StatusbarHwnd
		;SciTEOutput("Statusbarhwnd: " StatusbarHwnd "`nhCtrls: " hCtrl)
		Loop, Parse, hCtrl, `n
		{
			ControlGetText, Pages, % A_LoopField, % "ahk_id " StatusbarHwnd
			If RegExMatch(Pages, "(?<Disp>\d+)\s*\/\s*(?<Max>\d+)", Page) {
				PageDisp	:= StrLen(PageDisp) = 0	? 1 : PageDisp
				PageMax	:= StrLen(PageMax) = 0	? 1 : PageMax
				return {"disp":PageDisp, "max":PageMax}
			}
		}

return {"disp":1, "max":1} ; wenigsten eine 1 zurückgeben wenn nichts ermittelt werden konnte
}

FoxitReader_GetPDFPath() {                                                         	;-- den aktuellen Dokumentenpfad im 'Speichern unter' Dialog auslesen

	foxitSaveAs := "Speichern unter ahk_class #32770 ahk_exe FoxitReader.exe"
	If WinExist(foxitSaveAs) {
		WinGetText, allText, % foxitSaveAs
		RegExMatch(allText, "(?<Name>[\w+\s_\-\,]+\.pdf)\n.*Adresse\:\s*(?<Path>[A-Z]\:[\\\w\s_\-]+)\n", File)
		return FilePath "\" FileName
	}

return
}

FoxitReader_ToPrint(FoxitID="", Printer="") {                                 	;-- Druck Dialoghandler - Ausdruck auf übergebenen Drucker

		static foxitprint    	:= "i)[(Print)|(Drucken)]+ ahk_class i)#32770 ahk_exe i)FoxitReader.exe"

		If !FoxitID
			FoxitID := WinExist("ahk_class classFoxitReader")

		OldMatchMode := A_TitleMatchMode
		SetTitleMatchMode, RegEx                                                              	; RegEx Fenstervergleichsmodus einstellen

		FoxitInvoke("Print", FoxitID)                                                               	; 'Drucken' - Dialog wird aufgerufen
		WinWait, % foxitPrint,, 6                                                                 	; wartet 6 Sekunden auf das Dialogfenster
		hfoxitPrint	:= GetHex(WinExist(foxitPrint))                                        	; 'Drucken' - Dialog handle
		ItemNr  	:= VerifiedChoose("ComboBox1", hfoxitPrint, Printer)          	; Drucker auswählen
		If (ItemNr <> 0) {
			VerifiedClick("Button44", hfoxitPrint,,, true)                                    	; OK Button drücken
			WinWaitClose, % "ahk_id " hfoxitPrint,, 3                                    	; wartet max. 3 Sek. bis der Dialog geschlossen wurde
		}

		SetTitleMatchMode, % OldMatchMode                                            	; TitleMatchMode zurückstellen

return {"DialogID":hFoxitPrint, "ItemNr":ItemNr}                                   	; für Erfolgskontrolle und eventuelle weitere Abarbeitungen
}

FoxitReader_SignaturSetzen(FoxitID="") {                                		;-- zeichnet eine Signatur in die linke obere Ecke des Dokumentes

		; letzte Änderung: 19.09.2020

			CoordModeMouse_before :=  A_CoordModeMouse
			CoordMode, Mouse, Screen

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; Variablen
		;----------------------------------------------------------------------------------------------------------------------------------------------;{
			; ! NICHT ÄNDERN! dieser String wird für 'feiyus' FindText() Funktion benötigt, er entspricht der linken oberen Ecke des PDF-Frame (AfxWnd100su4)
			;~ static TopLeft :=	"|<>*210$71.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzw0000003zzzzs0000007"
			        	        	;~ . 	"zzzzk000000DzzzzU000000Tzzzz0000000zzzzy0000001zzzzw0000003zzzzs0000007zzzzk000000DzzzzU000000T"
			                		;~ . 	"zzzz0000000zzzzy0000001zzzzw0000003zzzzs0000007zzzzk000000DzzzzU000000Tzzzz0000000zzzzy0000001"

			static TopLeft :="|<>ABABAB-000000$47.000000000000000000000000000000000000000Tzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzk0007zzzU000Dzzy0000Tzzw0000zzzs0001zzzk0003zzzU0007zzz0000Dzzy0000Tzzw0000zzzs0001"
			static basetolerance := 0.5

			Addendum.PDF.RecentlySigned := true
		;}

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; Abbruch wenn kein FoxitReaderfenster
		;----------------------------------------------------------------------------------------------------------------------------------------------;{
			If !FoxitID
				If !(FoxitID := WinExist("ahk_class classFoxitReader"))
					return 0

			; PDF Backup!
			for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process") {
				If InStr(process.name, "FoxitReader") {
					RegExMatch(process.CommandLine, "\s\" q "(.*)?\" q, cmdline)
					break
				}
			}

			SplitPath, cmdline1, PDFName
			FileCopy, % cmdline1, % Addendum.BefundOrdner "\Backup\" PdfName, 1
		;}

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; Ermitteln des Childhandles (Docwindow) im Foxitreader (Bildschirmposition des Dokuments)
		;----------------------------------------------------------------------------------------------------------------------------------------------;{
			ActivateAndWait("ahk_id " FoxitID, 1)
			res	        	:= Controls("", "Reset", "")
			hDocTab      	:= Controls("", "GetActiveMDIChild, return hwnd", FoxitID)
			hDocWnd 	:= Controls("FoxitDocWnd", "hwnd", hDocTab)
			DocWnd      	:= GetWindowSpot(hDocWnd)

			;SciTEOutput("DocWnd handle: " hDocWnd " x" DocWnd.X  " y" DocWnd.Y " w" DocWnd.W " h" DocWnd.H )
		;}

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; FoxitReader vorbereiten für das Platzieren der Signatur
		;----------------------------------------------------------------------------------------------------------------------------------------------;{
			PraxTT("sende Befehl: 'Signatur platzieren'", "12 2")

			ActivateAndWait("ahk_id " FoxitID, 1)

			FoxitInvoke("SinglePage", FoxitID)
			FoxitInvoke("Fit_Page"  	, FoxitID)
			FoxitInvoke("FirstPage" 	, FoxitID)

			Sleep, 500
			MouseMove	, % DocWnd.X + 50, % DocWnd.Y + 50, 0
			MouseClick	, Left

		;}

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; Signatur setzen Menupunkt aufrufen
		;----------------------------------------------------------------------------------------------------------------------------------------------;{
			ActivateAndWait("ahk_id " FoxitID, 1)
			FoxitInvoke("Place_Signature", FoxitID)
		;}

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; sucht nach der linken oberen Ecke der Pdf Seite der 1.Seite, um dort im Anschluss die Signatur zu erstellen
		;----------------------------------------------------------------------------------------------------------------------------------------------;{
			PraxTT(" suche nach dem Signierbereich des Dokumentes.", "12 2")

			tryCount := 1
			FindSignatureRange:
			; Funktion FindText macht eine Bildsuche (entspricht der linken oberen Ecke des PDF Preview Bereiches)
			if (Ok := FindText(DocWnd.X, DocWnd.Y, DocWnd.W, DocWnd.H, basetolerance, 0, TopLeft)) {

				PraxTT("Fläche zum Signieren gefunden.", "4 0")
				X	:= ok.1.1 ;+ 5
				Y	:= ok.1.2 ;+ 5
				W	:= ok.1.3
				H	:= ok.1.4
				;X += W//2
				;Y += H//2
				Comment := ok.1.5
				;MouseMove, % X, % Y, 0
				;SciTEOutput("SBereich: x" X " y" Y " w" (X + Addendum.PDF.SignatureWidth) " h" (Y + Addendum.PDF.SignatureHeight))
				MouseMove, % X, % Y
				;MouseClickDrag, Left, % X, % Y, % (X + Addendum.PDF.SignatureWidth), % (Y + Addendum.PDF.SignatureHeight), 0
				SciTEOutput(">" tryCount ":->" X ", " Y)

			} else {

				sleep, 100
				tryCount ++
				basetolerance += 0.1
				If (tryCount < 30)
					goto FindSignatureRange
				else {

					PraxTT("", "off")
					return 0

				}
			}
		;}

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; sichern der akutellen FoxitID
		;----------------------------------------------------------------------------------------------------------------------------------------------;{
			Addendum.PDF.SignaturedID   	:= FoxitID
			Addendum.PDF.RecentlySigned 	:= true
		;}

			CoordMode, Mouse, % CoordModeMouse_before

return 1
}

FoxitReader_SignDoc(hDokSig) {		                       	                	;-- Bearbeiten des 'Dokument signieren' (Sign Document) Dialoges

		; letzte Änderung: 07.11.2021

		static appendix := "ahk_class #32770 ahk_exe FoxitReader.exe"

		Addendum.PDF.RecentlySigned := true
		FoxitID := GetParent(hDokSig)
		PraxTT("Das Fenster 'Dokument signieren' wird gerade bearbeitet....!'", "30 2")   ; {"timeout":30, "zoom":2, "position":"Bottom", "parent":FoxitID}) 	;"30 2")
		ActivateAndWait("ahk_id " hDokSig, 5)

	;{ Felder im Signierfenster auf die in der INI festgelegten Werte einstellen

		; Signieren als ------------------------------------------------------------------------------------------------------------
			ControlFocus 	, ComboBox1                                                                         	, % "ahk_id " hDokSig
			Control, ChooseString, % Addendum.PDF.SignierenAls , ComboBox1                 	, % "ahk_id " hDokSig
			sleep, 30

		; Darstellungstyp ---------------------------------------------------------------------------------------------------------
			ControlFocus 	, ComboBox4                                                                        	, % "ahk_id " hDokSig
			ControlGet, entryNr, FindString, % Addendum.PDF.Darstellungstyp, ComboBox4	, % "ahk_id " hDokSig  	; prüft Feld Signaturvorschau auf hinterlegte Signatur
			If entryNr
				Control, ChooseString, % Addendum.PDF.Darstellungstyp, ComboBox4        	, % "ahk_id " hDokSig
			else
				MsgBox, 4144, Addendum für AlbisOnWindows, % "Der gewünschte Darstellungstyp "
																						. 	Addendum.PDF.Darstellungstyp "`n"
																						. 	"ist nicht vorhanden"
			Sleep, 30

		; Ort: ----------------------------------------------------------------------------------------------------------------------
			ControlFocus 	, Edit2			                                                         	            	, % "ahk_id " hDokSig
			ControlSetText	, Edit2,  % Addendum.PDF.Ort                                                	, % "ahk_id " hDokSig  ; JEE_StrUtf8BytesToText(Addendum.PDF.Ort)
			ControlSend		, Edit2,  {Tab}	                                                                    	, % "ahk_id " hDokSig
			sleep, 30

		; Grund: -------------------------------------------------------------------------------------------------------------------
			ControlFocus 	, Edit3		                                                                     			, % "ahk_id " hDokSig
			ControlSetText	, Edit3, % Addendum.PDF.Grund                                            	, % "ahk_id " hDokSig
			ControlSend		, Edit3, {Tab}	                                                                    	, % "ahk_id " hDokSig
			sleep, 30

		; nach der Signierung sperren: ----------------------------------------------------------------------------------------
			If Addendum.PDF.DokumentSperren
				VerifiedCheck("Button4","","", hDokSig)
			sleep, 30
	;}

	;{ Signaturfenster schliessen
		SetTimer, PDFHelpCloseSaveDialogs, 100
		while isWindow(hDokSig) {
			If VerifiedClick("Button5", hDokSig)
				break
			else If (A_Index > 10)
				MsgBox, 4144, % "Addendum für AlbisOnWindows",  %	"Das Eintragen des Kennwortes hat nicht funktioniert.`n"
																								.	"Bitte tragen Sie es bitte manuell ein!`n"
																								.	"Drücken Sie danach bitte erst auf Ok."
			sleep, 50
		}
	;}

	;{ statistische Daten erfassen und speichern

		; Ermitteln der Seitenzahl
			PdfPages := FoxitReader_GetPages(FoxitID)                                     ; Seitenzahl des Dokumentes ermitteln (Statistik)
			If !RegExMatch(PdfPages.Max, "\d+")
				PdfPages.Max := 1

		; Signaturzähler erhöhen, Zähler sichern und kurz anzeigen
			Addendum.PDF.SignatureCount ++
			Addendum.PDF.SignaturePages += PdfPages.Max
			IniWrite, % Addendum.PDF.SignatureCount	, % AddendumDir "\Addendum.ini", % "ScanPool", % "SignatureCount"
			IniWrite, % Addendum.PDF.SignaturePages	, % AddendumDir "\Addendum.ini", % "ScanPool", % "SignaturePages"

		; Zählerstände per ToolTip einblenden
			foxitPos     	:= GetWindowSpot(WinExist("ahk_class classFoxitReader"))
			foxitAFXPos 	:= Controls("AfxFrameOrView140su1", "ControlPos", "ahk_class classFoxitReader")
			ToolTip, %	"Signature Nr: "   	Addendum.PDF.SignatureCount "`n"
						. 	"Seitenzahl: "     	PdfPages.Max                          	"`n"
						. 	"ges. Seiten: "    	Addendum.PDF.SignaturePages , foxitPos.X+foxitAFXPos.X, foxitPos.Y+foxitAFXPos.Y, 10

	;}

	; Dateidialog Routinen starten
		SetTimer, PDFNotRecentlySigned		, -20000
		PraxTT("'", "off")

return

PDFHelpCloseSaveDialogs:            	;{ - Notfalllösung für die immer noch unsichere Dialogerkennung
	If !WinExist("Speichern unter bestätigen " appendix) && (hwnd := WinExist("Speichern unter " appendix, "Speichern")) {
		SetTimer, PDFHelpCloseSaveDialogs, Off
		VerifiedClick("Speichern", hwnd)                                                                 	; Speichern Button drücken
		WinWait, % "Speichern unter bestätigen " appendix,, 4
	}
	If WinExist("Speichern unter bestätigen " appendix) {
		VerifiedClick("Ja", "Speichern unter bestätigen " appendix,,, true)    	; mit 'Ja' bestätigen
		If !WinExist("Speichern unter bestätigen " appendix) && !WinExist("Speichern unter " appendix) {
			Addendum.PDF.RecentlySigned := false
			SetTimer, PDFHelpCloseSaveDialogs,  Off
		}
	}
return ;}

PDFNotRecentlySigned: ;{
	Addendum.PDF.RecentlySigned	:= false
	SetTimer, PDFHelpCloseSaveDialogs,  Off
return ;}
}

JEE_StrUtf8BytesToText(vUtf8) {                                                    	;-- wandelt UTF8Bytes in Text (ini Dateien)
	if A_IsUnicode	{
		VarSetCapacity(vUtf8X, StrPut(vUtf8, "CP0"))
		StrPut(vUtf8, &vUtf8X, "CP0")
	return StrGet(&vUtf8X, "UTF-8")
	} 	else
		return StrGet(&vUtf8, "UTF-8")
}


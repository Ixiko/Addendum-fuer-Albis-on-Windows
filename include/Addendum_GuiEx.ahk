;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                            	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                                            	by Ixiko started in September 2017 - last change 07.10.2019 - this file runs under Lexiko's GNU Licence
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

PIC_GUI(GuiName, byref File, GDIx, GDIy , GDIw, GDIh, WTitle:="") {

				;global GGhdc
			If !pToken := Gdip_Startup()
			{
			   MsgBox, 0x40048, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
			   ExitApp
			}
			; Create a layered window (+E0x80000 : must be used for UpdateLayeredWindow to work!) that is always on top (+AlwaysOnTop), has no taskbar entry or caption
			Gui, %GuiName%:-Caption +E0x80000 +LastFound +AlwaysOnTop +HWNDhwnd1
			Gui, %GuiName%:+Owner
			Gui, %GuiName%: Show, Center , %WTitle%		; x%GDIx% y%GDIy%

			pBitmap1 := Gdip_CreateBitmapFromFile(file)
			; Check to ensure we actually got a bitmap from the file, in case the file was corrupt or some other error occured
			If !pBitmap1 {
				MsgBox, 0x40048, File loading error!, Could not load '%file%'
				ExitApp
			}
			;die Grafikdatei wird hinzugeladen
			IWidth := Gdip_GetImageWidth(pBitmap1), IHeight := Gdip_GetImageHeight(pBitmap1)
			hbm := CreateDIBSection(GDIw, GDIh)
			GGhdc := CreateCompatibleDC()
			obm := SelectObject(GGhdc, hbm)
			G1 := Gdip_GraphicsFromHDC(GGhdc)
			Gdip_SetInterpolationMode(G, 7)
			Gdip_DrawImage(G1, pBitmap1, 0, 0, GDIw, GDIh, 0, 0, GDIw, GDIh)
			UpdateLayeredWindow(hwnd1, GGhdc, GDIx, GDIy, GDIw, GDIh)

			SelectObject(hdc, obm)
			DeleteObject(hbm)
			DeleteDC(hdc)
			Gdip_DeleteGraphics(G1)

return hwnd1
}

GeistDerWarteZeit() {



}

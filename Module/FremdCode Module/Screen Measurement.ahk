CoordMode, Mouse, Screen ; Set to screen mode so you get accurate screen cordinates
Gui, 1:+AlwaysOnTop +ToolWindow -caption
Gui, 2:+AlwaysOnTop +ToolWindow -caption
gui, 1:color, FFFEAA ;Color of Ruler Box
gui, 2:color, FFFEAA ;Color of Dimentions Box
Gui 1:+LastFound

gui, 2:add, text, vdiment center w55 h20a
gui, 2:add, button, ghidewindow x+5 h20 w20, X
return

#s:: ;Currently Press WIN + S to activate ruler. Then left click to stop ruler. Press X to hide ruler overlay. 
MouseGetPos, origx, origy ;Get the Original Position of the mouse to calculate width and height

loop {
	; If you detect a left button click break the loop.
	KeyIsDown := GetKeyState("lbutton")
	if (keyisdown) 
		break 
	
	; Get teh current position of the mouse to calculate width and height
	MouseGetPos, curx,cury 
	
	;Calculate the distance from original mouse position
	guiw:=curx - origx
	guih:=cury - origy
	
	;update the dimentions
	gui, 2:default
	GuiControl,,diment,%guiw% x %guih%
	
	;psition the dimentions window left aligned to the selection box
	curxcount:=curx-100
	curycount:=cury+10
	gui, 2:show, x%curxcount% y%curycount%
	
	;Update the size of the ruler window
	gui, 1:show, x%origx% y%origy% w%guiw% h%guih%, Sizer
	WinSet, Transparent, 100, Sizer ; Make the window a little bit transparent.
	sleep, 10
	}
return

hidewindow:
gui, 2:hide
gui, 1:hide
return

esc::
GuiEscape:
exitapp
#NoEnv
#Persistent
#SingleInstance, Force
;#InstallKeybdHook
#Warn All, StdOut
SetTitleMatchMode, 2		;Fast is default
DetectHiddenWindows, Off	;Off is default
CoordMode, Mouse, Screen
CoordMode, Pixel, Screen
CoordMode, ToolTip, Screen
CoordMode, Caret, Screen
CoordMode, Menu, Screen
SetKeyDelay, -1
SetBatchLines, -1
SetWinDelay, -1
SetControlDelay, -1
SendMode, Input
FileEncoding, UTF-8

Größe:= 170								;cm
Gewicht:= 80								;kg
Geschlecht:= "w"							;(w)eiblich, (m)ännlich
PatientenHb:= ""							;

Hbmmol2gl_Faktor:= 18.0182
SollHb_M:= 10.9							;mmol/dl
SollHb_F:= 9.6								;mmol/dl
Gewicht_klein:= 35						;in kg
Reserveeisen_klein:= 15				;mg/kg

Hinweis = (
Berechnung des Eisenbedarfs nach Ganzoni
Bei einem Körpergewicht unter 35 kg wird das Reserveeisen mit 15 mg/kg KG berechnet,
ab einem Körpergewicht von über 35 kg wird ein Reserveeisen von pauschal 500 mg angenommen.

Formel: Gesamteisendefizit (mg) = [Soll-Hb – Patienten-Hb (g/dl)] x Körpergewicht (kg) x 2,4 + Reserveeisen (mg) (2,4 = Eisengehalt der Hämoglobins (3,49 mg/g) x Blutvolumen pro kg KG (0,07 l/kg))

Beim Körpergewicht sollte die fettfreie Körpermasse angenommen werden!
)

Gui, MedCalc: new
Gui, MedCalc: Add, Tab,


#SingleInstance Force
#NoEnv
SetWorkingDir %A_ScriptDir%
SetBatchLines -1
#MaxThreadsBuffer On




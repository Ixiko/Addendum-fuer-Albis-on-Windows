
#include ini.ahk
#include AddendumFunctions.ahk
#include Gui\PraxTT.ahk




PraxTTOff:                                    	;{ 	PraxTT - ToolTip im Addendum-Design - Timerlabel zum Ausblenden des Gui
	AnimateWindow(hPraxTT, PraxTTDuration, BLEND)
	Gui, PraxTT: Destroy
return
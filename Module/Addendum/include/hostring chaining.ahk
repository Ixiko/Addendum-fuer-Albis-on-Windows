
;   ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ 	⌨
;    ☆                                                                                                                                                                                                   	 ☆
;    ☆                                                                     H◯TSTRING ⚯ CHAINS                                                                                         	 ☆
;   ⌨                                                                                                                                                                                                 	⌨
;    ☆                                                                                                                                                                                                   	 ☆
;    ☆                                                                                                                                                                                                   	 ☆
;   ⌨                                                                                                                                                                                                   	⌨
;    ☆                                                                                                                                                                                                   	 ☆
;    ☆                                                                                                                                                                                                   	 ☆
;   ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆ ⌨ ☆ ☆	⌨

;	ℋ೦ፐႽТℛᏆΝᏀ
;	H◯TSTRING ⚯ CHAINS
;
;

CoordMode, ToolTip, Screen

family := { "G"  	: {"rpl" : "Groß"  	, "childs": {"m"	: "mutter", "v":"vater", "t":"tante", "o":"onkel", "ni":"nichte", "ne":"neffe"}}
				, "Ur" 	: {"rpl" : "Urgroß" 	, "childs": {"m"	: "mutter", "v":"vater", "t":"tante", "o":"onkel", "ni":"nichte", "ne":"neffe"}}
				, "H"    	: {"rpl" : "Halb"   	, "childs": {"b" 	:"bruder", "s":"schwester"}}
				, "E"    	: {"rpl" : "Enkel"   	, "childs": {"s" 	:"sohn", "t":"tochter"}}
				, "Sc" 	: {"rpl" : "Schw"		, "childs": {"a"	:"wager", "ä":"wägerin", "e":"ester"
																	, 	 "i"	: {"rpl" : "ieger", "childs" : {"m":"mutter", "v":"vater", "s":"sohn", "t":"tochter"}}}}}

stopKeys := ["Left", "Up", "Down", "Right","End", "Home", "Delete", "LButton", "MButton", "RButton", "Space", "Enter"]
;


fn_stopnext := Func("ChildHotstrings")
For each, key in stopKeys
	Hotkey, % "~" key , % fn_stopnext

ParentHotstrings(family)

return

^!ö::Reload

ParentHotstrings(allChains:="") {                                     	; starts the chain with the word beginnings

	global parentChains, wasSend
	global mainChainStatus := false


	; store allChains object in parentChains
		If !isObject(parentChains) && isObject(allChains)
			parentChains := allChains

	; function was called with object
		mainChainStatus := allChains ="ON" || IsObject(allChains) ? true : false
		For abbreviation, chain in parentChains
			Hotstring(":XC*:" abbreviation, Func("ChildHotstrings").Bind(chain.childs, chain.rpl), mainChainStatus )

		wasSend := RegExReplace(wasSend, "^.*?[\n]")
		wasSend := "base hotstrings: " (mainChainStatus ? "On" : "Off") "`n" wasSend
		ToolTip, % wasSend, 4200, 1,15

return
}
; Schwiegerester tochter
ChildHotstrings(newchain:="", parentReplacement:="") {         	; starts all further chains and stops the execution of the word start chain

	/* 	description

			This function is intended for the quick input of compound words in German. The longest word so far that has made it into the
			official language  has 67 letters  ("Grundstücksverkehrsgenehmigungszuständigkeitsübertragungsverordnung").  However, we
			are not talking about the longest words here, but about words that start with the same letters and branch out more and more
			later.  Noticeable hotstring abbreviations can hardly be invented for words similar to word stems.  Thus, the complete word is
			not written here, but only the part that is not found in any of the other words. Now you only have to press the 1 or 2 keys that
			are unique for the next chain.Chain link for chain link.

	*/

	; variables
		global parentChains, wasSend
		static lastchain, availableHS := {}

	; sends replacement text now
		wasSend .= parentReplacement ? parentReplacement "`n" : ""
		SendInput, % "{RAW}" parentReplacement
		wasSend .= !IsObject(newchain) ? newchain "`n" : ""
		ToolTip, % wasSend, 4200, 1,15

	; ―――――――――――――――――――――――――――――
	; IF THE FUNCTION IS CALLED INCORRECTLY, DO NOTHING
	; ―――――――――――――――――――――――――――――
		If (!isObject(newchain) && !isObject(lastchain)) {
			wasSend .= "  no objects`n"
			ToolTip, % wasSend, 4200, 1,15
			ParentHotstrings(parentChains)
			return
		}

	; ―――――――――――――――――――――――――――――
	; INITIAL CHAIN OF HOTSTRINGS  [INACTIVATION]
	;                                   (first child chain was invoked)
	; ―――――――――――――――――――――――――――――
		if (newchain && isObject(newchain) && !isObject(lastchain)) {
			lastchain := newchain
			ParentHotstrings()
		}

	; ―――――――――――――――――――――――――――――
	; INITIAL CHAIN OF HOTSTRINGS  [GRAND RESET]
	; (last chain replacement was called or a breaking hotkey was pressed)
	; ―――――――――――――――――――――――――――――
		else If (!isObject(newchain) && isObject(lastchain)) {

			SendInput, % "{RAW} "                                                                            	; sends Space

			For abbreviation, replacement in lastchain.childs                                     	; turns last chain hotstrings off
				Hotstring(":XC*:" abbreviation, "", 0)                                                  	; inactivates one hotstring
			For replacement, abbreviation in availableHS                                          	; delete all hotstrings for safety
				try Hotstring(":XC*:" abbreviation, "", 0)

			lastchain := newchain                                                                              	; delete lastchain object
			wasSend .= "grand reset for inital chain inactivated`n"
			ToolTip, % wasSend, 4200, 1,15
			ParentHotstrings(parentChains)                                                                 	; activates the hotstring initial chain
			return
		}

	; ―――――――――――――――――――――――――――――
	; PREVIOUS CHAIN OF HOTSTRINGS  [INACTIVATION]
	; (go on with activation of newchain hotstrings)
	; ―――――――――――――――――――――――――――――
		else if (isObject(newchain) && isObject(lastchain)) {

			For abbreviation, replacement in lastchain.childs
				 hotstring(":XC*:" abbreviation, "", 0)
			For replacement, abbreviation in availableHS                                          	; delete all hotstrings for safety
				try Hotstring(":XC*:" abbreviation, "", 0)

				wasSend .= "previous chain inactivated`n"
		}


    ; ―――――――――――――――――――――――――――――
    ; ACTIVATES A NEW HOTSTRING CHAIN
    ; ―――――――――――――――――――――――――――――
		If IsObject(newchain) {

			lastchain := newchain
			wasSend .= "newchain childs count: " newchain.Count() "`n"
			For abbreviation, replacement in newchain {


				If !isObject(replacement) {

					wasSend .= (!isObject(replacement)  ? abbreviation ": " replacement "`n" : "")
					If !availableHS.haskey(replacement)
						availableHS[replacement] := abbreviation
					Hotstring(":XC*:" abbreviation, Func("ChildHotstrings").Bind("⛨", replacement), 1)

				}
				else {

					wasSend .= abbreviation ": " replacement.rpl "  has " replacement.childs.Count() " childs`n"

					If !availableHS.haskey(replacement.rpl)
						availableHS[replacement.rpl] := abbreviation
					Hotstring(":XC*:" abbreviation, Func("ChildHotstrings").Bind(replacement.childs, replacement.rpl), 1)

				}
			}

		}

		ToolTip, % wasSend, 4200, 1,15
		ToolTip, % t, 500, 1, 14

}

/*


*/
				;~ Hotstring((!isObject(replacement) ? ":C*:" : ":XC*:") abbreviation, Func("ChildHotstrings", "", replacement))

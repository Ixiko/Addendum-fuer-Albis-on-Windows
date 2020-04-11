;------------------------------------------------------------------------------
; http://www.autohotkey.com/forum/topic24234.html
; ACC.ahk Standard Library
; by Sean
;------------------------------------------------------------------------------

ACC_Init(){
	COM_Init()
	If Not	DllCall("GetModuleHandle", "str", "oleacc")
	Return	DllCall("LoadLibrary"    , "str", "oleacc")
}
ACC_Term(){
	COM_Term()
	If   hModule :=	DllCall("GetModuleHandle", "str", "oleacc")
	Return		DllCall("FreeLibrary"    , "Uint", hModule)
}
ACC_AccessibleChildren(pacc, ByRef varChildren){
	VarSetCapacity(varChildren,16*(cChildren:=acc_ChildCount(pacc)),0)
	DllCall("oleacc\AccessibleChildren", "Uint", pacc, "int", 0, "int", cChildren, "Uint", &varChildren, "intP", cObtained)
	Return	cObtained
}
ACC_AccessibleObjectFromEvent(hWnd, idObject, idChild, ByRef _idChild_){
	VarSetCapacity(varChild,16,0)
	DllCall("oleacc\AccessibleObjectFromEvent", "Uint", hWnd, "Uint", idObject, "Uint", idChild, "UintP", pacc, "Uint", &varChild)
	_idChild_ := NumGet(varChild,8)
	Return	pacc
}
ACC_AccessibleObjectFromPoint(x = "", y = "", ByRef _idChild_ = ""){
	VarSetCapacity(varChild,16,0)
	x<>""&&y<>"" ? pt:=x&0xFFFFFFFF|y<<32 : DllCall("GetCursorPos", "int64P", pt)
	DllCall("oleacc\AccessibleObjectFromPoint", "int64", pt, "UintP", pacc, "Uint", &varChild)
	_idChild_ := NumGet(varChild,8)
	Return	pacc
}
ACC_AccessibleObjectFromWindow(hWnd = "", idObject = -4){
	If	(hWnd = "")
		hWnd := DllCall("GetForegroundWindow")
	DllCall("oleacc\AccessibleObjectFromWindow", "Uint", hWnd, "Uint", idObject, "Uint", COM_GUID4String(IID_IAccessible,"{618736E0-3C3D-11CF-810C-00AA00389B71}"), "UintP", pacc)
	Return	pacc
}
ACC_WindowFromAccessibleObject(pacc){
	DllCall("oleacc\WindowFromAccessibleObject", "Uint", pacc, "UintP", hWnd)
	Return	hWnd
}
ACC_ObjectFromWindow(hWnd, idObject = -4) {
	Acc_Init()
	If	DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", idObject&=0xFFFFFFFF, "Ptr", -VarSetCapacity(IID,16)+NumPut(idObject==0xFFFFFFF0?0x46000000000000C0:0x719B3800AA000C81,NumPut(idObject==0xFFFFFFF0?0x0000000000020400:0x11CF3C3D618736E0,IID,"Int64"),"Int64"), "Ptr*", pacc)=0
	Return	ComObjEnwrap(9,pacc,1)
}
ACC_GetRoleText(nRole){
	nSize := DllCall("oleacc\GetRoleTextA", "Uint", nRole, "Uint", 0, "Uint", 0)
	VarSetCapacity(sRole, nSize)
	DllCall("oleacc\GetRoleTextA", "Uint", nRole, "str", sRole, "Uint", nSize+1)
	Return	sRole
}
ACC_GetStateText(nState){
	nSize := DllCall("oleacc\GetStateTextA", "Uint", nState, "Uint", 0, "Uint", 0)
	VarSetCapacity(sState, nSize)
	DllCall("oleacc\GetStateTextA", "Uint", nState, "str", sState, "Uint", nSize+1)
	Return	sState
}
Acc_Get(Cmd, ChildPath="", ChildID=0, WinTitle="", WinText="", ExcludeTitle="", ExcludeText="") {
	static properties := {Action:"DefaultAction", DoAction:"DoDefaultAction", Keyboard:"KeyboardShortcut"}
	AccObj :=   IsObject(WinTitle)? WinTitle
			:   Acc_ObjectFromWindow( WinExist(WinTitle, WinText, ExcludeTitle, ExcludeText), 0 )
	if ComObjType(AccObj, "Name") != "IAccessible"
		ErrorLevel := "Could not access an IAccessible Object"
	else {
		StringReplace, ChildPath, ChildPath, _, %A_Space%, All
		AccError:=Acc_Error(), Acc_Error(true)
		Loop Parse, ChildPath, ., %A_Space%
			try {
				if A_LoopField is digit
					Children:=Acc_Children(AccObj), m2:=A_LoopField ; mimic "m2" output in else-statement
				else
					RegExMatch(A_LoopField, "(\D*)(\d*)", m), Children:=Acc_ChildrenByRole(AccObj, m1), m2:=(m2?m2:1)
				if Not Children.HasKey(m2)
					throw
				AccObj := Children[m2]
			} catch {
				ErrorLevel:="Cannot access ChildPath Item #" A_Index " -> " A_LoopField, Acc_Error(AccError)
				if Acc_Error()
					throw Exception("Cannot access ChildPath Item", -1, "Item #" A_Index " -> " A_LoopField)
				return
			}
		Acc_Error(AccError)
		StringReplace, Cmd, Cmd, %A_Space%, , All
		properties.HasKey(Cmd)? Cmd:=properties[Cmd]:
		try {
			if (Cmd = "Location")
				AccObj.accLocation(ComObj(0x4003,&x:=0), ComObj(0x4003,&y:=0), ComObj(0x4003,&w:=0), ComObj(0x4003,&h:=0), ChildId)
			  , ret_val := "x" NumGet(x,0,"int") " y" NumGet(y,0,"int") " w" NumGet(w,0,"int") " h" NumGet(h,0,"int")
			else if (Cmd = "Object")
				ret_val := AccObj
			else if Cmd in Role,State
				ret_val := Acc_%Cmd%(AccObj, ChildID+0)
			else if Cmd in ChildCount,Selection,Focus
				ret_val := AccObj["acc" Cmd]
			else
				ret_val := AccObj["acc" Cmd](ChildID+0)
		} catch {
			ErrorLevel := """" Cmd """ Cmd Not Implemented"
			if Acc_Error()
				throw Exception("Cmd Not Implemented", -1, Cmd)
			return
		}
		return ret_val, ErrorLevel:=0
	}
	if Acc_Error()
		throw Exception(ErrorLevel,-1)
}
Acc_Error(p="") {
	static setting:=0
	return p=""?setting:setting:=p
}
Acc_Children(Acc) {
	Acc_Init()
	cChildren:=Acc.accChildCount, Children:=[]
	if DllCall("oleacc\AccessibleChildren", "Ptr", ComObjValue(Acc), "Int", 0, "Int", cChildren, "Ptr", VarSetCapacity(varChildren,cChildren*(8+2*A_PtrSize),0)*0+&varChildren, "Int*", cChildren)=0 {
		Loop %cChildren%
			i:=(A_Index-1)*(A_PtrSize*2+8)+8, child:=NumGet(varChildren,i), Children.Insert(NumGet(varChildren,i-8)=3?child:Acc_Query(child)), ObjRelease(child)
		return Children
	}
	error:=Exception("",-1)
	MsgBox, 262420, Acc_Children Failed, % "File:  " error.file "`nLine: " error.line "`n`nContinue Script?"
	IfMsgBox, No
		ExitApp
}
Acc_ChildrenByRole(Acc, Role) {
	if ComObjType(Acc,"Name")!="IAccessible"
		ErrorLevel := "Invalid IAccessible Object"
	else {
		Acc_Init(), cChildren:=Acc.accChildCount, Children:=[]
		if DllCall("oleacc\AccessibleChildren", "Ptr",ComObjValue(Acc), "Int",0, "Int",cChildren, "Ptr",VarSetCapacity(varChildren,cChildren*(8+2*A_PtrSize),0)*0+&varChildren, "Int*",cChildren)=0 {
			Loop %cChildren% {
				i:=(A_Index-1)*(A_PtrSize*2+8)+8, child:=NumGet(varChildren,i)
				if NumGet(varChildren,i-8)=9
					AccChild:=Acc_Query(child), ObjRelease(child), Acc_Role(AccChild)=Role?Children.Insert(AccChild):
				else
					Acc_Role(Acc, child)=Role?Children.Insert(child):
			}
			return Children.MaxIndex()?Children:, ErrorLevel:=0
		} else
			ErrorLevel := "AccessibleChildren DllCall Failed"
	}
	if Acc_Error()
		throw Exception(ErrorLevel,-1)
}

; pCallback := RegisterCallback("ACC_WinEventProc")
ACC_SetWinEventHook(eventMin, eventMax, pCallback){
	Return	DllCall("SetWinEventHook", "Uint", eventMin, "Uint", eventMax, "Uint", 0, "Uint", pCallback, "Uint", 0, "Uint", 0, "Uint", 0)
}
ACC_UnhookWinEvent(hHook){
	Return	DllCall("UnhookWinEvent", "Uint", hHook)
}
ACC_WinEventProc(hHook, event, hWnd, idObject, idChild, eventThread, eventTime){
	Critical
	pacc := ACC_AccessibleObjectFromEvent(hWnd, idObject, idChild, _idChild_)
/*	Add custom codes here!
	...
*/
	COM_Release(pacc)
}
acc_Query(pacc, bunk = ""){
	If	DllCall(NumGet(NumGet(1*pacc)+0), "Uint", pacc, "Uint", COM_GUID4String(IID_IAccessible,bunk ? "{00020404-0000-0000-C000-000000000046}" : "{618736E0-3C3D-11CF-810C-00AA00389B71}"), "UintP", pobj)=0
		DllCall(NumGet(NumGet(1*pacc)+8), "Uint", pacc), pacc:=pobj
	Return	pacc
}
acc_Parent(pacc){
	If	DllCall(NumGet(NumGet(1*pacc)+28), "Uint", pacc, "UintP", paccParent)=0 && paccParent
	Return	acc_Query(paccParent)
}
acc_ChildCount(pacc){
	If	DllCall(NumGet(NumGet(1*pacc)+32), "Uint", pacc, "UintP", cChildren)=0
	Return	cChildren
}
acc_Child(pacc, idChild){
	If	DllCall(NumGet(NumGet(1*pacc)+36), "Uint", pacc, "int64", 3, "int64", idChild, "UintP", paccChild)=0 && paccChild
	Return	acc_Query(paccChild)
}
acc_Name(pacc, idChild = 0){
	If	DllCall(NumGet(NumGet(1*pacc)+40), "Uint", pacc, "int64", 3, "int64", idChild, "UintP", pName)=0 && pName
	Return	COM_Ansi4Unicode(pName) . SubStr(COM_SysFreeString(pName),1,0)
}
acc_Value(pacc, idChild = 0){
	If	DllCall(NumGet(NumGet(1*pacc)+44), "Uint", pacc, "int64", 3, "int64", idChild, "UintP", pValue)=0 && pValue
	Return	COM_Ansi4Unicode(pValue) . SubStr(COM_SysFreeString(pValue),1,0)
}
acc_Description(pacc, idChild = 0){
	If	DllCall(NumGet(NumGet(1*pacc)+48), "Uint", pacc, "int64", 3, "int64", idChild, "UintP", pDescription)=0 && pDescription
	Return	COM_Ansi4Unicode(pDescription) . SubStr(COM_SysFreeString(pDescription),1,0)
}
acc_Role(pacc, idChild = 0){
	VarSetCapacity(var,16,0)
	If	DllCall(NumGet(NumGet(1*pacc)+52), "Uint", pacc, "int64", 3, "int64", idChild, "Uint", &var)=0
	Return	ACC_GetRoleText(NumGet(var,8))
}
acc_State(pacc, idChild = 0){
/*
	STATE_SYSTEM_NORMAL	:= 0
	STATE_SYSTEM_UNAVAILABLE:= 0x1
	STATE_SYSTEM_SELECTED	:= 0x2
	STATE_SYSTEM_FOCUSED	:= 0x4
	STATE_SYSTEM_PRESSED	:= 0x8
	STATE_SYSTEM_CHECKED	:= 0x10
	STATE_SYSTEM_MIXED	:= 0x20
	STATE_SYSTEM_READONLY	:= 0x40
	STATE_SYSTEM_HOTTRACKED	:= 0x80
	STATE_SYSTEM_DEFAULT	:= 0x100
	STATE_SYSTEM_EXPANDED	:= 0x200
	STATE_SYSTEM_COLLAPSED	:= 0x400
	STATE_SYSTEM_BUSY	:= 0x800
	STATE_SYSTEM_FLOATING	:= 0x1000
	STATE_SYSTEM_MARQUEED	:= 0x2000
	STATE_SYSTEM_ANIMATED	:= 0x4000
	STATE_SYSTEM_INVISIBLE	:= 0x8000
	STATE_SYSTEM_OFFSCREEN	:= 0x10000
	STATE_SYSTEM_SIZEABLE	:= 0x20000
	STATE_SYSTEM_MOVEABLE	:= 0x40000
	STATE_SYSTEM_SELFVOICING:= 0x80000
	STATE_SYSTEM_FOCUSABLE	:= 0x100000
	STATE_SYSTEM_SELECTABLE	:= 0x200000
	STATE_SYSTEM_LINKED	:= 0x400000
	STATE_SYSTEM_TRAVERSED	:= 0x800000
	STATE_SYSTEM_MULTISELECTABLE	:= 0x1000000
	STATE_SYSTEM_EXTSELECTABLE	:= 0x2000000
	STATE_SYSTEM_ALERT_LOW	:= 0x4000000
	STATE_SYSTEM_ALERT_MEDIUM:=0x8000000
	STATE_SYSTEM_ALERT_HIGH	:= 0x10000000
	STATE_SYSTEM_PROTECTED	:= 0x20000000
	STATE_SYSTEM_HASPOPUP	:= 0x40000000
*/
	VarSetCapacity(var,16,0)
	If	DllCall(NumGet(NumGet(1*pacc)+56), "Uint", pacc, "int64", 3, "int64", idChild, "Uint", &var)=0
	Return	ACC_GetStateText(nState:=NumGet(var,8)) . "`t(" . acc_Hex(nState) . ")"
}
acc_Help(pacc, idChild = 0){
	If	DllCall(NumGet(NumGet(1*pacc)+60), "Uint", pacc, "int64", 3, "int64", idChild, "UintP", pHelp)=0 && pHelp
	Return	COM_Ansi4Unicode(pHelp) . SubStr(COM_SysFreeString(pHelp),1,0)
}
acc_HelpTopic(pacc, idChild = 0){
	If	DllCall(NumGet(NumGet(1*pacc)+64), "Uint", pacc, "UintP", pHelpFile, "int64", 3, "int64", idChild, "intP", idTopic)=0 && pHelpFile
	Return	COM_Ansi4Unicode(pHelpFile) . SubStr(COM_SysFreeString(pHelpFile),1,0) . "|" . idTopic
}
acc_KeyboardShortcut(pacc, idChild = 0){
	If	DllCall(NumGet(NumGet(1*pacc)+68), "Uint", pacc, "int64", 3, "int64", idChild, "UintP", pKeyboardShortcut)=0 && pKeyboardShortcut
	Return	COM_Ansi4Unicode(pKeyboardShortcut) . SubStr(COM_SysFreeString(pKeyboardShortcut),1,0)
}
acc_Focus(pacc){
	VarSetCapacity(var,16,0)
	If	DllCall(NumGet(NumGet(1*pacc)+72), "Uint", pacc, "Uint", &var)=0
	Return	(vtType:=NumGet(var,0,"Ushort"))=3 ? NumGet(var,8) : vtType=9 ? "+" . acc_Query(NumGet(var,8)) : ""
}
acc_Selection(pacc){
	VarSetCapacity(var,16,0)
	If	DllCall(NumGet(NumGet(1*pacc)+76), "Uint", pacc, "Uint", &var)=0
	Return	(vtType:=NumGet(var,0,"Ushort"))=3 ? NumGet(var,8) : vtType=9 ? "+" . acc_Query(NumGet(var,8)) : vtType=13 ? " " . acc_Query(NumGet(var,8),1) : ""
}
acc_DefaultAction(pacc, idChild = 0){
	If	DllCall(NumGet(NumGet(1*pacc)+80), "Uint", pacc, "int64", 3, "int64", idChild, "UintP", pDefaultAction)=0 && pDefaultAction
	Return	COM_Ansi4Unicode(pDefaultAction) . SubStr(COM_SysFreeString(pDefaultAction),1,0)
}
acc_Select(pacc, idChild = 0, nFlags = 3){
/*
	SELFLAG_NONE		:= 0x0
	SELFLAG_TAKEFOCUS	:= 0x1
	SELFLAG_TAKESELECTION	:= 0x2
	SELFLAG_EXTENDSELECTION	:= 0x4
	SELFLAG_ADDSELECTION	:= 0x8
	SELFLAG_REMOVESELECTION	:= 0x10
*/
	Return	DllCall(NumGet(NumGet(1*pacc)+84), "Uint", pacc, "int", nFlags, "int64", 3, "int64", idChild)
}
acc_Location(pacc, idChild = 0, ByRef l = "", ByRef t = "", ByRef w = "", ByRef h = ""){
	If	DllCall(NumGet(NumGet(1*pacc)+88), "Uint", pacc, "intP", l, "intP", t, "intP", w, "intP", h, "int64", 3, "int64", idChild)=0
	Return	l . "," . t . "," . w . "," . h
}
acc_Navigate(pacc, idChild = 0, nDir = 7){
/*
	NAVDIR_UP	:= 1
	NAVDIR_DOWN	:= 2
	NAVDIR_LEFT	:= 3
	NAVDIR_RIGHT	:= 4
	NAVDIR_NEXT	:= 5
	NAVDIR_PREVIOUS	:= 6
	NAVDIR_FIRSTCHILD:=7
	NAVDIR_LASTCHILD:= 8
*/
	VarSetCapacity(var,16,0)
	If	DllCall(NumGet(NumGet(1*pacc)+92), "Uint", pacc, "int", nDir, "int64", 3, "int64", idChild, "Uint", &var)=0
	Return	(vtType:=NumGet(var,0,"Ushort"))=3 ? NumGet(var,8) : vtType=9 ? "+" . acc_Query(NumGet(var,8)) : ""
}
acc_HitTest(pacc, x = "", y = ""){
	VarSetCapacity(var,16,0)
	x<>""&&y<>"" ? pt:=x&0xFFFFFFFF|y<<32 : DllCall("GetCursorPos", "int64P", pt)
	If	DllCall(NumGet(NumGet(1*pacc)+96), "Uint", pacc, "int64", pt, "Uint", &var)=0
	Return	(vtType:=NumGet(var,0,"Ushort"))=3 ? NumGet(var,8) : vtType=9 ? "+" . acc_Query(NumGet(var,8)) : ""
}
acc_DoDefaultAction(pacc, idChild = 0){
	Return	DllCall(NumGet(NumGet(1*pacc)+100), "Uint", pacc, "int64", 3, "int64", idChild)
}

; May be obsolete!
acc_Name_(pacc, idChild = 0, sName = ""){
	Return	DllCall(NumGet(NumGet(1*pacc)+104), "Uint", pacc, "int64", 3, "int64", idChild, "Uint", pName:=COM_SysAllocString(sName)) . SubStr(COM_SysFreeString(pName),1,0)
}

; May be obsolete!
acc_Value_(pacc, idChild = 0, sValue = ""){
	Return	DllCall(NumGet(NumGet(1*pacc)+108), "Uint", pacc, "int64", 3, "int64", idChild, "Uint", pValue:=COM_SysAllocString(sValue)) . SubStr(COM_SysFreeString(pValue),1,0)
}
acc_Hex(num){
	old := A_FormatInteger
	SetFormat, Integer, H
	num += 0
	SetFormat, Integer, %old%
	Return	num
}
/* Lexikos library stuff */
	/*
	LoadFile(Path [, EXE])

	Loads a script file as a child process and returns an object
	which can be used to call functions or get/set global vars.

	Path:
	The path of the script.
	EXE:
	The path of the AutoHotkey executable (defaults to A_AhkPath).

	Requirements:
	- AutoHotkey v1.1.17+    http://ahkscript.org/download/
	- ObjRegisterActive      http://goo.gl/wZsFLP
	- CreateGUID             http://goo.gl/obfmDc

	Version: 1.0
*/

LoadFile(path, exe:="", exception_level:=-1) {
	ObjRegisterActive(client := {}, guid := CreateGUID())
	code =
    (LTrim
    LoadFile.Serve("%guid%")
    #include %A_LineFile%
    #include %path%
    )
	try {
		exe := """" (exe="" ? A_AhkPath : exe) """"
		exec := ComObjCreate("WScript.Shell").Exec(exe " /ErrorStdOut *")
		exec.StdIn.Write(code)
		exec.StdIn.Close()
		while exec.Status = 0 && !client._proxy
			Sleep 10
		if exec.Status != 0 {
			err := exec.StdErr.ReadAll()
			ex := Exception("Failed to load file", exception_level)
			if RegExMatch(err, "Os)(.*?) \((\d+)\) : ==> (.*?)(?:\s*Specifically: (.*?))?\R?$", m)
				ex.Message .= "`n`nReason:`t" m[3] "`nLine text:`t" m[4] "`nFile:`t" m[1] "`nLine:`t" m[2]
			throw ex
		}
	}
	finally
		ObjRegisterActive(client, "")
	return client._proxy
}

class LoadFile {
	Serve(guid) {
		try {
			client := ComObjActive(guid)
			client._proxy := new this.Proxy
			client := ""
		}
		catch ex {
			stderr := FileOpen("**", "w")
			stderr.Write(format("{} ({}) : ==> {}`n     Specifically: {}"
                , ex.File, ex.Line, ex.Message, ex.Extra))
			stderr.Close()  ; Flush write buffer.
			ExitApp
		}
		; Rather than sleeping in a loop, make the script persistent
		; and then return so that the #included file is auto-executed.
		Hotkey IfWinActive, %guid%
		Hotkey vk07, #Persistent, Off
		#Persistent:
	}
	class Proxy {
		__call(name, args*) {
			if (name != "G")
				return %name%(args*)
		}
		G[name] { ; x.G[name] because x[name] via COM invokes __call.
			get {
				global
				return ( %name% )
			}
			set {
				global
				return ( %name% := value )
			}
		}
		__delete() {
			ExitApp
		}
	}
}

/*
	ObjRegisterActive(Object, CLSID, Flags:=0)

	Registers an object as the active object for a given class ID.
	Requires AutoHotkey v1.1.17+; may crash earlier versions.

	Object:
	Any AutoHotkey object.
	CLSID:
	A GUID or ProgID of your own making.
	Pass an empty string to revoke (unregister) the object.
	Flags:
	One of the following values:
	0 (ACTIVEOBJECT_STRONG)
	1 (ACTIVEOBJECT_WEAK)
	Defaults to 0.

	Related:
	http://goo.gl/KJS4Dp - RegisterActiveObject
	http://goo.gl/no6XAS - ProgID
	http://goo.gl/obfmDc - CreateGUID()
*/
ObjRegisterActive(Object, CLSID, Flags:=0) {
	static cookieJar := {}
	if (!CLSID) {
		if (cookie := cookieJar.Remove(Object)) != ""
			DllCall("oleaut32\RevokeActiveObject", "uint", cookie, "ptr", 0)
		return
	}
	if cookieJar[Object]
		throw Exception("Object is already registered", -1)
	VarSetCapacity(_clsid, 16, 0)
	if (hr := DllCall("ole32\CLSIDFromString", "wstr", CLSID, "ptr", &_clsid)) < 0
		throw Exception("Invalid CLSID", -1, CLSID)
	hr := DllCall("oleaut32\RegisterActiveObject"
        , "ptr", &Object, "ptr", &_clsid, "uint", Flags, "uint*", cookie
        , "uint")
	if hr < 0
		throw Exception(format("Error 0x{:x}", hr), -1)
	cookieJar[Object] := cookie
}

CreateGUID() {
	VarSetCapacity(pguid, 16, 0)
	if !(DllCall("ole32.dll\CoCreateGuid", "ptr", &pguid)) {
		size := VarSetCapacity(sguid, (38 << !!A_IsUnicode) + 1, 0)
		if (DllCall("ole32.dll\StringFromGUID2", "ptr", &pguid, "ptr", &sguid, "int", size))
			return StrGet(&sguid)
	}
	return ""
}
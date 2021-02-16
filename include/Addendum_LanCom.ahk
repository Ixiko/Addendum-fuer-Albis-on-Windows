; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                  Funktionsbibliothek für TCP - LAN Komunikation - benötigt class_TCP-UDP.ahk
;                                                                              	!diese Bibliothek enthält Funktionen für Einstellungen des Addendum Hauptskriptes!
;                                                            	by Ixiko started in September 2017 - last change 21.08.2020 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

admStartServer() {

		global DC_SERV := True
		global DC_CLI := False
		global admServer

	; admServer muss im aufrufenden Skript supergobal gemacht worden sein
		rmsgQueue := Object()

		admServer := new SocketTCP()
		admServer.Bind(["addr_any", 1337])
		admServer.Listen()
		admServer.OnAccept := Func("admOnAccept")

}

admOnAccept() {

		global admServer
		global rmsgQueue
		global DC_SERV
		global DC_CLI

		newTCP  	:= admServer.accept()
		newTCP.SendText("Successful Connection!")
		/*
		if IsObject(newTCP) {

			For key, val in newTCP {

				t.= A_Index ": " key ", values: "
				If IsObject(val) {

					t.= "isOBject with " val.Length() " key`n"
					For valkey, obj in val
						t.= "`t" valkey ", value: " (!IsObject(obj) ? obj : "isobject") "`n"

				} else
					t .= val "`n"
			}

		SciTEOutput(t)
		}
		 */
		;SciTEOutput("ProtocolID: " newTCP.ProtocolID ", newTCPType: " newTCP.SocketType)

		rmsg  	:= StrSplit(newTCP.RecvText(), "|")
		SciTEOutput(newTCP.RecvText())
		SciTEOutput("recv: " rmsg.Count())

		If IsFunc("adm" rmsg.1) {

			fnCall	:= Func("adm" rmsg.1).Bind(rmsg.2, rmsg.3, rmsg.4)
			SetTimer, % fnCall, -0

		} else if InStr(rmsg.1, "answer") {

			; rmsg.2 - Funktionsname welche die Anfrage gestellt hat
			SciTEOutput(rmsg.2 "|" rmsg.3 "|" rmsg.4)
			rmsgQueue.Push(rmsg.2 "|" rmsg.3 "|" rmsg.4 )

		}

		;newTCP.Disconnect()

return
}

admStatus(cmd, more, from) {


	switch cmd {

		case "AddendumExist":
			answer := "answer|true"
		default:
			answer := "answer|unknown command.."

	}

	admSendText(from, answer)

}

admSendText(to, Text) {

	global DC_SERV
	global DC_CLI

	x := new SocketTCP()
	Connected := x.Connect([to, 1337])
	x.SendText(Text)
	x.Disconnect()

	If Connected
		SciTEOutput("connected to: " to " (" Text ")")
	else
		SciTEOutput("can't connected to: " to)

}
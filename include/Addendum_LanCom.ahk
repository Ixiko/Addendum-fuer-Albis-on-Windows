; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                     	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                       Funktionsbibliothek für TCP - LAN Komunikation - benötigt lib\class_socket.ahk
;                  	   by Ixiko started in September 2017 - last change 21.12.2021 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

TCPStartServer()                                                          	{

	global connections

	If !IsObject(connections)
		connections := Object()

	Server := Addendum.LAN.Server.IamServer
	connections.Server := new SocketTCP()
	connections.Server.props := {"name":Server, "ip":Addendum.LAN.Server.ip, "port":Addendum.LAN.Server.port}
	connections.Server.bind("addr_any", 12345) ; Addendum.LAN.admServer
	connections.Server.listen()
	connections.Server.onAccept := Func("TCPServerAccept")
	connections.Server.onDisconnect("TCPServerDisconnect")

}

TCPServerAccept(this)                                                  	{

	global connections

	If !IsObject(connections.Server.clients)
		connections.Server.clients := Array()

	CI := connections.Server.clients.Push(this.accept())
	connections.Server.clients[CI].onrecv := Func("TCPReceive")
	msg := {"from":"Server", "cmd":"tell_props", "txt":CI}
	connections.Server.clients[CI].sendText(JSON.Dump(msg))

}

TCPServerDisconnect(client:="")                                   	{

	global admClients

}

TCPConnectTo(machine, ip, port)                                  	{

	global DC_SERV, DC_CLI
	global connections

	If !IsObject(connections[machine]) {

		connections[machine] := new SocketTCP()

	  ; Verbindung hergestellt
		If (connected := connections[machine].connect(ServerIP, ServerPort)) {
			connections[machine].onrecv := Func("TCPReceive").bind(machine)
		}
	  ; keine Verbindung - Objekt wieder entfernen
		else
			connections[machine] := ""
	}

return connections[machine]
}

TCPReceive(machine, answer)                                       	{                 	; empfängt Netzwerknachrichten

	msg := JSON.Load(answer.recvText(),, "UTF-8")
	SciTEOutput(" [" msg.from "] " msg.cmd ", " msg.txt)
	;~ SciTEOutput(" [LAN] recv " answer.recv())

}

TCPGetStates(cmd, more, from)                                 	{


	switch cmd {

		case "AddendumExist":
			answer := "answer|true"
		default:
			answer := "answer|unknown command.."

	}

	TCPSendTextTo(from, answer)

}

TCPSendTextTo(machine, Text)                                   	{

	global DC_SERV
	global DC_CLI

	Server := Addendum.LAN.Server
	If (compname<>Server.IamServer)
		connection := TCPConnectTo(Server, Server.ip, Server.port)
	msg := JSON.Dump({"to":machine, "txt":Text},, "UTF-8")
	connection.sendText(msg)

}


#include <TCP_V3.au3>
#include <date.au3>
#include <string.au3>
dim $sec=@SEC
dim $min=@MIN
dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR

;ToolTip("SERVER: Creating server...",10,30)
;traytip("TCP Server", "Port open at 88",5)
ConsoleWrite("SERVER: Creating server...")
;	 sleep(3000)
;	 tooltip("")
Global $hServer = _TCP_Server_Create(88); A server. Tadaa!

_TCP_RegisterEvent($hServer, $TCP_NEWCLIENT, "NewClient"); Whooooo! Now, this function (NewClient) get's called when a new client connects to the server.
_TCP_RegisterEvent($hServer, $TCP_DISCONNECT, "Disconnect"); And this,... this will get called when a client disconnects.
_TCP_RegisterEvent($hServer, $TCP_RECEIVE, "Received"); Function "Received" will get called when something is received
While 1

	sleep(1000)
WEnd

Func NewClient($hSocket,$iError); Yo, check this out! It's a $iError parameter! (In case you didn't noticed: It's in every function)
	;local $_TCP_ACTIVECLIENT
     ;traytip("TCP Server", "Client Connected",5)
	 ;ToolTip("SERVER: New client connected."&@CRLF&"Sending this: Bleh!",10,30)
	 ;ConsoleWrite("SERVER: New client connected."&@CRLF&"Sending this: Bleh!")
     ;_TCP_Server_Send("Hi there!"); Sending: "Bleh!" to the new client.
     ;$ip=_TCP_Server_ClientIP($_TCP_ACTIVECLIENT())
	  ;_TCP_Send($hSocket, "Bleh")
	 $ip=_TCP_Server_ClientIP($hSocket)
	 ToolTip("SERVER: New client connected."&$hSocket&@CRLF&" IP is: "&$ip,10,30)
;	 if not ProcessExists("Server_load_sms_2.exe") then 
;		run( @ScriptDir & "\Server_load_sms_2.exe" ) 
;	 EndIf	 
	 ;sleep(3000)
	 ;tooltip("")
EndFunc

Func Disconnect($hSocket,$iError); Damn, we lost a client. Time of death: @Hour & @Min & @Sec :P
     ToolTip("SERVER: Client disconnected.",10,30); Placing a tooltip right under the tooltips of the client.
	 sleep(1000)
	 tooltip("")
	;ConsoleWrite(@CRLF&"SERVER: Client disconnected.")
EndFunc

Func Received($hSocket, $sReceived,$iError); And we also registered this! Our homemade do-it-yourself function gets called when something is received.
     ;ToolTip("Server: We received this: "& $sReceived, 10,30); (and we'll display it)
	 ;ConsoleWrite(@CRLF&"Server Received from client:"& _HexToString($sReceived))
	 ;ConsoleWrite(@CRLF&"Server Received from client:"& ($sReceived))
	 ;$writefile=fileopen(@ScriptDir&"\"&$year&$month&$day&"_encode_info.txt",1)
	 ;FileWriteLine($writefile,@hour&":"&@MIN&":"&@SEC&", "&$sReceived )
	 ;FileWriteLine($writefile, _HexToString($sReceived ) &@CRLF )
	 ;FileWriteLine($writefile, ($sReceived ) &@CRLF )
	 ;FileClose($writefile)
	 ;Disconnect(1)
	 ;sleep(3000)
	 ;tooltip("")

	 $sReceived=_HexToString($sReceived)
	 if StringInStr($sReceived,"|*|") then
		 $SMS_send_date_EPOCH=stringleft($sReceived,10)
		 $SMS_send_date =  _DateAdd("s", Int($SMS_send_date_EPOCH), "1970/01/01 00:00:00")
		 $sReceived_truncated =StringTrimLeft($sReceived,13)
		 $username=StringLeft ( $sReceived_truncated, StringInStr( $sReceived_truncated,"|*|")-1)
		 $sReceived_truncated = StringTrimLeft ( $sReceived_truncated, StringInStr( $sReceived_truncated,"|*|")+2 )
		 ;$writefile=fileopen(@ScriptDir&"\"&$year&$month&$day&"_encode_info.txt",1
		 $writefile=fileopen(@ScriptDir&"\"& $username &"\"& stringleft($sReceived,10)&".sms",10)

		;FileWriteLine($writefile,@hour&":"&@MIN&":"&@SEC&", "&$sReceived )
		FileWriteLine($writefile, $SMS_send_date_EPOCH & @CRLF &$SMS_send_date & @CRLF & StringReplace( $sReceived_truncated,"|*|",@CRLF )  )
		;FileWriteLine($writefile, ($sReceived ) &@CRLF )
		FileClose($writefile)
		_TCP_Send($hSocket,"pointika")
		ConsoleWrite(@CRLF&"Server Received from client:"& $sReceived)
		_SMS_feed($SMS_send_date_EPOCH ,$username)
		;$writefile=fileopen(@ScriptDir&"\sms_feed",9)
		;FileWriteLine($writefile, $SMS_send_date_EPOCH &","&$username)
		;FileClose($writefile)
	 Else
		_TCP_Server_DisconnectClient($hSocket)
		ConsoleWrite(@CRLF&"Server disconnect the client:"& $hSocket )
	 EndIf
EndFunc


Func _SMS_feed($EPOCH , $name)
; Another sample which automatically creates the directory structure
While 1
$file = fileopen(@ScriptDir&"\sms_feed",9) ; which is similar to 1 + 8 (Append + create dir)
	If $file = -1 Then
		;MsgBox(0, "Error", "Unable to open file.")
		sleep(100)
		;Exit
	Else
		FileWriteLine($file, $EPOCH &","&$name)
		FileClose($file)
		ExitLoop
	EndIf
WEnd
EndFunc

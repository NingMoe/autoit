#include <TCP.au3>
#include <date.au3>
#include <sqlite.au3>
#include <sqlite.dll.au3>
#include <array.au3>
dim $sec=@SEC
dim $min=@MIN
dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR

_SQLite_Startup ()
If @error > 0 Then
    MsgBox(16, "SQLite Error", "SQLite.dll Can't be Loaded!")
    Exit - 1
EndIf
_SQLite_Open (@ScriptDir&"\encryption_sqlite") ; Open a  database file
If @error > 0 Then
    MsgBox(16, "SQLite Error", "Can't Load Database!")
    Exit - 1
EndIf
;;=======================================================================================
;;Creat DB file 
	;_SQLite_Exec(-1,"Create table Encryption (key ,name ,md5 );" )

		;_SQLite_Exec(-1,"Create table Encryption (INTIME ,MD ,EVENT );" )
		;_SQLite_Exec(-1,"insert into Encryption values ('A001', '53611'  ,'660');" )

	if _SQLite_Exec(-1,"select NAME from Encryption where MD5='33333';") <> $SQLITE_OK  then 
		
		MsgBox(0,"SQLite Error","Error Code: " & _SQLite_ErrCode() & @CR & "Error Message: " & _SQLite_ErrMsg(),5)
		_SQLite_Exec(-1,"Create table Encryption ( KEY ,NAME ,MD5 ,INTIME);" )
		_SQLite_Exec(-1,"insert into Encryption values ('11111', '22222'  ,'33333','20090910');" )
		;MsgBox(0,'No Data', 'Insert New data to the db Encryption',5)
	Else
		
	;MsgBox(0,'Update Data', 'Update data from Login',5)
	;_SQLite_Exec(-1,"insert into Login values ('A001', '53611'  ,'662');" )
	;_SQLite_Exec(-1,"Update Login set  EVENT='662' where INTIME='A001'  and  MD='53611';" )
	;_SQLite_Exec(-1,"Update Login set  EVENT='662' where INTIME='A002'  and  MD='53612';" )
	;_SQLite_Exec(-1,"Update Login set  EVENT='662' where INTIME='A003'  and  MD='53613';" )
	;_SQLite_Exec(-1,"Update Login set  EVENT='662' where INTIME='A004'  and  MD='53614';" )
		
	endif 
_SQLite_Close ()
_SQLite_Shutdown ()

;;=======================================================================================
;ToolTip("SERVER: Creating server...",10,30)
;traytip("TCP Server", "Port open at 88",5)
ConsoleWrite("SERVER: Creating server...")
;	 sleep(3000)
;	 tooltip("")
_TCP_Server_Create(88); A server. Tadaa!

_TCP_RegisterEvent($TCP_NEWCLIENT, "NewClient"); Whooooo! Now, this function (NewClient) get's called when a new client connects to the server.
_TCP_RegisterEvent($TCP_DISCONNECT, "Disconnect"); And this,... this will get called when a client disconnects.
_TCP_RegisterEvent($TCP_RECEIVE, "Received"); Function "Received" will get called when something is received
While 1
    
	sleep(500)
WEnd

Func NewClient($iError); Yo, check this out! It's a $iError parameter! (In case you didn't noticed: It's in every function)
     ;traytip("TCP Server", "Client Connected",5)
	 ;ToolTip("SERVER: New client connected."&@CRLF&"Sending this: Bleh!",10,30)
	 ;ConsoleWrite("SERVER: New client connected."&@CRLF&"Sending this: Bleh!")
     ;_TCP_Server_Send("Bleh!"); Sending: "Bleh!" to the new client.
     $ip=_TCP_Server_ClientIP($_TCP_ACTIVECLIENT())
	 ToolTip("SERVER: New client connected."&$_TCP_ACTIVECLIENT()&@CRLF&" IP is: "&$ip,10,30)
	 ;sleep(3000)
	 ;tooltip("")
EndFunc

Func Disconnect($iError); Damn, we lost a client. Time of death: @Hour & @Min & @Sec :P
     ToolTip("SERVER: Client disconnected.",10,30); Placing a tooltip right under the tooltips of the client.
	 sleep(1000)
	 tooltip("")
	ConsoleWrite(@CRLF&"SERVER: Client disconnected.")
EndFunc

Func Received($iError, $sReceived); And we also registered this! Our homemade do-it-yourself function gets called when something is received.
    dim $sec=@SEC
	dim $min=@MIN
	dim $hour=@HOUR
	Dim $day=@MDAY
	Dim $month=@MON
	DIM $year=@YEAR
	 ToolTip("Server: We received this: "& $sReceived, 10,30); (and we'll display it)
	 ConsoleWrite(@CRLF&"Server Received from client:"& $sReceived)
	 $writefile=fileopen(@ScriptDir&"\"&$year&$month&$day&"_encode_info.txt",1)
	 FileWriteLine($writefile,@hour&":"&@MIN&":"&@SEC&", "&$sReceived )
	 FileClose($writefile)
	 ;if StringInStr($sReceived,"Bryant") then _TCP_Server_Send("Get your point.")
	 ;Disconnect(1)
	 ;sleep(3000)
	 ;tooltip("")
		dim $dbrecord=StringSplit($sReceived,",")
		;MsgBox(0,"DBrecord",$dbrecord[1]&"<>"&$dbrecord[2]&"<>"&$dbrecord[3]&"<>"&$dbrecord[4])
		_SQLite_Startup ()	
		_SQLite_Open (@ScriptDir&"\encryption_sqlite") ; Open a  database file
		
			if $dbrecord[1]=1 then 
		;_SQLite_Exec(-1,"Create table Encryption ( KEY ,NAME ,MD5 ,INTIME);" )
		;_SQLite_Exec(-1,"insert into Encryption values ('"&$dbrecord[1]&"', '"&$dbrecord[1]&"', '"&$dbrecord[3]&"', '"&@YEAR&@MON&@MDAY&" "&@hour&@MIN&@SEC&"');" )
			_SQLite_Exec(-1,"insert into Encryption values ('"&$dbrecord[3]&"', '"&$dbrecord[2]&"', '"&$dbrecord[4]&"', '"&@YEAR&@MON&@MDAY&@hour&@MIN&@SEC&"');" )
		;_SQLite_Exec(-1,"insert into Encryption values ('11111', '55555555555555'  ,'33333','20090910 140000');" )	
		;if  _SQLite_Exec(-1,"select * from Encryption where KEY='"&$dbrecord[2]&"';")<> $SQLITE_OK  then 
		;
		;	MsgBox(0,"SQLite Error","Error Code: " & _SQLite_ErrCode() & @CR & "Error Message: " & _SQLite_ErrMsg(),5)
		;
		;	_SQLite_Exec(-1,"insert into Encryption values ('"&$dbrecord[2]&"', '"&$dbrecord[1]&"', '"&$dbrecord[3]&"', '"&@YEAR&@MON&@MDAY&@hour&@MIN&@SEC&"');" )
			;MsgBox(0,'No Data', 'Insert New data to the db Login',5)
			
		;Else
			;_SQLite_Exec(-1,"Update Encryption set  KEY='"&$dbrecord[2]&"' and MD5='"&$dbrecord[3]&"' and INTIME='"&@YEAR&@MON&@MDAY&@hour&@MIN&@SEC&"'  where NAME='"&$dbrecord[1]&"';" )	
			;MsgBox(0,'Update Data', 'Update data from Login',5)
		
		;endif
			EndIf
			if $dbrecord[1]=0 then 
				;MsgBox(0,"send to client","Hi,How are you.")
				;_TCP_Server_Send("Hi,How are you.")
				$n=_SQLite_Exec(-1,"select KEY , MD5 from Encryption where NAME='"&$dbrecord[3]&"';","_cb" ) 
				
				;sleep(500)
				;$n=_SQLite_Exec(-1,"select MD5 from Encryption where NAME='"&$dbrecord[3]&"';","_cb" ) 
				sleep(500)
				Disconnect(1)
			EndIf

		_SQLite_Close ()
		_SQLite_Shutdown ()
	
EndFunc


Func _cb($aRow)
	local $sentence=""
	
	;$sentence=_ArrayToString($aRow,"*")
	;$sentence=( StringReplace(StringReplace(StringStripWS($sentence,8),"KEY","") ,"MD5","") )
	ConsoleWrite(@LF)
	if $aRow[0] <> 'Key' then 
		For $s In $aRow
			;	ConsoleWrite($s & @TAB)
			;_TCP_Server_Send($s)
			;_TCP_Server_Send(StringReplace(StringReplace($s,"KEY","",0,1),"MD5","",0,1)&",")
			$sentence=$sentence&","&$s
			;ConsoleWrite($s)
	
		Next
	EndIf	
	;ConsoleWrite($sentence)
	;ConsoleWrite(@CRLF&  $sentence & @CRLF)
	_TCP_Server_Send(  $sentence )
     ConsoleWrite(@LF)
    ; Return $SQLITE_ABORT ; Would Abort the process and trigger an @error in _SQLite_Exec()
EndFunc

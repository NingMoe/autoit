#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here


#include <CommMG.au3> ;or if you save the commMg.dll in the @scripdir use #include @SciptDir & '\commmg.dll'
#include <windowsconstants.au3>
#include <file.au3>
#include <array.au3>
#include <read_parameters.au3>

Dim $sec = @SEC
Dim $min = @MIN
Dim $hour = @HOUR
Dim $day = @MDAY
Dim $month = @MON
Dim $year = @YEAR
Dim $today = $year & $month & $day

;dim $result = '';used for any returned error message setting port

;dim $setflow = 2;default to no flow control
;Dim $FlowType[3] = ["XOnXoff", "Hardware (RTS, CTS)", "NONE"]

Global $Port=12
Global $Baud=9600
Global $Data_Bits=8
Global $Stop_bits=1
Global $parity='none'
Global $flow_control="NONE"



dim $Dhcp=1
dim $Router="rt-n12c"
dim $UDPServer="202.168.197.244"
dim $UDPPort=514
dim $Command="at+ws"
dim $SecondCount=30

dim $connected=0
dim $first_connect=1
dim $ini_array
dim  $instr
$wireless_scan=''

$ini_array= _read_parameters (@ScriptName)
;_ArrayDisplay ($ini_array)

if IsArray( $ini_array ) then
	for $x=1 to UBound($ini_array)-1
	;ComPort=12
	if StringInStr ( $ini_array[$x],"ComPort=" ) then $ComPort=  StringTrimLeft( $ini_array[$x] ,8)
	; Baud=9600
	if StringInStr ( $ini_array[$x],"Baud=" ) then $Baud=  StringTrimLeft( $ini_array[$x] ,5)
	; Data_Bits=8
	if StringInStr ( $ini_array[$x],"Data_Bits=" ) then $Data_Bits=  StringTrimLeft( $ini_array[$x] ,10)
	; Stop_bits=1
	if StringInStr ( $ini_array[$x],"Stop_bits=" ) then $Stop_bits=  StringTrimLeft( $ini_array[$x] ,10)
	; Parity=none
	if StringInStr ( $ini_array[$x],"Parity=" ) then $parity=  StringTrimLeft( $ini_array[$x] ,7)
	; Flow_control=NONE
	if StringInStr ( $ini_array[$x],"Flow_control=" ) then $flow_control=  StringTrimLeft( $ini_array[$x] ,13)
	; Dhcp=1
	if StringInStr ( $ini_array[$x],"Dhcp=" ) then $Dhcp=  StringTrimLeft( $ini_array[$x] ,5)
	; Router=rt-n12c
	if StringInStr ( $ini_array[$x],"Router=" ) then $Router=  StringTrimLeft( $ini_array[$x] ,7)
	; UDPServer=202.168.197.244
	if StringInStr ( $ini_array[$x],"UDPServer=" ) then $UDPServer=  StringTrimLeft( $ini_array[$x] ,10)
	; UDPPort=514
	if StringInStr ( $ini_array[$x],"UDPPort=" ) then $UDPPort=  StringTrimLeft( $ini_array[$x] ,8)
	; Command=at+ws
	if StringInStr ( $ini_array[$x],"Command=" ) then $Command=  StringTrimLeft( $ini_array[$x] ,8)
	; SocketSend=1
	if StringInStr ( $ini_array[$x],"SocketSend=" ) then $SocketSend=  StringTrimLeft( $ini_array[$x] ,11)
	; SecondCount=100
	if StringInStr ( $ini_array[$x],"SecondCount=" ) then $SecondCount=  StringTrimLeft( $ini_array[$x] ,12)

	Next
EndIf

;MsgBox (0,"Parameter", $ComPort &" "& $Baud &" "& $Data_Bits &" "&$Stop_bits &" "& $parity &" "& $Dhcp &" "& $Router &" "&$UDPServer &" "& $UDPPort &" "&$Command &" "& $SocketSend &" "& $SecondCount )









UDPStartup()

; Register the cleanup function.
OnAutoItExitRegister("Cleanup")

Local $socket = UDPOpen($UDPServer, $UDPPort)
If @error <> 0 Then  MsgBox(0,"Error on TCP/UDP Connection", "Error on TCP/UDP Connection: " & $UDPServer &" : "& $UDPPort)



ConsoleWrite ( _CommGetVersion(1)  )

	if  Setport(1) and $first_connect=1 then
		;_CommSendstring("at+nclose=0" &@CRLF )
		;sleep(100)
		_CommSendstring("at+wa=" & $Router &@CRLF)
		sleep (100)
		_CommSendstring("at+ndhcp=" & $Dhcp &@CRLF)
		sleep(50)
		;_CommSendstring("at+ncudp=" &$UDPServer &"," & $UDPPort &@CRLF )
		;sleep(100)
		;_CommSendstring("at+cid=?" &@CRLF )

		;sleep(100)

	$first_connect=0
	EndIf

$x=0
While 1

	sleep(1000)

	_CommSendstring( $Command & @CR)

    ;gets characters received returning when one of these conditions is met:
    ;receive @CR, received 20 characters or 200ms has elapsed
   $instr = _CommGetString()

    If $instr <> '' Then;if we got something

        $instr = StringReplace($instr,@CR,@CRLF)

		$wireless_scan=$wireless_scan & $instr
		if StringInStr($wireless_scan ,"OK" )  then
			ConsoleWrite ( $wireless_scan  )
			FileWriteLine ( @ScriptDir &"\" & "Commg_send_" &$today &".log" , $wireless_scan & @CRLF)
			Local $status
			if $SocketSend=1 then
				$status = UDPSend($socket,  $wireless_scan)
				If $status = 0 Then MsgBox(0, "ERROR", "Error while sending UDP message: " & @error)
			EndIf

			$wireless_scan=''
			;_CommSendByte ("{ESC}",1)
			;_CommSendstring( chr(27) & @CRLF )
			;_CommSendstring( "U0202.168.197.244:514:" & $wireless_scan & @CRLF )
			;_CommSendByte ("{ESC}",1)
			;_CommSendstring( "chr(27)" & @CRLF )
			;_CommSendstring( "E" & @CRLF )
			;_CommSendByte ( chr(69),1)
			;U0202.168.197.244:514:I am evb
		EndIf

    EndIf
	$x=$x+1
	ConsoleWrite (@CRLF &"Count : "& $x )
	sleep(1000)
	if $x > $SecondCount  and $wireless_scan='' then ExitLoop
WEnd


Exit

;Func _sendudp()
;
;EndFunc

Func SetPort($mode = 1);if $mode = 1 then returns -1 if settings not made
	local $msg , $sportSetError

    local $portlist = _CommListPorts(0);find the available COM ports and write them into the ports combo
	;_ArrayDisplay($portlist)

    If @error = 1 Then
        MsgBox(0, 'trouble getting portlist', 'Program will terminate!')
        Exit
	EndIf

			 ;_CommSetPort($iPort, ByRef $sErr, $iBaud = 9600, $iBits = 8, $iPar = 0, $iStop = 1, $iFlow = 0, $RTSMode = 0, $DTRMode = 0)

            _CommSetPort($port, $sportSetError, $Baud, $Data_Bits, $parity, $Stop_bits, $flow_control)
            if $sportSetError = '' Then
				MsgBox(262144, 'Connected ','to COM' & $port, 2)
				return True
			Else
				MsgBox(262144, 'Setport error = ', $sportSetError)
				Return False
			EndIf





EndFunc   ;==>SetPort

Func Cleanup()
    UDPCloseSocket($socket)
    UDPShutdown()
EndFunc   ;==>Cleanup
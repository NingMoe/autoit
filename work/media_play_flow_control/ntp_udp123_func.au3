#include <Date.au3>

;MsgBox(0,'Now', _ntpnow(),5)

func _ntpnow()
;$ntpServer="pool.ntp.org"
;dim $ntpServer="stdtime.gov.tw"
dim $ntpServer="clock.cuhk.edu.hk"
dim $ntpServer="time-a.nist.gov"
dim $ntpServer="clock.stdtime.gov.tw"

UDPStartup()
Dim $socket = UDPOpen(TCPNameToIP($ntpServer), 123)
If @error <> 0 Then
    MsgBox(0,"","Can't open connection!",5)
    Exit
EndIf
;$status = UDPSend($socket, MakePacket("1b0e010000000000000000004c4f434ccb1eea7b866665cb00000000000000000000000000000000cb1eea7b866665cb "))
local $status
$status = UDPSend($socket, MakePacket("1b0e01000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"))
If $status = 0 Then
    MsgBox(0, "ERROR", "Error while connect to ntp server: " & @error,5)
	MsgBox(0, "ERROR", "Error while sending UDP message: " & @error,5)
	return(-1)
    Exit
EndIf
local $data=""
While $data=""
    $data = UDPRecv($socket, 100)
    sleep(100)
WEnd
UDPCloseSocket($socket)
UDPShutdown()
local $unsignedHexValue, $value , $TZinfo ,$TZoffset ,$UTC
$unsignedHexValue=StringMid($data,83,8); Extract time from packet. Disregards the fractional second.
;MsgBox(0, "UDP DATA", $unsignedHexValue)

$value=UnsignedHexToDec($unsignedHexValue)
$TZinfo = _Date_Time_GetTimeZoneInformation()
$TZoffset=$TZinfo[1]*-1
$UTC=_DateAdd("s",$value,"1900/01/01 00:00:00")
;MsgBox(0,"","Time from NTP Server UTC:  "&$UTC&@CRLF& _
;                "Time from NTP Server Local Offset:  "&_DateAdd("n",$TZoffset,$UTC)&@CRLF& _
;                "Time from Local Computer Clock:  "&_NowCalc())

return(_DateAdd("n",$TZoffset,$UTC))
EndFunc

;**************************************************************************************************
;** Fuctions **************************************************************************************
;**************************************************************************************************
Func MakePacket($d)
    Local $p=""
    While $d
        $p&=Chr(Dec(StringLeft($d,2)))
        $d=StringTrimLeft($d,2)
    WEnd
    Return $p
EndFunc
;**************************************************************************************************
Func UnsignedHexToDec($n)
	local $ones
    $ones=StringRight($n,1)
    $n=StringTrimRight($n,1)
    Return dec($n)*16+dec($ones)
EndFunc
;**************************************************************************************************
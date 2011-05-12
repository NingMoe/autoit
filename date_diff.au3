#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
#include <Date.au3>



dim $sec=@SEC
dim $min=@MIN
Dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR
; Script Start - Add your code below here

	$now=StringLeft(_NowCalc(),10)
	$date_addded = _DateAdd("d",2, $now )

	MsgBox(0,"DateAdd", $now & "-- " &$date_addded )

	if FileExists( @ScriptDir&"\hotel_info.txt" ) Then	
		;$hotel_info=_file2Array(@ScriptDir&"\hotel_info.txt",7,"^")
		;MsgBox(0,"Array Y ", UBound($hotel_info))
		;_ArrayDisplay($hotel_info)
		$diff=  StringReplace( StringLeft(_NowCalc(),10) ,"/","" ) -  StringLeft( FileGetTime( @ScriptDir&"\hotel_info.txt",0,1) ,8)
		; _DateDiff ( "D" , FileGetTime( @ScriptDir&"\hotel_info.txt",1,1) , _NowCalc())

	EndIf
	
	MsgBox(0,"DIFF" , $diff  & @CRLF & StringReplace( StringLeft(_NowCalc(),10) ,"/","" )  & @CRLF  & StringLeft( FileGetTime( @ScriptDir&"\hotel_info.txt",0,1) ,8) ) 
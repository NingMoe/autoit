#include <array.au3>
#include <File.au3>
#include <Date.au3>
;#include <CompInfo_win7.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GUIConstants.au3>
#include <mail_variable.au3>

_Gen_Test_Data()

Func _Gen_Test_Data()
	local $Data , $infome_data , $now_DateCalc , $file
	$now_DateCalc = ( _DateDiff( 's',"1970/01/01 00:00:00",_NowCalc()) ) + 100
	$data=$now_DateCalc &@CRLF &  _
			"2011/09/23 14:37:27" &@CRLF &  _
			"changtun@gmail.com"  &@CRLF & _
			"0928837823"  &@CRLF & _
			"當你有多個外面線路想要同時間對映到內部的某一台Server時，或許你是為了備援或Load Balance的考量。在Router" &@CRLF & _
			"姓名,行動電話"  &@CRLF &  _
			"changtun1,_0928837823"  &@CRLF & _
			"sean2,_0956330560"  &@CRLF & _
			"seanlu3,_0968269170"  &@CRLF & _ 
			"bryant4,_0928837823" 
	
	$infome_data=$now_DateCalc &",changtun"
	
	$file=fileopen (@ScriptDir & "\changtun\"& $now_DateCalc& ".sms",10)
	FileWrite($file, $data )
	FileClose($file)
	sleep(100)
	$file=fileopen (@ScriptDir & "\sms_feed",10)
	FileWrite($file,$infome_data)
	FileClose($file)
	sleep(100)
EndFunc


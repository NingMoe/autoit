#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.2.12.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <array.au3>
#include <mysql.au3>
#Include <File.au3>
;
;;
dim $db_ip="10.112.55.87"
dim  $my_http_base="st1.onlinebooking.com.tw"
dim $single_email="bryant@dynalab.com.tw"


dim $mymailbody 
$mymailbody=_read_direct_fly(@ScriptDir&"\direct_fly.htm", $my_http_base)



	if StringInStr($mymailbody,"[user_email]") then 
	$mymailbody=StringReplace($mymailbody,"[user_email]",$single_email)
	;MsgBox(0,"Message",$aRecords[$x])
	EndIf

MsgBox(0,"mail body",$mymailbody)


Func _read_direct_fly($directfly,$http_base)
	
local $aRecords
local $directfly_string=''
If Not _FileReadToArray($directfly,$aRecords) Then
   MsgBox(4096,"Error", " Error reading log to Array     error:" & @error)
   Exit
EndIf
For $x = 1 to $aRecords[0]
	$directfly_string=$directfly_string & $aRecords[$x]& @CRLF	
Next
	if StringInStr($directfly_string,"[http_base]") then 
	$directfly_string=StringReplace($directfly_string,"[http_base]",$http_base)
	;MsgBox(0,"Message",$aRecords[$x])
	EndIf
	;
	;if StringInStr($directfly_string,"[user_email]") then 
	;$directfly_string=StringReplace($directfly_string,"[user_email]",$single_email)
	;;MsgBox(0,"Message",$aRecords[$x])
	;EndIf

	;;Msgbox(0,'mail body', $directfly_string)
Return $directfly_string
EndFunc
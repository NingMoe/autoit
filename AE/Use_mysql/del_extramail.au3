#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

;ExtraMail_del_list.txt
#include <file.au3>
#include <array.au3>
#include <mysql.au3>

dim $sec=@SEC
dim $min=@MIN
Dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR

dim $db_ip="10.112.55.102"
dim  $my_http_base="st1.onlinebooking.com.tw"

Dim $ExtraMail_del_lists
If Not _FileReadToArray(@ScriptDir&"\ExtraMail_del_list2.txt",$ExtraMail_del_lists) Then
   ;MsgBox(4096,"Error", " Error reading log to Array     error:" & @error)
   Exit
EndIf
;For $x = 1 to $aRecords[0]
;    Msgbox(0,'Record:' & $x, $aRecords[$x])
;Next


;_ArrayDisplay($ExtraMail_del_lists)


;;;======================  
;;DB function
;;
;;
; MYSQL starten, DLL im PATH (enth‰lt auch @ScriptDir), sont Pfad zur DLL angeben. DLL muss libmysql.dll heiﬂen.
_MySQL_InitLibrary()
If @error Then Exit MsgBox(0, '', "")
;MsgBox(0, "DLL Version:",_MySQL_Get_Client_Version()&@CRLF& _MySQL_Get_Client_Info())

$MysqlConn = _MySQL_Init()

;Fehler Demo: C:\InstantRails-2.0-win\mysql\data\cookbook
;MsgBox(0,"Fehler-Demo","Fehler-Demo")
$connected = _MySQL_Real_Connect($MysqlConn,$db_ip,"alongbird","pkpkpk","test")
If $connected = 0 Then
	$errno = _MySQL_errno($MysqlConn)
	MsgBox(0,"Error:",$errno & @LF & _MySQL_error($MysqlConn))
	If $errno = $CR_UNKNOWN_HOST Then MsgBox(0,"Error:","$CR_UNKNOWN_HOST" & @LF & $CR_UNKNOWN_HOST)
Endif

for $i=1 to $ExtraMail_del_lists[0]
local $query1=""
;$query = "SELECT * FROM extra_email where extra_nbr='00000030640'"
;$query = "SELECT * FROM extra_email"
$query = "SELECT * FROM extra_email WHERE email='"& $ExtraMail_del_lists[$i] &"'" 
;$query = "SELECT * FROM extra_email WHERE email='aaa@abc.com'" 

$response= _MySQL_Real_Query($MysqlConn, $query)
;MsgBox(0, "_MySQL_Real_Query response", $response)

$res = _MySQL_Store_Result($MysqlConn)

;MsgBox(0, "_MySQL_Store_Result", $res)
;$fields = _MySQL_Num_Fields($res)

;$rows = _MySQL_Num_Rows($res)
;MsgBox(0, "", $rows & "-" & $fields)


;MsgBox(0, '', "Zugriff Methode 3 - alles in ein 2D Array")
$extramail_array = _MySQL_Fetch_Result_StringArray($res)
if UBound($extramail_array)=0 then 
	;MsgBox(0,"No data at the DB","NO data",2)
	_FileWriteLog(@ScriptDir&"\"&$month&$day&$hour&$min&"_Extra_mail_delete.log", $ExtraMail_del_lists[$i] &" not in DataBase" )
Else	
	;_ArrayDisplay($extramail_array)
	$query1="DELETE From extra_email WHERE email='"& $ExtraMail_del_lists[$i] &"'" 
	 _MySQL_Real_Query($MysqlConn, $query1)
EndIf

Next

;_ArrayDisplay($extramail_array)
;===== If you select all from DB then you will need to use these code for filter.
; ; Y is from 1 X is from 0
; ;MsgBox(0,"Array((y,x)",$extramail_array[5][1])
; ;MsgBox(0,"Array((5,1)",$extramail_array[5][1])
;
; ;MsgBox(0,"Array((5,1)",UBound($extramail_array,1))
;;=====
;dim $email[UBound($extramail_array,1)]
;
;for $r=1 to (UBound($extramail_array,1)-1)
;		$email[$r]=$extramail_array[$r][1]
;		;_FileWriteFromArray(@ScriptDir&"\email_output.txt",$email,1 )
;		;_FileWriteToLine(@ScriptDir&"\email_output.txt",$r , $email[$r],1)
;Next
;$email[0]=UBound($extramail_array,1)
;_ArrayDisplay($email)

; Abfrage freigeben
_MySQL_Free_Result($res)
; Verbindung beenden
_MySQL_Close($MysqlConn)
; MYSQL beenden
_MySQL_EndLibrary()

;MsgBox(0,"Email", "There are :" & (UBound($extramail_array,1)-1) &" email addresses")



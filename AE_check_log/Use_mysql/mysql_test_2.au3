; Test 20090408 by bryant
#cs ----------------------------------------------------------------------------
	
	AutoIt Version: 3.2.8.1 (beta)
	Author:         Prog@ndy
	
	Script Function:
	MySQL-Plugin Demo Script
	
#ce ----------------------------------------------------------------------------
;;
;;
;; 2011 0127 modify for internal network IP for mail server
;;
;;============================
;;20101116 modify for mail subject that is abstracted from mail body of the html file.
;;============================
;;20100721 modify for utf-8 encoding
;; Use onlinebooking.com.tw as mail server.
;; send mail by ae_direct_fly@onlinebooking.com.tw and return should be send to direct_send@onlinebooking.com.tw if need to analysis return mail.
;;============================
; Mail Account for AE
; MyOnlineBookingST1@gmail.com
; myonlinebookingst2@gmail.com
; myonlinebookingst3@gmail.com
; pass amextravel
;
#include <array.au3>
#include <mysql.au3>
#Include <File.au3>

dim $sec=@SEC
dim $min=@MIN
Dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR

;
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
Global $oMyRet[2]
;;
;; Connect My SQL for mail address.
;; DB is now at 10.112.55.87
;;
dim $db_ip="changtuntv.dyndns.org"
dim $db_ip="www.kitravel.com.tw"

dim $test_mode= _TEST_MODE()
;;
;; This is  connect to My SQL for user email
;;
; MYSQL starten, DLL im PATH (enth‰lt auch @ScriptDir), sont Pfad zur DLL angeben. DLL muss libmysql.dll heiﬂen.
_MySQL_InitLibrary()
If @error Then Exit MsgBox(0, '', "")
;MsgBox(0, "DLL Version:",_MySQL_Get_Client_Version()&@CRLF& _MySQL_Get_Client_Info())

$MysqlConn = _MySQL_Init()

;Fehler Demo: C:\InstantRails-2.0-win\mysql\data\cookbook
;MsgBox(0,"Fehler-Demo","Fehler-Demo")
$connected = _MySQL_Real_Connect($MysqlConn,$db_ip,"root","5loveyou","kitravel")
If $connected = 0 Then
	
	$errno = _MySQL_errno($MysqlConn)
	MsgBox(0,"Error:",$errno & @LF & _MySQL_error($MysqlConn))
	If $errno = $CR_UNKNOWN_HOST Then MsgBox(0,"Error:","$CR_UNKNOWN_HOST" & @LF & $CR_UNKNOWN_HOST)
Endif

_MySQL_Set_Character_Set($MysqlConn,"big5")

;Exit
; XAMPP cdcol
;MsgBox(0, "XAMPP-Cdcol-demo", "XAMPP-Cdcol-demo")

;$connected = _MySQL_Real_Connect($MysqlConn, "localhost", "root", "", "cookbook")
;If $connected = 0 Then Exit MsgBox(16, 'Connection Error', _MySQL_Error($MysqlConn))

;$query = "SELECT * FROM extra_email where extra_nbr='00000030640'"
;$query = "SELECT * FROM extra_email"
$query = "select HOTELID,HOTELNAME,TEL,FAX,EMAIL,WINDOW,STATUS from hotel"

_MySQL_Real_Query($MysqlConn, $query)


$res = _MySQL_Store_Result($MysqlConn)

;$fields = _MySQL_Num_Fields($res)

;$rows = _MySQL_Num_Rows($res)
;MsgBox(0, "", $rows & "-" & $fields)


;MsgBox(0, '', "Zugriff Methode 3 - alles in ein 2D Array")
$extramail_array = _MySQL_Fetch_Result_StringArray($res)
_ArrayDisplay($extramail_array)

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





Func _TEST_MODE()
	IF FileExists(@ScriptDir&"\TESTMODE.txt") Then
		$mode=FileReadLine(@ScriptDir&"\TESTMODE.txt",1)
		if $mode=1 then 
			MsgBox(0,"Test mode", "This is Test mode. All mail send to direct_send@onlinebooking.com.tw",5)
			
		Else
			MsgBox(0,"Delivery mode", "This is True mail delivery.",5)
			$mode=0
		EndIf
	
	Else
		MsgBox(0,"Delivery mode", "This is True mail delivery.",5)
		$mode=0
	EndIf
	
	return $mode
EndFunc
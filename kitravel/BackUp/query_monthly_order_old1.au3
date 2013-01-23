#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <array.au3>
#include <mysql.au3>
#Include <File.au3>
#Include <Date.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GUIConstants.au3>

dim $sec=@SEC
dim $min=@MIN
Dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR
dim $today=$year & $month & $day
dim $db_return_data , $return_data , $accbin_date
;
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
Global $oMyRet[2]
;;
;; dim $test_mode= _TEST_MODE()
local $sample
$sample = 'select a.HOTELID, a.ORDERNO, a.ORD, a.AMOUNT, a.PRICE, a.ITEMDESC, a.ADULT, a.CHILD, a.ADULTADD, a.ADULTADD_PRICE, a.CHILDADD, a.CHILDADD_PRICE,a.BF, a.BF_PRICE, a.FREEBF, a.MEMO, b.ORDERTIME,  b.USERNAME, b.GENDER, b.COUNTRY, b.IDNO, b.TEL, b.EMERGENCYTEL, b.FAX, b.PEOPLE, b.OPERATORNAME, b.EDITOR, b.MEMO, b.BALANCE, b.STATUS, b.DOWNPAYMENT, b.DPM_LIMIT, b.PAIEDDOWNPAY, b.DOWNPAYDESC, b.DOWNPAYTIME, b.DPM_METHOD, b.PAYOFF, b.PO_TIME, b.PO_METHOD, b.REIMBRUSEMENT, b.REFOUNDDESC, b.RBM_TIME, b.RBM_METHOD, b.EDITTIME, b.PURPOSE, b.ACCOUNTNO, b.ARRIVALTIME, c.EMAIL, c.MEMO, c.INTEREST, c.NICKNAME from hotel_orderitems a left join (hotel_order b, customer c) on b.HOTELID=a.HOTELID and b.ORDERNO=a.ORDERNO and c.HOTELID=b.HOTELID and c.UID=b.UID where b.ORDERTIME>="201201010000" and b.ORDERTIME<"201202010000"  and a.HOTELID=11 and a.ACTIVE=1 and a.AMOUNT=1 order by a.HOTELID, a.ORDERNO, a.ORD '


;$sample = 'select a.HOTELID, a.ORDERNO, a.ORD, a.AMOUNT, a.PRICE, a.ITEMDESC, a.ADULT, a.CHILD, a.ADULTADD, a.ADULTADD_PRICE, _
;         a.CHILDADD, a.CHILDADD_PRICE, a.BF, a.BF_PRICE, a.FREEBF, _
;         a.MEMO, b.ORDERTIME,  b.USERNAME, b.GENDER, b.COUNTRY, b.IDNO, _
;         b.TEL, b.EMERGENCYTEL, b.FAX, b.PEOPLE, b.OPERATORNAME, _
;         b.EDITOR, b.MEMO, b.BALANCE, b.STATUS, b.DOWNPAYMENT,  _
;         b.DPM_LIMIT, b.PAIEDDOWNPAY, b.DOWNPAYDESC, b.DOWNPAYTIME, b.DPM_METHOD, _
;         b.PAYOFF, b.PO_TIME, b.PO_METHOD, b.REIMBRUSEMENT, b.REFOUNDDESC, _
;         b.RBM_TIME, b.RBM_METHOD, b.EDITTIME, b.PURPOSE, b.ACCOUNTNO, _
;         b.ARRIVALTIME, c.EMAIL, c.MEMO, c.INTEREST, c.NICKNAME _
;         from hotel_orderitems a left join (hotel_order b, customer c) _
;         on b.HOTELID=a.HOTELID and b.ORDERNO=a.ORDERNO and c.HOTELID=b.HOTELID and c.UID=b.UID _
;         where b.ORDERTIME>="201201010000" and b.ORDERTIME<"201202010000"  and a.HOTELID=11 _
;         and a.ACTIVE=1 and a.AMOUNT=1 order by a.HOTELID, a.ORDERNO, a.ORD '



MsgBox (0,"Sample Query" , $sample)


















func _query_db()
;; Connect My SQL for mail address.
;; DB is now at 10.112.55.87
;;
;dim $db_ip="changtuntv.dyndns.org"
dim $db_ip="www.kitravel.com.tw"
;;
;; This is  connect to My SQL for user email
;;
_MySQL_InitLibrary()
If @error Then Exit MsgBox(0, 'Error', " mysql connect error. Check mysql DLL file.")
;MsgBox(0, "DLL Version:",_MySQL_Get_Client_Version()&@CRLF& _MySQL_Get_Client_Info())

$MysqlConn = _MySQL_Init()

;Fehler Demo: C:\InstantRails-2.0-win\mysql\data\cookbook
;MsgBox(0,"Fehler-Demo","Fehler-Demo")
$connected = _MySQL_Real_Connect($MysqlConn,$db_ip,"root","5*love*you","kitravel",6006)
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
;$query = 'select b.HOTELID, b.HOTELNAME, a.UID, a.USERNAME, a.ACCOUNT, a.PASSWD, a.TEL, a.EMAIL, b.IDNAME from manager a, hotel b, working c where c.HOTELID=b.HOTELID and a.UID=c.UID and a.STATUS="E"and b.HOTELID>11 order by b.HOTELID'

$query='select a.HOTELID, a.ORDERNO, a.ORD, a.AMOUNT, a.PRICE, a.ITEMDESC, a.ADULT, a.CHILD, a.ADULTADD, a.ADULTADD_PRICE, a.CHILDADD, a.CHILDADD_PRICE,a.BF, a.BF_PRICE, a.FREEBF, a.MEMO, b.ORDERTIME,  b.USERNAME, b.GENDER, b.COUNTRY, b.IDNO, b.TEL, b.EMERGENCYTEL, b.FAX, b.PEOPLE, b.OPERATORNAME, b.EDITOR, b.MEMO, b.BALANCE, b.STATUS, b.DOWNPAYMENT, b.DPM_LIMIT, b.PAIEDDOWNPAY, b.DOWNPAYDESC, b.DOWNPAYTIME, b.DPM_METHOD, b.PAYOFF, b.PO_TIME, b.PO_METHOD, b.REIMBRUSEMENT, b.REFOUNDDESC, b.RBM_TIME, b.RBM_METHOD, b.EDITTIME, b.PURPOSE, b.ACCOUNTNO, b.ARRIVALTIME, c.EMAIL, c.MEMO, c.INTEREST, c.NICKNAME from hotel_orderitems a left join (hotel_order b, customer c) on b.HOTELID=a.HOTELID and b.ORDERNO=a.ORDERNO and c.HOTELID=b.HOTELID and c.UID=b.UID where b.ORDERTIME>="201201010000" and b.ORDERTIME<"201202010000"  and a.HOTELID=11 and a.ACTIVE=1 and a.AMOUNT=1 order by a.HOTELID, a.ORDERNO, a.ORD '

_MySQL_Real_Query($MysqlConn, $query)
$res = _MySQL_Store_Result($MysqlConn)

;$fields = _MySQL_Num_Fields($res)

;$rows = _MySQL_Num_Rows($res)
;MsgBox(0, "", $rows & "-" & $fields)


;MsgBox(0, '', "Zugriff Methode 3 - alles in ein 2D Array")
$db_return_data = _MySQL_Fetch_Result_StringArray($res)
;_ArrayDisplay($db_return_data)

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

return $db_return_data
EndFunc
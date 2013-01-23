#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.8.1
	Author:         myName

	Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <array.au3>
#include <mysql.au3>
#include <File.au3>
#include <Date.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GUIConstants.au3>

Dim $sec = @SEC
Dim $min = @MIN
Dim $hour = @HOUR
Dim $day = @MDAY
Dim $month = @MON
Dim $year = @YEAR
Dim $today = $year & $month & $day
Dim $db_return_data, $return_data, $accbin_date
;
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
Global $oMyRet[2]
Local $output_header , $output_file
Local $start_month_m , $end_month_m
local $hid=208
local $hid=202



;;
;; dim $test_mode= _TEST_MODE()

; Subtract 2 weeks from today
local $sNewDate
$start_month_m =  StringTrimRight( StringTrimLeft( _DateAdd('m', -1, _NowCalcDate()) ,5 )  , 3)
;MsgBox(4096, "", "Today minus 1 month: " &  @YEAR & $start_month_m &"0000" )

$end_month_m = StringTrimRight( StringTrimLeft(  _NowCalcDate() ,5 )  , 3)
;MsgBox(4096, "", "Today minus 1 month: " &  @YEAR & $end_month_m & "0000" )

if $start_month_m=12 then
	$start_month_m= ( $year -1 ) & $start_month_m &"0000"
else
	$start_month_m= ( $year -1 ) & $start_month_m &"0000"
EndIf

$end_month_m =@YEAR& $end_month_m &"0000"
MsgBox(0,"start and end ",  $start_month_m &" -- "& $end_month_m,5)
;MsgBox(0,"start and end ",  $start_month_m &" -- "& $end_month_m)

$output_header= _db_header()
;_ArrayDisplay ($output_header)

$return_data=_query_order_db( $start_month_m,  $end_month_m  , $hid)
;$return_data=_query_order_db( 2012010000,  2012020000  , 11)
;_ArrayDisplay ($return_data)


if not IsArray ($return_data) then
	MsgBox(0,"No data", "No Business Data between " & $start_month_m & " and " & $end_month_m)
	Exit
EndIf


for $x=0 to 50
	; replace english header to chinese
	;MsgBox ( 0," replace header", $return_data[0][$x] & "  VS "  & $output_header[$x] )
	$return_data[0][$x]= $output_header[$x]
Next
;_ArrayDisplay ($return_data)


;MsgBox (0,"array", _array2string_tab($return_data,51))

$output_file =FileOpen ( @ScriptDir & "\monthly_order_" & $hid &"_"&$start_month_m&".txt",10)
FileWrite ( $output_file,   StringReplace( _array2string_tab($return_data,51) ,"^",";	")  )
FileClose ( $output_file)


Exit









Func _query_order_db( $start_month, $end_month , $hotel_id)
;Func _query_order_db( )
	;; Connect My SQL for mail address.
	;; DB is now at 10.112.55.87
	;;
	;dim $db_ip="changtuntv.dyndns.org"
	Dim $db_ip = "www.kitravel.com.tw"
	;;
	;; This is  connect to My SQL for user email
	;;
	_MySQL_InitLibrary()
	If @error Then Exit MsgBox(0, 'Error', " mysql connect error. Check mysql DLL file.")
	;MsgBox(0, "DLL Version:",_MySQL_Get_Client_Version()&@CRLF& _MySQL_Get_Client_Info())

	$MysqlConn = _MySQL_Init()

	;Fehler Demo: C:\InstantRails-2.0-win\mysql\data\cookbook
	;MsgBox(0,"Fehler-Demo","Fehler-Demo")
	$connected = _MySQL_Real_Connect($MysqlConn, $db_ip, "root", "5*love*you", "kitravel", 6006)
	If $connected = 0 Then

		$errno = _MySQL_errno($MysqlConn)
		MsgBox(0, "Error:", $errno & @LF & _MySQL_error($MysqlConn))
		If $errno = $CR_UNKNOWN_HOST Then MsgBox(0, "Error:", "$CR_UNKNOWN_HOST" & @LF & $CR_UNKNOWN_HOST)
	EndIf

	_MySQL_Set_Character_Set($MysqlConn, "big5")

	;Exit
	; XAMPP cdcol
	;MsgBox(0, "XAMPP-Cdcol-demo", "XAMPP-Cdcol-demo")

	;$connected = _MySQL_Real_Connect($MysqlConn, "localhost", "root", "", "cookbook")
	;If $connected = 0 Then Exit MsgBox(16, 'Connection Error', _MySQL_Error($MysqlConn))

	;$query = "SELECT * FROM extra_email where extra_nbr='00000030640'"
	;$query = "SELECT * FROM extra_email"
	;$query = 'select b.HOTELID, b.HOTELNAME, a.UID, a.USERNAME, a.ACCOUNT, a.PASSWD, a.TEL, a.EMAIL, b.IDNAME from manager a, hotel b, working c where c.HOTELID=b.HOTELID and a.UID=c.UID and a.STATUS="E"and b.HOTELID>11 order by b.HOTELID'

	$query  = 'select a.HOTELID, a.ORDERNO, a.ORD, a.AMOUNT, a.PRICE, a.ITEMDESC, ' &  _
'a.ADULT, a.CHILD, a.ADULTADD, a.ADULTADD_PRICE, a.CHILDADD, a.CHILDADD_PRICE, ' &  _
'a.BF, a.BF_PRICE, a.FREEBF, a.MEMO, b.ORDERTIME,  b.USERNAME, b.GENDER, b.COUNTRY,  ' &  _
'b.IDNO, b.TEL, b.EMERGENCYTEL, b.FAX, b.PEOPLE, b.OPERATORNAME, b.EDITOR,  ' &  _
'b.MEMO, b.BALANCE, b.STATUS, b.DOWNPAYMENT, b.DPM_LIMIT, b.PAIEDDOWNPAY,  ' &  _
'b.DOWNPAYDESC, b.DOWNPAYTIME, b.DPM_METHOD, b.PAYOFF, b.PO_TIME, b.PO_METHOD,  ' &  _
'b.REIMBRUSEMENT, b.REFOUNDDESC, b.RBM_TIME, b.RBM_METHOD, b.EDITTIME, b.PURPOSE,  ' &  _
'b.ACCOUNTNO, b.ARRIVALTIME, c.EMAIL, c.MEMO, c.INTEREST, c.NICKNAME from  ' &  _
'hotel_orderitems a left join (hotel_order b, customer c) on b.HOTELID=a.HOTELID  ' &  _
'and b.ORDERNO=a.ORDERNO and c.HOTELID=b.HOTELID and c.UID=b.UID where  ' &  _
'b.ORDERTIME>="'&$start_month&'" and b.ORDERTIME<"'&$end_month&'"  and a.HOTELID='&$hotel_id&' and  ' &  _
'a.ACTIVE=1 and a.AMOUNT=1 order by a.HOTELID, a.ORDERNO, a.ORD '

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

	Return $db_return_data
EndFunc   ;==>_query_db





Func _TEST_MODE()
	IF FileExists(@ScriptDir&"\TESTMODE.txt") Then
		$mode=FileReadLine(@ScriptDir&"\TESTMODE.txt",1)
		if $mode=1 then
			MsgBox(0,"Test mode", "This is Test mode. ",5)

		Else
			;MsgBox(0,"Delivery mode", "This is True delivery.",5)
			$mode=0
		EndIf

	Else
		;MsgBox(0,"Delivery mode", "This is True  delivery.",5)
		$mode=0
	EndIf

	return $mode
EndFunc


func _array2string_tab($array,$d)
	local $y, $x, $i, $string_to_return ,$a_line
	$string_to_return=''

	;MsgBox(0,"Dimention", "  $Y :" &UBound($array) )

	for $y=0 to UBound($array)-1
		$a_line=''
		if $d=1 then
			$a_line=$array[$y]

		Else
			for $x=0 to $d-1
				if $x < $d-1 then
					$a_line=$a_line &$array[$y][$x] &'^'
				else
					$a_line=$a_line &$array[$y][$x]
				EndIf
				;$string_to_return=$string_to_return  & $array[$y][1] &" , "& $array[$y][2] &" , " &$array[$y][6] &@CRLF
			next
			;$a_line=StringTrimRight($a_line,3) ; Cut off the last 3 character --> ","
		EndIf
		$string_to_return=$string_to_return & $a_line  & @CRLF
	Next


;MsgBox(0,"string", $string_to_return)
return $string_to_return
EndFunc




func _db_header()
	local $header[51]

		  $header[0]="HID"
		  $header[1]="訂單編號"
		  $header[2]="序"
		  $header[3]="房間數"
		  $header[4]="房價"

          $header[5]="內容"
		  $header[6]="大人"
		  $header[7]="小孩"
		  $header[8]="大人加床"
		  $header[9]="加床費"

          $header[10]="小孩加床"
		  $header[11]="加床費"
		  $header[12]="早餐"
		  $header[13]="早餐費"
		  $header[14]="免費早餐"

          $header[15]="房間入住摘要"
		  $header[16]="訂單時間"

		  $header[17]="客人姓名"
		  $header[18]="性別"
		  $header[19]="國家"
		  $header[20]="身份証/護照"

          $header[21]="電話"
		  $header[22]="緊急電話"
		  $header[23]="傳真"
		  $header[24]="同行人數"
		  $header[25]="接單人"

          $header[26]="最後修改"
		  $header[27]="訂單備註"
		  $header[28]="訂單餘額"
		  $header[29]="訂單狀態"
		  $header[30]="應付訂金"

          $header[31]="付訂期限"
		  $header[32]="已付訂金"
		  $header[33]="付訂備註"
		  $header[34]="付訂時間"
		  $header[35]="付訂方式"

          $header[36]="結帳付款"
		  $header[37]="結帳時間"
		  $header[38]="結帳方式"
		  $header[39]="退款金額"
		  $header[40]="退款備註"

          $header[41]="退款時間"
		  $header[42]="退款方式"
		  $header[43]="訂單最後修改時間"
		  $header[44]="出遊原因"
		  $header[45]="銀行帳號"

          $header[46]="預計到達時間"
		  $header[47]="EMAIL"
		  $header[48]="客人本身備註"
		  $header[49]="興趣"
		  $header[50]="暱稱"

	Return $header
EndFunc
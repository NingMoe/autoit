#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here



#include <mysql.au3>
#include <array.au3>
;InetGet("http://www.eztravel.com.tw/ezec/pkghsr/AeHsrInfo?queryDate=20101222", @ScriptDir&"\20101222.xml")
;exit
;;
;dim $test_mode= _TEST_MODE()
dim $DB_IP_bry= "127.0.0.1"
;dim $DB_IP_bry= "192.168.3.45"
dim $DB_NAME_bry ="test7"
dim $DB_ACCOUNT_bry ="amex"
dim $DB_PASS_bry ="amextravel"
dim $STATEMENT_bry 

dim $query_insert= "INSERT INTO hs_product (product_date,Tran_No,hs_class,DepTime,ArrTime,depCity,desCity) VALUES('20100101','399','N','0000','2229','110','115')"

dim $query_select = "SELECT * FROM hs_product where product_date<'20100201'"
dim $query_last ="SELECT * FROM `test7`.`hs_product` ORDER BY product_nbr DESC limit 1"

dim $query_update ="UPDATE `test7`.`hs_product` set product_date='20100105' where product_nbr='0006970685'"

dim $query_del = "DELETE from `test7`.`hs_product` where product_date<20100201"


;$STATEMENT_bry = $query_last
;$STATEMENT_bry = $query_select

;$STATEMENT_bry = $query_insert

;$STATEMENT_bry = $query_update

;$STATEMENT_bry = $query_del

;$my=_bryant_mysql($DB_IP_bry,$DB_NAME_bry,$DB_ACCOUNT_bry,$DB_PASS_bry, $STATEMENT_bry)
;if IsArray ($my) then 
;	_ArrayDisplay($my)
;Else
;	MsgBox(0,"Affect on DB ", "Affect in DB '" & $DB_NAME_bry & "' by " & $my & " rows")
;
;EndIf



Func _bryant_mysql($DB_IP ,$DB_NAME, $DB_ACCOUNT, $DB_PASS , $STATEMENT )

;; This is  connect to My SQL
;;
_MySQL_InitLibrary()
If @error Then Exit MsgBox(0, '', "")
$MysqlConn = _MySQL_Init()
;Fehler Demo: C:\InstantRails-2.0-win\mysql\data\cookbook
;MsgBox(0,"Fehler-Demo","Fehler-Demo")
$connected = _MySQL_Real_Connect($MysqlConn, $DB_IP, $DB_ACCOUNT ,$DB_PASS,$DB_NAME)
If $connected = 0 Then
	$errno = _MySQL_errno($MysqlConn)
	MsgBox(0,"Error:",$errno & @LF & _MySQL_error($MysqlConn))
	return ("MySQL_ERROR " &$errno)
	If $errno = $CR_UNKNOWN_HOST Then MsgBox(0,"Error:","$CR_UNKNOWN_HOST" & @LF & $CR_UNKNOWN_HOST)
Endif



;$query_select = "SELECT * FROM hs_product where product_date<'20100201'"
	if StringInStr($STATEMENT,"SELECT")>0 then 
		_MySQL_Real_Query($MysqlConn, $STATEMENT)
		$res = _MySQL_Store_Result($MysqlConn)
		;$fields = _MySQL_Num_Fields($res)
		;MsgBox(0, '', "Zugriff Methode 3 - alles in ein 2D Array")
		$result_array = _MySQL_Fetch_Result_StringArray($res)
		;_ArrayDisplay($result_array)
		return ($result_array) ;Return array
	else
		;$query_insert= 
		if  _MySQL_Real_Query($MysqlConn, $STATEMENT)=0 then 
			$affected_row= _MySQL_Affected_Rows($MysqlConn)
			;MsgBox(0, "Insert row", "Insert Total " & $affected_row & " rows" ,5)
			return ($affected_row) ; a number of row to return
		Else
			Return("ERROR")
		EndIf
	EndIf

; Abfrage freigeben
_MySQL_Free_Result($res)
; Verbindung beenden
_MySQL_Close($MysqlConn)
; MYSQL beenden
_MySQL_EndLibrary()


EndFunc



Func _bryant_mysql_array($DB_IP ,$DB_NAME, $DB_ACCOUNT, $DB_PASS , $STATEMENT )

;; This is  connect to My SQL
local $my_affected=0
;;
if  IsArray ($STATEMENT) then
	
	
	_MySQL_InitLibrary()
	If @error Then Exit MsgBox(0, '', "")
	$MysqlConn = _MySQL_Init()
	;Fehler Demo: C:\InstantRails-2.0-win\mysql\data\cookbook
	;MsgBox(0,"Fehler-Demo","Fehler-Demo")
	$connected = _MySQL_Real_Connect($MysqlConn, $DB_IP, $DB_ACCOUNT ,$DB_PASS,$DB_NAME)
	If $connected = 0 Then
		$errno = _MySQL_errno($MysqlConn)
		MsgBox(0,"Error:",$errno & @LF & _MySQL_error($MysqlConn))
		If $errno = $CR_UNKNOWN_HOST Then MsgBox(0,"Error:","$CR_UNKNOWN_HOST" & @LF & $CR_UNKNOWN_HOST)
	Endif


	for $h=1 to $STATEMENT[0]
		if MOD($h,1000)=0 then sleep(5)
		$insert_statement=$STATEMENT[$h]
		;$query_select = "SELECT * FROM hs_product where product_date<'20100201'"
		if StringInStr($insert_statement,"SELECT")>0 then 
			_MySQL_Real_Query($MysqlConn, $insert_statement)
			$res = _MySQL_Store_Result($MysqlConn)
			;$fields = _MySQL_Num_Fields($res)
			;MsgBox(0, '', "Zugriff Methode 3 - alles in ein 2D Array")
			$result_array = _MySQL_Fetch_Result_StringArray($res)
			_ArrayDisplay($result_array)
			return ($result_array) ;Return array
		else
			;$query_insert= 
			if  _MySQL_Real_Query($MysqlConn, $insert_statement)=0 then 
				$affected_row= _MySQL_Affected_Rows($MysqlConn)
				;MsgBox(0, "Insert row", "Insert Total " & $affected_row & " rows" ,5)
				$my_affected=$my_affected+$affected_row  ; a number of row to return
			Else
				Return("ERROR")
			EndIf
		EndIf


	Next
	return ($my_affected)
	; Abfrage freigeben
	_MySQL_Free_Result($res)
	; Verbindung beenden
	_MySQL_Close($MysqlConn)
	; MYSQL beenden
	_MySQL_EndLibrary()

EndIf


EndFunc
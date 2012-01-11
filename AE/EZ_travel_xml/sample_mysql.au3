#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <array.au3>
#include <file.au3>
#include <Date.au3>
#Include <_XMLDomWrapper.au3>
#include <mysql.au3>

dim $sec=@SEC
dim $min=@MIN
Dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR



;InetGet("http://www.eztravel.com.tw/ezec/pkghsr/AeHsrInfo?queryDate=20101222", @ScriptDir&"\20101222.xml")
;exit
;;
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
Global $oMyRet[2]
;;
;; Connect My SQL for mail address.
;; DB is now at 10.112.55.87
;;
;dim $test_mode= _TEST_MODE()
dim $db_ip="127.0.0.1"
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
$connected = _MySQL_Real_Connect($MysqlConn,$db_ip,"amex","amextravel","test7")
If $connected = 0 Then
	$errno = _MySQL_errno($MysqlConn)
	MsgBox(0,"Error:",$errno & @LF & _MySQL_error($MysqlConn))
	If $errno = $CR_UNKNOWN_HOST Then MsgBox(0,"Error:","$CR_UNKNOWN_HOST" & @LF & $CR_UNKNOWN_HOST)
Endif

;;
;$query_update= "Update corp set ename='ImageColor' where corp_id='T15'"
;
;_MySQL_Real_Query($MysqlConn, $query_update)
;$res_update= _MySQL_Store_Result($MysqlConn)
;$update_array = _MySQL_Fetch_Result_StringArray($res_update)
;_ArrayDisplay($update_array)
;

$query_insert= "INSERT INTO hs_product (product_date,Tran_No,hs_class,DepTime,ArrTime,depCity,desCity) VALUES('20100101','399','N','0000','2229','110','115')"
;
if  _MySQL_Real_Query($MysqlConn, $query_insert)=0 then 

$affected_row= _MySQL_Affected_Rows($MysqlConn)
MsgBox(0, "Insert row", "Insert Total " & $affected_row & " rows" ,5)
EndIf

$query = "SELECT * FROM hs_product where product_date<'20100201'"

_MySQL_Real_Query($MysqlConn, $query)


$res = _MySQL_Store_Result($MysqlConn)

;$fields = _MySQL_Num_Fields($res)

;MsgBox(0, '', "Zugriff Methode 3 - alles in ein 2D Array")
$extramail_array = _MySQL_Fetch_Result_StringArray($res)
_ArrayDisplay($extramail_array)



; Abfrage freigeben
_MySQL_Free_Result($res)
; Verbindung beenden
_MySQL_Close($MysqlConn)
; MYSQL beenden
_MySQL_EndLibrary()

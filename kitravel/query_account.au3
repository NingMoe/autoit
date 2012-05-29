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
dim $test_mode= _TEST_MODE()

Global $magic_word = "趣遊網"

Global $version 


;MsgBox(0,"temp dir ", @TempDir)

;; 這段是為了開放下傳與否而寫的； 如果這個檔案在伺服器上不在了或是版本不對了，則不執行了
;;
;MsgBox(0,"on info",$os_partial)
;MsgBox(0,"",@UserProfileDir)
Dim $aData = InetRead("http://ivan:9ps5678@202.133.232.82:8080/upload/kitravel_query.htm")
Dim $aBytesRead = @extended
;MsgBox(4096, "", "Bytes read: " & $aBytesRead & @CRLF & @CRLF & StringLeft( BinaryToString($aData),3) & @CRLF & StringTrimLeft( BinaryToString($aData),4) )
$version = StringLeft(BinaryToString($aData), 3)

If $aBytesRead = 0 Or $version = "000" Then
	FileDelete (@TempDir & "\acc.bin")
	MsgBox(0, "錯誤", "這個程式己經失效了，" & @CRLF & "謝謝支持。")
	Exit
EndIf



if FileExists (@TempDir & "\acc.bin") then 
	$accbin_date= FileGetTime ( @TempDir & "\acc.bin" )
	;_ArrayDisplay ( $accbin_date )	
	$iDateCalc = _DateDiff( 'n',$accbin_date[0]&"/"&$accbin_date[1]&"/"&$accbin_date[2]&" 00:00:00",_NowCalc())
	;MsgBox(0,"time diff" , $accbin_date[0]&"/"&$accbin_date[1]&"/"&$accbin_date[2]&" 00:00:00" &@CRLF& _NowCalc() & @CRLF &$iDateCalc)
	if $iDateCalc > 21600 then FileDelete (@TempDir & "\acc.bin")
EndIf

if FileExists (@TempDir & "\acc.bin") then 
	$return_data=_newfile2Array(@TempDir & "\acc.bin", 9, "^",0)
	Else
	$return_data=_query_db()
	
		for $y=0 to UBound($return_data)-1
			;for $x=0 to 8
			;MsgBox(0,"$return_data[$y][3]" , $return_data[$y][1] & " >>>> " &StringToBinary( $return_data[$y][1] ) & @CRLF &$return_data[$y][3] & " >>>> " &StringToBinary( $return_data[$y][3] ))
			$return_data[$y][1]= StringToBinary( $return_data[$y][1] )
			$return_data[$y][3]= StringToBinary( $return_data[$y][3] )
			$return_data[$y][6]= StringToBinary( $return_data[$y][6] )
			$return_data[$y][7]= StringToBinary( $return_data[$y][7] )
			$return_data[$y][8]= StringToBinary( $return_data[$y][8] )
		Next
		if IsString (  $return_data   ) then
			$logfile = FileOpen(@TempDir & "\acc.bin", 10)
			FileWriteLine ( $logfile, _array2string_tab( $return_data ,9)  )
			FileClose($logfile)
		EndIf
EndIf	
	;_ArrayDisplay ($return_data)
	
for $y=0 to UBound($return_data)-1
	;for $x=0 to 8
	;MsgBox(0,"$return_data[$y][3]" , $return_data[$y][1] & " >>>> " &StringToBinary( $return_data[$y][1] ) & @CRLF &$return_data[$y][3] & " >>>> " &StringToBinary( $return_data[$y][3] ))
	$return_data[$y][1]= BinaryToString( $return_data[$y][1] )
	$return_data[$y][3]= BinaryToString( $return_data[$y][3] )
	$return_data[$y][6]= BinaryToString( $return_data[$y][6] )
	$return_data[$y][7]= BinaryToString( $return_data[$y][7] )
	$return_data[$y][8]= BinaryToString( $return_data[$y][8] )
Next
	;$arrayfindall= _ArrayFindAll($return_data, "台北豬窩", 0, 0, 0, 0, 1)

	;_ArrayDisplay ($arrayfindall)
	;_ArrayDisplay ($return_data)
	_King_SelectFileGUI()

;MsgBox(0,"Email", "There are :" & (UBound($extramail_array,1)-1) &" email addresses")

Exit




Func _King_SelectFileGUI() ; 取得二個檔案的名字，文字內容，及名單。
	local $open_default
	Local $hotel_name, $file_csv, $btn, $msg, $btn_n, $aEnc_info, $rc , $searched_name , $arrayfindall 
	local $hotel_id , $sms_to_delete ,$file , $searched_id
	;$open_default=_Open_default()
	;if $open_default=1 then 
		;run ("notepad.exe " & @ScriptDir & "\SMS_text.txt" )
		;run ("notepad.exe " & @ScriptDir & "\SMS_name_list.csv")
	
	;EndIf
		
	GUICreate("  民宿名查詢  ", 180, 100, 10, 10  , -1, 0x00000018); WS_EX_ACCEPTFILES
	GUICtrlCreateLabel(" 民宿名查詢 ", 10, 10, 160, 40)
	$hotel_name = GUICtrlCreateInput("", 10, 25, 160, 30, 0x0004)
	GUICtrlSetState(-1, $GUI_DROPACCEPTED)

	;GUICtrlCreateLabel("2.拖放簡訊名單檔案到這個框，預設為 SMS_name_list.csv", 10, 75, 300, 40)
	;$file_csv = GUICtrlCreateInput("", 10, 90, 300, 30)
	;GUICtrlSetState(-1, $GUI_DROPACCEPTED)

	;GUICtrlCreateLabel("2.  民宿 ID 查詢  ", 10, 70, 460, 40)
	;$hotel_id = GUICtrlCreateInput("", 10, 85, 460, 30)
	;GUICtrlSetState(-1, $GUI_FOCUS)
	;GUICtrlCreateInput("", 10, 35, 300, 20) 	; will not accept drag&drop files
	$btn = GUICtrlCreateButton("是", 40, 70, 60, 20, 0x0001) ; Default button
	$btn_n = GUICtrlCreateButton("否", 100, 70, 60, 20)
	GUISetState()

	;$msg = 0
	While $msg <> $GUI_EVENT_CLOSE
		 $sec = @SEC
		 $min = @MIN
		 $hour = @HOUR
		 $day = @MDAY
		 $month = @MON
		 $year = @YEAR

		$msg = GUIGetMsg()
		Select
			Case $msg = $GUI_EVENT_CLOSE
                ;MsgBox(0, "", "Dialog was closed")
                Exit

			Case $msg = $btn
				If Not GUICtrlRead($hotel_name) = ""  Then
					$searched_name=StringReplace(  GUICtrlRead($hotel_name) , @CRLF , "")
					
					$arrayfindall= _ArrayFindAll($return_data, $searched_name, 0, 0, 0, 1, 1)
					
					;_ArrayDisplay ( $arrayfindall )
					if UBound ($arrayfindall)=0 then 
						MsgBox (0," 名稱查詢 " ,  " 名稱查詢 : " & $searched_name & "找不到 " )
						ExitLoop
					EndIf
					
					local $extracted_array[UBound ($arrayfindall)][9]
					
					;_ArrayDisplay ($extracted_array)
					;
					for $s =0 to UBound ($arrayfindall)-1
					;	
						for $sx=0 to 8
							;MsgBox(0,"in the array", $return_data[ $arrayfindall[$s] ] [1] )
							$extracted_array[$s][$sx]= $return_data[ $arrayfindall[$s] ][$sx]
							;MsgBox(0,"value at extrated_array", $extracted_array[$s][$sx] )
							
						Next	
					;;	;if StringInStr ( $return_data[$s][1] , StringReplace(  GUICtrlRead($hotel_name) , @CRLF , "") ) then 
					;;	 ;MsgBox 
					;;	
					next
					_ArrayDisplay ($extracted_array, " 名稱查詢 : " & $searched_name)
					;MsgBox(4096, " Field 1",  "Field 1 " & @CRLF & GUICtrlRead($hotel_name) )
				;Else
					;$SMS_text_file = @ScriptDir & "\SMS_text.txt"
					;$name_list = @ScriptDir & "\SMS_name_list.csv"
					;$SMS_send_date = $year & $month & $day
				EndIf

				;If not ( GUICtrlRead($hotel_id) = "") then
				;	;MsgBox(0,"Date diff",_DateDiff( 'D',_NowCalcDate() ,GUICtrlRead($hotel_id)) )
	
				;EndIf
				;
				;ExitLoop
			Case $msg = $btn_n
				;_sms_maintain()
				;MsgBox(0,"Check in GUI select 2", $sms_delete)
				$SMS_text_file = ""
				$name_list = ""
				ExitLoop
				;Exit

		EndSelect
	WEnd

;return ( $SMS_text_file , $name_list )
GUIDelete();
EndFunc   ;==>_king_SelectFileGUI

Func _Open_default()

	If FileExists(@ScriptDir & "\open_default.txt") Then
		$mode = FileReadLine(@ScriptDir & "\open_default.txt", 1)
		If $mode = 1 Then
			;MsgBox(0, "Open Default", " 開啟檔案模式" , 5)

		Else
			;MsgBox(0,"Process mode", " 高鐵車次資料會輸入資料庫 ",10)
			;$ans=InputBox("Process mode","高鐵車次資料會輸入資料庫 "&@CRLF& "輸入 N 可以離開")

			$mode = 0
			;MsgBox(0, "Open Default", "手動模式" & @CRLF & " 手動開啟檔案模式", 5)
			;if $ans="n" or $ans="N" or @error=1 then exit
		EndIf

	Else
		;MsgBox(0,"Process mode", " 高鐵車次資料會輸入資料庫 ",10)
		;$ans=InputBox("Process mode","高鐵車次資料會輸入資料庫 "&@CRLF& "輸入 N 可以離開")

		$mode = 0
		;MsgBox(0, "Open Default", "手動模式" & @CRLF & " 手動開啟檔案模式", 5)
		;if $ans="n" or $ans="N" or @error=1 then exit

	EndIf

	Return $mode
EndFunc   ;==>_TEST_MODE

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
	
	for $y=1 to UBound($array)-1
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



;; 2011 09 New Text to Two dimension array  with a start_line
Func _newFile2Array($PathnFile, $aColume, $delimiters, $Start_line)


	Local $aRecords
	If Not _FileReadToArray($PathnFile, $aRecords) Then
		MsgBox(4096, "Error", " Error reading file '" & $PathnFile & "' to Array   error:" & @error)
		Exit
	EndIf

	;_ArrayDisplay ($aRecords)
	For $x = $Start_line to 1 step -1
		_ArrayDelete($aRecords,$x)
	Next
	$aRecords[0]= $aRecords[0]-$Start_line
	;_ArrayDisplay ($aRecords)
	;c
	Local $TextToArray[$aRecords[0]][$aColume + 1]
	;$TextToArray[0][0]=$aRecords[0]
	Local $aRow
	For $y = 1 To $aRecords[0]
		;Msgbox(0,'Record:' & $y, $aRecords[$y])

		$aRow = StringSplit($aRecords[$y], $delimiters)
		;Msgbox(0,'X ,Colume :', $aRow[0])
		;_ArrayDisplay ($aRow)
		For $x = 1 To $aRow[0]
			;If StringInStr($aRow[$x], ",") Then

			;	$aRow[$x] = StringTrimLeft($aRow[$x], 1)
			;	;MsgBox(0, "after", $aRow[$x])
			;EndIf
			$TextToArray[$y -1][$x -1] = $aRow[$x]
		Next
	Next

	;_ArrayDisplay($TextToArray)
	Return $TextToArray

EndFunc   ;==>_file2Array



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
$query = 'select b.HOTELID, b.HOTELNAME, a.UID, a.USERNAME, a.ACCOUNT, a.PASSWD, a.TEL, a.EMAIL, b.IDNAME from manager a, hotel b, working c where c.HOTELID=b.HOTELID and a.UID=c.UID and a.STATUS="E"and b.HOTELID>11 order by b.HOTELID'
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
#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.6.1
	Author:         myName

	Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <array.au3>
#include <mysql.au3>
#include <File.au3>
#include <Date.au3>
#include <FTPEx.au3>
#Include <GDIPlus.au3>
#include <_print_word.au3>

Dim $sec = @SEC
Dim $min = @MIN
Dim $hour = @HOUR
Dim $day = @MDAY
Dim $month = @MON
Dim $year = @YEAR
Dim $today = $year & $month & $day
Dim $db_return_data, $return_hotel_data_array, $accbin_date , $istrue
;
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
Global $oMyRet[2]
Global $work_dir=@ScriptDir & "\hotel_data_files\"
global $finished_dir=@ScriptDir&"\finished\"
;;
Dim $test_mode = _TEST_MODE()

;_myftp_upload("E:\workspace", 111)


;MsgBox(0,"Wait ftp ", "Wait for FTP.")
If Not FileExists($work_dir) Then DirCreate($work_dir)
If Not FileExists($finished_dir) Then DirCreate($finished_dir)

If FileExists(@ScriptDir & "\hotel.txt") Then
	$accbin_date = FileGetTime(@ScriptDir & "\hotel.txt")
	;_ArrayDisplay ( $accbin_date )
	$iDateCalc = _DateDiff('n', $accbin_date[0] & "/" & $accbin_date[1] & "/" & $accbin_date[2] & " 00:00:00", _NowCalc())
	;MsgBox(0,"time diff" , $accbin_date[0]&"/"&$accbin_date[1]&"/"&$accbin_date[2]&" 00:00:00" &@CRLF& _NowCalc() & @CRLF &$iDateCalc)
	If $iDateCalc > 10080 Then FileDelete(@ScriptDir & "\hotel.txt")
EndIf

If FileExists(@ScriptDir & "\hotel.txt") Then
	$return_hotel_data_array = _newFile2Array(@ScriptDir & "\hotel.txt", 5, "^", 0)
Else
	$return_hotel_data_array = _query_db()


	If IsString($return_hotel_data_array) Then
		$logfile = FileOpen(@ScriptDir & "\hotel.txt", 10)
		FileWriteLine($logfile, _array2string_tab($return_hotel_data_array, 5))
		FileClose($logfile)
	EndIf
EndIf
	;_ArrayDisplay ($return_hotel_data_array)

While 1
		$istrue = _process_data_file()
	if $istrue=-1 then exitloop
WEnd

Exit


func _process_data_file ()
	local $showup_msg=""
	local  $data_files = _FileListToArray($work_dir) ; 移入的 doc 及 xls
	local $hotel_id, $hotel_name, $hotel_phone,$hotel_fax, $hotel_address, $no ,$no_e


	;_ArrayDisplay($data_files)
	if isarray ($data_files ) then


			$no = _ArraySearch($data_files, "-訂房須知.doc", 0, 0, 0, 1)
			$no_e = _ArraySearch($data_files, "-訂房須知_e.doc", 0, 0, 0, 1)

			;MsgBox(0,"DOC ", "-訂房須知.doc  :  " &  $no  &@CRLF  & "-訂房須知_e.doc : " & $no_e )
			;$no= $no*$no_e
			if $no > 0 or  $no_e>0  then


				if $no>0 then
					$the_doc_file=$data_files[$no]
					$hotel_name = StringLeft($data_files[$no], StringInStr($data_files[$no], "-訂房須知.doc") - 1)
				Else
					$the_doc_file=$data_files[$no_e]
					$hotel_name = StringLeft($data_files[$no], StringInStr($data_files[$no], "-訂房須知_e.doc") - 1)
				EndIf

				$file_same_hid= _ArrayFindAll ($data_files , $hotel_name ,0,0,0,1,1)
				;MsgBox(0, " Array search ", $hotel_name)
				;_ArrayDisplay ( $file_same_hid)
				$hotel_id =  	$return_hotel_data_array [  _ArraySearch ( $return_hotel_data_array, $hotel_name ,0,0,0,1,0,1) ][0]
				$hotel_phone =  $return_hotel_data_array [  _ArraySearch ( $return_hotel_data_array, $hotel_name ,0,0,0,1,0,1) ][2]
				$hotel_fax =  	$return_hotel_data_array [  _ArraySearch ( $return_hotel_data_array, $hotel_name ,0,0,0,1,0,1) ][3]
				$hotel_address= $return_hotel_data_array [  _ArraySearch ( $return_hotel_data_array, $hotel_name ,0,0,0,1,0,1) ][4]
				;MsgBox(0, " Array search ", $hotel_id &" -- "& $hotel_phone&" -- "& $hotel_fax &" -- "& $hotel_address)

				;for $x=0 to UBound($file_same_hid)-1
				;	$showup_msg= ($x+1)&","& $data_files[$x+1]& " ; " &$showup_msg
				;Next
				;MsgBox(0,"file name", $showup_msg)
				_move_and_print( $the_doc_file, $hotel_name , $hotel_id , $hotel_phone, $hotel_fax, $hotel_address)
				$no=1
			Else
				$no=-1
			EndIf
	EndIf
	return $no
EndFunc


Func _move_and_print($doc_file , $h_name ,$h_id, $hotel_phone, $hotel_fax, $hotel_address)
	local $all_doc_in_hid , $s, $s1 , $swap
	;local $doc_value=0


	if not FileExists ( $work_dir & $h_id) then DirCreate( $work_dir  & $h_id )

	FileMove ( $work_dir& "*"& $h_name &"*" , $work_dir & $h_id )


	$all_doc_in_hid=_FileListToArray($work_dir  & $h_id, "*.doc" )
	_ArraySort(  $all_doc_in_hid ,0 )

if UBound($all_doc_in_hid)>0 then
	$s=0
	$s1=0
	$s = _ArraySearch($all_doc_in_hid, "-入住須知.doc", 0, 0, 0, 1)
	;ConsoleWrite (@CRLF& _ArraySearch($all_doc_in_hid, "-入住須知.doc", 0, 0, 0, 1) & @CRLF)
	if $s >0 then
		$swap=""
		$s1 = _ArraySearch($all_doc_in_hid, "-訂房須知.doc", 0, 0, 0, 1)
			if $s1< $s then
				$swap=$all_doc_in_hid[$s1]
				$all_doc_in_hid[$s1]= $all_doc_in_hid[$s]
				$all_doc_in_hid[$s]=$swap
			EndIf

	EndIf

	$s=0
	$s1=0
	$s = _ArraySearch($all_doc_in_hid, "-入住須知_e.doc", 0, 0, 0, 1)
	if $s >0 then
		$swap=""
		$s1 = _ArraySearch($all_doc_in_hid, "-訂房須知_e.doc", 0, 0, 0, 1)
		if $s1< $s then
			$swap=$all_doc_in_hid[$s1]
			$all_doc_in_hid[$s1]= $all_doc_in_hid[$s]
			$all_doc_in_hid[$s]=$swap
		EndIf

	EndIf

	$s=0
	$s1=0
	$s = _ArraySearch($all_doc_in_hid, "-訂房須知_e.doc", 0, 0, 0, 1)
	if $s >0 then
		$swap=""
		$s1 = _ArraySearch($all_doc_in_hid, "-訂房須知.doc", 0, 0, 0, 1)
		if $s1< $s then
			$swap=$all_doc_in_hid[$s1]
			$all_doc_in_hid[$s1]= $all_doc_in_hid[$s]
			$all_doc_in_hid[$s]=$swap
		EndIf

	EndIf



EndIf
	;_ArrayDisplay( $all_doc_in_hid )


	if UBound($all_doc_in_hid)>0 then
		if $all_doc_in_hid[0]=1 then
		;MsgBox(0,"Print 1 file only", $all_doc_in_hid[1] &" ; "& $h_name &" ; "& $h_id , 5)
		_print_word($all_doc_in_hid[1] , $h_name ,$h_id, $hotel_phone, $hotel_fax, $hotel_address)
		endif
		;MsgBox(0,"Print file",$doc_file &" ; "& $h_name &" ; "& $h_id )
		if $all_doc_in_hid[0]>1 then
			for $a=1 to $all_doc_in_hid[0]
				; _print_word file 執行列印轉檔及產生 html
				;MsgBox(0,"Print files : " & $a &"/" & $all_doc_in_hid[0], $all_doc_in_hid[$a] &" ; "& $h_name &" ; "& $h_id )
				_print_word($all_doc_in_hid[$a] , $h_name ,$h_id, $hotel_phone, $hotel_fax, $hotel_address)
				sleep(3000)
			next
		EndIf
	EndIf
	;MsgBox(0,"ftp folder",  $work_dir & $h_id )
	_myftp_upload($work_dir & $h_id, $h_id)

	DirMove ($work_dir  , $finished_dir )
EndFunc

func _kiftp_upload($dir_2_ftp, $h_id)

	Local $server = '202.168.197.252'
	Local $username = 'bryant'
	Local $pass = '9ps567*9'

	Local $Open = _FTP_Open('MyFTP Control')
	Local $Conn = _FTP_Connect($Open, $server, $username, $pass,1,6002)
	_FTP_DirCreate ($Conn, "/kitravel_Tomcat6/webapps/ROOT/upload/" & $h_id )
	_FTP_DirPutContents( $Conn,$dir_2_ftp,"/kitravel_Tomcat6/webapps/ROOT/upload/"& $h_id,0 )
	; ...
	Local $Ftpc = _FTP_Close($Open)



EndFunc


func _myftp_upload($dir_2_ftp, $h_id)

	;Local $server = '202.168.197.243'
	;Local $username = 'ivan'
	;Local $pass = '9ps5678'


	Local $server = '202.168.197.243'
	Local $username = 'kitravel'
	Local $pass = '2796!233'

	Local $Open = _FTP_Open('MyFTP Control')
	Local $Conn = _FTP_Connect($Open, $server, $username, $pass,1,6002)
	;_FTP_DirCreate ($Conn, "/kitravel/" & $h_id )
	;_FTP_DirPutContents( $Conn,$dir_2_ftp,"/kitravel/"& $h_id,0 )
	_FTP_DirCreate ($Conn, "/upload/" & $h_id )
	_FTP_DirPutContents( $Conn,$dir_2_ftp,"/upload/"& $h_id,0 )
	; ...
	Local $Ftpc = _FTP_Close($Open)



EndFunc



Func _TEST_MODE()
	If FileExists(@ScriptDir & "\TESTMODE.txt") Then
		$mode = FileReadLine(@ScriptDir & "\TESTMODE.txt", 1)
		If $mode = 1 Then
			MsgBox(0, "Test mode", "This is Test mode. ", 5)

		Else
			;MsgBox(0,"Delivery mode", "This is True delivery.",5)
			$mode = 0
		EndIf

	Else
		;MsgBox(0,"Delivery mode", "This is True  delivery.",5)
		$mode = 0
	EndIf

	Return $mode
EndFunc   ;==>_TEST_MODE

Func _array2string_tab($array, $d)
	; $d 是傳入的 array 的欄位數目
	Local $y, $x, $i, $string_to_return, $a_line
	$string_to_return = ''

	;MsgBox(0, "Dimention", "  $Y :" & UBound($array))

	For $y = 1 To UBound($array) - 1
		$a_line = ''
		If $d = 1 Then
			$a_line = $array[$y]

		Else
			For $x = 0 To $d - 1
				If $x < $d - 1 Then
					$a_line = $a_line & $array[$y][$x] & '^'
				Else
					$a_line = $a_line & $array[$y][$x]
				EndIf
				;MsgBox(0," array 2 string tab", $a_line)
				;ConsoleWrite($a_line & @crlf)
				;$string_to_return=$string_to_return  & $array[$y][1] &" , "& $array[$y][2] &" , " &$array[$y][6] &@CRLF
			Next
			ConsoleWrite($a_line & @CRLF)
			;$a_line=StringTrimRight($a_line,3) ; Cut off the last 3 character --> ","
		EndIf
		$string_to_return = $string_to_return & $a_line & @CRLF
	Next


	;MsgBox(0,"string", $string_to_return)
	Return $string_to_return
EndFunc   ;==>_array2string_tab



;; 2011 09 New Text to Two dimension array  with a start_line
Func _newFile2Array($PathnFile, $aColume, $delimiters, $Start_line)


	Local $aRecords
	If Not _FileReadToArray($PathnFile, $aRecords) Then
		MsgBox(4096, "Error", " Error reading file '" & $PathnFile & "' to Array   error:" & @error)
		Exit
	EndIf

	;_ArrayDisplay ($aRecords)
	For $x = $Start_line To 1 Step -1
		_ArrayDelete($aRecords, $x)
	Next
	$aRecords[0] = $aRecords[0] - $Start_line
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
			$TextToArray[$y - 1][$x - 1] = $aRow[$x]
		Next
	Next

	;_ArrayDisplay($TextToArray)
	Return $TextToArray

EndFunc   ;==>_newFile2Array




Func _query_db()
	;; Connect My SQL for mail address.
	;; DB is now at 10.112.55.87
	;;
	;dim $db_ip="changtuntv.dyndns.org"
	Dim $db_ip = "www.kitravel.com.tw"
	;;
	;; This is  connect to My SQL for user email
	;;
	_MySQL_InitLibrary()
	If @error Then Exit MsgBox(0, 'DB error', " My SQL init error")
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
	$query = "SELECT hotel.HOTELID, hotel.HOTELNAME, hotel.TEL, hotel.FAX, hotel.ADDRESS from hotel "
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
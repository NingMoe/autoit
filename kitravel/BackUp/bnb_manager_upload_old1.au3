#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <array.au3>
#include <mysql.au3>
#Include <File.au3>
#Include <Date.au3>


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



	 if not FileExists (@ScriptDir & "\hotel_data_files\") then DirCreate( @ScriptDir & "\hotel_data_files")

if FileExists (@ScriptDir & "\hotel.txt") then 
	$accbin_date= FileGetTime ( @ScriptDir & "\hotel.txt" )
	;_ArrayDisplay ( $accbin_date )	
	$iDateCalc = _DateDiff( 'n',$accbin_date[0]&"/"&$accbin_date[1]&"/"&$accbin_date[2]&" 00:00:00",_NowCalc())
	;MsgBox(0,"time diff" , $accbin_date[0]&"/"&$accbin_date[1]&"/"&$accbin_date[2]&" 00:00:00" &@CRLF& _NowCalc() & @CRLF &$iDateCalc)
	if $iDateCalc > 10080 then FileDelete ( @ScriptDir & "\hotel.txt" )
EndIf

if FileExists ( @ScriptDir & "\hotel.txt") then 
	$return_data=_newfile2Array(@ScriptDir & "\hotel.txt", 5, "^",0)
	Else
	$return_data=_query_db()
	

		if IsString (  $return_data   ) then
			$logfile = FileOpen(@ScriptDir & "\hotel.txt", 10)
			FileWriteLine ( $logfile, _array2string_tab( $return_data ,5)  )
			FileClose($logfile)
		EndIf
EndIf	
	;_ArrayDisplay ($return_data)
	dim $data_files= _FileListToArray (@ScriptDir & "\hotel_data_files\")
	_ArrayDisplay ($data_files)
	$no = _ArraySearch( $data_files, "-訂房須知.doc" ,0,0,0,1)
	 $hotel_name= stringleft ( $data_files[$no] ,  StringInStr($data_files [$no], "-訂房須知.doc")-1  )
	 
	MsgBox (0," Array search ", $hotel_name )
	


Exit






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
	
	MsgBox(0,"Dimention", "  $Y :" &UBound($array) )
	
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
				;MsgBox(0," array 2 string tab", $a_line)
				;ConsoleWrite($a_line & @crlf)
				;$string_to_return=$string_to_return  & $array[$y][1] &" , "& $array[$y][2] &" , " &$array[$y][6] &@CRLF
			next
			ConsoleWrite($a_line & @crlf)
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
If @error Then Exit MsgBox(0, '', "")
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

return $db_return_data
EndFunc
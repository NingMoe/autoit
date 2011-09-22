#include <array.au3>
#include <File.au3>
#include <Date.au3>
;#include <CompInfo_win7.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GUIConstants.au3>
#include <mail_variable.au3>


Global $SMS_file ; =@ScriptDir&"\1316641706.sms" ; 這個由 _SelectFileGUI() 這個 func 得到
Dim $Name_list ; = @ScriptDir& "\SMS_name_list.csv"   這個由 _newFile2Array() 這個 func 得到
Dim $SMS_detail ;
Dim $sec = @SEC
Dim $min = @MIN
Dim $hour = @HOUR
Dim $day = @MDAY
Dim $month = @MON
Dim $year = @YEAR

Global $test_mode ;, $current_time
Global $SMS_Feed_FileTime=0
Global $SMS_Feed_List_PreProcess , $SMS_Feed_List[1]
Global $now_DateCalc
$SMS_Feed_List[0]=0
while 1
	if FileExists (@ScriptDir &"\sms_feed") then 
		;MsgBox(0,"File Get time", FileGetTime (@ScriptDir &"\sms_feed",0,1) )
		if $SMS_Feed_FileTime < FileGetTime (@ScriptDir &"\sms_feed",0,1)  then 
			$SMS_Feed_List_PreProcess=_Move_SMS_feed(@ScriptDir &"\sms_feed")
		EndIf
		_ArrayDelete($SMS_Feed_List_PreProcess,0 )
		;_ArrayDisplay ($SMS_Feed_List_PreProcess)
	EndIf	
	if  IsArray( $SMS_Feed_List_PreProcess ) then 
		_ArrayConcatenate ( $SMS_Feed_List , $SMS_Feed_List_PreProcess )
		$SMS_Feed_List_PreProcess=''
		_ArraySort ($SMS_Feed_List,0,1)
		$SMS_Feed_List[0]= UBound ($SMS_Feed_List)-1
		_ArrayDisplay ($SMS_Feed_List)
	EndIf
	if  $SMS_Feed_List[0] > 0 then 
		$now_DateCalc = _DateDiff( 's',"1970/01/01 00:00:00",_NowCalc())
		;MsgBox(0,"EPOCH", $now_DateCalc,3)
		if  ( StringLeft ( $SMS_Feed_List[1] ,10 ) - $now_DateCalc ) < -120 Then
			_ArrayDelete ($SMS_Feed_List,1)
			$SMS_Feed_List[0]= UBound ($SMS_Feed_List)-1
		EndIf
		
		if UBound($SMS_Feed_List )>1	then 
			if ( StringLeft ( $SMS_Feed_List[1] ,10 ) - $now_DateCalc ) >= -120  or ( StringLeft ( $SMS_Feed_List[1] ,10 ) - $now_DateCalc ) < 300 Then
				MsgBox(0, "Now and SendSMS time diff: " & $now_DateCalc , (   StringLeft ( $SMS_Feed_List[1] ,10 ) -$now_DateCalc  )  ,5)
				$SMS_file = @ScriptDir & "\" &  StringTrimLeft ($SMS_Feed_List[1], StringInStr ( $SMS_Feed_List[1],"," ) ) & "\" & StringLeft ($SMS_Feed_List[1], StringInStr ( $SMS_Feed_List[1],"," )-1 ) &".sms" 
				MsgBox (0,"SMS file: ",$SMS_file)
				$SMS_detail= _SMS_detail($SMS_file,1,5)
				$Name_list= _newFile2Array($SMS_file,2, ",", 6) 
			EndIf	
		EndIf
		
	EndIf
	sleep(333)
WEnd

;$SMS_detail= _SMS_detail($SMS_file,1,5)
;$Name_list= _newFile2Array($SMS_file,2, ",", 6) 
;;_ArrayDisplay ( $SMS_detail)
;;_ArrayDisplay ( $Name_list )
exit

Func _Move_SMS_feed( $PathnFile)
	Local $aRecords
; Another sample which automatically creates the directory structure
While 1
	If Not _FileReadToArray($PathnFile, $aRecords) Then
			MsgBox(4096, "Error", " Error reading file '" & $PathnFile & "' to Array   error:" & @error)
			;Exit
		sleep(75)
	Else
		$file_creat_time=FileGetTime (@ScriptDir &"\sms_feed",0,1)
		filemove ( @ScriptDir &"\sms_feed" , @ScriptDir &"\feed\"&$file_creat_time&".sms_feed" ,9)
		if FileExists (@ScriptDir &"\feed\"&$file_creat_time&".sms_feed") then ExitLoop
	EndIf

WEnd
_ArrayDisplay ( $aRecords)
return ($aRecords)
EndFunc


Func _SMS_detail ($PathnFile, $Start_line, $End_line)
	
	
	Local $aRecords
	If Not _FileReadToArray($PathnFile, $aRecords) Then
		MsgBox(4096, "Error", " Error reading file '" & $PathnFile & "' to Array   error:" & @error)
		Exit
	EndIf
	
	;_ArrayDisplay ($aRecords)
	For $x = $End_line+1 to $aRecords[0] 
		_ArrayDelete($aRecords,$x)
	Next
	$aRecords[0]= $End_line
	_ArrayDisplay ($aRecords)

return $aRecords
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
		For $x = 1 To $aRow[0]
			If StringInStr($aRow[$x], ",") Then
				
				$aRow[$x] = StringTrimLeft($aRow[$x], 1)
				;MsgBox(0, "after", $aRow[$x])
			EndIf
			$TextToArray[$y - 1][$x - 1] = $aRow[$x]
		Next
	Next
	
	;_ArrayDisplay($TextToArray)
	Return $TextToArray

EndFunc   ;==>_file2Array


Func _TEST_MODE()
	
	If FileExists(@ScriptDir & "\TESTMODE.txt") Then
		$mode = FileReadLine(@ScriptDir & "\TESTMODE.txt", 1)
		If $mode = 1 Then
			MsgBox(0, "Test mode", "測試模式" & @CRLF & "只會寄送到 Service Account ", 5)
			
		Else
			;MsgBox(0,"Process mode", " 高鐵車次資料會輸入資料庫 ",10)
			;$ans=InputBox("Process mode","高鐵車次資料會輸入資料庫 "&@CRLF& "輸入 N 可以離開")
			
			$mode = 0
			MsgBox(0, "Test mode", "正式模式" & @CRLF & "Order 備份檔案會寄送到所屬主人的信箱 ", 5)
			;if $ans="n" or $ans="N" or @error=1 then exit
		EndIf
		
	Else
		;MsgBox(0,"Process mode", " 高鐵車次資料會輸入資料庫 ",10)
		;$ans=InputBox("Process mode","高鐵車次資料會輸入資料庫 "&@CRLF& "輸入 N 可以離開")
		
		$mode = 0
		MsgBox(0, "Test mode", "正式模式" & @CRLF & "Order 備份檔案會寄送到所屬主人的信箱 ", 5)
		;if $ans="n" or $ans="N" or @error=1 then exit
		
	EndIf
	
	Return $mode
EndFunc   ;==>_TEST_MODE

Func _EPOCH ($DateToCalc)
; Calculated the number of seconds since EPOCH (1970/01/01 00:00:00) 
$iDateCalc = _DateDiff( 's',"1970/01/01 00:00:00",$DateToCalc ) ;_NowCalc())

;MsgBox( 4096, "_EPOCH Func", "EPOCH Time to send : " & $iDateCalc  & @CRLF  & "time lap to Current : " & ($iDateCalc - _DateDiff( 's',"1970/01/01 00:00:00", _NowCalc()) ) )
return $iDateCalc
EndFunc
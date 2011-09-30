;
;
;
; Once receive the upload , Send a text SMS to user. And then a mail.  Done
; After send all the SMS, Send a mail for the bill. and a SMS for notice.  Done
; After send mail , remove the Epoch.sms files. 
; Need to test array delete feature at line 50
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
dim $array_component_to_hunt, $is_delete_feed=0
Global $test_mode ;, $current_time
Global $SMS_Feed_FileTime=0
Global $SMS_Feed_List_PreProcess , $SMS_Feed_List[1]
Global $now_DateCalc

Dim $read_from_feed_text
$SMS_Feed_List[0]=0

$test_mode = _TEST_MODE()
if $test_mode then 
_Gen_Test_Data()
if not FileExists (@ScriptDir & "\sms_feed") then MsgBox(0, "Warning", "No test data Generated!" ,5)
EndIf
;MsgBox(0,"Wait", "Wait for check")
;

$read_from_feed_text= _read_feed_text(@ScriptDir & "\feed_text.txt")
if IsArray($read_from_feed_text) then 
	;_ArrayDisplay ($read_from_feed_text)
	_ArrayDelete($read_from_feed_text , 0)
	_ArrayConcatenate ( $SMS_Feed_List ,$read_from_feed_text )
		;$SMS_Feed_List_PreProcess=''
		_ArraySort ($SMS_Feed_List,0,1)
		$SMS_Feed_List[0]= UBound ($SMS_Feed_List)-1
	;_ArrayDisplay($SMS_Feed_List, "First read from feed_text.txt")
Else
	MsgBox(0,"No feed_text.txt file to read", "No file at " & @ScriptDir & "\feed_text.txt" , 5)
EndIf	
	

while 1
	if FileExists (@ScriptDir &"\sms_feed") then 
		;$now_DateCalc = _DateDiff( 's',"1970/01/01 00:00:00",_NowCalc())
		;MsgBox(0,"File Get time", FileGetTime (@ScriptDir &"\sms_feed",0,1) )
		if $SMS_Feed_FileTime < FileGetTime (@ScriptDir &"\sms_feed",0,1)  then 
			$SMS_Feed_List_PreProcess=_Move_SMS_feed(@ScriptDir &"\sms_feed")
			$SMS_Feed_FileTime = FileGetTime (@ScriptDir &"\sms_feed",0,1)
		EndIf
		_ArrayDelete($SMS_Feed_List_PreProcess,0 )
		;_ArrayDisplay ($SMS_Feed_List_PreProcess)
		$is_delete_feed=_check_sender_SMS_detail ($SMS_Feed_List_PreProcess)
		if not $is_delete_feed then 
		_make_sender_SMS_detail($SMS_Feed_List_PreProcess) ;; Think this is only one-dimention array, and with one record only.
		else 
			$array_component_to_hunt =_ArraySearch($SMS_Feed_List,$SMS_Feed_List_PreProcess[0],1,$SMS_Feed_List[0],1) 
			if  $array_component_to_hunt >=1 then 
				_ArrayDelete($SMS_Feed_List, $array_component_to_hunt )
				$SMS_Feed_List[0]-=1
				$SMS_Feed_List_PreProcess=''
			EndIf
				;_ArrayDisplay ($SMS_Feed_List,"After delete_feed action")
		EndIf
	EndIf	
	if  IsArray( $SMS_Feed_List_PreProcess ) then 
		_ArrayConcatenate ( $SMS_Feed_List , $SMS_Feed_List_PreProcess )
		$SMS_Feed_List_PreProcess=''
		_ArraySort ($SMS_Feed_List,0,1)
		$SMS_Feed_List[0]= UBound ($SMS_Feed_List)-1
		;_ArrayDisplay ($SMS_Feed_List)
		_write_feed_text($SMS_Feed_List)
	else 
		$SMS_Feed_List_PreProcess=''	
	EndIf
	
	if  $SMS_Feed_List[0] > 0 then 
		$now_DateCalc = _DateDiff( 's',"1970/01/01 00:00:00",_NowCalc())
		;MsgBox(0,"EPOCH", $now_DateCalc,3)
		_ArrayDisplay ($SMS_Feed_List)
		ToolTip("Load_to_sms. Next load: "& ( StringLeft ( $SMS_Feed_List[1] ,10 ) - $now_DateCalc ) & " Sec " &@CRLF & $SMS_Feed_List[1],300,700 )
		;MsgBox(0,"Time lap ", ( StringLeft ( $SMS_Feed_List[1] ,10 ) - $now_DateCalc ) )
		if  ( StringLeft ( $SMS_Feed_List[1] ,10 ) - $now_DateCalc ) < -120 Then
			FileMove ($SMS_file,@ScriptDir & "\" &  StringTrimLeft ($SMS_Feed_List[1], StringInStr ( $SMS_Feed_List[1],"," ) ) & "\" & StringLeft ($SMS_Feed_List[1], StringInStr ( $SMS_Feed_List[1],"," )-1 ) &".sms.omit")
			_ArrayDelete ($SMS_Feed_List,1)
			$SMS_Feed_List[0]= UBound ($SMS_Feed_List)-1
			_write_feed_text($SMS_Feed_List)
		EndIf
		
		if UBound($SMS_Feed_List )>1	then 
			if ( StringLeft ( $SMS_Feed_List[1] ,10 ) - $now_DateCalc ) >= -120  and ( StringLeft ( $SMS_Feed_List[1] ,10 ) - $now_DateCalc ) < 300 Then
				ToolTip("Now and SendSMS time diff: " &  StringLeft ( $SMS_Feed_List[1] ,10 ) -$now_DateCalc ,300,700)
				;MsgBox(0, "Now and SendSMS time diff: " & $now_DateCalc , (   StringLeft ( $SMS_Feed_List[1] ,10 ) -$now_DateCalc  )  ,5)
				$SMS_file = @ScriptDir & "\" &  StringTrimLeft ($SMS_Feed_List[1], StringInStr ( $SMS_Feed_List[1],"," ) ) & "\" & StringLeft ($SMS_Feed_List[1], StringInStr ( $SMS_Feed_List[1],"," )-1 ) &".sms" 
				;MsgBox (0,"SMS file: ",$SMS_file)
				ToolTip("Now Send SendSMS file: " & $SMS_file ,300,700)
				$SMS_detail= _SMS_detail($SMS_file,1,5)			  ; 	
				$Name_list= _newFile2Array($SMS_file,2, ",", 6)   ;  _newFile2Array($PathnFile, $aColume, $delimiters, $Start_line)
				
				
				_ProcessSendMail_to_SMS($Name_list ,$SMS_detail ,0)
				FileMove ($SMS_file,@ScriptDir & "\" &  StringTrimLeft ($SMS_Feed_List[1], StringInStr ( $SMS_Feed_List[1],"," ) ) & "\" & StringLeft ($SMS_Feed_List[1], StringInStr ( $SMS_Feed_List[1],"," )-1 ) &".sms.out")
				_ArrayDelete ($SMS_Feed_List,1)
				$SMS_Feed_List[0]= UBound ($SMS_Feed_List)-1
				_write_feed_text($SMS_Feed_List)
			EndIf	
		EndIf
		
		sleep(1333)
		ToolTip("")
	EndIf
	if $SMS_Feed_List[0]=0 then MsgBox(0,"Warning", "No sms_feed now" , 3)
	sleep(333)
	;_ArrayDisplay ($SMS_Feed_List)
WEnd

;$SMS_detail= _SMS_detail($SMS_file,1,5)
;$Name_list= _newFile2Array($SMS_file,2, ",", 6) 
;;_ArrayDisplay ( $SMS_detail)
;;_ArrayDisplay ( $Name_list )
exit

Func _write_feed_text($write_array)
	local $file , $r 
	if isarray( $write_array) then 
		$file = FileOpen(@ScriptDir & "\feed_text.txt",10)
		for $r =1 to $write_array[0]
			FileWriteLine($file, $write_array[$r])
		Next	
		FileClose ($file)
	EndIf
EndFunc

Func _read_feed_text($PathnFile)
	local $aRecords , $r ,$file
	if FileExists ($PathnFile) then 
		If Not _FileReadToArray($PathnFile, $aRecords) Then
			MsgBox(4096, "Error on _feed_text() function", " Error reading file '" & $PathnFile & "' to Array   error:" & @error, 10)
			;Exit
		EndIf
		;_ArrayDisplay($aRecords)
		for $r = UBound($aRecords)-1 to 1 step -1
			$file = @ScriptDir & "\" &  StringTrimLeft ($aRecords[$r], StringInStr ( $aRecords[$r],"," ) ) & "\" & StringLeft ($aRecords[$r], StringInStr ( $aRecords[$r],"," )-1 ) &".sms" 
			if not FileExists($file) then 
				_ArrayDelete( $aRecords, $r )
				$aRecords[0]= $aRecords[0]-1
			EndIf
		Next
	EndIf
	;_ArrayDisplay($aRecords)
return $aRecords
EndFunc

Func _check_sender_SMS_detail ($sender_sms_feed)
	local $SMS_file_senders, $SMS_detail_senders , $Name_list_senders[1][2] , $sender  , $delete_feed=0
	
				$SMS_file_senders = @ScriptDir & "\" &  StringTrimLeft ($sender_sms_feed[0], StringInStr ( $sender_sms_feed[0],"," ) ) & "\" & StringLeft ($sender_sms_feed[0], StringInStr ( $sender_sms_feed[0],"," )-1 ) &".sms" 
				if not FileExists ($SMS_file_senders) then  
					;MsgBox (0,"SMS file: ", "File is not exist :" & @CRLF &$SMS_file_senders,3)	; 需要再處理一下 error handeling.
					_FileWriteLog(@ScriptDir & "\logs\" & StringTrimRight(@ScriptName,4) & "_" & $year & $month & $day & ".log", " File is not exist : " & $SMS_file_senders )
				EndIf
				$SMS_detail_senders= _SMS_detail($SMS_file_senders,1,6)  ; 取得要發出的簡訊內容
				_ArrayDisplay($SMS_detail)
			if $SMS_detail_senders[5]=''  then
				filemove( $SMS_file_senders , $SMS_file_senders &".SenderOmitted ")
				$delete_feed=1
			EndIf		
return $delete_feed
EndFunc

Func _make_sender_SMS_detail($sender_sms_feed)
	local $SMS_file_senders, $SMS_detail_senders , $Name_list_senders[1][2] , $sender  
	
				$SMS_file_senders = @ScriptDir & "\" &  StringTrimLeft ($sender_sms_feed[0], StringInStr ( $sender_sms_feed[0],"," ) ) & "\" & StringLeft ($sender_sms_feed[0], StringInStr ( $sender_sms_feed[0],"," )-1 ) &".sms" 
				if not FileExists ($SMS_file_senders) then  
					;MsgBox (0,"SMS file: ", "File is not exist :" & @CRLF &$SMS_file_senders,3)	; 需要再處理一下 error handeling.
					_FileWriteLog(@ScriptDir & "\logs\" & StringTrimRight(@ScriptName,4) & "_" & $year & $month & $day & ".log", " File is not exist : " & $SMS_file_senders )
				EndIf
				$SMS_detail_senders= _SMS_detail($SMS_file_senders,1,5)  ; 取得要發出的簡訊內容
	
					
				;_ArrayDisplay($SMS_detail_senders)
				$sender= StringSplit("使用者,"&$SMS_detail_senders[4],"," )
				$Name_list_senders[0][0]= $sender[1]
				$Name_list_senders[0][1]= $sender[2]
				;$Name_list_senders= _newFile2Array($SMS_file_senders,2, ",", 6) ;; 這一行有問題，必需改為送件人的名字和行動電話
				;_ArrayDisplay($Name_list_senders)
				
				_ProcessSendMail_to_SMS($Name_list_senders , $SMS_detail_senders , 1)
EndFunc



Func _ProcessSendMail_to_SMS($name_list_array_parameter, $SMS_detail_parameter, $sender_feedback)  ;$name_list_array_parameter 是二維的 array  $SMS_detail_parameter 是一維的 Array.
	$m_BccAddress=""
	local $message , $sender_email_address , $sender_mobile, $name_list_title
	
	$name_list_array =$name_list_array_parameter   ; $name_list_array第一欄是名字，第二欄是電話
	$sender_email_address = $SMS_detail_parameter[3]
	$sender_mobile = $SMS_detail_parameter[4]
	$message = $SMS_detail_parameter[5]
	;$name_list_title = $SMS_detail_parameter[6]
	if not FileExists(@ScriptDir&"\logs\") then FileOpen(@ScriptDir&"\logs\",10)
For $r = 0 To (UBound($name_list_array, 1) - 1)
	
	;$m_AttachFiles = @ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"
	$as_Body = $name_list_array[$r][0] & "您好: " & $message

	;MsgBox (0,"This is name list array", "Send to Name : " & $name_list_array[$r][0]  & @CRLF & "Send to Mobile: " & $name_list_array[$r][1]  & @CRLF & " Mail body: " &$as_Body ,5)
	;$s_ToAddress = "sms@onlinebooking.com.tw"
	If StringInStr($name_list_array[$r][1], "_") Then
		$s_Subject = StringReplace($name_list_array[$r][1], "_", "")
		;$as_Body=StringReplace($as_Body,"[user_email]","ae@delta.com.tw") ; For test only.
	Else
		$s_Subject = $name_list_array[$r][1]
	EndIf
	;### This is for Corrrect SMS
	;$s_Subject="0919585516"
	if $test_mode=1 then 
		$m_Subject = "簡訊發送測試信件: " & $s_Subject ; subject from the email - can be anything you want it to be
		
		$rc = _INetSmtpMailCom($m_SmtpServer, $m_FromName, $m_FromAddress, $m_ToAddress, $m_Subject, $as_Body, $m_AttachFiles, $m_CcAddress, $m_BccAddress, $m_Username, $m_Password, $IPPort, $ssl)
		;$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $s_IPPort, $s_ssl)

	Else	
	MsgBox(0, "mail parameter", $s_SmtpServer & " / " & $s_FromName & " / " & $s_FromAddress & " / " & $s_ToAddress & " / " & $s_Subject & " / " & $as_Body & " / " & $s_AttachFiles & " / " & $s_CcAddress & " / " & $s_BccAddress & " / " & $s_Username & " / " & $s_Password & " / " & $s_IPPort & " / " & $s_ssl)
	;$rc = _INetSmtpMailCom_utf8($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $s_IPPort, $s_ssl)
	EndIf
	;### This is mail Test
	;$s_ToAddress="bryant@dynalab.com.tw"
	;MsgBox (0,"mail parameter", $s_SmtpServer&" / "& $s_FromName&" / "& $s_FromAddress&" / "& $s_ToAddress&" / "& $s_Subject&" / "& $as_Body&" / "& $s_AttachFiles&" / "& $s_CcAddress&" / "& $s_BccAddress&" / "& $s_Username&" / "& $s_Password&" / "& $s_IPPort&" / "& $s_ssl)
	;$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $s_IPPort, $s_ssl)
	;$rc = _INetSmtpMailCom($m_SmtpServer, $m_FromName, $m_FromAddress, $m_ToAddress, $m_Subject, $as_Body, $m_AttachFiles, $m_CcAddress, $m_BccAddress, $m_Username, $m_Password, $IPPort, $ssl)
	_FileWriteLog(@ScriptDir & "\logs\" & $sender_mobile & "_" & $year & $month & $day & ".log", $SMS_detail_parameter[1]& " SMS Send to : " & $s_Subject &" "& $name_list_array[$r][0])
	If $r > 0 And Mod($r, 5) = 0 Then
		Sleep(5000)
	EndIf
	If $r = (UBound($name_list_array, 1) - 1) and $sender_feedback=0 Then
		Dim $day = @MDAY
		Dim $month = @MON
		Dim $year = @YEAR
		
		$m_ToAddress= $sender_email_address
		$m_BccAddress = "bryant@net1.com.tw"
		$m_AttachFiles = @ScriptDir & "\logs\" & $sender_mobile  & "_" & $year & $month & $day & ".log"
		$rc = _INetSmtpMailCom($m_SmtpServer, $m_FromName, $m_FromAddress, $m_ToAddress, $m_Subject, $m_as_Body, $m_AttachFiles, $m_CcAddress, $m_BccAddress, $m_Username, $m_Password, $IPPort, $ssl)
		;$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $m_ToAddress, $m_Subject, $m_as_Body, $m_AttachFiles, $m_CcAddress, $s_BccAddress, $s_Username, $s_Password, $s_IPPort, $s_ssl)

		;Sleep(1000 * 60 * 5)
	EndIf


Next


EndFunc



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
;_ArrayDisplay ( $aRecords)
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
	;_ArrayDisplay ($aRecords)

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

Func _EPOCH ($DateToCalc)
; Calculated the number of seconds since EPOCH (1970/01/01 00:00:00) 
$iDateCalc = _DateDiff( 's',"1970/01/01 00:00:00",$DateToCalc ) ;_NowCalc())

;MsgBox( 4096, "_EPOCH Func", "EPOCH Time to send : " & $iDateCalc  & @CRLF  & "time lap to Current : " & ($iDateCalc - _DateDiff( 's',"1970/01/01 00:00:00", _NowCalc()) ) )
return $iDateCalc
EndFunc

Func _TEST_MODE()

	If FileExists(@ScriptDir & "\TESTMODE.txt") Then
		$mode = FileReadLine(@ScriptDir & "\TESTMODE.txt", 1)
		If $mode = 1 Then
			MsgBox(0, "Test mode", "測試模式" & @CRLF & "只會寄送到 Service Account ", 5)

		Else
			;MsgBox(0,"Process mode", " 高鐵車次資料會輸入資料庫 ",10)
			;$ans=InputBox("Process mode","高鐵車次資料會輸入資料庫 "&@CRLF& "輸入 N 可以離開")

			$mode = 0
			MsgBox(0, "Test mode", "正式模式" & @CRLF & "SMS 會寄送到所屬人的 Mobile", 5)
			;if $ans="n" or $ans="N" or @error=1 then exit
		EndIf

	Else
		;MsgBox(0,"Process mode", " 高鐵車次資料會輸入資料庫 ",10)
		;$ans=InputBox("Process mode","高鐵車次資料會輸入資料庫 "&@CRLF& "輸入 N 可以離開")

		$mode = 0
		MsgBox(0, "Test mode", "正式模式" & @CRLF & "SMS 會寄送到所屬人的 Mobile", 5)
		;if $ans="n" or $ans="N" or @error=1 then exit

	EndIf

	Return $mode
EndFunc   ;==>_TEST_MODE

Func _Gen_Test_Data()
	local $Data , $infome_data , $now_DateCalc , $file
	$now_DateCalc = ( _DateDiff( 's',"1970/01/01 00:00:00",_NowCalc()) ) + 100
	$data=$now_DateCalc &@CRLF &  _
			"2011/09/23 14:37:27" &@CRLF &  _
			"changtun@gmail.com"  &@CRLF & _
			"0928837823"  &@CRLF & _
			"當你有多個外面線路想要同時間對映到內部的某一台Server時，或許你是為了備援或Load Balance的考量。在Router" &@CRLF & _
			"姓名,行動電話"  &@CRLF &  _
			"changtun1,_0928837823"  &@CRLF & _
			"sean2,_0956330560"  &@CRLF & _
			"seanlu3,_0968269170"  &@CRLF & _ 
			"bryant4,_0928837823" 
	
	$infome_data=$now_DateCalc &",changtun"
	
	$file=fileopen (@ScriptDir & "\changtun\"& $now_DateCalc& ".sms",10)
	FileWrite($file, $data )
	FileClose($file)
	sleep(100)
	$file=fileopen (@ScriptDir & "\sms_feed",10)
	FileWrite($file,$infome_data)
	FileClose($file)
	sleep(100)
EndFunc


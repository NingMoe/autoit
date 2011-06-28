#include <array.au3>
#include <File.au3>
#include <Date.au3>
#include <CompInfo.au3>

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GUIConstants.au3>
#include <FTPEx.au3>



Global $SMS_text_file ; =@ScriptDir&"\SMS_text.txt" ; 這個由 _SelectFileGUI() 這個 func 得到
Global $name_list ; = @ScriptDir& "\SMS_name_list.csv"   這個由 _SelectFileGUI() 這個 func 得到
Global $oMyRet[2]
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")

Dim $sec = @SEC
Dim $min = @MIN
Dim $hour = @HOUR
Dim $day = @MDAY
Dim $month = @MON
Dim $year = @YEAR

Global $test_mode , $current_time

Global $magic_word = "民權國小天文社專用發簡訊密碼"
Global $astronomy = ".astronomy.txt"
Dim $os_partial , $email1 , $email2
Global $version , $user_name
Global $allow 
Global $Final_level_sms_list[1]
;$Final_level_sms_list[0]=0
;$test_mode=_TEST_MODE() ; return 1 means  Test mode.




while 1

scan_allow_SMS_list()
check_SMS_list()
;_open_sms_text_name(@ScriptDir&"\SMS_text.txt", @ScriptDir&"\SMS_name_list.csv")

WEnd


func check_SMS_list()
	local $current_name , $current_name_status ,$check_no
	for $check_no=0 to UBound ($Final_level_sms_list)-1
	$current_time = _DateDiff( 's',"1970/01/01 00:00:00",_NowCalc())
	;MsgBox(0,"", $current_time &"VS" & StringLeft( $Final_level_sms_list[$check_no], StringInStr($Final_level_sms_list[$check_no],"!" )-1 ) &"")
	  if  int ( $current_time ) >= int( StringLeft( $Final_level_sms_list[$check_no], StringInStr($Final_level_sms_list[$check_no],"!" )-1 ) )   Then
			MsgBox(0,"Send Message", $current_time &" VS " & StringLeft( $Final_level_sms_list[$check_no], StringInStr($Final_level_sms_list[$check_no],"!" )-1 )  & @CRLF &"Now Send Message.")
		EndIf
	next
EndFunc



Func scan_allow_SMS_list()
	; Scan new request and add to Final level sms list array
	dim $1st_level_sms_list , $2nd_level_sms_list , $load
	local $current_name , $current_name_status
	local $current_text , $current_text_status
	

;; First init the sw.  It will check if the directory which listed in allow.txt exists.	
	if FileExists(@ScriptDir&"\allow.txt") Then
		_FileReadToArray(@ScriptDir&"\allow.txt", $allow )
	Else
			MsgBox(0, "Error", "There is no allow.txt in script dir")	
			Exit
	EndIf
		;_ArrayDisplay($allow)
	if IsArray( $allow	) then 
		for $x=2 to UBound( $allow) -1
			local $file_list_array
			if not FileExists($allow[1]&"\"&$allow[$x]) then DirCreate($allow[1]&"\"&$allow[$x] )
		
			$file_list_array=_FileListToArray($allow[1]&"\"&$allow[$x],"*",1)
			;_ArrayDisplay($file_list_array)
		Next
	EndIf	
	
	if FileExists(@ScriptDir&"\load.txt") Then
		_FileReadToArray(@ScriptDir&"\load.txt", $Final_level_sms_list )
		if  IsInt( $Final_level_sms_list[0] ) then _ArrayDelete ($Final_level_sms_list,0)
		;_ArrayDisplay ($Final_level_sms_list)
	EndIf
	
	
	if FileExists($allow[1]& "\*.sms" ) then 
	$1st_level_sms_list= _FileListToArray($allow[1],"*.sms")
	;_ArrayDisplay($1st_level_sms_list)
		for $x=1 to $1st_level_sms_list[0]
			$current_time = _DateDiff( 's',"1970/01/01 00:00:00",_NowCalc())
			$current_text =  FileReadLine ($allow[1]&"\" &$1st_level_sms_list[$x] ,1)
				$current_text =  StringTrimLeft ( $current_text, StringInStr( $current_text ,"\",0,-1)  )
			
			$current_name = FileReadLine ($allow[1]&"\" &$1st_level_sms_list[$x] ,2)
				$current_name =  StringTrimLeft ( $current_name, StringInStr( $current_name ,"\",0,-1)  )
			;MsgBox(0,"read file",$allow[1]&"\" &$1st_level_sms_list[$x] )
			;MsgBox(0,"new request in " & $x, $current_time &"!" & $current_text & "!" &  $current_name & "!" &$1st_level_sms_list[$x] )
			;$2nd_level_sms_list=_FileListToArray($allow[1]&"\"& StringTrimRight($1st_level_sms_list[$x],4))
			;_ArrayDisplay($2nd_level_sms_list)
			if FileExists ( $allow[1]& "\" &StringTrimRight($1st_level_sms_list[$x] ,4) &"\" & $current_text ) then $current_text_status=1

			if FileExists ( $allow[1]& "\" &StringTrimRight($1st_level_sms_list[$x] ,4) &"\" & $current_name ) then $current_name_status=1

			if $current_name_status=1 and $current_text_status=1 then 
				$current_time=_DateAdd ( "s", 120, $current_time )
			_ArrayAdd($Final_level_sms_list, $current_time &"!" & $current_text & "!" &  $current_name & "!" &$1st_level_sms_list[$x] )
			;$Final_level_sms_list[0]= $Final_level_sms_list[0]+1
			FileMove($allow[1]&"\" &$1st_level_sms_list[$x] ,$allow[1]& "\" &StringTrimRight($1st_level_sms_list[$x] ,4))
			if  IsArray($Final_level_sms_list)=1 and $Final_level_sms_list[0]="" then  _ArrayDelete($Final_level_sms_list,0)
			$load = FileOpen( @ScriptDir&"\load.txt", 2) ; 1 = append
			_FileWriteFromArray($load, $Final_level_sms_list, 0)
			FileClose($load)
			
			EndIf
		Next
		;_ArrayDisplay($Final_level_sms_list)
	EndIf
	if  IsArray($Final_level_sms_list)=1 and $Final_level_sms_list[0]="" then  _ArrayDelete($Final_level_sms_list,0)
	if  IsArray($Final_level_sms_list)=1  then _ArrayDisplay($Final_level_sms_list)
EndFunc





Func _open_sms_text_name( $SMS_text_file , $name_list )

;;
;; Now is to open SMS Text file and name list
;; And show them is a msg box
;;

;$name_list=$name_list
;$name_list = "D:\AUTO\script\AE\2nd_1500.csv"
Dim $name_list_array
Dim $name_colume
Dim $mobile_colume

Dim $Show_name_phone = ""
Dim $button_return = 0

If FileExists($name_list) Then
	;$file=FileOpen(@ScriptDir&"\"&$name_list)
	$name_list_array = _file2Array($name_list, 4, ",")
	For $x = 0 To 3

		If StringInStr($name_list_array[0][$x], "學生姓名") Then 
			$name_colume = $x
		Else 
			if StringInStr($name_list_array[0][$x], "姓名") Then $name_colume = $x
		EndIf
		If StringInStr($name_list_array[0][$x], "手機") Then 
			$mobile_colume = $x
		Else
			If StringInStr($name_list_array[0][$x], "行動電話") Then $mobile_colume = $x
		EndIf
	Next
	;MsgBox(0,"name and mobile", $name_colume & "  " & $mobile_colume)
	;_ArrayDisplay($name_list_array)
	;MsgBox (0,"This is mobile table ", UBound($name_list_array,1) & @CRLF & " Record in total")
	
	For $y = 1 To UBound($name_list_array) - 1
		Local $mobile_phone_no = $name_list_array[$y][$mobile_colume]
		;if StringLeft ($mobile_phone_no,1)<>0 then $mobile_phone_no="0"&$mobile_phone_no
		;if StringLeft ($mobile_phone_no,2)<>09
		;MsgBox (0,"Array index : "& $y  , $name_list_array[$y][$name_colume] &" <> "& $name_list_array[$y][$mobile_colume])
		$Show_name_phone = $Show_name_phone & $name_list_array[$y][$name_colume] & "  :  " & $name_list_array[$y][$mobile_colume] & @CRLF
	Next
	$button_return = MsgBox(1, "Show Name and Phone for Check:", "目前是從這個  " & $name_list & @CRLF & "檔案中取得人名與電話  :  " & @CRLF & @CRLF & $Show_name_phone)
	If $button_return = 2 Then
		MsgBox(0, "請再檢查", "請重新執行程式")
		Exit
	EndIf
	
	_ArrayDisplay($name_list_array)
Else
	_FileWriteLog(@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & "_" & $year & $month & $day & ".log", $name_list & " is not at " & @ScriptDir)
	
EndIf



If FileExists($SMS_text_file) Then
	Dim $message = ""
	Dim $a_SMS_text_file
	If Not _FileReadToArray($SMS_text_file, $a_SMS_text_file) Then
		MsgBox(4096, "Error", " Error reading log to Array     error:" & @error)
		Exit
	EndIf
	For $x = 1 To $a_SMS_text_file[0]
		$message = $message & $a_SMS_text_file[$x]
		;Msgbox(0,'Record:' & $x, $$a_SMS_text_file[$x])
	Next
	If StringLen($message) > 63 Then
		$button_return = MsgBox(1, "目前簡訊文字長度:" & @CRLF & StringLen($message), "目前是從 " & $SMS_text_file & "  這個檔案中取得簡訊: " & @CRLF & @CRLF & "1. 簡訊原文:" & @CRLF & @CRLF & $message & @CRLF & @CRLF & @CRLF & @CRLF & "2. 簡訊發出會被截斷成為:" & @CRLF & @CRLF & StringLeft($message, 70))
	Else
		$button_return = MsgBox(1, "目前簡訊文字長度:" & @CRLF & StringLen($message), "目前是從 " & $SMS_text_file & "  這個檔案中取得簡訊: " & @CRLF & @CRLF & $message & @CRLF)
	EndIf
	If $button_return = 2 Then
		MsgBox(0, "", "請再檢查簡訊內容")
		Run("notepad.exe " & $SMS_text_file)
		Exit
	EndIf
	If $button_return = 1 Then $message = StringLeft($message, 63)
	MsgBox(0,"Message", $message)
EndIf

;_ArrayDisplay($a_SMS_text_file)

EndFunc






;; Two dimension array
Func _file2Array($PathnFile, $aColume, $delimiters)

	
	Local $aRecords
	If Not _FileReadToArray($PathnFile, $aRecords) Then
		MsgBox(4096, "Error", " Error reading file '" & $PathnFile & "' to Array   error:" & @error)
		Exit
	EndIf
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






Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Username = "", $s_Password = "", $IPPort = 25, $ssl = 0)
	$objEmail = ObjCreate("CDO.Message")
	$objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
	$objEmail.To = $s_ToAddress
	Local $i_Error = 0
	Local $i_Error_desciption = ""
	If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
	If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress
	$objEmail.Subject = $s_Subject
	If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
		;$objEmail.HTMLBodyPart.Charset="utf8";
		;$objEmail.BodyPart.Charset="utf-8";
		$objEmail.HTMLBody = $as_Body
	Else
		;$objEmail.bodyPart.Charset="utf8";
		$objEmail.Textbody = $as_Body & @CRLF
	EndIf
	If $s_AttachFiles <> "" Then
		Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
		For $x = 1 To $S_Files2Attach[0]
			$S_Files2Attach[$x] = _PathFull($S_Files2Attach[$x])
			If FileExists($S_Files2Attach[$x]) Then
				$objEmail.AddAttachment($S_Files2Attach[$x])
			Else
				$i_Error_desciption = $i_Error_desciption & @LF & 'File not found to attach: ' & $S_Files2Attach[$x]
				SetError(1)
				Return 0
			EndIf
		Next
	EndIf
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort
	;Authenticated SMTP
	If $s_Username <> "" Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
	EndIf
	If $ssl Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
	EndIf
	;Update settings
	$objEmail.BodyPart.Charset = "utf-8";
	$objEmail.Configuration.Fields.Update
	; Sent the Message
	$objEmail.Send
	If @error Then
		SetError(2)
		Return $oMyRet[1]
	EndIf
EndFunc   ;==>_INetSmtpMailCom
;
;
; Com Error Handler
Func MyErrFunc()
	$HexNumber = Hex($oMyError.number, 8)
	$oMyRet[0] = $HexNumber
	$oMyRet[1] = StringStripWS($oMyError.description, 3)
	ConsoleWrite("### COM Error !  Number: " & $HexNumber & "   ScriptLine: " & $oMyError.scriptline & "   Description:" & $oMyRet[1] & @LF)
	SetError(1); something to check for when this function returns
	Return
EndFunc   ;==>MyErrFunc



#include <array.au3>
#include <File.au3>
#include <Date.au3>
#include <CompInfo.au3>

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GUIConstants.au3>

;; �o�O���F�o�e²�T���{���A�D�n�O���F���v��p�ӧ諸
;; 1. �n���w�]����r�ɮסA�o�e��H�ɮ�
;; 2. �n���w���o�X����O
;; 3. �n���o�e��έp����O
;; 4. �i�H�ܤƬ� android ����o�e���i��
;; 5. V �@�}�l�|�ݤ@�ӱK�X
;;

Global $SMS_text_file ; =@ScriptDir&"\SMS_text.txt"
Global $name_list ; = @ScriptDir& "\SMS_name_list.csv"
Global $oMyRet[2]

Dim $sec = @SEC
Dim $min = @MIN
Dim $hour = @HOUR
Dim $day = @MDAY
Dim $month = @MON
Dim $year = @YEAR
Global $test_mode

Global $magic_word = "���v��p�Ѥ���M�εo²�T�K�X"
Global $astronomy = ".astronomy.txt"
Dim $os_partial
Global $version
;$test_mode=_TEST_MODE() ; return 1 means  Test mode.

;MsgBox(0,"on info",$os_partial)
;MsgBox(0,"",@UserProfileDir)
Dim $aData = InetRead("http://ivan:9ps5678@202.133.232.82:8080/upload/astronomy.htm")
Dim $aBytesRead = @extended
;MsgBox(4096, "", "Bytes read: " & $aBytesRead & @CRLF & @CRLF & StringLeft( BinaryToString($aData),3) & @CRLF & StringTrimLeft( BinaryToString($aData),4) )
$version = StringLeft(BinaryToString($aData), 3)

If $aBytesRead = 0 Or $version = "000" Then
	MsgBox(0, "���~", "�o�ӵ{���v�g���ĤF�A" & @CRLF & "�Э��s�U���C")
	Exit
EndIf


_SelectFileGUI()

;If Not FileExists(@ScriptDir & "\" & $astronomy) Then
If Not FileExists(@UserProfileDir & "\" & $astronomy) Then
	$os_partial = _get_os_partial()
	;Local $sData = InetRead("http://ivan:9ps5678@202.133.232.82:8080/upload/astronomy.htm") ;http://202.133.232.82:8080/upload/
	;Local $nBytesRead = @extended
	MsgBox(4096, "", "Bytes read: " & $nBytesRead & @CRLF & @CRLF & BinaryToString($sData) &@CRLF &StringLeft( BinaryToString($sData),4) & $os_partial )

	If $aBytesRead > 0 Then
		;dim $magicfile_name=BinaryToString($sData)&".txt"
		;Dim $magicfile = FileOpen(@ScriptDir & "\" & $astronomy, 10)
		Dim $magicfile = FileOpen(@UserProfileDir & "\" & $astronomy, 10)
		FileWriteLine($magicfile, StringLeft(BinaryToString($aData), 4) & $os_partial & @CRLF)
		FileWriteLine($magicfile, StringTrimLeft(BinaryToString($aData), 4))
		FileClose($magicfile)
		;$magic_word=BinaryToString($sData)

	EndIf
EndIf

If FileExists(@ScriptDir & "\" & $astronomy) Then
	Local $pass
	Local $line1 = FileReadLine(@ScriptDir & "\" & $astronomy, 1)
	If $version <> StringLeft($line1, 3) Then
		;FileMove(@ScriptDir&"\mail_sms.exe")
		InetGet("http://ivan:9ps5678@202.133.232.82:8080/upload/mail_sms.exe", @ScriptDir & "\mail_sms_1.exe", 1)
		MsgBox(0, "ĵ�i", "�o�ӵ{���v�g�L���F�A" & @CRLF & "�о��t���s�U���C")
	EndIf
	$os_partial = _get_os_partial()
	If StringTrimLeft($line1, 4) <> $os_partial Then
		FileDelete(@ScriptDir & "\" & $astronomy)
		MsgBox(0, "Restart the program", "�Э��}�o�ӵ{��", 10)
		Exit
	EndIf

	Local $line2 = FileReadLine(@ScriptDir & "\" & $astronomy, 2)
	;MsgBox(0,"stringinstring of line2",StringInStr ( $magic_word,  $line2 ))
	If StringInStr($magic_word, $line2) = 0 Then
		$input_pass = InputBox("�o�e²�T�ҨϥΪ��K�X", "�п�J")

		If $magic_word <> $input_pass Then
			FileDelete(@ScriptDir & "\" & $astronomy)
			MsgBox(0, "���~", "�K�X���~")
			Exit
		EndIf
	EndIf

EndIf

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

		If StringInStr($name_list_array[0][$x], "�ǥͩm�W") Then $name_colume = $x
		If StringInStr($name_list_array[0][$x], "�m�W") Then $name_colume = $x
		If StringInStr($name_list_array[0][$x], "���") Then $mobile_colume = $x
		If StringInStr($name_list_array[0][$x], "��ʹq��") Then $mobile_colume = $x
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
	$button_return = MsgBox(1, "Show Name and Phone for Check:", "�ثe�O�q�o��  " & $name_list & @CRLF & "�ɮפ����o�H�W�P�q��  :  " & @CRLF & @CRLF & $Show_name_phone)
	If $button_return = 2 Then
		MsgBox(0, "�ЦA�ˬd", "�Э��s����{��")
		Exit
	EndIf

	;_ArrayDisplay($name_list_array)
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
		$button_return = MsgBox(1, "�ثe²�T��r����:" & @CRLF & StringLen($message), "�ثe�O�q " & $SMS_text_file & "  �o���ɮפ����o²�T: " & @CRLF & @CRLF & "1. ²�T���:" & @CRLF & @CRLF & $message & @CRLF & @CRLF & @CRLF & @CRLF & "2. ²�T�o�X�|�Q�I�_����:" & @CRLF & @CRLF & StringLeft($message, 70))
	Else
		$button_return = MsgBox(1, "�ثe²�T��r����:" & @CRLF & StringLen($message), "�ثe�O�q " & $SMS_text_file & "  �o���ɮפ����o²�T: " & @CRLF & @CRLF & $message & @CRLF)
	EndIf
	If $button_return = 2 Then
		MsgBox(0, "", "�ЦA�ˬd²�T���e")
		Run("notepad.exe " & $SMS_text_file)
		Exit
	EndIf
	If $button_return = 1 Then $message = StringLeft($message, 63)

EndIf

;_ArrayDisplay($a_SMS_text_file)








MsgBox(0, "", "It is correct now. Process Send mail Now")
;Exit

; This is send gmail  function
;
;##################################
; Variables
;##################################
;
$s_SmtpServer = "onlinebooking.com.tw" ;"maild.digitshuttle.com"              ; address for the smtp-server to use - REQUIRED
$s_FromName = "ae_backup@onlinebooking.com.tw" ;"bryant@dynalab.com.tw"                      ; name from who the email was sent
$s_FromAddress = "ae_backup@onlinebooking.com.tw";"bryant@dynalab.com.tw" ;  address from where the mail should come
$s_ToAddress = "sms@onlinebooking.com.tw" ; destination address of the email - REQUIRED
$s_Subject = "0928837823"
$as_Body = "�D��" ; the messagebody from the mail - can be left blank but then you get a blank mail
$s_AttachFiles = "" ; the file you want to attach- leave blank if not needed
$s_CcAddress = "" ; address for cc - leave blank if not needed
$s_BccAddress = "" ; address for bcc - leave blank if not needed
;$s_Username = _Base64Encode("ae_direct_fly")                    ; username for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
;$s_Password = _Base64Encode("pkpkpk")                  ; password for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
$s_Password = "1qazxsw2"
$s_Username = "ae_backup"
$s_IPPort = 25 ; port used for sending the mail
$s_ssl = 1 ; Always use 1              ; enables/disables secure socket layer sending - put to 1 if using httpS
;$IPPort=465                            ; GMAIL port used for sending the mail
;$ssl=1                                 ; GMAILenables/disables secure socket layer sending - put to 1 if using httpS
;
;
;$s_SmtpServer = "smtp.gmail.com" ;"maild.digitshuttle.com"              ; address for the smtp-server to use - REQUIRED
;$s_FromName = "changtun" ;"bryant@dynalab.com.tw"                      ; name from who the email was sent
;$s_FromAddress = "changtun@gmail.com";"bryant@dynalab.com.tw" ;  address from where the mail should come
;$s_ToAddress = "sms@onlinebooking.com.tw" ; destination address of the email - REQUIRED
;$s_Subject = "0928837823"
;$as_Body = "�D��" ; the messagebody from the mail - can be left blank but then you get a blank mail
;$s_AttachFiles = "" ; the file you want to attach- leave blank if not needed
;$s_CcAddress = "" ; address for cc - leave blank if not needed
;$s_BccAddress = "" ; address for bcc - leave blank if not needed
;;$s_Username = _Base64Encode("ae_direct_fly")                    ; username for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
;;$s_Password = _Base64Encode("pkpkpk")                  ; password for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
;$s_Password = "9ps567*9"
;$s_Username = "changtun@gmail.com"
;$s_IPPort = 465 ; port used for sending the mail
;$s_ssl = 1 ; Always use 1              ; enables/disables secure socket layer sending - put to 1 if using httpS
;;$IPPort=465                            ; GMAIL port used for sending the mail
;;$ssl=1


;
$m_SmtpServer = "smtp.gmail.com" ; address for the smtp-server to use - REQUIRED
$m_FromName = "DSC" ; name from who the email was sent
$m_FromAddress = "bryant@digtishuttle.com" ;  address from where the mail should come

$m_ToAddress = "bryant@dynalab.com.tw" ; destination address of the email - REQUIRED
$m_Subject = "SMS-mail-status" ; subject from the email - can be anything you want it to be
$m_as_Body = "SMS-mail-status" ; the messagebody from the mail - can be left blank but then you get a blank mail
$m_AttachFiles = @ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & "_" & $year & $month & $day & ".log" ; the file you want to attach- leave blank if not needed sample :"d:\ibm240KB.jpg"
$m_CcAddress = "" ; address for cc - leave blank if not needed
$m_BccAddress = "" ; address for bcc - leave blank if not needed
$m_Username = "changtun@gmail.com" ; username for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
$m_Password = "9ps567*9" ; password for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
;$IPPort = 25                              ; port used for sending the mail
;$ssl = 0                                  ; enables/disables secure socket layer sending - put to 1 if using httpS
$IPPort = 465 ; GMAIL port used for sending the mail
$ssl = 1 ; GMAILenables/disables secure socket layer sending - put to 1 if using httpS
;
;


Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
;$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $IPPort, $ssl)
;
;;
;;
;dim $1st= "MyOnlineBookingST1@gmail.com"
;dim $2nd= "myonlinebookingst2@gmail.com"
;dim $3rd= "myonlinebookingst3@gmail.com"
;Dim $mymailbody = "�ϥΪ̥������ ���U/�n�J�A�n�J��~�i�H�o��Post�A���M�u���s���C�u���ۤv��Post�~��i��ק�P�R���C"

For $r = 1 To (UBound($name_list_array, 1) - 1)

	Dim $day = @MDAY
	Dim $month = @MON
	Dim $year = @YEAR
	;$m_AttachFiles = @ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"
	$as_Body = $name_list_array[$r][$name_colume] & "�z�n: " & $message

	;MsgBox (0,"This is mobile",$name_list_array[$r][2] & @CRLF & " This is going to send mail. Stop if you want")
	;$s_ToAddress = "sms@onlinebooking.com.tw"
	If StringInStr($name_list_array[$r][$mobile_colume], "_") Then
		$s_Subject = StringReplace($name_list_array[$r][$mobile_colume], "_", "")
		;$as_Body=StringReplace($as_Body,"[user_email]","ae@delta.com.tw") ; For test only.
	Else
		$s_Subject = $name_list_array[$r][$mobile_colume]
	EndIf
	;MsgBox (0,"This is mobile",$s_Subject & @CRLF & " This is going to send mail. Stop if you want")
	;MsgBox (0,"mail parameter", $s_SmtpServer&" / "& $s_FromName&" / "& $s_FromAddress&" / "& $s_ToAddress&" / "& $s_Subject&" / "& $as_Body&" / "& $s_AttachFiles&" / "& $s_CcAddress&" / "& $s_BccAddress&" / "& $s_Username&" / "& $s_Password&" / "& $s_IPPort&" / "& $s_ssl)
	;$m_ToAddress =	$name_list_array[$r][0] ;
	;$s_ToAddress = $name_list_array[$r][0] ;Correct mail to

	;$as_Body= "Updated at "&$year&$month&$day& @CRLF &$as_Body ; Correct sentence

	;$as_Body= "Updated at "&$year&$month&$day& @CRLF & $name_list_array[$r][0] &@CRLF &$as_Body ; for test only. To locate email address in mail body

	;if mod($r, 3)=1 	then
	;$s_Username =$1st
	;EndIf
	;
	;if mod($r, 3)=2 	Then
	;$s_Username =$2nd
	;EndIf
	;
	;if mod($r, 3)=0		then
	;$s_Username = $3rd
	;EndIf
	;$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $m_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $s_IPPort, $s_ssl)
	;### This is for Corrrect SMS
	;$s_Subject="0919585516"
	MsgBox(0, "mail parameter", $s_SmtpServer & " / " & $s_FromName & " / " & $s_FromAddress & " / " & $s_ToAddress & " / " & $s_Subject & " / " & $as_Body & " / " & $s_AttachFiles & " / " & $s_CcAddress & " / " & $s_BccAddress & " / " & $s_Username & " / " & $s_Password & " / " & $s_IPPort & " / " & $s_ssl)
	$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $s_IPPort, $s_ssl)
	;### This is mail Test
	;$s_ToAddress="bryant@dynalab.com.tw"
	;MsgBox (0,"mail parameter", $s_SmtpServer&" / "& $s_FromName&" / "& $s_FromAddress&" / "& $s_ToAddress&" / "& $s_Subject&" / "& $as_Body&" / "& $s_AttachFiles&" / "& $s_CcAddress&" / "& $s_BccAddress&" / "& $s_Username&" / "& $s_Password&" / "& $s_IPPort&" / "& $s_ssl)
	;$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $s_IPPort, $s_ssl)
	;$rc = _INetSmtpMailCom($m_SmtpServer, $m_FromName, $m_FromAddress, $m_ToAddress, $m_Subject, $as_Body, $m_AttachFiles, $m_CcAddress, $m_BccAddress, $m_Username, $m_Password, $IPPort, $ssl)
	_FileWriteLog(@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & "_" & $year & $month & $day & ".log", " Mail Send to " & $s_Subject & "  " & $s_ToAddress)
	If $r > 0 And Mod($r, 2) = 0 Then
		Sleep(20000)
	EndIf
	If $r = (UBound($name_list_array, 1) - 1) Then
		Dim $day = @MDAY
		Dim $month = @MON
		Dim $year = @YEAR
		Local $m_AttachFiles = @ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & "_" & $year & $month & $day & ".log"
		$rc = _INetSmtpMailCom($m_SmtpServer, $m_FromName, $m_FromAddress, $m_ToAddress, $m_Subject, $m_as_Body, $m_AttachFiles, $m_CcAddress, $m_BccAddress, $m_Username, $m_Password, $IPPort, $ssl)
		;$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $m_ToAddress, $m_Subject, $m_as_Body, $m_AttachFiles, $m_CcAddress, $s_BccAddress, $s_Username, $s_Password, $s_IPPort, $s_ssl)

		;Sleep(1000 * 60 * 5)
	EndIf


Next




Exit
;##################################
; Send gmail Script Sample
;##################################
Global $oMyRet[2]
;Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
;$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $IPPort, $ssl)
;
;If @error Then
;    MsgBox(0, "Error sending message", "Error code:" & @error & "  Rc:" & $rc)
;EndIf
;    MsgBox(0, "sending message", "Error code:" & @error & "  Rc:" & $rc)
;
;


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



Func _get_os_partial()
	Dim $OSs
	_ComputerGetOSs($OSs)
	If @error Then
		$error = @error
		$extended = @extended
		Switch $extended
			Case 1
				_ErrorMsg($ERR_NO_INFO)
			Case 2
				_ErrorMsg($ERR_NOT_OBJ)
		EndSwitch
	EndIf
	Dim $OSs_partial
	For $i = 1 To $OSs[0][0] Step 1
		;MsgBox(0, "Test _ComputerGetOSs", $i & _
		;		"CS Name: " & $OSs[$i][10] & @CRLF & _
		;		"Install Date: " & $OSs[$i][23] & @CRLF & _
		;		"OS Type: " & $OSs[$i][37] & @CRLF & _
		;		"Registered User: " & $OSs[$i][45] & @CRLF & _
		;		"Serial Number: " & $OSs[$i][46] & @CRLF & _
		;		"Version: " & $OSs[$i][58] & @CRLF )
	Next
	$i = 1
	$OSs_partial = $OSs[$i][10] & "$*$" & $OSs[$i][45] & "$*$" & $OSs[$i][46] & "$*$" & $OSs[$i][23]
	;MsgBox(0,"Info", $OSs_partial)
	Return (StringStripWS($OSs_partial, 8))
EndFunc   ;==>_get_os_partial
Func _ErrorMsg($message, $time = 0)
	MsgBox(48 + 262144, "Error!", $message, $time)
EndFunc   ;==>_ErrorMsg
;$user_id=$fundinfo
;$os_partial=_get_os_partial()
;MsgBox(0,"User ID",$user_id)


Func _TEST_MODE()

	If FileExists(@ScriptDir & "\TESTMODE.txt") Then
		$mode = FileReadLine(@ScriptDir & "\TESTMODE.txt", 1)
		If $mode = 1 Then
			MsgBox(0, "Test mode", "���ռҦ�" & @CRLF & "�u�|�H�e�� Service Account ", 5)

		Else
			;MsgBox(0,"Process mode", " ���K������Ʒ|��J��Ʈw ",10)
			;$ans=InputBox("Process mode","���K������Ʒ|��J��Ʈw "&@CRLF& "��J N �i�H���}")

			$mode = 0
			MsgBox(0, "Test mode", "�����Ҧ�" & @CRLF & "Order �ƥ��ɮ׷|�H�e����ݥD�H���H�c ", 5)
			;if $ans="n" or $ans="N" or @error=1 then exit
		EndIf

	Else
		;MsgBox(0,"Process mode", " ���K������Ʒ|��J��Ʈw ",10)
		;$ans=InputBox("Process mode","���K������Ʒ|��J��Ʈw "&@CRLF& "��J N �i�H���}")

		$mode = 0
		MsgBox(0, "Test mode", "�����Ҧ�" & @CRLF & "Order �ƥ��ɮ׷|�H�e����ݥD�H���H�c ", 5)
		;if $ans="n" or $ans="N" or @error=1 then exit

	EndIf

	Return $mode
EndFunc   ;==>_TEST_MODE




Func _SelectFileGUI()

	Local $file_txt, $file_csv, $btn, $msg, $btn_n, $aEnc_info, $rc

	GUICreate("��J�ɮ�", 320, 210, @DesktopWidth / 3 - 320, @DesktopHeight / 3 - 240, -1, 0x00000018); WS_EX_ACCEPTFILES
	GUICtrlCreateLabel("1.���²�T���e�ɮר�o�ӮءA�w�]�� SMS_text.txt", 10, 10, 300, 40)
	$file_txt = GUICtrlCreateInput("", 10, 25, 300, 40)
	GUICtrlSetState(-1, $GUI_DROPACCEPTED)
	GUICtrlCreateLabel("2.���²�T���e�ɮר�o�ӮءA�w�]�� SMS_name_list.csv", 10, 85, 300, 40)
	$file_csv = GUICtrlCreateInput("", 10, 100, 300, 40)
	GUICtrlSetState(-1, $GUI_DROPACCEPTED)
	;GUICtrlCreateInput("", 10, 35, 300, 20) 	; will not accept drag&drop files
	$btn = GUICtrlCreateButton("OK", 90, 180, 60, 20, 0x0001) ; Default button
	$btn_n = GUICtrlCreateButton("Exit", 160, 180, 60, 20)
	GUISetState()

	;$msg = 0
	While $msg <> $GUI_EVENT_CLOSE
		$msg = GUIGetMsg()
		Select
			Case $msg = $btn
				If Not (GUICtrlRead($file_txt) = "" And GUICtrlRead($file_csv) = "") Then
					;MsgBox(4096, "drag drop file", GUICtrlRead($file_txt) & "  " & GUICtrlRead($file_csv))
					$SMS_text_file = GUICtrlRead($file_txt)
					$name_list = GUICtrlRead($file_csv)
				Else
					$SMS_text_file = @ScriptDir & "\SMS_text.txt"
					$name_list = @ScriptDir & "\SMS_name_list.csv"
				EndIf
				ExitLoop
			Case $msg = $btn_n
				Exit
		EndSelect
	WEnd


EndFunc   ;==>_SelectFileGUI
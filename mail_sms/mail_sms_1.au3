#include <array.au3>
#Include <File.au3>		
#include <Date.au3>

;;
;;

dim $sec=@SEC
dim $min=@MIN
Dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR

$txt_file="2nd_1500.csv"
dim $extramail_array
if FileExists(@ScriptDir&"\"&$txt_file) then 
	;$file=FileOpen(@ScriptDir&"\"&$txt_file)
	$extramail_array=_file2Array(@ScriptDir&"\"&$txt_file ,3,",")
	;_ArrayDisplay($extramail_array)
	;MsgBox (0,"This is mobile table ", UBound($extramail_array,1) & @CRLF & " Record in total")
else 
	_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",$txt_file& " is not at " & @ScriptDir)
	
EndIf





; This is send gmail  function
;
;##################################
; Variables
;##################################
;
$s_SmtpServer = "onlinebooking.com.tw" ;"maild.digitshuttle.com"              ; address for the smtp-server to use - REQUIRED
$s_FromName = "ae_backup@onlinebooking.com.tw" ;"bryant@dynalab.com.tw"                      ; name from who the email was sent
$s_FromAddress = "ae_backup@onlinebooking.com.tw";"bryant@dynalab.com.tw" ;  address from where the mail should come
$s_ToAddress = "sms@onlinebooking.com.tw"   ; destination address of the email - REQUIRED
$s_Subject = "0928837823"   
$as_Body = "需要性福時偶爾力不從心?想使用藍色小丸子怕傷身又要先看醫生?合法通過SGS檢驗的天然產品,讓你硬景而升http://4fun.tw/VUQJ"                             ; the messagebody from the mail - can be left blank but then you get a blank mail
$s_AttachFiles = ""                       ; the file you want to attach- leave blank if not needed
$s_CcAddress = ""					      ; address for cc - leave blank if not needed
$s_BccAddress = ""     					  ; address for bcc - leave blank if not needed
;$s_Username = _Base64Encode("ae_direct_fly")                    ; username for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
;$s_Password = _Base64Encode("pkpkpk")                  ; password for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
$s_Password = "1234"
$s_Username = "ae_backup"
$s_IPPort = 25                              ; port used for sending the mail
$s_ssl = 1     ; Always use 1              ; enables/disables secure socket layer sending - put to 1 if using httpS
;$IPPort=465                            ; GMAIL port used for sending the mail
;$ssl=1                                 ; GMAILenables/disables secure socket layer sending - put to 1 if using httpS
;
;
;
$m_SmtpServer = "smtp.gmail.com"              ; address for the smtp-server to use - REQUIRED
$m_FromName = "DSC"                      ; name from who the email was sent
$m_FromAddress = "bryant@digtishuttle.com" ;  address from where the mail should come

$m_ToAddress = "bryant@dynalab.com.tw"   ; destination address of the email - REQUIRED
$m_Subject =  "SMS-mail-status"                   ; subject from the email - can be anything you want it to be
$m_as_Body =  "SMS-mail-status"             ; the messagebody from the mail - can be left blank but then you get a blank mail
$m_AttachFiles = @ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"             ; the file you want to attach- leave blank if not needed sample :"d:\ibm240KB.jpg"   
$m_CcAddress = ""       ; address for cc - leave blank if not needed
$m_BccAddress = ""     ; address for bcc - leave blank if not needed
$m_Username = "changtun@gmail.com"                    ; username for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
$m_Password = "9ps567*9"                  ; password for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
;$IPPort = 25                              ; port used for sending the mail
;$ssl = 0                                  ; enables/disables secure socket layer sending - put to 1 if using httpS
$IPPort=465                            ; GMAIL port used for sending the mail
$ssl=1                                 ; GMAILenables/disables secure socket layer sending - put to 1 if using httpS
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
dim $mymailbody="使用者必須能夠 註冊/登入，登入後才可以發表Post，不然只能瀏覽。只有自己的Post才能進行修改與刪除。"

for $r=0 to  1 ;(UBound($extramail_array,1)-1)
	
		Dim $day=@MDAY
		Dim $month=@MON
		DIM $year=@YEAR
		;$m_AttachFiles = @ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"
		;$as_Body=$mymailbody
		
		;MsgBox (0,"This is mobile",$extramail_array[$r][2] & @CRLF & " This is going to send mail. Stop if you want")
		$s_ToAddress = "sms@onlinebooking.com.tw" 
		if StringInStr($extramail_array[$r][2],"_") then 
		$s_Subject=StringReplace($extramail_array[$r][2],"_","")
			;$as_Body=StringReplace($as_Body,"[user_email]","ae@delta.com.tw") ; For test only.
			
		EndIf
		;MsgBox (0,"This is mobile",$s_Subject & @CRLF & " This is going to send mail. Stop if you want")
		;MsgBox (0,"mail parameter", $s_SmtpServer&" / "& $s_FromName&" / "& $s_FromAddress&" / "& $s_ToAddress&" / "& $s_Subject&" / "& $as_Body&" / "& $s_AttachFiles&" / "& $s_CcAddress&" / "& $s_BccAddress&" / "& $s_Username&" / "& $s_Password&" / "& $s_IPPort&" / "& $s_ssl)
		;$m_ToAddress =	$extramail_array[$r][0] ; 
		;$s_ToAddress = $extramail_array[$r][0] ;Correct mail to 
		
		;$as_Body= "Updated at "&$year&$month&$day& @CRLF &$as_Body ; Correct sentence
		
		;$as_Body= "Updated at "&$year&$month&$day& @CRLF & $extramail_array[$r][0] &@CRLF &$as_Body ; for test only. To locate email address in mail body
		
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
		;MsgBox (0,"mail parameter", $s_SmtpServer&" / "& $s_FromName&" / "& $s_FromAddress&" / "& $s_ToAddress&" / "& $s_Subject&" / "& $as_Body&" / "& $s_AttachFiles&" / "& $s_CcAddress&" / "& $s_BccAddress&" / "& $s_Username&" / "& $s_Password&" / "& $s_IPPort&" / "& $s_ssl)		
		$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $s_IPPort, $s_ssl)
	;### This is mail Test
		;$s_ToAddress="bryant@dynalab.com.tw"
		;MsgBox (0,"mail parameter", $s_SmtpServer&" / "& $s_FromName&" / "& $s_FromAddress&" / "& $s_ToAddress&" / "& $s_Subject&" / "& $as_Body&" / "& $s_AttachFiles&" / "& $s_CcAddress&" / "& $s_BccAddress&" / "& $s_Username&" / "& $s_Password&" / "& $s_IPPort&" / "& $s_ssl)		
		;$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $s_IPPort, $s_ssl)
			;$rc = _INetSmtpMailCom($m_SmtpServer, $m_FromName, $m_FromAddress, $m_ToAddress, $m_Subject, $as_Body, $m_AttachFiles, $m_CcAddress, $m_BccAddress, $m_Username, $m_Password, $IPPort, $ssl)
		_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"," Mail Send to "& $s_Subject& "  "&$s_ToAddress)
		if $r>0 and mod($r, 3)=0		then 
			sleep(5000)
		EndIf
		if $r>0 and mod($r, 500)=0 Then
			Dim $day=@MDAY
			Dim $month=@MON
			DIM $year=@YEAR
			;local $my_attach=@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"
		;$rc = _INetSmtpMailCom($m_SmtpServer, $m_FromName, $m_FromAddress, $m_ToAddress, $m_Subject, $m_as_Body, $m_AttachFiles, $m_CcAddress, $m_BccAddress, $m_Username, $m_Password, $IPPort, $ssl)
		$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $m_ToAddress, $m_Subject, $m_as_Body, $m_AttachFiles, $m_CcAddress, $s_BccAddress, $s_Username, $s_Password, $s_IPPort, $s_ssl)

			sleep(1000*60*5)
		EndIf

	
Next





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


Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Username = "", $s_Password = "",$IPPort=25, $ssl=0)
    $objEmail = ObjCreate("CDO.Message")
    $objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
    $objEmail.To = $s_ToAddress
    Local $i_Error = 0
    Local $i_Error_desciption = ""
    If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
    If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress
    $objEmail.Subject = $s_Subject
    If StringInStr($as_Body,"<") and StringInStr($as_Body,">") Then
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
            $S_Files2Attach[$x] = _PathFull ($S_Files2Attach[$x])
            If FileExists($S_Files2Attach[$x]) Then
                $objEmail.AddAttachment ($S_Files2Attach[$x])
            Else
                $i_Error_desciption = $i_Error_desciption & @lf & 'File not found to attach: ' & $S_Files2Attach[$x]
                SetError(1)
                return 0
            EndIf
        Next
    EndIf
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort
;Authenticated SMTP
    If $s_Username <> "" Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
    EndIf
    If $Ssl Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = false
    EndIf
;Update settings
	$objEmail.BodyPart.Charset="utf-8";
    $objEmail.Configuration.Fields.Update
; Sent the Message
    $objEmail.Send
    if @error then
        SetError(2)
        return $oMyRet[1]
    EndIf
EndFunc ;==>_INetSmtpMailCom
;
;
; Com Error Handler
Func MyErrFunc()
    $HexNumber = Hex($oMyError.number, 8)
    $oMyRet[0] = $HexNumber
    $oMyRet[1] = StringStripWS($oMyError.description,3)
    ConsoleWrite("### COM Error !  Number: " & $HexNumber & "   ScriptLine: " & $oMyError.scriptline & "   Description:" & $oMyRet[1] & @LF)
    SetError(1); something to check for when this function returns
    Return
EndFunc ;==>MyErrFunc



;; Two dimension array
Func _file2Array($PathnFile,$aColume,$delimiters)

	
	local $aRecords
	If Not _FileReadToArray($PathnFile,$aRecords) Then
		MsgBox(4096,"Error", " Error reading file '"&$PathnFile&"' to Array   error:" & @error)
		Exit
	EndIf
		;c
	local $TextToArray[$aRecords[0]][$aColume+1]
	;$TextToArray[0][0]=$aRecords[0]
	local $aRow
	For $y = 1 to $aRecords[0]
		;Msgbox(0,'Record:' & $y, $aRecords[$y])
		 
		$aRow=StringSplit($aRecords[$y],$delimiters)
		;Msgbox(0,'X ,Colume :', $aRow[0])
		For $x=1 to $aRow[0]
			if StringInStr($aRow[$x],",") then 
			
			$aRow[$x]=StringTrimLeft($aRow[$x],1)
			;MsgBox(0, "after", $aRow[$x])
			EndIf
			 $TextToArray[$y-1][$x-1]=$aRow[$x]
		next
	Next
	
	;_ArrayDisplay($TextToArray)
	Return $TextToArray

EndFunc
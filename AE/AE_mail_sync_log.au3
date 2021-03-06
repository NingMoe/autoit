;
;
; Send e-mail to AE staff for Faresheet sync 
; sync log  Faresheet2_to_Faresheet_Log_Page1.html
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
;##################################
; Include
;##################################
#include <file.au3>
#include <Array.au3>

;##################################
; Variables
;##################################
$s_SmtpServer = "smtp.gmail.com"              ; address for the smtp-server to use - REQUIRED
$s_FromName = "bryant_net1"                      ; name from who the email was sent
$s_FromAddress = "bryant@net1.com.tw" ;  address from where the mail should come
;$s_FromAddress = "changtun@gmail.com" ;  address from where the mail should come
$s_ToAddress = "sms@onlinebooking.com.tw"   ; destination address of the email - REQUIRED
$s_Subject = "0928837823"                   ; subject from the email - can be anything you want it to be
$as_Body = "我是新的簡訊"            ; the messagebody from the mail - can be left blank but then you get a blank mail
$s_AttachFiles = ""                        ; the file you want to attach- leave blank if not needed sample :"d:\ibm240KB.jpg"   
$s_CcAddress = ""       ; address for cc - leave blank if not needed
$s_BccAddress = ""     ; address for bcc - leave blank if not needed
$s_Username = "changtun2009@googlemail.com"                    ; username for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
$s_Password = "9ps56789"                  ; password for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
;$IPPort = 25                              ; port used for sending the mail
;$ssl = 0                                  ; enables/disables secure socket layer sending - put to 1 if using httpS
$IPPort=465                            ; GMAIL port used for sending the mail
$ssl=1                                 ; GMAILenables/disables secure socket layer sending - put to 1 if using httpS
;
Global $oMyRet[2]
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
;$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $IPPort, $ssl)


dim $sec=@SEC
dim $min=@MIN
dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR

if FileExists(@ScriptDir&"\Faresheet2_to_Faresheet_Log_Page1.html") then 
	
	$t=FileGetTime(@ScriptDir&"\Faresheet2_to_Faresheet_Log_Page1.html")
	;$yyyymd = $t[0]  & $t[1]  & $t[2] &" " & $t[3] & ":" & $t[4] & ":" & $t[5]
	$yyyymd = $t[0]  & $t[1]  & $t[2] ;&" " & $t[3] & ":" & $t[4] & ":" & $t[5]
	;MsgBox(0,"File time ",$yyyymd)
	$delta=int($year&$month&$day)-($yyyymd)
	;MsgBox(0,"Delta", $delta)
	if $delta=0 then 
						$s_ToAddress ="bryant@net1.com.tw"
						$s_CcAddress ="Pamela.H.Pan@aexp.com" 
						$s_BccAddress ="vivian.v.hsieh@aexp.com"
						$s_Subject = $year&$month&$day&" SyncBack log file"                  ; subject from the email - can be anything you want it to be
						$as_Body = $year&$month&$day&" SyncBack log file"            ; the messagebody from the mail - can be left blank but then you get a blank mail
						$s_AttachFiles = @ScriptDir&"\Faresheet2_to_Faresheet_Log_Page1.html"
						$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $IPPort, $ssl)

		
		;_AE_mail("FareSheet_Sync_log",$year&$month&$day&"_FareSheet_Sync_log","changtun@gmail.com","Pamela.H.Pan@aexp.com",@ScriptDir&"\Faresheet2_to_Faresheet_Log_Page1.html")
	Else
		
						$s_ToAddress ="bryant@net1.com.tw"
						$s_CcAddress ="Pamela.H.Pan@aexp.com" 
						$s_Subject = $year&$month&$day&"警告:FareSheet_Sync檔案有異常"                  ; subject from the email - can be anything you want it to be
						$as_Body = $year&$month&$day&"警告:FareSheet_Sync檔案有異常" 
						$s_BccAddress ="vivian.v.hsieh@aexp.com"
						; the messagebody from the mail - can be left blank but then you get a blank mail
						$s_AttachFiles = @ScriptDir&"\Faresheet2_to_Faresheet_Log_Page1.html" 
						$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $IPPort, $ssl)

		;_AE_mail("警告:FareSheet_Sync檔案有異常","警告:FareSheet_Sync檔案有異常","changtun@gmail.com","Pamela.H.Pan@aexp.com",@ScriptDir&"\Faresheet2_to_Faresheet_Log_Page1.html")
	
	EndIf
	
EndIf

	


Func _AE_mail($subject1,$text1,$sendto1,$bcc1,$attach1)
MsgBox(0,"Mail",$subject1&"_"&$text1,5)
Run (@ScriptDir&"\commail.exe -host=onlinebooking.com.tw  -from=ae_backup@onlinebooking.com.tw -to="&$sendto1&" -bcc="&$bcc1&"  -subject="&$subject1&"_"&$year&$month&$day&" -text="&$text1&" -attach="&$attach1)
;Run (@ScriptDir&"\commail.exe -host=maild.digitshuttle.com  -from=bryant@digitshuttle.com -to="&$sendto1&" -bcc="&$bcc1&"  -subject="&$subject1&"_"&$year&$month&$day&" -text="&$text1&" -attach="&$attach1)

EndFunc




;##################################
; Send gmail Script Sample
;##################################
;Global $oMyRet[2]
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
        $objEmail.HTMLBody = $as_Body
    Else
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
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
    EndIf
    If $Ssl Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    EndIf
;Update settings
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

#Include<file.au3>

;##################################
; Include
;##################################
;#Include<file.au3>
;##################################
; Variables
;##################################
$s_SmtpServer = "smtp.gmail.com"              ; address for the smtp-server to use - REQUIRED
$s_FromName = "changtun"                      ; name from who the email was sent
$s_FromAddress = "changtun@gmail.com" ;  address from where the mail should come
$s_ToAddress =""; "bryant@dynalab.com.tw"   ; destination address of the email - REQUIRED
$s_Subject =""; "我是新的信件_of new mail"                   ; subject from the email - can be anything you want it to be
$as_Body =""; "我是新的信件，可是這些是無家庭聯絡簿"            ; the messagebody from the mail - can be left blank but then you get a blank mail
$s_AttachFiles = "";"D:\AUTO\script\encryption\AES\1252946704.txt.enc"                        ; the file you want to attach- leave blank if not needed sample :"d:\ibm240KB.jpg"
$s_CcAddress = ""       ; address for cc - leave blank if not needed
$s_BccAddress = ""     ; address for bcc - leave blank if not needed
$s_Username = "changtun@gmail.com"                    ; username for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
$s_Password = "9ps567*9"                  ; password for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
;$IPPort = 25                              ; port used for sending the mail
;$ssl = 0                                  ; enables/disables secure socket layer sending - put to 1 if using httpS
 $IPPort=465                            ; GMAIL port used for sending the mail
 $ssl=1                                 ; GMAILenables/disables secure socket layer sending - put to 1 if using httpS

;##################################
; Script
;##################################
Global $oMyRet[2]
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")

func _bryant_sendmail($s_ToAddress_b, $s_Subject_b , $as_Body_b )
	$s_ToAddress =$s_ToAddress_b
	$s_Subject =$s_Subject_b
	$as_Body =$as_Body_b

$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $IPPort, $ssl)

If @error Then
    MsgBox(0, "Error sending message", "Error code:" & @error & "  Rc:" & $rc,5)
EndIf
    MsgBox(0, "sending message", "Error code:" & @error & "  Rc:" & $rc, 5)
EndFunc
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
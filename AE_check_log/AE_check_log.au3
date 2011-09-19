#include <array.au3>
#include <File.au3>
#include <Date.au3>
#include <string.au3> 

Global $test_mode , $mail_body
Global $sec=@SEC
Global $min=@MIN
Global $hour=@HOUR
Global $day=@MDAY
Global $month=@MON
Global $year=@YEAR
Global $WeekOfDay = @WDAY   ; Return 1-7  represent from Sunday , Monday ~ Saturday
Global $today=  $year&$month&$day
dim $date_format= StringTrimLeft($today,2)


$test_mode= _TEST_MODE()

Global $server_address[3]
$server_address[0]=2
if $test_mode=1 then 
	$server_address[1]="r:\"
	$server_address[2]="r:\"
else 	
	$server_address[1]="\\10.112.55.102\amex_data\mylog\"
	$server_address[2]="\\10.112.55.81\d$\amex_data\mylog\"
EndIf

;MsgBox(0,"Date format", $date_format)
; This is for St2 server
dim $st2_log_array[9]
$st2_log_array[0]=8
$st2_log_array[1]="["&$date_format&"]allpnr_find_uc_success.log"
$st2_log_array[2]="["&$date_format&"]allpnr_find_uc_fail.log"
$st2_log_array[3]="["&$date_format&"]direct_fly_success.log"
$st2_log_array[4]="["&$date_format&"]direct_fly_fail.log"
$st2_log_array[5]="["&$date_format&"]tinfo_success.log"
$st2_log_array[6]="["&$date_format&"]tinfo_fail.log"
$st2_log_array[7]="["&$date_format&"]ez_query_success.log"
$st2_log_array[8]="["&$date_format&"]reftkt_success.log"
;$st2_log_array[9]="Sync_log"
;$st2_log_array[10]="Mail_log"


dim $st2_logkeyword_array[9]
$st2_logkeyword_array[0]=8
$st2_logkeyword_array[1]="分送Qmail重要信件給TC完成"
$st2_logkeyword_array[2]="信件錯誤"
$st2_logkeyword_array[3]="與abacusTPE~SHA直航同步完成"
$st2_logkeyword_array[4]="沒有這個艙等"
$st2_logkeyword_array[5]="[tinfo_success]全部票資料完成"
$st2_logkeyword_array[6]="[tinfo_fail]機位異動之確認行程表需求寄發 失敗 pnr:"
$st2_logkeyword_array[7]="蒐集EZ資料日期"
$st2_logkeyword_array[8]="[reftkt_success]全部票資料完成"
;$st2_logkeyword_array[9]=""
;$st2_logkeyword_array[10]=""
;;;;
;;
;;
;; This is for ST1 server
dim $st1_log_array[17]
$st1_log_array[0]=16
$st1_log_array[1]= "["&$date_format&"]eticket_c_turn_fail.log"
$st1_log_array[2]= "["&$date_format&"]eticket_c_turn_success.log"
$st1_log_array[3]= "["&$date_format&"]eticket_e_turn_fail.log"
$st1_log_array[4]= "["&$date_format&"]eticket_e_turn_success.log"
$st1_log_array[5]= "["&$date_format&"]visa_fail.log"
$st1_log_array[6]= "["&$date_format&"]visa_success.log"
$st1_log_array[7]= "["&$date_format&"]confirm_turn_fail.log"
$st1_log_array[8]= "["&$date_format&"]confirm_turn_success.log"
$st1_log_array[9]= "["&$date_format&"]DelSeg_fail.log"
$st1_log_array[10]="["&$date_format&"]DelSeg_success.log"
$st1_log_array[11]="["&$date_format&"]DelSeg_error.log"
$st1_log_array[12]="["&$date_format&"]direct_fly_mail_fail.log"
$st1_log_array[13]="["&$date_format&"]travel_result_download_fail.log"
$st1_log_array[14]="["&$date_format&"]xo_send_irene_success.log"
$st1_log_array[15]="["&$date_format&"]bt_tran_auto_cancel_success.log"

dim $st1_logkeyword_array[17]
$st1_logkeyword_array[0]=16
$st1_logkeyword_array[1]="eticket_c_turn_fail"
$st1_logkeyword_array[2]="eticket_c_turn_success"
$st1_logkeyword_array[3]="eticket_e_turn_fail"
$st1_logkeyword_array[4]="eticket_e_turn_success"
$st1_logkeyword_array[5]="不是顧客不處理"
$st1_logkeyword_array[6]="visa_success]寄發簽證"
$st1_logkeyword_array[7]="來自ae88@onlinebooking.com.tw的confirm_turn要求命令"
$st1_logkeyword_array[8]="轉發確認行程表需求寄發 成功"
$st1_logkeyword_array[9]="DelSeg_fail"
$st1_logkeyword_array[10]="DelSeg_success]pnr"
$st1_logkeyword_array[11]="資料有誤此筆先跳過,下次再處理!"
$st1_logkeyword_array[12]="direct_fly_mail_fail"
$st1_logkeyword_array[13]="travel_result_download_fail]travel_result 採集檔案 程序  完成"
$st1_logkeyword_array[14]="xo_send_irene_success"
$st1_logkeyword_array[15]="bt_tran_auto_cancel_success"

;;
;;
;;
;;

Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
$s_SmtpServer = "smtp.gmail.com"              ; address for the smtp-server to use - REQUIRED
$s_FromName = "changtun"                      ; name from who the email was sent
$s_FromAddress = "changtun@gmail.com" ;  address from where the mail should come
$s_ToAddress = "bryant@dynalab.com.tw"   ; destination address of the email - REQUIRED
$s_Subject = "每日log檢查"                   ; subject from the email - can be anything you want it to be
$as_Body = "每日log檢查"          ; the messagebody from the mail - can be left blank but then you get a blank mail
$s_AttachFiles = ""                        ; the file you want to attach- leave blank if not needed sample :"d:\ibm240KB.jpg"   
$s_CcAddress = ""       ; address for cc - leave blank if not needed
$s_BccAddress = ""     ; address for bcc - leave blank if not needed
$s_Username = "changtun@gmail.com"                    ; username for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
$s_Password = "9ps567*9"                  ; password for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
;$IPPort = 25                              ; port used for sending the mail
;$ssl = 0                                  ; enables/disables secure socket layer sending - put to 1 if using httpS
 $IPPort=465                            ; GMAIL port used for sending the mail
 $ssl=1   
Global $oMyRet[2]
;;

dim $read_return

;MsgBox(0,"0 file is at ", $st2_log_array[1])
;MsgBox(0,"Server", $server_address[1])
;$read_return =  _open_and_look($server_address[1],$st2_log_array[1],$st2_logkeyword_array[1])
;MsgBox(0,'Return to real world', $read_return)

;; St2 [date] 的格式是 110703
for $x=1 to 8
	if FileExists($server_address[1] & $st2_log_array[$x] ) then 
		$read_return =  _open_and_look($server_address[1],$st2_log_array[$x],$st2_logkeyword_array[$x])	
		;MsgBox(0,$x & 'Return to real world', $read_return,10)
		if $read_return=0  then 
		$mail_body=$mail_body &	 " 伺服器: " &$server_address[1] &	 "檔案: " & $st2_log_array[$x] &  "  KeyWord: " & $st2_logkeyword_array[$x]  &	 "  不存在" & @CRLF
			
			
		EndIf
	Elseif not StringInStr($st2_log_array[$x],"fail.log")  then
	$mail_body = $mail_body & @CRLF &	" 伺服器: " &$server_address[1] &	 "   檔案: " & $st2_log_array[$x] &  "   不存在" & @CRLF & @CRLF
	Else
	EndIf
Next


; St1
for $x=1 to 15
	if FileExists($server_address[2] & $st1_log_array[$x] ) then 
		$read_return =  _open_and_look($server_address[2],$st1_log_array[$x],$st1_logkeyword_array[$x])	
		;MsgBox(0,$x & 'Return to real world', $read_return,10)
		if $read_return=0 then 
		$mail_body=$mail_body &	 " 伺服器: " &$server_address[2] &	 "  檔案: " & $st1_log_array[$x] &  "  KeyWord: " & $st1_logkeyword_array[$x]  &	 "  不存在" & @CRLF
			
			
		EndIf
	Elseif not StringInStr($st1_log_array[$x],"fail.log") then
	$mail_body = $mail_body & @CRLF &	" 伺服器: " &$server_address[2] &	 "   檔案: " & $st1_log_array[$x] &  "   不存在" & @CRLF & @CRLF
	Else
	EndIf
Next


	if $test_mode then 
		; Use default setting
	Else
		$s_ToAddress = "bryant@dynalab.com.tw";"davidliu@net1.com.tw;davidliu@dynalab.com.tw"
		$s_CcAddress = ""
		$s_BccAddress = ""
	EndIf
			
	if not $mail_body="" and $hour >=8 Then		
		$as_Body= $mail_body
		$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $IPPort, $ssl)
	Elseif not $mail_body="" and $hour <8 Then	
		_FileWriteLog(@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & "_" & $year & $month & $day & ".log", $mail_body )
	
	Else
		MsgBox(0,"無錯誤", @ScriptName & " 無錯誤",5)
	EndIf
	
	;run ( @ComSpec & " /c " & @ScriptDir& "\AE_mail_sync_log_0.1.exe"  ,"" , @SW_HIDE )
	;sleep(5000)
	;run ( @ComSpec & " /c " & @ScriptDir& "\analysis_mail_log_0.1.exe "  ,"" , @SW_HIDE )

Exit


;dim $logname, $keyword  11.05.16_
Func _open_and_look ($server_add, $logname, $keyword )
	;MsgBox(0,'1 file is at', $server_add & $logname  &"  "& $keyword)
	local $fileread_array
	_FileReadToArray ($server_add& $logname, $fileread_array)
	;_ArrayDisplay($fileread_array)
	;MsgBox(0,'log time format check', StringTrimLeft($year,2)&"."& $month &"."&$day)
	if not StringInStr ($fileread_array[1], StringTrimLeft($year,2)&"."& $month &"."&$day ) then return(0)
	
	for $x = UBound($fileread_array)-1 to 1 step -1
		
		if StringInStr( $fileread_array[$x], $keyword ) then 
			;MsgBox(0,'Line', $x)
			return ( $x )
			;ExitLoop
						
		EndIf	
			
	Next	
	Return( 0 )
EndFunc



Func _TEST_MODE()
	
	IF FileExists(@ScriptDir&"\TESTMODE.txt") Then
		$mode=FileReadLine(@ScriptDir&"\TESTMODE.txt",1)
		if $mode=1 then 
			MsgBox(0,"Test mode", "測試模式"&@CRLF& " ",5)
			
		Else
			;MsgBox(0,"Process mode", " 高鐵車次資料會輸入資料庫 ",10)
			;$ans=InputBox("Process mode","高鐵車次資料會輸入資料庫 "&@CRLF& "輸入 N 可以離開")
			
			$mode=0
			MsgBox(0,"Process mode", "正式模式"&@CRLF& " ",5)
			;if $ans="n" or $ans="N" or @error=1 then exit
		EndIf
	
	Else
		;MsgBox(0,"Process mode", " 高鐵車次資料會輸入資料庫 ",10)
			;$ans=InputBox("Process mode","高鐵車次資料會輸入資料庫 "&@CRLF& "輸入 N 可以離開")
			
			$mode=0
			MsgBox(0,"Process mode", "正式模式"&@CRLF& "",5)
			;if $ans="n" or $ans="N" or @error=1 then exit
		
	EndIf
	
	return $mode
EndFunc




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

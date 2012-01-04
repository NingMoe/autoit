


#include <array.au3>
#include <_pop3.au3>
#include <IE.au3>
#include <file.au3>
#include <Date.au3>
#include <string.au3> 

dim $sec=@SEC
dim $min=@MIN
Dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR
DIM $WeekOfDay = @WDAY   ; Return 1-7  represent from Sunday , Monday ~ Saturday
dim $report_mail_srv_down=0
dim $report_sms_srv_down=0
dim $restart_tomcat=0
Global $test_mode , $sleep_interval=60
Global $mail_count=0 , $last_mail_count=0 , $mailsrv_down_count=0

$test_mode=_TEST_MODE() ; return 1 means  Test mode.


Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
Global $oMyRet[2]

dim $batpath_a=@ScriptDir
dim $bat1="tomcat.bat"
;



	$s_SmtpServer = "smtp.gmail.com"              ; address for the smtp-server to use - REQUIRED
	$s_FromName = "changtun"                      ; name from who the email was sent
	$s_FromAddress = "changtun@gmail.com" ;  address from where the mail should come
	$s_ToAddress = "bryant@net1.com.tw"   ; destination address of the email - REQUIRED
	$s_Subject = "AE_SMS_Urgent通知信件"                   ; subject from the email - can be anything you want it to be
	$as_Body = "AE_SMS_Urgent通知信件"        ; the messagebody from the mail - can be left blank but then you get a blank mail
	$s_AttachFiles = ""                        ; the file you want to attach- leave blank if not needed sample :"d:\ibm240KB.jpg"   
	$s_CcAddress = ""       ; address for cc - leave blank if not needed
	$s_BccAddress = ""     ; address for bcc - leave blank if not needed
	$s_Username = "changtun@gmail.com"                    ; username for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
	$s_Password = "9ps567*9"                  ; password for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
	;$IPPort = 25                              ; port used for sending the mail
	;$ssl = 0                                  ; enables/disables secure socket layer sending - put to 1 if using httpS
	 $IPPort=465                            ; GMAIL port used for sending the mail
	 $ssl=1   



while 1 

	$mail_count=_get_email()
	
	Select
		case $mail_count >0
			$last_mail_count=$last_mail_count+1
				;MsgBox(0, "SMS server", "Mail server has "&  $mail_count &" mails",3)
			if $last_mail_count  >5 and $report_sms_srv_down=0 then 
				MsgBox(0, "SMS server", "AE SMS server己經連續 " & ($sleep_interval/60)*$last_mail_count & " 分鐘未發簡訊了" ,3)
				
				$s_Subject = "AE SMS server己經連續 " & ($sleep_interval/60)*$last_mail_count & " 分鐘未發簡訊了"                   ; subject from the email - can be anything you want it to be
				$as_Body =  "AE SMS server己經連續 " & ($sleep_interval/60)*$last_mail_count & " 分鐘未發簡訊了"   
				$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $IPPort, $ssl)
				$report_sms_srv_down=1
				
				if $restart_tomcat=0 then 
					WinMinimizeAll()
					_batch($batpath_a,$bat1)
					$restart_tomcat=1
				EndIf	
			EndIf
			
			if  mod($last_mail_count,100)=0 Then $report_sms_srv_down=0
	
		Case $mail_count=0 
				$last_mail_count=0
				$report_mail_srv_down=0
				$report_sms_srv_down=0
				$mailsrv_down_count=0
				$restart_tomcat=0
				;MsgBox(0, "Mail server", "SMS server Normal",3)
			
		case $mail_count< 0   ;  Now is only -1
 			$last_mail_count=0
			$mailsrv_down_count=$mailsrv_down_count+1
			;MsgBox(0, "Mail server", "Onlinebooking Mail server error",3)	
			if $mailsrv_down_count>5 and  $report_mail_srv_down=0 then
				$s_Subject = "Urgent! AE SMS server"               ; subject from the email - can be anything you want it to be
				$as_Body =  _NowDate() &" "&_NowTime()& "AE SMS server 無法連線到 onlinebooking.com.tw"      
				$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $IPPort, $ssl)
			EndIf
			
			if mod( $mailsrv_down_count,100)=0 then $mailsrv_down_count=0
	EndSelect
sleep(1000* $sleep_interval)
WEnd


Func _TEST_MODE()
	
	IF FileExists(@ScriptDir&"\TESTMODE.txt") Then
		$mode=FileReadLine(@ScriptDir&"\TESTMODE.txt",1)
		if $mode=1 then 
			MsgBox(0,"Test mode", "測試模式"&@CRLF,5)
			
		Else
			;MsgBox(0,"Process mode", " 高鐵車次資料會輸入資料庫 ",10)
			;$ans=InputBox("Process mode","高鐵車次資料會輸入資料庫 "&@CRLF& "輸入 N 可以離開")
			
			$mode=0
			MsgBox(0,"Normal mode", "正式模式"&@CRLF& " ",5)
			;if $ans="n" or $ans="N" or @error=1 then exit
		EndIf
	
	Else
		;MsgBox(0,"Process mode", " 高鐵車次資料會輸入資料庫 ",10)
			;$ans=InputBox("Process mode","高鐵車次資料會輸入資料庫 "&@CRLF& "輸入 N 可以離開")
			
			$mode=0
			MsgBox(0,"Normal mode", "正式模式"&@CRLF& " ",5)
			;if $ans="n" or $ans="N" or @error=1 then exit
		
	EndIf
	
	return $mode
EndFunc





Func _get_email()
;; mail download
;;
;ConsoleWrite(@AutoItVersion & @CRLF)
;~ See _pop3.au3 for a complete description of the pop3 functions
;~ Requires AU3 beta version 3.1.1.110 or newer.

;Global $MyPopServer = "56168.com.tw"
;Global $MyLogin = "bryant"
;Global $MyPasswd = "9ps56789"

;Global $MyPopServer = "onlinebooking.com.tw"
;Global $MyLogin = "direct_send"
;Global $MyPasswd = "amextravel"

;Global $MyPopServer = "maild.digitshuttle.com"
;Global $MyLogin = "servicemonitor"
;Global $MyPasswd = "9ps5678"
if not $test_mode then
	 $MyPopServer = "onlinebooking.com.tw"
	 $MyLogin = "sms"
	 $MyPasswd = "dynalab"
Else	
	$MyPopServer = "maild.digitshuttle.com"
	$MyLogin = "bryant"
	$MyPasswd = "9ps567*9"
EndIf

local $pop3_port=110 
local $_emails=0


_pop3Connect($MyPopServer, $MyLogin, $MyPasswd, $pop3_port)
If @error Then
	MsgBox(0, "Error", "Unable to connect to " & $MyPopServer & @CR & @error)
	Exit
Else
	ConsoleWrite("Connected to server pop3 " & $MyPopServer & @CR)
EndIf

;Local $stat = _Pop3Stat()
;If Not @error Then
;	_ArrayDisplay($stat, "Result of STAT COMMAND")
;Else
;	ConsoleWrite("Stat commande failed" & @CR)
;EndIf
;
; 取得信件的總數量，是一個二維陣列
 Local $list = _Pop3List()
 If Not @error Then
	 ;MsgBox(0,"The mails count:",UBound($list)-1,5)
 	;_ArrayDisplay($list, "")
	$_emails=$list[0]
 Else
 	ConsoleWrite("List commande failed" & @CR)
	$_emails=-1
 EndIf

;~ Local $noop = _Pop3Noop()
;~ If Not @error Then
;~ 	ConsoleWrite($noop & @CR)
;~ Else
;~ 	ConsoleWrite("List commande failed" & @CR)
;~ EndIf

;~ Local $uidl = _Pop3Uidl()
;~ If Not @error Then
;~ 	_ArrayDisplay($uidl, "")
;~ Else
;~ 	ConsoleWrite("Uidl commande failed" & @CR)
;~ EndIf


;~ Local $dele = _Pop3Dele(1)
;~ If Not @error Then
;~ 	ConsoleWrite($dele & @CR)
;~ Else
;~ 	ConsoleWrite("Dele commande failed" & @CR)
;~ EndIf

ConsoleWrite(_Pop3Quit() & @CR)
_pop3Disconnect()

Return $_emails
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



Func _batch($batpath,$bat)
;Local $bat
;local $batpath
local $pid

Opt("WinTitleMatchMode", -2) 
;; Opt("WinTitleMatchMode", 2) 2 is to set up Autoit to search title by any string in Window title.
;; Opt("WinTitleMatchMode", -2) -2 is to set up Autoit to search title by any string in Window title with Upper Case or Lower case.  
;;
if FileExists($batpath&"\"&$bat) then 
;_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",$bat&" is there")
 
	if WinExists ("- "&StringTrimRight($bat,4)) then 
		;MsgBox(0,"There your Are.", "The window exist");
		WinKill("- "&StringTrimRight($bat,4))
		;_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",$bat&" is killed")
		sleep(500)
		;BlockInput(1)
		
		$pid=run($batpath&"\"&$bat,$batpath)
		_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",$bat&" pid is "&$pid)
		;MsgBox(0,"New window","New Windows PID:"&$pid)
		sleep(200)
		WinSetTitle("C:\WINNT\system32\cmd.exe","","C:\WINNT\system32\cmd.exe - "&$bat)
		WinSetTitle("C:\WINDOWS\system32\cmd.exe","","C:\WINDOWS\system32\cmd.exe - "&$bat)	

		;run("cmd.exe",$batpath)
		;;WinActivate("C:\WINNT\system32\cmd.exe")
		;;if WinActive("C:\WINNT\system32\cmd.exe") then MsgBox(0,"Active","cmd.exe is active",3)
		;sleep(500)
		;send($bat&"{enter}")
		
		;BlockInput(0)
			if WinExists ("- "&$bat) then 
			_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",$bat&" is runing again.")
			EndIf
		;MsgBox(0,"There your Are.", "Wait for the window exist");
		;WinSetState("- "&$bat,"cmd.exe",@SW_MINIMIZE )
		Else
		BlockInput(1)
	
		$pid=run($batpath&"\"&$bat,$batpath)
		_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",$bat&" pid is "&$pid)
		sleep(200)
		;MsgBox(0,"New window","New Windows Pid:"&$pid,3)
		WinSetTitle("C:\WINNT\system32\cmd.exe","","C:\WINNT\system32\cmd.exe - "&$bat)
		WinSetTitle("C:\WINDOWS\system32\cmd.exe","","C:\WINDOWS\system32\cmd.exe - "&$bat)	
		
		;run("cmd.exe",$batpath)
		;;WinActivate("C:\WINNT\system32\cmd.exe")
		;sleep(500)
		;	;if WinActive("c:\winnt\system32\cmd.exe") then MsgBox(0,"Active","cmd.exe is active")
		;send($bat&"{enter}")
		BlockInput(0)
		if WinExists ("- "&$bat) then 
		_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",$bat&" is running now.")
		endif
	EndIf
	sleep(3000)	
	WinSetState("C:\WINNT\system32\cmd.exe - "&$bat, "", @SW_MINIMIZE)
	WinSetState("C:\WINDOWS\system32\cmd.exe - "&$bat, "", @SW_MINIMIZE)
Else
_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",$bat&" is not there")
	
EndIf
EndFunc

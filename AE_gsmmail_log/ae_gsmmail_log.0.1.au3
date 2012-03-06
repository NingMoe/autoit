; Test 20090408 by bryant
#cs ----------------------------------------------------------------------------
	
	AutoIt Version: 3.2.8.1 (beta)
	Author:         Prog@ndy
	
	Script Function:
	MySQL-Plugin Demo Script
	
#ce ----------------------------------------------------------------------------
;;
;;
;;
;;
;;============================
;; 這是為了寫出發簡訊時的每日通話記錄，為了 david 的對帳之用 2012/02


#include <array.au3>
#include <mysql.au3>
#Include <File.au3>
#include <date.au3>
#include<string.au3>
#include <_array2string_tab.au3>

dim $sec=@SEC
dim $min=@MIN
Dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR

;
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
Global $oMyRet[2]

dim $test_mode= _TEST_MODE()
dim $log_date=$year&$month&$day
;$log_date="20111002"
dim $log_range=0 
dim $a_log_range=""
dim $logfile

$log_range=_log_range()

if $log_range then 	
	$a_log_range=_log_range_date()
	;_ArrayDisplay($a_log_range)
EndIf	
	
;;
;; Connect My SQL for mail address.
;; DB is now at 10.112.55.87
;;
 dim $db_ip="10.112.55.87"
;dim $db_ip ="59.124.45.87"
dim  $my_http_base="";st1.onlinebooking.com.tw"
;;=============
;; This is to prepare mail body 
;; Need to change [user_email] bracket in the string
;dim $single_email="bryant@dynalab.com.tw"
dim  $mymailbody="";_read_direct_fly(@ScriptDir&"\direct_fly.htm", $my_http_base)
dim $mymailsubject="";_tract_mail_subject(@ScriptDir&"\direct_fly.htm")
;	if StringInStr($mymailbody,"[user_email]") then 
;	$mymailbody=StringReplace($mymailbody,"[user_email]",$single_email)
;	;MsgBox(0,"Message",$aRecords[$x])
;	EndIf
;MsgBox(0,"mail body",$mymailbody)
;;=============

;;
;; This is  connect to My SQL for user email
;;
; MYSQL starten, DLL im PATH (enth鄟t auch @ScriptDir), sont Pfad zur DLL angeben. DLL muss libmysql.dll hei絽n.
_MySQL_InitLibrary()
If @error Then Exit MsgBox(0, '', "")
;MsgBox(0, "DLL Version:",_MySQL_Get_Client_Version()&@CRLF& _MySQL_Get_Client_Info())

$MysqlConn = _MySQL_Init()

;Fehler Demo: C:\InstantRails-2.0-win\mysql\data\cookbook
;MsgBox(0,"Fehler-Demo","Fehler-Demo")
$connected = _MySQL_Real_Connect($MysqlConn,$db_ip,"alongbird","pkpkpk","test")
If $connected = 0 Then
	$errno = _MySQL_errno($MysqlConn)
	MsgBox(0,"Error:",$errno & @LF & _MySQL_error($MysqlConn))
	If $errno = $CR_UNKNOWN_HOST Then MsgBox(0,"Error:","$CR_UNKNOWN_HOST" & @LF & $CR_UNKNOWN_HOST)
Endif


if $log_range then 
	for $x =1 to UBound ($a_log_range)-1
	$log_date= $a_log_range[$x]
	;$query = "SELECT * FROM extra_email where extra_nbr='00000030640'"
	;$query = "SELECT * FROM extra_email"
	$query = 'SELECT * FROM ae_gsm_mail WHERE ftime>="' &$log_date& '000000" AND ftime<="' &$log_date& '235959" ORDER BY ftime'
	_MySQL_Real_Query($MysqlConn, $query)
	$res = _MySQL_Store_Result($MysqlConn)
	;$fields = _MySQL_Num_Fields($res)
	;$rows = _MySQL_Num_Rows($res)
	;MsgBox(0, "", $rows & "-" & $fields)

	;MsgBox(0, '', "Zugriff Methode 3 - alles in ein 2D Array")
	$db_return_data = _MySQL_Fetch_Result_StringArray($res)
	;_ArrayDisplay($db_return_data)
	
	;MsgBox(0,"string", _array2string_tab( $db_return_data ,10) )
	if IsString ( _array2string_tab( $db_return_data ,10)  ) then
		$logfile = FileOpen(@ScriptDir &"\"&$log_date&".txt", 10)
		FileWriteLine ( $logfile, _array2string_tab( $db_return_data ,10)  )
		
		FileClose($logfile)
	EndIf
	
	sleep(1000)
	Next
Else
	; Not an date range
	$query = 'SELECT * FROM ae_gsm_mail WHERE ftime>="' &$log_date& '000000" AND ftime<="' &$log_date& '235959" ORDER BY ftime'
	_MySQL_Real_Query($MysqlConn, $query)
	$res = _MySQL_Store_Result($MysqlConn)
	;$fields = _MySQL_Num_Fields($res)
	;$rows = _MySQL_Num_Rows($res)
	;MsgBox(0, "", $rows & "-" & $fields)

	;MsgBox(0, '', "Zugriff Methode 3 - alles in ein 2D Array")
	$db_return_data = _MySQL_Fetch_Result_StringArray($res)
	_ArrayDisplay($db_return_data)

EndIf


; Abfrage freigeben
_MySQL_Free_Result($res)
; Verbindung beenden
_MySQL_Close($MysqlConn)
; MYSQL beenden
_MySQL_EndLibrary()
;MsgBox(0,"Email", "There are :" & (UBound($db_return_data,1)-1) &" email addresses")


Exit


; This is send gmail  function
;
;##################################
; Variables
;##################################
;
$s_SmtpServer = "onlinebooking.com.tw" ;"maild.digitshuttle.com"              ; address for the smtp-server to use - REQUIRED
$s_FromName = "ae_direct_fly@onlinebooking.com.tw" ;"bryant@dynalab.com.tw"                      ; name from who the email was sent
$s_FromAddress = "ae_direct_fly@onlinebooking.com.tw";"bryant@dynalab.com.tw" ;  address from where the mail should come
$s_ToAddress = "direct_send@onlinebooking.com.tw"   ; destination address of the email - REQUIRED
$s_Subject = "美國運通旅遊服務部－24小時隨手查"   
$as_Body = "24小時隨手查"                             ; the messagebody from the mail - can be left blank but then you get a blank mail
$s_AttachFiles = ""                       ; the file you want to attach- leave blank if not needed
$s_CcAddress = ""					      ; address for cc - leave blank if not needed
$s_BccAddress = ""     					  ; address for bcc - leave blank if not needed
;$s_Username = _Base64Encode("ae_direct_fly")                    ; username for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
;$s_Password = _Base64Encode("pkpkpk")                  ; password for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
$s_Password = "pkpkpk"
$s_Username = "ae_direct_fly"
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
$m_Subject =  "24小時隨手查-mail-status"                   ; subject from the email - can be anything you want it to be
$m_as_Body =  "24小時隨手查-mail-status"             ; the messagebody from the mail - can be left blank but then you get a blank mail
$m_AttachFiles = @ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"             ; the file you want to attach- leave blank if not needed sample :"d:\ibm240KB.jpg"   
$m_CcAddress = "rita.j.liu@aexp.com"       ; address for cc - leave blank if not needed
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

	if $test_mode=1 then 
		$end_point=10
	Else
		$end_point=(UBound($db_return_data,1)-1)
	EndIf

for $r=1 to $end_point  
	
		Dim $day=@MDAY
		Dim $month=@MON
		DIM $year=@YEAR
		$m_AttachFiles = @ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"
		$as_Body=$mymailbody
		if StringInStr($as_Body,"[user_email]") then 
			$as_Body=StringReplace($as_Body,"[user_email]",$db_return_data[$r][0])
			;$as_Body=StringReplace($as_Body,"[user_email]","ae@delta.com.tw") ; For test only.
			
		EndIf
		If $test_mode=1 then
			MsgBox(0, "TestMode","Mails should send to "&$db_return_data[$r][0]&@CRLF&" Now send to direct_send@onlinebooking.com",5)
		else	
			$s_ToAddress = $db_return_data[$r][0] ;Correct mail to 
		EndIf
		;$m_ToAddress =	$db_return_data[$r][0] ; 
		$s_Subject = $mymailsubject
		;$as_Body= "Updated at "&$year&$month&$day& @CRLF &$as_Body ; Correct sentence
		
		;$as_Body= "Updated at "&$year&$month&$day& @CRLF & $db_return_data[$r][0] &@CRLF &$as_Body ; for test only. To locate email address in mail body
		
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
	$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $s_IPPort, $s_ssl)
		;$rc = _INetSmtpMailCom($m_SmtpServer, $m_FromName, $m_FromAddress, $m_ToAddress, $m_Subject, $as_Body, $m_AttachFiles, $m_CcAddress, $m_BccAddress, $m_Username, $m_Password, $IPPort, $ssl)
		_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"," Mail Send to "& $db_return_data[$r][0])
		if mod($r, 3)=0		then 
			sleep(5000)
		EndIf
		if mod($r, 1000)=0 Then
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
		;$objEmail.HTMLBodyPart.Charset="utf8";
		$objEmail.BodyPart.Charset="utf-8";
        $objEmail.HTMLBody = $as_Body
    Else
		$objEmail.bodyPart.Charset="utf8";
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



Func _TEST_MODE()
	IF FileExists(@ScriptDir&"\TESTMODE.txt") Then
		$mode=FileReadLine(@ScriptDir&"\TESTMODE.txt",1)
		if $mode=1 then 
			MsgBox(0,"Test mode", "This is Test mode.",5)
			
		Else
			MsgBox(0,"Delivery mode", "This will access True DB.",5)
			$mode=0
		EndIf
	
	Else
		MsgBox(0,"Delivery mode", "This will access True DB.",5)
		$mode=0
	EndIf
	
	return $mode
EndFunc



Func _log_range()
	IF FileExists(@ScriptDir&"\log_range.txt") Then
		$mode=FileReadLine(@ScriptDir&"\log_range.txt",1)
		if $mode=1 then 
			MsgBox(0,"log_range mode", "This is log range mode. ",5)
			
		Else
			MsgBox(0,"Delivery mode", "This is daily mode.",5)
			$mode=0
		EndIf
	
	Else
		MsgBox(0,"Delivery mode", "This is daily mode.",5)
		$mode=0
	EndIf
	
	return $mode
EndFunc


Func _log_range_date()
	local $log_range_date_array[2]  ,$end_date , $start_date , $pointer_date
		$log_range_date_array[0]=1
		$log_range_date_array[1]=0
		;$log_range_date_array[2]=0
	
		IF FileExists(@ScriptDir&"\log_range.txt") Then
			
		$log_range_date_array[1] =FileReadLine(@ScriptDir&"\log_range.txt",2)
		
		$start_date= _StringInsert(  _StringInsert( $log_range_date_array[1],"/",4 )  ,"/" , 7)
		$end_date= _StringInsert(  _StringInsert( FileReadLine(@ScriptDir&"\log_range.txt",3) ,"/",4 )  ,"/" , 7)  
		
		$pointer_date= $start_date
			while 1
				$pointer_date=_DateAdd("d",1,$pointer_date)
					if _DateDiff ("D", $pointer_date , $end_date) >=0 then
					_ArrayAdd($log_range_date_array, StringReplace ($pointer_date , "/","") )
		
					;_ArrayDisplay($log_range_date_array)
					;MsgBox(0,"Date Diff", _DateDiff ("D", $pointer_date , $end_date) ,2)
					$log_range_date_array[0]=$log_range_date_array[0]+1
					Else
						;MsgBox(0,"Date Diff", _DateDiff ("D", $pointer_date , $end_date) ,5)
						ExitLoop
					EndIf
			WEnd	
		EndIf
		;_ArrayDisplay($log_range_date_array)
		return $log_range_date_array
EndFunc
	
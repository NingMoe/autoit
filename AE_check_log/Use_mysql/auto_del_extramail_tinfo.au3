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
;;20101116 modify for mail subject that is abstracted from mail body of the html file.
;;============================
;;20100721 modify for utf-8 encoding
;; Use onlinebooking.com.tw as mail server.
;; send mail by ae_direct_fly@onlinebooking.com.tw and return should be send to direct_send@onlinebooking.com.tw if need to analysis return mail.
;;============================
; Mail Account for AE
; MyOnlineBookingST1@gmail.com
; myonlinebookingst2@gmail.com
; myonlinebookingst3@gmail.com
; pass amextravel
;
;;
;; 20110804 modify for AE staff to use this mail sender.
; Wait for decision .  Not possible to send mail from AE network.
;
; 2011/12/15 modify for delete out-dated name list of extramail in db.
; sql command '  select * from test.extra_email where corp_id="tinfo" and deadline<= "20120312"  '
;
; 2012/05/29 修改為直接刪除過期的 extramial 記錄

#include <array.au3>
#include <mysql.au3>
#Include <File.au3>
#include <_array2string.au3>

dim $sec=@SEC
dim $min=@MIN
Dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR
DIM $today= $year& $month & $day
dim $mailbody_in_main

;
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
Global $oMyRet[2]

dim $test_mode= _TEST_MODE()
;;
;; Connect My SQL for mail address.
;; DB is now at 10.112.55.87
;;
; dim $db_ip="10.112.55.87"
dim $db_ip ="10.112.55.87"
;dim $db_ip ="10.112.55.102"
dim  $my_http_base="st1.onlinebooking.com.tw"
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
If  FileExists(@ScriptDir & "\extra_email.csv")  Then
	$extramail_array =  _file2Array(@ScriptDir&"\extra_email.csv",2,";")
	;_ArrayDisplay($extramail_array)
	
	;MsgBox(0,"_array2string", _array2string($extramail_array,2) )
Else
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


;Exit
; XAMPP cdcol
;MsgBox(0, "XAMPP-Cdcol-demo", "XAMPP-Cdcol-demo")

;$connected = _MySQL_Real_Connect($MysqlConn, "localhost", "root", "", "cookbook")
;If $connected = 0 Then Exit MsgBox(16, 'Connection Error', _MySQL_Error($MysqlConn))

;$query = "SELECT * FROM extra_email where extra_nbr='00000030640'"
;$query = "SELECT * FROM extra_email"
;$query = "SELECT email FROM extra_email"

;$today="20120312"

$query = 'select * from extra_email where corp_id="tinfo" and deadline< ' & $today
;$query = 'select * from extra_email where corp_id="tinfo" and deadline> ' & $today

_MySQL_Real_Query($MysqlConn, $query)


$res = _MySQL_Store_Result($MysqlConn)

;$fields = _MySQL_Num_Fields($res)

;$rows = _MySQL_Num_Rows($res)
;MsgBox(0, "", $rows & "-" & $fields)


;MsgBox(0, '', "Zugriff Methode 3 - alles in ein 2D Array")
$extramail_array = _MySQL_Fetch_Result_StringArray($res)
;_ArrayDisplay($extramail_array)



	if UBound($extramail_array)-1 >0 then 
		;_ArrayDisplay ($extramail_array)
		;MsgBox(0,"extra_mail" , _array2string_tab($extramail_array,7) )
		;for $x=1 to UBound($extramail_array)
		;	
		;Next	
		
		;MsgBox(0,"Select from DB " , "Record no. out-dated: " & UBound ($extramail_array)-1  &@CRLF& "現在要刪除這些名單了。")
	
	
	
	
	;	$query2 = 'delete from extra_email where corp_id="tinfo" and deadline< ' & $today
	;	_MySQL_Real_Query($MysqlConn, $query2)
	;	$res2 = _MySQL_Store_Result($MysqlConn)
	
	;	$extramail_array = _MySQL_Fetch_Result_StringArray($res2)
		
		$mailbody_in_main= "Extra_mail-清單刪除"& @CRLF &_array2string($extramail_array, 7)
		MsgBox(0,"Delete list", $mailbody_in_main)
		
		
		_sendmail_func($mailbody_in_main)
	
	EndIf	
	if UBound($extramail_array)-1 <0 then MsgBox(0,"Report",  "Now, no out-dated name in DB extra_mail table.", 5)
;===== If you select all from DB then you will need to use these code for filter.
; ; Y is from 1 X is from 0
; ;MsgBox(0,"Array((y,x)",$extramail_array[5][1])
; ;MsgBox(0,"Array((5,1)",$extramail_array[5][1])
;
; ;MsgBox(0,"Array((5,1)",UBound($extramail_array,1))
;;=====
;dim $email[UBound($extramail_array,1)]
;
;for $r=1 to (UBound($extramail_array,1)-1)
;		$email[$r]=$extramail_array[$r][1]
;		;_FileWriteFromArray(@ScriptDir&"\email_output.txt",$email,1 )
;		;_FileWriteToLine(@ScriptDir&"\email_output.txt",$r , $email[$r],1)
;Next
;$email[0]=UBound($extramail_array,1)
;_ArrayDisplay($email)

; Abfrage freigeben
	_MySQL_Free_Result($res)
	; Verbindung beenden
	_MySQL_Close($MysqlConn)
	; MYSQL beenden
	_MySQL_EndLibrary()

;MsgBox(0,"Email", "There are :" & (UBound($extramail_array,1)-1) &" email addresses")

;_ArrayDisplay($extramail_array)
EndIf
exit


Func _sendmail_func($mailbody_infunc)
; This is send gmail  function
;
;##################################
; Variables
;##################################
;
;$s_SmtpServer = "onlinebooking.com.tw" ;"maild.digitshuttle.com"              ; address for the smtp-server to use - REQUIRED
;$s_FromName = "ae_direct_fly@onlinebooking.com.tw" ;"bryant@dynalab.com.tw"                      ; name from who the email was sent
;$s_FromAddress = "ae_direct_fly@onlinebooking.com.tw";"bryant@dynalab.com.tw" ;  address from where the mail should come
;$s_ToAddress = "direct_send@onlinebooking.com.tw"   ; destination address of the email - REQUIRED
;$s_Subject = "美國運通旅遊服務部－24小時隨手查"   
;$as_Body = "24小時隨手查"                             ; the messagebody from the mail - can be left blank but then you get a blank mail
;$s_AttachFiles = ""                       ; the file you want to attach- leave blank if not needed
;$s_CcAddress = ""					      ; address for cc - leave blank if not needed
;$s_BccAddress = ""     					  ; address for bcc - leave blank if not needed
;;$s_Username = _Base64Encode("ae_direct_fly")                    ; username for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
;;$s_Password = _Base64Encode("pkpkpk")                  ; password for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
;$s_Password = "pkpkpk"
;$s_Username = "ae_direct_fly"
;$s_IPPort = 25                              ; port used for sending the mail
;$s_ssl = 1     ; Always use 1              ; enables/disables secure socket layer sending - put to 1 if using httpS
;;$IPPort=465                            ; GMAIL port used for sending the mail
;;$ssl=1                                 ; GMAILenables/disables secure socket layer sending - put to 1 if using httpS
;
;
;
$m_SmtpServer = "smtp.gmail.com"              ; address for the smtp-server to use - REQUIRED
$m_FromName = "BryantLu"                      ; name from who the email was sent
$m_FromAddress = "changtun@gmail.com" ;  address from where the mail should come

$m_ToAddress = "bryant@net1.com.tw"   ; destination address of the email - REQUIRED
$m_Subject =  $today &"Extra_mail-清單刪除"                   ; subject from the email - can be anything you want it to be
$m_as_Body = $mailbody_infunc          ; the messagebody from the mail - can be left blank but then you get a blank mail
$m_AttachFiles = "";@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"             ; the file you want to attach- leave blank if not needed sample :"d:\ibm240KB.jpg"   
$m_CcAddress =""; "rita.j.liu@aexp.com"       ; address for cc - leave blank if not needed
$m_BccAddress = ""     ; address for bcc - leave blank if not needed
$m_Username = "changtun@gmail.com"                    ; username for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
$m_Password = "9ps567*9"                  ; password for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
;$IPPort = 25                              ; port used for sending the mail
;$ssl = 0                                  ; enables/disables secure socket layer sending - put to 1 if using httpS
$m_IPPort=465                            ; GMAIL port used for sending the mail
$m_ssl=1                                 ; GMAILenables/disables secure socket layer sending - put to 1 if using httpS
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

	;if $test_mode=1 then 
	;	$end_point=10
	;Else
	;	$end_point=(UBound($extramail_array,1)-1)
	;EndIf

;for $r=1 to $end_point  
	
;		Dim $day=@MDAY
;		Dim $month=@MON
;		DIM $year=@YEAR
;		$m_AttachFiles = @ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"
;		$as_Body=$mymailbody
;		if StringInStr($as_Body,"[user_email]") then 
;			$as_Body=StringReplace($as_Body,"[user_email]",$extramail_array[$r][0])
;			;$as_Body=StringReplace($as_Body,"[user_email]","ae@delta.com.tw") ; For test only.
			
;		EndIf
;		If $test_mode=1 then
;			MsgBox(0, "TestMode","Mails should send to "&$extramail_array[$r][0]&@CRLF&" Now send to direct_send@onlinebooking.com",5)
;		else	
;			$s_ToAddress = $extramail_array[$r][0] ;Correct mail to 
;		EndIf
;		;$m_ToAddress =	$extramail_array[$r][0] ; 
;		$s_Subject = $mymailsubject
;		;$as_Body= "Updated at "&$year&$month&$day& @CRLF &$as_Body ; Correct sentence
		
;		;$as_Body= "Updated at "&$year&$month&$day& @CRLF & $extramail_array[$r][0] &@CRLF &$as_Body ; for test only. To locate email address in mail body
		
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
;	$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $s_IPPort, $s_ssl)
		;$rc = _INetSmtpMailCom($m_SmtpServer, $m_FromName, $m_FromAddress, $m_ToAddress, $m_Subject, $as_Body, $m_AttachFiles, $m_CcAddress, $m_BccAddress, $m_Username, $m_Password, $IPPort, $ssl)
;		_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"," Mail Send to "& $extramail_array[$r][0])
;		if mod($r, 3)=0		then 
;			sleep(5000)
;		EndIf
;		if mod($r, 1000)=0 Then
;			Dim $day=@MDAY
;			Dim $month=@MON
;			DIM $year=@YEAR
;			;local $my_attach=@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"
;		;$rc = _INetSmtpMailCom($m_SmtpServer, $m_FromName, $m_FromAddress, $m_ToAddress, $m_Subject, $m_as_Body, $m_AttachFiles, $m_CcAddress, $m_BccAddress, $m_Username, $m_Password, $IPPort, $ssl)
		$rc = _INetSmtpMailCom($m_SmtpServer, $m_FromName, $m_FromAddress, $m_ToAddress, $m_Subject, $m_as_Body, $m_AttachFiles, $m_CcAddress, $m_BccAddress, $m_Username, $m_Password, $m_IPPort, $m_ssl)

;			sleep(1000*60*5)
;		EndIf

	
;Next

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
		;$objEmail.HTMLBodyPart.Charset="utf8";
		$objEmail.BodyPart.Charset="utf-8";
        $objEmail.HTMLBody = $as_Body
    Else
		$objEmail.bodyPart.Charset="utf-8";
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



Func _read_direct_fly($directfly,$http_base)
	
local $aRecords
local $directfly_string=''
If Not _FileReadToArray($directfly,$aRecords) Then
	_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"," Error Reading file at "& @ScriptDir&"\direct_fly.htm")
		
   MsgBox(4096,"Error", " Error reading file.     error:" & @error)
   Exit
EndIf
For $x = 1 to $aRecords[0]
	$directfly_string=$directfly_string & $aRecords[$x]& @CRLF	
Next
	if StringInStr($directfly_string,"[http_base]") then 
	$directfly_string=StringReplace($directfly_string,"[http_base]",$http_base)
	;MsgBox(0,"Message",$aRecords[$x])
	EndIf
	;
	;if StringInStr($directfly_string,"[user_email]") then 
	;$directfly_string=StringReplace($directfly_string,"[user_email]",$single_email)
	;;MsgBox(0,"Message",$aRecords[$x])
	;EndIf

	;;Msgbox(0,'mail body', $directfly_string)
Return $directfly_string
EndFunc




Func _tract_mail_subject($directfly)
	
local $aRecords
local $directfly_string=''
If Not _FileReadToArray($directfly,$aRecords) Then
	_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"," Error Reading file at "& @ScriptDir&"\direct_fly.htm")
		
   MsgBox(4096,"Error", " Error reading file.     error:" & @error)
   Exit
EndIf
For $x = 1 to $aRecords[0]
	$directfly_string=$aRecords[$x]& @CRLF	

	if StringInStr($directfly_string,"美國運通旅遊服務部24小時隨手查-") then 
	$mail_subject= StringReplace(StringReplace($directfly_string,"</strong>",""),"<strong>","")
	;MsgBox(0,"Message",$aRecords[$x])
	EndIf
	;
	;if StringInStr($directfly_string,"[user_email]") then 
	;$directfly_string=StringReplace($directfly_string,"[user_email]",$single_email)
	;;MsgBox(0,"Message",$aRecords[$x])
	;EndIf
Next
	;Msgbox(0,'mail subject', $mail_subject)
Return $mail_subject
EndFunc


Func _TEST_MODE()
	IF FileExists(@ScriptDir&"\TESTMODE.txt") Then
		$mode=FileReadLine(@ScriptDir&"\TESTMODE.txt",1)
		if $mode=1 then 
			MsgBox(0,"Test mode", "This is Test mode. mail send to bryant@net1.com.tw",5)
			
		Else
			MsgBox(0,"Delivery mode", "This is True mail delivery.",5)
			$mode=0
		EndIf
	
	Else
		MsgBox(0,"Delivery mode", "This is True mail delivery.",5)
		$mode=0
	EndIf
	
	return $mode
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



Func _array2string_tab($array, $d)
	; $d 是傳入的 array 的欄位數目
	Local $y, $x, $i, $string_to_return, $a_line
	$string_to_return = ''
	
	MsgBox(0, "Dimention", "  $Y :" & UBound($array))
	
	For $y = 1 To UBound($array) - 1
		$a_line = ''
		If $d = 1 Then
			$a_line = $array[$y]
			
		Else
			For $x = 0 To $d - 1
				If $x < $d - 1 Then
					$a_line = $a_line & $array[$y][$x] & '^'
				Else
					$a_line = $a_line & $array[$y][$x]
				EndIf
				;MsgBox(0," array 2 string tab", $a_line)
				;ConsoleWrite($a_line & @crlf)
				;$string_to_return=$string_to_return  & $array[$y][1] &" , "& $array[$y][2] &" , " &$array[$y][6] &@CRLF
			Next
			ConsoleWrite($a_line & @CRLF)
			;$a_line=StringTrimRight($a_line,3) ; Cut off the last 3 character --> ","
		EndIf
		$string_to_return = $string_to_return & $a_line & @CRLF
	Next


	;MsgBox(0,"string", $string_to_return)
	Return $string_to_return
EndFunc   ;==>_array2string_tab

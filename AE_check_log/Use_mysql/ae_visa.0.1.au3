; Test 20090408 by bryant
#cs ----------------------------------------------------------------------------
	
	AutoIt Version: 3.2.8.1 (beta)
	Author:         Prog@ndy
	
	Script Function:
	MySQL-Plugin Demo Script
	
#ce ----------------------------------------------------------------------------
;;
;;
; 20120308 修改程式，成為被其他程式呼叫的對象，不需要再有GUI
;
#include <array.au3>
#include <mysql.au3>
#Include <File.au3>

dim $sec=@SEC
dim $min=@MIN
Dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR
dim $today=$year & $month & $day
;
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
Global $oMyRet[2]
Global $visa_process_number=1  ; Add at the end of the mail

dim $gui_mode=1
dim $evisa_nbr_txt_array

$gui_mode= _GUI_mode()


if $gui_mode then 

	$evisa_nbr=InputBox("Input","這是用來補寄 evisa 的信件的小程式 "&@CRLF&"請輸入 evisa_nbr")
	if $evisa_nbr="" then Exit

Else

	if FileExists(@ScriptDir & "\_evisa_nbr.txt") then 
		_FileReadToArray(@ScriptDir & "\_evisa_nbr.txt", $evisa_nbr_txt_array )
		if IsArray( $evisa_nbr_txt_array ) then filemove (@ScriptDir & "\_evisa_nbr.txt", @ScriptDir & "\evisa_nbr_reprocessed\" & $today& "_evisa_nbr.txt",9 )
		;_ArrayDisplay($evisa_nbr_txt_array) 
		if $evisa_nbr_txt_array[1]<>"201222 ;Evisa補單專用，每一行一個ea_nbr" then 
			_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",' EVISA resend 檔案錯誤，沒有 ";Evisa補單專用，每一行一個ea_nbr" 這樣的文字 ')
			exit
		EndIf
		_ArrayDelete($evisa_nbr_txt_array,1)
		$evisa_nbr_txt_array[0]=$evisa_nbr_txt_array[0]-1
		;_ArrayDisplay($evisa_nbr_txt_array) 
		
	EndIf


EndIf

For $a =1 to $evisa_nbr_txt_array[0] 
	$evisa_nbr=$evisa_nbr_txt_array[$a]
	
	_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",' Prepare EVISA resend : ' & $evisa_nbr )
			
;;
;; Connect My SQL for mail address.
;; DB is now at 10.112.55.87
;;
dim $db_ip="10.112.55.87"
;dim $db_ip="60.250.18.187"

;dim  $my_http_base="st1.onlinebooking.com.tw"
;;=============
;; This is to prepare mail body 
;; Need to change [user_email] bracket in the string
;dim $single_email="bryant@dynalab.com.tw"
dim  $mymailbody="";_read_direct_fly(@ScriptDir&"\direct_fly.htm", $my_http_base)
dim $mymailsubject="EVISA";_tract_mail_subject(@ScriptDir&"\direct_fly.htm")
;	if StringInStr($mymailbody,"[user_email]") then 
;	$mymailbody=StringReplace($mymailbody,"[user_email]",$single_email)
;	;MsgBox(0,"Message",$aRecords[$x])
;	EndIf
;MsgBox(0,"mail body",$mymailbody)
;;=============
dim $test_mode= _TEST_MODE()
;;
;; This is  connect to My SQL for user email
;;
; MYSQL starten, DLL im PATH (enth鄟t auch @ScriptDir), sont Pfad zur DLL angeben. DLL muss libmysql.dll hei絽n.
_MySQL_InitLibrary()
If @error Then Exit MsgBox(0, 'Error while Init MySQL', "Error at Init MySQL Line 80")
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
_MySQL_Set_Character_Set($MysqlConn,"big5")

;Exit
; XAMPP cdcol
;MsgBox(0, "XAMPP-Cdcol-demo", "XAMPP-Cdcol-demo")

;$connected = _MySQL_Real_Connect($MysqlConn, "localhost", "root", "", "cookbook")
;If $connected = 0 Then Exit MsgBox(16, 'Connection Error', _MySQL_Error($MysqlConn))

;$query = "SELECT * FROM extra_email where extra_nbr='00000030640'"
;$query = "SELECT * FROM extra_email"

$query = "SELECT * FROM evisa_app where ea_nbr='"& $evisa_nbr &"'"

_MySQL_Real_Query($MysqlConn, $query)


$res = _MySQL_Store_Result($MysqlConn)

;$fields = _MySQL_Num_Fields($res)

;$rows = _MySQL_Num_Rows($res)
;MsgBox(0, "", $rows & "-" & $fields)


;MsgBox(0, '', "Zugriff Methode 3 - alles in ein 2D Array")
$DB_fetched_array = _MySQL_Fetch_Result_StringArray($res)
;_ArrayDisplay($DB_fetched_array)

;===== If you select all from DB then you will need to use these code for filter.
; ; Y is from 1 X is from 0
; ;MsgBox(0,"Array((y,x)",$DB_fetched_array[5][1])
; ;MsgBox(0,"Array((5,1)",$DB_fetched_array[5][1])
;
; ;MsgBox(0,"Array((5,1)",UBound($DB_fetched_array,1))
;;=====
;dim $email[UBound($DB_fetched_array,1)]
;
;for $r=1 to (UBound($DB_fetched_array,1)-1)
;		$email[$r]=$DB_fetched_array[$r][1]
;		;_FileWriteFromArray(@ScriptDir&"\email_output.txt",$email,1 )
;		;_FileWriteToLine(@ScriptDir&"\email_output.txt",$r , $email[$r],1)
;Next
;$email[0]=UBound($DB_fetched_array,1)
;_ArrayDisplay($email)

; Abfrage freigeben
_MySQL_Free_Result($res)
; Verbindung beenden
_MySQL_Close($MysqlConn)
; MYSQL beenden
_MySQL_EndLibrary()

;MsgBox(0,"Email", "There are :" & (UBound($DB_fetched_array,1)-1) &"  records")
if $gui_mode then
	
	dim $confirm=0
	$confirm=MsgBox(1,"Confirm",$DB_fetched_array[1][0] &"  "& $DB_fetched_array[1][12])
	if $confirm= 2 then Exit 

EndIf

if not $DB_fetched_array[1][23] ="" then 
	$visa_process_number=1
	if not $DB_fetched_array[1][26] ="" then 
		$visa_process_number=$visa_process_number+1
		
		if not $DB_fetched_array[1][29] ="" then 
			$visa_process_number=$visa_process_number+1
		EndIf	
	EndIf
EndIf

;MsgBox(0,"visa : at the end", $DB_fetched_array[1][23] &" ; "& $DB_fetched_array[1][26] &" ; " & $DB_fetched_array[1][29] &" --->  "& $visa_process_number)
for $x=0 to 49 ;$y=0 to (UBound($DB_fetched_array,1)-1)
	local $string_count
			;if $DB_fetched_array[0][$x]="visa_item" then 
			;$string_count=StringInStr($DB_fetched_array[1][$x],"type=checkbox>") +13
			
			;$DB_fetched_array[1][$x]=    StringTrimRight ( StringTrimLeft ( $DB_fetched_array[1][$x], $string_count )  ,26)
			
			;EndIf
			$mymailbody=$mymailbody & $DB_fetched_array[0][$x] &","& $DB_fetched_array[1][$x] & " ";@CRLF
	
	
Next	
$mymailbody= $mymailbody & "visa," &$visa_process_number & " " & "send_visa_date," & $DB_fetched_array[1][47] &@CRLF
ConsoleWrite ($mymailbody)
;_FileWriteLog(@ScriptDir & "\my.log", $mymailbody)

if $gui_mode then 
	$confirm=0
	$confirm=MsgBox(1,"mail body", $mymailbody,30)
	if $confirm= 2 then Exit 

EndIf
; This is send gmail  function
;
;##################################
; Variables
;##################################
;
$s_SmtpServer = "onlinebooking.com.tw" ;"maild.digitshuttle.com"              ; address for the smtp-server to use - REQUIRED
;$s_SmtpServer = "[10.112.55.105]" ;"maild.digitshuttle.com"  
$s_FromName = "ae@onlinebooking.com.tw" ;"bryant@dynalab.com.tw"                      ; name from who the email was sent
$s_FromAddress = "ae@onlinebooking.com.tw";"bryant@dynalab.com.tw" ;  address from where the mail should come
$s_ToAddress = "dtramex@tpe.atti.net.tw"   ; destination address of the email - REQUIRED
$s_Subject = "EVISA"   
$as_Body = ""                             ; the messagebody from the mail - can be left blank but then you get a blank mail
$s_AttachFiles = ""                       ; the file you want to attach- leave blank if not needed
$s_CcAddress = ""					      ; address for cc - leave blank if not needed
$s_BccAddress = "changtun@gmail.com"     					  ; address for bcc - leave blank if not needed
;$s_Username = _Base64Encode("ae_direct_fly")                    ; username for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
;$s_Password = _Base64Encode("pkpkpk")                  ; password for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
$s_Password = "pkpkpk"
$s_Username = "ae"
$s_IPPort = 25                              ; port used for sending the mail
$s_ssl = 1     ; Always use 1              ; enables/disables secure socket layer sending - put to 1 if using httpS
;$IPPort=465                            ; GMAIL port used for sending the mail
;$ssl=1                                 ; GMAILenables/disables secure socket layer sending - put to 1 if using httpS
;
;
;
$m_SmtpServer = "smtp.gmail.com"              ; address for the smtp-server to use - REQUIRED
$m_FromName = "Net1"                      ; name from who the email was sent
$m_FromAddress = "bryant@digtishuttle.com" ;  address from where the mail should come

$m_ToAddress = "bryant@net1.com.tw"   ; destination address of the email - REQUIRED
$m_Subject =  "EVISA-resend-mail-status"                   ; subject from the email - can be anything you want it to be
$m_as_Body =  "EVISA-resend-mail-status"             ; the messagebody from the mail - can be left blank but then you get a blank mail
$m_AttachFiles = @ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"             ; the file you want to attach- leave blank if not needed sample :"d:\ibm240KB.jpg"   
$m_CcAddress =""; "rita.j.liu@aexp.com"       ; address for cc - leave blank if not needed
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



for $r=1 to (UBound($DB_fetched_array,1)-1);$end_point  
	
		Dim $day=@MDAY
		Dim $month=@MON
		DIM $year=@YEAR
		$m_AttachFiles = @ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"
		$as_Body=$mymailbody

		If $test_mode=1 then
			MsgBox(0, "TestMode","Mails should send on ea_nbr : "&$DB_fetched_array[$r][0]&@CRLF&" Now send to changtun@gmail.com",5)
			$s_ToAddress = "changtun@gmail.com"
			$s_BccAddress = "";"changtun@gmail.com"  

		EndIf

		$s_Subject = $mymailsubject
	
	;$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $m_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $s_IPPort, $s_ssl)
	$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $s_IPPort, $s_ssl)
		;$rc = _INetSmtpMailCom($m_SmtpServer, $m_FromName, $m_FromAddress, $m_ToAddress, $m_Subject, $as_Body, $m_AttachFiles, $m_CcAddress, $m_BccAddress, $m_Username, $m_Password, $IPPort, $ssl)
		_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"," EVISA resend by ea_nbr "& $DB_fetched_array[$r][0])
		;_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"," EVISA resend Data : "& $as_Body )
		if mod($r, 3)=0		then 
			sleep(5000)
		EndIf
		if $r= ( UBound($DB_fetched_array,1)-1) Then
			
		;local $my_attach=@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"
		;$rc = _INetSmtpMailCom($m_SmtpServer, $m_FromName, $m_FromAddress, $m_ToAddress, $m_Subject, $m_as_Body, $m_AttachFiles, $m_CcAddress, $m_BccAddress, $m_Username, $m_Password, $IPPort, $ssl)
		$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $m_ToAddress, $m_Subject, $m_as_Body, $m_AttachFiles, $m_CcAddress, $s_BccAddress, $s_Username, $s_Password, $s_IPPort, $s_ssl)

			sleep(1000*60*5)
		EndIf

	
Next

sleep( 15000 )
Next

Exit


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
			MsgBox(0,"Test mode", "This is Test mode. All mail send to changtun@gmail.com",5)
			
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



Func _GUI_mode()
	IF FileExists(@ScriptDir&"\ae_visa_gui.txt") Then
		$mode=FileReadLine(@ScriptDir&"\ae_visa_gui.txt",1)
		if $mode=1 then 
			MsgBox(0,"GUI mode", "This is GUI mode. Input ea_nbr manually.",5)
			
		Else
			MsgBox(0,"silent mode", "Read ea_nbr from text file.",5)
			$mode=0
		EndIf
	
	Else
		MsgBox(0,"silent mode", "Read ea_nbr from text file.",5)
		$mode=0
	EndIf
	
	return $mode
EndFunc
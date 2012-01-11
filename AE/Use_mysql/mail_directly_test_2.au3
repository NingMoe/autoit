; Test 20090408 by bryant
#cs ----------------------------------------------------------------------------
	
	AutoIt Version: 3.2.8.1 (beta)
	Author:         Prog@ndy
	
	Script Function:
	MySQL-Plugin Demo Script
	
#ce ----------------------------------------------------------------------------
;
; Mail Account for AE
; Can not use googlemail to send.
; Use onlinebooking to send .
; Change mail client to blat.exe
;
; MyOnlineBookingST1@gmail.com
; myonlinebookingst2@gmail.com
; myonlinebookingst3@gmail.com
; pass amextravel
;


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

;
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
Global $oMyRet[2]
;;
;; Connect My SQL for mail address.
;; DB is now at 10.112.55.87
;;
dim $db_ip="10.112.55.87"
dim  $my_http_base="st1.onlinebooking.com.tw"


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
$query = "SELECT * FROM extra_email"
$query = "SELECT email FROM extra_email"

_MySQL_Real_Query($MysqlConn, $query)


$res = _MySQL_Store_Result($MysqlConn)

;$fields = _MySQL_Num_Fields($res)

;$rows = _MySQL_Num_Rows($res)
;MsgBox(0, "", $rows & "-" & $fields)


;MsgBox(0, '', "Zugriff Methode 3 - alles in ein 2D Array")
$extramail_array = _MySQL_Fetch_Result_StringArray($res)
;_ArrayDisplay($extramail_array)

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

MsgBox(0,"Email", "There are :" & (UBound($extramail_array,1)-1) &" email addresses")

;;=============
;; This is to prepare mail body 
;; Need to change [user_email] bracket in the string
;dim $single_email="bryant@dynalab.com.tw"
dim  $myreturn=_ReplaceStringInFile(@ScriptDir&"\direct_fly.htm", "[http_base]",$my_http_base)

if $myreturn = -1 then
	_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"," Can not write directly_fly.htm file")
		
EndIf
	if FileExists(@ScriptDir&"\direct_fly.htm") then FileCopy (@ScriptDir&"\direct_fly.htm",@ScriptDir&"\st1_direct_fly.htm" ,1)
;	if StringInStr($mymailbody,"[user_email]") then 
;	$mymailbody=StringReplace($mymailbody,"[user_email]",$single_email)
;	;MsgBox(0,"Message",$aRecords[$x])
;	EndIf
MsgBox(0,"mail body","Check file st1_direct_fly.htm ")
;;=============

; This is send gmail  function
;
;##################################
; Variables
;##################################
;$s_SmtpServer = "smtp.gmail.com"              ; address for the smtp-server to use - REQUIRED
;$s_FromName = "ae_direct_fly"                      ; name from who the email was sent
;$s_FromAddress = "ae_direct_fly@onlinebooking.com.tw" ;  address from where the mail should come
;;$s_FromAddress = "changtun@gmail.com" ;  address from where the mail should come
;$s_ToAddress = "changtun@gmail.com"   ; destination address of the email - REQUIRED
;$s_Subject = "美國運通旅遊服務部－24小時隨手查－大陸直飛班機資訊"                   ; subject from the email - can be anything you want it to be
;$as_Body = "ae_direct_fly"            ; the messagebody from the mail - can be left blank but then you get a blank mail
;$s_AttachFiles = ""                        ; the file you want to attach- leave blank if not needed sample :"d:\ibm240KB.jpg"   
;$s_CcAddress = ""       ; address for cc - leave blank if not needed
;$s_BccAddress = ""     ; address for bcc - leave blank if not needed
;$s_Username = "MyOnlineBookingST1@gmail.com"                    ; username for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
;$s_Password = "amextravel"                  ; password for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
;;$IPPort = 25                              ; port used for sending the mail
;;$ssl = 0                                  ; enables/disables secure socket layer sending - put to 1 if using httpS
;$IPPort=465                            ; GMAIL port used for sending the mail
;$ssl=1                                 ; GMAILenables/disables secure socket layer sending - put to 1 if using httpS
;
$s_SmtpServer = "onlinebooking.com.tw"              ; address for the smtp-server to use - REQUIRED
$s_FromName = "ae_direct_fly"                      ; name from who the email was sent
$s_FromAddress = "ae_direct_fly@onlinebooking.com.tw" ;  address from where the mail should come
;$s_FromAddress = "changtun@gmail.com" ;  address from where the mail should come
$s_ToAddress = "changtun@gmail.com"   ; destination address of the email - REQUIRED
$s_Subject = "美國運通旅遊服務部－24小時隨手查－大陸直飛班機資訊"                   ; subject from the email - can be anything you want it to be
$s_as_Body = @ScriptDir& "\st1_direct_fly.htm"            ; the messagebody from the mail - can be left blank but then you get a blank mail
$s_AttachFiles = ""                        ; the file you want to attach- leave blank if not needed sample :"d:\ibm240KB.jpg"   
$s_CcAddress = ""       ; address for cc - leave blank if not needed
$s_BccAddress = ""     ; address for bcc - leave blank if not needed
$s_Username = "ae_direct_fly"                    ; username for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
$s_Password = "pkpkpk"                  ; password for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
$s_IPPort = 25                              ; port used for sending the mail
$s_ssl = 0                                  ; enables/disables secure socket layer sending - put to 1 if using httpS
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

for $r=1 to  10 ;(UBound($extramail_array,1)-1)
	
		Dim $day=@MDAY
		Dim $month=@MON
		DIM $year=@YEAR
		
		if $r=1 then 
			$myreturn=_ReplaceStringInFile(@ScriptDir&"\st1_direct_fly.htm", "[user_email]",$extramail_array[$r][0])
			;$as_Body=StringReplace($as_Body,"[user_email]",$extramail_array[$r][0])
			;$as_Body=StringReplace($as_Body,"[user_email]","ae@delta.com.tw") ; For test only.			
		EndIf
		
		if $r>1 then 
			$myreturn=_ReplaceStringInFile(@ScriptDir&"\st1_direct_fly.htm", $extramail_array[$r-1][0],$extramail_array[$r][0])
			;$as_Body=StringReplace($as_Body,"[user_email]",$extramail_array[$r][0])
			;$as_Body=StringReplace($as_Body,"[user_email]","ae@delta.com.tw") ; For test only.			
		EndIf
		
		;$m_ToAddress =	$extramail_array[$r][0] ; 
		$s_ToAddress = $extramail_array[$r][0] ;Correct mail to 
		
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
	runwait(@ScriptDir&"\blat.exe " & @ScriptDir&"\st1_direct_fly.htm -charset big5 -subject 美國運通旅遊服務部24小時隨手查-兩岸直航班機資訊 -to changtun@gmail.com  -server onlinebooking.com.tw -f ae_direct_fly@onlinebooking.com.tw -i ae_direct_fly@onlinebooking.com.tw -port 25 -u ae_direct_fly -pw pkpkpk","",@SW_HIDE )
		
	;run(@ScriptDir&"\blat.exe " & @ScriptDir&"\st1_direct_fly.htm -charset big5 -subject 美國運通旅遊服務部24小時隨手查-兩岸直航班機資訊 -to " & $extramail_array[$r][0] & "  -server onlinebooking.com.tw -f ae_direct_fly@onlinebooking.com.tw -i ae_direct_fly@onlinebooking.com.tw -port 25 -u ae_direct_fly -pw pkpkpk")
	;$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $IPPort, $ssl)
	;$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $s_IPPort, $s_ssl)
		;$rc = _INetSmtpMailCom($m_SmtpServer, $m_FromName, $m_FromAddress, $m_ToAddress, $m_Subject, $as_Body, $m_AttachFiles, $m_CcAddress, $m_BccAddress, $m_Username, $m_Password, $IPPort, $ssl)
		_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log"," Mail Send to "& $extramail_array[$r][0])
		if mod($r, 3)=0		then 
			sleep(5000)
		EndIf
		if mod($r, 1000)=0 Then
		$rc = _INetSmtpMailCom($m_SmtpServer, $m_FromName, $m_FromAddress, $m_ToAddress, $m_Subject, $m_as_Body, $m_AttachFiles, $m_CcAddress, $m_BccAddress, $m_Username, $m_Password, $IPPort, $ssl)

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
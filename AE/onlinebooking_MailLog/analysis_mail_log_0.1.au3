#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
; 改 $MailLog_dir="\\10.112.55.105\Log"

#include <array.au3>
#include <_pop3.au3>
#include <file.au3>
#include <Date.au3>
#include <Zip.au3>
#include <string.au3> 

dim $sec=@SEC
dim $min=@MIN
Dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR
DIM $WeekOfDay = @WDAY   ; Return 1-7  represent from Sunday , Monday ~ Saturday

Global $test_mode , $continue_mode ,$FileList


Global $MailLog_dir=@ScriptDir
Global $MailLog_dir="\\10.112.55.105\Log";"c:\RaidenMAILD\log"

dim $today=  $year&$month&$day
dim $yesterday_1=   StringReplace (_DateAdd( 'd',-1, _NowCalcDate()) ,"/",""  )

Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")

$s_SmtpServer = "smtp.gmail.com"              ; address for the smtp-server to use - REQUIRED
$s_FromName = "changtun"                      ; name from who the email was sent
$s_FromAddress = "changtun@gmail.com" ;  address from where the mail should come
$s_ToAddress = "bryant@dynalab.com.tw"   ; destination address of the email - REQUIRED
$s_Subject = "ae_direct_fly每日信件統計"                   ; subject from the email - can be anything you want it to be
$as_Body = "ae_direct_fly每日信件統計"          ; the messagebody from the mail - can be left blank but then you get a blank mail
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


; MsgBox(0,"Yesterday", $MailLog_dir &"\" & $yesterday&".dtl")
$test_mode=_TEST_MODE() ; return 1 means  Test mode.
$continue_mode= _CONTINUE_MODE()

if $continue_mode then 
		$FileList= _FileListToArray ( $MailLog_dir, "*.dtl",1)
		_ArrayDelete($FileList,0)
		_ArraySort($FileList)
		;_ArrayDisplay ($FileList)
		for $x=0 to UBound($FileList)-1
			$yesterday_1= StringReplace($FileList[$x],".dtl","")
			;MsgBox (0,"File name as date" ,  $FileList[$x] &" --> "& $yesterday_1 )
			_analysis($yesterday_1)
			;sleep(30000)
		Next
Else
_analysis($yesterday_1)

EndIf


Func _analysis($yesterday )
local $MailLog , $analysis_result ,$analysis_string, $counter
$analysis_string=""
$counter=0
if FileExists($MailLog_dir &"\" & $yesterday&".dtl") Then
	;MsgBox(0,"Yesterday", $MailLog_dir &"\" & $yesterday&".dtl")
	 _FileReadToArray($MailLog_dir &"\" & $yesterday&".dtl" ,$MailLog) 
	;_ArrayDisplay ($MailLog)
	;MsgBox(0,"Ubound of $MailLog", UBound ( $MailLog ) )
	$analysis_result=FileOpen (@ScriptDir& "\"& $yesterday &"_maillog_analysis.txt",10  )
	for $x =0 to UBound ($MailLog)-1 
		if StringInStr ( $MailLog[$x],  "傳送郵件從 ae_direct_fly@onlinebooking.com.tw") then
			$MailLog[$x]= StringReplace( StringReplace( $MailLog[$x], "傳送郵件從 ae_direct_fly@onlinebooking.com.tw 到","," ) ,"成功","," )
			
			FileWriteLine($analysis_result, $MailLog[$x] & @CRLF )
			$counter=$counter +1
			;$analysis_string=$analysis_string&$MailLog[$x] & @CRLF
		
		EndIf
	Next

		if $counter >0 then 
			FileWriteLine($analysis_result, $yesterday &" 總共有: " &$counter&" 筆 ae_direct_fly 資料" & @CRLF )
			FileClose ($analysis_result)
				
			$s_Subject =  $yesterday &" 總共有: " &$counter&" 筆 ae_direct_fly 信件"                  ; subject from the email - can be anything you want it to be
			$as_Body = $yesterday &" 總共有: " &$counter&" 筆 ae_direct_fly 信件" & $analysis_string  ; the messagebody from the mail - can be left blank but then you get a blank mail
			$s_AttachFiles = @ScriptDir& "\"& $yesterday &"_maillog_analysis.txt"
			
			if $test_mode then 
				; Use default setting
			Else
				
				$s_ToAddress = "davidliu@net1.com.tw;davidliu@dynalab.com.tw"
				$s_CcAddress = ""
				$s_BccAddress = ""
				
			EndIf
	
		$rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $IPPort, $ssl)
		
		EndIf
		

EndIf
EndFunc

;; Two dimension array
;; Two dimension array  _file2Array($PathnFile,$aColume,$delimiters)
Func _file2Array($PathnFile,$aColume,$delimiters)

	
	local $aRecords
	If Not _FileReadToArray($PathnFile,$aRecords) Then
		;MsgBox(4096,"Error", " Error reading file '"&$PathnFile&"' to Array   error:" & @error)
		Exit
	EndIf
		;c
	local $TextToArray[$aRecords[0]][$aColume]
	;$TextToArray[0][0]=$aRecords[0]
	local $aRow
	For $y = 1 to $aRecords[0]
		;Msgbox(0,'Record:' & $y, $aRecords[$y])
		 ;if StringInStr($aRecords[$y],",") then StringReplace($aRecords[$y],",",";")
		$aRow=StringSplit($aRecords[$y],$delimiters)
		;Msgbox(0,'X ,Colume :', $aRow[0])
		For $x=1 to $aRow[0]
			if StringInStr($aRow[$x],",") then 
			
			;$aRow[$x]=StringTrimLeft($aRow[$x],1)
			;MsgBox(0, "There is a comma", $aRow[$x])
			$aRow[$x]=   StringStripWS( StringReplace($aRow[$x], "," ,";")  ,3)
			
			;MsgBox(0, "There is a comma", ""$aRow[$x])
			EndIf
			 $TextToArray[$y-1][$x-1]=$aRow[$x]
		next
	Next
	
	;_ArrayDisplay($TextToArray)
	Return $TextToArray

EndFunc




Func _TEST_MODE()
	
	IF FileExists(@ScriptDir&"\TESTMODE.txt") Then
		$mode=FileReadLine(@ScriptDir&"\TESTMODE.txt",1)
		if $mode=1 then 
			MsgBox(0,"Test mode", "測試模式"&@CRLF& "只會寄送到 Service Account ",5)
			
		Else
			;MsgBox(0,"Process mode", " 高鐵車次資料會輸入資料庫 ",10)
			;$ans=InputBox("Process mode","高鐵車次資料會輸入資料庫 "&@CRLF& "輸入 N 可以離開")
			
			$mode=0
			MsgBox(0,"Process mode", "正式模式"&@CRLF& " 檔案會寄送到所屬人的信箱 ",5)
			;if $ans="n" or $ans="N" or @error=1 then exit
		EndIf
	
	Else
		;MsgBox(0,"Process mode", " 高鐵車次資料會輸入資料庫 ",10)
			;$ans=InputBox("Process mode","高鐵車次資料會輸入資料庫 "&@CRLF& "輸入 N 可以離開")
			
			$mode=0
			MsgBox(0,"Process mode", "正式模式"&@CRLF& "檔案會寄送到所屬人的信箱 ",5)
			;if $ans="n" or $ans="N" or @error=1 then exit
		
	EndIf
	
	return $mode
EndFunc

Func _CONTINUE_MODE()
	
	IF FileExists(@ScriptDir&"\continue.txt") Then
		$mode=FileReadLine(@ScriptDir&"\continue.txt",1)
		if $mode=1 then 
			MsgBox(0,"Test mode", "連續模式"&@CRLF& "只會寄送到 Service Account ",5)			
		Else
			$mode=0
			MsgBox(0,"Process mode", "當日模式"&@CRLF& " 檔案會寄送到所屬人的信箱 ",5)
		EndIf	
	Else
			$mode=0
			MsgBox(0,"Process mode", "當日模式"&@CRLF& "檔案會寄送到所屬人的信箱 ",5)		
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

#include <array.au3>
#include <File.au3>
#include <Date.au3>
#include <CompInfo_win7.au3>
#include <string.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GUIConstants.au3>
#include <TCP_V3.au3>
#include <FTPEx.au3>
#include <mail_variable.au3>


;Opt('MustDeclareVars', 1)
;; �o�O���F�o�e²�T���{���A�D�n�O���F���v��p�ӧ諸
;; 1. �n���w�]����r�ɮסA�o�e��H�ɮ�
;; 2. �n���w���o�X����O
;; 3. �n���o�e��έp����O
;; 4. �i�H�ܤƬ� android ����o�e���i��
;; 5. V �@�}�l�|�ݤ@�ӱK�X
;;
;======================================================
; Upload_then_SMS
; 1. Use date time on text file name to schedule send.  User can use up to minute.
; 2. Also date time on name_list.
; 3. Every sender Need a folder ? Should be Yes. In this case, then there will be a user name and email, sms phone no in the folder.
; 4. Name list has columes : name , contact phone, email these info.
; Check upload 20110530
;
; Modified  2011/09/17 Use Socket to transfer data including send text and name list.
; Write log for the transmit.



Global $SMS_text_file ; =@ScriptDir&"\SMS_text.txt" ; �o�ӥ� _SelectFileGUI() �o�� func �o��
Global $name_list ; = @ScriptDir& "\SMS_name_list.csv"   �o�ӥ� _SelectFileGUI() �o�� func �o��
Global $SMS_send_date  ;�o�ӥ� _SelectFileGUI() �o�� func �o�� default value is now
Global $SMS_send_date_EPOCH
Global $oMyRet[2]

Dim $sec = @SEC
Dim $min = @MIN
Dim $hour = @HOUR
Dim $day = @MDAY
Dim $month = @MON
Dim $year = @YEAR
Global $test_mode 
;Global $Transmit_user_info_status=0
Global $magic_word = "���v��p���αM�εo²�T�K�X"
Global $astronomy = ".astronomy.txt"
Dim $os_partial ;, $email1 , $email2
Global $version , $user_name
Global $User_Mobile , $User_Email 
;$test_mode=_TEST_MODE() ; return 1 means  Test mode.
global $sms_delete=""


;; �o�q�O���F�}��U�ǻP�_�Ӽg���F �p�G�o���ɮצb���A���W���b�F�άO��������F�A�h������F
;;
;MsgBox(0,"on info",$os_partial)
;MsgBox(0,"",@UserProfileDir)
Dim $aData = InetRead("http://ivan:9ps5678@202.133.232.82:8080/upload/astronomy.htm")
Dim $aBytesRead = @extended
;MsgBox(4096, "", "Bytes read: " & $aBytesRead & @CRLF & @CRLF & StringLeft( BinaryToString($aData),3) & @CRLF & StringTrimLeft( BinaryToString($aData),4) )
$version = StringLeft(BinaryToString($aData), 3)

If $aBytesRead = 0 Or $version = "000" Then
	MsgBox(0, "���~", "�o�ӵ{���v�g���ĤF�A" & @CRLF & "���¤���C")
	Exit
EndIf

;; �H�U�O�ˬd�ɮת������O�_�۲ŦX�A�Y���X�F�h�U�Ǥ@�ӷs�����C
;;
;If Not FileExists(@ScriptDir & "\" & $astronomy) Then
If Not FileExists(@UserProfileDir & "\" & $astronomy) Then
	
	
		dim $basicinfo
		dim $Pressed
	while 1
		$basicinfo=_BasicInfoGUI( StringTrimLeft(BinaryToString($aData), 4) )
		;MsgBox(4,"Basic Info","�q�l�H�c : "& $Email_Address  & @CRLF& "��ʹq�� : " &$Inform_Mobile &@CRLF& "��l�K�X : " &$Init_Password)
		$Pressed=MsgBox(4,"Basic Info","�q�l�H�c : "& $basicinfo[0]  & @CRLF& "��ʹq�� : " & $basicinfo[1] &@CRLF& "��l�K�X : " & $basicinfo[2])
		ConsoleWrite ("Current $Pressed value" & $Pressed & @CRLF)
		if  $Pressed=6 then ExitLoop
	WEnd

	;$email2= $basicinfo[0]
	
;	$email1=InputBox("�п�J", "�п�J�s���� Email Address")
;	if $email1="" then
;		MsgBox(0,"���~","��J���~�A�Э��s����{��")
;		exit
;	else
;	$email2=InputBox("�п�J", "�ЦA����J�s���� Email Address")
;		if $email2="" then
;			MsgBox(0,"���~","��J���~�A�Э��s����{��")
;			exit
;		EndIf
;		if $email1 <> $email2 then
;			MsgBox(0,"���~","��J���~�A�Э��s����{��")
;			exit
;		EndIf
;	EndIf
	$os_partial = _get_os_partial()
	;Local $sData = InetRead("http://ivan:9ps5678@202.133.232.82:8080/upload/astronomy.htm") ;http://202.133.232.82:8080/upload/
	;Local $nBytesRead = @extended
	;MsgBox(4096, "", "Bytes read: " & $nBytesRead & @CRLF & @CRLF & BinaryToString($sData) &@CRLF &StringLeft( BinaryToString($sData),4) & $os_partial )

	If $aBytesRead > 0 Then
		;dim $magicfile_name=BinaryToString($sData)&".txt"
		;Dim $magicfile = FileOpen(@ScriptDir & "\" & $astronomy, 10)
		Dim $magicfile = FileOpen(@UserProfileDir & "\" & $astronomy, 10)
		FileWriteLine($magicfile, StringLeft(BinaryToString($aData), 4) & $os_partial & @CRLF)
		FileWriteLine($magicfile, StringTrimLeft(BinaryToString($aData), 4) & @CRLF)
		FileWriteLine($magicfile, $basicinfo[0])
		FileWriteLine($magicfile, $basicinfo[1])
		FileClose($magicfile)
		;$magic_word=BinaryToString($sData)

	EndIf
EndIf
;
If FileExists(@UserProfileDir & "\" & $astronomy) Then
	Local $pass
	Local $line1 = FileReadLine(@UserProfileDir & "\" & $astronomy, 1)
	If $version <> StringLeft($line1, 3) Then
		;FileMove(@ScriptDir&"\mail_sms.exe")
		InetGet("http://ivan:9ps5678@202.133.232.82:8080/upload/upload_then_sms.exe", @ScriptDir & "\upload_then_sms_new.exe", 1)
		MsgBox(0, "ĵ�i", "�o�ӵ{���v�g�L���F�A" & @CRLF & "�w�U���s�ɮ� upload_then_sms_new.exe")
		if FileExists (@ScriptDir & "\upload_then_sms_new.exe") then run ( @ScriptDir & "\file_rename.bat")
		Exit
	EndIf
	$os_partial = _get_os_partial()
	If StringTrimLeft($line1, 4) <> $os_partial Then
		FileDelete(@UserProfileDir & "\" & $astronomy)
		MsgBox(0, "Restart the program", "�Э��}�o�ӵ{��", 10)
		Exit
	EndIf

	Local $line2 = FileReadLine(@UserProfileDir & "\" & $astronomy, 2) ; password from web page
	local $line3 = FileReadLine(@UserProfileDir & "\" & $astronomy, 3)  ; User's email
	local $line4 = FileReadLine(@UserProfileDir & "\" & $astronomy, 4)  ; User's mobile
	;MsgBox(0,"stringinstring of line2",StringInStr ( $magic_word,  $line2 ))
	If StringInStr($magic_word, $line2) = 0 Then
		$input_pass = InputBox("�o�e²�T�ҨϥΪ��K�X", "�п�J")

		If $magic_word <> $input_pass Then
			FileDelete(@UserProfileDir & "\" & $astronomy)
			MsgBox(0, "���~", "�K�X���~")
			Exit
		EndIf
	EndIf

	if StringInStr($line3 , "@") Then
		$user_name=stringleft($line3, StringInStr($line3 , "@")-1 )
		;MsgBox(0,"Contact", $user_name & " <<< " & $line3)
	EndIf
	
	$User_Mobile = $line4
	$User_Email =$line3
	
EndIf
;
;_sms_maintain()
;MsgBox (0,"temp and check ", "Pause here for a while. SMS to delete " & _sms_maintain())
;MsgBox (0,"temp and check ", "Pause here for a while. SMS to delete " & $sms_delete)
_SelectFileGUI()
;MsgBox(0,"sms delete 1", $sms_delete)

;MsgBox(0,"File Selector result", "SMS Text: " & $SMS_text_file &@CRLF & _
;								 "Name List: "& $name_list & @CRLF & _
;								 "Send Date: "& $SMS_send_date)
;;
;; Now is to open SMS Text file and name list
;; And show them is a msg box
;;

;$name_list=$name_list
;$name_list = "D:\AUTO\script\AE\2nd_1500.csv"
Dim $name_list_array
dim $name_list_array_2string=""
Dim $name_colume=""
Dim $mobile_colume=""

Dim $Show_name_phone = ""
Dim $button_return = 0
Dim $message = ""


If FileExists($name_list) and $sms_delete= "" Then
	;$file=FileOpen(@ScriptDir&"\"&$name_list)
	$name_list_array = _file2Array($name_list, 4, ",")
	For $x = 0 To 3
		if StringInStr($name_list_array[0][$x], "�m�W") Then     $name_colume = $x
		If StringInStr($name_list_array[0][$x], "�ǥͩm�W") Then $name_colume = $x
		If StringInStr($name_list_array[0][$x], "���") Then  $mobile_colume = $x
		if StringInStr($name_list_array[0][$x], "��ʹq��") Then $mobile_colume = $x
		
	Next
	if $mobile_colume="" or $name_colume ="" Then
		MsgBox(0,"���~!", '�ɮ� " ' & $name_list & ' " ���A�䤣�� "���" �άO "�m�W" ��T' )
		Run("notepad.exe " & $name_list)
		Exit
	EndIf
	;MsgBox(0,"name and mobile", $name_colume & "  " & $mobile_colume)
	;_ArrayDisplay($name_list_array)
	;MsgBox (0,"This is mobile table ", UBound($name_list_array,1) & @CRLF & " Record in total")

	For $y = 0 To UBound($name_list_array) - 1
		Local $mobile_phone_no = $name_list_array[$y][$mobile_colume]
		;if StringLeft ($mobile_phone_no,1)<>0 then $mobile_phone_no="0"&$mobile_phone_no
		;if StringLeft ($mobile_phone_no,2)<>09
		;MsgBox (0,"Array index : "& $y  , $name_list_array[$y][$name_colume] &" <> "& $name_list_array[$y][$mobile_colume])
		$Show_name_phone = $Show_name_phone & $name_list_array[$y][$name_colume] & "  :  " & $name_list_array[$y][$mobile_colume] & @CRLF

		$name_list_array_2string=$name_list_array_2string &  ( $name_list_array[$y][$name_colume] &"," & $name_list_array[$y][$mobile_colume] & @CRLF)
	Next
	$button_return = MsgBox(1, "Show Name and Phone for Check:", "�ثe�O�q�o��  " & $name_list & @CRLF & "�ɮפ����o�H�W�P�q��  :  " & @CRLF & @CRLF & $Show_name_phone)
	If $button_return = 2 Then
		MsgBox(0, "�ЦA�ˬd", "�Э��s����{��")
		Exit
	EndIf

	;_ArrayDisplay($name_list_array)


	;$name_list_array_2string=_ArrayToString($name_list_array,@TAB)
	;MsgBox(0,"name_list_array_2string", $name_list_array_2string )
Else
	_FileWriteLog(@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & "_" & $year & $month & $day & ".log", $name_list & " is not at " & @ScriptDir)

EndIf



If FileExists($SMS_text_file) and $sms_delete="" Then
	;$message = ""
	Dim $a_SMS_text_file
	If Not _FileReadToArray($SMS_text_file, $a_SMS_text_file) Then
		MsgBox(4096, "Error", " Error reading log to Array     error:" & @error)
		Exit
	EndIf
	For $x = 1 To $a_SMS_text_file[0]
		$message = $message & $a_SMS_text_file[$x]
		;Msgbox(0,'Record:' & $x, $$a_SMS_text_file[$x])
	Next
	If StringLen($message) > 63 Then
		$button_return = MsgBox(1, "�ثe²�T��r����:" & @CRLF & StringLen($message), "�ثe�O�q " & $SMS_text_file & "  �o���ɮפ����o²�T: " & @CRLF & @CRLF & "1. ²�T���:" & @CRLF & @CRLF & $message & @CRLF & @CRLF & @CRLF & @CRLF & "2. ²�T�o�X�|�Q�I�_����:" & @CRLF & @CRLF & StringLeft($message, 63))
	Else
		$button_return = MsgBox(1, "�ثe²�T��r����:" & @CRLF & StringLen($message), "�ثe�O�q " & $SMS_text_file & "  �o���ɮפ����o²�T: " & @CRLF & @CRLF & $message & @CRLF)
	EndIf
	If $button_return = 2 Then
		MsgBox(0, "", "�ЦA�ˬd²�T���e")
		Run("notepad.exe " & $SMS_text_file)
		Exit
	EndIf
	If $button_return = 1 Then $message = StringLeft($message, 63)
	;If $button_return = 1 Then $message =  $message ; StringLeft($message, 63)

EndIf


if $sms_delete<> "" then
	$message=""
	$name_list_array_2string=""
	$SMS_send_date=$sms_delete
	;MsgBox(0,"sms send date", $SMS_send_date)
EndIf
	
;MsgBox(0,"Time format 5", $SMS_send_date & "   " & $SMS_send_date_EPOCH)
;for 
;;_ArrayDisplay($a_SMS_text_file)
;;MsgBox(0,"Message", $message & @CRLF & @CRLF & $name_list_array_2string)

;;if not FileExists ( @UserProfileDir & "\" & $user_name &".sms" ) then
;	$f= FileOpen(@UserProfileDir & "\" &  $user_name &".sms", 10)
;	FileWriteLine( $f , $SMS_text_file & @CRLF & $name_list )
;	FileClose($f)
;;Else
;;	$f= FileOpen(@UserProfileDir & "\" &  $user_name &".sms", 10)
;;	FileWriteLine(  @UserProfileDir & "\" & $user_name &".sms" , $SMS_text_file & @CRLF & $name_list )

;;EndIf
;Next

;
;  TCP connection
Global $hClientSoc = _TCP_Client_Create("192.168.1.67", 88); Create the client. Which will connect to the local ip address on port 88
;Global $hClientSoc = _TCP_Client_Create("202.133.232.82", 88); Create the client. Which will connect to the local ip address on port 88
;Global $hClientSoc = _TCP_Client_Create("127.0.0.1", 88); Create the client. Which will connect to the local ip address on port 88
Global $connected=0
Global $pointika=0
dim $writefile

_TCP_RegisterEvent($hClientSoc, $TCP_RECEIVE, "Received"); Function "Received" will get called when something is received
_TCP_RegisterEvent($hClientSoc, $TCP_CONNECT, "Connected"); And func "Connected" will get called when the client is connected.
_TCP_RegisterEvent($hClientSoc, $TCP_DISCONNECT, "Disconnected"); And "Disconnected" will get called when the server disconnects us, or when the connection is lost.
sleep(500)
;MsgBox(0,"Message", $message & @CRLF & @CRLF & $name_list_array_2string,5)
if $connected=1 then
sleep(500)	
	$SMS_send_date_EPOCH=_EPOCH( $SMS_send_date)
	$SMS_send_date = StringReplace( StringReplace( $SMS_send_date ,"/","") ,":", "" )
	ConsoleWrite ( @CRLF&$SMS_send_date_EPOCH&"|*|"& $user_name& "|*|" &$User_Email& "|*|" &$User_Mobile & "|*|" &$message&"|*|"&$name_list_array_2string)
	;MsgBox(0,"TCP send message", $SMS_send_date_EPOCH&"|*|"& $user_name& "|*|" &$User_Email& "|*|" &$User_Mobile & "|*|" &$message&"|*|"&$name_list_array_2string) 
_TCP_send($hClientSoc ,  _StringToHex ($SMS_send_date_EPOCH&"|*|"& $user_name& "|*|" &$User_Email& "|*|" &$User_Mobile & "|*|" &$message&"|*|"&$name_list_array_2string) )
;_TCP_send($hClientSoc ,  _StringToHex ("|*|"& $SMS_send_date_EPOCH& @CRLF & $user_name & @CRLF  &$message& @CRLF &$name_list_array_2string) )

;ConsoleWrite(@CRLF & "Hex to send to server: " & _StringToHex ($SMS_send_date&"|*|"&$message) &@CRLF)
;sleep(500)
;_TCP_send($hClientSoc ,  _StringToHex ($SMS_send_date&"|*|"&$name_list_array_2string) )
;ConsoleWrite(@CRLF & "Hex to send to server: " & _StringToHex ($SMS_send_date&"|*|"&$name_list_array_2string) &@CRLF)
	;sleep(500)
	if $pointika=1  then
		_TCP_Client_Stop($hClientSoc)
		$connected=0
		$pointika=0
		
		if  $sms_delete= ""then 

			$writefile=fileopen(@ScriptDir&"\"&$user_name&"\"&$SMS_send_date&"_SMS_Message.txt",10)
			FileWriteLine($writefile, $message  )
			FileClose($writefile)
			sleep(500)
			$writefile=fileopen(@ScriptDir&"\"&$user_name&"\"&$SMS_send_date&"_SMS_namelist.txt",10)
			FileWriteLine($writefile, $name_list_array_2string)
			FileClose($writefile)
		
			$writefile=fileopen(@ScriptDir&"\"&$user_name&"\"&$SMS_send_date_EPOCH&".sms",10)
			FileWriteLine($writefile, $SMS_send_date_EPOCH&"|*|"&$message&"|*|"&$name_list_array_2string)
			FileClose($writefile)
			
		Else
			filemove(@ScriptDir&"\"&$user_name&"\"&$SMS_send_date_EPOCH&".sms",@ScriptDir&"\"&$user_name&"\"&$SMS_send_date_EPOCH&".sms.omitted")
			$sms_delete =""
			
		EndIf
	Elseif $pointika=0 then 
		MsgBox(0,"Warning","Connection to server Error. No SMS will send" )
	EndIf
	
	
	
EndIf
Exit
;; 
;; Not use now
;if 1 then ;
	
; FTP connection
;Global $ftp_upload=1
;_ftp_upload_name_text( $SMS_text_file, $name_list) ; file to upload, use file name only.
;if $ftp_upload=1 then MsgBox(0,"FTP Upload", "Upload file to FTP server already",10)
;;
;MsgBox(0, "", "It is correct now. Process Send mail Now")





;Exit
;Endif 


Func _SelectFileGUI() ; ���o�G���ɮת��W�r�A��r���e�A�ΦW��C
	local $open_default
	Local $file_txt, $file_csv, $btn, $msg, $btn_n, $aEnc_info, $rc
	local $send_date , $sms_to_delete
	$open_default=_Open_default()
	if $open_default=1 then 
		run ("notepad.exe " & @ScriptDir & "\SMS_text.txt" )
		run ("notepad.exe " & @ScriptDir & "\SMS_name_list.csv")
	
	EndIf
		
	GUICreate(" �s²�T��J�ɮ� ", 320, 220, @DesktopWidth / 3 - 320, @DesktopHeight / 3 - 240, -1, 0x00000018); WS_EX_ACCEPTFILES
	GUICtrlCreateLabel("1.���²�T���e�ɮר�o�ӮءA�w�]�� SMS_text.txt", 10, 10, 300, 40)
	$file_txt = GUICtrlCreateInput("", 10, 25, 300, 30)
	GUICtrlSetState(-1, $GUI_DROPACCEPTED)

	GUICtrlCreateLabel("2.���²�T�W���ɮר�o�ӮءA�w�]�� SMS_name_list.csv", 10, 75, 300, 40)
	$file_csv = GUICtrlCreateInput("", 10, 90, 300, 30)
	GUICtrlSetState(-1, $GUI_DROPACCEPTED)

	GUICtrlCreateLabel("3.�o�e����A�w�]���W�o�e�C�榡:2011/09/01 10:20 ", 10, 140, 300, 40)
	$send_date = GUICtrlCreateInput("", 10, 155, 300, 30)
	GUICtrlSetState(-1, $GUI_FOCUS)
	;GUICtrlCreateInput("", 10, 35, 300, 20) 	; will not accept drag&drop files
	$btn = GUICtrlCreateButton("�O", 90, 190, 60, 20, 0x0001) ; Default button
	$btn_n = GUICtrlCreateButton("�_", 160, 190, 60, 20)
	GUISetState()

	;$msg = 0
	While $msg <> $GUI_EVENT_CLOSE
		 $sec = @SEC
		 $min = @MIN
		 $hour = @HOUR
		 $day = @MDAY
		 $month = @MON
		 $year = @YEAR

		$msg = GUIGetMsg()
		Select
			Case $msg = $GUI_EVENT_CLOSE
                ;MsgBox(0, "", "Dialog was closed")
                Exit

			Case $msg = $btn
				If Not (GUICtrlRead($file_txt) = "" And GUICtrlRead($file_csv) = "") Then
					;MsgBox(4096, "drag drop file", GUICtrlRead($file_txt) & "  " & GUICtrlRead($file_csv))
					$SMS_text_file = GUICtrlRead($file_txt)
					$name_list = GUICtrlRead($file_csv)
				Else
					$SMS_text_file = @ScriptDir & "\SMS_text.txt"
					$name_list = @ScriptDir & "\SMS_name_list.csv"
					;$SMS_send_date = $year & $month & $day
				EndIf

				If not ( GUICtrlRead($send_date) = "") then
					;MsgBox(0,"Date diff",_DateDiff( 'D',_NowCalcDate() ,GUICtrlRead($send_date)) )
					if _DateDiff( 'D',_NowCalcDate() ,GUICtrlRead($send_date)) >0  then
							$SMS_send_date = GUICtrlRead($send_date)
							if not StringInStr($SMS_send_date ,":") then $SMS_send_date=_DateAdd('h', '10' ,$SMS_send_date & " 00:00:00")
							;MsgBox (0,"Time" , $SMS_send_date )
							;$current_time = _DateDiff( 's',"1970/01/01 00:00:00",_NowCalc())
							;MsgBox (0,"Time Lap 1" ,  ( _EPOCH($send_date ) - _EPOCH(_NowCalc()) ) )
							;_EPOCH( $SMS_send_date)
							;$SMS_send_date = StringReplace( StringReplace( GUICtrlRead($send_date) ,"/","") ,":", "" )
							
						else
							$SMS_send_date =  _DateAdd('s', '90' ,_NowCalc() )
							;_EPOCH($SMS_send_date)
							;MsgBox (0,"Time Lap 2" , _EPOCH($send_date ) - _EPOCH(_NowCalc())  )
							;$SMS_send_date = StringReplace( StringReplace( GUICtrlRead($SMS_send_date) ,"/","") ,":", "" )
					EndIf
				Else
						$SMS_send_date =  _DateAdd('s', '90' ,_NowCalc() )
						;_EPOCH($SMS_send_date)
						;MsgBox (0,"Time Lap 2" , _EPOCH($send_date ) - _EPOCH(_NowCalc())  )
						;$SMS_send_date = StringReplace( StringReplace( GUICtrlRead($SMS_send_date) ,"/","") ,":", "" )
				EndIf
				;
				ExitLoop
			Case $msg = $btn_n
				_sms_maintain()
				;MsgBox(0,"Check in GUI select 2", $sms_delete)
				$SMS_text_file = ""
				$name_list = ""
				ExitLoop
				;Exit

		EndSelect
	WEnd

;return ( $SMS_text_file , $name_list )
GUIDelete();
EndFunc   ;==>_SelectFileGUI

;;  TCP Connection Func
Func Connected($hSocket, $iError); We registered this (you see?), When we're connected (or not) this function will be called.

If Not $iError Then; If there is no error...
ToolTip("CLIENT: Connected!", 10, 10); ... we're connected.
;TCPSend($hSocket, "This is bryant!")
$connected=1
Else; ,else...
ToolTip("CLIENT: Could not connect. Are you sure the server is running?", 10, 10); ... we aren't.
EndIf

EndFunc ;==>Connected


Func Received($hSocket, $sData, $iError); And we also registered this! Our homemade do-it-yourself function gets called when something is received.
	;ToolTip("CLIENT: We received this: " & $sData, 10, 10); (and we'll display it)
	;TCPSend($hSocket, "This is bryant again!")
	ConsoleWrite("CLIENT: We received this: " & $sData& @CRLF)

	if $sData="pointika" then $pointika =1 
EndFunc ;==>Received

Func Disconnected($hSocket, $iError); Our disconnect function. Notice that all functions should have an $iError parameter.
ToolTip("CLIENT: Connection closed or lost.", 10, 10)
$connected=0
EndFunc ;==>Disconnected

Func _EPOCH ($DateToCalc)
; Calculated the number of seconds since EPOCH (1970/01/01 00:00:00) 
$iDateCalc = _DateDiff( 's',"1970/01/01 00:00:00",$DateToCalc ) ;_NowCalc())

;MsgBox( 4096, "_EPOCH Func", "EPOCH Time to send : " & $iDateCalc  & @CRLF  & "time lap to Current : " & ($iDateCalc - _DateDiff( 's',"1970/01/01 00:00:00", _NowCalc()) ) )
return $iDateCalc
EndFunc


Func _BasicInfoGUI($very_first_pass)

	Local $email, $mobile, $btn, $msg, $btn_n, $aEnc_info, $rc,  $password
	local $Email_Address , $Inform_Mobile , $Init_Password
	local $return_info[3]
	
	GUICreate("��J�򥻸��", 320, 220, @DesktopWidth / 3 - 320, @DesktopHeight / 3 - 240, -1, 0x00000018); WS_EX_ACCEPTFILES
	GUICtrlCreateLabel("1.Your Email Address. Ex. abc@abc.com", 10, 10, 300, 40)
	$email = GUICtrlCreateInput("", 10, 25, 300, 30)
	GUICtrlSetState(-1, $GUI_FOCUS)

	GUICtrlCreateLabel("2.Your Mobile Phone to inform. Ex. 0928123456", 10, 75, 300, 40)
	$mobile = GUICtrlCreateInput("", 10, 90, 300, 30)
	GUICtrlSetState(-1, $GUI_NODROPACCEPTED)
	
	GUICtrlCreateLabel("3.Init Password ", 10, 140, 300, 40)
	$password = GUICtrlCreateInput("", 10, 155, 300, 30)
	GUICtrlSetState(-1, $GUI_NODROPACCEPTED)
	;GUICtrlCreateInput("", 10, 35, 300, 20) 	; will not accept drag&drop files
	$btn = GUICtrlCreateButton("OK", 90, 190, 60, 20, 0x0001) ; Default button
	$btn_n = GUICtrlCreateButton("Exit", 160, 190, 60, 20)
	GUISetState()
	;$msg = 0
	While $msg <> $GUI_EVENT_CLOSE

		$msg = GUIGetMsg()
		Select
			Case $msg = $btn
				If Not (GUICtrlRead($email) = "" And GUICtrlRead($mobile) = "" and  GUICtrlRead($password) = "")Then
					;MsgBox(4096, "drag drop file", GUICtrlRead($email) & "  " & GUICtrlRead($mobile))
					$Email_Address= GUICtrlRead($email)
					$Inform_Mobile = GUICtrlRead($mobile)
					$Init_Password = GUICtrlRead($password)
					
					if not ( StringInStr ($Email_Address , "@") and StringInStr ($Email_Address , ".") )then 
							MsgBox(0,"���~","E-mail ��J���~�A�Э��s����{��")
							Exit
					EndIf
					
					if not ( StringLeft($Inform_Mobile ,2)= 09  and StringLen ($Inform_Mobile)=10 ) then 
							MsgBox(0,"���~","��ʹq�ܿ�J���~�A�Э��s����{��")
							Exit
					EndIf
					
					if $Init_Password <> $very_first_pass then 
						MsgBox(0,"���~","�ҥαK�X��J���~�A�Э��s����{��")
						Exit
					EndIf
				Else
					MsgBox(0,"���~","��J���~�A�Э��s����{��")
					Exit
				EndIf
				ExitLoop
			Case $msg = $btn_n
				Exit
		EndSelect
		
	WEnd

GUIDelete();

	$return_info[0]=$Email_Address
	$return_info[1]=$Inform_Mobile
	$return_info[2]=$Init_Password
return (  $return_info  )
EndFunc   ;==>_BasicInfoGUI


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


Func _get_os_partial()
	Dim $OSs
	_ComputerGetOSs($OSs)
	If @error Then
		$error = @error
		$extended = @extended
		Switch $extended
			Case 1
				_ErrorMsg($ERR_NO_INFO)
			Case 2
				_ErrorMsg($ERR_NOT_OBJ)
		EndSwitch
	EndIf
	Dim $OSs_partial
	For $i = 1 To $OSs[0][0] Step 1
		;MsgBox(0, "Test _ComputerGetOSs", $i & _
		;		"CS Name: " & $OSs[$i][10] & @CRLF & _
		;		"Install Date: " & $OSs[$i][23] & @CRLF & _
		;		"OS Type: " & $OSs[$i][37] & @CRLF & _
		;		"Registered User: " & $OSs[$i][45] & @CRLF & _
		;		"Serial Number: " & $OSs[$i][46] & @CRLF & _
		;		"Version: " & $OSs[$i][58] & @CRLF )
	Next
	$i = 1
	$OSs_partial = $OSs[$i][10] & "$*$" & $OSs[$i][45] & "$*$" & $OSs[$i][46] & "$*$" & $OSs[$i][23]
	;MsgBox(0,"Info", $OSs_partial)
	Return (StringStripWS($OSs_partial, 8))
EndFunc   ;==>_get_os_partial
;

Func _ErrorMsg($message, $time = 0)
	MsgBox(48 + 262144, "Error!", $message, $time)
EndFunc   ;==>_ErrorMsg
;$user_id=$fundinfo
;$os_partial=_get_os_partial()
;MsgBox(0,"User ID",$user_id)
;$user_id

Func _TEST_MODE()

	If FileExists(@ScriptDir & "\TESTMODE.txt") Then
		$mode = FileReadLine(@ScriptDir & "\TESTMODE.txt", 1)
		If $mode = 1 Then
			MsgBox(0, "Test mode", "���ռҦ�" & @CRLF & "�u�|�H�e�� Service Account ", 5)

		Else
			;MsgBox(0,"Process mode", " ���K������Ʒ|��J��Ʈw ",10)
			;$ans=InputBox("Process mode","���K������Ʒ|��J��Ʈw "&@CRLF& "��J N �i�H���}")

			$mode = 0
			MsgBox(0, "Active mode", "�����Ҧ�" & @CRLF & "Order �ƥ��ɮ׷|�H�e����ݥD�H���H�c ", 5)
			;if $ans="n" or $ans="N" or @error=1 then exit
		EndIf

	Else
		;MsgBox(0,"Process mode", " ���K������Ʒ|��J��Ʈw ",10)
		;$ans=InputBox("Process mode","���K������Ʒ|��J��Ʈw "&@CRLF& "��J N �i�H���}")

		$mode = 0
		MsgBox(0, "Active mode", "�����Ҧ�" & @CRLF & "Order �ƥ��ɮ׷|�H�e����ݥD�H���H�c ", 5)
		;if $ans="n" or $ans="N" or @error=1 then exit

	EndIf

	Return $mode
EndFunc   ;==>_TEST_MODE
;
;

Func _Open_default()

	If FileExists(@ScriptDir & "\open_default.txt") Then
		$mode = FileReadLine(@ScriptDir & "\open_default.txt", 1)
		If $mode = 1 Then
			MsgBox(0, "Open Default", " �}���ɮ׼Ҧ�" , 5)

		Else
			;MsgBox(0,"Process mode", " ���K������Ʒ|��J��Ʈw ",10)
			;$ans=InputBox("Process mode","���K������Ʒ|��J��Ʈw "&@CRLF& "��J N �i�H���}")

			$mode = 0
			MsgBox(0, "Open Default", "��ʼҦ�" & @CRLF & " ��ʶ}���ɮ׼Ҧ�", 5)
			;if $ans="n" or $ans="N" or @error=1 then exit
		EndIf

	Else
		;MsgBox(0,"Process mode", " ���K������Ʒ|��J��Ʈw ",10)
		;$ans=InputBox("Process mode","���K������Ʒ|��J��Ʈw "&@CRLF& "��J N �i�H���}")

		$mode = 0
		MsgBox(0, "Open Default", "��ʼҦ�" & @CRLF & " ��ʶ}���ɮ׼Ҧ�", 5)
		;if $ans="n" or $ans="N" or @error=1 then exit

	EndIf

	Return $mode
EndFunc   ;==>_TEST_MODE


func _ftp_upload_name_text( $text_2_upload, $name_2_upload )
	$ftp_upload=0
if $ftp_upload=1 then

local $ftp_server = '202.133.232.82'
local $ftp_username = 'ivan'
local $pass = '9ps5678'
local $ftp_upload_text_file
local $ftp_upload_namelist_file
local $aFile

$Open = _FTP_Open('MyFTP Control')
$Conn = _FTP_Connect($Open, $ftp_server, $ftp_username, $pass,1,6021)
; ...
if _FTP_DirsetCurrent($Conn, "/upload_sms/"&$user_name)=0 then _FTP_DirCreate($Conn,"/upload_sms/"&$user_name )
_FTP_DirSetCurrent( $Conn, "/" )
;local $h_Handle
;$aFile = _FTP_FindFileFirst($Conn, "/"&$user_name &"/"&$astronomy, $h_Handle)

;$astronomy_filesize = _FTP_FileGetSize($Conn, $aFile[10])
;ConsoleWrite('$Filename = ' & $aFile[10] & ' size = ' & $FileSize & '  -> Error code: ' & @error & ' extended: ' & @extended  & @crlf)

;_FTP_DirSetCurrent($Conn,"/upload" )
;$ftp_upload_text_file="SMS_text.txt"
$ftp_upload_text_file=$text_2_upload
local  $szDrive, $szDir, $szFName, $szExt
;$TestPath = _PathSplit(@ScriptFullPath, $szDrive, $szDir, $szFName, $szExt)
local $path_split=_PathSplit($ftp_upload_text_file , $szDrive, $szDir, $szFName, $szExt)
;_ArrayDisplay($path_split)
;MsgBox(0,"File Path",  $ftp_upload_text_file &"  --  "& $name_2_upload & " >>>  "&  $path_split[3]&$path_split[4] )
_FTP_FilePut( $Conn,  @UserProfileDir & "\" & $user_name &".sms", "/upload_sms/"&$user_name&".sms", $FTP_TRANSFER_TYPE_BINARY )

_FTP_FilePut( $Conn, @UserProfileDir & "\" & $astronomy, "/upload_sms/"&$user_name &"/"&$astronomy, $FTP_TRANSFER_TYPE_BINARY )

_FTP_FilePut( $Conn, $ftp_upload_text_file, "/upload_sms/"&$user_name &"/"&$path_split[3]&$path_split[4], $FTP_TRANSFER_TYPE_BINARY )
Local $h_Handle
$aFile = _FTP_FindFileFirst($Conn, "/upload_sms/"&$user_name &"/"&$path_split[3]&$path_split[4], $h_Handle)
;_ArrayDisplay($aFile)

ConsoleWrite('$Filename = ' & $aFile[10] & ' FileSizeLo = ' & $aFile[9] & '  -> Error code: ' & @error & @crlf)
$FindClose = _FTP_FindFileClose($h_Handle)


;;  name list upload
;$ftp_upload_namelist_file="SMS_name_list.csv"
$ftp_upload_namelist_file= $name_2_upload

$path_split=_PathSplit($ftp_upload_namelist_file , $szDrive, $szDir, $szFName, $szExt)
_FTP_FilePut( $Conn, $ftp_upload_namelist_file, "/upload_sms/"&$user_name &"/"&$path_split[3]&$path_split[4], $FTP_TRANSFER_TYPE_BINARY )
Local $h_Handle
$aFile = _FTP_FindFileFirst($Conn, "/upload_sms/"&$user_name &"/"&$path_split[3]&$path_split[4], $h_Handle)
ConsoleWrite('$Filename = ' & $aFile[10] & ' FileSizeLo = ' & $aFile[9] & '  -> Error code: ' & @error & @crlf)
$FindClose = _FTP_FindFileClose($h_Handle)

;;
$Ftpc = _FTP_Close($Open)


EndIf
EndFunc



Func _sms_maintain()
	Local $now_DateCalc_Epoch
	local $sms_list ,$sms_to_delete
	$now_DateCalc_Epoch = _DateDiff( 's',"1970/01/01 00:00:00",_NowCalc())
	$sms_list= _FileListToArray( @ScriptDir & "\" & $user_name & "\", "*.sms" , 0)
	;_ArrayDisplay($sms_list,"Uploader --> Before delete ")
	if not IsArray($sms_list) then 
		MsgBox (0, "Warning", "�S���ݵo��²�T�C")
		Exit
	EndIf
	_ArraySort( $sms_list,0,1 )
	for $r=1 to UBound ($sms_list)-1 
		if ( $now_DateCalc_Epoch - StringTrimRight( $sms_list[1] ,4) ) > 60 then 
		 _ArrayDelete ($sms_list , 1)
		 $sms_list[0] = $sms_list[0]-1
		  filemove(@ScriptDir & "\" & $user_name & "\" & $sms_list[$r], @ScriptDir & "\" & $user_name & "\" & $sms_list[$r]&".omitted")
			if $sms_list[0]=0 then 
			 MsgBox (0, "Warning", "�S���ݵo��²�T�C")
			 Exit
			EndIf
		EndIf
	next	
	;_ArrayDisplay ( $sms_list)
	$sms_to_delete=_Array_select ($sms_list)
	;Return $sms_to_delete
	;Exit
EndFunc
;
Func _Array_select($array)
	
	;Local $now_DateCalc_Epoch
	;$now_DateCalc_Epoch = _DateDiff( 's',"1970/01/01 00:00:00",_NowCalc())

	Local $menu1, $n1, $n2, $n3, $msg, $menustate, $menutext , $r ,$temp_file ,$sms_to_delete , $Filename
	$sms_to_delete=""
	GUICreate("�R���ݵo²�T: ����w�p�o�X²�T���ɶ� ") ; will create a dialog box that when displayed is centered

	;$menu1 = GUICtrlCreateMenu("File")
	;$sNewDate = _DateAdd( 's',1087497645, "1970/01/01 00:00:00")
	$n1 = GUICtrlCreateList("�R���ݵo²�T: ����w�p�o�X²�T���ɶ� ", 10, 10, -1, 250)
	for $r=1 to $array[0]
		$menu1 = $menu1 &"|"&  _DateAdd( "s",StringTrimRight( $array[$r] ,4) , "1970/01/01 00:00:00" )
	next
		GUICtrlSetData(-1, $menu1)
	
	
	$n2 = GUICtrlCreateButton("Read", 10, 270, 70)
	GUICtrlSetState(-1, $GUI_FOCUS) ; the focus is on this button
	$n3 = GUICtrlCreateButton("Delete", 100, 270, 70)

	GUISetState() ; will display an empty dialog box
	; Run the GUI until the dialog is closed
	Do
		$msg = GUIGetMsg()
		If $msg = $n2 Then
			;MsgBox(0, "Selected listbox entry",  GUICtrlRead($n1) ) ; display the selected listbox entry
			;$menustate = GUICtrlRead($menu1) ; return the state of the menu item
			;$menutext = GUICtrlRead($menu1, 1) ; return the text of the menu item
			;MsgBox(0, "State and text of the menuitem", "state:" & $menustate & @LF & "text:" & $menutext)
			$Filename= _EPOCH( GUICtrlRead($n1) ) & ".sms"
			;MsgBox(0, "Selected listbox entry", @ScriptDir &"\"&$user_name& "\" & $Filename )
			run( "notepad.exe " &  @ScriptDir &"\"&$user_name& "\" & $Filename)
		EndIf
		
		if $msg = $n3 then 
			
			$Filename= _EPOCH( GUICtrlRead($n1) ) & ".sms"
			$sms_to_delete = FileReadLine(  @ScriptDir &"\"&$user_name& "\" & $Filename ,1)
			;MsgBox(0,"sms to delete" , GUICtrlRead($n1) & @CRLF& stringleft ($sms_to_delete , 10) )
			if FileExists( @ScriptDir &"\"&$user_name& "\" & $Filename) then 
				$sms_delete= GUICtrlRead($n1) ;& " <> " &$Filename
			Else
					MsgBox(0,"Warning", "�S������R�����ءC" & @CRLF &@ScriptDir &"\"&$user_name& "\" & $Filename,10)
					Exit
			EndIf
			;$temp_file=fileopen( @ScriptDir &"\"&$user_name& "\" & $Filename ,10)
			;FileWriteLine($temp_file , stringleft ($sms_to_delete , 13))
			;fileclose($temp_file)
			;FileMove(@ScriptDir &"\"&$user_name& "\" & $Filename, @ScriptDir &"\"&$user_name& "\" & $Filename &".omitted")
			exitloop
		EndIf
	
	Until $msg = $GUI_EVENT_CLOSE
GUIDelete();
;MsgBox(0,"SMS to Delete in Func", @ScriptDir &"\"&$user_name& "\" & $Filename)
if $sms_to_delete='' then 
	MsgBox(0,"Warning", "�S������R�����ءC",10)
	Exit
EndIf
;return    &$Filename
;MsgBox(0,"SMS to Delete in Func", $sms_delete )

EndFunc   ;==>Example
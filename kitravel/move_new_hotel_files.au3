#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here



;
;
;
;
;
dim $2bnb_demo="f:\Tomcat6\webapps\ROOT\demo"
dim $2bnb_rule="f:\Tomcat6\webapps\ROOT\hotel_rule"
dim $2kitravel_booking="d:\Tomcat6\webapps\ROOT\booking"
dim $2kitravel_rule="d:\Tomcat6\webapps\ROOT\hotel_rule"

dim $current_id


dim $new_hotel_files[9]
$new_hotel_files[0]=9
$new_hotel_files[1]="index.htm"
$new_hotel_files[2]="orderConfirm.htm"
$new_hotel_files[3]="b_index.htm"
$new_hotel_files[4]='orderRule.htm'
$new_hotel_files[5]='rule.htm'
$new_hotel_files[6]='orderRule_e.htm'
$new_hotel_files[7]='rule_en.htm'
$new_hotel_files[8]='orderRule_1.png'
	$current_id =  StringTrimLeft(@ScriptDir, StringInStr(@ScriptDir, "\" ,0,-1) )
	;MsgBox(0,"current dir", @ScriptDir & @CRLF & $current_id,5)
	Local $answer = InputBox("New hotel name_ID", "New name_ID of " &$current_id& "?" , $current_id, "", _
         - 1, -1, 0, 0)
	if $answer="" then
		MsgBox(0,"","No new hotel name ID. Exit")
		exit
	EndIf
;	MsgBox(0, "file copy 1", "File copy from "& @ScriptDir & "\" & $new_hotel_files[1] &" ---> " & $2kitravel_booking &"\"& $answer & @CRLF )
;	MsgBox(0, "file copy 2", "File copy from "& @ScriptDir & "\" & $new_hotel_files[4] &" ---> "& $2kitravel_rule &"\"& $current_id & @CRLF )
;	MsgBox(0, "file copy 3", "File copy from "& @ScriptDir & "\*.png"  &" ---> "& $2kitravel_rule &"\"& $current_id & @CRLF )

	MsgBox(0, "file copy 1", "File copy from "& @ScriptDir & "\" & $new_hotel_files[1] &" ---> " & $2kitravel_booking &"\"& $answer & @CRLF ,5)
	if FileExists(@ScriptDir & "\" & $new_hotel_files[1] ) then filecopy ( @ScriptDir & "\" & $new_hotel_files[1] , $2kitravel_booking &"\"& $answer&"\",10)
	if FileExists(@ScriptDir & "\" & $new_hotel_files[2] ) then filecopy ( @ScriptDir & "\" & $new_hotel_files[2] , $2kitravel_booking &"\"& $answer&"\",10)

	MsgBox(0, "file copy 2", "File copy from "& @ScriptDir & "\" & $new_hotel_files[4] &" ---> "& $2kitravel_rule &"\"& $current_id & @CRLF ,5)
	if FileExists(@ScriptDir & "\" & $new_hotel_files[4] ) then filecopy ( @ScriptDir & "\" & $new_hotel_files[4] , $2kitravel_rule &"\"& $current_id&"\",10)
	if FileExists(@ScriptDir & "\" & $new_hotel_files[5] ) then filecopy ( @ScriptDir & "\" & $new_hotel_files[5] , $2kitravel_rule &"\"& $current_id&"\",10)
	if FileExists(@ScriptDir & "\" & $new_hotel_files[6] ) then filecopy ( @ScriptDir & "\" & $new_hotel_files[6] , $2kitravel_rule &"\"& $current_id&"\",10)
	if FileExists(@ScriptDir & "\" & $new_hotel_files[7] ) then FileCopy ( @ScriptDir & "\" & $new_hotel_files[7] , $2kitravel_rule &"\"& $current_id&"\",10)

	MsgBox(0, "file copy 3", "File copy from "& @ScriptDir & "\*.png"  &" ---> "& $2kitravel_rule &"\"& $current_id & @CRLF ,5)
	if FileExists(@ScriptDir & "\" & $new_hotel_files[8] ) then filecopy ( @ScriptDir & "\*.png"  , $2kitravel_rule &"\"& $current_id&"\",10)

	if FileExists(@ScriptDir & "\" & $new_hotel_files[4] ) then filecopy ( @ScriptDir & "\" & $new_hotel_files[4] , $2bnb_rule &"\"& $current_id&"\",10)
	if FileExists(@ScriptDir & "\" & $new_hotel_files[5] ) then filecopy ( @ScriptDir & "\" & $new_hotel_files[5] , $2bnb_rule &"\"& $current_id&"\",10)
	if FileExists(@ScriptDir & "\" & $new_hotel_files[6] ) then filecopy ( @ScriptDir & "\" & $new_hotel_files[6] , $2bnb_rule &"\"& $current_id&"\",10)
	if FileExists(@ScriptDir & "\" & $new_hotel_files[7] ) then FileCopy ( @ScriptDir & "\" & $new_hotel_files[7] , $2bnb_rule &"\"& $current_id&"\",10)
	if FileExists(@ScriptDir & "\" & $new_hotel_files[8] ) then filecopy ( @ScriptDir & "\*.png"  , $2bnb_rule &"\"& $current_id&"\",10)
	if FileExists(@ScriptDir & "\" & $new_hotel_files[3] ) then filecopy ( @ScriptDir & "\" & $new_hotel_files[3] , $2bnb_demo &"\"& $answer&"\index.htm",10)

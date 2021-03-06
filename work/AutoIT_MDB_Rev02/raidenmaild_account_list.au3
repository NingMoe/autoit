; *******************************************************
; Examples for Access.au3
; *******************************************************
;
#include "Access.au3"
#include <array.au3>
#include <file.au3>

dim $sec=@SEC
dim $min=@MIN
Dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR
global $today=$year & $month & $day


dim $record_count
dim $field_array
dim $record_array
dim $array_html
dim $text_html
dim $r_Hide_list
dim $r_output_path

_DB_path()
$r_Hide_list= _HideList()
$r_output_path = _output_path()
	;MsgBox(0,"Output ", "Outpath : " & $r_output_path )
	;MsgBox (0, "Ex.6", "Example_AccessRecordsCount()")

	$record_count= Example_AccessRecordsCount( @ScriptDir & "\raidenmaild.mdb" , "usertable" )

	;MsgBox (0, "Ex.7", "Example_AccessTablesList()")
	;	Example_AccessTablesList()
	;MsgBox (0, "Ex.8", "Example_AccessFieldsList()")

	$field_array= Example_AccessFieldsList(@ScriptDir & "\raidenmaild.mdb" , "usertable")

	;MsgBox (0, "Ex.9", "Example_AccessRecordList()")
dim $array_account_list[$record_count+1][3]

for $x =1 to $record_count

	$record_array =  Example_AccessRecordList(  @ScriptDir & "\raidenmaild.mdb" , "usertable" , $x )
	if $record_array [10] = 0 then      ;  $record_array [10] = 0 表示這個人的 Disable_account 為 "否"
		;ConsoleWrite (@CRLF & $record_array[2] &" _ " & $record_array[1] &" _ " & $record_array[6])
		$array_account_list[$x][2]=$record_array[10]
		$array_account_list[$x][0]=$record_array[2]
		$array_account_list[$x][1]=$record_array[1] & "@kidsburgh.com.tw"

		;if $record_array[1]="bryant" then $array_account_list[$x][1]="-"
		;if $record_array[1]="ksbackup" then $array_account_list[$x][1]="-"
		;if $record_array[1]="service" then $array_account_list[$x][1]="-"
		for $r =1 to $r_Hide_list[0]
		if $record_array[1]= $r_Hide_list[$r] then $array_account_list[$x][1]="-"

		Next


	Else
		$array_account_list[$x][0]="-" ;$record_array[2]
		$array_account_list[$x][1]="-";$record_array[1]
		$array_account_list[$x][2]="-";$record_array[10]
	EndIf
	;if $record_array [10] = 1 then  MsgBox(0,"Dsiable Account", $record_array[2] &" , " &$record_array[1] & "@kidsburgh.com.tw"  );_ArrayDisplay ( $record_array )
	;_ArrayDisplay ($array_account_list)
next

_ArraySort ($array_account_list, 0,0,0,1 )

$array_html = _template()
for $x =1 to $record_count
;_ArrayDisplay ($array_account_list)

	if $array_account_list[$x][1]<> "-" then
	;$text_html = $text_html & $array_account_list[$x][0] & ' :  <a href="mailto:' &$array_account_list[$x][1]& '">' & $array_account_list[$x][1] & '</a><br>'  & @CRLF
	$text_html = $text_html &   StringReplace(  StringReplace ( $array_html[2] , "[name]", $array_account_list[$x][0] ) , "[email]" , $array_account_list[$x][1] ) & @CRLF
	EndIf

Next

	$text_html= $array_html[1]  & $text_html
	$text_html=$text_html &  $array_html[3]
	ConsoleWrite ($text_html)

$output=FileOpen ($r_output_path & "\contact.htm",10)
FileWrite ($output, $text_html )

FileClose($output)
;MsgBox (0, "Ex.10", "Example_AccessRecordAdd_Edit_Dele()")
;	Example_AccessRecordAdd_Edit_Dele()
;--------------------------------------------------------------
	;++++++++++++later
	; _AccessCreateMDB
	; _AccessCreateField ($oTable, "fld1" )
	; _AccessUpdateTable ($oNewDB, $oTable)
	; _AccessCreateTable (ByRef $o_object, $s_TableName ="")
;================================================================

Exit


; *******************************************************
; Example 6 - Count Number of fields in table
; *******************************************************
Func Example_AccessRecordsCount($access_mdb, $access_table)
	Local $o_DataBase = _AccessOpen($access_mdb)

	If $o_DataBase = 0 Then
		MsgBox(0, "Information", "Database file is not found :-" & @CR &  $access_mdb )
		Return
	EndIf

	Local $n_RecCount = _AccessRecordsCount ($o_DataBase, $access_table )

	;MsgBox(0, "Information", "No. Of Recordes in Table  is " & $n_RecCount)

	_AccessClose($o_DataBase)
	return $n_RecCount
EndFunc

; *******************************************************
; Example 7 - List tables in database
; *******************************************************

Func Example_AccessTablesList()
	Local $o_DataBase = _AccessOpen(@ScriptDir & "\raidenmaild.mdb")
	Local $avArray_Tables

	If $o_DataBase = 0 Then
		MsgBox(0, "Information", "Database file is not found :-" & @CR & @ScriptDir & "\raidenmaild.mdb")
		Return
	EndIf

	$avArray_Tables = _AccessTablesList ($o_DataBase)
	_ArrayDisplay ($avArray_Tables)

	$avArray_Tables = _AccessTablesList ($o_DataBase, False)
	_ArrayDisplay ($avArray_Tables)

	_AccessClose($o_DataBase)

EndFunc
; *******************************************************
; Example 8 - List Fields tables
; *******************************************************

Func Example_AccessFieldsList($access_mdb, $access_table)
	Local $o_DataBase = _AccessOpen($access_mdb)
	Local $O_TableObject, $avArray_Fields ;, $O_TableObject

	If $o_DataBase = 0 Then
		MsgBox(0, "Information", "Database file is not found :-" & @CR & $access_mdb)
		Return
	EndIf


	$avArray_Fields = _AccessFieldsList ($o_DataBase, $access_table)
	;_ArrayDisplay ($avArray_Fields)
	_AccessClose($o_DataBase)
	return ( $avArray_Fields )
EndFunc
; *******************************************************
; Example 9 - List the current Record
; *******************************************************
Func Example_AccessRecordList( $access_mdb, $access_table , $access_record_no)
	Local $o_DataBase = _AccessOpen($access_mdb )
	Local $O_TableObject, $avArray_Record ;, $O_TableObject

	If $o_DataBase = 0 Then
		MsgBox(0, "Information", "Database file is not found :-" & @CR & $access_mdb)
		Return
	EndIf

	$avArray_Record = _AccessRecordList ($o_DataBase, $access_table, $access_record_no)
	;_ArrayDisplay ($avArray_Record)
	_AccessClose($o_DataBase)
	return ( $avArray_Record )
EndFunc

; *******************************************************
; Example 10 - Add Record
; *******************************************************
Func Example_AccessRecordAdd_Edit_Dele()
	Local $o_DataBase = _AccessOpen(@ScriptDir & "\raidenmaild.mdb")
	Local $avData[5][2]
	Local $n_Feedback

	If $o_DataBase = 0 Then
		MsgBox(0, "Information", "Database file is not found :-" & @CR & @ScriptDir & "\raidenmaild.mdb")
		Return
	EndIf
	;---------- Data Setup
		$avData[0] [0] = "FLD1"
		$avData[1] [0] = "FLD2"
		$avData[2] [0] = "FLD3"
		$avData[3] [0] = "FLD4"

		$avData[4] [0] = "FLD1"


		$avData[0] [1] = "Will not be Written"
		$avData[1] [1] = 150
		$avData[2] [1] = @MDAY & "/" & @MON & "/" & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC
		$avData[3] [1] = 25

		$avData[4] [1] = "Fld1 is reset"
	;---------- Get the last Rec. position
	Local $n_RecCount = _AccessRecordsCount ($o_DataBase, "tblTable1")
	MsgBox (0,"Info.", "Before add data, Record Count =" & $n_RecCount )

	;---------- Save Data
	$n_Feedback = _AccessRecordAdd ($o_DataBase, "tblTable1", $avData)
	if $n_Feedback <> 1 Then
			MsgBox (0, "Info. feedback", "Error in writting Data")
			Return
	EndIf

	;---------- Get the last Rec. position
	Local $n_RecCount = _AccessRecordsCount ($o_DataBase, "tblTable1")
	MsgBox (0,"Info.", "After add data, Record Count =" & $n_RecCount )

	;----------- Show Last Entered Data
	Local $avArray_Record = _AccessRecordList ($o_DataBase, "tblTable1", $n_RecCount )
	_ArrayDisplay ($avArray_Record)

	;---------- Edit Data
		$avData[0] [1] = "Will not be Written"
		$avData[1] [1] = 11111
		$avData[2] [1] = @MDAY & "/" & @MON & "/" & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC
		$avData[3] [1] = 1111

		$avData[4] [1] = "All 1111"

	$n_Feedback = _AccessRecordEdit ($o_DataBase, "tblTable1", $avData, $n_RecCount)
	if $n_Feedback <> 1 Then
			MsgBox (0, "Info. feedback", "Error in writting Data")
			Return
	EndIf

	;----------- Show Last Entered Data
	Local $avArray_Record = _AccessRecordList ($o_DataBase, "tblTable1", $n_RecCount )
	_ArrayDisplay ($avArray_Record)

	;---------- Delete last Rec.
	Local $n_FeedBack = _AccessRecordDelete ($o_DataBase, "tblTable1", $n_RecCount  )
	if $n_FeedBack = 1 Then
		MsgBox (0,"info.", "Data was Deleted")
		$n_RecCount =  _AccessRecordsCount ($o_DataBase, "tblTable1")
		MsgBox (0,"Info.", "After Delete, Record Count =" & $n_RecCount )
	Else
		MsgBox (0,"info.", "Data was NOT Deleted")
	EndIf

	;---------- Close Database
	_AccessClose($o_DataBase)

EndFunc


func _template()
local $html_template[4]
$html_template[1]="<html>" & _
				  "<head>"  & _
				  '<meta http-equiv="Content-Language" content="zh-tw">'  & _
				  '<meta http-equiv="Content-Type" content="text/html; charset=big5">'  & _
				  "<title>Kidsburgh Email Address</title>"  & _
				  "</head>"  & _
				  "<body><center>"  & _
				  "<table width=50% height=100 border=1>" & _
				  "<caption>Kidsburgh Email Address </caption>"

$html_template[2]='<tr><th> [name] </th>  <td><a href="mailto:[email]">[email]</a> </td></tr>'
$html_template[3]="</table> Updated : " & $today& "</center></body>"  & _
				  "</html>"

Return($html_template)
EndFunc

; User DB format
;[0]|26
;[1]|accountname
;[2]|fullname
;[3]|domainname
;[4]|quota
;[5]|passwordtype
;[6]|userpassword
;[7]|autoforwardto
;[8]|backupto
;[9]|autoreplyflag
;[10]|inactivateflag
;[11]|lastlogin
;[12]|deletemailafterforwardingflag
;[13]|vipflag
;[14]|webmailflag
;[15]|icqnumber
;[16]|newmailnotificationpermission
;[17]|newmailnotificationflag
;[18]|popcmdpermission
;[19]|getcmdpermission
;[20]|cellularphonenumber
;[21]|icqsmspermission
;[22]|icqsmsflag
;[23]|createdate
;[24]|expirationcheckflag
;[25]|expirationdate
;[26]|accountneverexpiresflag


Func _DB_path()
	local $mode=""
	If FileExists(@ScriptDir & "\DB_path.txt") Then
		$mode = FileReadLine(@ScriptDir & "\DB_path.txt", 1)
		If $mode = "" Then
			if FileExists (@ScriptDir & "\raidenmaild.mdb" ) then
				MsgBox(0, "DB path", "Email Account DB :" & @ScriptDir & "\raidenmaild.mdb" , 5)
			Else
				MsgBox(0,"Error ", @ScriptDir & "\raidenmaild.mdb " & @CRLF  & "File not exists")
			EndIf
		Else
			;$mode <> ""
			if FileExists ($mode ) then
				MsgBox(0, "DB path", "Email Account DB at :" & $mode , 5)
				if FileExists (@ScriptDir & "\raidenmaild.mdb" ) then FileMove (@ScriptDir & "\raidenmaild.mdb" , @ScriptDir & "\backup_db\"& $today&"_raidenmaild.mdb" ,9)
				FileCopy ( $mode , @ScriptDir &"\raidenmaild.mdb" )
				if not FileExists (@ScriptDir & "\raidenmaild.mdb" ) then MsgBox(0,"Error ", "File at : "& $mode & @CRLF & "Can not copy to : " & @ScriptDir & "\raidenmaild.mdb "  )
			Else
				MsgBox(0,"Error 321 ", $mode & @CRLF  & "File not exists")
				$mode=""
			EndIf
		EndIf

	Else
		;MsgBox(0,"Process mode", " 高鐵車次資料會輸入資料庫 ",10)
		;$ans=InputBox("Process mode","高鐵車次資料會輸入資料庫 "&@CRLF& "輸入 N 可以離開")

		$mode = ""
		if FileExists (@ScriptDir & "\raidenmaild.mdb" ) then
			MsgBox(0, "DB path", "Email Account DB :" & @ScriptDir & "\raidenmaild.mdb" , 5)
		Else
			MsgBox(0,"Error ", @ScriptDir & "\raidenmaild.mdb " & @CRLF  & "File not exists")
		EndIf
		;if $ans="n" or $ans="N" or @error=1 then exit

	EndIf

	Return $mode
EndFunc   ;==>_TEST_MODE



Func _HideList()
	local $hidelist
	If FileExists(@ScriptDir & "\hidelist.txt") Then
		 _FileReadToArray(@ScriptDir & "\hidelist.txt",$hidelist)
		If  not IsArray ($hidelist) Then MsgBox(0, "Hide List", "No Hide name list", 5)



	EndIf
	;_ArrayDisplay ($hidelist)
	Return $hidelist
EndFunc   ;==>_TEST_MODE



Func _output_path()
	local $mode , $output_path
	If FileExists(@ScriptDir & "\output_path.txt") Then
		$output_path = FileReadLine(@ScriptDir & "\output_path.txt", 1)
		If $output_path = "" Then $output_path= @ScriptDir
	Else
		$output_path= @ScriptDir
	EndIf

	Return $output_path
EndFunc   ;==>_TEST_MODE


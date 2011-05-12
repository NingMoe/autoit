
#include <Array.au3>
#include <file.au3>
#include <Process.au3>
#include <Date.au3>
; Add date compare. 
; Not yet write a dame lin of code.
; 只有加入一些說明而已。

Dim	$min=@MIN
Dim	$hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR
dim $aFileList
dim $today=$year&$month&$day
dim $day_peroid=2

dim $test_mode=1
;; Test mode is not test by a file

if $test_mode=1 then 
	dim $mail_path="D:\downloads\eml"
	dim $move_path="D:\downloads\eml\moved"

Else
	dim $mail_path="C:\RaidenMAILD\Inboxes\ae_backup"
	dim $move_path="Z:\amex_data\Backup\ae_backup"

EndIf


$aFileList=_FileListToArray($mail_path, "*.eml",  1 )
if @error then Exit
_ArraySort($aFileList,0,1)
;_ArrayDisplay($aFileList)


for $x=$aFileList[0] to 1 step -1
	local $name
	local $creat_date
	$name=StringTrimLeft( $aFileList[$x] ,StringInStr($aFileList[$x],"_") )
	$creat_date = StringLeft($name,8)
	;MsgBox(0,"mail file name" , $name &@CRLF& $creat_date )
	if $today-$creat_date < $day_peroid  then 
		_ArrayDelete($aFileList,$x)
		$aFileList[0]=$aFileList[0]-1
	EndIf
Next

;_ArrayDisplay($aFileList)	


for $x=1 to $aFileList[0]
	local $name
	local $creat_date
	$name=StringTrimLeft( $aFileList[$x] ,StringInStr($aFileList[$x],"_") )
	$creat_date = StringLeft($name,8)

	;$t=filegettime($mail_path&"\"&$aFileList[$x],0,1)
	;MsgBox(0,"File time", $aFileList[$x]&@CRLF&  stringleft($t,8 )  ,5)
	if FileExists($mail_path&"\"&$aFileList[$x]) then 
		filemove($mail_path&"\"&$aFileList[$x], $move_path&"\"&$creat_date&"\",8)
		_WriteFile_append(@ScriptDir&"\"&$today&"_move_files.txt",$aFileList[$x] & " move to "&$move_path&"\"&$creat_date&"\"& $aFileList[$x] & @CRLF)
	EndIf
	sleep(100)
Next



Func _WriteFile( $_output_path ,$_line)
$_file = FileOpen($_output_path, 2)
;MsgBox(0,"Path",$date_output_path)

; Check if file opened for reading OK
If $_file = -1 Then
    MsgBox(0, "Error", "Unable to open file : " & $_output_path )
    Exit
EndIf
	FileWriteLine($_file,$_line)
FileClose($_file)

EndFunc

Func _WriteFile_append( $_output_path ,$_line)
$_file = FileOpen($_output_path, 1)
;MsgBox(0,"Path",$date_output_path)

; Check if file opened for reading OK
If $_file = -1 Then
    MsgBox(0, "Error", "Unable to open file : "& $_output_path)
    Exit
EndIf
	FileWriteLine($_file,$_line)
FileClose($_file)

EndFunc
	
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <array.au3>
;#include <mysql.au3>
#include <File.au3>
#include <Date.au3>
#include <gmail_func.au3>

Dim $sec = @SEC
Dim $min = @MIN
Dim $hour = @HOUR
Dim $day = @MDAY
Dim $month = @MON
Dim $year = @YEAR
Dim $today = $year & $month & $day

dim $ini_file= StringTrimRight ( @ScriptName ,4 )& ".txt"
dim $ini_array


dim $outbox, $validate   ,$keyword ,$keyword2 ,$report_to
dim $eml_list
dim $match_keyword=0
dim $test = _TEST_MODE()

if FileExists ( @ScriptDir &"\"& $ini_file  ) then

	_FileReadToArray(@ScriptDir &"\"& $ini_file,  $ini_array )

	;_ArrayDisplay ($ini_array)

	for $x=$ini_array[0] to 1 step -1
		if StringInStr( $ini_array[$x],";" )  then _ArrayDelete($ini_array, $x)
		if StringInStr ( $ini_array[$x],"outbox=" ) then $outbox=  StringTrimLeft( $ini_array[$x] ,7)
		if StringInStr ($ini_array[$x],"validate=" ) then $validate=  StringTrimLeft( $ini_array[$x] ,9)
		if StringInStr ($ini_array[$x],"keyword=" ) then $keyword=  StringTrimLeft( $ini_array[$x] ,8)
		if StringInStr ($ini_array[$x],"keyword2=" ) then $keyword2=  StringTrimLeft( $ini_array[$x] ,9)
		if StringInStr ($ini_array[$x],"report_to=" ) then $report_to=  StringTrimLeft( $ini_array[$x] ,10)
		if $ini_array[$x]="" then _ArrayDelete($ini_array, $x)

	Next
	$ini_array[0]=UBound($ini_array)-1
	;_ArrayDisplay ($ini_array)


EndIf
;MsgBox(0,"Var", $outbox &"  "& $validate&"  "&$keyword&"  "&$report_to )


dim $eml_create_time , $Date_diff
while 1
 $sec = @SEC
 $min = @MIN
 $hour = @HOUR
 $day = @MDAY
 $month = @MON
 $year = @YEAR
 $today = $year & $month & $day



$eml_list= _FileListToArray ( $outbox, "*",1)

	;_ArrayDisplay( $eml_list )
	if IsArray( $eml_list) then

		;_ArrayDisplay( $eml_list )
		for $x=1 to $eml_list[0]
			;MsgBox(0,"File",$outbox & $eml_list[$x] )
			if FileExists($outbox & $eml_list[$x]) then
				$eml_create_time =FileGetTime ($outbox & $eml_list[$x])

				if $test=1 then
					$Date_diff=100000
				Else
					$Date_diff = _DateDiff('h',$eml_create_time[0] &"/"& $eml_create_time[1] &"/" &$eml_create_time[2] &" "& $eml_create_time[3]&":"& $eml_create_time[4]&":"& $eml_create_time[5]  ,_NowCalc() )
					;MsgBox(0,"File and time", $outbox & $eml_list[$x] &@CRLF &  $eml_create_time[1] &" " &$eml_create_time[2] &"  "& $eml_create_time[3] )
					;MsgBox(0,"Date diff", $outbox & $eml_list[$x] &@CRLF &@CRLF & $Date_diff)
				EndIf

				if $Date_diff >= $validate then

					if $test=1 then
						$match_keyword=1
					Else
						$match_keyword= _read_mail ($outbox & $eml_list[$x], $keyword, $keyword2)
					EndIf

					if $match_keyword=1 then
						MsgBox(0,"Date diff", $outbox & $eml_list[$x] &@CRLF &@CRLF & $Date_diff ,10)
						_FileWriteLog(@ScriptDir &"\scan_mailoutbox_" &$today &".log", $outbox & $eml_list[$x] &" stay more then "& $Date_diff & " Hour.")
						_bryant_sendmail($report_to, "scan_mailoutbox " &$today &" notice." , $outbox & $eml_list[$x] &" stay more then "& $Date_diff & " Hour."  )
						 Run(@ComSpec & " /c " & @ScriptDir&'\move_file.bat ' &  $outbox & $eml_list[$x] , "", @SW_HIDE)

					EndIf
				EndIf

			EndIf



		Next
	EndIf

	if $test=1 then
		sleep(30*1000)
	Else
		Sleep(1800*1000)
	EndIf
WEnd

Exit

func _read_mail($eml_file_inf ,$keyword_inf, $keyword2_inf)
	local $line , $f , $bingo
	$bingo=0
	if FileExists ($eml_file_inf) then
		for $f=1 to 20
			$line=FileReadLine($eml_file_inf, $f)
			if StringInStr ($line,$keyword_inf) and StringInStr ($line,$keyword2_inf) then $bingo=1
		Next
	EndIf
return $bingo
EndFunc


Func _TEST_MODE()
	If FileExists(@ScriptDir & "\TESTMODE.txt") Then
		$mode = FileReadLine(@ScriptDir & "\TESTMODE.txt", 1)
		If $mode = 1 Then
			MsgBox(0, "Test mode", "This is Test mode. ", 5)

		Else
			;MsgBox(0,"Delivery mode", "This is True delivery.",5)
			$mode = 0
		EndIf

	Else
		;MsgBox(0,"Delivery mode", "This is True  delivery.",5)
		$mode = 0
	EndIf

	Return $mode
EndFunc   ;==>_TEST_MODE

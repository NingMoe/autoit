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
#include <FTPEx.au3>



Dim $sec = @SEC
Dim $min = @MIN
Dim $hour = @HOUR
Dim $day = @MDAY
Dim $month = @MON
Dim $year = @YEAR
Dim $today = $year & $month & $day

dim $aFilelist_real
dim $aFilelist_rental

dim $aRental
Dim $aReal_estate
dim $count=0

dim $real_estate = "real_estate-" ;"real_estate-2012-11-01.csv"
dim $rental ="rentals-" ;"rentals-2012-11-01.csv"
dim $date_between=7
dim $data_path="E:\workarround\房屋資料\2012-12-01全"
dim $line_per_batch=1000
dim $aSplit_files_Returned


dim $return_txt_data_array

If FileExists(@ScriptDir & "\real_estate.txt") Then
	_FileReadToArray( @ScriptDir & "\real_estate.txt", $return_txt_data_array )
EndIf
	;_ArrayDisplay($return_txt_data_array)
	;Data_Path
	;Real_Estate_Name
	;Rental_Name
	;DateBetween
	for $x=1 to $return_txt_data_array[0]
		;ConsoleWrite(  $return_txt_data_array[0])
		if StringInStr ($return_txt_data_array[$x] ,"Data_Path" ) then
			$count=0
			$count= StringInStr($return_txt_data_array[$x],"=" )
			;ConsoleWrite( @CRLF & $count)
			if $count>0 then $return_txt_data_array[$x]= StringTrimLeft ($return_txt_data_array[$x], $count )
			$data_path=$return_txt_data_array[$x]
		EndIf

		if StringInStr ($return_txt_data_array[$x] ,"Real_Estate_Name" ) then
			$count=0
			$count= StringInStr($return_txt_data_array[$x],"=" )
			;ConsoleWrite( @CRLF & $count)
			if $count>0 then $return_txt_data_array[$x]= StringTrimLeft ($return_txt_data_array[$x], $count )
			$real_estate=$return_txt_data_array[$x]
		EndIf

		if StringInStr ($return_txt_data_array[$x] ,"Rental_Name" ) then
			$count=0
			$count= StringInStr($return_txt_data_array[$x],"=" )
			;ConsoleWrite( @CRLF & $count)
			if $count>0 then $return_txt_data_array[$x]= StringTrimLeft ($return_txt_data_array[$x], $count )
			$rental =$return_txt_data_array[$x]
		EndIf

		if StringInStr ($return_txt_data_array[$x] ,"DateBetween" ) then
			$count=0
			$count= StringInStr($return_txt_data_array[$x],"=" )
			;ConsoleWrite( @CRLF & $count)
			if $count>0 then $return_txt_data_array[$x]= StringTrimLeft ($return_txt_data_array[$x], $count )
			$date_between =$return_txt_data_array[$x]
		EndIf
		;
		if StringInStr ($return_txt_data_array[$x] ,"LinePerBatch" ) then
			$count=0
			$count= StringInStr($return_txt_data_array[$x],"=" )
			;ConsoleWrite( @CRLF & $count)
			if $count>0 then $return_txt_data_array[$x]= StringTrimLeft ($return_txt_data_array[$x], $count )
			$line_per_batch =$return_txt_data_array[$x]
		EndIf

	Next
	;ConsoleWrite ( @CRLF & $date_between &" , " & $rental &" , "& $real_estate &" , "& $data_path&" , "& $line_per_batch & @CRLF )
	;_ArrayDisplay($return_txt_data_array)

; This for test purpose
;$rental="real_estate-2012-11-00.csv"
$rental="real_estate-2012-11-01.csv"
global $process_file_date = "2012/11/01"
Global $processd_file =  $rental
;$rental="test - 複製.txt"



	$aSplit_files_Returned =_split_txt($data_path, $rental , $line_per_batch )
	if not IsArray ($aSplit_files_Returned ) then MsgBox (0,"Error!", "Split file with error." &@CRLF & " Check them at " & $data_path& "\" & $processd_file )
	_ArrayDisplay ($aSplit_files_Returned)
	;;rem _split_txt($f_data_path ,$f_file, $f_line_per_batch)
	;Local $begin = TimerInit()
	;_newFile2Array(@ScriptDir & "\hotel.txt", 5, "^", 0)

;$aRental = _special_File2Array($data_path & "\" & $rental, 21 , ',' ,1 )

;sleep(5*1000)
;Local $dif = TimerDiff($begin)
;ConsoleWrite ( @CRLF &"Load " & $rental& " Need "& ($dif/1000) &" Sec" & @CRLF )
;_ArraySort( $aRental,0,0,0,1)


;_ArrayDisplay($aRental,"First 1000 Line" ,1000)

; $dif = TimerDiff($begin)
;ConsoleWrite ( @CRLF &"Display array Need "& ($dif/1000) &" Sec" & @CRLF )
Exit

Func _split_txt($f_data_path ,$input, $f_line_per_batch)

	;local 	$input
	local   $f_file
	local	$output = "_" & $input
	local 	$split = $f_line_per_batch
	local $l_counter=0
	local $output_f
	local $output_f_counter=1
	local  $s , $find_EndOfRecord
	local $aSplit_files
	local $aSplit_files_error=0


	if FileExists ( $f_data_path&"\aSplit_files.txt") then
		_FileReadToArray ($f_data_path&"\aSplit_files.txt", $aSplit_files)
		Return $aSplit_files

	EndIf

;
;	$aSplit_files =  _FileListToArray ($f_data_path &"\"  , StringtrimRight($output,4) &"*" ,1 )
;		_ArrayDisplay ($aSplit_files)
;		MsgBox (0,"Wait","Wait a while")
	$f_file = FileOpen($f_data_path&"\"&$input, 0)
	;$s = StringRegExpReplace($f_file, chr(10), "")
	;ConsoleWrite(" Replace @LF: " &$s)

	$output_f = FileOpen($f_data_path&"\"&$output, 10)

	; Check if file opened for reading OK
	If $f_file = -1 Then
		MsgBox(0, "Error", "Unable to open file: "&$input)
		Exit
	EndIf
	local $begin  = TimerInit()
	; Read in lines of text until the EOF is reached
	While 1

		Local $dif = TimerDiff($begin)
		$line = FileReadLine($f_file)
			If @error = -1 Then ExitLoop
			;MsgBox(0, "Line read:", $line)

			$l_counter=$l_counter+1

			if mod($l_counter,$split)=0 then $find_EndOfRecord=1

			FileWriteLine($output_f,$line)

			if $find_EndOfRecord=1 then
				;ConsoleWrite( @CRLF & "find_EndOfRecord at Line: "& $l_counter & @CRLF &$line )
				;MsgBox(0,"end of reocrd ??? "," ??? Line: "& $l_counter & @CRLF& $line)
				;$find_EndOfRecord=1
				; \x0d => @CR , \x0a => @LF
				 if  StringRegExp($line, "([0-9]{1,2})\:([0-9]{1,2})([0-9]{1,2})\:([0-9]{1,2})" ) then
					;ConsoleWrite ("Bingo! Line: "& $l_counter & @CRLF &$line &@CRLF)
					;MsgBox(0,"Bingo! Find end of reocrd", @CRLF & @CRLF &"Line: "& $l_counter & @CRLF& $line &@CRLF)
					FileClose($output_f)
					FileMove($f_data_path&"\"&$output  , $f_data_path &"\" & StringTrimRight($output,4) & "_" & ( 1000+ $output_f_counter ) & StringRight($output,4) ,9)
					 ;MsgBox(0,"Wait!","Check file from  : " & $f_data_path&"\"&$output &@CRLF &" to " & @CRLF & $f_data_path &"\" & StringTrimRight($output,4) &$output_f_counter & StringRight($output,4)  )
					sleep(500)
					$output_f = FileOpen($f_data_path&"\"&$output, 10)
					$output_f_counter=$output_f_counter+1
					$find_EndOfRecord=0

					;local $dif = TimerDiff($begin)
					;ConsoleWrite ("Finish partial Split. It costs : "& ($dif/1000) &" Sec" & @CRLF )
				EndIf


			EndIf
	Wend
	FileClose($output_f)
	FileMove($f_data_path&"\"&$output  , $f_data_path &"\" & StringTrimRight($output,4) & "_" & ( 1000+ $output_f_counter ) & StringRight($output,4) ,9)
	FileClose($f_file)

	$dif = TimerDiff($begin)
	ConsoleWrite ("Finish Split. It costs : "& ($dif/1000) &" Sec" & @CRLF )

	$aSplit_files =  _FileListToArray ($f_data_path &"\"  , StringtrimRight($output,4) &"*" ,1 )
	;_ArrayDisplay ($aSplit_files)
	if  IsArray ($aSplit_files) then
		Return $aSplit_files
	else
		return $aSplit_files_error
	EndIf

EndFunc


;; 2011 09 New Text to Two dimension array  with a start_line
Func _special_File2Array($PathnFile, $aColume, $delimiters, $Start_line)
	;local $replace_status
	;  $replace_status= _ReplaceStringInFile ($PathnFile,"\r","")
	;	ConsoleWrite (@CRLF & "String replace: " & $replace_status)
	local $f_file =FileOpen ( $PathnFile,0)
	local $f_string=""
	local $f_line=""

	Local $begin = TimerInit()
	ConsoleWrite ("Start read file " & @CRLF )
	While 1
		Local $f_line = FileReadline($f_file)
		If @error = -1 Then ExitLoop
		;MsgBox(0, "Char read:", $line)
		$f_string=$f_string& $f_line & "# "
	WEnd
	FileClose($f_file)
	;FileWrite (@ScriptDir &"\output.txt", @CRLF&"Just Readin :" &$string )
	Local $dif = TimerDiff($begin)
	ConsoleWrite ("Start to replace. Load File costs : "& ($dif/1000) &" Sec" & @CRLF )
		; \x0d => @CR , \x0a => @LF
	$sNew_Text = StringRegExpReplace($f_string, "([0-9]{1,2})\:([0-9]{1,2})([0-9]{1,2})\:([0-9]{1,2})\#\x20", "," &$process_file_date&"#!^")
	$sNew_Text = StringStripCR ($sNew_Text)
	$dif = TimerDiff($begin)
	;ConsoleWrite ($sNew_Text )
	ConsoleWrite ("Finish RegExpReplace. It costs : "& ($dif/1000) &" Sec" & @CRLF )

	Local $aRecords

	$aRecords = StringSplit ( $sNew_Text ,"^" )

	;_ArrayDisplay ($aRecords,"First 1000 line", 1000)

	;MsgBox

	For $x = $Start_line To 1 Step -1
		_ArrayDelete($aRecords, $x)
	Next
	$aRecords[0] = $aRecords[0] - $Start_line
	;_ArrayDisplay ($aRecords)
	;c

	Local $TextToArray[$aRecords[0]][$aColume + 1]

	;$TextToArray[0][0]=$aRecords[0]
	Local $aRow
	For $y = 1 To $aRecords[0]
		if StringInStr ( $aRecords[$y],"," ,0,$aColume ) >0 then
			if StringInStr($aRecords[$y],'"',0,2) >0  then
				local $quote1=StringInStr($aRecords[$y],'"',0,1)
				local $quote2=StringInStr($aRecords[$y],'"',0,2)
				local $comma_at = StringInStr($aRecords[$y],',',0,1, $quote1, $quote2)
				local $f_temp_string = StringMid ( $aRecords[$y], $quote1, ($quote2- $quote1)+1 )
				;local $f_abc
				;MsgBox(0,"Text in quote", StringReplace ( $f_temp_string,',',' ' ) )
				$aRecords[$y] =StringLeft($aRecords[$y],$quote1-1 ) &  StringReplace ( $f_temp_string,',',' ' )  & StringTrimLeft($aRecords[$y],$quote2 )



				;$f_abc =StringLeft($aRecords[$y],$quote1-1 ) &  StringReplace ( $f_temp_string,',',' ' )  & StringTrimLeft($aRecords[$y],$quote2 )
				;ConsoleWrite ( @CRLF&$aRecords[$y]  &@CRLF& $f_abc&@CRLF)
				;MsgBox (0,"conbined text", $aRecords[$y]  &@CRLF& $f_abc)
				;$aRecords[$y]= StringReplace ($aRecords[$y], $comma_at ,"_"  )
				;ConsoleWrite ($comma_at & @CRLF&  StringReplace ($aRecords[$y], $comma_at ,"_"  ) )
				;MsgBox (0,"original and replaced " & $y, $comma_at & @CRLF&  StringReplace ($aRecords[$y], $comma_at ,"_"  ) )
				;local $replaced_string= StringRegExpReplace( $aRecords[$y], ' \"*\,*\"', "replaced")
				;MsgBox (0,"original and replaced " & $y, @CRLF& $aRecords[$y] & @CRLF & $replaced_string )
			EndIf

		EndIf

		if StringInStr ( $aRecords[$y],"," ,0,$aColume ) >0 then
			;ConsoleWrite(@ScriptDir &"\error_data.txt")
			;;  "," 會比資料中多一個，是在 replace 中加入的
			;MsgBox(0,"More then 21 comma at row : " & $y &" "& StringInStr ( $aRecords[$y],"," ,0,20), $aRecords[$y])
				if not FileExists (@ScriptDir &"\error_data.txt") then
					$t= FileOpen (@ScriptDir &"\error_data.txt",10)
					FileClose($t)
				EndIf
				FileWrite (@ScriptDir &"\error_data.txt",@CRLF &  $aRecords[$y] )
				ContinueLoop
			EndIf

		;Msgbox(0,'Record:' & $y, $aRecords[$y])

		$aRow = StringSplit($aRecords[$y], $delimiters)
		;Msgbox(0,'X ,Colume :', $aRow[0])
		;_ArrayDisplay ($aRow)
		For $x = 1 To $aRow[0]
			;If StringInStr($aRow[$x], ",") Then

			;	$aRow[$x] = StringTrimLeft($aRow[$x], 1)
			;	;MsgBox(0, "after", $aRow[$x])
			;EndIf
			;ConsoleWrite( "Row No: " & $y & " Colume no: "& $x &" # ")
			$TextToArray[$y - 1][$x - 1] = $aRow[$x]
		Next
		;ConsoleWrite(@CRLF)
	Next
	;ConsoleWrite (@CRLF)
	;_ArrayDisplay($TextToArray)
	Return $TextToArray

EndFunc   ;==>_newFile2Array

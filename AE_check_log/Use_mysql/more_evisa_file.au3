#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <file.au3>
#include <array.au3>

dim $sec=@SEC
dim $min=@MIN
Dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR
dim $today=$year & $month & $day


dim $return_evisa_file_array =_more_evisa_files() 

if IsArray( $return_evisa_file_array ) then  

	if $return_evisa_file_array[0]>0 then  _combine_evisa_files ($return_evisa_file_array)
 Else
	 MsgBox(0, "Error in main " , "no file for  process or file in error format. ")
EndIf


Exit


Func _more_evisa_files()
	local $fit_pattern
	local $search = FileFindFirstFile(@ScriptDir&"\*_evisa_nbr.txt")  
	local $evisa_file_array[1]
	$evisa_file_array[0]=0

	; Check if the search was successful
	If $search = -1 Then
		MsgBox(0, "Error _more_evisa_file ", "No files/directories matched pattern '*_evisa_nbr.txt' ",5)
		;Exit
		return 0
	Else
		
		While 1
			$evisa_file = FileFindNextFile($search) 
				If @error Then 
					FileClose($search)
					;_ArrayDisplay($evisa_file_array)
					return $evisa_file_array
				EndIf	
				$fit_pattern=_check_line_pattern ( $evisa_file, 1 , ";Evisa補單專用，每一行一個ea_nbr" )
				
			if $fit_pattern then ;MsgBox(4096, "File:", $evisa_file)
				_ArrayAdd($evisa_file_array, $evisa_file)
				$evisa_file_array[0]=$evisa_file_array[0]+1
			
			;Else
				;MsgBox(0,"more evisa file", "_ArraySearch($evisa_file_array, $evisa_file) "& _ArraySearch($evisa_file_array, $evisa_file) )
				;_ArrayDelete ($evisa_file_array, _ArraySearch($evisa_file_array, $evisa_file) )
				;$evisa_file_array[0]=$evisa_file_array[0]-1
				;MsgBox(0,"File move in _more_evisa_file" , @ScriptDir & "\" &$evisa_file & "  move >>>> " &@ScriptDir & "\evisa_nbr_reprocessed\" & $today& "_error_" &$evisa_file)
				;filemove (@ScriptDir & "\" &$evisa_file, @ScriptDir & "\evisa_nbr_reprocessed\" & $today& "_error_" &$evisa_file,9 )
				;_ArrayDisplay( $evisa_file_array )
			EndIf
		WEnd
	EndIf

	;_ArrayDisplay( $evisa_file_array )
EndFunc


func _check_line_pattern ( $check_file, $line_no , $line_pattern)
		if FileExists(@ScriptDir&"\"&$check_file ) then 
				if FileReadLine( $check_file , $line_no  ) <> $line_pattern then 
					_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",'_check_line_pattern() , EVISA resend 檔案錯誤，沒有 "' & $line_pattern & '" 這樣的文字 ')
					;MsgBox(0,"File move in _check_line_pattern" , @ScriptDir&"\"&$check_file & "  move >>>> " &@ScriptDir & "\evisa_nbr_reprocessed\" & $today& "_error_" &$check_file)
					filemove (@ScriptDir&"\"&$check_file, @ScriptDir & "\evisa_nbr_reprocessed\" & $today& "_error_" &$check_file,9 )
					return 0
				else 
					return 1
				EndIf
		Else
			return 0
		EndIf

EndFunc



Func _combine_evisa_files ($evisa_file_array)
	local $func_string=""
	local $func1_evisa_nbr_txt_array
	local $func2_evisa_nbr_txt_array
	local $open_file
	
	;_ArrayDisplay($evisa_file_array)
	for $x=1 to $evisa_file_array[0]
		
		if $x=1 then 
			if FileExists(@ScriptDir & "\" &$evisa_file_array[$x]) then _FileReadToArray(@ScriptDir & "\" &$evisa_file_array[$x], $func1_evisa_nbr_txt_array )
			if IsArray( $func1_evisa_nbr_txt_array ) then filemove (@ScriptDir & "\" &$evisa_file_array[$x], @ScriptDir & "\evisa_nbr_reprocessed\" & $today& "_" &$evisa_file_array[$x],9 )
				;_ArrayDisplay($evisa_nbr_txt_array) 
		EndIf
		
		if $x>=2 then
				if FileExists(@ScriptDir & "\" &$evisa_file_array[$x]) then _FileReadToArray(@ScriptDir & "\" &$evisa_file_array[$x], $func2_evisa_nbr_txt_array )
				if IsArray( $func2_evisa_nbr_txt_array ) then filemove (@ScriptDir & "\" &$evisa_file_array[$x], @ScriptDir & "\evisa_nbr_reprocessed\" & $today& "_" &$evisa_file_array[$x],9 )

				for $v=2 to  $func2_evisa_nbr_txt_array[0]
					if _ArraySearch($func1_evisa_nbr_txt_array, $func2_evisa_nbr_txt_array[$v] ) =-1 then 
						_ArrayAdd( $func1_evisa_nbr_txt_array, $func2_evisa_nbr_txt_array[$v] )
					EndIf
				next
					
		
		EndIf

	Next
	_ArrayDelete($func1_evisa_nbr_txt_array,0)
	$func_string= _ArrayToString ( $func1_evisa_nbr_txt_array, @CRLF )
	$open_file=FileOpen (@ScriptDir & "\_evisa_nbr.txt" , 10)
	FileWrite ($open_file, $func_string )
	FileClose($open_file)

if FileExists (@ScriptDir & "\_evisa_nbr.txt") then 
	Return 1
Else
	Return 0
EndIf

EndFunc
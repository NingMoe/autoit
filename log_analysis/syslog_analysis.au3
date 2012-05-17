#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here


#include <array.au3>
#include <Date.au3>
#Include <File.au3>


dim $sec=@SEC
dim $min=@MIN
Dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR
DIM $today= $year& $month & $day
dim $file_list 
Global $current_file=""

$file_list=_FileListToArray(@ScriptDir, "SyslogCatchAll*.txt",1)

;_ArrayDisplay($file_list)


If  IsArray ($file_list)  Then
	
	for $x =1 to $file_list[0]
		$current_file= $file_list[$x]
		$array =  _logsplit( @ScriptDir & "\" &$file_list[$x],"	",2)
		;_ArrayDisplay($array, _NowTime() )
	
		;MsgBox(0,"_array2string", _array2string($extramail_array,2) )
		MsgBox(0,"Move" , "Now is " &$file_list[$x] & ", Move to next log file " ,5)
	Next

EndIf

func _logsplit(  $PathnFile, $delimiters, $f_count)
	local $f_logfile , $line , $f_search_word 
	local $filedate=""
	
if FileExists ($PathnFile) then 
	local $CountLines ,$counter
		$counter=0
		$filedate=  StringReplace( StringReplace( StringTrimRight( $current_file , 4) ,"-","")  ,"SyslogCatchAll_","")
		FileMove($PathnFile , @ScriptDir & "\syslog_backup\SyslogCatchAll_"& $filedate & ".txt",9)
		$CountLines = _FileCountLines(@ScriptDir & "\syslog_backup\SyslogCatchAll_"& $filedate & ".txt")
		$f_logfile=FileOpen(@ScriptDir & "\syslog_backup\SyslogCatchAll_"& $filedate & ".txt",0)
		
	While 1
		$counter=$counter+1
		$line = FileReadLine($f_logfile)
		If @error = -1 Then ExitLoop
			;MsgBox(0, "Line read:", $line)
		if StringInStr( $line , $delimiters) >0 then 
			$f_search_word=  Stringleft ( $line , StringInStr( $line , $delimiters,0,$f_count+1)-1 )
			;MsgBox(0, "Line read:", "Line >>> " & $line & @CRLF &  $f_search_word )
			
			$f_search_word=  StringTrimLeft ( $f_search_word , StringInStr( $line , $delimiters,0,$f_count)  )
		
			;MsgBox(0, "Line read:", "Line >>> " & $line & @CRLF &  $f_search_word )

			;MsgBox(0,"Current file" , $current_file & " >> " & $filedate )
			;MsgBox(0, "Line read:", "Line >>> " & $line & @CRLF &  $f_search_word  & @CRLF &  @ScriptDir &"\" &$filedate&"\" &  StringStripWS( StringReplace( $f_search_word,".","_" ) ,8) &".log")
			$file_to_write= fileopen( @ScriptDir &"\" &$filedate&"\" &  StringStripWS( StringReplace( $f_search_word,".","_" ) ,8) &".log" ,9)
			
			FileWriteLine ( $file_to_write  , $line )
			
			FileClose ($file_to_write )
			
		EndIf
		If MOD ($counter, 500 )=0 then TrayTip ("", "Porcess " & $counter &" / "& $CountLines ,5) 
	Wend
	;MsgBox(0,"Move Message",  $PathnFile & @CRLF & @ScriptDir & "\syslog_backup\SyslogCatchAll_"& $filedate & ".txt")

	
endif

EndFunc


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
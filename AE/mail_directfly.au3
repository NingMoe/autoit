#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.2.12.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <array.au3>
#include <file.au3>
#include <ExcelCOM_UDF.au3>  ; Include the collection
#include <Date.au3>

Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR

dim $title_row=0

dim $source_excel="\extra_email.csv"
	;$source_excel="\t.csv"
dim $pass="amextravel"
dim $account="ae_direct_fly"



$excel_array= _file2Array(@ScriptDir&"\"&$source_excel,2,";")
	if $title_row<>0 then 
		_ArrayDelete($excel_array,$title_row)
		$excel_array[0][0]=$excel_array[0][0]-$title_row
	EndIf	
;_ArrayDisplay($excel_array)





;; Two dimension array
Func _file2Array($PathnFile,$aColume,$delimiters)

	
	local $aRecords
	If Not _FileReadToArray($PathnFile,$aRecords) Then
		MsgBox(4096,"Error", " Error reading file '"&$PathnFile&"' to Array   error:" & @error)
		Exit
	EndIf
		;
		;MsgBox(0,"Line No ", $aRecords[0])
;;
;;  Delete empty line in the Text file. Usually at the end of file.
	For $y = $aRecords[0] to 1 step -1 
		local $minus=0
		if $aRecords[$y]="" then 
			_ArrayDelete($aRecords, $y)
			;MsgBox(0,"Array delete", "Delete Line : " &$y)
			$minus=$minus+1
		EndIf
		
	Next
	$aRecords[0]=$aRecords[0]-$minus
	;_ArrayDisplay($aRecords)
	
;;
;; Split line to 2-dimention array
	local $TextToArray[$aRecords[0]+1][$aColume]
		$TextToArray[0][0]=$aRecords[0]+1
		;_ArrayDisplay($TextToArray)
	
	
	local $aRow
	;Msgbox(0,'Record No from : ' ,"1 to " &$aRecords[0]+1)
	For $y = 1 to $aRecords[0]+1
		;Msgbox(0,'Record:' & $y, $aRecords[$y])
		 
		$aRow=StringSplit($aRecords[$y],$delimiters)
		;if $aRow[0]>2 then ;_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log","Line: "& $y& " >> " & $aRow[0] & " - "&$aRow[1] & " - "&$aRow[2])
		;Msgbox(0,'X ,Colume :', $aRow[0] & " - "&$aRow[1] & " - "&$aRow[2])
		;EndIf
		;_ArrayDisplay($aRow)
		
		;Msgbox(0,'X ,Colume :', $aRow[0] & " - "&$aRow[1] & " - "&$aRow[2])
		For $x=1 to $aRow[0]
			if StringInStr($aRow[$x],",") then 			
			$aRow[$x]=StringReplace($aRow[$x],',','')
				if StringInStr($aRow[$x],'"') then StringReplace($aRow[$x],'"','')
			;MsgBox(0, "after", $aRow[$x])
			EndIf
			 $TextToArray[$y][$x]=$aRow[$x]
			; MsgBox(0, "Line: "& $y,  $TextToArray[$y][$x])
		next
		_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log","Line: "& $y& " >> " &$TextToArray[$y][2])

		 if $y= $aRecords[0]+1 Then MsgBox(0, "Line: "& $y,  $TextToArray[$y][2])
	Next
	
	_ArrayDisplay($TextToArray)
	Return $TextToArray

EndFunc

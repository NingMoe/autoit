#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include<file.au3>
#include<array.au3>


dim $path="E:\firefox_download\rejected_20120229\20120229"  ; 退信所在的路俓
dim $path="R:\20120301"
dim $aMail_List
dim $to_find="TO:"
dim $reason1="Status:"
dim $reason2="Reason:"
dim $find_it=0

dim $omit="service@kidsburgh.com.tw"

dim $to=""
dim $status=""
dim $reason=""


dim $lines_of_mail

dim $x , $m 

dim $file_to_write

$file_to_write=FileOpen(@ScriptDir & "\file.txt",9 )

$aMail_List=_FileListToArray ($path,"*.eml")
;_ArrayDisplay($aMail_List)

for $x =1 to ($aMail_List[0] )
	$find_it=0	
	$to=""
	$status=""
	$reason=""
	
	_FileReadToArray($path & "\" & $aMail_List[$x],$lines_of_mail)
	
	;_ArrayDisplay($lines_of_mail)
	for $m=1 to $lines_of_mail[0]
		if StringInStr( $lines_of_mail[$m],$to_find) then 
			if not StringInStr($lines_of_mail[$m],$omit) then 
				;MsgBox(0,"To:", $lines_of_mail[$m] ,5)
				;FileWrite($file_to_write, StringReplace ($lines_of_mail[$m] , $to_find, "")  & " ; " )
				$to=StringReplace ($lines_of_mail[$m] , $to_find, "")
				$find_it=1
			EndIf
		EndIf	
		if StringInStr( $lines_of_mail[$m],$reason1) then 
			;MsgBox(0,"Reason1:", $lines_of_mail[$m] ,5)
			;FileWrite($file_to_write, $lines_of_mail[$m] & " ; ")
			$status= $lines_of_mail[$m] 
		EndIf	
		
			if StringInStr( $lines_of_mail[$m],$reason2) then 
			;MsgBox(0,"Reason2:", $lines_of_mail[$m] ,5)
			;FileWrite($file_to_write, $lines_of_mail[$m] & " ; " )
			$reason=$lines_of_mail[$m] 
		EndIf	
		
	Next
	FileWrite($file_to_write, $to & " ; " &  $status & " ; " & $reason & @CRLF)
	
	if $find_it=1 then FileMove ($path & "\" & $aMail_List[$x] , $path & "\processed\",9 )
Next
FileClose($file_to_write)

Exit
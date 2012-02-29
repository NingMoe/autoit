#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
; 20120228 ªìª©¡A©|¥¼§¹¦¨
#include<array.au3>



$section = IniReadSectionNames(@ScriptDir&"\scanfolder.ini")
If @error Then 
    MsgBox(4096, "", "Error occurred, probably no INI file.")
	Exit

EndIf

;_ArrayDisplay($section)


for $x=1 to $section[0]

	$var = IniReadSection(@ScriptDir&"\scanfolder.ini", $section[$x])
	If @error Then 
		MsgBox(4096, "", "Error occurred, probably no INI file.")
	Else
		;For $i = 1 To $var[0][0]
			;MsgBox(4096, "", "Key: " & $var[$i][0] & @CRLF & "Value: " & $var[$i][1])
		;Next
		;_ArrayDisplay($var)
		_scanfolder($var)
	EndIf
next

Exit


Func _scanfolder ($ini_array)
	;_ArrayDisplay($ini_array)
	local  $scandir, $file_ext,$content_format,$omit_ext,$action_path,$action,$move_path,$netdrive
	local $mount_info
	$scandir=$ini_array[1][1] 
	$file_ext=$ini_array[2][1]
	$content_format=$ini_array[3][1]   ; Not apply now
	$omit_ext=$ini_array[4][1]
	$action_path=$ini_array[5][1]
	$action=$ini_array[6][1]  			
	$move_path=$ini_array[7][1]
	$netdrive=$ini_array[8][1]
	
	
	if $netdrive<>"" then 
	$mount_info=StringSplit($netdrive,";")	
		if not FileExists($mount_info[1]) then DriveMapAdd($mount_info[1],$mount_info[2],0,$mount_info[3],$mount_info[4])
	EndIf
	
	
	if FileExists($scandir&"\*."&$file_ext) then 
		
		filemove($scandir&"\*."&$file_ext, $action_path,9) 	
		sleep(1000)
		run ($action)
		;Run(@ComSpec & " /c " & 'commandName', "", @SW_HIDE) 
		if not FileExists($scandir&"\*."&$file_ext) then FileMove($scandir&"\*."&$file_ext,$move_path,9)
		
	EndIf	



EndFunc


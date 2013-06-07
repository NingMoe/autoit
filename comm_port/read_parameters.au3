#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <array.au3>
#include <file.au3>
;_read_parameters ("EVB_commg_noui.au3")

func _read_parameters( $script_name )

local $ini_file= StringTrimRight ( $script_name ,4 )& ".ini"
local $ini_array


;MsgBox (0,"File is there?", @ScriptDir &"\"& $ini_file )

if FileExists ( @ScriptDir &"\"& $ini_file  ) then

	_FileReadToArray(@ScriptDir &"\"& $ini_file,  $ini_array )

	;_ArrayDisplay ($ini_array)

	for $x=$ini_array[0] to 1 step -1
			if $ini_array[$x]="" then _ArrayDelete($ini_array, $x)
	Next
	$ini_array[0]=UBound($ini_array)-1
	;_ArrayDisplay ($ini_array)


EndIf

Return ($ini_array)

EndFunc

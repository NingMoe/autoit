#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
#include <Constants.au3>
#include <file.au3>
#include <array.au3>
#Include <GDIPlus.au3>


$image= @ScriptDir&'\p1.jpg'
 MsgBox (0,"Image HxW", _get_image_width_hight($image) )

Exit
; Script Start - Add your code below here
;ConsoleWrite(@CRLF&@ScriptDir&'\identify.exe '& @ScriptDir&'\rule1.png' &@CRLF)
;$a = Run (@ComSpec & " /c " & @ScriptDir&'\identify.exe '& @ScriptDir&'\rule1.png', "", @SW_SHOW, $STDERR_MERGED )
$a = Run (@ComSpec & " /c " & @ScriptDir&'\identify.exe '& @ScriptDir&'\rule1.png', "", @SW_HIDE,  $STDERR_CHILD + $STDOUT_CHILD )

;$line=StdoutRead( $a )
local $output, $line
While 1
	$line = StdoutRead($a)
	If @error Then ExitLoop
	;MsgBox(0, "STDOUT read:", $line)
	if $line<>""  then $output= $output & $line
WEnd

ConsoleWrite(@CRLF& ": "&$output )
 ;MsgBox (0,"STD read",$a & "  "&$line ,5)





func _get_image_width_hight($image_file)
_GDIPlus_Startup ()
Local $hImage = _GDIPlus_ImageLoadFromFile( $image_file )
If @error Then
    MsgBox(16, "Error", "Does the "& $image_file &" exist?")
    Exit 1
EndIf
$w=_GDIPlus_ImageGetWidth($hImage)
ConsoleWrite(_GDIPlus_ImageGetWidth($hImage) & @CRLF)
$h=_GDIPlus_ImageGetHeight($hImage)
ConsoleWrite(_GDIPlus_ImageGetHeight($hImage) & @CRLF)
_GDIPlus_ImageDispose ($hImage)
_GDIPlus_ShutDown ()
return ( $w&" ; "&$h   )
;return ( $h )
EndFunc
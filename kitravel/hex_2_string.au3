#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <string.au3>
dim $String = ""

$file = FileOpen("test.txt", 0)

; Read in 1 character at a time until the EOF is reached
While 1
    $chars = FileRead($file, 1)
    If @error = -1 Then ExitLoop
    ;MsgBox(0, "Char read:", $chars)
	$String=$String& $chars
Wend
FileClose($file)


$file =FileOpen( @ScriptDir &"\t.bin" , 18 )
FileWrite ( $file ,$String )
FileClose($file)

$file = FileOpen(@ScriptDir & "\t.bin", 16)

; Read in 1 character at a time until the EOF is reached
While 1
    $chars = FileRead($file, 1)
    If @error = -1 Then ExitLoop
    ;MsgBox(0, "Char read:", $chars)
	$String=$String& $chars
Wend
FileClose($file)


MsgBox(0, "", $String)

Exit

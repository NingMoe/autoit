#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#Include <File.au3>
#include <Date.au3>
#include <array.au3>
#include <ntp_udp123_func.au3>

#include <_Zip.au3>

global $sec = @SEC
global $min = @MIN
global $hour = @HOUR
global $day = @MDAY
global $month = @MON
global $year = @YEAR
global $today = $year & $month & $day

local $upgrade_confirm=0
local $upgrade_filepath=@ScriptDir
local $upgrade_file_list_in_func
local $aFile_in_download_upgrade
local $base_dir

if FileExists (@ScriptDir &"\Upgrade_FileList_"& $today ) Then
	$upgrade_confirm= FileReadLine (@ScriptDir &"\Upgrade_FileList_"& $today,1)
	$upgrade_filepath= FileReadLine (@ScriptDir &"\Upgrade_FileList_"& $today,2)

	if $upgrade_confirm = ( $today &"=1")  then
		_FileReadToArray (@ScriptDir &"\Upgrade_FileList_"& $today, $aFile_in_download_upgrade )
		$base_dir= $aFile_in_download_upgrade[2]

		_ArrayDelete($aFile_in_download_upgrade,1)
		_ArrayDelete($aFile_in_download_upgrade,1)
		$aFile_in_download_upgrade[0]=$aFile_in_download_upgrade[0]-2
		;_ArrayDisplay($aFile_in_download_upgrade ,"Copy Upgrade Files")
		for $u =1 to $aFile_in_download_upgrade[0]

			if FileExists ( @ScriptDir &"\" & $aFile_in_download_upgrade[$u] ) then FileMove (@ScriptDir &"\" & $aFile_in_download_upgrade[$u],@ScriptDir &"\bak\" & $today &"_"&$aFile_in_download_upgrade[$u],9)
			ConsoleWrite(@CRLF & "Copy :" & $base_dir &"\" & $aFile_in_download_upgrade[$u]  &" to " & @ScriptDir &"\" & $aFile_in_download_upgrade[$u] & @CRLF)
			FileCopy ($base_dir &"\" & $aFile_in_download_upgrade[$u] , @ScriptDir &"\" & $aFile_in_download_upgrade[$u] ,9)
		Next

	Else
		MsgBox (0, "Upgrade Error", "No file to Upgrade to.")
		Exit
	EndIf

	FileMove (@ScriptDir &"\Upgrade_FileList_"& $today, @ScriptDir &"\bak" )
Else

	MsgBox (0, "Upgrade Error", "No file to Upgrade to.")

EndIf

if FileExists (@ScriptDir &"\test.txt") then
	sleep (3*1000)
	if FileExists ( $base_dir &"\" & $aFile_in_download_upgrade[1] )  then
		ConsoleWrite ( "notepad "  & $base_dir &"\" & $aFile_in_download_upgrade[1] )
		run( "notepad "  & $base_dir &"\" & $aFile_in_download_upgrade[1] )
		WinWaitActive( $aFile_in_download_upgrade[1] &" - °O¨Æ¥»" )
		send ("please restart 'event_parser.exe' ")
	EndIf


Else
	sleep (10*1000)
	if FileExists (@ScriptDir &"\event_parser.exe") then
		run(@ScriptDir &"\event_parser.exe")
	EndIf

EndIf

Exit
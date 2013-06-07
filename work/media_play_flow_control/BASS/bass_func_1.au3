#include <Bass.au3>
#include <BassConstants.au3>
#include <array.au3>

$Music_file_inMain='E:\workarround\音樂\吳金黛-綠色方舟ok\01 溼地小精靈.mp3'
$Music_file_inMain="E:\workarround\音樂\寶寶的音樂花園ok\01.美麗的眼睛睡著了\01 Pour toi maman.mp3"

_my_bass_play ($Music_file_inMain)

func _my_bass_play ($Music_file)
local  $playing_state = -1
local $file = $Music_file
;Open Bass.DLL.  Required for all function calls.
_BASS_STARTUP ("BASS.dll")

;Initalize bass.  Required for most functions.
_BASS_Init(0, -1, 44100, 0, "")

;Check if bass iniated.  If not, we cannot continue.
If @error Then
	MsgBox(0, "Error", "Could not initialize audio")
	Exit
EndIf



;Create a stream from that file.
$MusicHandle = _BASS_StreamCreateFile(False, $file, 0, 0, 0)
 $aMusic= _BASS_ChannelGetInfo($MusicHandle)
 ;_ArrayDisplay ( $aMusic  )

;Check if we opened the file correctly.
If @error Then
	MsgBox(0, "Error", "Could not load audio file" & @CR & "Error = " & @error)
	Exit
EndIf

;Iniate playback
_BASS_ChannelPlay($MusicHandle, 1)
$playing_state =1

;Get the length of the song in bytes.
$song_length = _BASS_ChannelGetLength($MusicHandle, $BASS_POS_BYTE)
 ConsoleWrite ("Song Length: "  & $song_length & @CRLF)
 ConsoleWrite ( "Music length in sec:  "&  _BASS_ChannelBytes2Seconds ($MusicHandle,$song_length) & @CRLF)
 ConsoleWrite ("Playing State 1 : " & $playing_state)
 dim $stop_once=0
 dim $original_volume

While 1
	Sleep(20)
	;Get the current position in bytes
	$current = _BASS_ChannelGetPosition($MusicHandle, $BASS_POS_BYTE)
	;Calculate the percentage
	$percent = Round(($current / $song_length) * 100, 0)
	if $percent = 6 and $stop_once=0 then _BASS_ChannelSetVolume ($MusicHandle,50)
	if $percent = 8 and $stop_once=0 then _BASS_ChannelSetVolume ($MusicHandle,20)
	if $percent = 10 and $stop_once=0 then
		;$original_volume= _BASS_ChannelGetVolume($MusicHandle)
		_BASS_ChannelSetVolume ($MusicHandle,10)
		;ConsoleWrite ( @CRLF & "original_volume : " & $original_volume )

		_BASS_ChannelPause($MusicHandle)
		ConsoleWrite (@CRLF &"Sleep 3 Sec")
		sleep (3*1000)

		$stop_once=1
		$playing_state =0
		 ConsoleWrite (@CRLF& "Playing State 2 : " & $playing_state)
	EndIf


	if  $stop_once=1 and  $playing_state =0 then

		_BASS_ChannelPlay($MusicHandle, 0)
		if $percent = 10 then _BASS_ChannelSetVolume ($MusicHandle,20)
		;$stop_once=0
		$playing_state =1
		 ConsoleWrite (@CRLF& "Playing State 3 : " & $playing_state)
	 EndIf

	 if $percent = 12 then _BASS_ChannelSetVolume ($MusicHandle,50)
	 if $percent = 14 then _BASS_ChannelSetVolume ($MusicHandle,100)
		;
	;Display that to the user
	ToolTip("Completed " & $percent & "%", 0, 0)
	;If the song is complete, then exit.
	If $current >= $song_length Then ExitLoop
WEnd
EndFunc

Func OnAutoItExit()
	;Free Resources
	_BASS_Free()
EndFunc   ;==>OnAutoItExit

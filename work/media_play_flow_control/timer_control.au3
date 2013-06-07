#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
; ; 2012 12 28 這支程式是用來讀取每小時的 PlayList( *.lst ) 並且送指令到 Player 播出的
;1 每小時的 play list (在 event_parser 的 playlist merge 中己處理了)
;2 小時提醒及安全宣導 play list 這是獨立於每小時的 PlayList之外
;3 特殊節日的 PlayList
;4 同步伺服器上的新增檔案
;5 檢查同步是否完成
;6 auto upgrade
;7 Read setting for Video/Audio player dir
;8


#Include <File.au3>
#include <Date.au3>
#include <array.au3>
#include <ntp_udp123_func.au3>
#include <Bass.au3>
#include <BassConstants.au3>
;#include <gen_playlist_t1.au3>

Dim $sec = @SEC
Dim $min = @MIN
Dim $hour = @HOUR
Dim $day = @MDAY
Dim $month = @MON
Dim $year = @YEAR
Dim $today = $year & $month & $day
dim $today_in_slash=$year &"/"& $month &"/"& $day
dim $current_time
dim $current_time_with_colon= $hour &":"& $min &":"& $sec
dim $switch_playlist_time

dim $setting_file=@ScriptDir & "\timer_control.ini"
dim $burgh_setting=@ScriptDir & "\sync_delay.txt"
dim $burgh=""
dim $aSetting
Global $video_player_dir=@ScriptDir
Global $video_file_root_directory=@ScriptDir
Global $announce_file_root_dir=@ScriptDir
dim $sync_delay=0

dim $hour_start_from=0

dim $current_hourly_playlist , $next_hourly_palylist
dim $hourly_playlist

Global $aRepeatList_main
dim $aPlaylist  ; playlist in array
dim $aAnnounceList ; announce list ib array
Global $music_start_hour=10
Global $announce_start_hour="10:00"
Global $music_end_hour=22
Global $announce_end_hour=22
dim $in_playing_time=0
dim $before_playing_time=0
dim $diff_in_sec_to_start_hour=0
dim $diff_in_sec_to_end_hour=0
Global $announce_delta=15*60
global $vlc_volume=200
dim $close_announce="."

dim $timer
dim $action_exe="vlc.exe"
dim $next_event
dim $diff_in_sec=0
dim $init_process_id=0
dim $PID
dim $process_by_this_circle=0
dim $launch_player_time=":59:59"



Dim $test_mode = _TEST_MODE()

;MsgBox (0,"",  ' s '  &   $today_in_slash &" "& $music_start_hour &":00:00 " & _NowCalc()   )

;MsgBox (0,"",  ( _DateDiff('s' , _NowCalc() ,   $today_in_slash &" "& $music_start_hour &":00:00"  ) ) )
;MsgBox (0,"",  ( _DateDiff('s'  , _NowCalc() ,   $today_in_slash &" "& $music_end_hour &":00:00" )  ) )


;if ( _DateDiff('s'   , _NowCalc(),   $today_in_slash &" "& $music_end_hour &":00:00"  ) < 0)  then

;	;$before_playing_time =_DateDiff('n' , _NowCalc() ,   $today_in_slash &" "& $music_start_hour &":00:00" )
;	MsgBox(0,"Now time", "Now is after play_time"  )
;	Exit
;EndIf

if ( _DateDiff('s' , _NowCalc() ,   $today_in_slash &" "& $music_start_hour &":00:00" ) > 0)  then

	$before_playing_time =_DateDiff('n' , _NowCalc() ,   $today_in_slash &" "& $music_start_hour &":00:00" )
	MsgBox(0,"Now time", "Now is before play_time: " &  $before_playing_time )
EndIf



if ( _DateDiff('s' , _NowCalc() ,   $today_in_slash &" "& $music_start_hour &":00:00" ) < 0)  and ( _DateDiff('s'   , _NowCalc(),   $today_in_slash &" "& $music_end_hour &":00:00"  ) > 0) then
	$in_playing_time=1
EndIf

;MsgBox(0,"Now time", "Now is before play_time? " &  $before_playing_time  & ".  in_playing_time: "&$in_playing_time)



;$in_playing_time=1
if $in_playing_time=0 then
	MsgBox(0,"請稍待", "播音時間為 : " & $today_in_slash &" "& $music_start_hour &":00:00 之後以及" & @CRLF &  $today_in_slash &" "& $music_end_hour &":00:00 之前"   )
	exit
EndIf

if FileExists ($burgh_setting ) then $burgh=FileReadLine ($burgh_setting,2)

if FileExists ($setting_file ) then

	Local $aSetting = IniReadSection( $setting_file, $burgh)

		;_ArrayDisplay($aSetting)

	if IsArray($aSetting) then
		;_ArrayDisplay($aSetting)

		for $s=1 to $aSetting[0][0]
				;MsgBox(0," aSetting", "Key: " & $aSetting[$s][0] & @CRLF & "Value: " & $aSetting[$s][1])
				if  StringInStr ($aSetting[$s][0],"sync delay parameter") then $sync_delay= $aSetting[$s][1]
				if  StringInStr ($aSetting[$s][0],"Video player directory") then $video_player_dir= $aSetting[$s][1]
				if  StringInStr ($aSetting[$s][0],"video file root directory") then $video_file_root_directory=$aSetting[$s][1]
				if  StringInStr ($aSetting[$s][0],"announce file root directory") then $announce_file_root_dir=$aSetting[$s][1]
				if  StringInStr ($aSetting[$s][0],"close_announce") then $close_announce=$aSetting[$s][1]


		Next
	EndIf
	;MsgBox(0,"Setting",  $burgh &" ; " & $video_player_dir &" ; "& $video_file_root_directory &" ; "& $sync_delay )
	;MsgBox (0,"Setting ", "announce file root directory : " & $announce_file_root_dir )
EndIf



dim $m_music_root=$video_file_root_directory ;"E:\workarround\音樂\ready\"
dim $m_play_album= "";"1002;3001;3005"

$m_play_album= _read_play_album ()
if $m_play_album<>0 then
	$a = _gen_playlist( $m_music_root , $m_play_album  )
	;_ArrayDisplay ( $a  )
EndIf




		$hour_start_from= $hour
		;$hourly_playlist=_FileListToArray (@ScriptDir &"\","*.lst",1)
		 ;_ArrayDisplay ($hourly_playlist)
		_FileReadToArray (@ScriptDir &"\playlist.txt", $aPlaylist)

		if IsArray( $aPlaylist ) Then
			$delta=0
			$music_start_hour=  StringReplace( StringReplace ( $aPlaylist[2] ,"<#","") ,">","")
			$music_end_hour=  StringReplace( StringReplace ( $aPlaylist[ $aPlaylist[0] ] ,"</#","") ,">","")
			;MsgBox (0,"Music start from to : " ,  $music_start_hour & " >>>> "& $music_end_hour )
			_ArrayDelete ( $aPlaylist , $aPlaylist[0] ) ;  delete line of start time
			_ArrayDelete ( $aPlaylist , 2 ) 			;  delete line of start time
			_ArrayDelete ( $aPlaylist , 1 ) 			;  delete line of start time
			$aPlaylist[0]=  $aPlaylist[0]-3

			for $x=$aPlaylist[0] to 1 step  -1
				if $aPlaylist[$x]="" then
					_ArrayDelete($aPlaylist, $x)
					$aPlaylist[0]=  $aPlaylist[0]-1
				EndIf
			Next

		EndIf
		;_ArrayDisplay ($aPlaylist)

		_FileReadToArray (@ScriptDir &"\announcelist.txt", $aAnnounceList)
		$aRepeatList_main = _cal_announcelist($aAnnounceList)
		;_ArrayDisplay ($aRepeatList_main)
;================================================================================
		; _vlc_play ($aRepeatList_main[1])

		;if IsArray($aRepeatList_main) then
		;	ConsoleWrite ( @CRLF & $video_player_dir &" " & $announce_file_root_dir &"\"& $aRepeatList_main[1] )
		;	$aRepeatList_main[1]="29 堡內上課須知.mp3"
		;	Run(@ComSpec & " /c " & $video_player_dir &'  --volume=350 --play-and-exit "'& $announce_file_root_dir &'\'& $aRepeatList_main[1], "", @SW_HIDE)
		;	sleep(3*1000)
		;	$init_process_id =ProcessWait($action_exe )
		;	if $init_process_id  then
		;		ConsoleWrite ( @CRLF&"Current time: " &  $current_time & ".  $init_process_id: " & $init_process_id  )
		;		ConsoleWrite ( @CRLF & " $init_process_id: " & $init_process_id  )
		;	EndIf
		;EndIf

		;if IsArray( $aAnnounceList ) Then
		;	$announce_start_hour=  StringReplace( StringReplace ( $aAnnounceList[2] ,"<#","") ,">","")
		;	$announce_end_hour=  StringReplace( StringReplace ( $aAnnounceList[ $aAnnounceList[0] ] ,"</#","") ,">","")
		;
		;	_ArrayDelete ( $aAnnounceList , $aAnnounceList[0] ) ;  delete line of start time
		;	_ArrayDelete ( $aAnnounceList , 2 ) 			;  delete line of start time
		;	_ArrayDelete ( $aAnnounceList , 1 ) 			;  delete line of start time
		;	$aAnnounceList[0]=  $aAnnounceList[0]-3
		;	for $x=$aAnnounceList[0] to 1 step  -1
		;		if $aAnnounceList[$x]="" then
		;			_ArrayDelete($aAnnounceList, $x)
		;			$aAnnounceList[0]=  $aAnnounceList[0]-1
		;		EndIf
		;	Next
		;EndIf
		;_ArrayDisplay($aAnnounceList )
		;MsgBox(0,"", _ArraySearch( $hourly_playlist, "15.lst" ,0,0,1) )
		;Exit
		;MsgBox(0,"Play Music", $video_file_root_directory &"\"& $aPlaylist[3] )
		;Exit
		;if FileExists ($video_file_root_directory &"\" & $aPlaylist[3] ) then  MsgBox(0,"Play Music", $video_file_root_directory &"\"& $aPlaylist[3] )
;=======================================

local  $playing_state = -1
local $file
;Open Bass.DLL.  Required for all function calls.
_BASS_STARTUP ("BASS.dll")

;Initalize bass.  Required for most functions.
_BASS_Init(0, -1, 44100, 0, "")

;Check if bass iniated.  If not, we cannot continue.
If @error Then
	MsgBox(0, "Error", "Could not initialize audio")
	Exit
EndIf


dim $start_announcelist_index=0 ; next event by id not time
dim $next_announcelist_index=0
dim $next_announcelist_time=0
local $sec_to_next_event=0

while 1
	 $sec = @SEC
	 $min = @MIN
	 $hour = @HOUR

	 $current_time = $hour & $min & $sec
	 $current_time_with_colon= $hour &":"& $min &":"& $sec

	;ConsoleWrite ("Video player dir: " & $video_player_dir & @CRLF)
	;ConsoleWrite ( @CRLF & $current_time_with_colon & @CRLF)
	;$diff_in_sec_to_start_hour = _DateDiff('s' , _NowCalc(), $today_in_slash &" " & $music_start_hour &":00:00"  )
	;$diff_in_sec_to_end_hour =   _DateDiff('s' , _NowCalc(), $today_in_slash &" " & $music_end_hour &":00:00"  )
	;MsgBox ( 1, "Time diff" , "$diff_in_sec_to_start_hour : " & $diff_in_sec_to_start_hour & @CRLF & "$diff_in_sec_to_end_hour: " &  ($diff_in_sec_to_end_hour /60) & " "& mod($diff_in_sec_to_end_hour,60 ) )
	;ConsoleWrite ( @CRLF & $diff_in_sec_to_start_hour & @CRLF  & $diff_in_sec_to_end_hour & @CRLF)
	;
	;這是電腦起始時間點
	if $start_announcelist_index=0  then
		for $d= 1 to $aRepeatList_main[0][0]

			if  $start_announcelist_index=0 and ( _DateDiff('s' , _NowCalc() ,  $aRepeatList_main[$d][0] ) > 0) then
				$start_announcelist_index= $d
				$next_announcelist_index = $d
				$next_announcelist_time = $aRepeatList_main[$d][0]
				$sec_to_next_event = _DateDiff('s' , _NowCalc() ,  $next_announcelist_time )
				;MsgBox(0,"next event in sec : " &$next_announcelist_time , $sec_to_next_event &" Sec to : " & $aRepeatList_main[$d][0] &" ; "& $aRepeatList_main[$d][1] )
			EndIf
		Next
	EndIf
	;MsgBox(0,"next event in sec: " & $next_announcelist_time , $sec_to_next_event &" Sec to : " & $aRepeatList_main[$next_announcelist_index][0] &" ; "& $aRepeatList_main[$next_announcelist_index][1] )


	ConsoleWrite ( @CRLF & "Music start from : " & $music_start_hour &@CRLF&"Next event "&$next_event &" will on in " & $sec_to_next_event  &" Sec" & @CRLF)
	;_ArrayDisplay ($aRepeatList_main )
	;MsgBox(0,"Next event", "Music start from : " & $music_start_hour &@CRLF&"Next event "&$next_event &" will on in " & $sec_to_next_event  &" Sec")
	;MsgBox (0,"Time diff to start hour",  $today_in_slash &" " & $music_start_hour & @CRLF & $diff_in_sec )


	for $x=1 to $aPlaylist[0]
			;ConsoleWrite (@CRLF & "Now Playing : "& $video_file_root_directory &"\"& $aPlaylist[$x] &@CRLF )
			;MsgBox(0,"Play Music",  $video_file_root_directory &"\"& $aPlaylist[$x] , 3)
			ConsoleWrite (@CRLF & "Now Playing : "& $aPlaylist[$x] &@CRLF )
			MsgBox(0,"Play Music",   $aPlaylist[$x] , 3)

		 ;_my_bass_play ( $video_file_root_directory &"\"& $aPlaylist[$x] )

		 ;$MusicHandle = _BASS_StreamCreateFile(False, $video_file_root_directory &"\"& $aPlaylist[$x] , 0, 0, 0)
		 $MusicHandle = _BASS_StreamCreateFile(False,  $aPlaylist[$x] , 0, 0, 0)
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
			ConsoleWrite ( @CRLF& "Song Length: "  & $song_length & @CRLF)
			ConsoleWrite ( @CRLF&"Music length in sec:  "&  _BASS_ChannelBytes2Seconds ($MusicHandle,$song_length) & @CRLF)
			ConsoleWrite ( @CRLF&"Playing State 1 : " & $playing_state & @CRLF)
			dim $stop_once=0
			dim $original_volume

			while 1
				Sleep(200)
				;Get the current position in bytes
				$current = _BASS_ChannelGetPosition($MusicHandle, $BASS_POS_BYTE)
				;Calculate the percentage
				$percent = Round(($current / $song_length) * 100, 0)

				$sec_to_next_event = _DateDiff('s' , _NowCalc() ,  $next_announcelist_time )
				ConsoleWrite ($sec_to_next_event & "  Next announce at : "& $next_announcelist_time &" Next announce event: "& @CRLF)
				if $sec_to_next_event > 30 then  sleep ( 3000 )

				if $sec_to_next_event <=30 then
					$original_volume= _BASS_ChannelGetVolume($MusicHandle)
					sleep (50)

					if $sec_to_next_event <=8 and $sec_to_next_event > 5 then  _BASS_ChannelSetVolume ($MusicHandle,50)

					if $sec_to_next_event <=5 and $sec_to_next_event > 1 then _BASS_ChannelSetVolume ($MusicHandle,20)

					if $sec_to_next_event <=1  then
						_BASS_ChannelPause($MusicHandle)
						$stop_once=1
						$playing_state =0
						ConsoleWrite (@CRLF& "Playing State 2 : " & $playing_state)

						if _vlc_play( $aRepeatList_main[$next_announcelist_index][1] ) then

							$next_announcelist_index= $next_announcelist_index+1
							$next_announcelist_time = $aRepeatList_main[$next_announcelist_index][0]
							$sec_to_next_event = _DateDiff('s' , _NowCalc() ,  $next_announcelist_time )
							ConsoleWrite ( @CRLF & $sec_to_next_event &";;"& $next_announcelist_time & @CRLF)
						EndIf
					EndIf
				EndIf

				if ProcessExists ("vlc.exe") then ProcessClose("vlc.exe")

				if  $stop_once=1 and  $playing_state =0 then

					_BASS_ChannelPlay($MusicHandle, 0)
					sleep(2000)
					_BASS_ChannelSetVolume ($MusicHandle,20)
					sleep(2000)
					_BASS_ChannelSetVolume ($MusicHandle,50)
					sleep(2000)
					_BASS_ChannelSetVolume ($MusicHandle, 100)
					;$stop_once=0
					$playing_state =1
					 ConsoleWrite (@CRLF& "Playing State 3 : " & $playing_state)
				 EndIf

				;if $percent = 6 and $stop_once=0 then _BASS_ChannelSetVolume ($MusicHandle,50)
				;if $percent = 8 and $stop_once=0 then _BASS_ChannelSetVolume ($MusicHandle,20)
				;if $percent = 10 and $stop_once=0 then
				;	;$original_volume= _BASS_ChannelGetVolume($MusicHandle)
				;	_BASS_ChannelSetVolume ($MusicHandle,10)
				;	;ConsoleWrite ( @CRLF & "original_volume : " & $original_volume )

				;	_BASS_ChannelPause($MusicHandle)
				;	ConsoleWrite (@CRLF &"Sleep 3 Sec")
				;	sleep (3*1000)

				;	$stop_once=1
				;	$playing_state =0
				;	 ConsoleWrite (@CRLF& "Playing State 2 : " & $playing_state)
				;EndIf
				;if  $stop_once=1 and  $playing_state =0 then
				;
				;	_BASS_ChannelPlay($MusicHandle, 0)
				;	if $percent = 10 then _BASS_ChannelSetVolume ($MusicHandle,20)
				;	;$stop_once=0
				;	$playing_state =1
				;	 ConsoleWrite (@CRLF& "Playing State 3 : " & $playing_state)
				; EndIf

				; if $percent = 12 then _BASS_ChannelSetVolume ($MusicHandle,50)
				; if $percent = 14 then _BASS_ChannelSetVolume ($MusicHandle,100)
				;
				;Display that to the user
				if  mod ($percent,10 )=0 then  TrayTip( "Play percentage:"& $percent & "%",$aPlaylist[$x] & @CRLF &"Next play: " & $aPlaylist[$x+1], 3 )
				;If the song is complete, then exit.
				If $current >= $song_length Then ExitLoop

		WEnd
		$stop_once=0
		if $x= $aPlaylist[0] then $x=1
	Next




;=============================================================
	$init_process_id=1
	if $init_process_id=0 then  ;; Just start this exe to run first playlist


			if ProcessExists ( $action_exe ) then
				$PID = ProcessExists( $action_exe )
				If $PID Then ProcessClose( $PID )
			EndIf

		;Run(@ComSpec & " /c " & $video_player_dir &" " & @ScriptDir &"\"&$current_hourly_playlist, "", @SW_HIDE)
		;sleep(3*1000)
		;$init_process_id =ProcessWait($action_exe )
		;if $init_process_id  then
		;	;ConsoleWrite ( @CRLF&"Current time: " &  $current_time & ".  $init_process_id: " & $init_process_id  )
		;	;ConsoleWrite ( @CRLF & " $init_process_id: " & $init_process_id  )
		;EndIf


	Else

		$diff_in_sec =_DateDiff('s' , _NowCalc(), $today_in_slash &" " &$next_event  )

		;MsgBox(0,"$diff_in_sec : ", "$diff_in_sec : " & $diff_in_sec,2)
		if $diff_in_sec >= 60 then
			$process_by_this_circle=0

				ConsoleWrite ( @CRLF &"Current time: " &  $current_time & ".  Still Need to wait :" &  $diff_in_sec  &" Sec, About " & Floor ($diff_in_sec/60)& " Min and "&  mod($diff_in_sec,60) &" Sec" &" to run "&$next_hourly_palylist )

			sleep (30*1000)
		EndIf

		if $diff_in_sec <59 then
			if $diff_in_sec > 5 then sleep (1000)
			if $diff_in_sec <=5 then sleep(500)
			ConsoleWrite ( @CRLF &"Current time: " &  $current_time & ".  Still Need to wait :" &  $diff_in_sec  &" Sec, About " & Floor ($diff_in_sec/60)& " Min and "&  mod($diff_in_sec,60) &" Sec" &" to run "&$next_hourly_palylist )
			if  $diff_in_sec <= 2 then

				;$current_hourly_playlist= $hourly_playlist[$timer+1]
				if $process_by_this_circle=0 and ProcessExists ( $init_process_id ) then ProcessClose( $init_process_id )
				if $process_by_this_circle=0 then Run(@ComSpec & " /c " & $video_player_dir &" " & @ScriptDir &"\"&$next_hourly_palylist, "", @SW_HIDE)
				$init_process_id =ProcessWait($action_exe )

				ConsoleWrite (@CRLF & "New Process started at "  & _NowCalc() & " by  $init_process_id = "& $init_process_id )
				if $init_process_id then $process_by_this_circle=1


			EndIf

		EndIf

	if  $diff_in_sec_to_start_hour <0 and $diff_in_sec_to_end_hour < 0 then ExitLoop
	EndIf


;=============================================================



WEnd

Exit


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
	;if $percent = 6 and $stop_once=0 then _BASS_ChannelSetVolume ($MusicHandle,50)
	;if $percent = 8 and $stop_once=0 then _BASS_ChannelSetVolume ($MusicHandle,20)
	;if $percent = 10 and $stop_once=0 then
	;	;$original_volume= _BASS_ChannelGetVolume($MusicHandle)
	;	_BASS_ChannelSetVolume ($MusicHandle,10)
	;	;ConsoleWrite ( @CRLF & "original_volume : " & $original_volume )

	;	_BASS_ChannelPause($MusicHandle)
	;	ConsoleWrite (@CRLF &"Sleep 3 Sec")
	;	sleep (3*1000)

	;	$stop_once=1
	;	$playing_state =0
	;	 ConsoleWrite (@CRLF& "Playing State 2 : " & $playing_state)
	;EndIf


	;if  $stop_once=1 and  $playing_state =0 then
    ;
	;	_BASS_ChannelPlay($MusicHandle, 0)
	;	if $percent = 10 then _BASS_ChannelSetVolume ($MusicHandle,20)
	;	;$stop_once=0
	;	$playing_state =1
	;	 ConsoleWrite (@CRLF& "Playing State 3 : " & $playing_state)
	; EndIf

	; if $percent = 12 then _BASS_ChannelSetVolume ($MusicHandle,50)
	; if $percent = 14 then _BASS_ChannelSetVolume ($MusicHandle,100)
		;
	;Display that to the user
	ToolTip("Completed " & $percent & "%", 0, 0)
	;If the song is complete, then exit.
	If $current >= $song_length Then ExitLoop
WEnd
	_BASS_stop_it()

EndFunc

Func _BASS_stop_it()
	_BASS_ChannelSetVolume ($MusicHandle,5)
	_BASS_Free()
EndFunc

Func OnAutoItExit()
;	;Free Resources
	_BASS_ChannelSetVolume ($MusicHandle,20)
	_BASS_Free()
EndFunc   ;==>OnAutoItExit


func _cal_announcelist($announce_list_inFunc)
	local $cal , $1_lead , $1_end , $2_lead, $2_end, $3_lead, $3_end, $4_lead, $4_end , $ann_start, $ann_end , $ann_h_diff
	Local $aRepeatList[1]
	$aRepeatList[0]=0
	$1_lead =0
	$1_end =0
	$2_lead=0
	$2_end=0
	$3_lead=0
	$3_end=0
	$4_lead=0
	$4_end=0



	;ConsoleWrite ($1_lead &" ... "  & $1_end & @CRLF  & $2_lead&" ... "  & $2_end& @CRLF & $3_lead&" ... "  & $3_end& @CRLF  & $4_lead&" ... "  & $4_end )

	;MsgBox(0,"Status", "Now is in _cal_announcelist function start")

	for $cal=1 to $announce_list_inFunc[0]
		if $1_lead=0 and StringInStr ( $announce_list_inFunc[$cal] ,"<1#" ) then  $1_lead= $cal
		if $1_end=0 and  StringInStr ( $announce_list_inFunc[$cal] ,"<1/#" ) then  $1_end= $cal

		if $1_lead <>0  then
			if  StringInStr ( $announce_list_inFunc[$cal] ,"<2#" ) then  $2_lead= $cal
		EndIf
		if $1_end <>0  then
			if  StringInStr ( $announce_list_inFunc[$cal] ,"<2/#" ) then  $2_end= $cal
		EndIf

		if $1_lead >0 and $2_lead >0 and $3_lead=0  and $cal > $2_end  then
			if  StringInStr ( $announce_list_inFunc[$cal] ,"<3#" ) then  $3_lead= $cal
		EndIf

		if $1_end >0 and $2_end >0 and $3_end=0  and $cal > $3_lead  then
			if  StringInStr ( $announce_list_inFunc[$cal] ,"<3/#" ) then  $3_end= $cal
		EndIf

		if $1_lead >0 and $2_lead >0 and $3_lead >0 and $4_lead=0 and $cal > $3_end  then
			if  StringInStr ( $announce_list_inFunc[$cal] ,"<4#" ) then  $4_lead= $cal
		EndIf
		if  $1_end >0 and $2_end >0 and $3_end >0 and $4_end=0  and $cal > $4_lead  then
			if  StringInStr ( $announce_list_inFunc[$cal] ,"<4/#" ) then  $4_end= $cal
		EndIf

	next

	ConsoleWrite ("group1 Line No. " & $1_lead &" ... "  & $1_end & @CRLF & "group2 Line No. "  & $2_lead&" ... "  & $2_end& @CRLF & "group3 Line No. "& $3_lead&" ... "  & $3_end& @CRLF & "group4 Line No. " & $4_lead&" ... "  & $4_end )
	;MsgBox(0,"Start and end position", $1_lead &" ... "  & $1_end & @CRLF  & $2_lead&" ... "  & $2_end& @CRLF & $3_lead&" ... "  & $3_end& @CRLF  & $4_lead&" ... "  & $4_end )
	;_ArrayDisplay ($announce_list_inFunc)
	;MsgBox (0,"1  $announce_start_hour", $announce_start_hour)
	if $1_lead>0 and $1_end >0 Then
		$ann_start= StringReplace ( StringReplace($announce_list_inFunc[$1_lead] ,"<1#","")  ,">","")
		$announce_start_hour= $ann_start
		$ann_end = StringReplace ( StringReplace( $announce_list_inFunc[$1_end] ,"<1/#","") ,">","")
		$ann_h_diff= _DateDiff('h' , $today_in_slash &" " & $ann_start&":00", $today_in_slash &" " & $ann_end&":00"  )
		;MsgBox (0,"2  $announce_start_hour", $announce_start_hour)
		if $ann_h_diff >=1 then

			 for $z=1 to $ann_h_diff

				for $y=$1_lead+1  to $1_end-1
				_arrayadd( $aRepeatList , $announce_list_inFunc[$y] )
				Next
				$aRepeatList[0]=$aRepeatList[0] + ($1_end- $1_lead-1)
				;_ArrayDisplay( $aRepeatList)
				if $z=1 then $announce_delta=( 3600/$aRepeatList[0] )
				;MsgBox (0,"delta", $announce_delta)
			 Next
			ConsoleWrite (@CRLF & "aRepeatList : $ann_start ->" &$ann_start & ".  aRepeatList : $ann_end ->" & $ann_end &".  aRepeatList : $ann_h_diff ->"& $ann_h_diff &".  And Array size is: " & $ann_h_diff &" x "  &($1_end- $1_lead-1) & @CRLF )
			;_ArrayDisplay( $aRepeatList)
		EndIf
		;MsgBox (0,"diff ", _DateDiff('h' , $today_in_slash &" " & $ann_start, $today_in_slash &" " & $ann_end  ))
	EndIf


	if $2_lead>0 and $2_end >0 Then
		$ann_start= StringReplace ( StringReplace($announce_list_inFunc[$2_lead] ,"<2#","")  ,">","")
		$ann_end = StringReplace ( StringReplace( $announce_list_inFunc[$2_end] ,"<2/#","") ,">","")
		$ann_h_diff= _DateDiff('h' , $today_in_slash &" " & $ann_start, $today_in_slash &" " & $ann_end  )

		if $ann_h_diff >=1 then

			 for $z=1 to $ann_h_diff

				for $y=$2_lead+1  to $2_end-1
				_arrayadd( $aRepeatList , $announce_list_inFunc[$y] )
				Next
				$aRepeatList[0]=$aRepeatList[0] + ($2_end- $2_lead-1)
				;_ArrayDisplay( $aRepeatList)
			 Next
			ConsoleWrite ("aRepeatList repeat : " & $ann_h_diff &" times. And Array size is: " & $ann_h_diff &" x "  &($2_end- $2_lead-1) )
			;_ArrayDisplay( $aRepeatList)
		EndIf
		;MsgBox (0,"diff ", _DateDiff('h' , $today_in_slash &" " & $ann_start, $today_in_slash &" " & $ann_end  ))
	EndIf


	if $3_lead>0 and $3_end >0 Then
		$ann_start= StringReplace ( StringReplace($announce_list_inFunc[$3_lead] ,"<3#","")  ,">","")
		$ann_end = StringReplace ( StringReplace( $announce_list_inFunc[$3_end] ,"<3/#","") ,">","")
		$ann_h_diff= _DateDiff('h' , $today_in_slash &" " & $ann_start, $today_in_slash &" " & $ann_end  )

		if $ann_h_diff >=1 then

			 for $z=1 to $ann_h_diff

				for $y=$3_lead+1  to $3_end-1
				_arrayadd( $aRepeatList , $announce_list_inFunc[$y] )
				Next
				$aRepeatList[0]=$aRepeatList[0] + ($3_end- $3_lead-1)
				;_ArrayDisplay( $aRepeatList)
			 Next
			ConsoleWrite ("aRepeatList repeat : " & $ann_h_diff &" times. And Array size is: " & $ann_h_diff &" x "  &($3_end- $3_lead-1) )
			;_ArrayDisplay( $aRepeatList)
		EndIf
		;MsgBox (0,"diff ", _DateDiff('h' , $today_in_slash &" " & $ann_start, $today_in_slash &" " & $ann_end  ))
	EndIf


	if $4_lead>0 and $4_end >0 Then
		$ann_start= StringReplace ( StringReplace($announce_list_inFunc[$4_lead] ,"<4#","")  ,">","")
		$ann_end = StringReplace ( StringReplace( $announce_list_inFunc[$4_end] ,"<4/#","") ,">","")
		$ann_h_diff= _DateDiff('h' , $today_in_slash &" " & $ann_start, $today_in_slash &" " & $ann_end  )

		if $ann_h_diff >=1 then

			 for $z=1 to $ann_h_diff

				for $y=$4_lead+1  to $4_end-1
				_arrayadd( $aRepeatList , $announce_list_inFunc[$y] )
				Next
				$aRepeatList[0]=$aRepeatList[0] + ($4_end- $4_lead-1)
				;_ArrayDisplay( $aRepeatList)
			 Next
			ConsoleWrite ("aRepeatList repeat : " & $ann_h_diff &" times. And Array size is: " & $ann_h_diff &" x "  &($4_end- $4_lead-1) )
			;_ArrayDisplay( $aRepeatList)
		EndIf
		;MsgBox (0,"diff ", _DateDiff('h' , $today_in_slash &" " & $ann_start, $today_in_slash &" " & $ann_end  ))
	EndIf


	;_ArrayDisplay( $aRepeatList)
	;ConsoleWrite (@CRLF & $announce_delta)
	local $2d_announce_list[$aRepeatList[0]+2][2]  ; $aRepeatList[0]+1 is for arraysize cell, $aRepeatList[0]+2 is for add 10:00 play music
	local $hour_announce
	;_ArrayDisplay ($2d_announce_list)

	for $d=1 to $aRepeatList[0]
		;MsgBox (0,"date add" ,   's , '& ($d-1)*$announce_delta &" , " &$today_in_slash&" "& $ann_start &":00:00  ~ "  & _DateAdd ('s' , ($d-1)*$announce_delta, $today_in_slash&" "& $ann_start &":00:00" ) )
		;$2d_announce_list[$d][0]=  _DateAdd ('s', ($d-1)*$announce_delta, $today_in_slash&" "&$ann_start &":00" )

		$2d_announce_list[$d][0]=  _DateAdd ('s', ($d-1)*$announce_delta, $today_in_slash&" "&$announce_start_hour &":00:00" )
		if mod($d-1,4)=0 then
			;MsgBox(0,"$announce_start_hour " & $d-1, $announce_start_hour &".  '" &  StringLeft( StringReplace( StringReplace( _DateAdd ('h', ($d-1)/4 , $today_in_slash&" "&$announce_start_hour &":00") ,$today_in_slash,"") ,":","") ,5 )&"'"   )
			$hour_announce= StringStripWS( StringLeft( StringReplace( StringReplace( _DateAdd ('h', ($d-1)/4 , $today_in_slash&" "&$announce_start_hour &":00") ,$today_in_slash,"") ,":","") ,5 ) ,8 )

			$2d_announce_list[$d][1]= $hour_announce&".mp3;"&$aRepeatList[$d]
		Else
			$2d_announce_list[$d][1]= $aRepeatList[$d]
		EndIf

	Next
	$2d_announce_list[$aRepeatList[0]+1][0]=  _DateAdd ('s', ($d-1)*$announce_delta, $today_in_slash&" "&$announce_start_hour &":00:00" )
	$2d_announce_list[$aRepeatList[0]+1][1]= "2200.mp3;" &$close_announce


	$2d_announce_list[0][0]=$aRepeatList[0]+1
	;_ArrayAdd()
	;_ArrayDisplay ($2d_announce_list)
	;MsgBox(0,"Status", "Now is in _cal_announcelist function end  ")
	;Exit
	Return $2d_announce_list
EndFunc


func _vlc_play($announce_list_vlc)
	local  $this_time_announce_list , $this_vlc_id, $winwait_vlc_pid
	local $scan_urgent_once, $scan_urgent_always

	ConsoleWrite(@CRLF & $announce_list_vlc)

	if StringInStr ( $announce_list_vlc ,";" ) then
		$this_time_announce_list= StringSplit ( $announce_list_vlc,";")
	Else
		$this_time_announce_list=$announce_list_vlc
	EndIf


		;以下這段是為了即時掃瞄檔案夾內是不是有臨時宣佈
			$scan_urgent_once=_FileListToArray ( $announce_file_root_dir & "\urgent_announce\once", "*.mp3" )
			$scan_urgent_always= _FileListToArray ( $announce_file_root_dir & "\urgent_announce\always", "*.mp3" )

			if IsArray( $scan_urgent_always ) then

				if IsArray( $this_time_announce_list ) then
					$this_time_announce_list[0]=$this_time_announce_list[0]+$scan_urgent_always[0]
					_ArrayDelete($scan_urgent_always,0)
					_ArrayConcatenate ( $this_time_announce_list , $scan_urgent_always )
				Else
					_ArrayDisplay($scan_urgent_always)
					_ArrayInsert( $scan_urgent_always,1,$this_time_announce_list)

					$scan_urgent_always[0]=$scan_urgent_always[0]+1
					;_ArrayDisplay($scan_urgent_always)

					$this_time_announce_list = $scan_urgent_always
					;_ArrayDisplay ($this_time_announce_list)
				EndIf

			EndIf

			if IsArray( $scan_urgent_once ) then

				if IsArray( $this_time_announce_list ) then
					$this_time_announce_list[0]=$this_time_announce_list[0]+$scan_urgent_once[0]
					_ArrayDelete($scan_urgent_once,0)
					_ArrayConcatenate ( $this_time_announce_list , $scan_urgent_once )
				Else
					_ArrayInsert( $scan_urgent_once,1,$this_time_announce_list)
					$scan_urgent_once[0]=$scan_urgent_once[0]+1
					$this_time_announce_list=$scan_urgent_once
					;_ArrayDisplay ($this_time_announce_list)
				EndIf
			EndIf
		;;
		;;
		;以上這段是為了即時掃瞄檔案夾內是不是有臨時宣佈



	if isarray($this_time_announce_list) then
		;_ArrayDisplay ($this_time_announce_list)
			for $v=1 to $this_time_announce_list[0]

			ConsoleWrite ( @CRLF & $video_player_dir &'  --volume='&$vlc_volume&' --play-and-exit "'& $announce_file_root_dir &'\'& $this_time_announce_list[$v]  & @CRLF)
			;$announce_list_vlc="29 堡內上課須知.mp3"
			$winwait_vlc_pid=run( @ScriptDir&"\winwait_vlc.exe","",@SW_MINIMIZE)
			$init_process_id =Run(@ComSpec & " /c " & $video_player_dir &'  --volume='&$vlc_volume&' --play-and-exit "'& $announce_file_root_dir &'\'& $this_time_announce_list[$v], "", @SW_HIDE)
			sleep(2*1000)
			;ProcessWait($action_exe )
			;if $init_process_id  then
			;	ConsoleWrite ( @CRLF&"Current time: " &  $current_time & ".  $init_process_id: " & $init_process_id  )
			;	ConsoleWrite ( @CRLF & " $init_process_id: " & $init_process_id  )
			;EndIf
			ProcessWaitClose($init_process_id)
			if ProcessExists ( "winwait_vlc.exe" ) Then ProcessClose ( "winwait_vlc.exe" )
			if ProcessExists ( $winwait_vlc_pid ) Then ProcessClose ( $winwait_vlc_pid)

			Next
	Else
			ConsoleWrite ( @CRLF & $video_player_dir &'  --volume='&$vlc_volume&' --play-and-exit "'& $announce_file_root_dir &'\'& $announce_list_vlc  & @CRLF)
			;$announce_list_vlc="29 堡內上課須知.mp3"
			$winwait_vlc_pid=run( @ScriptDir&"\winwait_vlc.exe","",@SW_MINIMIZE)
			$init_process_id=Run(@ComSpec & " /c " & $video_player_dir &'  --volume='&$vlc_volume&' --play-and-exit "'& $announce_file_root_dir &'\'& $announce_list_vlc, "", @SW_HIDE)
			sleep(3*1000)
			;ProcessWait($action_exe )
			ProcessWaitClose($init_process_id)
			if ProcessExists ( "winwait_vlc.exe" ) Then ProcessClose ( "winwait_vlc.exe" )
			if ProcessExists ( $winwait_vlc_pid ) Then ProcessClose ( $winwait_vlc_pid)

	EndIf
	if ProcessExists ($init_process_id) then ProcessClose($init_process_id)

	$scan_urgent_once=_FileListToArray ( $announce_file_root_dir & "\urgent_announce\once", "*" )
	if IsArray( $scan_urgent_once ) then
		for $x=1 to $scan_urgent_once[0]
			FileMove ( $announce_file_root_dir & "\urgent_announce\once\" & $scan_urgent_once[$x], $announce_file_root_dir & "\urgent_announce",1)
		next
	EndIf

	Return true
EndFunc




;;以下是產生 playlist 的過程
;; 先讀取 play_album.txt 的檔案，取得專輯的編號後，掃瞄檔案清單，再編為 array
;; 最後再用 arrayswap 取得兩個 ramdom number 再用 swap 的方法改變位置
func _read_play_album ()
	local $play_album_txt , $jan_start , $feb_start , $this_m_start ,$this_m_end , $weekday, $this_m_play_album
	local $weekday=_DateDayOfWeek(@WDAY, 1)
	local $this_m_play_album=0
	local $play_album_txt_date1=0
	local $playlist_txt_date1=1
	local $month = @MON
	;MsgBox (0,"month and Week ",  mod($month,2)  &" "& _DateDayOfWeek(@WDAY, 1) )
	; mod($month,2)=1 為單數月 mod($month,2)=0 為雙數月
	;if FileExists (@ScriptDir & "\play_album.txt" ) then $play_album_txt_date1=FileReadLine(@ScriptDir & "\play_album.txt",1)
	;if FileExists (@ScriptDir & "\playlist.txt" ) then $playlist_txt_date1=FileReadLine(@ScriptDir &"\playlist.txt",1)
 if $play_album_txt_date1<>$playlist_txt_date1 then

	if FileExists (@ScriptDir & "\play_album.txt" ) then
	_FileReadToArray (@ScriptDir &"\play_album.txt", $play_album_txt)
	;_ArrayDisplay ( $play_album_txt )

		for $u=0 to $play_album_txt[0]
			if StringInStr ( $play_album_txt[$u], "<#1>") then $jan_start= $u
			if StringInStr ( $play_album_txt[$u], "<#2>") then $feb_start= $u

		Next
		;MsgBox(0," jan and feb ", "jan start: " &$jan_start & "  jan end : " & $feb_start -1  & ",   Feb start: " & $feb_start  &" Feb end: " & $play_album_txt[0]  )

		if mod($month,2)=1 then
			$this_m_start= $jan_start
			$this_m_end = $feb_start -1

			;if _DateDayOfWeek(@WDAY, 1) )
			;MsgBox (0," 4" , $this_m_start  & " , "& $this_m_end  & " , "& $weekday)
			for $u= $this_m_start to $this_m_end
				if  StringInStr ($play_album_txt[$u], $weekday) then
					;MsgBox (0,"Play album list" ,StringTrimLeft($play_album_txt[$u] , StringInStr ($play_album_txt[$u], ",") )   )
					$this_m_play_album=StringTrimLeft($play_album_txt[$u] , StringInStr ($play_album_txt[$u], ",") )
				EndIf
			Next

		Else
			$this_m_start= $feb_start
			$this_m_end = $play_album_txt[0]

			for $u= $this_m_start to $this_m_end
				if  StringInStr ($play_album_txt[$u], $weekday) then
					;MsgBox (0,"Play album list" ,StringTrimLeft($play_album_txt[$u] , StringInStr ($play_album_txt[$u], ",") )   )
					$this_m_play_album=StringTrimLeft($play_album_txt[$u] , StringInStr ($play_album_txt[$u], ",") )
				EndIf
			Next
		EndIf


	EndIf

 EndIf

Return $this_m_play_album
EndFunc

func _gen_playlist( $music_root , $play_album  )

local  $music_array
$music_array= _FileListToArray ( $music_root , "*")
local  $music_array_2d[$music_array[0]+1 ][2]

local $today_play[1]
local $temp_play

local  $aPlay_album= stringsplit( $play_album, ";" )

local $this_m_playlist_txt
local $this_play_txt_start_date=""

;_ArrayDisplay ($music_array_2d)

for $x=1 to $music_array[0]
	if StringInStr ( $music_array[$x], "__" ) then
		;MsgBox (0,"string left " ,StringLeft ( $music_array[$x] , ( StringInStr ( $music_array[$x], "__" )  -1 ) ) )
		$music_array_2d[$x][1] = $music_array[$x]
		$music_array_2d[$x][0] = StringLeft ( $music_array[$x] , ( StringInStr ( $music_array[$x], "__" )  -1 ) )

	;_ArrayDisplay ($music_array_2d)

	EndIf
Next

;_ArrayDisplay ($music_array_2d)


;_ArrayDisplay ($aPlay_album)



for $y=1 to $aPlay_album[0]
	for $x=1 to $music_array[0]
		if StringInStr ( $music_array_2d[$x][0] , $aPlay_album[$y] ) then
		;MsgBox(0," ", $music_root &$music_array_2d[$x][1] )

		$temp_play= _FileListToArray ( $music_root & $music_array_2d[$x][1], "*.mp3" ,0 )


			;_ArrayDisplay($temp_play)

		if IsArray ($temp_play) then

			for $z=1 to $temp_play[0]
				$temp_play[$z]= $music_root & $music_array_2d[$x][1] & "\" & $temp_play[$z]
			Next
			_ArrayDelete ($temp_play, 0 )
			_ArrayConcatenate( $today_play , $temp_play  )


		EndIf

		;_ArrayDisplay ($today_play)

		EndIf
	Next
Next
$today_play[0]=( UBound ($today_play)-1  )

$today_play= _random_today_play ( $today_play )

;_ArrayDisplay ($today_play)


;;=====================
;
if FileExists (@ScriptDir & "\play_album.txt" ) then $this_play_txt_start_date=FileReadLine ( @ScriptDir & "\play_album.txt",1 )

ConsoleWrite ( @CRLF &" start date : " & $this_play_txt_start_date)

if $this_play_txt_start_date="" then
	$this_play_txt_start_date= "<"& $today &">"
EndIf

$this_m_playlist_txt = fileopen (@ScriptDir &"\playlist.txt" ,10)

FileWriteLine ( $this_m_playlist_txt , $this_play_txt_start_date)
FileWriteLine ( $this_m_playlist_txt , "<#09:00>")


For $x=1 to $today_play[0]
		FileWriteLine ($this_m_playlist_txt , $today_play[$x] )

Next
FileWriteLine ( $this_m_playlist_txt, "</#22:00>")


FileClose ( $this_m_playlist_txt)
sleep(50)
FileCopy ( @ScriptDir &"\playlist.txt", @ScriptDir &"\playlist"&$today&".txt")
sleep(50)
return $today_play

EndFunc

func _random_today_play( $aToday_play )
	local $ran_play[$aToday_play[0]+1]
	local $ran_1st, $ran_2nd, $this_max
	$this_max= $aToday_play[0]
	_ArrayDelete ( $aToday_play,0)
	;_ArrayDisplay ($aToday_play )

	for $r=1 to 100 ;$this_max
		$ran_1st= Random( 0, $this_max-1 ,1 )
		;sleep(10)
		$ran_2nd= Random( 0, $this_max-1 ,1 )
		;ConsoleWrite (@CRLF  &" cycle: "& $r &" ran 1st: " & $ran_1st & "  vs ran 2nd: " & $ran_2nd )
		if $ran_1st<> $ran_2nd then
		_ArraySwap( $aToday_play[$ran_1st], $aToday_play[$ran_2nd] )
		_ArrayReverse( $aToday_play )
		;_ArraySwap( $aToday_play[$ran_1st], $aToday_play[$ran_2nd] )
		EndIf
	Next
	_ArrayReverse( $aToday_play )
	_ArrayAdd ( $aToday_play, UBound ($aToday_play) )
	_ArrayReverse ( $aToday_play )

return $aToday_play

EndFunc

;單月
; mon, 1002;3001;3005
; tue, 1003;3002;3004
; wen, 1004;2004;3003
; thr, 1001;2001;3004
; fri, 1002;2002;3005
; sat, 1003;2003;3003
; sun, 1004;2004;3002
; 雙月
; mon, 1001;2001;3001
; tue, 1002;2002;3005
; wen, 1003;2003;3004
; thr, 2001;2004;3003
; fri, 1004;3001;3005
; sat, 1001;2002;3002
; sun, 1002;2003;3004

Func _TEST_MODE()
	If FileExists(@ScriptDir & "\TESTMODE.txt") Then
		$mode = FileReadLine(@ScriptDir & "\TESTMODE.txt", 1)
		If $mode = 1 Then
			;MsgBox(0, "Test mode", "This is Test mode. ", 5)

		Else
			;MsgBox(0,"Delivery mode", "This is True delivery.",5)
			$mode = 0
		EndIf

	Else
		;MsgBox(0,"Delivery mode", "This is True  delivery.",5)
		$mode = 0
	EndIf

	Return $mode
EndFunc   ;==>_TEST_MODE

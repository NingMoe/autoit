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

Dim $sec = @SEC
Dim $min = @MIN
Dim $hour = @HOUR
Dim $day = @MDAY
Dim $month = @MON
Dim $year = @YEAR
Dim $today = $year & $month & $day
dim $today_in_slash=$year &"/"& $month &"/"& $day
dim $current_time
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
dim $music_start_hour=10
dim $announce_start_hour=10
dim $music_end_hour=22
dim $announce_end_hour=22
dim $in_playing_time=0
dim $diff_in_sec_to_start_hour=0
dim $diff_in_sec_to_end_hour=0
dim $delta=0


dim $timer
dim $action_exe="vlc.exe"
dim $next_event
dim $diff_in_sec=0
dim $init_process_id=0
dim $PID
dim $process_by_this_circle=0
dim $launch_player_time=":59:59"


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

		Next
	EndIf
	;MsgBox(0,"Setting",  $burgh &" ; " & $video_player_dir &" ; "& $video_file_root_directory &" ; "& $sync_delay )
	;MsgBox (0,"Setting ", "announce file root directory : " & $announce_file_root_dir )
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

		 ;_vlc_play ($aRepeatList_main[1])

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
while 1
	 $sec = @SEC
	 $min = @MIN
	 $hour = @HOUR

	 $current_time = $hour & $min & $sec
	 $current_time_with_colon= $hour &":"& $min &":"& $sec
	;ConsoleWrite ("Video player dir: " & $video_player_dir & @CRLF)
	ConsoleWrite ($current_time_with_colon )
	$diff_in_sec_to_start_hour = _DateDiff('s' , _NowCalc(), $today_in_slash &" " & $music_start_hour  )
	$diff_in_sec_to_end_hour =   _DateDiff('s' , _NowCalc(), $today_in_slash &" " & $music_end_hour  )

	;MsgBox (0,"Time diff to start hour",  $today_in_slash &" " & $music_start_hour & @CRLF & $diff_in_sec )
	if  $diff_in_sec_to_start_hour <=0 and $diff_in_sec_to_end_hour >=0 then
		$in_playing_time=1
		MsgBox (0,"Time diff to start and end hour", $diff_in_sec_to_start_hour  & @CRLF  & $diff_in_sec_to_end_hour )
	Else
		$in_playing_time=0
	EndIf
		;$in_playing_time=1
	if $in_playing_time then
		for $x=1 to $aPlaylist[0]
			ConsoleWrite ("Now Playing : "& $video_file_root_directory &"\"& $aPlaylist[$x])
		 _my_bass_play ( $video_file_root_directory &"\"& $aPlaylist[$x] )


		Next
	EndIf



;=============================================================
		$init_process_id=1
	if $init_process_id=0 then  ;; Just start this exe to run first playlist


			if ProcessExists ( $action_exe ) then
				$PID = ProcessExists( $action_exe )
				If $PID Then ProcessClose( $PID )
			EndIf

		Run(@ComSpec & " /c " & $video_player_dir &" " & @ScriptDir &"\"&$current_hourly_playlist, "", @SW_HIDE)
		sleep(3*1000)
		$init_process_id =ProcessWait($action_exe )
		if $init_process_id  then
			;ConsoleWrite ( @CRLF&"Current time: " &  $current_time & ".  $init_process_id: " & $init_process_id  )
			;ConsoleWrite ( @CRLF & " $init_process_id: " & $init_process_id  )
		EndIf


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


	EndIf


;=============================================================

if  $diff_in_sec_to_start_hour <0 and $diff_in_sec_to_end_hour < 0 then ExitLoop

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
	_BASS_Free()

EndFunc

;Func _BASS_stop_it()
;;	;Free Resources
;	_BASS_Free()
;EndFunc   ;==>OnAutoItExit


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

	ConsoleWrite ($1_lead &" ... "  & $1_end & @CRLF  & $2_lead&" ... "  & $2_end& @CRLF & $3_lead&" ... "  & $3_end& @CRLF  & $4_lead&" ... "  & $4_end )
	;MsgBox(0,"Start and end position", $1_lead &" ... "  & $1_end & @CRLF  & $2_lead&" ... "  & $2_end& @CRLF & $3_lead&" ... "  & $3_end& @CRLF  & $4_lead&" ... "  & $4_end )
	;_ArrayDisplay ($announce_list_inFunc)

	if $1_lead>0 and $1_end >0 Then
		$ann_start= StringReplace ( StringReplace($announce_list_inFunc[$1_lead] ,"<1#","")  ,">","")
		$ann_end = StringReplace ( StringReplace( $announce_list_inFunc[$1_end] ,"<1/#","") ,">","")
		$ann_h_diff= _DateDiff('h' , $today_in_slash &" " & $ann_start, $today_in_slash &" " & $ann_end  )

		if $ann_h_diff >=1 then

			 for $z=1 to $ann_h_diff

				for $y=$1_lead+1  to $1_end-1
				_arrayadd( $aRepeatList , $announce_list_inFunc[$y] )
				Next
				$aRepeatList[0]=$aRepeatList[0] + ($1_end- $1_lead-1)
				;_ArrayDisplay( $aRepeatList)
			 Next
			ConsoleWrite ("aRepeatList repeat : " & $ann_h_diff &" times. And Array size is: " & $ann_h_diff &" x "  &($1_end- $1_lead-1) )
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



	;MsgBox(0,"Status", "Now is in _cal_announcelist function end  ")
	;Exit
	Return $aRepeatList
EndFunc


func _vlc_play($announce_list_vlc)
	local  $this_time_announce_list
		ConsoleWrite(@CRLF & $announce_list_vlc)
	if StringInStr ( $announce_list_vlc ,";" ) then $this_time_announce_list= StringSplit ( $announce_list_vlc,";")


	if isarray($this_time_announce_list) then
		_ArrayDisplay ($this_time_announce_list)

			for $v=1 to $this_time_announce_list[0]

			ConsoleWrite ( @CRLF & $video_player_dir &'  --volume=350 --play-and-exit "'& $announce_file_root_dir &'\'& $this_time_announce_list[$v]  & @CRLF)
			;$announce_list_vlc="29 堡內上課須知.mp3"
			Run(@ComSpec & " /c " & $video_player_dir &'  --volume=350 --play-and-exit "'& $announce_file_root_dir &'\'& $this_time_announce_list[$v], "", @SW_HIDE)
			sleep(3*1000)
			$init_process_id =ProcessWait($action_exe )
			;if $init_process_id  then
			;	ConsoleWrite ( @CRLF&"Current time: " &  $current_time & ".  $init_process_id: " & $init_process_id  )
			;	ConsoleWrite ( @CRLF & " $init_process_id: " & $init_process_id  )
			;EndIf
			ProcessWaitClose($init_process_id)

			Next
	Else
			ConsoleWrite ( @CRLF & $video_player_dir &'  --volume=350 --play-and-exit "'& $announce_file_root_dir &'\'& $announce_list_vlc  & @CRLF)
			;$announce_list_vlc="29 堡內上課須知.mp3"
			Run(@ComSpec & " /c " & $video_player_dir &'  --volume=350 --play-and-exit "'& $announce_file_root_dir &'\'& $announce_list_vlc, "", @SW_HIDE)
			sleep(3*1000)
			$init_process_id =ProcessWait($action_exe )
			ProcessWaitClose($init_process_id)

	EndIf
	if ProcessExists ($init_process_id) then ProcessClose($init_process_id)
	Return true
EndFunc
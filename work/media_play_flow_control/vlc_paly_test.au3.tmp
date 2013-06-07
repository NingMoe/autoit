#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
#include <file.au3>
#include <array.au3>
; Script Start - Add your code below here

Global $video_player_dir="E:\green_apps\VLCPortable_2.0.5\VLCPortable\App\vlc\vlc.exe"
Global $video_file_root_directory="E:\workarround\音樂\ready\"
Global $announce_file_root_dir="E:\workarround\播音檔\京華堡播音"
dim $action_exe="vlc.exe"
dim $vlc_volume=200

dim $abc="1700.mp3"

_vlc_play ($abc)

Exit

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

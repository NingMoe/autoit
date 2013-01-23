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
dim $video_player_dir=@ScriptDir
dim $video_file_root_directory=@ScriptDir
dim $sync_delay=0
dim $hourly_playlist
dim $hour_start_from=0
dim $current_hourly_playlist , $next_hourly_palylist
dim $timer
dim $action_exe
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
		Next
	EndIf
	;MsgBox(0,"Setting",  $burgh &" ; " & $video_player_dir &" ; "& $video_file_root_directory &" ; "& $sync_delay )

EndIf
		$hour_start_from= $hour
		$hourly_playlist=_FileListToArray (@ScriptDir &"\","*.lst",1)
		 ;_ArrayDisplay ($hourly_playlist)

		;MsgBox(0,"", _ArraySearch( $hourly_playlist, "15.lst" ,0,0,1) )
		;Exit
While 1
	 $sec = @SEC
	 $min = @MIN
	 $hour = @HOUR

	 $current_time = $hour & $min & $sec
	;ConsoleWrite ("Video player dir: " & $video_player_dir & @CRLF)
	Select
		case $hour <= 9
			$current_hourly_playlist= "before_openning.lst"
			$next_hourly_palylist = "10.lst"
			$next_event=09 & $launch_player_time
		case $hour >= 22
			$current_hourly_playlist= "after_closing.lst"
			$next_hourly_palylist = "after_closing.lst"
			$next_event=$hour & $launch_player_time

		case $hour >= 10 and  $hour <22

				$current_hourly_playlist= $hour &".lst"  ;$hourly_playlist[$timer]
				$next_hourly_palylist = ($hour+1) &".lst"
				$next_event= $hour & $launch_player_time

		EndSelect
		ConsoleWrite (@CRLF & "Current time :"& $current_time&" $current_hourly_playlist :" & $current_hourly_playlist &". $next_hourly_palylist: "& $next_hourly_palylist &".  And Next event " &  $next_event )
		;MsgBox(0,"$current_hourly_playlist", $current_hourly_playlist)

		$action_exe="vlc.exe"
		$action_exe="notepad.exe"



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

WEnd
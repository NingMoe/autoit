#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.8.1
	Author:         myName

	Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
; 2012 12 28
; 這支程式是要讀取現在的時間然後取得現在的 PlayList後，交由 vlc 之類的 Player 程式撥出音樂
; 1 系統對時，如果可以取得 NTP 伺服器的時間那就最好了 V
; 2 開機時，由 server 讀取 PlayList ，若無法連線那就用現在的 Play List  V
; 3 檢查現行程式內有沒有特定節日的 event list
; 4 寫出今天的每個小時的 PlayList，若網路不通，則不覆寫每個小時的 PlayList
; 5 檢查是否要同步伺服器上的新增檔案;;; 這個能力可能是在定時檢查伺服器上的檔案清單
; 6 讀取本機設定值決定取得新媒體檔案的 delay time
; 7 讀取安全宣導小時提醒的排程原則
; 8
; 9
; 10



#include <File.au3>
#include <Date.au3>
#include <array.au3>
#include <ntp_udp123_func.au3>

#include <_Zip.au3>
;Opt("WinTitleMatchMode", 2)
Global $sec = @SEC
Global $min = @MIN
Global $hour = @HOUR
Global $day = @MDAY
Global $month = @MON
Global $year = @YEAR
Global $today = $year & $month & $day

Global $version = 0.1

global $shop_id=001
Global $connect_to_inet = 0
Global $version_url = "http://www.kidsburgh.com.tw/media_sync/version.htm"
Global $media_to_sync_url = "http://www.kidsburgh.com.tw/media_sync/kidsburgh_media_sync.html"
Global $playlist_to_sync_url = "http://www.kidsburgh.com.tw/media_sync/kidsburgh_playlist_sync.html"
Global $announcelist_to_sync_url = "http://www.kidsburgh.com.tw/media_sync/kidsburgh_announcelist_sync.html"
Dim $my_playlist = @ScriptDir & "\play_album.txt"
Dim $my_announcelist = @ScriptDir & "\announcelist.txt"
Dim $net_announcelist, $net_announcelist_date
Dim $net_playlist, $net_playlist_date, $my_playlist_date, $my_announcelist_date, $net_version

global $program_upgrade = 0
global $upgrade_date = 0
global $sync_delay_parameter=0
global $video_file_root_directory ="";"E:\workarround\音樂\ready\"
global $announce_file_root_directory ="";"E:\workarround\播音檔\京華堡播音"
Global $media_sync_status=0

Dim $aReturn_Playlist, $aReturn_announcelist
Global $valiade_date = 0
Dim $return_hour_list
Dim $new_plylist = 1
Global $this_app_start_time=0


Dim $test_mode = _TEST_MODE()
;MsgBox(0,"Test Mods", $test_mode)
dim $today_in_slash=$year &"/"& $month &"/"& $day
dim $current_time_with_colon= $hour &":"& $min &":"& $sec

$this_app_start_time= $today_in_slash &" "&$current_time_with_colon
;(0," time start", $this_app_start_time )




;MsgBox (0,"",  ( _DateDiff('n'   ,   $this_app_start_time , _NowCalc() )  ) )

if FileExists(@ScriptDir &"\global.ini" ) then
	$shop_id = IniRead (@ScriptDir &"\global.ini","shop_id","shop_id",$shop_id )
	$media_to_sync_url= IniRead (@ScriptDir &"\global.ini","media_to_sync_url","media_to_sync_url",$media_to_sync_url )
	$playlist_to_sync_url= IniRead (@ScriptDir &"\global.ini","playlist_to_sync_url","playlist_to_sync_url",$playlist_to_sync_url )
	$announcelist_to_sync_url =IniRead (@ScriptDir &"\global.ini","announcelist_to_sync_url","announcelist_to_sync_url",$announcelist_to_sync_url )

EndIf
	ConsoleWrite ("Global parameter" & @CRLF & $shop_id & @CRLF & $media_to_sync_url & @CRLF & $playlist_to_sync_url &@CRLF& $announcelist_to_sync_url)
	;MsgBox(0,"Global parameter", $shop_id & @CRLF & $media_to_sync_url & @CRLF & $playlist_to_sync_url &@CRLF& $announcelist_to_sync_url)


If Not FileExists(@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & ".ini") Then FileOpen(@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & ".ini", 10)
IniWrite(@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & ".ini", "version", "version", $version)
$upgrade_date=IniRead(@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & ".ini", "upgrade", "upgrade_date", $upgrade_date)
;$sync_delay_parameter=IniRead (@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & ".ini", "syncmedia", "sync_delay_parameter", $sync_delay_parameter)
$video_file_root_directory =IniRead (@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & ".ini", "syncmedia", "video_file_root_directory", $video_file_root_directory)
$announce_file_root_directory =IniRead (@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & ".ini", "syncmedia", "announce_file_root_directory", $announce_file_root_directory)
FileClose(@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & ".ini")
;MsgBox(0,"Parameter read ", $version &" , "& $upgrade_date &" , "& $sync_delay_parameter &" , "&$video_file_root_directory&" , "&$announce_file_root_directory)


_ping_to_inet()
if not $test_mode then _sync_inet_time()
;_MediaSync($media_to_sync_url)
_PlaylistSync($playlist_to_sync_url)


;Exit
;; check internet  play list and local playlist date difference
;; check if Upgrade needed.
;;
If $connect_to_inet = 1 Then
	$net_version = StringTrimLeft( _PlaylistSync($version_url), 8)
	;if IsString ($net_version) then $net_version= StringTrimLeft ( $net_version ,8)
	;MsgBox (0,"Net Version", $net_version)
	If $net_version > $version And $upgrade_date <> $today Then
		$upgrade = $today
		;if $program_upgrade then MsgBox (0," Need to upgrade", @ScriptDir &"\"&$net_version &".zip" ,5 )
		IniWrite(@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & ".ini", "upgrade", "upgrade_date", $today)
		$program_upgrade = _program_upgrade(StringTrimRight($version_url, 11), $net_version)
		;if FileExists (  @ScriptDir &"\"&$net_version &".zip" ) then MsgBox (0," Need to upgrade", @ScriptDir &"\"&$net_version &".zip" )
	EndIf
	$net_playlist = _PlaylistSync($playlist_to_sync_url)
	;MsgBox(0,"$net_playlist", $net_playlist)

	$net_announcelist = _PlaylistSync($announcelist_to_sync_url)
	;MsgBox(0,"$net_announcelist", $net_announcelist )

	If IsString($net_playlist) Then
		If StringInStr($net_playlist, "<#1>") Then
			$net_playlist_date = StringReplace(StringReplace(StringLeft($net_playlist, StringInStr($net_playlist, "<#1>") - 1), "<", ""), ">", "")
			;MsgBox(0,"Net playlist date", $net_playlist_date)
		EndIf
		if FileExists ($my_playlist) then
			$my_playlist_date = StringReplace(StringReplace(FileReadLine($my_playlist, 1), "<", ""), ">", "")
			;MsgBox(0,"$my_playlist_date vs Net playlist date:", $my_playlist_date &" VS " &$net_playlist_date )
			If ($net_playlist_date > $my_playlist_date) And $net_playlist_date < $today Then
				;MsgBox (0,"file move ? ", $my_playlist &@CRLF &$my_playlist & "_" &$today )
				FileMove( $my_playlist, @ScriptDir & "\bak\play_album.txt." & $today & ".bak" , 9)

				;MsgBox (0,"file write ? ", $my_playlist &@CRLF &$net_playlist)
				FileWrite($my_playlist, $net_playlist)
				;MsgBox(0,"Pause 1  Move file. $my_playlist_date vs Net playlist date:", $my_playlist_date &" VS " &$net_playlist_date )
			EndIf
		Else
			FileWrite($my_playlist, $net_playlist)
		EndIf

	EndIf
	;MsgBox(0,"Pause 2  $my_playlist_date vs Net playlist date:", $my_playlist_date &" VS " &$net_playlist_date )

	If IsString($net_announcelist) Then
		If StringInStr($net_announcelist, "<1#") Then
			$net_announcelist_date = StringReplace(StringReplace(StringLeft($net_announcelist, StringInStr($net_announcelist, "<1#") - 1), "<", ""), ">", "")
			;MsgBox(0,"$net announcelist date", $net_announcelist_date)
		EndIf
		if FileExists ($my_announcelist) then
			$my_announcelist_date = StringReplace(StringReplace(FileReadLine($my_announcelist, 1), "<", ""), ">", "")
			;MsgBox(0,"$my_announcelist_date VS $net_announcelist_date", $my_announcelist_date &" VS "& $net_announcelist_date)
			If ($net_announcelist_date > $my_announcelist_date) And $net_announcelist_date < $today Then
				;MsgBox(0,"File move?", $my_announcelist &@CRLF & @ScriptDir & "\bak\announcelist.txt." & $today & ".bak" )
				FileMove($my_announcelist, @ScriptDir & "\bak\announcelist.txt." & $today & ".bak", 9)
				FileWrite($my_announcelist, $net_announcelist)
			EndIf
		Else
			FileWrite($my_announcelist, $net_announcelist)
		EndIf

	EndIf
	Sleep(1000)

EndIf

; Read local playlist
;$aReturn_announcelist = _read_palylist($my_announcelist)
;if  IsArray( $aReturn_announcelist ) then _ArrayDisplay ( $aReturn_announcelist )

;Read local announce list
;$aReturn_Playlist = _read_palylist($my_playlist)
;if  IsArray( $aReturn_Playlist ) then _ArrayDisplay ( $aReturn_Playlist )




;If $new_plylist = 1 Then $return_hour_list = _merge_palylist($aReturn_Playlist, $aReturn_announcelist)
;_ArrayDisplay ($return_hour_list)

;Run(@ScriptDir & "\timer_control.exe")
;_MediaSync($media_to_sync_url)

while 1
if ProcessExists ("timer_control.exe") then
	$hour = @HOUR
	if $hour > 22 then Exit

	;MsgBox(0,"Shop id ", $shop_id  & " VS "& Int(  StringReplace($shop_id,"0","")  )  )
	;MsgBox (0,"time diff to sync ", " shop_id: " &$shop_id& " This ap start from: " & $this_app_start_time &"  and  laps "   &( _DateDiff('n'  ,  $this_app_start_time , _NowCalc() )  ) ,3)

	if $test_mode and $media_sync_status=0 then _MediaSync($media_to_sync_url)
	if $media_sync_status=0 and   ( _DateDiff('n'  ,  $this_app_start_time , _NowCalc() )  ) > ( Int(  StringReplace($shop_id,"0","")  ) * 10  )  then
		_MediaSync($media_to_sync_url)
	EndIf
	sleep (1000*10)
Else
	sleep(100)
	$hour = @HOUR
	if $hour > 22 then Exit
	Run(@ScriptDir & "\timer_control.exe")

EndIf

WEnd

Exit

Func _merge_palylist($playlist, $announcelist)
	Local $current_hour_playlist, $current_hour_announcelist
	Local $hour_list ;  排定 repeat 後的  10:00 ~ 22:00 的清單
	Local $repeat_hour = 0 ; 目前 playlist 內的所排定的清單的總時數，從 10:00 開始排起，有 1,2,3,4,5,6 hr 等

	$current_hour_announcelist = $announcelist
	For $x = $current_hour_announcelist[0] To 1 Step -1
		If StringInStr($current_hour_announcelist[$x], "~") Then _ArrayDelete($current_hour_announcelist, $x)
	Next
	$current_hour_announcelist[0] = UBound($current_hour_announcelist) - 1
	;_ArrayDisplay ($current_hour_announcelist)
	For $x = $current_hour_announcelist[0] To 1 Step -1
		If StringInStr($current_hour_announcelist[$x], "#") Then _ArrayDelete($current_hour_announcelist, $x)
	Next
	$current_hour_announcelist[0] = UBound($current_hour_announcelist) - 1
	;_ArrayDisplay ($current_hour_announcelist)

	For $x = 1 To $playlist[0] ; 這是取得目前未 repeat 前的排表時數之用
		If StringInStr($playlist[$x], "<#") Then
			;_ArrayAdd ( $hour_list, $playlist[$x])
			$repeat_hour = $repeat_hour + 1
		EndIf
	Next
	If $repeat_hour > 6 Then $repeat_hour = 6
	$hour_list = _repeat_playlist($repeat_hour) ; 這是取得 repeat 後的排表清單
	ConsoleWrite("$hour_list Ubound count: " & UBound($hour_list))
	;_ArrayDisplay($hour_list)


	; 以下才是以每個小時為單位，排出單一小時的 playlist ，並且完成檔案
	; 共有  10 , 11, 12, 13, 14, 15 這些
	;local $10_oc[1]
	;local $11_oc[1]
	;local $12_oc[1]
	;local $13_oc[1]
	;local $14_oc[1]
	;local $15_oc[1]
	;MsgBox (0, "Array search ",_ArraySearch( $playlist, "#10:00" ) )
	;_ArrayDisplay ($playlist)
	;MsgBox (0,"array search <#10:00>",  _ArraySearch( $playlist, "<#10:00>" )  &  _ArraySearch( $playlist, "</#10:00>" ))
	;MsgBox(0,"Hour to repeat", $repeat_hour)

	For $r = 1 To $repeat_hour
		; "<#1"&($r-1)&":00>" 就是隨著 $r 而變化的小時數，從 10:00 開始到 playlist 內的小時數目 。結果就是 <#10:00> ... <#15:00>
		If _ArraySearch($playlist, "<#" & ($r - 1 + 10) & ":00>") > 0 And _ArraySearch($playlist, "</#" & ($r - 1 + 10) & ":00>") > 0 Then
			Local $10_oc[1]

			For $x = _ArraySearch($playlist, "<#" & ($r - 1 + 10) & ":00>") + 1 To _ArraySearch($playlist, "</#" & ($r - 1 + 10) & ":00>") - 1
				_ArrayAdd($10_oc, $playlist[$x])
			Next
			$10_oc[0] = UBound($10_oc) - 1
			;_ArrayDisplay($10_oc)
		ElseIf _ArraySearch($playlist, "<#" & ($r - 1 + 10) & ":00>") > 0 And _ArraySearch($playlist, "~") Then
			Local $10_oc[1]
			For $x = _ArraySearch($playlist, "<#" & ($r - 1 + 10) & ":00>") + 1 To _ArraySearch($playlist, "~") - 1
				_ArrayAdd($10_oc, $playlist[$x])
			Next
			$10_oc[0] = UBound($10_oc) - 1
			;_ArrayDisplay($10_oc)
		Else
			Local $10_oc = 0
		EndIf
		;; 以下這段是把 announce list 平均插入 platlist 之中
		If IsArray($10_oc) Then
			Local $step_unmber = Floor($10_oc[0] / $current_hour_announcelist[0])
			Local $mod_no = Mod($10_oc[0], $current_hour_announcelist[0])
			Local $split_mod = 0
			;MsgBox (0,"The devided number",  $step_unmber  & "  " & $mod_no & "  " & $split_mod)
			For $x = $current_hour_announcelist[0] To 1 Step -1
				;if $mod_no>0 and $x>1 then  $minus_mod=1
				If $mod_no > 0 And $x > 1 Then
					$split_mod = 1
				Else
					$split_mod = 0
				EndIf
				;ConsoleWrite("The devided number: " & $step_unmber  & "  " & $mod_no & "  " & $split_mod & @CRLF)
				;ConsoleWrite(" array insert: 1+ $step_unmber*($x-1)+ ($mod_no-$split_mod)  " &@CRLF& "1+ " & $step_unmber &" * ("&$x&"-1) +"& ($mod_no-$split_mod) &"="&( 1+$step_unmber*($x-1)+ ($mod_no-$split_mod)) &@CRLF )

				_ArrayInsert($10_oc, 1 + $step_unmber * ($x - 1) + ($mod_no - $split_mod), $current_hour_announcelist[$x])
				;$minus_mod=$minus_mod+1
				If $mod_no > 0 Then $mod_no = $mod_no - 1
			Next

			;; 檢查 announce list 中是否有同一行之中有多個檔案要播音，如果有就插入在同一 array的欄位之後
			;_ArrayDisplay($10_oc)
			For $x = UBound($10_oc) - 1 To 1 Step -1
				Local $split_announcelist
				Local $split_announcelist_needed = 0
				If StringInStr($10_oc[$x], ";") Or StringInStr($10_oc[$x], ",") Then $split_announcelist_needed = 1
				If $split_announcelist_needed Then
					$split_announcelist = StringSplit($10_oc[$x], ";")
					For $z = 1 To $split_announcelist[0]
						_ArrayInsert($10_oc, $x + $z, $split_announcelist[$z])
					Next
					_ArrayDelete($10_oc, $x)
					;_ArrayDisplay($10_oc)

				EndIf

			Next
			;;以下是寫成檔案

			ConsoleWrite(@CRLF & @ScriptDir & "\" & ($r - 1 + 10) & ".lst" & "  --> " & @ScriptDir & "\bak\" & ($r - 1 + 10) & ".lst." & $today & $hour & $min & $sec & ".bak" & @CRLF)
			If FileExists(@ScriptDir & "\" & ($r - 1 + 10) & ".lst") Then FileMove(@ScriptDir & "\" & ($r - 1 + 10) & ".lst", @ScriptDir & "\bak\" & ($r - 1 + 10) & ".lst." & $today & $hour & $min & $sec & ".bak", 9)


			For $x = 1 To UBound($10_oc) - 1
				FileOpen(@ScriptDir & "\" & ($r - 1 + 10) & ".lst", 9)

				FileWriteLine(@ScriptDir & "\" & ($r - 1 + 10) & ".lst", $10_oc[$x])

			Next
			FileClose(@ScriptDir & "\" & ($r - 1 + 10) & ".lst")
			If FileExists(@ScriptDir & "\" & ($r - 1 + 10) & ".lst") Then ConsoleWrite("File exist : " & @ScriptDir & "\" & ($r - 1 + 10) & ".lst" & @CRLF)
			;_ArrayDisplay($10_oc, "count: 1" &$r-1 &":00")

		EndIf
	Next
	;;;;;;
	For $x = $repeat_hour To UBound($hour_list) - 1
		FileCopy(@ScriptDir & "\" & $hour_list[$x] & ".lst", ($x + 10) & ".lst", 9)
		If FileExists(($x + 10) & ".lst") Then ConsoleWrite($hour_list[$x] & " --> " & $x & "  " & ($x + 10) & ".lst" & @CRLF)
	Next
	;_ArrayDisplay($hour_list)
	Return $hour_list
	;;;;

EndFunc   ;==>_merge_palylist

Func _repeat_playlist($hr_to_repeat)
	Local $repeat_playlist
	If $hr_to_repeat = 1 Then Local $repeat_playlist[13] = [10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10]

	If $hr_to_repeat = 2 Then Local $repeat_playlist[13] = [10, 11, 10, 11, 10, 11, 10, 11, 10, 11, 10, 11, 10]

	If $hr_to_repeat = 3 Then Local $repeat_playlist[13] = [10, 11, 12, 10, 11, 12, 10, 11, 12, 10, 11, 12, 10]

	If $hr_to_repeat = 4 Then Local $repeat_playlist[13] = [10, 11, 12, 13, 10, 11, 12, 13, 10, 11, 12, 13, 10]

	If $hr_to_repeat = 5 Then Local $repeat_playlist[13] = [10, 11, 12, 13, 14, 10, 11, 12, 13, 14, 10, 11, 12]

	If $hr_to_repeat = 6 Then Local $repeat_playlist[13] = [10, 11, 12, 13, 14, 15, 10, 11, 12, 13, 14, 15, 10]

	Return $repeat_playlist
EndFunc   ;==>_repeat_playlist

Func _read_palylist($playlist)


	; read playlist and check media file exist?
	;
	Local $playlist_start_date = 0
	Local $aPlaylist
	Local $playlist_last_modify = 0
	If FileExists($playlist) Then
		$playlist_last_modify = FileGetTime($playlist, "", 1)
		;ConsoleWrite("File exist at : " & $playlist & @CRLF )
		_FileReadToArray($playlist, $aPlaylist)
		;_ArrayDisplay ($aPlaylist)

		If IsArray($aPlaylist) Then
			For $x = UBound($aPlaylist) - 1 To 0 Step -1
				If $aPlaylist[$x] = "" Then
					_ArrayDelete($aPlaylist, $x)
					$aPlaylist[0] = $aPlaylist[0] - 1
				EndIf
			Next
		Else
			$aPlaylist = 0
		EndIf

		If IsArray($aPlaylist) Then
			$playlist_start_date = StringStripWS(StringReplace(StringReplace($aPlaylist[1], "<", ""), ">", ""), 8)
			;MsgBox(0,"Start date" , $playlist_start_date &"  "& $year & $month & $day )
			IniWrite(@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & ".ini", "playlist", "playlist", $playlist)
			IniWrite(@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & ".ini", "playlist", "valiade_date", $playlist_start_date)
			IniWrite(@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & ".ini", "playlist", "modified_date", $playlist_last_modify)

			If $playlist_start_date <= $year & $month & $day Then
				$valiade_date = $playlist_start_date
				;MsgBox(0,"valiade_date" , $valiade_date)
				_ArrayDelete($aPlaylist, 1)
				$aPlaylist[0] = $aPlaylist[0] - 1
			Else
				MsgBox(0, "Attention _read_Playlist()", "The date on this play list is not valiade")
				;Exit
				$new_plylist = 0
			EndIf
		EndIf

		;_ArrayDisplay($aPlaylist)
	Else
		$aPlaylist = 0
	EndIf
	Return ($aPlaylist)
EndFunc   ;==>_read_palylist


Func _ping_to_inet()
	$ping = InetGetSize("http://www.kidsburgh.com.tw/images/top_logo.gif", 1)
	;$ping=Ping("168.95.192.1", 250)
	If $ping > 0 Then
		$connect_to_inet = 1
		ConsoleWrite("connect_to_inet=1" & @CRLF)
	ElseIf $ping = 0 Then
		ConsoleWrite("connect_to_inet=0" & @CRLF)
	EndIf
EndFunc   ;==>_ping_to_inet

Func _PlaylistSync($url)
	Local $playlist_to_sync = 0
	If $connect_to_inet = 1 Then
		Local $aData = InetRead($url, 1)
		Local $aBytesRead = @extended
		;MsgBox(4096, "", "Bytes read: " & $aBytesRead & @CRLF & @CRLF & StringLeft( BinaryToString($aData),3) & @CRLF & StringTrimLeft( BinaryToString($aData),4) )
		;$version = StringLeft(BinaryToString($aData), 3)
		$playlist_to_sync = BinaryToString($aData)
		;MsgBox(0,"$playlist_to_sync", $playlist_to_sync,5)
		ConsoleWrite ( $playlist_to_sync& @CRLF)
	EndIf
	Return ($playlist_to_sync)
EndFunc   ;==>_PlaylistSync

Func _MediaSync($url)
	Local $media_to_sync = 0
	local $this_mediasync_dir=""
	Local $this_mediasync_string=""
	local $check_media_sync
	local $this_winscp_id=0
	local $winwait_winscp_pid=0

	;local
	; winscp.exe /console /log=log.txt /command "option confirm off " "open ftp://kbuser:25176455@mail.kidsburgh.com.tw:6002" "synchronize local E:\workarround\音樂\syncsample /kshare/mis/syncsample -delete"  "close " "exit"
	;if not FileExists(@ScriptDir &"\winscp.exe") then MsgBox (0,"Error, File not exist", @ScriptDir &"\winscp.exe is not here." )

	If $connect_to_inet = 1 Then
		Local $aData = InetRead($url, 1)
		Local $aBytesRead = @extended
		;MsgBox(4096, "", "Bytes read: " & $aBytesRead & @CRLF & @CRLF & StringLeft( BinaryToString($aData),3) & @CRLF & StringTrimLeft( BinaryToString($aData),4) )
		;$version = StringLeft(BinaryToString($aData), 3)
		$media_to_sync = BinaryToString($aData)
		;MsgBox(0, "$media_to_sync", $media_to_sync, 2)
		;MsgBox(0, "$media_to_sync", $media_to_sync)
		ConsoleWrite( @CRLF& $media_to_sync & @CRLF)

		$this_mediasync_string= StringSplit( StringStripCR( $media_to_sync ) ,"media_sync" ,1)
		for $x=1 to $this_mediasync_string[0]
			if StringInStr( $this_mediasync_string[$x] , "_"& $shop_id ) then
				$this_mediasync_dir= StringStripWS( StringTrimLeft( $this_mediasync_string[$x] , 5) ,3)
				;ConsoleWrite ( @CRLF & $x & ": $this_mediasync_dir: '" & $this_mediasync_dir &"' "& @CRLF )
				;Else
				;$this_mediasync_dir= StringStripWS( StringTrimLeft( $this_mediasync_string[2] , 5) ,3)
				;ConsoleWrite ( @CRLF &"$this_mediasync_dir: '" & $this_mediasync_dir &"' "& @CRLF )
			EndIf
		next
		if $this_mediasync_dir="" then
			$this_mediasync_dir=StringStripWS( StringTrimLeft( $this_mediasync_string[2] , 1) ,3)
			ConsoleWrite ( @CRLF &"2: $this_mediasync_dir: '" & $this_mediasync_dir &"' "& @CRLF )
		EndIf
		;_ArrayDisplay($this_mediasync_string)
		;MsgBox(0," Run winscp.exe ", @ComSpec & " /c " & @ScriptDir &'\winscp.exe /console /log=mediasync_log_'&$today&'.txt /command "option confirm off " "open ftp://kbuser:25176455@mail.kidsburgh.com.tw:6002" "synchronize local ' &" E:\workarround\音樂\syncsample "& $this_mediasync_dir &' -delete"  "close " "exit"' & '"" @SW_HIDE ')

		; winscp.exe /console /log=log.txt /command "option confirm off " "open ftp://kbuser:25176455@mail.kidsburgh.com.tw:6002" "synchronize local E:\workarround\音樂\syncsample /kshare/mis/syncsample -delete"  "close " "exit"
		if not FileExists(@ScriptDir &"\winscp.exe") then
			MsgBox (0,"Error, File not exist", @ScriptDir &"\winscp.exe is not here." )
		Else
			;MsgBox(0," Run winscp.exe ", @ComSpec & " /c " & @ScriptDir &'\winscp.exe /console /log=mediasync_log_'&$today&'.txt /command "option confirm off " "open ftp://kbuser:25176455@mail.kidsburgh.com.tw:6002" "synchronize local ' &" E:\workarround\音樂\syncsample "& $this_mediasync_dir &' -delete"  "close " "exit"' & '"" @SW_HIDE ')
			if FileExists ( @ScriptDir&'\mediasync_log_'&$today&'.txt') then FileMove ( @ScriptDir&'\mediasync_log_'&$today&'.txt' , @ScriptDir&'\bak\mediasync_log_'&$today&'.txt' )
			if $media_sync_status=0 then
				$winwait_winscp_pid=run( @ScriptDir&"\winwait_winscp.exe","",@SW_MINIMIZE)
				$this_winscp_id=Run(@ComSpec & " /c " & @ScriptDir &'\winscp.exe /console /log=mediasync_log_'&$today&'.txt /command "option confirm off " "open ftp://kbuser:25176455@mail.kidsburgh.com.tw:6002" "synchronize local ' &" E:\workarround\音樂\syncsample "& $this_mediasync_dir &' -delete"  "close " "exit"', "", @SW_MINIMIZE)
				sleep(100)
				Opt("WinTitleMatchMode", 2)
				ProcessWait ("winscp.exe")
				WinSetState("WinSCP","",@SW_MINIMIZE)

				if ProcessExists ("winscp.exe") then ProcessWaitClose ( "winscp.exe")
				if ProcessExists ( $winwait_winscp_pid ) Then ProcessClose ( $winwait_winscp_pid)
			EndIf

			if FileExists ( @ScriptDir&'\mediasync_log_'&$today&'.txt') then
				_FileReadToArray ( @ScriptDir&'\mediasync_log_'&$today&'.txt' , $check_media_sync )
				;_ArrayDisplay ( $check_media_sync )

				if IsArray( $check_media_sync ) Then
					;MsgBox(0,"check 1" , "check if log file has 'Nothing to synchronize.'")
					;Nothing to synchronize.
					for $abc= $check_media_sync[0] to 1  step -1
						if StringInStr ( $check_media_sync[$abc] , "Nothing to synchronize." ) then
							;MsgBox(0,"check 2" , "log file has 'Nothing to synchronize.'")
							$media_sync_status=1
						EndIf
					Next

				EndIf
			EndIf




		EndIf
		;Run(@ComSpec & " /c " & $video_player_dir &'  --volume='&$vlc_volume&' --play-and-exit "'& $announce_file_root_dir &'\'& $this_time_announce_list[$v], "", @SW_HIDE)
		;sleep(2*1000)

	EndIf
	Return ($media_to_sync)
EndFunc   ;==>_MediaSync

Func _sync_inet_time()
	If $connect_to_inet = 0 Then Return
	Local $net_time
	If _ntpnow() = -1 Then
		MsgBox(0, "NTP Error", "無法使用網路對時，使用本機時間", 5)
	Else
		Local $iDateCalc
		;$iDateCalc = _DateDiff( 's',"1970/01/01 00:00:00",_ntpnow())
		;MsgBox( 4096, "", "Number of seconds since EPOCH: " & $iDateCalc )
		;MsgBox(0,'Now', _ntpnow(),5)
		$net_time = _ntpnow()
		Run(@ComSpec & " /c " & 'time ' & StringTrimLeft($net_time, 11), "", @SW_HIDE) ;
		Run(@ComSpec & " /c " & 'date ' & StringLeft($net_time, 10), "", @SW_HIDE)
		;ConsoleWrite (@CRLF &  StringTrimLeft($net_time,11 ) &@CRLF)
		;ConsoleWrite (@CRLF &  StringLeft($net_time,10 ) &@CRLF)
	EndIf
EndFunc   ;==>_sync_inet_time


Func _array2string_tab($array, $d)
	; $d 是傳入的 array 的欄位數目
	Local $y, $x, $i, $string_to_return, $a_line
	$string_to_return = ''

	;MsgBox(0, "Dimention", "  $Y :" & UBound($array))

	For $y = 1 To UBound($array) - 1
		$a_line = ''
		If $d = 1 Then
			$a_line = $array[$y]

		Else
			For $x = 0 To $d - 1
				If $x < $d - 1 Then
					$a_line = $a_line & $array[$y][$x] & '^'
				Else
					$a_line = $a_line & $array[$y][$x]
				EndIf
				;MsgBox(0," array 2 string tab", $a_line)
				;ConsoleWrite($a_line & @crlf)
				;$string_to_return=$string_to_return  & $array[$y][1] &" , "& $array[$y][2] &" , " &$array[$y][6] &@CRLF
			Next
			ConsoleWrite($a_line & @CRLF)
			;$a_line=StringTrimRight($a_line,3) ; Cut off the last 3 character --> ","
		EndIf
		$string_to_return = $string_to_return & $a_line & @CRLF
	Next


	;MsgBox(0,"string", $string_to_return)
	Return $string_to_return
EndFunc   ;==>_array2string_tab


Func _program_upgrade($net_version_path_in_fun, $net_version_in_fun)
	Local $aFile_in_download_upgrade, $upgrade_confirm
	Local $upgrade_file_list_in_func
	; $version_url="http://www.kidsburgh.com.tw/media_sync/version.htm"
	;MsgBox ( 0, "net Version path",StringTrimRight ( $version_url,11 ) )
	InetGet($net_version_path_in_fun & $net_version_in_fun & ".zip", @ScriptDir & "\" & $net_version_in_fun & ".zip")
	If FileExists(@ScriptDir & "\" & $net_version_in_fun & ".zip") Then

		_Zip_Unzipall(@ScriptDir & "\" & $net_version_in_fun & ".zip", @ScriptDir & "\" & $net_version_in_fun & "\")
		$aFile_in_download_upgrade = _FileListToArray(@ScriptDir & "\" & $net_version_in_fun & "\")
		;_ArrayDisplay($aFile_in_download_upgrade)
		$upgrade_file_list_in_func = FileOpen(@ScriptDir & "\Upgrade_FileList_" & $today, 10)
		FileWriteLine($upgrade_file_list_in_func, $today & "=1")
		FileWriteLine($upgrade_file_list_in_func, @ScriptDir & "\" & $net_version_in_fun)

		For $u = 1 To $aFile_in_download_upgrade[0]
			;if FileExists (@ScriptDir &"\"&$net_version_in_fun&"\" & $aFile_in_download_upgrade[$u] ) then
			;FileCopy (@ScriptDir &"\"&$net_version_in_fun&"\" & $aFile_in_download_upgrade[$u] ,)
			;EndIf
			FileWriteLine($upgrade_file_list_in_func, $aFile_in_download_upgrade[$u])

		Next
		FileClose($upgrade_file_list_in_func)


		$upgrade_confirm = MsgBox(4, "Version Upgrade", "Program Upgrade to Version : " & $net_version_in_fun)
		If $upgrade_confirm Then
			If FileExists(@ScriptDir & "\copy_upgrade_files.exe") Then
				Run(@ScriptDir & "\copy_upgrade_files.exe")
			EndIf
			ProcessWait("copy_upgrade_files.exe")
			If ProcessExists("copy_upgrade_files.exe") Then Exit

		Else
			MsgBox(0, "Upgrade error", "Look for system admin help!!")
			Return 0
		EndIf

	Else
		MsgBox(0, "Upgrade error", "Look for system admin help!!")
		Return 0
	EndIf
EndFunc   ;==>_program_upgrade




Func _TEST_MODE()
	If FileExists(@ScriptDir & "\TESTMODE.txt") Then
		$mode = FileReadLine(@ScriptDir & "\TESTMODE.txt", 1)
		If $mode = 1 Then
			MsgBox(0, "Test mode", "This is Test mode. ", 5)

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

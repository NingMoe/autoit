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

Global $version=0.1

Global $connect_to_inet=0
Global $version_url="http://www.kidsburgh.com.tw/media_sync/version.htm"
Global $media_to_sync_rul="http://www.kidsburgh.com.tw/media_sync/kidsburgh_media_sync.html"
Global $playlist_to_sync_rul="http://www.kidsburgh.com.tw/media_sync/kidsburgh_playlist_sync.html"
global $announcelist_to_sync_url="http://www.kidsburgh.com.tw/media_sync/kidsburgh_announcelist_sync.html"
dim $my_playlist= @ScriptDir &"\playlist.txt"
dim $my_announcelist= @ScriptDir &"\announcelist.txt"
dim $net_announcelist , $net_announcelist_date
dim $net_playlist , $net_playlist_date , $my_playlist_date , $my_announcelist_date , $net_version

dim $program_upgrade=0
dim $upgrade_date=0

dim $aReturn_Playlist , $aReturn_announcelist
Global $valiade_date=0
dim $return_hour_list
dim $new_plylist=1
	if not FileExists (@ScriptDir &"\" & StringTrimRight(@ScriptName,4) &".ini") then FileOpen (@ScriptDir &"\" & StringTrimRight(@ScriptName,4) &".ini",10)
	IniWrite (@ScriptDir &"\" & StringTrimRight(@ScriptName,4) &".ini" ,"version", "version", $version)
	IniRead (@ScriptDir &"\" & StringTrimRight(@ScriptName,4) &".ini" ,"upgrade", "upgrade_date", $upgrade_date)
	FileClose ( @ScriptDir &"\" & StringTrimRight(@ScriptName,4) &".ini" )
	_ping_to_inet()
	_sync_inet_time()
	_MediaSync($media_to_sync_rul)

;; check internet  play list and local playlist date difference
;; check if Upgrade needed.
;;
if $connect_to_inet=1  then
	$net_version= StringTrimLeft( _PlaylistSync($version_url) ,8 )
		;if IsString ($net_version) then $net_version= StringTrimLeft ( $net_version ,8)
		;MsgBox (0,"Net Version", $net_version)
		if $net_version > $version   and  $upgrade_date<> $today then
			$upgrade=$today
			;if $program_upgrade then MsgBox (0," Need to upgrade", @ScriptDir &"\"&$net_version &".zip" ,5 )
			IniWrite (@ScriptDir &"\" & StringTrimRight(@ScriptName,4) &".ini" ,"upgrade", "upgrade_date", $today)
			$program_upgrade=  _program_upgrade( StringTrimRight ( $version_url,11 ) ,$net_version )
			;if FileExists (  @ScriptDir &"\"&$net_version &".zip" ) then MsgBox (0," Need to upgrade", @ScriptDir &"\"&$net_version &".zip" )
		EndIf
	$net_playlist=_PlaylistSync($playlist_to_sync_rul)
	;MsgBox(0,"$net_playlist", $net_playlist)

	$net_announcelist=_PlaylistSync($announcelist_to_sync_url)
	;MsgBox(0,"$net_announcelist", $net_announcelist )

	if IsString ($net_playlist ) then
		if StringInStr( $net_playlist, "<#10:00>" ) then
			$net_playlist_date =  StringReplace( StringReplace( stringleft ( $net_playlist, StringInStr( $net_playlist, "<#10:00>" )-1 )  ,"<","")  ,">","")
			;MsgBox(0,"Net playlist date", $net_playlist_date)
		EndIf
		$my_playlist_date=  StringReplace( StringReplace( FileReadLine ( $my_playlist , 1 )   ,"<","")  ,">","")
			;MsgBox(0,"$my_playlist_date", $my_playlist_date)
		if ($net_playlist_date > $my_playlist_date) and $net_playlist_date > $today then
			filemove ( $my_playlist , @ScriptDir &"\bak\playlist.txt." & $today &".bak" ,9)
			FileWrite ( $my_playlist , $net_playlist )
		EndIf
	endif

	if IsString ($net_announcelist ) then
		if StringInStr( $net_announcelist, "<#10:00>" ) then
			$net_announcelist_date =  StringReplace( StringReplace( stringleft ( $net_announcelist, StringInStr( $net_announcelist, "<#10:00>" )-1 )  ,"<","")  ,">","")
			;MsgBox(0,"$net announcelist date", $net_announcelist_date)
		EndIf
		$my_announcelist_date=  StringReplace( StringReplace( FileReadLine ( $my_announcelist , 1 )   ,"<","")  ,">","")
			;MsgBox(0,"$my_announcelist_date", $my_announcelist_date)
		if ($net_announcelist_date > $my_announcelist_date) and $net_announcelist_date > $today then
			filemove ( $my_announcelist , @ScriptDir &"\bak\announcelist.txt." & $today &".bak" ,9)
			FileWrite ( $my_announcelist , $net_announcelist )
		EndIf
	endif
	sleep(1000)

EndIf

	; Read local playlist
	$aReturn_announcelist = _read_palylist($my_announcelist)
	;if  IsArray( $aReturn_announcelist ) then _ArrayDisplay ( $aReturn_announcelist )

	;Read local announce list
	$aReturn_Playlist= _read_palylist($my_playlist)
	;if  IsArray( $aReturn_Playlist ) then _ArrayDisplay ( $aReturn_Playlist )




	if $new_plylist =1 then $return_hour_list=_merge_palylist($aReturn_Playlist, $aReturn_announcelist)
	;_ArrayDisplay ($return_hour_list)

	run (@ScriptDir &"\timer_control.exe")

Exit

func _merge_palylist( $playlist , $announcelist)
	local $current_hour_playlist , $current_hour_announcelist
	local $hour_list    ;  排定 repeat 後的  10:00 ~ 22:00 的清單
	Local $repeat_hour=0 ; 目前 playlist 內的所排定的清單的總時數，從 10:00 開始排起，有 1,2,3,4,5,6 hr 等

	$current_hour_announcelist =  $announcelist
	for $x=$current_hour_announcelist[0] to 1 step -1
			if StringInStr( $current_hour_announcelist[$x], "~" ) then _ArrayDelete ($current_hour_announcelist, $x)
	Next
		$current_hour_announcelist[0]= UBound ($current_hour_announcelist)-1
		;_ArrayDisplay ($current_hour_announcelist)
	for $x=$current_hour_announcelist[0] to 1 step -1
			if StringInStr( $current_hour_announcelist[$x], "#" ) then _ArrayDelete ($current_hour_announcelist, $x)
	Next
		$current_hour_announcelist[0]= UBound ($current_hour_announcelist)-1
		;_ArrayDisplay ($current_hour_announcelist)

	for  $x=1  to $playlist[0]  ; 這是取得目前未 repeat 前的排表時數之用
		if StringInStr ($playlist[$x] , "<#" )  then
			;_ArrayAdd ( $hour_list, $playlist[$x])
			$repeat_hour= $repeat_hour+1
		EndIf
	Next
		if $repeat_hour >6 then $repeat_hour=6
		$hour_list=_repeat_playlist ($repeat_hour)  ; 這是取得 repeat 後的排表清單
		ConsoleWrite ("$hour_list Ubound count: " & UBound($hour_list))
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

	for $r=1 to $repeat_hour
			; "<#1"&($r-1)&":00>" 就是隨著 $r 而變化的小時數，從 10:00 開始到 playlist 內的小時數目 。結果就是 <#10:00> ... <#15:00>
		if  _ArraySearch( $playlist, "<#"&($r-1+10)&":00>" )>0 and _ArraySearch( $playlist, "</#"&($r-1+10)&":00>" ) >0  then
			local $10_oc[1]

			for $x =_ArraySearch( $playlist, "<#"&($r-1+10)&":00>" )+1 to _ArraySearch( $playlist, "</#"&($r-1+10)&":00>" )-1
				_ArrayAdd($10_oc, $playlist[$x])
			Next
				$10_oc[0]= UBound ($10_oc)-1
				;_ArrayDisplay($10_oc)
			ElseIf _ArraySearch( $playlist, "<#"&($r-1+10)&":00>" ) >0 and _ArraySearch( $playlist, "~" ) then
				local $10_oc[1]
			for $x =_ArraySearch( $playlist, "<#"&($r-1+10)&":00>" )+1 to _ArraySearch( $playlist, "~" )-1
				_ArrayAdd($10_oc, $playlist[$x])
			Next
				$10_oc[0]= UBound ($10_oc)-1
				;_ArrayDisplay($10_oc)
		Else
			local $10_oc=0
		EndIf
		;; 以下這段是把 announce list 平均插入 platlist 之中
		if IsArray ($10_oc) then
			local $step_unmber=Floor ( $10_oc[0]/$current_hour_announcelist[0] )
			local $mod_no = mod( $10_oc[0], $current_hour_announcelist[0] )
			local $split_mod=0
			 ;MsgBox (0,"The devided number",  $step_unmber  & "  " & $mod_no & "  " & $split_mod)
				for $x=$current_hour_announcelist[0] to 1 step -1
				;if $mod_no>0 and $x>1 then  $minus_mod=1
				if $mod_no>0 and $x>1 then
					$split_mod=1
				Else
					$split_mod=0
				EndIf
				;ConsoleWrite("The devided number: " & $step_unmber  & "  " & $mod_no & "  " & $split_mod & @CRLF)
				;ConsoleWrite(" array insert: 1+ $step_unmber*($x-1)+ ($mod_no-$split_mod)  " &@CRLF& "1+ " & $step_unmber &" * ("&$x&"-1) +"& ($mod_no-$split_mod) &"="&( 1+$step_unmber*($x-1)+ ($mod_no-$split_mod)) &@CRLF )

				_ArrayInsert ( $10_oc,  1+$step_unmber*($x-1)+ ($mod_no-$split_mod)  , $current_hour_announcelist[$x])
				;$minus_mod=$minus_mod+1
				if $mod_no>0 then $mod_no=$mod_no-1
			Next

			;; 檢查 announce list 中是否有同一行之中有多個檔案要播音，如果有就插入在同一 array的欄位之後
			;_ArrayDisplay($10_oc)
			for $x=UBound ($10_oc)-1 to 1 step -1
				local $split_announcelist
				local $split_announcelist_needed=0
				if StringInStr ($10_oc[$x] ,";" ) or StringInStr ($10_oc[$x] ,"," ) then $split_announcelist_needed=1
				if $split_announcelist_needed then
						$split_announcelist= StringSplit ($10_oc[$x],";")
						for $z=1 to $split_announcelist[0]
						_ArrayInsert ($10_oc,$x+$z, $split_announcelist[$z] )
						Next
						_ArrayDelete ($10_oc, $x)
						;_ArrayDisplay($10_oc)

				EndIf

			Next
			;;以下是寫成檔案

			ConsoleWrite ( @CRLF& @ScriptDir &"\"& ( $r-1 +10)&".lst" & "  --> "& @ScriptDir &"\bak\"& ( $r-1 +10)&".lst."&$today &$hour&$min & $sec & ".bak" &@CRLF)
			if FileExists (@ScriptDir &"\"& ( $r-1 +10)&".lst")  then FileMove (@ScriptDir &"\"& ( $r-1 +10)&".lst" , @ScriptDir &"\bak\"& ( $r-1 +10)&".lst."&$today &$hour&$min & $sec &".bak",9)


			for $x=1 to UBound ($10_oc)-1
				FileOpen (@ScriptDir &"\"& ( $r-1 +10)&".lst",9)

				FileWriteLine( @ScriptDir &"\"& ( $r-1+10)&".lst", $10_oc[$x] )

			Next
			FileClose (@ScriptDir &"\"& ( $r-1+10)&".lst")
			if FileExists (@ScriptDir &"\"& ( $r-1+10)&".lst" ) then ConsoleWrite ("File exist : " &@ScriptDir &"\"& ( $r-1+10)&".lst" & @CRLF)
			;_ArrayDisplay($10_oc, "count: 1" &$r-1 &":00")

		EndIf
	Next
;;;;;;
	for $x= $repeat_hour  to UBound($hour_list)-1
		filecopy ( @ScriptDir &"\"& $hour_list[$x] &".lst" , ( $x +10 )&".lst" ,9 )
		if FileExists ( ( $x +10 )&".lst" ) then  ConsoleWrite ( $hour_list[$x] &" --> "& $x  & "  " &( $x +10 )&".lst" &@CRLF)
	Next
	;_ArrayDisplay($hour_list)
	return $hour_list
;;;;

EndFunc

func _repeat_playlist($hr_to_repeat )
	local $repeat_playlist
	if $hr_to_repeat =1 then local $repeat_playlist[13]=[10,10,10,10,10,10,10,10,10,10,10,10,10]

	if $hr_to_repeat =2 then local $repeat_playlist[13]=[10,11,10,11,10,11,10,11,10,11,10,11,10]

	if $hr_to_repeat =3 then local $repeat_playlist[13]=[10,11,12,10,11,12,10,11,12,10,11,12,10]

	if $hr_to_repeat =4 then local $repeat_playlist[13]=[10,11,12,13,10,11,12,13,10,11,12,13,10]

	if $hr_to_repeat =5 then local $repeat_playlist[13]=[10,11,12,13,14,10,11,12,13,14,10,11,12]

	if $hr_to_repeat =6 then local $repeat_playlist[13]=[10,11,12,13,14,15,10,11,12,13,14,15,10]

return $repeat_playlist
EndFunc

func _read_palylist($playlist)


	; read playlist and check media file exist?
	;
	local $playlist_start_date=0
	Local $aPlaylist
	local $playlist_last_modify=0
	if FileExists ($playlist) then
		$playlist_last_modify=FileGetTime ( $playlist,"", 1)
		;ConsoleWrite("File exist at : " & $playlist & @CRLF )
		_FileReadToArray ($playlist , $aPlaylist )
		;_ArrayDisplay ($aPlaylist)

		if IsArray($aPlaylist ) then
			for $x= UBound ($aPlaylist)-1 to 0 step -1
				if $aPlaylist[$x]="" then
					_ArrayDelete ($aPlaylist, $x)
					$aPlaylist[0]=$aPlaylist[0]-1
				EndIf
			next
		else
			$aPlaylist=0
		EndIf

		If IsArray($aPlaylist ) then
			$playlist_start_date= StringStripWS( StringReplace( StringReplace( $aPlaylist[1],"<","" ) ,">","" ) ,8)
				;MsgBox(0,"Start date" , $playlist_start_date &"  "& $year & $month & $day )
				IniWrite (@ScriptDir &"\" & StringTrimRight(@ScriptName,4) &".ini" ,"playlist", "playlist", $playlist)
				IniWrite (@ScriptDir &"\" & StringTrimRight(@ScriptName,4) &".ini" ,"playlist", "valiade_date", $playlist_start_date)
				IniWrite (@ScriptDir &"\" & StringTrimRight(@ScriptName,4) &".ini" ,"playlist", "modified_date", $playlist_last_modify)

			if $playlist_start_date<= $year & $month & $day Then
				$valiade_date=$playlist_start_date
				;MsgBox(0,"valiade_date" , $valiade_date)
				_ArrayDelete ($aPlaylist, 1)
				$aPlaylist[0]=$aPlaylist[0]-1
			else
				MsgBox(0,"Attention _read_Playlist()", "The date on this play list is not valiade")
				;Exit
				$new_plylist=0
			EndIf
		EndIf

		;_ArrayDisplay($aPlaylist)
	Else
		$aPlaylist=0
	EndIf
	return ($aPlaylist)
EndFunc


func _ping_to_inet()
	$ping=InetGetSize ("http://www.kidsburgh.com.tw/images/top_logo.gif",1)
	;$ping=Ping("168.95.192.1", 250)
	if $ping>0 then
		$connect_to_inet=1
		ConsoleWrite  ("connect_to_inet=1" & @CRLF)
	ElseIf $ping=0 then
		ConsoleWrite  ("connect_to_inet=0" & @CRLF)
	EndIf
EndFunc

func _PlaylistSync($url)
	local $playlist_to_sync=0
	if $connect_to_inet=1 then
		local $aData = InetRead($url,1)
		Local $aBytesRead = @extended
		;MsgBox(4096, "", "Bytes read: " & $aBytesRead & @CRLF & @CRLF & StringLeft( BinaryToString($aData),3) & @CRLF & StringTrimLeft( BinaryToString($aData),4) )
		;$version = StringLeft(BinaryToString($aData), 3)
		$playlist_to_sync=BinaryToString($aData)
		;MsgBox(0,"$playlist_to_sync", $playlist_to_sync,5)
		;ConsoleWrite ( $playlist_to_sync& @CRLF)
	EndIf
    return ($playlist_to_sync)
EndFunc

func _MediaSync($url)
	local $media_to_sync=0
	if $connect_to_inet=1 then
		local $aData = InetRead($url,1)
		Local $aBytesRead = @extended
		;MsgBox(4096, "", "Bytes read: " & $aBytesRead & @CRLF & @CRLF & StringLeft( BinaryToString($aData),3) & @CRLF & StringTrimLeft( BinaryToString($aData),4) )
		;$version = StringLeft(BinaryToString($aData), 3)
		$media_to_sync=BinaryToString($aData)
		MsgBox(0,"$media_to_sync", $media_to_sync,2)
		ConsoleWrite ( $media_to_sync& @CRLF)
	EndIf
    return ($media_to_sync)
EndFunc

func _sync_inet_time()
	if $connect_to_inet=0  then return
	local  $net_time
	if _ntpnow() = -1 then
	MsgBox(0,"NTP Error","無法使用網路對時，使用本機時間",5)
	else
		local $iDateCalc
		;$iDateCalc = _DateDiff( 's',"1970/01/01 00:00:00",_ntpnow())
		;MsgBox( 4096, "", "Number of seconds since EPOCH: " & $iDateCalc )
		;MsgBox(0,'Now', _ntpnow(),5)
		$net_time=_ntpnow()
		Run(@ComSpec & " /c " & 'time ' & StringTrimLeft($net_time,11 ), "", @SW_HIDE) ;
		Run(@ComSpec & " /c " & 'date ' & StringLeft($net_time,10 ), "", @SW_HIDE)
		;ConsoleWrite (@CRLF &  StringTrimLeft($net_time,11 ) &@CRLF)
		;ConsoleWrite (@CRLF &  StringLeft($net_time,10 ) &@CRLF)
	EndIf
EndFunc


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


func _program_upgrade( $net_version_path_in_fun , $net_version_in_fun)
		local $aFile_in_download_upgrade , $upgrade_confirm
		local $upgrade_file_list_in_func
		; $version_url="http://www.kidsburgh.com.tw/media_sync/version.htm"
		;MsgBox ( 0, "net Version path",StringTrimRight ( $version_url,11 ) )
		inetget( $net_version_path_in_fun & $net_version_in_fun &".zip" , @ScriptDir &"\"&$net_version_in_fun &".zip")
		if FileExists ( @ScriptDir &"\"&$net_version_in_fun &".zip" ) then

			_Zip_Unzipall ( @ScriptDir & "\"&$net_version_in_fun&".zip" ,@ScriptDir &"\"&$net_version_in_fun&"\" )
			$aFile_in_download_upgrade =_FileListToArray ( @ScriptDir &"\"&$net_version_in_fun&"\" )
			;_ArrayDisplay($aFile_in_download_upgrade)
			$upgrade_file_list_in_func= FileOpen ( @ScriptDir &"\Upgrade_FileList_"& $today  ,10 )
			FileWriteLine ( $upgrade_file_list_in_func ,  $today &"=1")
			FileWriteLine ( $upgrade_file_list_in_func ,  @ScriptDir &"\"&$net_version_in_fun)

			for $u=1 to $aFile_in_download_upgrade[0]
				;if FileExists (@ScriptDir &"\"&$net_version_in_fun&"\" & $aFile_in_download_upgrade[$u] ) then
					;FileCopy (@ScriptDir &"\"&$net_version_in_fun&"\" & $aFile_in_download_upgrade[$u] ,)
				;EndIf
				FileWriteLine ( $upgrade_file_list_in_func ,$aFile_in_download_upgrade [$u])

			Next
			FileClose ($upgrade_file_list_in_func)


			$upgrade_confirm= MsgBox(4,"Version Upgrade", "Program Upgrade to Version : "  & $net_version_in_fun )
			if $upgrade_confirm then
				if FileExists (@ScriptDir &"\copy_upgrade_files.exe") then
					run (@ScriptDir &"\copy_upgrade_files.exe")
				EndIf
				ProcessWait ( "copy_upgrade_files.exe" )
				if ProcessExists ( "copy_upgrade_files.exe" ) then exit

			Else
				MsgBox (0,"Upgrade error", "Look for system admin help!!")
				Return 0
			endif

		Else
			MsgBox (0,"Upgrade error", "Look for system admin help!!")
			Return 0
		EndIf
EndFunc

;#include <file.au3>
;#include <array.au3>
;#include <date.au3>


;Dim $sec = @SEC
;Dim $min = @MIN
;Dim $hour = @HOUR
;Dim $day = @MDAY
;Dim $month = @MON
;Dim $year = @YEAR
;Dim $today = $year & $month & $day
;
;dim $m_music_root="E:\workarround\音樂\ready\"
;dim $m_play_album= "";"1002;3001;3005"

;$m_play_album= _read_play_album ()
;if $m_play_album<>0 then
;	$a = _gen_playlist( $m_music_root , $m_play_album  )
;	;_ArrayDisplay ( $a  )
;EndIf

;Exit


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


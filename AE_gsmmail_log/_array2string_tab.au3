

func _array2string_tab($array,$d)
	local $y, $x, $i, $string_to_return ,$a_line
	$string_to_return=''
	
	;MsgBox(0,"Dimention", "  $Y :" &UBound($array) )
	
	for $y=1 to UBound($array)-1
		$a_line='"'
		if $d=1 then 
			$a_line=$array[$y]
			
		Else	
			for $x=0 to $d-1
				if $x < $d-1 then
					$a_line=$a_line &$array[$y][$x] &'"	"'
				else
					$a_line=$a_line &$array[$y][$x]& '"'
				EndIf
				;$string_to_return=$string_to_return  & $array[$y][1] &" , "& $array[$y][2] &" , " &$array[$y][6] &@CRLF
			next
			;$a_line=StringTrimRight($a_line,3) ; Cut off the last 3 character --> "," 
		EndIf
		$string_to_return=$string_to_return & $a_line  & @LF	
	Next


;MsgBox(0,"string", $string_to_return)
return $string_to_return
EndFunc
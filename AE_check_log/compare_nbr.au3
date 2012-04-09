#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include<file.au3>
#include<array.au3>

dim $evisa_items
dim $english="A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
dim $letters , $compare_string

$letters=StringSplit($english,",")
;_ArrayDisplay($letters)

_FileReadToArray( @ScriptDir &"\1.txt", $evisa_items )
;_ArrayDisplay($evisa_items)

;for $i =1 to $letters[0]	
	; StringLeft($evisa_items[$i],1) 
;Next

for $i=1 to $evisa_items[0]
		$compare_string=$evisa_items[$i]
		;MsgBox(0,"Comare from Line " & $i)
	for $j=$i+1 to  $evisa_items[0]
		;if $i=2 then MsgBox(0,"Compare" , "Line " & $i & " Compare string : " & $compare_string & " Same to Line : " & $j &  "  " & $evisa_items[$j] )
		if $evisa_items[$j] =$compare_string then 
			ConsoleWrite( "Line " & $i & " Compare string : " & $compare_string & "Same to Line : " & $j &  "  " & $evisa_items[$j] )
			MsgBox(0,"Compare" , "Line " & $i & " Compare string : " & $compare_string & " Same to Line : " & $j &  "  " & $evisa_items[$j] )
		endif	
	next
	

Next
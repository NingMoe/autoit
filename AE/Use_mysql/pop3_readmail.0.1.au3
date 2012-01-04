#include <array.au3>
#include <_pop3.au3>
#include <File.au3>

;ConsoleWrite(@AutoItVersion & @CRLF)
;~ See _pop3.au3 for a complete description of the pop3 functions
;~ Requires AU3 beta version 3.1.1.110 or newer.

Global $MyPopServer = "56168.com.tw"
Global $MyLogin = "bryant"
Global $MyPasswd = "9ps56789"



Global $MyPopServer = "onlinebooking.com.tw"
Global $MyLogin = "direct_send"
Global $MyPasswd = "amextravel"




_pop3Connect($MyPopServer, $MyLogin, $MyPasswd)
If @error Then
	MsgBox(0, "Error", "Unable to connect to " & $MyPopServer & @CR & @error)
	Exit
Else
	ConsoleWrite("Connected to server pop3 " & $MyPopServer & @CR)
EndIf

;;================================
;; Single mail retrive
;;
local $no=-1

if $no<> -1 then 
$single_retr = _Pop3Retr( $no )
;MsgBox(0,"Retr", $single_retr)

$file = FileOpen(@ScriptDir& "\"&$no&"_test.txt", 1)

; Check if file opened for writing OK
If $file = -1 Then
    MsgBox(0, "Error", "Unable to open file.")
    Exit
EndIf
FileWrite($file, $single_retr)
FileClose($file)

$single_retr= _retr_split_3( $no  )
_ArrayDisplay( $single_retr )

MsgBox(0,"This is a sigle message", "Wait if you going to real message")

EndIf
;;================================


;Local $stat = _Pop3Stat()
;If Not @error Then
;	_ArrayDisplay($stat, "Result of STAT COMMAND")
;Else
;	ConsoleWrite("Stat commande failed" & @CR)
;EndIf
;
; 取得信件的總數量，是一個二維陣列
 Local $list = _Pop3List()
 If Not @error Then
	 MsgBox(0,"The mails count:",UBound($list)-1,5)
 	_ArrayDisplay($list, "")
 Else
 	ConsoleWrite("List commande failed" & @CR)
 EndIf

;dim $split_mails_array[ $list[0] ][4]



for $i =1 to  $list[0]
;for $i =146 to  148 ; $list[0]
	
$single_retr= _retr_split( $i  )
if $single_retr[0]=0 Then
	$single_retr=_retr_split_2( $i )
EndIf
;_ArrayDisplay( $single_retr )
if $single_retr[0]=0 Then
	$single_retr=_retr_split_3( $i )
EndIf


;Example: Write to line 3 of c:\test.txt REPLACING line 3
_FileWriteToLine(@ScriptDir&"\directmail_retr.txt", $i, $i &" ; "& $single_retr[0] &" ; "& $single_retr[1] &" ; "& $single_retr[2] , 1)

;$split_mails_array[ $i ][1]=$single_retr[0]
;$split_mails_array[ $i ][2]=$single_retr[1]
;$split_mails_array[ $i ][3]=$single_retr[2]

next
;_ArrayDisplay( $split_mails_array )




;~ Local $noop = _Pop3Noop()
;~ If Not @error Then
;~ 	ConsoleWrite($noop & @CR)
;~ Else
;~ 	ConsoleWrite("List commande failed" & @CR)
;~ EndIf

;~ Local $uidl = _Pop3Uidl()
;~ If Not @error Then
;~ 	_ArrayDisplay($uidl, "")
;~ Else
;~ 	ConsoleWrite("Uidl commande failed" & @CR)
;~ EndIf

;~ Local $top = _Pop3Top(1, 0)
;~ If Not @error Then
;~ 	ConsoleWrite(StringStripCR($top) & @CR)
;~ Else
;~ 	ConsoleWrite("top commande failed" & @CR)
;~ EndIf

;~ Local $dele = _Pop3Dele(1)
;~ If Not @error Then
;~ 	ConsoleWrite($dele & @CR)
;~ Else
;~ 	ConsoleWrite("Dele commande failed" & @CR)
;~ EndIf

ConsoleWrite(_Pop3Quit() & @CR)
_pop3Disconnect()



func _retr_split( $no )
 Local $retr = _Pop3Retr( $no)
 local $reason[3]
 local $tmp2
 local $tmp1
 
 
	If Not @error Then
		;ConsoleWrite(StringStripCR($retr) & @CR)
		;MsgBox(0,"Retr mail",StringStripCR($retr) & @CR)
		Else
		ConsoleWrite("Retr commande failed" & @CR)
	EndIf

	if StringInStr($retr,"Reason:",1,2,2 ) then 
		$retr=StringTrimleft($retr, StringInStr($retr,"Reason:",1,2)-1 )
		;MsgBox(0,"Mail retr", $retr)
		if  StringInStr($retr,"<",1 ) then 
		  
			if StringInStr($retr,"<",1 ) then
				$reason[0]=StringLeft($retr, StringInStr($retr,"<") -1  )
				;MsgBox(0,"reason 0", $reason[0])
				$tmp1=StringTrimLeft($retr, StringInStr($retr,"<") )
				;MsgBox(0,"tmp", $tmp1)
			EndIf
			if StringInStr($tmp1,">",1 ) then
				;MsgBox(0,"tmp", StringInStr($tmp1,">:",1 )-1)
				;MsgBox(0,"tmp", $tmp1)
				$reason[1]=Stringleft($tmp1, StringInStr($tmp1,">",1) -1  )
				;MsgBox(0 ,"reason 1", $reason[1])
				$tmp2=StringTrimLeft( $tmp1, StringInStr($tmp1,">",1) +1)
			EndIf
	
			$reason[2]=StringStripWS( StringStripCR ($tmp2),7) 
			;MsgBox(0 ,"reason 2", $reason[2])
		Else
			;$tmp2=StringTrimLeft($retr, StringInStr($retr,"=_NextPart_",1 ) +1)
			;$reason[0]=StringStripWS( StringStripCR ($tmp2),4) 
			$reason[0]=StringStripWS( StringStripCR ($retr),7) 
		EndIf 
	  
		return ($reason)
	Else
		
		$reason[0]=0
		return ($reason)	
	EndIf

EndFunc


func _retr_split_2( $no )
	
	; StringStripws( StringStripCR (  StringReplace  ($single_retr  ,chr(10) ," ") ),7 )
 Local $retr =   StringStripws( StringStripCR (  StringReplace  ( _Pop3Retr( $no) ,chr(10) ," ") ),7 ) ;( _Pop3Retr( $no))
 local $reason[3]
 local $tmp2
 local $tmp1
 
 
	If Not @error Then
		;ConsoleWrite(StringStripCR($retr) & @CR)
		;MsgBox(0,"Retr mail",StringStripCR($retr) & @CR)
		Else
		ConsoleWrite("Retr commande failed" & @CR)
	EndIf

	if StringInStr($retr,"Response Message : To:",1,2,1 ) then 
		$retr=StringTrimleft($retr, StringInStr($retr,"Response Message : To:",1,2)+21 )
		  ;MsgBox(0,"Mail retr 1", $retr)
		if  StringInStr($retr,"Subject:",1 ) then 
		  
			if StringInStr($retr,"Subject:",1 ) then
				$reason[1]=StringLeft($retr, StringInStr($retr,"Subject:") -1  )
				;MsgBox(0,"reason 0", $reason[0])
				$tmp1=StringTrimLeft($retr, StringInStr($retr,"Subject:") )
				;MsgBox(0,"tmp", $tmp1)
			EndIf
			if StringInStr($tmp1,"Reason:",1 ) then
				;MsgBox(0,"tmp", StringInStr($tmp1,">:",1 )-1)
				;MsgBox(0,"tmp", $tmp1)
				$reason[0]=StringTrimLeft($tmp1, StringInStr($tmp1,"Reason:",1) -1  )
				;MsgBox(0 ,"reason 1", $reason[1])
				$tmp2=StringTrimLeft( $tmp1, StringInStr($tmp1,")",1) +1)
			EndIf
	
			$reason[2]=StringStripWS( StringStripCR ($tmp2),7) 
			;MsgBox(0 ,"reason 2", $reason[2])
		Else
			;$tmp2=StringTrimLeft($retr, StringInStr($retr,"=_NextPart_",1 ) +1)
			;$reason[0]=StringStripWS( StringStripCR ($tmp2),4) 
			$reason[0]=StringStripWS( StringStripCR ($retr),7) 
		EndIf 
	  
		return ($reason)
	Else
		
		$reason[0]=0
		return ($reason)	
	EndIf

EndFunc



func _retr_split_3( $no )
	
	; StringStripws( StringStripCR (  StringReplace  ($single_retr  ,chr(10) ," ") ),7 )
 Local $retr =   StringStripws( StringStripCR (  StringReplace  ( _Pop3Retr( $no) ,chr(10) ," ") ),7 ) ;( _Pop3Retr( $no))
 local $reason[3]
 local $tmp2
 local $tmp1
 
 
	If Not @error Then
		;ConsoleWrite(StringStripCR($retr) & @CR)
		;MsgBox(0,"Retr mail",StringStripCR($retr) & @CR)
		Else
		ConsoleWrite("Retr commande failed" & @CR)
	EndIf

	if StringInStr($retr,"To:",1,2,2 ) then 
		$retr=StringTrimleft($retr, StringInStr($retr,"To:",1,2,2)+3 )
		  ;MsgBox(0,"Mail retr 1", $retr)
		if  StringInStr($retr,"Subject:",1 ) then 
		  
			if StringInStr($retr,"Subject:",1 ) then
				$reason[1]=StringLeft($retr, StringInStr($retr,"Subject:") -1  )
				;MsgBox(0,"reason 0", $reason[0])
				$tmp1=StringTrimLeft($retr, StringInStr($retr,"Subject:") )
				;MsgBox(0,"tmp", $tmp1)
			EndIf
			if StringInStr($tmp1,"Reason:",1 ) then
				;MsgBox(0,"tmp", StringInStr($tmp1,">:",1 )-1)
				;MsgBox(0,"tmp", $tmp1)
				$reason[0]=StringTrimLeft($tmp1, StringInStr($tmp1,"Reason:",1) -1  )
				;MsgBox(0 ,"reason 1", $reason[1])
				$tmp2=StringTrimLeft( $tmp1, StringInStr($tmp1,")",1) +1)
			EndIf
	
			$reason[2]=StringStripWS( StringStripCR ($tmp2),7) 
			;MsgBox(0 ,"reason 2", $reason[2])
		Else
			;$tmp2=StringTrimLeft($retr, StringInStr($retr,"=_NextPart_",1 ) +1)
			;$reason[0]=StringStripWS( StringStripCR ($tmp2),4) 
			$reason[0]=StringStripWS( StringStripCR ($retr),7) 
		EndIf 
	  
		return ($reason)
	Else
		
		$reason[0]=0
		return ($reason)	
	EndIf

EndFunc

#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

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




dim $dest_dir=@ScriptDir
if FileExists($dest_dir&"\3rd_2_del.txt") then 
	
$return=_file2Array($dest_dir&"\3rd_2_del.txt", 2 ,",")
;; Two dimension array  _file2Array($PathnFile,$aColume,$delimiters)
;Msgbox(0,'Array :' , UBound($return))
;_ArrayDisplay($return)
EndIf

;MsgBox(0,"Array size", UBound($return))
_ArrayDisplay($return)



_pop3Connect($MyPopServer, $MyLogin, $MyPasswd)
If @error Then
	MsgBox(0, "Error", "Unable to connect to " & $MyPopServer & @CR & @error)
	Exit
Else
	ConsoleWrite("Connected to server pop3 " & $MyPopServer & @CR)
EndIf



;Local $stat = _Pop3Stat()
;If Not @error Then
;	_ArrayDisplay($stat, "Result of STAT COMMAND")
;Else
;	ConsoleWrite("Stat commande failed" & @CR)
;EndIf
;
; 取得信件的總數量，是一個二維陣列
; Local $list = _Pop3List()
; If Not @error Then
;	 MsgBox(0,"The mails count:",UBound($list)-1,5)
; 	_ArrayDisplay($list, "")
; Else
; 	ConsoleWrite("List commande failed" & @CR)
; EndIf

;dim $split_mails_array[ $list[0] ][4]






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

;for $i =1 to $mail_list_2_d[0]

for $i=0 to UBound($return)-1
	;MsgBox(0,"List mail no and name", $return[$i][0]&" <-----> "&$return[$i][1], 1)
	$single_retr = _Pop3Retr( $return[$i][0] )
	if  StringInStr(  StringStripws( StringStripCR (  StringReplace  ($single_retr  ,chr(10) ," ") ),7 ) , "<"&$return[$i][1])&">" Then
		;MsgBox(0,"Good for one mail", $return[$i][0])
		_pop_retr_single(  $return[$i][0] )
		 Local $dele = _Pop3Dele( $return[$i][0]  )
			If Not @error Then
				ConsoleWrite($dele & @CR)
			EndIf
	Else
		_pop_retr_single( $single_retr  )
	EndIf
next

;~ Local $dele = _Pop3Dele(1)
;~ If Not @error Then
;~ 	ConsoleWrite($dele & @CR)
;~ Else
;~ 	ConsoleWrite("Dele commande failed" & @CR)
;~ EndIf

ConsoleWrite(_Pop3Quit() & @CR)
_pop3Disconnect()



func _pop_retr_single( $single_retr_1 )
	$a= StringStripws( StringStripCR (  StringReplace  ($single_retr_1  ,chr(10) ," ") ),7 )
;local $no  ;=499
;$single_retr = _Pop3Retr( $no )
;MsgBox(0,"Retr", $single_retr)

$file = FileOpen(@ScriptDir& "\error_mail\test.txt", 9)

; Check if file opened for writing OK
If $file = -1 Then
    MsgBox(0, "Error", "Unable to open file.")
    Exit
EndIf
FileWrite($file, $a & @CRLF)
FileClose($file)

;$single_retr= _retr_split_2( $no  )
;_ArrayDisplay( $single_retr )

;MsgBox(0,"This is a sigle message", "Wait if you going to real message")
EndFunc
;;================================





;; Two dimension array
Func _file2Array($PathnFile,$aColume,$delimiters)

	
	local $aRecords
	If Not _FileReadToArray($PathnFile,$aRecords) Then
		MsgBox(4096,"Error", " Error reading file '"&$PathnFile&"' to Array   error:" & @error)
		Exit
	EndIf
		;c
	local $TextToArray[$aRecords[0]][$aColume+1]
	;$TextToArray[0][0]=$aRecords[0]
	local $aRow
	For $y = 1 to $aRecords[0]
		;Msgbox(0,'Record:' & $y, $aRecords[$y])
		 
		$aRow=StringSplit($aRecords[$y],$delimiters)
		;Msgbox(0,'X ,Colume :', $aRow[0])
		For $x=1 to $aRow[0]
			if StringInStr($aRow[$x],",") then 
			
			$aRow[$x]=StringTrimLeft($aRow[$x],1)
			;MsgBox(0, "after", $aRow[$x])
			EndIf
			 $TextToArray[$y-1][$x-1]=$aRow[$x]
		next
	Next
	
	;_ArrayDisplay($TextToArray)
	Return $TextToArray

EndFunc


#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

;#include <Word.au3>
#include <file.au3>
#include <array.au3>
#Include <GDIPlus.au3>

;$oWordApp = _WordCreate ("")
;$oDoc = _WordDocOpen ($oWordApp, @ScriptDir & "\test2.doc")
;_WordDocPrint ($oDoc)
;_WordQuit ($oWordApp, 0)


Func _print_word($doc_file , $h_name ,$h_id, $hotel_phone, $hotel_fax, $hotel_address)

local $work_dir=@ScriptDir & "\hotel_data_files\"
local $img_dir;="d:\pdf_test\"
local $word_exe_path;="D:\Green_app\Office_2003_Portable\office2003\winword.exe "
;local $doc_file , $h_name ,$h_id
local $sleep_time=5000
local $png_lst
local $index, $b_index, $order_confirm, $orderRule, $orderRule_e , $rule, $rule_e , $image_url, $image_hight, $image_hight_sum , $br_counter , $br_string



if FileExists (@ScriptDir & "\print_word.txt") then
$img_dir=  StringTrimLeft ( FileReadLine( @ScriptDir & "\print_word.txt" , 1) , 8 )
$word_exe_path = StringTrimLeft (  FileReadLine( @ScriptDir & "\print_word.txt" , 2) , 14 )

;MsgBox(0,"read from file",  $img_dir & @CRLF  & $word_exe_path)
EndIf

;  $doc_file="test3.doc"
;  $h_name="¤s¤ô¥Ð¶®ªÙ"
;  $h_id=221
;  $hotel_phone="0912345678"
;  $hotel_fax="021234567"
;  $hotel_address="abcdef sdfsdlfksdflsdlsd"




;MsgBox(0,"Print and wait", $img_dir &  StringTrimRight( $doc_file, 4 ) & "_1.png" )


			;$png_lst=_FileListToArray ( $img_dir , "*.png" )
			;_ArrayDisplay ( $png_lst )
	;MsgBox(0, "run" , $word_exe_path &"  " & $work_dir  & $h_id &"\"& $doc_file &  " /q /n /mFilePrintDefault /mFileExit")

	run( $word_exe_path &"  " & $work_dir  & $h_id &"\"& $doc_file &  " /q /n /mFilePrintDefault /mFileExit")

	ProcessWait("gswin32c.exe")
	sleep(5000)
	;MsgBox(0,"Print and wait", $img_dir &  StringTrimRight( $doc_file,4 ) & "_1.png" )


	$doc_file=StringTrimRight( $doc_file,4 )
		ProcessWaitClose ("gswin32c.exe" )
		MsgBox(0,"gswin32c.exe ", "Gswin32c.exe closed", 5)

	if FileExists ($img_dir &  $doc_file &"_1.png") then
		$png_lst=_FileListToArray ( $img_dir,"*.png")
		;_ArrayDisplay ( $png_lst )
	EndIf

  $image_hight=0
  $image_hight_sum=0

for $i=1 to UBound ($png_lst)-1
		;MsgBox( 0,"2nd run ", "convert.exe -trim " & $img_dir & $png_lst[$i] &" -bordercolor white -border 20x20 " & $work_dir & $h_id &"\orderRule_"& $i &".png" )
		;run ("convert.exe -trim " & $png_lst[$i] &" -bordercolor white -border 20x20 " &  StringTrimRight( $png_lst[$i] ,4) & "_c.png")
		run ("convert.exe -trim " & $img_dir & $png_lst[$i] &" -bordercolor white -border 10x10 " & $work_dir & $h_id &"\orderRule_"& $i &".png" )
		ProcessWaitClose ("convert.exe" )
		sleep(1000)
		;if FileExists ($img_dir & StringTrimRight( $png_lst[($png_lst)-1] ,4) & "_c.png") then exitloop

		$image_hight= _get_image_height( $work_dir & $h_id &"\orderRule_"& $i &".png" )
		$image_hight_sum=$image_hight_sum+ $image_hight
next


for $i=1 to UBound ($png_lst)-1
	;	sleep(1000)
		if FileExists ($img_dir & $png_lst[$i]) then  FileDelete ( $img_dir & $png_lst[$i] )
		$png_lst[$i]="orderRule_"&$i &".png"
	;	;if FileExists ($img_dir & StringTrimRight( $png_lst[($png_lst)-1] ,4) & "_c.png") then exitloop
next

 ;_ArrayDisplay ( $png_lst)
 ; create file
 ; orderRule.htm , rule.htm , index.htm , k_index.htm, orderConfirm.htm
 $index =_ki_template_html( "k_index.htm" ,$h_name, $h_id )
 $b_index =_ki_template_html( "b_index.htm" ,$h_name, $h_id )
 $order_confirm= StringReplace( _ki_template_html( "k_orderConfirm.htm" ,$h_name, $h_id )  ,"[$tel_mobile_address]" , $hotel_phone &" "& $hotel_fax &"  "& $hotel_address  )
; $orderRule=_ki_template_html( "orderRule.htm" ,$h_name, $h_id )
 ;$orderRule_e=_ki_template_html( "orderRule_e.htm" ,$h_name, $h_id )
 ;$rule=_ki_template_html( "rule.htm" ,$h_name, $h_id )
 ;$rule_e=_ki_template_html( "rule_e.htm" ,$h_name, $h_id )

 ;MsgBox (0, "template", _ki_template_html( "k_index.htm" ,$h_name, $h_id )  )
	if FileExists ($work_dir & $h_id & "\index.htm" )  then FileDelete ( $work_dir & $h_id & "\index.htm" )
	FileWrite(  $work_dir & $h_id & "\index.htm", $index )

	if FileExists ($work_dir & $h_id & "\b_index.htm" )  then FileDelete ( $work_dir & $h_id & "\b_index.htm" )
	FileWrite(  $work_dir & $h_id & "\b_index.htm", $b_index )

	if FileExists ($work_dir & $h_id & "\orderConfirm.htm" )  then FileDelete ( $work_dir & $h_id & "\orderConfirm.htm" )
	FileWrite(  $work_dir & $h_id & "\orderComfirm.htm", $order_confirm )


  $orderRule=_ki_template_html( "orderRule.htm" ,$h_name, $h_id )
  ;MsgBox(0,"", _get_image_height (@ScriptDir&"\o.jpg") )
		;(x-110-17*y)/17
	$br_counter=  int( ( $image_hight_sum -110 -( 17 * UBound($png_lst)-2 ) )/17  ) +2
	$br_string=""
	;MsgBox(0, " image height and  br counter", $image_hight_sum & " ; " & $br_counter )
	for $y=1 to $br_counter
		$br_string = $br_string & "<br>"
		if mod($y,10)=0 then  $br_string=$br_string & @CRLF
	next

  for $y=1 to UBound($png_lst)-1
		$image_url='<p><img src="http://www.bnbhome.com.tw/hotel_rule/[h_id]/[png_name]"></p>'
		$orderRule=StringReplace ( $orderRule, $image_url &@CRLF &"</body></html>" ,"")
		$image_url=StringReplace($image_url, "[h_id]", $h_id)

		$image_url= StringReplace( $image_url , "[png_name]", 'orderRule_'& $y &'.png')
		;MsgBox (0,"y " , 'orderRule_'& $y &'.png' )
		$orderRule=$orderRule &$image_url  & @CRLF

	next

	if FileExists ($work_dir & $h_id & "\rule.htm" )  then FileDelete ( $work_dir & $h_id & "\rule.htm" )
	FileWrite(  $work_dir & $h_id & "\rule.htm", $orderRule & @CRLF &"</body></html>" )


	$orderRule =$orderRule & @CRLF & $br_string & @CRLF &"</body></html>"
	;$orderRule=  StringReplace ($orderRule ,$image_url, '<p><img src="http://www.bnbhome.com.tw/hotel_rule/'&$h_id&'/orderRule_'&$p&'.png"></p>' )

	;MsgBox( 0,"orderRule.htm", $orderRule )
	if FileExists ($work_dir & $h_id & "\orderRule.htm" )  then FileDelete ( $work_dir & $h_id & "\orderRule.htm" )
	FileWrite(  $work_dir & $h_id & "\orderRule.htm", $orderRule )





EndFunc
;Exit


func _ki_template_html( $html_file_name, $h_name, $h_id)
	local $k_index , $temp_array;, $b_index, $order_confirm, $orderRule, $orderRule_e , $rule, $rule_e
	_FileReadToArray ( @ScriptDir &"\"&$html_file_name  ,$temp_array)
	for $in=1 to UBound ($temp_array )-1
		;$temp_array[$in]=StringReplace ( StringReplace($temp_array[$in], "[hotel_name]", $h_name) , "[h_id]" , $h_id )
		$k_index= $k_index & $temp_array[$in] & @CRLF
	next
	;_ArrayDisplay ($temp_array )
	$k_index= StringReplace ( StringReplace($k_index, "[$hotel_name]", $h_name) , "[$h_id]" , $h_id )

	;MsgBox(0,"k_index.htm", $k_index)
	return $k_index
EndFunc


func _ki_template_html_direct($h_name, $h_id)
	local $k_index , $b_index, $order_confirm, $orderRule, $orderRule_e , $rule, $rule_e , $temp_array

	_FileReadToArray ( @ScriptDir &"\k_index.htm"  ,$temp_array)
	for $in=1 to UBound ($temp_array )-1
		;$temp_array[$in]=StringReplace ( StringReplace($temp_array[$in], "[hotel_name]", $h_name) , "[h_id]" , $h_id )
		$k_index= $k_index & $temp_array[$in] & @CRLF
	next
	;_ArrayDisplay ($temp_array )
	$k_index= StringReplace ( StringReplace($k_index, "[hotel_name]", $h_name) , "[h_id]" , $h_id )

	MsgBox(0,"k_index.htm", $k_index)



EndFunc







func _get_image_height($image_file)
_GDIPlus_Startup ()
Local $hImage = _GDIPlus_ImageLoadFromFile( $image_file )
If @error Then
    MsgBox(16, "Error", "Does the "& $image_file &" exist?")
    Exit 1
EndIf

;ConsoleWrite(_GDIPlus_ImageGetWidth($hImage) & @CRLF)
$h=_GDIPlus_ImageGetHeight($hImage)
;ConsoleWrite(_GDIPlus_ImageGetHeight($hImage) & @CRLF)
_GDIPlus_ImageDispose ($hImage)
_GDIPlus_ShutDown ()
return ( $h )
EndFunc
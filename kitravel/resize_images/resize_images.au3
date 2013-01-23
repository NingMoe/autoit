

#include <file.au3>
#include <array.au3>
#include <GDIPlus.au3>
#include <Constants.au3>

Local $work_dir = @ScriptDir&"\output" ; output dir
Local $img_dir = @ScriptDir;&"\input" ; orighnal dir
Local $imgbak_dir = @ScriptDir&"\image_bak" ; orighnal dir
Local $sleep_time = 5000
Local $png_lst
Local $png_name_by_converter = ""
Local $image_type[4]
Local $resize[3]
Local $image_list[1]
$image_list[0] = 0

local $image_w_h
; Resize image


$resize[0] = "450x330" ;(檔案不得大於 3M)
$resize[1] = "400x200" ;(檔案不得大於 5M)
$resize[2] = "120x130" ;(檔案不得大於 1M)

$size_boundry_0_w=450
$size_boundry_0_h=330

$size_boundry_1_w=400
$size_boundry_1_h=200

$size_boundry_2_w=120
$size_boundry_2_h=130

$image_type[0] = "jpg"
$image_type[1] = "jpeg"
$image_type[2] = "png"
$image_type[3] = "bmp"


if not FileExists ($work_dir) then DirCreate( $work_dir )
;if not FileExists ($img_dir) then DirCreate( $img_dir )
if not FileExists ($imgbak_dir) then DirCreate( $imgbak_dir )

For $x = 0 To 3
	$png_lst=''
	If FileExists($img_dir & "\*." & $image_type[$x]) Then
		$png_lst = _FileListToArray($img_dir, "*." & $image_type[$x], 1)
		;_ArrayDisplay($png_lst, "png_list" & $x )
	EndIf

	If IsArray($png_lst) Then
		;_ArrayDisplay ( $png_lst, "$png_lst ")
		;MsgBox(0,"png_list", "PNG file no in List: "&UBound($png_lst)-1 )

		 $image_list[0]= $image_list[0]  + $png_lst [0]
		 _ArrayDelete ($png_lst, 0)
		_ArrayConcatenate($image_list, $png_lst)

		;_ArrayDisplay($image_list, "$image_list " & $x )
		;MsgBox(0, "image_list", "Image files no in List: " & UBound($image_list) - 1)


	EndIf

Next


;_ArrayDisplay($image_list, "$image_list "  )


For $x = 1 To UBound($image_list) - 1
	;MsgBox( 0,"2nd run ", "convert.exe -trim " & $img_dir & $png_lst[$i] &" -bordercolor white -border 20x20 " & $work_dir & $h_id &"\orderRule_"& $i &".png" )
	;run ("convert.exe -trim " & $png_lst[$i] &" -bordercolor white -border 20x20 " &  StringTrimRight( $png_lst[$i] ,4) & "_c.png")
	;ConsoleWrite( @CRLF &"convert.exe " & $img_dir & "\" & StringStripWS( $image_list[$x],8 ) & "-encoding utf8  -resize " & $resize[0] & "^ " & $work_dir & "\" &   Stringleft ($image_list[$x], StringInStr( $image_list[$x],".")-1 )  & "_" & $resize[0] &"."& StringTrimLeft ($image_list[$x], StringInStr( $image_list[$x],".") ) )
	;MsgBox(0,"WxH", _get_image_width_hight($img_dir & '\' & $image_list[$x] ) )
	$image_w_h=_get_image_width_hight($img_dir & '\' & $image_list[$x] )
	;_ArrayDisplay ($a)
	;MsgBox(0, "convert " & $x , "convert.exe " & $img_dir & "\" & $image_list[$x] & " -encoding utf8 -resize " & $resize[0] & "^ " & $work_dir & "\" &   StringLeft ($image_list[$x], StringInStr( $image_list[$x],".")-1 )  & "_" & $resize[0] &"."& StringTrimLeft ($image_list[$x], StringInStr( $image_list[$x],".") ) )
	if not ( ($image_w_h[0] < $size_boundry_0_w ) and ($image_w_h[1] < $size_boundry_0_h ) )then
		run ('convert.exe "' & $img_dir & '\' & $image_list[$x] & '" -encoding utf8 -resize ' & $resize[0] & ' "' & $work_dir & '\' &   StringLeft ($image_list[$x], StringInStr( $image_list[$x],'.')-1 )  & '_' & $resize[0] &'.'& StringTrimLeft ($image_list[$x], StringInStr( $image_list[$x],'.')&'"' ),"",  $STDERR_CHILD + $STDOUT_CHILD )
		ProcessWaitClose ("convert.exe" )
		Sleep(1000)
	Else
			FileCopy ($img_dir & '\' & $image_list[$x], $work_dir & '\' & $image_list[$x])
	EndIf
	;if FileExists ($img_dir & StringTrimRight( $png_lst[($png_lst)-1] ,4) & "_c.png") then exitloop
	if not( ($image_w_h[0] < $size_boundry_1_w ) and ($image_w_h[1] < $size_boundry_1_h ) )then
		run ('convert.exe "' & $img_dir & '\' & $image_list[$x] & '" -encoding utf8 -resize ' & $resize[1] & ' "' & $work_dir & '\' &   StringLeft ($image_list[$x], StringInStr( $image_list[$x],'.')-1 )  & '_' & $resize[1] &'.'& StringTrimLeft ($image_list[$x], StringInStr( $image_list[$x],'.')&'"' ),"",  $STDERR_CHILD + $STDOUT_CHILD )
		ProcessWaitClose ("convert.exe" )
		Sleep(1000)
	EndIf
	if not ( ($image_w_h[0] < $size_boundry_2_w ) and ($image_w_h[1] < $size_boundry_2_h ) ) then
		run ('convert.exe "' & $img_dir & '\' & $image_list[$x] & '" -encoding utf8 -resize ' & $resize[2] & ' "' & $work_dir & '\' &   StringLeft ($image_list[$x], StringInStr( $image_list[$x],'.')-1 )  & '_' & $resize[2] &'.'& StringTrimLeft ($image_list[$x], StringInStr( $image_list[$x],'.')&'"' ),"",  $STDERR_CHILD + $STDOUT_CHILD )
		ProcessWaitClose ("convert.exe" )
		Sleep(1000)
	EndIf
	FileMove( $img_dir & '\' & $image_list[$x] , $imgbak_dir&'\' ,9 )
Next



Exit



func _get_image_width_hight($image_file)
_GDIPlus_Startup ()
Local $hImage = _GDIPlus_ImageLoadFromFile( $image_file )
local $w, $h, $wh[2]

If @error Then
    MsgBox(16, "Error", "Does the "& $image_file &" exist?")
    Exit 1
EndIf
$w=_GDIPlus_ImageGetWidth($hImage)
ConsoleWrite(_GDIPlus_ImageGetWidth($hImage) & @CRLF)
$h=_GDIPlus_ImageGetHeight($hImage)
ConsoleWrite(_GDIPlus_ImageGetHeight($hImage) & @CRLF)
_GDIPlus_ImageDispose ($hImage)
_GDIPlus_ShutDown ()
$wh[0]=$w
$wh[1]=$h
;return ( $w&"x"&$h   )
return ( $wh )

EndFunc
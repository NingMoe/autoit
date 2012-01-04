#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.2.12.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#Include <File.au3>
#include <array.au3>
dim $pnr_path=@ScriptDir&"\pnr\"
dim $pnr_path="z:\travel_result\"
dim $pnr_list_2_rename[1]
dim $increase=0
$pnr_list=_FileListToArray($pnr_path,"*.rtf",1)


for $i= $pnr_list[0] to 1  step -1
	
	if StringInStr($pnr_list[$i],"_",0,3) Then
		;MsgBox( 0, "PNR 2 rename", $pnr_list[$i]&" ---> "& StringLeft($pnr_list[$i],6)&".rtf" )
		_Arrayadd($pnr_list_2_rename,$pnr_list[$i])
		$increase=$increase+1
		EndIf
Next

$pnr_list_2_rename[0]=$increase
;_ArrayDisplay($pnr_list)
;_ArrayDisplay($pnr_list_2_rename)

for $i=1 to $pnr_list_2_rename[0]
	
	;MsgBox(0," no", StringInStr($pnr_list_2_rename[$i],"_") )
	if StringInStr($pnr_list_2_rename[$i],"_")=7 Then 
	;MsgBox(0," to rename",$pnr_path&$pnr_list_2_rename[$i] &" ----> "& $pnr_path& StringLeft($pnr_list_2_rename[$i],6)&".rtf" )
	
	Filecopy($pnr_path&$pnr_list_2_rename[$i] , $pnr_path& StringLeft($pnr_list_2_rename[$i],6)&".rtf",1)
	if FileExists($pnr_path& StringLeft($pnr_list_2_rename[$i],6)&".rtf") then FileMove($pnr_path & $pnr_list_2_rename[$i],  $pnr_path&"\renamed\",8)
	EndIf
Next
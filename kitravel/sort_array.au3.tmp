; sort array

#include <array.au3>

dim $all_doc_in_hid[5]

dim $s ,$s1
dim $swap

;$all_doc_in_hid

; 山水田雅舍-入住須知.doc
; 山水田雅舍-入住須知_e.doc
; 山水田雅舍-訂房須知.doc
; 山水田雅舍-訂房須知_e.doc

$all_doc_in_hid[0]=4
$all_doc_in_hid[1]="山水田雅舍-訂房須知.doc"
$all_doc_in_hid[2]=  "山水田雅舍-訂房須知_e.doc"
$all_doc_in_hid[3]= "山水田雅舍-入住須知.doc"
$all_doc_in_hid[4]= "山水田雅舍-入住須知_e.doc"


;_ArrayDisplay($all_doc_in_hid)

;$s = _ArraySearch($all_doc_in_hid, "-入住須知_d.doc", 0, 0, 0, 1)
;ConsoleWrite (@CRLF& _ArraySearch($all_doc_in_hid, "-入住須知_d.doc", 0, 0, 0, 1) & @CRLF)
$s=0
$s1=0
$s = _ArraySearch($all_doc_in_hid, "-入住須知.doc", 0, 0, 0, 1)
;ConsoleWrite (@CRLF& _ArraySearch($all_doc_in_hid, "-入住須知.doc", 0, 0, 0, 1) & @CRLF)
if $s >0 then
	$swap=""
	$s1 = _ArraySearch($all_doc_in_hid, "-訂房須知.doc", 0, 0, 0, 1)
	if $s1< $s then
		$swap=$all_doc_in_hid[$s1]
		$all_doc_in_hid[$s1]= $all_doc_in_hid[$s]
		$all_doc_in_hid[$s]=$swap
	EndIf

EndIf

$s=0
$s1=0
$s = _ArraySearch($all_doc_in_hid, "-入住須知_e.doc", 0, 0, 0, 1)
if $s >0 then
	$swap=""
	$s1 = _ArraySearch($all_doc_in_hid, "-訂房須知_e.doc", 0, 0, 0, 1)
	if $s1< $s then
		$swap=$all_doc_in_hid[$s1]
		$all_doc_in_hid[$s1]= $all_doc_in_hid[$s]
		$all_doc_in_hid[$s]=$swap
	EndIf

EndIf

_ArrayDisplay($all_doc_in_hid)




; sort array

#include <array.au3>

dim $all_doc_in_hid[5]

dim $s ,$s1
dim $swap

;$all_doc_in_hid

; �s���ж���-�J����.doc
; �s���ж���-�J����_e.doc
; �s���ж���-�q�ж���.doc
; �s���ж���-�q�ж���_e.doc

$all_doc_in_hid[0]=4
$all_doc_in_hid[1]="�s���ж���-�q�ж���.doc"
$all_doc_in_hid[2]=  "�s���ж���-�q�ж���_e.doc"
$all_doc_in_hid[3]= "�s���ж���-�J����.doc"
$all_doc_in_hid[4]= "�s���ж���-�J����_e.doc"


;_ArrayDisplay($all_doc_in_hid)

;$s = _ArraySearch($all_doc_in_hid, "-�J����_d.doc", 0, 0, 0, 1)
;ConsoleWrite (@CRLF& _ArraySearch($all_doc_in_hid, "-�J����_d.doc", 0, 0, 0, 1) & @CRLF)
$s=0
$s1=0
$s = _ArraySearch($all_doc_in_hid, "-�J����.doc", 0, 0, 0, 1)
;ConsoleWrite (@CRLF& _ArraySearch($all_doc_in_hid, "-�J����.doc", 0, 0, 0, 1) & @CRLF)
if $s >0 then
	$swap=""
	$s1 = _ArraySearch($all_doc_in_hid, "-�q�ж���.doc", 0, 0, 0, 1)
	if $s1< $s then
		$swap=$all_doc_in_hid[$s1]
		$all_doc_in_hid[$s1]= $all_doc_in_hid[$s]
		$all_doc_in_hid[$s]=$swap
	EndIf

EndIf

$s=0
$s1=0
$s = _ArraySearch($all_doc_in_hid, "-�J����_e.doc", 0, 0, 0, 1)
if $s >0 then
	$swap=""
	$s1 = _ArraySearch($all_doc_in_hid, "-�q�ж���_e.doc", 0, 0, 0, 1)
	if $s1< $s then
		$swap=$all_doc_in_hid[$s1]
		$all_doc_in_hid[$s1]= $all_doc_in_hid[$s]
		$all_doc_in_hid[$s]=$swap
	EndIf

EndIf

_ArrayDisplay($all_doc_in_hid)




#include "..\MPDF_UDF.au3"

_SelectImages()

Func _SelectImages()
	Local $var = FileOpenDialog("Select images", @ScriptDir & "\", "Images (*.jpg;*.bmp;*gif;*png;*tif;*ico)", 4)
	If @error Then
		MsgBox(4096,"","No File(s) chosen")
	Else
		Local $aImgs = StringSplit($var, "|", 3)
		;set the properties for the pdf
		_SetTitle("Image2PDF")
		_SetSubject("Convert image(s) to pdf")
		_SetKeywords("pdf, AutoIt")
		_OpenAfter(True);open after generation
		_SetUnit($PDF_UNIT_CM)
		_SetPaperSize("a4")
		_SetZoomMode($PDF_ZOOM_CUSTOM,90)
		_SetOrientation($PDF_ORIENTATION_PORTRAIT)
		_SetLayoutMode($PDF_LAYOUT_CONTINOUS)

		;initialize the pdf
		_InitPDF(@ScriptDir & "\Image2PDF.pdf")

		If UBound($aImgs)<>1 Then
			;=== load resources used in pdf ===
			For $i=1 To UBound($aImgs)-1
				_LoadResImage("img"&$i, $aImgs[0] & "\" & $aImgs[$i])
			Next
			;load each image on it's own page
			For $i = 1 To UBound($aImgs)-1
				_BeginPage()
					;scale image to paper size!
					_InsertImage("img"&$i, 0, 0, _GetPageWidth()/_GetUnit(), _GetPageHeight()/_GetUnit())
				_EndPage()
			Next
		Else
			_LoadResImage("taietel", $aImgs[0])
			_BeginPage()
			;scale image to paper size!
			_InsertImage("taietel", 0, 0, _GetPageWidth()/_GetUnit(), _GetPageHeight()/_GetUnit())
			_EndPage()
		EndIf
		;then, finally, write the buffer to disk
		_ClosePDFFile()
	EndIf
EndFunc

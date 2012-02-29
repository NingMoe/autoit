#include "..\MPDF_UDF.au3"
#include <File.au3>

Global $sF = FileOpenDialog("Choose a text file", @ScriptDir & "\", "Text file (*.au3;*.txt;*.ini;*cfg)", 1)

If @error Then
	MsgBox(4096, "", "No File(s) chosen")
Else
	;set the properties for the pdf
	_SetTitle("Txt2PDF")
	_SetSubject("Convert text file to pdf")
	_SetKeywords("pdf, AutoIt")
	_OpenAfter(True);open after generation
	_SetUnit($PDF_UNIT_CM)
	_SetPaperSize("A4")
	_SetZoomMode($PDF_ZOOM_CUSTOM, 90)
	_SetOrientation($PDF_ORIENTATION_PORTRAIT)
	_SetLayoutMode($PDF_LAYOUT_CONTINOUS)

	;initialize the pdf
	_InitPDF(@ScriptDir & "\Txt2Pdf.pdf")
	_LoadFontTT("F1", $PDF_FONT_CALIBRI,$PDF_FONT_ITALIC)

	_Txt2PDF($sF, "F1")

	;write the buffer to disk
	_ClosePDFFile()
EndIf

; #FUNCTION# ====================================================================================================================
; Name ..........: _Txt2PDF
; Description ...: Convert a text file to pdf
; Syntax ........: _Txt2PDF( $sText , $sFontAlias  )
; Parameters ....: $sText               -  file path.
;                  $sFontAlias          -  font alias.
; Return values .: None
; Author(s) .....: Mihai Iancu (taietel at yahoo dot com)
; Modified ......:
; Remarks .......: If the string is very long, it will be scaled to paper width
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Txt2PDF($sFile, $sFontAlias)
	Local $hFile = FileOpen($sFile)
	Local $sText = FileRead($hFile)
	FileClose($hFile)

	Local $iUnit = Ceiling(_GetUnit())
	Local $iX = 2
	Local $iY = Ceiling(_GetPageHeight() / _GetUnit()) - 1.5
	Local $iPagina = Ceiling(_GetPageWidth() / $iUnit) - $iX
	Local $iWidth = Ceiling($iPagina - $iX);, 1)
	Local $lScale
	Local $iRanduri = StringSplit($sText & @CRLF & @CRLF & @CRLF & @CRLF, @CRLF, 3)
	Local $iHR = 0.5 * Ceiling($iY / (10 * $iUnit))
	Local $iPages = Ceiling((UBound($iRanduri)) * $iHR / $iY)
	Local $iNrRanduri = Ceiling(UBound($iRanduri) / $iPages-2)
	Local $nrp
	For $j = 0 To $iPages + 2
		$nrp = _BeginPage()
		_DrawText(_GetPageWidth()/_GetUnit()-1, 1, $nrp, "F1", 10, $PDF_ALIGN_CENTER)
		For $i = 0 To $iNrRanduri - 1
			Local $sLength = Round(_GetTextLength($iRanduri[$i + $j * $iNrRanduri], $sFontAlias, 10))
			Local $iH = $iY - $iHR * ($i + 1)
			Select
				Case $iH < 1
					_EndPage()
				Case $i + $j * $iNrRanduri = UBound($iRanduri) - 1
					_EndPage()
					Return
				Case $sLength > $iWidth - 1
					$lScale = Ceiling($iWidth * 100 / $sLength)
					_SetTextHorizontalScaling($lScale)
					_DrawText($iX, $iH, $iRanduri[$i + $j * $iNrRanduri], $sFontAlias, 10, $PDF_ALIGN_LEFT, 0)
					_SetTextHorizontalScaling(100)
				Case Else
					_DrawText($iX, $iH, $iRanduri[$i + $j * $iNrRanduri], $sFontAlias, 10, $PDF_ALIGN_LEFT, 0)
			EndSelect
		Next
		_EndPage()
	Next
EndFunc   ;==>_Txt2PDF
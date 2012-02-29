#include "..\MPDF_UDF.au3"

;set the properties for the pdf
_SetTitle("Demo Table PDF in AutoIt")
_SetSubject("Demo Table PDF in AutoIt, with formating")
_SetKeywords("pdf, demo, table, AutoIt")
_OpenAfter(True);open after generation
_SetUnit($PDF_UNIT_CM)
_SetPaperSize("CUSTOM",841.890, 595.276); A4 landscape
_SetZoomMode($PDF_ZOOM_FULLPAGE)
_SetOrientation($PDF_ORIENTATION_PORTRAIT)
_SetLayoutMode($PDF_LAYOUT_CONTINOUS)

;initialize the pdf
_InitPDF(@ScriptDir & "\Example_Table.pdf")

;=== load used font(s) ===
;fonts: Garamond
_LoadFontTT("_CalibriB", $PDF_FONT_CALIBRI, $PDF_FONT_BOLD)
_LoadFontTT("_CalibriI", $PDF_FONT_CALIBRI, $PDF_FONT_ITALIC)
_LoadFontTT("_Calibri", $PDF_FONT_CALIBRI)

;begin page
_BeginPage()
	_InsertTable(2, 2,0,0,10,10);table with 10x10 cells, 2cm from the bottom-left, width/height auto
	$sTit = "Sample pdf table generated with AutoIt"
	_SetTextRenderingMode(1)
	_InsertRenderedText((_GetPageWidth()/_GetUnit())/2, _GetPageHeight()/_GetUnit()-1.5, $sTit, "_CalibriI", 32, 100, $PDF_ALIGN_CENTER, 0xbbbbbb, 0x202020)
	_SetTextRenderingMode(0)
_EndPage()
_BeginPage()
	_InsertTable(2, 8,0,0,8,5);table with 8x5 cells, 2cm from the left,8cm from bottom, width/height auto
	$sTit = "Sample pdf table generated with AutoIt"
	_SetTextRenderingMode(2)
	_InsertRenderedText((_GetPageWidth()/_GetUnit())/2, _GetPageHeight()/_GetUnit()-1.5, $sTit, "_CalibriB", 32, 100, $PDF_ALIGN_CENTER, 0xbbbbbb, 0x202020)
	_SetTextRenderingMode(0)
_EndPage()
_BeginPage()
	_InsertTable(2, 8,20,0,8,5)
	$sTit = "Sample pdf table generated with AutoIt"
	_SetTextRenderingMode(3)
	_InsertRenderedText((_GetPageWidth()/_GetUnit())/2, _GetPageHeight()/_GetUnit()-1.5, $sTit, "_CalibriI", 32, 100, $PDF_ALIGN_CENTER, 0xbbbbbb, 0x202020)
	_SetTextRenderingMode(0)
_EndPage()
_BeginPage()
	_InsertTable(8, 2,0,15,8,15)
	$sTit = "Sample pdf table generated with AutoIt"
	_SetTextRenderingMode(5)
	_InsertRenderedText((_GetPageWidth()/_GetUnit())/2, _GetPageHeight()/_GetUnit()-1.5, $sTit, "_CalibriB", 32, 100, $PDF_ALIGN_CENTER, 0xbbbbbb, 0x202020)
	_SetTextRenderingMode(0)
_EndPage()
;write the buffer to disk
_ClosePDFFile()

Func _InsertTable($iX, $iY, $iW=0, $iH=0, $iCols=0, $iRows=0,$lTxtColor = 0x000000, $lBorderColor = 0xdddddd)
	Local $iPgW = Round(_GetPageWidth()/_GetUnit(),1)
	Local $iPgH = Round(_GetPageHeight()/_GetUnit(),1)
	If $iW = 0 Then $iW = $iPgW - $iX -2
	If $iH = 0 Then $iH = $iPgH - $iY -2
	_SetColourStroke($lBorderColor)
	_Draw_Rectangle($iX, $iY, $iW, $iH, $PDF_STYLE_STROKED, 0, 0xfefefe, 0.01)
	_SetColourStroke(0)
	Local $iColW = $iW/$iCols
	Local $iRowH = $iH/$iRows
	Local $lRGB
	For $i = 0 To $iRows-1
		For $j = 0 To $iCols-1
			If $i=0 Then
				$lRGB = 0xefefef
			Else
				$lRGB = 0xfefefe
			EndIf
			_SetColourStroke($lBorderColor)
			_Draw_Rectangle($iX+$j*$iColW, $iY+$iH-($i+1)*$iRowH, $iColW, $iRowH, $PDF_STYLE_STROKED, 0, $lRGB, 0.01)
			_SetColourStroke(0)
			Local $sText = "Row "&$i&": Col "&$j
			Local $sLength = Round(_GetTextLength($sText, "_Calibri", 10),1)
			$lScale = Ceiling(0.75*$iColW * 100/ $sLength)
			_SetColourFill($lTxtColor)
			_SetTextHorizontalScaling($lScale)
			_DrawText($iX+$j*$iColW+$iColW/10, $iY+$iH-($i+1)*$iRowH + ($iRowH-10/_GetUnit())/2, $sText, "_Calibri", 10, $PDF_ALIGN_LEFT, 0)
			_SetTextHorizontalScaling(100)
			_SetColourFill(0)
		Next
	Next
EndFunc

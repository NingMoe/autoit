#include "..\MPDF_UDF.au3"

;set the properties for the pdf
_SetTitle("Demo PDF in AutoIt")
_SetSubject("Demo PDF in AutoIt, without any ActiveX or DLL...")
_SetKeywords("pdf, demo, AutoIt")
_OpenAfter(True);open after generation
_SetUnit($PDF_UNIT_CM)
_SetPaperSize("A4")
_SetZoomMode($PDF_ZOOM_CUSTOM,90)
_SetOrientation($PDF_ORIENTATION_PORTRAIT)
_SetLayoutMode($PDF_LAYOUT_CONTINOUS)

;initialize the pdf
_InitPDF(@ScriptDir & "\Example_mixed.pdf")

;=== load resources used in pdf ===
;images:
_LoadResImage("taietel", @ScriptDir & "\Images\bmp.bmp")
_LoadResImage("taietel1", @ScriptDir & "\Images\png.png")
_LoadResImage("taietel2", @ScriptDir & "\Images\gif.gif")
_LoadResImage("taietel3", @ScriptDir & "\Images\jpg.jpg")
_LoadResImage("taietel4", @ScriptDir & "\Images\ico.ico")
_LoadResImage("taietel5", @ScriptDir & "\Images\tif.tif")
;fonts:
;_LoadFontStandard("_Times", $PDF_FONT_STD_TIMES)
_LoadFontTT("_Arial", $PDF_FONT_ARIAL)
_LoadFontTT("_TimesT", $PDF_FONT_TIMES)
_LoadFontTT("_Calibri", $PDF_FONT_CALIBRI)
_LoadFontTT("_Garamond", $PDF_FONT_GARAMOND)
;_LoadFontTT("_Symbol", $PDF_FONT_SYMBOL)

;=== create objects that are used in multiple pages ===
;create a header on all pages, except the first:
_StartObject("Antet", $PDF_OBJECT_ALLPAGES);NOTFIRSTPAGE)
	;insert a image already loaded
	_InsertImage("taietel", 2, 25, 3, 3)
	;change the colour of the text that follows
	_SetColourFill(0x323232)
	;stretch it a bit, down to 90%
	_SetTextHorizontalScaling(90)
	;and begin writting some data
	_DrawText(5.2, 27.6, StringUpper("Et adipiscing nec nisi elementum natoque!"), "_Garamond", 14, $PDF_ALIGN_LEFT)
	_DrawText(5.2, 26.9, StringUpper("Dapibus scelerisque vel rhoncus porttitor!"), "_Garamond", 16, $PDF_ALIGN_LEFT)
	_SetTextHorizontalScaling(80)
	_DrawText(5.2, 26.2, "Dapibus scelerisque vel rhoncus porttitor!", "_TimesT", 12, $PDF_ALIGN_LEFT)
	_DrawText(5.2, 25.6, "Rhoncus a vut natoque pellentesque", "_TimesT", 12, $PDF_ALIGN_LEFT)
	_DrawText(5.2, 25, "taietel@yahoo.com" & "; " & "http://autoitscript.com/forum/topic/118827-create-pdf-from-your-application/", "_TimesT", 11, $PDF_ALIGN_LEFT)
	;get the scalling back to default value
	_SetTextHorizontalScaling(100)
	;and colour also
	_SetColourFill(0)
	;that's the end of our header!
_EndObject()

;start a page
_BeginPage()
	;put some graphics, text etc (see the rest)
	_InsertImage("taietel1", 2, 10, 7, 7)

	_SetColourFill(0xFF0000)
	_DrawText(3, 21, "Demo PDF Arial TT", "_Arial", 12, $PDF_ALIGN_LEFT, 0)

	_SetColourFill(0xFFFF00)
	_SetWordSpacing(50)
	_DrawText(3, 20, "Demo PDF Times TT", "_Times", 18, $PDF_ALIGN_LEFT, 0)
	_SetWordSpacing(0)

	_SetTextHorizontalScaling(200)
	_SetColourFill(0x0000FF)
	_DrawText(3, 19, "Demo PDF Times Standard", "Times", 12, $PDF_ALIGN_LEFT, 0)
	_SetTextHorizontalScaling(100)

	_SetColourFill(0xFF00FF)
	_DrawText(3, 18, "Demo PDF Courier TT", "_Garamond", 30, $PDF_ALIGN_LEFT, 0)
	_SetColourFill(0)

	_SetTextRenderingMode(2)
	_SetColourStroke(0x808080)
	_DrawText(5, 15, "Demo PDF rotated ", "_Arial", 22, $PDF_ALIGN_LEFT, 40)
	_SetColourStroke(0)
	_SetTextRenderingMode(0)

	_SetColourFill(0x111111)
	_DrawText(10, 15, "AutoIt center", "_TimesT", 22, $PDF_ALIGN_CENTER, 0)
	_DrawText(10, 14, "AutoIt left", "_TimesT", 22, $PDF_ALIGN_LEFT, 0)
	_DrawText(10, 13, "AutoIt right", "_TimesT", 22, $PDF_ALIGN_RIGHT, 0)
	_SetColourFill(0)
_EndPage()

;begin another page
_BeginPage(270)
	_Draw_Rectangle(3, 17, 5, 5, $PDF_STYLE_FILLED, 0, 0xee0000, 0.05)
	_Draw_Rectangle(9, 17, 5, 5, $PDF_STYLE_STROKED, 0, 0xeeee00, 0.05)
	_DrawCircle(13, 17, 2.5)
	_InsertImage("taietel2", 2, 7, 5, 7)
_EndPage()

_BeginPage()
	_Insert3DPie(4, 18, 2.5, 200, 270, 0x996600)
	_Insert3DPie(4, 18, 2.5, 90, 200, 0xeeee00)
	_Insert3DPie(4, 18, 2.5, 270, 300, 0x0000ff)
	_Insert3DPie(4, 18, 2.5, 300, 360, 0xee00ee)
	_Insert3DPie(4, 18, 2.5, 60, 90, 0x00ee00)
	_Insert3DPie(4, 18, 2.5, 0, 60, 0xff0000)
	_InsertImage("taietel5", 2, 10.5, 7, 5)
_EndPage()

_BeginPage(180)
	_SetTextRenderingMode(5)
	_InsertRenderedText(9, 23.5, "DEMO pdf", "_Times", 32, 100, $PDF_ALIGN_CENTER, 0xF00000, 0x202020)
	_SetTextRenderingMode(0)
	_InsertImage("taietel3", 2, 10.5, 7, 5)
	$sText = "Anna Montes ridiculus, in penatibus, in aliquam enim sagittis pellentesque? Mattis duis et ut nunc sagittis enim "
	$sText &= "tortor urna, eros? Scelerisque? Dapibus scelerisque vel rhoncus porttitor! Porttitor ridiculus. In adipiscing augue "
	$sText &= "vel pellentesque tortor porta hac tristique turpis placerat scelerisque elementum hac pulvinar mid dolor pellentesque "
	$sText &= "lundium mattis nec. Nec sed. Et adipiscing nec nisi elementum natoque! Turpis penatibus est dictumst magnis integer "
	$sText &= "scelerisque, sociis, risus scelerisque ultrices auctor porta, enim? Ac montes pellentesque cum enim augue penatibus "
	$sText &= "pulvinar? Vel mid, cum habitasse etiam urna? In? Et, natoque! Integer odio egestas! Rhoncus a vut natoque pellentesque "
	$sText &= "diam lundium augue in mus. Auctor, dictumst lacus turpis phasellus etiam, proin mauris. Natoque ultricies turpis nisi "
	$sText &= "platea parturient? Nunc nascetur est, adipiscing enim turpis Mihai."

	_Paragraph($sText, 2, 22, 8, "_Calibri", 10, 0)
	_Paragraph($sText, 11, 22, 8, "_Garamond", 13.5, 0)
	_Paragraph($sText, 2, 9.5, 15, "_TimesT", 10, 0)
_EndPage()

_BeginPage(0)
	_Insert3DCube(5, 15, 1, 4, 0xFF0000, 0.5, 0.1, 0.02)
	_Insert3DCube(7, 15, 1, 6, 0xFFFF00, 0.5, 0.1, 0.02)
	_Insert3DCube(9, 15, 1, 5, 0x0000FF, 0.5, 0.1, 0.02)
	_InsertImage("taietel4", 2, 10.5,1,1)
	_DrawLine(2, 2, 12, 6, $PDF_STYLE_STROKED, 10, 0.1, 0x996600, 0, 0)
_EndPage()

;and put another 10 blank pages at the end
For $i = 1 To 10
	_BeginPage()
	_EndPage()
Next

;then, finally, write the buffer to disk
_ClosePDFFile()

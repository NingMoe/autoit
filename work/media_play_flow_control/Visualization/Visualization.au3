;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;Copyright notice;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; Copyright © All rights reserved Andreas Karlsson 2008
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; This source is provided for educational purposes.
; If you wish to use parts of this source you need to credit me in your application.
; Contact: andreas.karlsson3@gmail.com
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; Notice that the above copyright does not apply to the following function:
; _GDIPlus_CreateLineBrushFromRect
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

#include <misc.au3>
#include <GDIPlus.au3>
#include <array.au3>
#include <file.au3>
#include <windowsconstants.au3>
Global $bass
Global Const $PI = 3.14159
Bass_Start()
$pluginhandle = Bass_LoadPlugin("bassflac.dll")

Opt("GUIOnEventMode", 1)
Global Const $width = 800
Global Const $height = 600
$hwnd = GUICreate("Visualization", $width, $height, -1, -1, -1, $WS_EX_ACCEPTFILES)
GUISetOnEvent(-3, "close")
GUISetState()


GUIRegisterMsg(563, "WM_DROPFILES_FUNC")

_GDIPlus_Startup()
$graphics = _GDIPlus_GraphicsCreateFromHWND($hwnd)
$bitmap = _GDIPlus_BitmapCreateFromGraphics($width, $height, $graphics)
$vizbitmap = _GDIPlus_BitmapCreateFromGraphics($width, $height, $graphics)
$backbuffer = _GDIPlus_ImageGetGraphicsContext($bitmap)
$vizbuffer = _GDIPlus_ImageGetGraphicsContext($vizbitmap)
$brush = _GDIPlus_BrushCreateSolid(0xFF22FF22)
$pen = _GDIPlus_PenCreate(0xFF22AA22, 2)

Global $blacktrans = _GDIPlus_BrushCreateSolid(0x10000000)

$family = _GDIPlus_FontFamilyCreate("Arial")
$font = _GDIPlus_FontCreate($family, 26)
$format = _GDIPlus_StringFormatCreate()
_GDIPlus_StringFormatSetAlign($format, 1)
$rect = _GDIPlus_RectFCreate(0, 150, $width, $height)
Global $RandomBrushes[128]
For $i = 0 To UBound($RandomBrushes) - 1
	$RandomBrushes[$i] = _GDIPlus_BrushCreateSolid(Random(0xAA000000, 0xAAFFFFFF, 1))
Next
Global $WhiteTransBrushes[256]
For $i = UBound($WhiteTransBrushes) - 1 To 0 Step -1
	$WhiteTransBrushes[$i] = _GDIPlus_BrushCreateSolid("0x" & Hex($i, 2) & "FFFFFF")
Next
Local $aFact[4] = [0.0, 0.01, 0.02, 1.0]
$lgbrush = _GDIPlus_CreateLineBrushFromRect(0, 00, $width, $height, $aFact, -1, 0xFFAA0000, 0xFF00AA00)

_AntiAlias($backbuffer, 4)
_AntiAlias($vizbuffer, 4)

Global $stream
Global $user32 = DllOpen("user32.dll")
Global $ID3 = "char id[3];char title[30];char artist[30];char album[30];char year[4];char comment[30];ubyte genre;"
Global $SongString = "Audio Visualization with GDI+" & @CRLF & "Drag 'n drop audio file to start playback"
Global $SongStringOpacity = 255
Global $active = 3
Global $groundangle = 0
Global $towerscount = 32
Global $roofs[$towerscount + 1][2]
Global $released = True


Global $seed = Random(0, 10000, 1)


$b = DllStructCreate("float[128]")

#Region Globals for _ScopeViz
Global $scrollbm1 = _GDIPlus_BitmapCreateFromGraphics($width, $height, $graphics)
Global $scrollbackbuffer1 = _GDIPlus_ImageGetGraphicsContext($scrollbm1)
_GDIPlus_GraphicsClear($scrollbackbuffer1)
Global $scrollbm2 = _GDIPlus_BitmapCreateFromGraphics($width, $height, $graphics)
Global $scrollbackbuffer2 = _GDIPlus_ImageGetGraphicsContext($scrollbm2)
_GDIPlus_GraphicsClear($scrollbackbuffer2)

Global $activeg = $scrollbackbuffer1
Global $totaloffset = 0
Global $step = 2
Global $_oldx=0
Global $_oldy=0
Global $changingbrush=_GDIPlus_BrushCreateSolid()

;~ _AntiAlias($scrollbackbuffer1,4)
;~ _AntiAlias($scrollbackbuffer2,4)
#EndRegion Globals for _ScopeViz

Global $momento=100
;~ Global $


Do
	Sleep(10)

	_GDIPlus_GraphicsClear($backbuffer)
	$call = DllCall($bass, "dword", "BASS_ChannelGetData", "dword", $stream, "ptr", DllStructGetPtr($b), "dword", 0x80000000)

	Switch $active
		Case 0
			_CircleViz($vizbuffer, $b, $brush)
		Case 1
			_SinViz($vizbuffer, $b, $pen)
		Case 2
			_TowerViz($vizbuffer, $b, $lgbrush, $pen)
		Case 3
			_TriangleViz($vizbuffer, $b, $pen)
		Case 4
			_BubbleSleep($vizbuffer, $b)
		Case 5
			_SpeakerViz($vizbuffer, $b, $pen)
		Case 6
			_3DTowerViz($vizbuffer, $b, $lgbrush)
		Case 7
			_ScopeViz($vizbuffer, $b)
		Case 8
			_TestViz($vizbuffer,$b,$pen)

		Case Else
			$active = 0
	EndSwitch
	_GDIPlus_GraphicsDrawImageRect($backbuffer, $vizbitmap, 0, 0, $width, $height)

	If _IsPressed("01", $user32) And WinActive($hwnd) Then
		If $released Then
			For $i = 0 To UBound($roofs) - 1
				$roofs[$i][0] = $height / 2
			Next

			$active += 1
			$released = False
		EndIf
	Else
		$released = True
	EndIf

	If $SongStringOpacity > 0 Then
		_GDIPlus_GraphicsDrawStringEx($backbuffer, $SongString, $font, $rect, $format, $WhiteTransBrushes[$SongStringOpacity])
		$SongStringOpacity -= 1
	EndIf





	_GDIPlus_GraphicsDrawImageRect($graphics, $bitmap, 0, 0, $width, $height)


Until False

Func _ScopeViz($surface, $fftstruct)
	If $totaloffset = 0 Then _GDIPlus_GraphicsClear($surface)
	$Sum = 0
	For $i = 1 To 128
		$Sum += DllStructGetData($fftstruct, 1, $i)
	Next

	$totaloffset += $step

	$activex1 = Mod($totaloffset, $width * 2)
	$activex2 = Mod($totaloffset - $width, $width * 2)


	If $activex2 = $width Then
		_GDIPlus_GraphicsClear($scrollbackbuffer1)
		$activeg = $scrollbackbuffer1
	ElseIf $activex1 = $width Then
		$activeg = $scrollbackbuffer2
		_GDIPlus_GraphicsClear($scrollbackbuffer2)
	EndIf

	$xpaint = Mod($totaloffset, $width)

	$r = $Sum * $height / 4
	
	$tred=Hex((Sin($totaloffset/1500)+1)/2*255,2)
	$tgreen=Hex((Sin($totaloffset/1500*2)+1)/2*255,2)
	$tblue=Hex((Sin($totaloffset/1500*3)+1)/2*255,2)
	
	$tcolor="0xFF"&$tred&$tgreen&$tblue
	_GDIPlus_BrushSetSolidColor_FromBeta($changingbrush,$tcolor)

	_GDIPlus_GraphicsFillRect($activeg, $xpaint, $height / 2 - $r / 2, $step, $r, $changingbrush)
	_GDIPlus_GraphicsDrawImageRect($surface, $scrollbm1, $width - $activex1, 0, $width, $height)
	_GDIPlus_GraphicsDrawImageRect($surface, $scrollbm2, $width - $activex2, 0, $width, $height)

EndFunc   ;==>_ScopeViz



Func close()
	Bass_FreeStream($stream)
	Bass_UnloadPlugin($pluginhandle)
	_GDIPlus_BrushDispose($lgbrush)
	For $i = 0 To UBound($WhiteTransBrushes) - 1
		_GDIPlus_BrushDispose($WhiteTransBrushes[$i])
	Next
	For $i = 0 To UBound($RandomBrushes) - 1
		_GDIPlus_BrushDispose($RandomBrushes[$i])
	Next
	_GDIPlus_StringFormatDispose($format)
	_GDIPlus_FontDispose($font)
	_GDIPlus_FontFamilyDispose($family)
	_GDIPlus_BrushDispose($blacktrans)
	_GDIPlus_PenDispose($pen)
	_GDIPlus_BrushDispose($brush)
	_GDIPlus_GraphicsDispose($vizbuffer)
	_GDIPlus_GraphicsDispose($backbuffer)
	_GDIPlus_BitmapDispose($vizbitmap)
	_GDIPlus_BitmapDispose($bitmap)
	_GDIPlus_GraphicsDispose($graphics)
	_GDIPlus_Shutdown()

	Exit
EndFunc   ;==>close




Func _3DTowerViz($surface, $fftstruct, $brush)
	_GDIPlus_GraphicsClear($surface, 0xFF000000)
	Local $tcount = 10
	Local $towerw = $width / $tcount - 5
	For $i = 0 To $tcount Step 1
		$fft = DllStructGetData($fftstruct, 1, ($i / $tcount) * 128 + 1)
		$h = (Sqrt($fft) ^ 0.75) * $height
		$x = 5 + $i * ($towerw + 5)
		_Fill3DStaple($surface, $x, $height - $h, $towerw - $towerw / 4, $height + $h, $towerw / 4, $brush, $brush, $brush)
	Next
EndFunc   ;==>_3DTowerViz


Func _BubbleSleep($surface, $fftstruct)
	SRandom($seed)
	_GDIPlus_GraphicsClear($surface, 0xFF000000)

	For $i = 1 To 110

		$fft = DllStructGetData($fftstruct, 1, $i)

		$w = Sqrt($fft) * 200 + 10
		$h = Sqrt($fft) * 200 + 10
		$x = Random(0, $width, 1) - $w / 2

		$y = Random(0, $height, 1) - $h / 2

		_GDIPlus_GraphicsFillEllipse($surface, $x, $y, $w, $h, $RandomBrushes[$i - 1])

	Next



EndFunc   ;==>_BubbleSleep


Func _CircleViz($surface, $fftstruct, $brush)
	_GDIPlus_GraphicsClear($surface, 0xFF000000)
	SRandom($seed)
	Local $dots = 100
	For $i = 1 To $dots
		$fft = DllStructGetData($fftstruct, 1, Random(1, 100, 1));$randvalues[$i-1])
		If Mod($i, 2) = 0 Then
			$x = (Cos($PI * ($i / $dots)) * 30 * Sqrt(Sqrt($fft * 100000))) + $width / 2
			$y = (Sin($PI * ($i / $dots)) * 30 * Sqrt(Sqrt($fft * 100000))) + $height / 2
		Else
			$x = (Cos(-1 * $PI * ($i / $dots)) * 30 * Sqrt(Sqrt($fft * 100000))) + $width / 2
			$y = (Sin(-1 * $PI * ($i / $dots)) * 30 * Sqrt(Sqrt($fft * 100000))) + $height / 2
		EndIf

		_GDIPlus_GraphicsFillEllipse($surface, $x, $y, 2, 2, $brush)

	Next


EndFunc   ;==>_CircleViz



Func _SpeakerViz($surface, $fftstruct, $pen)
;~ 	_GDIPlus_GraphicsClear($surface, 0xFF000000)
	_GDIPlus_GraphicsFillRect($surface, 0, 0, $width, $height, $blacktrans)
	$Sum = 0
	For $i = 1 To 128
		$Sum += DllStructGetData($fftstruct, 1, $i)
	Next
	$size = $Sum * 105 + 10
	_GDIPlus_GraphicsDrawArc($surface, 10, $height / 2 - $size / 2, $size, $size, 270, 180, $pen)
	_GDIPlus_GraphicsDrawArc($surface, $width - 10 - $size, $height / 2 - $size / 2, $size, $size, 90, 180, $pen)

EndFunc   ;==>_SpeakerViz


Func _TriangleViz($surface, $fftstruct, $pen)
	$Sum = 0
;~ 	_GDIPlus_GraphicsFillRect($surface,0,0,$width,$height,$blacktrans)
	_GDIPlus_GraphicsFillRect($surface, 0, 0, $width, $height, $blacktrans)
;~ 	_GDIPlus_GraphicsClear($surface)


	For $i = 1 To 128
		$Sum += DllStructGetData($fftstruct, 1, $i)
	Next
	$groundangle += $PI / 50
	$size = 50 + $Sum * 100
	$x1 = Cos($groundangle + $PI / 2) * $size + $width / 2
	$y1 = Sin($groundangle + $PI / 2) * $size + $height / 2
	$x2 = Cos($groundangle + $PI / 2 + (2 * $PI) / 3) * $size + $width / 2
	$y2 = Sin($groundangle + $PI / 2 + (2 * $PI) / 3) * $size + $height / 2
	$x3 = Cos($groundangle + $PI / 2 + ((2 * $PI) / 3) * 2) * $size + $width / 2
	$y3 = Sin($groundangle + $PI / 2 + ((2 * $PI) / 3) * 2) * $size + $height / 2
	_GDIPlus_GraphicsDrawLine($surface, $x1, $y1, $x2, $y2, $pen)
	_GDIPlus_GraphicsDrawLine($surface, $x2, $y2, $x3, $y3, $pen)
	_GDIPlus_GraphicsDrawLine($surface, $x3, $y3, $x1, $y1, $pen)



EndFunc   ;==>_TriangleViz

Func _TestViz($surface, $fftstruct, $pen)
;~ 	_GDIPlus_GraphicsFillRect($surface, 0, 0, $width, $height, $blacktrans)
	_GDIPlus_GraphicsClear($surface)	
	$Sum = 0



	For $i = 1 To 128
		$Sum += DllStructGetData($fftstruct, 1, $i)
	Next
	$momento+=$Sum*50
	
	
	$groundangle+=( $Sum/100+$momento/1000+0.01)/2
	$momento/=2
	
	$x1=Cos($groundangle)*200+$width/2
	$y1=Sin($groundangle)*200+$height/2
	$x2=Cos($groundangle+$PI)*200+$width/2
	$y2=Sin($groundangle+$PI)*200+$height/2
	
	
	$x3=Cos($groundangle+$PI/2)*200+$width/2
	$y3=Sin($groundangle+$PI/2)*200+$height/2
	$x4=Cos($groundangle+$PI+$PI/2)*200+$width/2
	$y4=Sin($groundangle+$PI+$PI/2)*200+$height/2
	
	_GDIPlus_GraphicsDrawLine($surface,$x1,$y1,$x2,$y2,$pen)
	_GDIPlus_GraphicsDrawLine($surface,$x1,$y1,$x3,$y3,$pen)
;~ 	_GDIPlus_GraphicsDrawLine($surface,$x3,$y3,$x2,$y2,$pen)
	_GDIPlus_GraphicsDrawLine($surface,$x2,$y2,$x4,$y4,$pen)
;~ 	_GDIPlus_GraphicsDrawLine($surface,$x1,$y1,$x4,$y4,$pen)

EndFunc   ;==>_TowerViz


Func _TowerViz($surface, $fftstruct, $brush, $pen)
	_GDIPlus_GraphicsClear($surface, 0xFF000000)

	Local $towerw = $width / $towerscount



	For $i = 0 To $towerscount Step 1
		$fft = DllStructGetData($fftstruct, 1, ($i / $towerscount) * 128 + 1)

		$h = (Sqrt($fft) ^ 0.75) * $height
		;$h = (Log($fft*100)) * $height
		$x = 1 + $i * ($towerw)
		If $roofs[$i][0] < $h Then
			$roofs[$i][0] = $h
			$roofs[$i][1] = 13
		Else
			$roofs[$i][1] -= 1

			If $roofs[$i][1] < 0 Then $roofs[$i][0] -= 5
		EndIf

		_GDIPlus_GraphicsFillRect($surface, $x, $height - $h, $towerw, $height + $h, $brush)
		_GDIPlus_GraphicsDrawLine($surface, $x, $height - $roofs[$i][0], $x + $towerw, $height - $roofs[$i][0], $pen)
	Next
;~ 	Sleep(10)

EndFunc   ;==>_TowerViz

Func _SinViz($surface, $fftstruct, $pen)
	Local $oldx = 0, $oldy = $height / 2
	_GDIPlus_GraphicsClear($surface, 0xFF000000)
	For $i = 1 To 128

	Next

	For $i = 0 To $width Step 5
		$fft = DllStructGetData($fftstruct, 1, $i / 6)
		$y = $height / 2 + Sin($i) * Sqrt($fft) * 500
		_GDIPlus_GraphicsDrawLine($surface, $oldx, $oldy, $i, $y, $pen)
		$oldx = $i
		$oldy = $y
	Next


EndFunc   ;==>_SinViz



Func WM_DROPFILES_FUNC($hwnd, $msgID, $wParam, $lParam)
	Local $nSize, $pFileName
	Local $nAmt = DllCall("shell32.dll", "int", "DragQueryFile", "hwnd", $wParam, "int", 0xFFFFFFFF, "ptr", 0, "int", 255)
	For $i = 0 To $nAmt[0] - 1
		$nSize = DllCall("shell32.dll", "int", "DragQueryFile", "hwnd", $wParam, "int", $i, "ptr", 0, "int", 0)
		$nSize = $nSize[0] + 1
		$pFileName = DllStructCreate("char[" & $nSize & "]")
		DllCall("shell32.dll", "int", "DragQueryFile", "hwnd", $wParam, "int", $i, "ptr", DllStructGetPtr($pFileName), "int", $nSize)
		Bass_FreeStream($stream)
		$stream = Bass_StreamCreateFile(DllStructGetData($pFileName, 1))


		If $stream = 0 Then Return
		If StringRight(DllStructGetData($pFileName, 1), 4) = "flac" Or StringRight(DllStructGetData($pFileName, 1), 4) = ".ogg" Then
			$ptr = Bass_ChannelGetTags($stream, 2)
			$temp = _GetID3StructFromOGGComment($ptr)
			$SongString = DllStructGetData($temp, "Title")
			If StringLen(DllStructGetData($temp, "Artist")) > 1 Then $SongString &= " - " & DllStructGetData($temp, "Artist")
			$SongStringOpacity = 255
		Else
			$ptr = Bass_ChannelGetTags($stream, 0)
			$temp = DllStructCreate($ID3, $ptr)
			$SongString = DllStructGetData($temp, "Title")
			If StringLen(DllStructGetData($temp, "Artist")) > 1 Then $SongString &= " - " & DllStructGetData($temp, "Artist")
			$SongStringOpacity = 255
		EndIf


		Bass_ChannelPlay($stream)

	Next
EndFunc   ;==>WM_DROPFILES_FUNC

Func _GDIPlus_BrushSetSolidColor_FromBeta($hBrush, $iARGB = 0xFF000000)
    Local $aResult
    $aResult = DllCall($ghGDIPDll, "int", "GdipSetSolidFillColor", "hwnd", $hBrush, "int", $iARGB)
    If @error Then Return SetError(@error, @extended, 0)
    Return SetError($aResult[0], 0, $aResult[0] = 0)
EndFunc ;==>_GDIPlus_BrushSetSolidColor

Func Bass_UnloadPlugin($plugin)
	DllCall($bass, "int", "BASS_PluginFree", "dword", $plugin)
EndFunc   ;==>Bass_UnloadPlugin

Func _AntiAlias($hGraphics, $iMode)
	Local $aResult

	$aResult = DllCall($ghGDIPDll, "int", "GdipSetSmoothingMode", "hwnd", $hGraphics, "int", $iMode)
	If @error Then Return SetError(@error, @extended, False)
	Return SetError($aResult[0], 0, $aResult[0] = 0)
EndFunc   ;==>_AntiAlias

Func Bass_LoadPlugin($fname)
	$str = DllStructCreate("char[255];")
	DllStructSetData($str, 1, $fname)
	$call = DllCall($bass, "dword", "BASS_PluginLoad", "ptr", DllStructGetPtr($str), "dword", 0)
	Return $call[0]
EndFunc   ;==>Bass_LoadPlugin

Func Bass_StreamCreateFile($fname)
	$str = DllStructCreate("char[255];")
	DllStructSetData($str, 1, $fname)

	$call = DllCall($bass, "int", "BASS_StreamCreateFile", "int", 0, "ptr", DllStructGetPtr($str), "uint64", 0, "uint64", 0, "dword", 0);
	Return $call[0]
EndFunc   ;==>Bass_StreamCreateFile


Func Bass_Start()
	$bass = DllOpen("bass.dll")
	$call = DllCall($bass, "int", "BASS_Init", "int", -1, "dword", 44100, "dword", 0, "hwnd", 0, "ptr", 0)
EndFunc   ;==>Bass_Start


Func Bass_ChannelPlay($stream)
	$call = DllCall($bass, "int", "BASS_ChannelPlay", "dword", $stream, "int", 1);
EndFunc   ;==>Bass_ChannelPlay

Func Bass_FreeStream($stream)
	DllCall($bass, "int", "BASS_StreamFree", "dword", $stream)
EndFunc   ;==>Bass_FreeStream
Func Bass_ChannelGetTags($stream, $flag)
	$call = DllCall($bass, "ptr", "BASS_ChannelGetTags", "dword", $stream, "dword", $flag)
	Return $call[0]
EndFunc   ;==>Bass_ChannelGetTags


Func _GetID3StructFromOGGComment($ptr)
	$tags = DllStructCreate($ID3)
	Do
		$s = DllStructCreate("char[255];", $ptr)
		$string = DllStructGetData($s, 1)
		If StringLeft($string, 1) = Chr(0) Then ExitLoop
;~ 		MsgBox(0, "", $string)


		Switch StringLeft($string, StringInStr($string, "=") - 1)
			Case "title"
				DllStructSetData($tags, "title", StringTrimLeft($string, StringInStr($string, "=")))
			Case "artist"
				DllStructSetData($tags, "artist", StringTrimLeft($string, StringInStr($string, "=")))
			Case "album"
				DllStructSetData($tags, "album", StringTrimLeft($string, StringInStr($string, "=")))
			Case "date"
				DllStructSetData($tags, "year", StringTrimLeft($string, StringInStr($string, "=")))
			Case "genre"
				DllStructSetData($tags, "genre", StringTrimLeft($string, StringInStr($string, "=")))
			Case "comment"
				DllStructSetData($tags, "comment", StringTrimLeft($string, StringInStr($string, "=")))
		EndSwitch
		$ptr += StringLen($string) + 1
	Until False

	Return $tags
EndFunc   ;==>_GetID3StructFromOGGComment




;==== GDIPlus_CreateLineBrushFromRect ===
;Description - Creates a LinearGradientBrush object from a set of boundary points and boundary colors.
; $aFactors - If non-array, default array will be used.
;           Pointer to an array of real numbers that specify blend factors. Each number in the array
;           specifies a percentage of the ending color and should be in the range from 0.0 through 1.0.
;$aPositions - If non-array, default array will be used.
;            Pointer to an array of real numbers that specify blend factors' positions. Each number in the array
;            indicates a percentage of the distance between the starting boundary and the ending boundary
;            and is in the range from 0.0 through 1.0, where 0.0 indicates the starting boundary of the
;            gradient and 1.0 indicates the ending boundary. There must be at least two positions
;            specified: the first position, which is always 0.0, and the last position, which is always
;            1.0. Otherwise, the behavior is undefined. A blend position between 0.0 and 1.0 indicates a
;            line, parallel to the boundary lines, that is a certain fraction of the distance from the
;            starting boundary to the ending boundary. For example, a blend position of 0.7 indicates
;            the line that is 70 percent of the distance from the starting boundary to the ending boundary.
;            The color is constant on lines that are parallel to the boundary lines.
; $iArgb1    - First Top color in 0xAARRGGBB format
; $iArgb2    - Second color in 0xAARRGGBB format
; $LinearGradientMode -  LinearGradientModeHorizontal       = 0x00000000,
;                        LinearGradientModeVertical         = 0x00000001,
;                        LinearGradientModeForwardDiagonal  = 0x00000002,
;                        LinearGradientModeBackwardDiagonal = 0x00000003
; $WrapMode  - WrapModeTile       = 0,
;              WrapModeTileFlipX  = 1,
;              WrapModeTileFlipY  = 2,
;              WrapModeTileFlipXY = 3,
;              WrapModeClamp      = 4
; GdipCreateLineBrushFromRect(GDIPCONST GpRectF* rect, ARGB color1, ARGB color2,
;             LinearGradientMode mode, GpWrapMode wrapMode, GpLineGradient **lineGradient)
; Reference:  http://msdn.microsoft.com/en-us/library/ms534043(VS.85).aspx
;
Func _GDIPlus_CreateLineBrushFromRect($iX, $iY, $iWidth, $iHeight, $aFactors, $aPositions, _
		$iArgb1 = 0xFF0000FF, $iArgb2 = 0xFFFF0000, $LinearGradientMode = 0x00000001, $WrapMode = 0)

	Local $tRect, $pRect, $aRet, $tFactors, $pFactors, $tPositions, $pPositions, $iCount

	If $iArgb1 = -1 Then $iArgb1 = 0xFF0000FF
	If $iArgb2 = -1 Then $iArgb2 = 0xFFFF0000
	If $LinearGradientMode = -1 Then $LinearGradientMode = 0x00000001
	If $WrapMode = -1 Then $WrapMode = 1

	$tRect = DllStructCreate("float X;float Y;float Width;float Height")
	$pRect = DllStructGetPtr($tRect)
	DllStructSetData($tRect, "X", $iX)
	DllStructSetData($tRect, "Y", $iY)
	DllStructSetData($tRect, "Width", $iWidth)
	DllStructSetData($tRect, "Height", $iHeight)

	;Note: Withn _GDIPlus_Startup(), $ghGDIPDll is defined
	$aRet = DllCall($ghGDIPDll, "int", "GdipCreateLineBrushFromRect", "ptr", $pRect, "int", $iArgb1, _
			"int", $iArgb2, "int", $LinearGradientMode, "int", $WrapMode, "int*", 0)

	If IsArray($aFactors) = 0 Then Dim $aFactors[4] = [0.0, 0.4, 0.6, 1.0]
	If IsArray($aPositions) = 0 Then Dim $aPositions[4] = [0.0, 0.3, 0.7, 1.0]

	$iCount = UBound($aPositions)
	$tFactors = DllStructCreate("float[" & $iCount & "]")
	$pFactors = DllStructGetPtr($tFactors)
	For $iI = 0 To $iCount - 1
		DllStructSetData($tFactors, 1, $aFactors[$iI], $iI + 1)
	Next
	$tPositions = DllStructCreate("float[" & $iCount & "]")
	$pPositions = DllStructGetPtr($tPositions)
	For $iI = 0 To $iCount - 1
		DllStructSetData($tPositions, 1, $aPositions[$iI], $iI + 1)
	Next

	$hStatus = DllCall($ghGDIPDll, "int", "GdipSetLineBlend", "hwnd", $aRet[6], _
			"ptr", $pFactors, "ptr", $pPositions, "int", $iCount)
	Return $aRet[6] ; Handle of Line Brush
EndFunc   ;==>_GDIPlus_CreateLineBrushFromRect



Func _Fill3DStaple($surface, $x, $y, $w, $h, $3D, $b1, $b2, $b3)
	Local $va1[5][2]
	Local $va2[5][2]
	Local $va3[5][2]
	$va1[0][0] = 4
	$va1[1][0] = $x
	$va1[1][1] = $y
	$va1[2][0] = $x + $w
	$va1[2][1] = $y
	$va1[3][0] = $x + $w + $3D
	$va1[3][1] = $y + $3D
	$va1[4][0] = $x + $3D
	$va1[4][1] = $y + $3D
	$va2[0][0] = 4
	$va2[1][0] = $x + $3D
	$va2[1][1] = $y + $3D
	$va2[2][0] = $x + $w + $3D
	$va2[2][1] = $y + $3D
	$va2[3][0] = $x + $w + $3D
	$va2[3][1] = $y + $h
	$va2[4][0] = $x + $3D
	$va2[4][1] = $y + $h
	$va3[0][0] = 4
	$va3[1][0] = $x + $3D
	$va3[1][1] = $y + $h
	$va3[2][0] = $x
	$va3[2][1] = $y + $h - $3D
	$va3[3][0] = $x
	$va3[3][1] = $y
	$va3[4][0] = $x + $3D
	$va3[4][1] = $y + $3D
	_GDIPlus_GraphicsFillPolygon_FromBeta($surface, $va1, $b1)
	_GDIPlus_GraphicsFillPolygon_FromBeta($surface, $va2, $b2)
	_GDIPlus_GraphicsFillPolygon_FromBeta($surface, $va3, $b3)
EndFunc   ;==>_Fill3DStaple


; #FUNCTION# ===================================================================================
; Name...........: _GDIPlus_GraphicsFillPolygon
; Description ...: Fill a polygon
; Syntax.........: _GDIPlus_GraphicsFillPolygon($hGraphics, $aPoints[, $hBrush = 0])
; Parameters ....: $hGraphics   - Handle to a Graphics object
;                  $aPoints     - Array that specify the vertices of the polygon:
;                  |[0][0] - Number of vertices
;                  |[1][0] - Vertice 1 X position
;                  |[1][1] - Vertice 1 Y position
;                  |[2][0] - Vertice 2 X position
;                  |[2][1] - Vertice 2 Y position
;                  |[n][0] - Vertice n X position
;                  |[n][1] - Vertice n Y position
;                  $hBrush      - Handle to a brush object that is used to fill the polygon.
;                               - If $hBrush is 0, a solid black brush is used.
; Return values .: Success      - True
;                  Failure      - False
; Author ........:
; Modified.......: smashly
; Remarks .......:
; Related .......:
; Link ..........; @@MsdnLink@@ GdipFillPolygonI
; Example .......; Yes
; ===============================================================================================
Func _GDIPlus_GraphicsFillPolygon_FromBeta($hGraphics, $aPoints, $hBrush = 0)
	Local $iI, $iCount, $pPoints, $tPoints, $aResult, $tmpError, $tmpExError

	$iCount = $aPoints[0][0]
	$tPoints = DllStructCreate("int[" & $iCount * 2 & "]")
	$pPoints = DllStructGetPtr($tPoints)
	For $iI = 1 To $iCount
		DllStructSetData($tPoints, 1, $aPoints[$iI][0], (($iI - 1) * 2) + 1)
		DllStructSetData($tPoints, 1, $aPoints[$iI][1], (($iI - 1) * 2) + 2)
	Next

	_GDIPlus_BrushDefCreate($hBrush)
	$aResult = DllCall($ghGDIPDll, "int", "GdipFillPolygonI", "hWnd", $hGraphics, "hWnd", $hBrush, _
			"ptr", $pPoints, "int", $iCount, "int", "FillModeAlternate")
	$tmpError = @error
	$tmpExError = @extended
	_GDIPlus_BrushDefDispose()
	If $tmpError Then Return SetError($tmpError, $tmpExError, False)
	Return SetError($aResult[0], 0, $aResult[0] = 0)
EndFunc   ;==>_GDIPlus_GraphicsFillPolygon_FromBeta
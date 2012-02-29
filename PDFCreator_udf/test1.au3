#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

;; Testpage2PDF script
; Part of $PDFCreator
; License: GPL
; Homepage: http://www.sf.net/projects/PDFCreator
; Version: 1.0.0.0
; Date: September, 1. 2005
; Author: Frank Heindorfer
; Comments: Save the test page as pdf-file using the com interface of $PDFCreator.

; Translated by ptrex

AutoItSetOption("MustDeclareVars", 1)

Const $maxTime = 10    ; in seconds
Const $sleepTime = 250 ; in milliseconds

Dim $fso, $WshShell, $PDFCreator, $DefaultPrinter, $ReadyState, $c, _
    $Scriptname, $Scriptbasename

$fso = ObjCreate("Scripting.FileSystemObject")

$Scriptbasename = $fso.GetBaseName(@ScriptFullPath)


$WshShell = ObjCreate("WScript.Shell")

$PDFCreator = ObjCreate("PDFCreator.clsPDFCreator")
$PDFCreator.cStart ("/NoProcessingAtStartup")

$ReadyState = 0
With $PDFCreator
 .cOption("UseAutosave") = 1
 .cOption("UseAutosaveDirectory") = 1
 .cOption("AutosaveDirectory") = $fso.GetParentFolderName(@ScriptFullPath)
 .cOption("AutosaveFilename") = "Testpage - PDFCreator"
 .cOption("AutosaveFormat") = 0                        ; 0=PDF, 1=PNG, 2=JPG, 3=BMP, 4=PCX, 5=TIFF, 6=PS, 7= EPS, 8=ASCII
 $DefaultPrinter = .cDefaultprinter
 .cDefaultprinter = "PDFCreator"
 .cClearcache()
 .cPrintPDFCreatorTestpage()
 .cPrinterStop = 0
EndWith

$c = 0

Do 
 $c = $c + 1
    Sleep ($sleepTime)
Until ($ReadyState = 0) and ($c < ($maxTime * 1000 / $sleepTime))

With $PDFCreator
 .cDefaultprinter = $DefaultPrinter
    Sleep( 200)
 .cClose()
EndWith

If $ReadyState = 0 then
 Consolewrite ("Creating test page as pdf." & @CRLF & @CRLF & "An error is occured: Time is up!"& @CR)
 ProcessClose("PDFCreator.exe")
EndIf

;--- $PDFCreator events ---

Func PDFCreator_eReady()
 $ReadyState = 1
EndFunc

Func PDFCreator_eError()
 Consolewrite ("An error is occured!" & @CRLF & @CRLF & _
  "Error [" & $PDFCreator.cErrorDetail("Number") & "]: " & $PDFCreator.cErrorDetail("Description")& @CR)
;VA  Wscript.Quit
EndFunc
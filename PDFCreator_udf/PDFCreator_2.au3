#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here


; Convert2PDF script
; Part of $PDFCreator
; License: GPL
; Homepage: http://www.sf.net/projects/pdfcreator
; Version: 1.0.0.0
; Date: September, 1. 2005
; Author: Frank Heindorfer
; Comments: This script convert a printable file in a pdf-file using the com interface of $PDFCreator.

; Translated by ptrex

AutoItSetOption("MustDeclareVars", 1)

Const $maxTime = 30    ; in seconds
Const $sleepTime = 250 ; in milliseconds

Dim $objArgs, $ifname, $fso, $PDFCreator, $DefaultPrinter, $ReadyState, _
 $i, $c, $AppTitle, $Scriptname, $ScriptBasename, $File


$fso = ObjCreate("Scripting.FileSystemObject")

$Scriptname = $fso.GetFileName(@ScriptFullPath)
$ScriptBasename = $fso.GetFileName(@ScriptFullPath)

$AppTitle = "PDFCreator - " & $ScriptBasename

;$File = InputBox("FileName","Fill in the Path and filename","C:\_\Apps\AutoIT3\COM\PDFCreator\VBScripts\GUI.vbs")
$File="r:\2.txt"
$PDFCreator = ObjCreate("PDFCreator.clsPDFCreator")
$PDFCreator.cStart ("/NoProcessingAtStartup")

With $PDFCreator
  .cOption("UseAutosave") = 1
 .cOption("UseAutosaveDirectory") = 1
 .cOption("AutosaveDirectory") = $fso.GetParentFolderName(@ScriptFullPath)
 .cOption("AutosaveFilename") = "PDFCreator_abc"
 .cOption("AutosaveFormat") = 1                     ; 0=PDF, 1=PNG, 2=JPG, 3=BMP, 4=PCX, 5=TIFF, 6=PS, 7= EPS, 8=ASCII
 $DefaultPrinter = .cDefaultprinter
 .cDefaultprinter = "PDFCreator"
 .cClearcache()
EndWith

; For $i = 0 to $objArgs.Count - 1
 With $PDFCreator
  $ifname = $File ;"C:\Tmp\Test.xls" ;$objArgs($i)
  If Not $fso.FileExists($ifname) Then
   MsgBox (0,"Error","Can't find the file: " & $ifname & @CR & $AppTitle)
    Exit
  EndIf
  If Not .cIsPrintable(String($ifname)) Then
   ConsoleWrite("Converting: " & $ifname & @CRLF & @CRLF & _
    "An error is occured: File is not printable!" & @CRLF &  $AppTitle & @CR)
  EndIf

  $ReadyState = 0
  .cOption("AutosaveDirectory") = $fso.GetParentFolderName($ifname)
  .cOption("AutosaveFilename") = $fso.GetBaseName($ifname)
  .cPrintfile (String($ifname))
  .cPrinterStop = 0

  $c = 0
  Do 
   $c = $c + 1
   Sleep ($sleepTime)
  Until ($ReadyState = 0) and ($c < ($maxTime * 1000 / $sleepTime))
  
  If $ReadyState = 0 then
   ConsoleWrite("Converting: " & $ifname & @CRLF & @CRLF & _
    "An error is occured: File is not printable!" & @CRLF &  $AppTitle & @CR)
    Exit
  EndIf
 EndWith
;Next

With $PDFCreator
 .cDefaultprinter = $DefaultPrinter
 .cClearcache()
    Sleep (200)
 .cClose()
EndWith

ProcessClose("PDFCreator.exe")

;--- $PDFCreator events ---

Func PDFCreator_eReady()
 $ReadyState = 1
EndFunc

Func PDFCreator_eError()
 MsgBox(0, "An error is occured!" , "Error [" & $PDFCreator.cErrorDetail("Number") & "]: " & $PDFCreator.cErrorDetail("Description")& @CR)
EndFunc
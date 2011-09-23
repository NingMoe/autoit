

AutoItSetOption("MustDeclareVars", 1)

Const $maxTime = 30 ; in seconds
Const $sleepTime = 250 ; in milliseconds

Dim $objArgs, $ifname, $fso, $PDFCreator, $DefaultPrinter, $ReadyState, _
$i, $c, $AppTitle, $Scriptname, $ScriptBasename,$PDFCObject


$fso = ObjCreate("Scripting.FileSystemObject")

$Scriptname = $fso.GetFileName(@ScriptFullPath)
$ScriptBasename = $fso.GetFileName(@ScriptFullPath)

$AppTitle = "PDFCreator - " & $ScriptBasename

$PDFCreator = ObjCreate("PDFCreator.clsPDFCreator")
$PDFCreator.cStart ("/NoProcessingAtStartup")
$PDFCObject=ObjEvent($PDFCreator,"PDFCreator_") ; Start receiving Events.


With $PDFCreator
.cOption("UseAutosave") = 1
.cOption("UseAutosaveDirectory") = 1
.cOption("AutosaveFormat") = 2 ; 0=PDF, 1=PNG, 2=JPG, 3=BMP, 4=PCX, 5=TIFF, 6=PS, 7= EPS, 8=ASCII
$DefaultPrinter = .cDefaultprinter
.cDefaultprinter = "PDFCreator"
.cClearcache()
EndWith

; For $i = 0 to $objArgs.Count - 1
With $PDFCreator
$ifname = "r:\2.txt" ;$objArgs($i)
If Not $fso.FileExists($ifname) Then
MsgBox (0,"Error","Can't find the file: " & $ifname & @CR & $AppTitle)
Exit
EndIf
If Not .cIsPrintable($ifname) Then
ConsoleWrite("Converting: " & $ifname & @CRLF & @CRLF & _
"An error is occured: File is not printable!" & @CRLF & $AppTitle & @CR)
EndIf

$ReadyState = 0
.cOption("AutosaveDirectory") = $fso.GetParentFolderName($ifname)
.cOption("AutosaveFilename") = $fso.GetBaseName($ifname)
.cPrintfile($ifname)
.cPrinterStop = 0

$c = 0
Do
$c = $c + 1
Sleep ($sleepTime)
Until ($ReadyState <> 0) or ($c >= ($maxTime * 1000 / $sleepTime))

If $ReadyState = 0 then
ConsoleWrite("Converting: " & $ifname & @CRLF & @CRLF & _
"An error is occured: File is not printable!" & @CRLF & $AppTitle & @CR)
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

;--- $PDFCreator events ---

Func PDFCreator_eReady()
$ReadyState = 1
EndFunc

Func PDFCreator_eError()
MsgBox(0, "An error is occured!" , "Error [" & $PDFCreator.cErrorDetail("Number") & "]: " & $PDFCreator.cErrorDetail("Description")& @CR)
EndFunc
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

; Set a COM Error handler -- only one can be active at a time (see helpfile)
$oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")

$sNumberToDial = "5551212"
Dial($sNumberToDial)

Exit

Func Dial($pNum, $time2wait = 5000)

	dim $FromModem = ""
	$DialString = "ATDT" & $pNum & ";" & @CR

	$com = ObjCreate ("NETCommOCX.NETComm")

	With $com
		.CommPort = 3
		.PortOpen = True
		.Settings = "9600,N,8,1"
		.InBufferCount = 0
		.Output = $DialString
	EndWith

	$begin = TimerInit()
	While 1
		If $com.InBufferCount Then
			$FromModem = $FromModem & $com.InputData;* Check for "OK".
			If StringInStr($FromModem, "OK") Then;* Notify the user to pick up the phone.
				MsgBox(0, "Phone", "Please pick up the phone and either press Enter or click OK")
				ExitLoop
			EndIf
		EndIf
		If (TimerDiff($begin) > $time2wait) Then
			MsgBox(0, "Error", "No response from modem on COM" & $com.CommPort & " in " & $time2wait & " milliseconds")
			Exit
		EndIf
	WEnd

	$com.Output = "ATH" & @CR
	$com.PortOpen = False
	$com = 0

EndFunc;==>Dial

Func MyErrFunc()
; Set @ERROR to 1 and return control to the program following the trapped error
	SetError(1)
	MsgBox(0, "Error", "Error communicating with modem on COM" );& $com.CommPort)
	Exit
EndFunc;==>MyErrFunc
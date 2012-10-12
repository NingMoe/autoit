#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here


#include <Inet.au3>

Local $PublicIP = _GetIP()
MsgBox(0, "IP Address", "Your IP Address is: " & $PublicIP)
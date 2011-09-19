#Include <WinAPI.au3>

WinClose("[CLASS:Progman]", "") 
;WinLogOnOff()

Func WinLogOnOff()
Local $nCount
    $nCount = DllCall("shell32.dll", "int", 60, "hwnd", 0)
        Return $nCount[0]
EndFunc

ConsoleWrite( "DllCall Error " & _WinAPI_GetLastError () & @CRLF & _WinAPI_GetLastErrorMessage ()) ;Error 13: The data is invalid.
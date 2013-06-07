#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
Opt("WinTitleMatchMode", 2)
while 1
				ProcessWait ("winscp.exe")
				;if ProcessExists ("winscp.exe") then
					;sleep(1000*3)
					WinSetState("WinSCP","",@SW_MINIMIZE)
					sleep(1000*3)
					;ProcessWaitClose ( "winscp.exe")
				if not ProcessExists("winscp.exe" ) then Exit
				;EndIf


WEnd
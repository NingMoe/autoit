#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
Opt("WinTitleMatchMode", 2)
while 1
				ProcessWait ("vlc.exe")
				;if ProcessExists ("winscp.exe") then
					;sleep(1000*3)
					WinSetState("VLC","",@SW_MINIMIZE)
					sleep(1000*3)
					;ProcessWaitClose ( "vlc.exe")
				if not ProcessExists("vlc.exe" ) then Exit
				;EndIf


WEnd



#include <array.au3>
#include <_pop3.au3>
#include <IE.au3>
#include <file.au3>
#include <Date.au3>
#include <string.au3> 

dim $sec=@SEC
dim $min=@MIN
Dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR
DIM $WeekOfDay = @WDAY   ; Return 1-7  represent from Sunday , Monday ~ Saturday


dim $batpath=@ScriptDir
dim $bat="tomcat.bat"




Opt("WinTitleMatchMode", 3) 
if ProcessExists("java.exe") and WinExists ("Start Tomcat") then 
	
	MsgBox(0,"Hi there", "tomcat is there.")
	WinKill("Start Tomcat","")
	sleep(500)
	run($batpath&"\"&$bat,$batpath)
	
	sleep(200)
	WinSetTitle("C:\WINNT\system32\cmd.exe","","Start Tomcat")
	WinSetTitle("C:\WINDOWS\system32\cmd.exe","","Start Tomcat")	
	;_batch($batpath,$bat)
EndIf


Exit
	
Func _batch($batpath,$bat)
;Local $bat
;local $batpath
local $pid

Opt("WinTitleMatchMode", -2) 
;; Opt("WinTitleMatchMode", 2) 2 is to set up Autoit to search title by any string in Window title.
;; Opt("WinTitleMatchMode", -2) -2 is to set up Autoit to search title by any string in Window title with Upper Case or Lower case.  
;;
if FileExists($batpath&"\"&$bat) then 
;_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",$bat&" is there")
 
	if WinExists ("- "&StringTrimRight($bat,4)) then 
		;MsgBox(0,"There your Are.", "The window exist");
		WinKill("- "&StringTrimRight($bat,4))
		;_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",$bat&" is killed")
		sleep(500)
		;BlockInput(1)
		
		$pid=run($batpath&"\"&$bat,$batpath)
		_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",$bat&" pid is "&$pid)
		;MsgBox(0,"New window","New Windows PID:"&$pid)
		sleep(200)
		WinSetTitle("C:\WINNT\system32\cmd.exe","","C:\WINNT\system32\cmd.exe - "&$bat)
		WinSetTitle("C:\WINDOWS\system32\cmd.exe","","C:\WINDOWS\system32\cmd.exe - "&$bat)	

		;run("cmd.exe",$batpath)
		;;WinActivate("C:\WINNT\system32\cmd.exe")
		;;if WinActive("C:\WINNT\system32\cmd.exe") then MsgBox(0,"Active","cmd.exe is active",3)
		;sleep(500)
		;send($bat&"{enter}")
		
		;BlockInput(0)
			if WinExists ("- "&$bat) then 
			_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",$bat&" is runing again.")
			EndIf
		;MsgBox(0,"There your Are.", "Wait for the window exist");
		;WinSetState("- "&$bat,"cmd.exe",@SW_MINIMIZE )
		Else
		BlockInput(1)
	
		$pid=run($batpath&"\"&$bat,$batpath)
		_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",$bat&" pid is "&$pid)
		sleep(200)
		;MsgBox(0,"New window","New Windows Pid:"&$pid,3)
		WinSetTitle("C:\WINNT\system32\cmd.exe","","C:\WINNT\system32\cmd.exe - "&$bat)
		WinSetTitle("C:\WINDOWS\system32\cmd.exe","","C:\WINDOWS\system32\cmd.exe - "&$bat)	
		
		;run("cmd.exe",$batpath)
		;;WinActivate("C:\WINNT\system32\cmd.exe")
		;sleep(500)
		;	;if WinActive("c:\winnt\system32\cmd.exe") then MsgBox(0,"Active","cmd.exe is active")
		;send($bat&"{enter}")
		BlockInput(0)
		if WinExists ("- "&$bat) then 
		_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",$bat&" is running now.")
		endif
	EndIf
	sleep(3000)	
	WinSetState("C:\WINNT\system32\cmd.exe - "&$bat, "", @SW_MINIMIZE)
	WinSetState("C:\WINDOWS\system32\cmd.exe - "&$bat, "", @SW_MINIMIZE)
Else
_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",$bat&" is not there")
	
EndIf
EndFunc

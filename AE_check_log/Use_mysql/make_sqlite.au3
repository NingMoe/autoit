#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.2.12.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <TCP.au3>
#include <date.au3>
#include <sqlite.au3>
#include <sqlite.dll.au3>
#include <array.au3>
#include <new_user.au3>

Global $ip
dim $sec=@SEC
dim $min=@MIN
dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR


_SQLite_Startup ()
If @error > 0 Then
    MsgBox(16, "SQLite Error", "SQLite.dll Can't be Loaded!")
    Exit - 1
EndIf
_SQLite_Open (@ScriptDir&"\encryption_sqlite") ; Open a  database file
If @error > 0 Then
    MsgBox(16, "SQLite Error", "Can't Load Database!")
    Exit - 1
EndIf
;;=======================================================================================
;;Creat DB file 
	;_SQLite_Exec(-1,"Create table Encryption (key ,name ,md5 );" )

		;_SQLite_Exec(-1,"Create table Encryption (INTIME ,MD ,EVENT );" )
		;_SQLite_Exec(-1,"insert into Encryption values ('A001', '53611'  ,'660');" )

	if _SQLite_Exec(-1,"select NAME from Encryption where MD5='33333';") <> $SQLITE_OK  then 
		
		MsgBox(0,"SQLite Error","Error Code: " & _SQLite_ErrCode() & @CR & "Error Message: " & _SQLite_ErrMsg(),5)
		_SQLite_Exec(-1,"Create table Encryption ( KEY ,NAME ,MD5 ,INTIME);" )
		_SQLite_Exec(-1,"insert into Encryption values ('11111', '22222'  ,'33333','20090910');" )
		;MsgBox(0,'No Data', 'Insert New data to the db Encryption',5)
	Else
		
	;MsgBox(0,'Update Data', 'Update data from Login',5)
	;_SQLite_Exec(-1,"insert into Login values ('A001', '53611'  ,'662');" )
	;_SQLite_Exec(-1,"Update Login set  EVENT='662' where INTIME='A001'  and  MD='53611';" )
	;_SQLite_Exec(-1,"Update Login set  EVENT='662' where INTIME='A002'  and  MD='53612';" )
	;_SQLite_Exec(-1,"Update Login set  EVENT='662' where INTIME='A003'  and  MD='53613';" )
	;_SQLite_Exec(-1,"Update Login set  EVENT='662' where INTIME='A004'  and  MD='53614';" )
		
	endif 
_SQLite_Close ()
_SQLite_Shutdown ()

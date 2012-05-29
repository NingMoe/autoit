; *******************************************************
; Examples for Access.au3
; *******************************************************
;
#include "Access.au3"

MsgBox (0, "Ex.1", "Example_AccessOpen()")
	Example_AccessOpen()
MsgBox (0, "Ex.2", "Example_AccessTableExists()")
	Example_AccessTableExists()
MsgBox (0, "Ex.3", "Example_AccessTablesCount()")
	Example_AccessTablesCount()
MsgBox (0, "Ex.4", "Example_AccessFieldsCount()")
	Example_AccessFieldsCount()
MsgBox (0, "Ex.5", "Example_AccessFieldExists()")
	Example_AccessFieldExists()
MsgBox (0, "Ex.6", "Example_AccessRecordsCount()")
	Example_AccessRecordsCount()
MsgBox (0, "Ex.7", "Example_AccessTablesList()")
	Example_AccessTablesList()
MsgBox (0, "Ex.8", "Example_AccessFieldsList()")
	Example_AccessFieldsList()
MsgBox (0, "Ex.9", "Example_AccessRecordList()")
	Example_AccessRecordList()
MsgBox (0, "Ex.10", "Example_AccessRecordAdd_Edit_Dele()")
	Example_AccessRecordAdd_Edit_Dele()
;--------------------------------------------------------------
	;++++++++++++later
	; _AccessCreateMDB
	; _AccessCreateField ($oTable, "fld1" )
	; _AccessUpdateTable ($oNewDB, $oTable)
	; _AccessCreateTable (ByRef $o_object, $s_TableName ="")
;================================================================

Exit

; *******************************************************
; Example 1 - Open existing Access file
; *******************************************************
;
Func Example_AccessOpen()
	Local $o_DataBase = _AccessOpen(@ScriptDir & "\Test.mdb")
	If $o_DataBase = 0 Then
		MsgBox(0, "Information", "Database file is not found :-" & @CR & @ScriptDir & "\Test.mdb")
		Return
	Else
		MsgBox(0, "Information", "Database file was opened :-" & @CR & @ScriptDir & "\Test.mdb")
	EndIf

	_AccessClose($o_DataBase)
EndFunc   ;==>Example_AccessOpen

; *******************************************************
; Example 2 - Open existing Access file, Check if table is Exisits
; *******************************************************
;
Func Example_AccessTableExists()
	Local $o_DataBase = _AccessOpen(@ScriptDir & "\Test.mdb")
	Local $O_TableObject

	If $o_DataBase = 0 Then
		MsgBox(0, "Information", "Database file is not found :-" & @CR & @ScriptDir & "\Test.mdb")
		Return
	EndIf

	$O_TableObject = _AccessTableExists ($o_DataBase, "tblTable1")
	If $O_TableObject = 0 Then
		MsgBox(0, "Information", "Table is not Exists. ")
		Return
	Else
		MsgBox(0, "Information", "Table is Exists. ")
	EndIf

	_AccessClose($o_DataBase)
EndFunc   ;==>Example_AccessTableExists

; *******************************************************
; Example 3 - Open existing Access file, Count Tables
; *******************************************************
;
Func Example_AccessTablesCount()
	Local $o_DataBase = _AccessOpen(@ScriptDir & "\Test.mdb")
	Local $n_UserTables, $n_All_Tables

	If $o_DataBase = 0 Then
		MsgBox(0, "Information", "Database file is not found :-" & @CR & @ScriptDir & "\Test.mdb")
		Return
	EndIf

	$n_UserTables = _AccessTablesCount ($o_DataBase)
	$n_All_Tables = _AccessTablesCount ($o_DataBase, False)

	MsgBox(0, "Information", "User Table count is " & $n_UserTables)

	MsgBox(0, "Information", " All Table count is " & $n_All_Tables)



	_AccessClose($o_DataBase)
EndFunc   ;==>Example_AccessTableExists

; *******************************************************
; Example 4 - Count Number of fields in table
; *******************************************************
;
Func Example_AccessFieldsCount()
	Local $o_DataBase = _AccessOpen(@ScriptDir & "\Test.mdb")
	;Local $O_TableObject

	If $o_DataBase = 0 Then
		MsgBox(0, "Information", "Database file is not found :-" & @CR & @ScriptDir & "\Test.mdb")
		Return
	EndIf

	$n_Fields_Count = _AccessFieldsCount ($o_DataBase, "tblTable1")

	MsgBox (0, "Information", "No of Fields in Table1 is " & $n_Fields_Count)

	_AccessClose($o_DataBase)

EndFunc

; *******************************************************
; Example 5 - Check if a field is Exists
; *******************************************************
;
Func Example_AccessFieldExists()
	Local $o_DataBase = _AccessOpen(@ScriptDir & "\Test.mdb")
	Local $O_TableObject

	If $o_DataBase = 0 Then
		MsgBox(0, "Information", "Database file is not found :-" & @CR & @ScriptDir & "\Test.mdb")
		Return
	EndIf

	$O_TableObject = _AccessTableExists ($o_DataBase, "tblTable1")
	If $O_TableObject = 0 Then
		MsgBox(0, "Information", "tblTable1 Table is not Exists. ")
		Return
	EndIf

	Local $b_FieldExists = _AccessFieldExists ($O_TableObject, "fld1")
	if $b_FieldExists = 0 Then
		MsgBox(0, "Information", "Field fld1 is not Exists. ")
		Return
	Else
		MsgBox (0, "Information", "Field fld1 is Exists.")
	EndIf

	_AccessClose($o_DataBase)
EndFunc

; *******************************************************
; Example 6 - Count Number of fields in table
; *******************************************************
Func Example_AccessRecordsCount()
	Local $o_DataBase = _AccessOpen(@ScriptDir & "\Test.mdb")

	If $o_DataBase = 0 Then
		MsgBox(0, "Information", "Database file is not found :-" & @CR & @ScriptDir & "\Test.mdb")
		Return
	EndIf

	Local $n_RecCount = _AccessRecordsCount ($o_DataBase, "tblTable1")

	MsgBox(0, "Information", "No. Of Recordes in Table  is " & $n_RecCount)

	_AccessClose($o_DataBase)
EndFunc

; *******************************************************
; Example 7 - List tables in database
; *******************************************************

Func Example_AccessTablesList()
	Local $o_DataBase = _AccessOpen(@ScriptDir & "\Test.mdb")
	Local $avArray_Tables

	If $o_DataBase = 0 Then
		MsgBox(0, "Information", "Database file is not found :-" & @CR & @ScriptDir & "\Test.mdb")
		Return
	EndIf

	$avArray_Tables = _AccessTablesList ($o_DataBase)
	_ArrayDisplay ($avArray_Tables)

	$avArray_Tables = _AccessTablesList ($o_DataBase, False)
	_ArrayDisplay ($avArray_Tables)

	_AccessClose($o_DataBase)

EndFunc
; *******************************************************
; Example 8 - List Fields tables
; *******************************************************

Func Example_AccessFieldsList()
	Local $o_DataBase = _AccessOpen(@ScriptDir & "\Test.mdb")
	Local $O_TableObject, $avArray_Fields ;, $O_TableObject

	If $o_DataBase = 0 Then
		MsgBox(0, "Information", "Database file is not found :-" & @CR & @ScriptDir & "\Test.mdb")
		Return
	EndIf


	$avArray_Fields = _AccessFieldsList ($o_DataBase, "tblTable1")
	_ArrayDisplay ($avArray_Fields)
	_AccessClose($o_DataBase)
EndFunc
; *******************************************************
; Example 9 - List the current Record
; *******************************************************
Func Example_AccessRecordList()
	Local $o_DataBase = _AccessOpen(@ScriptDir & "\Test.mdb")
	Local $O_TableObject, $avArray_Record ;, $O_TableObject

	If $o_DataBase = 0 Then
		MsgBox(0, "Information", "Database file is not found :-" & @CR & @ScriptDir & "\Test.mdb")
		Return
	EndIf

	$avArray_Record = _AccessRecordList ($o_DataBase, "tblTable1", 2)
	_ArrayDisplay ($avArray_Record)
	_AccessClose($o_DataBase)
EndFunc

; *******************************************************
; Example 10 - Add Record
; *******************************************************
Func Example_AccessRecordAdd_Edit_Dele()
	Local $o_DataBase = _AccessOpen(@ScriptDir & "\Test.mdb")
	Local $avData[5][2]
	Local $n_Feedback

	If $o_DataBase = 0 Then
		MsgBox(0, "Information", "Database file is not found :-" & @CR & @ScriptDir & "\Test.mdb")
		Return
	EndIf
	;---------- Data Setup
		$avData[0] [0] = "FLD1"
		$avData[1] [0] = "FLD2"
		$avData[2] [0] = "FLD3"
		$avData[3] [0] = "FLD4"

		$avData[4] [0] = "FLD1"


		$avData[0] [1] = "Will not be Written"
		$avData[1] [1] = 150
		$avData[2] [1] = @MDAY & "/" & @MON & "/" & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC
		$avData[3] [1] = 25

		$avData[4] [1] = "Fld1 is reset"
	;---------- Get the last Rec. position
	Local $n_RecCount = _AccessRecordsCount ($o_DataBase, "tblTable1")
	MsgBox (0,"Info.", "Before add data, Record Count =" & $n_RecCount )

	;---------- Save Data
	$n_Feedback = _AccessRecordAdd ($o_DataBase, "tblTable1", $avData)
	if $n_Feedback <> 1 Then
			MsgBox (0, "Info. feedback", "Error in writting Data")
			Return
	EndIf

	;---------- Get the last Rec. position
	Local $n_RecCount = _AccessRecordsCount ($o_DataBase, "tblTable1")
	MsgBox (0,"Info.", "After add data, Record Count =" & $n_RecCount )

	;----------- Show Last Entered Data
	Local $avArray_Record = _AccessRecordList ($o_DataBase, "tblTable1", $n_RecCount )
	_ArrayDisplay ($avArray_Record)

	;---------- Edit Data
		$avData[0] [1] = "Will not be Written"
		$avData[1] [1] = 11111
		$avData[2] [1] = @MDAY & "/" & @MON & "/" & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC
		$avData[3] [1] = 1111

		$avData[4] [1] = "All 1111"

	$n_Feedback = _AccessRecordEdit ($o_DataBase, "tblTable1", $avData, $n_RecCount)
	if $n_Feedback <> 1 Then
			MsgBox (0, "Info. feedback", "Error in writting Data")
			Return
	EndIf

	;----------- Show Last Entered Data
	Local $avArray_Record = _AccessRecordList ($o_DataBase, "tblTable1", $n_RecCount )
	_ArrayDisplay ($avArray_Record)

	;---------- Delete last Rec.
	Local $n_FeedBack = _AccessRecordDelete ($o_DataBase, "tblTable1", $n_RecCount  )
	if $n_FeedBack = 1 Then
		MsgBox (0,"info.", "Data was Deleted")
		$n_RecCount =  _AccessRecordsCount ($o_DataBase, "tblTable1")
		MsgBox (0,"Info.", "After Delete, Record Count =" & $n_RecCount )
	Else
		MsgBox (0,"info.", "Data was NOT Deleted")
	EndIf

	;---------- Close Database
	_AccessClose($o_DataBase)

EndFunc


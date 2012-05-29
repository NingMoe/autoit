
; --- Rev 02--------------
; Fix bug in _AccessRecordList, bad return value
; Bug in Error Tape for _AccessRecordsCount

; #INDEX# =======================================================================================================================
; Title .........: Microsoft Access Automation UDF Library for AutoIt3
; AutoIt Version : 3.2.0.
; Language ......: English
; Description ...: A collection of functions for creating, attaching to, reading from and manipulating Microsoft Access .
; Author(s) .....: Ayman Henry (aymhenry@gmail.com)
; Dll ...........: user32.dll
; ===============================================================================================================================

; ------------------------------------------------------------------------------
;	Version: V0.0
;	Last Update: 4/May/2012
;	Requirements: AutoIt v3.2.0.1 or higher
;	Notes: Errors associated with incorrect objects will be common user errors.
;		Creating it agine after this UDF is now out of development and is no longer supported.
; 		http://www.autoitscript.com/forum/topic/40397-msaccess-udf-updated/page__p__318703__hl__msaccess%20phone__fromsearch__1#entry318703
;	Update History:
; ------------------------------------------------------------------------------

#include-once
#include <Array.au3>

; #VARIABLES# ===================================================================================================================

Global Enum _; Error Status Types
		$_AccessStatus_Success = 0, _
		$_AccessStatus_GeneralError, _
		$_AccessStatus_ComError, _
		$_AccessStatus_InvalidDataType, _
		$_AccessStatus_InvalidObjectType, _
		$_AccessStatus_InvalidValue, _
		$_AccessStatus_ReadOnly, _
		$_AccessStatus_NoMatch, _
		$_AccessStatus_DataTypeMisMatch, _
		$_AccessStatus_TableExists

Global $oAccessErrorHandler, $sAccessUserErrorHandler
Global $_AccessErrorNotify = True
Global $__AccessAU3Debug = False
Global _; Com Error Handler Status Strings
		$AccessComErrorNumber, _
		$AccessComErrorNumberHex, _
		$AccessComErrorDescription, _
		$AccessComErrorScriptline, _
		$AccessComErrorWinDescription, _
		$AccessComErrorSource, _
		$AccessComErrorHelpFile, _
		$AccessComErrorHelpContext, _
		$AccessComErrorLastDllError, _
		$AccessComErrorComObj, _
		$AccessComErrorOutput


Global $dbBigInt = 16
Global $dbBinary = 9
Global $dbBoolean = 1
Global $dbByte = 2
Global $dbChar = 18
Global $dbCurrency = 5
Global $dbDate = 8
Global $dbDecimal = 20
Global $dbDouble = 7
Global $dbFloat = 21
Global $dbGUID = 15
Global $dbInteger = 3
Global $dbLong = 4
Global $dbLongBinary = 11
Global $dbMemo = 12
Global $dbNumeric = 19
Global $dbSingle = 6
Global $dbText = 10
Global $dbTime = 22
Global $dbTimeStamp = 23
Global $dbVarBinary = 17


; #Later# =====================================================================================================================

; _AccessCreateTable()
; _AccessUpdateTable($oNewDB, $oTable)
; _AccessTabelDelete()
; _AccessQueryLike()
; _AccessQueryStr()

; #CURRENT# =====================================================================================================================
; _AccessOpen()
; _AccessClose()

; _AccessTableExists()
; _AccessTablesCount()
; _AccessTablesList()

; _AccessFieldExists
; _AccessFieldsList
; _AccessFieldsCount()

; _AccessRecordsCount()
; _AccessRecordList ()
; _AccessRecordAdd()
; _AccessRecordEdit()
; _AccessRecordDelete ()

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __AccessIsObjType
; __AccessCreateDB()


; #FUNCTION# ====================================================================================================================
; Name...........: _AccessOpen
; Description ...: Open Microsoft Office Access File
; Parameters ....: $s_FilePath	- File path and name
;					$Options	- Sets various options for the database, as specified in Remarks
;									True Opens the database in exclusive (exclusive: A type of access to data in a database that is shared over a network.
;										When you open a database in exclusive mode, you prevent others from opening the database.) mode.
;									False (Default) Opens the database in shared mode.
;					$ReadOnly   - True if you want to open the database with read-only access, or False (default) if you want to open the database with read/write access.
;					$Connect	- Specifies various connection information, including passwords.
; Return values .: On Success	- Returns database object linked to the opened file
;                  On Failure	- Returns 0 and sets @ERROR
;					@ERROR		- 0 ($_AccessStatus_Success) = No Error
;								- 1 ($_AccessStatus_GeneralError) = General Error
;								- 3 ($_AccessStatus_InvalidDataType) = Invalid Data Type
;								- 4 ($_AccessStatus_InvalidObjectType) = Invalid Object Type
; Author ........: Ayman Henry
; ADO Command ...: expression.OpenDatabase(Name, Options, ReadOnly, Connect)
; ===============================================================================================================================
Func _AccessOpen($s_FilePath = "", $Options = False, $ReadOnly = False, $Connect = ";pwd=")
	Local $o_object, $o_MDBObj

	If $s_FilePath = "" Then Return 0 ; There is currently no way of read non existing db.
	If FileExists($s_FilePath) = 0 Then Return 0

	Local $result, $f_mustUnlock = 0, $i_ErrorStatusCode = $_AccessStatus_Success

	; Setup internal error handler to Trap COM errors, turn off error notification
	Local $status = __AccessInternalErrorHandlerRegister()
	If Not $status Then __AccessErrorNotify("Warning", "_AccessOpen", _
			"Cannot register internal error handler, cannot trap COM errors", _
			"Use _AccessErrorHandlerRegister() to register a user error handler")
	Local $f_NotifyStatus = _AccessErrorNotify() ; save current error notify status

	_AccessErrorNotify(False)
	$o_object = ObjCreate("DAO.DBEngine.36")

	If Not IsObj($o_object) Or @error = $_AccessStatus_ComError Then
		$i_ErrorStatusCode = $_AccessStatus_NoMatch
	EndIf

	; restore error notify and error handler status
	_AccessErrorNotify($f_NotifyStatus) ; restore notification status
	__AccessInternalErrorHandlerDeRegister()

	If Not $i_ErrorStatusCode = $_AccessStatus_Success Then
		;$o_object = ObjCreate("Access.Application")
		;$o_object = ObjCreate("DAO.DBEngine.36")

		If Not IsObj($o_object) Then
			__AccessErrorNotify("Error", "_AccessOpen", "", "Access Object Creation Failed")
			Return SetError($_AccessStatus_GeneralError, 0, 0)
		EndIf
	EndIf

	$o_MDBObj = __AccessOpenDB($o_object, $s_FilePath, $Options, $ReadOnly, $Connect)

	Return SetError(@error, 0, $o_MDBObj)
EndFunc   ;==>_AccessOpen

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __AccessInternalErrorHandlerRegister
; Description ...: Register and enable the internal Access COM error handler
; Parameters ....: None
; Return values .: Success  - 1
;                  Failure  - 0
; Author ........: Ayman Henry (based on Code written on Word.au3)
; Modified.......:
; Remarks .......:
; ===============================================================================================================================
Func __AccessInternalErrorHandlerRegister()
	Local $sCurrentErrorHandler = ObjEvent("AutoIt.Error")
	If $sCurrentErrorHandler <> "" And Not IsObj($oAccessErrorHandler) Then
		; We've got trouble... User COM Error handler assigned without using _AccessErrorHandlerRegister
		Return SetError($_AccessStatus_GeneralError, 0, 0)
	EndIf
	$oAccessErrorHandler = ""
	$oAccessErrorHandler = ObjEvent("AutoIt.Error", "__AccessInternalErrorHandler")
	If IsObj($oAccessErrorHandler) Then Return SetError($_AccessStatus_Success, 0, 1)
	Return SetError($_AccessStatus_GeneralError, 0, 0)
EndFunc   ;==>__AccessInternalErrorHandlerRegister

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __AccessInternalErrorHandlerDeRegister
; Description ...: DeRegister the internal Access COM error handler and register User Access COM error handler if defined
; Parameters ....: None
; Return values .: 1
; Author ........: Ayman Henry (based on Code written on Word.au3)
; Modified.......:
; Remarks .......:
; ===============================================================================================================================
Func __AccessInternalErrorHandlerDeRegister()
	$oAccessErrorHandler = ""
	If $sAccessUserErrorHandler <> "" Then
		$oAccessErrorHandler = ObjEvent("AutoIt.Error", $sAccessUserErrorHandler)
	EndIf
	Return SetError($_AccessStatus_Success, 0, 1)
EndFunc   ;==>__AccessInternalErrorHandlerDeRegister

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __AccessInternalErrorHandler
; Description ...: Update $AccessComError... global variables
; Parameters ....: None
; Return values .: @error = $_AccessStatus_ComError
; Author ........: Ayman Henry (based on Code written on Word.au3)
; Modified.......:
; Remarks .......:
; ===============================================================================================================================
Func __AccessInternalErrorHandler()
	$AccessComErrorScriptline = $oAccessErrorHandler.scriptline
	$AccessComErrorNumber = $oAccessErrorHandler.number
	$AccessComErrorNumberHex = Hex($oAccessErrorHandler.number, 8)
	$AccessComErrorDescription = StringStripWS($oAccessErrorHandler.description, 2)
	$AccessComErrorWinDescription = StringStripWS($oAccessErrorHandler.WinDescription, 2)
	$AccessComErrorSource = $oAccessErrorHandler.Source
	$AccessComErrorHelpFile = $oAccessErrorHandler.HelpFile
	$AccessComErrorHelpContext = $oAccessErrorHandler.HelpContext
	$AccessComErrorLastDllError = $oAccessErrorHandler.LastDllError
	$AccessComErrorOutput = ""
	$AccessComErrorOutput &= "--> COM Error Encountered in " & @ScriptName & @CRLF
	$AccessComErrorOutput &= "----> $AccessComErrorScriptline = " & $AccessComErrorScriptline & @CRLF
	$AccessComErrorOutput &= "----> $AccessComErrorNumberHex = " & $AccessComErrorNumberHex & @CRLF
	$AccessComErrorOutput &= "----> $AccessComErrorNumber = " & $AccessComErrorNumber & @CRLF
	$AccessComErrorOutput &= "----> $AccessComErrorWinDescription = " & $AccessComErrorWinDescription & @CRLF
	$AccessComErrorOutput &= "----> $AccessComErrorDescription = " & $AccessComErrorDescription & @CR
	$AccessComErrorOutput &= "----> $AccessComErrorSource = " & $AccessComErrorSource & @CRLF
	$AccessComErrorOutput &= "----> $AccessComErrorHelpFile = " & $AccessComErrorHelpFile & @CRLF
	$AccessComErrorOutput &= "----> $AccessComErrorHelpContext = " & $AccessComErrorHelpContext & @CRLF
	$AccessComErrorOutput &= "----> $AccessComErrorLastDllError = " & $AccessComErrorLastDllError & @CRLF
	If $_AccessErrorNotify Or $__AccessAU3Debug Then ConsoleWrite($AccessComErrorOutput & @CRLF)
	Return SetError($_AccessStatus_ComError)
EndFunc   ;==>__AccessInternalErrorHandler

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __AccessOpenDB
; Description ...:
; Syntax ........: __AccessOpenDB(Byref $o_object, $s_FilePath[, $Options = False[, $ReadOnly = False[, $Connect = ";pwd="]]])
; Parameters ....: $o_object            - [in/out] Object variable of a Access.Application object.
;                  $s_FilePath          - A string value. Full path of the document to open
;                  $Options             - [optional] Default is False.
;											Sets various options for the database, as specified in Remarks
;											True Opens the database in exclusive (exclusive: A type of access to data in a database that is shared over a network.
;										When you open a database in exclusive mode, you prevent others from opening the database.) mode.
;                  $ReadOnly            - [optional] Default is False.True if you want to open the database with read-only access, or False (default) if you want to open the database with read/write access.
;                  $Connect             - [optional] Default is ";pwd=". Specifies various connection information, including passwords.
; Return values .: None
; Author ........: Ayman Henry
; Modified ......:
; Remarks .......: Under modification - not used
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __AccessOpenDB(ByRef $o_object, $s_FilePath, $Options = False, $ReadOnly = False, $Connect = ";pwd=")
	If Not FileExists($s_FilePath) Then Return 0
	;--------
	Local $o_doc

	If Not IsObj($o_object) Then
		__AccessErrorNotify("Error", "_AccessOpenDB", "$_AccessStatus_InvalidDataType")
		Return SetError($_AccessStatus_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __AccessIsObjType($o_object, "application") Then
		__AccessErrorNotify("Error", "_AccessOpenDB", "$_AccessStatus_InvalidObjectType")
		Return SetError($_AccessStatus_InvalidObjectType, 1, 0)
	EndIf
	;
	;$o_doc = $o_object.OpenDatabase ( $s_FilePath); ,True, True, ";pwd=" )
	$o_doc = $o_object.OpenDatabase($s_FilePath, $Options, $ReadOnly, $Connect)

	If Not IsObj($o_doc) Then
		__AccessErrorNotify("Error", "_AccessOpenDB", "", "Document Object Creation Failed")
		Return SetError($_AccessStatus_GeneralError, 0, 0)
	EndIf

	Return SetError($_AccessStatus_Success, 0, $o_doc)
EndFunc   ;==>__AccessOpenDB

; #FUNCTION# ====================================================================================================================
; Name ..........: _AccessClose
; Description ...:
; Syntax ........: _AccessClose(Byref $o_object)
; Parameters ....: $o_object            - [in/out] database object to close.
; Return values .: On Success 	- Returns 1
;                  On Failure	- Returns 0
; Author ........: Ayman Henry
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AccessClose(ByRef $o_object)

	If Not IsObj($o_object) Then
		__AccessErrorNotify("Error", "_AccessClose", "$_AccessStatus_InvalidDataType")
		Return SetError($_AccessStatus_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __AccessIsObjType($o_object, "database") Then
		__AccessErrorNotify("Error", "_AccessClose", "$_AccessStatus_InvalidObjectType")
		Return SetError($_AccessStatus_InvalidObjectType, 1, 0)
	EndIf
	;
	$o_object.Close

	Return SetError($_AccessStatus_Success, 1, 0)
EndFunc   ;==>_AccessClose

; #FUNCTION# ====================================================================================================================
; Function Name:   _AccessErrorNotify
; Description ...: Specifies whether Access.au3 automatically notifies of Warnings and Errors (to the console)
; Parameters ....: $f_notify	- Optional: specifies whether notification should be on or off
;								- -1 = (Default) return current setting
;								- True = Turn On
;								- False = Turn Off
; Return values .: On Success 	- If $f_notify = -1, returns the current notification setting, else returns 1
;                  On Failure	- Returns 0
; Author ........: Ayman Henry (based on Code written  on Word.au3)
; ===============================================================================================================================
Func _AccessErrorNotify($f_notify = -1)
	Switch Number($f_notify)
		Case -1
			Return $_AccessErrorNotify
		Case 0
			$_AccessErrorNotify = False
			Return 1
		Case 1
			$_AccessErrorNotify = True
			Return 1
		Case Else
			__AccessErrorNotify("Error", "_AccessErrorNotify", "$_AccessStatus_InvalidValue")
			Return 0
	EndSwitch
EndFunc   ;==>_AccessErrorNotify

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __AccessErrorNotify
; Description ...:
; Syntax ........: __AccessErrorNotify($s_severity, $s_func[, $s_status = ""[, $s_message = ""]])
; Parameters ....: $s_severity          - A string value.
;                  $s_func              - A string value.
;                  $s_status            - [optional] A string value. Default is "".
;                  $s_message           - [optional] A string value. Default is "".
; Return values .: None
; Author ........: Ayman Henry (based on Code written  on Word.au3)
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __AccessErrorNotify($s_severity, $s_func, $s_status = "", $s_message = "")
	If $_AccessErrorNotify Or $__AccessAU3Debug Then
		Local $sStr = "--> Access.au3 " & $s_severity & " from function " & $s_func
		If Not $s_status = "" Then $sStr &= ", " & $s_status
		If Not $s_message = "" Then $sStr &= " (" & $s_message & ")"
		ConsoleWrite($sStr & @CRLF)
	EndIf
	Return 1
EndFunc   ;==>__AccessErrorNotify

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __AccessIsObjType
; Description ...: Check to see if an object variable is of a specific type
; Syntax ........: __AccessIsObjType(Byref $o_object, $s_type)
; Parameters ....: $o_object            - [in/out] an object to be checked.
;                  $s_type              - Type Name, that the object has to be of its type.
; Return values .: True / False the check result
; Author ........: Ayman Henry (based on code in word.au3)
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __AccessIsObjType(ByRef $o_object, $s_type)
	If Not IsObj($o_object) Then Return SetError($_AccessStatus_InvalidDataType, 1, 0)

	; Setup internal error handler to Trap COM errors, turn off error notification
	Local $status = __AccessInternalErrorHandlerRegister()
	If Not $status Then __AccessErrorNotify("Warning", "internal function __AccessIsObjType", _
			"Cannot register internal error handler, cannot trap COM errors", _
			"Use _AccessErrorHandlerRegister() to register a user error handler")
	Local $f_NotifyStatus = _AccessErrorNotify() ; save current error notify status
	_AccessErrorNotify(False)
	;
	Local $s_Name = StringLower(ObjName($o_object)), $objectOK = False

	Switch StringLower($s_type)
		Case "application"
			If $s_Name = "_dbengine" Then $objectOK = True
		Case "database"
			If $s_Name = "database" Then $objectOK = True
		Case "_tabledef"
			If $s_Name = "_tabledef" Then $objectOK = True
		Case "_field"
			If $s_Name = "_field" Then $objectOK = True
		Case "RecordSet"
			If $s_Name = "RecordSet" Then $objectOK = True

		Case Else
			; Unsupported ObjType specified
			Return SetError($_AccessStatus_InvalidValue, 2, 0)
	EndSwitch

	; restore error notify and error handler status
	_AccessErrorNotify($f_NotifyStatus) ; restore notification status
	__AccessInternalErrorHandlerDeRegister()

	If $objectOK Then
		Return SetError($_AccessStatus_Success, 0, 1)
	Else
		Return SetError($_AccessStatus_InvalidObjectType, 1, 0)
	EndIf
EndFunc   ;==>__AccessIsObjType

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __AccessCreateDB
; Description ...: Returns an object variable representing a new empty document
; Syntax ........: __AccessCreateDB(Byref $o_object, $sFileLocation[, $Options = ""[, $ReadOnly = False[, $Connect = ""]]])
; Parameters ....: $o_object            - [in/out] Object variable of a Access.Application object.
;                  $s_FilePath          - A string value. Full path of the document to open
;                  $Options             - [optional] Default is False.
;											Sets various options for the database, as specified in Remarks
;											True Opens the database in exclusive (exclusive: A type of access to data in a database that is shared over a network.
;										When you open a database in exclusive mode, you prevent others from opening the database.) mode.
;                  $ReadOnly            - [optional] Default is False.True if you want to open the database with read-only access, or False (default) if you want to open the database with read/write access.
;                  $Connect             - [optional] Default is ";pwd=". Specifies various connection information, including passwords.
; Return values .: On Success	- Returns an object variable pointing to a Access.Application, document object
;                  On Failure	- Returns 0 and sets @ERROR
;					@ERROR		- 0 ($_AccessStatus_Success) = No Error
;								- 1 ($_AccessStatus_GeneralError) = General Error
;								- 3 ($_AccessStatus_InvalidDataType) = Invalid Data Type
;								- 4 ($_AccessStatus_InvalidObjectType) = Invalid Object Type
; Author ........: Ayman Henry
; Modified ......:
; Remarks .......: Under modification - not used
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __AccessCreateDB(ByRef $o_object, $sFileLocation, $Options = "", $ReadOnly = False, $Connect = "")

	If Not IsObj($o_object) Then
		__AccessErrorNotify("Error", "__AccessCreateDB", "$_AccessStatus_InvalidDataType")
		Return SetError($_AccessStatus_InvalidDataType, 1, 0)
	EndIf

	If Not __AccessIsObjType($o_object, "application") Then
		__AccessErrorNotify("Error", "__AccessCreateDB", "$_AccessStatus_InvalidObjectType")
		Return SetError($_AccessStatus_InvalidObjectType, 1, 0)
	EndIf

	Local $o_doc = $o_object.CreateDatabase($sFileLocation, ";LANGID=0x0409;CP=1252;COUNTRY=0") ; dbLangGeneral

	If Not IsObj($o_doc) Then
		__AccessErrorNotify("Error", "__AccessCreateDB", "", "Document Object Creation Failed")
		Return SetError($_AccessStatus_GeneralError, 0, 0)
	EndIf

	Return SetError($_AccessStatus_Success, 0, $o_doc)
EndFunc   ;==>__AccessCreateDB

; #FUNCTION# ====================================================================================================================
; Name ..........: _AccessCreateTable
; Description ...:
; Syntax ........: _AccessCreateTable(Byref $o_object[, $s_TableName = ""])
; Parameters ....: $o_object            - [in/out] databse object.
;                  $s_TableName         - [optional] A string value. Default is "". Tavle name to be craeted
; Return values .: None
; Author ........: Ayman Henry
; Modified ......:
; Remarks .......: Under modification - not used
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AccessCreateTable(ByRef $o_object, $s_TableName = "")
	Local $o_doc

	If $s_TableName = "" Then Return 0

	If _AccessTableExists($o_object, $s_TableName) = True Then
		__AccessErrorNotify("Error", "_AccessCreateTable", "$_AccessStatus_TableExists")
		Return SetError($_AccessStatus_TableExists, 1, 0)
	EndIf

	If Not IsObj($o_object) Then
		__AccessErrorNotify("Error", "_AccessCreateTable", "$_AccessStatus_InvalidDataType")
		Return SetError($_AccessStatus_InvalidDataType, 1, 0)
	EndIf

	If Not __AccessIsObjType($o_object, "database") Then
		__AccessErrorNotify("Error", "_AccessCreateTable", "$_AccessStatus_InvalidObjectType")
		Return SetError($_AccessStatus_InvalidObjectType, 1, 0)
	EndIf


	Local $o_doc = $o_object.CreateTableDef($s_TableName)

	If Not IsObj($o_doc) Then
		__AccessErrorNotify("Error", "_AccessCreateTable", "", "Document Object Creation Failed")
		Return SetError($_AccessStatus_GeneralError, 0, 0)
	EndIf

	If Not __AccessIsObjType($o_doc, "_TableDef") Then
		__AccessErrorNotify("Error", "_AccessCreateTable", "$_AccessStatus_InvalidObjectType")
		Return SetError($_AccessStatus_InvalidObjectType, 1, 0)
	EndIf

	Return SetError($_AccessStatus_Success, 0, $o_doc)

EndFunc   ;==>_AccessCreateTable

; #FUNCTION# ====================================================================================================================
; Name ..........: _AccessCreateField
; Description ...:
; Syntax ........: _AccessCreateField(Byref $o_object[, $s_FieldName = ""[, $s_FieldType = 10]])
; Parameters ....: $o_object            - [in/out] Database object.
;                  $s_FieldName         - [optional] A string value. Default is "". Field Name
;                  $s_FieldType         - [optional] A string value. Default is 10. i.e Text Type Field
; Return values .: None
; Author ........: Ayman Henry
; Modified ......:
; Remarks .......: Under modification - not used
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AccessCreateField(ByRef $o_object, $s_FieldName = "", $s_FieldType = 10)
	Local $o_doc

	If $s_FieldName = "" Then Return 0

	If _AccessFieldExists($o_object, $s_FieldName) = True Then Return 0

	If Not IsObj($o_object) Then
		__AccessErrorNotify("Error", "_AccessCreateField", "$_AccessStatus_InvalidDataType")
		Return SetError($_AccessStatus_InvalidDataType, 1, 0)
	EndIf

	If Not __AccessIsObjType($o_object, "_TableDef") Then
		__AccessErrorNotify("Error", "_AccessCreateField", "$_AccessStatus_InvalidObjectType")
		Return SetError($_AccessStatus_InvalidObjectType, 1, 0)
	EndIf


	Local $o_doc = $o_object.CreateField($s_FieldName, $s_FieldType)

	If Not IsObj($o_doc) Then
		__AccessErrorNotify("Error", "_AccessCreateField", "", "Document Object Creation Failed")
		Return SetError($_AccessStatus_GeneralError, 1, 0)
	EndIf


	If Not __AccessIsObjType($o_doc, "_Field") Then
		__AccessErrorNotify("Error", "_AccessCreateField", "$_AccessStatus_InvalidObjectType")
		Return SetError($_AccessStatus_InvalidObjectType, 1, 0)
	EndIf

	$o_object.Fields.Append($o_doc)
	;$o_object.TableDefs.Append ($o_doc)

	Return SetError($_AccessStatus_Success, 0, $o_doc)

EndFunc   ;==>_AccessCreateField

; #FUNCTION# ====================================================================================================================
; Name ..........: _AccessUpdateTable
; Description ...:
; Syntax ........: _AccessUpdateTable(Byref $o_object, Byref $o_Table)
; Parameters ....: $o_object            - [in/out] Database Object.
;                  $o_Table             - [in/out] Table name to update.
; Return values .: None
; Author ........: Ayman Henry
; Modified ......:
; Remarks .......: Under modification - not used
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AccessUpdateTable(ByRef $o_object, ByRef $o_Table)
	Local $o_doc

	If Not IsObj($o_object) Then
		__AccessErrorNotify("Error", "_AccessUpdateTable", "$_AccessStatus_InvalidDataType")
		Return SetError($_AccessStatus_InvalidDataType, 1, 0)
	EndIf

	If Not __AccessIsObjType($o_object, "database") Then
		__AccessErrorNotify("Error", "_AccessUpdateTable", "$_AccessStatus_InvalidObjectType")
		Return SetError($_AccessStatus_InvalidObjectType, 1, 0)
	EndIf

	If Not IsObj($o_Table) Then
		__AccessErrorNotify("Error", "_AccessUpdateTable", "$_AccessStatus_InvalidDataType")
		Return SetError($_AccessStatus_InvalidDataType, 1, 0)
	EndIf

	If Not __AccessIsObjType($o_Table, "_TableDef") Then
		__AccessErrorNotify("Error", "_AccessUpdateTable", "$_AccessStatus_InvalidObjectType")
		Return SetError($_AccessStatus_InvalidObjectType, 1, 0)
	EndIf

	Local $o_doc = $o_object.TableDefs.Append($o_Table)

	If Not IsObj($o_doc) Then
		__AccessErrorNotify("Error", "_AccessCreateTable", "", "Document Object Creation Failed")
		Return SetError($_AccessStatus_GeneralError, 0, 0)
	EndIf

	Return SetError($_AccessStatus_Success, 0, $o_doc)

EndFunc   ;==>_AccessUpdateTable

; #FUNCTION# ====================================================================================================================
; Name ..........: _AccessTablesCount
; Description ...: Return number of tables in a database
; Syntax ........: _AccessTablesCount(Byref $o_DataBaseObject[, $b_OnlyNonSysTable = True])
; Parameters ....: $o_DataBaseObject    - [in/out] database object.
;                  $b_OnlyNonSysTable   - [optional] A binary value. Default is True. if True: do not Count also system tables
; Return values .: Number of tables on the given database
; Author ........: Ayman Henry
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AccessTablesCount(ByRef $o_DataBaseObject, $b_OnlyNonSysTable = True)
	Local $n_Tables, $nCnt, $n_TablesCount = 0

	If Not IsObj($o_DataBaseObject) Then
		__AccessErrorNotify("Error", "_AccessTablesCount", "$_AccessStatus_InvalidDataType")
		Return SetError($_AccessStatus_InvalidDataType, 1, 0)
	EndIf

	If Not __AccessIsObjType($o_DataBaseObject, "database") Then
		__AccessErrorNotify("Error", "_AccessTablesCount", "$_AccessStatus_InvalidObjectType")
		Return SetError($_AccessStatus_InvalidObjectType, 1, 0)
	EndIf
	;if BitAND ($o_DataBaseObject.TableDefs($nCnt).Attributes,-2147483646) =0  Then ; dbSystemObject -2147483646

	$n_Tables = $o_DataBaseObject.TableDefs.Count

	If $b_OnlyNonSysTable = False Then
		Return $n_Tables
	EndIf

	For $nCnt = 0 To $n_Tables - 1
		If BitAND($o_DataBaseObject.TableDefs($nCnt).Attributes, -2147483646) = 0 Then ; dbSystemObject -2147483646
			$n_TablesCount = $n_TablesCount + 1
		EndIf
	Next
	Return $n_TablesCount

EndFunc   ;==>_AccessTablesCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _AccessTableExists
; Description ...: Checks if a table is exisits on a database
; Syntax ........: _AccessTableExists($o_Database[, $s_TableName = ""])
; Parameters ....: $o_Database          - database object.
;                  $s_TableName         - [optional] A string value. Default is "". table name
; Return values .: Sucess 	- The table object.,
;					Fail 	- Return 0
; Author ........: Ayman Henry
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AccessTableExists($o_Database, $s_TableName = "")
	Local $n_TableNo = _AccessTablesCount($o_Database, False)
	If $s_TableName = "" Or $n_TableNo = 0 Then Return 0

	For $nCnt = 0 To $n_TableNo - 1
		If $o_Database.Tabledefs($nCnt).Name = $s_TableName Then
			Return $o_Database.Tabledefs($s_TableName)
		EndIf
	Next

	Return 0

EndFunc   ;==>_AccessTableExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _AccessFieldsCount
; Description ...: Return number of Fields on a table
; Syntax ........: _AccessFieldsCount(Byref $o_DataBaseObject, $s_Table)
; Parameters ....: $o_DataBaseObject    - [in/out] Database object.
;                  $s_Table             - A string value.Table name
; Return values .: Sucess		- No. of fields in table
;				   Fail			- Retrn 0
; Author ........: Ayman Henry
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AccessFieldsCount(ByRef $o_DataBaseObject, $s_Table)
	Local $o_TableObject, $n_doc

	If $s_Table = "" Then Return 0
	$o_TableObject = _AccessTableExists($o_DataBaseObject, $s_Table)

	If $o_TableObject = 0 Then Return 0

	If Not IsObj($o_DataBaseObject) Then
		__AccessErrorNotify("Error", "_AccessFieldsCount", "$_AccessStatus_InvalidDataType")
		Return SetError($_AccessStatus_InvalidDataType, 1, 0)
	EndIf

	If Not __AccessIsObjType($o_DataBaseObject, "database") Then
		__AccessErrorNotify("Error", "_AccessFieldsCount", "$_AccessStatus_InvalidObjectType")
		Return SetError($_AccessStatus_InvalidObjectType, 1, 0)
	EndIf

	;$n_doc = $o_DataBaseObject.TableDefs($s_Table).Fields.Count ; db.TableDefs(0).Fields.Count
	$n_doc = $o_TableObject.Fields.Count ; db.TableDefs(0).Fields.Count

	Return $n_doc
EndFunc   ;==>_AccessFieldsCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _AccessFieldExists
; Description ...: Checks if a Field is Exisys in a table
; Syntax ........: _AccessFieldExists(Byref $o_Table[, $s_Field = ""])
; Parameters ....: $o_Table             - [in/out] Table Object.
;                  $s_Field             - [optional] A string value. Default is "". Field Name
; Return values .: Sucess 1, Fail 0
; Author ........: Ayman Henry
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AccessFieldExists(ByRef $o_Table, $s_Field = "")

	Local $n_FieldsNo

	If $s_Field = "" Then Return 0

	If Not IsObj($o_Table) Then
		__AccessErrorNotify("Error", "_AccessFieldExists", "$_AccessStatus_InvalidDataType")
		Return SetError($_AccessStatus_InvalidDataType, 1, 0)
	EndIf

	If Not __AccessIsObjType($o_Table, "_TableDef") Then
		__AccessErrorNotify("Error", "_AccessFieldExists", "$_AccessStatus_InvalidObjectType")
		Return SetError($_AccessStatus_InvalidObjectType, 1, 0)
	EndIf


	$n_FieldsNo = $o_Table.Fields.Count

	For $nCnt = 0 To $n_FieldsNo - 1
		;if $o_Database.TableDefs($s_TableName).Fields($nCnt).Name = $s_Field then Return True
		If $o_Table.Fields($nCnt).Name = $s_Field Then
			Return $o_Table.Fields($nCnt)
		EndIf
	Next

	Return False

EndFunc   ;==>_AccessFieldExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _AccessRecordsCount
; Description ...: Count the number of Records on a table
; Syntax ........: _AccessRecordsCount(Byref $o_Database[, $s_TableName = ""])
; Parameters ....: $o_Database          - [in/out] database object.
;                  $s_TableName         - [optional] A string value. Default is "".Table Name
; Return values .: Sucess	- No of recordes
;					Fail   Return 0
; Author ........: Ayman Henry
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AccessRecordsCount(ByRef $o_Database, $s_TableName = "")
	Local $nErr
	If $s_TableName = "" Then Return 0
	If _AccessTableExists($o_Database, $s_TableName) = 0 Then Return 0
	;-------
	Local $n_RecCount, $o_Rec
	$o_Rec = $o_Database.OpenrecordSet($s_TableName)

	If Not IsObj($o_Rec) Then
		__AccessErrorNotify("Error", "_AccessRecordsCount", "$_AccessStatus_InvalidDataType")
		Return SetError($_AccessStatus_InvalidDataType, 1, 0)
	EndIf

	If Not __AccessIsObjType($o_Rec, "RecordSet") Then
		__AccessErrorNotify("Error", "_AccessRecordsCount", "$_AccessStatus_InvalidObjectType")
		Return SetError($_AccessStatus_InvalidObjectType, 1, 0)
	EndIf



	; Setup internal error handler to Trap COM errors, turn off error notification
	Local $status = __AccessInternalErrorHandlerRegister()
	If Not $status Then __AccessErrorNotify("Warning", "_AccessRecordAdd", _
			"Cannot register internal error handler, cannot trap COM errors", _
			"Use _AccessErrorHandlerRegister() to register a user error handler")
	Local $f_NotifyStatus = _AccessErrorNotify() ; save current error notify status
	$_AccessErrorNotify = False

	$o_Rec.MoveFirst
	$nErr = @error
	$n_RecCount = $o_Rec.RecordCount
	$nErr = @error + $nErr

	$_AccessErrorNotify = True

	; restore error notify and error handler status
	_AccessErrorNotify($f_NotifyStatus) ; restore notification status
	__AccessInternalErrorHandlerDeRegister()

	If $nErr = 0 Then
		Return $n_RecCount
	Else
		Return 0
	EndIf
EndFunc   ;==>_AccessRecordsCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _AccessTablesList
; Description ...: Returns an array has a list of tables on a database
; Syntax ........: _AccessTablesList(Byref $o_DataBaseObject[, $b_OnlyNonSysTable = True])
; Parameters ....: $o_DataBaseObject    - [in/out] Database Object.
;                  $b_OnlyNonSysTable   - [optional] A binary value. Default is True. if true, ignores the system tables
; Return values .: Array with tables name, the first array item has the number of tables.
; Author ........: Ayman Henry
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AccessTablesList(ByRef $o_DataBaseObject, $b_OnlyNonSysTable = True)
	Local $avArray[1], $n_Count = 0
	Local $n_TablesCount = _AccessTablesCount($o_DataBaseObject, False)
	$avArray[0] = $n_TablesCount
	If $n_TablesCount = 0 Then Return $avArray

	For $nCnt = 0 To $n_TablesCount - 1
		If $b_OnlyNonSysTable = True Then
			If BitAND($o_DataBaseObject.TableDefs($nCnt).Attributes, -2147483646) = 0 Then ; dbSystemObject -2147483646
				$n_Count = $n_Count + 1
				_ArrayAdd($avArray, $o_DataBaseObject.TableDefs($nCnt).Name)
			EndIf
		Else
			$n_Count = $n_Count + 1
			_ArrayAdd($avArray, $o_DataBaseObject.TableDefs($nCnt).Name)
		EndIf
	Next
	$avArray[0] = $n_Count
	Return $avArray

EndFunc   ;==>_AccessTablesList

; #FUNCTION# ====================================================================================================================
; Name ..........: _AccessFieldsList
; Description ...: Returns an array has a list of fields on a database
; Syntax ........: _AccessFieldsList(Byref $o_DataBaseObject, $s_Table)
; Parameters ....: $o_DataBaseObject    - [in/out] database name.
;                  $s_Table             - A string value. table object.
; Return values .: Array with files name, the first array item has the number of fields.
; Author ........: Ayman Henry
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AccessFieldsList(ByRef $o_DataBaseObject, $s_Table)
	Local $avArray[1]
	Local $n_FieldsCount = _AccessFieldsCount($o_DataBaseObject, $s_Table)
	$avArray[0] = $n_FieldsCount

	If $n_FieldsCount = 0 Then Return $avArray

	For $nCnt = 0 To $n_FieldsCount - 1
		;_ArrayAdd ($avArray, $o_Table.Fields($nCnt).Name )
		_ArrayAdd($avArray, $o_DataBaseObject.TableDefs($s_Table).Fields($nCnt).Name)
	Next

	Return $avArray

EndFunc   ;==>_AccessFieldsList

; #FUNCTION# ====================================================================================================================
; Name ..........: _AccessRecordList
; Description ...: Returns an array has a list of record fields value.
; Syntax ........: _AccessRecordList(Byref $o_DataBaseObject, $s_Table, $nRecNo)
; Parameters ....: $o_DataBaseObject    - [in/out] Database object.
;                  $s_Table             - A string value. Table Name
;                  $nRecNo              - An integer number value. The abslout position for the required fields.
; Return values .: Array with one record data, the first array item has the number of fields.
; Author ........: Ayman Henry
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AccessRecordList(ByRef $o_DataBaseObject, $s_Table, $nRecNo)
	Local $avArray[1], $o_Rec
	Local $n_FieldsCount = _AccessFieldsCount($o_DataBaseObject, $s_Table)
	$avArray[0] = 0
	If $n_FieldsCount = 0 Then Return $avArray
	If Not IsInt($nRecNo) Then Return $avArray
	If $nRecNo < 1 Then Return $avArray

	$o_Rec = $o_DataBaseObject.OpenRecordSet($s_Table)

	If Not IsObj($o_Rec) Then
		__AccessErrorNotify("Error", "_AccessRecordList", "$_AccessStatus_InvalidDataType")
		Return SetError($_AccessStatus_InvalidDataType, 1, 0)
	EndIf

	If Not __AccessIsObjType($o_Rec, "RecordSet") Then
		__AccessErrorNotify("Error", "_AccessRecordList", "$_AccessStatus_InvalidObjectType")
		Return SetError($_AccessStatus_InvalidObjectType, 1, 0)
	EndIf
	;-------
	$o_Rec.MoveFirst
	$o_Rec.Move($nRecNo - 1)

	If $o_Rec.Eof = -1 Or $o_Rec.Bof = -1 Then
		$o_Rec.close
		Return $avArray
	EndIf

	$avArray[0] = $n_FieldsCount
	For $nCnt = 0 To $n_FieldsCount - 1
		;_ArrayAdd ($avArray, $o_Table.Fields($nCnt).Name )
		_ArrayAdd($avArray, $o_Rec.Fields($nCnt).Value)
	Next

	$o_Rec.close
	Return $avArray

EndFunc   ;==>_AccessRecordList

; #FUNCTION# ====================================================================================================================
; Name ..........: _AccessRecordAdd
; Description ...: Add a record to a table
; Syntax ........: _AccessRecordAdd(Byref $o_DataBaseObject, $s_Table, $av_Fields)
; Parameters ....: $o_DataBaseObject    - [in/out] Database object.
;                  $s_Table             - A string value. table name to add a record to
;                  $av_Fields           - An array of variants. Two dim. array, col.1 has the field name col.2 the value.
; Return values .: Success	 1
;				   Falis any other value.
; Author ........: Ayman Henry
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AccessRecordAdd(ByRef $o_DataBaseObject, $s_Table, $av_Fields)
	Return __AccessRecordAddEdit($o_DataBaseObject, $s_Table, $av_Fields)
EndFunc   ;==>_AccessRecordAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _AccessRecordEdit
; Description ...: Edit a record to a table
; Syntax ........: _AccessRecordEdit(Byref $o_DataBaseObject, $s_Table, $av_Fields, $nRecNo)
; Parameters ....: $o_DataBaseObject    - [in/out] Database object.
;                  $s_Table             - A string value. table name to edit.
;                  $av_Fields           - An array of variants.An array of variants. Two dim. array, col.1 has the field name col.2 the value.
;                  $nRecNo              - An integer number value. the absoulte position of the record to be edit
; Return values .: Success	 1
;				   Falis any other value.
; Author ........: Ayman Henry
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AccessRecordEdit(ByRef $o_DataBaseObject, $s_Table, $av_Fields, $nRecNo)
	Return __AccessRecordAddEdit($o_DataBaseObject, $s_Table, $av_Fields, $nRecNo)
EndFunc   ;==>_AccessRecordEdit

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __AccessRecordAddEdit
; Description ...: Add or Edit a record if $nRecNo is given i.e it is in Edit mode
; Syntax ........: __AccessRecordAddEdit(Byref $o_DataBaseObject, $s_Table, $av_Fields[, $nRecNo = Default])
; Parameters ....: $o_DataBaseObject    - [in/out] Database object.
;                  $s_Table             - A string value. table name to edit.
;                  $av_Fields           - An array of variants.An array of variants. Two dim. array, col.1 has the field name col.2 the value.
;                  $nRecNo              - An integer number value. the absoulte position of the record to be edit
; Return values .: Success	 1
;				   Falis any other value.
; Author ........: Ayaman Henry
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __AccessRecordAddEdit(ByRef $o_DataBaseObject, $s_Table, $av_Fields, $nRecNo = Default)

	If $s_Table = "" Then Return -1
	If Not IsArray($av_Fields) Then Return -2
	If UBound($av_Fields, 0) <> 2 Then Return -3

	If $nRecNo <> Default Then
		If Not IsInt($nRecNo) Then Return 0
		If $nRecNo < 1 Then Return 0
	EndIf
	;------------
	Local $o_Rec, $b_Err = False
	Local $o_Table = _AccessTableExists($o_DataBaseObject, $s_Table)
	If $o_Table = 0 Then Return -4

	;--------
	For $nCnt = 0 To UBound($av_Fields, 0) - 1
		If 0 = _AccessFieldExists($o_Table, $av_Fields[$nCnt][0]) Then Return -5
	Next
	;--------
	$o_Rec = $o_DataBaseObject.OpenRecordSet($s_Table)

	If Not IsObj($o_Rec) Then
		__AccessErrorNotify("Error", "_AccessRecordList", "$_AccessStatus_InvalidDataType")
		Return SetError($_AccessStatus_InvalidDataType, 1, 0)
	EndIf

	If Not __AccessIsObjType($o_Rec, "RecordSet") Then
		__AccessErrorNotify("Error", "_AccessRecordList", "$_AccessStatus_InvalidObjectType")
		Return SetError($_AccessStatus_InvalidObjectType, 1, 0)
	EndIf
	;-------

	; Setup internal error handler to Trap COM errors, turn off error notification
	Local $status = __AccessInternalErrorHandlerRegister()
	If Not $status Then __AccessErrorNotify("Warning", "_AccessRecordAdd", _
			"Cannot register internal error handler, cannot trap COM errors", _
			"Use _AccessErrorHandlerRegister() to register a user error handler")
	Local $f_NotifyStatus = _AccessErrorNotify() ; save current error notify status
	_AccessErrorNotify(False)

	If $nRecNo = Default Then
		$o_Rec.AddNew
	Else

		$o_Rec.MoveFirst
		If @error <> 0 Then $b_Err = True

		$o_Rec.Move($nRecNo - 1)
		If @error <> 0 Then $b_Err = True

		$o_Rec.Edit
		If @error <> 0 Then $b_Err = True
	EndIf

	If $b_Err = False Then

		For $nCnt = 0 To UBound($av_Fields, 1) - 1
			$o_Rec.Fields($av_Fields[$nCnt][0]).Value = $av_Fields[$nCnt][1]

			;MsgBox (0,0, $o_Rec.Fields( $av_Fields[$nCnt][0] ).Name & "  Value=" & $av_Fields[$nCnt][1] )

			If @error <> 0 Then ;$_AccessStatus_ComError
				;MsgBox (0,0, $o_Rec.Fields( $av_Fields[$nCnt][0] ).Name & "  @error=" & @error )

				$b_Err = True
				ExitLoop
			EndIf
		Next
	EndIf

	; restore error notify and error handler status
	_AccessErrorNotify($f_NotifyStatus) ; restore notification status
	__AccessInternalErrorHandlerDeRegister()

	If $b_Err = True Then
		$o_Rec.CancelUpdate
		__AccessErrorNotify("Error", "_AccessRecordAdd", "$_AccessStatus_DataTypeMisMatch, Field =" & $o_Rec.Fields($av_Fields[$nCnt][0]).Name)
		Return SetError($_AccessStatus_DataTypeMisMatch, 1, 0)
	Else
		$o_Rec.Update
	EndIf

	$o_Rec.close
	Return 1

EndFunc   ;==>__AccessRecordAddEdit

; #FUNCTION# ====================================================================================================================
; Name ..........: _AccessRecordDelete
; Description ...: Delete a record
; Syntax ........: _AccessRecordDelete(Byref $o_DataBaseObject, $s_Table, $nRecNo)
; Parameters ....: $o_DataBaseObject    - [in/out] database object.
;                  $s_Table             - A string value. table name
;                  $nRecNo              - An integer number value. Absoult position of record.
; Return values .: None
; Author ........: Ayman Henry
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AccessRecordDelete(ByRef $o_DataBaseObject, $s_Table, $nRecNo)
	Local $o_Rec, $b_Err = 0
	Local $n_FieldsCount = _AccessFieldsCount($o_DataBaseObject, $s_Table)

	If $n_FieldsCount = 0 Then Return 0
	If Not IsInt($nRecNo) Then Return 0
	If $nRecNo < 1 Then Return 0

	$o_Rec = $o_DataBaseObject.OpenRecordSet($s_Table)

	If Not IsObj($o_Rec) Then
		__AccessErrorNotify("Error", "_AccessRecordList", "$_AccessStatus_InvalidDataType")
		Return SetError($_AccessStatus_InvalidDataType, 1, 0)
	EndIf

	If Not __AccessIsObjType($o_Rec, "RecordSet") Then
		__AccessErrorNotify("Error", "_AccessRecordList", "$_AccessStatus_InvalidObjectType")
		Return SetError($_AccessStatus_InvalidObjectType, 1, 0)
	EndIf
	;-------
	$o_Rec.MoveFirst
	$o_Rec.Move($nRecNo - 1)

	If $o_Rec.Eof = -1 Or $o_Rec.Bof = -1 Then
		$o_Rec.close
		Return 0
	EndIf

	; Setup internal error handler to Trap COM errors, turn off error notification
	Local $status = __AccessInternalErrorHandlerRegister()
	If Not $status Then __AccessErrorNotify("Warning", "_AccessRecordAdd", _
			"Cannot register internal error handler, cannot trap COM errors", _
			"Use _AccessErrorHandlerRegister() to register a user error handler")
	Local $f_NotifyStatus = _AccessErrorNotify() ; save current error notify status

	$o_Rec.Delete
	If @error <> 0 Then ;$_AccessStatus_ComError
		$b_Err = True
	EndIf

	; restore error notify and error handler status
	_AccessErrorNotify($f_NotifyStatus) ; restore notification status
	__AccessInternalErrorHandlerDeRegister()
	$o_Rec.close

	If $b_Err = 1 Then Return 0
	Return 1

EndFunc   ;==>_AccessRecordDelete

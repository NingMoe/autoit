#Tidy_Parameters= /gd 1  /gds 1 /nsdp
#AutoIt3Wrapper_Au3Check_Parameters= -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Au3Check_Stop_OnWarning=Y
#include-once
#include <Array.au3>
#include <Date.au3>
#include <File.au3>
#include <OutlookExConstants.au3>
; #INDEX# =======================================================================================================================
; Title .........: Microsoft Outlook Function Library (MS Outlook 2003 and later)
; AutoIt Version : 3.1.1 (AutoIt3 with COM support)
; UDF Version ...: 0.7.1.1
; Language ......: English
; Description ...: A collection of functions for accessing and manipulating Microsoft Outlook
; Author(s) .....: wooltown, water
; Modified.......: 20120419 (YYYYMMDD)
; Contributors ..: progandy (CSV functions taken and modified from http://www.autoitscript.com/forum/topic/114406-csv-file-to-multidimensional-array)
;                  Ultima, PsaltyDS for the basis of the _OL_ArrayConcatenate function
; Resources .....: Outlook 2003 Visual Basic Reference: http://msdn.microsoft.com/en-us/library/aa271384(v=office.11).aspx
;                  Outlook 2007 Developer Reference:    http://msdn.microsoft.com/en-us/library/bb177050(v=office.12).aspx
;                  Outlook 2010 Developer Reference:    http://msdn.microsoft.com/en-us/library/ff870432.aspx
;                  Outlook Examples:                    http://www.vboffice.net/sample.html?cmd=list&mnu=2
; ===============================================================================================================================

#region #VARIABLES#
; #VARIABLES# ===================================================================================================================
Global $iOL_Debug = 0 ; Debug level. 0 = no debug information, 1 = to console, 2 = to MsgBox, 3 = into File
Global $sOL_DebugFile = @ScriptDir & "\OutlookDebug.txt" ; Debug file if $iOL_Debug is set to 3
Global $oOL_Error ; COM Error handler
Global $bOL_AlreadyRunning = False ; Is Outlook already running?
; ===============================================================================================================================
#endregion #VARIABLES#

; #CURRENT# =====================================================================================================================
;_OL_Open
;_OL_Close
;_OL_AccountGet
;_OL_AddressListGet
;_OL_AddressListMemberGet
;_OL_ApplicationGet
;_OL_BarGroupAdd
;_OL_BarGroupDelete
;_OL_BarGroupGet
;_OL_BarShortcutAdd
;_OL_BarShortcutDelete
;_OL_BarShortcutGet
;_OL_CategoryAdd
;_OL_CategoryDelete
;_OL_CategoryGet
;_OL_DistListMemberAdd
;_OL_DistListMemberDelete
;_OL_DistListMemberGet
;_OL_FolderAccess
;_OL_FolderArchiveGet
;_OL_FolderArchiveSet
;_OL_FolderCopy
;_OL_FolderCreate
;_OL_FolderDelete
;_OL_FolderExists
;_OL_FolderFind
;_OL_FolderGet
;_OL_FolderModify
;_OL_FolderMove
;_OL_FolderRename
;_OL_FolderSelectionGet
;_OL_FolderSet
;_OL_FolderTree
;_OL_Item2Task
;_OL_ItemAttachmentAdd
;_OL_ItemAttachmentDelete
;_OL_ItemAttachmentGet
;_OL_ItemAttachmentSave
;_OL_ItemConflictGet
;_OL_ItemCopy
;_OL_ItemCreate
;_OL_ItemDelete
;_OL_ItemDisplay
;_OL_ItemExport
;_OL_ItemFind
;_OL_ItemForward
;_OL_ItemGet
;_OL_ItemImport
;_OL_ItemModify
;_OL_ItemMove
;_OL_ItemPrint
;_OL_ItemRecipientAdd
;_OL_ItemRecipientDelete
;_OL_ItemRecipientGet
;_OL_ItemRecurrenceDelete
;_OL_ItemRecurrenceExceptionGet
;_OL_ItemRecurrenceExceptionSet
;_OL_ItemRecurrenceGet
;_OL_ItemRecurrenceSet
;_OL_ItemReply
;_OL_ItemSave
;_OL_ItemSend
;_OL_ItemSendReceive
;_OL_ItemSync
;_OL_MailheaderGet
;_OL_MaiLSignatureCreate
;_OL_MaiLSignatureDelete
;_OL_MaiLSignatureGet
;_OL_MaiLSignatureSet
;_OL_OOFGet
;_OL_OOFSet
;_OL_PSTAccess
;_OL_PSTClose
;_OL_PSTCreate
;_OL_PSTGet
;_OL_RecipientFreeBusyGet
;_OL_ReminderDelay
;_OL_ReminderDismiss
;_OL_ReminderGet
;_OL_RuleActionGet
;_OL_RuleActionSet
;_OL_RuleAdd
;_OL_RuleConditionGet
;_OL_RuleConditionSet
;_OL_RuleDelete
;_OL_RuleExecute
;_OL_RuleGet
;_OL_StoreGet
;_OL_VersionInfo
;_OL_Wrapper_CreateAppointment
;_OL_Wrapper_SendMail
; ===============================================================================================================================

; #INTERNAL_USE_ONLY#============================================================================================================
;_OL_ArrayConcatenate
;_OL_CheckProperties
;_OL_COMError
;_ReadCSV
;_WriteCSV
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name...........: _OL_Open
; Description ...: Opens a connection to Microsoft Outlook.
; Syntax.........: _OL_Open([$bOL_WarningClick = False[, $sOL_WarningProgram = ""[, $iOL_WinCheckTime = 1000[, $iOL_CtrlCheckTime = 1000[, $sOL_ProfileName = ""[,  $sOL_Password = ""]]]]])
; Parameters ....: $bOL_WarningClick   - Optional: If set to True a function will click away the Outlook security warning messages (default = False)
;                  $sL_WarningProgram  - Optional: Complete path to the WarningProgram, e.g.: "C:\OLext\_OL_Warnings.exe"
;                  + (default = @ScriptDir & "\_OL_Warnings.exe" which is part of this UDF)
;                  $iOL_WinCheckTime   - Optional: Time in milliseconds to wait before a check for a warning window is performed (default = 1000 milliseconds = 1 second)
;                  $iOL_CtrlCheckTime  - Optional: How long, in milliseconds, we will wait before we check that the controls we click are enabled (default = 1000)
;                  $sOL_ProfileName    - Optional: Name of the profile to be used for logon. The default profile will be used if none is specified (default = "")
;                  $sOL_Password       - Optional: The password associated with the profile (default = "")
; Return values .: Success - New object identifier, sets @extended to:
;                  |0 - No COM error handler has been initialized for this UDF because another COM error handler was already active
;                  |1 - A COM error handler has been initialized for this UDF
;                  Failure - Returns 0 and sets @error:
;                  |1 - Unable to create Outlook Object. See @extended for details (@error returned by ObjCreate)
;                  |2 - Unable to create custom error handler. See @extended for details (@error returned by ObjEvent)
;                  |3 - $bOL_WarningClick is invalid. Has to be True or False
;                  |4 - $iOL_WinCheckTime is invalid. Has to be an integer
;                  |5 - File $sOL_WarningProgram does not exist. Has to be an executable that can be startet using the run command
;                  |6 - Run error when running $sOL_WarningProgram. Please check @extended for more information
;                  |7 - $iOL_CtrlCheckTime is invalid. Has to be an integer
;                  |8 - Unable to logon to the default profile. See @extended for details (@error returned by Logon)
;                  |9 - No profile has been configured or there is no default profile for Outlook
;                  |10 - You specified a profile name to logon to but Outlook is already running
; Author ........: Wooltown
; Modified ......: water
; Remarks .......: Create your own $sOL_WarningProgram Exe if the window title or messages have changed or are displayed in another language.
;                  If Outlook is already running the global variable $bOL_AlreadyRunning is set to True.
;                  If Outlook is already running a specified profile name is ignored because no logon is needed.
;+
;                  The COM error handler will be initialized only if there doesn't already exist another error handler.
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _OL_Open($bOL_WarningClick = False, $sOL_WarningProgram = "", $iOL_WinCheckTime = 1000, $iOL_CtrlCheckTime = 1000, $sOL_ProfileName = "", $sOL_Password = "")

	Local $iOL_ErrorHandler = 0
	Local $oOL = ObjGet("", "Outlook.Application")
	If $oOL <> 0 Then $bOL_AlreadyRunning = True
	If Not IsBool($bOL_WarningClick) Then Return SetError(3, 0, 0)
	If Not IsInt($iOL_WinCheckTime) Then Return SetError(4, 0, 0)
	If Not IsInt($iOL_CtrlCheckTime) Then Return SetError(7, 0, 0)
	If $bOL_AlreadyRunning And $sOL_ProfileName <> "" Then Return SetError(10, 0, 0)
	If Not $bOL_AlreadyRunning Then
		$oOL = ObjCreate("Outlook.Application")
		If @error Or Not IsObj($oOL) Then Return SetError(1, @error, 0)
	EndIf
	If $bOL_WarningClick Then
		If StringStripWS($sOL_WarningProgram, 3) = "" Then $sOL_WarningProgram = @ScriptDir & "\_OL_Warnings.exe"
		If Not FileExists($sOL_WarningProgram) Then Return SetError(5, 0, 0)
		Run($sOL_WarningProgram & " " & @AutoItPID & " " & $iOL_WinCheckTime & " " & $iOL_CtrlCheckTime & " " & $oOL.Version & " " & $oOL.LanguageSettings.LanguageID($msoLanguageIDUI), "", @SW_HIDE)
		If @error Then Return SetError(6, @error, 0)
	EndIf
	; A COM error handler will be initialised only if one does not exist
	If ObjEvent("AutoIt.Error") = "" Then
		$oOL_Error = ObjEvent("AutoIt.Error", "_OL_COMError") ; Creates a custom COM error handler
		If @error Then Return SetError(2, @error, 0)
		$iOL_ErrorHandler = 1
	EndIf
	Local $aVersion = StringSplit($oOL.Version, '.')
	; Logon to the specified or the default profile if Outlook wasn't already running (for Outlook 2007 and later)
	If Not $bOL_AlreadyRunning And Int($aVersion[1]) >= 12 Then
		If StringStripWS($sOL_ProfileName, 3) = "" Then $sOL_ProfileName = $oOL.DefaultProfileName
		If StringStripWS($sOL_ProfileName, 3) = "" Then Return SetError(9, 0, 0)
		Local $oOL_Namespace = $oOL.GetNamespace("MAPI")
		$oOL_Namespace.Logon($sOL_ProfileName, $sOL_Password, False, False)
		If @error Then Return SetError(8, @error, 0)
	EndIf
	Return SetError(0, $iOL_ErrorHandler, $oOL)

EndFunc   ;==>_OL_Open

; #FUNCTION# ====================================================================================================================
; Name...........: _OL_Close
; Description ...: Close the connection to Microsoft Outlook.
; Syntax.........: _OL_Close($oOL[, $bOL_ForceClose = False])
; Parameters ....: $oOL            - Outlook object returned by a preceding call to _OL_Open()
;                  $bOL_ForceClose - Optional: If True Outlook is closed even when it was running at _OL_Open time (default = False)
; Return values .: Success - Returns 1 and sets @extended to 1 if Outlook was already running
;                  Failure - Returns 0 and sets @error
;                  |1 - Error with Session.Logoff. Please check @extended for more information
;                  |2 - Error with Application.Quit. Please check @extended for more information
;                  |3 - Error with Quit. Please check @extended for more information
; Author ........: Wooltown
; Modified ......: water
; Remarks .......: If Outlook was already running when _OL_Open was called you have to use flag $bOL_ForceClose to close Outlook.
;                  If Outlook was already running when _OL_Open was called @extended is set to 1 to indicate this
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _OL_Close(ByRef $oOL, $bOL_ForceClose = False)

	If Not $bOL_AlreadyRunning Or $bOL_ForceClose Then
		$oOL.Session.Logoff
		If @error Then Return SetError(1, @error, 0)
		$oOL.Application.Quit
		If @error Then Return SetError(2, @error, 0)
		; The associated Outlook session is closed completely.
		; The user is logged out of the messaging system and any changes to items not already saved are discarded
		$oOL.Quit
		If @error Then Return SetError(3, @error, 0)
	EndIf
	$oOL = 0
	$iOL_Debug = 0
	$sOL_DebugFile = @ScriptDir & "\OutlookDebug.txt"
	$oOL_Error = 0
	If $bOL_AlreadyRunning Then
		$bOL_AlreadyRunning = False
		Return SetError(0, 1, 1)
	EndIf
	$bOL_AlreadyRunning = False
	Return 1

EndFunc   ;==>_OL_Close

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_AccountGet
; Description ...: Get information about the accounts available for the current profile.
; Syntax.........: _OL_AccountGet($oOL)
; Parameters ....: $oOL - Outlook object returned by a preceding call to _OL_Open()
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - AccountType
;                  |1 - Displayname
;                  |2 - SMTPAddress
;                  |3 - Username
;                  Failure - Returns "" and sets @error:
;                  |1 - Function is only supported for Outlook 2007 and later
; Author ........: water
; Modified ......:
; Remarks .......: This function only works for Outlook 2007 and later.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_AccountGet($oOL)

	Local $aVersion = StringSplit($oOL.Version, '.')
	If Int($aVersion[1]) < 12 Then Return SetError(1, 0, "")
	Local $iOL_Index = 0
	Local $aOL_Account[$oOL.Session.Accounts.Count + 1][4] = [[$oOL.Session.Accounts.Count, 4]]
	For $oOL_Account In $oOL.Session.Accounts
		$iOL_Index = $iOL_Index + 1
		$aOL_Account[$iOL_Index][0] = $oOL_Account.AccountType
		$aOL_Account[$iOL_Index][1] = $oOL_Account.DisplayName
		$aOL_Account[$iOL_Index][2] = $oOL_Account.SMTPAddress
		$aOL_Account[$iOL_Index][3] = $oOL_Account.UserName
	Next
	Return $aOL_Account

EndFunc   ;==>_OL_AccountGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_AddressListGet
; Description ...: Returns all Addresslists.
; Syntax.........: _OL_AddressListGet($oOL[, $bOL_Resolve = True])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $bOL_Resolve - Optional: If True only addresslists that are used when resolving recipient names are returned (default = True)
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - Constant from the OlAddressListType enumeration representing the type of the Addresslist
;                  |1 - Display name for the object
;                  |2 - Index indicating the position of the AddressList within the collection
;                  |3 - Integer that represents the order of this Addresslist to be used when resolving recipient names
;                  +    -1 means the Addresslist is not used to resolve addresses
;                  |4 - A string representing the unique identifier for the addresslist
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_AddressListGet($oOL, $bOL_Resolve = True)

	Local $iOL_Index = 1, $iOL_Index1, $oOL_AddressList
	Local $aOL_AddressLists[$oOL.Session.AddressLists.Count + 1][5]
	For $iOL_Index1 = 1 To $oOL.Session.AddressLists.Count
		$oOL_AddressList = $oOL.Session.AddressLists($iOL_Index1)
		If $bOL_Resolve = False Or $oOL_AddressList.ResolutionOrder <> -1 Then
			$aOL_AddressLists[$iOL_Index][0] = $oOL_AddressList.AddressListType
			$aOL_AddressLists[$iOL_Index][1] = $oOL_AddressList.Name
			$aOL_AddressLists[$iOL_Index][2] = $iOL_Index1
			$aOL_AddressLists[$iOL_Index][3] = $oOL_AddressList.ResolutionOrder
			$aOL_AddressLists[$iOL_Index][4] = $oOL_AddressList.ID
			$iOL_Index += 1
		EndIf
	Next
	ReDim $aOL_AddressLists[$iOL_Index][UBound($aOL_AddressLists, 2)]
	$aOL_AddressLists[0][0] = $iOL_Index - 1
	$aOL_AddressLists[0][1] = UBound($aOL_AddressLists, 2)
	Return $aOL_AddressLists

EndFunc   ;==>_OL_AddressListGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_AddressListMemberGet
; Description ...: Gets all members of an address list.
; Syntax.........: _OL_AddressListMemberGet($oOL, $vOL_ID)
; Parameters ....: $oOL    - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_ID - Number or name of an address list in the address lists collection as returned by _OL_AddressListGet
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - E-mail address of the AddressEntry
;                  |1 - Display name for the AddressEntry
;                  |2 - Constant from the OlAddressEntryUserType enumeration representing the user type of the AddressEntry
;                  |3 - Unique identifier for the object (string)
;                  |4 - Object of the AddressEntry
;                  Failure - Returns "" and sets @error:
;                  |1 - No address list index specified
;                  |2 - Address list specified by $vOL_ID could not be found
; Author ........: water
; Modified.......:
; Remarks .......: To access an AddressList by number please use the Index returned by _OL_AddressListGet in column 3
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_AddressListMemberGet($oOL, $vOL_ID)

	If StringStripWS($vOL_ID, 3) = "" Then Return SetError(1, 0, "")
	Local $oOL_Items = $oOL.Session.AddressLists.Item($vOL_ID).AddressEntries
	If @error Then Return SetError(2, @error, 0)
	Local $aOL_Members[$oOL_Items.Count + 1][5] = [[$oOL_Items.Count, 5]], $iOL_Index = 1
	For $oOL_Item In $oOL_Items
		$aOL_Members[$iOL_Index][0] = $oOL_Item.Address ; <== ??
		$aOL_Members[$iOL_Index][1] = $oOL_Item.Name
		$aOL_Members[$iOL_Index][2] = $oOL_Item.AddressEntryUserType
		$aOL_Members[$iOL_Index][3] = $oOL_Item.ID
		; Exchange user that belongs to the same or a different Exchange forest
		If $oOL_Item.AddressEntryUserType = $olExchangeUserAddressEntry Or $oOL_Item.AddressEntryUserType = $olExchangeRemoteUserAddressEntry Then
			$aOL_Members[$iOL_Index][4] = $oOL_Item.GetExchangeUser
			$aOL_Members[$iOL_Index][0] = $aOL_Members[$iOL_Index][4] .PrimarySmtpAddress
		EndIf
		; Address entry in an Outlook Contacts folder
		If $oOL_Item.AddressEntryUserType = $olOutlookContactAddressEntry Then
			$aOL_Members[$iOL_Index][4] = $oOL_Item.GetContact
			$aOL_Members[$iOL_Index][0] = $aOL_Members[$iOL_Index][4] .Email1Address
		EndIf
		$iOL_Index += 1
	Next
	Return $aOL_Members

EndFunc   ;==>_OL_AddressListMemberGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ApplicationGet
; Description ...: Get information about the Outlook application.
; Syntax.........: _OL_ApplicationGet($oOL)
; Parameters ....: $oOL - Outlook object returned by a preceding call to _OL_Open()
; Return values .: Success - one-dimensional one based array with the following information:
;                  |1 - Name of the default profile
;                  |2 - LanguageSettings: Execution mode language
;                  |3 - LanguageSettings: Help language
;                  |4 - LanguageSettings: Install language
;                  |5 - LanguageSettings: User interface language
;                  |6 - Name of the application
;                  |7 - Product code. String specifying the Microsoft Outlook globally unique identifier (GUID
;                  |8 - Product version (n.n.n.n)
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ApplicationGet($oOL)

	Local $aOL_Application[9] = [8]
	$aOL_Application[1] = $oOL.DefaultProfileName
	$aOL_Application[2] = $oOL.LanguageSettings.LanguageID($msoLanguageIDExeMode)
	$aOL_Application[3] = $oOL.LanguageSettings.LanguageID($msoLanguageIDHelp)
	$aOL_Application[4] = $oOL.LanguageSettings.LanguageID($msoLanguageIDInstall)
	$aOL_Application[5] = $oOL.LanguageSettings.LanguageID($msoLanguageIDUI)
	$aOL_Application[6] = $oOL.Name
	$aOL_Application[7] = $oOL.ProductCode
	$aOL_Application[8] = $oOL.Version
	Return $aOL_Application

EndFunc   ;==>_OL_ApplicationGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_BarGroupAdd
; Description ...: Add a group to the OutlookBar.
; Syntax.........: _OL_BarGroupAdd($oOL, $sOL_Groupname[, $iOL_Position = 1])
; Parameters ....: $oOL           - Outlook object returned by a preceding call to _OL_Open()
;                  $oOL_Groupname - Name of the group to be created
;                  $iOL_Position  - Optional: Position at which the new group will be inserted in the Shortcuts pane (default = 1 = at the top of the bar)
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - Error accessing the Outlookbar pane. For details please see @extended
;                  |2 - Error creating the group. For details please see @extended
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_BarGroupAdd($oOL, $sOL_Groupname, $iOL_Position = 1)

	Local $oOL_Pane = $oOL.ActiveExplorer.Panes("OutlookBar")
	If @error Then Return SetError(1, @error, 0)
	$oOL_Pane.Contents.Groups.Add($sOL_Groupname, $iOL_Position)
	If @error Then Return SetError(2, @error, 0)
	Return 1

EndFunc   ;==>_OL_BarGroupAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_BarGroupDelete
; Description ...: Deletes a group from the OutlookBar.
; Syntax.........: _OL_BarGroupDelete($oOL, $vOL_Groupname)
; Parameters ....: $oOL           - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Groupname - Name or 1-based index value of the group to be deleted
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $vOL_Groupname is empty
;                  |2 - Error accessing the specified group. For details please see @extended
;                  |3 - Error removing the specified group. For details please see @extended
; Author ........: water
; Modified ......:
; Remarks .......: To delete a group by name isn't possible in Outlook 2002
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_BarGroupDelete($oOL, $vOL_Groupname)

	If StringStripWS($vOL_Groupname, 3) = "" Then Return SetError(1, 0, 0)
	Local $oOL_Groups = $oOL.ActiveExplorer.Panes.Item("OutlookBar").Contents.Groups
	If @error Then Return SetError(2, @error, 0)
	$oOL_Groups.Remove($vOL_Groupname)
	If @error Then Return SetError(3, @error, 0)
	Return 1

EndFunc   ;==>_OL_BarGroupDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_BarGroupGet
; Description ...: Returns all groups in the OutlookBar.
; Syntax.........: _OL_BarGroupGet($oOL)
; Parameters ....: $oOL    - Outlook object returned by a preceding call to _OL_Open()
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - Display name of the group
;                  |1 - $OlOutlookBarViewType constant representing the view type of the group
;                  Failure - Returns "" and sets @error:
;                  |1 - Error accessing the Outlookbar pane. For details please see @extended
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_BarGroupGet($oOL)

	Local $iOL_Index = 1
	Local $oOL_Pane = $oOL.ActiveExplorer.Panes("OutlookBar")
	If @error Then Return SetError(1, @error, "")
	Local $aOL_Groups[$oOL_Pane.Contents.Groups.Count + 1][2]
	For $oOL_Group In $oOL_Pane.Contents.Groups
		$aOL_Groups[$iOL_Index][0] = $oOL_Group.Name
		$aOL_Groups[$iOL_Index][1] = $oOL_Group.ViewType
		$iOL_Index += 1
	Next
	$aOL_Groups[0][0] = $iOL_Index - 1
	$aOL_Groups[0][1] = UBound($aOL_Groups, 2)
	Return $aOL_Groups

EndFunc   ;==>_OL_BarGroupGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_BarShortcutAdd
; Description ...: Add a shortcut to a group in the OutlookBar.
; Syntax.........: _OL_BarShortcutAdd($oOL, $vOL_Groupname, $sOL_Shortcutname, $sOL_Target[, $iOL_Position = 1[, $sOL_Icon = ""]])
; Parameters ....: $oOL              - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Groupname    - Name or 1-based index value of the group where the shortcut will be created
;                  $oOL_Shortcutname - Name of the shortcut to be created
;                  $sOL_Target       - Target of the shortcut being created. Can be an Outlook folder, filesystem folder, filesystem path or URL
;                  $iOL_Position     - Optional: Position at which the new shortcut will be inserted in the group (default = 0 = first shortcut in an empty group)
;                  $sOL_Icon         - Optional: The path of the icon file e.g. C:\temp\sample.ico (default = no icon)
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - Error accessing the group specified by $vOL_Groupname. For details please see @extended
;                  |2 - Error creating the shortcut. For details please see @extended
;                  |3 - Specified icon file could not be found
;                  |4 - Error setting the icon for the created shortcut. For details please see @extended
; Author ........: water
; Modified ......:
; Remarks .......: Specify $iOL_Position = 1 to position a shortcut at the top of a non empty group
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_BarShortcutAdd($oOL, $vOL_Groupname, $sOL_Shortcutname, $sOL_Target, $iOL_Position = 0, $sOL_Icon = "")

	Local $oOL_Group = $oOL.ActiveExplorer.Panes.Item("OutlookBar").Contents.Groups.Item($vOL_Groupname)
	If @error Or Not IsObj($oOL_Group) Then Return SetError(1, @error, 0)
	Local $oOL_ShortCut = $oOL_Group.Shortcuts.Add($sOL_Target, $sOL_Shortcutname, $iOL_Position)
	If @error Then Return SetError(2, @error, 0)
	If $sOL_Icon <> "" Then
		If Not FileExists($sOL_Icon) Then Return SetError(3, 0, 0)
		$oOL_ShortCut.SetIcon($sOL_Icon)
		If @error Then Return SetError(4, @error, 0)
	EndIf
	Return 1

EndFunc   ;==>_OL_BarShortcutAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_BarShortcutDelete
; Description ...: Deletes a Shortcut from the OutlookBar.
; Syntax.........: _OL_BarShortcutDelete($oOL, $vOL_Groupname, $vOL_Shortcutname)
; Parameters ....: $oOL              - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Groupname    - Name or 1-based index value of the group from where the shortcut will be deleted
;                  $vOL_Shortcutname - Name or 1-based index value of the shortcut to be deleted
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $vOL_Shortcutname is empty
;                  |2 - $vOL_Groupname is empty
;                  |3 - Error accessing the specified group. For details please see @extended
;                  |4 - Error removing the specified Shortcut. For details please see @extended
; Author ........: water
; Modified ......:
; Remarks .......: To delete a shortcut by name isn't possible in Outlook 2002
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_BarShortcutDelete($oOL, $vOL_Groupname, $vOL_Shortcutname)

	If StringStripWS($vOL_Shortcutname, 3) = "" Then Return SetError(1, 0, 0)
	If StringStripWS($vOL_Groupname, 3) = "" Then Return SetError(2, 0, 0)
	Local $oOL_Group = $oOL.ActiveExplorer.Panes.Item("OutlookBar").Contents.Groups.Item($vOL_Groupname)
	If @error Then Return SetError(3, @error, 0)
	$oOL_Group.Shortcuts.Remove($vOL_Shortcutname)
	If @error Then Return SetError(4, @error, 0)
	Return 1

EndFunc   ;==>_OL_BarShortcutDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_BarShortcutGet
; Description ...: Returns all shortcuts of a group in the OutlookBar.
; Syntax.........: _OL_BarShortcutGet($oOL, $vOL_Group)
; Parameters ....: $oOL       - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Group - Name or 1-based index value of the group
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - Display name of the shortcut
;                  |1 - Variant indicating the target of the specified shortcut in a Shortcuts pane group
;                  Failure - Returns "" and sets @error:
;                  |1 - $vOL_Group is empty
;                  |2 - Error accessing the specified group. For details please see @extended
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: http://msdn.microsoft.com/en-us/library/bb176723(v=office.12).aspx
; Example .......: Yes
; ===============================================================================================================================
Func _OL_BarShortcutGet($oOL, $vOL_Group)

	Local $iOL_Index = 1
	If StringStripWS($vOL_Group, 3) = "" Then Return SetError(1, 0, "")
	Local $oOL_Group = $oOL.ActiveExplorer.Panes.Item("OutlookBar").Contents.Groups.Item($vOL_Group)
	If @error Then Return SetError(2, @error, "")
	Local $aOL_Shortcuts[$oOL_Group.Shortcuts.Count + 1][2]
	For $oOL_ShortCut In $oOL_Group.Shortcuts
		$aOL_Shortcuts[$iOL_Index][0] = $oOL_ShortCut.Name
		$aOL_Shortcuts[$iOL_Index][1] = $oOL_ShortCut.Target
		$iOL_Index += 1
	Next
	$aOL_Shortcuts[0][0] = $iOL_Index - 1
	$aOL_Shortcuts[0][1] = UBound($aOL_Shortcuts, 2)
	Return $aOL_Shortcuts

EndFunc   ;==>_OL_BarShortcutGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_CategoryAdd
; Description ...: Add a category.
; Syntax.........: _OL_CategoryAdd($oOL, $sOL_Category[, $iOL_Color = $olCategoryColorNone[, $sOL_Shortcut = $olCategoryShortcutKeyNone]])
; Parameters ....: $oOL          - Outlook object returned by a preceding call to _OL_Open()
;                  $sOL_Category - Name of the category to be created
;                  $iOL_Color    - Optional: Color for the new category (default = OlCategoryColorNone)
;                  $iOL_Shortcut - Optional: Shortcut key for the new category (default = OlCategoryShortcutKeyNone)
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - Error accessing the Session.Categories object. For details please see @extended
;                  |2 - Error creating the category. For details please see @extended
;                  |3 - Specified category already exists
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_CategoryAdd($oOL, $sOL_Category, $iOL_Color = $olCategoryColorNone, $iOL_Shortcut = $olCategoryShortcutKeyNone)

	Local $oOL_Categories = $oOL.Session.Categories
	If @error Then Return SetError(1, @error, 0)
	; Check if category already exists
	Local $aOL_Categories = _OL_CategoryGet($oOL)
	If IsArray($aOL_Categories) Then
		For $iIndex = 1 To $aOL_Categories[0][0]
			If $aOL_Categories[$iIndex][5] = $sOL_Category Then Return SetError(3, 0, 0)
		Next
	EndIf
	$oOL_Categories.Add($sOL_Category, $iOL_Color, $iOL_Shortcut)
	If @error Then Return SetError(2, @error, 0)
	Return 1

EndFunc   ;==>_OL_CategoryAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_CategoryDelete
; Description ...: Deletes a category.
; Syntax.........: _OL_CategoryDelete($oOL, $sOL_Category)
; Parameters ....: $oOL          - Outlook object returned by a preceding call to _OL_Open()
;                  $sOL_Category - Name, CategoryID or 1-based index value of the category to be deleted
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $sOL_Category is empty
;                  |2 - Error accessing the categories. For details please see @extended
;                  |3 - Specified category does not exist
;                  |4 - Error removing the specified category. For details please see @extended
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_CategoryDelete($oOL, $sOL_Category)

	If StringStripWS($sOL_Category, 3) = "" Then Return SetError(1, 0, 0)
	Local $oOL_Categories = $oOL.Session.Categories
	If @error Then Return SetError(2, @error, 0)
	; Check if category exists
	Local $bOL_Found = False
	Local $aOL_Categories = _OL_CategoryGet($oOL)
	If IsArray($aOL_Categories) Then
		For $iIndex = 1 To $aOL_Categories[0][0]
			If $aOL_Categories[$iIndex][5] = $sOL_Category Then
				$bOL_Found = True
				ExitLoop
			EndIf
		Next
	EndIf
	If $bOL_Found = False Then Return SetError(3, 0, 0)
	$oOL_Categories.Remove($sOL_Category)
	If @error Then Return SetError(4, @error, 0)
	Return 1

EndFunc   ;==>_OL_CategoryDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_CategoryGet
; Description ...: Returns all categories by which Outlook items can be grouped.
; Syntax.........: _OL_CategoryGet($oOL)
; Parameters ....: $oOL    - Outlook object returned by a preceding call to _OL_Open()
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - CategoryBorderColor: OLE_COLOR value that represents the border color of the color swatch for a category
;                  |1 - CategoryGradientBottomColor: OLE_COLOR value that represents the bottom gradient color of the color swatch for a category
;                  |2 - CategoryGradientTopColor: OLE_COLOR value that represents the top gradient color of the color swatch for a category
;                  |3 - CategoryID: String value that represents the unique identifier for the category
;                  |4 - Color: OlCategoryColor constant that indicates the color used by the category object
;                  |5 - Name: Display name for the category
;                  |6 - ShortcutKey: OlCategoryShortcutKey constant that specifies the shortcut key used by the category
;                  Failure - Returns "" and sets @error:
;                  |1 - Error accessing the Session.Categories object. For details please see @extended
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_CategoryGet($oOL)

	Local $iOL_Index = 1
	Local $oOL_Categories = $oOL.Session.Categories
	If @error Then Return SetError(1, @error, "")
	Local $aOL_Categories[$oOL_Categories.Count + 1][7]
	For $oOL_Category In $oOL_Categories
		$aOL_Categories[$iOL_Index][0] = $oOL_Category.CategoryBorderColor
		$aOL_Categories[$iOL_Index][1] = $oOL_Category.CategoryGradientBottomColor
		$aOL_Categories[$iOL_Index][2] = $oOL_Category.CategoryGradientTopColor
		$aOL_Categories[$iOL_Index][3] = $oOL_Category.CategoryID
		$aOL_Categories[$iOL_Index][4] = $oOL_Category.Color
		$aOL_Categories[$iOL_Index][5] = $oOL_Category.Name
		$aOL_Categories[$iOL_Index][6] = $oOL_Category.ShortcutKey
		$iOL_Index += 1
	Next
	$aOL_Categories[0][0] = $iOL_Index - 1
	$aOL_Categories[0][1] = UBound($aOL_Categories, 2)
	Return $aOL_Categories

EndFunc   ;==>_OL_CategoryGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_DistListMemberAdd
; Description ...: Adds one or multiple members to a distribution list.
; Syntax.........: _OL_DistListMemberAdd($oOL, $vOL_Item, $sOL_StoreID, $vOL_P1 = ""[, $vOL_P2 = ""[, $vOL_P3 = ""[, $vOL_P4 = ""[, $vOL_P5 = ""[, $vOL_P6 = ""[, $vOL_P7 = ""[, $vOL_P8 = ""[, $vOL_P9 = ""[, $vOL_P10 = ""]]]]]]]]])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the distribution list item
;                  $sOL_StoreID - StoreID where the EntryID is stored. Use the keyword "Default" to use the users mailbox
;                  $vOL_P1      - Member to add to the distribution list. Either a recipient object or the recipients name to be resolved
;                  +              or a zero based one-dimensional array with unlimited number of members
;                  $vOL_P2      - Optional: member to add to the distribution list. Either a recipient object or the recipients name to be resolved
;                  $vOL_P3      - Optional: Same as $vOL_P2
;                  $vOL_P4      - Optional: Same as $vOL_P2
;                  $vOL_P5      - Optional: Same as $vOL_P2
;                  $vOL_P6      - Optional: Same as $vOL_P2
;                  $vOL_P7      - Optional: Same as $vOL_P2
;                  $vOL_P8      - Optional: Same as $vOL_P2
;                  $vOL_P9      - Optional: Same as $vOL_P2
;                  $vOL_P10     - Optional: Same as $vOL_P2
; Return values .: Success - Distribution list object
;                  Failure - Returns 0 and sets @error:
;                  |1 - No distribution list item specified
;                  |2 - Item could not be found. EntryID might be wrong
;                  |3 - Error adding member to the distribution list. @extended = number of the invalid member (zero based)
;                  |4 - Member name could not be resolved. @extended = number of the invalid member (zero based)
; Author ........: water
; Modified.......:
; Remarks .......: $vOL_P2 to $vOL_P10 will be ignored if $vOL_P1 is an array of members
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_DistListMemberAdd($oOL, $vOL_Item, $sOL_StoreID, $vOL_P1, $vOL_P2 = "", $vOL_P3 = "", $vOL_P4 = "", $vOL_P5 = "", $vOL_P6 = "", $vOL_P7 = "", $vOL_P8 = "", $vOL_P9 = "", $vOL_P10 = "")

	Local $oOL_Recipient, $aOL_Recipients[10]
	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	; Move members into an array
	If Not IsArray($vOL_P1) Then
		$aOL_Recipients[0] = $vOL_P1
		$aOL_Recipients[1] = $vOL_P2
		$aOL_Recipients[2] = $vOL_P3
		$aOL_Recipients[3] = $vOL_P4
		$aOL_Recipients[4] = $vOL_P5
		$aOL_Recipients[5] = $vOL_P6
		$aOL_Recipients[6] = $vOL_P7
		$aOL_Recipients[7] = $vOL_P8
		$aOL_Recipients[8] = $vOL_P9
		$aOL_Recipients[9] = $vOL_P10
	Else
		$aOL_Recipients = $vOL_P1
	EndIf
	; add members to the distribution list
	For $iOL_Index = 0 To UBound($aOL_Recipients) - 1
		; member is an object = recipient name already resolved
		If IsObj($aOL_Recipients[$iOL_Index]) Then
			$vOL_Item.AddMember($aOL_Recipients[$iOL_Index])
			If @error Then Return SetError(3, $iOL_Index, 0)
		Else
			If StringStripWS($aOL_Recipients[$iOL_Index], 3) = "" Then ContinueLoop
			$oOL_Recipient = $oOL.Session.CreateRecipient($aOL_Recipients[$iOL_Index])
			If @error Or Not IsObj($oOL_Recipient) Then Return SetError(4, $iOL_Index, 0)
			$oOL_Recipient.Resolve
			If @error Or Not $oOL_Recipient.Resolved Then Return SetError(4, $iOL_Index, 0)
			$vOL_Item.AddMember($oOL_Recipient)
			If @error Then Return SetError(3, $iOL_Index, 0)
		EndIf
	Next
	$vOL_Item.Close(0)
	Return $vOL_Item

EndFunc   ;==>_OL_DistListMemberAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_DistListMemberDelete
; Description ...: Deletes one or multiple members from a distribution list.
; Syntax.........: _OL_DistListMemberDelete($oOL, $vOL_Item, $sOL_StoreID, $vOL_P1 = ""[, $vOL_P2 = ""[, $vOL_P3 = ""[, $vOL_P4 = ""[, $vOL_P5 = ""[, $vOL_P6 = ""[, $vOL_P7 = ""[, $vOL_P8 = ""[, $vOL_P9 = ""[, $vOL_P10 = ""]]]]]]]]])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the distribution list item. Use the keyword "Default" to use the users mailbox
;                  $sOL_StoreID - StoreID where the EntryID is stored. Use the keyword "Default" to use the users mailbox
;                  $vOL_P1      - Member to delete from the distribution list. Either a recipient object or the recipients name to be resolved
;                  +              or a zero based one-dimensional array with unlimited number of members
;                  $vOL_P2      - Optional: member to delete from the distribution list. Either a recipient object or the recipients name to be resolved
;                  $vOL_P3      - Optional: Same as $vOL_P2
;                  $vOL_P4      - Optional: Same as $vOL_P2
;                  $vOL_P5      - Optional: Same as $vOL_P2
;                  $vOL_P6      - Optional: Same as $vOL_P2
;                  $vOL_P7      - Optional: Same as $vOL_P2
;                  $vOL_P8      - Optional: Same as $vOL_P2
;                  $vOL_P9      - Optional: Same as $vOL_P2
;                  $vOL_P10     - Optional: Same as $vOL_P2
; Return values .: Success - Distribution list object
;                  Failure - Returns 0 and sets @error:
;                  |1 - No distribution list item specified
;                  |2 - Distribution list item could not be found. EntryID might be wrong
;                  |3 - Error removing member from the distribution list. @extended = number of the invalid member (zero based)
;                  |4 - Member name could not be resolved. @extended = number of the invalid member (zero based)
; Author ........: water
; Modified.......:
; Remarks .......: $vOL_P2 to $vOL_P10 will be ignored if $vOL_P1 is an array of members
;+
;                  No error is returned if a specified member is not a member of this distribution list
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_DistListMemberDelete($oOL, $vOL_Item, $sOL_StoreID, $vOL_P1 = "", $vOL_P2 = "", $vOL_P3 = "", $vOL_P4 = "", $vOL_P5 = "", $vOL_P6 = "", $vOL_P7 = "", $vOL_P8 = "", $vOL_P9 = "", $vOL_P10 = "")

	Local $oOL_Recipient, $aOL_Recipients[10]
	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	; Move members into an array
	If Not IsArray($vOL_P1) Then
		$aOL_Recipients[0] = $vOL_P1
		$aOL_Recipients[1] = $vOL_P2
		$aOL_Recipients[2] = $vOL_P3
		$aOL_Recipients[3] = $vOL_P4
		$aOL_Recipients[4] = $vOL_P5
		$aOL_Recipients[5] = $vOL_P6
		$aOL_Recipients[6] = $vOL_P7
		$aOL_Recipients[7] = $vOL_P8
		$aOL_Recipients[8] = $vOL_P9
		$aOL_Recipients[9] = $vOL_P10
	Else
		$aOL_Recipients = $vOL_P1
	EndIf
	; Delete members from the distribution list
	For $iOL_Index = 0 To UBound($aOL_Recipients) - 1
		; member is an object = recipient name already resolved
		If IsObj($aOL_Recipients[$iOL_Index]) Then
			$vOL_Item.RemoveMember($aOL_Recipients[$iOL_Index])
			If @error Then Return SetError(3, $iOL_Index, 0)
		Else
			If StringStripWS($aOL_Recipients[$iOL_Index], 3) = "" Then ContinueLoop
			$oOL_Recipient = $oOL.Session.CreateRecipient($aOL_Recipients[$iOL_Index])
			If @error Or Not IsObj($oOL_Recipient) Then Return SetError(4, $iOL_Index, 0)
			$oOL_Recipient.Resolve
			If @error Or Not $oOL_Recipient.Resolved Then Return SetError(4, $iOL_Index, 0)
			$vOL_Item.RemoveMember($oOL_Recipient)
			If @error Then Return SetError(3, $iOL_Index, 0)
		EndIf
	Next
	$vOL_Item.Close(0)
	Return $vOL_Item

EndFunc   ;==>_OL_DistListMemberDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_DistListMemberGet
; Description ...: Gets all members of a distribution list.
; Syntax.........: _OL_DistListMemberGet($oOL, $vOL_Item[, $sOL_StoreID = Default])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the distribution list item
;                  $sOL_StoreID - Optional: StoreID where the EntryID is stored. Use the keyword "Default" to use the users mailbox
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - Recipient object of the member
;                  |1 - Name of the member
;                  |2 - EntryID of the member
;                  Failure - Returns "" and sets @error:
;                  |1 - No distribution list item specified
;                  |2 - Item could not be found. EntryID might be wrong
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_DistListMemberGet($oOL, $vOL_Item, $sOL_StoreID = Default)

	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, "")
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	Local $aOL_Members[$vOL_Item.MemberCount + 1][3] = [[$vOL_Item.MemberCount, 3]]
	For $iOL_Index = 1 To $vOL_Item.MemberCount
		$aOL_Members[$iOL_Index][0] = $vOL_Item.GetMember($iOL_Index)
		$aOL_Members[$iOL_Index][1] = $vOL_Item.GetMember($iOL_Index).Name
		$aOL_Members[$iOL_Index][2] = $vOL_Item.GetMember($iOL_Index).EntryID
	Next
	Return $aOL_Members

EndFunc   ;==>_OL_DistListMemberGet

; #FUNCTION# ====================================================================================================================
; Name...........: _OL_FolderAccess
; Description ...: Access a folder.
; Syntax.........: _OL_FolderAccess($oOL[, $sOL_Folder = "" [, $iOL_FolderType = Default[, $iOL_ItemType = Default]]])
; Parameters ....: $oOL            - Outlook object returned by a preceding call to _OL_Open()
;                  $sOL_Folder     - Optional: Name of folder to access (default = default folder of current user (class specified by $iOL_FolderType))
;                  |  "rootfolder\subfolder\...\subfolder" to access any public folder or any folder of the current user
;                  +      "rootfolder" for the current user can be replaced by "*"
;                  |  "\\firstname name" to access the default folder of another user (class specified by $iOL_FolderType)
;                  |  "\\firstname name\\subfolder\...\subfolder" to access a subfolder of the default folder of another user (class specified by $iOL_FolderType)
;                  |  "\\firstname name\subfolder\..\subfolder" to access any subfolder of another user
;                  +      "firstname name" for the current user can be replaced by "*"
;                  |  "" to access the default folder of the current user (class specified by $iOL_FolderType)
;                  |  "\subfolder" to access a subfolder of the default folder of the current user (class specified by $iOL_FolderType)
;                  $iOL_FolderType - Optional: Type of folder if you want to access a default folder. Is defined by the Outlook OlDefaultFolders enumeration (default = Default)
;                  $iOL_ItemType   - Optional: Type of item which is used to select the default folder. Is defined by the Outlook OlItemType enumeration (default = Default)
; Return values .: Success - one-dimensional one based array with the following information:
;                  |1 - Object to the folder
;                  |2 - Default item type (integer) for the specified folder. Defined by the Outlook OlItemType enumeration
;                  |3 - StoreID (string) of the store to access the folder by ID
;                  |4 - EntryID (string) of the folder to access the folder by ID
;                  |5 - Folderpath (string)
;                  Failure - Returns "" and sets @error:
;                  |1 - $iOL_FolderType is missing or not a number
;                  |2 - Could not resolve specified User in $sOL_Folder
;                  |3 - Error accessing specified folder
;                  |4 - Specified folder could not be found
;                  |5 - Neither $sOL_Folder, $iOL_FolderType nor $iOL_ItemType was specified
;                  |6 - No valid $iOL_ItemType was found to set the default folder $iOL_FolderType accordingly
; Author ........: water
; Modified.......:
; Remarks .......: If you only specify $iOL_ItemType then $iOL_FolderType is set to the default folder for this item type.
;                  Supported item types are: $olAppointmentItem, $olContactItem, $olDistributionListItem, $olJournalItem, $olMailItem, $olNoteItem and $olTaskItem
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_FolderAccess($oOL, $sOL_Folder = "", $iOL_FolderType = Default, $iOL_ItemType = Default)

	Local $oOL_Folder, $aOL_Folders, $aOL_Result[6] = [5]
	; Set $iOL_FolderType based on $iOL_ItemType
	If $sOL_Folder = "" And $iOL_FolderType = Default Then
		If $iOL_ItemType = Default Then Return SetError(5, 0, "")
		Local $aFolders[8][2] = [[7, 2],[$olAppointmentItem, $olFolderCalendar],[$olContactItem, $olFolderContacts],[$olDistributionListItem, $olFolderContacts], _
				[$olJournalItem, $olFolderJournal],[$olMailItem, $olFolderDrafts],[$olNoteItem, $olFolderNotes],[$olTaskItem, $olFolderTasks]]
		Local $bOL_Found = False
		For $iIndex = 1 To $aFolders[0][0]
			If $iOL_ItemType = $aFolders[$iIndex][0] Then
				$iOL_FolderType = $aFolders[$iIndex][1]
				$bOL_Found = True
				ExitLoop
			EndIf
		Next
		If $bOL_Found = False Then SetError(6, 0, "")
	EndIf
	Local $oOL_Namespace = $oOL.GetNamespace("MAPI")
	If $sOL_Folder = "" Or (StringLeft($sOL_Folder, 1) = "\" And _ ; No folder specified. Use default folder depending on $iOL_FolderType
			StringMid($sOL_Folder, 2, 1) <> "\") Then ; Folder starts with "\" = subfolder in default folder depending on $iOL_FolderType
		If $iOL_FolderType = Default Or Not IsNumber($iOL_FolderType) Then Return SetError(1, 0, "") ; Required $iOL_FolderType is missing
		$oOL_Folder = $oOL_Namespace.GetDefaultFolder($iOL_FolderType)
		If @error Or Not IsObj($oOL_Folder) Then Return SetError(3, @error, "")
		If $sOL_Folder <> "" Then
			$aOL_Folders = StringSplit(StringMid($sOL_Folder, 2), "\")
			SetError(0) ; Reset @error possibly set by StringSplit
			For $iOL_Index = 1 To $aOL_Folders[0]
				$oOL_Folder = $oOL_Folder.Folders($aOL_Folders[$iOL_Index])
				If @error Or Not IsObj($oOL_Folder) Then Return SetError(4, @error, "")
			Next
		EndIf
	Else
		If StringLeft($sOL_Folder, 2) = "\\" Then ; Access a folder of another user
			If $iOL_FolderType = Default Or Not IsNumber($iOL_FolderType) Then Return SetError(1, 0, "") ; Required $iOL_FolderType is missing
			$aOL_Folders = StringSplit(StringMid($sOL_Folder, 3), "\") ; Split off Recipient
			SetError(0) ; Reset @error possibly set by StringSplit
			If $aOL_Folders[1] = "*" Then $aOL_Folders[1] = $oOL.GetNameSpace("MAPI").CurrentUser.Name
			Local $oOL_Dummy = $oOL_Namespace.CreateRecipient("=" & $aOL_Folders[1]) ; Create Recipient. "=" sets resolve to strict
			$oOL_Dummy.Resolve ; Resolve
			If Not $oOL_Dummy.Resolved Then Return SetError(2, 0, "")
			If $aOL_Folders[0] > 1 And StringStripWS($aOL_Folders[2], 3) = "" Then ; Access a subfolder of the specified default folder of another user
				$oOL_Folder = $oOL_Namespace.GetSharedDefaultFolder($oOL_Dummy, $iOL_FolderType)
				If @error Or Not IsObj($oOL_Folder) Then Return SetError(3, @error, "")
			Else ; Access any folder of another user
				$oOL_Folder = $oOL_Namespace.GetSharedDefaultFolder($oOL_Dummy, $iOL_FolderType).Parent
				If @error Or Not IsObj($oOL_Folder) Then Return SetError(3, @error, "")
			EndIf
		Else
			$aOL_Folders = StringSplit($sOL_Folder, "\") ; Folder specified. Split and get the object
			SetError(0) ; Reset @error possibly set by StringSplit
			If $aOL_Folders[1] = "*" Then $aOL_Folders[1] = $oOL_Namespace.GetDefaultFolder($olFolderInbox).Parent.Name
			$oOL_Folder = $oOL_Namespace.Folders($aOL_Folders[1])
			If @error Or Not IsObj($oOL_Folder) Then Return SetError(4, @error, "")
		EndIf
		If $aOL_Folders[0] > 1 Then ; Access subfolders
			For $iOL_Index = 2 To $aOL_Folders[0]
				If $aOL_Folders[$iOL_Index] <> "" Then
					$oOL_Folder = $oOL_Folder.Folders($aOL_Folders[$iOL_Index])
					If @error Or Not IsObj($oOL_Folder) Then Return SetError(4, @error, "")
				EndIf
			Next
		EndIf
	EndIf
	$aOL_Result[1] = $oOL_Folder
	$aOL_Result[2] = $oOL_Folder.DefaultItemType
	$aOL_Result[3] = $oOL_Folder.StoreID
	$aOL_Result[4] = $oOL_Folder.EntryID
	$aOL_Result[5] = $oOL_Folder.FolderPath
	Return $aOL_Result

EndFunc   ;==>_OL_FolderAccess

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_FolderArchiveGet
; Description ...: Returns the auto-archive properties of a folder.
; Syntax.........: _OL_FolderArchiveGet($oOL_Folder)
; Parameters ....: $oOL_Folder - Folder object of the folder to be changed as returned by _OL_FolderAccess
; Return values .: Success - one-dimensional one based array with the following information:
;                  |1 - AgeFolder:   TRUE: Archive or delete items in the folder as specified
;                  |2 - DeleteItems: TRUE: Delete, instead of archive, items that are older than the aging period
;                  |3 - FileName:    File for archiving aged items
;                  |4 - Granularity: Unit of time for aging, whether archiving is to be calculated in units of months, weeks, or days.
;                  +Valid granularity: 0=Months, 1=Weeks, 2=Days
;                  |5 - Period :     Amount of time in the given granularity. Value between 1 and 999
;                  |6 - Default:     Indicates which settings should be set to the default.
;                  |0: Nothing assumes a default value
;                  |1: Only the file location assumes a default value.
;                  +This is the same as checking Archive this folder using these settings and Move old items to default archive folder in the AutoArchive
;                  +tab of the Properties dialog box for the folder
;                  |3: All settings assume a default value. This is the same as checking Archive items in this folder using default settings in the AutoArchive
;                  +tab of the Properties dialog box for the folder
;                  Failure - Returns 0 and sets @error:
;                  |1 - Error creating $oOL_Storage. Please check @extended for details
;                  |2 - Error creating $oOL_PA. Please check @extended for details
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_FolderArchiveGet($oOL_Folder)

	Local Const $sProptagURL = "http://schemas.microsoft.com/mapi/proptag/"
	Local Const $sPR_AGING_AGE_FOLDER = $sProptagURL & "0x6857000B"
	Local Const $sPR_AGING_PERIOD = $sProptagURL & "0x36EC0003"
	Local Const $sPR_AGING_GRANULARITY = $sProptagURL & "0x36EE0003"
	Local Const $sPR_AGING_DELETE_ITEMS = $sProptagURL & "0x6855000B"
	Local Const $sPR_AGING_FILE_NAME_AFTER9 = $sProptagURL & "0x6859001E"
	Local Const $sPR_AGING_DEFAULT = $sProptagURL & "0x685E0003"

	Local $aOL_AutoArchive[7] = [6]
	; Create or get solution storage in given folder by message class
	Local $oOL_Storage = $oOL_Folder.GetStorage("IPC.MS.Outlook.AgingProperties", $olIdentifyByMessageClass)
	If @error Or Not IsObj($oOL_Storage) Then Return SetError(1, @error, 0)
	Local $oOL_PA = $oOL_Storage.PropertyAccessor
	If @error Or Not IsObj($oOL_PA) Then Return SetError(2, @error, 0)
	$aOL_AutoArchive[1] = $oOL_PA.GetProperty($sPR_AGING_AGE_FOLDER)
	$aOL_AutoArchive[2] = $oOL_PA.GetProperty($sPR_AGING_GRANULARITY)
	$aOL_AutoArchive[3] = $oOL_PA.GetProperty($sPR_AGING_DELETE_ITEMS)
	$aOL_AutoArchive[4] = $oOL_PA.GetProperty($sPR_AGING_PERIOD)
	$aOL_AutoArchive[5] = $oOL_PA.GetProperty($sPR_AGING_FILE_NAME_AFTER9)
	$aOL_AutoArchive[6] = $oOL_PA.GetProperty($sPR_AGING_DEFAULT)
	Return $aOL_AutoArchive

EndFunc   ;==>_OL_FolderArchiveGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_FolderArchiveSet
; Description ...: Sets the auto-archive properties of a folder and (optional) all subfolders.
; Syntax.........: _OL_FolderArchiveSet($oOL_Folder, $bOL_Recursive, $bOL_AgeFolder[, $bOL_DeleteItems = Default[, $sOL_FileName = Default[, $iOL_Granularity = Default[, $iOL_Period = Default[, $iOL_Default = Default]]]]])
; Parameters ....: $oOL_Folder      - Folder object of the folder to be changed as returned by _OL_FolderAccess
;                  $bOL_Recursive   - TRUE: Set properties for the specified folder and all subfolders
;                  $bOL_AgeFolder   - TRUE: Archive or delete items in the folder as specified
;                  $bOL_DeleteItems - Optional: TRUE: Delete, instead of archive, items that are older than the aging period (default = Default)
;                  $sOL_FileName    - Optional: File for archiving aged items. If this is an empty string, the default archive file, archive.pst, will be used (default = Default)
;                  $iOL_Granularity - Optional: Unit of time for aging, whether archiving is to be calculated in units of months, weeks, or days (default = Default).
;                  +Valid granularity: 0=Months, 1=Weeks, 2=Days
;                  $iOL_Period      - Optional: Amount of time in the given granularity. Valid period: 1-999 (default = Default)
;                  $iOL_Default     - Optional: Indicates which settings should be set to the default (default = Default):
;                  |0: Nothing assumes a default value
;                  |1: Only the file location assumes a default value.
;                  +This is the same as checking Archive this folder using these settings and Move old items to default archive folder in the AutoArchive
;                  +tab of the Properties dialog box for the folder
;                  |3: All settings assume a default value. This is the same as checking Archive items in this folder using default settings in the AutoArchive
;                  +tab of the Properties dialog box for the folder
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oOL_Folder is not an object
;                  |2 - $bOL_Recursive is not boolean
;                  |3 - $bOL_AgeFolder is not boolean
;                  |4 - $bOL_DeleteItems is not boolean
;                  |5 - $iOL_Granularity is not an integer or <0 or > 2
;                  |6 - $iOL_Period is not an integer or < 1 or > 999
;                  |7 - $iOL_Default is not an integer or an invalid number (must be 0, 1 or 3)
;                  |8 - Error creating $oOL_Storage. Please check @extended for details
;                  |9 - Error creating $oOL_PA. Please check @extended for details
;                  |10 - Error saving changed properties. Please check @extended for details
; Author ........: water
; Modified ......:
; Remarks .......: More links:
;                  http://msdn.microsoft.com/en-us/library/ff870123.aspx (Outlook 2010)
;                  https://blogs.msdn.com/b/jmazner/archive/2006/10/30/setting-autoarchive-properties-on-a-folder-hierarchy-in-outlook-2007.aspx?Redirected=true
; Related .......:
; Link ..........: http://msdn.microsoft.com/en-us/library/bb176434(v=office.12).aspx (Outlook 2007)
; Example .......: Yes
; ===============================================================================================================================
Func _OL_FolderArchiveSet($oOL_Folder, $bOL_Recursive, $bOL_AgeFolder, $bOL_DeleteItems = Default, $sOL_FileName = Default, $iOL_Granularity = Default, $iOL_Period = Default, $iOL_Default = Default)

	Local Const $sProptagURL = "http://schemas.microsoft.com/mapi/proptag/"
	Local Const $sPR_AGING_AGE_FOLDER = $sProptagURL & "0x6857000B"
	Local Const $sPR_AGING_PERIOD = $sProptagURL & "0x36EC0003"
	Local Const $sPR_AGING_GRANULARITY = $sProptagURL & "0x36EE0003"
	Local Const $sPR_AGING_DELETE_ITEMS = $sProptagURL & "0x6855000B"
	Local Const $sPR_AGING_FILE_NAME_AFTER9 = $sProptagURL & "0x6859001E"
	Local Const $sPR_AGING_DEFAULT = $sProptagURL & "0x685E0003"

	If Not IsObj($oOL_Folder) Then Return SetError(1, 0, 0)
	If Not IsBool($bOL_Recursive) Then Return SetError(2, 0, 0)
	If Not IsBool($bOL_AgeFolder) Then Return SetError(3, 0, 0)
	If $bOL_DeleteItems <> Default And Not IsBool($bOL_DeleteItems) Then Return SetError(4, 0, 0)
	If $iOL_Granularity <> Default And Not IsInt($iOL_Granularity) Or $iOL_Granularity < 0 Or $iOL_Granularity > 2 Then Return SetError(5, 0, 0)
	If $iOL_Period <> Default And (Not IsInt($iOL_Period) Or $iOL_Period < 1 Or $iOL_Period > 999) Then Return SetError(6, 0, 0)
	If $iOL_Default <> Default And (Not IsInt($iOL_Default) Or ($iOL_Default <> 0 And $iOL_Default <> 1 And $iOL_Default <> 3)) Then Return SetError(7, 0, 0)
	; Create or get solution storage in given folder by message class
	Local $oOL_Storage = $oOL_Folder.GetStorage("IPC.MS.Outlook.AgingProperties", $olIdentifyByMessageClass)
	If @error Or Not IsObj($oOL_Storage) Then Return SetError(8, @error, 0)
	Local $oOL_PA = $oOL_Storage.PropertyAccessor
	If @error Or Not IsObj($oOL_PA) Then Return SetError(9, @error, 0)
	; Set the 6 aging properties in the solution storage
	$oOL_PA.SetProperty($sPR_AGING_AGE_FOLDER, $bOL_AgeFolder)
	If $iOL_Granularity <> Default Then $oOL_PA.SetProperty($sPR_AGING_GRANULARITY, $iOL_Granularity)
	If $bOL_DeleteItems <> Default Then $oOL_PA.SetProperty($sPR_AGING_DELETE_ITEMS, $bOL_DeleteItems)
	If $iOL_Period <> Default Then $oOL_PA.SetProperty($sPR_AGING_PERIOD, $iOL_Period)
	If $sOL_FileName <> Default Then $oOL_PA.SetProperty($sPR_AGING_FILE_NAME_AFTER9, $sOL_FileName)
	If $iOL_Default <> Default Then $oOL_PA.SetProperty($sPR_AGING_DEFAULT, $iOL_Default)
	; Save changes as hidden messages to the associated portion of the folder
	$oOL_Storage.Save
	If @error Then Return SetError(10, @error, 0)
	; Process subfolders
	If $bOL_Recursive Then
		For $oOL_SubFolder In $oOL_Folder.Folders
			_OL_FolderArchiveSet($oOL_SubFolder, $bOL_Recursive, $bOL_AgeFolder, $bOL_DeleteItems, $sOL_FileName, $iOL_Granularity, $iOL_Period, $iOL_Default)
			If @error Then Return SetError(@error, @extended, 0)
		Next
	EndIf
	Return 1

EndFunc   ;==>_OL_FolderArchiveSet

; #FUNCTION# ====================================================================================================================
; Name...........: _OL_FolderCopy
; Description ...: Copies a folder, all subfolders and all contained items.
; Syntax.........: _OL_FolderCopy($oOL, $vOL_SourceFolder, $vOL_TargetFolder)
; Parameters ....: $oOL              - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_SourceFolder - Source folder name or object of the folder to be copied
;                  $vOL_TargetFolder - Target folder name or object of the folder to be copied to
; Return values .: Success - Folder object of the copied folder
;                  Failure - Returns 0 and sets @error:
;                  |1 - Error accessing the specified source folder. See @extended for errorcode returned by _OL_FolderAccess
;                  |2 - Error accessing the specified target folder. See @extended for errorcode returned by _OL_FolderAccess
;                  |3 - Source folder has not been specified or is empty
;                  |4 - Target folder has not been specified or is empty
;                  |5 - Source and target folder are the same
;                  |6 - Error copying the folder to the target folder. For details check @extended
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_FolderCopy($oOL, $vOL_SourceFolder, $vOL_TargetFolder)

	Local $aOL_Temp

	If Not IsObj($vOL_SourceFolder) Then
		If StringStripWS($vOL_SourceFolder, 3) = "" Then Return SetError(3, 0, 0)
		$aOL_Temp = _OL_FolderAccess($oOL, $vOL_SourceFolder)
		If @error Then Return SetError(1, @error, 0)
		$vOL_SourceFolder = $aOL_Temp[1]
	EndIf
	If Not IsObj($vOL_TargetFolder) Then
		If StringStripWS($vOL_TargetFolder, 3) = "" Then Return SetError(4, 0, 0)
		$aOL_Temp = _OL_FolderAccess($oOL, $vOL_TargetFolder)
		If @error Then Return SetError(2, @error, 0)
		$vOL_TargetFolder = $aOL_Temp[1]
	EndIf
	If $vOL_SourceFolder = $vOL_TargetFolder Then Return SetError(5, 0, 0)
	Local $vOL_Folder = $vOL_SourceFolder.CopyTo($vOL_TargetFolder)
	If @error Then Return SetError(6, @error, 0)
	Return $vOL_Folder

EndFunc   ;==>_OL_FolderCopy

; #FUNCTION# ====================================================================================================================
; Name...........: _OL_FolderCreate
; Description ...: Create a folder and subfolders.
; Syntax.........: _OL_FolderCreate($oOL, $sOL_Folder, $iOL_FolderType[, $vOL_StartFolder = ""])
; Parameters ....: $oOL             - Outlook object returned by a preceding call to _OL_Open()
;                  $sOL_Folder      - Folder(s) to be created
;                  $iOL_FolderType  - Type of folder(s) to be created. Is defined by the Outlook OlDefaultFolders enumeration
;                  $vOL_StartFolder - Optional: Folder object as returned by _OL_FolderAccess or full name of folder to create the new
;                  +folder in (default is root folder)
; Return values .: Success - Folder object of the created folder
;                  Failure - Returns 0 and sets @error:
;                  |1 - $iOL_FolderType is missing or not a number
;                  |2 - Folder could not be created. See @extended for COM error code
;                  |3 - Error accessing specified folder. See @extended for the error code of _OL_AccessFolder
;                  |4 - Folder already exists
;                  |5 - Error adding folder. See @extended for the error code of the Add method
; Author ........: water
; Modified.......:
; Remarks .......: The folder and subfolders all have the same type specified by $iOL_FolderType.
;                  To set properties of a folder please use _OL_FolderModfiy
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_FolderCreate($oOL, $sOL_Folder, $iOL_FolderType, $vOL_StartFolder = "")

	If Not IsNumber($iOL_FolderType) Then Return SetError(1, 0, 0) ; Required $iOL_FolderType is missing
	Local $oOL_Namespace = $oOL.GetNamespace("MAPI")
	If Not IsObj($vOL_StartFolder) Then
		If StringStripWS($vOL_StartFolder, 3) = "" Then ; Startfolder is not specified - use root folder
			Local $oOL_Inbox = $oOL_Namespace.GetDefaultFolder($olFolderInbox)
			$vOL_StartFolder = $oOL_Inbox.Parent
		Else
			Local $aOL_Temp = _OL_FolderAccess($oOL, $vOL_StartFolder)
			If @error Then Return SetError(3, @error, 0)
			$vOL_StartFolder = $aOL_Temp[1]
		EndIf
	EndIf
	Local $aOL_SubFolders = StringSplit($sOL_Folder, "\")
	SetError(0)
	For $iOL_Index = 1 To $aOL_SubFolders[0]
		; Check if folder already exists
		For $oOL_Folder In $vOL_StartFolder.Folders
			If $oOL_Folder.Name = $aOL_SubFolders[$iOL_Index] Then Return SetError(4, 0, 0)
		Next
		$vOL_StartFolder = $vOL_StartFolder.Folders.Add($aOL_SubFolders[$iOL_Index], $iOL_FolderType)
		If @error Or Not IsObj($vOL_StartFolder) Then Return SetError(5, @error, 0)
	Next
	Return $vOL_StartFolder

EndFunc   ;==>_OL_FolderCreate

; #FUNCTION# ====================================================================================================================
; Name...........: _OL_FolderDelete
; Description ...: Deletes a folder, all subfolders and all contained items.
; Syntax.........: _OL_FolderDelete($oOL, $sOL_Folder[, $iOL_Flags = 0])
; Parameters ....: $oOL        - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Folder - Folder object as returned by _OL_FolderAccess or full name of folder to be deleted
;                  $iOL_Flags  - Optional: Specifies what should be deleted. Can be a combination of the following:
;                  |0: Deletes the folder, all subfolders and all contained items (default)
;                  |1: Deletes all items (but no folders) in the specified folder
;                  |2: Recursively deletes all items (but no folders) in the specified folder and all subfolders
;                  |4: Deletes all subfolders and their items in the specified folder (but not the items in the specified folder)
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - Error accessing specified folder. See @extended for the error code of _OL_AccessFolder
;                  |2 - Folder could not be deleted. See @extended for COM error code
;                  |3 - Folder has not been specified or is empty
;                  |4 - Subfolder could not be deleted. See @extended for COM error code
;                  |5 - Item could not be deleted. See @extended for COM error code
; Author ........: water
; Modified.......:
; Remarks .......: Flag usage:
;                  To empty the trash folder (or any Outlook system folder) and delete all items plus all subfolders use $iOL_Flags = 5
;                  To delete all items in all folders and subfolders but retain the folder structure use $iOL_Flags = 3
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_FolderDelete($oOL, $vOL_Folder, $iOL_Flags = 0)

	If Not IsObj($vOL_Folder) Then
		If StringStripWS($vOL_Folder, 3) = "" Then Return SetError(3, 0, 0)
		Local $aOL_Temp = _OL_FolderAccess($oOL, $vOL_Folder)
		If @error Then Return SetError(1, @error, 0)
		$vOL_Folder = $aOL_Temp[1]
	EndIf
	; Delete the folder, all subfolders and all contained items
	If $iOL_Flags = 0 Then
		$vOL_Folder.Delete
		If @error Then Return SetError(2, @error, 0)
		Return 1
	EndIf
	; Delete items recursively
	If BitAND($iOL_Flags, 2) = 2 Then
		For $oOL_SubFolder In $vOL_Folder.Folders
			$aOL_Temp = _OL_FolderDelete($oOL, $oOL_SubFolder, $iOL_Flags)
			If @error Then Return SetError(2, @error, "")
		Next
	EndIf
	; Just delete all items in the specified folder
	If BitAND($iOL_Flags, 1) = 1 Or BitAND($iOL_Flags, 2) = 2 Then
		For $iOL_Index = $vOL_Folder.Items.Count To 1 Step -1
			$vOL_Folder.Items($iOL_Index).Delete
			If @error Then Return SetError(5, @error, 0)
		Next
	EndIf
	; Delete all subfolders and all contained items
	If BitAND($iOL_Flags, 4) = 4 Then
		For $iOL_Index = $vOL_Folder.Folders.Count To 1 Step -1
			$vOL_Folder.Folders($iOL_Index).Delete
			If @error Then Return SetError(4, @error, 0)
		Next
	EndIf
	Return 1

EndFunc   ;==>_OL_FolderDelete

; #FUNCTION# ====================================================================================================================
; Name...........: _OL_FolderExists
; Description ...: Checks if the specified folder exists.
; Syntax.........: _OL_FolderExists($oOL, $sOL_Folder)
; Parameters ....: $oOL        - Outlook object returned by a preceding call to _OL_Open()
;                  $sOL_Folder - Full name of folder to be checked
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - Error accessing specified folder. See @extended for the error code of _OL_AccessFolder
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_FolderExists($oOL, $sOL_Folder)

	_OL_FolderAccess($oOL, $sOL_Folder)
	If @error Then Return SetError(1, @error, 0)
	Return 1

EndFunc   ;==>_OL_FolderExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_FolderFind
; Description ...: Finds folders filtered by name and/or default item type.
; Syntax.........: _OL_FolderFind($oOL, $vOL_Folder[, $iOL_Recursionlevel = 0[, $sOL_FolderName = ""[, $iOL_StringMatch = 1[, $iOL_DefaultItemType = Default]]]])
; Parameters ....: $oOL                  - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Folder           - Folder object as returned by _OL_FolderAccess or full name of folder where the search will be started.
;                  +If you want to search a default folder you have to specify the folder object.
;                  $iOL_Recursionlevel   - Optional: Number of subfolders to search. 0 means only the specified folder is searched (default = 0)
;                  $sOL_FolderName       - Optional: String to search for in the folder name. The matching mode (exact or substring) is specified by the next parameter (default = "")
;                  +Can be combined with $iOL_DefaultItemType
;                  $iOL_StringMatch      - Optional: Matching mode (default = 1). Can be one of the following:
;                  |  1: Exact match
;                  |  2: Substring
;                  $iOL_DefaultItemType  - Optional: Only return folders which can hold items of the following item type. Is defined by the Outlook OlItemType enumeration.
;                  +Can be combined with $sOL_FolderName (default = Default)
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - Object of the folder
;                  |1 - FolderPath
;                  |2 - Name
;                  Failure - Returns "" and sets @error:
;                  |1 - $sOL_FolderName and $iOL_DefaultItemType have not been set
;                  |2 - Error accessing the specified folder. See @extended for errorcode returned by _OL_FolderAccess
; Author ........: water
; Modified ......:
; Remarks .......: You have to specify at least $sOL_FolderName or $iOL_DefaultItemType
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_FolderFind($oOL, $vOL_Folder, $iOL_Recursionlevel = 0, $sOL_FolderName = "", $iOL_StringMatch = 1, $iOL_DefaultItemType = Default)

	Local $iCount1 = 1, $aOL_Temp, $bOL_Found
	If $vOL_Folder = "" And $iOL_DefaultItemType = Default Then Return SetError(1, 0, "")
	If Not IsObj($vOL_Folder) Then
		$aOL_Temp = _OL_FolderAccess($oOL, $vOL_Folder)
		If @error Then Return SetError(2, @error, "")
		$vOL_Folder = $aOL_Temp[1]
	EndIf
	Local $aOL_Folders[$vOL_Folder.Folders.Count + 1][3]
	For $vOL_Folder In $vOL_Folder.Folders
		$bOL_Found = False
		If $sOL_FolderName <> "" Then
			If $iOL_StringMatch = 1 And $vOL_Folder.Name == $sOL_FolderName Then $bOL_Found = True
			If $iOL_StringMatch = 2 And StringInStr($vOL_Folder.Name, $sOL_FolderName) > 0 Then $bOL_Found = True
		EndIf
		If $iOL_DefaultItemType <> Default And $vOL_Folder.DefaultItemType = $iOL_DefaultItemType Then $bOL_Found = True
		If $bOL_Found Then
			$aOL_Folders[$iCount1][0] = $vOL_Folder
			$aOL_Folders[$iCount1][1] = $vOL_Folder.FolderPath
			$aOL_Folders[$iCount1][2] = $vOL_Folder.Name
			$iCount1 += 1
		EndIf
		If $iOL_Recursionlevel > 0 Then
			$aOL_Temp = _OL_FolderFind($oOL, $vOL_Folder, $iOL_Recursionlevel - 1, $sOL_FolderName, $iOL_StringMatch, $iOL_DefaultItemType)
			_OL_ArrayConcatenate($aOL_Folders, $aOL_Temp, 0)
		EndIf
	Next
	If UBound($aOL_Folders, 1) > 1 Then
		_ArraySort($aOL_Folders, 1, 1, 0, 1)
		For $iCount1 = 1 To UBound($aOL_Folders, 1) - 1
			If $aOL_Folders[$iCount1][0] = "" Then
				ReDim $aOL_Folders[$iCount1][UBound($aOL_Folders, 2)]
				ExitLoop
			EndIf
		Next
		_ArraySort($aOL_Folders, 0, 1, 0, 1)
	EndIf
	$aOL_Folders[0][0] = UBound($aOL_Folders, 1) - 1
	$aOL_Folders[0][1] = UBound($aOL_Folders, 2)
	Return $aOL_Folders

EndFunc   ;==>_OL_FolderFind

; #FUNCTION# ====================================================================================================================
; Name...........: _OL_FolderGet
; Description ...: Get information about the current or any other folder.
; Syntax.........: _OL_FolderGet($oOL[, $vOL_Folder = ""])
; Parameters ....: $oOL        - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Folder - Optional: Folder object as returned by _OL_FolderAccess or full name of folder (default = "" = current folder)
; Return values .: Success - one-dimensional one based array with the following information:
;                  |1  - Object to the folder
;                  |2  - Default item type (integer) for the specified folder. Defined by the Outlook OlItemType enumeration
;                  |3  - StoreID (string) of the store to access the folder by ID
;                  |4  - EntryID (string) of the folder to access the folder by ID
;                  |5  - Display name of the folder
;                  |6  - The path of the selected folder
;                  |7  - Number of unread items in the folder
;                  |8  - Total number of items in the folder
;                  |9  - Address Book Name for a contacts folder
;                  |10 - Determines which views are displayed on the View menu
;                  |11 - Default message class for items in the folder
;                  |12 - Description of the folder
;                  |13 - Determines if the folder will be synchronized with the e-mail server
;                  |14 - Determines if the folder is a Microsoft SharePoint Server folder
;                  |15 - Specifies if the contact items folder will be displayed as an address list in the Outlook Address Book
;                  |16 - Indicates if to display the number of unread messages in the folder or the total number of items in the folder in the Navigation Pane
;                  |17 - Indicates the Web view state for the folder
;                  |18 - URL of the Web page that is assigned with the folder
;                  Failure - Returns "" and sets @error:
;                  |1 - Error accessing specified folder. See @extended for the error code of _OL_AccessFolder
; Author ........: water
; Modified.......:
; Remarks .......: The current folder is the one displayed in the active explorer
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_FolderGet($oOL, $vOL_Folder = "")

	If Not IsObj($vOL_Folder) Then
		If StringStripWS($vOL_Folder, 3) = "" Then $vOL_Folder = $oOL.ActiveExplorer.CurrentFolder
		Local $aOL_Temp = _OL_FolderAccess($oOL, $vOL_Folder)
		If @error Then Return SetError(1, @error, 0)
		$vOL_Folder = $aOL_Temp[1]
	EndIf
	Local $aOL_Folder[19] = [18]
	$aOL_Folder[1] = $vOL_Folder
	$aOL_Folder[2] = $vOL_Folder.DefaultItemType
	$aOL_Folder[3] = $vOL_Folder.StoreID
	$aOL_Folder[4] = $vOL_Folder.EntryID
	$aOL_Folder[5] = $vOL_Folder.Name
	$aOL_Folder[6] = $vOL_Folder.FolderPath
	$aOL_Folder[7] = $vOL_Folder.UnReadItemCount
	$aOL_Folder[8] = $vOL_Folder.Items.Count
	$aOL_Folder[9] = $vOL_Folder.AddressBookName
	$aOL_Folder[10] = $vOL_Folder.CustomViewsOnly
	$aOL_Folder[11] = $vOL_Folder.DefaultMessageClass
	$aOL_Folder[12] = $vOL_Folder.Description
	$aOL_Folder[13] = $vOL_Folder.InAppFolderSyncObject
	$aOL_Folder[14] = $vOL_Folder.IsSharePointFolder
	$aOL_Folder[15] = $vOL_Folder.ShowAsOutlookAB
	$aOL_Folder[16] = $vOL_Folder.ShowItemCount
	$aOL_Folder[17] = $vOL_Folder.WebViewOn
	$aOL_Folder[18] = $vOL_Folder.WebViewURL
	Return $aOL_Folder

EndFunc   ;==>_OL_FolderGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_FolderModify
; Description ...: Modifies the properties of a folder.
; Syntax.........: _OL_FolderModify($oOL, $vOL_Folder[, $sOL_AddressBookName = ""[, $bOL_CustomViewsOnly = Default[, $sOL_Description = ""[, $bOL_InAppFolderSyncObject = Default[, $sOL_Name = ""[, $bOL_ShowAsOutlookAB = Default[, $iOL_ShowItemCount = Default[, $bOL_WebViewOn = Default[, $sOL_WebViewURL = ""]]]]]]]]])
; Parameters ....: $oOL                       - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Folder                - Folder object as returned by _OL_FolderAccess or full name of folder to modify
;                  $sOL_AddressBookName       - Address Book name if the folder represents a contacts folder (default = "" = do not change property)
;                  $bOL_CustomViewsOnly       - True/False. Determines which views are displayed on the view menu (default = keyword "Default" = do not change property)
;                  $sOL_Description           - Description of the folder (default = "" = do not change property)
;                  $bOL_InAppFolderSyncObject - True/False. Determines if the folder will be synchronized with the e-mail server. (default = keyword "Default" = do not change property)
;                  $sOL_Name                  - Display name for the folder (default = "" = do not change property)
;                  $bOL_ShowAsOutlookAB       - True/False. Specifies whether the folder will be displayed as an address list in the Outlook Address Book (folder thas to be a contacts folder) (default = keyword "Default" = do not change property)
;                  $iOL_ShowItemCount         - OlShowItemCount enumeration. Indicates the itemcount to display - if any (default = keyword "Default" = do not change property)
;                  $bOL_WebViewOn             - True/False. Indicates the web view state (default = keyword "Default" = do not change property)
;                  $sOL_WebViewURL            - URL of the Web page for this folder (default = "" = do not change property)
; Return values .: Success - Folder object of the created folder
;                  Failure - Returns 0 and sets @error:
;                  |1 - $vOL_Folder has not been specified
;                  |2 - Error accessing the specified folder. See @extended for errorcode returned by GetFolderFromID
;                  |3 - Error setting propery $sOL_AddressBookName. See @extended for more details
;                  |4 - Error setting propery $bOL_CustomViewsOnly. See @extended for more details
;                  |5 - Error setting propery $sOL_Description. See @extended for more details
;                  |6 - Error setting propery $bOL_InAppFolderSyncObject. See @extended for more details
;                  |7 - Error setting propery $sOL_Name. See @extended for more details
;                  |8 - Error setting propery $bOL_ShowAsOutlookAB. See @extended for more details
;                  |9 - Error setting propery $iOL_ShowItemCount. See @extended for more details
;                  |10 - Error setting propery $bOL_WebViewOn. See @extended for more details
;                  |11 - Error setting propery $sOL_WebViewURL. See @extended for more details
; Author ........: water
; Modified ......:
; Remarks .......: To reset a string property set the corresponding value to " ".
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_FolderModify($oOL, $vOL_Folder, $sOL_AddressBookName = "", $bOL_CustomViewsOnly = Default, $sOL_Description = "", $bOL_InAppFolderSyncObject = Default, $sOL_Name = "", $bOL_ShowAsOutlookAB = Default, $iOL_ShowItemCount = Default, $bOL_WebViewOn = Default, $sOL_WebViewURL = "")

	If Not IsObj($vOL_Folder) Then
		If StringStripWS($vOL_Folder, 3) = "" Then Return SetError(1, 0, 0)
		Local $aFolder = _OL_FolderAccess($oOL, $vOL_Folder)
		If @error Then Return SetError(2, @error, 0)
		$vOL_Folder = $aFolder[1]
	EndIf
	If $sOL_AddressBookName <> "" Then
		$vOL_Folder.AddressBookName = $sOL_AddressBookName
		If @error Then Return SetError(3, @error, 0)
	EndIf
	If $bOL_CustomViewsOnly <> Default Then
		$vOL_Folder.CustomViewsOnly = $bOL_CustomViewsOnly
		If @error Then Return SetError(4, @error, 0)
	EndIf
	If $sOL_Description <> "" Then
		$vOL_Folder.Description = $sOL_Description
		If @error Then Return SetError(5, @error, 0)
	EndIf
	If $bOL_InAppFolderSyncObject <> Default Then
		$vOL_Folder.InAppFolderSyncObject = $bOL_InAppFolderSyncObject
		If @error Then Return SetError(6, @error, 0)
	EndIf
	If $sOL_Name <> "" Then
		$vOL_Folder.Name = $sOL_Name
		If @error Then Return SetError(7, @error, 0)
	EndIf
	If $bOL_ShowAsOutlookAB <> Default Then
		$vOL_Folder.ShowAsOutlookAB = $bOL_ShowAsOutlookAB
		If @error Then Return SetError(8, @error, 0)
	EndIf
	If $iOL_ShowItemCount <> Default Then
		$vOL_Folder.ShowItemCount = $iOL_ShowItemCount
		If @error Then Return SetError(9, @error, 0)
	EndIf
	If $bOL_WebViewOn <> Default Then
		$vOL_Folder.WebViewOn = $bOL_WebViewOn
		If @error Then Return SetError(10, @error, 0)
	EndIf
	If $sOL_WebViewURL <> "" Then
		$vOL_Folder.WebViewURL = $sOL_WebViewURL
		If @error Then Return SetError(11, @error, 0)
	EndIf
	Return $vOL_Folder

EndFunc   ;==>_OL_FolderModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_FolderMove
; Description ...: Moves a folder plus subfolders to a new target folder.
; Syntax.........: _OL_FolderMove($oOL, $vOL_SourceFolder, $vOL_TargetFolder)
; Parameters ....: $oOL              - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_SourceFolder - Folder object as returned by _OL_FolderAccess or full name of folder to move
;                  $vOL_TargetFolder - Folder object as returned by _OL_FolderAccess or full name of folder to move to
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - Error accessing the specified source folder. See @extended for errorcode returned by _OL_FolderAccess
;                  |2 - Error accessing the specified target folder. See @extended for errorcode returned by _OL_FolderAccess
;                  |3 - Source folder has not been specified or is empty
;                  |4 - Target folder has not been specified or is empty
;                  |5 - Source and target folder are the same
;                  |6 - Error moving the folder to the target folder. For details check @extended
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_FolderMove($oOL, $vOL_SourceFolder, $vOL_TargetFolder)

	Local $aOL_Temp
	If Not IsObj($vOL_SourceFolder) Then
		If StringStripWS($vOL_SourceFolder, 3) = "" Then Return SetError(3, 0, 0)
		$aOL_Temp = _OL_FolderAccess($oOL, $vOL_SourceFolder)
		If @error Then Return SetError(1, @error, 0)
		$vOL_SourceFolder = $aOL_Temp[1]
	EndIf
	If Not IsObj($vOL_TargetFolder) Then
		If StringStripWS($vOL_TargetFolder, 3) = "" Then Return SetError(4, 0, 0)
		$aOL_Temp = _OL_FolderAccess($oOL, $vOL_TargetFolder)
		If @error Then Return SetError(2, @error, 0)
		$vOL_TargetFolder = $aOL_Temp[1]
	EndIf
	If $vOL_SourceFolder = $vOL_TargetFolder Then Return SetError(5, 0, 0)
	$vOL_SourceFolder.MoveTo($vOL_TargetFolder)
	If @error Then Return SetError(6, @error, 0)
	Return 1

EndFunc   ;==>_OL_FolderMove

; #FUNCTION# ====================================================================================================================
; Name...........: _OL_FolderRename
; Description ...: Renames a folder.
; Syntax.........: _OL_FolderRename($oOL, $sOL_Folder, $sOL_Name)
; Parameters ....: $oOL        - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Folder - Folder object as returned by _OL_FolderAccess or full name of folder to be renamed
;                  $sOL_Name   - New display name of the folder
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - Error accessing specified folder. See @extended for the error code of _OL_AccessFolder
;                  |2 - Folder could not be renamed. See @extended for COM error code
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_FolderRename($oOL, $vOL_Folder, $sOL_Name)

	If Not IsObj($vOL_Folder) Then
		Local $aOL_Temp = _OL_FolderAccess($oOL, $vOL_Folder)
		If @error Then Return SetError(1, @error, 0)
		$vOL_Folder = $aOL_Temp[1]
	EndIf
	$vOL_Folder.Name = $sOL_Name
	If @error Then Return SetError(2, @error, 0)
	Return 1

EndFunc   ;==>_OL_FolderRename

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_FolderSelectionGet
; Description ...: Gets all items selected in the active explorer (folder).
; Syntax.........: _OL_FolderSelectionGet($oOL)
; Parameters ....: $oOL - Outlook object returned by a preceding call to _OL_Open()
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - Object of the selected item
;                  |1 - EntryID of the selected item
;                  |2 - OlObjectClass constant indicating the object's class
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_FolderSelectionGet($oOL)

	Local $oOL_Selection = $oOL.ActiveExplorer.Selection
	Local $aOL_Selection[$oOL_Selection.Count + 1][3] = [[$oOL_Selection.Count, 2]]
	For $iOL_Index = 1 To $oOL_Selection.Count
		$aOL_Selection[$iOL_Index][0] = $oOL_Selection.Item($iOL_Index)
		$aOL_Selection[$iOL_Index][1] = $oOL_Selection.Item($iOL_Index).EntryId
		$aOL_Selection[$iOL_Index][2] = $oOL_Selection.Item($iOL_Index).Class
	Next
	Return $aOL_Selection

EndFunc   ;==>_OL_FolderSelectionGet

; #FUNCTION# ====================================================================================================================
; Name...........: _OL_FolderSet
; Description ...: Sets a new folder as the current folder.
; Syntax.........: _OL_FolderSet($oOL, $vOL_Folder)
; Parameters ....: $oOL        - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Folder - Folder object as returned by _OL_FolderAccess or full name of folder that will become the new current folder
; Return values .: Success - Object of the folder
;                  Failure - Returns 0 and sets @error:
;                  |1 - Folder has not been specified or is empty
;                  |2 - Error accessing specified folder. See @extended for the error code of _OL_AccessFolder
;                  |3 - Error setting the current folder. See @extended for more error information
; Author ........: water
; Modified.......:
; Remarks .......: The current folder is the one displayed in the active explorer
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_FolderSet($oOL, $vOL_Folder)

	If Not IsObj($vOL_Folder) Then
		If StringStripWS($vOL_Folder, 3) = "" Then Return SetError(1, 0, 0)
		Local $aOL_Temp = _OL_FolderAccess($oOL, $vOL_Folder)
		If @error Then Return SetError(2, @error, 0)
		$vOL_Folder = $aOL_Temp[1]
	EndIf
	$oOL.ActiveExplorer.CurrentFolder = $vOL_Folder
	If @error Then Return SetError(3, @error, 0)
	Return $vOL_Folder

EndFunc   ;==>_OL_FolderSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_FolderTree
; Description ...: Lists all folders and subfolders starting with a specified folder.
; Syntax.........: _OL_FolderTree($oOL, $vOL_Folder[, $iOL_Level = 9999])
; Parameters ....: $oOL        - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Folder - Folder object as returned by _OL_FolderAccess or full name of folder to start
;                  $iOL_Level  - Optional: Number of levels to list (default = 9999).
;                  |1 = just the level specified in $vOL_Folder
;                  |2 = The level specified in $vOL_Folder plus the next level
; Return values .: Success - one-dimensional zero based array with the folderpath of each folder
;                  Failure - Returns "" and sets @error:
;                  |1 - Source folder has not been specified or is empty
;                  |2 - Error accessing a folder. See @extended for errorcode returned by _OL_FolderAccess
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_FolderTree($oOL, $vOL_Folder, $iOL_Level = 9999)

	Local $aOL_Temp, $aOL_FolderTree[1]
	If Not IsObj($vOL_Folder) Then
		If StringStripWS($vOL_Folder, 3) = "" Then Return SetError(1, 0, "")
		$aOL_Temp = _OL_FolderAccess($oOL, $vOL_Folder)
		If @error Then Return SetError(2, @error, "")
		$vOL_Folder = $aOL_Temp[1]
	EndIf
	$aOL_FolderTree[0] = $vOL_Folder.FolderPath
	$iOL_Level = $iOL_Level - 1
	If $iOL_Level > 0 Then
		For $oOL_Folder In $vOL_Folder.Folders
			$aOL_Temp = _OL_FolderTree($oOL, $oOL_Folder, $iOL_Level)
			If @error Then Return SetError(2, @error, "")
			_ArrayConcatenate($aOL_FolderTree, $aOL_Temp)
		Next
	EndIf
	Return $aOL_FolderTree

EndFunc   ;==>_OL_FolderTree

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_Item2Task
; Description ...: Marks an item as a task and assigns a task interval for the item.
; Syntax.........: _OL_Item2Task($oOL, $vOL_Item, $sOL_StoreID, $iOL_Interval)
; Parameters ....: $oOL          - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item     - EntryID or object of the item
;                  $sOL_StoreID  - StoreID where the EntryID is stored. Use the keyword "Default" to use the users mailbox
;                  $iOL_Interval - Time period for which the item is marked as a task. Defined by the $OlMarkInterval Enumeration
; Return values .: Success - Item object
;                  Failure - Returns 0 and sets @error:
;                  |1 - No Outlook item specified
;                  |2 - Item could not be found. EntryID might be wrong. Check @extended for more information
;                  |3 - $iOL_Interval is not a number
;                  |4 - Method MarkAsTask returned an error. Check @extended for more information
; Author ........: water
; Modified ......:
; Remarks .......: This function sets the value of several other properties, depending on the value provided in $iOL_Interval.
;                  For more information about the properties set see the link below (OlMarkInterval Enumeration)
;                  +
;                  To change this or set further properties please call _OL_ItemModify
; Related .......:
; Link ..........: http://msdn.microsoft.com/en-us/library/bb208108(v=office.12).aspx
; Example .......: Yes
; ===============================================================================================================================
Func _OL_Item2Task($oOL, $vOL_Item, $sOL_StoreID, $iOL_Interval)

	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	If Not IsInt($iOL_Interval) Then SetError(3, 0, 0)
	$vOL_Item.MarkAsTask($iOL_Interval)
	If @error Then Return SetError(4, @error, 0)
	Return $vOL_Item

EndFunc   ;==>_OL_Item2Task

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemAttachmentAdd
; Description ...: Adds one or more attachments to an item.
; Syntax.........: _OL_ItemAttachmentAdd($oOL, $vOL_Item, $sOL_StoreID, $vOL_P1 = ""[, $vOL_P2 = ""[, $vOL_P3 = ""[, $vOL_P4 = ""[, $vOL_P5 = ""[, $vOL_P6 = ""[, $vOL_P7 = ""[, $vOL_P8 = ""[, $vOL_P9 = ""[, $vOL_P10 = ""]]]]]]]]])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the item
;                  $sOL_StoreID - StoreID where the EntryID is stored. Use the keyword "Default" to use the users mailbox
;                  $vOL_P1      - The source of the attachment. This can be a file (represented by the full file system path (drive letter or UNC path) with a file name) or an
;                  +Outlook item (EntryId or object) that constitutes the attachment
;                  +or a zero based one-dimensional array with unlimited number of attachments.
;                  |Every attachment parameter can consist of up to 4 sub-parameters separated by commas:
;                  | 1 - Source: The source of the attachment as described above
;                  | 2 - (Optional) Type: The type of the attachment. Can be one of the OlAttachmentType constants (default = $olByValue)
;                  | 3 - (Optional) Position: For RTF format. Position where the attachment should be placed within the body text (default = Beginning of the item)
;                  | 4 - (Optional) DisplayName: For RTF format and Type = $olByValue. Name is displayed in an Inspector object or when viewing the properties of the attachment
;                  $vOL_P2      - Optional: Same as $vOL_P1 but no array is allowed
;                  $vOL_P3      - Optional: Same as $vOL_P2
;                  $vOL_P4      - Optional: Same as $vOL_P2
;                  $vOL_P5      - Optional: Same as $vOL_P2
;                  $vOL_P6      - Optional: Same as $vOL_P2
;                  $vOL_P7      - Optional: Same as $vOL_P2
;                  $vOL_P8      - Optional: Same as $vOL_P2
;                  $vOL_P9      - Optional: Same as $vOL_P2
;                  $vOL_P10     - Optional: Same as $vOL_P2
; Return values .: Success - Item object
;                  Failure - Returns 0 and sets @error:
;                  |1 - No Outlook item specified
;                  |2 - Item could not be found. EntryID might be wrong
;                  |3 - Error adding attachment to the item list. @extended = number of the invalid attachment (zero based)
;                  |4 - Attachment could not be found. @extended = number of the invalid attachment (zero based)
; Author ........: water
; Modified.......:
; Remarks .......: $vOL_P2 to $vOL_P10 will be ignored if $vOL_P1 is an array of attachments.
;                  For more details about sub-parameters 2-4 please check MSDN for the Attachments.Add method
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemAttachmentAdd($oOL, $vOL_Item, $sOL_StoreID, $vOL_P1, $vOL_P2 = "", $vOL_P3 = "", $vOL_P4 = "", $vOL_P5 = "", $vOL_P6 = "", $vOL_P7 = "", $vOL_P8 = "", $vOL_P9 = "", $vOL_P10 = "")

	Local $aOL_Attachments[10]
	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	; Move attachments into an array
	If Not IsArray($vOL_P1) Then
		$aOL_Attachments[0] = $vOL_P1
		$aOL_Attachments[1] = $vOL_P2
		$aOL_Attachments[2] = $vOL_P3
		$aOL_Attachments[3] = $vOL_P4
		$aOL_Attachments[4] = $vOL_P5
		$aOL_Attachments[5] = $vOL_P6
		$aOL_Attachments[6] = $vOL_P7
		$aOL_Attachments[7] = $vOL_P8
		$aOL_Attachments[8] = $vOL_P9
		$aOL_Attachments[9] = $vOL_P10
	Else
		$aOL_Attachments = $vOL_P1
	EndIf
	; add attachments to the item
	For $iOL_Index = 0 To UBound($aOL_Attachments) - 1
		Local $aOL_Temp = StringSplit($aOL_Attachments[$iOL_Index], ",")
		ReDim $aOL_Temp[5] ; Make sure the array has 4 elements (element 2-4 might be empty)
		If StringMid($aOL_Temp[1], 2, 1) = ":" Or StringLeft($aOL_Temp[1], 2) = "\\" Then ; Attachment specified as file (drive letter or UNC path)
			If Not FileExists($aOL_Temp[1]) Then Return SetError(4, $iOL_Index, 0)
		ElseIf Not IsObj($aOL_Temp[1]) Then ; Attachment specified as EntryID
			If StringStripWS($aOL_Attachments[$iOL_Index], 3) = "" Then ContinueLoop
			$aOL_Temp[1] = $oOL.Session.GetItemFromID($aOL_Temp[1], $sOL_StoreID)
			If @error Then Return SetError(4, $iOL_Index, 0)
		EndIf
		If $aOL_Temp[2] = "" Then $aOL_Temp[2] = $olByValue ; The attachment is a copy of the original file
		If $aOL_Temp[3] = "" Then $aOL_Temp[3] = 1 ; The attachment should be placed at the beginning of the message body
		$vOL_Item.Attachments.Add($aOL_Temp[1], $aOL_Temp[2], $aOL_Temp[3], $aOL_Temp[4])
		If @error Then Return SetError(3, $iOL_Index, 0)
	Next
	$vOL_Item.Close(0)
	Return $vOL_Item

EndFunc   ;==>_OL_ItemAttachmentAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemAttachmentDelete
; Description ...: Deletes one or multiple attachments from an item.
; Syntax.........: _OL_ItemAttachmentDelete($oOL, $vOL_Item, $sOL_StoreID, $vOL_P1 = ""[, $vOL_P2 = ""[, $vOL_P3 = ""[, $vOL_P4 = ""[, $vOL_P5 = ""[, $vOL_P6 = ""[, $vOL_P7 = ""[, $vOL_P8 = ""[, $vOL_P9 = ""[, $vOL_P10 = ""]]]]]]]]])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the item
;                  $sOL_StoreID - StoreID where the EntryID is stored. Use the keyword "Default" to use the users mailbox
;                  $vOL_P1      - Number of the attachment to delete from the attachments collection
;                  +or a zero based one-dimensional array with unlimited number of attachments
;                  $vOL_P2      - Optional: Number of the attachment to delete from the attachments collection
;                  $vOL_P3      - Optional: Same as $vOL_P2
;                  $vOL_P4      - Optional: Same as $vOL_P2
;                  $vOL_P5      - Optional: Same as $vOL_P2
;                  $vOL_P6      - Optional: Same as $vOL_P2
;                  $vOL_P7      - Optional: Same as $vOL_P2
;                  $vOL_P8      - Optional: Same as $vOL_P2
;                  $vOL_P9      - Optional: Same as $vOL_P2
;                  $vOL_P10     - Optional: Same as $vOL_P2
; Return values .: Success - Item object
;                  Failure - Returns 0 and sets @error:
;                  |1 - No item specified
;                  |2 - Item could not be found. EntryID might be wrong
;                  |3 - Error removing attachment from the item. @extended = number of the invalid attachment parameter (zero based)
; Author ........: water
; Modified.......:
; Remarks .......: $vOL_P2 to $vOL_P10 will be ignored if $vOL_P1 is an array of numbers
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemAttachmentDelete($oOL, $vOL_Item, $sOL_StoreID, $vOL_P1 = "", $vOL_P2 = "", $vOL_P3 = "", $vOL_P4 = "", $vOL_P5 = "", $vOL_P6 = "", $vOL_P7 = "", $vOL_P8 = "", $vOL_P9 = "", $vOL_P10 = "")

	Local $aOL_Attachments[10]
	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	; Move numbers into an array
	If Not IsArray($vOL_P1) Then
		$aOL_Attachments[0] = $vOL_P1
		$aOL_Attachments[1] = $vOL_P2
		$aOL_Attachments[2] = $vOL_P3
		$aOL_Attachments[3] = $vOL_P4
		$aOL_Attachments[4] = $vOL_P5
		$aOL_Attachments[5] = $vOL_P6
		$aOL_Attachments[6] = $vOL_P7
		$aOL_Attachments[7] = $vOL_P8
		$aOL_Attachments[8] = $vOL_P9
		$aOL_Attachments[9] = $vOL_P10
	Else
		$aOL_Attachments = $vOL_P1
	EndIf
	; Delete attachments from the item
	For $iOL_Index = 0 To UBound($aOL_Attachments) - 1
		If StringStripWS($aOL_Attachments[$iOL_Index], 3) = "" Then ContinueLoop
		$vOL_Item.Attachments.Remove($aOL_Attachments[$iOL_Index])
		If @error Then Return SetError(3, $iOL_Index, 0)
	Next
	$vOL_Item.Close(0)
	Return $vOL_Item

EndFunc   ;==>_OL_ItemAttachmentDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemAttachmentGet
; Description ...: Get a list of attachments of an item.
; Syntax.........: _OL_ItemAttachmentGet($oOL, $vOL_Item[, $sOL_StoreID = Default])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the item
;                  $sOL_StoreID - Optional: StoreID where the EntryID is stored (default = users mailbox)
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - Object to the attachment
;                  |1 - DisplayName: String representing the name, which does not need to be the actual file name, displayed below the icon representing the embedded attachment
;                  |2 - FileName: String representing the file name of the attachment
;                  |3 - PathName: String representing the full path to the linked attached file
;                  |4 - Position: Integer indicating the position of the attachment within the body of the item
;                  |5 - Size: Integer indicating the size (in bytes) of the attachment
;                  |6 - Type: OlAttachmentType constant indicating the type of the specified object
;                  Failure - Returns 0 and sets @error:
;                  |1 - No Outlook item specified
;                  |2 - Item could not be found. EntryID might be wrong
;                  |3 - Item has no attachments
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemAttachmentGet($oOL, $vOL_Item, $sOL_StoreID = Default)

	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	If $vOL_Item.Attachments.Count = 0 Then Return SetError(3, 0, 0)
	Local $aOL_Attachments[$vOL_Item.Attachments.Count + 1][7] = [[$vOL_Item.Attachments.Count, 7]]
	Local $iOL_Index = 1
	For $oOL_Attachment In $vOL_Item.Attachments
		$aOL_Attachments[$iOL_Index][0] = $oOL_Attachment
		$aOL_Attachments[$iOL_Index][1] = $oOL_Attachment.DisplayName
		$aOL_Attachments[$iOL_Index][2] = $oOL_Attachment.FileName
		$aOL_Attachments[$iOL_Index][3] = $oOL_Attachment.PathName
		$aOL_Attachments[$iOL_Index][4] = $oOL_Attachment.Position
		$aOL_Attachments[$iOL_Index][5] = $oOL_Attachment.Size
		$aOL_Attachments[$iOL_Index][6] = $oOL_Attachment.Type
		$iOL_Index += 1
	Next
	Return $aOL_Attachments

EndFunc   ;==>_OL_ItemAttachmentGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemAttachmentSave
; Description ...: Saves a single attachment of an item in the specified path.
; Syntax.........: _OL_ItemAttachmentSave($oOL, $vOL_Item, $sOL_StoreID, $iOL_Attachment, $sOL_Path)
; Parameters ....: $oOL            - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item       - EntryID or object of the item of which to save the attachment
;                  $sOL_StoreID    - StoreID of the source store as returned by _OL_FolderAccess. Use the keyword "Default" to use the users mailbox
;                  $iOL_Attachment - Number of the attachment to save as returned by _OL_ItemAttachmentGet (one based)
;                  $sOL_Path       - Path (drive, directory[, filename]) where to save the item.
;                                    If filename or extension is missing it is set to the filename/extension of the attachment.
;                                    In this case the directory needs a trailing backslash.
;                                    If the directory does not exist it is created
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $sOL_Path is missing
;                  |2 - Specified directory does not exist. It could not be created
;                  |3 - Specified item could not be found
;                  |4 - Output file already exists
;                  |5 - Error saving an attachment. For details check @extended
;                  |6 - No item has been specified
;                  |7 - $sOL_Path not specified completely. Drive, dir, name or extension is missing
;                  |8 - $iOL_Attachment is either not numeric or < 1 or > # of attachments as returned by _OL_ItemAttachmentGet. @extended is the number of attachments
; Author ........: water
; Modified ......:
; Remarks .......: If the file you want the attachment to be saved already exists an error is returned.
;                  _OL_ItemSave saves all attachments but can create distinct filenames by adding a number between 00 and 99 at the end
; Related .......: _OL_ItemSave
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemAttachmentSave($oOL, $vOL_Item, $sOL_StoreID, $iOL_Attachment, $sOL_Path)

	Local $sOL_Drive, $sOL_Dir, $sOL_FName, $sOL_Ext
	If StringStripWS($sOL_Path, 3) = "" Then Return SetError(1, 0, 0)
	_PathSplit($sOL_Path, $sOL_Drive, $sOL_Dir, $sOL_FName, $sOL_Ext)
	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(6, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If Not IsObj($vOL_Item) Then Return SetError(3, 0, 0)
	EndIf
	Local $aOL_Attachments = _OL_ItemAttachmentGet($oOL, $vOL_Item, $sOL_StoreID)
	If Not (IsNumber($iOL_Attachment)) Or $iOL_Attachment < 1 Or $iOL_Attachment > $aOL_Attachments[0][0] Then Return SetError(8, $aOL_Attachments[0][0], 0)
	; Set filename/extension to name/extension of the attachment
	Local $iPos = StringInStr($aOL_Attachments[$iOL_Attachment][2], ".")
	If $sOL_FName = "" Then $sOL_FName = StringLeft($aOL_Attachments[$iOL_Attachment][2], $iPos - 1)
	If $sOL_Ext = "" Then $sOL_Ext = StringMid($aOL_Attachments[$iOL_Attachment][2], $iPos)
	; Replace invalid characters from filename with underscore
	$sOL_FName = StringRegExpReplace($sOL_FName, '[ \/:*?"<>|]', '_')
	If $sOL_Drive = "" Or $sOL_Dir = "" Or $sOL_FName = "" Or $sOL_Ext = "" Then Return SetError(7, 0, 0)
	If Not FileExists($sOL_Drive & $sOL_Dir) Then
		If DirCreate($sOL_Drive & $sOL_Dir) = 0 Then Return SetError(2, 0, 0)
	EndIf
	$sOL_Path = $sOL_Drive & $sOL_Dir & $sOL_FName & $sOL_Ext
	If FileExists($sOL_Path) = 1 Then Return SetError(4, 0, 0)
	; Save attachment
	$aOL_Attachments[$iOL_Attachment][0] .SaveAsFile($sOL_Path)
	If @error Then Return SetError(5, @error, 0)
	Return 1

EndFunc   ;==>_OL_ItemAttachmentSave

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemConflictGet
; Description ...: Get a list of items that are in conflict with the selected item.
; Syntax.........: _OL_ItemConflictGet($oOL, $vOL_Item[, $sOL_StoreID = Default])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the item
;                  $sOL_StoreID - Optional: StoreID where the EntryID is stored (default = users mailbox)
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - Object of the item in conflict
;                  |1 - Class of the object in conflict. Defined by the  OlObjectClass enumeration
;                  |2 - Name of the object in conflict
;                  Failure - Returns 0 and sets @error:
;                  |1 - No Outlook item specified
;                  |2 - Item could not be found. EntryID might be wrong
;                  |3 - Item has no conflicts
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemConflictGet($oOL, $vOL_Item, $sOL_StoreID = Default)

	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	If $vOL_Item.Conflicts.Count = 0 Then Return SetError(3, 0, 0)
	Local $aOL_Conflicts[$vOL_Item.Conflicts.Count + 1][3] = [[$vOL_Item.Conflicts.Count, 3]]
	Local $iOL_Index = 1
	For $oOL_Conflict In $vOL_Item.Conflicts
		$aOL_Conflicts[$iOL_Index][0] = $oOL_Conflict
		$aOL_Conflicts[$iOL_Index][1] = $oOL_Conflict.Type
		$aOL_Conflicts[$iOL_Index][2] = $oOL_Conflict.Name
		$iOL_Index += 1
	Next
	Return $aOL_Conflicts

EndFunc   ;==>_OL_ItemConflictGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemCopy
; Description ...: Copies an item (contact, appointment ...) using the specified EntryID and StoreID.
; Syntax.........: _OL_ItemCopy($oOL, $vOL_Item[, $sOL_StoreID = Default[, $vOL_TargetFolder = ""]])
; Parameters ....: $oOL              - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item         - EntryID or object of the item to copy
;                  $sOL_StoreID      - Optional: StoreID of the source store as returned by _OL_FolderAccess (default = users mailbox)
;                  $vOL_TargetFolder - Optional: Target folder (object) as returned by _OL_FolderAccess or full name of folder
; Return values .: Success - Item object of the copied item
;                  Failure - Returns 0 and sets @error:
;                  |1 - No or an invalid item has been specified
;                  |2 - Error accessing the specified target folder. See @extended for errorcode returned by _OL_FolderAccess
;                  |3 - Source and target folder are of different types
;                  |4 - Error moving the copied item to the target folder. For details check @extended
; Author ........: water
; Modified ......:
; Remarks .......: If $vOL_TargetFolder is omitted the copy is created in the same folder as the source item
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemCopy($oOL, $vOL_Item, $sOL_StoreID = Default, $vOL_TargetFolder = "")

	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(1, @error, 0)
	EndIf
	Local $oOL_SourceFolder = $vOL_Item.Parent
	If Not IsObj($vOL_TargetFolder) Then
		If StringStripWS($vOL_TargetFolder, 3) = "" Then $vOL_TargetFolder = $oOL_SourceFolder
		Local $aOL_Temp = _OL_FolderAccess($oOL, $vOL_TargetFolder)
		If @error Then Return SetError(2, @error, 0)
		$vOL_TargetFolder = $aOL_Temp[1]
	EndIf
	If $oOL_SourceFolder.DefaultItemType <> $vOL_TargetFolder.DefaultItemType Then Return SetError(3, 0, 0)
	Local $vOL_ItemCopied = $vOL_Item.Copy
	$vOL_ItemCopied.Close(0)
	; Move the copied item to another folder
	If $oOL_SourceFolder <> $vOL_TargetFolder Then
		$vOL_ItemCopied = _OL_ItemMove($oOL, $vOL_ItemCopied, $sOL_StoreID, $vOL_TargetFolder)
		If @error Then Return SetError(4, @error, 0)
	EndIf
	Return $vOL_ItemCopied

EndFunc   ;==>_OL_ItemCopy

; #FUNCTION# ====================================================================================================================
; Name...........: _OL_ItemCreate
; Description ...: Create an item.
; Syntax.........: _OL_ItemCreate($oOL, $iOL_ItemType, $vOL_Folder = ""[, $sOL_Template = ""[,$sOL_P1 = ""[, $sOL_P2 = ""[, $sOL_P3 = ""[, $sOL_P4 = ""[, $sOL_P5 = ""[, $sOL_P6 = ""[, $sOL_P7 = ""[, $sOL_P8 = ""[, $sOL_P9 = ""[, $sOL_P10 = ""]]]]]]]]]]])
; Parameters ....: $oOL            - Outlook object returned by a preceding call to _OL_Open()
;                  $iOL_ItemType   - Type of item to create. Is defined by the Outlook OlItemType enumeration
;                  $vOL_Folder     - Optional: Folder object as returned by _OL_FolderAccess or full name of folder where the item will be created.
;                  |If not specified the default folder for the item type specified by $iOL_ItemType will be selected
;                  $sOL_Template   - Optional: Path and file name of the Outlook template for the new item
;                  $sOL_P1         - Optional: Item property in the format: propertyname=propertyvalue
;                  |or a zero based one-dimensional array with unlimited number of properties in the same format
;                  $sOL_P2         - Optional: Item property in the format: propertyname=propertyvalue
;                  $sOL_P3         - Optional: Item property in the format: propertyname=propertyvalue
;                  $sOL_P4         - Optional: Item property in the format: propertyname=propertyvalue
;                  $sOL_P5         - Optional: Item property in the format: propertyname=propertyvalue
;                  $sOL_P6         - Optional: Item property in the format: propertyname=propertyvalue
;                  $sOL_P7         - Optional: Item property in the format: propertyname=propertyvalue
;                  $sOL_P8         - Optional: Item property in the format: propertyname=propertyvalue
;                  $sOL_P9         - Optional: Item property in the format: propertyname=propertyvalue
;                  $sOL_P10        - Optional: Item property in the format: propertyname=propertyvalue
; Return values .: Success - Item object of the created item
;                  Failure - Returns 0 and sets @error:
;                  |1 - Error accessing specified folder. See @extended for errorcode returned by _OL_FolderAccess
;                  |2 - Error moving the Item to the specified folder. See @extended for errorcode returned by _OL_ItemMove
;                  |3 - Property doesn't contain a "=" to separate name and value. @extended = number of property in error (zero based)
;                  |4 - Error creating the item. @extended = error returned by the COM interface
;                  |5 - Invalid or no $iOL_ItemType specified
;                  |6 - Specified template file does not exist
;                  |1xx - Error checking the properties $sOL_P1 to $sOL_P10 as returned by _OL_CheckProperties.
;                  +      @extended is the number of the property in error (zero based)
; Author ........: water
; Modified.......:
; Remarks .......: $sOL_P2 to $sOL_P10 will be ignored if $sOL_P1 is an array of properties
;                  Be sure to specify the properties in correct case e.g. "FirstName" is valid, "Firstname" is invalid
;                  +
;                  If you want to create a meeting request and send it to some attendees you have to create an appointment and set property
;                  +MeetingStatus to one of the OlMeetingStatus enumeration
;                  +
;                  Note: Mails are created in the drafts folder if you do not specify $vOL_Folder
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemCreate($oOL, $iOL_ItemType, $vOL_Folder = "", $sOL_Template = "", $sOL_P1 = "", $sOL_P2 = "", $sOL_P3 = "", $sOL_P4 = "", $sOL_P5 = "", $sOL_P6 = "", $sOL_P7 = "", $sOL_P8 = "", $sOL_P9 = "", $sOL_P10 = "")

	Local $aOL_Properties[10], $iOL_Pos, $oOL_Item
	If Not IsNumber($iOL_ItemType) Then Return SetError(5, 0, 0)
	If $sOL_Template <> "" And Not FileExists($sOL_Template) Then Return SetError(6, 0, 0)
	If Not IsObj($vOL_Folder) Then
		Local $aFolderToAccess = _OL_FolderAccess($oOL, $vOL_Folder, Default, $iOL_ItemType)
		If @error Then Return SetError(1, @error, 0)
		$vOL_Folder = $aFolderToAccess[1]
	EndIf
	If StringStripWS($sOL_Template, 3) = "" Then
		$oOL_Item = $vOL_Folder.Items.Add($iOL_ItemType)
	Else
		$oOL_Item = $oOL.CreateItemFromTemplate($sOL_Template, $vOL_Folder) ; create item based on a template
	EndIf
	If @error Then Return SetError(4, @error, 0)
	; Move property parameters into an array
	If Not IsArray($sOL_P1) Then
		$aOL_Properties[0] = $sOL_P1
		$aOL_Properties[1] = $sOL_P2
		$aOL_Properties[2] = $sOL_P3
		$aOL_Properties[3] = $sOL_P4
		$aOL_Properties[4] = $sOL_P5
		$aOL_Properties[5] = $sOL_P6
		$aOL_Properties[6] = $sOL_P7
		$aOL_Properties[7] = $sOL_P8
		$aOL_Properties[8] = $sOL_P9
		$aOL_Properties[9] = $sOL_P10
	Else
		$aOL_Properties = $sOL_P1
	EndIf
	; Check properties
	If Not _OL_CheckProperties($oOL_Item, $aOL_Properties) Then Return SetError(@error, @extended, 0)
	; Set properties of the Item
	For $iOL_Index = 0 To UBound($aOL_Properties) - 1
		If $aOL_Properties[$iOL_Index] <> "" Then
			$iOL_Pos = StringInStr($aOL_Properties[$iOL_Index], "=")
			If $iOL_Pos <> 0 Then
				$oOL_Item.ItemProperties.Item(StringStripWS(StringLeft($aOL_Properties[$iOL_Index], $iOL_Pos - 1), 3)).value = _
						StringStripWS(StringMid($aOL_Properties[$iOL_Index], $iOL_Pos + 1), 3)
			Else
				Return SetError(3, $iOL_Index, 0)
			EndIf
		EndIf
	Next
	$oOL_Item.Close(0)
	; Mails: Move the Item from the drafts folder to another folder if folder was specified and sourcefolder <> targetfolder
	If IsObj($vOL_Folder) And $vOL_Folder.FolderPath <> $oOL_Item.Parent.FolderPath Then
		$oOL_Item = _OL_ItemMove($oOL, $oOL_Item, $oOL_Item.Parent.StoreID, $vOL_Folder)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	Return $oOL_Item

EndFunc   ;==>_OL_ItemCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemDelete
; Description ...: Delete an item (contact, appointment ...) using the specified EntryID and StoreID.
; Syntax.........: _OL_ItemDelete($oOL, $sOL_EntryId[, $sOL_StoreID = Default)
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the item to delete
;                  $sOL_StoreID - Optional: StoreID where the EntryID is stored (default = the users mailbox)
; Return values .: Success - Item object
;                  Failure - Returns 0 and sets @error:
;                  |1 - No item has been specified
;                  |2 - Item could not be found. EntryID might be wrong
;                  |3 - Item could not be deleted. Please see @extended for more information
; Author ........: water
; Modified ......:
; Remarks .......: To cancel a meeting you have to set property "MeetingStatus" to $olMeetingCanceled and send the meeting again
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemDelete($oOL, $vOL_Item, $sOL_StoreID = Default)

	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	$vOL_Item.Delete
	If @error Then Return SetError(3, @error, 0)
	Return $vOL_Item

EndFunc   ;==>_OL_ItemDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemDisplay
; Description ...: Displays an item (contact, appointment ...) using the specified EntryID and StoreID.
; Syntax.........: _OL_ItemDisplay($oOL, $sOL_EntryId[, $sOL_StoreID = Default[, $iOL_Width = 0[, $iOL_Height = 0[, $iOL_Left = 0[, $iOL_Top = 0[, $iOL_State = $olNormalWindow]]]]]])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the item to display
;                  $sOL_StoreID - Optional: StoreID where the EntryID is stored (default = keyword "Default" = the users mailbox)
;                  $iOL_Width   - Optional: The width of the window in pixel (default = 0 = Use Outlook default)
;                  $iOL_Height  - Optional: The height of the window in pixel (default = 0 = Use Outlook default)
;                  $iOL_Left    - Optional: The left position of the window in pixel (default = 0 = Use Outlook default)
;                  $iOL_Top     - Optional: The top position of the window in pixel (default = 0 = Use Outlook default)
;                  $iOL_State   - Optional: State of the window. Defined by the Outlook OlWindowState enumeration (default = $olNormalWindow)
; Return values .: Success - Object of the Inspector where the item is displayed
;                  Failure - Returns 0 and sets @error:
;                  |1 - No item has been specified
;                  |2 - Item could not be found. EntryID might be wrong
;                  |3 - Item could not be displayed. Please see @extended for more information
;                  |4 - Error setting properties of the window. Please see @extended for more information
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemDisplay($oOL, $vOL_Item, $sOL_StoreID = Default, $iOL_Width = 0, $iOL_Height = 0, $iOL_Left = 0, $iOL_Top = 0, $iOL_State = $olNormalWindow)

	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	$vOL_Item.Display()
	If @error Then Return SetError(3, @error, 0)
	If $iOL_Width > 0 Then $vOL_Item.GetInspector.Width = $iOL_Width
	If $iOL_Height > 0 Then $vOL_Item.GetInspector.Height = $iOL_Height
	If $iOL_Left > 0 Then $vOL_Item.GetInspector.left = $iOL_Left
	If $iOL_Top > 0 Then $vOL_Item.GetInspector.Top = $iOL_Top
	$vOL_Item.GetInspector.WindowState = $iOL_State
	If @error Then Return SetError(4, @error, 0)
	Return $vOL_Item.GetInspector

EndFunc   ;==>_OL_ItemDisplay

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemExport
; Description ...: Exports items from an array to a file.
; Syntax.........: _OL_ItemExport($sOL_Path, $sOL_Delimiter, $sOL_Quote, $iOL_Format, $sOL_Header, $aOL_Data)
; Parameters ....: $sOL_Path      - Drive, Directory, Filename and Extension of the output file
;                  $sOL_Delimiter - Optional: Fieldseparator (default = , (comma))
;                  $sOL_Quote     - Optional: Quote character (default = " (double quote))
;                  $iOL_Format    - Character encoding of file:
;                  |0 or 1 - ASCII writing
;                  |2      - Unicode UTF16 Little Endian writing (with BOM)
;                  |3      - Unicode UTF16 Big Endian writing (with BOM)
;                  |4      - Unicode UTF8 writing (with BOM)
;                  |5      - Unicode UTF8 writing (without BOM)
;                  $sOL_Header    - Header line with comma separated list of properties to export
;                  $aOL_Data      - 1-based two-dimensional array
; Return values .: Success - Number of records exported
;                  Failure - Returns 0 and sets @error:
;                  |1 - Parameter $sOL_Path is empty
;                  |2 - File $sOL_Path already exists
;                  |3 - $iOL_Type is not numeric or an invalid number
;                  |4 - $sOL_Header is empty
;                  |5 - $aOL_Data is empty or not a two-dimensional array
;                  |6 - Error writing header line to file $sOL_Path. Please see @extended for error of function _WriteCSV
;                  |7 - Error writing data lines to file $sOL_Path. Please see @extended for error of function _WriteCSV
; Author ........: water
; Modified ......:
; Remarks .......: Fill the array with data using _OL_ItemFind
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemExport($sOL_Path, $sOL_Delimiter, $sOL_Quote, $iOL_Format, $sOL_Header, ByRef $aOL_Data)

	If StringStripWS($sOL_Path, 3) = "" Then Return SetError(1, 0, 0)
	If FileExists($sOL_Path) Then Return SetError(2, 0, 0)
	If Not IsNumber($iOL_Format) Or $iOL_Format > 5 Then Return SetError(3, 0, 0)
	If StringStripWS($sOL_Header, 3) = "" Then Return SetError(4, 0, 0)
	If Not IsArray($aOL_Data) Or UBound($aOL_Data, 0) <> 2 Then Return SetError(5, 0, 0)
	If $sOL_Delimiter = "" Or IsKeyword($sOL_Delimiter) Then $sOL_Delimiter = ","
	If $sOL_Quote = "" Or IsKeyword($sOL_Quote) Then $sOL_Quote = '"'
	; Write header to file
	Local $aOL_HeaderSplit = StringSplit($sOL_Header, ",")
	Local $aOL_HeaderTab[2][$aOL_HeaderSplit[0]] = [[1, $aOL_HeaderSplit[0]]]
	For $iIndex = 1 To $aOL_HeaderSplit[0]
		$aOL_HeaderTab[1][$iIndex - 1] = $aOL_HeaderSplit[$iIndex]
	Next
	Local $iOL_Result = _WriteCSV($sOL_Path, $aOL_HeaderTab, $sOL_Delimiter, $sOL_Quote, $iOL_Format)
	If @error Then Return SetError(6, @error, 0)
	; Write data to file
	$iOL_Result = _WriteCSV($sOL_Path, $aOL_Data, $sOL_Delimiter, $sOL_Quote, $iOL_Format)
	If @error Then Return SetError(7, @error, 0)
	Return $iOL_Result

EndFunc   ;==>_OL_ItemExport

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemFind
; Description ...: Find items (contacts, appointments ...) returning an array of all specified properties.
; Syntax.........: _OL_ItemFind($oOL, $vOL_Folder[, $iOL_ObjectClass = Default[, $sOL_Restrict = ""[, $sOL_SearchName = ""[, $sOL_SearchValue = ""[, $sOL_ReturnProperties = ""[, $sOL_Sort = ""[, $iOL_Flags = 0[, $sOL_WarningClick = ""]]]]]]]])
; Parameters ....: $oOL                  - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Folder           - Folder object as returned by _OL_FolderAccess or full name of folder where the search will be started.
;                  +If you want to search a default folder you have to specify the folder object.
;                  $iOL_ObjectClass      - Optional: Class of items to search for. Defined by the Outlook OlObjectClass enumeration (default = Default = $olContact)
;                  $sOL_Restrict         - Optional: Filter text to restrict number of items returned (exact match). For details please see Remarks
;                  $sOL_SearchName       - Optional: Name of the property to search for (without brackets)
;                  $sOL_SearchValue      - Optional: String value of the property to search for (partial match)
;                  $sOL_ReturnProperties - Optional: Comma separated list of properties to return (default = depending on $iOL_ObjectClass. Please see Remarks)
;                  $sOL_Sort             - Optional: Property to sort the result on plus optional flag to sort descending (default = None). E.g. "[Subject], True" sorts the result descending on the subject
;                  $iOL_Flags            - Optional: Flags to set different processing options. Can be a combination of the following:
;                  |  1: Subfolders will be included
;                  |  2: Row 1 contains column headings. Therefore the number of rows/columns in the table has to be calculated using UBound
;                  |  4: Just return the number of records. You don't get an array, just a single integer denoting the total number of records found
;                  $sOL_WarningClick     - Optional: The Entire SearchString to 'OutlookWarning2.exe' (default = None)
; Return values .: Success - One based two-dimensional array with the properties specified by $sOL_ReturnProperties
;                  Failure - Returns "" and sets @error:
;                  |1 - You have to specifiy $sOL_SearchName AND $sOL_SearchValue or none of them
;                  |2 - $sOL_WarningClick not found
;                  |3 - Error accessing the specified folder. See @extended for errorcode returned by _OL_FolderAccess
;                  |4 - Error accessing specified property. For details check @extended
;                  |1xx - Error checking $sOL_ReturnProperties as returned by _OL_CheckProperties. @extended is the number of the property in error (zero based)
; Author ........: water
; Modified ......:
; Remarks .......: Be sure to specify the values in $sOL_ReturnProperties and $sOL_SearchName in correct case e.g. "FirstName" is valid, "Firstname" is invalid
;+
;                  $sOL_Restrict: Filter can be a Jet query or a DASL query with the @SQL= prefix. Jet query language syntax:
;                  Restrict filter:  Filter LogicalOperator Filter ...
;                  LogicalOperator:  And, Or, Not. Use ( and ) to change the processing order
;                  Filter:           "[property] operator 'value'" or '[property] operator "value"'
;                  Operator:         <, >, <=, >=, <>, =
;                  Example:          "[Start]='2011-02-21 08:00' And [End]='2011-02-21 10:00' And [Subject]='Test'"
;                  See: http://msdn.microsoft.com/en-us/library/cc513841.aspx              - "Searching Outlook Data"
;                       http://msdn.microsoft.com/en-us/library/bb220369(v=office.12).aspx - "Items.Restrict Method"
;+
;                  N.B.: Pass time as HH:MM, HH:MM:SS is invalid and returns no result
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemFind($oOL, $vOL_Folder, $iOL_ObjectClass = Default, $sOL_Restrict = "", $sOL_SearchName = "", $sOL_SearchValue = "", $sOL_ReturnProperties = "", $sOL_Sort = "", $iOL_Flags = 0, $sOL_WarningClick = "")

	Local $bOL_Checked = False, $oOL_Items, $aOL_Temp, $iOL_RecCounter = 0, $iOL_Counter = 0
	If $sOL_WarningClick <> "" Then
		If FileExists($sOL_WarningClick) = 0 Then Return SetError(2, 0, "")
		Run($sOL_WarningClick)
	EndIf
	If $iOL_ObjectClass = Default Then $iOL_ObjectClass = $olContact ; Set Default ObjectClass
	; Set default return properties depending on the class of items
	If StringStripWS($sOL_ReturnProperties, 3) = "" Then
		Switch $iOL_ObjectClass
			Case $olContact
				$sOL_ReturnProperties = "FirstName,LastName,Email1Address,Email2Address,MobileTelephoneNumber"
			Case $olDistributionList
				$sOL_ReturnProperties = "Subject,Body,MemberCount"
			Case $olNote, $olMail
				$sOL_ReturnProperties = "Subject,Body,CreationTime,LastModificationTime,Size"
			Case Else
		EndSwitch
	EndIf
	If Not IsObj($vOL_Folder) Then
		$aOL_Temp = _OL_FolderAccess($oOL, $vOL_Folder)
		If @error Then Return SetError(3, @error, "")
		$vOL_Folder = $aOL_Temp[1]
	EndIf
	If ($sOL_SearchName <> "" And $sOL_SearchValue = "") Or ($sOL_SearchName = "" And $sOL_SearchValue <> "") Then Return SetError(1, 0, "")
	Local $aOL_ReturnProperties = StringSplit(StringStripWS($sOL_ReturnProperties, 8), ",")
	Local $iOL_Index = $aOL_ReturnProperties[0]
	If $aOL_ReturnProperties[0] < 2 Then $iOL_Index = 2
	Local $aOL_Items[$vOL_Folder.Items.Count + 1][$iOL_Index] = [[0, $aOL_ReturnProperties[0]]]
	If StringStripWS($sOL_Restrict, 3) = "" Then
		$oOL_Items = $vOL_Folder.Items
	Else
		$oOL_Items = $vOL_Folder.Items.Restrict($sOL_Restrict)
	EndIf
	If BitAND($iOL_Flags, 4) <> 4 And $sOL_Sort <> "" Then
		$aOL_Temp = StringSplit($sOL_Sort, ",")
		If $aOL_Temp[0] = 1 Then
			$oOL_Items.Sort($sOL_Sort)
		Else
			$oOL_Items.Sort($aOL_Temp[1], True)
		EndIf
	EndIf
	For $oOL_Item In $oOL_Items
		If $oOL_Item.Class <> $iOL_ObjectClass Then ContinueLoop
		; Get all properties of first item and check for existance and correct case
		If BitAND($iOL_Flags, 4) <> 4 And Not $bOL_Checked Then
			If Not _OL_CheckProperties($oOL_Item, $aOL_ReturnProperties, 1) Then Return SetError(@error, @extended, "")
			$bOL_Checked = True
		EndIf
		If $sOL_SearchName <> "" And StringInStr($oOL_Item.ItemProperties.Item($sOL_SearchName).value, $sOL_SearchValue) = 0 Then ContinueLoop
		; Fill array with the specified properties
		$iOL_Counter += 1
		If BitAND($iOL_Flags, 4) <> 4 Then
			For $iOL_Index = 1 To $aOL_ReturnProperties[0]
				$aOL_Items[$iOL_Counter][$iOL_Index - 1] = $oOL_Item.ItemProperties.Item($aOL_ReturnProperties[$iOL_Index]).value
				If @error Then Return SetError(4, @error, "")
				If BitAND($iOL_Flags, 2) = 2 And $iOL_Counter = 1 Then $aOL_Items[0][$iOL_Index - 1] = $oOL_Item.ItemProperties.Item($aOL_ReturnProperties[$iOL_Index]).name
			Next
		EndIf
		If BitAND($iOL_Flags, 4) <> 4 And BitAND($iOL_Flags, 2) <> 2 Then $aOL_Items[0][0] = $iOL_Counter
	Next
	If BitAND($iOL_Flags, 4) = 4 Then
		$iOL_RecCounter += $oOL_Items.Count
		; Process subfolders
		If BitAND($iOL_Flags, 1) = 1 Then
			For $vOL_Folder In $vOL_Folder.Folders
				$iOL_RecCounter += _OL_ItemFind($oOL, $vOL_Folder, $iOL_ObjectClass, $sOL_Restrict, $sOL_SearchName, $sOL_SearchValue, $sOL_ReturnProperties, $sOL_Sort, $iOL_Flags, $sOL_WarningClick)
			Next
		EndIf
		Return $iOL_RecCounter
	Else
		ReDim $aOL_Items[$iOL_Counter + 1][$aOL_ReturnProperties[0]] ; Process subfolders
		If BitAND($iOL_Flags, 1) = 1 Then
			For $vOL_Folder In $vOL_Folder.Folders
				$aOL_Temp = _OL_ItemFind($oOL, $vOL_Folder, $iOL_ObjectClass, $sOL_Restrict, $sOL_SearchName, $sOL_SearchValue, $sOL_ReturnProperties, $sOL_Sort, $iOL_Flags, $sOL_WarningClick)
				_OL_ArrayConcatenate($aOL_Items, $aOL_Temp, $iOL_Flags)
			Next
		EndIf
		Return $aOL_Items
	EndIf

EndFunc   ;==>_OL_ItemFind

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemForward
; Description ...: Forwards an item (contact, appointment ...) using the specified EntryID and StoreID.
; Syntax.........: _OL_ItemForward($oOL, $vOL_Item, $sOL_StoreID, $iOL_Type)
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the item
;                  $sOL_StoreID - StoreID where the EntryID is stored. Use "Default" to access the users mailbox
;                  $iOL_Type    - Type of forwarded item. Valid values are:
;                  |0 - $iOL_Type is ignored for mail items
;                  |1 - ForwardAsVcal: Item is forwarded as virtual calendar item. Valid for appointment and contact items
;                  |    In Outlook 2002 a contact is forwarded in the VCard format
;                  |2 - ForwardAsBusinessCard: Item is forwarded as Electronic Business Card (EBC). Valid for contact items
; Return values .: Success - Object of the forwarded item
;                  Failure - Returns 0 and sets @error:
;                  |1 - No item has been specified
;                  |2 - Item could not be found. EntryID might be wrong
;                  |3 - Item could not be forwarded. @extended = error returned by .Forward
;                  |4 - Item could not be saved. @extended = error returned by .Close
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemForward($oOL, $vOL_Item, $sOL_StoreID, $iOL_Type)

	Local $vOL_ItemForward
	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	; Mail: Simple forward
	; Appointment: ForwardAsVcal
	; Contact: ForwardAsVcal (Outlook 2002 as Vcard) or ForwardAsBusinessCard
	If $vOL_Item.Class = $olMail Then
		$vOL_ItemForward = $vOL_Item.Forward
	ElseIf $vOL_Item.Class = $olContact Then
		If $iOL_Type = 1 Then
			If $vOL_Item.OutlookVersion = "10.0" Then
				$vOL_ItemForward = $vOL_Item.ForwardAsVcard
			Else
				$vOL_ItemForward = $vOL_Item.ForwardAsVcal
			EndIf
		EndIf
		If $iOL_Type = 2 Then $vOL_ItemForward = $vOL_Item.ForwardAsBusinessCard
	Else
		$vOL_ItemForward = $vOL_Item.ForwardAsVcal
	EndIf
	If @error Then Return SetError(3, @error, 0)
	$vOL_Item.Close(0)
	If @error Then Return SetError(4, @error, 0)
	Return $vOL_ItemForward

EndFunc   ;==>_OL_ItemForward

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemGet
; Description ...: Get all or selected properties of an item (contact, appointment ...) using the specified EntryID and StoreID.
; Syntax.........: _OL_ItemGet($oOL, $vOL_Item[, $sOL_StoreID = Default[, $sOL_Properties = ""]])
; Parameters ....: $oOL            - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item       - EntryID or object of the item
;                  $sOL_StoreID    - Optional: StoreID where the EntryID is stored (default = keyword "Default" = the users mailbox)
;                  $sOL_Properties - Optional: Comma separated list of properties to return (default = "" = return all properties)
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - Name of the property
;                  |1 - Value of the property
;                  |2 - Type of the property. Defined by the Outlook OlUserPropertyType enumeration
;                  Failure - Returns "" and sets @error:
;                  |1 - No item has been specified
;                  |2 - Item could not be found. EntryID might be wrong
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemGet($oOL, $vOL_Item, $sOL_StoreID = Default, $sOL_Properties = "")

	Local $vOL_Value
	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, "")
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, "")
	EndIf
	$sOL_Properties = "," & StringReplace($sOL_Properties, " ", "") & ","
	Local $aOL_Properties[$vOL_Item.ItemProperties.Count + 1][3] = [[$vOL_Item.ItemProperties.Count, 3]]
	Local $iOL_Counter = 1
	For $oOL_Property In $vOL_Item.ItemProperties
		If Not ($sOL_Properties = ",," Or StringInStr($sOL_Properties, "," & $oOL_Property.name & ",") > 0) Then ContinueLoop
		$aOL_Properties[$iOL_Counter][0] = $oOL_Property.name
		$aOL_Properties[$iOL_Counter][2] = $oOL_Property.type
		Switch $oOL_Property.type
			Case $olKeywords
				$vOL_Value = $oOL_Property.value
				$aOL_Properties[$iOL_Counter][1] = _ArrayToString($vOL_Value)
			Case Else
				$aOL_Properties[$iOL_Counter][1] = $oOL_Property.value
		EndSwitch
		$iOL_Counter += 1
	Next
	ReDim $aOL_Properties[$iOL_Counter][UBound($aOL_Properties, 2)]
	$aOL_Properties[0][0] = UBound($aOL_Properties, 1) - 1
	_ArraySort($aOL_Properties, 0, 1)
	Return $aOL_Properties

EndFunc   ;==>_OL_ItemGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemImport
; Description ...: Import items from a file.
; Syntax.........: _OL_ItemImport($oOL, $sOL_Path, $sOL_Delimiters, $sOL_Quote, $iOL_Format, $vOL_Folder, $iOL_ItemType)
; Parameters ....: $oOL            - Outlook object
;                  $sOL_Path       - Path (drive, directory, filename) where the data to be imported is stored
;                  $sOL_Delimiters - Optional: Fieldseparators of CSV, multiple are allowed (default = ,;)
;                  $sOL_Quote      - Optional: Character to quote strings (default = ")
;                  $iOL_Format     - Character encoding of file:
;                  |0 or 1 - ASCII writing
;                  |2      - Unicode UTF16 Little Endian writing (with BOM)
;                  |3      - Unicode UTF16 Big Endian writing (with BOM)
;                  |4      - Unicode UTF8 writing (with BOM)
;                  |5      - Unicode UTF8 writing (without BOM)
;                  $vOL_Folder   - Folder object as returned by _OL_FolderAccess or full name of folder where the objects will be stored
;                  $iOL_ItemType - Type of the items that will be created in the $vOL_Folder. Defined by the Outlook OlItemType enumeration
; Return values .: Success - Number of records imported
;                  Failure - Returns 0 and sets @error:
;                  |1 - Parameter $sOL_Path is empty
;                  |2 - File $sOL_Path does not exist
;                  |3 - $vOL_Folder is empty
;                  |4 - $iOL_ItemType is not numeric
;                  |5 - Error processing input file $sOL_Path. Please see @extended for the returncode of _ParseCSV
;                  |6 - Error accessing folder $vOL_Folder. Please see @extended for more information
;                  |7 - Error creating item in folder $vOL_Folder. Please see @extended for more information
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemImport($oOL, $sOL_Path, $sOL_Delimiters, $sOL_Quote, $iOL_Format, $vOL_Folder, $iOL_ItemType)

	If StringStripWS($sOL_Path, 3) = "" Then Return SetError(1, 0, 0)
	If Not FileExists($sOL_Path) Then Return SetError(2, 0, 0)
	If Not IsObj($vOL_Folder) Then
		If StringStripWS($vOL_Folder, 3) = "" Then Return SetError(3, 0, 0)
		Local $aOL_Temp = _OL_FolderAccess($oOL, $vOL_Folder)
		If @error Then Return SetError(6, @error, "")
		$vOL_Folder = $aOL_Temp[1]
	EndIf
	If Not IsNumber($iOL_ItemType) Then Return SetError(4, 0, 0)
	Local $aOL_Data, $sOL_String, $aOL_ItemData
	$aOL_Data = _ParseCSV($sOL_Path, $sOL_Delimiters, $sOL_Quote, $iOL_Format)
	If @error Then Return SetError(5, @error, 0)
	For $iIndex1 = 1 To UBound($aOL_Data, 1) - 1
		$sOL_String = ""
		For $iIndex2 = 0 To UBound($aOL_Data, 2) - 1
			$sOL_String = $sOL_String & "|" & $aOL_Data[0][$iIndex2] & "=" & $aOL_Data[$iIndex1][$iIndex2]
		Next
		$aOL_ItemData = StringSplit($sOL_String, "|", 2)
		$sOL_String = StringMid($sOL_String, 2) ; Get rid of first |
		_OL_ItemCreate($oOL, $iOL_ItemType, $vOL_Folder, "", $aOL_ItemData)
		If @error Then Return SetError(7, @error, 0)
	Next
	Return UBound($aOL_Data, 1) - 1

EndFunc   ;==>_OL_ItemImport

; #FUNCTION# ====================================================================================================================
; Name...........: _OL_ItemModify
; Description ...: Modifies an item by setting the specified properties to the specified values.
; Syntax.........: _OL_ItemModify($oOL, $vOL_Item[, $oOL_StoreID = Default, $sOL_P1 = ""[, $sOL_P2 = ""[, $sOL_P3 = ""[, $sOL_P4 = ""[, $sOL_P5 = ""[, $sOL_P6 = ""[, $sOL_P7 = ""[, $sOL_P8 = ""[, $sOL_P9 = ""[, $sOL_P10 = ""]]]]]]]]]])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the item
;                  $sOL_StoreID - Optional: StoreID where the EntryID is stored. Use the keyword "Default" to use the users mailbox
;                  $sOL_P1      - Property to modify in the format: propertyname=propertyvalue
;                  +or a zero based one-dimensional array with unlimited number of properties in the same format
;                  $sOL_P2      - Optional: Property to modify in the format: propertyname=propertyvalue
;                  $sOL_P3      - Optional: Property to modify in the format: propertyname=propertyvalue
;                  $sOL_P4      - Optional: Property to modify in the format: propertyname=propertyvalue
;                  $sOL_P5      - Optional: Property to modify in the format: propertyname=propertyvalue
;                  $sOL_P6      - Optional: Property to modify in the format: propertyname=propertyvalue
;                  $sOL_P7      - Optional: Property to modify in the format: propertyname=propertyvalue
;                  $sOL_P8      - Optional: Property to modify in the format: propertyname=propertyvalue
;                  $sOL_P9      - Optional: Property to modify in the format: propertyname=propertyvalue
;                  $sOL_P10     - Optional: Property to modify in the format: propertyname=propertyvalue
; Return values .: Success - Item object
;                  Failure - Returns 0 and sets @error:
;                  |1 - No item has been specified
;                  |2 - Item could not be found. EntryID might be wrong
;                  |3 - Property doesn't contain a "=" to separate name and value. @extended = number of property in error (zero based)
;                  |4 - Item could not be saved. @extended = error returned by .save
;                  |1xx - Error checking the properties $sOL_P1 to $sOL_P10 as returned by _OL_CheckProperties.
;                  +      @extended is the number of the property in error (zero based)
; Author ........: water
; Modified.......:
; Remarks .......: $sOL_P2 to $sOL_P10 will be ignored if $sOL_P1 is an array of properties
;                  Be sure to specify the properties in correct case e.g. "FirstName" is valid, "Firstname" is invalid
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemModify($oOL, $vOL_Item, $sOL_StoreID, $sOL_P1, $sOL_P2 = "", $sOL_P3 = "", $sOL_P4 = "", $sOL_P5 = "", $sOL_P6 = "", $sOL_P7 = "", $sOL_P8 = "", $sOL_P9 = "", $sOL_P10 = "")

	Local $aOL_Properties[10], $iOL_Pos
	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	; Move property parameters into an array
	If Not IsArray($sOL_P1) Then
		$aOL_Properties[0] = $sOL_P1
		$aOL_Properties[1] = $sOL_P2
		$aOL_Properties[2] = $sOL_P3
		$aOL_Properties[3] = $sOL_P4
		$aOL_Properties[4] = $sOL_P5
		$aOL_Properties[5] = $sOL_P6
		$aOL_Properties[6] = $sOL_P7
		$aOL_Properties[7] = $sOL_P8
		$aOL_Properties[8] = $sOL_P9
		$aOL_Properties[9] = $sOL_P10
	Else
		$aOL_Properties = $sOL_P1
	EndIf
	; Check properties
	If Not _OL_CheckProperties($vOL_Item, $aOL_Properties) Then Return SetError(@error, @extended, "")
	; Set properties of the item
	For $iOL_Index = 0 To UBound($aOL_Properties) - 1
		If $aOL_Properties[$iOL_Index] <> "" Then
			$iOL_Pos = StringInStr($aOL_Properties[$iOL_Index], "=")
			If $iOL_Pos <> 0 Then
				$vOL_Item.ItemProperties.Item(StringStripWS(StringLeft($aOL_Properties[$iOL_Index], $iOL_Pos - 1), 3)).value = StringStripWS(StringMid($aOL_Properties[$iOL_Index], $iOL_Pos + 1), 3)
			Else
				Return SetError(3, $iOL_Index, 0)
			EndIf
		EndIf
	Next
	$vOL_Item.Save
	If @error Then Return SetError(4, @error, 0)
	Return $vOL_Item

EndFunc   ;==>_OL_ItemModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemMove
; Description ...: Moves an item (contact, appointment ...) using the specified EntryID and StoreID.
; Syntax.........: _OL_ItemMove($oOL, $vOL_Item, $sOL_StoreID, $vOL_TargetFolder)
; Parameters ....: $oOL              - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item         - EntryID or object of the item to move
;                  $sOL_StoreID      - StoreID of the source store as returned by _OL_FolderAccess. Use "Default" to access the users mailbox
;                  $vOL_TargetFolder - Target folder object as returned by _OL_FolderAccess or full name of folder
; Return values .: Success - Item object of the moved item
;                  Failure - Returns 0 and sets @error:
;                  |1 - Error accessing the specified target folder. See @extended for errorcode returned by _OL_FolderAccess
;                  |2 - Source and target folder are of different types
;                  |3 - Source and target folder are the same
;                  |4 - Target folder has not been specified or is empty
;                  |5 - Error moving the item to the target folder. For details check @extended
;                  |6 - No or an invalid item has been specified
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemMove($oOL, $vOL_Item, $sOL_StoreID, $vOL_TargetFolder)

	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(6, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(6, @error, 0)
	EndIf
	Local $oOL_SourceFolder = $vOL_Item.Parent
	If Not IsObj($vOL_TargetFolder) Then
		If StringStripWS($vOL_TargetFolder, 3) = "" Then Return SetError(4, 0, 0)
		Local $aOL_Temp = _OL_FolderAccess($oOL, $vOL_TargetFolder)
		If @error Then Return SetError(1, @error, 0)
		$vOL_TargetFolder = $aOL_Temp[1]
	EndIf
	If $oOL_SourceFolder = $vOL_TargetFolder Then Return SetError(3, 0, 0)
	If $oOL_SourceFolder.DefaultItemType <> $vOL_TargetFolder.DefaultItemType Then Return SetError(2, 0, 0)
	Local $ol_ItemMoved = $vOL_Item.Move($vOL_TargetFolder)
	If @error Then Return SetError(5, @error, 0)
	Return $ol_ItemMoved

EndFunc   ;==>_OL_ItemMove

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemPrint
; Description ...: Prints an item (contact, appointment ...) using all the default settings.
; Syntax.........: _OL_ItemPrint($oOL, $vOL_Item, $sOL_StoreID)
; Parameters ....: $oOL              - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item         - EntryID or object of the item to print
;                  $sOL_StoreID      - Optional: StoreID of the source store as returned by _OL_FolderAccess (default = keyword "Default" = the users mailbox)
; Return values .: Success - Item object of the printed item
;                  Failure - Returns 0 and sets @error:
;                  |1 - No item has been specified
;                  |2 - Item could not be found. EntryId/StoreId might be invalid. For details please check @extended
;                  |3 - Error printing the specified item. For details please check @extended
; Author ........: water
; Modified ......:
; Remarks .......: Item is printed on the default printer with default settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemPrint($oOL, $vOL_Item, $sOL_StoreID = Default)

	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	Local $ol_ItemPrinted = $vOL_Item.PrintOut()
	If @error Then Return SetError(3, @error, 0)
	Return $ol_ItemPrinted

EndFunc   ;==>_OL_ItemPrint

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemRecipientAdd
; Description ...: Adds one or multiple recipients to an item.
; Syntax.........: _OL_ItemRecipientAdd($oOL, $vOL_Item, $sOL_StoreID, $iOL_Type, $vOL_P1 = ""[, $vOL_P2 = ""[, $vOL_P3 = ""[, $vOL_P4 = ""[, $vOL_P5 = ""[, $vOL_P6 = ""[, $vOL_P7 = ""[, $vOL_P8 = ""[, $vOL_P9 = ""[, $vOL_P10 = ""]]]]]]]]])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the item item
;                  $sOL_StoreID - StoreID where the EntryID is stored. Use the keyword "Default" to use the users mailbox
;                  $iOL_Type    - Integer representing the type of recipient. For details see Remarks
;                  $vOL_P1      - Recipient to add to the item. Either a recipient object or the recipients name to be resolved
;                  +or a zero based one-dimensional array with unlimited number of recipients
;                  $vOL_P2      - Optional: recipient to add to the item. Either a recipient object or the recipients name to be resolved
;                  $vOL_P3      - Optional: Same as $vOL_P2
;                  $vOL_P4      - Optional: Same as $vOL_P2
;                  $vOL_P5      - Optional: Same as $vOL_P2
;                  $vOL_P6      - Optional: Same as $vOL_P2
;                  $vOL_P7      - Optional: Same as $vOL_P2
;                  $vOL_P8      - Optional: Same as $vOL_P2
;                  $vOL_P9      - Optional: Same as $vOL_P2
;                  $vOL_P10     - Optional: Same as $vOL_P2
; Return values .: Success - Item object
;                  Failure - Returns 0 and sets @error:
;                  |1 - No item specified
;                  |2 - Item could not be found. EntryID might be wrong
;                  |3 - Error adding recipient to the item. @extended = number of the invalid recipient (zero based)
;                  |4 - Recipient name could not be resolved. @extended = number of the invalid recipient (zero based)
;                  |5 - $iOL_Type is missing or not a number
; Author ........: water
; Modified.......:
; Remarks .......: $vOL_P2 to $vOL_P10 will be ignored if $vOL_P1 is an array of recipients
;                  +
;                  Valid $iOL_Type parameters:
;                  MailItem recipient: one of the following OlMailRecipientType constants: olBCC, olCC, olOriginator, or olTo
;                  MeetingItem recipient: one of the following OlMeetingRecipientType constants: olOptional, olOrganizer, olRequired, or olResource
;                  TaskItem recipient: one of the following OlTaskRecipientType constants: olFinalStatus, or olUpdate
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemRecipientAdd($oOL, $vOL_Item, $sOL_StoreID, $iOL_Type, $vOL_P1, $vOL_P2 = "", $vOL_P3 = "", $vOL_P4 = "", $vOL_P5 = "", $vOL_P6 = "", $vOL_P7 = "", $vOL_P8 = "", $vOL_P9 = "", $vOL_P10 = "")

	Local $oOL_Recipient, $aOL_Recipients[10], $oOL_TempRecipient
	If Not IsNumber($iOL_Type) Then Return SetError(5, 0, 0)
	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	; Move members into an array
	If Not IsArray($vOL_P1) Then
		$aOL_Recipients[0] = $vOL_P1
		$aOL_Recipients[1] = $vOL_P2
		$aOL_Recipients[2] = $vOL_P3
		$aOL_Recipients[3] = $vOL_P4
		$aOL_Recipients[4] = $vOL_P5
		$aOL_Recipients[5] = $vOL_P6
		$aOL_Recipients[6] = $vOL_P7
		$aOL_Recipients[7] = $vOL_P8
		$aOL_Recipients[8] = $vOL_P9
		$aOL_Recipients[9] = $vOL_P10
	Else
		$aOL_Recipients = $vOL_P1
	EndIf
	; add recipients to the item
	For $iOL_Index = 0 To UBound($aOL_Recipients) - 1
		; recipient is an object = recipient name already resolved
		If IsObj($aOL_Recipients[$iOL_Index]) Then
			$vOL_Item.Recipients.Add($aOL_Recipients[$iOL_Index])
			If @error Then Return SetError(3, $iOL_Index, 0)
		Else
			If StringStripWS($aOL_Recipients[$iOL_Index], 3) = "" Then ContinueLoop
			$oOL_Recipient = $oOL.Session.CreateRecipient($aOL_Recipients[$iOL_Index])
			If @error Or Not IsObj($oOL_Recipient) Then Return SetError(4, $iOL_Index, 0)
			$oOL_Recipient.Resolve
			If @error Or Not $oOL_Recipient.Resolved Then Return SetError(4, $iOL_Index, 0)
			#forceref $oOL_TempRecipient ; to prevent the AU3Check warning: $oOL_TempRecipient: declared, but not used in func.
			$oOL_TempRecipient = $vOL_Item.Recipients.Add($oOL_Recipient)
			If @error Then Return SetError(3, $iOL_Index, 0)
			$oOL_TempRecipient.Type = $iOL_Type
		EndIf
	Next
	$vOL_Item.Close(0)
	Return $vOL_Item

EndFunc   ;==>_OL_ItemRecipientAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemRecipientDelete
; Description ...: Deletes one or multiple recipients from an item.
; Syntax.........: _OL_ItemRecipientDelete($oOL, $vOL_Item, $sOL_StoreID, $vOL_P1 = ""[, $vOL_P2 = ""[, $vOL_P3 = ""[, $vOL_P4 = ""[, $vOL_P5 = ""[, $vOL_P6 = ""[, $vOL_P7 = ""[, $vOL_P8 = ""[, $vOL_P9 = ""[, $vOL_P10 = ""]]]]]]]]])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the item
;                  $sOL_StoreID - StoreID where the EntryID is stored. Use the keyword "Default" to use the users mailbox
;                  $vOL_P1      - Number of the recipient to delete from the recipients collection
;                  +or a zero based one-dimensional array with unlimited number of recipients
;                  $vOL_P2      - Optional: Number of the recipient to delete from the recipients collection
;                  $vOL_P3      - Optional: Same as $vOL_P2
;                  $vOL_P4      - Optional: Same as $vOL_P2
;                  $vOL_P5      - Optional: Same as $vOL_P2
;                  $vOL_P6      - Optional: Same as $vOL_P2
;                  $vOL_P7      - Optional: Same as $vOL_P2
;                  $vOL_P8      - Optional: Same as $vOL_P2
;                  $vOL_P9      - Optional: Same as $vOL_P2
;                  $vOL_P10     - Optional: Same as $vOL_P2
; Return values .: Success - Item object
;                  Failure - Returns 0 and sets @error:
;                  |1 - No item specified
;                  |2 - Item could not be found. EntryID might be wrong
;                  |3 - Error removing recipient from the item. @extended = number of the invalid recipient parameter (zero based)
; Author ........: water
; Modified.......:
; Remarks .......: $vOL_P2 to $vOL_P10 will be ignored if $vOL_P1 is an array of numbers
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemRecipientDelete($oOL, $vOL_Item, $sOL_StoreID, $vOL_P1 = "", $vOL_P2 = "", $vOL_P3 = "", $vOL_P4 = "", $vOL_P5 = "", $vOL_P6 = "", $vOL_P7 = "", $vOL_P8 = "", $vOL_P9 = "", $vOL_P10 = "")

	Local $aOL_Recipients[10]
	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	; Move recipients into an array
	If Not IsArray($vOL_P1) Then
		$aOL_Recipients[0] = $vOL_P1
		$aOL_Recipients[1] = $vOL_P2
		$aOL_Recipients[2] = $vOL_P3
		$aOL_Recipients[3] = $vOL_P4
		$aOL_Recipients[4] = $vOL_P5
		$aOL_Recipients[5] = $vOL_P6
		$aOL_Recipients[6] = $vOL_P7
		$aOL_Recipients[7] = $vOL_P8
		$aOL_Recipients[8] = $vOL_P9
		$aOL_Recipients[9] = $vOL_P10
	Else
		$aOL_Recipients = $vOL_P1
	EndIf
	; Delete recipients from the item
	For $iOL_Index = 0 To UBound($aOL_Recipients) - 1
		If StringStripWS($aOL_Recipients[$iOL_Index], 3) = "" Then ContinueLoop
		$vOL_Item.Recipients.Remove($aOL_Recipients[$iOL_Index])
		If @error Then Return SetError(3, $iOL_Index, 0)
	Next
	$vOL_Item.Close(0)
	Return $vOL_Item

EndFunc   ;==>_OL_ItemRecipientDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemRecipientGet
; Description ...: Gets all recipients of an item.
; Syntax.........: _OL_ItemRecipientGet($oOL, $vOL_Item[, $sOL_StoreID = Default])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the item
;                  $sOL_StoreID - Optional: StoreID where the EntryID is stored (default = keyword "Default" = the users mailbox)
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - Recipient object
;                  |1 - Name of the recipient
;                  |2 - EntryID of the recipient
;                  Failure - Returns "" and sets @error:
;                  |1 - No item specified
;                  |2 - Item could not be found. EntryID might be wrong
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemRecipientGet($oOL, $vOL_Item, $sOL_StoreID = Default)

	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, "")
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	Local $aOL_Members[$vOL_Item.Recipients.Count + 1][3] = [[$vOL_Item.Recipients.Count, 3]]
	For $iOL_Index = 1 To $vOL_Item.Recipients.Count
		$aOL_Members[$iOL_Index][0] = $vOL_Item.Recipients.Item($iOL_Index)
		$aOL_Members[$iOL_Index][1] = $vOL_Item.Recipients.Item($iOL_Index).Name
		$aOL_Members[$iOL_Index][2] = $vOL_Item.Recipients.Item($iOL_Index).EntryID
	Next
	Return $aOL_Members

EndFunc   ;==>_OL_ItemRecipientGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemRecurrenceDelete
; Description ...: Delete recurrence information of an item (appointment or task).
; Syntax.........: _OL_ItemRecurrenceDelete($oOL, $vOL_Item[, $sOL_StoreID = Default])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the appointment or task item
;                  $sOL_StoreID - Optional: StoreID where the EntryID is stored (default = keyword "Default" = the users mailbox)
; Return values .: Success - Item object
;                  Failure - Returns 0 and sets @error:
;                  |1 - No appointment or task item specified
;                  |2 - Item could not be found. EntryID might be wrong
;                  |3 - Item has no recurrence information
;                  |4 - Error with ClearRecurrencePattern. For more info please see @extended
;                  |5 - Error with Save. For more info please see @extended
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemRecurrenceDelete($oOL, $vOL_Item, $sOL_StoreID = Default)

	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	; Recurrence object of the appointment
	If $vOL_Item.IsRecurring = False Then Return SetError(3, 0, 0)
	$vOL_Item.ClearRecurrencePattern
	If @error Then Return SetError(4, @error, 0)
	$vOL_Item.Save
	If @error Then Return SetError(5, @error, 0)
	Return $vOL_Item

EndFunc   ;==>_OL_ItemRecurrenceDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemRecurrenceExceptionGet
; Description ...: Return all exceptions in the recurrence information of an item (appointment or task).
; Syntax.........: _OL_ItemRecurrenceExceptionGet($oOL, $vOL_Item[, $sOL_StoreID = Default])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the appointment or task item
;                  $sOL_StoreID - Optional: StoreID where the EntryID is stored (default = keyword "Default" = the users mailbox)
; Return values .: Success - one-dimensional one based array with the following information:
;                  |1 - AppointmentItem: The AppointmentItem object that is the exception. Not valid for deleted appointments
;                  |2 - Deleted:         Returns True if the AppointmentItem was deleted from the recurring pattern
;                  |3 - OriginalDate:    A Date indicating the original date and time of an AppointmentItem before it was altered.
;                  +Will return the original date even if the AppointmentItem has been deleted.
;                  +However, it will not return the original time if deletion has occurred
;                  Failure - Returns "" and sets @error:
;                  |1 - No appointment or task item specified
;                  |2 - Item could not be found. EntryID might be wrong
;                  |3 - Item has no recurrence information
;                  |4 - Error with GetRecurrencePattern. For more info please see @extended
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemRecurrenceExceptionGet($oOL, $vOL_Item, $sOL_StoreID = Default)

	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, "")
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, "")
	EndIf
	; Recurrence object of the appointment
	If $vOL_Item.IsRecurring = False Then Return SetError(3, 0, "")
	Local $oOL_Recurrence = $vOL_Item.GetRecurrencePattern
	If @error Then Return SetError(4, @error, "")
	Local $aOL_Exceptions[$oOL_Recurrence.Exceptions.Count + 1][3] = [[$oOL_Recurrence.Exceptions.Count, 3]]
	Local $iOL_Index = 1
	For $oOL_Exception In $oOL_Recurrence.Exceptions
		$aOL_Exceptions[$iOL_Index][0] = $oOL_Exception.AppointmentItem
		$aOL_Exceptions[$iOL_Index][1] = $oOL_Exception.Deleted
		$aOL_Exceptions[$iOL_Index][2] = $oOL_Exception.OriginalDate
		$iOL_Index += 1
	Next
	Return $aOL_Exceptions

EndFunc   ;==>_OL_ItemRecurrenceExceptionGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemRecurrenceExceptionSet
; Description ...: Define an exception in the recurrence information of an item (appointment or task).
; Syntax.........: _OL_ItemRecurrenceExceptionSet($oOL, $vOL_Item, $sOL_StoreID, $sOL_StartDate[, $sOL_NewStartDate = ""[, $sOL_NewEndDate = ""[, $sOL_NewSubject = ""[, $sOL_NewBody = ""]]]]
; Parameters ....: $oOL              - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item         - EntryID or object of the appointment or task item
;                  $sOL_StoreID      - StoreID where the EntryID is stored. Use "Default" if you use the users mailbox
;                  $sOL_StartDate    - Start date and time of the item to be changed
;                  $sOL_NewStartDate - Optional: New start date and time
;                  $sOL_NewEndDate   - Optional: New end date and time or duration in minutes
;                  $sOL_NewSubject   - Optional: New subject
;                  $sOL_NewBody      - Optional: New body
; Return values .: Success - item object of the exception item
;                  Failure - Returns 0 and sets @error:
;                  |1 - No appointment or task item specified
;                  |2 - Item could not be found. EntryID might be wrong
;                  |3 - Item has no recurrence information
;                  |4 - Error with GetRecurrencePattern. For more info please see @extended
;                  |5 - Error accessing the specified occurrence. Date/time might be invalid. For more info please see @extended
;                  |6 - Error saving the exception. For more info please see @extended
; Author ........: water
; Modified.......:
; Remarks .......: To change more properties please use _OL_ItemModify
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemRecurrenceExceptionSet($oOL, $vOL_Item, $sOL_StoreID, $sOL_StartDate, $sOL_NewStartDate = "", $sOL_NewEndDate = "", $sOL_NewSubject = "", $sOL_NewBody = "")

	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	; Recurrence object of the appointment
	If $vOL_Item.IsRecurring = False Then Return SetError(3, 0, 0)
	Local $oOL_Recurrence = $vOL_Item.GetRecurrencePattern
	If @error Then Return SetError(4, @error, 0)
	Local $oOL_OccurrenceItem = $oOL_Recurrence.GetOccurrence($sOL_StartDate)
	If @error Then Return SetError(5, @error, 0)
	If $sOL_NewStartDate <> "" Then $oOL_OccurrenceItem.Start = $sOL_NewStartDate
	If $sOL_NewEndDate <> "" Then
		If IsNumber($sOL_NewEndDate) Then
			$oOL_OccurrenceItem.Duration = $sOL_NewEndDate
		Else
			$oOL_OccurrenceItem.End = $sOL_NewEndDate
		EndIf
	EndIf
	If $sOL_NewSubject <> "" Then $oOL_OccurrenceItem.Subject = $sOL_NewSubject
	If $sOL_NewBody <> "" Then $oOL_OccurrenceItem.Body = $sOL_NewBody
	$oOL_OccurrenceItem.Save
	If @error Then Return SetError(6, @error, 0)
	Return $oOL_OccurrenceItem

EndFunc   ;==>_OL_ItemRecurrenceExceptionSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemRecurrenceGet
; Description ...: Gets recurrence information of an item (appointment or task).
; Syntax.........: _OL_ItemRecurrenceGet($oOL, $vOL_Item[, $sOL_StoreID = Default])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the appointment or task item
;                  $sOL_StoreID - Optional: StoreID where the EntryID is stored (default = keyword "Default" = the users mailbox)
; Return values .: Success - one-dimensional one based array with the following information:
;                  |1  - DayOfMonth:       Integer indicating the day of the month on which the recurring appointment or task occurs
;                  |2  - DayOfWeekMask:    OlDaysOfWeek constant representing the mask for the days of the week on which the recurring appointment or task occurs
;                  |3  - Duration:         Integer indicating the duration (in minutes) of the RecurrencePattern
;                  |4  - EndTime:          Time indicating the end time for a recurrence pattern
;                  |5  - Instance:         Integer specifying the count for which the recurrence pattern is valid for a given interval
;                  |6  - Interval:         Integer specifying the number of units of a given recurrence type between occurrences
;                  |7  - MonthOfYear:      Integer indicating which month of the year is valid for the specified recurrence pattern
;                  |8  - NoEndDate:        Boolean value that indicates True if the recurrence pattern has no end date
;                  |9  - Occurrences:      Integer indicating the number of occurrences of the recurrence pattern
;                  |10 - PatternEndDate:   Date indicating the end date for the recurrence pattern
;                  |11 - PatternStartDate: Date indicating the start date for the recurrence pattern
;                  |12 - RecurrenceType:   OlRecurrenceType constant specifying the frequency of occurrences for the recurrence pattern
;                  |13 - StartTime:        Time indicating the start time for a recurrence pattern
;                  Failure - Returns "" and sets @error:
;                  |1 - No appointment or task item specified
;                  |2 - Item could not be found. EntryID might be wrong
;                  |3 - Item has no recurrence information
;                  |4 - Error with GetRecurrencePattern. For more info please see @extended
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemRecurrenceGet($oOL, $vOL_Item, $sOL_StoreID = Default)

	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, "")
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, "")
	EndIf
	; Recurrence object of the appointment
	If $vOL_Item.IsRecurring = False Then Return SetError(3, 0, "")
	Local $oOL_Recurrence = $vOL_Item.GetRecurrencePattern
	If Not IsObj($oOL_Recurrence) Or @error Then Return SetError(4, @error, "")
	Local $aOL_Pattern[14] = [14]
	$aOL_Pattern[1] = $oOL_Recurrence.DayOfMonth
	$aOL_Pattern[2] = $oOL_Recurrence.DayOfWeekMask
	$aOL_Pattern[3] = $oOL_Recurrence.Duration
	$aOL_Pattern[4] = $oOL_Recurrence.EndTime
	$aOL_Pattern[5] = $oOL_Recurrence.Instance
	$aOL_Pattern[6] = $oOL_Recurrence.Interval
	$aOL_Pattern[7] = $oOL_Recurrence.MonthOfYear
	$aOL_Pattern[8] = $oOL_Recurrence.NoEndDate
	$aOL_Pattern[9] = $oOL_Recurrence.Occurrences
	$aOL_Pattern[10] = $oOL_Recurrence.PatternEndDate
	$aOL_Pattern[11] = $oOL_Recurrence.PatternStartDate
	$aOL_Pattern[12] = $oOL_Recurrence.RecurrenceType
	$aOL_Pattern[13] = $oOL_Recurrence.StartTime
	Return $aOL_Pattern

EndFunc   ;==>_OL_ItemRecurrenceGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemRecurrenceSet
; Description ...: Sets recurrence information of an item (appointment or task).
; Syntax.........: _OL_ItemRecurrenceSet($oOL, $vOL_Item, $sOL_StoreID, $sOL_PatternStartDate, $sOL_StartTime, $vOL_PatternEndDate, $vOL_EndTime, $iOL_RecurrenceType, $iOL_DayOf, $iOL_Interval, $iOL_Instance, $iOL_Occurrences)
; Parameters ....: $oOL                  - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item             - EntryID or object of the appointment or task item
;                  $sOL_StoreID          - StoreID where the EntryID is stored. Use "Default" if you use the users mailbox
;                  $sOL_PatternStartDate - Date indicating the start date for the recurrence pattern
;                  $sOL_StartTime        - Time indicating the start time for the recurrence pattern
;                  $vOL_PatternEndDate   - Date indicating the end date for the recurrence pattern OR
;                  +                       "" that indicates the recurrence pattern has no end date OR
;                  +                       an integer indicating the number of occurrences of the recurrence pattern
;                  $vOL_EndTime          - Time indicating the end time for the recurrence pattern OR
;                  +                       an integer indicating the duration (in minutes) of the recurrence pattern
;                  $iOL_RecurrenceType   - Constant specifying the frequency of occurrences for the recurrence pattern.
;                  +                       Is defined by the Outlook OlRecurrenceType enumeration
;                  $iOL_DayOf            - DayOfWeekMask (mask for the days of the week on which the recurring appointment or task occurs) OR
;                  +                       DayOfMonth (integer indicating the day of the month on which the recurring appointment or task occurs) if $sOL_RecurrenceType = $olRecursMonthly OR
;                  +                       DayOfMonth/MonthOfYear (integer indicating the day of the month and month of the year on which the recurring appointment or task occurs) if $sOL_RecurrenceType = $olRecursYearly
;                  $iOL_Interval         - Integer specifying the number of units of a given recurrence type between occurrences
;                  $iOL_Instance         - Integer specifying the count for which the recurrence pattern is valid for a given interval.
;                                          Only valid for $sOL_RecurrenceType $olRecursMonthNth and $olRecursYearNth
; Return values .: Success - Item object
;                  Failure - Returns 0 and sets @error:
;                  |1 - No appointment or task item specified
;                  |2 - Item could not be found. EntryID might be wrong
;                  |3 - Error with Save. For more info please see @extended
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemRecurrenceSet($oOL, $vOL_Item, $sOL_StoreID, $sOL_PatternStartDate, $sOL_StartTime, $vOL_PatternEndDate, $vOL_EndTime, $iOL_RecurrenceType, $iOL_DayOf, $iOL_Interval, $iOL_Instance)

	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	; Recurrence object of the item
	Local $oOL_Recurrence = $vOL_Item.GetRecurrencePattern
	#forceref $oOL_Recurrence ; to prevent the AU3Check warning: $oOL_Recurrence: declared, but not used in func.
	; Set properties of the reccurrence
	$oOL_Recurrence.RecurrenceType = $iOL_RecurrenceType
	$oOL_Recurrence.PatternStartDate = $sOL_PatternStartDate
	$oOL_Recurrence.StartTime = $sOL_StartTime
	; Set PatternEndDate to date, number of occurrences or NoEndDate
	If IsInt($vOL_PatternEndDate) Then
		$oOL_Recurrence.Occurrences = $vOL_PatternEndDate
	ElseIf $vOL_PatternEndDate <> "" Then
		$oOL_Recurrence.PatternEndDate = $vOL_PatternEndDate
	Else
		$oOL_Recurrence.NoEndDate = True
	EndIf
	; Set PatternEndTime to time or duration
	If IsInt($vOL_EndTime) Then
		$oOL_Recurrence.Duration = $vOL_EndTime
	Else
		$oOL_Recurrence.EndTime = $vOL_EndTime
	EndIf
	; Set DayOfWeekMask or DayOfMonth and MonthOfYear
	If $iOL_RecurrenceType = $olRecursYearly Then
		Local $aOL_Temp = StringSplit($iOL_DayOf, "/")
		$oOL_Recurrence.DayOfMonth = $aOL_Temp[1]
		$oOL_Recurrence.MonthofYear = $aOL_Temp[2]
	EndIf
	If $iOL_RecurrenceType = $olRecursWeekly Or $iOL_RecurrenceType = $olRecursMonthNth Or $iOL_RecurrenceType = $olRecursYearNth And $iOL_DayOf <> "" Then $oOL_Recurrence.DayOfWeekMask = $iOL_DayOf
	If $iOL_RecurrenceType = $olRecursMonthly And $iOL_DayOf <> "" Then $oOL_Recurrence.DayOfMonth = $iOL_DayOf
	; Set Interval
	If $iOL_Interval <> 0 Then $oOL_Recurrence.Interval = $iOL_Interval
	; Set Instance
	If $iOL_RecurrenceType = $olRecursMonthNth Or $iOL_RecurrenceType = $olRecursYearNth And $iOL_Instance <> 0 Then $oOL_Recurrence.Instance = $iOL_Instance
	$vOL_Item.Save
	If @error Then Return SetError(3, @error, 0)
	Return $vOL_Item

EndFunc   ;==>_OL_ItemRecurrenceSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemReply
; Description ...: Replies/responds to an item.
; Syntax.........: _OL_ItemReply($oOL, $vOL_Item[, $sOL_StoreID[, $bOL_ReplyAll = False[, $iOL_Response = $olMeetingAccepted]]])
; Parameters ....: $oOL          - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item     - EntryID or object of the item to move
;                  $sOL_StoreID  - Optional: StoreID of the source store as returned by _OL_FolderAccess (default = keyword "Default" = the users mailbox)
;                  $bOL_ReplyAll - Optional: False: reply to the original sender (default), True: reply to all recipients
;                  $iOL_Response - Optional: Indicates the response to a meeting request. Is defined by the Outlook OlMeetingResponse enumeration
;                  +(default = $olMeetingAccepted = The meeting was accepted)
; Return values .: Success - object of the created item
;                  Failure - Returns 0 and sets @error:
;                  |1 - No item has been specified
;                  |2 - Item could not be found. EntryID might be wrong
;                  |3 - Error with method .Reply, .ReplyAll or .Respond. For more info please see @extended
;                  |4 - Error with Save. For more info please see @extended
; Author ........: water
; Modified ......:
; Remarks .......: $bOL_ReplyAll is used for mail items and ignored for all other items
;                  $iOL_Response is used for meeting and task items and ignored for all other items
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemReply($oOL, $vOL_Item, $sOL_StoreID = Default, $bOL_ReplyAll = False, $iOL_Response = Default)

	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	; Mail: reply or replyall
	If $vOL_Item.Class = $olMail Then
		If $bOL_ReplyAll Then
			$vOL_Item = $vOL_Item.Reply
			If @error Then Return SetError(3, @error, 0)
		Else
			$vOL_Item = $vOL_Item.Reply
			If @error Then Return SetError(3, @error, 0)
		EndIf
	EndIf
	; Meeting request: Respond
	If $vOL_Item.Class = $olAppointment Then
		If $iOL_Response = Default Then $iOL_Response = $olMeetingAccepted
		$vOL_Item = $vOL_Item.Respond($iOL_Response)
		If @error Then Return SetError(3, @error, 0)
	EndIf
	; Task: Respond
	If $vOL_Item.Class = $olTask Then
		If $iOL_Response = Default Then $iOL_Response = $olTaskAccept
		$vOL_Item = $vOL_Item.Respond($iOL_Response)
		If @error Then Return SetError(3, @error, 0)
	EndIf
	$vOL_Item.Close(0)
	If @error Then Return SetError(4, @error, 0)
	Return $vOL_Item

EndFunc   ;==>_OL_ItemReply

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemSave
; Description ...: Saves an item (contact, appointment ...) and/or all attachments in the specified path with the specified type.
; Syntax.........: _OL_ItemSave($oOL, $vOL_Item, $sOL_StoreID, $sOL_Path, $iOL_Type[, $iOL_Flags = 0])
; Parameters ....: $oOL              - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item         - EntryID or object of the item to save
;                  $sOL_StoreID      - StoreID of the source store as returned by _OL_FolderAccess. Use the keyword "Default" to use the users mailbox
;                  $sOL_Path         - Path (drive, directory[, filename]) where to save the item.
;                                      If the filename is missing it is set to the item subject. In this case the directory needs a trailing backslash.
;                                      The extension is always set according to $iOL_Type.
;                                      If the directory does not exist it is created
;                  $iOL_Type         - The file type to save. Is defined by the Outlook OlSaveAsType enumeration
;                  $iOL_Flags        - Optional: Flags to set different processing options. Can be a combination of the following:
;                  |  1: Save the item (default)
;                  |  2: Save attachments. Will be saved into the same directory as the item itself.
;                  +Name is Filename of the item, underscore plus name of attachment plus (optional) unterscore plus integer so multiple att. with the same name
;                  +can be saved
; Return values .: Success - Object of the saved item
;                  Failure - Returns 0 and sets @error:
;                  |1 - $sOL_Path is missing
;                  |2 - Specified directory does not exist. It could not be created
;                  |3 - $iOL_Type is missing or invalid
;                  |4 - Error saving the item. For details check @extended
;                  |5 - Error saving an attachment. For details check @extended
;                  |6 - No or an invalid item has been specified
;                  |7 - Invalid $iOL_Type specified
;                  |8 - Could not save attachment. More than 99 files with the same filename encountered. @extended is set to the attachment number in error
; Author ........: water
; Modified ......:
; Remarks .......: $iOL_Flags: 1 = save the item without attachments, 2 = save attachments only, 3 = save item + attachments
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemSave($oOL, $vOL_Item, $sOL_StoreID, $sOL_Path, $iOL_Type, $iOL_Flags = 1)

	Local $aOL_Type2Ext[11][2] = [[$olDoc, ".doc"],[$olHTML, ".html"],[$olICal, ".ics"],[$olMHTML, ".mht"],[$olMSG, ".msg"],[$olMSGUnicode, ".msg"], _
			[$olRTF, ".rtf"],[$olTemplate, ".oft"],[$olTXT, ".txt"],[$olVCal, ".vcs"],[$olVCard, "vcf"]]
	Local $sOL_Drive, $sOL_Dir, $sOL_FName, $sOL_Ext
	If StringStripWS($sOL_Path, 3) = "" Then Return SetError(1, 0, 0)
	_PathSplit($sOL_Path, $sOL_Drive, $sOL_Dir, $sOL_FName, $sOL_Ext)
	If Not FileExists($sOL_Drive & $sOL_Dir) Then
		If DirCreate($sOL_Drive & $sOL_Dir) = 0 Then Return SetError(2, 0, 0)
	EndIf
	If Not IsNumber($iOL_Type) Then Return SetError(3, 0, 0)
	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(6, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(6, @error, 0)
	EndIf
	; Set filename to item subject if filename is empty
	If $sOL_FName = "" Then $sOL_FName = $vOL_Item.Subject
	; Replace invalid characters from filename with underscore
	$sOL_FName = StringRegExpReplace($sOL_FName, '[ \/:*?"<>|]', '_')
	; Select extension according to $iOL_Type
	For $iOL_Index = 0 To UBound($aOL_Type2Ext) - 1
		If $aOL_Type2Ext[$iOL_Index][0] = $iOL_Type Then $sOL_Ext = $aOL_Type2Ext[$iOL_Index][1]
	Next
	If $sOL_Ext = "" Then Return SetError(7, 0, 0)
	$sOL_Path = $sOL_Drive & $sOL_Dir & $sOL_FName & $sOL_Ext
	; Save item
	If BitAND($iOL_Flags, 1) = 1 Then
		$vOL_Item.SaveAs($sOL_Path, $iOL_Type)
		If @error Then Return SetError(4, @error, 0)
	EndIf
	; Save attachments
	If BitAND($iOL_Flags, 2) = 2 Then
		Local $aOL_Attachments = _OL_ItemAttachmentGet($oOL, $vOL_Item, $sOL_StoreID)
		If @error = 0 Then
			For $iOL_Index = 1 To $aOL_Attachments[0][0]
				If FileExists($sOL_Drive & $sOL_Dir & $sOL_FName & "_" & $aOL_Attachments[$iOL_Index][2]) = 1 Then
					Local $aOL_Temp = StringSplit($aOL_Attachments[$iOL_Index][2], ".")
					For $iOL_Index2 = 1 To 99
						If FileExists($sOL_Drive & $sOL_Dir & $sOL_FName & "_" & $aOL_Temp[1] & "_" & $iOL_Index2 & "." & $aOL_Temp[2]) = 0 Then ExitLoop
					Next
					If $iOL_Index2 > 99 Then Return SetError(8, $iOL_Index, 0)
					$aOL_Attachments[$iOL_Index][0] .SaveAsFile($sOL_Drive & $sOL_Dir & $sOL_FName & "_" & $aOL_Temp[1] & "_" & $iOL_Index2 & "." & $aOL_Temp[2])
					If @error Then Return SetError(5, @error, 0)
				Else
					$aOL_Attachments[$iOL_Index][0] .SaveAsFile($sOL_Drive & $sOL_Dir & $sOL_FName & "_" & $aOL_Attachments[$iOL_Index][2])
					If @error Then Return SetError(5, @error, 0)
				EndIf
			Next
		EndIf
	EndIf
	Return $vOL_Item

EndFunc   ;==>_OL_ItemSave

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemSend
; Description ...: Sends an item (appointment, mail, task) using the specified EntryID and StoreID.
; Syntax.........: _OL_ItemSend($oOL, $vOL_Item[, $sOL_StoreID = Default])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the item to send
;                  $sOL_StoreID - Optional: StoreID of the source store as returned by _OL_FolderAccess (default = keyword "Default" = the users mailbox)
; Return values .: Success - Object of the item
;                  Failure - Returns 0 and sets @error:
;                  |1 - No or an invalid item has been specified
;                  |2 - Error sending the item. For details check @extended
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemSend($oOL, $vOL_Item, $sOL_StoreID = Default)

	If Not IsObj($vOL_Item) Then
		If StringStripWS($vOL_Item, 3) = "" Then Return SetError(1, 0, 0)
		$vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
		If @error Then Return SetError(1, @error, 0)
	EndIf
	$vOL_Item.Send()
	If @error Then Return SetError(2, @error, 0)
	Return $vOL_Item

EndFunc   ;==>_OL_ItemSend

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemSendReceive
; Description ...: Initiates immediate delivery of all undelivered messages and immediate receipt of mail for all accounts in the current profile.
; Syntax.........: _OL_ItemSendReceive($oOL[, $bOL_ShowProgress = False])
; Parameters ....: $oOL              - Outlook object returned by a preceding call to _OL_Open()
;                  $bOL_ShowProgress - Optional: If True show the Outlook Send/Receive progress dialog box (default = False)
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1  - Error executing the SendAndReceive method. For details check @extended
;                  |99 - Function not available for this Outlook version. @extended denotes the lowest required Outlook version to run the function
; Author ........: water
; Modified ......:
; Remarks .......: Only available for Outlook 2007 and later
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemSendReceive($oOL, $bOL_ShowProgress = False)

	Local $aVersion = StringSplit($oOL.Version, '.')
	If Int($aVersion[1]) < 12 Then Return SetError(99, 12, 0)
	Local $oOL_Namespace = $oOL.GetNamespace("MAPI")
	$oOL_Namespace.SendAndReceive($bOL_ShowProgress)
	If @error Then Return SetError(1, @error, 0)
	Return 1

EndFunc   ;==>_OL_ItemSendReceive

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ItemSync
; Description ...: Starts synchronization for all or a single Send/Receive group(s) set up for the user.
; Syntax.........: _OL_ItemSync($oOL[, $sOL_Group = ""])
; Parameters ....: $oOL       - Outlook object returned by a preceding call to _OL_Open()
;                  $sOL_Group - Optional: Name of the Send/Receive group to be synchronized (default = all)
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - Error returned by method Start. For details please check @extended
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ItemSync($oOL, $sOL_Group = "")

	Local $oOL_Namespace = $oOL.GetNamespace("MAPI")
	For $iOL_Index = 1 To $oOL_Namespace.SyncObjects.Count
		If $sOL_Group = "" Or $sOL_Group = $oOL_Namespace.SyncObjects.Item($iOL_Index).Name Then
			$oOL_Namespace.SyncObjects.Item($iOL_Index).Start
			If @error Then Return SetError(1, @error, 0)
		EndIf
	Next
	Return 1

EndFunc   ;==>_OL_ItemSync

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_MailheaderGet
; Description ...: Get the headers of a mail item using the specified EntryID and StoreID.
; Syntax.........: _OL_MailheaderGet($oOL, $vOL_Item[, $sOL_StoreID])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Item    - EntryID or object of the mail item
;                  $sOL_StoreID - Optional: StoreID of the source store as returned by _OL_FolderAccess (default = keyword "Default" = the users mailbox)
; Return values .: Success - Returns a string with the mail headers
;                  Failure - Returns "" and sets @error:
;                  |1 - Error getting the mail object from the specified EntryID and StoreID
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_MailheaderGet($oOL, $vOL_Item, $sOL_StoreID = Default)

	Local $sPR_MAIL_HEADER_TAG = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
	If Not IsObj($vOL_Item) Then $vOL_Item = $oOL.Session.GetItemFromID($vOL_Item, $sOL_StoreID)
	If @error Then Return SetError(1, @error, "")
	Local $oOL_PA = $vOL_Item.PropertyAccessor
	Return $oOL_PA.GetProperty($sPR_MAIL_HEADER_TAG)

EndFunc   ;==>_OL_MailheaderGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_MailSignatureCreate
; Description ...: Creates a new/modifies an existing e-mail signature.
; Syntax.........: _OL_MailSignatureCreate($sOL_Name, $oOL_Word, $oOL_Range[, $bOL_NewMessage = False[, $bOL_ReplyMessage = False]])
; Parameters ....: $sOL_Name          - Name of the signature to be created/modified.
;                  $oOL_Word          - Object of an already running Word Application
;                  $oOL_Range         - Range (as defined by the word range method) that contains the signature text + formatting
;                  $bOL_NewMessage    - Optional: True sets the signature as the default signature to be added to new email messages (default = False)
;                  $bOL_ReplyMessage  - Optional: True sets the signature as the default signature to be added when you reply to an email messages (default = False)
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oOL_Word is not an object
;                  |2 - $sOL_Name is empty
;                  |3 - $oOL_Range is not an object
;                  |4 - Error adding signature. For details check @extended
; Author ........: water
; Modified ......:
; Remarks .......: If the signature already exist $bOL_NewMessage and $bOL_ReplyMessage can be set but not unset. Use _OL_MailSignatureSet in this case.
; Related .......:
; Link ..........: http://technet.microsoft.com/en-us/magazine/2006.10.heyscriptingguy.aspx
; Example .......: Yes
; ===============================================================================================================================
Func _OL_MailSignatureCreate($sOL_Name, $oOL_Word, $oOL_Range, $bOL_NewMessage = False, $bOL_ReplyMessage = False)

	If Not IsObj($oOL_Word) Then Return SetError(1, 0, "")
	If StringStripWS($sOL_Name, 3) = "" Then Return SetError(2, 0, "")
	If Not IsObj($oOL_Range) Then Return SetError(3, 0, "")
	Local $oOL_EmailOptions = $oOL_Word.EmailOptions
	Local $oOL_SignatureObject = $oOL_EmailOptions.EmailSignature
	Local $oOL_SignatureEntries = $oOL_SignatureObject.EmailSignatureEntries
	$oOL_SignatureEntries.Add($sOL_Name, $oOL_Range)
	If @error Then Return SetError(4, @error, 0)
	If $bOL_NewMessage Then $oOL_SignatureObject.NewMessageSignature = $sOL_Name
	If $bOL_ReplyMessage Then $oOL_SignatureObject.ReplyMessageSignature = $sOL_Name
	Return 1

EndFunc   ;==>_OL_MailSignatureCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_MailSignatureDelete
; Description ...: Deletes an existing e-mail signature.
; Syntax.........: _OL_MailSignatureDelete($sOL_Signature[, $oOL_Word = 0])
; Parameters ....: $sOL_Signature - Name of the signature to be created/modified
;                  $oOL_Word      - Optional: Object of an already running Word Application (default = 0 = no Word Application running)
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oOL_Word is not an object
;                  |2 - $sOL_Signature is empty
;                  |3 - $sOL_Signature does not exist
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_MailSignatureDelete($sOL_Signature, $oOL_Word = 0)

	If StringStripWS($sOL_Signature, 3) = "" Then Return SetError(2, 0, "")
	Local $bOL_WordStart = False
	If $oOL_Word = 0 Then
		$oOL_Word = ObjCreate("Word.Application")
		$bOL_WordStart = True
	EndIf
	If @error Or Not IsObj($oOL_Word) Then Return SetError(1, @error, "")
	; Check if the specified signatures exist
	_OL_MaiLSignatureGet($sOL_Signature, $oOL_Word)
	If @error Then
		If $bOL_WordStart = True Then
			$oOL_Word.Quit
			$oOL_Word = 0
		EndIf
		Return SetError(3, 0, 0)
	EndIf
	Local $oOL_EmailOptions = $oOL_Word.EmailOptions
	Local $oOL_SignatureObject = $oOL_EmailOptions.EmailSignature
	Local $oOL_SignatureEntries = $oOL_SignatureObject.EmailSignatureEntries
	$oOL_SignatureEntries.Item($sOL_Signature).Delete
	Local $iOL_Error = @error, $iOL_Extended = @extended
	If $bOL_WordStart = True Then
		$oOL_Word.Quit
		$oOL_Word = 0
	EndIf
	If $iOL_Error <> 0 Then Return SetError($iOL_Error, $iOL_Extended, 0)
	Return 1

EndFunc   ;==>_OL_MailSignatureDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_MailSignatureGet
; Description ...: Gets a list of e-mail signatures used when you create/edit e-mail messages and replies.
; Syntax.........: _OL_MailSignatureGet([$sOL_Signature = ""[, $oOL_Word = 0]])
; Parameters ....: $sOL_Signature - Optional: Name of a signature to check for existance. The result contains this single signature or is set to error.
;                  $oOL_Word      - Optional: Object of an already running Word Application (default = 0 = no Word Application running)
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - Name of the signature
;                  |1 - True if the signature is used when creating new messages
;                  |2 - True if the signature is used when replying to a message
;                  Failure - Returns "" and sets @error:
;                  |1 - Error accessing word object. For details check @extended
;                  |2 - Specified signature does not exist
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_MailSignatureGet($sOL_Signature = "", $oOL_Word = 0)

	Local $bOL_WordStart = False
	If $oOL_Word = 0 Then
		$oOL_Word = ObjCreate("Word.Application")
		$bOL_WordStart = True
	EndIf
	If @error Or Not IsObj($oOL_Word) Then Return SetError(1, @error, "")
	Local $oOL_EmailOptions = $oOL_Word.EmailOptions
	Local $oOL_SignatureObject = $oOL_EmailOptions.EmailSignature
	Local $oOL_SignatureEntries = $oOL_SignatureObject.EmailSignatureEntries
	Local $sOL_NewMessageSig = $oOL_SignatureObject.NewMessageSignature
	Local $sOL_ReplyMessageSig = $oOL_SignatureObject.ReplyMessageSignature
	Local $aOL_Signatures[$oOL_SignatureEntries.Count + 1][3]
	Local $iOL_Index = 0
	For $oOL_SignatureEntry In $oOL_SignatureEntries
		If $sOL_Signature = "" Or $sOL_Signature == $oOL_SignatureEntry.Name Then
			$iOL_Index = $iOL_Index + 1
			$aOL_Signatures[$iOL_Index][0] = $oOL_SignatureEntry.Name
			If $aOL_Signatures[$iOL_Index][0] = $sOL_NewMessageSig Then
				$aOL_Signatures[$iOL_Index][1] = True
			Else
				$aOL_Signatures[$iOL_Index][1] = False
			EndIf
			If $aOL_Signatures[$iOL_Index][0] = $sOL_ReplyMessageSig Then
				$aOL_Signatures[$iOL_Index][2] = True
			Else
				$aOL_Signatures[$iOL_Index][2] = False
			EndIf
		EndIf
	Next
	ReDim $aOL_Signatures[$iOL_Index + 1][3]
	$aOL_Signatures[0][0] = $iOL_Index
	$aOL_Signatures[0][1] = UBound($aOL_Signatures, 2)
	If $bOL_WordStart = True Then
		$oOL_Word.Quit
		$oOL_Word = 0
	EndIf
	If $sOL_Signature <> "" And $aOL_Signatures[0][0] = 0 Then Return SetError(2, 0, "")
	Return $aOL_Signatures

EndFunc   ;==>_OL_MailSignatureGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_MailSignatureSet
; Description ...: Sets the signature to be added to new email messages and/or when you reply to an email message.
; Syntax.........: _OL_MailSignatureSet($sOL_NewMessage , $sOL_ReplyMessage[, $oOL_Word = 0])
; Parameters ....: $sOL_NewMessage   - Name of the signature to be added to new email messages. "" removes the default signature. Keyword Default leaves the signature unchanged
;                  $sOL_ReplyMessage - Name of the signature to be added when you reply to an email messages. "" removes the default signature. Keyword Default leaves the signature unchanged
;                  $oOL_Word         - Optional: Object of an already running Word Application (default = 0 = no Word Application running)
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oOL_Word is not an object
;                  |2 - Error getting list of signatures using _OL_MailSignatureGet. Please check @extended
;                  |3 - $sOL_NewMessage could not be found in the list of already defined signatures.
;                  |4 - $sOL_ReplyMessage could not be found in the list of already defined signatures.
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_MailSignatureSet($sOL_NewMessage, $sOL_ReplyMessage, $oOL_Word = 0)

	Local $bOL_WordStart = False
	If $oOL_Word = 0 Then
		$oOL_Word = ObjCreate("Word.Application")
		$bOL_WordStart = True
	EndIf
	If @error Or Not IsObj($oOL_Word) Then Return SetError(1, @error, "")
	; Check if the specified signatures exist
	If $sOL_NewMessage <> Default And $sOL_NewMessage <> "" Then
		_OL_MaiLSignatureGet($sOL_NewMessage, $oOL_Word)
		If @error Then
			If $bOL_WordStart = True Then
				$oOL_Word.Quit
				$oOL_Word = 0
			EndIf
			Return SetError(3, 0, 0)
		EndIf
	EndIf
	If $sOL_ReplyMessage <> Default And $sOL_ReplyMessage <> "" Then
		_OL_MaiLSignatureGet($sOL_ReplyMessage, $oOL_Word)
		If @error Then
			If $bOL_WordStart = True Then
				$oOL_Word.Quit
				$oOL_Word = 0
			EndIf
			Return SetError(4, 0, 0)
		EndIf
	EndIf
	; Set Signatures
	Local $oOL_EmailOptions = $oOL_Word.EmailOptions
	Local $oOL_SignatureObject = $oOL_EmailOptions.EmailSignature
	#forceref $oOL_SignatureObject
	If $sOL_NewMessage <> Default Then $oOL_SignatureObject.NewMessageSignature = $sOL_NewMessage
	If $sOL_ReplyMessage <> Default Then $oOL_SignatureObject.ReplyMessageSignature = $sOL_ReplyMessage
	If $bOL_WordStart = True Then
		$oOL_Word.Quit
		$oOL_Word = 0
	EndIf
	Return 1

EndFunc   ;==>_OL_MailSignatureSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_OOFGet
; Description ...: Get information about the OOF (Out of Office) setting of the specified store.
; Syntax.........: _OL_OOFGet($oOL[, $sOL_Store = "*"])
; Parameters ....: $oOL       - Outlook object returned by a preceding call to _OL_Open()
;                  $sOL_Store - Optional: Store for which the OOF should be retrieved.
;                               Use "*" to denote your default store or specify the store of another user
; Return values .: Success - one-dimensional one based array with the following information:
;                  |0 - State of the OOF. True = OOF is set, False = OOF is not set
;                  |1 - OOF text for internal senders
;                  Failure - Returns "" and sets @error:
;                  |1 - The specified store could not be accessed
;                  |2 - Error accessing the internal OOF mail item. For details please see @extended
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: http://social.msdn.microsoft.com/Forums/en-US/outlookdev/thread/3e3dd60b-a9ce-4484-b974-6b78766e376b
; Example .......: Yes
; ===============================================================================================================================
Func _OL_OOFGet($oOL, $sOL_Store = "*")

	Local $oOL_Item, $aOL_OOF[3] = [2]
	Local $aOL_Folder = _OL_FolderAccess($oOL, "\\" & $sOL_Store, $olFolderInbox)
	If @error Then Return SetError(1, @error, 0)
	If $sOL_Store = "*" Then $sOL_Store = $aOL_Folder[1] .Parent.Name
	; Get the status of the OOF for the specified store
	$aOL_OOF[1] = $oOL.Session.Stores.Item($sOL_Store).PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x661D000B")
	; Get the text of the internal OOF
	$oOL_Item = $aOL_Folder[1] .GetStorage("IPM.Note.Rules.OofTemplate.Microsoft", $olIdentifyByMessageClass)
	If @error Then Return SetError(2, @error, 0)
	$aOL_OOF[2] = $oOL_Item.Body
	Return $aOL_OOF

EndFunc   ;==>_OL_OOFGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_OOFSet
; Description ...: Sets the OOF (Out of Office) message for your or another users Exchange Store and/or activates/deactivates the OOF.
; Syntax.........: _OL_OOFSet($oOL, $sOL_Store, $bOL_OOFActivate, $sOL_OOFText)
; Parameters ....: $oOL             - Outlook object returned by a preceding call to _OL_Open()
;                  $sOL_Store       - Store for which the OOF should be set. Use "*" to denote your default store or specify the store of another user if you have write permission
;                  $bOL_OOFActivate - If set to True the OOF is activated. Keyword Default leaves the status unchanged
;                  $sOL_OOFText     - OOF reply text for internal messages. "" clears the text. Keyword Default leaves the text unchanged
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - Error returned by _OL_FolderAccess (the error code of this function can be found in @extended)
;                  |2 - Invalid StoreType. Has to be $olPrimaryExchangeMailbox or $olExchangeMailbox
;                  |3 - Error returned by Outlook GetStorage method for the internal OOF. For details please see @extended
;                  |4 - Error returned by Outlook Save method for the internal OOF. For details please see @extended
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: http://social.msdn.microsoft.com/Forums/en-US/outlookdev/thread/99b07ca3-e26c-4eab-b644-2c7749638f0e
; Example .......: Yes
; ===============================================================================================================================
Func _OL_OOFSet($oOL, $sOL_Store, $bOL_OOFActivate, $sOL_OOFText)

	Local $oOL_Item
	Local $aOL_Folder = _OL_FolderAccess($oOL, "\\" & $sOL_Store, $olFolderInbox)
	If @error Then Return SetError(1, @error, 0)
	If $sOL_Store = "*" Then $sOL_Store = $aOL_Folder[1] .Parent.Name
	Local $iOL_StoreType = $oOL.Session.Stores.Item($sOL_Store).ExchangeStoreType
	If $iOL_StoreType <> $olPrimaryExchangeMailbox And $iOL_StoreType <> $olExchangeMailbox Then Return SetError(2, 0, 0)
	; Set the text of the internal OOF
	If $sOL_OOFText <> Default Then
		$oOL_Item = $aOL_Folder[1] .GetStorage("IPM.Note.Rules.OofTemplate.Microsoft", $olIdentifyByMessageClass)
		If @error Then Return SetError(3, @error, 0)
		$oOL_Item.Body = $sOL_OOFText
		$oOL_Item.Save
		If @error Then Return SetError(4, @error, 0)
	EndIf
	; Set the status of the OOF for the specified store
	If $bOL_OOFActivate <> Default Then _
			$oOL.Session.Stores.Item($sOL_Store).PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x661D000B", $bOL_OOFActivate)
	Return 1

EndFunc   ;==>_OL_OOFSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_PSTAccess
; Description ...: Accesses a PST file so Outlook can access it as a folder.
; Syntax.........: _OL_PSTAccess($oOL, $sOL_PSTPath[, $sOL_DisplayName = ""])
; Parameters ....: $oOL             - Outlook object returned by a preceding call to _OL_Open()
;                  $sOL_PSTPath     - Path of the PST file (including filename & extension)
;                  $sOL_DisplayName - Optional: Displayname of the resulting Outlook folder (default = let Outlook set the display name)
; Return values .: Success - Object to the folder
;                  Failure - Returns 0 and sets @error:
;                  |1 - PST file $sOL_PSTPath does not exist
;                  |2 - Error accessing namespace object. For details check @extended
;                  |3 - Error adding the PST file as an Outlook folder. For details check @extended
; Author ........: water
; Modified ......:
; Remarks .......: You can pass element 1 of the resulting array to _OL_Folderget to get further information.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_PSTAccess($oOL, $sOL_PSTPath, $sOL_DisplayName = "")

	If FileExists($sOL_PSTPath) = 0 Then Return SetError(1, 0, 0)
	Local $oOL_Namespace = $oOL.GetNamespace("MAPI")
	If @error Or Not IsObj($oOL_Namespace) Then Return SetError(2, 0, 0)
	$oOL_Namespace.AddStore($sOL_PSTPath)
	If @error Then Return SetError(3, @error, 0)
	If $sOL_DisplayName <> "" Then $oOL_Namespace.Folders.GetLast.Name = $sOL_DisplayName ; Set Displayname
	Return $oOL_Namespace.Folders.GetLast

EndFunc   ;==>_OL_PSTAccess

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_PSTClose
; Description ...: Closes a PST file and removes the Outlook folder.
; Syntax.........: _OL_PSTClose($oOL, $oOL_Folder)
; Parameters ....: $oOL        - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Folder - Object of the Outlook folder representing the PST file or the displayname of the folder
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - Error accessing Namespace object. For details check @extended
;                  |2 - Error accessing the specified folder. For details check @extended
;                  |3 - Error removing the specified folder. For details check @extended
; Author ........: water
; Modified ......:
; Remarks .......: You can pass element 1 of the resulting array to _OL_Folderget to get further information.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_PSTClose($oOL, $oOL_Folder)

	Local $oOL_Namespace = $oOL.GetNamespace("MAPI")
	If @error Or Not IsObj($oOL_Namespace) Then Return SetError(1, 0, 0)
	If Not IsObj($oOL_Folder) Then
		$oOL_Folder = $oOL_Namespace.Folders.Item($oOL_Folder)
		If @error Or Not IsObj($oOL_Folder) Then Return SetError(2, @error, 0)
	EndIf
	$oOL_Namespace.RemoveStore($oOL_Folder)
	If @error Then Return SetError(3, @error, 0)
	Return 1

EndFunc   ;==>_OL_PSTClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_PSTCreate
; Description ...: Create a new (empty) PST file and access it in Outlook as a folder.
; Syntax.........: _OL_PSTCreate($oOL, $sOL_PSTPath[, $sOL_DisplayName = ""[, $iOL_PSTType = $olStoreANSI]])
; Parameters ....: $oOL             - Outlook object returned by a preceding call to _OL_Open()
;                  $sOL_PSTPath     - Path of the PST file (including filename & extension)
;                  $sOL_DisplayName - Optional: Displayname of the resulting Outlook folder (default = let Outlook set the display name)
;                  $iOL_PSTType     - Optional: Type of the PST file. Possible values:
;                  |$olStoreANSI    - ANSI format compatible with all previous versions of Microsoft Office Outlook format (default)
;                  |$olStoreDefault - Default format compatible with the mailbox mode in which Microsoft Office Outlook runs on the Microsoft Exchange Server
;                  |$olStoreUnicode - Unicode format compatible with Microsoft Office Outlook 2003 and later
; Return values .: Success - Object to the folder
;                  Failure - Returns 0 and sets @error:
;                  |1 - PST file $sOL_PSTPath already exists
;                  |2 - Error accessing Namespace object. For details check @extended
;                  |3 - Error creating the PST file. For details check @extended
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_PSTCreate($oOL, $sOL_PSTPath, $sOL_DisplayName = "", $iOL_PSTType = $olStoreANSI)

	If FileExists($sOL_PSTPath) = 1 Then Return SetError(1, 0, 0)
	Local $oOL_Namespace = $oOL.GetNamespace("MAPI")
	If @error Or Not IsObj($oOL_Namespace) Then Return SetError(2, 0, 0)
	$oOL_Namespace.AddStoreEx($sOL_PSTPath, $iOL_PSTType)
	If @error Then Return SetError(3, @error, 0)
	If $sOL_DisplayName <> "" Then $oOL_Namespace.Folders.GetLast.Name = $sOL_DisplayName ; Set Displayname
	Return $oOL_Namespace.Folders.GetLast

EndFunc   ;==>_OL_PSTCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_PSTGet
; Description ...: Gets a list of currently accessed PST files.
; Syntax.........: _OL_PSTGet($oOL)
; Parameters ....: $oOL - Outlook object returned by a preceding call to _OL_Open()
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - Displayname of the folder
;                  |1 - Object of the folder
;                  |2 - Path to the PST file in the filesystem
;                  Failure - Returns "" and sets @error:
;                  |1 - Error accessing namespace object. For details check @extended
; Author ........: water
; Modified ......:
; Remarks .......: You can pass element 1 of the resulting array to _OL_Folderget to get further information.
; Related .......:
; Link ..........: http://www.visualbasicscript.com/Find-PST-files-configured-in-outlook-m44947.aspx
; Example .......: Yes
; ===============================================================================================================================
Func _OL_PSTGet($oOL)

	Local $sOL_FolderSubString, $sOL_Path, $iIndex1 = 0, $iIndex2, $iOL_Pos, $aOL_PST[1][3] = [[0, 3]]
	Local $oOL_Namespace = $oOL.GetNamespace("MAPI")
	If @error Or Not IsObj($oOL_Namespace) Then Return SetError(1, 0, "")
	For $oOL_Folder In $oOL_Namespace.Folders
		$sOL_Path = ""
		For $iIndex2 = 1 To StringLen($oOL_Folder.StoreID) Step 2
			$sOL_FolderSubString = StringMid($oOL_Folder.StoreID, $iIndex2, 2)
			If $sOL_FolderSubString <> "00" Then $sOL_Path &= Chr(Dec($sOL_FolderSubString))
		Next
		If StringInStr($sOL_Path, "mspst.dll") > 0 Then ; PST file
			$iOL_Pos = StringInStr($sOL_Path, ":\")
			If $iOL_Pos > 0 Then
				$sOL_Path = StringMid($sOL_Path, $iOL_Pos - 1)
			Else
				$iOL_Pos = StringInStr($sOL_Path, "\\")
				If $iOL_Pos > 0 Then $sOL_Path = StringMid($sOL_Path, $iOL_Pos)
			EndIf
			ReDim $aOL_PST[UBound($aOL_PST, 1) + 1][UBound($aOL_PST, 2)]
			$iIndex1 = $iIndex1 + 1
			$aOL_PST[$iIndex1][0] = $oOL_Folder.Name
			$aOL_PST[$iIndex1][1] = $oOL_Namespace.GetFolderFromID($oOL_Folder.EntryID, $oOL_Folder.StoreID)
			$aOL_PST[$iIndex1][2] = $sOL_Path
			$aOL_PST[0][0] = $iIndex1
		EndIf
	Next
	Return $aOL_PST

EndFunc   ;==>_OL_PSTGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_RecipientFreeBusyGet
; Description ...: Returns free/busy information for the recipient.
; Syntax.........: _OL_RecipientFreeBusyGet($oOL, $vOL_Recipient, $sOL_Start[, $iOL_MinPerChar = 30[, $bOL_CompleteFormat = False]])
; Parameters ....: $oOL                - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_Recipient      - Name of a recipient or resolved object of a recipient
;                  $sOL_Start          - The start date for the returned period of free/busy information
;                  $iOL_MinPerChar     - Optional: The number of minutes per character represented in the returned free/busy string (default = 30)
;                  $bOL_CompleteFormat - Optional: True if the returned string should contain not only free/busy information, but also values for
;                  +each character according to the OlBusyStatus constants (default = False)
; Return values .: Success - String of free/busy information
;                  Failure - Returns "" and sets @error:
;                  |1 - No recipient has been specified
;                  |2 - Error creating recipient object. For details check @extended
;                  |3 - Recipient could not be resolved. For details check @extended
;                  |4 - Error retrieving the free/busy inforamtion. For details check @extended
; Author ........: water
; Modified ......:
; Remarks .......: The default is to return a string representing one month of free/busy information.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_RecipientFreeBusyGet($oOL, $vOL_Recipient, $sOL_Start, $iOL_MinPerChar = 30, $bOL_CompleteFormat = False)

	; Recipient specified as name - resolve
	If Not IsObj($vOL_Recipient) Then
		If StringStripWS($vOL_Recipient, 3) = "" Then Return SetError(1, 0, "")
		$vOL_Recipient = $oOL.Session.CreateRecipient($vOL_Recipient)
		If @error Or Not IsObj($vOL_Recipient) Then Return SetError(2, @error, "")
		$vOL_Recipient.Resolve
		If @error Or Not $vOL_Recipient.Resolved Then Return SetError(3, @error, "")
	EndIf
	Local $sOL_FreeBusy = $vOL_Recipient.FreeBusy($sOL_Start, $iOL_MinPerChar, $bOL_CompleteFormat)
	If @error Or $sOL_FreeBusy = "" Then Return SetError(4, @error, "")
	Return $sOL_FreeBusy

EndFunc   ;==>_OL_RecipientFreeBusyGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ReminderDelay
; Description ...: Delays the reminder by a specified time.
; Syntax.........: _OL_ReminderDelay($oOL_Reminder[, $iOL_DelayTime = 5])
; Parameters ....: $oOL_Reminder  - Represents a reminder object
;                  $iOL_DelayTime - Optional: amount of time (in minutes) to delay the reminder (default = 5)
; Return values .: Success - 1
;                  Failure - 0 and sets @error:
;                  |1 - You didn't specify our you specified an invalid object
;                  |2 - $iOL_DelayTime is not an integer
;                  |3 - Error returned by method .Snooze. For more information check @extended
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ReminderDelay($oOL_Reminder, $iOL_DelayTime = 5)

	If Not IsObj($oOL_Reminder) Then Return SetError(1, 0, 0)
	If Not IsInt($iOL_DelayTime) Then Return SetError(2, 0, 0)
	$oOL_Reminder.Snooze($iOL_DelayTime)
	If @error Then Return SetError(3, @error, 0)
	Return 1

EndFunc   ;==>_OL_ReminderDelay

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ReminderDismiss
; Description ...: Dismisses the specified reminder.
; Syntax.........: _OL_ReminderDismiss($oOL, $iOL_Reminder)
; Parameters ....: $oOL          - Outlook object returned by a preceding call to _OL_Open()
;                  $iOL_Reminder - Index number of the object in the reminders collection
; Return values .: Success - 1
;                  Failure - 0 and sets @error:
;                  |1 - The reminder has to be visible to be dismissed
;                  |2 - Error returned by method .Dismiss. For more information check @extended
; Author ........: water
; Modified ......:
; Remarks .......: The Dismiss method will fail if there is no visible reminder
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ReminderDismiss($oOL, $iOL_Reminder)

	If $oOL.Reminders.Item($iOL_Reminder).IsVisible = False Then Return SetError(1, 0, 0)
	$oOL.Reminders.Item($iOL_Reminder).Dismiss()
	If @error Then Return SetError(2, @error, 0)
	Return 1

EndFunc   ;==>_OL_ReminderDismiss

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_ReminderGet
; Description ...: Returns all or only visible reminders.
; Syntax.........: _OL_ReminderGet($oOL[, $bOL_IsVisible = True])
; Parameters ....: $oOL           - Outlook object returned by a preceding call to _OL_Open()
;                  $bOL_IsVisible - Optional: Only return visible reminders (default = True)
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - String representing the title
;                  |1 - OlObjectClass constant indicating the object's class of the specified outlook item (see element 4)
;                  |2 - Boolean that determines if the reminder is currently visible
;                  |3 - Object corresponding to the Reminder
;                  |4 - Object corresponding to the specified Outlook item (AppointmentItem, MailItem, ContactItem, TaskItem)
;                  |5 - Date that indicates the next date and time the specified reminder will occur
;                  |6 - Date that specifies the original date and time that the specified reminder is set to occur
;                  Failure - Returns "" and sets @error:
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_ReminderGet($oOL, $bOL_IsVisible = True)

	Local $iOL_Index = 1, $aOL_Reminders[$oOL.Reminders.Count + 1][7]
	For $oOL_Reminder In $oOL.Reminders
		If $bOL_IsVisible = False Or ($bOL_IsVisible = True And $oOL_Reminder.IsVisible) Then
			$aOL_Reminders[$iOL_Index][0] = $oOL_Reminder.Caption
			$aOL_Reminders[$iOL_Index][1] = $oOL_Reminder.Item.Class
			$aOL_Reminders[$iOL_Index][2] = $oOL_Reminder.IsVisible
			$aOL_Reminders[$iOL_Index][3] = $oOL_Reminder
			$aOL_Reminders[$iOL_Index][4] = $oOL_Reminder.Item
			$aOL_Reminders[$iOL_Index][5] = $oOL_Reminder.NextReminderDate
			$aOL_Reminders[$iOL_Index][6] = $oOL_Reminder.OriginalReminderDate
			$iOL_Index += 1
		EndIf
	Next
	ReDim $aOL_Reminders[$iOL_Index][UBound($aOL_Reminders, 2)]
	$aOL_Reminders[0][0] = $iOL_Index - 1
	$aOL_Reminders[0][1] = UBound($aOL_Reminders, 2)
	Return $aOL_Reminders

EndFunc   ;==>_OL_ReminderGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_RuleActionGet
; Description ...: Gets all actions for a specified rule.
; Syntax.........: _OL_RuleActionGet($oOL_Rule[, $bOL_Enabled = True])
; Parameters ....: $oOL_Rule    - Rule object returned by a preceding call to _OL_RuleGet in element 0
;                  $bOL_Enabled - Optional: Only returns enabled actions if set to True (default = True)
; Return values .: Success - two-dimensional one based array with the following information:
;                  Elements 0 - 2 are the same for every action type. The other elements (if any) depend on the action type.
;                  |0 - OlRuleActionType constant indicating the type of action that is taken by the rule action
;                  |1 - OlObjectClass constant indicating the class of the rule action
;                  |2 - True if the action is enabled
;                  |AssignToCategoryRuleAction
;                  |3 - Categories assigned to the message separated by the pipe character
;                  |MoveOrCopyRuleAction
;                  |3 - Object of the folder where the message will be copied/moved to
;                  |4 - Name of the folder where the message will be copied/moved to
;                  |SendRuleAction
;                  |3 - Recipients collection (object) that represents the recipient list for the cc/forward/redirect action
;                  |4 - Recipients (string) separated by the pipe character
;                  |MarkAsTaskRuleAction
;                  |3 - String that represents the label of the flag for the message
;                  |4 - constant in the OlMarkInterval enumeration representing the interval before the task is due
;                  |NewItemAlertRuleAction
;                  |3 - Text to be displayed in the new item alert dialog box
;                  |PlaySoundRuleAction
;                  |3 - Full file path to a sound file (.wav)
;                  Failure - Returns "" and sets @error:
;                  |1 - The ActionType can not be handled by this function. @extended contains the ActionType in error
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_RuleActionGet($oOL_Rule, $bOL_Enabled = True)

	Local $iOL_Index = 1
	Local $aOL_Actions[$oOL_Rule.Actions.Count + 1][5] = [[$oOL_Rule.Actions.Count, 5]]
	For $oOL_Action In $oOL_Rule.Actions
		If $bOL_Enabled = False Or $oOL_Action.Enabled = True Then
			; Properties that apply to all action types
			$aOL_Actions[$iOL_Index][0] = $oOL_Action.ActionType
			$aOL_Actions[$iOL_Index][1] = $oOL_Action.Class
			$aOL_Actions[$iOL_Index][2] = $oOL_Action.Enabled
			; Properties that apply to individual action types
			Switch $oOL_Action.ActionType
				Case $olRuleActionAssignToCategory ; AssignToCategoryRuleAction object
					Local $aOL_Categories = $oOL_Action.Categories ; array of strings representing the categories assigned to the message
					$aOL_Actions[$iOL_Index][3] = _ArrayToString($aOL_Categories)
				Case $olRuleActionMoveToFolder, $olRuleActionCopyToFolder ; MoveOrCopyRuleAction object
					$aOL_Actions[$iOL_Index][3] = $oOL_Action.Folder ; Folder object that represents the folder to which the rule moves or copies the message
					If IsObj($oOL_Action.Folder) Then $aOL_Actions[$iOL_Index][4] = $oOL_Action.Folder.Name
				Case $olRuleActionCcMessage, $olRuleActionForward, $olRuleActionForwardAsAttachment, $olRuleActionRedirect ; SendRuleAction object
					$aOL_Actions[$iOL_Index][3] = $oOL_Action.Recipients ; collection that represents the recipient list for the send action
					Local $sOL_Recipients
					For $oOL_Recipient In $oOL_Action.Recipients
						$sOL_Recipients = $sOL_Recipients & $oOL_Recipient.Name & "|"
					Next
					$aOL_Actions[$iOL_Index][4] = StringLeft($sOL_Recipients, StringLen($sOL_Recipients) - 1)
				Case $olRuleActionMarkAsTask ; MarkAsTaskRuleAction object
					$aOL_Actions[$iOL_Index][3] = $oOL_Action.FlagTo ; String that represents the label of the flag for the message
					$aOL_Actions[$iOL_Index][4] = $oOL_Action.MarkInterval ; constant in the OlMarkInterval enumeration representing the interval before the task is due
				Case $olRuleActionNewItemAlert ; NewItemAlertRuleAction object
					$aOL_Actions[$iOL_Index][3] = $oOL_Action.Text ; String that represents the text displayed in the new item alert dialog box
				Case $olRuleActionPlaySound ; PlaySoundRuleAction object
					$aOL_Actions[$iOL_Index][3] = $oOL_Action.FilePath ; Full file path to a sound file (.wav)
				Case $olRuleActionClearCategories, $olRuleActionDelete, $olRuleActionDeletePermanently, _ ; Actions without additional properties
						$olRuleActionDesktopAlert, $olRuleActionNotifyDelivery, $olRuleActionNotifyRead, $olRuleActionStop
				Case $olRuleActionServerReply, $olRuleActionTemplate, $olRuleActionFlagForActionInDays, _ ; Types not yet handled by Outlook object model
						$olRuleActionFlagColor, $olRuleActionFlagClear, $olRuleActionImportance, $olRuleActionSensitivity, _
						$olRuleActionPrint, $olRuleActionMarkRead, $olRuleActionDefer, $olRuleActionStartApplication
				Case Else
					Return SetError(1, $oOL_Action.ActionType, "")
			EndSwitch
			$iOL_Index += 1
		EndIf
	Next
	ReDim $aOL_Actions[$iOL_Index][UBound($aOL_Actions, 2)]
	$aOL_Actions[0][0] = $iOL_Index - 1
	Return $aOL_Actions

EndFunc   ;==>_OL_RuleActionGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_RuleActionSet
; Description ...: Adds a new or overwrites an existing action of an existing rule of the specified store.
; Syntax.........: _OL_RuleActionSet($oOL, $sOL_Store, $sOL_RuleName, $iOL_RuleActionType, $bOL_Enabled[, $sOL_P1 = ""[, $sOL_P2 = ""]])
; Parameters ....: $oOL                - Outlook object returned by a preceding call to _OL_Open()
;                  $sOL_Store          - Name of the Store where the rule will be defined. "*" = your default store
;                  $sOL_RuleName       - Name of the rule
;                  $iOL_RuleActionType - Type of the rule action. Please see the OlRuleActionType enumeration
;                  $bOL_Enabled        - True sets the rule action to enabled
;                  $sOL_P1             - Optional: Data to create the rule action depending on $iOL_RuleActionType. Please check remarks for details
;                  $sOL_P2             - Optional: Same as $sOL_P1
; Return values .: Success - Object of the added action
;                  Failure - Returns 0 and sets @error:
;                  |1 - Error accessing specified store. For details please check @extended
;                  |2 - Error accessing the rule collection. For details please check @extended
;                  |3 - Error accessing the specified rule. For details please check @extended
;                  |4 - Error creating the action for the specified rule. For details please check @extended
;                  |5 - Error saving the specified rule. For details please check @extended
;                  |6 - $sOL_P1 is not an folder object for rule action type $olRuleActionMoveToFolder or $olRuleActionCopyToFolder
;                  |7 - Error adding a recipient. For details please check @extended
;                  |8 - Error resolving recipients. @extended is the 1-based number of the recipient in error
;                  |9 - The specified rule action is not valid for the rule type (send/receive)
;                  |10 - The specified wav sound file could not be found
;                  |11 - The specified $iOL_RuleActionType is invalid
;                  |12 - The specified $iOL_RuleActionType is not supported by the Outlook object model at the moment
; Author ........: water
; Modified ......:
; Remarks .......: Not all possible rule actions can be created using the COM model.
;                  To remove an action from a rule set $bOL_Enabled to False.
;                  Remarks for different types of actions:
;+
;                  $olRuleActionAssignToCategory:
;                  $sOL_P1: Specify a string of categories to be assigned to the message separated by the pipe character e.g. "Birthday|Private"
;+
;                  $olRuleActionMoveToFolder, $olRuleActionCopyToFolder:
;                  $sOL_P1: Folder object that represents the folder to which the rule moves or copies the message
;+
;                  $olRuleActionCcMessage, $olRuleActionForward, $olRuleActionForwardAsAttachment, $olRuleActionRedirect
;                  $sOL_P1: collection that represents the recipient list for the send action e.g. "George Smith;John Doe"
;+
;                  $olRuleActionMarkAsTask:
;                  $sOL_P1: String that represents the label of the flag for the message e.g. "Very urgent!"
;                  $sOL_P2: constant in the OlMarkInterval enumeration representing the interval before the task is due
;+
;                  $olRuleActionNewItemAlert:
;                  $sOL_P1: String that represents the text displayed in the new item alert dialog box
;+
;                  $olRuleActionPlaySound:
;                  $sOL_P1: Full file path to a sound file (.wav) e.g. "C:\Windows\Media\Tada.wav"
;+
;                  $olRuleActionClearCategories, $olRuleActionDelete, $olRuleActionDeletePermanently, $olRuleActionDesktopAlert,
;                  $olRuleActionNotifyDelivery, $olRuleActionNotifyRead, $olRuleActionStop:
;                  No parameters need to be set
; Related .......:
; Link ..........: http://msdn.microsoft.com/en-us/library/bb206764(v=office.12).aspx
; Example .......: Yes
; ===============================================================================================================================
Func _OL_RuleActionSet($oOL, $sOL_Store, $sOL_RuleName, $iOL_RuleActionType, $bOL_Enabled, $sOL_P1 = "", $sOL_P2 = "")

	Local $oOL_Action
	If $sOL_Store = "*" Then $sOL_Store = $oOL.Session.DefaultStore.DisplayName
	Local $oOL_Store = $oOL.Session.Stores.Item($sOL_Store)
	If @error Then Return SetError(1, @error, 0)
	Local $oOL_Rules = $oOL_Store.GetRules()
	If @error Then Return SetError(2, @error, 0)
	Local $oOL_Rule = $oOL_Rules.Item($sOL_RuleName)
	If @error Then Return SetError(3, @error, 0)
	; Properties that apply to individual action types
	Switch $iOL_RuleActionType
		Case $olRuleActionAssignToCategory ; AssignToCategoryRuleAction object
			$oOL_Action = $oOL_Rule.Actions.AssignToCategory
			If @error Then Return SetError(4, @error, 0)
			$oOL_Action.Enabled = $bOL_Enabled
			$oOL_Action.Categories = StringSplit($sOL_P1, "|", 2) ; array of strings representing the categories assigned to the message
			SetError(0) ; Reset an error raised by StringSplit when nothing to split
		Case $olRuleActionMoveToFolder, $olRuleActionCopyToFolder ; MoveOrCopyRuleAction object
			If $oOL_Rule.RuleType = $olRuleSend And $iOL_RuleActionType = $olRuleActionMoveToFolder Then Return SetError(9, 0, 0)
			If Not IsObj($sOL_P1) Then Return SetError(6, 0, 0)
			If $iOL_RuleActionType = $olRuleActionMoveToFolder Then $oOL_Action = $oOL_Rule.Actions.MoveToFolder
			If $iOL_RuleActionType = $olRuleActionCopyToFolder Then $oOL_Action = $oOL_Rule.Actions.CopyToFolder
			If @error Then Return SetError(4, @error, 0)
			$oOL_Action.Enabled = $bOL_Enabled
			$oOL_Action.Folder = $sOL_P1 ; Folder object that represents the folder to which the rule moves or copies the message
		Case $olRuleActionCcMessage, $olRuleActionForward, $olRuleActionForwardAsAttachment, $olRuleActionRedirect ; SendRuleAction object
			If $oOL_Rule.RuleType = $olRuleReceive And $iOL_RuleActionType = $olRuleActionCcMessage Then Return SetError(9, 0, 0)
			If $oOL_Rule.RuleType = $olRuleSend And $iOL_RuleActionType = $olRuleActionForward Then Return SetError(9, 0, 0)
			If $oOL_Rule.RuleType = $olRuleSend And $iOL_RuleActionType = $olRuleActionForwardAsAttachment Then Return SetError(9, 0, 0)
			If $oOL_Rule.RuleType = $olRuleSend And $iOL_RuleActionType = $olRuleActionRedirect Then Return SetError(9, 0, 0)
			If $iOL_RuleActionType = $olRuleActionCcMessage Then $oOL_Action = $oOL_Rule.Actions.CC
			If $iOL_RuleActionType = $olRuleActionForward Then $oOL_Action = $oOL_Rule.Actions.Forward
			If $iOL_RuleActionType = $olRuleActionForwardAsAttachment Then $oOL_Action = $oOL_Rule.Actions.ForwardAsAttachment
			If $iOL_RuleActionType = $olRuleActionRedirect Then $oOL_Action = $oOL_Rule.Actions.Redirect
			If @error Then Return SetError(4, @error, 0)
			$oOL_Action.Enabled = $bOL_Enabled
			Local $aOL_Recipients = StringSplit($sOL_P1, ";")
			Local $ol_Recipient
			SetError(0) ; Reset an error raised by StringSplit when nothing to split
			For $iOL_Index = 1 To $aOL_Recipients[0] ; collection that represents the recipient list for the send action
				$ol_Recipient = $oOL_Action.Recipients.Add($aOL_Recipients[$iOL_Index])
				If @error Then Return SetError(7, @error, 0)
				If $ol_Recipient.Resolve = False Then Return SetError(8, $iOL_Index, 0)
			Next
		Case $olRuleActionMarkAsTask ; MarkAsTaskRuleAction object
			If $oOL_Rule.RuleType = $olRuleSend Then Return SetError(9, 0, 0)
			$oOL_Action = $oOL_Rule.Actions.MarkAsTask
			If @error Then Return SetError(4, @error, 0)
			$oOL_Action.Enabled = $bOL_Enabled
			$oOL_Action.FlagTo = $sOL_P1 ; String that represents the label of the flag for the message
			$oOL_Action.MarkInterval = $sOL_P2 ; constant in the OlMarkInterval enumeration representing the interval before the task is due
		Case $olRuleActionNewItemAlert ; NewItemAlertRuleAction object
			If $oOL_Rule.RuleType = $olRuleSend Then Return SetError(9, 0, 0)
			$oOL_Action = $oOL_Rule.Actions.NewItemAlert
			If @error Then Return SetError(4, @error, 0)
			$oOL_Action.Enabled = $bOL_Enabled
			$oOL_Action.Text = $sOL_P1 ; String that represents the text displayed in the new item alert dialog box
		Case $olRuleActionPlaySound ; PlaySoundRuleAction object
			If $oOL_Rule.RuleType = $olRuleSend Then Return SetError(9, 0, 0)
			If FileExists($sOL_P1) = 0 Then Return SetError(10, 0, 0)
			$oOL_Action = $oOL_Rule.Actions.PlaySound
			If @error Then Return SetError(4, @error, 0)
			$oOL_Action.Enabled = $bOL_Enabled
			$oOL_Action.FilePath = $sOL_P1 ; Full file path to a sound file (.wav)
		Case $olRuleActionClearCategories, $olRuleActionDelete, $olRuleActionDeletePermanently, _ ; Actions without additional properties
				$olRuleActionDesktopAlert, $olRuleActionNotifyDelivery, $olRuleActionNotifyRead, $olRuleActionStop
			If $oOL_Rule.RuleType = $olRuleReceive And $iOL_RuleActionType = $olRuleActionNotifyDelivery Then Return SetError(9, 0, 0)
			If $oOL_Rule.RuleType = $olRuleReceive And $iOL_RuleActionType = $olRuleActionNotifyRead Then Return SetError(9, 0, 0)
			If $oOL_Rule.RuleType = $olRuleSend And $iOL_RuleActionType = $olRuleActionDelete Then Return SetError(9, 0, 0)
			If $oOL_Rule.RuleType = $olRuleSend And $iOL_RuleActionType = $olRuleActionDeletePermanently Then Return SetError(9, 0, 0)
			If $oOL_Rule.RuleType = $olRuleSend And $iOL_RuleActionType = $olRuleActionDesktopAlert Then Return SetError(9, 0, 0)
			If $iOL_RuleActionType = $olRuleActionNotifyDelivery Then $oOL_Action = $oOL_Rule.Actions.NotifyDelivery
			If $iOL_RuleActionType = $olRuleActionNotifyRead Then $oOL_Action = $oOL_Rule.Actions.NotifyRead
			If $iOL_RuleActionType = $olRuleActionDelete Then $oOL_Action = $oOL_Rule.Actions.Delete
			If $iOL_RuleActionType = $olRuleActionDeletePermanently Then $oOL_Action = $oOL_Rule.Actions.DeletePermanently
			If $iOL_RuleActionType = $olRuleActionDesktopAlert Then $oOL_Action = $oOL_Rule.Actions.DesktopAlert
			If @error Then Return SetError(4, @error, 0)
			$oOL_Action.Enabled = $bOL_Enabled
		Case $olRuleActionServerReply, $olRuleActionTemplate, $olRuleActionFlagForActionInDays, _ ; Types not yet handled by Outlook object model
				$olRuleActionFlagColor, $olRuleActionFlagClear, $olRuleActionImportance, $olRuleActionSensitivity, _
				$olRuleActionPrint, $olRuleActionMarkRead, $olRuleActionDefer, $olRuleActionStartApplication
			Return SetError(12, $iOL_RuleActionType, 0)
		Case Else
			Return SetError(11, $iOL_RuleActionType, 0)
	EndSwitch
	; Update the server
	$oOL_Rules.Save
	If @error Then Return SetError(5, @error, 0)
	Return $oOL_Action

EndFunc   ;==>_OL_RuleActionSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_RuleAdd
; Description ...: Adds a new rule to the specified store.
; Syntax.........: _OL_RuleAdd($oOL, $sOL_Store, $sOL_RuleName[, $bOL_Enabled = True[, $iOL_RuleType = $olRuleReceive[, $iOL_ExecutionOrder = 0]]])
; Parameters ....: $oOL                - Outlook object returned by a preceding call to _OL_Open()
;                  $sOL_Store          - Name of the Store where the rule will be defined. "*" = your default store
;                  $sOL_RuleName       - Name of the rule
;                  $bOL_Enabled        - Optional: True sets the rule to enabled (default = True)
;                  $iOL_RuleType       - Optional: Can be $olRuleSend or $olRuleReceive (default = $olRuleReceive)
;                  $iOL_ExecutionOrder - Optional: Integer indicating the order of execution of the rule among other rules (default = 1)
; Return values .: Success - Object of the created rule
;                  Failure - Returns 0 and sets @error:
;                  |1 - Rule already exists for the specified store
;                  |2 - Error returned by method .GetRules. For more information check @extended
;                  |3 - Error creating the rule. For more information check @extended
;                  |4 - Error saving the rule collection. For more information check @extended
; Author ........: water
; Modified ......:
; Remarks .......: A newly added rule is always a client rule till you add actions which can be executed on the server
; Related .......:
; Link ..........: http://www.outlookpower.com/issues/issue200904/00002353001.html
; Example .......: Yes
; ===============================================================================================================================
Func _OL_RuleAdd($oOL, $sOL_Store, $sOL_RuleName, $bOL_Enabled = True, $iOL_RuleType = $olRuleReceive, $iOL_ExecutionOrder = 1)

	If $sOL_Store = "*" Then $sOL_Store = $oOL.Session.DefaultStore.DisplayName
	Local $oOL_Rules = $oOL.Session.Stores.Item($sOL_Store).GetRules
	If @error Then Return SetError(2, @error, 0)
	For $oOL_Rule In $oOL_Rules
		If $oOL_Rule.Name = $sOL_RuleName Then Return SetError(1, 0, 0)
	Next
	$oOL_Rule = $oOL_Rules.Create($sOL_RuleName, $iOL_RuleType)
	If @error Then Return SetError(3, @error, 0)
	$oOL_Rule.Enabled = $bOL_Enabled
	$oOL_Rule.ExecutionOrder = $iOL_ExecutionOrder
	$oOL_Rules.Save
	If @error Then Return SetError(4, @error, 0)
	Return $oOL_Rule

EndFunc   ;==>_OL_RuleAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_RuleConditionGet
; Description ...: Gets all conditions or condition exceptions for a specified rule.
; Syntax.........: _OL_RuleConditionGet($oOL_Rule[, $bOL_Enabled = True[, $bOL_Exceptions = False]])
; Parameters ....: $oOL_Rule       - Rule object returned by a preceding call to _OL_RuleGet in element 0
;                  $bOL_Enabled    - Optional: Only returns enabled conditions if set to True (default = True)
;                  $bOL_Exceptions - Optional: Only returns defined exceptions to the conditions if set to True (default = False)
; Return values .: Success - two-dimensional one based array with the following information:
;                  Elements 0 - 2 are the same for every condition type. The other elements (if any) depend on the condition type.
;                  |0 - OlRuleConditionType constant indicating the type of condition that is taken by the rule condition
;                  |1 - OlObjectClass constant indicating the class of the rule condition
;                  |2 - True if the condition is enabled
;                  |AccountRuleCondition
;                  |3 - Account object that represents the account used to evaluate the rule condition
;                  |AddressRuleCondition
;                  |3 - array of strings to evaluate the address rule condition
;                  |CategoryRuleCondition
;                  |3 - array of strings representing the categories evaluated by the rule condition
;                  |FormNameRuleCondition
;                  |3 - array of form identifiers
;                  |FromRssFeedRuleCondition
;                  |3 - array of String elements that represent the RSS subscriptions
;                  |ImportanceRuleCondition
;                  |3 - OlImportance constant indicating the relative level of importance for the message
;                  |SenderInAddressListRuleCondition
;                  |3 - AddressList object that represents the address list
;                  |4 - Name of the addresslist object
;                  |TextRuleCondition
;                  |3 - array of String elements that represents the text to be evaluated
;                  |ToOrFromRuleCondition
;                  |3 - collection that represents the recipient list for the evaluation of the rule condition
;                  |4 - Recipients (string) separated by the pipe character
;                  Failure - Returns "" and sets @error:
;                  |1 - The ConditionType can not be handled by this function. @extended contains the ConditionType in error
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_RuleConditionGet($oOL_Rule, $bOL_Enabled = True, $bOL_Exceptions = False)

	Local $iOL_Index = 1
	Local $oOL_ConditionOrException = $oOL_Rule.Conditions
	If $bOL_Exceptions = True Then $oOL_ConditionOrException = $oOL_Rule.Exceptions
	Local $aOL_Conditions[$oOL_ConditionOrException.Count + 1][5] = [[$oOL_ConditionOrException.Count, 5]]
	For $oOL_Object In $oOL_ConditionOrException
		If $bOL_Enabled = False Or $oOL_Object.Enabled = True Then
			; Properties that apply to all condition types
			$aOL_Conditions[$iOL_Index][0] = $oOL_Object.ConditionType
			$aOL_Conditions[$iOL_Index][1] = $oOL_Object.Class
			$aOL_Conditions[$iOL_Index][2] = $oOL_Object.Enabled
			; Properties that apply to individual condition types
			Switch $oOL_Object.ConditionType
				Case $olConditionAccount ; AccountRuleCondition object
					$aOL_Conditions[$iOL_Index][3] = $oOL_Object.Account ; Account object that represents the account used to evaluate the rule condition
				Case $olConditionRecipientAddress, $olConditionSenderAddress ; AddressRuleCondition object
					Local $aOL_Address = $oOL_Object.Address ; array of strings to evaluate the address rule condition
					$aOL_Conditions[$iOL_Index][3] = _ArrayToString($aOL_Address)
				Case $olConditionCategory ; CategoryRuleCondition object
					Local $aOL_Categories = $oOL_Object.Categories ; array of strings representing the categories evaluated by the rule condition
					$aOL_Conditions[$iOL_Index][3] = _ArrayToString($aOL_Categories)
				Case $olConditionFormName ; FormNameRuleCondition object
					Local $aOL_Forms = $oOL_Object.FormName ; array of form identifiers
					$aOL_Conditions[$iOL_Index][3] = _ArrayToString($aOL_Forms)
				Case $olConditionFromRssFeed ; FromRssFeedRuleCondition object
					Local $aOL_Feeds = $oOL_Object.FromRssFeed ; array of String elements that represent the RSS subscriptions
					$aOL_Conditions[$iOL_Index][3] = _ArrayToString($aOL_Feeds)
				Case $olConditionImportance ; ImportanceRuleCondition object
					$aOL_Conditions[$iOL_Index][3] = $oOL_Object.Importance ; OlImportance constant indicating the relative level of importance for the message
				Case $olConditionSenderInAddressBook ; SenderInAddressListRuleCondition object
					$aOL_Conditions[$iOL_Index][3] = $oOL_Object.AddressList ; AddressList object that represents the address list
					If IsObj($oOL_Object.AddressList) Then $aOL_Conditions[$iOL_Index][4] = $oOL_Object.AddressList.Name
				Case $olConditionBody, $olConditionBodyOrSubject, $olConditionMessageHeader, $olConditionSubject ; TextRuleCondition object
					Local $aOL_Text = $oOL_Object.Text ; array of String elements that represents the text to be evaluated
					$aOL_Conditions[$iOL_Index][3] = _ArrayToString($aOL_Text)
					; Conditions that the Rules object model supports for rules created by the Wizard but not for those created by the object model
				Case $olConditionSentTo, $olConditionFrom ; ToOrFromRuleCondition object
					$aOL_Conditions[$iOL_Index][3] = $oOL_Object.Recipients
					Local $sOL_Recipients
					For $oOL_Recipient In $oOL_Object.Recipients
						$sOL_Recipients = $sOL_Recipients & $oOL_Recipient.Name & "|"
					Next
					$aOL_Conditions[$iOL_Index][4] = StringLeft($sOL_Recipients, StringLen($sOL_Recipients) - 1)
				Case $olConditionAnyCategory, $olConditionCc, $olConditionFromAnyRssFeed, $olConditionHasAttachment, _ ; Conditions without additional properties
						$olConditionLocalMachineOnly, $olConditionMeetingInviteOrUpdate, $olConditionNotTo, $olConditionOnlyToMe, _
						$olConditionOtherMachine, $olConditionTo, $olConditionToOrCc
				Case $olConditionSensitivity, $olConditionFlaggedForAction, $olConditionOOF, $olConditionSizeRange, _ ; Types not yet handled by Outlook object model
						$olConditionDateRange, $olConditionProperty
				Case Else
					Return SetError(1, $oOL_Object.ConditionType, "")
			EndSwitch
			$iOL_Index += 1
		EndIf
	Next
	ReDim $aOL_Conditions[$iOL_Index][UBound($aOL_Conditions, 2)]
	$aOL_Conditions[0][0] = $iOL_Index - 1
	Return $aOL_Conditions

EndFunc   ;==>_OL_RuleConditionGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_RuleConditionSet
; Description ...: Adds a new or overwrites an existing condition or condition exception to an existing rule of the specified store.
; Syntax.........: _OL_RuleConditionSet($oOL, $sOL_Store, $sOL_RuleName, $iOL_RuleConditionType, $bOL_Enabled[, $sOL_P1 = ""])
; Parameters ....: $oOL                   - Outlook object returned by a preceding call to _OL_Open()
;                  $sOL_Store             - Name of the Store where the rule will be defined. "*" = your default store
;                  $sOL_RuleName          - Name of the rule
;                  $iOL_RuleConditionType - Type of the rule condition. Please see the OlRuleCOnditionType enumeration
;                  $bOL_Enabled           - Optional: True sets the rule condition to enabled (default = True)
;                  $bOL_Exceptions        - Optional: Sets exceptions to the rule conditions if set to True (default = False)
;                  $sOL_P1                - Optional: Data to create the rule condition depending on $iOL_RuleConditionType
; Return values .: Success - Object of the added condition
;                  Failure - Returns 0 and sets @error:
;                  |1 - Error accessing specified store. For details please check @extended
;                  |2 - Error accessing the rule collection. For details please check @extended
;                  |3 - Error accessing the specified rule. For details please check @extended
;                  |4 - Error creating the condition for the specified rule. For details please check @extended
;                  |5 - Error saving the specified rule. For details please check @extended
;                  |6 - The specified rule condition is not valid for the rule type (send/receive)
;                  |7 - Error adding a recipient. For details please check @extended
;                  |8 - Error resolving recipients. @extended is the 1-based number of the recipient in error
;                  |9 - The specified $iOL_RuleConditionType is invalid
;                  |10 - The specified $iOL_RuleConditionType is not supported by the Outlook object model at the moment
; Author ........: water
; Modified ......:
; Remarks .......: Not all possible rule conditions can be created using the COM model.
;                  To remove an action from a rule set $bOL_Enabled to False.
;                  Remarks for different types of conditions:
;+
;                  $olConditionAccount:
;                  $sOL_P1: Account object that represents the account used to evaluate the rule condition
;+
;                  $olConditionBody, $olConditionBodyOrSubject, $olConditionMessageHeader, $olConditionSubject:
;                  $sOL_P1: Specify a string of elements that represent the text to be evaluated separated by the pipe character e.g. "Vacation|return"
;+
;                  $olConditionCategory:
;                  $sOL_P1: Specify a string of elements that represent the categories separated by the pipe character e.g. "Birthday|Private"
;+
;                  $olConditionFormName:
;                  $sOL_P1: Specify a string of form identifiers to be evaluated by the rule condition separated by the pipe character
;+
;                  $olConditionFrom, $olConditionSentTo:
;                  $sOL_P1: Specify a string of elements that represents the recipient list separated by ";" e.g. "George Smith;John Doe"
;+
;                  $olFromRSSFeed:
;                  $sOL_P1: Specify a string of elements that represent the RSS subscriptions separated by the pipe character
;+
;                  $olConditionImportance:
;                  $sOL_P1: OlImportance constant indicating the relative level of importance
;+
;                  $olConditionRecipientAddress, $olConditionSenderAddress:
;                  $sOL_P1: Specify a string of elements to evaluate the address rule condition separated by ";"
;+
;                  $olConditionSenderInAddressList:
;                  $sOL_P1: AddressList object that represents the address list used to evaluate the rule condition
;+
;                  $olConditionAnyCategory, $olConditionCC, $olConditionFromAnyRSSFeed, $olConditionHasAttachment, $olConditionMeetingInviteOrUpdate,
;                  $olConditionNotTo, $olConditionLocalMachineOnly, $olConditionOnlyToMe, $olConditionTo, $olConditionToOrCC:
;                  No parameters need to be set
; Related .......:
; Link ..........: http://msdn.microsoft.com/en-us/library/bb206766(v=office.12).aspx
; Example .......: Yes
; ===============================================================================================================================
Func _OL_RuleConditionSet($oOL, $sOL_Store, $sOL_RuleName, $iOL_RuleConditionType, $bOL_Enabled = True, $bOL_Exceptions = False, $sOL_P1 = "")

	Local $oOL_Object
	If $sOL_Store = "*" Then $sOL_Store = $oOL.Session.DefaultStore.DisplayName
	Local $oOL_Store = $oOL.Session.Stores.Item($sOL_Store)
	If @error Then Return SetError(1, @error, 0)
	Local $oOL_Rules = $oOL_Store.GetRules()
	If @error Then Return SetError(2, @error, 0)
	Local $oOL_Rule = $oOL_Rules.Item($sOL_RuleName)
	If @error Then Return SetError(3, @error, 0)
	Local $oOL_ConditionOrException = $oOL_Rule.Conditions
	If $bOL_Exceptions = True Then $oOL_ConditionOrException = $oOL_Rule.Exceptions
	; Properties that apply to individual condition types
	Switch $iOL_RuleConditionType
		Case $olConditionAccount ; AccountRuleCondition object
			$oOL_Object = $oOL_ConditionOrException.Account
			If @error Then Return SetError(4, @error, 0)
			$oOL_Object.Enabled = $bOL_Enabled
			$oOL_Object.Account = $sOL_P1 ; Account object that represents the account used to evaluate the rule condition
		Case $olConditionBody, $olConditionBodyOrSubject, $olConditionMessageHeader, $olConditionSubject ; TextRuleCondition object
			If $oOL_Rule.RuleType = $olRuleSend And $iOL_RuleConditionType = $olConditionMessageHeader Then Return SetError(6, 0, 0)
			If $iOL_RuleConditionType = $olConditionBody Then $oOL_Object = $oOL_ConditionOrException.Body
			If $iOL_RuleConditionType = $olConditionBodyOrSubject Then $oOL_Object = $oOL_ConditionOrException.BodyOrSubject
			If $iOL_RuleConditionType = $olConditionMessageHeader Then $oOL_Object = $oOL_ConditionOrException.MessageHeader
			If $iOL_RuleConditionType = $olConditionSubject Then $oOL_Object = $oOL_ConditionOrException.Subject
			If @error Then Return SetError(4, @error, 0)
			$oOL_Object.Enabled = $bOL_Enabled
			$oOL_Object.Text = StringSplit($sOL_P1, "|", 2) ; array of string elements that represents the text to be evaluated
			SetError(0) ; Reset an error raised by StringSplit when nothing to split
		Case $olConditionCategory ; CategoryRuleCondition object
			$oOL_Object = $oOL_ConditionOrException.Category
			If @error Then Return SetError(4, @error, 0)
			$oOL_Object.Enabled = $bOL_Enabled
			$oOL_Object.Categories = StringSplit($sOL_P1, "|", 2) ; array of strings representing the categories assigned to the message
			SetError(0) ; Reset an error raised by StringSplit when nothing to split
		Case $olConditionFormName ; FormNameRuleCondition object
			$oOL_Object = $oOL_ConditionOrException.FormName
			If @error Then Return SetError(4, @error, 0)
			$oOL_Object.Enabled = $bOL_Enabled
			$oOL_Object.FormName = StringSplit($sOL_P1, "|", 2) ; represents an array of form identifiers to be evaluated by the rule condition
			SetError(0) ; Reset an error raised by StringSplit when nothing to split
		Case $olConditionFrom, $olConditionSentTo ; ToOrFromRuleCondition object
			If $oOL_Rule.RuleType = $olRuleSend And $iOL_RuleConditionType = $olConditionFrom Then Return SetError(6, 0, 0)
			If $iOL_RuleConditionType = $olConditionFrom Then $oOL_Object = $oOL_ConditionOrException.From
			If $iOL_RuleConditionType = $olConditionFrom Then $oOL_Object = $oOL_ConditionOrException.SentTo
			If @error Then Return SetError(4, @error, 0)
			$oOL_Object.Enabled = $bOL_Enabled
			Local $aOL_Recipients = StringSplit($sOL_P1, ";")
			Local $ol_Recipient
			SetError(0) ; Reset an error raised by StringSplit when nothing to split
			For $iOL_Index = 1 To $aOL_Recipients[0] ; collection that represents the recipient list
				$ol_Recipient = $oOL_Object.Recipients.Add($aOL_Recipients[$iOL_Index])
				If @error Then Return SetError(7, @error, 0)
				If $ol_Recipient.Resolve = False Then Return SetError(8, $iOL_Index, 0)
			Next
		Case $olConditionFromRssFeed ; FromRSSFeedRuleCondition object
			If $oOL_Rule.RuleType = $olRuleSend Then Return SetError(6, 0, 0)
			$oOL_Object = $oOL_ConditionOrException.FromRSSFeed
			If @error Then Return SetError(4, @error, 0)
			$oOL_Object.Enabled = $bOL_Enabled
			$oOL_Object.FromRSSFeed = StringSplit($sOL_P1, "|", 2) ; array of string elements that represent the RSS subscriptions
			SetError(0) ; Reset an error raised by StringSplit when nothing to split
		Case $olConditionImportance ; ImportanceRuleCondition object
			$oOL_Object = $oOL_ConditionOrException.Importance
			If @error Then Return SetError(4, @error, 0)
			$oOL_Object.Enabled = $bOL_Enabled
			$oOL_Object.Importance = $sOL_P1 ; OlImportance constant indicating the relative level of importance
		Case $olConditionRecipientAddress, $olConditionSenderAddress ; AddressRuleCondition object
			If $oOL_Rule.RuleType = $olRuleSend And $iOL_RuleConditionType = $olConditionSenderAddress Then Return SetError(6, 0, 0)
			If $iOL_RuleConditionType = $olConditionRecipientAddress Then $oOL_Object = $oOL_ConditionOrException.RecipientAddress
			If $iOL_RuleConditionType = $olConditionSenderAddress Then $oOL_Object = $oOL_ConditionOrException.SenderAddress
			If @error Then Return SetError(4, @error, 0)
			$oOL_Object.Enabled = $bOL_Enabled
			$oOL_Object.Address = StringSplit($sOL_P1, ";", 2) ; array of string elements to evaluate the address rule condition
			SetError(0) ; Reset an error raised by StringSplit when nothing to split
		Case $olConditionSenderInAddressBook ; SenderInAddressListRuleCondition object
			If $oOL_Rule.RuleType = $olRuleSend Then Return SetError(6, 0, 0)
			$oOL_Object = $oOL_ConditionOrException.SenderInAddressList
			If @error Then Return SetError(4, @error, 0)
			$oOL_Object.Enabled = $bOL_Enabled
			$oOL_Object.AddressList = $sOL_P1 ; AddressList object that represents the address list used to evaluate the rule condition
		Case $olConditionAnyCategory, $olConditionCc, $olConditionFromAnyRssFeed, $olConditionHasAttachment, _ ; Conditions without additional properties
				$olConditionMeetingInviteOrUpdate, $olConditionNotTo, $olConditionLocalMachineOnly, $olConditionOnlyToMe, $olConditionTo, _
				$olConditionToOrCc
			If $oOL_Rule.RuleType = $olRuleSend And $iOL_RuleConditionType = $olConditionFromAnyRssFeed Then Return SetError(6, 0, 0)
			If $oOL_Rule.RuleType = $olRuleSend And $iOL_RuleConditionType = $olConditionNotTo Then Return SetError(6, 0, 0)
			If $oOL_Rule.RuleType = $olRuleSend And $iOL_RuleConditionType = $olConditionOnlyToMe Then Return SetError(6, 0, 0)
			If $oOL_Rule.RuleType = $olRuleSend And $iOL_RuleConditionType = $olConditionTo Then Return SetError(6, 0, 0)
			If $oOL_Rule.RuleType = $olRuleSend And $iOL_RuleConditionType = $olConditionToOrCc Then Return SetError(6, 0, 0)
			If $iOL_RuleConditionType = $olConditionAnyCategory Then $oOL_Object = $oOL_ConditionOrException.AnyCategory
			If $iOL_RuleConditionType = $olConditionCc Then $oOL_Object = $oOL_ConditionOrException.CC
			If $iOL_RuleConditionType = $olConditionFromAnyRssFeed Then $oOL_Object = $oOL_ConditionOrException.FromAnyRSSFeed
			If $iOL_RuleConditionType = $olConditionHasAttachment Then $oOL_Object = $oOL_ConditionOrException.HasAttachment
			If $iOL_RuleConditionType = $olConditionMeetingInviteOrUpdate Then $oOL_Object = $oOL_ConditionOrException.MeetingInviteOrUpdate
			If $iOL_RuleConditionType = $olConditionNotTo Then $oOL_Object = $oOL_ConditionOrException.NotTo
			If $iOL_RuleConditionType = $olConditionLocalMachineOnly Then $oOL_Object = $oOL_ConditionOrException.OnLocalMachine
			If $iOL_RuleConditionType = $olConditionOnlyToMe Then $oOL_Object = $oOL_ConditionOrException.OnlyToMe
			If $iOL_RuleConditionType = $olConditionTo Then $oOL_Object = $oOL_ConditionOrException.ToMe
			If $iOL_RuleConditionType = $olConditionToOrCc Then $oOL_Object = $oOL_ConditionOrException.ToOrCC
			If @error Then Return SetError(4, @error, 0)
			$oOL_Object.Enabled = $bOL_Enabled
		Case $olConditionDateRange, $olConditionFlaggedForAction, $olConditionOOF, $olConditionOtherMachine, _ ; Types not yet handled by Outlook object model
				$olConditionProperty, $olConditionSensitivity, $olConditionSizeRange, $olConditionUnknown
			Return SetError(10, $iOL_RuleConditionType, 0)
		Case Else
			Return SetError(9, $iOL_RuleConditionType, 0)
	EndSwitch
	; Update the server
	$oOL_Rules.Save
	If @error Then Return SetError(5, @error, 0)
	Return $oOL_Object

EndFunc   ;==>_OL_RuleConditionSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_RuleDelete
; Description ...: Deletes a rule from the specified store.
; Syntax.........: _OL_RuleDelete($oOL, $sOL_Store, $sOL_RuleName)
; Parameters ....: $oOL          - Outlook object returned by a preceding call to _OL_Open()
;                  $sOL_Store    - Name of the Store where the rule will be deleted from. "*" = your default store
;                  $sOL_RuleName - Name of the rule to be deleted
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - Rule doesn't exist in the specified store
;                  |2 - Error returned by method .GetRules. For more information check @extended
;                  |3 - Error deleting the rule. For more information check @extended
;                  |4 - Error saving the changed rules. For details please check @extended
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_RuleDelete($oOL, $sOL_Store, $sOL_RuleName)

	If $sOL_Store = "*" Then $sOL_Store = $oOL.Session.DefaultStore.DisplayName
	Local $oOL_Rules = $oOL.Session.Stores.Item($sOL_Store).GetRules
	If @error Then Return SetError(2, @error, 0)
	Local $bOL_Found = False
	For $oOL_Rule In $oOL_Rules
		If $oOL_Rule.Name = $sOL_RuleName Then $bOL_Found = True
	Next
	If $bOL_Found = False Then Return SetError(1, 0, 0)
	$oOL_Rules.Remove($sOL_RuleName)
	If @error Then Return SetError(3, @error, 0)
	; Update the server
	$oOL_Rules.Save
	If @error Then Return SetError(4, @error, 0)
	Return 1

EndFunc   ;==>_OL_RuleDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_RuleExecute
; Description ...: Applies a rule as an one-off operation.
; Syntax.........: _OL_RuleExecute($oOL, $sOL_Store, $sOL_RuleName, $oOL_Folder[, $bOL_IncludeSubfolders = False[, $iOL_ExecuteOption = $olRuleExecuteAllMessages[, $bOL_ShowProgress = False]]])
; Parameters ....: $oOL                   - Outlook object returned by a preceding call to _OL_Open()
;                  $sOL_Store             - Name of the Store where the rule will be deleted from. "*" = your default store
;                  $sOL_RuleName          - Name of the rule to be deleted
;                  $oOL_Folder            - Object of the folder to which to apply the rule
;                  $bOL_IncludeSubfolders - Optional: Subfolders will be included if set to True (default = False)
;                  $iOL_ExecuteOption     - Optional: Specifies the type of messages in the specified folder or folders that a rule should be applied to (default = $olRuleExecuteAllMessages)
;                  $bOL_ShowProgress      - Optional: When set to True displays the progress dialog box when the rule is executed (default = False)
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - Rule doesn't exist in the specified store
;                  |2 - Error returned by method .GetRules. For more information check @extended
;                  |3 - Error executing the rule. For more information check @extended
;                  |4 - $oOL_Folder is not of type object
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_RuleExecute($oOL, $sOL_Store, $sOL_RuleName, $oOL_Folder, $bOL_IncludeSubfolders = False, $iOL_ExecuteOption = $olRuleExecuteAllMessages, $bOL_ShowProgress = False)

	If $sOL_Store = "*" Then $sOL_Store = $oOL.Session.DefaultStore.DisplayName
	Local $oOL_Rules = $oOL.Session.Stores.Item($sOL_Store).GetRules
	If @error Then Return SetError(2, @error, 0)
	If Not IsObj($oOL_Folder) Then Return SetError(4, 0, 0)
	Local $bOL_Found = False
	For $oOL_Rule In $oOL_Rules
		If $oOL_Rule.Name = $sOL_RuleName Then
			$bOL_Found = True
			ExitLoop
		EndIf
	Next
	If $bOL_Found = False Then Return SetError(1, 0, 0)
	$oOL_Rule.Execute($bOL_ShowProgress, $oOL_Folder, $bOL_IncludeSubfolders, $iOL_ExecuteOption)
	If @error Then Return SetError(3, @error, 0)
	Return 1

EndFunc   ;==>_OL_RuleExecute

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_RuleGet
; Description ...: Returns a list of rules for the specified store.
; Syntax.........: _OL_RuleGet($oOL[, sOL_Store = "*" [, $bOL_Enabled = True]])
; Parameters ....: $oOL         - Outlook object returned by a preceding call to _OL_Open()
;                  $sOL_Store   - Optional: Store to query for rules. Use "*" to denote the default store or specify the name of another store (default = "*")
;                  $bOL_Enabled - Optional: Only returns enabled rules if set to True (default = True)
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - Rule object
;                  |1 - Boolean value that determines if the rule is to be applied
;                  |2 - Integer indicating the order of execution of the rule in the rules collection
;                  |3 - Boolean value that indicates if the rule executes as a client-side rule
;                  |4 - String representing the name of the rule
;                  |5 - Constant from the OlRuleType enumeration indicating if the rule applies to messages being sent or received
;                  Failure - Returns "" and sets @error:
;                  |1 - No rules found for the specified store
;                  |2 - Error returned by method .GetRules. For more information check @extended
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_RuleGet($oOL, $sOL_Store = "*", $bOL_Enabled = True)

	If $sOL_Store = "*" Then $sOL_Store = $oOL.Session.DefaultStore.DisplayName
	Local $oOL_Rules = $oOL.Session.Stores.Item($sOL_Store).GetRules
	If @error Then Return SetError(2, @error, "")
	If $oOL_Rules.Count = 0 Then Return SetError(1, 0, "")
	Local $aOL_Rules[$oOL_Rules.Count + 1][6] = [[$oOL_Rules.Count, 6]]
	Local $iOL_Index = 1
	For $oOL_Rule In $oOL_Rules
		If $bOL_Enabled = False Or $oOL_Rule.Enabled = True Then
			$aOL_Rules[$iOL_Index][0] = $oOL_Rule
			$aOL_Rules[$iOL_Index][1] = $oOL_Rule.Enabled
			$aOL_Rules[$iOL_Index][2] = $oOL_Rule.ExecutionOrder
			$aOL_Rules[$iOL_Index][3] = $oOL_Rule.IsLocalRule
			$aOL_Rules[$iOL_Index][4] = $oOL_Rule.Name
			$aOL_Rules[$iOL_Index][5] = $oOL_Rule.RuleType
			$iOL_Index += 1
		EndIf
	Next
	ReDim $aOL_Rules[$iOL_Index][UBound($aOL_Rules, 2)]
	$aOL_Rules[0][0] = $iOL_Index - 1
	Return $aOL_Rules

EndFunc   ;==>_OL_RuleGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OL_StoreGet
; Description ...: Get information about the Stores in the current profile.
; Syntax.........: _OL_StoreGet($oOL)
; Parameters ....: $oOL - Outlook object returned by a preceding call to _OL_Open()
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - display name of the Store object
;                  |1 - Constant in the OlExchangeStoreType enumeration that indicates the type of an Exchange store
;                  |2 - Full file path for a Personal Folders File (.pst) or an Offline Folder File (.ost) store
;                  |3 - True if the store is a cached Exchange store
;                  |4 - True if the store is a store for an Outlook data file (Personal Folders File (.pst) or Offline Folder File (.ost))
;                  |5 - True if Instant Search is enabled and operational
;                  |6 - True if the Store is open
;                  |7 - String identifying the Store
;                  |8 - True if the OOF (Out Of Office) is set for this store
;                  Failure - Returns "" and sets @error:
;                  |1 - Function is only supported for Outlook 2007 and later
; Author ........: water
; Modified ......:
; Remarks .......: This function only works for Outlook 2007 and later.
;                  It always returns a valid filepath for PST files where function _OL_PSTGet might not (hebrew characters in filename etc.)
;                  +
;                  A store object represents a file on the local computer or a network drive that stores e-mail messages and other items.
;                  If you use an Exchange server, you can have a store on the server, in an Exchange Public folder, or on a local computer
;                  in a Personal Folders File (.pst) or Offline Folder File (.ost).
;                  For a POP3, IMAP, and HTTP e-mail server, a store is a .pst file.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_StoreGet($oOL)

	Local $aVersion = StringSplit($oOL.Version, '.')
	If Int($aVersion[1]) < 12 Then Return SetError(1, 0, "")
	Local $iOL_Index = 0
	Local $aOL_Store[$oOL.Session.Stores.Count + 1][9] = [[$oOL.Session.Stores.Count, 9]]
	For $oOL_Store In $oOL.Session.Stores
		$iOL_Index = $iOL_Index + 1
		$aOL_Store[$iOL_Index][0] = $oOL_Store.DisplayName
		$aOL_Store[$iOL_Index][1] = $oOL_Store.ExchangeStoreType
		$aOL_Store[$iOL_Index][2] = $oOL_Store.FilePath
		$aOL_Store[$iOL_Index][3] = $oOL_Store.IsCachedExchange
		$aOL_Store[$iOL_Index][4] = $oOL_Store.IsDataFileStore
		$aOL_Store[$iOL_Index][5] = $oOL_Store.IsInstantSearchEnabled
		$aOL_Store[$iOL_Index][6] = $oOL_Store.IsOpen
		$aOL_Store[$iOL_Index][7] = $oOL_Store.StoreId
		If $oOL_Store.ExchangeStoreType = $olExchangeMailbox Or $oOL_Store.ExchangeStoreType = $olPrimaryExchangeMailbox Then
			$aOL_Store[$iOL_Index][8] = $oOL_Store.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x661D000B")
		EndIf
	Next
	Return $aOL_Store

EndFunc   ;==>_OL_StoreGet

; #FUNCTION# ====================================================================================================================
; Name...........: _OL_VersionInfo
; Description ...: Returns an array of information about the OutlookEX.au3 UDF.
; Syntax.........: _OL_VersionInfo()
; Parameters ....: None
; Return values .: Success - one-dimensional one based array with the following information:
;                  |1 - Release Type (T=Test or V=Production)
;                  |2 - Major Version
;                  |3 - Minor Version
;                  |4 - Sub Version
;                  |5 - Release Date (YYYYMMDD)
;                  |6 - AutoIt version required
;                  |7 - List of authors separated by ","
;                  |8 - List of contributors separated by ","
; Author ........: water
; Modified.......:
; Remarks .......: Based on function _IE_VersionInfo written bei Dale Hohm
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_VersionInfo()

	Local $aOL_VersionInfo[9] = [8, "V", 0, 7, 1.1, "20120419", "3.1.1", "wooltown, water", _
			"progandy (CSV functions), Ultima, PsaltyDS (basis of the _OL_ArrayConcatenate function)"]
	Return $aOL_VersionInfo

EndFunc   ;==>_OL_VersionInfo

; #FUNCTION# ====================================================================================================================
; Name...........: _OL_Wrapper_CreateAppointment
; Description ...: Create an appointment (wrapper function).
; Syntax.........: _OL_Wrapper_CreateAppointment($oOL, $sSubject, $sStartDate[, $vEndDate = ""[, $sLocation = ""[, $bAllDayEvent = False[, $sBody = ""[, $sReminder = 15[, $sShowTimeAs = ""[, $iImportance = ""[, $iSensitivity = ""[, $iRecurrenceType = ""[, $sPatternStartDate = ""[, $sPatternEndDate = ""[, $iInterval = ""[, $iDayOfWeekMask = ""[, $iDay_MonthOfMonth_Year = ""[, $iInstance = ""]]]]]]]]]]]]]]])
; Parameters ....: $oOL                    - Outlook object returned by a preceding call to _OL_Open()
;                  $sSubject               - The Subject of the Appointment.
;                  $sStartDate             - Start date & time of the Appointment, format YYYY-MM-DD HH:MM - or what is set locally.
;                  $vEndDate               - Optional: End date & time of the Appointment, format YYYY-MM-DD HH:MM - or what is set locally OR
;                                            Number of minutes. If not set 30 minutes is used.
;                  $sLocation              - Optional: The location where the meeting is going to take place.
;                  $bAllDayEvent           - Optional: True or False(default), if set to True and the appointment is lasting for more than one day, end Date
;                                            must be one day higher than the actual end Date.
;                  $sBody                  - Optional: The Body of the Appointment.
;                  $sReminder              - Optional: Reminder in Minutes before start, 0 for no reminder
;                  $sShowTimeAs            - Optional: $olBusy=2 (default), $olFree=0, $olOutOfOffice=3, $olTentative=1
;                  $iImportance            - Optional: $olImportanceNormal=1 (default), $olImportanceHigh=2, $olImportanceLow=0
;                  $iSensitivity           - Optional: $olNormal=0, $olPersonal=1, $olPrivate=2, $olConfidential=3
;                  $iRecurrenceType        - Optional: $olRecursDaily=0, $olRecursWeekly=1, $olRecursMonthly=2, $olRecursMonthNth=3, $olRecursYearly=5, $olRecursYearNth=6
;                  $sPatternStartDate      - Optional: Start Date of the Reccurent Appointment, format YYYY-MM-DD - or what is set locally.
;                  $sPatternEndDate        - Optional: End Date of the Reccurent Appointment, format YYYY-MM-DD - or what is set locally.
;                  $iInterval              - Optional: Interval between the Reccurent Appointment
;                  $iDayOfWeekMask         - Optional: Add the values of the days the appointment shall occur. $olSunday=1, $olMonday=2, $olTuesday=4, $olWednesday=8, $olThursday=16, $olFriday=32, $olSaturday=64
;                  $iDay_MonthOfMonth_Year - Optional: DayOfMonth or MonthOfYear, Day of the month or month of the year on which the recurring appointment or task occurs
;                  $iInstance              - Optional: This property is only valid for recurrences of the $olRecursMonthNth and $olRecursYearNth type and allows the definition of a recurrence pattern that is only valid for the Nth occurrence, such as "the 2nd Sunday in March" pattern. The count is set numerically: 1 for the first, 2 for the second, and so on through 5 for the last. Values greater than 5 will generate errors when the pattern is saved.
; Return values .: Success - Object of the appointment
;                  Failure - Returns 0 and sets @error:
;                  |1    - $sStartDate is invalid
;                  |2    - $sOL_Body is missing
;                  |4    - $sTo, $sCc and $sBCc are missing
;                  |1xxx - Error returned by function _OL_FolderAccess
;                  |2xxx - Error returned by function _OL_ItemCreate
;                  |3xxx - Error returned by function _OL_ItemModify
;                  |4xxx - Error returned by function _OL_ItemRecurrenceSet
; Author ........: water
; Modified.......:
; Remarks .......: This is a wrapper function to simplify creating an appointment. If you have to set more properties etc. you have to do all steps yourself
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_Wrapper_CreateAppointment($oOL, $sSubject, $sStartDate, $vEndDate = "", $sLocation = "", $bAllDayEvent = False, $sBody = "", $sReminder = 15, $sShowTimeAs = "", $iImportance = "", $iSensitivity = "", $iRecurrenceType = "", $sPatternStartDate = "", $sPatternEndDate = "", $iInterval = "", $iDayOfWeekMask = "", $iDay_MonthOfMonth_Year = "", $iInstance = "")

	If Not _DateIsValid($sStartDate) Then Return SetError(1, 0, 0)
	Local $sEnd, $oOL_Item
	; Access the default calendar
	Local $aOL_Folder = _OL_FolderAccess($oOL, "", $olFolderCalendar)
	If @error Then Return SetError(@error + 1000, @extended, 0)
	; Create an appointment item in the default calendar and set properties
	If _DateIsValid($vEndDate) Then
		$sEnd = "End=" & $vEndDate
	Else
		$sEnd = "Duration=" & Number($vEndDate)
	EndIf
	$oOL_Item = _OL_ItemCreate($oOL, $olAppointmentItem, $aOL_Folder[1], "", "Subject=" & $sSubject, "Location=" & $sLocation, "AllDayEvent=" & $bAllDayEvent, _
			"Start=" & $sStartDate, "Body=" & $sBody, "Importance=" & $iImportance, "BusyStatus=" & $sShowTimeAs, $sEnd, "Sensitivity=" & $iSensitivity)
	If @error Then Return SetError(@error + 2000, @extended, 0)
	; Set reminder properties
	If $sReminder <> 0 Then
		$oOL_Item = _OL_ItemModify($oOL, $oOL_Item, Default, "ReminderSet=True", "ReminderMinutesBeforeStart=" & $sReminder)
		If @error Then Return SetError(@error + 3000, @extended, 0)
	Else
		$oOL_Item = _OL_ItemModify($oOL, $oOL_Item, Default, "ReminderSet=False")
		If @error Then Return SetError(@error + 3000, @extended, 0)
	EndIf
	; Set recurrence
	$iDayOfWeekMask = ""
	If $iRecurrenceType <> "" Then
		Local $sSDate, $sSTime, $sEDate, $sETime
		$sSDate = StringLeft($sPatternStartDate, 10)
		$sSTime = StringStripWS(StringMid($sPatternStartDate, 11), 3)
		$sEDate = StringLeft($sPatternEndDate, 10)
		$sETime = StringStripWS(StringMid($sPatternEndDate, 11), 3)
		If $iDayOfWeekMask <> "" Then $iDay_MonthOfMonth_Year = $iDayOfWeekMask
		$oOL_Item = _OL_ItemRecurrenceSet($oOL, $oOL_Item, Default, $sSDate, $sSTime, $sEDate, $sETime, $iRecurrenceType, $iDay_MonthOfMonth_Year, $iInterval, $iInstance)
		If @error Then Return SetError(@error + 4000, @extended, 0)
	EndIf
	Return $oOL_Item

EndFunc   ;==>_OL_Wrapper_CreateAppointment

; #FUNCTION# ====================================================================================================================
; Name...........: _OL_Wrapper_SendMail
; Description ...: Create and send a mail (wrapper function).
; Syntax.........: _OL_Wrapper_SendMail($oOL[, $sTo = ""[, $sCc= ""[, $sBCc = ""[, $sSubject = ""[, $sBody = ""[, $sAttachments = ""[, $iBodyFormat = $olFormatUnspecified[, $iImportance = $olImportanceNormal]]]]]]]])
; Parameters ....: $oOL          - Outlook object returned by a preceding call to _OL_Open()
;                  $sTo          - Optional: The recipiant(s), separated by ;
;                  $sCc          - Optional: The CC recipiant(s) of the mail, separated by ;
;                  $sBCc         - Optional: The BCC recipiant(s) of the mail, separated by ;
;                  $sSubject     - Optional: The Subject of the mail
;                  $sBody        - Optional: The Body of the mail
;                  $sAttachments - Optional: Attachments, separated by ;
;                  $iBodyFormat  - Optional: The Bodyformat of the mail as defined by the OlBodyFormat enumeration (default = $olFormatUnspecified)
;                  $iImportance  - Optional: The Importance of the mail as defined by the OlImportance enumeration (default = $olImportanceNormal)
; Return values .: Success - Object of the sent mail
;                  Failure - Returns 0 and sets @error:
;                  |1    - $iOL_BodyFormat is not a number
;                  |2    - $sOL_Body is missing
;                  |4    - $sTo, $sCc and $sBCc are missing
;                  |1xxx - Error returned by function _OL_FolderAccess
;                  |2xxx - Error returned by function _OL_ItemCreate
;                  |3xxx - Error returned by function _OL_ItemModify
;                  |4xxx - Error returned by function _OL_ItemRecipientAdd
;                  |5xxx - Error returned by function _OL_ItemAttachmentAdd
;                  |6xxx - Error returned by function _OL_ItemSend
; Author ........: water
; Modified.......:
; Remarks .......: This is a wrapper function to simplify sending an email. If you have to set more properties etc. you have to do all steps yourself
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _OL_Wrapper_SendMail($oOL, $sTo = "", $sCc = "", $sBCc = "", $sSubject = "", $sBody = "", $sAttachments = "", $iBodyFormat = $olFormatPlain, $iImportance = $olImportanceNormal)

	If Not IsInt($iBodyFormat) Then SetError(1, 0, 0)
	If StringStripWS($sBody, 3) = "" Then SetError(2, 0, 0)
	If StringStripWS($sTo, 3) = "" And StringStripWS($sCc, 3) = "" And StringStripWS($sBCc, 3) = "" Then SetError(3, 0, 0)
	; Access the default outbox folder
	Local $aOL_Folder = _OL_FolderAccess($oOL, "", $olFolderOutbox)
	If @error Then Return SetError(@error + 1000, @extended, 0)
	; Create a mail item in the default folder
	Local $oOL_Item = _OL_ItemCreate($oOL, $olMailItem, $aOL_Folder[1], "", "Subject=" & $sSubject, "BodyFormat=" & $iBodyFormat, "Importance=" & $iImportance)
	If @error Then Return SetError(@error + 2000, @extended, 0)
	; Set the body according to $iOL_BodyFormat
	If $iBodyFormat = $olFormatHTML Then
		_OL_ItemModify($oOL, $oOL_Item, Default, "HTMLBody=" & $sBody)
	Else
		_OL_ItemModify($oOL, $oOL_Item, Default, "Body=" & $sBody)
	EndIf
	If @error Then Return SetError(@error + 3000, @extended, 0)
	; Add recipients (to, cc and bcc)
	Local $aOL_Recipients
	If $sTo <> "" Then
		$aOL_Recipients = StringSplit($sTo, ";", 2)
		_OL_ItemRecipientAdd($oOL, $oOL_Item, Default, $olTo, $aOL_Recipients)
		If @error Then Return SetError(@error + 4000, @extended, 0)
	EndIf
	If $sCc <> "" Then
		$aOL_Recipients = StringSplit($sCc, ";", 2)
		_OL_ItemRecipientAdd($oOL, $oOL_Item, Default, $olCC, $aOL_Recipients)
		If @error Then Return SetError(@error + 4000, @extended, 0)
	EndIf
	If $sBCc <> "" Then
		$aOL_Recipients = StringSplit($sBCc, ";", 2)
		_OL_ItemRecipientAdd($oOL, $oOL_Item, Default, $olBCC, $aOL_Recipients)
		If @error Then Return SetError(@error + 4000, @extended, 0)
	EndIf
	; Add attachments
	If $sAttachments <> "" Then
		Local $aOL_Attachments = StringSplit($sAttachments, ";", 2)
		_OL_ItemAttachmentAdd($oOL, $oOL_Item, Default, $aOL_Attachments)
		If @error Then Return SetError(@error + 5000, @extended, 0)
	EndIf
	; Send mail
	_OL_ItemSend($oOL, $oOL_Item, Default)
	If @error Then Return SetError(@error + 6000, @extended, 0)
	Return $oOL_Item

EndFunc   ;==>_OL_Wrapper_SendMail

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: _OL_ArrayConcatenate
; Description ...: Concatenate 2D arrays.
; Syntax.........: _OL_ArrayConcatenate(ByRef $avArrayTarget, Const ByRef $avArraySource)
; Parameters ....: $avArrayTarget - The array to concatenate onto
;                  $avArraySource - The array to concatenate from - Must be 1D or 2D to match $avArrayTarget,
;                                   and if 2D, then Ubound($avArraySource, 2) <= Ubound($avArrayTarget, 2).
;                  $iOL_Flags     - Flags as defined in function call for _OL_ItemFind
; Return values .: Success - Index of last added item
;                  Failure - -1, sets @error to 1 and @extended per failure (see code below)
; Author ........: Ultima
; Modified.......: PsaltyDS - 1D/2D version, changed return value and @error/@extended to be consistent with __ArrayAdd()
;                  water - removed 1D array support, support for row 1 containing the number of rows/columns
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================
Func _OL_ArrayConcatenate(ByRef $avArrayTarget, Const ByRef $avArraySource, $iOL_Flags)

	If Not IsArray($avArrayTarget) Then Return SetError(1, 1, -1); $avArrayTarget is not an array
	If Not IsArray($avArraySource) Then Return SetError(1, 2, -1); $avArraySource is not an array
	Local $iUBoundTarget0 = UBound($avArrayTarget, 0), $iUBoundSource0 = UBound($avArraySource, 0)
	If $iUBoundTarget0 <> $iUBoundSource0 Then Return SetError(1, 3, -1); 1D/2D dimensionality did not match
	If $iUBoundTarget0 > 2 Then Return SetError(1, 4, -1); At least one array was 3D or more
	Local $iUBoundTarget1 = UBound($avArrayTarget, 1), $iUBoundSource1 = UBound($avArraySource, 1)
	Local $iNewSize = $iUBoundTarget1 + $iUBoundSource1 - 1
	Local $iUBoundTarget2 = UBound($avArrayTarget, 2), $iUBoundSource2 = UBound($avArraySource, 2)
	If $iUBoundSource2 > $iUBoundTarget2 Then Return SetError(1, 5, -1); 2D boundry of source too large for target
	ReDim $avArrayTarget[$iNewSize][$iUBoundTarget2]
	For $r = 1 To $iUBoundSource1 - 1
		For $c = 0 To $iUBoundSource2 - 1
			$avArrayTarget[$iUBoundTarget1 + $r - 1][$c] = $avArraySource[$r][$c]
		Next
	Next
	If BitAND($iOL_Flags, 2) <> 2 Then
		$avArrayTarget[0][0] = $iNewSize - 1
		If UBound($avArrayTarget) > 1 Then $avArrayTarget[0][1] = UBound($avArrayTarget, 2)
	EndIf
	Return $iNewSize - 1

EndFunc   ;==>_OL_ArrayConcatenate

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: _OL_CheckProperties
; Description ...: Check if specified properties exist for the item and have the correct case.
; Syntax.........: _OL_CheckProperties($oOL_Item, $aOL_Properties)
; Parameters ....: $oOL_Item       - Object of the item to check
;                  $aOL_Properties - Zero based array of property names. Format "propertyname=propertyvalue" is valid as well
;                  $iOL_Flag       - 0: Array $aOL_Properties is zero based, 1: Array $aOL_Properties is one based
; Return values .: Success - 1
;                  Failure - 0 and sets @error:
;                  |100 - Property is not valid for this type of item. @extended = number of property in error (zero based)
;                  |101 - Property has wrong case (e.g. "firstname" is wrong, "FirstName" is correct). @extended = number of property in error (zero based)
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================
Func _OL_CheckProperties($oOL_Item, $aOL_Properties, $iOL_Flag = 0)

	Local $sOL_ItemProperties = ",", $aOL_Temp, $iOL_Index, $iOL_End
	For $oOL_Property In $oOL_Item.ItemProperties
		$sOL_ItemProperties = $sOL_ItemProperties & $oOL_Property.name & ","
	Next
	If $iOL_Flag = 1 Then
		$iOL_Index = 1
		$iOL_End = $aOL_Properties[0]
	Else
		$iOL_Index = 0
		$iOL_End = UBound($aOL_Properties) - 1
	EndIf
	For $iOL_Index = $iOL_Index To $iOL_End
		$aOL_Temp = StringSplit($aOL_Properties[$iOL_Index], "=")
		If $aOL_Temp[1] <> "" And StringInStr($sOL_ItemProperties, "," & $aOL_Temp[1] & ",", 0) = 0 Then Return SetError(100, $iOL_Index, 0)
		If $aOL_Temp[1] <> "" And StringInStr($sOL_ItemProperties, "," & $aOL_Temp[1] & ",", 1) = 0 Then Return SetError(101, $iOL_Index, 0)
	Next
	Return 1

EndFunc   ;==>_OL_CheckProperties

; #INTERNAL_USE_ONLY#============================================================================================================
; Name ..........: _OL_COMError
; Description ...: Called if an ObjEvent error occurs.
; Syntax.........: _OL_COMError()
; Parameters ....: None
; Return values .: Sets @error to 999 and @error to the COM error number (decimal)
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _OL_COMError()

	Local $bHexNumber = Hex($oOL_Error.number, 8)
	Local $aOL_VersionInfo = _OL_VersionInfo()
	Local $sOL_Error = "COM Error Encountered in " & @ScriptName & @CRLF & _
			"OutlookEx UDF version = " & $aOL_VersionInfo[2] & "." & $aOL_VersionInfo[3] & "." & $aOL_VersionInfo[4] & @CRLF & _
			"@AutoItVersion = " & @AutoItVersion & @CRLF & _
			"@AutoItX64 = " & @AutoItX64 & @CRLF & _
			"@Compiled = " & @Compiled & @CRLF & _
			"@OSArch = " & @OSArch & @CRLF & _
			"@OSVersion = " & @OSVersion & @CRLF & _
			"Scriptline = " & $oOL_Error.scriptline & @CRLF & _
			"NumberHex = " & $bHexNumber & @CRLF & _
			"Number = " & $oOL_Error.number & @CRLF & _
			"WinDescription = " & StringStripWS($oOL_Error.WinDescription, 2) & @CRLF & _
			"Description = " & StringStripWS($oOL_Error.description, 2) & @CRLF & _
			"Source = " & $oOL_Error.Source & @CRLF & _
			"HelpFile = " & $oOL_Error.HelpFile & @CRLF & _
			"HelpContext = " & $oOL_Error.HelpContext & @CRLF & _
			"LastDllError = " & $oOL_Error.LastDllError
	If $iOL_Debug > 0 Then
		If $iOL_Debug = 1 Then ConsoleWrite($sOL_Error & @CRLF & "========================================================" & @CRLF)
		If $iOL_Debug = 2 Then MsgBox(64, "Outlook UDF - Debug Info", $sOL_Error)
		If $iOL_Debug = 3 Then FileWrite($sOL_DebugFile, @YEAR & "." & @MON & "." & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & " " & @CRLF & _
				"-------------------" & @CRLF & $sOL_Error & @CRLF & "========================================================" & @CRLF)
	EndIf
	Return SetError(999, $oOL_Error.number, 0)

EndFunc   ;==>_OL_COMError

; #INTERNAL_USE_ONLY#============================================================================================================
; Name ..........: _OL_TestEnvironmentCreate
; Description ...: Delete and recreate the OutlookEX UDF test environment.
; Syntax.........: _OL_TestEnvironmentCreate($oOL[, $vOL_DontAsk = ""[, $vOL_DontDelete= ""]])
; Parameters ....: $oOL            - Outlook object returned by a preceding call to _OL_Open()
;                  $vOL_DontAsk    - Optional: Return error if folder already exists? 1 = no, 4 = yes, "" = read value from ini file
;                  $vOL_DontDelete - Optional: Delete test environment if it already exists? 1 = no, 4 = yes, "" = read value from ini file
; Return values .: Success - 1
;                  Failure - 0 and sets @error:
;                  |1 - Folder *\Outlook-UDF-Test already exists and user denied to delete/recreate the testenvironment
;                  |2 - Error deleting folder. @extended is @error as returned by _OL_FolderDelete
;                  |3xx - Error creating source folder. @extended is set to @error as returned by _OL_FolderCreate
;                  |4xx - Error creating target folder. @extended is set to @error as returned by _OL_FolderCreate
;                  |5xx - Error creating item in source folder. @extended is set to @error as returned by _OL_<itemtype>Create
;                  |6xx - Error creating item in target folder. @extended is set to @error as returned by _OL_<itemtype>Create
; Author ........: water
; Modified ......:
; Remarks .......: The test environment consists of the following items:
;                    * Folder Outlook-UDF-Test, subfolders plus items
;                    * Group "Outlook-UDF-Test" in the Outlook bar
;                    * Shortcut int the "Outlook-UDF-Test" group in the Outlook bar
;                    * Category "Outlook-UDF-Test" in the Outlook bar
;                    * Mail signature "Outlook-UDF-Test"
;                    * PST "Outlook-UDF-Test" in "C:\temp\Outlook-UDF-Test.pst"
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _OL_TestEnvironmentCreate($oOL, $vOL_DontAsk = "", $vOL_DontDelete = "")

	Local $vOL_Result, $sOL_CurrentUser = $oOL.GetNameSpace("MAPI").CurrentUser.Name
	;---------------------------------
	; Delete existing folder structure
	;---------------------------------
	If _OL_FolderExists($oOL, "*\Outlook-UDF-Test") Then
		If StringStripWS($vOL_DontAsk, 3) = "" Then $vOL_DontAsk = IniRead("_OL_TestEnvironment.ini", "Configuration", "DontAsk", "4") ; checked = 1, unchecked = 4
		If StringStripWS($vOL_DontDelete, 3) = "" Then $vOL_DontDelete = IniRead("_OL_TestEnvironment.ini", "Configuration", "DontDelete", "4") ; checked = 1, unchecked = 4
		If $vOL_DontDelete = 4 Then
			If $vOL_DontAsk = 4 Then
				Local $iResult = MsgBox(35, "OutlookEX UDF - Create Test Environment", "Testenvironment already exists. Should it be deleted and recreated?")
				If $iResult = 2 Then Return SetError(1, 0, 0)
				If $iResult = 7 Then Return 1
			EndIf
			If $vOL_DontAsk = 1 Or $iResult = 6 Then _OL_FolderDelete($oOL, "*\Outlook-UDF-Test")
		Else
			Return 1
		EndIf
		If @error Then Return SetError(2, @error, 0)
	EndIf
	;------------------------------------------------------------
	; Create Source Folder plus one subfolder for every item type
	;------------------------------------------------------------
	Local $oSourceFolderCalendar = _OL_FolderCreate($oOL, "Outlook-UDF-Test\SourceFolder\Calendar", $olFolderCalendar)
	If @error Then Return SetError(300, @error, 0)
	Local $oSourceFolderContact = _OL_FolderCreate($oOL, "Contacts", $olFolderContacts, "*\Outlook-UDF-Test\SourceFolder")
	If @error Then Return SetError(301, @error, 0)
	; Mark the folder as address book and change the name
	$oSourceFolderContact.ShowAsOutlookAB = True
	$oSourceFolderContact.AddressBookName = "Outlook-UDF-Test"
	If @error Then Return SetError(302, @error, 0)
	Local $oSourceFolderMail = _OL_FolderCreate($oOL, "Mail", $olFolderInbox, "*\Outlook-UDF-Test\SourceFolder")
	If @error Then Return SetError(303, @error, 0)
	Local $oSourceFolderNotes = _OL_FolderCreate($oOL, "Notes", $olFolderNotes, "*\Outlook-UDF-Test\SourceFolder")
	If @error Then Return SetError(304, @error, 0)
	Local $oSourceFolderTasks = _OL_FolderCreate($oOL, "Tasks", $olFolderTasks, "*\Outlook-UDF-Test\SourceFolder")
	If @error Then Return SetError(305, @error, 0)
	;-----------------------------------
	; Create test items in Source Folder
	;-----------------------------------
	; Appointment
	Local $sOL_StartTime = StringMid(_DateAdd("n", -10, _NowCalc()), 12, 5)
	Local $sOL_EndTime = StringMid(_DateAdd("h", 3, _NowCalc()), 12, 5)
	Local $oOL_Item = _OL_ItemCreate($oOL, $olAppointmentItem, $oSourceFolderCalendar, "", "Subject=TestAppointment", _
			"Start=" & _NowCalcDate() & " " & $sOL_StartTime, "End=" & _NowCalcDate() & " " & $sOL_EndTime, _
			"ReminderMinutesBeforeStart=15", "ReminderSet=True", "Location=Building A, Room 10")
	If @error Then Return SetError(500, @error, 0)
	; Appointment: Add optional recipient
	_OL_ItemRecipientAdd($oOL, $oOL_Item, Default, $olOptional, $sOL_CurrentUser)
	If @error Then Return SetError(501, @error, 0)
	; Appointment: Add recurrence: Daily with defined start and end date/time
	_OL_ItemRecurrenceSet($oOL, $oOL_Item, Default, _NowCalcDate(), $sOL_StartTime, _DateAdd("D", 14, _NowCalcDate()), $sOL_EndTime, $olRecursDaily, "", "", "")
	If @error Then Return SetError(502, @error, 0)
	; Define exception
	Local $sOL_Temp = _DateAdd("D", 1, _NowCalcDate())
	_OL_ItemRecurrenceExceptionSet($oOL, $oOL_Item, Default, $sOL_Temp & " " & $sOL_StartTime & ":00", $sOL_Temp & " 08:00:00", $sOL_Temp & " 14:00:00", "TestException", "ExceptionBody")
	If @error Then Return SetError(503, @error, 0)
	; Appointment: Create a conflict
	$oOL_Item = _OL_ItemCreate($oOL, $olAppointmentItem, $oSourceFolderCalendar, "", "Subject=TestAppointment-Conflict", _
			"Start=" & _NowCalcDate() & " " & $sOL_StartTime, "End=" & _NowCalcDate() & " " & $sOL_EndTime)
	If @error Then Return SetError(504, @error, 0)
	; Contact
	_OL_ItemCreate($oOL, $olContactItem, $oSourceFolderContact, "", "FirstName=TestFirstName", "LastName=TestLastName")
	If @error Then Return SetError(505, @error, 0)
	; Distribution List + Member
	$vOL_Result = _OL_ItemCreate($oOL, $olDistributionListItem, $oSourceFolderContact, "", "Subject=TestDistributionList", "Importance=" & $olImportanceHigh)
	If @error Then Return SetError(506, @error, 0)
	_OL_DistListMemberAdd($oOL, $vOL_Result, Default, $sOL_CurrentUser)
	If @error Then Return SetError(507, @error, 0)
	; Mail plus attachment
	$vOL_Result = _OL_ItemCreate($oOL, $olMailItem, $oSourceFolderMail, "", "Subject=TestMail", "BodyFormat=" & $olFormatHTML, "HTMLBody=Bodytext in <b>bold</b>.", "To=" & $sOL_CurrentUser)
	If @error Then Return SetError(508, @error, 0)
	_OL_ItemAttachmentAdd($oOL, $vOL_Result, Default, @ScriptDir & "\The_Outlook.jpg")
	If @error Then Return SetError(509, @error, 0)
	; Note
	_OL_ItemCreate($oOL, $olNoteItem, $oSourceFolderNotes, "", "Body=TestNote", "Width=350")
	If @error Then Return SetError(510, @error, 0)
	; Task
	_OL_ItemCreate($oOL, $olTaskItem, $oSourceFolderTasks, "", "Subject=TestSubject", "StartDate=" & _NowDate())
	If @error Then Return SetError(511, @error, 0)
	;------------------------------------------------------------
	; Create Target Folder plus one subfolder for every item type
	;------------------------------------------------------------
	_OL_FolderCreate($oOL, "TargetFolder\Calendar", $olFolderCalendar, "*\Outlook-UDF-Test")
	If @error Then Return SetError(400, @error, 0)
	_OL_FolderCreate($oOL, "Contacts", $olFolderContacts, "*\Outlook-UDF-Test\TargetFolder")
	If @error Then Return SetError(401, @error, 0)
	_OL_FolderCreate($oOL, "Mail", $olFolderInbox, "*\Outlook-UDF-Test\TargetFolder")
	If @error Then Return SetError(402, @error, 0)
	_OL_FolderCreate($oOL, "Notes", $olFolderNotes, "*\Outlook-UDF-Test\TargetFolder")
	If @error Then Return SetError(403, @error, 0)
	_OL_FolderCreate($oOL, "Tasks", $olFolderTasks, "*\Outlook-UDF-Test\TargetFolder")
	If @error Then Return SetError(404, @error, 0)
	Return 1

EndFunc   ;==>_OL_TestEnvironmentCreate

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: _ParseCSV
; Description ...: Reads a CSV-file
; Syntax.........: _ParseCSV($sFile, $sDelimiters=',', $sQuote='"', $iFormat=0)
; Parameters ....: $sFile       - File to read or string to parse
;                  $sDelimiters - [optional] Fieldseparators of CSV, multiple are allowed (default = ,;)
;                  $sQuote      - [optional] Character to quote strings (default = ")
;                  $iFormat     - [optional] Encoding of the file (default = 0):
;                  |-1     - No file, plain data given
;                  |0 or 1 - automatic (ASCII)
;                  |2      - Unicode UTF16 Little Endian reading
;                  |3      - Unicode UTF16 Big Endian reading
;                  |4 or 5 - Unicode UTF8 reading
; Return values .: Success - 2D-Array with CSV data (0-based)
;                  Failure - 0, sets @error to:
;                  |1 - could not open file
;                  |2 - error on parsing data
;                  |3 - wrong format chosen
; Author ........: ProgAndy
; Modified.......:
; Remarks .......:
; Related .......: _WriteCSV
; Link ..........: http://www.autoitscript.com/forum/topic/114406-csv-file-to-multidimensional-array
; Example .......:
; ===============================================================================================================================
Func _ParseCSV($sFile, $sDelimiters = ',;', $sQuote = '"', $iFormat = 0)

	Local Static $aEncoding[6] = [0, 0, 32, 64, 128, 256]
	If $iFormat < -1 Or $iFormat > 6 Then
		Return SetError(3, 0, 0)
	ElseIf $iFormat > -1 Then
		Local $hFile = FileOpen($sFile, $aEncoding[$iFormat])
		If @error Then Return SetError(1, @error, 0)
		$sFile = FileRead($hFile)
		FileClose($hFile)
	EndIf
	If $sDelimiters = "" Or IsKeyword($sDelimiters) Then $sDelimiters = ',;'
	If $sQuote = "" Or IsKeyword($sQuote) Then $sQuote = '"'
	$sQuote = StringLeft($sQuote, 1)
	Local $srDelimiters = StringRegExpReplace($sDelimiters, '[\\\^\-\[\]]', '\\\0')
	Local $srQuote = StringRegExpReplace($sQuote, '[\\\^\-\[\]]', '\\\0')
	Local $sPattern = StringReplace(StringReplace('(?m)(?:^|[,])\h*(["](?:[^"]|["]{2})*["]|[^,\r\n]*)(\v+)?', ',', $srDelimiters, 0, 1), '"', $srQuote, 0, 1)
	Local $aREgex = StringRegExp($sFile, $sPattern, 3)
	If @error Then Return SetError(2, @error, 0)
	$sFile = '' ; save memory
	Local $iBound = UBound($aREgex), $iIndex = 0, $iSubBound = 1, $iSub = 0
	Local $aResult[$iBound][$iSubBound]
	For $i = 0 To $iBound - 1
		Select
			Case StringLen($aREgex[$i]) < 3 And StringInStr(@CRLF, $aREgex[$i])
				$iIndex += 1
				$iSub = 0
				ContinueLoop
			Case StringLeft(StringStripWS($aREgex[$i], 1), 1) = $sQuote
				$aREgex[$i] = StringStripWS($aREgex[$i], 3)
				$aResult[$iIndex][$iSub] = StringReplace(StringMid($aREgex[$i], 2, StringLen($aREgex[$i]) - 2), $sQuote & $sQuote, $sQuote, 0, 1)
			Case Else
				$aResult[$iIndex][$iSub] = $aREgex[$i]
		EndSelect
		$aREgex[$i] = 0 ; save memory
		$iSub += 1
		If $iSub = $iSubBound Then
			$iSubBound += 1
			ReDim $aResult[$iBound][$iSubBound]
		EndIf
	Next
	If $iIndex = 0 Then $iIndex = 1
	ReDim $aResult[$iIndex][$iSubBound]
	Return $aResult

EndFunc   ;==>_ParseCSV

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: _WriteCSV
; Description ...: Writes a CSV-file (appends to an existing file)
; Syntax.........: _WriteCSV($sFile, Const ByRef $aData, $sDelimiter, $sQuote, $iFormat=0)
; Parameters ....: $sFile      - Destination file
;                  $aData      - [Const ByRef] 1-based 2D-Array with data
;                  $sDelimiter - [optional] Fieldseparator (default = ,)
;                  $sQuote     - [optional] Quote character (default = ")
;                  $iFormat    - [optional] character encoding of file (default = 0)
;                  |0 or 1 - ASCII writing
;                  |2      - Unicode UTF16 Little Endian writing (with BOM)
;                  |3      - Unicode UTF16 Big Endian writing (with BOM)
;                  |4      - Unicode UTF8 writing (with BOM)
;                  |5      - Unicode UTF8 writing (without BOM)
; Return values .: Success - Number of records written
;                  Failure - 0, sets @error to:
;                  |1 - No valid 2D-Array
;                  |2 - Could not open file
; Author ........: ProgAndy
; Modified.......: water
; Remarks .......:
; Related .......: _ParseCSV
; Link ..........: http://www.autoitscript.com/forum/topic/114406-csv-file-to-multidimensional-array
; Example .......:
; ===============================================================================================================================
Func _WriteCSV($sFile, Const ByRef $aData, $sDelimiter = ',', $sQuote = '"', $iFormat = 0)

	Local Static $aEncoding[6] = [9, 9, 41, 73, 137, 265] ; Mode + 1 (append to end of file) + 8 (create dir structure)
	If $sDelimiter = "" Or IsKeyword($sDelimiter) Then $sDelimiter = ','
	If $sQuote = "" Or IsKeyword($sQuote) Then $sQuote = '"'
	Local $iBound = UBound($aData, 1), $iSubBound = UBound($aData, 2)
	If Not $iSubBound Then Return SetError(1, 0, 0)
	Local $hFile = FileOpen($sFile, $aEncoding[$iFormat])
	If @error Then Return SetError(2, @error, 0)
	For $i = 1 To $iBound - 1
		For $j = 0 To $iSubBound - 1
			FileWrite($hFile, $sQuote & StringReplace($aData[$i][$j], $sQuote, $sQuote & $sQuote, 0, 1) & $sQuote)
			If $j < $iSubBound - 1 Then FileWrite($hFile, $sDelimiter)
		Next
		FileWrite($hFile, @CRLF)
	Next
	FileClose($hFile)
	Return $iBound - 1

EndFunc   ;==>_WriteCSV

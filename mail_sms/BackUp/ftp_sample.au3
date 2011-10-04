;
func _ftp_upload_name_text( $text_2_upload, $name_2_upload )
	local $ftp_upload
if $ftp_upload=1 then

local $ftp_server = '202.133.232.82'
local $ftp_username = 'ivan'
local $pass = '9ps5678'
local $ftp_upload_text_file
local $ftp_upload_namelist_file
local $aFile

$Open = _FTP_Open('MyFTP Control')
$Conn = _FTP_Connect($Open, $ftp_server, $ftp_username, $pass,1,6021)
; ...
if _FTP_DirsetCurrent($Conn, "/upload_sms/"&$user_name)=0 then _FTP_DirCreate($Conn,"/upload_sms/"&$user_name )
_FTP_DirSetCurrent( $Conn, "/" )
;local $h_Handle
;$aFile = _FTP_FindFileFirst($Conn, "/"&$user_name &"/"&$astronomy, $h_Handle)

;$astronomy_filesize = _FTP_FileGetSize($Conn, $aFile[10])
;ConsoleWrite('$Filename = ' & $aFile[10] & ' size = ' & $FileSize & '  -> Error code: ' & @error & ' extended: ' & @extended  & @crlf)

;_FTP_DirSetCurrent($Conn,"/upload" )
;$ftp_upload_text_file="SMS_text.txt"
$ftp_upload_text_file=$text_2_upload
local  $szDrive, $szDir, $szFName, $szExt
;$TestPath = _PathSplit(@ScriptFullPath, $szDrive, $szDir, $szFName, $szExt)
local $path_split=_PathSplit($ftp_upload_text_file , $szDrive, $szDir, $szFName, $szExt)
;_ArrayDisplay($path_split)
;MsgBox(0,"File Path",  $ftp_upload_text_file &"  --  "& $name_2_upload & " >>>  "&  $path_split[3]&$path_split[4] )
_FTP_FilePut( $Conn,  @UserProfileDir & "\" & $user_name &".sms", "/upload_sms/"&$user_name&".sms", $FTP_TRANSFER_TYPE_BINARY )

_FTP_FilePut( $Conn, @UserProfileDir & "\" & $astronomy, "/upload_sms/"&$user_name &"/"&$astronomy, $FTP_TRANSFER_TYPE_BINARY )

_FTP_FilePut( $Conn, $ftp_upload_text_file, "/upload_sms/"&$user_name &"/"&$path_split[3]&$path_split[4], $FTP_TRANSFER_TYPE_BINARY )
Local $h_Handle
$aFile = _FTP_FindFileFirst($Conn, "/upload_sms/"&$user_name &"/"&$path_split[3]&$path_split[4], $h_Handle)
;_ArrayDisplay($aFile)

ConsoleWrite('$Filename = ' & $aFile[10] & ' FileSizeLo = ' & $aFile[9] & '  -> Error code: ' & @error & @crlf)
$FindClose = _FTP_FindFileClose($h_Handle)


;;  name list upload
;$ftp_upload_namelist_file="SMS_name_list.csv"
$ftp_upload_namelist_file= $name_2_upload

$path_split=_PathSplit($ftp_upload_namelist_file , $szDrive, $szDir, $szFName, $szExt)
_FTP_FilePut( $Conn, $ftp_upload_namelist_file, "/upload_sms/"&$user_name &"/"&$path_split[3]&$path_split[4], $FTP_TRANSFER_TYPE_BINARY )
Local $h_Handle
$aFile = _FTP_FindFileFirst($Conn, "/upload_sms/"&$user_name &"/"&$path_split[3]&$path_split[4], $h_Handle)
ConsoleWrite('$Filename = ' & $aFile[10] & ' FileSizeLo = ' & $aFile[9] & '  -> Error code: ' & @error & @crlf)
$FindClose = _FTP_FindFileClose($h_Handle)

;;
$Ftpc = _FTP_Close($Open)


EndIf
EndFunc#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.2.12.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here


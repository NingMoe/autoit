#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <FTPEx.au3>

$dir_2_ftp="E:\autoit.git\kitravel\hotel_data_files\221"
$h_id=221

_kiftp_upload( $dir_2_ftp, $h_id )

func _kiftp_upload($dir_2_ftp, $h_id)

	Local $server = '202.168.197.252'
	Local $username = 'bryant_test'
	Local $pass = '9ps5678'

	Local $Open = _FTP_Open('MyFTP Control')
	Local $Conn = _FTP_Connect($Open, $server, $username, $pass,0,6002)
	_FTP_DirCreate ($Conn, "/" & $h_id )
	_FTP_DirPutContents( $Conn,$dir_2_ftp,"/"& $h_id,0 )
	; ...
	Local $Ftpc = _FTP_Close($Open)



EndFunc
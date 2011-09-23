#include <array.au3>
#include <File.au3>
#include <Date.au3>
#include <CompInfo_win7.au3>
#include <string.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GUIConstants.au3>


;$Pressed= MsgBox(4,"", "Yes or No")

;ConsoleWrite ("Current $Pressed value" & $Pressed & @CRLF)
;if  $Pressed=7 then 
	
;	Exit
;EndIf

dim $basicinfo
dim $Pressed
while 1
	$basicinfo=_BasicInfoGUI()
	;MsgBox(4,"Basic Info","電子信箱 : "& $Email_Address  & @CRLF& "行動電話 : " &$Inform_Mobile &@CRLF& "初始密碼 : " &$Init_Password)
	$Pressed=MsgBox(4,"Basic Info","電子信箱 : "& $basicinfo[0]  & @CRLF& "行動電話 : " & $basicinfo[1] &@CRLF& "初始密碼 : " & $basicinfo[2])
	ConsoleWrite ("Current $Pressed value" & $Pressed & @CRLF)
	if  $Pressed=6 then ExitLoop
WEnd

Exit

Func _BasicInfoGUI()

	Local $email, $mobile, $btn, $msg, $btn_n, $aEnc_info, $rc,  $password
	local $Email_Address , $Inform_Mobile , $Init_Password
	local $return_info[3]
	
	GUICreate("輸入基本資料", 320, 220, @DesktopWidth / 3 - 320, @DesktopHeight / 3 - 240, -1, 0x00000018); WS_EX_ACCEPTFILES
	GUICtrlCreateLabel("1.Your Email Address. Ex. abc@abc.com", 10, 10, 300, 40)
	$email = GUICtrlCreateInput("", 10, 25, 300, 30)
	GUICtrlSetState(-1, $GUI_FOCUS)

	GUICtrlCreateLabel("2.Your Mobile Phone to inform. Ex. 0928123456", 10, 75, 300, 40)
	$mobile = GUICtrlCreateInput("", 10, 90, 300, 30)
	GUICtrlSetState(-1, $GUI_NODROPACCEPTED)
	
	GUICtrlCreateLabel("3.Init Password ", 10, 140, 300, 40)
	$password = GUICtrlCreateInput("", 10, 155, 300, 30)
	GUICtrlSetState(-1, $GUI_NODROPACCEPTED)
	;GUICtrlCreateInput("", 10, 35, 300, 20) 	; will not accept drag&drop files
	$btn = GUICtrlCreateButton("OK", 90, 190, 60, 20, 0x0001) ; Default button
	$btn_n = GUICtrlCreateButton("Exit", 160, 190, 60, 20)
	GUISetState()
	;$msg = 0
	While $msg <> $GUI_EVENT_CLOSE

		$msg = GUIGetMsg()
		Select
			Case $msg = $btn
				If Not (GUICtrlRead($email) = "" And GUICtrlRead($mobile) = "" and  GUICtrlRead($password) = "")Then
					;MsgBox(4096, "drag drop file", GUICtrlRead($email) & "  " & GUICtrlRead($mobile))
					$Email_Address= GUICtrlRead($email)
					$Inform_Mobile = GUICtrlRead($mobile)
					$Init_Password = GUICtrlRead($password)
					
					if not ( StringInStr ($Email_Address , "@") and StringInStr ($Email_Address , ".") )then 
							MsgBox(0,"錯誤","E-mail 輸入錯誤，請重新執行程式")
							Exit
					EndIf
					
					if not ( StringLeft($Inform_Mobile ,2)= 09  and StringLen ($Inform_Mobile)=10 ) then 
							MsgBox(0,"錯誤","行動電話輸入錯誤，請重新執行程式")
							Exit
					EndIf
				Else
					MsgBox(0,"錯誤","輸入錯誤，請重新執行程式")
					Exit
				EndIf
				ExitLoop
			Case $msg = $btn_n
				Exit
		EndSelect
		
	WEnd

GUIDelete();

	$return_info[0]=$Email_Address
	$return_info[1]=$Inform_Mobile
	$return_info[2]=$Init_Password
return (  $return_info  )
EndFunc   ;==>_BasicInfoGUI


#Include <String.au3>
dim $text="�@�G�T�|�����C�K�E�s�@�G�T�|�����C�K�E�s�@�G�T�|�����C�K�E�s�@�G�T�|�����C�K�E�s�@�G�T�|�����C�K�E�s�@�G�T�|�����C�K�E�s�@�G�T�|�����C�K�E�s�@�G�T�|�����C�K�E�s�@�G�T�|�����C�K�E�s�@�G�T�|�����C�K�E�s"
dim $text="������Search Engine Land��Ĭ�Q���ܡAGoogle���Ѫ��s�����̷s���e�A�i��b�o�ͦa�_����o�ƥ�ɳ̦��ΡA�]�����ͧƱ椣�α��y�h���ӷ��N��x���̷s�����C"
dim $ENC_text= _StringToHex($text)
dim $DEC_text= _HexToString($ENC_text)

MsgBox(0,"Encript", "ENC text :"& $ENC_text & @CRLF & "DEC text :" & $DEC_text )
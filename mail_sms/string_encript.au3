
#Include <String.au3>
dim $text="一二三四五六七八九零一二三四五六七八九零一二三四五六七八九零一二三四五六七八九零一二三四五六七八九零一二三四五六七八九零一二三四五六七八九零一二三四五六七八九零一二三四五六七八九零一二三四五六七八九零"
dim $text="部落格Search Engine Land的蘇利文表示，Google提供社群網站最新內容，可能在發生地震等突發事件時最有用，因為網友希望不用掃描多重來源就能掌握最新消息。"
dim $ENC_text= _StringToHex($text)
dim $DEC_text= _HexToString($ENC_text)

MsgBox(0,"Encript", "ENC text :"& $ENC_text & @CRLF & "DEC text :" & $DEC_text )
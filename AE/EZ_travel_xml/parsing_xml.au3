;
; For Parsing EZ-travel XML
;
;
#include <array.au3>
#include <file.au3>
#include <Date.au3>
#Include <_XMLDomWrapper.au3>
#include <bryant_mysql.au3>


dim $sec=@SEC
dim $min=@MIN
Dim $hour=@HOUR
Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR


;===============================================================================
Global $sXmlFile
Global $query_date='0'
;===============================================================================
dim $test_mode= _TEST_MODE()
;
;dim $DB_IP_bry= "127.0.0.1"
;dim $DB_IP_bry= "192.168.3.45"
dim $DB_IP_bry= "140.119.116.137"

;dim $DB_NAME_bry ="test7"
dim $DB_NAME_bry ="test"

;dim $DB_ACCOUNT_bry ="amex"
dim $DB_ACCOUNT_bry ="ivan"
;dim $DB_PASS_bry ="amextravel"
dim $DB_PASS_bry ="9ps5678"

dim $STATEMENT_bry 

dim $query_insert= "INSERT INTO hs_product (product_date,Tran_No,hs_class,DepTime,ArrTime,depCity,desCity) VALUES('20100101','399','N','0000','2229','110','115')"

dim $query_select = "SELECT * FROM hs_product where product_date<'20100201'"
dim $query_last ="SELECT * FROM `test7`.`hs_product` ORDER BY product_nbr DESC limit 1"

dim $query_update ="UPDATE `test7`.`hs_product` set product_date='20100105' where product_nbr='0006970685'"

dim $query_del = "DELETE from `test7`.`hs_product` where product_date<20100201"


;$STATEMENT_bry = $query_last
;$STATEMENT_bry = $query_select

;$STATEMENT_bry = $query_insert

;$STATEMENT_bry = $query_update

;$STATEMENT_bry = $query_del


;$my=_bryant_mysql($DB_IP_bry,$DB_NAME_bry,$DB_ACCOUNT_bry,$DB_PASS_bry, $STATEMENT_bry)
;if IsArray ($my) then 
;	_ArrayDisplay($my)
;Else
;	MsgBox(0,"Affect on DB ", "Affect in DB '" & $DB_NAME_bry & "' by " & $my & " rows")

;EndIf

; This is to process XML file from EZ-Travel
;
;$sXmlFile = FileOpenDialog("Open XML", @ScriptDir, "XML (*.XML)", 1)
;If @error Then
;    MsgBox(4096, "File Open", "No file chosen , Exiting")
;    Exit
;EndIf

;InetGet("http://www.eztravel.com.tw/ezec/pkghsr/AeHsrInfo?queryDate=20101222", @ScriptDir&"\20101222.xml")
;$process_date=InputBox("今天是 : " &$year &"/"&$month &"/"& $day ,"需要處理的高鐵行程日期"&@CRLF& "範例：輸入格式為 : " &$year&$month&$day)
do  
	$process_date=InputBox("今天是 : " &$year &"/"&$month &"/"& $day ,"需要處理的高鐵行程日期"&@CRLF& "範例：輸入格式為 : " &$year&$month&$day & @CRLF &"輸入 N 可以離開")
	if $process_date="n" or $process_date="N" or @error =1 then exit
until $process_date>=$year&$month&$day 

if Not FileExists (@ScriptDir&"\ez-travel-"& $process_date &".xml") then 
	InetGet("http://www.eztravel.com.tw/ezec/pkghsr/AeHsrInfo?queryDate="&$process_date, @ScriptDir&"\ez-travel-"& $process_date &".xml")
EndIf

	$sXmlFile =@ScriptDir&"\ez-travel-"& $process_date &".xml"

$query_select="SELECT * FROM hs_product where product_date="&$process_date
$STATEMENT_bry = $query_select
$my=_bryant_mysql($DB_IP_bry,$DB_NAME_bry,$DB_ACCOUNT_bry,$DB_PASS_bry, $STATEMENT_bry)
if IsArray ($my) then 
	;_ArrayDisplay($my)
	;MsgBox(0,"hs_product in DB", "資料庫中尚有當日其他資料"& UBound ($my))
	$ans=InputBox("資料庫中 hs_product" ," 資料庫 "&$DB_IP_bry &":/"& $DB_NAME_bry &" 中還有 "& $process_date &" 其他資料: " & UBound ($my) &" 筆" &@CRLF & "輸入 N 可以離開")
	if $ans="n" or $ans="N" or @error =1 then exit
Else
	;MsgBox(0,"Affect on DB ", "Affect in DB '" & $DB_NAME_bry & "' by " & $my & " rows")
	if stringinstr($my,"MySQL_ERROR") >0  then MsgBox(0, "資料庫",  "資料庫 "&$DB_IP_bry &" 連線錯誤")
EndIf





dim $dest_dir=@ScriptDir
;if FileExists($dest_dir&"\test_data.xml") then 
;	$sXmlFile=$dest_dir&"\test_data.xml"
    $oXml = _XMLFileOpen($sXmlFile)
;EndIf


	;dim $getpath=_XMLGetPath('//ezProdDetail/*')
	;MsgBox(0,"paht",$getpath)
$ezXPathRoot = '//ezProdDetail'
$root_node= _XMLGetValue($ezXPathRoot & '/*')  
	;_ArrayDisplay($root_node)
if  IsArray ($root_node) and $root_node[0]=1 then
		;MsgBox(0,"Query Date", $root_node[1] ,5)
		$query_date= $root_node[1]
EndIf

$ezXPath = '//ezProdDetail/prod'
;Dim $aAttrName[4], $aAttrValue[4], $node

; 計算出這個 XML文件中有多少個 prod
$prod_node_count=_XMLGetNodeCount($ezXPath)  
;MsgBox(0, "Node Count", $prod_node_count)
   
$total_schedule_node_count= _XMLGetNodeCount($ezXPath & "/scheduleInfo/*") 
;MsgBox(0, "total schedule node count", $total_schedule_node_count)   

dim $hs_product[1+$total_schedule_node_count][7]
	$hs_product[0][0]= $total_schedule_node_count

dim $h=1 ; for hs_product counter


if $prod_node_count > 0 Then
	;Local $rtn_array[$prod_node_count]
	;$rtn_array=_XMLSelectNodes ($ezXPath &'/prodNo')
	;$rtn_array=_XMLSelectNodes ($ezXPath &'')
	;$rtn_array= _XMLGetValue($ezXPath & "/depCityCode")
	;$rtn_array= _XMLGetValue($ezXPath & "/*")
	;_ArrayDisplay($rtn_array)
	for $x=1 to $prod_node_count 
			; 取得目前 prod 中的文字資訊 ─包括 起迄站名及編號
			$prod_node= _XMLGetValue($ezXPath & '[' & $x & ']/*')   
			;_ArrayDisplay($prod_node)
			
			; 計算這個 XML 中的目前這個  prod 中有多少個 schedule
			$schedule_node_count=_XMLGetNodeCount($ezXPath & "[" & $x & "]/scheduleInfo/*") 
			;MsgBox(0, "2nd Node Count", $schedule_node_count)
				;$aNodeName = _XMLGetChildNodes ($ezXPath&"/scheduleInfo/schedule"); get a list of node names under this path
				;_ArrayDisplay($aNodeName)

			
			for $n=1 to $schedule_node_count  
				
					;	local $schedule_info[4]
					;$schedule_info=_XMLGetValue($ezXPath & "[" & $x & "]/scheduleInfo["& $n &"]/*")
					;$schedule_info=_XMLGetChildNodes($ezXPath & "[" & $x & "]/scheduleInfo["& $n &"]/*")
					;_ArrayDisplay($schedule_info)
				
				; 開始取出目前 prod中的 schedileInfo/schedule 中的三項資訊。
				$schedule_node= _XMLGetValue($ezXPath & '[' & $x & ']/scheduleInfo/schedule['&$n&']/*') 
				;_ArrayDisplay($schedule_node,$n& " of schedule_node")
				
				
				if  $schedule_node[3] -$schedule_node[2] =0  or $schedule_node[3] -$schedule_node[2] > 300 then 
					;$hs_product[0][0]= $hs_product[0][0] -1
					;ConsoleWrite ( $n & " : " &$schedule_node_count&" > " &$schedule_node[1] &" , "& $schedule_node[2] &" , "& $schedule_node[3] & @CRLF)
					
					$hs_product[$h][0]=""
					
					;$hs_product[$h][0]= $query_date								; For Date info 
					$hs_product[$h][1]= _TrainClass ($schedule_node[1],0)		; Train No
					$hs_product[$h][2]= _TrainClass ($schedule_node[1],1)		; Train Class
					$hs_product[$h][3]= $schedule_node[2]						; Train Depature time
					$hs_product[$h][4]= $schedule_node[3]						; Train Arrival Time
					$hs_product[$h][5]= $prod_node[3]							; Train Depature City Code
					$hs_product[$h][6]= $prod_node[5]							; Train Arrival City Code
					$h=$h+1
					continueLoop
					
				Else
				$hs_product[$h][0]= $query_date								; For Date info 
				$hs_product[$h][1]= _TrainClass ($schedule_node[1],0)		; Train No
				$hs_product[$h][2]= _TrainClass ($schedule_node[1],1)		; Train Class
				$hs_product[$h][3]= $schedule_node[2]						; Train Depature time
				$hs_product[$h][4]= $schedule_node[3]						; Train Arrival Time
				$hs_product[$h][5]= $prod_node[3]							; Train Depature City Code
				$hs_product[$h][6]= $prod_node[5]							; Train Arrival City Code
				$h=$h+1
				EndIf

				
					;_XMLGetAllAttrib($ezXPath  & "[" & $x & "]/scheduleInfo/schedule["& $n &"]/*",$aAttrName,$aAttrValue)
					;_XMLGetAllAttrib($ezXPath  & "/scheduleInfo/schedule["& $n &"]",$aAttrName,$aAttrValue)
					;_XMLGetAllAttrib($ezXPath  & "/scheduleInfo["& $n &"]",$aAttrName,$aAttrValue)
					;_ArrayDisplay($aAttrName,$ezXPath & "[" & $x & "]/scheduleInfo/schedule["& $n &"]/*")
					;_ArrayDisplay($aAttrValue,$ezXPath  & "[" & $x & "]/scheduleInfo/schedule["& $n &"]/*")
				
			Next

	Next

;_ArrayDisplay( $hs_product )





EndIf
;_ArrayDisplay($rtn_array)
if IsArray($hs_product) then
	local $aSTATMENT [ $hs_product[0][0] +1 ]
	;local $a_counter=0
	
	$afile=FileOpen(@ScriptDir&"\hs_product"&$query_date&".txt",130)
	$dfile=FileOpen(@ScriptDir&"\hs_product_del_"&$query_date&".txt",130)
	
	if $test_mode=0 then 
		
	endif 	
	
	for $h=1 to $hs_product[0][0]
		$aLine= $hs_product[$h][0] &","&  $hs_product[$h][1] &","& $hs_product[$h][2] &","& $hs_product[$h][3] &","& $hs_product[$h][4] &","& $hs_product[$h][5] &","& $hs_product[$h][6]
		
		;$STATEMENT_bry=$insert_statement	
		;$my=_bryant_mysql($DB_IP_bry,$DB_NAME_bry,$DB_ACCOUNT_bry,$DB_PASS_bry, $STATEMENT_bry)
			
		
		if $hs_product[$h][0]=""  then
			;ConsoleWrite( $aLine &@CRLF)
			FileWriteLine($dfile, $aLine )
			
			ContinueLoop
		Else	
			
			;$a_counter=$a_counter+1
			;_ArrayDisplay($my)
			$aSTATMENT[$h]="INSERT INTO hs_product (product_date,Tran_No,hs_class,DepTime,ArrTime,depCity,desCity) VALUES('"& $hs_product[$h][0] &"','"& $hs_product[$h][1] &"','"& $hs_product[$h][2] &"','"& $hs_product[$h][3] &"','"& $hs_product[$h][4] &"','"& $hs_product[$h][5] &"','"& $hs_product[$h][6] &"')"

			FileWriteLine($afile, $aLine )
		EndIf
	Next
	FileClose( $afile )
	FileClose( $dfile )
	
	for $h= $hs_product[0][0] to 1 step -1
		;$a_counter=$a_counter+1
		if $aSTATMENT[$h]="" then _ArrayDelete($aSTATMENT, $h)
	Next
	$aSTATMENT[0]= UBound($aSTATMENT)-1  ;$hs_product[0][0]-$a_counter
EndIf	

_ArrayDisplay($aSTATMENT)
	$STATEMENT_bry=$aSTATMENT
	$my=_bryant_mysql_array($DB_IP_bry,$DB_NAME_bry,$DB_ACCOUNT_bry,$DB_PASS_bry, $STATEMENT_bry)
	if IsArray ($my) then 
		_ArrayDisplay($my)
	Else
		MsgBox(0,"Affect on DB ", "Affect in DB '" & $DB_NAME_bry & "' by " & $my & " rows")
	;
	EndIf
	
;;
;
Exit







Func _TrainClass ( $train,$class )
	; $class is 0 means to return Train No.  $class =1 then return TrainClass
	;
	if $class=0 then 
		$train=StringTrimRight($train,1)
	
	elseif $class=1 Then
		$train=StringRight ($train,1)
	EndIf
	Return $train
EndFunc



Func _TEST_MODE()
	IF FileExists(@ScriptDir&"\TESTMODE.txt") Then
		$mode=FileReadLine(@ScriptDir&"\TESTMODE.txt",1)
		if $mode=1 then 
			MsgBox(0,"Test mode", "測試模式"&@CRLF& "高鐵車次資料只會寫入文字檔案 ",5)
			
		Else
			;MsgBox(0,"Process mode", " 高鐵車次資料會輸入資料庫 ",10)
			$ans=InputBox("Process mode","高鐵車次資料會輸入資料庫 "&@CRLF& "輸入 N 可以離開")
			
			$mode=0
			if $ans="n" or $ans="N" or @error=1 then exit
		EndIf
	
	Else
		;MsgBox(0,"Process mode", " 高鐵車次資料會輸入資料庫 ",10)
			$ans=InputBox("Process mode","高鐵車次資料會輸入資料庫 "&@CRLF& "輸入 N 可以離開")
			
			$mode=0
			if $ans="n" or $ans="N" or @error=1 then exit
		
	EndIf
	
	return $mode
EndFunc
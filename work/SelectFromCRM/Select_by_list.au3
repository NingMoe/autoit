
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.2.12.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
; From mssql_execute_array.au3  to select data
; 2013/03/07
;
#include <array.au3>
#include <file.au3>

Dim $day=@MDAY
Dim $month=@MON
DIM $year=@YEAR


If FileExists(@ScriptDir&"\sp_insert.ini") Then

	$sp_file= IniRead(@ScriptDir&"\sp_insert.ini", "sp", "sp", "NotFound")
	$start_d= IniRead(@ScriptDir&"\sp_insert.ini", "start", "start","NotFound")
	$end_d= IniRead(@ScriptDir&"\sp_insert.ini", "end", "end","NotFound")
	$program=StringMid($sp_file,3,1)
	;MsgBox(0," parameter", '"'&$sp_file &'"  <>  "'&$start_d&'" <> "'&$end_d&'" Program ID='&$program)

Else
    ; Log to file about missing .ini file
	MsgBox(0, "Not Found", @ScriptDir&"\sp_insert.ini not found.",5)
	 _FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log",@ScriptDir&"\sp_insert.ini not found.")

	Exit
EndIf

;; Check date
;if not (int($year&$month&$day)-int($end_d))=0 then
;	_FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log","Not for today. Please Check sp_insert.ini file")
;	MsgBox(0," Error ", "File is not for today",20)
;	Exit
;endif
;;
dim $md
if FileExists(@ScriptDir&"\"&$sp_file) then
	_FileReadToArray(@ScriptDir&"\"&$sp_file,$md)
	;_ArrayDisplay ($md)

	dim $aExecute[$md[0]+1]
	$aExecute[0]=$md[0]
	for $i=1 to $md[0]
		$aExecute[$i]="select * from CRMGG where GG115='"& $md[$i] &"'"
	Next
Else
	 _FileWriteLog(@ScriptDir&"\"&StringTrimRight(@ScriptName,4)&"_"&$year&$month&$day&".log", $sp_file&" files not found ")


endif
;_ArrayDisplay($aExecute)

; ($PathnFile, $aColume, $delimiters, $Start_line)
local $schema, $schema_txt
if FileExists (@ScriptDir &"\schema.txt") then
	$schema_txt=''
	$schema=_newFile2Array(@ScriptDir &"\schema.txt",3,";",0)
	;_ArrayDisplay ($schema)
	for $y=0 to UBound ($schema)-1
		;MsgBox(0, $y, $schema[$y][2])

		$schema_txt=$schema_txt &'";"'& StringStripWS( $schema[$y][2] ,8)
	Next
	$schema_txt=$schema_txt & @CRLF
	for $y=0 to UBound ($schema)-1
		$schema_txt=$schema_txt &'";"'& StringStripWS( $schema[$y][1] ,8 )
	Next

	;MsgBox(0,"Schema_txt",$schema_txt)

EndIf

Dim $server ="mail.kidsburgh.com.tw"; "140.119.116.208"; server address
Dim $user = "sa";The user
Dim $pass = "dsc@123";The ......... password !!!
Dim $dbname = "kidsburgh";"hinet_mod";the database
Dim $oConn1
Dim $oRS1 ;For Stored Proc test !
Dim $oRS2 ; For Sql Query test !
dim $return

dim $express_code="1100009239"

dim $query = "select * from CRMGG where GG115='" &$express_code & "'"

;dim $SQLQuery = "INSERT INTO yoho_coupon VALUES('53611',0)"
;dim $SQLQuery = "Update yoho_coupon set M_type=1 where MD='53611'"
dim $SQLQuery =""; "DELETE FROM yoho_coupon where M_TYPE='"&$program&"'"
dim $DSN = "Provider=SQLOLEDB;Server=" & $server & ";Database="& $dbname &";Uid=" &$user& ";Pwd=" &$pass
;==============
;dim $aExecute[4]
;$aExecute[0]=3
;$aExecute[1]="INSERT INTO yoho_coupon VALUES('53611123',0)"
;$aExecute[2]="INSERT INTO yoho_coupon VALUES('53611124',0)"
;$aExecute[3]="INSERT INTO yoho_coupon VALUES('53611125',0)"
;==========
;_ExecuteQuery($DSN, $SQLQuery)
;_Execute_a_Query($DSN,$aExecute)
;sleep(10000)
;$return=_Query($query)
$return= _aQuery($aExecute)

;_ArrayDisplay($return)

Func _ExecuteQuery ($DSN,$SQLQuery)
$adoSQL = $SQLQuery
$adoCon = ObjCreate ("ADODB.Connection")
$adoCon.Open ($DSN)
$adoCon.Execute($adoSQL)
$adoCon.Close
EndFunc

Func _Execute_a_Query ($DSN,$aExecute1)

$adoCon = ObjCreate ("ADODB.Connection")
$adoCon.Open ($DSN)
for $i=1 to $aExecute1[0]
	$adoSQL = $aExecute1[$i]
	$adoCon.Execute($adoSQL)
next
$adoCon.Close
EndFunc


;;  Query from an array data to another array data
Func _aQuery($query1)
Dim $adOpenForwardOnly = 0 ;Valeur par defaut. Utilise un curseur a defilement en avant. Identique a un curseur statique mais ne permettant que de faire defiler les enregistrements vers l'avant. Ceci accroit les performances lorsque vous ne devez effectuer qu'un seul passage dans un Recordset.
Dim $adOpenKeyset = 1 ;Utilise un curseur a jeu de cles. Identique a un curseur dynamique mais ne permettant pas de voir les enregistrements ajoutes par d'autres utilisateurs (les enregistrements supprimes par d'autres utilisateurs ne sont pas accessibles a partir de votre Recordset). Les modifications de donnees effectuees par d'autres utilisateurs demeurent visibles.
Dim $adOpenDynamic = 2 ;Utilise un curseur dynamique. Les ajouts, modifications et suppressions effectues par d'autres utilisateurs sont visibles et tous les deplacements sont possibles dans le Recordset a l'exception des signets, s'ils ne sont pas pris en charge par le fournisseur.
Dim $adOpenStatic = 3 ;Utilise un curseur statique. Copie statique d'un jeu d'enregistrements qui permet de trouver des donnees ou de generer des etats. Les ajouts, modifications ou suppressions effectues par d'autres utilisateurs ne sont pas visibles.
Dim $adOpenUnspecified = -1


Dim $adLockBatchOptimistic =4 ;Mise a jour par lot optimiste. Obligatoire en mode de mise a jour par lots.
Dim $adLockOptimistic =3 ;Verrouillage optimiste, un enregistrement a la fois. Le fournisseur utilise le verrouillage optimiste et ne verrouille les enregistrements qu'a l'appel de la methode Update.
Dim $adLockPessimistic =2 ;Verrouillage pessimiste, un enregistrement a la fois. Le fournisseur assure une modification correcte des enregistrements, et les verrouille generalement dans la source de donnees des l'edition.
Dim $adLockReadOnly =1 ;Enregistrements en lecture seule. Vous ne pouvez pas modifier les donnees.
Dim $adLockUnspecified =-1

Dim $adCmdUnspecified = -1 ;Ne specifie pas l'argument type de la commande.
Dim $adCmdText = 1 ;CommandText correspond a la definition textuelle d'une commande ou d'un appel de procedure stockee.
Dim $adCmdTable = 2 ;CommandText correspond au nom de la table dont les colonnes sont toutes renvoyees par une requete SQL generee en interne.
Dim $adCmdStoredProc = 4 ;CommandText correspond au nom d'une procedure stockee.
Dim $adCmdUnknown = 8 ;Valeur par defaut. Indique que le type de commande specifie dans la propriete CommandText n'est pas connu.
Dim $adCmdFile = 256 ;CommandText correspond au nom de fichier d'un Recordset stocke de facon permanente.
Dim $adCmdTableDirect = 512 ;CommandText correspond au nom d'une table dont les colonnes sont toutes renvoyees.

Dim $IntRetour = 0
Dim $intNbLignes = 0
local $retrived_data_in_string=$schema_txt&@CRLF
;local $return_is_array=0

	;~ Dim $Sepa = String(@TAB); for having tab for separator !
	Dim $Sepa = '";"' ;like CSV

Dim $strConnexion = "DRIVER={SQL Server};SERVER=" & $server & ";DATABASE=" & $dbname & ";uid=" & $user & ";pwd=" & $pass & ";"
$oConn1 = ObjCreate("ADODB.Connection")
$oConn1.open($strConnexion)

for $x=1 to UBound ($query1)-1

	;MsgBox(0,"Array data to query  No:" & $x & " / "& ( UBound($query1)-1), $query1[$x] )
	;For a test of a simple text Query
	;Dim $strSqlQuerySelect = "select VID from edu_media where LEAF_ID like '0000%' AND ORIG='DVD'"
	Dim $strSqlQuerySelect = $query1[$x]
		;Dim $strSqlQuerySelect = "select * from edu_media where MS_IP1='172.17.247.22:5004/asset/mds%3a/mds/video3/'"
	$oRS2 = ObjCreate("ADODB.Recordset")
	$oRS2.open($strSqlQuerySelect,$oConn1,$adOpenStatic,$adLockOptimistic,$adCmdText)
	$intNbLignes = $oRS2.recordCount ;i re-use $intNbLignes !
		;For getting Fields Header Names...
		;For $j = 0 To $oRS2.Fields.count -1
		;ConsoleWrite($oRS2.Fields($j).name & $Sepa)
		;Next
		;ConsoleWrite (@CRLF);to rewind
	if $intNbLignes>0 then
		local $query_return[$oRS2.recordCount][$oRS2.Fields.count]
		For $i = 0 To $intNbLignes -1 ; or "For $i = 0 To ($oRS2.recordCount - 1)" if you want

			For $j = 0 To $oRS2.Fields.count -1
				;ConsoleWrite($oRS2.Fields($j).value & $Sepa)
				$query_return[$i][$j]=$oRS2.Fields($j).value

				$retrived_data_in_string=$retrived_data_in_string &$Sepa& $oRS2.Fields($j).value
			Next
			;ConsoleWrite (@CRLF);to begin the next line
			$oRS2.MoveNext ; DON'T FORGET THIS OR CRASH...!!!
			$retrived_data_in_string=$retrived_data_in_string &"^^^"&@CRLF
		Next


		;_ArrayDisplay($query_return)

		;MsgBox(0,"Retrived data in string , Total : " & $i & " Line", $retrived_data_in_string )
		;return($retrived_data_in_string)
	else
		MsgBox(0,"No Data", "No Data From DataBase",3)
		FileWriteLine(@ScriptDir &"\query_without_data.txt", $strSqlQuerySelect & @CRLF)
	EndIf


Next
	;MsgBox(0,"Retrived data in string , Total : " & $i & " Line", $retrived_data_in_string )
	FileWrite(@ScriptDir &"\query_output.txt", $retrived_data_in_string)
	return($retrived_data_in_string)

$oRS2.close
$oConn1.close
EndFunc






Func _Query($query1)
Dim $adOpenForwardOnly = 0 ;Valeur par defaut. Utilise un curseur a defilement en avant. Identique a un curseur statique mais ne permettant que de faire defiler les enregistrements vers l'avant. Ceci accroit les performances lorsque vous ne devez effectuer qu'un seul passage dans un Recordset.
Dim $adOpenKeyset = 1 ;Utilise un curseur a jeu de cles. Identique a un curseur dynamique mais ne permettant pas de voir les enregistrements ajoutes par d'autres utilisateurs (les enregistrements supprimes par d'autres utilisateurs ne sont pas accessibles a partir de votre Recordset). Les modifications de donnees effectuees par d'autres utilisateurs demeurent visibles.
Dim $adOpenDynamic = 2 ;Utilise un curseur dynamique. Les ajouts, modifications et suppressions effectues par d'autres utilisateurs sont visibles et tous les deplacements sont possibles dans le Recordset a l'exception des signets, s'ils ne sont pas pris en charge par le fournisseur.
Dim $adOpenStatic = 3 ;Utilise un curseur statique. Copie statique d'un jeu d'enregistrements qui permet de trouver des donnees ou de generer des etats. Les ajouts, modifications ou suppressions effectues par d'autres utilisateurs ne sont pas visibles.
Dim $adOpenUnspecified = -1


Dim $adLockBatchOptimistic =4 ;Mise a jour par lot optimiste. Obligatoire en mode de mise a jour par lots.
Dim $adLockOptimistic =3 ;Verrouillage optimiste, un enregistrement a la fois. Le fournisseur utilise le verrouillage optimiste et ne verrouille les enregistrements qu'a l'appel de la methode Update.
Dim $adLockPessimistic =2 ;Verrouillage pessimiste, un enregistrement a la fois. Le fournisseur assure une modification correcte des enregistrements, et les verrouille generalement dans la source de donnees des l'edition.
Dim $adLockReadOnly =1 ;Enregistrements en lecture seule. Vous ne pouvez pas modifier les donnees.
Dim $adLockUnspecified =-1

Dim $adCmdUnspecified = -1 ;Ne specifie pas l'argument type de la commande.
Dim $adCmdText = 1 ;CommandText correspond a la definition textuelle d'une commande ou d'un appel de procedure stockee.
Dim $adCmdTable = 2 ;CommandText correspond au nom de la table dont les colonnes sont toutes renvoyees par une requete SQL generee en interne.
Dim $adCmdStoredProc = 4 ;CommandText correspond au nom d'une procedure stockee.
Dim $adCmdUnknown = 8 ;Valeur par defaut. Indique que le type de commande specifie dans la propriete CommandText n'est pas connu.
Dim $adCmdFile = 256 ;CommandText correspond au nom de fichier d'un Recordset stocke de facon permanente.
Dim $adCmdTableDirect = 512 ;CommandText correspond au nom d'une table dont les colonnes sont toutes renvoyees.

Dim $IntRetour = 0
Dim $intNbLignes = 0

;~ Dim $Sepa = String(@TAB); for having tab for separator !
Dim $Sepa = ";" ;like CSV

Dim $strConnexion = "DRIVER={SQL Server};SERVER=" & $server & ";DATABASE=" & $dbname & ";uid=" & $user & ";pwd=" & $pass & ";"
$oConn1 = ObjCreate("ADODB.Connection")
$oConn1.open($strConnexion)


;For a test of a simple text Query
;Dim $strSqlQuerySelect = "select VID from edu_media where LEAF_ID like '0000%' AND ORIG='DVD'"
Dim $strSqlQuerySelect = $query1
;Dim $strSqlQuerySelect = "select * from edu_media where MS_IP1='172.17.247.22:5004/asset/mds%3a/mds/video3/'"
$oRS2 = ObjCreate("ADODB.Recordset")
$oRS2.open($strSqlQuerySelect,$oConn1,$adOpenStatic,$adLockOptimistic,$adCmdText)
$intNbLignes = $oRS2.recordCount ;i re-use $intNbLignes !
;For getting Fields Header Names...
;For $j = 0 To $oRS2.Fields.count -1
;ConsoleWrite($oRS2.Fields($j).name & $Sepa)
;Next
;ConsoleWrite (@CRLF);to rewind
if $intNbLignes>0 then
	local $query_return[$oRS2.recordCount][$oRS2.Fields.count]
	For $i = 0 To $intNbLignes -1 ; or "For $i = 0 To ($oRS2.recordCount - 1)" if you want
		For $j = 0 To $oRS2.Fields.count -1
			;ConsoleWrite($oRS2.Fields($j).value & $Sepa)
			$query_return[$i][$j]=$oRS2.Fields($j).value
		Next
		;ConsoleWrite (@CRLF);to begin the next line
		$oRS2.MoveNext ; DON'T FORGET THIS OR CRASH...!!!
	Next
	;_ArrayDisplay($query_return)
	return($query_return)
else
	MsgBox(0,"No Data", "No Data From DataBase")


EndIf

$oRS2.close
$oConn1.close
EndFunc



;; 2011 09 New Text to Two dimension array  with a start_line
Func _newFile2Array($PathnFile, $aColume, $delimiters, $Start_line)


	Local $aRecords
	If Not _FileReadToArray($PathnFile, $aRecords) Then
		MsgBox(4096, "Error", " Error reading file '" & $PathnFile & "' to Array   error:" & @error)
		Exit
	EndIf

	;_ArrayDisplay ($aRecords)
	For $x = $Start_line To 1 Step -1
		_ArrayDelete($aRecords, $x)
	Next
	$aRecords[0] = $aRecords[0] - $Start_line
	;_ArrayDisplay ($aRecords)
	;c
	Local $TextToArray[$aRecords[0]][$aColume + 1]
	;$TextToArray[0][0]=$aRecords[0]
	Local $aRow
	For $y = 1 To $aRecords[0]
		;Msgbox(0,'Record:' & $y, $aRecords[$y])

		$aRow = StringSplit($aRecords[$y], $delimiters)
		;Msgbox(0,'X ,Colume :', $aRow[0])
		;_ArrayDisplay ($aRow)
		For $x = 1 To $aRow[0]
			;If StringInStr($aRow[$x], ",") Then

			;	$aRow[$x] = StringTrimLeft($aRow[$x], 1)
			;	;MsgBox(0, "after", $aRow[$x])
			;EndIf
			$TextToArray[$y - 1][$x - 1] = $aRow[$x]
		Next
	Next

	;_ArrayDisplay($TextToArray)
	Return $TextToArray

EndFunc   ;==>_newFile2Array
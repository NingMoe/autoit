#include <array.au3>

dim $array[6]

$array[0]="5"
$array[1]="AA"
$array[2]="aa"
$array[3]="CC"
$array[4]="DD"
$array[5]="AA"

for $r=1 to $array[0]
MsgBox ( 0,"Search in an array", "component AA is at :" &_ArraySearch($array,"AA",1,$array[0],1) )

$array_component_to_hunt =_ArraySearch($array,"AA",1,$array[0],1) 
if  $array_component_to_hunt >=1 then 
_ArrayDelete($array, $array_component_to_hunt )
$array[0]-=1
EndIf
_ArrayDisplay ($array)

Next
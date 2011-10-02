#include <Date.au3>
dim $now_DateCalc

while 1
$now_DateCalc = _DateDiff( 's',"1970/01/01 00:00:00",_NowCalc())
ToolTip("EPOCH current :"& $now_DateCalc, 10, 10); ... we're connected.
sleep(2000)
WEnd
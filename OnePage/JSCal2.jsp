<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
	<head>
	<link type="text/css" rel="stylesheet" href="css/JSCal2/css/jscal2.css" />
	<link type="text/css" rel="stylesheet" href="css/JSCal2/css/border-radius.css" />
	<!-- <link type="text/css" rel="stylesheet" href="css/JSCal2/css/reduce-spacing.css" /> -->

	<link id="skin-win2k" title="Win 2K" type="text/css" rel="alternate stylesheet" href="css/JSCal2/css/win2k/win2k.css" />
	<link id="skin-steel" title="Steel" type="text/css" rel="alternate stylesheet" href="css/JSCal2/css/steel/steel.css" />
	<link id="skin-gold" title="Gold" type="text/css" rel="alternate stylesheet" href="css/JSCal2/css/gold/gold.css" />
	<link id="skin-matrix" title="Matrix" type="text/css" rel="alternate stylesheet" href="css/JSCal2/css/matrix/matrix.css" />

	<link id="skinhelper-compact" type="text/css" rel="alternate stylesheet" href="css/JSCal2/css/reduce-spacing.css" />

	<script src="css/JSCal2/js/jscal2.js"></script>
	<script src="css/JSCal2/js/unicode-letter.js"></script>

	<!-- you actually only need to load one of these; we put them all here for demo purposes -->
	<script src="css/JSCal2/js/lang/ca.js"></script>
	<script src="css/JSCal2/js/lang/cn.js"></script>
	<script src="css/JSCal2/js/lang/cz.js"></script>
	<script src="css/JSCal2/js/lang/de.js"></script>
	<script src="css/JSCal2/js/lang/es.js"></script>
	<script src="css/JSCal2/js/lang/fr.js"></script>
	<script src="css/JSCal2/js/lang/hr.js"></script>
	<script src="css/JSCal2/js/lang/it.js"></script>
	<script src="css/JSCal2/js/lang/jp.js"></script>
	<script src="css/JSCal2/js/lang/nl.js"></script>
	<script src="css/JSCal2/js/lang/pl.js"></script>
	<script src="css/JSCal2/js/lang/pt.js"></script>
	<script src="css/JSCal2/js/lang/ro.js"></script>
	<script src="css/JSCal2/js/lang/ru.js"></script>
	<script src="css/JSCal2/js/lang/sk.js"></script>
	<script src="css/JSCal2/js/lang/sv.js"></script>

	<!-- this must stay last so that English is the default one -->
	<script src="css/JSCal2/js/lang/en.js"></script>

	<link type="text/css" rel="stylesheet" href="demopage.css" />
	</head>
	<body style="background-color: #fff">
			<form name="test">
			<%for(int i = 0 ; i < 10 ; i++){ %>
			<input id="actual_expense_date<%=i %>" name="actual_expense_date" style="text-align:right;"/>
			<button id="f_clearRangeStart" onclick="clearRangeStart(<%=i %>)">clear</button>
			<br>
			<%} %>
			</form>
			<script type="text/javascript">
				var e = document.getElementsByName("actual_expense_date");
				var RANGE_CAL = new Array(e.length);
				for(var i = 0 ; i < e.length ; i++){
					RANGE_CAL[i] = new Calendar({
						inputField: e[i].id,
						dateFormat: "%Y/%m/%d",
						trigger: e[i].id,
						bottomBar: false,
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
						}
					});
				}
				function clearRangeStart(n) {
					var e = document.forms['test'];
					e.elements["actual_expense_date"][n].value = "";
				};
				//
				var links = document.getElementsByTagName("link");
				var skins = {};
				for (var i = 0; i < links.length; i++) {
					if (/^skin-(.*)/.test(links[i].id)) {
						var id = RegExp.$1;
						skins[id] = links[i];
					}
				}
				var skin = "gold";
				for (var i in skins) {
					if (skins.hasOwnProperty(i))
						skins[i].disabled = true;
				}
				if (skins[skin])
					skins[skin].disabled = false;
			</script>
	</body>
</html>
<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>OnePage Login</title>
<link rel="stylesheet" type="text/css" href="css/css.css"/>
</head>

<body>
<div id="wrapper">
	<div id="main" class="clearfix">
		<div id="header">
			<div id="inheader">
				<table width="100%">
					<tr>
					<td width="100px"></td>
					<td><h3>OnePage Login</h3></td>
					<td width="100px" valign="bottom"><p align="right"><a href="Login.jsp">Logout</a>　</p></td>
					</tr>
				</table>
			</div>
		</div><!--end header-->
		<div id="content" >
			<form action="GotoLogin?status=login" method="post" style="background-color:#FFF">
			<table width="100%" border="0" align="center">
						<tr>
							<td width="15%"></td>
							<td width="70%" align="center">
							帳號  : <input name="userid" type="text" style="width:80%;" />
							</td>
							<td width="15%"></td>
						</tr>
						<tr>
							<td width="15%"></td>
							<td width="70%" align="center">
							密碼 : <input name="pwd" type="password" style="width:80%;" />
							</td>
							<td width="15%"></td>
						</tr>
						<tr>
							<td width="15%"></td>
							<td width="70%" align="right">
							<input name="" type="submit" value="登入" />
							</td>
							<td width="15%"></td>
						</tr>
					</table>
			</form>	
		</div><!--end content-->
	</div><!--end main-->
</div><!--end wrapper-->
<div id="footer">
	<div id="infoot">
		<p><font size="+1">美國運通旅遊部 版權所有 © 2010</font></p>
	</div><!--end infoot-->
</div><!--end footer-->
</body>
</html>

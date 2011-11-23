<%@page import="java.util.HashMap" %>
<%@page import="java.util.ArrayList" %>
<%@page import="java.sql.ResultSet" %>

<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
    
<%
ArrayList UserBudget = null;//user 的預算資料

if(request.getAttribute("UserBudget") != null){
	UserBudget = (ArrayList)request.getAttribute("UserBudget");
}
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>OnePage</title>
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
					<td>
					<h3>
					OnePage / 
					<%
					out.println(session.getAttribute("corp_abbr")+" / "+session.getAttribute("user_cname"));
					switch(Integer.parseInt((String)session.getAttribute("user_authority"))){
						//管理:4決行:2批核:1員工:0
						case 0:
							
							break;
						case 1:
							out.println("(批核)");
							break;
						case 2:
							out.println("(決行)");
							break;
						case 4:
							out.println("(管理)");
							break;
					}
					%>
					</h3>
					</td>
					<td width="100px" valign="bottom">
						<p align="right">
							<font size=3><a href="GotoLogin?status=logout">Logout</a></font>
						</p>
					</td>
					</tr>
				</table>
			</div>
		</div><!--end header-->
		<div id="content">
			<table width="100%" border="0" align="center">
				<tr>
					<td width="15%" align="center"><a href="User">[個人出差]</a></td>
					<td width="15%" align="center">[個人資料維護]</td>
					<%if(session.getAttribute("user_authority") != null 
						&& session.getAttribute("user_authority").equals("4")){ %>
						<td width="15%" align="center"><a href="Manager?">[員工資料維護]</a></td>
						<td width="15%" align="center">[費用資料維護]</td>
						<td width="40%"></td>
					<%}else if(session.getAttribute("user_authority") != null 
						&& session.getAttribute("user_authority").equals("2")){ %>
						<td width="15%" align="center"><a href="TopUser?">[預算維護]</a></td>
						<td width="55%"></td>
					<%}else{%>
						<td width="70%"></td>
					<%}%>
				</tr>
			</table>
			<br>
			budget_nbr=<%=((HashMap)UserBudget.get(UserBudget.size()-1)).get("budget_nbr") %><br>
			user_nbr=<%=((HashMap)UserBudget.get(UserBudget.size()-1)).get("user_nbr") %><br>
			budget_total=<%=((HashMap)UserBudget.get(UserBudget.size()-1)).get("budget_total") %><br>
			<form method="post" action="TopUser">
				<input name="isStatus" type="hidden" value="edit" />
				<input name="budget_nbr" type="hidden" value="<%=((HashMap)UserBudget.get(UserBudget.size()-1)).get("budget_nbr") %>" />
				<table width="100%" border="0" align="center">
					<tr>
						<td width="10%" align="right">總預算</td>
						<td width="20%" align="right" style="padding-right:5px;">
							<input type="text" style="text-align:right; width: 99%;" 
								id="budget_total" name="budget_total" 
								value="<%=((HashMap)UserBudget.get(UserBudget.size()-1)).get("budget_total") %>">
						</td>
						<td width="70%">
						<input type="submit" value="修改" name="submit">
						</td>
					</tr>
					<tr>
						<td align="right">己決行預算</td>
						<td align="right">
						<%=((HashMap)UserBudget.get(UserBudget.size()-1)).get("budget_grant") %>
						</td>
						<td></td>
					</tr>
					<tr>
						<td align="right">已報帳預算</td>
						<td align="right">
						<%=((HashMap)UserBudget.get(UserBudget.size()-1)).get("budget_booked") %>
						</td>
						<td></td>
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

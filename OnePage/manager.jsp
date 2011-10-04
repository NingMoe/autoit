<%@page import="java.sql.ResultSet" %>
<%@page import="java.sql.ResultSetMetaData" %>
<%@page import="java.lang.*" %>
<%@page import="java.util.HashMap" %>

<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<%
response.setContentType("text/html;charset=UTF-8");
request.setCharacterEncoding("UTF-8");

ResultSet data = (ResultSet)request.getAttribute("data");
int data_cont = 0;
while(data.next()){
	data_cont++;
}
data.beforeFirst();
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
		<div id="content" style="width:1550px;">
			
			<table width="100%" border="0" align="center">
				<tr>
					<td width="15%" align="center"><a href="User">[個人出差]</a></td>
					<td width="15%" align="center">[個人資料維護]</td>
					<%if(session.getAttribute("user_authority") != null 
						&& session.getAttribute("user_authority").equals("4")){ %>
						<td width="15%" align="center"><a href="Manager?">[員工資料維護]</a></td>
						<td width="15%" align="center">[費用資料維護]</td>
						<td width="40%"></td>
					<%}else{%>
						<td width="70%"></td>
					<%}%>
				</tr>
			</table>
			
			<br>
			
			<%if(request.getParameter("status") == null 
					|| request.getParameter("status").equals("list") 
					|| request.getParameter("status").equals("edit")
					|| request.getParameter("status").equals("newUser")){%>
			
				<table width="100%" border="0">
					<%if(data != null){%>
					<tr>
						<td width="4%" align="center">
						</td>
						<td width="4%" align="center">名字</td>
						<td width="18%" align="center">帳號 ( eMail )</td>
						<td width="5%" align="center">密碼</td>
						<td width="7%" align="center">電話</td>
						<td width="4%" align="center">職別</td>
						<td width="4%" align="center">職等</td>
						<td width="4%" align="center">權限</td>
						<td width="6%" align="center">啟用日期</td>
						<td width="6%" align="center">停用日期</td>
						<td width="5%" align="center">LastName</td>
						<td width="5%" align="center">FirstName</td>
						<td width="6%" align="center">身分證字號</td>
						<td width="5%" align="center">上次登入</td>
						<td width="6%" align="center">上次批核者</td>
						<td width="6%" align="center">上次決行者</td>
					</tr>
					<form action="Manager?status=newUser" method="post">
						<tr>
							<td align="center">
								<input name="submit" type="submit" value="新增" />
							</td>
							<td style="padding-right:5px;">
								<input name="cname" type="text" value="" id="cname" style="width:99%;"/>
							</td>
							<td style="padding-right:5px;">
								<input name="login" type="text" value="" id="login" style="width:99%;"/>
							</td>
							<td style="padding-right:5px;">
								<input name="pswd" type="text" value="" id="pswd" style="width:99%;"/>
							</td>
							<td style="padding-right:5px;">
								<input name="phone" type="text" value="" id="phone" style="width:99%;"/>
							</td>
							<td style="padding-right:5px;">
								<input name="role" type="text" value="" id="role" style="width:99%;"/>
							</td>
							<td style="padding-right:5px;">
								<input name="rank" type="text" value="" id="rank" style="width:99%;"/>
							</td>
							<td align="center">
								<select name="authority" id="authority">
									<option value="0" selected="selected">員工</option>
									<option value="2">決行</option>
									<option value="4">管理</option>
								</select>
							</td>
							<td></td>
							<td></td>
							<td style="padding-right:5px;">
								<input name="lastname" type="text" value="" id="lastname" style="width:99%;"/>
							</td>
							<td style="padding-right:5px;">
								<input name="firstname" type="text" value="" id="firstname" style="width:99%;"/>
							</td>
							<td style="padding-right:5px;">
								<input name="id" type="text" value="" id="id" style="width:99%;"/>
							</td>
							<td></td>
							<td></td>
							<td width="6%" ></td>
						</tr>
					</form>
						<%while(data.next()){ %>
						<form action="Manager?status=edit" method="post">
							<input type="hidden" name="user_nbr" value="<%=data.getString("user_nbr")%>">
							<input type="hidden" name="login_old" value="<%=data.getString("login")%>">
							<tr>
								<td align="center">
									<%
									if(data.getString("date_end") != null 
										&& !data.getString("date_end").equals("")
										&& data.getString("date_end").split("-") != null){
										
									}else{
									%>
										<input name="submit" type="submit" value="更新" />
									<%
									}
									%>
								</td>
								<td style="padding-right:5px;">
									<input name="cname" type="text" 
									value="<%=data.getString("cname")!=null?data.getString("cname"):""%>" 
									id="cname" style="width:99%;"/>
								</td>
								<td style="padding-right:5px;">
									<input name="login" type="text" 
									value="<%=data.getString("login")!=null?data.getString("login"):""%>" 
									id="login" style="width:99%;"/>
								</td>
								<td style="padding-right:5px;">
									<input name="psw" type="text"
									value="****" 
									id="psw" style="width:99%;"/>
								</td>
								<td style="padding-right:5px;">
									<input name="phone" type="text" 
									value="<%=data.getString("phone")!=null?data.getString("phone"):""%>" 
									id="phone" style="width:99%;"/>
								</td>
								<td style="padding-right:5px;">
									<input name="role" type="text" 
									value="<%=data.getString("role")!=null?data.getString("role"):""%>" 
									id="role" style="width:99%;"/>
								</td>
								<td style="padding-right:5px;">
									<input name="rank" type="text" 
									value="<%=data.getString("rank")!=null?data.getString("rank"):""%>" 
									id="rank" style="width:99%;"/>
								</td>
								<td align="center">
									<%if(data.getString("authority") != null && !data.getString("authority").equals("")){
										int aut = Integer.parseInt(data.getString("authority"));
									%>
										<select name="authority" id="authority">
											<option value="0" <%=aut==0?"selected=\"selected\"":""%>>員工</option>
											<option value="2" <%=aut==2?"selected=\"selected\"":""%>>決行</option>
											<option value="4" <%=aut==4?"selected=\"selected\"":""%>>管理</option>
										</select>
									<%}%>
								</td>
								<td align="center">
								<%
								if(data.getString("date_start") != null 
										&& !data.getString("date_start").equals("")
										&& data.getString("date_start").split("-") != null){
									String tmp_data[] = data.getString("date_start").split("-");
									if(Integer.parseInt(tmp_data[1]) < 10){
										tmp_data[1] = "0" + tmp_data[1];
									}
									if(Integer.parseInt(tmp_data[2]) < 10){
										tmp_data[2] = "0" + tmp_data[2];
									}
									out.println(tmp_data[0]+tmp_data[1]+tmp_data[2]);
								}
								%>
								</td>
								<td align="center">
								<%
								if(data.getString("date_end") != null 
										&& !data.getString("date_end").equals("")
										&& data.getString("date_end").split("-") != null){
									String tmp_data[] = data.getString("date_end").split("-");
									if(Integer.parseInt(tmp_data[1]) < 10){
										tmp_data[1] = "0" + tmp_data[1];
									}
									if(Integer.parseInt(tmp_data[2]) < 10){
										tmp_data[2] = "0" + tmp_data[2];
									}
									out.println(tmp_data[0]+tmp_data[1]+tmp_data[2]);
								}else{
								%>	
									<input type="checkbox" name="date_end" id="date_end" value="<%=data.getString("user_nbr")%>"/>
								<%	
								}
								%>
								</td>
								<td style="padding-right:5px;">
									<input name="lastname" type="text" 
									value="<%=data.getString("lastname")!=null?data.getString("lastname"):""%>" 
									id="lastname" style="width:99%;"/>
								</td>
								<td style="padding-right:5px;">
									<input name="firstname" type="text" 
									value="<%=data.getString("firstname")!=null?data.getString("firstname"):""%>" 
									id="firstname" style="width:99%;"/>
								</td>
								<td style="padding-right:5px;">
									<input name="id" type="text" 
									value="<%=data.getString("id")!=null?data.getString("id"):""%>" 
									id="id" style="width:99%;"/>
								</td>
								<td>
								<%
								if(data.getString("date_ack") != null && data.getString("date_ack").split("-") != null){
									String tmp_data[] = data.getString("date_ack").split("-");
									if(Integer.parseInt(tmp_data[1]) < 10){
										tmp_data[1] = "0" + tmp_data[1];
									}
									if(Integer.parseInt(tmp_data[2]) < 10){
										tmp_data[2] = "0" + tmp_data[2];
									}
									out.println(tmp_data[0]+tmp_data[1]+tmp_data[2]);
								}
								%>
								</td>
								<td><%out.print(data.getString("approver")); %></td>
								<td><%out.print(data.getString("controller")); %></td>
							</form>
						</tr>
						<%}%>
					<%}%>
				</table>
			
			<%}%>
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


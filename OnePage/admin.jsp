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
ResultSetMetaData rm = (ResultSetMetaData)request.getAttribute("rm");
int data_cont = 0;
while(data.next()){
	data_cont++;
}
data.beforeFirst();
/*
HashMap cManagers = (HashMap)request.getAttribute("CorpManagers");
if(cManagers == null){
	cManagers = new HashMap();
}
*/
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
						<td><h3><a href="Admin">Admin/${msg}</a></h3></td>
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
			
			<%if(request.getParameter("status") == null || request.getParameter("status").equals("list")){%>
				<form action="Admin" method="get">
					<% 
					int n = 0;
					int n1 , n2;
					if(request.getParameter("page") != null){
						if(Integer.parseInt(request.getParameter("page")) > 0){
							n = Integer.parseInt(request.getParameter("page"));
						}
					}
					n1 = n+100;
					n2 = n-100;
					if(n2 < 0)
						n2 = 0;
					%>
					<table width="100%" align="center">
						<tr>
							<%if(request.getParameter("select") == null){%>
								<td width="10%" align="center">
									<input name="select" type="text" id="select" />
								</td>
								<td width="5%" align="center">
									<input name="submit" type="submit" value="搜尋" />
								</td>
								<td width="65%"></td>
								<td width="10%" align="center">
									<a href="Admin?status=list&page=<%=n2%>">上一頁</a>
								</td>
								<td width="10%" align="center">
									<%if(!(data_cont < 100)){%>
									<a href="Admin?status=list&page=<%=n1%>">下一頁</a>
									<%}%>
								</td>
							<%}else{%>
								<td width="10%" align="center">
									<%if(!request.getParameter("select").equals("")){ %>
									<input name="select" type="text" id="select" value='<%=request.getParameter("select")%>' />
									<%}else{%>
									<input name="select" type="text" id="select" />
									<%} %>
								</td>
								<td width="5%" align="center">
									<input name="submit" type="submit" value="搜尋" />
								</td>
								<td width="65%"></td>
								<td width="10%" align="center">
									<a href="Admin?status=list&select=<%=request.getParameter("select")%>&page=<%=n2%>">上一頁</a>
								</td>
								<td width="10%" align="center">
									<%if(!(data_cont < 100)){%>
									<a href="Admin?status=list&select=<%=request.getParameter("select")%>&page=<%=n1%>">下一頁</a>
									<%}%>
								</td>
							<%}%>
						</tr>
					</table>
				</form>
				
				${Error_msg}<br>
				<%if(data != null){%>
					<table width="100%">
						<tr>
							<td width="5%" align="center">代碼</td>
							<td width="40%">公司名稱(中文/英文)</td>
							<td width="10%" align="center">統編/簡稱</td>
							<td width="25%">管理者聯絡方式</td>
							<td width="10%" align="center">啟用日期</td>
							<td width="10%" align="center">停用日期</td>
						</tr>
						<%while(data.next()){%>
						<tr>
							<td><% out.print(data.getString("corp_id")); %></td>
							<td>
								<% 
								if(data.getString("cname").equals(data.getString("ename")))
									out.print(data.getString("cname"));
								else
									out.print(data.getString("cname")+"<br>"+data.getString("ename")); 
								%>
							</td>
							<td align="center">
								<%out.print(data.getString("idno"));%><br>
								<%out.print(data.getString("abbr"));%>
							</td>
							<td>
								<% 
								out.println(data.getString("cont") +" , "+data.getString("tel") +"<br>");
								String mail = data.getString("email");
								if(mail != null){
									mail = mail.toLowerCase();
									if(mail.indexOf(";") != -1){
										out.print(mail.substring(0,mail.indexOf(";"))+"<br>");
										out.print(mail.substring(mail.indexOf(";")+1,mail.length()));
									}else if(data.getString("email").indexOf(",") != -1){
										out.print(mail.substring(0,mail.indexOf(","))+"<br>");
										out.print(mail.substring(mail.indexOf(",")+1,mail.length()));
									}else{
										out.print(mail);
									}
								}
								
								/*
								if(cManagers != null){
									mail = (String)cManagers.get(data.getString("idno"));
									if(mail != null){
										mail = mail.toLowerCase();
										if(mail.indexOf(";") != -1){
											out.print(mail.substring(0,mail.indexOf(";"))+"<br>");
											out.print(mail.substring(mail.indexOf(";")+1,mail.length()));
										}else if(data.getString("email").indexOf(",") != -1){
											out.print(mail.substring(0,mail.indexOf(","))+"<br>");
											out.print(mail.substring(mail.indexOf(",")+1,mail.length()));
										}else{
											out.print(mail);
										}
									}
								}
								*/
								%>
							</td>
							<td align="center">
								<%if(data.getString("idno") != null && !data.getString("idno").equals("")
										&& data.getString("email") != null && !data.getString("email").equals("")){%>
									<%if(data.getString("start_date").equals("")){%>
										<a href="Admin?status=list
											<%=
											"&page="+n+
											"&onservices=on"+
											"&user_id="+data.getString("email")+
											"&user_pwd=admin"+
											"&corp_id="+data.getString("corp_id")+
											"&corp_idno="+data.getString("idno")+
											((request.getParameter("select")!=null?"&select="+request.getParameter("select"):""))%>">
										啟用
										</a>
									<%}else{%>
										<a href="Admin?status=list
											<%=
											"&page="+n+
											"&onservices=on"+
											"&user_id="+data.getString("email")+
											"&user_pwd=admin"+
											"&corp_id="+data.getString("corp_id")+
											"&corp_idno="+data.getString("idno")+
											((request.getParameter("select")!=null?"&select="+request.getParameter("select"):""))%>">
										<%
										if(data.getString("start_date").split("-") != null){
											String tmp_data[] = data.getString("start_date").split("-");
											if(Integer.parseInt(tmp_data[1]) < 10){
												tmp_data[1] = "0" + tmp_data[1];
											}
											if(Integer.parseInt(tmp_data[2]) < 10){
												tmp_data[2] = "0" + tmp_data[2];
											}
											out.println(tmp_data[0]+tmp_data[1]+tmp_data[2]);
										}
										%>
										</a>
									<%}%>
								<%}%>
							</td>
							<td align="center">
								<%if(data.getString("idno") != null && !data.getString("idno").equals("")
										&& data.getString("email") != null && !data.getString("email").equals("")){%>
									<%if(data.getString("stop_date").equals("")){%>
										<a href="Admin?status=list
											<%="&page="+n+"&onservices=off&corp_id="+data.getString("corp_id")+
											(request.getParameter("select")!=null?"&select="+request.getParameter("select"):"")%>">
										停用
										</a>
									<%}else{%>
										<a href="Admin?status=list
											<%="&page="+n+"&onservices=off&corp_id="+data.getString("corp_id")+
											(request.getParameter("select")!=null?"&select="+request.getParameter("select"):"")%>">
										<%
										if(data.getString("stop_date").split("-") != null){
											String tmp_data[] = data.getString("stop_date").split("-");
											if(Integer.parseInt(tmp_data[1]) < 10){
												tmp_data[1] = "0" + tmp_data[1];
											}
											if(Integer.parseInt(tmp_data[2]) < 10){
												tmp_data[2] = "0" + tmp_data[2];
											}
											out.println(tmp_data[0]+tmp_data[1]+tmp_data[2]);
										}
										%>
										</a>
									<%}%>
								<%}%>
							</td>
						</tr>
						<%}%>
					</table>
				<%}%>
				<table width="100%">
					<td width="80%"></td>
					<%if(request.getParameter("select") == null){%>
						<td width="10%" align="center">
							<a href="Admin?status=list&page=<%=n2%>">上一頁</a>
						</td>
						<td width="10%" align="center">
							<%if(!(data_cont < 100)){%>
							<a href="Admin?status=list&page=<%=n1%>">下一頁</a>
							<%}%>
						</td>
					<%}else{%>
						<td width="10%" align="center">
							<a href="Admin?status=list&select=<%=request.getParameter("select")%>&page=<%=n2%>">上一頁</a>
						</td>
						<td width="10%" align="center">
							<%if(!(data_cont < 100)){%>
							<a href="Admin?status=list&select=<%=request.getParameter("select")%>&page=<%=n1%>">下一頁</a>
							<%}%>
						</td>
					<%}%>
				</table>
			<%}%>
			<%if(request.getParameter("status") != null && request.getParameter("status").equals("edit")){%>
				<form action="Admin" method="get">
					<table width="350" border="1" align="center">
						<input type="hidden" name="onservices" value="on">
						<input type="hidden" name="select" value="<%=request.getParameter("corp_idno") %>">
						<tr>
							<td align="center" colspan="2"><b>啟用公司</b></td>
						</tr>
						<tr>
							<td width="150" align="right">公司中文名稱</td>
							<td><input name="corp_name" type="text" /></td>
						</tr>
						<tr>
							<td width="150" align="right">公司編號</td>
							<td><input name="corp_id" type="text" value='<%=request.getParameter("corp_id") %>'/></td>
						</tr>
						<tr>
							<td width="150" align="right">公司統編</td>
							<td><input name="corp_idno" type="text" value='<%=request.getParameter("corp_idno") %>'/></td>
						</tr>
						<tr>
							<td width="150" align="right">管理人Email</td>
							<td><input name="user_id" type="text" /></td>
						</tr>
						<tr>
							<td width="150" align="right">管理人密碼</td>
							<td><input name="user_pwd" type="text" id="user_pwd" value='admin'/></td>
						</tr>
						<tr>
							<td align="right" colspan="2"><input name="submit" type="submit" value="新增" /></td>
						</tr>
					</table>
				</form>
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
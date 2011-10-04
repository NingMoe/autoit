<%@page import="java.util.HashMap" %>
<%@page import="java.util.ArrayList" %>
<%@page import="java.sql.ResultSet" %>

<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
    
<%
HashMap AirMap = (HashMap)this.getServletContext().getAttribute("AirMap");
String[] AirName = {
		"中國地區","港澳及台灣地區","美洲地區 ","歐洲地區",
		"東北亞地區","東南亞地區 ","大洋洲地區","中南美洲地區 ","中東及非洲地區"};
//
//---決行者才有的資訊---
ArrayList UserBudget;
int UserBudget_totle = -1;
int UserBudget_grant = -1;
if(request.getAttribute("UserBudget") != null){
	UserBudget = (ArrayList)request.getAttribute("UserBudget");
	UserBudget_totle = 
		Integer.parseInt((String)(((HashMap)UserBudget.get(UserBudget.size()-1)).get("budget_total")));
	UserBudget_grant = 
		Integer.parseInt((String)(((HashMap)UserBudget.get(UserBudget.size()-1)).get("budget_grant")));
}
//---------------------
//
ArrayList trip_flow = (ArrayList)request.getAttribute("trip_flow");
ArrayList trip_flow_tripGroups = (ArrayList)request.getAttribute("trip_flow_tripGroups");
ArrayList trip_flow_userGroups = (ArrayList)request.getAttribute("trip_flow_userGroups"); 

ResultSet trip_list = (ResultSet)request.getAttribute("trip_list");

HashMap edit_trip_list = null;
ArrayList edit_tripRoute_list = null;
ArrayList edit_tripExpense_list = null;
ArrayList edit_tripFlow_list = null;
ArrayList edit_tripFlowUser_list = null;
if(request.getAttribute("edit_trip_list") != null){
	edit_tripFlow_list = (ArrayList)request.getAttribute("edit_tripFlow_list");
	edit_tripFlowUser_list = (ArrayList)request.getAttribute("edit_tripFlowUser_list");
	edit_trip_list = (HashMap)request.getAttribute("edit_trip_list");
	edit_tripRoute_list = (ArrayList)request.getAttribute("edit_tripRoute_list");
	edit_tripExpense_list = (ArrayList)request.getAttribute("edit_tripExpense_list");
}

HashMap show_trip_list = null;
ArrayList show_tripRoute_list = null;
ArrayList show_tripExpense_list = null;
ArrayList show_tripFlow_list = null;
ArrayList show_tripFlowUser_list = null;
if(request.getAttribute("show_trip_list") != null){
	show_trip_list = (HashMap)request.getAttribute("show_trip_list");
	show_tripRoute_list = (ArrayList)request.getAttribute("show_tripRoute_list");
	show_tripExpense_list = (ArrayList)request.getAttribute("show_tripExpense_list");
	edit_tripRoute_list = (ArrayList)request.getAttribute("show_tripRoute_list");
	edit_tripExpense_list = (ArrayList)request.getAttribute("show_tripExpense_list");
	show_tripFlow_list = (ArrayList)request.getAttribute("show_tripFlow_list");
	show_tripFlowUser_list = (ArrayList)request.getAttribute("show_tripFlowUser_list");
}

HashMap actual_trip_list = null;
ArrayList actual_tripRoute_list = null;
ArrayList actual_tripExpense_list = null;
ArrayList actual_tripFlow_list = null;
ArrayList actual_tripFlowUser_list = null;
ArrayList actual_TripActual_Expense_list = null;
if(request.getAttribute("actual_tripExpense_list") != null){
	actual_trip_list = (HashMap)request.getAttribute("actual_trip_list");
	actual_tripRoute_list = (ArrayList)request.getAttribute("actual_tripRoute_list");
	actual_tripExpense_list = (ArrayList)request.getAttribute("actual_tripExpense_list");
	actual_tripFlow_list = (ArrayList)request.getAttribute("actual_tripFlow_list");
	actual_tripFlowUser_list = (ArrayList)request.getAttribute("actual_tripFlowUser_list");
	actual_TripActual_Expense_list = (ArrayList)request.getAttribute("actual_TripActual_Expense_list");
}

HashMap user_date = (HashMap)request.getAttribute("user_date");
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>OnePage</title>
	<link rel="stylesheet" type="text/css" href="css/css.css"/>
	
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
	<style type="text/css">
	.DynarchCalendar-topBar{
		text-align:center;
	}
	</style>
	<script>
	function isValue(v){
		if(v == 'NaN' || v ==  null || v == '')
			return 0;
		else
			return v;
	}
	function sum_estim_f(n){
		var e = document.forms['estim'];
		e.elements['estim_sum'][n].value=
			parseInt(isValue(e.elements['estim_hotel'][n].value) ,10)+
			parseInt(isValue(e.elements['estim_traffic'][n].value) ,10)+
			parseInt(isValue(e.elements['estim_meal'][n].value) ,10)+
			parseInt(isValue(e.elements['estim_extra'][n].value) ,10)+
			parseInt(isValue(e.elements['estim_other'][n].value) ,10);
		if(e.elements['estim_sum'][n].value == 0)
			e.elements['estim_sum'][n].value = '';
	}
	function sum_sum_estim_totle_f(){
		var e = document.forms['estim'];
		e.elements['sum_estim_totle'].value=
			parseInt(isValue(e.elements['sum_estim_hotel'].value),10)+
			parseInt(isValue(e.elements['sum_estim_traffic'].value),10)+
			parseInt(isValue(e.elements['sum_estim_meal'].value),10)+
			parseInt(isValue(e.elements['sum_estim_extra'].value),10)+
			parseInt(isValue(e.elements['sum_estim_other'].value),10);
		if(e.elements['sum_estim_totle'].value == 0)
			e.elements['sum_estim_totle'].value = '';
	}
	function sum_estim(){
		var e = document.forms['estim'];
		var name=new Array(
				'sum_estim_duration' , 'sum_estim_hotel', 
				'sum_estim_traffic' , 'sum_estim_meal', 
				'sum_estim_extra', 'sum_estim_other');
		var name2=new Array(
				'estim_duration' , 'estim_hotel', 
				'estim_traffic' , 'estim_meal', 
				'estim_extra', 'estim_other');
		for(var i = 0 ; i < name.length ; i++){
			e.elements[name[i]].value=0;

			if(i == 0){
				for(var j = 0 ; j < e.elements[name2[0]].length ; j++){
					e.elements[name[0]].value=
						parseInt(isValue(e.elements[name[0]].value),10)+
						parseInt(isValue(e.elements[name2[0]][j].value),10);
					
				}
			}else{
				for(var j = 0 ; j < e.elements[name2[i]].length ; j++){
					e.elements[name[i]].value=
						parseInt(isValue(e.elements[name[i]].value),10)+
						(parseInt(isValue(e.elements[name2[i]][j].value),10)*parseInt(isValue(e.elements[name2[0]][j].value)));
				}
			}
			if(e.elements[name[i]].value == 0)
				e.elements[name[i]].value = '';
			sum_estim_f(i);
			sum_sum_estim_totle_f();
		}
	}
	
	
	function actual_totle(n,m){
		var e = document.forms['actual'];
		//totle
		//e.elements['actual_estim_city_totle'].value = e.elements['actual_estim_city_h'][n].value;
		e.elements['actual_estim_hotel_totle'].value = e.elements['actual_trip_hotel_estim_h'].value;
		e.elements['actual_estim_traffic_totle'].value = e.elements['actual_trip_traffic_estim_h'].value;
		e.elements['actual_estim_meal_totle'].value = e.elements['actual_trip_meal_estim_h'].value;
		e.elements['actual_estim_extra_totle'].value = e.elements['actual_trip_extra_estim_h'].value;
		e.elements['actual_estim_other_totle'].value = e.elements['actual_trip_other_estim_h'].value;
		e.elements['actual_estim_sum_totle'].value = e.elements['actual_trip_expense_estimate_h'].value;
		
		//全部差異比
		var actual_estim_hotel_avg = 0;
		var actual_estim_traffic_avg = 0;
		var actual_estim_meal_avg = 0;
		var actual_estim_extra_avg = 0;
		var actual_estim_other_avg = 0;
		var actual_estim_sum_avg = 0;

		for(var i = 0 ; i < e.elements['actual_expense_sum'].length ; i++){
			e.elements['actual_expense_sum'][i].value = 
				parseInt(isValue(e.elements['actual_estim_hotel'][i].value))
				+ parseInt(isValue(e.elements['actual_expense_traffic'][i].value))
				+ parseInt(isValue(e.elements['actual_expense_meal'][i].value))
				+ parseInt(isValue(e.elements['actual_expense_extra'][i].value))
				+ parseInt(isValue(e.elements['actual_expense_other'][i].value)); 
		}
		
		for(var i = 0 ; i < e.elements['actual_estim_hotel'].length ; i++){
			actual_estim_hotel_avg
				+= parseInt(isValue(e.elements['actual_estim_hotel'][i].value));
			actual_estim_traffic_avg
				+= parseInt(isValue(e.elements['actual_expense_traffic'][i].value));
			actual_estim_meal_avg
				+= parseInt(isValue(e.elements['actual_expense_meal'][i].value));
			actual_estim_extra_avg
				+= parseInt(isValue(e.elements['actual_expense_extra'][i].value));
			actual_estim_other_avg
				+= parseInt(isValue(e.elements['actual_expense_other'][i].value));
			actual_estim_sum_avg
				+= parseInt(isValue(e.elements['actual_expense_sum'][i].value));
		}
		
		actual_estim_hotel_avg
			= actual_estim_hotel_avg/parseInt(isValue(e.elements['actual_trip_hotel_estim_h'].value))*100;
		actual_estim_traffic_avg
			= actual_estim_traffic_avg/parseInt(isValue(e.elements['actual_trip_traffic_estim_h'].value))*100;
		actual_estim_meal_avg
			= actual_estim_meal_avg/parseInt(isValue(e.elements['actual_trip_meal_estim_h'].value))*100;
		actual_estim_extra_avg
			= actual_estim_extra_avg/parseInt(isValue(e.elements['actual_trip_extra_estim_h'].value))*100;
		actual_estim_other_avg
			= actual_estim_other_avg/parseInt(isValue(e.elements['actual_trip_other_estim_h'].value))*100;
		actual_estim_sum_avg
			= actual_estim_sum_avg/parseInt(isValue(e.elements['actual_trip_expense_estimate_h'].value))*100;
			
		actual_estim_hotel_avg = actual_estim_hotel_avg.toFixed(2);//.toFixed(2)取小數兩位
		actual_estim_traffic_avg = actual_estim_traffic_avg.toFixed(2);
		actual_estim_meal_avg = actual_estim_meal_avg.toFixed(2);
		actual_estim_extra_avg = actual_estim_extra_avg.toFixed(2);
		actual_estim_other_avg = actual_estim_other_avg.toFixed(2);
		actual_estim_sum_avg = actual_estim_sum_avg.toFixed(2);
		
		//e.elements['actual_estim_city_avg'].value = e.elements['actual_estim_city_h'][n].value;
		e.elements['actual_estim_hotel_avg'].value = actual_estim_hotel_avg+"%";
		e.elements['actual_estim_traffic_avg'].value = actual_estim_traffic_avg+"%";
		e.elements['actual_estim_meal_avg'].value = actual_estim_meal_avg+"%";
		e.elements['actual_estim_extra_avg'].value = actual_estim_extra_avg+"%";
		e.elements['actual_estim_other_avg'].value = actual_estim_other_avg+"%";
		e.elements['actual_estim_sum_avg'].value = actual_estim_sum_avg+"%";

		//處理票據張數
		var num = 0;
		var actual_receipt_totle = 0;
		for(var i = 0 ; i < e.elements['actual_expense_receipt'].length ; i++){
			actual_receipt_totle 
				+= parseInt(isValue(e.elements['actual_expense_receipt'][i].value),10);
		}
		for(var i = 0 ; i < 2 ; i++){
			//前兩個張數是全部總和
			e.elements['actual_expense_receipt_totle'][i].value = actual_receipt_totle;
		}
		for(var i = 0 ; i < e.elements['actual_duration_h'].length ; i++){
			//處理接下來每個行程段的分別張數
			actual_receipt_totle = 0;
			for(var j = num ; j < num+parseInt(isValue(e.elements['actual_duration_h'][n].value),10) ; j++){
				actual_receipt_totle 
					+= parseInt(isValue(e.elements['actual_expense_receipt'][j].value),10);
			}
			num += parseInt(isValue(e.elements['actual_duration_h'][i].value),10);
			e.elements['actual_expense_receipt_totle'][i+2].value = actual_receipt_totle;
		}
	}
	</script>
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
			
			<%if(trip_flow != null){ %>
				<table width="100%">
					<%if(session.getAttribute("user_authority").equals("2")){ %>
					<tr>
						<td width="150" align="center"><strong><font color="#cc0000">決行區</font></strong></td>
						<td width="100" align="center"><strong><font color="#cc0000">申請者</font></strong></td>
						<td width="100" align="center"><strong><font color="#cc0000">申請日期</font></strong></td>
						<td width="320" align="center"><strong><font color="#cc0000">說明</font></strong></td>
						<td width="50" align="center"><strong><font color="#cc0000">天數</font></strong></td>
						<td width="100" align="center"><strong><font color="#cc0000">預估費用</font></strong></td>
						<td width="110" align="center"><strong><font color="#cc0000">狀態</font></strong></td>
					</tr>
					<%}else{%>
					<tr>
						<td width="150" align="center"><strong><font color="#000066">核准區</font></strong></td>
						<td width="100" align="center"><strong><font color="#000066">申請者</font></strong></td>
						<td width="100" align="center"><strong><font color="#000066">申請日期</font></strong></td>
						
						<%
						if(request.getParameter("islist") != null){
							//非 已批核 已決行 則 continue
							//if(request.getParameter("islist").equals("ok")){
							%>
								<td width="320" align="center"><strong><font color="#000066">說明</font></strong></td>
								<td width="50" align="center"><strong><font color="#000066">天數</font></strong></td>
								<td width="100" align="center"><strong><font color="#000066">預估費用</font></strong></td>
								<td width="110" align="center"><strong><font color="#000066">行程狀態</font></strong></td>
							<%
							//}
						}else{
						%>
							<td width="220" align="center"><strong><font color="#000066">說明</font></strong></td>
							<td width="100" align="center"><strong><font color="#000066">送批核決行</font></strong></td>
							<td width="50" align="center"><strong><font color="#000066">天數</font></strong></td>
							<td width="100" align="center"><strong><font color="#000066">預估費用</font></strong></td>
							<td width="110" align="center"><strong><font color="#000066">狀態</font></strong></td>
						<%
						}
						%>
					</tr>
					<%} %>
					<%
					boolean isTrip_flow[] = new boolean[trip_flow.size()];
					for(int i = 0 ; i < trip_flow.size() ; i++){
						HashMap map = (HashMap)trip_flow.get(i); 
						isTrip_flow[i] = true;
						for(int j = i+1 ; j < trip_flow.size() ; j++){
							HashMap map2 = (HashMap)trip_flow.get(j);
							if(map.get("trip_nbr").equals(map2.get("trip_nbr"))){
								isTrip_flow[i] = false;
								break;
							}
						}
					}
					%>
					<%for(int i = 0 ; i < trip_flow.size() ; i++){ %>
					<%
					HashMap map = (HashMap)trip_flow.get(i); 
					HashMap trmap = (HashMap)trip_flow_tripGroups.get(i);
					HashMap usermap = (HashMap)trip_flow_userGroups.get(i);
					
					if(isTrip_flow[i] == false)
						continue;
					
					if(request.getParameter("islist") == null){
						//非 待批核 待決行 則 continue
						if(!(map.get("tripflow_status").equals("1") || map.get("tripflow_status").equals("2"))){
							continue;
						}
					}else{
						if(request.getParameter("islist") != null){
							if(request.getParameter("islist").equals("ok")){
								//非 已批核 已決行 則 continue
								if(!(map.get("tripflow_status").equals("3") 
										|| map.get("tripflow_status").equals("4"))){
									
									//if(trmap.get("trip_status").equals("6")){
										continue;
									//}
								}
							}else if(request.getParameter("islist").equals("no")){
								//非 退回 則 continue
								if(!(map.get("tripflow_status").equals("5"))){
									//if(!trmap.get("trip_status").equals("6")){
										continue;
									//}
								}
								//if(trmap.get("trip_status").equals("6")){
								//	continue;
								//}
							}
						}
					}
					%>
						<tr>
							</td><td align="center">
								<table width="100%">
								<tr>
								<td width="33%">
									<form id="sub_form_list" name="sub_form_list" method="get" action="User">
										<input name="isStatus" type="hidden" value="show" />
										<input name="Trip" type="hidden" value="<%=map.get("trip_nbr") %>" />
										<input name="user_nbr" type="hidden" value="<%=map.get("request_nbr") %>" />
										<input name="trip_flow_nbr" type="hidden" value="<%=i %>" />
										
										<input type="submit" value="詳細" name="submit">
									</form>
								</td>
								<td width="33%">
									<form id="form_list" name="form_list" method="post" action="User">
									<input name="isAction" type="hidden" value="approved" />
									<input name="upTripFlow" type="hidden" id="upTripFlow" value="<%=map.get("tripflow_nbr") %>" />
									<input name="upTrip" type="hidden" id="upTripFlow" value="<%=map.get("trip_nbr") %>" />
									<input name="request_nbr" type="hidden" id="request_nbr" value="<%=map.get("request_nbr") %>" />
									<%if(session.getAttribute("user_authority").equals("2")){ %>
										<input name="isStatus" type="hidden" id="isStatus" value="1" />
										<input name="upStatus" type="hidden" id="upTripFlow" value="4" />
									<%}else{ %>
										<input name="isStatus" type="hidden" id="isStatus" value="1" />
										<input name="upStatus" type="hidden" id="upTripFlow" value="3" />
									<%} %> 
									<%if(session.getAttribute("user_authority").equals("2")){ %>
										<%if(map.get("tripflow_status") != null && map.get("tripflow_status").equals("2")){ %>
											<input type="submit" value="決行" name="submit">
											</td>
											<td width="33%">
											<input type="submit" value="退回" name="submit">
											<!--  
											<a href="User?
												<%="isStatus=1&upTripFlow="+map.get("tripflow_nbr")+
												"&upTrip="+map.get("trip_nbr")+"&upStatus=3" %>">批核</a> 
											<a href="User?
												<%="isStatus=1&upTripFlow="+map.get("tripflow_nbr")+
												"&upTrip="+map.get("trip_nbr")+"&upStatus=5" %>">退回</a>
											-->
										<%}else{ %>
											決行
											</td>
											<td width="33%">
											退回
										<%} %>	
									<%}else{ %>
										<%if(map.get("tripflow_status") != null && map.get("tripflow_status").equals("1")){ %>
											<input type="submit" value="批核" name="submit">
											</td>
											<td width="33%">
											<input type="submit" value="退回" name="submit">
											<!--  
											<a href="User?
												<%="isStatus=1&upTripFlow="+map.get("tripflow_nbr")+
												"&upTrip="+map.get("trip_nbr")+"&upStatus=3" %>">批核</a> 
											<a href="User?
												<%="isStatus=1&upTripFlow="+map.get("tripflow_nbr")+
												"&upTrip="+map.get("trip_nbr")+"&upStatus=5" %>">退回</a>
											-->
										<%}else{ %>
											批核
											</td>
											<td width="33%">
											退回
										<%} %>	
									<%} %>
								</td>
								</tr>
								</table>
							</td>
							<td><%=usermap.get("cname")!=null ? usermap.get("cname") :"" %></td>
							<td align="center">
							<%
							if(map.get("tripflow_time_apply") != null){
							
								//申請日期
								String tmp_data[] = ((String)map.get("tripflow_time_apply")).split("-");
								if(Integer.parseInt(tmp_data[1]) < 10){
									tmp_data[1] = "0" + tmp_data[1];
								}
								if(Integer.parseInt(tmp_data[2]) < 10){
									tmp_data[2] = "0" + tmp_data[2];
								}
								out.println(tmp_data[0]+"/"+tmp_data[1]+"/"+tmp_data[2]);
							}
							%>
							</td>
							<td><%=trmap.get("trip_description") %></td>
							
							<%if(request.getParameter("islist") == null){ %>
								<%if(!session.getAttribute("user_authority").equals("2")){ %>
								<td style="padding-right:5px;">
									<input type="text" style="text-align: left; width: 99%;" 
									id="nextUser" value="<%=user_date.get("approver_mail") %>" name="nextUser">
								</td>
								<%} %>
							<%} %>
							
							<td align="center">
								<%
									if(trmap.get("trip_interval").equals(""))
										out.print("0 天"); 
									else
										out.print(trmap.get("trip_interval")+" 天"); 
								%> 
							</td>
							<td align="right">
								$<% 
								out.println(
									(trmap.get("expense_estimate")!=null && !trmap.get("expense_estimate").equals("")?
											Integer.parseInt((String)trmap.get("expense_estimate")): 0) 
									+ (trmap.get("tkt_estim")!=null && !trmap.get("tkt_estim").equals("")?
											Integer.parseInt((String)trmap.get("tkt_estim")) : 0)); 
								%>
							</td>
							<td align="center">
							<%
							//狀態--待批核:1待決行:2核准:3決行:4取消:5
							
							if(request.getParameter("islist") != null){
								switch(Integer.parseInt((String)trmap.get("trip_status"))){
								case 0:
									out.println("申請中");
									break;
								case 1:
									out.println("申請中");
									break;
								case 2:
									out.println("批核中");
									break;
								case 3:
									out.println("已決行");
									break;
								case 4:
									out.println("出差中");
									break;
								case 5:
									out.println("結案");
									break;
								case 6:
									out.println("被退回");
									break;
								}
							}else{
								switch(Integer.parseInt((String)map.get("tripflow_status"))){
								case 0:
									out.println("被取消");
									break;
								case 1:
									out.println("待批核");
									break;
								case 2:
									out.println("待決行");
									break;
								case 3:
									out.println("已核准");
									break;
								case 4:
									out.println("已決行");
									break;
								case 5:
									out.println("已退回");
									break;
								}
							}
							%>
							</td>
						</tr>
					</form>
					<%} %>
					<tr>
						<td></td>
						<td></td>
						<td colspan="6">
							<a href="User?islist=ok">
							<%
							if(session.getAttribute("user_authority").equals("2")){
								for(int i = 0 , cont = 0; i < trip_flow.size() ; i++) {
									HashMap map = (HashMap)trip_flow.get(i);
									if(isTrip_flow[i] == false)
										continue;
									if(map.get("tripflow_status").equals("4")){
										cont++;
									}
									if(i == trip_flow.size()-1)
										out.println("[已決行 "+cont+" 件]");
								}
							}else{
								for(int i = 0 , cont = 0; i < trip_flow.size() ; i++) {
									HashMap map = (HashMap)trip_flow.get(i);
									if(isTrip_flow[i] == false)
										continue;
									if(map.get("tripflow_status").equals("3")){
										cont++;
									}
									if(i == trip_flow.size()-1)
										out.println("[已核准 "+cont+" 件]");
								}
							}
							%>
							</a>　<a href="User?islist=no">
							<%
							for(int i = 0 , cont = 0; i < trip_flow.size() ; i++) {
								HashMap map = (HashMap)trip_flow.get(i);
								if(isTrip_flow[i] == false)
									continue;
								if(map.get("tripflow_status").equals("5")){
									cont++;
								}
								if(i == trip_flow.size()-1)
									out.println("[已退回 "+cont+" 件]");
							}
							%>
							</a>
						</td>
					</tr>
				</table>
				<br>
				
				<%
				if(request.getParameter("isStatus") != null && request.getParameter("isStatus").equals("show")){
				//顯示詳細資訊
				HashMap trmap = (HashMap)trip_flow.get(Integer.parseInt(request.getParameter("trip_flow_nbr"))); 
				HashMap trgmap = (HashMap)trip_flow_tripGroups.get(Integer.parseInt(request.getParameter("trip_flow_nbr")));
				HashMap usermap = (HashMap)trip_flow_userGroups.get(Integer.parseInt(request.getParameter("trip_flow_nbr")));
				%>
					
					<table width="100%">
						<tr>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
							<td width="5%" height="0"></td>
						</tr>
						<tr>
							
						<tr>
							<td colspan="3" align="center">
							<b><font color="#000066">出差詳細資訊</font></b></td>
							<td colspan="3">
								<%=usermap.get("cname") %>
							</td>
							<td colspan="14" align="left">
								<%=show_trip_list.get("trip_description") %>
							</td>
						</tr>
						<tr>
							<td rowspan="4">機票資訊</td>
							<td rowspan="4" colspan="12">
								<div style="overflow:auto; height:100px; margin:0px; padding:0px; background:#999;">
								<table>
									<tr>
										<td width="15%">日期</td>
										<td width="15%">出發機場</td>
										<td width="15%">落地機場</td>
										<td width="10%">航班</td>
										<td width="5%" align="center">轉機</td>
									</tr>
									<%
									if(show_tripRoute_list != null){
										for(int i = 0 ; i < show_tripRoute_list.size() ; i++){
											HashMap map = (HashMap)show_tripRoute_list.get(i);
									%>
										<tr>
											<td width="15%" style="padding-right:5px;">
												<%=map.get("date_rout_dep") %>
											</td>
											<td width="15%" style="padding-right:5px; text-align:left;">
												<%=map.get("rout_dep_port") %>
											</td>
											<td width="15%" style="padding-right:5px; text-align:left;">
												<%=map.get("rout_ariv_port") %>
											</td>
											<td width="10%" style="padding-right:5px; text-align:left;">
												<%=map.get("rout_flight") %>
											</td>
											<td width="5%" align="center">
												<%
												if(map.get("rout_transfer").equals("y"))
													out.print("yes");
												else
													out.print("no");
												%>
											</td>
										</tr>
									<%
										}
									}
									%>
								</table>
								</div>
							</td>
							<td colspan="2" align="right">PNR:</td>
							<td colspan="2" style="padding-right:5px;">
								<%=show_trip_list.get("trip_pnr") %>
							</td>
							<%if(show_trip_list.get("trip_status").equals("2") ){ %>
							<td colspan="3">部門預算餘額 </td>
							<%}else{ %>
							<td colspan="3" ROWSPAN=4 ></td>
							<%} %>
						</tr>
						<tr>
							<td colspan="2" align="right">票價預估:</td>
							<td colspan="2" style="padding-right:5px;">
								<%=show_trip_list.get("tkt_estim") %>
							</td>
							<%if(show_trip_list.get("trip_status").equals("2") ){ %>
							<td colspan="3" align="right">
								$<%=UserBudget_totle-UserBudget_grant %>
							</td>
							<%} %>
						</tr>
						<tr>
							
							<td colspan="2" align="right">出差天數:</td>
							<td colspan="2" style="padding-right:5px;">
								<%=show_trip_list.get("trip_interval") %>
							</td>
							<%if(show_trip_list.get("trip_status").equals("2") ){ %>
							<td colspan="3">扣除本次費用</td>
							<%} %>
						</tr>
						<tr>
							<td colspan="2" align="right">總預算:</td>
							<td colspan="2">
							<%
							int tkt_estim_tmp = 0;;
							if(show_trip_list.get("tkt_estim") != null 
									&& !show_trip_list.get("tkt_estim").equals("")
									&& show_trip_list.get("expense_estimate") != null 
									&& !show_trip_list.get("expense_estimate").equals("")){
								int i = Integer.parseInt((String)show_trip_list.get("expense_estimate"));
								int j = Integer.parseInt((String)show_trip_list.get("tkt_estim"));
								tkt_estim_tmp = i+j;
								out.println("$"+(i+j));
							}
							%>
							</td>
							<%if(show_trip_list.get("trip_status").equals("2") ){ %>
							<td colspan="3" align="right">
							$<%=(UserBudget_totle-UserBudget_grant-tkt_estim_tmp) %>
							</td>
							<%} %>
						</tr>
						<tr>
							<td colspan="1">預算</td>
							<td colspan="2" align="center">城市</td>
							<td colspan="1" align="center">天數</td>
							<td colspan="2" align="center">住宿</td>
							<td colspan="2" align="center">交通</td>
							<td colspan="2" align="center">膳食</td>
							<td colspan="2" align="center">雜項</td>
							<td colspan="2" align="center">其他&票價</td>
							<td colspan="2" align="center">總計</td>
							<td colspan="4" align="center">備註</td>
						</tr>
						<tr>
							<td colspan="1">小計</td>
							<td colspan="2"></td>
							<td colspan="1" align="right" style="padding-right:5px; text-align:right; ">
								<%
								if(show_tripExpense_list != null){
									int sum = 0;
									for(int i = 0 ; i < show_tripExpense_list.size() ; i++){ 
										HashMap map = (HashMap)show_tripExpense_list.get(i);
										sum += Integer.parseInt((String)map.get("duration"));
									}
									out.print(sum);
								}
								%>
							</td>
							<td colspan="2" style="padding-right:5px; text-align:right; ">
								<%=show_trip_list.get("hotel_estim") %>
							</td>
							<td colspan="2" style="padding-right:5px; text-align:right; ">
								<%=show_trip_list.get("traffic_estim") %>
							</td>
							<td colspan="2" style="padding-right:5px; text-align:right; ">
								<%=show_trip_list.get("meal_estim") %>
							</td>
							<td colspan="2" style="padding-right:5px; text-align:right; ">
								<%=show_trip_list.get("extra_estim") %>
							</td>
							<td colspan="2" style="padding-right:5px; text-align:right; ">
								<%=show_trip_list.get("other_estim") %>
							</td>
							<td colspan="2" align="right" style="padding-right:5px; text-align:right; ">
								<%=show_trip_list.get("expense_estimate") %>
							</td>
							<td colspan="4" style="padding-right:5px;">
							</td>
						</tr>
						
						<%for(int i = 0 ; i < show_tripExpense_list.size() ; i++){ %>
						<%HashMap map = (HashMap)show_tripExpense_list.get(i); %>
						<tr>
							<td colspan="1" align="center"><%=i+1 %></td>
							<td colspan="2">
								<%
								for(int j = 0 ; j < AirName.length ; j++){
									ArrayList arr = (ArrayList)AirMap.get(AirName[j]);
									if(arr != null){
										for(int k = 0 ; k < arr.size() ; k++){
											String val[] = ((String)arr.get(k)).split(",");
											//val[0]=中文  val[0]=英文   val[0]=縮寫 
											if(map.get("expense_city") != null && map.get("expense_city").equals(val[2])){
												out.println(val[0]);
												break;
											}
										}
									}
								}
								%>
							</td>
							<td colspan="1" align="center" style="padding-right:5px; text-align:right; ">
								<%=map.get("duration") %>
							</td>
							<td colspan="2" style="padding-right:5px; text-align:right; ">
								<%=map.get("estim_hotel") %>
							</td>
							<td colspan="2" style="padding-right:5px; text-align:right; ">
								<%=map.get("estim_traffic") %>
							</td>
							<td colspan="2" style="padding-right:5px; text-align:right; ">
								<%=map.get("estim_meal") %>
							</td>
							<td colspan="2" style="padding-right:5px; text-align:right; ">
								<%=map.get("estim_extra") %>
							</td>
							<td colspan="2" style="padding-right:5px; text-align:right; ">
								<%=map.get("estim_other") %>
							</td>
							<td colspan="2" align="right" name="estim_sum" style="padding-right:5px; text-align:right; ">
								<% 
								out.println(
										Integer.parseInt((String)map.get("estim_hotel"))+
										Integer.parseInt((String)map.get("estim_traffic"))+
										Integer.parseInt((String)map.get("estim_meal"))+
										Integer.parseInt((String)map.get("estim_extra"))+
										Integer.parseInt((String)map.get("estim_other")));
								%>
							</td>
							<td colspan="4" style="padding-right:5px;">
								<%=map.get("expense_note") %>
							</td>
						</tr>
						<%} %>
					</table>
					<table width="100%">
						<tr>
							<td width="10%" align="center">申請流程</td>
							<td width="10%" align="center">立案</td>
							<td width="10%" align="center">申請</td>
							<td width="10%" align="center">批核</td>
							<td width="10%" align="center">決行</td>
							<td width="10%" align="center">報帳</td>
							<td width="10%" align="center"></td>
							<td width="10%" align="center"></td>
							<td width="10%" align="center"></td>
							<td width="10%" align="center">取消</td>
						</tr>
						<tr>
							<td colspan="1" align="center">日期</td>
							<td colspan="1">
							<%
							//建立時間
							if(show_trip_list.get("time_create") != null 
							&& !((String)show_trip_list.get("time_create")).equals("")){
								String tmp_data[] = ((String)show_trip_list.get("time_create")).split("-");
								if(Integer.parseInt(tmp_data[1]) < 10){
									tmp_data[1] = "0" + tmp_data[1];
								}
								if(Integer.parseInt(tmp_data[2]) < 10){
									tmp_data[2] = "0" + tmp_data[2];
								}
								out.println(tmp_data[0]+"/"+tmp_data[1]+"/"+tmp_data[2]);
							}else
								out.println("尚無記錄");
							%>
							</td>
							<td colspan="1">
							<%
							//申請日期
							if(show_trip_list.get("time_apply") != null 
							&& !((String)show_trip_list.get("time_apply")).equals("")){
								String tmp_data[] = ((String)show_trip_list.get("time_apply")).split("-");
								if(Integer.parseInt(tmp_data[1]) < 10){
									tmp_data[1] = "0" + tmp_data[1];
								}
								if(Integer.parseInt(tmp_data[2]) < 10){
									tmp_data[2] = "0" + tmp_data[2];
								}
								out.println(tmp_data[0]+"/"+tmp_data[1]+"/"+tmp_data[2]);
							}else
								out.println("尚無記錄");
							%>
							</td>
							<td colspan="1">
							<%
							//核准時間
							if(show_trip_list.get("time_agree") != null 
							&& !((String)show_trip_list.get("time_agree")).equals("")){
								String tmp_data[] = ((String)show_trip_list.get("time_agree")).split("-");
								if(Integer.parseInt(tmp_data[1]) < 10){
									tmp_data[1] = "0" + tmp_data[1];
								}
								if(Integer.parseInt(tmp_data[2]) < 10){
									tmp_data[2] = "0" + tmp_data[2];
								}
								out.println(tmp_data[0]+"/"+tmp_data[1]+"/"+tmp_data[2]);
							}else
								out.println("尚無記錄");
							%>
							</td>
							<td colspan="1">
							<%
							//准行時間
							if(show_trip_list.get("time_final") != null 
							&& !((String)show_trip_list.get("time_final")).equals("")){
								String tmp_data[] = ((String)show_trip_list.get("time_final")).split("-");
								if(Integer.parseInt(tmp_data[1]) < 10){
									tmp_data[1] = "0" + tmp_data[1];
								}
								if(Integer.parseInt(tmp_data[2]) < 10){
									tmp_data[2] = "0" + tmp_data[2];
								}
								out.println(tmp_data[0]+"/"+tmp_data[1]+"/"+tmp_data[2]);
							}else
								out.println("尚無記錄");
							%>
							</td>
							<td colspan="1">尚無記錄</td>
							<td colspan="1"></td>
							<td colspan="1"></td>
							<td colspan="1"></td>
							<td colspan="1">
							<%
							//取消時間
							if(show_trip_list.get("time_close") != null 
							&& !((String)show_trip_list.get("time_close")).equals("")){
								String tmp_data[] = ((String)show_trip_list.get("time_close")).split("-");
								if(Integer.parseInt(tmp_data[1]) < 10){
									tmp_data[1] = "0" + tmp_data[1];
								}
								if(Integer.parseInt(tmp_data[2]) < 10){
									tmp_data[2] = "0" + tmp_data[2];
								}
								out.println(tmp_data[0]+"/"+tmp_data[1]+"/"+tmp_data[2]);
							}else
								out.println("尚無記錄");
							%>
							</td>
						</tr>
						<tr>
							<td colspan="1" align="center"></td>
							<td colspan="1"></td>
							<td colspan="1" style="padding-right:5px;">
								<%=usermap.get("cname") %><font size=2 color="#227700">提出</font>
							</td>
							<td colspan="1" style="padding-right:5px;" align="center">
								<%
									//列出所有的批何者
									if(show_tripFlowUser_list != null && show_tripFlow_list != null){
										for(int i = 0 ; i < show_tripFlowUser_list.size() ; i++){
											HashMap map = (HashMap)show_tripFlowUser_list.get(i);
											HashMap map2 = (HashMap)show_tripFlow_list.get(i); 
											
											int n = Integer.parseInt((String)map2.get("tripflow_status"));
											if(n == 1){
												out.println(map.get("cname"));
												out.println("<font size=2 color=\"#227700\">待批<br></font>");
											}else if(n == 3){
												out.println(map.get("cname"));
												out.println("<font size=2 color=\"#000066\">已批<br></font>");
											}else if(n == 5){
												out.println(map.get("cname"));
												out.println("<font size=2 color=\"#AA0000\">退回<br></font>");
											}
										}
									}
								 %>
							</td>
							<td colspan="1" style="padding-right:5px;" align="center">
								<%
									//列出所有的准行者
									if(show_tripFlowUser_list != null && show_tripFlow_list != null){
										for(int i = 0 ; i < show_tripFlowUser_list.size() ; i++){
											HashMap map = (HashMap)show_tripFlowUser_list.get(i);
											HashMap map2 = (HashMap)show_tripFlow_list.get(i); 
											int n = Integer.parseInt((String)map2.get("tripflow_status"));
											if(n == 2){
												out.println(map.get("cname"));
												out.println("<font size=2 color=\"#227700\">待准<br></font>");
											}else if(n == 4){
												out.println(map.get("cname"));
												out.println("<font size=2 color=\"#000066\">已准<br></font>");
											}
										}
									}
								 %>
							</td>
							<td colspan="1"></td>
							<td colspan="1"></td>
							<td colspan="1"></td>
							<td colspan="1"></td>
							<td colspan="1">
							<%
							if(show_trip_list.get("time_close") != null 
								&& !((String)show_trip_list.get("time_close")).equals("")){
								
								out.println((String)show_trip_list.get("close_describe"));
							}
							%>
							</td>
						</tr>
						<tr>
							<td align="center">備註</td>
							<td colspan="9" style="padding-right:5px;">
								<%=show_trip_list.get("trip_note") %>
							</td>
						</tr>
						<tr>
							<td align="center">狀態</td>
							<td colspan="4">尚未申請批核 / 等待機票報價</td>
							<form name="estim_down" action="User" method="post">
								<input name="isAction" type="hidden" value="approved" />
								<input name="upTripFlow" type="hidden" id="upTripFlow" value="<%=trmap.get("tripflow_nbr") %>" />
								<input name="upTrip" type="hidden" id="upTripFlow" value="<%=trmap.get("trip_nbr") %>" />
								<input name="request_nbr" type="hidden" id="request_nbr" value="<%=trmap.get("request_nbr") %>" />
								<%if(session.getAttribute("user_authority").equals("2")){ %>
									<input name="isStatus" type="hidden" id="isStatus" value="1" />
									<input name="upStatus" type="hidden" id="upTripFlow" value="4" />
								<%}else{ %>
									<input name="isStatus" type="hidden" id="isStatus" value="1" />
									<input name="upStatus" type="hidden" id="upTripFlow" value="3" />
								<%} %>
								<%
								boolean isViewTD = true;
								for(int i = 0 ; i < trip_flow.size() ; i++){
									HashMap map = (HashMap)trip_flow.get(i);
									String s = 
										Integer.toString(Integer.parseInt((String)show_trip_list.get("trip_nbr")));
									if(map.get("trip_nbr").equals(s)){
										if(map.get("tripflow_status").equals("1") 
												|| map.get("tripflow_status").equals("2")){
											if(session.getAttribute("user_authority").equals("2")){
												%>
												<td></td>
												<td></td>
												<td colspan="1" align="center" style="padding-right:5px;" >
													<input type="submit" value="決行" name="submit">
												</td>
												<td colspan="1" align="center" style="padding-right:5px;" >
													<input type="submit" value="退回" name="submit">
												</td>
												<%
											}else{
												%>
												<td colspan="3" align="center" style="padding-right:5px;" >
													<input type="submit" value="批核" name="submit">
													<input name="nextUser" type="text" value="<%=user_date.get("approver_mail") %>" 
															id="nextUser" style="text-align:left; width:75%;"/>
												</td>
												<td colspan="1" align="center" style="padding-right:5px;" >
													<input type="submit" value="退回" name="submit">
												</td>
												<%
											}
											isViewTD = false;
										}
									}
								}
								if(isViewTD == true){
									%>
									<td colspan="3"></td>
									<td></td>
									<%
								}
								%>
								<td colspan="1" align="center">
									<input name="submit" type="submit" value="返回" />
								</td>
							</form>
						</tr>
					</table>
					<br>
				<%} %>
			<%} %>
			
			<table width="100%">
				<tr>
					<td width="125" align="center"><strong><font color="#006633">員工區</font></strong></td>
					<td width="100" align="center"><strong><font color="#006633">申請者</font></strong></td>
					<td width="100" align="center"><strong><font color="#006633">申請日期</font></strong></td>
					<td width="330" align="center"><strong><font color="#006633">說明</font></strong></td>
					<td width="50" align="center"><strong><font color="#006633">天數</font></strong></td>
					<td width="100" align="center"><strong><font color="#006633">預估費用</font></strong></td>
					<td width="110" colspan="3" align="center"><strong><font color="#006633">狀態</font></strong></td>
				</tr>
				<%if(trip_list != null){%>
					<%while(trip_list.next()){ %>
						<%
						if(request.getParameter("isUserList") != null){
							if(request.getParameter("isUserList").equals("6")){
								if(!trip_list.getString("trip_status").equals("6")){
									continue;//顯示被退回
								}
							}else if(request.getParameter("isUserList").equals("5")){
								if(!trip_list.getString("trip_status").equals("5")){
									continue;//顯示已結案
								}
							}
						}else{
							if(trip_list.getString("trip_status").equals("5")){
								continue;
							}
						}
						%>
						<tr>
							<td align="center">
							<a href="User?isTrip=edit&trop_id=<%=trip_list.getString("trip_nbr")%>">
							<%
							//申請:1批核:2決行:3出差:4結案:5退回:6
							switch(Integer.parseInt(trip_list.getString("trip_status"))){
							case 0:
								out.println("[修改]");
								break;
							case 1:
								out.println("[修改]");
								break;
							case 2:
								out.println("[詳細]");
								break;
							case 3:
								out.println("[詳細]");
								break;
							case 5:
								out.println("[詳細]");
								break;
							case 6:
								out.println("[修改]");
								break;
							}
							%>
							</a>
							<a href="User?isTrip=actual&trip_id=<%=trip_list.getString("trip_nbr")%>">
							<%
							//已決行 才可報帳
							if(trip_list.getString("trip_status").equals("3")){
								out.println("[報帳]");
							}
							%>
							</a>
							</td>
							<td><%=session.getAttribute("user_cname") /*user cname*/ %></td>
							<td align="center">
							<%
							if(Integer.parseInt(trip_list.getString("trip_status")) < 2){
								out.println("未提出申請");
							}else{
								//申請時間
								String tmp_data[] = ((String)trip_list.getString("time_create")).split("-");
								if(Integer.parseInt(tmp_data[1]) < 10){
									tmp_data[1] = "0" + tmp_data[1];
								}
								if(Integer.parseInt(tmp_data[2]) < 10){
									tmp_data[2] = "0" + tmp_data[2];
								}
								out.println(tmp_data[0]+"/"+tmp_data[1]+"/"+tmp_data[2]);
							}
							%>
							</td>
							<td><%=trip_list.getString("trip_description") /*行程目的*/ %></td>
							<td align="center"><%=trip_list.getString("trip_interval") /*天數*/ %>天</td>
							<td align="right">
								$<% 
								out.println(
									(trip_list.getString("expense_estimate")!=null && !trip_list.getString("expense_estimate").equals("")?
											Integer.parseInt(trip_list.getString("expense_estimate")): 0)
									+ (trip_list.getString("tkt_estim")!=null && !trip_list.getString("tkt_estim").equals("")?
											Integer.parseInt(trip_list.getString("tkt_estim")) : 0)); /*總預算 */ 
								%>
							</td><td colspan="3" align="center">
							<%
							//申請:1批核:2決行:3出差:4結案:5退回:6
							switch(Integer.parseInt(trip_list.getString("trip_status"))){
							case 0:
								out.println("申請中");
								break;
							case 1:
								out.println("申請中");
								break;
							case 2:
								out.println("批核中");
								break;
							case 3:
								out.println("已決行");
								break;
							case 4:
								out.println("出差中");
								break;
							case 5:
								out.println("結案");
								break;
							case 6:
								out.println("被退回");
								break;
							}
							%>
							</td>
						</tr>
					<%} %>
				<%}%>
				<tr>
					<td></td>
					<td align="center"><a href="User?isTrip=show"">[申請出差]</a></td>
					<td colspan="8">
						<%
						if(trip_list != null){
							trip_list.beforeFirst();
							int a1 = 0 , a2 = 0; 
							while(trip_list.next()){
								if(trip_list.getString("trip_status").equals("6")){
									// 退回 計數器
									a1++;
								}
								else if(trip_list.getString("trip_status").equals("5")){
									// 以報帳  計數器
									a2++;
								}
							}
						%>
							<a href="User?isUserList=6">[被退回 <%=a1 %> 件]</a>　
							<a href="User?isUserList=5">[已結案 <%=a2 %> 件]</a>
						<%
						}
						%>
					</td>
				</tr>
			</table>
			
			<br>
			
			<!--報帳trip  -->
			<%if(request.getParameter("isTrip") != null && request.getParameter("isTrip").equals("actual")){ %>
			actual
			<%
			int[][] actual = new int[actual_tripExpense_list.size()][];
			String[] city_name = new String[actual_tripExpense_list.size()];
			//
			for(int i = 0 ; i < actual_tripExpense_list.size() ; i++){
				HashMap map = (HashMap)actual_tripExpense_list.get(i);
				actual[i] = new int[Integer.parseInt((String)map.get("duration"))];
				for(int j = 0 ; j < AirName.length ; j++){
					ArrayList arr = (ArrayList)AirMap.get(AirName[j]);
					if(arr != null){
						for(int k = 0 ; k < arr.size() ; k++){
							String val[] = ((String)arr.get(k)).split(",");
							//val[0]=中文  val[0]=英文   val[0]=縮寫 
							if(map.get("expense_city") != null && map.get("expense_city").equals(val[2])){
								city_name[i] = val[0];
								break;
							}
						}
					}
				}
			}
			%>
			<form name="actual" action="User?actual" method="post">
			<input name="isAction" type="hidden" value="actual" />
			<input name="trip_id" type="hidden" value="<%=request.getParameter("trip_id") %>" />
			<input name="actual_trip_expense_estimate_h" id="actual_trip_expense_estimate_h" 
					type="hidden" value="<%=actual_trip_list.get("expense_estimate") %>">
			<input name="actual_trip_hotel_estim_h" id="actual_trip_hotel_estim_h" 
					type="hidden" value="<%=actual_trip_list.get("hotel_estim") %>">
			<input name="actual_trip_traffic_estim_h" id="actual_trip_traffic_estim_h" 
					type="hidden" value="<%=actual_trip_list.get("traffic_estim") %>">
			<input name="actual_trip_meal_estim_h" id="actual_trip_meal_estim_h" 
					type="hidden" value="<%=actual_trip_list.get("meal_estim") %>">
			<input name="actual_trip_extra_estim_h" id="actual_trip_extra_estim_h" 
					type="hidden" value="<%=actual_trip_list.get("extra_estim") %>">
			<input name="actual_trip_other_estim_h" id="actual_trip_other_estim_h" 
					type="hidden" value="<%=actual_trip_list.get("other_estim") %>">
			<table width="100%">
				<tr>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
				</tr>
				<tr>
					<td colspan="1" align="center"><B>報帳</B></td>
					<td colspan="2" align="right"><%=session.getAttribute("user_cname") %></td>
					<td colspan="17">出差目的 : <%=actual_trip_list.get("trip_description") %></td>
				</tr>
				<tr>
					<td colspan="6" align="right"></td>
					<!-- 
					<td colspan="2" align="center">城市</td>
					 -->
					<td colspan="2" align="center">票據張數</td>
					<td colspan="2" align="center">住宿</td>
					<td colspan="2" align="center">交通</td>
					<td colspan="2" align="center">膳食</td>
					<td colspan="2" align="center">雜項</td>
					<td colspan="2" align="center">其他&票價</td>
					<td colspan="2" align="center">總計</td>
				</tr>
				<tr>
					<td colspan="6" align="right">總預算</td>
					<!-- 
					<td colspan="2" align="center">
						<input name="actual_estim_city_totle" id="actual_estim_city_totle" 
							type="text" value=""
							style="text-align:center; width:99%; border:0px;" readonly="readonly"
						/>
					</td>
					 -->
					<td colspan="2" align="right">
						<input name="actual_expense_receipt_totle" id="actual_expense_receipt_totle" 
							type="text" value=""
							style="text-align:right; width:99%; border:0px;" readonly="readonly"
						/>	
					</td>
					<td colspan="2" align="right">
						<input name="actual_estim_hotel_totle" id="actual_estim_hotel_totle" 
							type="text" value=""
							style="text-align:right; width:99%; border:0px;" readonly="readonly"
						/>
					</td>
					<td colspan="2" align="right">
						<input id="actual_estim_traffic_totle" name="actual_estim_traffic_totle" 
							type="text" value=""
							style="text-align:right; width:99%; border:0px;" readonly="readonly"
						/>
					</td>
					<td colspan="2" align="right">
						<input id="actual_estim_meal_totle" name="actual_estim_meal_totle"
							type="text" value=""
							style="text-align:right; width:99%; border:0px;" readonly="readonly"
						/>
					</td>
					<td colspan="2" align="right">
						<input id="actual_estim_extra_totle" name="actual_estim_extra_totle"
							type="text" value=""
							style="text-align:right; width:99%; border:0px;" readonly="readonly"
						/>
					</td>
					<td colspan="2" align="right">
						<input id="actual_estim_other_totle" name="actual_estim_other_totle"
							type="text" value=""
							style="text-align:right; width:99%; border:0px;" readonly="readonly"
						/>
					</td>
					<td colspan="2" align="right">
						<input id="actual_estim_sum_totle" name="actual_estim_sum_totle"
							type="text" value=""
							style="text-align:right; width:99%; border:0px;" readonly="readonly"
						/>
					</td>
				</tr>
				<tr>
					<td colspan="6" align="right">百分比</td>
					<!-- 
					<td colspan="2" align="center">
						<input name="actual_estim_city_avg" id="actual_estim_city_avg" 
							type="text" value=""
							style="text-align:center; width:99%; border:0px;" readonly="readonly"
						/>
					</td>
					 -->
					<td colspan="2" align="right">
						<input name="actual_expense_receipt_totle" id="actual_expense_receipt_totle" 
							type="text" value=""
							style="text-align:right; width:99%; border:0px;" readonly="readonly"
						/>
					</td>
					<td colspan="2" align="right">
						<input name="actual_estim_hotel_avg" id="actual_estim_hotel_avg" 
							type="text" value=""
							style="text-align:right; width:99%; border:0px;" readonly="readonly"
						/>
					</td>
					<td colspan="2" align="right">
						<input id="actual_estim_traffic_avg" name="actual_estim_traffic_avg" 
							type="text" value=""
							style="text-align:right; width:99%; border:0px;" readonly="readonly"
						/>
					</td>
					<td colspan="2" align="right">
						<input id="actual_estim_meal_avg" name="actual_estim_meal_avg"
							type="text" value=""
							style="text-align:right; width:99%; border:0px;" readonly="readonly"
						/>
					</td>
					<td colspan="2" align="right">
						<input id="actual_estim_extra_avg" name="actual_estim_extra_avg"
							type="text" value=""
							style="text-align:right; width:99%; border:0px;" readonly="readonly"
						/>
					</td>
					<td colspan="2" align="right">
						<input id="actual_estim_other_avg" name="actual_estim_other_avg"
							type="text" value=""
							style="text-align:right; width:99%; border:0px;" readonly="readonly"
						/>
					</td>
					<td colspan="2" align="right">
						<input id="actual_estim_sum_avg" name="actual_estim_sum_avg"
							type="text" value=""
							style="text-align:right; width:99%; border:0px;" readonly="readonly"
						/>
					</td>
				</tr>
				<%for(int city_n = 0 , n = 0 , j = 0  , cont = 0; city_n < actual.length ; city_n++){ %>
					<%
						HashMap CityMap = (HashMap)actual_tripExpense_list.get(city_n);
						HashMap RouteMap = (HashMap)actual_tripRoute_list.get(city_n);
						String city = "null";
						if(CityMap != null)
							city  = (String)CityMap.get("expense_city"); 
						
						
						HashMap temap =  (HashMap)actual_tripExpense_list.get(city_n);
						int actual_estim_hotel_h 
							= temap.get("estim_hotel").equals("")?0:Integer.parseInt((String)temap.get("estim_hotel"));;
						int actual_estim_traffic_h
							= temap.get("estim_traffic").equals("")?0:Integer.parseInt((String)temap.get("estim_traffic"));
						int actual_estim_meal_h
							= temap.get("estim_meal").equals("")?0:Integer.parseInt((String)temap.get("estim_meal"));
						int actual_estim_extra_h
							= temap.get("estim_extra").equals("")?0:Integer.parseInt((String)temap.get("estim_extra"));
						int actual_estim_other_h
							= temap.get("estim_other").equals("")?0:Integer.parseInt((String)temap.get("estim_other"));
						int actual_estim_sum 
							= actual_estim_hotel_h + actual_estim_traffic_h
							+ actual_estim_meal_h + actual_estim_extra_h + actual_estim_other_h;
						int actual_duration_h
							= temap.get("duration").equals("")?0:Integer.parseInt((String)temap.get("duration"));
						
					%>
					<input type="hidden" id="actual_estim_city_h<%=city_n %>" 
						name="actual_estim_city_h" value="<%=city_name[city_n] %>">
					<input type="hidden" id="actual_estim_hotel_h<%=city_n %>" 
						name="actual_estim_hotel_h" value="<%=actual_estim_hotel_h %>">
					<input type="hidden" id="actual_estim_traffic_h<%=city_n %>" 
						name="actual_estim_traffic_h" value="<%=actual_estim_traffic_h %>">
					<input type="hidden" id="actual_estim_meal_h<%=city_n %>" 
						name="actual_estim_meal_h" value="<%=actual_estim_meal_h %>">
					<input type="hidden" id="actual_estim_extra_h<%=city_n %>" 
						name="actual_estim_extra_h" value="<%=actual_estim_extra_h %>">
					<input type="hidden" id="actual_estim_other_h<%=city_n %>" 
						name="actual_estim_other_h" value="<%=actual_estim_other_h %>">
					<input type="hidden" id="actual_estim_sum_h<%=city_n %>" 
						name="actual_estim_sum_h" value="<%=actual_estim_sum %>">
					<input type="hidden" id="actual_duration_h<%=city_n %>" 
						name="actual_duration_h" value="<%=actual_duration_h %>">
					
					<tr></tr>
					
					<tr>
						<td colspan="6" align="right"><b>城市 : <%=city_name[city_n] %></b></td>
						<td colspan="2" align="right">
							<input name="actual_expense_receipt_totle" id="actual_expense_receipt_totle" 
								type="text" value=""
								style="text-align:right; width:99%; border:0px;" readonly="readonly"
							/>
						</td>
						<td colspan="2" align="right"><%=actual_estim_hotel_h/actual_duration_h %></td>
						<td colspan="2" align="right"><%=actual_estim_traffic_h/actual_duration_h %></td>
						<td colspan="2" align="right"><%=actual_estim_meal_h/actual_duration_h %></td>
						<td colspan="2" align="right"><%=actual_estim_extra_h/actual_duration_h %></td>
						<td colspan="2" align="right"><%=actual_estim_other_h/actual_duration_h %></td>
						<td colspan="2" align="right"><%=actual_estim_sum/actual_duration_h %></td>
					</tr>
					<tr>
						<td colspan="3" align="center">天數</td>
						<td colspan="3" align="center">日期</td>
						<td colspan="2" align="center">票據張數</td>
						<td colspan="2" align="center">住宿</td>
						<td colspan="2" align="center">交通</td>
						<td colspan="2" align="center">膳食</td>
						<td colspan="2" align="center">雜項</td>
						<td colspan="2" align="center">其他&票價</td>
						<td colspan="2" align="center">總計</td>
					</tr>
					<%for(int i = 0 ; i < actual[city_n].length ; i++ , n++){ %>
					<%
					HashMap acmap = null;
					String actual_expense_date = "";
					String actual_expense_receipt = "0";
					String actual_estim_hotel = "0";
					String actual_expense_traffic = "0";
					String actual_expense_meal = "0";
					String actual_expense_extra = "0";
					String actual_expense_other = "0";
					String actual_expense_sum = "0";
					if(actual_TripActual_Expense_list != null && j < actual_TripActual_Expense_list.size()){
						acmap = (HashMap)actual_TripActual_Expense_list.get(j);
					}
					if(acmap != null && acmap.get("day_nbr").equals(""+(i+1))){
						actual_expense_date = (String)acmap.get("expense_date");
						actual_expense_receipt = (String)acmap.get("expense_receipt");
						actual_estim_hotel = (String)acmap.get("estim_hotel");
						actual_expense_traffic = (String)acmap.get("expense_traffic");
						actual_expense_meal = (String)acmap.get("expense_meal");
						actual_expense_extra = (String)acmap.get("expense_extra");
						actual_expense_other = (String)acmap.get("expense_other");
						actual_expense_sum = (String)acmap.get("expense_sum");
						j++;
					}
					//cont
					if(actual_expense_date.equals("")){
						//如果資料庫理沒有日期的話 就依據triproute 的 date_rout_dep 來填入日期
						String datas[] = ((String)RouteMap.get("date_rout_dep")).split("/");
						java.util.Date d1 = new java.util.Date();
						java.text.SimpleDateFormat sdfmt = new java.text.SimpleDateFormat("yyyy/MM/dd");
						java.util.Calendar cal = java.util.Calendar.getInstance();
						cal.set(Integer.parseInt(datas[0]),Integer.parseInt(datas[1]),Integer.parseInt(datas[2]));
						d1 = cal.getTime();
					    // Calendar.YEAR 代表加減年
					    // Calendar.MONTH 代表加減月份
					    // Calendar.DATE 代表加減天數
					    // Calendar.HOUR 代表加減小時數
					    // Calendar.MINUTE 代表加減分鐘數
					    // Calendar.SECOND 代表加減秒數
						cal.add(java.util.Calendar.DATE,i+1);
						d1 = cal.getTime();
						actual_expense_date = sdfmt.format(d1);
					}
					%>
					<tr>
						<td colspan="3" align="center">
							第<%=i+1 %>天
							<input name="actual_day_nbr" type="hidden" value="<%=i+1 %>" />
							<input name="actual_expense_city" type="hidden" value="<%=city %>" />
						</td>
						<td colspan="3" align="center" style="padding-right:5px;">
							<!-- 日期 --><input name="actual_expense_date" id="actual_expense_date<%=cont++ %>" type="text" 
							value="<%=actual_expense_date %>" 
							id="actual_expense_date" style="text-align:right; width:99%;"/>
						</td>
						<td colspan="2" style="padding-right:5px;">
							<!-- 票據張數 --><input name="actual_expense_receipt" type="text" 
							value="<%=actual_expense_receipt %>" 
							id="actual_expense_receipt" style="text-align:right; width:99%;"/>
						</td>
						<td colspan="2" style="padding-right:5px;">
							<!-- 住宿 --><input name="actual_estim_hotel" type="text" 
							value="<%=actual_estim_hotel %>" onchange="actual_sum(<%=n %>)" onClick="actual_totle(<%=city_n+","+(cont-1) %>)"
							id="actual_estim_hotel" style="text-align:right; width:99%;"/>
						</td>
						<td colspan="2" style="padding-right:5px;">
							<!-- 交通 --><input name="actual_expense_traffic" type="text" 
							value="<%=actual_expense_traffic %>" onchange="actual_sum(<%=n %>)" onClick="actual_totle(<%=city_n+","+(cont-1) %>)"
							id="actual_expense_traffic" style="text-align:right; width:99%;"/>
						</td>
						<td colspan="2" style="padding-right:5px;">
							<!-- 膳食 --><input name="actual_expense_meal" type="text" 
							value="<%=actual_expense_meal %>" onchange="actual_sum(<%=n %>)" onClick="actual_totle(<%=city_n+","+(cont-1) %>)"
							id="actual_expense_meal" style="text-align:right; width:99%;"/>
						</td>
						<td colspan="2" style="padding-right:5px;">
							<!-- 雜項 --><input name="actual_expense_extra" type="text" 
							value="<%=actual_expense_extra %>" onchange="actual_sum(<%=n %>)" onClick="actual_totle(<%=city_n+","+(cont-1) %>)"
							id="actual_expense_extra" style="text-align:right; width:99%;"/>
						</td>
						<td colspan="2" style="padding-right:5px;">
							<!-- 其他&票價 --><input name="actual_expense_other" type="text" 
							value="<%=actual_expense_other %>" onchange="actual_sum(<%=n %>)" onClick="actual_totle(<%=city_n+","+(cont-1) %>)"
							id="actual_expense_other" style="text-align:right; width:99%;"/>
						</td>
						<td colspan="2" align="right">
							<!-- 總計 --><input name="actual_expense_sum" type="text" onClick="actual_totle(<%=city_n+","+(cont-1) %>)"
							value="<%=actual_expense_sum %>"
							id="actual_expense_sum" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
						</td>
					</tr>
					<%} %>
				<%} %>
				<script type="text/javascript">
					//日期js 報帳
					var f = document.forms['actual'];
					var e = f.elements["actual_expense_date"];
					//
					var RANGE_CAL = new Array(e.length);
					for(var i = 0 ; i < e.length ; i++){
						RANGE_CAL[i] = new Calendar({
							inputField: e[i].id,
							dateFormat: "%Y/%m/%d",
							trigger: e[i].id,
							bottomBar: true,
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
							}
						});
					}
					function clearRangeStart(n) {
						var e = document.forms['actual'];
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

					//預先取得一次數值
					actual_totle(1,3);
				</script>
				<tr>
					<td colspan="3"></td>
					<td colspan="3" align="center" style="padding-right:5px;"></td>
					<td colspan="2" style="padding-right:5px;"></td>
					<td colspan="2" style="padding-right:5px;"></td>
					<td colspan="2" style="padding-right:5px;"></td>
					<td colspan="4" align="center" style="padding-right:5px;">
					<font color="#cc0000"><b>※送出後不得修改※</b></font>
					</td>
					<td colspan="2" align="right" style="padding-right:5px;"></td>
					<td colspan="2" align="right"></td>
				</tr>
				<tr>
					<td colspan="3"></td>
					<td colspan="3" align="center" style="padding-right:5px;"></td>
					<td colspan="2" style="padding-right:5px;"></td>
					<td colspan="2" style="padding-right:5px;"></td>
					<td colspan="2" style="padding-right:5px;"></td>
					<td colspan="4" align="center" style="padding-right:5px;">
						<input name="submit" type="submit" value="暫存" />
						<input name="submit" type="submit" value="送出" />
						<input name="submit" type="submit" value="返回" />
					</td>
					<td colspan="2" align="right" style="padding-right:5px;"></td>
					<td colspan="2" align="right"></td>
				</tr>
			</table>
			</form>
			<%} %>
			
			<!--編輯trip  -->
			<%if(request.getParameter("isTrip") != null && request.getParameter("isTrip").equals("edit")){ %>
			<%if(edit_trip_list != null){ %>
			<form name="estim" action="User" method="post">
			<input name="isAction" type="hidden" value="trip" />
			<input name="status" type="hidden" value="<%=edit_trip_list.get("trip_nbr") %>" />
			<table width="100%">
				<tr>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
				</tr>
				<tr>
					<td colspan="20">申請出差</td>
				</tr>
				<tr>
					<td colspan="3">修改申請</td>
					<td colspan="3"><%=session.getAttribute("user_cname")%></td>
					<td colspan="14">
						<%if(edit_trip_list.get("trip_status").equals("1") 
							|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="trip_description" type="text" value="<%=edit_trip_list.get("trip_description") %>" 
							id="trip_description" style="text-align:left; width:99%;"/>
						<%}else {
							out.println(edit_trip_list.get("trip_description"));
						}%>
					</td>
				</tr>
				<tr>
					<td rowspan="4">機票資訊</td>
					<td rowspan="4" colspan="12">
						<div style="overflow:auto; height:100px; margin:0px; padding:0px; background:#999;">
						<table>
							<tr>
								<td width="15%">日期</td>
								<td width="15%">出發機場</td>
								<td width="15%">落地機場</td>
								<td width="10%">航班</td>
								<td width="5%" align="center">轉機</td>
							</tr>
							<%
							if(edit_tripRoute_list != null){
								for(int i = 0 ; i < edit_tripRoute_list.size() ; i++){
									HashMap map = (HashMap)edit_tripRoute_list.get(i);
							%>
								<tr>
									<td width="15%" style="padding-right:5px;">
										<%if(edit_trip_list.get("trip_status").equals("1") 
												|| edit_trip_list.get("trip_status").equals("6")){ %>
											<input name="air_Date" type="text" value="<%=map.get("date_rout_dep") %>" 
											id="air_Date<%=i %>" style="text-align:left; width:99%;"/>
										<%}else {
											out.println(map.get("date_rout_dep"));
										}%>
									</td>
									<td width="15%" style="padding-right:5px;">
										<%if(edit_trip_list.get("trip_status").equals("1") 
												|| edit_trip_list.get("trip_status").equals("6")){ %>
											<input name="air_Origin" type="text" value="<%=map.get("rout_dep_port") %>" 
											id="air_Origin" style="text-align:left; width:99%;"/>
										<%}else{
											out.println(map.get("rout_dep_port"));
										}%>
									</td>
									<td width="15%" style="padding-right:5px;">
										<%if(edit_trip_list.get("trip_status").equals("1") 
												|| edit_trip_list.get("trip_status").equals("6")){ %>
											<input name="air_Destination" type="text" value="<%=map.get("rout_ariv_port") %>" 
											id="air_Destination" style="text-align:left; width:99%;"/>
										<%}else{
											out.println(map.get("rout_ariv_port"));
										}%>
									</td>
									<td width="10%" style="padding-right:5px;">
										<%if(edit_trip_list.get("trip_status").equals("1") 
												|| edit_trip_list.get("trip_status").equals("6")){ %>
											<input name="air_Flight" type="text" value="<%=map.get("rout_flight") %>" 
											id="air_Flight" style="text-align:left; width:99%;"/>
										<%}else{
											out.println(map.get("rout_flight"));
										}%>
									</td>
									<td width="5%" align="center">
										<%if(edit_trip_list.get("trip_status").equals("1") 
											|| edit_trip_list.get("trip_status").equals("6")){ %>
											<input name="air_transfer" type="checkbox" value="<%=i%>" id="air_transfer"
											<%
											if(map.get("rout_transfer").equals("y"))
												out.print("checked=\"checked\" ");
											%>
											/>
										<%}else{
											if(map.get("rout_transfer").equals("y")){
												out.println("yes");
											}else{
												out.println("no");
											}
										}%>
									</td>
								</tr>
							<%
								}
							}
							%>
							<%if(edit_trip_list.get("trip_status").equals("1") || edit_trip_list.get("trip_status").equals("6")){ %>
								<%for(int i = edit_tripRoute_list.size() ; i < 6 ; i++){ %>
								<tr>
									<td width="15%" style="padding-right:5px;">
										<input name="air_Date" type="text" value="" 
										id="air_Date<%=i %>" style="text-align:left; width:99%;"/>
									</td>
									<td width="15%" style="padding-right:5px;">
										<input name="air_Origin" type="text" value="" 
										id="air_Origin" style="text-align:left; width:99%;"/>
									</td>
									<td width="15%" style="padding-right:5px;">
										<input name="air_Destination" type="text" value="" 
										id="air_Destination" style="text-align:left; width:99%;"/>
									</td>
									<td width="10%" style="padding-right:5px;">
										<input name="air_Flight" type="text" value="" 
										id="air_Flight" style="text-align:left; width:99%;"/>
									</td>
									<td width="5%" align="center">
										<input name="air_transfer" type="checkbox" value="<%=i%>" id="air_transfer"/>
									</td>
								</tr>
								<%} %>
							<%} %>
							<script type="text/javascript">
								//日期js 修改trip
								var f = document.forms['estim'];
								var e = f.elements["air_Date"];
								//
								var RANGE_CAL = new Array(e.length);
								for(var i = 0 ; i < e.length ; i++){
									RANGE_CAL[i] = new Calendar({
										inputField: e[i].id,
										dateFormat: "%Y/%m/%d",
										trigger: e[i].id,
										bottomBar: true,
										onSelect: function() {
											var date = Calendar.intToDate(this.selection.get());
										}
									});
								}
								function clearRangeStart(n) {
									var e = document.forms['actual'];
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
						</table>
						</div>
					</td>
					<td colspan="2" align="right">PNR:</td>
					<td colspan="2" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
							|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="air_pnt" type="text" value="<%=edit_trip_list.get("trip_pnr") %>" 
							id="air_pnt" style="text-align:left; width:99%;"/>
						<%}else{
							out.println(edit_trip_list.get("trip_pnr"));
						}%>
					</td>
					<td colspan="3">[航班訂位]</td>
				</tr>
				<tr>
					<td colspan="2" align="right">票價預估:</td>
					<td colspan="2" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
							|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="air_estim" type="text" value="<%=edit_trip_list.get("tkt_estim") %>" 
							id="air_estim" style="text-align:left; width:99%;"/>
						<%}else{
							out.println(edit_trip_list.get("tkt_estim"));
						}%>
					</td>
					<td colspan="3">
						<a href="http://et.3shop.com.tw/AmexEitinerary/searchAv_new.jsp?pnrCode=MKZTKJ&amp;emailAddr=davidliu@dynalab.com.tw;bryant@net1.com.tw">
							[航班查詢]
						</a>
					</td>
				</tr>
				<tr>
					
					<td colspan="2" align="right">出差天數:</td>
					<td colspan="2" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
							|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="air_interval" type="text" value="<%=edit_trip_list.get("trip_interval") %>" 
							id="air_interval" style="text-align:left; width:99%;"/>
						<%}else{
							out.println(edit_trip_list.get("trip_interval"));
						}%>
					</td>
					<td colspan="3">[票價查詢]</td>
				</tr>
				<tr>
					<td colspan="2" align="right">總費用:</td>
					<td colspan="2">
					<%
					if(edit_trip_list.get("tkt_estim") != null 
							&& !edit_trip_list.get("tkt_estim").equals("")
							&& edit_trip_list.get("expense_estimate") != null 
							&& !edit_trip_list.get("expense_estimate").equals("")){
						int i = Integer.parseInt((String)edit_trip_list.get("expense_estimate"));
						int j = Integer.parseInt((String)edit_trip_list.get("tkt_estim"));
						
						out.println("$"+(i+j));
					}
					%>
					</td>
					<td colspan="4"></td>
				</tr>
				<tr>
					<td colspan="1">預算</td>
					<td colspan="2" align="center">城市</td>
					<td colspan="1" align="center">天數</td>
					<td colspan="2" align="center">住宿</td>
					<td colspan="2" align="center">交通</td>
					<td colspan="2" align="center">膳食</td>
					<td colspan="2" align="center">雜項</td>
					<td colspan="2" align="center">其他&票價</td>
					<td colspan="2" align="center">總計</td>
					<td colspan="4" align="center">備註</td>
				</tr>
				<tr>
					<td colspan="1">小計</td>
					<td colspan="2"></td>
					<td colspan="1" align="right" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
							|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="sum_estim_duration" type="text"
							value="<%
							if(edit_tripExpense_list != null){
								int sum = 0;
								for(int i = 0 ; i < edit_tripExpense_list.size() ; i++){ 
									HashMap map = (HashMap)edit_tripExpense_list.get(i);
									sum += Integer.parseInt((String)map.get("duration"));
								}
								out.print(sum);
							}
							%>"
							id="sum_estim_duration" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
						<%}else{
							//out.println(edit_trip_list.get("trip_interval"));
							if(edit_tripExpense_list != null){
								int sum = 0;
								for(int i = 0 ; i < edit_tripExpense_list.size() ; i++){ 
									HashMap map = (HashMap)edit_tripExpense_list.get(i);
									sum += Integer.parseInt((String)map.get("duration"));
								}
								out.print(sum);
							}
						}%>
					</td>
					<td colspan="2" align="right" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
							|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="sum_estim_hotel" type="text" value="<%=edit_trip_list.get("hotel_estim") %>"
							id="sum_estim_hotel" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
						<%}else{
							out.println(edit_trip_list.get("hotel_estim"));
						}%>
					</td>
					<td colspan="2" align="right" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
							|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="sum_estim_traffic" type="text" value="<%=edit_trip_list.get("traffic_estim") %>"
							id="sum_estim_traffic" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
						<%}else{
							out.println(edit_trip_list.get("traffic_estim"));
						}%>
					</td>
					<td colspan="2" align="right" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
							|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="sum_estim_meal" type="text" value="<%=edit_trip_list.get("meal_estim") %>"
							id="sum_estim_meal" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
						<%}else{
							out.println(edit_trip_list.get("meal_estim"));
						}%>
					</td>
					<td colspan="2" align="right" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
								|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="sum_estim_extra" type="text" value="<%=edit_trip_list.get("extra_estim") %>"
							id="sum_estim_extra" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
						<%}else{
							out.println(edit_trip_list.get("extra_estim"));
						}%>
					</td>
					<td colspan="2" align="right" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
								|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="sum_estim_other" type="text" value="<%=edit_trip_list.get("other_estim") %>""
							id="sum_estim_other" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
						<%}else{
							out.println(edit_trip_list.get("other_estim"));
						}%>
					</td>
					<td colspan="2" align="right" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
								|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="sum_estim_totle" type="text" value="<%=edit_trip_list.get("expense_estimate") %>" 
							id="sum_estim_totle" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
						<%}else{
							out.println(edit_trip_list.get("expense_estimate"));
						}%>
					</td>
					<td colspan="4" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
								|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="Expense_note" type="text" 
							value="" 
							id="Expense_note" style="text-align:left; width:99%; border:0px;"/>
						<%}else{
							
						}%>
					</td>
				</tr>
				
				<%for(int i = 0 ; i < edit_tripExpense_list.size() ; i++){ %>
				<%HashMap map = (HashMap)edit_tripExpense_list.get(i); %>
				<tr>
					<td colspan="1" align="center"><%=i+1 %></td>
					<td colspan="2">
						<%if(edit_trip_list.get("trip_status").equals("1") 
							|| edit_trip_list.get("trip_status").equals("6")){ %>
							<select name="expense_city" id="expense_city">
								<option value="0"></option>
								<%for(int j = 0 ; j < AirName.length ; j++){%>
									<%
									ArrayList arr = (ArrayList)AirMap.get(AirName[j]);
									if(arr != null){
										for(int k = 0 ; k < arr.size() ; k++){
											String val[] = ((String)arr.get(k)).split(",");
											//val[0]=中文  val[0]=英文   val[0]=縮寫 
									%>
										<option value="<%=val[2] %>" 
										<%=map.get("expense_city") != null && map.get("expense_city").equals(val[2])?"selected=\"selected\" ":" " %>>
										<%=val[0] %>
										</option>
									<%
										}
									}
									%>
								<%}%>
							</select>
						<%}else{
							for(int j = 0 ; j < AirName.length ; j++){
								ArrayList arr = (ArrayList)AirMap.get(AirName[j]);
								if(arr != null){
									for(int k = 0 ; k < arr.size() ; k++){
										String val[] = ((String)arr.get(k)).split(",");
										//val[0]=中文  val[0]=英文   val[0]=縮寫 
								
										if(map.get("expense_city") != null && map.get("expense_city").equals(val[2])){
											out.println(val[0]);
										}
									}
								}
							}
						} %>
					</td>
					<td colspan="1" align="right" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
							|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="estim_duration" type="text" value="<%=map.get("duration") %>" onchange="sum_estim()"
							id="estim_duration" style="text-align:right; width:99%;"/>
						<%}else{
							out.println(map.get("duration"));
						} %>
					</td>
					<td colspan="2" align="right" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
							|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="estim_hotel" type="text" value="<%=map.get("estim_hotel") %>" onchange="sum_estim()"
							id="estim_hotel" style="text-align:right; width:99%;"/>
						<%}else{
							out.println(map.get("estim_hotel"));
						} %>
					</td>
					<td colspan="2" align="right" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
							|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="estim_traffic" type="text" value="<%=map.get("estim_traffic") %>" onchange="sum_estim()"
							id="estim_traffic" style="text-align:right; width:99%;"/>
						<%}else{
							out.println(map.get("estim_traffic"));
						} %>
					</td>
					<td colspan="2" align="right" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
							|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="estim_meal" type="text" value="<%=map.get("estim_meal") %>" onchange="sum_estim()"
							id="estim_meal" style="text-align:right; width:99%;"/>
						<%}else{
							out.println(map.get("estim_meal"));
						} %>
					</td>
					<td colspan="2" align="right" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
							|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="estim_extra" type="text" value="<%=map.get("estim_extra") %>" onchange="sum_estim()"
							id="estim_extra" style="text-align:right; width:99%;"/>
						<%}else{
							out.println(map.get("estim_extra"));
						} %>
					</td>
					<td colspan="2" align="right" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
							|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="estim_other" type="text" value="<%=map.get("estim_other") %>" onchange="sum_estim()"
							id="estim_other" style="text-align:right; width:99%;"/>
						<%}else{
							out.println(map.get("estim_other"));
						} %>
					</td>
					<td colspan="2" align="right" name="estim_sum">
						<%if(edit_trip_list.get("trip_status").equals("1") 
							|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="estim_sum" type="text" 
							value="<% 
							out.println(
									Integer.parseInt((String)map.get("estim_hotel"))+
									Integer.parseInt((String)map.get("estim_traffic"))+
									Integer.parseInt((String)map.get("estim_meal"))+
									Integer.parseInt((String)map.get("estim_extra"))+
									Integer.parseInt((String)map.get("estim_other")));
							%>"
							id="estim_sum" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
						<%}else{
							out.println(
									Integer.parseInt((String)map.get("estim_hotel"))+
									Integer.parseInt((String)map.get("estim_traffic"))+
									Integer.parseInt((String)map.get("estim_meal"))+
									Integer.parseInt((String)map.get("estim_extra"))+
									Integer.parseInt((String)map.get("estim_other")));
						} %>
					</td>
					<td colspan="4" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
							|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="expense_note" type="text" value="<%=map.get("expense_note") %>" 
							id="expense_note" style="text-align:left; width:99%;"/>
						<%}else{
							out.println(map.get("expense_note"));
						} %>
					</td>
				</tr>
				<%} %>
				<%if(edit_trip_list.get("trip_status").equals("1") || edit_trip_list.get("trip_status").equals("6")){ %>
					<%for(int i = edit_tripExpense_list.size() ; i < 6 ; i++){ %>
					<tr>
						<td colspan="1" align="center"><%=i+1 %></td>
						<td colspan="2">
							<select name="expense_city" id="expense_city">
								<option value="0"></option>
								<%for(int j = 0 ; j < AirName.length ; j++){%>
									<%
									ArrayList arr = (ArrayList)AirMap.get(AirName[j]);
									if(arr != null){
										for(int k = 0 ; k < arr.size() ; k++){
											String val[] = ((String)arr.get(k)).split(",");
											//val[0]=中文  val[0]=英文   val[0]=縮寫 
									%>
										<option value="<%=val[2] %>"><%=val[0] %></option>
									<%
										}
									}
									%>
								<%}%>
							</select>
						</td>
						<td colspan="1" align="center" style="padding-right:5px;">
							<input name="estim_duration" type="text" value="" onchange="sum_estim()"
							id="estim_duration" style="text-align:right; width:99%;"/>
						</td>
						<td colspan="2" style="padding-right:5px;">
							<input name="estim_hotel" type="text" value="" onchange="sum_estim()"
							id="estim_hotel" style="text-align:right; width:99%;"/>
						</td>
						<td colspan="2" style="padding-right:5px;">
							<input name="estim_traffic" type="text" value="" onchange="sum_estim()"
							id="estim_traffic" style="text-align:right; width:99%;"/>
						</td>
						<td colspan="2" style="padding-right:5px;">
							<input name="estim_meal" type="text" value="" onchange="sum_estim()"
							id="estim_meal" style="text-align:right; width:99%;"/>
						</td>
						<td colspan="2" style="padding-right:5px;">
							<input name="estim_extra" type="text" value="" onchange="sum_estim()"
							id="estim_extra" style="text-align:right; width:99%;"/>
						</td>
						<td colspan="2" style="padding-right:5px;">
							<input name="estim_other" type="text" value="" onchange="sum_estim()"
							id="estim_other" style="text-align:right; width:99%;"/>
						</td>
						<td colspan="2" align="right" name="estim_sum">
							<input name="estim_sum" type="text" value=""
							id="estim_sum" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
						</td>
						<td colspan="4" style="padding-right:5px;">
							<input name="expense_note" type="text" value="" 
							id="expense_note" style="text-align:left; width:99%;"/>
						</td>
					</tr>
					<%} %>
				<%} %>
			</table>
			<table width="100%">
				<tr>
					<td width="10%" align="center">申請流程</td>
					<td width="10%" align="center">立案</td>
					<td width="10%" align="center">申請</td>
					<td width="10%" align="center">批核</td>
					<td width="10%" align="center">決行</td>
					<td width="10%" align="center">報帳</td>
					<td width="10%" align="center"></td>
					<td width="10%" align="center"></td>
					<td width="10%" align="center"></td>
					<td width="10%" align="center">取消</td>
				</tr>
				<tr>
					<td colspan="1" align="center">日期</td>
					<td colspan="1">
					<%
					//建立時間
					if(edit_trip_list.get("time_create") != null 
					&& !((String)edit_trip_list.get("time_create")).equals("")){
						String tmp_data[] = ((String)edit_trip_list.get("time_create")).split("-");
						if(Integer.parseInt(tmp_data[1]) < 10){
							tmp_data[1] = "0" + tmp_data[1];
						}
						if(Integer.parseInt(tmp_data[2]) < 10){
							tmp_data[2] = "0" + tmp_data[2];
						}
						out.println(tmp_data[0]+"/"+tmp_data[1]+"/"+tmp_data[2]);
					}else
						out.println("尚無記錄");
					%>
					</td>
					<td colspan="1">
					<%
					//申請日期
					if(edit_trip_list.get("time_apply") != null 
					&& !((String)edit_trip_list.get("time_apply")).equals("")){
						String tmp_data[] = ((String)edit_trip_list.get("time_apply")).split("-");
						if(Integer.parseInt(tmp_data[1]) < 10){
							tmp_data[1] = "0" + tmp_data[1];
						}
						if(Integer.parseInt(tmp_data[2]) < 10){
							tmp_data[2] = "0" + tmp_data[2];
						}
						out.println(tmp_data[0]+"/"+tmp_data[1]+"/"+tmp_data[2]);
					}else
						out.println("尚無記錄");
					%>
					</td>
					<td colspan="1">
					<%
					//核准時間
					if(edit_trip_list.get("time_agree") != null 
					&& !((String)edit_trip_list.get("time_agree")).equals("")){
						String tmp_data[] = ((String)edit_trip_list.get("time_agree")).split("-");
						if(Integer.parseInt(tmp_data[1]) < 10){
							tmp_data[1] = "0" + tmp_data[1];
						}
						if(Integer.parseInt(tmp_data[2]) < 10){
							tmp_data[2] = "0" + tmp_data[2];
						}
						out.println(tmp_data[0]+"/"+tmp_data[1]+"/"+tmp_data[2]);
					}else
						out.println("尚無記錄");
					%>
					</td>
					<td colspan="1">
					<%
					//准行時間
					if(edit_trip_list.get("time_final") != null 
					&& !((String)edit_trip_list.get("time_final")).equals("")){
						String tmp_data[] = ((String)edit_trip_list.get("time_final")).split("-");
						if(Integer.parseInt(tmp_data[1]) < 10){
							tmp_data[1] = "0" + tmp_data[1];
						}
						if(Integer.parseInt(tmp_data[2]) < 10){
							tmp_data[2] = "0" + tmp_data[2];
						}
						out.println(tmp_data[0]+"/"+tmp_data[1]+"/"+tmp_data[2]);
					}else
						out.println("尚無記錄");
					%>
					</td>
					<td colspan="1">尚無記錄</td>
					<td colspan="1"></td>
					<td colspan="1"></td>
					<td colspan="1"></td>
					<td colspan="1">
					<%
					//取消時間
					if(edit_trip_list.get("time_close") != null 
					&& !((String)edit_trip_list.get("time_close")).equals("")){
						String tmp_data[] = ((String)edit_trip_list.get("time_close")).split("-");
						if(Integer.parseInt(tmp_data[1]) < 10){
							tmp_data[1] = "0" + tmp_data[1];
						}
						if(Integer.parseInt(tmp_data[2]) < 10){
							tmp_data[2] = "0" + tmp_data[2];
						}
						out.println(tmp_data[0]+"/"+tmp_data[1]+"/"+tmp_data[2]);
					}else
						out.println("尚無記錄");
					%>
					</td>
				</tr>
				<tr>
					<td colspan="1" align="center"></td>
					<td colspan="1"></td>
					<td colspan="1" style="padding-right:5px;">
						<%if(edit_trip_list.get("time_apply") != null 
							&& !((String)edit_trip_list.get("time_apply")).equals("")){%>
							<%=session.getAttribute("user_cname")%><font size=2 color="#227700">提出</font>
						<%}else{
							out.print("　");
						}%>
					</td>
					<td colspan="1" style="padding-right:5px;">
							
						<%
						//列出所有的批何者
						if(edit_tripFlowUser_list != null && edit_tripFlow_list != null){
							for(int i = 0 ; i < edit_tripFlowUser_list.size() ; i++){
								HashMap map = (HashMap)edit_tripFlowUser_list.get(i);
								HashMap map2 = (HashMap)edit_tripFlow_list.get(i); 
								int n = Integer.parseInt((String)map2.get("tripflow_status"));
								if(n == 1){
									out.println(map.get("cname"));
									out.println("<font size=2 color=\"#227700\">待批<br></font>");
								}else if(n == 3){
									out.println(map.get("cname"));
									out.println("<font size=2 color=\"#000066\">已批<br></font>");
								}else if(n == 5){
									out.println(map.get("cname"));
									out.println("<font size=2 color=\"#AA0000\">退回<br></font>");
								}
							}
						}
						%>
					</td>
					<td colspan="1">
							
						<%
						//列出所有的批何者
						if(edit_tripFlowUser_list != null && edit_tripFlow_list != null){
							for(int i = 0 ; i < edit_tripFlowUser_list.size() ; i++){
								HashMap map = (HashMap)edit_tripFlowUser_list.get(i);
								HashMap map2 = (HashMap)edit_tripFlow_list.get(i); 
								int n = Integer.parseInt((String)map2.get("tripflow_status"));
								if(n == 2){
									out.println(map.get("cname"));
									out.println("<font size=2 color=\"#227700\">待准<br></font>");
								}else if(n == 4){
									out.println(map.get("cname"));
									out.println("<font size=2 color=\"#000066\">已准<br></font>");
								}
							}
						}
						%>
					</td>
					<td colspan="1"></td>
					<td colspan="1"></td>
					<td colspan="1"></td>
					<td colspan="1"></td>
					<td colspan="1">
					<%
					if(edit_trip_list.get("time_close") != null 
						&& !((String)edit_trip_list.get("time_close")).equals("")){
						out.println((String)edit_trip_list.get("close_describe"));
					}
					%>
					</td>
				</tr>
				<tr>
					<td align="center">備註</td>
					<td colspan="9" style="padding-right:5px;">
						<%if(edit_trip_list.get("trip_status").equals("1") 
								|| edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="trip_note" type="text" value="<%=edit_trip_list.get("trip_note") %>" 
							id="trip_note" style="text-align:left; width:99%;"/>
						<%}else{
							out.println(edit_trip_list.get("trip_note"));
						} %>
					</td>
				</tr>
				<tr>
					<td align="center">狀態</td>
					<td colspan="4">尚未申請批核 / 等待機票報價</td>
					<td colspan="1" align="center">
						<%if(edit_trip_list.get("trip_status").equals("1") || edit_trip_list.get("trip_status").equals("6")){ %>
							<input name="submit" type="submit" value="暫存" />
						<%} %>
					</td>
					
					<%if(edit_trip_list.get("trip_status").equals("1") || edit_trip_list.get("trip_status").equals("6")){ %>
						<td colspan="2" align="center">
							<input name="submit" type="submit" value="申請" />
							<input name="nextUser" type="text" value="<%=user_date.get("approver_mail") %>" 
							id="nextUser" style="font-size:10px; text-align:left; width:70%;"/>
						</td>
					<%}else{ %>
						<td></td>
						<td></td>
					<%} %>
					<td colspan="1" align="center">
						<%if(!edit_trip_list.get("trip_status").equals("5")){ %>
							<input name="submit" type="submit" value="取消" />
						<%} %>
					</td>
					<td colspan="1" align="center">
						<input name="submit" type="submit" value="返回" />
					</td>
				</tr>
			</table>
			</form>
			<%} %>
			<%} %>
			
			<%if(request.getParameter("isTrip") != null && request.getParameter("isTrip").equals("show")){ %>
			<form name="estim" action="User" method="post">
			<input name="isAction" type="hidden" value="trip" />
			<input name="status" type="hidden" value="new" />
			<table width="100%">
				<tr>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
					<td width="5%" height="0"></td>
				</tr>
				<tr>
					<td colspan="20">申請出差</td>
				</tr>
				<tr>
					<td colspan="3">出差申請</td>
					<td colspan="3"><%=session.getAttribute("user_cname")%></td>
					<td colspan="14" align="center">
						<input name="trip_description" type="text" value="填入出差說明" 
						id="trip_description" style="text-align:left; width:99%;"/>
					</td>
				</tr>
				<tr>
					<td rowspan="4">機票資訊</td>
					<td rowspan="4" colspan="12">
						<div style="overflow:auto; height:100px; margin:0px; padding:0px; background:#999;">
						<table>
							<tr>
								<td width="15%">日期</td>
								<td width="15%">出發機場</td>
								<td width="15%">落地機場</td>
								<td width="10%">航班</td>
								<td width="5%" align="center">轉機</td>
							</tr>
							<%for(int i = 0 ; i < 6 ; i++){ %>
							<tr>
								<td width="15%" style="padding-right:5px;">
									<input name="air_Date" type="text" value="" 
									id="air_Date<%=i %>" style="text-align:left; width:99%;"/>
								</td>
								<td width="15%" style="padding-right:5px;">
									<input name="air_Origin" type="text" value="" 
									id="air_Origin" style="text-align:left; width:99%;"/>
								</td>
								<td width="15%" style="padding-right:5px;">
									<input name="air_Destination" type="text" value="" 
									id="air_Destination" style="text-align:left; width:99%;"/>
								</td>
								<td width="10%" style="padding-right:5px;">
									<input name="air_Flight" type="text" value="" 
									id="air_Flight" style="text-align:left; width:99%;"/>
								</td>
								<td width="5%" align="center">
									<input name="air_transfer" type="checkbox" value="<%=i%>" id="air_transfer"/>
								</td>
							</tr>
							<%} %>
							<script type="text/javascript">
								//日期js 新增trip
								var f = document.forms['estim'];
								var e = f.elements["air_Date"];
								//
								var RANGE_CAL = new Array(e.length);
								for(var i = 0 ; i < e.length ; i++){
									RANGE_CAL[i] = new Calendar({
										inputField: e[i].id,
										dateFormat: "%Y/%m/%d",
										trigger: e[i].id,
										bottomBar: true,
										onSelect: function() {
											var date = Calendar.intToDate(this.selection.get());
										}
									});
								}
								function clearRangeStart(n) {
									var e = document.forms['actual'];
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
						</table>
						</div>
					</td>
					<td colspan="2" align="right">PNR:</td>
					<td colspan="2" style="padding-right:5px;">
						<input name="air_pnt" type="text" value="" 
						id="air_pnt" style="text-align:left; width:99%;"/>
					</td>
					<td colspan="3">[航班訂位]</td>
				</tr>
				<tr>
					<td colspan="2" align="right">票價預估:</td>
					<td colspan="2" style="padding-right:5px;">
						<input name="air_estim" type="text" value="" 
						id="air_estim" style="text-align:left; width:99%;"/>
					</td>
					<td colspan="3">
						<a href="http://et.3shop.com.tw/AmexEitinerary/searchAv_new.jsp?pnrCode=MKZTKJ&amp;emailAddr=davidliu@dynalab.com.tw;bryant@net1.com.tw">
							[航班查詢]
						</a>
					</td>
				</tr>
				<tr>
					
					<td colspan="2" align="right">出差天數:</td>
					<td colspan="2" style="padding-right:5px;">
						<input name="air_interval" type="text" value="" 
						id="air_interval" style="text-align:left; width:99%;"/>
					</td>
					<td colspan="3">[票價查詢]</td>
				</tr>
				<tr>
					<td colspan="7">　</td>
				</tr>
				<tr>
					<td colspan="1">預算</td>
					<td colspan="2" align="center">城市</td>
					<td colspan="1" align="center">天數</td>
					<td colspan="2" align="center">住宿</td>
					<td colspan="2" align="center">交通</td>
					<td colspan="2" align="center">膳食</td>
					<td colspan="2" align="center">雜項</td>
					<td colspan="2" align="center">其他&票價</td>
					<td colspan="2" align="center">總計</td>
					<td colspan="4" align="center">備註</td>
				</tr>
				<tr>
					<td colspan="1">小計</td>
					<td colspan="2"></td>
					<td colspan="1" align="right" style="padding-right:5px;">
						<input name="sum_estim_duration" type="text" value=""
						id="sum_estim_duration" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
					</td>
					<td colspan="2" style="padding-right:5px;">
						<input name="sum_estim_hotel" type="text" value=""
						id="sum_estim_hotel" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
					</td>
					<td colspan="2" style="padding-right:5px;">
						<input name="sum_estim_traffic" type="text" value=""
						id="sum_estim_traffic" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
					</td>
					<td colspan="2" style="padding-right:5px;">
						<input name="sum_estim_meal" type="text" value=""
						id="sum_estim_meal" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
					</td>
					<td colspan="2" style="padding-right:5px;">
						<input name="sum_estim_extra" type="text" value=""
						id="sum_estim_extra" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
					</td>
					<td colspan="2" style="padding-right:5px;">
						<input name="sum_estim_other" type="text" value=""
						id="sum_estim_other" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
					</td>
					<td colspan="2" align="right" style="padding-right:5px;">
						<input name="sum_estim_totle" type="text" value="" 
						id="sum_estim_totle" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
					</td>
					<td colspan="4" style="padding-right:5px;">
						<input name="Expense_note" type="text" value="" 
						id="Expense_note" style="text-align:left; width:99%; border:0px;"/>
					</td>
				</tr>
				<%for(int i = 0 ; i < 6 ; i++){ %>
				<tr>
					<td colspan="1" align="center"><%=i+1 %></td>
					<td colspan="2">
						<select name="expense_city" id="expense_city">
							<option value="0"></option>
							<%for(int j = 0 ; j < AirName.length ; j++){%>
								<%
								ArrayList arr = (ArrayList)AirMap.get(AirName[j]);
								if(arr != null){
									for(int k = 0 ; k < arr.size() ; k++){
										String val[] = ((String)arr.get(k)).split(",");
										//val[0]=中文  val[0]=英文   val[0]=縮寫 
								%>
									<option value="<%=val[2] %>"><%=val[0] %></option>
								<%
									}
								}
								%>
							<%}%>
						</select>
					</td>
					<td colspan="1" align="center" style="padding-right:5px;">
						<input name="estim_duration" type="text" value="" onchange="sum_estim()"
						id="estim_duration" style="text-align:right; width:99%;"/>
					</td>
					<td colspan="2" style="padding-right:5px;">
						<input name="estim_hotel" type="text" value="" onchange="sum_estim()"
						id="estim_hotel" style="text-align:right; width:99%;"/>
					</td>
					<td colspan="2" style="padding-right:5px;">
						<input name="estim_traffic" type="text" value="" onchange="sum_estim()"
						id="estim_traffic" style="text-align:right; width:99%;"/>
					</td>
					<td colspan="2" style="padding-right:5px;">
						<input name="estim_meal" type="text" value="" onchange="sum_estim()"
						id="estim_meal" style="text-align:right; width:99%;"/>
					</td>
					<td colspan="2" style="padding-right:5px;">
						<input name="estim_extra" type="text" value="" onchange="sum_estim()"
						id="estim_extra" style="text-align:right; width:99%;"/>
					</td>
					<td colspan="2" style="padding-right:5px;">
						<input name="estim_other" type="text" value="" onchange="sum_estim()"
						id="estim_other" style="text-align:right; width:99%;"/>
					</td>
					<td colspan="2" align="right" name="estim_sum">
						<input name="estim_sum" type="text" value=""
						id="estim_sum" style="text-align:right; width:99%; border:0px;" readonly="readonly"/>
					</td>
					<td colspan="4" style="padding-right:5px;">
						<input name="expense_note" type="text" value="" 
						id="expense_note" style="text-align:left; width:99%;"/>
					</td>
				</tr>
				<%} %>
			</table>
			<table width="100%">
				<tr>
					<td width="10%" align="center">申請流程</td>
					<td width="10%" align="center">立案</td>
					<td width="10%" align="center">申請</td>
					<td width="10%" align="center">批核</td>
					<td width="10%" align="center">決行</td>
					<td width="10%" align="center">報帳</td>
					<td width="10%" align="center"></td>
					<td width="10%" align="center"></td>
					<td width="10%" align="center"></td>
					<td width="10%" align="center">取消</td>
				</tr>
				<tr>
					<td colspan="1" align="center">日期</td>
					<td colspan="1"></td>
					<td colspan="1"></td>
					<td colspan="1"></td>
					<td colspan="1"></td>
					<td colspan="1"></td>
					<td colspan="1"></td>
					<td colspan="1"></td>
					<td colspan="1"></td>
					<td colspan="1"></td>
				</tr>
				<tr>
					<td colspan="1" align="center"></td>
					<td colspan="1"></td>
					<td colspan="1" style="padding-right:5px;">
					</td>
					<td colspan="1"></td>
					<td colspan="1"></td>
					<td colspan="1"></td>
					<td colspan="1"></td>
					<td colspan="1"></td>
					<td colspan="1"></td>
					<td colspan="1"></td>
				</tr>
				<tr>
					<td align="center">備註</td>
					<td colspan="9" style="padding-right:5px;">
						<input name="trip_note" type="text" value="" 
						id="trip_note" style="text-align:left; width:99%;"/>
					</td>
				</tr>
				<tr>
					<td align="center">狀態</td>
					<td colspan="4">尚未申請批核 / 等待機票報價</td>
					<td colspan="1" align="center">
						<input name="submit" type="submit" value="暫存" />
					</td>
					<td colspan="2" align="center">
						<input name="submit" type="submit" value="申請" />
						<input name="nextUser" type="text" value="<%=user_date.get("approver_mail") %>" 
						id="nextUser" style="font-size:10px; text-align:left; width:70%;"/>
					</td>
					<td colspan="1" align="center">
						<input name="submit" type="submit" value="取消" />
					</td>
					<td colspan="1" align="center">
						<a href="User">[返回]</a>
					</td>
				</tr>
			</table>
			</form>
			<%} %>
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

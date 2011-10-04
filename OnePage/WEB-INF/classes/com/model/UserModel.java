package com.model;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;

import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import com.web.Singleton;

public class UserModel {
	HttpServletRequest request = null;
	HttpServletResponse response = null;
	HttpServlet servlet = null;
	HttpSession session = null;
	TopUserModel topmodel = new TopUserModel(); 
	
	HashMap CorpMap = null;
	Statement CorpStmt = null;
	
	HashMap DBMap = null;
	
	//HashMap DB = null;
	Statement DBStmt = null;
	
	Calendar calendar = Calendar.getInstance();
	
	Singleton singleton;
	public void set( HttpServlet servlet , HttpServletRequest request , HttpServletResponse response){
		this.servlet = servlet;
		this.request = request;
		this.response = response;
		session = request.getSession();
		
		topmodel.set(servlet, request, response);
		
		singleton = Singleton.getInstance(servlet.getServletContext());
		//if(CorpMap == null){
			//CorpMap = (HashMap) servlet.getServletContext().getAttribute("CorpDB");
			//CorpStmt = (Statement) CorpMap.get("Statement");
			
			DBMap = (HashMap) servlet.getServletContext().getAttribute("DBMap");
			
			//DB = (HashMap) servlet.getServletContext().getAttribute("DB");
			//DBStmt = (Statement) DB.get("Statement");
			
			DBStmt = singleton.getDBStatement((String)session.getAttribute("user_nbr"));
			CorpMap = singleton.DBMap;
			DBMap = singleton.DBMap;
			CorpStmt = singleton.getDB2Statement((String)session.getAttribute("user_nbr"));
		//}
	}
	
	public String getDate(){
		String date= 
			calendar.get(Calendar.YEAR) + "-"+(calendar.get(Calendar.MONTH)+1) + 
			"-" + calendar.get(Calendar.DAY_OF_MONTH);
		return date;
	}
	
	//用 TripFlowList 查詢 群組 user_nbr 的 資料  以 ArrayList 回傳
	public ArrayList getUserGroups(ArrayList TripFlowList , String user_lv){
		ArrayList arr = new ArrayList();
		HashMap map = null;
		for(int i = 0 ; i < TripFlowList.size() ; i++){
			map = (HashMap) TripFlowList.get(i);
			arr.add(getUser((String)map.get(user_lv)));
		}
		return arr;
	}
	
	//用 TripFlowList 查詢 群組 user_nbr 的 資料  以 ArrayList 回傳
	public ArrayList getUserGroups(ArrayList TripFlowList){
		ArrayList arr = new ArrayList();
		HashMap map = null;
		for(int i = 0 ; i < TripFlowList.size() ; i++){
			map = (HashMap) TripFlowList.get(i);
			arr.add(getUser((String)map.get("request_nbr")));
		}
		return arr;
	}
	
	//取出 指定 user_nbr 的 全部資訊 用HashMap
	public HashMap getUser(String user_nbr){
		String sql = "SELECT * FROM user where " + 
			"user_nbr='"+user_nbr+ "' and corp_id='"+session.getAttribute("user_corp_id")+"'";
		try {
			ResultSet rs = DBStmt.executeQuery(sql);
			String s[] = {
					"user_nbr" , "login" , "pswd" , "corp_id" , 
					"cname" , "lastname" , "firstname" , "id" , 
					"phone" , "role" , "rank" , "authority" , "date_start" , 
					"date_end" , "date_ack" , "approver" , "controller"};
			HashMap map = null;
			rs.absolute(-1);//移動至最後一筆資料
			map = new HashMap();
			for(int i = 0 ; i < s.length ; i++){
				map.put(s[i], rs.getString(s[i]));
			}
			if(map != null){
				String approver = rs.getString("approver");
				String controller = rs.getString("controller");
				map.put("approver_mail", getUserEmail(approver));
				map.put("controller_mail", getUserEmail(controller));
			}
			
			return map;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	//取出 指定 user_nbr 的 email
	public String getUserEmail(String user_nbr){
		String sql = "SELECT * FROM user where " + 
			"user_nbr='"+user_nbr+ "' and corp_id='"+session.getAttribute("user_corp_id")+"'";
		try {
			ResultSet rs = DBStmt.executeQuery(sql);
			String mail = null;
			while(rs.next()){
				mail = rs.getString("login");
			}
			return mail;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	//取出 指定 email 的 user_nbr
	public String getUserNBR(String login){
		String sql = "SELECT * FROM user where " + 
			"login='"+login+ "' and corp_id='"+session.getAttribute("user_corp_id")+"'";
		try {
			ResultSet rs = DBStmt.executeQuery(sql);
			String a_nbr = null;
			while(rs.next()){
				a_nbr = rs.getString("user_nbr");
			}
			return a_nbr;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	//以Trip_nbr取出 指定使用者 的 行程事件表
	public ArrayList getTripFlow_Trip(String Trip_nbr ){
		/*
		V tripflow_nbr (行程批核編號)
		V request_nbr (申請者編號)
		V aprover_nbr (批核決行者編號)
		*/
		String sql;
		try {
			ResultSet rs;
			sql = "SELECT * FROM tripflow where trip_nbr='"+Trip_nbr+"'";
			rs = DBStmt.executeQuery(sql);
			String[] s = {
					"tripflow_nbr","request_nbr","aprover_nbr","trip_nbr","tripflow_status",
					"tripflow_time_apply","tripflow_time_process"};
			ArrayList arr = new ArrayList();
			HashMap map;
			while(rs.next()){
				map = new HashMap();
				for(int i = 0 ; i < s.length ; i++){
					map.put(s[i], rs.getString(s[i]));
				}
				arr.add(map);
			}
			return arr;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	//以aprover_nbr(user_nbr)取出 指定使用者 的 行程事件表
	public ArrayList getTripFlow(String aprover_nbr ){
		/*
		V tripflow_nbr (行程批核編號)
		V request_nbr (申請者編號)
		V aprover_nbr (批核決行者編號)
		*/
		String sql;
		try {
			ResultSet rs;
			sql = "SELECT * FROM tripflow where aprover_nbr='"+aprover_nbr+"'";
			rs = DBStmt.executeQuery(sql);
			String[] s = {
					"tripflow_nbr","request_nbr","aprover_nbr","trip_nbr","tripflow_status",
					"tripflow_time_apply","tripflow_time_process"};
			ArrayList arr = new ArrayList();
			HashMap map;
			while(rs.next()){
				map = new HashMap();
				for(int i = 0 ; i < s.length ; i++){
					map.put(s[i], rs.getString(s[i]));
				}
				arr.add(map);
			}
			return arr;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	//以aprover_nbr(user_nbr)取出 指定使用者 的 行程事件表 並指定狀態
	public ArrayList getTripFlow(String aprover_nbr , String status){
		/*
		V tripflow_nbr (行程批核編號)
		V request_nbr (申請者編號)
		V aprover_nbr (批核決行者編號)
		*/
		String sql;
		try {
			ResultSet rs;
			
			if(status.equals("1") || status.equals("2")){
				sql = "SELECT * FROM tripflow where " +
					"aprover_nbr='"+aprover_nbr+"' and (tripflow_status='1' or tripflow_status='2')";
			}else if(status.equals("3") || status.equals("4")){
				sql = "SELECT * FROM tripflow where " +
					"aprover_nbr='"+aprover_nbr+"' and (tripflow_status='3' or tripflow_status='4')";
			}else if(status.equals("5")){
				sql = "SELECT * FROM tripflow where " +
					"aprover_nbr='"+aprover_nbr+"' and tripflow_status='5'";
			}else{
				sql = "SELECT * FROM tripflow where aprover_nbr='"+aprover_nbr+"'";
			}
			
			rs = DBStmt.executeQuery(sql);
			String[] s = {
					"tripflow_nbr","request_nbr","aprover_nbr","trip_nbr","tripflow_status",
					"tripflow_time_apply","tripflow_time_process"};
			ArrayList arr = new ArrayList();
			HashMap map;
			while(rs.next()){
				map = new HashMap();
				for(int i = 0 ; i < s.length ; i++){
					map.put(s[i], rs.getString(s[i]));
				}
				arr.add(map);
			}
			return arr;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	//修改 行程事件
	public void editTripFlow(String tripflow_nbr , String trip_nbr , String tripflow_status ){
		/*
		V tripflow_nbr (行程批核編號)
		V request_nbr (申請者編號)
		V aprover_nbr (批核決行者編號)
		V trip_nbr (行程編號)
		V tripflow_status (狀態--待批核:1待決行:2核准:3決行:4取消:5)
		V tripflow_time_apply (事件申請時間)
		V tripflow_time_process (事件處理時間)
		*/
		String sql;
		ResultSet rs;
		try {
			//更新行程(tripflow)事件
			sql = "SELECT * FROM tripflow " + "where tripflow_nbr='"+tripflow_nbr+"' and trip_nbr='"+trip_nbr+"'";
			rs = DBStmt.executeQuery(sql);
			//System.out.println(sql);
			String r_nbr = "";
			//
			while(rs.next()){
				r_nbr = rs.getString("aprover_nbr");
			}
			
			rs = DBStmt.executeQuery(sql);
			rs.last();
			rs.updateString("tripflow_status",tripflow_status);//(狀態--待批核:1待決行:2核准:3決行:4取消:5)
			rs.updateString("tripflow_time_process",getDate());//(事件處理時間)
			rs.updateRow();//修改資料
			
			
			
			//更新行程(trip)資料
			if(tripflow_status.equals("1") || tripflow_status.equals("3") || 
					tripflow_status.equals("4") || tripflow_status.equals("5")){
				sql = "SELECT * FROM trip where trip_nbr='"+trip_nbr+"'";
				rs = DBStmt.executeQuery(sql);
				rs.last();
				if(tripflow_status.equals("1")){
					//申請
					rs.updateString("time_apply",getDate());//申請時間
					rs.updateString("trip_status","1");//(行程狀態--申請中:1批核中:2已決行:3出差中:4結案:5被退回:6)
				}else if(tripflow_status.equals("3")){
					//核准
					rs.updateString("time_agree",getDate());//核准時間
					rs.updateString("agree_nbr",(String)session.getAttribute("user_nbr"));//批何者編號
					rs.updateString("trip_status","2");//(行程狀態--申請中:1批核中:2已決行:3出差中:4結案:5被退回:6)
				}else if(tripflow_status.equals("4")){
					//准行
					rs.updateString("time_final",getDate());//准行時間
					rs.updateString("controller_nbr",(String)session.getAttribute("user_nbr"));//決行者編號
					rs.updateString("trip_status","3");//(行程狀態--申請中:1批核中:2已決行:3出差中:4結案:5被退回:6)
					
					//修改 已決行預算
					ArrayList ub =  topmodel.getUserBudget((String)session.getAttribute("user_nbr"));
					String budget_nbr = (String)((HashMap)ub.get(ub.size()-1)).get("budget_nbr");
					String user_nbr = (String)((HashMap)ub.get(ub.size()-1)).get("user_nbr");
					int n = (rs.getString("expense_estimate")!=null && !rs.getString("expense_estimate").equals("")?
							Integer.parseInt(rs.getString("expense_estimate")): 0) 
							+ (rs.getString("tkt_estim")!=null && !rs.getString("tkt_estim").equals("")?
									Integer.parseInt(rs.getString("tkt_estim")) : 0);
					topmodel.editUserBudgetGrant(budget_nbr, user_nbr, n);
				}else if(tripflow_status.equals("5")){
					//退回
					rs.updateString("time_agree",getDate());//核准時間
					rs.updateString("trip_status","6");//(行程狀態--申請中:1批核中:2已決行:3出差中:4結案:5被退回:6)
				}
				rs.updateRow();
				
				
				//更新User
				String a_nbr = "";
				if(getUserNBR(request.getParameter("nextUser")) != null)
					a_nbr = getUserNBR(request.getParameter("nextUser"));
				if(!a_nbr.equals("") && !r_nbr.equals("")){
					sql = "SELECT * FROM user where user_nbr='"+r_nbr+"'";
					rs = DBStmt.executeQuery(sql);
					while(rs.next()){
						rs.last();
						rs.updateString("approver",a_nbr);//寫入上次批何者號
						rs.updateRow();
					}
				}
			}
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}
	
	//新增 行程事件
	public void newTripFlow(
			String request_nbr , //(申請者編號)
			String aprover_nbr , //(批核決行者編號)
			String trip_nbr , //(行程編號)
			String tripflow_status, //(狀態--待批核:1待決行:2核准:3決行:4取消:5)
			boolean isUpApprover){//判斷是否更新  uesr table 的 approver
		/*
		V tripflow_nbr (行程批核編號)
		V request_nbr (申請者編號)
		V aprover_nbr (批核決行者編號)
		V trip_nbr (行程編號)
		V tripflow_status (狀態--待批核:1待決行:2核准:3決行:4取消:5)
		V tripflow_time_apply (事件申請時間)
		V tripflow_time_process (事件處理時間)
		*/
		String sql;
		try {
			ResultSet rs;
			
			String a_nbr = getUserNBR(aprover_nbr);
			
			/*
			//查詢是否已發出
			sql = "SELECT * FROM tripflow where " +
					"request_nbr='"+request_nbr+"' and "+
					"aprover_nbr='"+a_nbr+"' and "+
					"trip_nbr='"+trip_nbr+"'";
			rs = DBStmt.executeQuery(sql);
			String status = "";
			while(rs.next()){
				status = rs.getString("tripflow_status");
			}
			*/
			//需未發出  或者  已經退回才可再次發出
			//if(status.equals("") || status.equals("5")){
				//新增 行程事件
				if(a_nbr != null && !a_nbr.equals("")){
					sql = "SELECT * FROM tripflow Limit 0,1";
					rs = DBStmt.executeQuery(sql);
					rs.moveToInsertRow();
					rs.updateString("request_nbr",request_nbr);//(申請者編號)
					rs.updateString("aprover_nbr",a_nbr);//(批核決行者編號)
					rs.updateString("trip_nbr",trip_nbr);//(行程編號)
					rs.updateString("tripflow_status",tripflow_status);//(狀態--待批核:1待決行:2核准:3決行:4取消:5)
					rs.updateString("tripflow_time_apply",getDate());//(事件申請時間)
					rs.insertRow();//新增資料
				}
			//}
			
			//更新行程(trip) 以及 User資料
			if(tripflow_status.equals("1") || tripflow_status.equals("2")){
				//更新行程
				sql = "SELECT * FROM trip where trip_nbr='"+trip_nbr+"'";
				rs = DBStmt.executeQuery(sql);
				rs.last();
				//申請
				rs.updateString("time_apply",getDate());//申請時間
				rs.updateString("trip_status","2");
				//(行程狀態--申請中:1批核中:2已決行:3出差中:4結案:5被退回:6)
				rs.updateRow();
				
				if(isUpApprover){
					//更新User
					sql = "SELECT * FROM user where user_nbr='"+request_nbr+"'";
					rs = DBStmt.executeQuery(sql);
					rs.last();
					rs.updateString("approver",a_nbr);//寫入上次批何者號
					rs.updateRow();
					//System.out.println(a_nbr);
				}
			}
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}
	
	//用 TripFlowList 查詢 群組 trip_nbr 的 資料  以 ArrayList 回傳
	public ArrayList getTripGroups(ArrayList TripFlowList){
		ArrayList arr = new ArrayList();
		HashMap map = null;
		for(int i = 0 ; i < TripFlowList.size() ; i++){
			map = (HashMap) TripFlowList.get(i);
			arr.add(getTripMap((String)map.get("request_nbr"),(String)map.get("trip_nbr")));
		}
		return arr;
	}
	
	//查詢 某 trip_nbr 的 資料  以HashMap回傳
	public HashMap getTripMap(String usr_nbr , String trip_nbr){
		String sql = "SELECT * FROM trip where usr_nbr='"+usr_nbr+"' and trip_nbr='"+trip_nbr+"'";
		//System.out.println(sql);
		try {
			ResultSet rs = DBStmt.executeQuery(sql);
			String[] s = {
					"trip_nbr","usr_nbr","corp_id","trip_description",
					"trip_note","dep_date","rtn_date","trip_interval",
					"trip_pnr","date_ack","trip_status","time_create",
					"time_apply","time_agree","time_final","time_close",
					"close_describe","agree_nbr","controller_nbr","tkt_estim",
					"hotel_estim","traffic_estim","meal_estim","extra_estim",
					"other_estim","expense_estimate","tkt","hotel","traffic","meal",
					"extra","other","expense_actual"};
			HashMap map = new HashMap();
			while(rs.next()){
				for(int i = 0 ; i < s.length ; i++){
					map.put(s[i], rs.getString(s[i]));
				}
			}
			return map;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	//查詢 某 trip_nbr 的 資料  以ResultSet回傳
	public ResultSet getTrip(String usr_nbr , String trip_nbr){
		String sql = "SELECT * FROM trip where usr_nbr='"+usr_nbr+"' and trip_nbr='"+trip_nbr+"'";
		//System.out.println(sql);
		try {
			ResultSet rs = DBStmt.executeQuery(sql);
			//System.out.println("getTrip");
			return rs;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	public ResultSet getTrip(String usr_nbr){
		String sql = "SELECT * FROM trip where usr_nbr='"+usr_nbr+"'";
		try {
			ResultSet rs = DBStmt.executeQuery(sql);
			return rs;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	//停止 行程trip
	public void stopTrip(String trip_nbr , String close_describe){
		String sql;
		try {
			sql = "SELECT * FROM trip where trip_nbr='"+trip_nbr+"'";
			ResultSet rs = DBStmt.executeQuery(sql);
			rs.last();
			rs.updateString("trip_status","5");//(行程狀態--申請中:1 批核中:2 已決行:3 出差中:4 結案:5 被退回:6)
			rs.updateString("close_describe",close_describe);//(結案原因--報帳結案/取消byxxxx)
			rs.updateString("time_close",getDate());//(結案時間--結案狀態開始)
			rs.updateRow();//更新資料
			rs.close();
			
			sql = "SELECT * FROM tripflow " + "where trip_nbr='"+trip_nbr+"'";
			rs = DBStmt.executeQuery(sql);
			while(rs.next()){
				if(rs.getString("tripflow_status").equals("1") 
						|| rs.getString("tripflow_status").equals("2")){
					rs.last();
					rs.updateString("tripflow_status", "0");
					rs.updateRow();//更新資料
				}
			}
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}
	
	//更新 行程trip
	public String editTrip(String trip_nbr){
		//處理接值
		String usr_nbr = (String) session.getAttribute("user_nbr");//(用戶編號)
		String corp_id = (String) session.getAttribute("user_corp_id");//(公司編號)
		String trip_description = request.getParameter("trip_description");//(行程目的描述)
		String trip_note = request.getParameter("trip_note");//(行程附言)
		String dep_date ;//(出發日)
		String rtn_date ;//(返回日)
		String trip_interval = request.getParameter("air_interval");//(出差天數預設是rtn_date~dep_date，可容修改)
		String trip_pnr = request.getParameter("air_pnt");//(機票pnr)
		String date_ack;//(上次異動時間)
		String trip_status = request.getParameter("submit");//(行程狀態--申請中:1批核中:2決行中:3出差中:4結案:5被退回:6)
		String[] t_status = {"存回","申請","取消"};//行程狀態  狀態機
		String time_create;//(立案時間--申請狀態開始)
		String time_apply ;//(申請時間--批核狀態開始)
		String time_agree ;//(批核時間--決行狀態開始)
		String time_final ;//(決行時間--出差狀態開始)
		String time_close ;//(結案時間--結案狀態開始)
		String agree_nbr = request.getParameter("nextUser");//(批核者編號)
		String controller_nbr ;//( 決行者編號 )
		String tkt_estim = request.getParameter("air_estim");//(機票費用預估)
		String hotel_estim = request.getParameter("sum_estim_hotel");//(住宿費用預估)
		String traffic_estim = request.getParameter("sum_estim_traffic");//(交通費用預估)
		String meal_estim = request.getParameter("sum_estim_meal");//(膳食費用預估)
		String extra_estim = request.getParameter("sum_estim_extra");//(雜項費用預估)
		String other_estim = request.getParameter("sum_estim_other");//(其他費用預估)
		String expense_estimate = request.getParameter("sum_estim_totle");//(總費用預估)
		
		//更新 trip 資料
		date_ack = getDate();//(上次異動時間)
		String sql = "SELECT * FROM trip where trip_nbr='"+trip_nbr+"'";
		//System.out.println(sql);
		try {
			ResultSet rs = DBStmt.executeQuery(sql);
			rs.last();
			rs.updateString("usr_nbr",usr_nbr);//(用戶編號)
			rs.updateString("corp_id",corp_id);//(公司編號)
			rs.updateString("trip_description",trip_description);//(行程目的描述)
			rs.updateString("trip_note",trip_note);//(行程附言)
			//rs.updateString("dep_date",dep_date) ;//("出發日)
			//rs.updateString("rtn_date",rtn_date) ;//("返回日)
			rs.updateString("trip_interval",trip_interval);//(出差天數預設是rtn_date~dep_date，可容修改)
			rs.updateString("trip_pnr",trip_pnr);//(機票pnr))
			rs.updateString("date_ack",date_ack);//(上次異動時間)
			//rs.updateString("agree_nbr",agree_nbr);//(批核者編號)
			//rs.updateString("controller_nbr",controller_nbr);//(決行者編號 )
			rs.updateString("tkt_estim",tkt_estim);//(機票費用預估)
			rs.updateString("hotel_estim",hotel_estim);//(住宿費用預估)
			rs.updateString("traffic_estim",traffic_estim);//(交通費用預估)
			rs.updateString("meal_estim",meal_estim);//(膳食費用預估)
			rs.updateString("extra_estim",extra_estim);//(雜項費用預估)
			rs.updateString("other_estim",other_estim);//(其他費用預估)
			rs.updateString("expense_estimate",expense_estimate);//(總費用預估)
			rs.updateRow();//更新資料
			rs.close();
			return trip_nbr;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return "";
	}
	
	//新增行程trip
	public String newTrip(){
		//處理接值
		String usr_nbr = (String) session.getAttribute("user_nbr");//(用戶編號)
		String corp_id = (String) session.getAttribute("user_corp_id");//(公司編號)
		String trip_description = request.getParameter("trip_description");//(行程目的描述)
		String trip_note = request.getParameter("trip_note");//(行程附言)
		String dep_date ;//(出發日)
		String rtn_date ;//(返回日)
		String trip_interval = request.getParameter("air_interval");//(出差天數預設是rtn_date~dep_date，可容修改)
		String trip_pnr = request.getParameter("air_pnt");//(機票pnr)
		String date_ack;//(上次異動時間)
		String trip_status = request.getParameter("submit");//(行程狀態--申請中:1批核中:2決行中:3出差中:4結案:5被退回:6)
		String[] t_status = {"存回","申請","取消"};//行程狀態  狀態機
		String time_create;//(立案時間--申請狀態開始)
		String time_apply ;//(申請時間--批核狀態開始)
		String time_agree ;//(批核時間--決行狀態開始)
		String time_final ;//(決行時間--出差狀態開始)
		String time_close ;//(結案時間--結案狀態開始)
		String agree_nbr = request.getParameter("nextUser");//(批核者編號)
		String controller_nbr ;//( 決行者編號 )
		String tkt_estim = request.getParameter("air_estim");//(機票費用預估)
		String hotel_estim = request.getParameter("sum_estim_hotel");//(住宿費用預估)
		String traffic_estim = request.getParameter("sum_estim_traffic");//(交通費用預估)
		String meal_estim = request.getParameter("sum_estim_meal");//(膳食費用預估)
		String extra_estim = request.getParameter("sum_estim_extra");//(雜項費用預估)
		String other_estim = request.getParameter("sum_estim_other");//(其他費用預估)
		String expense_estimate = request.getParameter("sum_estim_totle");//(總費用預估)
		
		//新增 trip 資料
		time_create = getDate();//(立案時間--申請狀態開始)
		date_ack = getDate();//(上次異動時間)
		String sql = "SELECT * FROM trip Limit 0,1";
		try {
			ResultSet rs = DBStmt.executeQuery(sql);
			rs.moveToInsertRow();
			rs.updateString("usr_nbr",usr_nbr);//(用戶編號)
			rs.updateString("corp_id",corp_id);//(公司編號)
			rs.updateString("trip_description",trip_description);//(行程目的描述)
			rs.updateString("trip_note",trip_note);//(行程附言)
			//rs.updateString("dep_date",dep_date) ;//("出發日)
			//rs.updateString("rtn_date",rtn_date) ;//("返回日)
			rs.updateString("trip_interval",trip_interval);//(出差天數預設是rtn_date~dep_date，可容修改)
			rs.updateString("trip_pnr",trip_pnr);//(機票pnr))
			rs.updateString("date_ack",date_ack);//(上次異動時間)
			rs.updateString("trip_status","1");//(行程狀態--申請:1批核:2決行:3出差:4結案:5)
			rs.updateString("time_create",time_create);//(立案時間--申請狀態開始)
			//rs.updateString("agree_nbr",agree_nbr);//(批核者編號)
			//rs.updateString("controller_nbr",controller_nbr);//(決行者編號 )
			rs.updateString("tkt_estim",tkt_estim);//(機票費用預估)
			rs.updateString("hotel_estim",hotel_estim);//(住宿費用預估)
			rs.updateString("traffic_estim",traffic_estim);//(交通費用預估)
			rs.updateString("meal_estim",meal_estim);//(膳食費用預估)
			rs.updateString("extra_estim",extra_estim);//(雜項費用預估)
			rs.updateString("other_estim",other_estim);//(其他費用預估)
			rs.updateString("expense_estimate",expense_estimate);//(總費用預估)
			rs.insertRow();//新增資料
			
			sql = "SELECT * FROM trip where usr_nbr='"+usr_nbr+"'";
			rs = DBStmt.executeQuery(sql);
			ArrayList<String> arr = new ArrayList();
			while(rs.next()){
				arr.add(rs.getString("trip_nbr"));
			}
			rs.close();
			return (String) arr.get(arr.size()-1);
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return "";
	}
	
	//查詢 某 trip_nbr 的 航段資料  以ArrayList回傳
	public ArrayList getTripRouteArr(String trip_nbr){
		String sql = "SELECT * FROM triproute where trip_nbr='"+trip_nbr+"'";
		//System.out.println(sql);
		try {
			ResultSet rs = DBStmt.executeQuery(sql);
			//System.out.println("getTripRoute");
			ArrayList arr = new ArrayList();
			String[] s = {
					"rout_nbr","trip_nbr","rout_note","date_rout_dep",
					"rout_dep_port","rout_ariv_port","rout_flight","rout_transfer"};
			HashMap map;
			while(rs.next()){
				map = new HashMap();
				for(int i = 0 ; i < s.length ; i++){
					map.put(s[i], rs.getString(s[i]));
				}
				arr.add(map);
			}
			return arr;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	//查詢 某 trip_nbr 的 航段資料  以ResultSet回傳
	public ResultSet getTripRoute(String trip_nbr){
		String sql = "SELECT * FROM triproute where trip_nbr='"+trip_nbr+"'";
		//System.out.println(sql);
		try {
			ResultSet rs = DBStmt.executeQuery(sql);
			//System.out.println("getTripRoute");
			return rs;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	//新增行程航段資料
	public void newTripRoute(String trip_nbr){
		//處理接值
		//String[] rout_note = request.getParameterValues("air_Date");//(航段附言)
		String[] date_rout_dep = request.getParameterValues("air_Date");//(出發日期)
		String[] rout_dep_port = request.getParameterValues("air_Origin");//(出發機場通常是桃園機場)
		String[] rout_ariv_port = request.getParameterValues("air_Destination");//(落地機場例如是上海浦東機場)
		String[] rout_flight = request.getParameterValues("air_Flight");//(航班)
		String[] rout_transfer = request.getParameterValues("air_transfer");//(轉機)
		
		//
		String sql = "";
		try {
			//先刪除所有資料
			sql = "SELECT * FROM triproute where trip_nbr='"+trip_nbr+"'";
			ResultSet rs = DBStmt.executeQuery(sql);
			if(rs != null){
				sql = "DELETE FROM `triproute` WHERE `trip_nbr`='"+trip_nbr+"'";
				//System.out.println(sql);
				DBStmt.execute(sql);
			}
			//依序新增
			sql = "SELECT * FROM triproute where trip_nbr='"+trip_nbr+"'";
			rs = DBStmt.executeQuery(sql);
			for(int i = 0 , j = 0 ; i < date_rout_dep.length ; i++){
				if(date_rout_dep[i] != null && !date_rout_dep[i].equals("") && !date_rout_dep[i].equals(" ")){
					rs.moveToInsertRow();
					rs.updateString("trip_nbr",trip_nbr);//行程編號
					rs.updateString("date_rout_dep",date_rout_dep[i]);//(出發日期)
					rs.updateString("rout_dep_port",rout_dep_port[i]);//(出發機場通常是桃園機場)
					rs.updateString("rout_ariv_port",rout_ariv_port[i]);//(落地機場例如是上海浦東機場)
					rs.updateString("rout_flight",rout_flight[i]);//(航班)
					if(rout_transfer!= null && j < rout_transfer.length &&
							rout_transfer[j] != null && rout_transfer[j].equals(Integer.toString(i))){
						rs.updateString("rout_transfer","y");//(轉機)
						j++;
					}else{
						rs.updateString("rout_transfer","n");//(轉機)
					}
					rs.insertRow();//新增資料
				}
			}
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}
	
	//新增 報帳資訊(actual_expense)
	public void newTripActual_ExpenseArr(String trip_nbr){
		String[] day_nbr = request.getParameterValues("actual_day_nbr");//(日程序號)
		String[] expense_date = request.getParameterValues("actual_expense_date");//(日期)
		String[] expense_city = request.getParameterValues("actual_expense_city");//(出差城市)
		String[] date_ack = request.getParameterValues("actual_date_ack");//(上次異動時間)
		String[] estim_hotel = request.getParameterValues("actual_estim_hotel");//(住宿費用)
		String[] expense_traffic = request.getParameterValues("actual_expense_traffic");//(交通費用)
		String[] expense_meal = request.getParameterValues("actual_expense_meal");//(膳食費用)
		String[] expense_extra = request.getParameterValues("actual_expense_extra");//(雜項費用)
		String[] expense_other = request.getParameterValues("actual_expense_other");//(其他費用)
		String[] expense_sum = request.getParameterValues("actual_expense_sum");//(費用)
		String[] expense_receipt = request.getParameterValues("actual_expense_receipt");//(檢付收據--張數)	
		
		String sql = "SELECT * FROM actual_expense where trip_nbr='"+trip_nbr+"'";
		//System.out.println(sql);
		try {
			ResultSet rs;
			ArrayList arr = new ArrayList();
			HashMap map;
			//先刪除所有資料
			sql = "SELECT * FROM actual_expense where trip_nbr='"+trip_nbr+"'";
			rs = DBStmt.executeQuery(sql);
			if(rs != null){
				sql = "DELETE FROM `actual_expense` WHERE `trip_nbr`='"+trip_nbr+"'";
				//System.out.println(sql);
				DBStmt.execute(sql);
			}
			//依序新增
			sql = "SELECT * FROM actual_expense where trip_nbr='"+trip_nbr+"'";
			rs = DBStmt.executeQuery(sql);
			for(int i = 0 , j = 0 ; i < expense_date.length ; i++){
				if(expense_date[i] != null && !expense_date[i].equals("") && !expense_date[i].equals(" ")){
					rs.moveToInsertRow();
					rs.updateString("trip_nbr", trip_nbr);//
					rs.updateString("day_nbr", day_nbr[i]);//(日程序號)
					rs.updateString("expense_date", expense_date[i]);//(日期)
					rs.updateString("expense_city", expense_city[i]);//(出差城市)
					rs.updateString("date_ack", getDate());//(上次異動時間)
					rs.updateString("estim_hotel", estim_hotel[i]);//(住宿費用)
					rs.updateString("expense_traffic", expense_traffic[i]);//(交通費用)
					rs.updateString("expense_meal", expense_meal[i]);//(膳食費用)
					rs.updateString("expense_extra", expense_extra[i]);//(雜項費用)
					rs.updateString("expense_other", expense_other[i]);//(雜項費用)
					rs.updateString("expense_sum", expense_sum[i]);//(費用總額)
					rs.updateString("expense_receipt", expense_receipt[i]);//(檢付收據--張數)
					rs.insertRow();//新增資料
				}
			}
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}
	
	//查詢 某 trip_nbr的 報帳資訊(actual_expense)  以ArrayList回傳
	public ArrayList getTripActual_ExpenseArr(String trip_nbr){
		String sql = "SELECT * FROM actual_expense where trip_nbr='"+trip_nbr+"'";
		//System.out.println(sql);
		try {
			ResultSet rs = DBStmt.executeQuery(sql);
			ArrayList arr = new ArrayList();
			HashMap map;
			while(rs.next()){
				map = new HashMap();
				map.put("expense_nbr", rs.getString("expense_nbr"));//(費用編號)
				map.put("trip_nbr", rs.getString("trip_nbr"));//(行程編號)
				map.put("day_nbr", rs.getString("day_nbr"));//(日程序號)
				map.put("expense_date", rs.getString("expense_date"));//(日期)
				map.put("expense_city", rs.getString("expense_city"));//(出差城市)
				map.put("expense_note", rs.getString("expense_note"));//((附註))
				map.put("date_ack", rs.getString("date_ack"));//(上次異動時間)
				map.put("estim_hotel", rs.getString("estim_hotel"));//(住宿費用)
				map.put("expense_traffic", rs.getString("expense_traffic"));//(交通費用)
				map.put("expense_meal", rs.getString("expense_meal"));//(膳食費用)
				map.put("expense_extra", rs.getString("expense_extra"));//(雜項費用)
				map.put("expense_other", rs.getString("expense_other"));//(其他費用)
				map.put("expense_sum", rs.getString("expense_sum"));//(費用總額)
				map.put("expense_receipt", rs.getString("expense_receipt"));//(檢付收據--張數)
				arr.add(map);
			}
			return arr;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	//查詢 某 trip_nbr的 航段預算  以ArrayList回傳
	public ArrayList getTripEstim_ExpenseArr(String trip_nbr){
		String sql = "SELECT * FROM estim_expense where trip_nbr='"+trip_nbr+"'";
		//System.out.println(sql);
		try {
			ResultSet rs = DBStmt.executeQuery(sql);
			//System.out.println("getTripEstim_Expense");
			ArrayList arr = new ArrayList();
			String[] s = {
					"estim_nbr","trip_nbr","rout_nbr","duration",
					"expense_city","expense_note","date_ack","estim_hotel",
					"estim_traffic","estim_meal","estim_extra","estim_other"};
			HashMap map;
			while(rs.next()){
				map = new HashMap();
				for(int i = 0 ; i < s.length ; i++){
					map.put(s[i], rs.getString(s[i]));
				}
				arr.add(map);
			}
			return arr;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	//查詢 某 trip_nbr的 航段預算  以ResultSet回傳
	public ResultSet getTripEstim_Expense(String trip_nbr){
		String sql = "SELECT * FROM estim_expense where trip_nbr='"+trip_nbr+"'";
		//System.out.println(sql);
		try {
			ResultSet rs = DBStmt.executeQuery(sql);
			//System.out.println("getTripEstim_Expense");
			return rs;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	//新增 航段 預估費用資料
	public void newTripEstim_Expense(String trip_nbr){
		//處理接值
		String[] estim_nbr ;//(預估費用編號)]
		String[] rout_nbr;//(行程段序號)
		String[] duration = request.getParameterValues("estim_duration");//(停留天數)
		String[] expense_city = request.getParameterValues("expense_city");//(出差城市)
		String[] expense_note = request.getParameterValues("expense_note");//(附註)
		String[] date_ack ;//(上次異動時間)
		String[] estim_hotel = request.getParameterValues("estim_hotel");//(住宿費用預估一天)
		String[] estim_traffic = request.getParameterValues("estim_traffic");//(交通費用預估一天)
		String[] estim_meal = request.getParameterValues("estim_meal");//(膳食費用預估一天)
		String[] estim_extra = request.getParameterValues("estim_extra");//(雜項費用預估一天)
		String[] estim_other = request.getParameterValues("estim_other");//(其他費用預估一天)
		
		String sql = "";
		sql = "SELECT * FROM triproute where trip_nbr='"+trip_nbr+"'";
		try {
			//取得航段序號
			ResultSet rs = DBStmt.executeQuery(sql);
			ArrayList<String> arr = new ArrayList();
			while(rs.next()){
				arr.add(rs.getString("rout_nbr"));
			}
			//清掉該行程所有航段預算
			sql = "SELECT * FROM estim_expense where trip_nbr='"+trip_nbr+"'";
			rs = DBStmt.executeQuery(sql);
			if(rs != null){
				sql = "DELETE FROM `estim_expense` WHERE (`trip_nbr`='"+trip_nbr+"')";
				//System.out.println(sql);
				DBStmt.execute(sql);
			}
			//依序新增
			sql = "SELECT * FROM estim_expense where trip_nbr='"+trip_nbr+"'";
			rs = DBStmt.executeQuery(sql);
			for(int i = 0 ; i < duration.length ; i++){
				if(duration[i] != null && !duration[i].equals("") && !duration[i].equals(" ")){
					rs.moveToInsertRow();
					rs.updateString("trip_nbr",trip_nbr);//行程編號
					rs.updateString("rout_nbr",Integer.toString(i));//行程段序號
					rs.updateString("expense_city",expense_city[i]);//停留城市
					rs.updateString("duration",duration[i]);//停留天數
					rs.updateString("expense_note",expense_note[i]);//附註
					rs.updateString("date_ack",getDate());//上次異動時間
					rs.updateString("estim_hotel",estim_hotel[i]);//住宿費用預估一天
					rs.updateString("estim_traffic",estim_traffic[i]);//交通費用預估一天
					rs.updateString("estim_meal",estim_meal[i]);//膳食費費用預估一天
					rs.updateString("estim_extra",estim_extra[i]);//雜項費用預估一天
					rs.updateString("estim_other",estim_other[i]);//其他費用預估一天
					rs.insertRow();//新增資料
				}
			}
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (IndexOutOfBoundsException e) {
			e.printStackTrace();
		}
	}
}

package com.model;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Calendar;
import java.util.HashMap;

import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import com.web.Singleton;

public class ManagerModel {
	HttpServletRequest request = null;
	HttpServletResponse response = null;
	HttpServlet servlet = null;
	HttpSession session = null;
	
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
		
		singleton = Singleton.getInstance(servlet.getServletContext());
		//if(CorpMap == null){
			//CorpMap = (HashMap) servlet.getServletContext().getAttribute("CorpDB");
			//CorpStmt = (Statement) CorpMap.get("Statement");
			
			//DBMap = (HashMap) servlet.getServletContext().getAttribute("DBMap");
			
			//DB = (HashMap) servlet.getServletContext().getAttribute("DB");
			//DBStmt = (Statement) DB.get("Statement");
			
			DBStmt = singleton.getDBStatement((String)session.getAttribute("user_nbr"));
			CorpMap = singleton.DBMap;
			DBMap = singleton.DBMap;
			CorpStmt = singleton.getDB2Statement((String)session.getAttribute("user_nbr"));
		//}
	}
	
	public void UpDataUser(){
		String[] user_nbr = request.getParameterValues("user_nbr");//User系統編號
		String[] login = request.getParameterValues("login");//登入帳號
		String[] login_old = request.getParameterValues("login_old");//原本的帳號
		String[] cname = request.getParameterValues("cname");//中文名
		String[] psw = request.getParameterValues("psw");//密碼
		String[] phone = request.getParameterValues("phone");//電話
		String[] role = request.getParameterValues("role");//職別
		String[] rank = request.getParameterValues("rank");//職級
		String[] authority = request.getParameterValues("authority");//操作等級
		String[] date_end = request.getParameterValues("date_end");//是否停止  有停止的才會在陣列內 需比對資料
		String[] lastname = request.getParameterValues("lastname");//英文名
		String[] firstname = request.getParameterValues("firstname");//英文名
		String[] id = request.getParameterValues("id");//身分證
		try {
			String sql = "";
			int data_end_i = 0;
			for(int i = 0 ; i < user_nbr.length ; i++){
				ResultSet rs ;
				if(login[i] != null && !login[i].equals("")){
					sql = "SELECT * FROM user where " + "login='"+login[i]+ "'";
					rs = DBStmt.executeQuery(sql);
					int cont = 0;
					while(rs.next()){
						cont++;
					}
					if(cont == 0 || login_old[i].equals(login[i])){//沒有重複的mail才修改
						sql = "SELECT * FROM user where " +
							"user_nbr='"+Integer.parseInt(user_nbr[i])+ "'";
						rs = DBStmt.executeQuery(sql);
						while(rs.next()){
							rs.last(); 
							if(login[i] != null && !login[i].equals(""))
								rs.updateString("login",login[i]);//登入帳號
							if(cname[i] != null && !cname[i].equals(""))
								rs.updateString("cname",cname[i]);//中文名
							if(psw[i] != null && !psw[i].equals("") && !psw[i].equals("****"))
								rs.updateString("pswd",psw[i]);//密碼
							if(phone[i] != null && !phone[i].equals(""))
								rs.updateString("phone",phone[i]);//電話
							if(role[i] != null && !role[i].equals(""))
								rs.updateString("role",role[i]);//職別
							if(rank[i] != null && !rank[i].equals(""))
								rs.updateString("rank",rank[i]);//職級
							if(authority[i] != null && !authority[i].equals(""))
								rs.updateString("authority",authority[i]);//操作等級
							if(lastname[i] != null && !lastname[i].equals(""))
								rs.updateString("lastname",lastname[i]);//英文名
							if(firstname[i] != null && !firstname[i].equals(""))
								rs.updateString("firstname",firstname[i]);//英文名
							if(id[i] != null && !id[i].equals(""))
								rs.updateString("id",id[i]);//身分證
							if(date_end != null && data_end_i < date_end.length){
								if(date_end[data_end_i] != null && !date_end[data_end_i].equals("")){
									if(date_end[data_end_i].equals(user_nbr[i])){//停止
										String date= 
											calendar.get(Calendar.YEAR) + "-"+(calendar.get(Calendar.MONTH)+1) + 
											"-" + calendar.get(Calendar.DAY_OF_MONTH);
										rs.updateString("date_end",date);//
										data_end_i++;
									}
								}
							}
							rs.updateRow();//更新user資料
						}
					}
				}
			}
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}
	
	public void newUser(
			String login , String pswd , String authority ,
			String cname , String ename , String phone , String role , String rank ,
			String lastname , String firstname , String id){
		try {
			String sql;
			ResultSet rs;
			sql = "SELECT * FROM user where " + "login='"+login+ "'";
			rs = DBStmt.executeQuery(sql);
			int cont = 0;
			while(rs.next()){
				cont++;
			}
			if(cont == 0){//該使用者沒有在名單內 可以新增該使用者
				sql = "SELECT * FROM user where " +
				"user_nbr='"+Integer.parseInt((String)session.getAttribute("user_nbr"))+ "'";
				rs = DBStmt.executeQuery(sql);
				rs.next();
				String corp_id = rs.getString("corp_id");
				String date= 
					calendar.get(Calendar.YEAR) + "-"+(calendar.get(Calendar.MONTH)+1) + 
					"-" + calendar.get(Calendar.DAY_OF_MONTH);
				
				rs.moveToInsertRow();
				rs.updateString("login", login);
				rs.updateString("pswd", pswd);
				rs.updateString("corp_id", corp_id);
				rs.updateString("authority", authority);
				rs.updateString("date_start",date);
				if(cname != null && !cname.equals(""))
					rs.updateString("cname",cname);
				if(ename != null && !ename.equals(""))
					rs.updateString("ename",cname);
				if(phone != null && !phone.equals(""))
					rs.updateString("phone",phone);
				if(role != null && !role.equals(""))
					rs.updateString("role",role);
				if(rank != null && !rank.equals(""))
					rs.updateString("rank",rank);
				if(lastname != null && !lastname.equals(""))
					rs.updateString("lastname",lastname);
				if(firstname != null && !firstname.equals(""))
					rs.updateString("firstname",firstname);
				if(id != null && !id.equals(""))
					rs.updateString("id",id);
				
				rs.insertRow();
			}
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}
	
	public ResultSet getUserListRs(int start , int num , String Corp_id){
		HashMap TableMap = (HashMap)DBMap.get("user");
		
		try {
			//案 login 排序 小~大
			String sql = "SELECT * FROM user where " +
							"corp_id="+Corp_id+" ORDER BY \"login\" ASC LIMIT "+start+","+num;
			
			ResultSet rs = DBStmt.executeQuery(sql);
			
			return rs;
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}
}

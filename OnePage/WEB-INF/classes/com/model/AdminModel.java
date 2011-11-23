package com.model;

import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;

import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import com.controller.var;
import com.mysql.jdbc.Connection;
import com.web.Singleton;

public class AdminModel {
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
			
			DBStmt = singleton.getDBStatement("hello");
			CorpMap = singleton.DBMap;
			DBMap = singleton.DBMap;
			CorpStmt = singleton.getDB2Statement("hello");
		//}
	}
	
	//開啟服務   corp_id=公司編號不是統編     user_id=申請的管理者mail
	public boolean OnServices(String corp_id , String user_id){
		if(corp_id != null){
			String id = corp_id;
			String sql = "";
			
			String date= 
				calendar.get(Calendar.YEAR) + "-"+(calendar.get(Calendar.MONTH)+1) + 
				"-" + calendar.get(Calendar.DAY_OF_MONTH);
	    	
			//處理同時有N個管理者的申請
			if(user_id != null){
				String mail = user_id.toLowerCase();
				String tmp[] = null;
				if(mail.indexOf(";") != -1){
					tmp = mail.split(";");
				}else if(mail.indexOf(",") != -1){
					tmp = mail.split(",");
				}else if(mail.indexOf(" ") != -1){
					tmp = mail.split(" ");
				}
				if(tmp != null && tmp.length > 0){
					user_id = tmp[0];
					for(int i = 1 ; i < tmp.length ; i++){
						OnServices(corp_id , tmp[i]);
					}
				}
			}
			
			ResultSet rs;
			int cont = 0;
			boolean isUpCorpStartDate = false;//判別 corp 是否更新  新增日期
			try {
				//先找尋使用者
				sql = "SELECT * FROM user where " +
					"login='"+user_id+ "'";
				rs = DBStmt.executeQuery(sql);
				while(rs.next()){
					cont++;
				}
				rs.beforeFirst();
				rs.next();
				if(cont != 0){
					//找到該使用者
					if(rs.getString("corp_id").equals(request.getParameter("corp_idno"))){
						//找到該使用者 並且也是該公司成員
						
						//將該使用者設為管理者 authority=4
						sql = "SELECT * FROM user where " +
						"corp_id='"+request.getParameter("corp_idno") + "' and " +
						"login='"+user_id+"'";
						rs = DBStmt.executeQuery(sql);
						rs.last(); 
						rs.updateString("authority","4");//權限
						rs.updateRow();//更新user資料
						isUpCorpStartDate = true;
					}
				}else{
					//找不到該使用者 則新增使用者
					sql = "SELECT * FROM user where " +
						"corp_id='"+request.getParameter("corp_idno") + "' and " +
						"login='"+user_id+"'";
					rs = DBStmt.executeQuery(sql);
					cont = 0;
					while(rs.next()){
						cont++;
					}
					rs.beforeFirst();//回到第一筆資料
					if(cont != 0){
						// 已存在的 user
						rs.last(); 
						rs.updateString("authority","4");//權限
						rs.updateRow();//更新user資料
						isUpCorpStartDate = true;
					}else{
						// 還沒有此 user
						rs.moveToInsertRow(); 
						rs.updateString("login", user_id);//帳號
						rs.updateString("pswd", request.getParameter("user_pwd"));//密碼
						rs.updateString("corp_id", request.getParameter("corp_idno"));//公司統編
						rs.updateString("date_start", date);//公司統編
						rs.updateString("authority", "4");//權限
						rs.insertRow();//新增user資料
						isUpCorpStartDate = true;
					}
				}
				
				//更新 corp table 公司啟動時間
				if(isUpCorpStartDate == true){
					sql = "UPDATE corp SET start_date='" + date + "' WHERE corp_id='" + id + "'";
					CorpStmt.executeUpdate(sql);
					
					//洗掉停用記錄
					sql = "UPDATE corp SET stop_date='' WHERE corp_id='" + id + "'";
					CorpStmt.executeUpdate(sql);
					
					return true;
				}else{
					return false;
				}
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
		return false;
	}
	
	//關閉服務
	public void OffServices(String corp_id){
		if(corp_id != null){
			String id = corp_id;
			String date= 
				calendar.get(Calendar.YEAR) + "-"+(calendar.get(Calendar.MONTH)+1) + 
				"-" + calendar.get(Calendar.DAY_OF_MONTH);
			
			String sql = "UPDATE corp SET stop_date='" + date + "' WHERE corp_id='" + id + "'";
			
			try {
				CorpStmt.executeUpdate(sql);
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}
	
	//取得當前 ResultSet 下的管理者名單
	public HashMap getCorpManagers(ResultSet rs){
		HashMap map = new HashMap();
		String sql = "";
		
		ResultSet tmprs = rs;
		try {
			while(tmprs.next()){
				if(map.get(tmprs.getString("idno")) == null && !tmprs.getString("idno").equals("")){
					sql = "SELECT * FROM user where " +
							"corp_id="+tmprs.getString("idno")+" and "+"authority=4";
					ResultSet dbrs = DBStmt.executeQuery(sql);
					while(dbrs.next()){
						map.put(tmprs.getString("idno"), dbrs.getString("login"));
					}
				}
			}
			return map;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	//可搜尋方法 id=欄位 var=參數
	public ResultSet getCorpListRs(int start , int num , String id , String var){
		HashMap TableMap = (HashMap)DBMap.get("Corp");
		
		try {
			//System.out.println(sql);
			//SELECT * FROM corp where corp_id=0 Limit 0,10
			//SELECT * FROM corp where corp_id like '%B01%' Limit 0,100; 
			/*
			SELECT * FROM corp where corp_id like '%台灣%' Limit 0,5 
			UNION 
			SELECT * FROM corp where cname like '%台灣%' Limit 0,5 
			UNION  
			SELECT * FROM corp where addr like '%台灣%' Limit 0,5
			*/
	        String sql = "";
	        
	        //搜尋特定欄位 "corp_id","cname","ename","idno","abbr","email","start_date","stop_date
	        String s[] = {
	        		"corp_id","cname","ename","idno","abbr","email",
	        		"start_date","stop_date","tel","cont"};
	        for(int i = 0 ; i < s.length ; i++){
	        	sql += "SELECT * FROM corp where "+
        			s[i]+" like '%"+var+"%' Limit "+start+","+num+" UNION ";
	        }
	        /*
	        ResultSet rs2 = CorpStmt.executeQuery("EXPLAIN corp");
	        while(rs2.next()){
	        	sql += "SELECT * FROM corp where "+
	        		rs2.getString("Field")+" like '%"+var+"%' Limit "+start+","+num+" UNION ";
	        }
	        rs2.close();
	        */
			sql = sql.substring(0 , sql.length()-6);
			//System.out.println(sql);
			
			ResultSet rs = CorpStmt.executeQuery(sql);
			
			return rs;
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}
	
	//回傳公司列表方法
	public ResultSet getCorpListRs(int start , int num){
		HashMap TableMap = (HashMap)DBMap.get("Corp");
		
		try {
			String sql = "SELECT * FROM corp LIMIT "+start+","+num;
			
			ResultSet rs = CorpStmt.executeQuery(sql);
			
			return rs;
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}
	
	//取得欄位資料
	public ResultSetMetaData getCorpForme(ResultSet rs){
		try {
			return rs.getMetaData();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}
	
	public String[][] getCorpListArray(int start , int num){
		HashMap TableMap = (HashMap)DBMap.get("Corp");
		
		try {
			ResultSet rs = getCorpListRs(start , num);
			
			ArrayList data;
			ArrayList arr = new ArrayList();
			
			ResultSetMetaData rm = getCorpForme(rs);
			//rm.getColumnLabel(j) 可取得藍位名稱
			//rm.getColumnCount() 取得藍位總數
			
			
			//加入資料
			int n = start;
			while(rs.next()){
				data = new ArrayList();
				for(int i = 1 ; i < rm.getColumnCount()+1 ; i++){
					if(rs.getString(rm.getColumnLabel(i)) != null){
						//System.out.println(i+""+(rm.getColumnLabel(i)+" "+rs.getString(rm.getColumnLabel(i))));
						data.add(rs.getString(rm.getColumnLabel(i)));
					}
				}
				arr.add(data);
			}
			
			//
			String[][] re = new String[arr.size()][rm.getColumnCount()];
			for(int i = 0 ; i < arr.size() ; i++){
				for(int j = 0 ; j < ((ArrayList)arr.get(i)).size() ; j++){
					re[i][j] = (String)((ArrayList)arr.get(i)).get(j);
				}
			}
			return re;//
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}
}

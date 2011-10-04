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

public class TopUserModel {
	HttpServletRequest request = null;
	HttpServletResponse response = null;
	HttpServlet servlet = null;
	HttpSession session = null;
	
	HashMap CorpMap = null;
	Statement CorpStmt = null;
	
	HashMap DBMap = null;
	Statement DBStmt = null;
	
	Calendar calendar = Calendar.getInstance();
	
	Singleton singleton;
	public void set( HttpServlet servlet , HttpServletRequest request , HttpServletResponse response){
		this.servlet = servlet;
		this.request = request;
		this.response = response;
		session = request.getSession();
		
		singleton = Singleton.getInstance(servlet.getServletContext());
		DBMap = (HashMap) servlet.getServletContext().getAttribute("DBMap");
		DBStmt = singleton.getDBStatement((String)session.getAttribute("user_nbr"));
		CorpMap = singleton.DBMap;
		DBMap = singleton.DBMap;
		CorpStmt = singleton.getDB2Statement((String)session.getAttribute("user_nbr"));
	}
	
	
	
	public ArrayList getCorpBudget(String corp_id){
		ArrayList arr = new ArrayList();
		HashMap map;
		String sql = "SELECT * FROM budget where " + 
			"corp_id='"+session.getAttribute("user_corp_id")+"'";
		try {
			ResultSet rs = DBStmt.executeQuery(sql);
			while(rs.next()){
				map = new HashMap();
				map.put("budget_nbr", rs.getString("budget_nbr"));//預算編號
				map.put("user_nbr", rs.getString("user_nbr"));//決行者編號
				map.put("corp_id", rs.getString("corp_id"));//公司名
				map.put("budget_total", rs.getString("budget_total"));//總預算決行者自行維護
				map.put("budget_grant", rs.getString("budget_grant"));//已決行預算己報帳預算+己決行預算
				map.put("budget_booked", rs.getString("budget_booked"));//已報帳預算
				arr.add(map);
			}
			return arr;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	//修改 以決行預算
	public void editUserBudgetGrant(String budget_nbr , String user_nbr , int budget_grant){
		String sql = "SELECT * FROM budget where budget_nbr='"+budget_nbr+ "'";
		//System.out.println(sql);
		try {
			ResultSet rs = DBStmt.executeQuery(sql);
			rs.absolute(-1);
			rs.last();
			int n = Integer.parseInt(rs.getString("budget_grant")) + budget_grant;
			rs.updateString("budget_grant",Integer.toString(n));
			rs.updateRow();
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}
	
	//修改 總預算
	public void editUserBudget(String budget_nbr , String user_nbr , String budget_total){
		String sql = "SELECT * FROM budget where budget_nbr='"+budget_nbr+ "'";
		//System.out.println(sql);
		try {
			ResultSet rs = DBStmt.executeQuery(sql);
			rs.absolute(-1);
			rs.last();
			rs.updateString("budget_total",budget_total);
			rs.updateRow();
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}
	
	public ArrayList getUserBudget(String user_nbr){
		ArrayList arr = new ArrayList();
		HashMap map;
		String sql = "SELECT * FROM budget where " + 
			"user_nbr='"+user_nbr+ "' and corp_id='"+session.getAttribute("user_corp_id")+"'";
		try {
			ResultSet rs = DBStmt.executeQuery(sql);
			while(rs.next()){
				map = new HashMap();
				map.put("budget_nbr", rs.getString("budget_nbr"));//預算編號
				map.put("user_nbr", rs.getString("user_nbr"));//決行者編號
				map.put("corp_id", rs.getString("corp_id"));//公司名
				map.put("budget_total", rs.getString("budget_total"));//總預算決行者自行維護
				map.put("budget_grant", rs.getString("budget_grant"));//已決行預算己報帳預算+己決行預算
				map.put("budget_booked", rs.getString("budget_booked"));//已報帳預算
				arr.add(map);
			}
			return arr;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
}

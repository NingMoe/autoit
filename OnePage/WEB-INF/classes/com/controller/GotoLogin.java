package com.controller;

import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Calendar;
import java.util.HashMap;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import com.web.Singleton;

/**
 * Servlet implementation class GotoLogin
 */
public class GotoLogin extends HttpServlet {
	private static final long serialVersionUID = 1L;
    
	HttpSession session;
	Singleton singleton;

	String uid = null , pwd = null;
	String login = null , status = null;
	
	HashMap DB = null;
	Statement DBStmt = null;
	HashMap CorpMap = null;
	Statement CorpStmt = null;
	
	Calendar calendar = Calendar.getInstance();
    /**
     * @see HttpServlet#HttpServlet()
     */
    public GotoLogin() {
        super();
        // TODO Auto-generated constructor stub
    }
    
    //判斷是否已登入
    public void isLogin(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		session = request.getSession();
		singleton = Singleton.getInstance(this.getServletContext());
		//if(DB == null){
			///*
			//DB = (HashMap) this.getServletContext().getAttribute("DB");
			DBStmt = singleton.getDBStatement("hello");
			CorpMap = singleton.DBMap;
			CorpStmt = singleton.getDB2Statement("hello");
			//*/
			/*
			DB = (HashMap) this.getServletContext().getAttribute("DB");
			DBStmt = (Statement) DB.get("Statement");
			CorpMap = (HashMap) this.getServletContext().getAttribute("CorpDB");
			CorpStmt = (Statement) CorpMap.get("Statement");
			*/
		//}
		
		login = (String)session.getAttribute("Login");
		status = request.getParameter("status");
		
		if(login != null && !login.equals("")){
			//已登入
			response.sendRedirect("user.jsp");
		}else{
			//未登入
			if(status != null){
				if(status.equals("login")){//欲登入
					this.Login(request, response);
				}else if(status.equals("logout")){//欲登出
					this.Logout(request, response);
				}
			}else{
				response.sendRedirect("Login.jsp");
			}
		}
	}
    
    //登入方法
    public void Login(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
    	if(request.getParameter("userid") != null){
			uid = request.getParameter("userid");
		}
		if(request.getParameter("pwd") != null){
			pwd = request.getParameter("pwd");
		}
		HashMap sys = singleton.SysMap;
		if((uid != null && uid.equals(sys.get("ADMINname"))) && 
				(pwd != null && pwd.equals((String)sys.get("ADMINpwd")))){
			//登入
			session.setAttribute("user", uid);//加入 登入資訊
			response.sendRedirect("Admin");//送到 Admin 頁面
		}else{
			//其他身分
			String sql = "SELECT * FROM user where " +
			"login='" + uid + "' and " + "pswd='" + pwd + "'";
			
			try {
				ResultSet rs = DBStmt.executeQuery(sql);
				ResultSet corp_rs;
				int cont = 0;
				while(rs.next()){
					cont++;
					session.setAttribute("user_authority", rs.getString("authority"));//系統內身分 加入 登入資訊
					session.setAttribute("user_corp_id", rs.getString("corp_id"));//公司統編 加入 登入資訊
					session.setAttribute("user_cname", rs.getString("cname"));//中文名 加入 登入資訊
					session.setAttribute("user_nbr", rs.getString("user_nbr"));//User 系統編號
					sql = "SELECT * FROM corp where idno='"+rs.getString("corp_id")+"'";
					corp_rs = CorpStmt.executeQuery(sql);
					corp_rs.next();
					session.setAttribute("corp_cname", corp_rs.getString("cname"));//公司 中文名稱
					session.setAttribute("corp_ename", corp_rs.getString("ename"));//公司 英文名稱
					session.setAttribute("corp_abbr", corp_rs.getString("abbr"));//公司 縮寫
					
					rs.last();
					String date= 
						calendar.get(Calendar.YEAR) + "-"+(calendar.get(Calendar.MONTH)+1) + 
						"-" + calendar.get(Calendar.DAY_OF_MONTH);
					rs.updateString("date_ack", date);//登記登入時間
					rs.updateRow();//更新user資料
				}
				if(cont > 0){
					rs.beforeFirst();
					session.setAttribute("user", uid);//加入 登入資訊
					response.sendRedirect("User");//送到 User 頁面
					
					String user_nbr = (String)session.getAttribute("user_nbr");
					session.setAttribute("user_DBCon",singleton.getDBConnection(user_nbr));
					session.setAttribute("user_CDBCon",singleton.getDB2Connection(user_nbr));
					return;
				}
			} catch (SQLException e) {
				e.printStackTrace();
			}
			
			//找不到登入資訊就跳回Login.jsp
			response.sendRedirect("Login.jsp");
		}
    }
    
    public void Logout(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
    	singleton.releaseDBConnection((String)session.getAttribute("user_nbr"));
    	singleton.releaseDB2Connection((String)session.getAttribute("user_nbr"));
    	session.invalidate();//清除 session
    	response.sendRedirect("Login.jsp");
    }
    
	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		this.isLogin(request, response);
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		this.isLogin(request, response);
	}

}

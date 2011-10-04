package com.debug;

import java.io.IOException;
import java.io.PrintWriter;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.web.Singleton;

/**
 * Servlet implementation class HelloInit
 */
public class HelloDebug extends HttpServlet {
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public HelloDebug() {
        super();
        // TODO Auto-generated constructor stub
    }

    Singleton singleton;
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		
		request.setCharacterEncoding("UTF-8");
		response.setContentType("text/html;charset=UTF-8");
		PrintWriter out = response.getWriter();
		singleton = Singleton.getInstance(this.getServletContext());
		out.print("HelloDebug 系統參數確認畫面<br>");
		
		out.print("<h4>系統參數</h4>");
		
		
		HashMap SysMap = singleton.SysMap;
        String[] tmp = {
				"DB","DBport","DBname","DBpwd","DBuser","",
				"CorpDBname","",
				"ADMINname","ADMINpwd","",
				"SystemEmail","MailServer"
				};
        for(int i = 0 ; i < tmp.length ; i++){
        	if(!tmp[i].equals(""))
        		out.print(tmp[i]+" : "+SysMap.get(tmp[i])+"<br>");
        	else
        		out.print("<br>");
        }
        //
        out.print("<hr>");
        //
        HashMap DBMap = singleton.DBMap;
        
        HashMap DBSys = (HashMap) this.getServletContext().getAttribute("DB");
        Statement Stmt = singleton.getDBStatement("hello");
        //HashMap CorpDBSys = (HashMap) this.getServletContext().getAttribute("CorpDB");
        Statement CorpStmt = singleton.getDB2Statement("hello");
        
        if(Stmt != null){
        	out.print("<h4>DB 成功連線</h4>");
	        try{
	        	Statement stmt2 = singleton.getDBStatement("hello");
		        String sql;
				
				sql = "show table status";
				ResultSet rs = Stmt.executeQuery(sql);
				HashMap TableMap;
				while(rs.next()){
	                out.println("<h4>Table : "+rs.getString("Name")+"</h4>");
	                ResultSet rs2 = stmt2.executeQuery("EXPLAIN "+rs.getString("Name"));
	                TableMap = (HashMap)DBMap.get(rs.getString("Name"));
	                if(TableMap != null){
		                while(rs2.next()){
		                	if(TableMap.get(rs2.getString(1)) != null){
		                		out.print(
		                				"<B><font color=\"#227700\">V </font></B>"+
		                				rs2.getString("Field")+" "+
		                				TableMap.get(rs2.getString("Field"))+"<br>");
		                	}else{
		                		out.print(
		                				"<font color=\"#AA0000\"><B>X </B>"+
		                				rs2.getString("Field")+" "+
		                				TableMap.get(rs2.getString("Field"))+"</font><br>");
		                	}
		                }
	                }else{
	                	out.print(
                				"<font color=\"#AA0000\"><B>X</B> 同步 Table "+rs.getString("Name")+"失敗</font><br>");
	                }
	            }
				
				sql = "EXPLAIN corp";
				rs = CorpStmt.executeQuery(sql);
				TableMap = (HashMap)DBMap.get("Corp");
				out.println("<h4>Table : corp</h4>");
				while(rs.next()){
                	if(TableMap.get(rs.getString(1)) != null){
                		out.print(
                				"<B><font color=\"#227700\">V </font></B>"+
                				rs.getString("Field")+" "+
                				TableMap.get(rs.getString("Field"))+"<br>");
                	}else{
                		out.print(
                				"<font color=\"#AA0000\"><B>X </B>"+
                				rs.getString("Field")+" "+
                				TableMap.get(rs.getString("Field"))+"</font><br>");
                	}
                }
				
				out.print("<br>");
	        }catch(SQLException e){
				out.print("SQLException : " + e.getMessage());
			}
        }else{
        	out.print("DB 尚未連線<br>");
        	/*
        	if(DBSys.get("ClassNotFoundException") != null){
        		ClassNotFoundException e = 
        			(ClassNotFoundException) DBSys.get("ClassNotFoundException");
        		out.print("ClassNotFoundException : "+e.getMessage()+"<br>");
        	}
        	if(DBSys.get("SQLException") != null){
        		SQLException e = (SQLException) DBSys.get("SQLException");
        		out.println("SQLException : "+e.getMessage()+"<br>");
        	}
        	*/
        }
        //
        out.print("<hr>");
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
	}

}

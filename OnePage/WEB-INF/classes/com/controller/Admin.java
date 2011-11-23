package com.controller;

import java.io.IOException;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.HashMap;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import com.model.AdminModel;

/**
 * Servlet implementation class Admin
 */
public class Admin extends HttpServlet {
	private static final long serialVersionUID = 1L;
	HttpSession session;   
	AdminModel model = new AdminModel();
	//
	String page = "0";
	String status = "list"; // list 瀏覽 列表 , new 新增公司
	String select = ""; 
    /**
     * @see HttpServlet#HttpServlet()
     */
    public Admin() {
        super();
        
        // TODO Auto-generated constructor stub
    }
    
    protected void processRequest(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
    	request.setCharacterEncoding("UTF-8");
		response.setContentType("text/html;charset=UTF-8");
    	session = request.getSession();
    	
		if(session.getAttribute("user") == null){
			response.sendRedirect("GotoLogin");
		}else{
			if(request.getParameter("page") != null){
				if(Integer.parseInt(request.getParameter("page")) > -1){
					page = request.getParameter("page");
				}
			}else{
				page = "0";
			}
			if(request.getParameter("status") != null){
				status = request.getParameter("status");
			}else{
				status = "list";
			}
			if(request.getParameter("select") != null){
				select = request.getParameter("select");
			}else{
				select = "";
			}
			
			request.setAttribute("msg", session.getAttribute("user"));//登入成功
			model.set(this, request, response);
			
			if(status.equals("list")){//列出公司
				this.gotoList(request, response);
			}else if(status.equals("edit")){//新增公司
				this.gotoEditCorp(request, response);
			}else{
				this.gotoList(request, response);
			}
		}
    }
    
    public void gotoList(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
    	ResultSet rs = null;
    	ResultSetMetaData rm = null;
    	
    	//更新資料
    	if(request.getParameter("onservices") != null && request.getParameter("corp_id") != null){
    		if(request.getParameter("onservices").equals("on")){
    			if(!model.OnServices(request.getParameter("corp_id") , request.getParameter("user_id"))){
    				//啟用失敗
    				select = request.getParameter("user_id");
    				request.setAttribute("Error_msg", "啟用服務錯誤");
    			}
    		}else if(request.getParameter("onservices").equals("off")){
    			model.OffServices(request.getParameter("corp_id"));
    		}
    	}
    	
    	//取得列表
    	if(!select.equals("")){
    		//有搜尋
    		rs = model.getCorpListRs(Integer.parseInt(page), 100 , "corp_id" , select);
    	}else{
    		//沒有搜尋
    		rs = model.getCorpListRs(Integer.parseInt(page), 100);
    	}
    	if(rs != null){
    		rm = model.getCorpForme(rs);
    	}
    	
		request.setAttribute("data", rs);
		request.setAttribute("rm", rm);
		
		/*
		HashMap cManagers = model.getCorpManagers(rs);
		if(cManagers != null){
			request.setAttribute("CorpManagers", cManagers);
			try {
				//rs.absolute(int row)//到指定的資料
				rs.beforeFirst();//回到第一筆資料
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
		*/
		
		request.getRequestDispatcher("admin.jsp").forward(request, response);
    }
    
    public void gotoEditCorp(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
    	
		request.getRequestDispatcher("admin.jsp").forward(request, response);
    }
    
	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		processRequest(request, response);
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		processRequest(request, response);
	}

}

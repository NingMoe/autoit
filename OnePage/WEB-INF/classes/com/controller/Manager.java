package com.controller;

import java.io.IOException;
import java.sql.ResultSet;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import com.model.AdminModel;
import com.model.ManagerModel;

/**
 * Servlet implementation class Manager
 */
public class Manager extends HttpServlet {
	private static final long serialVersionUID = 1L;
	HttpSession session; 
	ManagerModel model = new ManagerModel();
	//
	String page = "0";
	String status = "list"; // list 瀏覽 列表 
	String select = ""; 
    /**
     * @see HttpServlet#HttpServlet()
     */
    public Manager() {
        super();
        // TODO Auto-generated constructor stub
    }
    protected void processRequest(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
    	request.setCharacterEncoding("UTF-8");
		response.setContentType("text/html;charset=UTF-8");
		session = request.getSession();
		ResultSet rs = null;
		
		if(session.getAttribute("user") == null){
			response.sendRedirect("GotoLogin");//尚未登入
		}else if(session.getAttribute("user_authority").equals("4")){
			
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
			
			if(status.equals("edit")){
				//System.out.println("edit");
				model.UpDataUser();
				status = "list";
			}else if(status.equals("newUser")){
				// login , pswd , authority 三欄位必填 其他選填
				if( request.getParameter("login")!=null && !request.getParameter("login").equals("") &&
					request.getParameter("pswd")!=null && !request.getParameter("pswd").equals("") &&
					request.getParameter("authority")!=null && !request.getParameter("authority").equals("")){
					
					model.newUser(request.getParameter("login"), request.getParameter("pswd"), 
							request.getParameter("authority"), 
							request.getParameter("cname"), 
							request.getParameter("ename"), 
							request.getParameter("phone"), 
							request.getParameter("role"), 
							request.getParameter("rank"), 
							request.getParameter("lastname"), 
							request.getParameter("firstname"),
							request.getParameter("id"));
				}
				status = "list";
			}
			if(status.equals("list")){
				rs = model.getUserListRs(Integer.parseInt(page), 100 , (String) session.getAttribute("user_corp_id"));
				request.setAttribute("data", rs);
				request.setAttribute("status", status);
			}
			request.getRequestDispatcher("manager.jsp").forward(request, response);
		}else{
			response.sendRedirect("User");//非管理者送回 User
		}
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

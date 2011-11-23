package com.controller;

import java.io.IOException;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import com.model.TopUserModel;
import com.model.UserModel;

/**
 * Servlet implementation class TopUser
 */
public class TopUser extends HttpServlet {
	private static final long serialVersionUID = 1L;
	HttpSession session;
	TopUserModel model = new TopUserModel(); 
	HttpServletRequest request;
	HttpServletResponse response;
	
    public TopUser() {
    	super();
	}
	
	protected void processRequest(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		response.setContentType("text/html;charset=UTF-8");
		session = request.getSession();
		model.set(this, request, response);
		this.request = request;
		this.response = response;
		
		if(session.getAttribute("user") == null){
			response.sendRedirect("GotoLogin");
		}else if(session.getAttribute("user_authority").equals("2")){
			String user_nbr = (String)session.getAttribute("user_nbr");
			String corp_id = (String)session.getAttribute("corp_id");
			
			if(request.getParameter("isStatus") != null){
				if(request.getParameter("isStatus").equals("edit")){
					String budget_nbr = request.getParameter("budget_nbr");
					String budget_total = request.getParameter("budget_total");
					//System.out.println(budget_nbr + " , " +user_nbr+" , "+ budget_total);
					model.editUserBudget(budget_nbr, user_nbr, budget_total);
				}
			}
			
			request.setAttribute("UserBudget", model.getUserBudget(user_nbr));
			request.getRequestDispatcher("topUser.jsp").forward(request, response);
		}else{
			response.sendRedirect("GotoLogin");
		}
	}
	
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		this.processRequest(request, response);
	}
	
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		this.processRequest(request, response);
	}
}

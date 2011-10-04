package com.controller;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import com.model.AdminModel;
import com.model.TopUserModel;
import com.model.UserModel;

/**
 * Servlet implementation class User
 */
public class User extends HttpServlet {
	private static final long serialVersionUID = 1L;
	HttpSession session;
	UserModel model = new UserModel(); 
	TopUserModel topmodel = new TopUserModel(); 
	HttpServletRequest request;
	HttpServletResponse response;
    /**
     * @see HttpServlet#HttpServlet()
     */
    public User() {
        super();
        // TODO Auto-generated constructor stub
    }
    
    protected void processRequest(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
    	request.setCharacterEncoding("UTF-8");
		response.setContentType("text/html;charset=UTF-8");
		session = request.getSession();
		model.set(this, request, response);
		topmodel.set(this, request, response);
		this.request = request;
		this.response = response;
		
		if(session.getAttribute("user") == null){
			response.sendRedirect("GotoLogin");
		}else{
			String isAction = request.getParameter("isAction");
			if(isAction != null){
				if(isAction.equals("trip")){
					doTrip();//trip 處理
				}else if(isAction.equals("approved")){
					doApproved();//批核/決行 處理
				}else if(isAction.equals("actual")){
					doActual();//報帳
				}
			}
			
			String isStatus = request.getParameter("isStatus");
			if(isStatus !=null && isStatus.equals("show")){
				//檢視  詳細資料
				request.setAttribute("show_trip_list", 
						model.getTripMap(request.getParameter("user_nbr"), request.getParameter("Trip")));
				
				request.setAttribute("show_tripRoute_list", model.getTripRouteArr(request.getParameter("Trip")));
				request.setAttribute("show_tripExpense_list", model.getTripEstim_ExpenseArr(request.getParameter("Trip")));
				
				ArrayList show_tripFlow_list = model.getTripFlow_Trip(request.getParameter("Trip"));
				request.setAttribute("show_tripFlow_list", show_tripFlow_list);
				request.setAttribute("show_tripFlowUser_list", model.getUserGroups(show_tripFlow_list , "aprover_nbr"));
			}else{
				if(request.getParameter("trop_id") != null){
					//編輯 詳細資料
					request.setAttribute("edit_trip_list", 
							model.getTripMap((String)session.getAttribute("user_nbr"), request.getParameter("trop_id")));
					
					ArrayList edit_tripFlow_list = model.getTripFlow_Trip(request.getParameter("trop_id"));
					request.setAttribute("edit_tripFlow_list", edit_tripFlow_list);
					request.setAttribute("edit_tripFlowUser_list", model.getUserGroups(edit_tripFlow_list , "aprover_nbr"));
					request.setAttribute("edit_tripRoute_list", model.getTripRouteArr(request.getParameter("trop_id")));
					request.setAttribute("edit_tripExpense_list", model.getTripEstim_ExpenseArr(request.getParameter("trop_id")));
				}
			}
			
			String isTrip = request.getParameter("isTrip");
			if(isTrip != null && isTrip.equals("actual")){
				//報帳資訊
				String trip_id = request.getParameter("trip_id");
				if(trip_id != null){
					request.setAttribute("actual_trip_list", model.getTripMap((String)session.getAttribute("user_nbr"),trip_id));
					request.setAttribute("actual_tripRoute_list", model.getTripRouteArr(trip_id));
					request.setAttribute("actual_tripExpense_list", model.getTripEstim_ExpenseArr(trip_id));
					ArrayList actual_tripFlow_list = model.getTripFlow_Trip(trip_id);
					request.setAttribute("actual_tripFlow_list", actual_tripFlow_list);
					request.setAttribute("actual_tripFlowUser_list", model.getUserGroups(actual_tripFlow_list , "aprover_nbr"));
					request.setAttribute("actual_TripActual_Expense_list", model.getTripActual_ExpenseArr(trip_id));
				}
			}
			
			if(session.getAttribute("user_authority").equals("2")){
				//決行者才有的資訊
				request.setAttribute("UserBudget", topmodel.getUserBudget((String)session.getAttribute("user_nbr")));
			}
			
			ArrayList trip_flow ;
			trip_flow = model.getTripFlow((String)session.getAttribute("user_nbr"));
			request.setAttribute("user_date", model.getUser((String)session.getAttribute("user_nbr")));
			request.setAttribute("trip_flow", trip_flow);
			request.setAttribute("trip_flow_tripGroups", model.getTripGroups(trip_flow));
			request.setAttribute("trip_flow_userGroups", model.getUserGroups(trip_flow));
			request.setAttribute("trip_list", model.getTrip((String)session.getAttribute("user_nbr")));
			request.setAttribute("msg", session.getAttribute("user"));
			request.getRequestDispatcher("user.jsp").forward(request, response);
		}
    }
    
    public void doActual(){
    	if(request.getParameter("isAction").equals("actual")){
    		if(request.getParameter("submit").equals("暫存")){
    			model.newTripActual_ExpenseArr(request.getParameter("trip_id"));
    		}else if(request.getParameter("submit").equals("送出")){
    			
    		}
    	}
    }
    
    public void doApproved(){
    	if(request.getParameter("isStatus") != null){
			if(request.getParameter("isStatus").equals("1")){
				//待批核
				String upStatus = "";
				if(request.getParameter("submit") != null){
					if(request.getParameter("submit").equals("批核") || 
							request.getParameter("submit").equals("決行")){
						
						if(request.getParameter("submit").equals("決行"))
							upStatus = "4";
						else if(request.getParameter("submit").equals("批核"))
							upStatus = "3";
						
						String unbr = model.getUserNBR(request.getParameter("nextUser"));
						//判斷 nextUserTop 欄位是否有資料
						if(request.getParameter("nextUserTop") != null && 
								!request.getParameter("nextUserTop").equals("")){
							HashMap map = model.getUser((String)session.getAttribute("user_nbr"));
							String s = model.getUserNBR((String)map.get("approver_mail"));
							String top = model.getUserNBR(request.getParameter("nextUserTop"));
							//System.out.println(request.getParameter("nextUserTop"));
							//System.out.println(s + "," + top);
							if(!s.equals(top)){
								unbr = top;
							}
						}
						
						//批核
						model.editTripFlow(
								request.getParameter("upTripFlow"), 
								request.getParameter("upTrip"), 
								upStatus);
						if(request.getParameter("submit").equals("批核")){
							//新增 轉發
							HashMap umap = model.getUser(unbr);
							//System.out.println(umap.get("login") + "," + umap.get("authority"));
							if(umap.get("authority").equals("2")){
								//轉發給 決行
								model.newTripFlow(
										request.getParameter("request_nbr"), 
										(String)umap.get("login"), 
										request.getParameter("upTrip"), 
										"2",
										false);
							}else{
								//轉發給 批核
								model.newTripFlow(
										request.getParameter("request_nbr"), 
										(String)umap.get("login"), 
										request.getParameter("upTrip"), 
										"1",
										false);
							}
						}
					}else if(request.getParameter("submit").equals("退回")){
						upStatus = "5";
						model.editTripFlow(
								request.getParameter("upTripFlow"), 
								request.getParameter("upTrip"), 
								upStatus);
					}
				}else{
					model.editTripFlow(
							request.getParameter("upTripFlow"), 
							request.getParameter("upTrip"), 
							request.getParameter("upStatus"));
				}
			}else if(request.getParameter("isStatus").equals("2")){
				//待決行
				model.editTripFlow(
						request.getParameter("upTripFlow"), 
						request.getParameter("upTrip"), 
						request.getParameter("upStatus"));
			}
		}
    }
    
    public void doTrip(){
    	if(request.getParameter("submit") != null && request.getParameter("submit").equals("暫存")){
			if(request.getParameter("status").equals("new")){
				//新增
				String trip_nbr = model.newTrip();
				model.newTripRoute(trip_nbr);
				model.newTripEstim_Expense(trip_nbr);
			}else{
				//更新
				String trip_nbr = model.editTrip(request.getParameter("status"));
				model.newTripRoute(trip_nbr);
				model.newTripEstim_Expense(trip_nbr);
			}
		}else if(request.getParameter("submit") != null && request.getParameter("submit").equals("申請")){
			String trip_nbr;
			if(request.getParameter("status").equals("new")){
				//尚無新增 trip
				trip_nbr = model.newTrip();
				model.newTripRoute(trip_nbr);
				model.newTripEstim_Expense(trip_nbr);
			}else{
				//已新增 trip
				trip_nbr = model.editTrip(request.getParameter("status"));
				model.newTripRoute(trip_nbr);
				model.newTripEstim_Expense(trip_nbr);
			}
			//System.out.println(trip_nbr);
			//提出核准
			if (trip_nbr != null && !trip_nbr.equals("")){
				String unbr = model.getUserNBR(request.getParameter("nextUser"));
				HashMap umap = model.getUser(unbr);
				if(umap.get("authority").equals("2")){
					//決行者
					model.newTripFlow(
							(String)session.getAttribute("user_nbr"), 
							request.getParameter("nextUser"), 
							trip_nbr,
							"2",
							true);
				}else{
					//非決行者
					model.newTripFlow(
							(String)session.getAttribute("user_nbr"), 
							request.getParameter("nextUser"), 
							trip_nbr, 
							"1",
							true);
				}
			}
		}else if(request.getParameter("submit") != null && request.getParameter("submit").equals("取消")){
			//取消trip
			model.stopTrip(request.getParameter("status"), "申請人取消");
		}else if(request.getParameter("submit") != null && request.getParameter("submit").equals("返回")){
			
		}
    }
    
	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		this.processRequest(request, response);
	}
	
	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		this.processRequest(request, response);
	}

}

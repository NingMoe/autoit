package com.web;
/*
 * 
 * 此 class 做廢
 * 
 * */
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.servlet.ServletContext;
import javax.servlet.ServletContextEvent;
import javax.servlet.ServletContextListener;

import com.controller.var;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;


public class Hello implements ServletContextListener {
	
	private HashMap SysMap;
	
	@Override
	public void contextDestroyed(ServletContextEvent sce) {
		// TODO Auto-generated method stub
		
	}
	
	@Override
	public void contextInitialized(ServletContextEvent sce) {
		// TODO Auto-generated method stub
		BufferedReader reader = null;
		try{
			ServletContext context = sce.getServletContext();
			String SysFile = context.getInitParameter("sys");//載入書籤檔案初始參數
			reader = new BufferedReader(
					new InputStreamReader(context.getResourceAsStream(SysFile),"UTF-8")
					);
			String input = null;
			
			SysMap = new HashMap();
			while((input = reader.readLine()) != null){//讀取每筆書籤紀錄
				String[] tokens = input.split(",");
				if(tokens != null){
					SysMap.put(tokens[0], tokens[1]);
				}
			}
			context.setAttribute("SysMap", SysMap);//設定為 Sys 屬性
			context.setAttribute("DB", getDBSys(context));//設定為 DB連線 屬性
			context.setAttribute("CorpDB", getCorpDBSys(context));//設定為 CropDB連線 屬性
			context.setAttribute("DBMap", getDBMap(context));//設定為 DB Table 屬性
			context.setAttribute("AirMap", getAirMap(context));//設定為 Air Map 機場列表  屬性
		}catch(IOException e){
			Logger.getLogger(Hello.class.getName()).log(Level.SEVERE,null,e);
		}finally{
			try{
				reader.close();
			}catch(IOException e){
				Logger.getLogger(Hello.class.getName()).log(Level.SEVERE,null,e);
			}
		}
	}
	
	public HashMap getCorpDBSys(ServletContext context){
		HashMap CorpDBSys = new HashMap();
		try{
			HashMap map = (HashMap) context.getAttribute("SysMap");
			Class.forName("com.mysql.jdbc.Driver");
			String Driver = 
				"jdbc:mysql://"+
				map.get("DB")+":"+map.get("DBport")+"/"+map.get("CorpDBname")+"?" +
				"user="+map.get("DBuser")+"&" +
				"password="+map.get("DBpwd")+"&" +
				"useUnicode=true&characterEncoding=utf-8" +
				"&autoReconnect=true&initialTimeout=0";
			CorpDBSys.put("Driver", Driver);
			
			Connection con = DriverManager.getConnection(Driver);
			CorpDBSys.put("Connection", con);
			
			Statement stmt = con.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE,ResultSet.CONCUR_UPDATABLE);
			CorpDBSys.put("Statement", stmt);
		}catch(ClassNotFoundException e){
			CorpDBSys.put("ClassNotFoundException", e);
		}catch(SQLException e){
			CorpDBSys.put("SQLException", e);
		}
		return CorpDBSys;
	}
	
	public HashMap getDBSys(ServletContext context){
		HashMap DBSys = new HashMap();
		try{
			HashMap map = (HashMap) context.getAttribute("SysMap");
			Class.forName("com.mysql.jdbc.Driver");
			String Driver = 
				"jdbc:mysql://"+
				map.get("DB")+":"+map.get("DBport")+"/"+map.get("DBname")+"?" +
				"user="+map.get("DBuser")+"&" +
				"password="+map.get("DBpwd")+"&" +
				"useUnicode=true&characterEncoding=utf-8" +
				"&autoReconnect=true&initialTimeout=0";
			DBSys.put("Driver", Driver);
			
			Connection con = DriverManager.getConnection(Driver);
			DBSys.put("Connection", con);
			
			Statement stmt = con.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE,ResultSet.CONCUR_UPDATABLE);
			DBSys.put("Statement", stmt);
		}catch(ClassNotFoundException e){
			DBSys.put("ClassNotFoundException", e);
		}catch(SQLException e){
			DBSys.put("SQLException", e);
		}
		return DBSys;
	}
	public HashMap getDBMap(ServletContext context){
		BufferedReader reader = null;
		HashMap DBMap = new HashMap();
		try{
			String SysFile = context.getInitParameter("DBtxt");//載入DB檔案初始參數
			reader = new BufferedReader(
					new InputStreamReader(context.getResourceAsStream(SysFile),"UTF-8")
					);
			String input = null;
			
			HashMap TableMap = new HashMap();
			while((input = reader.readLine()) != null){
				String[] tokens = input.split(",");
				if(tokens != null){
					//System.out.println(tokens[0]+" , "+tokens[1]);
					if(tokens[0].equals("table")){//新的table
						TableMap = new HashMap();
						DBMap.put(tokens[1],TableMap);//存入欄位
					}else{
						TableMap.put(tokens[0], tokens[1]);
					}
				}
			}
		}catch(IOException e){
			Logger.getLogger(Hello.class.getName()).log(Level.SEVERE,null,e);
		}finally{
			try{
				reader.close();
			}catch(IOException e){
				Logger.getLogger(Hello.class.getName()).log(Level.SEVERE,null,e);
			}
		}
		return DBMap;
	}
	public HashMap getAirMap(ServletContext context){
		BufferedReader reader = null;
		HashMap AirMap = new HashMap();
		try{
			String SysFile = context.getInitParameter("air");//載入air檔案初始參數
			
			reader = new BufferedReader(
					new InputStreamReader(context.getResourceAsStream(SysFile),"UTF-8")
					);
			String input = null;
			
			ArrayList<String> arr = new ArrayList<String>();
			//String[] s = new String[9];
			//int i = 0;
			while((input = reader.readLine()) != null){
				String[] tokens = input.split(",");
				if(tokens != null){
					if(tokens.length == 1){//
						arr = new ArrayList<String>();
						AirMap.put(tokens[0],arr);//存入列表
						//s[i++] = tokens[0];
						//System.out.println(tokens[0]);
					}else{
						arr.add(tokens[0]+","+tokens[1]+","+tokens[2]);
						//System.out.println(tokens[0]+" , "+tokens[1]+" , "+tokens[2]);
					}
				}
			}
			//for(int j = 0 ; j < s.length ; j++){
			//	ArrayList a = (ArrayList) AirMap.get(s[j]);
			//	System.out.println(s[j] + " , "+a.get(0));
			//}
		}catch(IOException e){
			Logger.getLogger(Hello.class.getName()).log(Level.SEVERE,null,e);
		}finally{
			try{
				reader.close();
			}catch(IOException e){
				Logger.getLogger(Hello.class.getName()).log(Level.SEVERE,null,e);
			}
		}
		return AirMap;
	}
}
package com.web;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.servlet.ServletContext;

public class Singleton {
	private static Singleton instance = null;
	
	private Connection DBCon;//數個?
	private Connection DB2Con;//數個?
	private HashMap DBConMap = new HashMap();
	private HashMap DB2ConMap = new HashMap();
	private ServletContext context;
	
	private HashMap DBTimeMap = new HashMap();
	
	public HashMap SysMap;
	public HashMap DBMap;
	public HashMap AirMap;
	public static Singleton getInstance(ServletContext context) {
		if (instance == null){
			synchronized(Singleton.class){
				if(instance == null) {
					instance = new Singleton(context);
				}
			}
		}
		return instance;
	}
	private Singleton(ServletContext context){
		//同步區
		this.context = context;
		
		SysMap = getSysMap(context);//設定為 DB Table 屬性
		AirMap = getAirMap(context);//設定為 Air Map 機場列表  屬性
		DBMap = getDBMap(context);
		
		context.setAttribute("SysMap", SysMap);//設定為 Sys 屬性
		context.setAttribute("DBMap", DBMap);//設定為 DB Table 屬性
		context.setAttribute("AirMap", AirMap);//設定為 Air Map 機場列表  屬性
	}
	
	public void isDate(String user){
		//判斷 使用者是否超過 20分鐘 1200秒 操作Connection
		Date date = null;
		date = (Date) DBTimeMap.get(user);
		if(date == null){
			date = Calendar.getInstance().getTime();
			//System.out.println(user+" 建立 time");
		}else{
			long d = Calendar.getInstance().getTime().getTime() - date.getTime();
			//System.out.println(user+" "+(d/1000));
			date = Calendar.getInstance().getTime();
			if(d/1000 > 1200-5){//超過20分鐘
				this.releaseDBConnection(user);
				this.releaseDB2Connection(user);
			}
		}
		DBTimeMap.put(user, date);
	}
	
	public Connection getDBConnection(String user){
		isDate(user);
		Connection con = (Connection) DBConMap.get(user);
		try {
			if(con == null){
				//建立DBCon
				HashMap map = SysMap;
				Class.forName("com.mysql.jdbc.Driver");
				String Driver = 
					"jdbc:mysql://"+
					map.get("DB")+":"+map.get("DBport")+"/"+map.get("DBname")+"?" +
					"user="+map.get("DBuser")+"&" +
					"password="+map.get("DBpwd")+"&" +
					"useUnicode=true&characterEncoding=utf-8" +
					"&autoReconnect=true&initialTimeout=0";
				con = DriverManager.getConnection(Driver);
				
				DBConMap.put(user, con);
			}
		} catch (ClassNotFoundException e) {
			e.printStackTrace();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return con;
	}
	public Statement getDBStatement(String user){
		Connection con = this.getDBConnection(user);
		Statement stmt = null;
		try {
			stmt = con.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE,ResultSet.CONCUR_UPDATABLE);
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return stmt;
	}
	public void releaseDBConnection(String user){
		DBConMap.remove(user);//釋放DBCon
	}
	public Connection getDB2Connection(String user){
		isDate(user);
		Connection con = (Connection) DB2ConMap.get(user);
		try{
			if(con == null){
				//建立DB2Con
				HashMap map = SysMap;
				Class.forName("com.mysql.jdbc.Driver");
				String Driver = 
					"jdbc:mysql://"+
					map.get("DB")+":"+map.get("DBport")+"/"+map.get("CorpDBname")+"?" +
					"user="+map.get("DBuser")+"&" +
					"password="+map.get("DBpwd")+"&" +
					"useUnicode=true&characterEncoding=utf-8" +
					"&autoReconnect=true&initialTimeout=0";
				con = DriverManager.getConnection(Driver);
				
				DB2ConMap.put(user, con);
			}
		}catch(ClassNotFoundException e){
			e.printStackTrace();
		}catch(SQLException e){
			e.printStackTrace();
		}
		return con;
	}
	public Statement getDB2Statement(String user){
		Connection con = this.getDB2Connection(user);
		Statement stmt = null;
		try {
			stmt = con.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE,ResultSet.CONCUR_UPDATABLE);
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return stmt;
	}
	public void releaseDB2Connection(String user){
		DB2ConMap.remove(user);//釋放DB2Con
	}
	
	//
	public HashMap getSysMap(ServletContext context){
		BufferedReader reader = null;
		HashMap SysMap = new HashMap();
		try{
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
		}catch(IOException e){
			Logger.getLogger(Hello.class.getName()).log(Level.SEVERE,null,e);
		}finally{
			try{
				reader.close();
			}catch(IOException e){
				Logger.getLogger(Hello.class.getName()).log(Level.SEVERE,null,e);
			}
		}
		return SysMap;
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

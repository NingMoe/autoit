package com.model;

import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;

public class MD5 {
	static MessageDigest algorithm;
	static byte[] defaultBytes;
	static byte messageDigest[];
	public static String getMD5(String str){
		defaultBytes = str.getBytes();
		try {
			if(algorithm == null)
				algorithm = (MessageDigest) MessageDigest.getInstance("MD5");
	    	algorithm.reset();
	    	algorithm.update(defaultBytes);
	    	messageDigest = algorithm.digest();
	 
	    	StringBuffer hexString = new StringBuffer();
	    	for (int i = 0; i < messageDigest.length; i++) {
	    	  hexString.append(Integer.toHexString(0xFF & messageDigest[i]));
	    	}
	    	String foo = messageDigest.toString();
	    	str = hexString + "";
	    } catch (NoSuchAlgorithmException nsae) {
	    	nsae.printStackTrace();
	    }
	    return str;
	} 
}

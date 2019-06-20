package com.test;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.test.excelUtils.ExcelSpreadsheetProcessor;
import com.test.form.UserInfoForm;

public class ExcelOperation {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		ExcelSpreadsheetProcessor exp = new ExcelSpreadsheetProcessor();
		
		Set<String> shippingList= new HashSet<String>();
		shippingList=exp.loadShippingList(UserInfoForm.class,"SHIPPINGLINE");
		System.out.println("shippingList"+shippingList);
		 for(String s:shippingList)
			{
	        	System.out.println("*"+s);
	        	List<UserInfoForm> resultSet =exp.ExcelDatatoClassObject(UserInfoForm.class, s);
	    		for(UserInfoForm userInfoForm: resultSet)
	    		{
	    			System.out.println(userInfoForm.getCONTAINER());
	    			System.out.println(userInfoForm.getSITE());
	    		}
			}
		
		
		
		//List<UserInfoForm> resultSet =exp.ExcelDatatoClassObject(UserInfoForm.class, "COSCO");
		//List<UserInfoForm> resultSet =exp.ExcelDatatoClassObject(UserInfoForm.class, "MSC");
		/*List<UserInfoForm> resultSet =exp.ExcelDatatoClassObject(UserInfoForm.class, "HYUNDAI");
		for(UserInfoForm userInfoForm: resultSet)
		{
			System.out.println(userInfoForm.getCONTAINER());
			System.out.println(userInfoForm.getSITE());
		}*/
		
	}

}

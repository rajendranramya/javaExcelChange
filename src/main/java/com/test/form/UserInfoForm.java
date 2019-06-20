package com.test.form;

import com.fasterxml.jackson.annotation.JsonGetter;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.test.excelUtils.DataMapping;

@DataMapping(ExcelFilePath="V:\\test\\InputExcel.xlsx", ExcelWorrkSheetName="UserInfo")
public class UserInfoForm {

	@JsonProperty	
	public String SHIPPINGLINE;
	private String CONTAINER;
	private String SITE;
	
	@JsonGetter("CONTAINER")
	public String getCONTAINER() {
		return CONTAINER;
	}
	public void setCONTAINER(String cONTAINER) {
		CONTAINER = cONTAINER;
	}
	@JsonGetter("SITE")
	public String getSITE() {
		return SITE;
	}
	public void setSITE(String sITE) {
		SITE = sITE;
	}
	
	
	
}

package com.test.excelUtils;

import java.lang.annotation.Annotation;

public class AnnotationHandle {

	public AnnotationHandle(){	
	}
	
	public <T> String ExcelWorrkSheetName(Class<T> objClass) {
		if(objClass.isAnnotationPresent(DataMapping.class)){
			Annotation ann = objClass.getAnnotation(DataMapping.class);
			DataMapping dataMap = (DataMapping) ann;
			return dataMap.ExcelWorrkSheetName();
		}		
		return null;		
	}
	
	public <T> String ExcelFilePath(Class<T> objClass) {
		if(objClass.isAnnotationPresent(DataMapping.class)){
			Annotation ann = objClass.getAnnotation(DataMapping.class);
			DataMapping dataMap = (DataMapping) ann;
			return dataMap.ExcelFilePath();
		}		
		return null;		
	}
}

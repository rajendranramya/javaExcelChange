package com.test.excelUtils;

import java.lang.annotation.*;

@Target({ElementType.METHOD, ElementType.FIELD, ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
public @interface DataMapping {
	String ExcelWorrkSheetName();
	String ExcelFilePath();
}

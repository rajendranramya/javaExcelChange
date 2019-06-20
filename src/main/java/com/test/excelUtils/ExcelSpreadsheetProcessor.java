package com.test.excelUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.json.simple.JSONObject;
import org.openqa.selenium.ElementNotVisibleException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * @author n0217408
 *
 */
public class ExcelSpreadsheetProcessor {
	private static final Logger LOG = LoggerFactory.getLogger(ExcelSpreadsheetProcessor.class);
    private  Workbook workbook = null;
    private  Sheet sheet = null;
    private  String returnJSON = null;
    private  List<String> resultSet = new ArrayList<String>();
  
    private  String TestScenarioHeader = "SHIPPINGLINE";
    
    public <T> List<T> ExcelDatatoClassObject(Class<T> objClass, String scenarioFilter)throws IOException, InvalidFormatException
    {
    	 List<T> resultSet1 = new ArrayList<T>();
    	 resultSet = new ArrayList<String>();
    	try{
	    	LOG.info("Loading test data from spreadsheet...");
	    	AnnotationHandle annotation = new AnnotationHandle();
	        FileInputStream excelInputStream = new FileInputStream(new File(annotation.ExcelFilePath(objClass)));
	        workbook = WorkbookFactory.create(excelInputStream);
	        //sheet = workbook.getSheetAt(0);
	        sheet = workbook.getSheet(annotation.ExcelWorrkSheetName(objClass));
	        loadFromSpreadsheet(scenarioFilter);
	      
	        for(String s:resultSet)
			{
	        	resultSet1.add((T) new ObjectMapper().readerFor(objClass).readValue(s));
			}
	        return resultSet1;
    	}
    	catch(Exception ex){
    		LOG.info("JSON created ..." + returnJSON);
    		throw new ElementNotVisibleException(ex.getMessage());
    		
    	}
    }
    
    private  void loadFromSpreadsheet(String sceanrioFilter) throws IOException, InvalidFormatException
	{
    	returnJSON = null;
		try{
			int numberOfColumns = countNonEmptyColumns(sheet);
			int testScenarioColumnCount = getColumnofScenarioFilter(numberOfColumns);
			if(testScenarioColumnCount < 0){
				LOG.info("Please verify the Excel Sheet does not contain the Column '" + TestScenarioHeader + "'" );
				return;
			}
			Row HeaderRow = null;
			for (Row row : sheet)
			{
			    if (isEmpty(row))
			    {
			        break;
			    }
			    else
			    {
			        // Row 0 will be Header Row
			        if (row.getRowNum() != 0)
			        {
			        	if(row.getCell(testScenarioColumnCount).getRichStringCellValue().toString().equals(sceanrioFilter) == false){
			        		continue;
			        	}
			        		
			        	JSONObject jObject = new JSONObject();
			            
			            for (int column = 0; column < numberOfColumns; column++)
			            {
			            	
			            	String strCellValue = null;
			            	String Headercell = HeaderRow.getCell(column).getRichStringCellValue().toString();
			                Cell cell = row.getCell(column);
			                if(cell == null)
			                	strCellValue = "";
			                else{
			                	Object cellObject = objectFrom(workbook, cell);
			                	if(cellObject == null)
			                		strCellValue = "";
			                	else{
			                		strCellValue = cellObject.toString();
			                	}
			                }
			                jObject.put(Headercell, strCellValue);
			                
			            }
			            returnJSON=jObject.toJSONString();
			            resultSet.add(returnJSON);
			        }
			        else {
			        	HeaderRow= row;
			        }
			        
			    }
			}
		}
		catch(Exception ex){
			LOG.info(ex.getMessage());
			throw new ElementNotVisibleException(ex.getMessage());
		}
		
	}
    
    /*public <T> T JSONtoClassMapping(Class<T> objClass)
    {
    	try
    	{
    		return new ObjectMapper().readerFor(objClass).readValue(returnJSON);
    	}
    	catch(Exception ex)
    	{
    		return null;
    	}
    }*/
    
    private  int getColumnofScenarioFilter(int noofColumns)
    {
    	int columnCntr = -1;
    	Row headerRow = sheet.getRow(0);
    	for (int column = 0; column < noofColumns; column++)
        {
    		Cell cell = headerRow.getCell(column);
            Object cellObject = objectFrom(workbook, cell);
            if(cellObject != null){
            	if(cellObject.toString().equals(TestScenarioHeader))
            		return column;
            }
        }
    	return columnCntr;
    }

    /**
     * 
     * @param row
     * @return
     */
    private  boolean isEmpty(final Row row)
    {
        Cell firstCell = row.getCell(0);
        return (firstCell == null) || (firstCell.getCellType() == Cell.CELL_TYPE_BLANK);
    }

    /**
     * Count the number of columns, using the number of non-empty cells in the
     * first row.
     */
    private  int countNonEmptyColumns(final Sheet sheet)
    {
        Row firstRow = sheet.getRow(0);
        return firstEmptyCellPosition(firstRow);
    }

    /**
     * 
     * @param cells
     * @return
     */
    private  int firstEmptyCellPosition(final Row cells)
    {
        int columnCount = 0;
        for (Cell cell : cells)
        {
            if (cell.getCellType() == Cell.CELL_TYPE_BLANK)
            {
                break;
            }
            columnCount++;
        }
        return columnCount;
    }

    /**
     * 
     * @param workbook
     * @param cell
     * @return
     */
    private  Object objectFrom(final Workbook workbook, final Cell cell)
    {
        Object cellValue = null;

        if (cell.getCellType() == Cell.CELL_TYPE_STRING)
        {
            cellValue = cell.getRichStringCellValue().getString();
        }
        else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
        {
            cellValue = getNumericCellValue(cell);
        }
        else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN)
        {
            cellValue = cell.getBooleanCellValue();
        }
        else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA)
        {
            cellValue = evaluateCellFormula(workbook, cell);
        }

        return cellValue;

    }

    /**
     * 
     * @param cell
     * @return
     */
    private  Object getNumericCellValue(final Cell cell)
    {
        Object cellValue;
        if (DateUtil.isCellDateFormatted(cell))
        {
            cellValue = new Date(cell.getDateCellValue().getTime());
        }
        else
        {
            cellValue = cell.getNumericCellValue();
        }
        return cellValue;
    }

    /**
     * 
     * @param workbook
     * @param cell
     * @return
     */
    private   Object evaluateCellFormula(final Workbook workbook, final Cell cell)
    {
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        CellValue cellValue = evaluator.evaluate(cell);
        Object result = null;

        if (cellValue.getCellType() == Cell.CELL_TYPE_BOOLEAN)
        {
            result = cellValue.getBooleanValue();
        }
        else if (cellValue.getCellType() == Cell.CELL_TYPE_NUMERIC)
        {
            result = cellValue.getNumberValue();
        }
        else if (cellValue.getCellType() == Cell.CELL_TYPE_STRING)
        {
            result = cellValue.getStringValue();
        }

        return result;
    }
    public  <T> Set<String> loadShippingList(Class<T> objClass,String sceanrioFilter) throws IOException, InvalidFormatException
	{
    	AnnotationHandle annotation = new AnnotationHandle();
        FileInputStream excelInputStream = new FileInputStream(new File(annotation.ExcelFilePath(objClass)));
        workbook = WorkbookFactory.create(excelInputStream);
        //sheet = workbook.getSheetAt(0);
        sheet = workbook.getSheet(annotation.ExcelWorrkSheetName(objClass));
    	Set<String> shippingList= new HashSet<String>();
    			try{
			int numberOfColumns = countNonEmptyColumns(sheet);
			int testScenarioColumnCount =  getColumnofScenarioFilter(numberOfColumns);
			if(testScenarioColumnCount < 0){
				LOG.info("Please verify the Excel Sheet does not contain the Column '" + TestScenarioHeader + "'" );
				return shippingList; 
			}
			Row HeaderRow = null;
			for (Row row : sheet)
			{
			    if (isEmpty(row))
			    {
			        break;
			    }
			    else
			    {
			        // Row 0 will be Header Row
			        if (row.getRowNum() != 0)
			        {
			        	/*if(row.getCell(testScenarioColumnCount).getRichStringCellValue().toString().equals(sceanrioFilter) == false){
			        		continue;
			        	}*/
			        	
			            for (int column = 0; column < 1; column++)
			            {
			            	
			            	String strCellValue = null;
			            	Cell cell = row.getCell(column);
			                if(cell == null)
			                	strCellValue = "";
			                else{
			                	Object cellObject = objectFrom(workbook, cell);
			                	if(cellObject == null)
			                		strCellValue = "";
			                	else{
			                		strCellValue = cellObject.toString();
			                	}
			                }
			                shippingList.add(strCellValue);
			                
			            }
			            
			        }
			        else {
			        	HeaderRow= row;
			        }
			        
			    }
			}
		}
		catch(Exception ex){
			LOG.info(ex.getMessage());
			throw new ElementNotVisibleException(ex.getMessage());
		}
    			return shippingList;
		
	}
    
       
 

}

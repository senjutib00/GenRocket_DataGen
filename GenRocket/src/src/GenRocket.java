package src;

import java.io.*;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import org.openqa.selenium.Platform;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.ElementNotSelectableException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;


import com.genRocket.GenRocketException;
import com.genRocket.engine.EngineAPI;
import com.genRocket.engine.EngineManual;
import com.microsoft.schemas.office.visio.x2012.main.CellType;

/**
 * Class to manage the batch execution of test scripts within the framework
 * 
 * @author Cognizant
 */


	/**
	 * The entry point of the test batch execution <br>
	 * Exits with a value of 0 if the test passes and 1 if the test fails
	 * 
	 * @param args
	 *            Command line arguments to the Allocator (Not applicable)
	 */
	
	public class GenRocket {
		
		public static void main(String[] args) throws IOException {
			
			GenRocket.gen();	
			System.out.println("Hello");
		}
			
		
		public static void gen() throws IOException {
	        String scenario = "C:\\Users\\sbhadra\\Eclipse_Workspace\\GenRocket\\Datatables\\HCVImportFileScenario.grs";
	        String domainName = "";
	        EngineAPI api = new EngineManual();
	        
	        try {
	              // Initialize Scenario
	        	
	              api.scenarioLoad(scenario);                      
	             // api.initialize(scenario);
	              List<String> domainsName = api.domains();
	                                     if(domainsName.size()!=0) {
	                                                   domainName = domainsName.get(0);
	                                                   System.out.println("scenario fetched : " +domainName);    
	                                     }
	                                     else if(domainsName.size()==0) {
	                                 }
	              api.receiverParameterSet(domainName, "ExcelFileReceiver", "path", "C:\\Users\\sbhadra\\Eclipse_Workspace\\GenRocket\\Datatables");                                   
	              api.receiverParameterSet(domainName, "ExcelFileReceiver", "fileName", "Backup");       
	              api.receivers(domainName);
	              // Run Scenario
	              System.out.println("scenario started running");  
	             
	              api.scenarioRun();   
	              System.out.println("scenario  completed");  
	              
	             // CopyingExcelData();
	              copyDataInExcel();
	            }
	                      catch (GenRocketException e) {
	                    	  System.out.println(e.getMessage());
	         }
	}

			public static void copyDataInExcel(){

		        String strSourceWb = "C:\\Users\\sbhadra\\Eclipse_Workspace\\GenRocket\\Datatables\\Backup.xlsx";
		        String strSourceSheet = "GenRocket";
		        String strDesWb = "C:\\Users\\sbhadra\\Eclipse_Workspace\\GSNAP_Test_Run - Copy\\Datatables\\Automated Scripts.xls";
		        String strDesSheet = "Automated Test Data";

		        try {
					copyExcelSheetDataFromOnetoAnotherWorkBook(strSourceWb, strSourceSheet, strDesWb, strDesSheet);
				} catch (IOException e) {
					e.printStackTrace();
				}

		    }
			public static void copyExcelSheetDataFromOnetoAnotherWorkBook(String strSourceWb, String strSourceSheet,
		            String strDesWb, String strDesSheet) throws IOException {
		        File srcExcel = new File(strSourceWb);
		        FileInputStream srcFis = new FileInputStream(srcExcel);
		        XSSFWorkbook srcWb = new XSSFWorkbook(srcFis);
		        Sheet srcSheet = srcWb.getSheet(strSourceSheet);
		        int intSrcRowLength = srcSheet.getLastRowNum() + 1;
		        int intSrcColumnLength = srcSheet.getRow(0).getPhysicalNumberOfCells();

		        File desExcel = new File(strDesWb);
		        FileInputStream desFis = new FileInputStream(desExcel);
		        HSSFWorkbook desWb = new HSSFWorkbook(desFis);
		        Sheet desSheet = null;
		        if ((desWb.getNumberOfSheets() != 0)) {
		            for (int intNoofSheet = 0; intNoofSheet < desWb.getNumberOfSheets(); intNoofSheet++) {
		                if (desWb.getSheetName(intNoofSheet).equalsIgnoreCase(strDesSheet)) {
		                    desSheet = desWb.getSheet(strDesSheet);
		                    break;
		                } 
		            }
		        } else {
		            desSheet = desWb.createSheet(strDesSheet);
		            desSheet = desWb.getSheet(strDesSheet);
		        }

		        FileOutputStream outFile = null;
		        for (int intRow = 0; intRow < intSrcRowLength; intRow++) {
		            Row srcRow = srcSheet.getRow(intRow);
		            Row desRow = desSheet.createRow(intRow);
		            for (int intColumn = 0; intColumn < intSrcColumnLength; intColumn++) {
		                srcRow.getCell(intColumn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
		                switch (srcRow.getCell(intColumn).getCellType()) {
		                case Cell.CELL_TYPE_STRING:
		                    desRow.createCell(intColumn).setCellValue(srcRow.getCell(intColumn).getStringCellValue());
		                    break;
		                case Cell.CELL_TYPE_BOOLEAN:
		                    desRow.createCell(intColumn).setCellValue(srcRow.getCell(intColumn).getBooleanCellValue());
		                    break;
		                case Cell.CELL_TYPE_NUMERIC:
		                    desRow.createCell(intColumn).setCellValue(srcRow.getCell(intColumn).getNumericCellValue());
		                    break;
		                case Cell.CELL_TYPE_FORMULA:
		                    switch (srcRow.getCell(intColumn).getCachedFormulaResultType()) {
		                    case Cell.CELL_TYPE_STRING:
		                        desRow.createCell(intColumn).setCellValue(srcRow.getCell(intColumn).getStringCellValue());
		                        break;
		                    case Cell.CELL_TYPE_NUMERIC:
		                        desRow.createCell(intColumn).setCellValue(srcRow.getCell(intColumn).getNumericCellValue());
		                        break;
		                    default:
		                        System.out.println("Undefined formula field value");
		                        break;
		                    }
		                default:
		                    System.out.println("Undefined field");
		                    break;
		                }

		            }
		        }

		        outFile = new FileOutputStream(desExcel);
		        desWb.write(outFile);
		        desWb.close();
		        srcWb.close();
		        outFile.close();

		    }

			
	}





		
	
	

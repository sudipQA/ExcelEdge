package excelpoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
/**
 * A dirty simple program that reads an Excel file.
 * @author www.codejava.net
 *
 */
public class SimpleExcelReaderExample{
	static String inputauto;
	static String inputcontrol;
    public static void main(String[] args) throws IOException   {
    	
    	SimpleExcelReaderExample excelreader = new SimpleExcelReaderExample();
    	    	//Demo exceldemo = new Demo();
    	//exceldemo.autodataSh
        //automation data sheet
    //	System.out.println(inputk);
System.out.println(inputauto);
System.out.println(inputcontrol);
        String excelFilePathauto = "C:\\Users\\sudip dutta\\Desktop\\tool\\datasheetauto.xlsx";
    	// String excelFilePathauto = inputauto;
        FileInputStream inputStreamauto = new FileInputStream(new File(excelFilePathauto));
        Workbook workbookauto = new XSSFWorkbook(inputStreamauto);
        int numberOfSheetsauto = workbookauto.getNumberOfSheets();
       //Control datasheet
        
       String excelFilePath = "C:\\Users\\sudip dutta\\Desktop\\tool\\control.xlsx";
     //  String excelFilePath = inputcontrol;
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        int numberOfSheets = workbook.getNumberOfSheets();
 
        
       for(int i = 0; i < numberOfSheets; i++) { 
    	   Sheet aSheet = workbook.getSheetAt(i);
    	     	   
	        for (int t = 0; t < numberOfSheetsauto; t++) {
	        	Sheet aSheetauto = workbookauto.getSheetAt(t);
	        	
	        	
	        	//check both sheet name equal
	        	if(aSheet.getSheetName().equals(aSheetauto.getSheetName())) {
	        		int columnIndex = 0;
	        		//System.out.println("&&&&&&&&&&&&&&&&&&&&&&&");
	        		//System.out.println(aSheetauto.getSheetName());
	        		//System.out.println(aSheetauto.getLastRowNum());
	        		//check the test case name
	        		for(int rowIndex =0;rowIndex<=aSheetauto.getLastRowNum();rowIndex++) {
	        			String testNm = aSheetauto.getRow(rowIndex).getCell(columnIndex).getStringCellValue();
	        			
	        			
	        			//System.out.println(testNm + " " + rowIndex + " " + columnIndex);
	        			
	        			if(testNm.equals("sudip1")) {
	        				
	        				switch(aSheet.getSheetName()) {
	        				case "owner":
	        					String value = aSheetauto.getRow(rowIndex).getCell(1).getStringCellValue();
	        				//System.out.println(value);
	        				int actresult = excelreader.owner(value);
	        				if(actresult==2){
	        					System.out.println("pass1");
	        					//aSheetauto.getRow(rowIndex).getCell(1).setCellValue("Javatpoint");
	        					
	        				}
	        				
	        				break;
	        				 case "beneficiary":
	        				break;
	        				default:
	        					System.out.println("no data");
	        					
	        				}
	        				 break; 
	        				
	        			}
	        			
	        		}
	        		
	        	}
	        	
	        }
	        

       }
       
       workbook.close();
       inputStream.close();
       workbookauto.close();
       inputStreamauto.close();
    	
    	
    }
    
    //Owner
    public static int owner(String fieldNm) {
    	
    	if(fieldNm =="sukanya1") {
    		return 1;
    	}
    	else{
    	    return 2;
    	  }
		
   }  
    

}  

    
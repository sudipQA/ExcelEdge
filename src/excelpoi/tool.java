package excelpoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class tool {

	public static void main(String[] args) throws IOException {
		String excelFilePathauto = "C:\\Users\\sudip dutta\\Desktop\\tool\\datasheetauto.xlsx";
        FileInputStream inputStreamauto = new FileInputStream(new File(excelFilePathauto));
        Workbook workbookauto = new XSSFWorkbook(inputStreamauto);
        int numberOfSheetsauto = workbookauto.getNumberOfSheets();
        for(int t = 0; t < numberOfSheetsauto; t++) {
        Sheet aSheetauto = workbookauto.getSheetAt(t);
        int columnIndex = 0;
        
        for(int rowIndex =0;rowIndex<aSheetauto.getLastRowNum();rowIndex++) {
        	
        	String data0= aSheetauto.getRow(rowIndex).getCell(columnIndex).getStringCellValue();
            System.out.println(data0);
        }
        
       
        

	}

	}
}

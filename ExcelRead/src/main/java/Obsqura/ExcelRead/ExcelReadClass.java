package Obsqura.ExcelRead;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*dependencies required for excel operations
 * 1.poi
 * 2.poi-ooxml
 * */

public class ExcelReadClass {

	public static void main(String[] args)throws IOException {
	
        //Reading file using fileinput stream
        FileInputStream fileIn = new FileInputStream("Student.xlsx");
        XSSFWorkbook wbRead = new XSSFWorkbook(fileIn);//mapping the file to workbook

        XSSFSheet readSheet= wbRead.getSheet("Sheet1");//mapping the first sheet 
        
        //Reading content of a sheet
        for (Row row:readSheet) {//iterating through  each row
        	
        	for(Cell cell:row) {//iterating through  each cell
        		
               if(cell.getCellType()==CellType.STRING)
                        System.out.print(cell.getStringCellValue());
               else if(cell.getCellType()==CellType.NUMERIC)
                        System.out.print((int)cell.getNumericCellValue()); 
                        
                System.out.print (" ");
	
        	}
        	
        	System.out.println();
        }
    	wbRead.close();


	}

}

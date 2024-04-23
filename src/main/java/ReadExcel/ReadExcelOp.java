package ReadExcel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelOp {
	
	 public static void main(String []args) {
		 
		 ReadExcelOp read = new ReadExcelOp();
		 try {
			read.ReadExcelOp();
		} catch (IOException e) {
			
			e.printStackTrace();
		}
		  
		 
	 }

	 public void ReadExcelOp() throws IOException  {
		 
		 //setting path to open the stream
		 FileInputStream fis = new FileInputStream("C:\\Users\\Admin\\Desktop\\Learning\\Task11.xlsx");
			
		 // open the work book
		 XSSFWorkbook xlworkbook = new XSSFWorkbook(fis);
			
		 // open the sheet
		 XSSFSheet xlsheet = xlworkbook.getSheetAt(0);
			
		 
		 //Get print
		 int lastrow = xlsheet.getLastRowNum();
		 //iterate through the rows
		 for(int i = 0; i<=lastrow; i++) {
			 XSSFRow xlrow = xlsheet.getRow(i);
			
			 
			 int lastColumn = xlrow.getLastCellNum();
			 for(int k=0; k<lastColumn; k++) {
				 XSSFCell xlcell = xlrow.getCell(k);
				 String cellvalue = xlcell.getStringCellValue(); 
				 System.out.print(cellvalue+ " ");
			 }
			 System.out.println(" ");
		 
		 }
		 	 
		  fis.close();
		  xlworkbook.close();
	 }
	 
}

package ReadExcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelop {

	public static void main(String[] args) {
		
		WriteExcelop excel = new WriteExcelop();
		excel.writingToExcelop();

	}
		public void writingToExcelop() {
			//set the stream to connect excel file
			try {
				FileInputStream fis = new FileInputStream("C:\\Users\\Admin\\Desktop\\Learning\\Task11.xlsx");
				
				//open the workbook
				XSSFWorkbook xlworkbook = new XSSFWorkbook(fis);
				
				//open the sheet
				XSSFSheet xlsheet = xlworkbook.createSheet("finalsheet");
				
				//get hold of the rows in the particular
				XSSFRow xlrow = xlsheet.createRow(0);
				
				//Now use cells and write your date in to the cell
				XSSFCell xlcell = xlrow.createCell(0);
				
				xlcell.setCellValue("Name");
				
				xlcell = xlrow.createCell(1);
				xlcell.setCellValue("Age");
				
				xlcell = xlrow.createCell(2);
				xlcell.setCellValue("Email");
				
				
				//Row number2
				xlrow = xlsheet.createRow(1);
				
				xlcell = xlrow.createCell(0);
				xlcell.setCellValue("John Doi");
				
				xlcell = xlrow.createCell(1);
				xlcell.setCellValue("30");
				
				xlcell = xlrow.createCell(2);
				xlcell.setCellValue("john@test.com");
				
				//Row numeber3
				xlrow = xlsheet.createRow(2);
				
				xlcell = xlrow.createCell(0);
				xlcell.setCellValue("Jane Doi");
				
				xlcell = xlrow.createCell(1);
				xlcell.setCellValue("28");
				
				xlcell = xlrow.createCell(2);
				xlcell.setCellValue("john@test.com");
				
				//Row number4
				xlrow = xlsheet.createRow(3);
				
				xlcell = xlrow.createCell(0);
				xlcell.setCellValue("Bob Smith");
				
				xlcell = xlrow.createCell(1);
				xlcell.setCellValue("35");
				
				xlcell = xlrow.createCell(2);
				xlcell.setCellValue("jacky@example.com");
				
				//Row number5
				xlrow = xlsheet.createRow(4);
				
				xlcell = xlrow.createCell(0);
				xlcell.setCellValue("Swapnil");
				
				xlcell = xlrow.createCell(1);
				xlcell.setCellValue("37");
				
				xlcell = xlrow.createCell(2);
				xlcell.setCellValue("swapnil@example.com");
				
				//close.
				FileOutputStream fos = new FileOutputStream("C:\\Users\\Admin\\Desktop\\Learning\\Task11.xlsx");
				xlworkbook.write(fos);
				fis.close();
				fos.close();
				xlworkbook.close();
				
			} catch (FileNotFoundException e) {
		
				e.printStackTrace();
			} catch (IOException e) {
				
				e.printStackTrace();
			}
			
		}
}
 
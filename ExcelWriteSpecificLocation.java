package sampleread;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelWriteSpecificLocation {
	
	private static final String FILE_NAME = "D:\\Java\\workspace\\ExcelReadExample\\src\\main\\resources\\TestFile3.xlsx";
	

public static void ExcelWritespecific(String s,int r,int c) {	
	System.out.println("Creating excel");
	  int rowNum = r-1;
	   int colNum = c-1;
	    XSSFWorkbook workbook = new XSSFWorkbook();// work book creation
	       XSSFSheet sheet = workbook.createSheet("Datatypes in Java"); //  sheet creation
	         Row row = sheet.createRow(rowNum); //   row creation
	          Cell cell = row.createCell(colNum);
	              cell.setCellValue((String) s);// string data is written
	                           

	       try {
	           FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
	           workbook.write(outputStream);
	           workbook.close();
	       } catch (FileNotFoundException e) {
	           e.printStackTrace();
	       } catch (IOException e) {
	           e.printStackTrace();
	       }

	       System.out.println("Done");
	   }

public static void main(String[] args) {
		ExcelWritespecific("sachin",1,7);

}
}
package sampleread;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelWrite {
	private static final String FILE_NAME = "D:\\Java\\workspace\\ExcelReadExample\\src\\main\\resources\\TestFile2.xlsx";
	public static void ExcelWrite() {
	
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Details of registered users");

		Object[][] datatypes = { { "Name", "Age", "Place" }, { "Thomas", 22, "Newyork" }, { "Sachin", 28, "Bangalore" },
				{ "Elias", 32, "Kochi" }, { "Jackson", 45, "Trivandrum" }, { "Mic", 52, "Kozhikode" } };

		int rowNum = 0;
		System.out.println("Creating excel");

			for (Object[] datatype : datatypes) {
				Row row = sheet.createRow(rowNum++);
				int colNum = 0;
				for (Object field : datatype) {
					Cell cell = row.createCell(colNum++);
					if (field instanceof String) {
						cell.setCellValue((String) field);
					} else if (field instanceof Integer) {
						cell.setCellValue((Integer) field);
					}
				}
			}

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
		//ExcelRead();
		ExcelWrite();
		//	ExcelReadWriteMain.ExcelWritespecific("sachin",0,1);

	}
	}
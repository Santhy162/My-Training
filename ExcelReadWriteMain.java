package sampleread;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadWriteMain {
	
	private static final String FILE_NAME = "D:\\Java\\workspace\\My_Maven\\src\\main\\resources\\TestFile1.xlsx";
	// public static ArrayList<String> arrl = new ArrayList<String>();
	public static void main(String args[]) {
		ExcelRead();
	}

	public static void ExcelRead() {

	// String[][] xData = new String[3][3];
	// String a = null;

	try {

	FileInputStream excelFile = new FileInputStream(FILE_NAME);
	XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
	XSSFSheet datatypeSheet = workbook.getSheetAt(0);
	//Sheet datatypeSheet = (Sheet) workbook.getSheetAt(0);
	Iterator<Row> iterator = datatypeSheet.iterator();

	while (iterator.hasNext()) {
	Row currentRow = iterator.next();
	Iterator<Cell> cellIterator = currentRow.iterator();

	while (cellIterator.hasNext()) {

	Cell currentCell = cellIterator.next();
	// getCellTypeEnum shown as deprecated for version 3.15
	// getCellTypeEnum ill be renamed to getCellType starting from version 4.0
	if (currentCell.getCellType() == CellType.STRING) {
	// a=currentCell.getStringCellValue();
	// System.out.print(a + "--");
	System.out.print(currentCell.getStringCellValue() + "======");
	// arrl.add( currentCell.getStringCellValue());
	}

	//else if (DateUtil.isCellDateFormatted(currentCell)) {
	// SimpleDateFormat datevar = new SimpleDateFormat("dd-MM-yyyy");
	// System.out.println("The cell contains a date value: ");
	// System.out.print(datevar.format(currentCell.getDateCellValue()) + "--");

	//} //commented for obsqurea class only
	else if (currentCell.getCellType() == CellType.NUMERIC) {
	System.out.print(currentCell.getNumericCellValue() + "--");
	// String aa= currentCell.getNumericCellValue().toString();
	// System.out.print(aa + "--");
	// arrl.add( currentCell.getStringCellValue());
	}

	}
	System.out.println();

	}
	} catch (FileNotFoundException e) {
	e.printStackTrace();
	} catch (IOException e) {
	e.printStackTrace();
	}
	}
}

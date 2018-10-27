package utils;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	static String projectPath;
	static XSSFWorkbook wb;
	static XSSFSheet sheet;
	
	public ExcelUtils(String excelPath, String sheetName){
		try {
			wb = new XSSFWorkbook(excelPath);
			sheet = wb.getSheet(sheetName);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	//Get Row count
	public static void getRowCount() {
		try {
			int rows = sheet.getPhysicalNumberOfRows();
			System.out.println("# of rows : " +rows);
			
		} catch (Exception e) {
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}
	}
	
	//Get cell data string
	public static void getCellDataString(int rowNum, int colNum){
		try {
			String cellData = sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
			System.out.println(cellData);
			
		} catch (Exception e) {
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}
	}
	
	//Get cell data numeric
	public static void getCellDataNumeric(int rowNum, int colNum){
		try {
			double cellData = sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
			System.out.println(cellData);
			
		} catch (Exception e) {
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}
	}
}

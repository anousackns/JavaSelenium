package utils;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	
	static String projectPath;
	static XSSFWorkbook wb;
	static XSSFSheet sheet;
	
	public static void main(String[] args) {
		getRowCount();
	}
	
	public static void getRowCount() {
		try {
			projectPath = System.getProperty("user.dir");
			wb = new XSSFWorkbook(projectPath+"\\excel\\Test.xlsx");
			sheet = wb.getSheet("S2");
			int rows = sheet.getPhysicalNumberOfRows();
			
			System.out.println("# of rows : " +rows);
			
		} catch (Exception e) {
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}
	}
	
	public static void getCellData(){
		try {
			projectPath = System.getProperty("user.dir");
			wb = new XSSFWorkbook(projectPath+"\\excel\\Test.xlsx");
			sheet = wb.getSheet("S2");
		} catch (Exception e) {
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}
		
		
		
	}
}

package utils;

public class ExcelUtilsDemo {
	
	public static void main(String[] args) {
		
		String projectPath = System.getProperty("user.dir");
		ExcelUtils excel = new ExcelUtils(projectPath+"/excel/Test.xlsx", "S2");
		
		ExcelUtils.getRowCount();
		ExcelUtils.getCellDataNumeric(1, 6);
		ExcelUtils.getCellDataString(0, 0);
	}

}

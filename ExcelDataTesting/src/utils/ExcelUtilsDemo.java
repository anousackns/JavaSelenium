package utils;

public class ExcelUtilsDemo {
	
	public static void main(String[] args) {
		//set working directory path
		String projectPath = System.getProperty("user.dir");
		//parameters for contructor
		new ExcelUtils(projectPath+"/excel/Test.xlsx", "S2");
		
		ExcelUtils.getRowCount();
		ExcelUtils.getCellDataNumeric(1, 6);
		ExcelUtils.getCellDataString(0, 0);
	}

}

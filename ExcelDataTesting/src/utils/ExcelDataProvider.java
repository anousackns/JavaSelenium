package utils;

public class ExcelDataProvider {
	
	public void testData(String excelPath, String sheetName) {
		
		ExcelUtils excel = new ExcelUtils(excelPath, sheetName);
		
		int rowCount = excel.getRowCount();
		int colCount = excel.getColCount();

		for(int i=1; i < rowCount; i++) {
			for(int j=1; j < colCount; j++) {
				
			}
		}
	}
	
}

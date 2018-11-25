package dataProviders;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;

public class ExcelReader {

	public static String DefaultPath = "/home/machine-ia/Documents/Git/MicromaxInfo/MicromaxInfo/DataSet/";
	// Below method is used for passing all data in sheet, during test.

	@DataProvider(name = "dataForLogin")
	public static Object[][] excelDataProvider1() throws IOException {
		Object[][] obj = getData("document.xlsx", "Sheet1");
		return obj;
	}

	@DataProvider(name = "dataForLogin2")
	public static Object[][] excelDataProvider2() throws IOException {
		Object[][] obj = getData("ExcelName2", "SheetName");
		return obj;
	}

	@SuppressWarnings("resource")
	public static Object[][] getData(String ExcelName, String SheetName) throws IOException  
	{
		//we are using Apache poi because it is more advaced than jxl(last updated 2009) and
		//poi can support xlsx format as well and all office 2007 excel formats
		int i,j;
		int rowFirst, rowLast, rowCount=0, cellCount=0;
		FileInputStream file = new FileInputStream(DefaultPath + ExcelName);
		String[][] Data=null;
		XSSFWorkbook workbook;
		XSSFSheet sheet;
				
		
		workbook = new XSSFWorkbook(file);
		sheet = workbook.getSheet(SheetName);
		//dataFormatter will convert data from excel to String
		//df = new DataFormatter();
		//getFirstRowNum & getLastRowNum() starts counting from 0.
		rowFirst = sheet.getFirstRowNum();
		rowLast = sheet.getLastRowNum();
		rowCount = (rowLast-rowFirst)+1;
		if(rowCount==0)
			System.out.println("Sheet has 0 rows");
		else
		{
			//getLastCellNum starts counting from 1 
			cellCount = sheet.getRow(0).getLastCellNum();
			Data = new String[rowCount][cellCount];
		}
		
		
		for(i=0; i<rowCount; i++)
		{
			for(j=0; j<cellCount; j++)
			{
				XSSFRow row = sheet.getRow(i);
				Data[i][j] = row.getCell(j).getStringCellValue();
			}
		}

		System.out.println("Data array values are :");
		for(String[] array : Data)
			System.out.println(Arrays.toString(array));
			//System.out.println(Arrays.deepToString(Data));
		return Data;
    }
	
}

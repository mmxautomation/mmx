package dataProviders;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelHandler {

	
	private static XSSFWorkbook Workbook;
	private static XSSFSheet Sheet;
	private static XSSFRow Row;
	private static XSSFCell Cell;
	private static String FilePath;
	private static FileInputStream ExcelFile;
	
	
	public ExcelHandler(String Path, String SheetName ) throws IOException
	{
		FilePath = Path;
		ExcelFile = new FileInputStream(Path);
		Workbook = new XSSFWorkbook(ExcelFile);
		Sheet = Workbook.getSheet(SheetName);
	}

	
	public static String getCellData(int RowNum, int ColNum)
	{
		Cell = Sheet.getRow(RowNum).getCell(ColNum);
		String CellData = Cell.getStringCellValue();
		return CellData;
	}
	
	
	public static void setCellData(Object Result, int RowNum, int ColNum) throws IOException
	{
		Row = Sheet.getRow(RowNum);
		Cell = Row.getCell(ColNum);
		if(Cell==null)
		{
			Cell = Row.createCell(ColNum);
			if(Result instanceof String)
				Cell.setCellValue((String)Result);
			if(Result instanceof Integer)
				Cell.setCellValue((Integer) Result);
			}
		
		else 
		{
			if(Result instanceof String)
			Cell.setCellValue((String)Result);
			if(Result instanceof Integer)
			Cell.setCellValue((Integer) Result);
			}
		
		FileOutputStream fileOut = new FileOutputStream(FilePath);
		Workbook.write(fileOut);
		
		
	}
	
}

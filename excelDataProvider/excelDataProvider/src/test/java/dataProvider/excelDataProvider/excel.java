package dataProvider.excelDataProvider;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

/* each  row data in xcel is to be considered as 1 array , 
then we will store the rows or arrays into a multidimensional array so that we can send it from dataprovider(@dataprovider annotation) to the test (@test)*/ 

class excel {
	
	@Test
	public void getExcel() throws IOException
	{
		
		FileInputStream fis=new FileInputStream("C:\\Users\\rahul\\Documents\\excelDriven.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheetAt(0);
		int rowCount=sheet.getPhysicalNumberOfRows();
		XSSFRow row=sheet.getRow(0);
		int colCount=row.getLastCellNum();
		Object data[][]=new Object[rowCount-1][colCount];
		for(int i=0;i<rowCount-1;i++)
		{
			System.out.println("outer loop started");
			row=sheet.getRow(i+1);
			for(int j=0;j<colCount;j++)
			{
				System.out.println(row.getCell(j));
			}
			System.out.println("outer loop Ended");
		}
	}

}

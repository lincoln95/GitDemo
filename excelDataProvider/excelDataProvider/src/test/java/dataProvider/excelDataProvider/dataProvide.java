package dataProvider.excelDataProvider;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

//We are providing the data to test using data provider

//Actual test should read the data from data provider and execute the tests
//if we provide 5 sets of data as 5 arrays from data provider then test will run 5 times 
//then your test will run 5 times with 5 separate set of data
//dataprovider treats one complete array as one test data
//we will add testng dependencies and make this project completely maven portable 

public class dataProvide {

	DataFormatter formatter=new DataFormatter();
	
	
	/*we need to provide the name of the dataprovider we r listening to,i.e "drivetest" as we have mentioned the
    the data provider name as driveTest .Testng @Test will first go tho dataprovider with name drivetest and 
    will execute dat 1st and if any data is getting returned,then it returns to the test.*/
	
	@Test(dataProvider="driveTest") 
	public void testCaseData(String greeting,String communication,String id)
	{
	
	System.out.println(greeting+communication+id);
	}
	
	
	
	@DataProvider(name="driveTest")
	public Object[][] getData() throws IOException
	{
		FileInputStream fis=new FileInputStream("C:\\Users\\BeheraS\\Documents\\excelDriven.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheetAt(0); // getting the sheet present at 0th index ,means the 1st sheet
		int rowCount=sheet.getPhysicalNumberOfRows(); // getting the number of rows present in that sheet 
		XSSFRow row=sheet.getRow(0);
		int colCount=row.getLastCellNum(); // getting the no. of cell having data for a particular row ,which will indirectly give the column count
		Object data[][]=new Object[rowCount-1][colCount]; // declaring multidimensional array data[row][col] with the no. of rows and columns which contains data in xcel
		                                                  //0th row,0th col=hello 0th row 1st column=text 0th row 2nd column=1
		                                                  // similarly 1st row 0th col=bye 1st row 1st col=message 1st row 2nd col=143
		
		//For every outer loop (rowCount,inner loop will  iterate the no. of tyms as columns for that row 
		for(int i=0;i<rowCount-1;i++)   
		{
			row=sheet.getRow(i+1); //as we dunt want the header row   of the sheet so we put i+1
			for(int j=0;j<colCount;j++)
			{
				XSSFCell cell=row.getCell(j); // for each row , it will capture that particular rows cell value and put in it a cell so that we can format it to our desired datatype
				
				data[i][j]=formatter.formatCellValue(cell); //one by one we get each rows each cell data and will put in the multidimentional array
				//one complete outer loop row will be 1 complete array
				/* we are also formatting the cell data using data formatter so that even if the data has some other datatype ,
				   it will be converted to string as we are passing string data as formal parameters */
			
			}
		}
		return data;
		
		
			
		
		
	
		
		
		
		
		
		
	//	return data;
		
		
	}
	
}

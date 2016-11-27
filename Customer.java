package vw;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class Customer {
  
	@Test(dataProvider = "getData")
	public void LoginTest (String sUsername, String sPassword) {
		WebDriver driver = new FirefoxDriver();
		driver.get("http://www.volkswagen.co.uk");
		
		driver.findElement(By.linkText("My VW Login / Signup")).click();
	    driver.findElement(By.id("username")).click();
	    driver.findElement(By.id("username")).clear();
	    driver.findElement(By.id("username")).sendKeys("kajan.222@hotmail.com");
	    driver.findElement(By.id("password")).clear();
	    driver.findElement(By.id("password")).sendKeys("Thili1981");
	    driver.findElement(By.id("login-button")).click();
	    driver.findElement(By.linkText("Find a retailer")).click();
	    driver.findElement(By.id("searchTerm")).clear();
	    driver.findElement(By.id("searchTerm")).sendKeys("london");
	    driver.findElement(By.id("searchSubmit")).click();
	    driver.findElement(By.linkText("[logout]")).click();
	    driver.close();
	}
	
	@DataProvider
	public String[][] getData() throws Exception {   
		File excel = new File("C:/Selenium/VW/VW/TestData4.xls");
		FileInputStream fis = new FileInputStream(excel);
		HSSFWorkbook wb = new HSSFWorkbook(fis);
		HSSFSheet ws = wb.getSheet("TestData");
		int rowNum = ws.getLastRowNum() + 1;
		int colNum = ws.getRow(0).getLastCellNum();
		System.out.println(rowNum);
		System.out.println(colNum);
	
		String[][] data = new String[rowNum][colNum];
		for(int i=1; i<rowNum; i++){
			HSSFRow row = ws.getRow(i);
			for (int j=0; j<colNum; j++){
				HSSFCell cell = row.getCell(j);
				String Value = cellToString(cell);
				data[i][j] = Value;
				//System.out.println(Value);
				
			}
		}
		return data;
	}

	public static String cellToString(HSSFCell cell) {
		// TODO Auto-generated method stub
		int type;
		Object result;
		type = cell.getCellType();
		switch (type) {
		case 0 : //numeric value in excel
		result = cell.getNumericCellValue();
		break;
		case 1 : //string value in excel
		result = cell.getStringCellValue();
		break;
		default :
		throw new RuntimeException("There is no support for this type in POI");
		
		}
		return result.toString();
	}
	
	
}
	
	
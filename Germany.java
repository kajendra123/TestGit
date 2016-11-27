package vw;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

//import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;



import org.openqa.selenium.*;
import org.openqa.selenium.firefox.FirefoxDriver;


// We will learn more about classes in the near future.
public class Germany {
	
  // Global class variables
	  private WebDriver driver; // Object pointing to the Browser Object
	  private String baseUrl;
	  private boolean acceptNextAlert = true;
	  private StringBuffer verificationErrors = new StringBuffer();
  
  int xRows, xCols;
  String xlPath, xlSheet, xlPath_Res;
  String[][] xlData;
  
	// Declare Test Data Variables
	String vTDID, vURL, vEmail, vPassword, vLocation;
	String vExecute;


  // @Before JUnit Annotation 
  @Before
  public void setUp() throws Exception {
    driver = new FirefoxDriver(); // May be an instruction to run on Firefox.
    baseUrl = "http://www.volkswagen.co.uk/"; // Storing the Base URL.
    //driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS); // Timeout before an error occurs.
    
	xlPath = "C:/Selenium/VW/VW/TestData4.xls";
	xlPath_Res = "C:/Selenium/VW/VW/TestData4_Res.xls";
	xlSheet = "TestData";
	
	xlData = readXL(xlPath, xlSheet);

  }

  // @Test Main code to execute our Test Case SW_003
  @Test
  
 
  public void testProject1IDESW003() throws Exception {
	
	// Go through each row within the Test Data
	
	for (int vRow=1; vRow<xRows; vRow++){
		// Get the data and put into the variables accordingly
		// Initialize Test Data Variables
		vTDID = xlData[vRow][0]; 
		vExecute = xlData[vRow][1]; 
		vURL = xlData[vRow][2]; 
		vEmail = xlData[vRow][3]; 
		vPassword = xlData[vRow][4]; 
		vLocation = xlData[vRow][5];
		
		
		if (vExecute.equals("Y")){
			myTC001(vRow);
		} else {
			xlData[vRow][6] = "Skipped";
			System.out.println(xlData);
			
		}
	}
		
    }

  // @After Another JUnit annotation towards the end.
  @After
  public void tearDown() throws Exception {
    driver.close();
	writeXL(xlPath_Res, xlSheet, xlData);
  }

  // IGNORE EVERYTHING DOWN BELOW FOR NOW.
  private boolean isElementPresent(By by) {
    try {
      driver.findElement(by);
      return true;
    } catch (NoSuchElementException e) {
      return false;
    }
  }

  private boolean isAlertPresent() {
    try {
      driver.switchTo().alert();
      return true;
    } catch (NoAlertPresentException e) {
      return false;
    }
  }

  private String closeAlertAndGetItsText() {
    try {
      Alert alert = driver.switchTo().alert();
      String alertText = alert.getText();
      if (acceptNextAlert) {
        alert.accept();
      } else {
        alert.dismiss();
      }
      return alertText;
    } finally {
      acceptNextAlert = true;
    }
  }
  
  // Teach Java to R/W from MS Excel
	// Method to read XL
	public String[][] readXL(String fPath, String fSheet) throws Exception{
		// Inputs : XL Path and XL Sheet name
		// Output : 
		
			String[][] xData;   

			File myxl = new File(fPath);                                
			FileInputStream myStream = new FileInputStream(myxl);                                
			HSSFWorkbook myWB = new HSSFWorkbook(myStream);                                
			HSSFSheet mySheet = myWB.getSheet(fSheet);                                 
			xRows = mySheet.getLastRowNum()+1;                                
			xCols = mySheet.getRow(0).getLastCellNum();   
			System.out.println("Total Rows in Excel are " + xRows);
			System.out.println("Total Cols in Excel are " + xCols);
			xData = new String[xRows][xCols];        
			for (int i = 0; i < xRows; i++) {                           
					HSSFRow row = mySheet.getRow(i);
					for (int j = 0; j < xCols; j++) {                               
						HSSFCell cell = row.getCell(j);
						String value = "-";
						if (cell!=null){
							value = cellToString(cell);
						}
						xData[i][j] = value;      
						System.out.print(value);
						System.out.print("----");
						}        
					System.out.println("");
					}    
			myxl = null; // Memory gets released
			return xData;
	}
	
	//Change cell type
	public static String cellToString(HSSFCell cell) { 
		// This function will convert an object of type excel cell to a string value
		int type = cell.getCellType();                        
		Object result;                        
		switch (type) {                            
			case HSSFCell.CELL_TYPE_NUMERIC: //0                                
				result = cell.getNumericCellValue();                                
				break;                            
			case HSSFCell.CELL_TYPE_STRING: //1                                
				result = cell.getStringCellValue();                                
				break;                            
			case HSSFCell.CELL_TYPE_FORMULA: //2                                
				throw new RuntimeException("We can't evaluate formulas in Java");  
			case HSSFCell.CELL_TYPE_BLANK: //3                                
				result = "%";                                
				break;                            
			case HSSFCell.CELL_TYPE_BOOLEAN: //4     
				result = cell.getBooleanCellValue();       
				break;                            
			case HSSFCell.CELL_TYPE_ERROR: //5       
				throw new RuntimeException ("This cell has an error");    
			default:                  
				throw new RuntimeException("We don't support this cell type: " + type); 
				}                        
		return result.toString();      
		}
	
	// Method to write into an XL
	public void writeXL(String fPath, String fSheet, String[][] xData) throws Exception{

	    	File outFile = new File(fPath);
	        HSSFWorkbook wb = new HSSFWorkbook();
	        HSSFSheet osheet = wb.createSheet(fSheet);
	        int xR_TS = xData.length;
	        int xC_TS = xData[0].length;
	    	for (int myrow = 0; myrow < xR_TS; myrow++) {
		        HSSFRow row = osheet.createRow(myrow);
		        for (int mycol = 0; mycol < xC_TS; mycol++) {
		        	HSSFCell cell = row.createCell(mycol);
		        	cell.setCellType(HSSFCell.CELL_TYPE_STRING);
		        	cell.setCellValue(xData[myrow][mycol]);
		        }
		        FileOutputStream fOut = new FileOutputStream(outFile);
		        wb.write(fOut);
		        fOut.flush();
		        fOut.close();
	    	}
		}
	
	public void myTC001 (int fRow) throws InterruptedException{
		// Run the Test Data on the AUT
		System.out.println("Now executing Test Data : " + vTDID);
		
	    driver.get(vURL); // Go to Volkawagen.co.uk
	    
	    driver.manage().window().maximize();
	    Thread.sleep(13000);
	    
	    driver.findElement(By.linkText("My VW Login / Signup")).click();
	    Thread.sleep(5000);
	    driver.findElement(By.id("username")).click();
	    driver.findElement(By.id("username")).clear();
	    driver.findElement(By.id("username")).sendKeys(vEmail);
	    
	    driver.findElement(By.id("password")).clear();
	    driver.findElement(By.id("password")).sendKeys(vPassword);
	    Thread.sleep(2000);
	    
	    driver.findElement(By.id("login-button")).click();
	    Thread.sleep(2000);
	    driver.findElement(By.linkText("Find a retailer")).click();
	    Thread.sleep(2000);
	    
	    driver.findElement(By.id("searchTerm")).clear();
	    driver.findElement(By.id("searchTerm")).sendKeys(vLocation);
	    
	    driver.findElement(By.id("searchSubmit")).click();
	    Thread.sleep(2000);
	    driver.findElement(By.linkText("[logout]")).click();
	    Thread.sleep(5000);
	    
	  
	   xlData[fRow][6] = "Pass";
	   
	   System.out.println(xlData);
	    
	    

	}
	
}

package vw;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class Login {

	public static void main(String[] args) throws InterruptedException {
		// TODO Auto-generated method stub

		WebDriver driver = new FirefoxDriver();
		driver.manage().window().maximize();
		
		driver.get("http://www.volkswagen.co.uk");
		Thread.sleep(10000);
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

}

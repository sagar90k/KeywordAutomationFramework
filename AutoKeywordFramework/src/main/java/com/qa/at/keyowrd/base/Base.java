package com.qa.at.keyowrd.base;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class Base {

	public WebDriver driver;
	public Properties prop;

	public WebDriver init_driver(String browserName) {
		
		if (browserName.equals("chrome")) 
		{
				System.setProperty("webdriver.chrome.driver","C:\\Users\\AMAR\\git\\KeywordAutomationFramework\\AutoKeywordFramework\\src\\main\\resources\\drivers\\chromedriver.exe");			
				driver = new ChromeDriver();	
				
				ChromeOptions options = new ChromeOptions();
				options.addArguments("disable-infobars"); 
		}
		return driver;
}

	
	
	public Properties init_properties() {
		
		prop = new Properties();
		
		try {
			FileInputStream ip = new FileInputStream(
					"C:\\Users\\AMAR\\git\\KeywordAutomationFramework\\AutoKeywordFramework\\src\\main\\java\\com\\qa\\at\\keyword\\config\\config.properties");
			prop.load(ip);			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return prop;
	}

}

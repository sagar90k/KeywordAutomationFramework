package com.qa.at.keyowrd.base;

import java.io.FileInputStream;
import java.io.File;
import java.io.IOException;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Properties;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import java.util.*;

public class Base {

	public WebDriver driver;
	public Properties prop;

	public WebDriver init_driver(String browserName) {

		if (browserName.equals("chrome")) {
			System.setProperty("webdriver.chrome.driver",
					"C:\\Users\\AMAR\\git\\KeywordAutomationFramework\\AutoKeywordFramework\\src\\main\\resources\\drivers\\chromedriver.exe");

			ChromeOptions options = new ChromeOptions();
			options.setExperimentalOption("excludeSwitches", Collections.singletonList("enable-automation"));
			options.setExperimentalOption("useAutomationExtension", false);

			driver = new ChromeDriver(options);

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

	public void take_screenshot(WebDriver driver, String ssname, XWPFDocument document ) {
		
		File src = null;
		
		 src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		

		try {
			FileUtils.copyFile(src, new File("F:\\Automation Study\\screenshots\\" + ssname + ".png"));
			
			
		} catch (IOException e) {
			System.out.println(e.getMessage());

		}
		String fsrc ="F:\\Automation Study\\screenshots\\" + ssname + ".png";
	
		add_to_word_doc(ssname, fsrc, document);
		
	}

	public void add_to_word_doc(String ssname, String fsrc, XWPFDocument document) {
		
		
		// Write the Document in file system
		FileOutputStream doc_out = null;
		try {
			doc_out = new FileOutputStream(new File("F:\\Automation Study\\screenshots\\" + ssname + ".docx"));
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		XWPFParagraph paragraph = document.createParagraph();
		XWPFRun run = paragraph.createRun();
		run.setText(" TC:" + ssname);
		run.addCarriageReturn();
		
		String imgFile = fsrc;

		// run.addPicture(sip_stream, pictureType, filename, width, height)
		FileInputStream fins=null;
		try {
			fins = new FileInputStream(imgFile);
		} catch (FileNotFoundException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}
		try {
			run.addPicture(fins, Document.PICTURE_TYPE_PNG, ssname, Units.toEMU(500), Units.toEMU(500));
			
			
		} catch (InvalidFormatException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		try {
			
			document.write(doc_out);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		System.out.println("createparagraph.docx written successfully");
		try {
			fins.close();
			doc_out.close();
			imgFile = null;
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		};
	}

	/*
	 * 
	 * public void wite_doc(FileOutputStream doc_out,String ssname){
	 * 
	 * File fsrc= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
	 * 
	 * try { FileUtils.copyFile(fsrc, new File(
	 * "F:\\Automation Study\\screenshots\\"+ssname+".png"));
	 * 
	 * } catch (IOException e) { System.out.println(e.getMessage());
	 * 
	 * }
	 * 
	 * XWPFDocument document = new XWPFDocument(); XWPFParagraph paragraph =
	 * document.createParagraph(); XWPFRun run = paragraph.createRun();
	 * run.setText(" TC:" + ssname); //run.addPicture(pictureData, pictureType,
	 * filename, width, height)
	 * 
	 * try { document.write(doc_out); } catch (IOException e) { // TODO
	 * Auto-generated catch block e.printStackTrace(); }
	 * 
	 * }
	 */
}

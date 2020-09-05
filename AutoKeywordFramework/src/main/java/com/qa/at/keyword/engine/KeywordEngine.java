package com.qa.at.keyword.engine;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Timestamp;
import java.util.Properties;
import java.util.concurrent.TimeUnit;
import java.util.Date;
import java.util.Date.*;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.qa.at.keyowrd.base.Base;

public class KeywordEngine {

	public WebDriver driver;
	public Properties prop;

	public static Workbook book;
	public static Sheet teststep_sheet, scenario_sheet;

	public Base base;

	public WebElement element;

	public final String SCENARIO_SHEET_PATH = "C:\\Users\\AMAR\\git\\KeywordAutomationFramework\\AutoKeywordFramework\\src\\main\\java\\com\\qa\\at\\keyword\\scenarios\\scenarios.xlsx";

	public void startExecution(String test_Step_SheetName, String scenarios_Sheet_Name) {

		String locatorType = null;
		String locatorValue = null;

		FileInputStream file = null;
		
		FileOutputStream fos =null; 
		

		try {
			file = new FileInputStream(SCENARIO_SHEET_PATH);
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		try {
			book = WorkbookFactory.create(file);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}


		scenario_sheet = book.getSheet(scenarios_Sheet_Name);
		teststep_sheet = book.getSheet(test_Step_SheetName);

		Date date = new Date();
		Timestamp cst= new Timestamp(date.getTime());

		int sccol = 0;
		String testcase_number;
		int testcase_stepnumber = 0;
		String testcase_name_to_refer;
		String test_data_col_to_refer;
		String doc_name = null;
		int test_to_execute_or_not = 0;

		String test_start_row = null;
		String test_end_row = null;
		String ssname = null;
		FileOutputStream doc_op = null;
		int start_row = 0;
		int end_row = 0;
		int test_or_not = 0;
		int data_to_refer = 0;

		for (int j = 0; j < scenario_sheet.getLastRowNum(); j++) {

			XWPFDocument passdocument = new XWPFDocument();
			XWPFDocument faildocument = new XWPFDocument();
			
			
			
			
			testcase_name_to_refer = scenario_sheet.getRow(j + 1).getCell(sccol + 1).toString().trim();
			testcase_number = scenario_sheet.getRow(j + 1).getCell(sccol + 0).toString().trim();
			test_or_not = (int) scenario_sheet.getRow(j + 1).getCell(sccol + 2).getNumericCellValue();
			start_row = (int) scenario_sheet.getRow(j + 1).getCell(sccol + 3).getNumericCellValue();
			end_row = (int) scenario_sheet.getRow(j + 1).getCell(sccol + 4).getNumericCellValue();
			test_data_col_to_refer = scenario_sheet.getRow(j + 1).getCell(sccol + 5).toString().trim();
			
			// base.create_results_word_doc(ssname);
			// base.create_results_word_doc(testcase_name_to_refer);

			try {
				if (test_or_not == 1) {

					int k = 0;
					for (int i = start_row; i <= end_row; i++) {

						try {
							ssname = "TCNo-" + testcase_number + "-" + testcase_name_to_refer;
							doc_name=ssname;
							testcase_stepnumber = (int) teststep_sheet.getRow(i).getCell(k + 1).getNumericCellValue();
								
								
							locatorType = teststep_sheet.getRow(i).getCell(k + 3).toString().trim();
							locatorValue = teststep_sheet.getRow(i).getCell(k + 4).toString().trim();

							String action = teststep_sheet.getRow(i).getCell(k + 5).toString().trim();
							String value = teststep_sheet.getRow(i)
									.getCell(k + Integer.parseInt(test_data_col_to_refer)).toString();

							// System.out.println("locatorValue:"+locatorValue);

							switch (locatorType) {

							case "xpath":

								WebDriverWait wait1 = new WebDriverWait(driver, 10);
								element = wait1.until(ExpectedConditions.elementToBeClickable(By.xpath(locatorValue)));

								// element =
								// driver.findElement(By.className(locatorValue));
								locatorType = null;
								break;

							case "id":

								WebDriverWait wait2 = new WebDriverWait(driver, 10);
								element = wait2.until(ExpectedConditions.elementToBeClickable(By.id(locatorValue)));

								// element =
								// driver.findElement(By.className(locatorValue));
								locatorType = null;
								break;

							case "linkText":
								WebDriverWait wait3 = new WebDriverWait(driver, 10);
								element = wait3.until(ExpectedConditions.elementToBeClickable(By.linkText(locatorValue)));

								//element = driver.findElement(By.linkText(locatorValue));
								element.click();
								locatorType = null;

								break;
								
							case "class":
								WebDriverWait wait4 = new WebDriverWait(driver, 10);
								element = wait4.until(ExpectedConditions.elementToBeClickable(By.className(locatorValue)));
						
								element.click();
								locatorType = null;

								break;
								
							default:
								break;

							}

							switch (action) {

							case "open browser":
								base = new Base();
								prop = base.init_properties();
								if (value.isEmpty() || value.equals("NA")) {
									driver = base.init_driver(prop.getProperty("browser"));
								} else {
									driver = base.init_driver(value);
								}
								break;

							case "sendkeys":
								if (action.equalsIgnoreCase("sendkeys")) {
									element.sendKeys(value);
								} else if (action.equalsIgnoreCase("click")) {
									element.click();
								}
								break;

							case "click":
								if (action.equalsIgnoreCase("click")) {
									element.click();
								}
								break;

							case "enter url":
								if (value.isEmpty() || value.equals("NA")) {
									driver.get(prop.getProperty("url"));
								} else {
									driver.get(value);
								}
								break;

							case "verifyOutputValue":

								String Elementvalue = element.getText();

								if (value.isEmpty() || value.equals("NA")) {
									continue;
								} else if (value.equals(Elementvalue)) {
									System.out.println("call Set_status('Passed')");
									System.out.println("call take_screenshot()");

									ssname = ssname + "-TS-" + testcase_stepnumber + " - Passed";
									base.take_screenshot(driver,doc_name,ssname,passdocument);
									cst= new Timestamp(date.getTime());
									
									//setCellType(cell.CELL_TYPE_STRING);
									//cell.setCellValue("SoftwareTestingMaterial.com");
									
									teststep_sheet.getRow(i).getCell(k + 7).setCellValue("Pass");
									teststep_sheet.getRow(i).getCell(k + 8).setCellValue(cst);

								} else {
									System.out.println("call Set_status('Failed')");
									System.out.println("call take_screenshot()");
									
								
									
									ssname = ssname + "-TS-" + testcase_stepnumber + " - Failed";
									base.take_screenshot(driver, doc_name,ssname, faildocument);
									try {
										faildocument.close();
									} catch (IOException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}

								}
								ssname = null;
								break;
								
							case "verifyOutputElementPresent":
								
								if (value.isEmpty() || value.equals("NA") && (element.isDisplayed())) 
								{
									 
									System.out.println("call Set_status('Passed')");
									System.out.println("call take_screenshot()");

									ssname = ssname + "-TS-" + testcase_stepnumber + " - Passed";
									base.take_screenshot(driver,doc_name, ssname,passdocument);

								} else {
									System.out.println("call Set_status('Failed')");
									System.out.println("call take_screenshot()");
									
									ssname = ssname + "-TS-" + testcase_stepnumber + " - Failed";
									base.take_screenshot(driver,doc_name, ssname, faildocument);
									try {
										faildocument.close();
									} catch (IOException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
									
								}
								ssname = null;
								
								break;
								
							default:
								break;
							}

						}
						catch (Exception e) {
							e.printStackTrace();
						}

					}
				}

				else {
					continue;
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
			
			try {
				passdocument.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			try {
				book.write(fos);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
		}
				
	}
}

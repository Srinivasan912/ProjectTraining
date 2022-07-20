package org.utility;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {
	
	public static WebDriver driver;
	
	public static WebDriver launchBrowser(String browsername) {
		if(browsername.equalsIgnoreCase("chrome")) {
			WebDriverManager.chromedriver().setup();
			driver= new ChromeDriver();	
		}
		else if(browsername.equalsIgnoreCase("firefox")) {
			WebDriverManager.firefoxdriver().setup();
			driver= new FirefoxDriver();
		}
		else if(browsername.equalsIgnoreCase("edge")) {
			WebDriverManager.edgedriver().setup();
			driver= new EdgeDriver();
		}
		return driver;
		}
	
	public static void maximize() {
		driver.manage().window().maximize();
	}
	
	public static void implicitwait(long secs) {
		driver.manage().timeouts().implicitlyWait(secs, TimeUnit.SECONDS);
	}
	
	public static void urlLaunch(String url) {
		driver.get(url);
	}
	
	public static void sendkeys(WebElement e,String value) {
		e.sendKeys(value);
	}
	
	public static void click(WebElement e) {
		e.click();
	}
	
	public static String currentUrl() {
		String url = driver.getCurrentUrl();
		return url;
	}
	
	public static String title() {
		String title = driver.getCurrentUrl();
		return title;
	}
	
	public static void quit() {
		driver.quit();
	}
	
	public static String  getAttribute(WebElement e) {
		String att=e.getAttribute("value");
		return att;		
	}
	
	public static void moveToElement(WebElement target) {
		Actions a = new Actions(driver);
		a.moveToElement(target).perform();
	}
	
	public static void selectByIndex(WebElement e,int index) {
		Select s=new Select(e);
		s.selectByIndex(index);
	}
	public static void tabKey() throws AWTException {
		Robot r = new Robot();
		r.keyPress(KeyEvent.VK_TAB);
		r.keyRelease(KeyEvent.VK_TAB);
	}
	
	public static void createExcel(String filename, String sheetname, int row, int cell, String givenvalue) throws IOException {
		File f = new File("D:\\workplace\\eclipse-wrokspace\\Maven-Demo\\src\\test\\resources\\"+filename+".xlsx");
		Workbook w = new XSSFWorkbook();
		Sheet cs =w.createSheet(sheetname);
		Row cr =cs.createRow(row);
		Cell cc= cr.createCell(cell);
		cc.setCellValue(givenvalue);
		FileOutputStream fo = new FileOutputStream(f);
		w.write(fo);
	}
	
	public static void addExcelCellValues(String filename, String sheetname, int row, int cell, String givenvalue) throws IOException {
		File f = new File("D:\\workplace\\eclipse-wrokspace\\Maven-Demo\\src\\test\\resources\\"+filename+".xlsx");
		FileInputStream fi = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fi);
		Sheet cs =w.getSheet(sheetname);
		Row cr =cs.getRow(row);
		Cell cc= cr.createCell(cell);
		cc.setCellValue(givenvalue);
		FileOutputStream fo = new FileOutputStream(f);
		w.write(fo);
	}
	
	public static void addExcelRowValues(String filename, String sheetname, int row, int cell, String givenvalue) throws IOException {
		File f = new File("D:\\workplace\\eclipse-wrokspace\\Maven-Demo\\src\\test\\resources\\"+filename+".xlsx");
		FileInputStream fi = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fi);
		Sheet cs =w.getSheet(sheetname);
		Row cr =cs.createRow(row);
		Cell cc= cr.createCell(cell);
		cc.setCellValue(givenvalue);
		FileOutputStream fo = new FileOutputStream(f);
		w.write(fo);
	}
	
	public static String getExcel(String filename, String sheetname, int row, int cell) throws IOException {
		File f = new File("D:\\workplace\\eclipse-wrokspace\\Maven-Demo\\src\\test\\resources\\"+filename+".xlsx");
		FileInputStream fi = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fi);
		Sheet s =w.getSheet(sheetname);
		Row r =s.getRow(row);
		Cell c =r.getCell(cell);
		int type = c.getCellType();
		String value;
		if(type==1) {
			value=c.getStringCellValue();
		}
		else {
			if(DateUtil.isCellDateFormatted(c)) {
				Date d = c.getDateCellValue();
				SimpleDateFormat si = new SimpleDateFormat("dd-MM-yy");
				value = si.format(d);
			}
			else {
				double d = c.getNumericCellValue();
				long l=(long) d;
				value =String.valueOf(l);
			}
		}
		return value;
	}
}

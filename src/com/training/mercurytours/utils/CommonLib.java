package com.training.mercurytours.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.tools.ant.types.CommandlineJava.SysProperties;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;
import org.testng.Reporter;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Parameters;

public class CommonLib {

	public static WebDriver oBrowser;

	public static String openBrowser(String sBrowserName) {
		try {
			switch (sBrowserName.toUpperCase()) {
			case "IE":
			case "INTERNET EXPLORER": {
				System.setProperty("webdriver.ie.driver",
						Constants.sDriverPath+"IEDriverServer.exe");
				oBrowser = new InternetExplorerDriver();
				break;
			}
			case "CHROME": {
				System.setProperty("webdriver.chrome.driver",
						Constants.sDriverPath+"chromedriver.exe");
				oBrowser = new ChromeDriver();
				break;
			}
			case "FF":
			case "FIREFOX": {
				System.setProperty("webdriver.gecko.driver",
						Constants.sDriverPath+"geckodriver.exe");
				oBrowser = new FirefoxDriver();
				break;
			}
			default: {
				System.setProperty("webdriver.chrome.driver",
						Constants.sDriverPath+"chromedriver.exe");
				oBrowser = new ChromeDriver();
				break;
			}
			}
			return "Pass";
	
		} catch (Throwable t) {
			System.out.println("Unable to launch "+sBrowserName+" browser");
			takeScreenshot("OpenBrowser.png");
			return "Failed"; 
		}

	}

	public static String navigateToURL(String sURL) {
		try {
			oBrowser.get(sURL);
			oBrowser.manage().window().maximize();
			//oBrowser.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
			oBrowser.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
			oBrowser.manage().timeouts().setScriptTimeout(60, TimeUnit.SECONDS);
			return "Pass";
	} catch (Throwable t) {
			takeScreenshot("navigateURL.png");
			System.out.println("Navigating URL is failed");
			return "Failed";
		}

	}
	
	public static void pauseScript(int iSec)
	{
		try {
			Thread.sleep(iSec*1000);
		
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		
	}
	
	
	public static String closeBrowser() {
		
		try
		{
		oBrowser.close();
		oBrowser.quit();
		return "Pass";
		}
		catch(Throwable t)
		{
			takeScreenshot("CloseBrowser.png");
			System.out.println("Unable to close the browser");
			return "Fail";
		}
	}
	
	
	public static void takeScreenshot(String sFileName) {
		try
		{
		File ScrFile = ((TakesScreenshot) oBrowser)
				.getScreenshotAs(OutputType.FILE);
		FileUtils.copyFile(ScrFile, new File(Constants.sReportPath+sFileName));
		}
		catch(Throwable t)
		{
			System.out.println("Unable to take screenshot");
		}
	}
	
	
	public static void selectValFromListbox(WebElement oElem,String sVal)
	{
		try
		{
			Select oLstBox=new Select(oElem);
			oLstBox.selectByVisibleText(sVal);
		}
		catch(Throwable t)
		{
			System.out.println(t.getMessage());
		}
	}
	
	public static void enterText(By oBy,String sVal)
	{
		try
		{
			oBrowser.findElement(oBy).clear();
			oBrowser.findElement(oBy).sendKeys(sVal);
		}
		catch(Throwable t)
		{
			takeScreenshot("EnterText.png");
			System.out.println("Error while entering text into text box");
		}
	}
	
	public static void clickElem(By oBy)
	{
		try
		{
			oBrowser.findElement(oBy).click();
			
		}
		catch(Throwable t)
		{
			takeScreenshot("ClickElem.png");
			System.out.println(t.getMessage());
		}
	}

	
	public static void switchTOFrame(String sID)
	{
		try
		{
		oBrowser.switchTo().frame(sID);
		}
		catch(Throwable e)
		{
			System.out.println(e.getMessage());
		}
	}
	
	public static void handleAlert()
	{
		try
		{
			Alert oAlt=oBrowser.switchTo().alert();
			oAlt.accept();
			
		}
		catch(Throwable t)
		{
			takeScreenshot("HandleAlerts");
			System.out.println(t.getMessage());
		}
	}
	
	
	public static String getDateTimeStamp()
	{
		try
		{
			Date oDate = new Date();
			SimpleDateFormat oDateFormat = new SimpleDateFormat("YYYY_MM_dd_hh_mma");
			String sStamp;
			sStamp = oDateFormat.format(oDate);
			System.out.println("Current Date Time Stamp =" + sStamp);
			return sStamp;

		}
		catch(Throwable t)
		{
			System.out.println(t.getMessage());
			return null;
		}
	}
	
	public static void clickOnElemByusingName(List<WebElement> oElems,String sVal)
	{
		try
		{
			for(WebElement oElem:oElems)
			{
				String sElemName=oElem.getAttribute("value").replace("$", " ");
				if(sElemName.contains(sVal))
				{
				  oElem.click();
				  break;
				}
			}
		}
		catch(Throwable t)
		{
			System.out.println(t.getMessage());
		}
		
	}
	
	
	public static Properties readDataFromPropertiesFile(String sFilePath)
	{
		try
		{
		FileInputStream oFIP=new FileInputStream(sFilePath);
		Properties oProp=new Properties();
		oProp.load(oFIP);
		return oProp;
		}
		catch(Exception e)
		{
			System.out.println(e.getMessage());
			return null;
		}
	}
	
	
	public boolean isElemDisplayed(WebElement oElem)
	{
		return oElem.isDisplayed();
	}
	
	public WebElement getElem(By oBy)
	{
		return oBrowser.findElement(oBy);
	}
	
	public List<WebElement> getElems(By oBy)
	{
		return oBrowser.findElements(oBy);
	}
	
	

	

}

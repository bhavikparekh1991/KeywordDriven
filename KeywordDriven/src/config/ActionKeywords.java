package config;

import java.util.concurrent.TimeUnit;

import static executionEngine.DriverScript.OR;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;

import executionEngine.DriverScript;
import utility.Log;

public class ActionKeywords {
	public static WebDriver driver;
	public static WebElement getLocator(String object, String data) throws Exception
	{
		/*System.setProperty("webdriver.chrome.driver", "C:/Users/pbhavik/Downloads/chromedriver_win32/chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.get("https://www.google.co.in/?gfe_rd=ctrl&amp;ei=bXAwU8jYN4W6iAf8zIDgDA&amp;gws_rd=cr");*/
		
		String locatorType = object.split("_")[0];
		String locatorValue = object.split("_")[1];
		
		if(locatorType.toLowerCase().equals("id"))
		return driver.findElement(By.id(locatorValue));
		else if(locatorType.toLowerCase().equals("name"))
			return driver.findElement(By.name(locatorValue));
		else if((locatorType.toLowerCase().equals("classname")) || (locatorType.toLowerCase().equals("class")))
			return driver.findElement(By.className(locatorValue));
		else if((locatorType.toLowerCase().equals("tagname")) || (locatorType.toLowerCase().equals("tag")))
			return driver.findElement(By.tagName(locatorValue));
		else if((locatorType.toLowerCase().equals("linktext")) || (locatorType.toLowerCase().equals("link")))
			return driver.findElement(By.linkText(locatorValue));
		else if(locatorType.toLowerCase().equals("partiallinktext"))
			return driver.findElement(By.partialLinkText(locatorValue));
		else if((locatorType.toLowerCase().equals("cssselector")) || (locatorType.toLowerCase().equals("css")))
			return driver.findElement(By.cssSelector(locatorValue));
		else if(locatorType.toLowerCase().equals("xpath"))
			return driver.findElement(By.xpath(locatorValue));
		else
			throw new Exception("Unknown locator type '" + locatorType + "'");
		
	}
	
	public static WebElement getWebElement(String object, String data) throws Exception{
		return getLocator(OR.getProperty(object), data);
	}
	
	/*public static void openBrowser(String object){		
		try
		{
			Log.info("Opening Browser");
			String driverPath = "C:/Users/pbhavik/Downloads/geckodriver-v0.14.0-win64/geckodriver.exe";
			System.setProperty("webdriver.firefox.marionette", driverPath);
			driver = new FirefoxDriver();
			driver.manage().window().maximize();
		}
		catch(Exception e)
		{
			Log.info("Not able to open Browser --- " + e.getMessage());
			DriverScript.bResult=false;
		}
	}*/
	
	//This block of code will decide which browser type to start
	public static void openBrowser(String object,String data){		
			Log.info("Opening Browser");
			try{
				//If value of the parameter is Mozilla, this will execute
				if(data.equals("Mozilla")){
					String driverPath = "C:/Users/PBHAVIK/Downloads/geckodriver-v0.14.0-win64/geckodriver.exe";
					
																//geckodriver-v0.16.1-win64
					System.setProperty("webdriver.firefox.marionette", driverPath);
					driver = new FirefoxDriver();
					driver.manage().window().maximize();
					Log.info("Mozilla browser started");
					}
				else if(data.equals("IE")){
					//You may need to change the code here to start IE Driver
					driver=new InternetExplorerDriver();
					Log.info("IE browser started");
					}
				else if(data.equals("Chrome")){
					
					 System.setProperty("webdriver.chrome.driver", "C:/Users/PBHAVIK/Downloads/chromedriver_win32/chromedriver.exe");
					 driver=new ChromeDriver();
					Log.info("Chrome browser started");
					}

				int implicitWaitTime=(10);
				driver.manage().timeouts().implicitlyWait(implicitWaitTime, TimeUnit.SECONDS);
			}catch (Exception e){
				Log.info("Not able to open the Browser --- " + e.getMessage());
				DriverScript.bResult = false;
			}
		}

	public static void navigate(String object, String data){
			try{
				Log.info("Navigating to URL");
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				//Constant Variable is used in place of URL
				driver.get(Constants.URL);
			}catch(Exception e){
				Log.info("Not able to navigate --- " + e.getMessage());
				DriverScript.bResult = false;
				}
			}

	public static void click(String object, String data) throws Exception{
		try
		{
			//driver.findElement(By.xpath(OR.getProperty(object))).click();
			Log.info("Click on WebElement :" + object);
			getWebElement(object, data).click();	
		}
		catch(Exception e)
		{
			Log.info("Not able to click --- " + e.getMessage());
			DriverScript.bResult=false;
		}
	}

	/*public static void input_UserName(String object, String data) throws Exception{
		try
		{	Log.info("Entering the text in UserName");
			getWebElement(object, data).sendKeys(Constants.UserName);
		}
		catch(Exception e)
		{
			Log.info("Not able to enter UserName --- " + e.getMessage());
			DriverScript.bResult=false;
		}
	}*/

	//Now this method accepts two value (Object name & Data)
	public static void input(String object, String data){
			try{
				Log.info("Entering the text in " + object);
				driver.findElement(By.xpath(OR.getProperty(object))).sendKeys(data);
			 }catch(Exception e){
				 Log.error("Not able to Enter UserName --- " + e.getMessage());
				 DriverScript.bResult = false;
			 	}
			}
	
	/*public static void input_Password(String object, String data) throws Exception{
		try
		{
			Log.info("Entering the text in Password");
			getWebElement(object, data).sendKeys(Constants.Password);
		}
		catch(Exception e)
		{
			Log.info("Not able to enter Password --- " + e.getMessage());
			DriverScript.bResult=false;
		}
	}*/

	public static void waitfor(String object,  String data) throws Exception{
		try
		{
			Log.info("Wait for 5 seconds");
			Thread.sleep(5000);
		}
		catch(Exception e)
		{
			Log.info("Not able to wait --- " + e.getMessage());
			DriverScript.bResult=false;
		}
	}

	public static void closeBrowser(String object,  String data){
		try
		{
			Log.info("Closing the browser");
			driver.quit();
		}
		catch(Exception e)
		{
			Log.info("Not able to close the Browser --- " + e.getMessage());
			DriverScript.bResult=false;
		}
	}

}

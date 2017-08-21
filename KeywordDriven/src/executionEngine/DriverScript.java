package executionEngine;

import java.io.FileInputStream;
import java.lang.reflect.Method;
import java.util.Properties;
//import java.util.concurrent.TimeUnit;


import org.apache.log4j.xml.DOMConfigurator;
//import org.openqa.selenium.By;
//import org.openqa.selenium.WebDriver;
//import org.openqa.selenium.firefox.FirefoxDriver;

import org.openqa.selenium.WebDriver;

import config.ActionKeywords;
import config.Constants;
import utility.ExcelUtils;
import utility.Log;

public class DriverScript {
	
	//private static final int Col_TestStepFailScreenShot = 0;
	public static Properties OR;
	public static ActionKeywords actionKeywords;
	public static String sActionKeyword;
	public static String sPageObject;
	public static Method method[];
	public static int iTestStep;
	public static int iTestLastStep;
	public static String sTestCaseID;
	public static String sRunMode;
	public static String sData;
	public static boolean bResult;
	public DriverScript() throws NoSuchMethodException, SecurityException{
		actionKeywords = new ActionKeywords();
		method = actionKeywords.getClass().getMethods();		
	}

    public static void main(String[] args) throws Exception {
    	ExcelUtils.setExcelFile(Constants.Path_TestData);
    	DOMConfigurator.configure("log4j.xml");
    	String Path_OR = Constants.Path_OR;
		FileInputStream fs = new FileInputStream(Path_OR);
		OR= new Properties(System.getProperties());
		OR.load(fs);

		DriverScript startEngine = new DriverScript();
		startEngine.execute_TestCase();
    }

    private void execute_TestCase() throws Exception {
		//This will return the total number of test cases mentioned in the Test cases sheet
    	int iTotalTestCases = ExcelUtils.getRowCount(Constants.Sheet_TestCases);
		//This loop will execute number of times equal to Total number of test cases
		for(int iTestcase=1;iTestcase<=iTotalTestCases-1;iTestcase++){
			//This is to get the Test case name from the Test Cases sheet
			
			bResult = true;
			sTestCaseID = ExcelUtils.getCellData(iTestcase, Constants.Col_TestCaseID, Constants.Sheet_TestCases); 
			//This is to get the value of the Run Mode column for the current test case
			sRunMode = ExcelUtils.getCellData(iTestcase, Constants.Col_RunMode,Constants.Sheet_TestCases);
			//This is the condition statement on RunMode value
			if (sRunMode.equals("Yes")){
				//Only if the value of Run Mode is 'Yes', this part of code will execute
				iTestStep = ExcelUtils.getRowContains(sTestCaseID, Constants.Col_TestCaseID, Constants.Sheet_TestSteps);
				iTestLastStep = ExcelUtils.getTestStepsCount(Constants.Sheet_TestSteps, sTestCaseID, iTestStep);
				Log.startTestCase(sTestCaseID);
				//Setting the value of bResult variable to 'true' before starting every test step
				bResult=true;
				//This loop will execute number of times equal to Total number of test steps
				for (;iTestStep<=iTestLastStep;iTestStep++){
		    		sActionKeyword = ExcelUtils.getCellData(iTestStep, Constants.Col_ActionKeyword,Constants.Sheet_TestSteps);
		    		sPageObject = ExcelUtils.getCellData(iTestStep, Constants.Col_PageObject, Constants.Sheet_TestSteps);
		    		sData = ExcelUtils.getCellData(iTestStep, Constants.Col_DataSet, Constants.Sheet_TestSteps);
		    		execute_Actions();
		    		
					if(bResult==false){
						//If 'false' then store the test case result as Fail
						ExcelUtils.setCellData(Constants.KEYWORD_FAIL,iTestcase,Constants.Col_Result,Constants.Sheet_TestCases);
						//End the test case in the logs
						Log.endTestCase(sTestCaseID);
						//By this break statement, execution flow will not execute any more test step of the failed test case
						break;
						}

					}
				//This will only execute after the last step of the test case, if value is not 'false' at any step	
				if(bResult==true){
				//Storing the result as Pass in the excel sheet
				ExcelUtils.setCellData(Constants.KEYWORD_PASS,iTestcase,Constants.Col_Result,Constants.Sheet_TestCases);
				Log.endTestCase(sTestCaseID);	
					}
				}
    		}
    	}

    private static void execute_Actions() throws Exception {
		for(int i=0;i<method.length;i++){			
			if(method[i].getName().equals(sActionKeyword)){
				method[i].invoke(actionKeywords,sPageObject, sData);
				//This code block will execute after every test step
				if(bResult==true){
					//If the executed test step value is true, Pass the test step in Excel sheet
					ExcelUtils.setCellData(Constants.KEYWORD_PASS, iTestStep, Constants.Col_TestStepResult, Constants.Sheet_TestSteps);
					break;
				}else{
					//If the executed test step value is false, Fail the test step in Excel sheet
					ExcelUtils.setCellData(Constants.KEYWORD_FAIL, iTestStep, Constants.Col_TestStepResult, Constants.Sheet_TestSteps);
					/*ExcelUtils.setScreenShot(Constants.screenshotName, iTestStep, Constants.Col_TestStepFailScreenShot, Constants.Sheet_TestSteps);*/
					ExcelUtils.setScreenShotLink(sTestCaseID, sActionKeyword, sPageObject, iTestStep, Constants.Col_TestStepFailScreenShot, Constants.Sheet_TestSteps);
					//In case of false, the test execution will not reach to last step of closing browser
					//So it make sense to close the browser before moving on to next test case
					ActionKeywords.closeBrowser("","");
					break;
				}
			}
		}
		}

}

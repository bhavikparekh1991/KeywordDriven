package utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.commons.io.FileUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.events.EventFiringWebDriver;

import config.ActionKeywords;
import config.Constants;
import executionEngine.DriverScript;

public class ExcelUtils {
	 private static XSSFSheet ExcelWSheet;
     private static XSSFWorkbook ExcelWBook;
     private static org.apache.poi.ss.usermodel.Cell Cell;
     private static XSSFRow Row;

     //static WebDriver driver;
     //private static XSSFCell Cell;

     	public static void setExcelFile(String Path) throws Exception {
			try {
				FileInputStream ExcelFile = new FileInputStream(Path);
				ExcelWBook = new XSSFWorkbook(ExcelFile);
			} catch (Exception e){
				Log.error("Class Utils | Method setExcelFile | Exception desc : "+e.getMessage());
				DriverScript.bResult = false;
				}
			}

		public static String getCellData(int RowNum, int ColNum, String SheetName ) throws Exception{
			try{
				ExcelWSheet = ExcelWBook.getSheet(SheetName);
				Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
				String CellData = Cell.getStringCellValue();
				return CellData;
			 }catch (Exception e){
				 Log.error("Class Utils | Method getCellData | Exception desc : "+e.getMessage());
				 DriverScript.bResult = false;
				 return"";
				 }
			 }

		public static int getRowCount(String SheetName){
			int iNumber=0;
			try {
				ExcelWSheet = ExcelWBook.getSheet(SheetName);
				iNumber=ExcelWSheet.getLastRowNum()+1;
			} catch (Exception e){
				Log.error("Class Utils | Method getRowCount | Exception desc : "+e.getMessage());
				DriverScript.bResult = false;
				}
			return iNumber;
			}

		public static int getRowContains(String sTestCaseName, int colNum,String SheetName) throws Exception{
			int iRowNum=0;	
			try {
				//ExcelWSheet = ExcelWBook.getSheet(SheetName);
				int rowCount = ExcelUtils.getRowCount(SheetName);
				for (; iRowNum<rowCount; iRowNum++){
					if  (ExcelUtils.getCellData(iRowNum,colNum,SheetName).equalsIgnoreCase(sTestCaseName)){
						break;
					}
				}       			
			} catch (Exception e){
				Log.error("Class Utils | Method getRowContains | Exception desc : "+e.getMessage());
				DriverScript.bResult = false;
				}
			return iRowNum;
			}

		public static int getTestStepsCount(String SheetName, String sTestCaseID, int iTestCaseStart) throws Exception{
			try {
				int rowcount = ExcelUtils.getRowCount(SheetName);
				for(int i=iTestCaseStart;i<=rowcount;i++){
					if(!sTestCaseID.equals(ExcelUtils.getCellData(i, Constants.Col_TestCaseID, SheetName))){
						int number = i-1;
						return number;      				
						}
					}
				ExcelWSheet = ExcelWBook.getSheet(SheetName);
				int number=ExcelWSheet.getLastRowNum()+1;
				return number;
			} catch (Exception e){
				Log.error("Class Utils | Method getRowContains | Exception desc : "+e.getMessage());
				DriverScript.bResult = false;
				return 0;
			}
		}
		
		public static void setCellData(String Result,  int RowNum, int ColNum, String SheetName) throws Exception    {
			   try{
		 
				   ExcelWSheet = ExcelWBook.getSheet(SheetName);
				   Row  = ExcelWSheet.getRow(RowNum);
				   Cell = Row.getCell(ColNum, Row.RETURN_BLANK_AS_NULL);
				   if (Cell == null) {
					   Cell = Row.createCell(ColNum);
					   Cell.setCellValue(Result);
				   } else {
						Cell.setCellValue(Result);
						}
					// Constant variables Test Data path and Test Data file name
					FileOutputStream fileOut = new FileOutputStream(Constants.Path_TestData);
					ExcelWBook.write(fileOut);
					//fileOut.flush();
					fileOut.close();
					ExcelWBook = new XSSFWorkbook(new FileInputStream(Constants.Path_TestData));
				 }catch(Exception e){
					DriverScript.bResult = false;
					}
				}

		
		public static void setScreenShotLink(String sTestCaseID, String ActionKeyword, String PageObject, int RowNum, int ColNum, String SheetName) throws Exception    {
			try
			{
				System.out.println("test");
				
				File screenshot = ((TakesScreenshot)ActionKeywords.driver).getScreenshotAs(OutputType.FILE);
				String FileName = sTestCaseID+"_"+ActionKeyword+"_"+PageObject+".png";
				FileUtils.copyFile(screenshot, new File("C:/Workspace/Workspace/KeywordDriven/ScreentShots/"+FileName));
				ExcelWSheet = ExcelWBook.getSheet(SheetName);
				Row  = ExcelWSheet.getRow(RowNum);
				Cell = Row.getCell(ColNum, Row.RETURN_BLANK_AS_NULL);
				if (Cell == null) {
					CreationHelper createHelper = ExcelWBook.getCreationHelper();
					XSSFCellStyle hlinkstyle = ExcelWBook.createCellStyle();
					XSSFFont hlinkfont = ExcelWBook.createFont();
					hlinkfont.setUnderline(XSSFFont.U_SINGLE);
					hlinkfont.setColor(HSSFColor.BLUE.index);
					hlinkstyle.setFont(hlinkfont);
					Cell = Row.createCell((short)ColNum);//.createCell((short) 1);
					
					Cell.setCellValue("C:/Workspace/Workspace/KeywordDriven/ScreentShots/"+FileName);
					
					XSSFHyperlink link = (XSSFHyperlink)createHelper.createHyperlink(Hyperlink.LINK_FILE);
					link.setAddress("C:/Workspace/Workspace/KeywordDriven/ScreentShots/"+FileName);
					Cell.setHyperlink(link);
					Cell.setCellStyle(hlinkstyle);
			       
				 }
				 else {
					 CreationHelper createHelper = ExcelWBook.getCreationHelper();
						XSSFCellStyle hlinkstyle = ExcelWBook.createCellStyle();
						XSSFFont hlinkfont = ExcelWBook.createFont();
						hlinkfont.setUnderline(XSSFFont.U_SINGLE);
						hlinkfont.setColor(HSSFColor.BLUE.index);
						hlinkstyle.setFont(hlinkfont);
						
						Cell.setCellValue("C:/Workspace/Workspace/KeywordDriven/ScreentShots/"+FileName);
						
						XSSFHyperlink link = (XSSFHyperlink)createHelper.createHyperlink(Hyperlink.LINK_FILE);
						link.setAddress("C:/Workspace/Workspace/KeywordDriven/ScreentShots/"+FileName);
						Cell.setHyperlink(link);
						Cell.setCellStyle(hlinkstyle);
				 	}
					   
				
				FileOutputStream fileOut = new FileOutputStream(Constants.Path_TestData);
				ExcelWBook.write(fileOut);
				fileOut.close();
				ExcelWBook = new XSSFWorkbook(new FileInputStream(Constants.Path_TestData));
			}
			catch(Exception e)
			{
				DriverScript.bResult = false;
			}
		}

		public static void setScreenShot(String screenshotName, int RowNum, int ColNum, String SheetName) throws Exception    {
	//ActionKeywords test = new ActionKeywords();
	//WebDriver driver = new ChromeDriver();
		   System.out.println("test");
		
		   File screenshot = ((TakesScreenshot)ActionKeywords.driver).getScreenshotAs(OutputType.FILE);
		  
		      String FullAddress=System.getProperty("user.dir")+"/ScreentShots/"+screenshotName+".png";
			//String FullAddress=System.getProperty("C:/Workspace/Workspace/KeywordDriven/ScreentShots/"+screenshotName+".png");
		      FileUtils.copyFile(screenshot, new File(FullAddress));
		      
		   ExcelWSheet = ExcelWBook.getSheet(SheetName);
		  /* Row  = ExcelWSheet.getRow(RowNum);
		   Cell = Row.getCell(ColNum, Row.RETURN_BLANK_AS_NULL);
		   if (Cell == null) {
			   Cell = Row.createCell(ColNum);*/
			  InputStream inputStream = new FileInputStream(FullAddress);
			   //Get the contents of an InputStream as a byte[].
			   byte[] bytes = IOUtils.toByteArray(inputStream);
			   //Adds a picture to the workbook
			   int pictureIdx = ExcelWBook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
			   //close the input stream
			   inputStream.close();
			   //Returns an object that handles instantiating concrete classes
			   CreationHelper helper = ExcelWBook.getCreationHelper();
			   //Creates the top-level drawing patriarch.
			   Drawing drawing = ExcelWSheet.createDrawingPatriarch();

			   //Create an anchor that is attached to the worksheet
			   ClientAnchor anchor = helper.createClientAnchor();

			   //create an anchor with upper left cell _and_ bottom right cell
			 /*  anchor.setCol1(ColNum); //Column B
			   anchor.setRow1(RowNum); //Row 3
			   anchor.setCol2(ColNum+1); //Column C
			   anchor.setRow2(RowNum+1); //Row 4
*/			   
			   anchor.setCol1(7); //Column B
			   anchor.setCol2(7+1); //Column C
			   anchor.setRow1(12); //Row 3
			   anchor.setRow2(12); //Row 4

			   //Creates a picture
			   Picture pict = drawing.createPicture(anchor, pictureIdx);

			   //Reset the image to the original size
			   //pict.resize(); //don't do that. Let the anchor resize the image!

			   //Create the Cell B3
			   Cell cell = ExcelWSheet.createRow(RowNum).createCell(ColNum);

			   //set width to n character widths = count characters * 256
			   int widthUnits = 20*256;
			   ExcelWSheet.setColumnWidth(1, widthUnits);

			   //set height to n points in twips = n * 20
			   short heightUnits = 60*20;
			   cell.getRow().setHeight(heightUnits);
		   
			// Constant variables Test Data path and Test Data file name
			FileOutputStream fileOut = new FileOutputStream(Constants.Path_TestData);
			ExcelWBook.write(fileOut);
			//fileOut.flush();
			fileOut.close();
			ExcelWBook = new XSSFWorkbook(new FileInputStream(Constants.Path_TestData));
		 
		}



}

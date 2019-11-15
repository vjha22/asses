package asses;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class Search {
	public static String vSearch;
	public static String [][] xData;
	public static int xlRows,XlCols;
	
	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		String vBrowser="gc";
		WebDriver driver = null;
		
		if (vBrowser.equals("FF"))
		{	
		System.setProperty("webdriver.gecko.driver", "C:\\testing selenium jar files\\geckodriver.exe");
		driver = new FirefoxDriver();}
		
		else if (vBrowser.contentEquals("gc"))
		{System.setProperty("webdriver.chrome.driver", "C:\\testing selenium jar files\\chromedriver.exe");
		driver = new ChromeDriver();}
		driver.manage().window().maximize();
		driver.get("https://www.yahoo.com/");
		for(int i=1;i<xlRows;i++) {
			vSearch = xData[i][0];
		driver.findElement(By.id("uh-search-box")).sendKeys(vSearch);
		Thread.sleep(1000);
		driver.findElement(By.id("uh-search-button")).click();
		Thread.sleep(1000);
		xData[i][1] = driver.getTitle();
		}}
	public static void xlRead(String sPath) throws Exception
	{
		File myFile=new File(sPath);
		FileInputStream myStream=new FileInputStream(myFile);
		HSSFWorkbook myworkbook=new HSSFWorkbook(myStream);
		HSSFSheet mySheet=myworkbook.getSheetAt(0);
		xlRows=mySheet.getLastRowNum()+1;
		XlCols=mySheet.getRow(0).getLastCellNum();
		xData=new String[xlRows][XlCols];
		for(int i=0;i<xlRows;i++)
		{
			HSSFRow row=mySheet.getRow(i);
			for(short j=0;j<XlCols;j++)
			{
				HSSFCell cell=row.getCell(j);
				String value=cellToString(cell);
				xData[i][j]=value;
				System.out.print("-"+xData[i][j]);
			}
			System.out.println();
		}
	}
		public static String cellToString(HSSFCell cell)
		{
			int type=cell.getCellType();
			Object result;
			switch(type)
			{
			case HSSFCell.CELL_TYPE_NUMERIC:
			result=cell.getNumericCellValue();
			break;
			case HSSFCell.CELL_TYPE_STRING:
			result=cell.getStringCellValue();
			break;
			case HSSFCell.CELL_TYPE_FORMULA:
			throw new RuntimeException("We cannot evaluate formula");
			case HSSFCell.CELL_TYPE_BLANK:
			result="-";
			case HSSFCell.CELL_TYPE_BOOLEAN:
			result=cell.getBooleanCellValue();
			case HSSFCell.CELL_TYPE_ERROR:
			result="This cell has some error";
			default:
			throw new RuntimeException("We do not support this cell type");
			}
			return result.toString();
			
		}
		
		public static void xlwrite(String xlpath1,String[][] xData) throws Exception
		{
			System.out.println("Inside XL Write");
			File myFile1=new File(xlpath1);
			FileOutputStream fout=new FileOutputStream(myFile1);
			HSSFWorkbook wb=new HSSFWorkbook();
			HSSFSheet mySheet1=wb.createSheet("TestResults");
			for(int i=0;i<xlRows;i++)
			{
				HSSFRow row1=mySheet1.createRow(i);
				for(short j=0;j<XlCols;j++)
				{
					HSSFCell cell1=row1.createCell(j);
					cell1.setCellType(HSSFCell.CELL_TYPE_STRING);
					cell1.setCellValue(xData[i][j]);
				}
			}
			wb.write(fout);
			fout.flush();
			fout.close();
		}

}

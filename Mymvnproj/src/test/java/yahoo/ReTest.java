package yahoo;

import java.io.FileInputStream;
import java.lang.reflect.Method;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.Listeners;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import atu.testng.reports.listeners.ATUReportsListener;
import atu.testng.reports.listeners.ConfigurationListener;
import atu.testng.reports.listeners.MethodListener;


@Listeners({ATUReportsListener.class,MethodListener.class,ConfigurationListener.class})
public class ReTest extends MainClass
{	
 
	{
		System.setProperty("atu.reporter.config", "E:\\inrhythm\\atu.properties");
	}
  @Test
  @Parameters({"browser"})
  public void retesting(String br) throws Exception
  {
	  if(br.matches("firefox"))
	  {
		driver=new FirefoxDriver();				  
	  }
	  if(br.matches("chrome"))
	  {
		 System.setProperty("webdriver.chrome.driver","E:\\inrhythm\\chromedriver.exe");
		 driver=new ChromeDriver();	  
	  }
	  
	  FileInputStream fin=new FileInputStream("E:\\inrhythm\\testdata.xlsx");
	  XSSFWorkbook wb=new XSSFWorkbook(fin);  //create workbook object
	  XSSFSheet ws=wb.getSheet("retest");
	  
	  Row row;
	  String classname,methodname;
	  
	  for(int r=1;r<=ws.getLastRowNum();r++)
	  {
		  row=ws.getRow(r);
		  if(row.getCell(4).getStringCellValue().matches("yes"))
		  {
			  classname=row.getCell(2).getStringCellValue();
			  methodname=row.getCell(3).getStringCellValue();
			  
			  Class c=Class.forName(classname);   
			  Method m=c.getMethod(methodname,null);   
			  Object obj=c.newInstance();   
			  m.invoke(obj, null); 
			   
		  }
	  }
	  fin.close();
		
	 /* Home h=new Home(driver);
	  h.createacc();
	  h.login();
	  
	  Inbox i=new Inbox(driver);
	  i.deletemail();
	  
	  Compose c=new Compose(driver);
	  c.signout();*/
   }
  }


	 












package Sample;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;

public class ExcelReadingusingFillo {
  @Test
  public void f() throws FilloException {
	   
	  System.setProperty("webdriver.chrome.driver",".\\driver\\chromedriver.exe");
	  WebDriver driver=new ChromeDriver();
	  driver.navigate().to("https://login.microsoftonline.com/");
	  driver.manage().window().maximize();
	  
	  String excelPath = ".\\testdata\\Test.xlsx";
	  System.out.println(excelPath);
	  
	  Fillo fillo = new Fillo();
	  Connection connection = fillo.getConnection(excelPath);
	  
	  String strSelectQuerry = "Select * from  Sheet1";
	  System.out.println(strSelectQuerry);
	  
	  Recordset recordset =null;
	  recordset = connection.executeQuery(strSelectQuerry);
	  
	  while(recordset.next())
	  {
	         
	     System.out.println("Column 1 = " +recordset.getField("SiteTitle"));
	     String siteTitle = recordset.getField("SiteTitle");
	 
	             
	     System.out.println("Column 2 = " +recordset.getField("SiteTopic"));
	     String siteTopic = recordset.getField("SiteTopic");
	           
	             
	   }
	  
	  connection.close();
	  
	  //Use update query to update the details in excel  file
	    Connection connection1 = fillo.getConnection(excelPath);
	     
	    System.out.println("Column 1 value before Update clause = " +recordset.getField("SiteTitle"));      
	    String strUpdateQuerry = "Update Sheet1 Set SiteTitle = 'SoftwareTestingHelp.com' ";
	     
	    System.out.println(strUpdateQuerry);
	    connection1.executeUpdate(strUpdateQuerry);
	     
	    System.out.println("Column 1 value after Update clause = " +recordset.getField("SiteTitle"));
	  
	  

	  
  }
}


package reader;


//http://www.vogella.com/tutorials/JavaExcel/article.html

import java.io.File;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;


import java.io.IOException;
import java.util.Locale;

import java.io.File;
import java.io.IOException;
import jxl.*;
import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.read.biff.BiffException;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import com.gargoylesoftware.htmlunit.javascript.host.Console;



public class login {
	
	public static String inputFile;
	public static String npa;
	public static String prfx;
	public static String lnr;
	public static String userId;
	public static String pwd;
	public static String siteIp;
	
	private static WritableCellFormat timesBoldUnderline;
	  private static WritableCellFormat times;
	  
	
	public void setInputFile(String inputFile) {
	    this.inputFile = inputFile;
	  }
	private static void createLabel(WritableSheet sheet)
		      throws WriteException {
		    // Lets create a times font
		    WritableFont times10pt = new WritableFont(WritableFont.TIMES, 10);
		    // Define the cell format
		    times = new WritableCellFormat(times10pt);
		    // Lets automatically wrap the cells
		    times.setWrap(true);

		    // create create a bold font with unterlines
		    WritableFont times10ptBoldUnderline = new WritableFont(WritableFont.TIMES, 10, WritableFont.BOLD, false,
		        UnderlineStyle.SINGLE);
		    timesBoldUnderline = new WritableCellFormat(times10ptBoldUnderline);
		    // Lets automatically wrap the cells
		    timesBoldUnderline.setWrap(true);

		    CellView cv = new CellView();
		    cv.setFormat(times);
		    cv.setFormat(timesBoldUnderline);
		    cv.setAutosize(true);

		    // Write a few headers
		    addCaption(sheet, 0, 0, "TELEPHONE NUM");
		    addCaption(sheet, 1, 0, "LSMS LRN");
		    addCaption(sheet, 2, 0, "SPID");
		    addCaption(sheet, 3, 0, "LSMS DATE");
		    addCaption(sheet, 4, 0, "ACTION");

		  }
	private static void addCaption(WritableSheet sheet, int column, int row, String s)
		      throws RowsExceededException, WriteException {
		    Label label;
		    label = new Label(column, row, s, timesBoldUnderline);
		    sheet.addCell(label);
		  }


	
	/**
	 * @param args
	 * @throws IOException 
	 * @throws BiffException 
	 */
	
	
	public static void main(String[] args) throws InterruptedException, BiffException, IOException {
		// TODO Auto-generated method stub
		
		

		System.out.println("LNPUPDATE tool start");
		
		
		
		File file = new File("./chromedriver.exe");
        System.setProperty("webdriver.chrome.driver", file.getAbsolutePath());
     
        
      
        WebDriver driver = new ChromeDriver();
        
        
        File userIP = new File("C:/lnp/userprofip.xls");
        System.out.println("yahan??");
        Workbook w1;
        
        w1 = Workbook.getWorkbook(userIP);
        System.out.println("ya yahan ??");
		Sheet sheet1 = w1.getSheet(0);
		System.out.println("sheet1.getColumns() "+sheet1.getColumns());
		System.out.println("sheet1.getRows() "+sheet1.getRows());
		//for (int y = 1; y < sheet1.getRows(); y++) {
			//System.out.println("y "+y);
			for (int x = 0; x < sheet1.getColumns(); x++) {
				System.out.println("x "+x);
				Cell cell1 = sheet1.getCell(x, 1);
      			//CellType type1 = cell1.getType();
      			
      			System.out.println("cell1 contents: "+cell1.getContents().toString());
		
      			if(x==0)
  				{
  					userId = cell1.getContents().toString();
  					System.out.println("userId "+userId);
  				}
  				if(x==1)
  				{
  					pwd = cell1.getContents().toString();
  					System.out.println("pwd "+pwd);
  				}
  				if(x==2)
  				{
  					siteIp = cell1.getContents().toString();
  					System.out.println("siteIp "+siteIp);
  				}
			}
		//}
        //driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

        //driver.get("http://clph105.sldc.sbc.com/lnp/");
        //driver.get("http://intranet.att.com/lnptool/");
        driver.get(siteIp);
WebElement userid = driver.findElement(By.name("userid"));
userid.sendKeys(userId);
WebElement password = driver.findElement(By.name("password"));
password.sendKeys(pwd);
WebElement submit = driver.findElement(By.name("btnSubmit"));
submit.click();
WebElement ok = driver.findElement(By.name("successOK"));
ok.click();
Thread.sleep(3000);
WebElement login = driver.findElement(By.xpath("html/body/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td/center/a/img"));
login.click();
Thread.sleep(3000);

//WebElement westcoast = driver.findElement(By.name("region"));
WebElement westcoast = driver.findElement(By.xpath("html/body/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td/center/table/tbody/tr/td[1]/table[@class='subtitle']/tbody/tr/td/form/input[2]"));
westcoast.click();


File outFile = new File("./lnpoutput.xls");
WorkbookSettings wbSettings = new WorkbookSettings();

wbSettings.setLocale(new Locale("en", "EN"));

File inputWorkbook = new File("./lnpinput.xls");

Workbook w;

System.out.println("main yahan hoon yahan hoon yahan hoon yahan");
try {
	
	try
	{
		System.out.println("yahan aaya kya??");
		
		WritableWorkbook workbook = Workbook.createWorkbook(outFile, wbSettings);
		workbook.createSheet("Report", 0);
		WritableSheet excelSheet = workbook.getSheet(0);
		createLabel(excelSheet);
		
		
				
		w = Workbook.getWorkbook(inputWorkbook);
		//Get the first sheet
		Sheet sheet = w.getSheet(0);
		System.out.println("sheet.getColumns() "+sheet.getColumns());
		System.out.println("sheet.getRows() "+sheet.getRows());
		//Loop over first 10 column and lines
		System.out.println("Workbook to khulgayi lagta hai");
			for (int j = 1; j < sheet.getRows(); j++) {
				System.out.println("j "+j);
				WebElement lsms = driver.findElement(By.xpath("/html/body/table/tbody/tr[2]/td/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td[2]/table/tbody/tr[19]/td/div[@class='menuspace']/a[@class='leftnav']/font"));
				//WebElement lsms = driver.findElement(By.xpath("//a[@class='leftnav' and @href='SV.cfm?db=1']"));
				lsms.click();
				
				System.out.println("lsms click ni ho ra kya ??");
				
					for (int i = 1; i < sheet.getColumns(); i++) {
						System.out.println("i "+i);
					Cell cell = sheet.getCell(i, j);
	      			CellType type = cell.getType();
	      			
	      			System.out.println("cell contents: "+cell.getContents().toString());
	      			
	      			System.out.println("yahan hai ??");
	      			
	      			if (type == CellType.LABEL) {
	      			System.out.println("Abe number chahiye text nahi !!  "
	    	  			+ cell.getContents());
	      			}

	      			if (type == CellType.NUMBER) {
	      				System.out.println("I got a number "
	      						+ cell.getContents());
	      				
	      			}
      		
	      			if(i==1)
      				{
      					npa = cell.getContents().toString();
      				}
      				if(i==2)
      				{
      					prfx = cell.getContents().toString();
      				}
      				if(i==3)
      				{
      					lnr = cell.getContents().toString();
      				}
				}
      		System.out.println("npa: "+npa);
      		System.out.println("prfx: "+prfx);
      		System.out.println("lnr: "+lnr);
      		
      		String strFirst=npa;
      		String strmiddle=prfx;
      		String strLast=lnr;
      		String TN=strFirst+strmiddle+strLast;
      		WebElement first = driver.findElement(By.name("npa"));
      		first.sendKeys(strFirst);
      		WebElement middle = driver.findElement(By.name("nxx"));
      		middle.sendKeys(strmiddle);
      		WebElement last = driver.findElement(By.name("loline"));
      		last.sendKeys(strLast);
      		WebElement status = driver.findElement(By.name("status"));
      		Select select = new Select(status);
      		select.selectByVisibleText("Active");
      		WebElement sub = driver.findElement(By.xpath("/html/body/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[1]/table/tbody/tr/td/center/form/table/tbody/tr[3]/td/table/tbody/tr/td[1]/input"));
      		//WebElement sub = driver.findElement(By.xpath("//input[@src='/lnp/images/submit.gif']"));
      		sub.click();

      		//Find Telephone Number
      		WebElement telNoMatch=driver.findElement(By.xpath("//a[text()='"+TN+"']"));
      		String telNo = telNoMatch.getText();
      		System.out.println("TN "+telNo);

      		//Find LSMS LRN of TN
      		WebElement lrnMatch=driver.findElement(By.xpath("//table[@class='hometext']//td[contains(.,'"+TN+"')]/following-sibling::td[4]"));
      		String lrn = lrnMatch.getText();
      		System.out.println("LSMS LRN "+lrn);

      		//Find SPID of TN
      		WebElement spidMatch=driver.findElement(By.xpath("//table[@class='hometext']//td[contains(.,'"+TN+"')]/following-sibling::td[6]"));
      		String spid=spidMatch.getText();
      		//Extract only the Numbers from New SP field
      		String finalSpid=spid.substring(0, spid.indexOf('(')).trim();
      		System.out.println("SPID "+finalSpid);

      		//Find LSMS Date of TN
      		WebElement dateMatch=driver.findElement(By.xpath("//table[@class='hometext']//td[contains(.,'"+TN+"')]/following-sibling::td[9]"));
      		String date=dateMatch.getText();
      		String finalDate=date.substring(0, date.indexOf(' ')).trim();
      		System.out.println("LSMS Date "+finalDate);
      		
      		
      		for(int i = 0; i < 5; i++) {
      		Label label;
      		if(i==0)
      		{
      			label = new Label(i,j,telNo,times);
      		excelSheet.addCell(label);
      		}
      		if(i==1)
      		{
      			label = new Label(i,j,lrn,times);
      		excelSheet.addCell(label);
      		}
      		if(i==2)
      		{
      			label = new Label(i,j,finalSpid,times);
      		excelSheet.addCell(label);
      		}
      		if(i==3)
      		{
      			label = new Label(i,j,finalDate,times);
      		excelSheet.addCell(label);
      		}
      		if(i==4)
      		{
      			label = new Label(i,j,"UPDATE",times);
      		excelSheet.addCell(label);
      		}
      		}
	    	
		}   
			
			workbook.write();
			workbook.close();
	}
	
	catch(Exception e)
	{
		System.out.println("Workbook yar");
	}
  
}
 
 catch (Exception e) {
  e.printStackTrace();
  System.out.println("confusing try catch inception :-P");
}


driver.quit();

System.out.println("Please check the result file under ./lnpoutput.xls ");
	}

}
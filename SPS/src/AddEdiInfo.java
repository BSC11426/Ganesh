import java.io.File;
import java.util.Properties;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.sikuli.script.Pattern;
import org.sikuli.script.Screen;

import jxl.Sheet;
import jxl.Workbook;

public class AddEdiInfo {
	public static void main(String[] args) throws Exception {

		Properties CONFIG = null;

		Workbook workbook = Workbook.getWorkbook(new File("D:/Ganesh/AddEDIInfo1.xls"));
		Sheet sheet1 = workbook.getSheet(0);

//		File propfile = new File("C:/Users/ganesh.pirale/workspace/SPSCommerce/src/Data.properties");

//		CONFIG = new Properties();
//		CONFIG.load(new FileInputStream(propfile));
//
//		String str = CONFIG.getProperty("name");
//		System.out.println(str);
		
		WebDriver driver = new FirefoxDriver();
		driver.manage().window().maximize();

//		driver.get("http://boon.mps.spscommerce.net:7777/sso/jsp/login.jsp?site2pstoretoken=v1.4~E1CED8E6~4FCEE530A480A5EB284A5BCE7884E2351F6617B53021BACC523F359BA841C59F4D5D351B93228EA0BA1ED62A6C6911739276042AD4A822880CEBE9EE2CA43D0B389940B006717CB24B59B7EB08FFCB4E6009B9E6A4340E7DAAC67BAB5B4AC0C726F6CE3D31CA6BCD0ECB0F75849E9BDF94F7B768D3025EED1167E67E71E6353EADD97EF166144FBFBCA4AAF2ACC6536550A9D3F5D998547EEBED6AC81071A7A2DA7F6D1B8E92B3BA30A53340C920E2597063AA84C1991B6A0E34534327B6F3E51B19188E51B0FDBD8ECC5863EAF5ED27&p_error_code=&p_submit_url=http%3A%2F%2Fboon.mps.spscommerce.net%3A7777%2Fsso%2Fauth&p_cancel_url=http%3A%2F%2Fdc4ui.p01.ppd%3A7777&ssousername=ypawar");
		
		driver.get(sheet1.getCell(0, 0).getContents());
		driver.findElement(By.xpath("//input[@id='c3']")).sendKeys(sheet1.getCell(0, 1).getContents()); 
		driver.findElement(By.cssSelector("input[type='submit']:nth-child(1)")).click();   // Login 
		
			
		Thread.sleep(2000);

		Screen screen = new Screen();

		// Pattern image = new
		// Pattern("C://Users//ganesh.pirale//Desktop//sikulisnap//gmail.PNG");
		// Pattern image = new
		// Pattern("C://Users/ganesh.pirale/Desktop/sikulisnap/gmail.PNG"); //
		// it works
		// Pattern image = new
		// Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\gmail.PNG");
		// // it works
		// screen.click(image);
		// Pattern image1 = new
		// Pattern("C://Users/ganesh.pirale/Desktop/sikulisnap/gmail1.PNG");
		// screen.click(image1);

		Pattern pwd = new Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\pwd.PNG");
		Pattern login = new Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\login.PNG");
		// Pattern dc4 = new Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\dc4.PNG");
		Pattern EDIDesc = new Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\EDIDesc.PNG");
		Pattern EnvQlf = new Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\EnvQlf.PNG");
		Pattern EnvID = new Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\EnvID.PNG");
		Pattern GrpID = new Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\GrpID.PNG");
		Pattern GrpQlf = new Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\GrpQlf.PNG");

		Pattern CommID = new Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\CommID.PNG");
		Pattern EDIPwd = new Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\EDIPwd.PNG");
		Pattern SegDelimeter = new Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\SegDelimeter.PNG");
		Pattern ElementDelimiter = new Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\ElementDelimiter.PNG");
		Pattern SubeleDelimiter = new Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\SubeleDelimiter.PNG");
		Pattern RelChar = new Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\RelChar.PNG");
		Pattern DecChar = new Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\DecChar.PNG");
//		Pattern Cancel = new Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\Cancel.PNG");
		Pattern save = new Pattern("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap\\save.PNG");
		
//		screen.type(pwd, "canvas5233");
		

		Sheet sheet2 = workbook.getSheet(1);
		// driver.findElement(By.xpath("//a[@id='menuTabs1:0:commandMenuItem1'][@name='menuTabs1:0:commandMenuItem1']")).click();
		driver.findElement(By.linkText("DC4")).click();
		Thread.sleep(1000);
		driver.findElement(By.xpath("//input[@id='inputText41']")).sendKeys(sheet2.getCell(0, 0).getContents());  // Type name very carefully here that should be test data not real company name  
		driver.findElement(By.xpath("//button[@id='commandButton12']")).click();
		driver.findElement(By.linkText(sheet2.getCell(0, 1).getContents())).click();  // click very carefully here that should be test data not real company name 
		driver.findElement(By.linkText("EDI Info")).click();
		// driver.findElement(By.linkText("Add EDI Info")).click();
		// driver.findElement(By.xpath("//*[@id='_idJsp4:_idJsp19']/table[1]/tbody/tr[1]/td/table/tbody/tr/td[1]/button")).click();

		Thread.sleep(4000);
		// driver.switchTo().frame("generator");
		// driver.switchTo().defaultContent();

//		Robot robot = new Robot();
		// robot.keyPress(java.awt.event.KeyEvent.VK_TAB);
		// robot.keyRelease(java.awt.event.KeyEvent.VK_TAB);
	 

		System.out.println("Switched to frame");

		// try {
		// wait.until(ExpectedConditions.class (by));
		// } catch (TimeoutException e) {
		// e.printStackTrace();
		// }
 
		// screen.type(DecChar, "11");

		int r = sheet2.getRows();
		int c = sheet2.getColumns();

		System.out.println("Rows:" + r);
		System.out.println("Columns:" + c);

		
		for (int i = 3; i < r; i++) {   // i is row number, starts with 0 i.e index starts with 0 
			
			driver.findElement(By.xpath("//*[@id='_idJsp4:_idJsp19']/table[1]/tbody/tr[1]/td/table/tbody/tr/td[1]/button")).click();  // Add EDI Info

			screen.type(EDIDesc, sheet2.getCell(0, i).getContents());
			screen.type(EnvQlf, sheet2.getCell(1, i).getContents());
			screen.type(EnvID, sheet2.getCell(2, i).getContents());
			screen.type(GrpID, sheet2.getCell(3, i).getContents());
			screen.type(GrpQlf, sheet2.getCell(4, i).getContents());
			screen.type(CommID, sheet2.getCell(5, i).getContents());
			screen.type(EDIPwd, sheet2.getCell(6, i).getContents());
			screen.type(SegDelimeter, sheet2.getCell(7, i).getContents());
			screen.type(ElementDelimiter, sheet2.getCell(8, i).getContents());
			screen.type(SubeleDelimiter, sheet2.getCell(9, i).getContents());
			screen.type(RelChar, sheet2.getCell(10, i).getContents());
			screen.type(DecChar, sheet2.getCell(11, i).getContents());

			screen.click(RelChar);

			screen.click(save);

			Thread.sleep(6000);

		}

		// screen.type(EDIDesc, sheet1.getCell(0, 0).getContents());

		// screen.type("C:\\Users\\ganesh.pirale\\Desktop\\sikulisnap1\\EDI.xlsx",
		// "Code");

	}
}

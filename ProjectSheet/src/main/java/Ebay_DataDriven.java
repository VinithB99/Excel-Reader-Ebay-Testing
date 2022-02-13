import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.time.Duration;
import java.util.List;
import java.util.Properties;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;

public class Ebay_DataDriven {

	public static WebDriver driver;
	public static String Browser = "Chrome";
	public static int iRow, iTotalRow, iCell, iTotalCell;
	public static String sExcelFile = "./data/New.xlsx";
	static Properties oPro = new Properties();

	public static void main(String[] args) throws IOException {
		String sSheet, sSearchTxt;
		sSheet = "Sheet1";
		Page_Info();
		try {
			InputStream oFile = new FileInputStream(sExcelFile);
			XSSFWorkbook oExcel = new XSSFWorkbook(oFile);
			XSSFSheet oSheet = oExcel.getSheet(sSheet);
			Row oRow;
			Cell oCell;
			iTotalRow = oSheet.getLastRowNum();
			for (iRow = 1; iRow <= iTotalRow; iRow++) {
				oRow = oSheet.getRow(iRow);
				iTotalCell = oRow.getLastCellNum();
				for (iCell = 0; iCell < iTotalCell; iCell++) {
					oCell = oRow.getCell(iCell);
					sSearchTxt = oCell.getStringCellValue();
					Search_Product(sSearchTxt, "Electronics");
					Get_Match(sSearchTxt);
				}
			}
			oExcel.close();
			oFile.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static void readPropertyFile() throws IOException {
		String location = "D:\\JAVA\\eclipse\\ProjectSheet\\data\\Environment.properties";
		FileInputStream oFile = new FileInputStream(location);
		oPro.load(oFile);
		System.out.println("Sit URL is : " + oPro.getProperty("sit_url"));
		System.out.println("Sit name is : " + oPro.getProperty("sit_name"));
		oFile.close();
	}

	public static void invokeBrowser() {
		switch (Browser) {
		case "Chrome":
			System.out.println("User option is " + oPro.getProperty("browser1") + ",So invoking Chrome browser!!!");
			System.setProperty("webdriver.chrome.driver", "./driver/chromedriver.exe");
			driver = new ChromeDriver();
			break;
		case "Firefox":
			System.out.println("User option is " + oPro.getProperty("browser2") + ",So invoking Firefox browser!!!");
			System.setProperty("webdriver.gecko.driver", "./driver/geckodriver.exe");
			driver = new FirefoxDriver();
			break;
		case "Edge":
			System.out.println("User option is " + oPro.getProperty("browser2") + ",So invoking Edge browser!!!");
			System.setProperty("webdriver.edge.driver", "./driver/msedgedriver.exe");
			driver = new EdgeDriver();
			break;

		default:
			System.out.println("User option is wrong " + oPro.getProperty("browser1") + ",So invoking default Chrome browser!!!");
			System.setProperty("webdriver.chrome.driver", "./driver/chromedriver.exe");
			driver = new ChromeDriver();
			break;
		}
	}

	public static void Page_Info() throws IOException {
		readPropertyFile();
		invokeBrowser();
		driver.manage().window().maximize();
		driver.get(oPro.getProperty("sit_url"));
		driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(20));
	}

	public static void Search_Product(String sTxt, String sCat) {
		WebElement oText, oBtn, oDropDown;
		oText = driver.findElement(By.xpath("//input[@id='twotabsearchtextbox']"));
		oText.clear();
		oText.sendKeys(sTxt);
		oDropDown = driver.findElement(By.xpath("//select[@id='searchDropdownBox']"));
		Select oSelect = new Select(oDropDown);
		oSelect.selectByVisibleText(sCat);
		oBtn = driver.findElement(By.id("nav-search-submit-button"));
		oBtn.click();
		driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(30));
	}

	public static void Get_Match(String sTxt) throws Exception {
		WebElement oText, oProduct;
		Thread.sleep(5000);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
		oText = driver.findElement(By.xpath("//div[@class='a-section a-spacing-small a-spacing-top-small']/span"));
		String sText = oText.getText();
		System.out.println("Search Result is : " + sText);
		sText = sText.replaceAll("[^0-9]", "").trim();
		int iText = Integer.parseInt(sText);
		if (iText > 0) {
			System.out.println("Search Result is Listed");
		} else {
			System.out.println("No Search Result");
		}
		List<WebElement> oList = driver.findElements(By.xpath("//div[@class='a-section a-spacing-medium']"));
		System.out.println("Total Value is : " + oList.size());
		for (int i = 0; i < oList.size(); i++) {
			oProduct = oList.get(i);
			String sVal1 = oProduct.findElement(By.xpath(".//a[@class='a-link-normal a-text-normal']/span")).getText();
			String sVal2 = oList.get(i).findElement(By.xpath(".//span[@class='a-price']//span[@class='a-price-whole']"))
					.getText();
			System.out.println(sVal1);
			System.out.println(sVal2);
			Write_Cell_Value_To_Excel(sExcelFile, sTxt, i, 0, sVal1);
			Write_Cell_Value_To_Excel(sExcelFile, sTxt, i, 1, sVal2);
		}
		ScrollPageto(0, 0);
	}

	public static void ScrollPageto(int x, int y) {
		JavascriptExecutor oJs;
		String sCmd;
		oJs = (JavascriptExecutor) driver;
		sCmd = String.format("window.scrollTo(%d,%d)", x, y);
		oJs.executeScript(sCmd);
	}

	public static void Write_Cell_Value_To_Excel(String sFile, String sSheet, int iRow, int iCell, String sValue) {
		InputStream oFile;
		XSSFWorkbook oExcel;
		XSSFSheet oSheet;
		Row oRow;
		Cell oCell;
		try {
			oFile = new FileInputStream(sFile);
			oExcel = new XSSFWorkbook(oFile);
			oSheet = oExcel.getSheet(sSheet);
			if (oSheet == null) {
				oExcel.createSheet(sSheet);
				oSheet = oExcel.getSheet(sSheet);
			}
			oRow = oSheet.getRow(iRow);
			if (oRow == null) {
				oSheet.createRow(iRow);
				oRow = oSheet.getRow(iRow);
			}
			oCell = oRow.getCell(iCell);
			if (oCell == null) {
				oRow.createCell(iCell);
				oCell = oRow.getCell(iCell);
			}
			oCell.setCellValue(sValue);
			OutputStream oFileWrite = new FileOutputStream(sFile);
			oExcel.write(oFileWrite);
			oFileWrite.close();
			oExcel.close();
			oFile.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}

package headingTags;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.time.Duration;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.opencsv.CSVWriter;

public class EmpulsWebsite {
	WebDriver driver;
	
	@BeforeTest
	public void launchBrowser()
	{
		System.setProperty("webdriver.chrome.driver", "/Users/kailash.k/Downloads/chromedriver_mac64 (4)/chromedriver");
		driver = new ChromeDriver();
		Date d = new Date();
		System.out.println("Test Execution Date : " + d.toString());
	}
	
	@Test
	public void countOfH1HomePage() {
		try {
			
				driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
				//String url = "https://www.empuls.io/";
				//driver.get(url);
				//System.out.println("Page link : " + url);
			
				FileInputStream fs = new FileInputStream("/Users/kailash.k/Documents/EmpulsWebsitePage links.xlsx");
				//Creating a workbook
				XSSFWorkbook wb = new XSSFWorkbook(fs);
				XSSFSheet sheet = wb.getSheetAt(0);
				//List<WebElement> pageUrl;
				int totRows = sheet.getLastRowNum();
				System.out.println("total rows : " + totRows);
				
			
				
				FileWriter outputfile = new FileWriter("/Users/kailash.k/Documents/CountOfHeadingTages.csv");
				CSVWriter writer = new CSVWriter(outputfile);
				
				String[] header = {"pageLink","Title Count","H1 Tag","H2 Tag","H3 Tag","H4 Tag","H5 Tag","H6 Tag"};
				writer.writeNext(header);
				int firstRow = sheet.getFirstRowNum();
				System.out.println("First Row Number : " + firstRow);
				System.out.println("Code running before loop");
				for(int i=firstRow;i<=totRows;i++) {
					Row row = sheet.getRow(i);
					Cell cell = row.getCell(0);
					String pageLink = cell.getStringCellValue();
					System.out.println("Page link : " + pageLink);
					driver.get(pageLink);
			
					List<WebElement> heading1 = driver.findElements(By.tagName("h1"));
					List<WebElement> heading2 = driver.findElements(By.tagName("h2"));
					List<WebElement> heading3 = driver.findElements(By.tagName("h3"));
					List<WebElement> heading4 = driver.findElements(By.tagName("h4"));
					List<WebElement> heading5 = driver.findElements(By.tagName("h5"));
					List<WebElement> heading6 = driver.findElements(By.tagName("h6"));
					List<WebElement> titleCount = driver.findElements(By.tagName("title"));
					int H1Count = heading1.size();
					int H2Count = heading2.size();
					int H3Count = heading3.size();
					int H4Count = heading4.size();
					int H5Count = heading5.size();
					int H6Count = heading6.size();
					int title = titleCount.size(); 
			
					String heading1Count = Integer.toString(H1Count);
					String heading2Count = Integer.toString(H2Count);
					String heading3Count = Integer.toString(H3Count);
					String heading4Count = Integer.toString(H4Count);
					String heading5Count = Integer.toString(H5Count);
					String heading6Count = Integer.toString(H6Count);
					String titleCnt = Integer.toString(title);
					
					
					String[] data1 = {pageLink,titleCnt,heading1Count,heading2Count,heading3Count,heading4Count,heading5Count,heading6Count};
					writer.writeNext(data1);
					
					
					
		
			
				}
				
				
					wb.close();
					writer.close(); 
				
		}
		catch(Exception e) {
			System.out.println(e);
		}
	
	}
	
	@AfterTest
	public void closeBrowser() {
	
		driver.close();
	}

}

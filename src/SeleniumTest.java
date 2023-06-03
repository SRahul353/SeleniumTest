import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import java.io.File;
import java.io.FileInputStream;
import java.util.List;
import java.util.ArrayList;
import java.util.Collections;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SeleniumTest {

    public static void main(String[] args) {

    	System.setProperty("webdriver.chrome.driver", "C:\\Users\\sahar\\eclipse-workspace\\Selenium\\Selenium WebDriver\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.navigate().to("http://www.google.com/");

        
        try {
            File file = new File("C:\\Users\\sahar\\eclipse-workspace\\Selenium\\Selenium WebDriver\\Excel.xlsx");
            FileInputStream fis = new FileInputStream(file);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            for (Sheet sheet : wb) {
            	
                int row = 2, col = 2, loop=1;
                Cell testdataCell = sheet.getRow(row).getCell(col);
                String testdata = testdataCell.toString();

                while (loop==1) {
                    WebElement searchBar = driver.findElement(By.name("q"));
                    searchBar.clear();
                    searchBar.sendKeys(testdata);

                    Thread.sleep(3000);

                    List<WebElement> suggestionElements = driver.findElements(By.xpath("//div[@class='wM6W7d' and @role='presentation']/span"));
                    List<String> suggestionList = new ArrayList<>();
                    for (WebElement element : suggestionElements) {
                        String text = element.getText().trim();
                        if (!text.isEmpty()) {
                            suggestionList.add(text);
                        }
                    }
                    String maxop = Collections.max(suggestionList);
                    String minop = Collections.min(suggestionList);
                    
                    
                    Cell c = sheet.getRow(row).getCell(col+1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    c.setCellValue(maxop);
                    c = sheet.getRow(row).getCell(col+2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    c.setCellValue(minop);


                    row++;
                    try {
                    Cell testdatanew = sheet.getRow(row).getCell(col);
                    testdata = testdatanew.toString();
                    }
                    catch(Exception e) {
                    	loop=0;
                    }
                }

            }
        fis.close();
        FileOutputStream fos = new FileOutputStream(file);
        wb.write(fos);
        fos.close();
        wb.close();

        } 
        catch (Exception e) {
            e.printStackTrace();
        }

       driver.quit();
    }
}

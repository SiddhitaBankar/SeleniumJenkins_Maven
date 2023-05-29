import static org.testng.Assert.assertEquals;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

public class ReadingExcelSelenium_Ex1 {
WebDriver driver;
String projectPath = System.getProperty("user.dir");
String excelPath = projectPath + "\\data\\seltestdata1.xlsx";
String chromeDriverPath = System.setProperty("webdriver.chrome.driver",projectPath + "\\drivers\\chromedriver.exe");
String url = "https://parabank.parasoft.com/parabank/index.htm";

@Test
public void registration() throws InterruptedException, IOException {
FileInputStream fis = new FileInputStream(excelPath);

XSSFWorkbook wb = new XSSFWorkbook(fis);
XSSFSheet sheet = wb.getSheet("Sheet1");

int rows = sheet.getLastRowNum() - sheet.getFirstRowNum();
System.out.println("Row count = " + rows);
driver = new ChromeDriver();
driver.get(url);
Thread.sleep(2000);
DataFormatter formatter = new DataFormatter();
for (int i = 1; i <= rows; i++) {

WebElement registerLink = driver.findElement(By.linkText("Register"));
registerLink.click();

WebElement firstNameElement = driver.findElement(By.id("customer.firstName"));
String firstName = sheet.getRow(i).getCell(0).toString();
firstNameElement.sendKeys(firstName);
WebElement lastNameElement = driver.findElement(By.id("customer.lastName"));

String lastName = sheet.getRow(i).getCell(1).toString();
lastNameElement.sendKeys(lastName);

WebElement addressElement = driver.findElement(By.id("customer.address.street"));
String address = sheet.getRow(i).getCell(2).toString();
addressElement.sendKeys(address);

WebElement cityElement = driver.findElement(By.id("customer.address.city"));
String city = sheet.getRow(i).getCell(3).toString();
cityElement.sendKeys(city);
WebElement stateElement = driver.findElement(By.id("customer.address.state"));
String state = sheet.getRow(i).getCell(4).toString();
stateElement.sendKeys(state);

WebElement zipCodeElement = driver.findElement(By.id("customer.address.zipCode"));
String zipCode = sheet.getRow(i).getCell(5).toString();
zipCodeElement.sendKeys(zipCode);

WebElement phoneElement = driver.findElement(By.id("customer.phoneNumber"));

XSSFCell cell = sheet.getRow(i).getCell(6);
String phone = formatter.formatCellValue(cell);
phoneElement.sendKeys(phone);
WebElement ssnElement = driver.findElement(By.id("customer.ssn"));
String ssn = sheet.getRow(i).getCell(7).toString();
ssnElement.sendKeys(ssn);

WebElement userNameElement = driver.findElement(By.id("customer.username"));
String userName = sheet.getRow(i).getCell(8).toString();
userNameElement.sendKeys(userName);
WebElement passwordElement = driver.findElement(By.id("customer.password"));
String password = sheet.getRow(i).getCell(9).toString();
passwordElement.sendKeys(password);
WebElement confirmPasswordElement = driver.findElement(By.id("repeatedPassword"));

String confirmPassword = sheet.getRow(i).getCell(10).toString();
confirmPasswordElement.sendKeys(confirmPassword);

WebElement register = driver.findElement(By.xpath("//input[@value='Register']"));
register.click();
assertEquals(driver.getTitle(), "ParaBank | Customer Created");
System.out.println(driver.getTitle());

Thread.sleep(5000);
WebElement logout = driver.findElement(By.linkText("Log Out"));
logout.click();

}

wb.close();

Thread.sleep(2000);
driver.quit();
}
}

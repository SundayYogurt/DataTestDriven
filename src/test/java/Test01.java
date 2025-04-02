import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.*;
import org.openqa.selenium.*;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.concurrent.TimeUnit;

public class Test01 {

    private WebDriver driver;
    private WebDriverWait wait;

    @BeforeAll
    static void setUpClass() {
        System.setProperty("webdriver.edge.driver", "./edgedriver_win64/msedgedriver2.exe");
    }

    @BeforeEach
    void setUp() {
        driver = new EdgeDriver();
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        wait = new WebDriverWait(driver, 10); // Correct usage with long (seconds)

    }

    @Test
    void test01() throws IOException {
        String path = "./exel/Sci.xlsx";
        try (FileInputStream fs = new FileInputStream(path);
             XSSFWorkbook workbook = new XSSFWorkbook(fs)) {

            XSSFSheet sheet = workbook.getSheetAt(0);
            int rowNum = sheet.getLastRowNum() + 1;

            for (int i = 1; i < rowNum ; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                driver.get("http://localhost/sc_shortcourses/signup");

                // Get values from the Excel sheet
                String nameTitleTha = getCellValue(row.getCell(1));
                String firstnameTha = getCellValue(row.getCell(2));
                String lastnameTha = getCellValue(row.getCell(3));
                String nameTitleEng = getCellValue(row.getCell(4));
                String firstnameEng = getCellValue(row.getCell(5));
                String lastnameEng = getCellValue(row.getCell(6));
                String birthDate = getCellValue(row.getCell(7));
                String birthMonth = getCellValue(row.getCell(8));
                String birthYear = getCellValue(row.getCell(9));
                String idCard = getCellValue(row.getCell(10));
                String password = getCellValue(row.getCell(11));
                String mobile = getCellValue(row.getCell(12));
                String email = getCellValue(row.getCell(13));
                String address = getCellValue(row.getCell(14));
                String province = getCellValue(row.getCell(15));
                String district = getCellValue(row.getCell(16));
                String subDistrict = getCellValue(row.getCell(17));
                String postalCode = getCellValue(row.getCell(18));
                String accept = getCellValue(row.getCell(19));

                fillForm(nameTitleTha, firstnameTha, lastnameTha, nameTitleEng, firstnameEng, lastnameEng, birthDate, birthMonth, birthYear, idCard, password, mobile, email, address, province, district, subDistrict, postalCode);

                // Handle checkbox explicitly via WebDriverWait and JavaScriptExecutor
                WebElement acceptCheckbox = driver.findElement(By.id("accept"));
                if (!acceptCheckbox.isSelected()) {
                    ((JavascriptExecutor) driver).executeScript("arguments[0].click();", acceptCheckbox);
                }

                WebElement submitButton = driver.findElement(By.xpath("/html/body/section/div/div/form/div[6]/button"));
                Actions actions = new Actions(driver);
                actions.moveToElement(submitButton).click().perform();



            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            driver.quit();
        }
    }

    private void fillForm(String nameTitleTha, String firstnameTha, String lastnameTha, String nameTitleEng, String firstnameEng, String lastnameEng, String birthDate, String birthMonth, String birthYear, String idCard, String password, String mobile, String email, String address, String province, String district, String subDistrict, String postalCode) {
        new Select(driver.findElement(By.id("nameTitleTha"))).selectByVisibleText(nameTitleTha);
        driver.findElement(By.id("firstnameTha")).sendKeys(firstnameTha);
        driver.findElement(By.id("lastnameTha")).sendKeys(lastnameTha);
        new Select(driver.findElement(By.id("nameTitleEng"))).selectByVisibleText(nameTitleEng);
        driver.findElement(By.id("firstnameEng")).sendKeys(firstnameEng);
        driver.findElement(By.id("lastnameEng")).sendKeys(lastnameEng);
        driver.findElement(By.id("birthDate")).sendKeys(birthDate);
        driver.findElement(By.id("birthMonth")).sendKeys(birthMonth);
        driver.findElement(By.id("birthYear")).sendKeys(birthYear);
        driver.findElement(By.id("idCard")).sendKeys(idCard);
        driver.findElement(By.id("password")).sendKeys(password);
        driver.findElement(By.id("mobile")).clear();
        driver.findElement(By.id("mobile")).sendKeys(mobile);
        driver.findElement(By.id("email")).sendKeys(email);
        driver.findElement(By.id("address")).clear();
        driver.findElement(By.id("address")).sendKeys(address);
        driver.findElement(By.id("province")).sendKeys(province);
        driver.findElement(By.id("district")).sendKeys(district);
        driver.findElement(By.id("subDistrict")).clear();
        driver.findElement(By.id("subDistrict")).sendKeys(subDistrict);
        driver.findElement(By.id("postalCode")).sendKeys(postalCode);
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue().trim();
            case NUMERIC: return String.valueOf((int) cell.getNumericCellValue());
            default: return "";
        }
    }

    @AfterAll
    static void tearDown() {
        // Optionally clean up any resources if necessary
    }
}

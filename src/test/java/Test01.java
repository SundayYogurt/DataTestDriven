import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.concurrent.TimeUnit;

public class Test01 {
    @Test
    void test01() throws IOException {
        System.setProperty("webdriver.edge.driver", "./edgedriver_win64/msedgedriver.exe");

        // อ่านไฟล์ Excel
        String path = "./exel/Sci.xlsx";
        FileInputStream fs = new FileInputStream(path);
        XSSFWorkbook workbook = new XSSFWorkbook(fs);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int rowNum = sheet.getLastRowNum() + 1; // นับจำนวนแถวทั้งหมด

        // ลูปอ่านข้อมูลจาก Excel และกรอกลงในฟอร์ม
        for (int i = 1; i < rowNum-1; i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            // เปิด WebDriver (EdgeDriver)
            WebDriver driver = new EdgeDriver();
            driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
            driver.get("https://sc.npru.ac.th/sc_shortcourses/signup");

            // สร้าง WebDriverWait สำหรับใช้ในกรอกข้อมูลที่ต้องรอให้ element ปรากฏ
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10).toMillis());

            // ดึงค่าจาก Excel
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

            System.out.println(mobile);
            // กรอกข้อมูลในฟอร์ม
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

// เคลียร์ค่าก่อนกรอกเบอร์มือถือใหม่
            driver.findElement(By.id("mobile")).clear();  // เคลียร์ค่าเก่า
            driver.findElement(By.id("mobile")).sendKeys(mobile);  // กรอกเบอร์ใหม่

            driver.findElement(By.id("email")).sendKeys(email);

// เคลียร์ค่าก่อนกรอกที่อยู่ใหม่
            driver.findElement(By.id("address")).clear();  // เคลียร์ค่าเก่า
            driver.findElement(By.id("address")).sendKeys(address);

            driver.findElement(By.id("province")).sendKeys(province);
            driver.findElement(By.id("district")).sendKeys(district);

// เคลียร์ค่าก่อนกรอกตำบลใหม่
            driver.findElement(By.id("subDistrict")).clear();  // เคลียร์ค่าเก่า
            driver.findElement(By.id("subDistrict")).sendKeys(subDistrict);  // กรอกตำบลใหม่

            driver.findElement(By.id("postalCode")).sendKeys(postalCode);

// ใช้ JavaScriptExecutor เพื่อคลิก checkbox "ยอมรับข้อตกลง"
            ((org.openqa.selenium.JavascriptExecutor) driver)
                    .executeScript("document.getElementById('accept').click();");

// รอ 5 วินาทีเพื่อดูผล
            try {
                Thread.sleep(5000);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            // ปิด WebDriver
            driver.quit();
        }



        // ปิดไฟล์ Excel
        workbook.close();
        fs.close();
    }

    // ฟังก์ชันแปลงค่าจากเซลล์
    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            default:
                return "";
        }
    }
}

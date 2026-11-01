package practical7;

import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class ExcelIntegration {
    public static void main(String[] args) throws Exception {
        Path input = Path.of("data", "students.xlsx");
        Path output = Path.of("output", "students-updated.xlsx");
        Files.createDirectories(output.getParent());
        try (InputStream stream = Files.newInputStream(input); Workbook book = new XSSFWorkbook(stream)) {
            Sheet sheet = book.getSheetAt(0);
            if (sheet.getLastRowNum() != 10) throw new IllegalStateException("Exactly 10 students required");
            for (int i = 1; i <= 10; i++) {
                Row row = sheet.getRow(i);
                double marks = row.getCell(1).getNumericCellValue();
                row.getCell(1).setCellValue(marks + 5);
            }
            try (OutputStream streamOut = Files.newOutputStream(output)) {
                book.write(streamOut);
            }
        }
        ChromeOptions options = new ChromeOptions();
        if (Boolean.getBoolean("headless")) options.addArguments("--headless=new");
        WebDriver driver = new ChromeDriver(options);
        try (InputStream stream = Files.newInputStream(output); Workbook book = new XSSFWorkbook(stream)) {
            driver.get(Path.of("web", "students.html").toUri().toString());
            Sheet sheet = book.getSheetAt(0);
            for (int i = 1; i <= 10; i++) {
                Row row = sheet.getRow(i);
                String name = row.getCell(0).getStringCellValue();
                String marks = Integer.toString((int) row.getCell(1).getNumericCellValue());
                WebElement nameInput = driver.findElement(By.id("name"));
                WebElement marksInput = driver.findElement(By.id("marks"));
                nameInput.clear(); nameInput.sendKeys(name);
                marksInput.clear(); marksInput.sendKeys(marks);
                driver.findElement(By.id("save")).click();
                WebElement savedRow = driver.findElement(By.id("student-" + i));
                if (!savedRow.findElements(By.tagName("td")).get(0).getText().equals(name)
                        || !savedRow.findElements(By.tagName("td")).get(1).getText().equals(marks)) {
                    throw new IllegalStateException("Web verification failed for " + name);
                }
                System.out.println(name + " -> Updated marks: " + marks + " -> Web PASS");
            }
            if (driver.findElements(By.cssSelector("tbody tr")).size() != 10) {
                throw new IllegalStateException("Expected 10 web records");
            }
            System.out.println("PASS - 10 Excel records updated, reopened and verified on the web page");
        } finally {
            driver.quit();
        }
    }
}

package me.xeynq4.edugrades;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.util.stream.Collectors;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Main {
    public static void main(String[] args) {

        WebDriver driver = new ChromeDriver();

        driver.get("https://smnd.edupage.org/");
        driver.manage().window().maximize();
        WebElement name = driver.findElement(By.id("home_Login_1e1"));
        WebElement password = driver.findElement(By.id("home_Login_1e2"));
        WebElement login = driver.findElement(By.className("skgdFormSubmit"));

        name.sendKeys(Login.name);
        password.sendKeys(Login.password);
        login.click();

        WebElement gradesPage = driver.findElement(By.xpath("//*[@id=\"edubar\"]/div[2]/div[1]/div/ul/li[5]/a"));
        gradesPage.click();

        driver.manage().timeouts().implicitlyWait(1, TimeUnit.SECONDS);

        List<WebElement> predmety = driver.findElements(By.className("predmetRow"));

        List<Float> predmetyPercentage = predmety.stream()
                .map(row -> row.findElement(By.className("znPriemerCell")).getText().replace("%", "")).map(a -> {
                    if (a.length() > 0)
                        return a;
                    else
                        return "-1";
                }).map(a -> Float.parseFloat(a)).collect(Collectors.toList());

        List<String> predmetyNames = predmety.stream().map(row -> row.findElement(By.className("fixedCell")).getText())
                .collect(Collectors.toList());

        driver.close();

        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("0");

        int sum = 0;
        int numOfSums = 0;
        for (int i = 0; i < predmetyNames.size() + 1; i++) {
            Row row = sheet.createRow(i);
            Cell cell1 = row.createCell(0);
            Cell cell2 = row.createCell(1);
            Cell cell3 = row.createCell(2);
            if (i == predmetyNames.size()) {
                cell1.setCellValue("Priemer");
                cell2.setCellValue(sum / numOfSums);
                if (cell2.getNumericCellValue() < 45)
                    cell3.setCellValue(5);
                else if (cell2.getNumericCellValue() < 60)
                    cell3.setCellValue(4);
                else if (cell2.getNumericCellValue() < 75)
                    cell3.setCellValue(3);
                else if (cell2.getNumericCellValue() < 90)
                    cell3.setCellValue(2);
                else
                    cell3.setCellValue(1);
            } else {
                cell1.setCellValue(predmetyNames.get(i));
                if (predmetyPercentage.get(i) == -1)
                    cell2.setCellValue("");
                else {
                    cell2.setCellValue(predmetyPercentage.get(i));
                    sum += predmetyPercentage.get(i);
                    numOfSums++;
                    if (cell2.getNumericCellValue() < 45)
                        cell3.setCellValue(5);
                    else if (cell2.getNumericCellValue() < 60)
                        cell3.setCellValue(4);
                    else if (cell2.getNumericCellValue() < 75)
                        cell3.setCellValue(3);
                    else if (cell2.getNumericCellValue() < 90)
                        cell3.setCellValue(2);
                    else
                        cell3.setCellValue(1);
                }
            }
        }
        try {
            OutputStream fileOut = new FileOutputStream("grades.xls");
            wb.write(fileOut);
            wb.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

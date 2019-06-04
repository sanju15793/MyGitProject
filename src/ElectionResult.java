import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

/**
* Author: Sanjay S
* Date: 04/06/2019
* Description: Utility to get election result into an excel file
**/

public class ElectionResult {
    private static WebDriver driver;
    private final String EC_URL = "http://results.eci.gov.in/pc/en/partywise/index.htm";
    private final String chromedriverPath = "C:\\Software Downloads\\Selenium\\chromedriver\\chromedriver.exe";

    private final ArrayList ar1 = new ArrayList();
    private final ArrayList ar2 = new ArrayList();
    private final ArrayList ar3 = new ArrayList();
    private final ArrayList ar4 = new ArrayList();

    public static void main(String[] args) {
        ElectionResult electionResult = new ElectionResult();

        electionResult.startBrowser();
        electionResult.ConstituencyWiseSearch();
        electionResult.selectStateAndConstituency();
        electionResult.getResultDataFromTable();
        electionResult.printElectionResult();
        electionResult.closeBrowser();
        electionResult.createElectionResultExcel();
    }

    private void startBrowser() {
        System.setProperty("webdriver.chrome.driver", chromedriverPath);
        driver = new ChromeDriver();
        driver.get(EC_URL);
        driver.manage().window().maximize();
    }

    private void ConstituencyWiseSearch() {
        driver.findElement(By.xpath("(//a[@class='ctl00_Menu1_1 ctl00_Menu1_3'])[3]")).click();
    }

    private void selectStateAndConstituency() {
        Scanner scan = new Scanner(System.in);
        System.out.println("Please enter the state code: ");
        String state = scan.next();
        System.out.println("Please enter the constituency code: ");
        int constituency = scan.nextInt();
        Select statedropdown = new Select(driver.findElement(By.id("ddlState")));
        statedropdown.selectByValue(state);
        Select district = new Select(driver.findElement(By.id("ddlAC")));
        district.selectByIndex(constituency);
    }

    private void getResultDataFromTable() {
        int totalrows = driver.findElements(By.xpath("//div[@id='div1']/table[1]/tbody/tr")).size();
        for (int i = 4; i < totalrows; i++) {
            ar1.add(driver.findElement(By.xpath("//div[@id='div1']/table[1]/tbody/tr[" + i + "]/td[2]")).getText());
            ar2.add(driver.findElement(By.xpath("//div[@id='div1']/table[1]/tbody/tr[" + i + "]/td[3]")).getText());
            ar3.add(driver.findElement(By.xpath("//div[@id='div1']/table[1]/tbody/tr[" + i + "]/td[6]")).getText());
            ar4.add(driver.findElement(By.xpath("//div[@id='div1']/table[1]/tbody/tr[" + i + "]/td[7]")).getText());
        }
    }

    private void printElectionResult() {
        for (int i = 0; i < ar1.size(); i++) {
            System.out.println("Candidate : " + ar1.get(i));
            System.out.println("party :" + ar2.get(i));
            System.out.println("votes :" + ar3.get(i));
            System.out.println("% of votes :" + ar4.get(i));
            System.out.println("**************");
            System.out.println(" ");
        }
    }

    private void createElectionResultExcel() {

        try {
            File f = new File("C:/Users/Sanjay.S/ProperGitProject/files/ECResult1.xlsx");
            FileInputStream fin = new FileInputStream(f);
            XSSFWorkbook wb = new XSSFWorkbook(fin);
            XSSFSheet sh1 = wb.getSheetAt(0);

            int totalrowsinexcel = sh1.getLastRowNum();
            for (int i = 0; i <= totalrowsinexcel; i++) {
                sh1.removeRow(sh1.getRow(i));
            }

            System.out.println("removed all previous data in the excel file");

            sh1.createRow(0).createCell(0).setCellValue("Candidate");
            sh1.getRow(0).createCell(1).setCellValue("Party");
            sh1.getRow(0).createCell(2).setCellValue("Total Votes");

            sh1.createRow(1).createCell(0).setCellValue((String) ar1.get(0));
            sh1.getRow(1).createCell(1).setCellValue((String) ar2.get(0));
            sh1.getRow(1).createCell(2).setCellValue((String) ar3.get(0));

            for (int x = 2; x <= ar1.size(); x++) {
                sh1.createRow(x).createCell(0).setCellValue((String) ar1.get(x - 1));
                sh1.getRow(x).createCell(1).setCellValue((String) ar2.get(x - 1));
                sh1.getRow(x).createCell(2).setCellValue((String) ar3.get(x - 1));
            }

            FileOutputStream fout = new FileOutputStream(f);
            fin.close();
            wb.write(fout);
            fout.close();
            System.out.println("Election result details written into the excel file successfully");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void closeBrowser() {
        driver.quit();
    }
}


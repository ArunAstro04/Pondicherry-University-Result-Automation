import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class UniversityResult {
	public static void main(String a[]) throws FileNotFoundException, IOException{
		System.setProperty("webdriver.chrome.driver","D:\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("http://exam.pondiuni.edu.in/oresults/");
		File excelFile=new File("C:\\Users\\Hp\\Downloads\\test\\studentmark.xlsx");
		InputStream fin=new FileInputStream(excelFile);
		XSSFWorkbook wb = new XSSFWorkbook(fin);
		XSSFSheet sh = wb.getSheetAt(0);
		XSSFWorkbook workbook = new XSSFWorkbook();
		FileOutputStream fileOut=new FileOutputStream("C:\\Users\\Hp\\Downloads\\test\\2.xlsx");;
		XSSFSheet sheet = workbook.createSheet("Student Data");
		int x=1,lineNo=0;
		ArrayList<String> head = new ArrayList<String>();
		//ArrayList<String> data = new ArrayList<String>();
		Iterator<Row> rows = sh.rowIterator();
		XSSFRow  Row=null;
		XSSFCell cell=null;
		while(rows.hasNext()){
			XSSFCell rollNo=wb.getSheetAt(0).getRow(x).getCell(0);
			//System.out.println(rollNo.toString());
			WebElement registerno = driver.findElement(By.id("reg_no"));
			registerno.sendKeys(rollNo.toString());
			WebElement semester = driver.findElement(By.id("exam"));
			semester.sendKeys("Sixth");
			WebElement submit = driver.findElement(By.xpath("//*[@id=\"print_app_form\"]"));
			submit.click(); 
			WebDriverWait wait = new WebDriverWait(driver, 60); 
			wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("/html/body/div[1]/div/div[3]/div/div[2]/div[1]/div[2]/table[1]/tbody/tr[2]/td"))); 
			// To find subject name
			WebElement sub1 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[2]/td[2]"));
			String subject1 = sub1.getText();
			
			WebElement sub2 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[3]/td[2]"));
			String subject2 = sub2.getText();
			
			WebElement sub3 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[4]/td[2]"));
			String subject3 = sub3.getText();
			
			WebElement sub4 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[5]/td[2]"));
			String subject4 = sub4.getText();
			
			WebElement sub5 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[6]/td[2]"));
			String subject5 = sub5.getText();
			
			WebElement sub6 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[7]/td[2]"));
			String subject6 = sub6.getText();
			
			WebElement sub7 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[8]/td[2]"));
			String subject7 = sub7.getText();
			
			WebElement sub8 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[9]/td[2]"));
			String subject8 = sub8.getText();
			
			WebElement sub9 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[10]/td[2]"));
			String subject9 = sub9.getText();
			
			WebElement sub10 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[11]/td[2]"));
			String subject10 = sub10.getText();
			
			// To find Name and Grade
			WebElement nam = driver.findElement(By.xpath("//*[@id=\"student_info\"]/tbody/tr[3]/td"));
			String name = nam.getText();
			
			WebElement gra1 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[2]/td[7]/div"));
			String grade1 = gra1.getText();
			
			WebElement gra2 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[3]/td[7]/div"));
			String grade2 = gra2.getText();
			
			WebElement gra3 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[4]/td[7]"));
			String grade3 = gra3.getText();
			
			WebElement gra4 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[5]/td[7]"));
			String grade4 = gra4.getText();
			
			WebElement gra5 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[6]/td[7]"));
			String grade5 = gra5.getText();
			
			WebElement gra6 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[7]/td[7]"));
			String grade6  = gra6.getText();
			
			WebElement gra7 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[8]/td[7]"));
			String grade7 = gra7.getText();
			
			WebElement gra8 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[9]/td[7]"));
			String grade8 = gra8.getText();
			
			WebElement gra9 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[10]/td[7]"));
			String grade9 = gra9.getText();
			
			WebElement gra10 = driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[11]/td[7]"));
			String grade10 = gra10.getText();
			
			driver.navigate().back();
			if(lineNo==0){
				head.add("Register Number");
				head.add("Name");
	        	head.add(subject1);
	        	head.add(subject2);
	        	head.add(subject3);
	        	head.add(subject4);
	        	head.add(subject5);
	        	head.add(subject6);
	        	head.add(subject7);
	        	head.add(subject8);
	        	head.add(subject9);
	        	head.add(subject10);
	        	Row  = sheet.createRow(lineNo);
	        	for (int i = 0; i < head.size(); i++) 
				{	
	        		cell = Row.createCell(i);
	        		cell.setCellValue(head.get(i));
	        		sheet.autoSizeColumn(i);
				}
	        	head.remove("Register Number");
				head.remove("Name");
	        	head.remove(subject1);
	        	head.remove(subject2);
	        	head.remove(subject3);
	        	head.remove(subject4);
	        	head.remove(subject5);
	        	head.remove(subject6);
	        	head.remove(subject7);
	        	head.remove(subject8);
	        	head.remove(subject9);
	        	head.remove(subject10);
	        	lineNo++;
			}
			
			head.add(rollNo.toString());
			head.add(name.substring(21));
			head.add(grade1);
			head.add(grade2);
			head.add(grade3);
			head.add(grade4);
			head.add(grade5);
			head.add(grade6);
			head.add(grade7);
			head.add(grade8);
			head.add(grade9);
			head.add(grade10);
			Row  = sheet.createRow(lineNo);
			
			for (int j=0; j < head.size(); j++) 
			{	
        		cell = Row.createCell(j);
        		cell.setCellValue(head.get(j));
        		sheet.autoSizeColumn(j);
        		System.out.print(head.get(j)+" ");
			}
			
			head.remove(rollNo.toString());
			head.remove(name.substring(21));
			head.remove(grade1);
			head.remove(grade2);
			head.remove(grade3);
			head.remove(grade4);
			head.remove(grade5);
			head.remove(grade6);
			head.remove(grade7);
			head.remove(grade8);
			head.remove(grade9);
			head.remove(grade10);
			lineNo++;
        	x++;
			System.out.println("Ex");
			if(lineNo==47){
				workbook.write(fileOut);
				break;
			}
		}
		//fin.close();
		wb.close();
		workbook.close();
		fileOut.close();
	}
}

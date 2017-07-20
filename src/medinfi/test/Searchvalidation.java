package medinfi.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;

public class Searchvalidation {
	String Excel1,Excel2;
	void retreivetestdata()throws Exception{
		File input=new File("./Testdata/Input.xlsx");
		FileInputStream fis=new FileInputStream(input);
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheetAt(0);
		Excel1=sheet.getRow(0).getCell(0).toString();
		Excel2=sheet.getRow(0).getCell(1).toString();
		wb.close();
		
	}
	public static ArrayList retreiveDoctors() throws IOException
	{
		File doctor=new File("./Testdata/doctor's name.xlsx");
		FileInputStream fis1=new FileInputStream(doctor);
		XSSFWorkbook wb1=new XSSFWorkbook(fis1);
		XSSFSheet Excel3=wb1.getSheetAt(0);
		ArrayList<String> var = new ArrayList<String>();
		for(int i=0;Excel3.getRow(i)!=null;i++)
		{
			var.add(Excel3.getRow(i).getCell(0).getStringCellValue());
		}
		wb1.close();
		return var;
	}
	public static void main(String[] args) throws Exception  {
		Searchvalidation obj=new Searchvalidation();
		obj.retreivetestdata();
		System.setProperty("webdriver.gecko.driver","./exe file/geckodriver.exe");
		boolean validation=true;
		WebDriver driver=new FirefoxDriver();
		driver.get("http://www.medinfi.com/");
		driver.findElement(By.xpath("html/body/div[1]/div/div/div/div[2]/div[1]/div/div[3]/div[1]/div[1]/div/input")).sendKeys(obj.Excel1);
		Thread.sleep(3000);
		WebElement locality = driver.findElement(By.id("autoCityResult1"));
		if(locality.isDisplayed())
		{
			locality.click();

		}
		driver.findElement(By.xpath("html/body/div[1]/div/div/div/div[2]/div[1]/div/div[3]/div[2]/div[1]/div[1]/input")).sendKeys(obj.Excel2);
		Thread.sleep(8000);
		WebElement doctorname = driver.findElement(By.id("autoResult1"));
		ArrayList savedDoctorList=retreiveDoctors();
		ArrayList<String> mismatchedDoctorList=new ArrayList<String>();
		String displayedDoctorList[] = null;
		if(doctorname.isDisplayed())
		{
			
			displayedDoctorList=doctorname.getText().split("\n");
			
		}
		for(String doctor:displayedDoctorList)
		{
			if(!savedDoctorList.contains(doctor))
			{
				mismatchedDoctorList.add(doctor);
				validation=false;
			}
		}
		driver.manage().timeouts().implicitlyWait(100, TimeUnit.SECONDS);
		XWPFDocument document = new XWPFDocument();
        FileOutputStream out = new FileOutputStream(
        new File("./TestOutput/TestReport" + ".docx"));
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("Test Report for MedInfi");
        run.addBreak();
        
        
        if (validation)
		{
        	run.setText("Values in the drop down match with expected output Data set");
        	run.addBreak();
        	run.setText("Testing completed Succesfully");
		}
		else
		{
			run.setText("Values in the drop down do not match with expected output Data set");
			run.addBreak();
			run.setText("Below are the mismatched values");
			for(String mismatchedDoctor:mismatchedDoctorList)
			{
				run.addBreak();
				run.setText(mismatchedDoctor);
				
			}
		}
		driver.close();
		document.write(out);
	    System.out.println("Test report generated successfully");
        out.close();
	 }
	}

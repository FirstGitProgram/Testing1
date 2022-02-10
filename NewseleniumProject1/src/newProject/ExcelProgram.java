package newProject;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelProgram
{
	public static void main(String[] args) throws Exception 
	{
		FileInputStream fs=new  FileInputStream("C:\\Users\\LENOVO\\Documents\\ExcelFile\\Book.xlsx");
		XSSFWorkbook aws= new XSSFWorkbook(fs);
		
		XSSFSheet sheet=aws.getSheetAt(0);
		
		XSSFRow row= sheet.getRow(0);
		
		XSSFCell col=row.getCell(2);
		
		System.out.println(col.getStringCellValue());
		
		int rowcount= sheet.getLastRowNum();
		
		System.out.println("Row count"+rowcount);
		
		int totalrow=rowcount+1;
		System.out.println("total row count"+ totalrow);
		
		 int colcount=sheet.getRow(rowcount).getLastCellNum();
		 
		 System.out.println("column is "+colcount);
		 for(int i=0; i< totalrow ;i++)
		 {
			 for(int j=0; j<colcount ; j++)
			 {
				 System.out.println(sheet.getRow(i).getCell(j));
			 }
		 }
		
	}

}

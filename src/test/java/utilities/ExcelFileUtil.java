package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtil {
	
	Workbook wb;
	//to read path of excel constructor
	// y I am writing constructor here, the path of excel is every where require, If I create constructor here where ever I want I can invoke it by initialization object for it
	//By default constructor maintain public access modifier
	public ExcelFileUtil(String Excelpath)throws Throwable
	{
		FileInputStream fi = new FileInputStream(Excelpath);
		//from that file we need to get workbook
		wb = WorkbookFactory.create(fi);
	}
	//method for counting rows in sheet
	public int rowCount(String sheetName)
	{
		return wb.getSheet(sheetName).getLastRowNum();
	}
	//method for reading cell data
	public String getCellData(String sheetName,int row,int column)
	{
		String data ="";
		//**********************************************************************************************************************************************
		//go to my shet in that sheet if any rows if any columns type is numeric read that numaric cell in to one variable and then convert in to string type
		if(wb.getSheet(sheetName).getRow(row).getCell(column).getCellType()==CellType.NUMERIC)
		{
			//
			int celldata =(int) wb.getSheet(sheetName).getRow(row).getCell(column).getNumericCellValue();
			//
			data = String.valueOf(celldata);
					
		}
		//when else part execute if any cell contain string type data
		else
		{
			data =wb.getSheet(sheetName).getRow(row).getCell(column).getStringCellValue();
		}
		return data;
		
	}
	public void setCellData(String sheetName,int row,int column,String status,String WriteExcel)throws Throwable
	{
		//get sheet from wb
		//Sheet interface,not get new keyword here
		Sheet ws = wb.getSheet(sheetName);
		//get row from sheet
		//Row interface,not get new keyword here
		Row rowNum =ws.getRow(row);
		//create cell in row
		Cell cell = rowNum.createCell(column);
		//write status
		//cell interface, not get new keyword here
		cell.setCellValue(status);
		//equalsIgnoreCase means case sensitive small or capital
		if(status.equalsIgnoreCase("Pass"))
		{
			CellStyle style = wb.createCellStyle();
			//Font interface, not get new keyword here
			Font font = wb.createFont();
			font.setColor(IndexedColors.GREEN.getIndex());
			font.setBold(true);
			
			style.setFont(font);
			ws.getRow(row).getCell(column).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("Fail"))
		{
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			font.setColor(IndexedColors.RED.getIndex());
			font.setBold(true);
			style.setFont(font);
			ws.getRow(row).getCell(column).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("Blocked"))
		{
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			font.setColor(IndexedColors.BLUE.getIndex());
			font.setBold(true);
			style.setFont(font);
			ws.getRow(row).getCell(column).setCellStyle(style);
		}
		FileOutputStream fo = new FileOutputStream(WriteExcel);
		wb.write(fo);
	}
	
	public static void main(String[] args) throws Throwable  {
		
		ExcelFileUtil xl = new ExcelFileUtil("D:/Selenium/selenium live project/Sample Excel.xlsx");
		//count number of rows in sheet
		int rc = xl.rowCount("Employ");
		System.out.println(rc);
		for(int i=1;i<=rc;i++) {
			String fname = xl.getCellData("Employ", i, 0);
			String mname = xl.getCellData("Employ", i, 1);
			String lname = xl.getCellData("Employ", i, 2);
			String eid = xl.getCellData("Employ", i, 3);
			System.out.println(fname+" "+mname+" "+lname+" "+eid);
			//xl.setCellData("Employ", i, 4, "Fail", "D:/Selenium/selenium live project/Results.xlsx");
			//xl.setCellData("Employ", i, 4, "Pass", "D:/Selenium/selenium live project/Results.xlsx");
			xl.setCellData("Employ", i, 4, "Blocked", "D:/Selenium/selenium live project/Results.xlsx");
		}
		
	}


}

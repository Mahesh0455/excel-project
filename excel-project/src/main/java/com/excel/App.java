package com.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.io.*;

public class App {
	public static void main(String[] args) throws Exception {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet spreadsheet = workbook.createSheet(" Student-Data ");
		
	
		XSSFRow row;
		
		Map<String, Object[]> student_data = new TreeMap<String, Object[]>();

		student_data.put("0", new Object[] { "Roll No","Name", "Year","Result" });
		student_data.put("1", new Object[] { "104","Mahesh", "2020","Pass" });
		student_data.put("2", new Object[] { "107","Ramesh", "2022","Pass" });
		student_data.put("3", new Object[] { "108","Mohan", "2021","Fail	" });
		student_data.put("4", new Object[] { "109","Gopal", "2022" });
		student_data.put("5", new Object[] { "109","Satish", "2022","Fail" });

		int rowid = 0;
		

		Set<String> keys = student_data.keySet();

		for (String key : keys) {
			row = spreadsheet.createRow(rowid++);
			Object obj[] = student_data.get(key);
			int cellid = 0;
	
			for (Object o : obj) {
				Cell cell = row.createCell(cellid++);
				cell.setCellValue((String) o);
			}
		}
		
		//System.out.println(spreadsheet.getRow(2).getCell(1).getStringCellValue());

		FileOutputStream out = new FileOutputStream(
				new File("C:/Users/MPawar/Desktop/mahesh/learnings/java/alt/Student_Details.xlsx"));
		workbook.write(out);
		out.flush();
		out.close();
		
		File file=new File("C:/Users/MPawar/Desktop/mahesh/learnings/java/alt/Student_Details.xlsx");
		
		ReadExcel read=new ReadExcel();
		read.readExcel(file);
	}
}

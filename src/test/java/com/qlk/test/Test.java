package com.qlk.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {

	public static void main(String[] args) throws Exception {
		Request request=new Request();
		//System.out.println(request.getRequestURL(18));
		//System.out.println(request.sendPost(18));
		//request.writeResult("qqqqq", 2);
		File file = new File("C:\\Users\\cbhbit\\Desktop\\test.xls");
		// 下面尝试更改第一行第一列的单元格的值
		updateExcel(file, "Sheet1", 0, 0, "hehe");

	}

	public static void updateExcel(File exlFile, String sheetName, int col, int row, String value) throws Exception {
		FileInputStream fis = new FileInputStream(exlFile);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		// workbook.
		XSSFSheet sheet = workbook.getSheet(sheetName);
		XSSFCell mycell = sheet.createRow(row).createCell(col);
		mycell.setCellValue(value);
		XSSFRow r = sheet.getRow(row);
		XSSFCell cell = r.getCell(col);
		// int type=cell.getCellType();
		String str1 = cell.getStringCellValue();
		// 这里假设对应单元格原来的类型也是String类型
		cell.setCellValue(value);
		System.out.println("单元格原来值为" + str1);
		System.out.println("单元格值被更新为" + value);

		fis.close();// 关闭文件输入流

		FileOutputStream fos = new FileOutputStream(exlFile);
		workbook.write(fos);
		fos.close();// 关闭文件输出流
	}

}

package com.qlk.test;

public class Test {

	public static void main(String[] args) {
		ExcelReader excelReader=new ExcelReader("C:\\Users\\cbhbit\\Desktop\\ODC.xlsx");
		try {
			System.out.println(excelReader.readExcelTitle()[0]);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}

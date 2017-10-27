package com.qlk.test;

public class Test {

	public static void main(String[] args){
		//String filePath="C:\\Users\\cbhbit\\Desktop\\test.xlsx";
		Worker worker=new Worker("C:\\Users\\cbhbit\\Desktop\\test.xlsx");
		//worker.work();
		//worker.dataCheck();
		worker.dataCheckByRows();
		
	}
}

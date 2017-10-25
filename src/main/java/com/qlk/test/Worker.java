package com.qlk.test;

import java.util.List;
import java.util.Map;

public class Worker {
	private String filePath = "C:\\Users\\cbhbit\\Desktop\\ODC.xlsx";;

	Excel excel = new Excel();
	List<Map<Integer, Map<Integer, Object>>> content;

	HTTPRequest request = new HTTPRequest();

	Worker() {
		content = excel.readExcelXLSXContent(filePath);

	}
	Worker(String filePath) {
		this.filePath=filePath;
		if(filePath.endsWith("xlsx")){
			content = excel.readExcelXLSXContent(filePath);
		}else if(filePath.endsWith("xls")){
			content = excel.readExcelXLSContent(filePath);
		}else
			System.out.println("文件类型不支持！");

	}

	public String[] mapToString(Map<Integer, Map<Integer, Object>> content, int rowNumber) {
		String[] row = new String[content.get(rowNumber).size()];
		for (int i = 0; i < content.get(rowNumber).size(); i++) {
			row[i] = String.valueOf(content.get(rowNumber).get(i));
		}
		return row;
	}

	public void work() {
		if(content.isEmpty()){
			System.out.println("文件内容为空！");
			return ;
		}
		String host = mapToString(content.get(0), 1)[0];
		String[] actualResult = new String[content.get(1).size()];
		actualResult[0] = mapToString(content.get(1), 0)[7];

		for (int i = 1; i < content.get(1).size(); i++) {
			String url = host + mapToString(content.get(1), i)[0] + mapToString(content.get(1), i)[1];
			String params = mapToString(content.get(1), i)[2];
			String method = mapToString(content.get(1), i)[5];

			actualResult[i] = request.sendrequest(url, params, method);
			System.out.println("第"+i+"条记录正在执行，请稍后!");
		}
		System.out.println("Start to write to excel...");
		if(filePath.endsWith("xlsx")){
			excel.writeExcelXLSXByColumn(actualResult, filePath, 1, 7);
		}else if(filePath.endsWith("xls")){
			excel.writeExcelXLSByColumn(actualResult, filePath, 1, 7);
		}		
		System.out.println("Done !");

		String[] result = new String[content.get(1).size()];
		result[0] = mapToString(content.get(1), 0)[8];

		for (int i = 1; i < content.get(1).size(); i++) {
			if (actualResult[i].contains(mapToString(content.get(1), i)[6]))
				result[i] = "Pass";
			else
				result[i] = "Fail";
		}

		System.out.println("Start to write to excel...");
		if(filePath.endsWith("xlsx")){
			excel.writeExcelXLSXByColumnPattern(result, filePath, 1, 8);
		}else if(filePath.endsWith("xls")){
			excel.writeExcelXLSByColumnPattern(result, filePath, 1, 8);
		}
		System.out.println("Done !");
	}

}

package com.qlk.test;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.codehaus.jackson.map.ObjectMapper;

public class Worker {
	private String filePath = "C:\\Users\\cbhbit\\Desktop\\ODC.xlsx";;

	Excel excel = new Excel();
	List<Map<Integer, Map<Integer, Object>>> content;

	HTTPRequest request = new HTTPRequest();

	Worker() {
		content = excel.readExcelXLSXContent(filePath);

	}

	Worker(String filePath) {
		this.filePath = filePath;
		if (filePath.endsWith("xlsx")) {
			content = excel.readExcelXLSXContent(filePath);
		} else if (filePath.endsWith("xls")) {
			content = excel.readExcelXLSContent(filePath);
		} else
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
		if (content.isEmpty()) {
			System.out.println("文件内容为空！");
			return;
		}
		String host = mapToString(content.get(0), 1)[0];
		String[] actualResult = new String[content.get(1).size()];
		actualResult[0] = mapToString(content.get(1), 0)[7];

		for (int i = 1; i < content.get(1).size(); i++) {
			String url = host + mapToString(content.get(1), i)[0] + mapToString(content.get(1), i)[1];
			String params = mapToString(content.get(1), i)[2];
			String method = mapToString(content.get(1), i)[5];

			actualResult[i] = request.sendrequest(url, params, method);
			System.out.println("第" + i + "条记录正在执行，请稍后!");
		}
		System.out.println("Start to write to excel...");
		if (filePath.endsWith("xlsx")) {
			excel.writeExcelXLSXByColumn(actualResult, filePath, 1, 7);
		} else if (filePath.endsWith("xls")) {
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
		if (filePath.endsWith("xlsx")) {
			excel.writeExcelXLSXByColumnPattern(result, filePath, 1, 8);
		} else if (filePath.endsWith("xls")) {
			excel.writeExcelXLSByColumnPattern(result, filePath, 1, 8);
		}
		System.out.println("Done !");
	}

	public void dataCheck() {
		if (content.isEmpty()) {
			System.out.println("文件内容为空！");
			return;
		}
		String url = mapToString(content.get(0), 1)[0];
		String method = mapToString(content.get(0), 1)[1];
		String param = mapToString(content.get(0), 1)[2];
		String affect = mapToString(content.get(0), 1)[3];

		String[] actualResult = new String[content.get(1).size()];
		String[] result = new String[content.get(1).size()];
		String[] actual_result = new String[content.get(1).size()];
		ObjectMapper mapper = new ObjectMapper();
		// 循环列数
		for (int j = 0; j < mapToString(content.get(1), 0).length - 1; j++) {
			actualResult[0] = mapToString(content.get(1), 0)[j + 1] + "_result";
			int divNum = param.indexOf(mapToString(content.get(1), 0)[0]) + mapToString(content.get(1), 0)[0].length()
					+ 3;
			String param1 = param.substring(0, divNum);
			String param2 = param.substring(divNum);
			// 每行记录发送请求，获得response,循环行数
			for (int i = 1; i < content.get(1).size(); i++) {

				String params = param1 + mapToString(content.get(1), i)[0] + param2;
				// System.out.println(params);
				actualResult[i] = request.sendrequest(url, params, method);
				// System.out.println(request.sendrequest(url, params, method));
				// System.out.println("---------------------------------------");
				System.out.println("第" + i + "条记录正在执行，请稍后!");
			}

			result[0] = actualResult[0];
			actual_result[0] = "actual_" + mapToString(content.get(1), 0)[j + 1];
			String keyToValue = mapToString(content.get(1), 0)[j + 1] + "=";
			//
			for (int i = 1; i < content.get(1).size(); i++) {
				try {
					Map<?, ?> m = mapper.readValue(actualResult[i], Map.class);
					actual_result[i] = getValue(m.toString(), keyToValue);
					// System.out.println(actual_result[i]);
					if (m.toString().contains(keyToValue)) {
						if (m.toString().contains(keyToValue + mapToString(content.get(1), i)[j + 1]))
							result[i] = "Pass";
						else
							result[i] = "Fail";

					} else
						result[i] = "Fail";
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
					result[i] = "Fail";
				}

			}

			// System.out.println("Start to write to excel " + j + " column");
			if (filePath.endsWith("xlsx")) {
				excel.writeExcelXLSXByColumnPattern(result, filePath, 1, content.get(1).get(0).size() + j * 2 + 1);
				excel.writeExcelXLSXByColumn(actual_result, filePath, 1, content.get(1).get(0).size() + j * 2);
			} else if (filePath.endsWith("xls")) {
				excel.writeExcelXLSByColumnPattern(result, filePath, 1, content.get(1).get(0).size() + j * 2 + 1);
				excel.writeExcelXLSByColumn(actual_result, filePath, 1, content.get(1).get(0).size() + j * 2);
			}

		}
		System.out.println("Done !");
	}

	/**
	 * 返回jason串中的指定key值的value
	 * 
	 * @param jasonString
	 * @param keyToValue
	 * @return
	 */
	public String getValue(String jasonString, String keyToValue) {
		// System.out.println(jasonString);
		// System.out.println(keyToValue);
		String value = null;
		int index = jasonString.indexOf(keyToValue);
		int end = index + keyToValue.length();
		for (int i = end; i < jasonString.length(); i++) {
			if (jasonString.charAt(i) == ',' || jasonString.charAt(i) == '{' || jasonString.charAt(i) == '}'
					|| jasonString.charAt(i) == '[' || jasonString.charAt(i) == ']') {
				end = i;
				break;
			}
		}
		// System.out.println(index + " " + end);
		value = jasonString.substring(index, end);
		return value;
	}

	/**
	 * 返回每个接口对应数组
	 * 
	 * @param rowNumber
	 * @return
	 */
	public int[] getColumNumbers(int rowNumber) {
		String s = mapToString(content.get(0), rowNumber)[3];
		String[] columString = s.split(",");
		int[] colums = new int[columString.length];
		for (int i = 0; i < columString.length; i++) {
			colums[i] = Integer.parseInt(columString[i]);
		}
		return colums;
	}

	public void dataCheckByRows() {
		if (content.isEmpty()) {
			System.out.println("文件内容为空！");
			return;
		}
		// 数据行数(包含表头)
		int dataRows = content.get(1).size();		
		String[] actualResult = new String[dataRows];
		int cols=mapToString(content.get(1), 0).length;
		String[][] result = new String[dataRows][cols];
		String[][] actual_result = new String[dataRows][cols];
		ObjectMapper mapper = new ObjectMapper();
		
		String url ;
		String method ;
		String param ;
		String affect ;
		//根据第一个表里的行数来确定接口数目
		for(int k=1;k<content.get(0).size();k++){
		url = mapToString(content.get(0), k)[0];
		method = mapToString(content.get(0), k)[1];
		param = mapToString(content.get(0), k)[2];
		//affect = mapToString(content.get(0), k)[3];
		int[] coluNums=getColumNumbers(k);
		
		// 循环列数
		//for (int j = 0; j < mapToString(content.get(1), 0).length - 1; j++) {
			actualResult[0] = "result";
			int divNum = param.indexOf(mapToString(content.get(1), 0)[0]) + mapToString(content.get(1), 0)[0].length()
					+ 3;
			String param1 = param.substring(0, divNum);
			String param2 = param.substring(divNum);
			// 每行记录发送请求，获得response,循环行数
			for (int i = 1; i < content.get(1).size(); i++) {

				String params = param1 + mapToString(content.get(1), i)[0] + param2;
				actualResult[i] = request.sendrequest(url, params, method);
				System.out.println("第" + i + "条记录正在执行，请稍后!");
			}
			//根据接口里的列的数值来决定每次请求验证的列数
			for (int j = 0; j < coluNums.length ; j++) {
			result[0][coluNums[j]] = mapToString(content.get(1), 0)[coluNums[j]] + "_result";
			actual_result[0][coluNums[j]] = "actual_" + mapToString(content.get(1), 0)[coluNums[j]];
			String keyToValue = mapToString(content.get(1), 0)[coluNums[j]] + "=";
			//
			for (int i = 1; i < dataRows; i++) {
				try {
					Map<?, ?> m = mapper.readValue(actualResult[i], Map.class);
					actual_result[i][coluNums[j]] = getValue(m.toString(), keyToValue);
					if (m.toString().contains(keyToValue)) {
						if (m.toString().contains(keyToValue + mapToString(content.get(1), i)[coluNums[j]]))
							result[i][coluNums[j]] = "Pass";
						else
							result[i][coluNums[j]] = "Fail";

					} else
						result[i][coluNums[j]] = "Fail";
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
					result[i][coluNums[j]] = "Fail";
				}

			}}
		}
		for(int j=1;j<dataRows;j++){
			for(int i=1;i<cols;i++){
				//System.out.println(actual_result[j][i]+" "+result[j][i]);
				if (filePath.endsWith("xlsx")) {
					excel.writeExcelXLSXResultPass(actual_result[j][i],result[j][i], filePath, 1, j,content.get(1).get(0).size() + (i-1) * 2 );
				} else if (filePath.endsWith("xls")) {
					excel.writeExcelXLSResultPass(actual_result[j][i],result[j][i], filePath, 1, j,content.get(1).get(0).size() + (i-1) * 2 );
				}
			}
		}
		System.out.println("Done !");
	}
}

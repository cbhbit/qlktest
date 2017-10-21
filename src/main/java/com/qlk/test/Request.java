package com.qlk.test;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.net.URL;
import java.net.URLConnection;

public class Request {
	private String filePath = "D:\\Backup\\桌面\\ODC.xlsx";
	private String[] requestRow;

	public Request() {

	}

	public Request(String filePath) {
		this.filePath = filePath;
	}

	public String[] setRequestMethod(String filePath, int sheetNumber, int rowNumber) {
		Excel excelReader = new Excel(filePath);
		try {
			requestRow=excelReader.getContentRow(sheetNumber, rowNumber);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		//System.out.println(requestRow.length);
		return requestRow;
	}

	private String requestURL;

	public String getRequestURL(int rowNumber) {
		requestURL = setRequestMethod(filePath, 0, 1)[0] + setRequestMethod(filePath, 1, rowNumber)[0]
				+ setRequestMethod(filePath, 1, rowNumber)[1];
		return requestURL;
	}

	public String sendGet(int rowNumber) {
		String result = "";
		String url = getRequestURL(rowNumber);
		String param = setRequestMethod(filePath, 1, rowNumber)[2];
		BufferedReader in = null;
		String urlNameString;
		try {
			if (param == "")
				urlNameString = url;
			else
				urlNameString = url + "?" + param;
			URL realUrl = new URL(urlNameString);
			// System.out.println(urlNameString);
			// �򿪺�URL֮�������
			URLConnection connection = realUrl.openConnection();
			// ����ͨ�õ���������
			connection.setRequestProperty("accept", "*/*");
			connection.setRequestProperty("connection", "Keep-Alive");
			connection.setRequestProperty("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1;SV1)");
			// ����ʵ�ʵ�����
			connection.connect();
			/*
			 * // ��ȡ������Ӧͷ�ֶ� Map<String, List<String>> map =
			 * connection.getHeaderFields(); // �������е���Ӧͷ�ֶ� for (String key :
			 * map.keySet()) { System.out.println(key + "--->" + map.get(key)); }
			 */
			// ���� BufferedReader����������ȡURL����Ӧ
			in = new BufferedReader(new InputStreamReader(connection.getInputStream()));
			String line;
			while ((line = in.readLine()) != null) {
				result += line;
			}
		} catch (Exception e) {
			System.out.println("����GET��������쳣��" + e);
			e.printStackTrace();
		}
		// ʹ��finally�����ر�������
		finally {
			try {
				if (in != null) {
					in.close();
				}
			} catch (Exception e2) {
				e2.printStackTrace();
			}
		}
		return result;
	}

	public String sendPost(int rowNumber) {
		PrintWriter out = null;
		String url = getRequestURL(rowNumber);
		String param = setRequestMethod(filePath, 1, rowNumber)[2];
		BufferedReader in = null;
		String result = "";
		String urlNameString;
		try {
			urlNameString = url;

			URL realUrl = new URL(urlNameString);
			// �򿪺�URL֮�������
			URLConnection conn = realUrl.openConnection();
			// ����ͨ�õ���������
			conn.setRequestProperty("accept", "*/*");
			conn.setRequestProperty("connection", "Keep-Alive");
			conn.setRequestProperty("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1;SV1)");
			conn.setRequestProperty("Content-Type", "application/json");
			// ����POST�������������������
			conn.setDoOutput(true);
			conn.setDoInput(true);
			// ��ȡURLConnection�����Ӧ�������
			out = new PrintWriter(conn.getOutputStream());
			// �����������
			out.print(param);
			// flush������Ļ���
			out.flush();
			// ����BufferedReader����������ȡURL����Ӧ
			in = new BufferedReader(new InputStreamReader(conn.getInputStream()));
			String line;
			while ((line = in.readLine()) != null) {
				result += line;
			}
		} catch (Exception e) {
			System.out.println("���� POST ��������쳣��" + e);
			e.printStackTrace();
		}
		// ʹ��finally�����ر��������������
		finally {
			try {
				if (out != null) {
					out.close();
				}
				if (in != null) {
					in.close();
				}
			} catch (IOException ex) {
				ex.printStackTrace();
			}
		}
		return result;
	}

	public void writeResult(String value, int rowNumber) throws Exception {
		Excel excelReader = new Excel(filePath);
		excelReader.writeExcel(1, rowNumber, 7, value);
		
	}
	
}

package com.qlk.test;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.net.URL;
import java.net.URLConnection;

public class HTTPRequest {
	private String result = "";
	String urlNameString;
	BufferedReader in = null;
	
	public String getResult() {
		return result;
	}
	
	public String sendrequest(String url,String params,String method){
		if(method.equals("GET"))
			return sendGet(url,params,method);
		else if(method.equals("POST"))
			return sendPost(url,params,method);
		else
			return "HTTP method is wrong!";			
	}

	public String sendGet(String url,String params,String method) {
		try {
			if (params == "")
				urlNameString = url;
			else
				urlNameString = url + "?" + params;
			URL realUrl = new URL(urlNameString);
			//System.out.println(urlNameString);
			URLConnection connection = realUrl.openConnection();

			connection.setRequestProperty("accept", "*/*");
			connection.setRequestProperty("connection", "Keep-Alive");
			connection.setRequestProperty("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1;SV1)");

			connection.connect();

			in = new BufferedReader(new InputStreamReader(connection.getInputStream()));
			String line;
			while ((line = in.readLine()) != null) {
				result += line;
			}
		} catch (Exception e) {
			//System.out.println("There is something wrong in GET request!" + e);
			return e.toString();
		}

		finally {
			try {
				if (in != null) {
					in.close();
				}
			} catch (Exception e2) {
				//e2.printStackTrace();
			}
		}
		return result;
	}

	public String sendPost(String url,String params,String method) {
		PrintWriter out = null;

		try {
			urlNameString = url;

			URL realUrl = new URL(urlNameString);
			//System.out.println(urlNameString);
			URLConnection conn = realUrl.openConnection();

			conn.setRequestProperty("accept", "*/*");
			conn.setRequestProperty("connection", "Keep-Alive");
			conn.setRequestProperty("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1;SV1)");
			conn.setRequestProperty("Content-Type", "application/json");

			conn.setDoOutput(true);
			conn.setDoInput(true);

			out = new PrintWriter(conn.getOutputStream());

			out.print(params);

			out.flush();

			in = new BufferedReader(new InputStreamReader(conn.getInputStream()));
			String line;
			while ((line = in.readLine()) != null) {
				result += line;
			}
		} catch (Exception e) {
			//System.out.println("There is something wrong in POST request!" + e);
			return e.toString();
		} finally {
			try {
				if (out != null) {
					out.close();
				}
				if (in != null) {
					in.close();
				}
			} catch (IOException ex) {
				//ex.printStackTrace();
			}
		}
		return result;
	}

}

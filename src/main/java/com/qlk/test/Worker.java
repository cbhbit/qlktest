package com.qlk.test;

import java.util.Map;

public class Worker {
	private String filePath="C:\\Users\\cbhbit\\Desktop\\ODC.xlsx";;
	
	Excel excel=new Excel();
	Map<Integer, Map<Integer, Object>> content1=excel.readExcelXLSXContentBySheet(filePath, 0);
	Map<Integer, Map<Integer, Object>> content2=excel.readExcelXLSXContentBySheet(filePath, 1);
	
	HTTPRequest request=new HTTPRequest();
	
	Worker(){

	}
	public String[] mapToString(Map<Integer, Map<Integer, Object>> content,int rowNumber){
		String[] row=new String[content.get(rowNumber).size()];
		for(int i=0;i<content.get(rowNumber).size();i++){
			row[i]=String.valueOf(content.get(rowNumber).get(i));
		}
		return row;
	}
	public boolean isEqual(Map<Integer, Map<Integer, Object>> content,int rowNumber){
		if(mapToString(content,rowNumber)[7].equals(mapToString(content,rowNumber)[6]))
    		return true;
    	else
    		return false;
	}
	public boolean isEqual(Map<Integer, Map<Integer, Object>> content,int rowNumber,int modelNumber){
		String s=mapToString(content,rowNumber)[6];
		if(mapToString(content,rowNumber)[7].contains(s))
    		return true;
    	else
    		return false;
	}
	
	public void work(){
		String host=mapToString(content1,1)[0];
		for(int i=1;i<content2.size();i++){
			String url = host+mapToString(content2,i)[0]+mapToString(content2,i)[1];
			String params = mapToString(content2,i)[2];
			String method = mapToString(content2,i)[5];
			excel.writeExcelXLSXByCell(filePath, 1, i, 7, request.sendrequest(url,params,method));
		}
	}

}

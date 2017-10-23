package com.qlk.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;  
import java.util.Date;  
import java.util.HashMap;  
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;  
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
import org.slf4j.Logger;  
import org.slf4j.LoggerFactory;

public class Excel{
	private Logger logger = LoggerFactory.getLogger(Excel.class);
	
    public String[] readExcelXLSTitleBySheet(String filePath,int sheetNumber){
    	FileInputStream fileIn=null;
    	HSSFWorkbook wb=null;
		try {
			fileIn = new FileInputStream(filePath);
			wb=new HSSFWorkbook(fileIn);
	    	HSSFSheet sheet=wb.getSheetAt(sheetNumber);
	    	HSSFRow row=sheet.getRow(0);

	    	int colNum = row.getPhysicalNumberOfCells();
	    	String[] title=new String[colNum];
	    	for (int i = 0; i < colNum; i++) {
	        	if(row.getCell(i)==null)
	        		title[i]="";
	        	else
	        		title[i] = row.getCell(i).getStringCellValue();
	        }
	    	return title;
		} catch (FileNotFoundException e) {
			logger.error("FileNotFound",e);
		} catch (IOException e) {
			logger.error("IOException",e);
		}finally{
			try {
				wb.close();
				fileIn.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
    	return null;   	
    }
    public Map<Integer, Map<Integer,Object>> readExcelXLSContentBySheet(String filePath,int sheetNumber){  
    	FileInputStream fileIn=null;
    	HSSFWorkbook wb=null;
		try {
			fileIn = new FileInputStream(filePath);
			wb=new HSSFWorkbook(fileIn);
	    	HSSFSheet sheet=wb.getSheetAt(sheetNumber);
	    	HSSFRow row=sheet.getRow(0);
	    	
	    	Map<Integer, Map<Integer,Object>> content = new HashMap<Integer, Map<Integer,Object>>();    
	        // 得到总行数  
	        int rowNum = sheet.getLastRowNum();  
	        row = sheet.getRow(0);  
	        int colNum = row.getPhysicalNumberOfCells();  
	        // 正文内容应该从第二行开始,第一行为表头的标题  
	        for (int i = 1; i <= rowNum; i++) {  
	            row = sheet.getRow(i);  
	            int j = 0;  
	            Map<Integer,Object> cellValue = new HashMap<Integer, Object>();  
	            while (j < colNum) {  
	                Object obj = getCellFormatValue(row.getCell(j));  
	                cellValue.put(j, obj);  
	                j++;  
	            }  
	            content.put(i, cellValue);  
	        }  
	        return content;
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			try {
				wb.close();
				fileIn.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		return null;
    }
    public void writeExcelXLSByCell(String filePath,int sheetNumber,int rowNumber,int cellNumber,String value){
    	FileInputStream fileIn=null;
    	FileOutputStream fileOut=null;
    	HSSFWorkbook wb=null;
		try {
			fileIn = new FileInputStream(filePath);
			fileOut=new FileOutputStream(filePath);
	    	
	    	wb=new HSSFWorkbook(fileIn);
	    	HSSFSheet sheet=wb.getSheetAt(sheetNumber);
	    	HSSFRow row=sheet.getRow(rowNumber);
	    	HSSFCell cell=row.createCell(cellNumber);
			cell.setCellValue(value);
			
			wb.write(fileOut);
			fileOut.flush();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			try {
				wb.close();
				fileIn.close();
				fileOut.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}		
    }
    public String[] readExcelXLSXTitleBySheet(String filePath,int sheetNumber){
    	FileInputStream fileIn=null;
    	XSSFWorkbook wb=null;
		try {
			fileIn = new FileInputStream(filePath);
			wb=new XSSFWorkbook(fileIn);
	    	XSSFSheet sheet=wb.getSheetAt(sheetNumber);
	    	XSSFRow row=sheet.getRow(0);

	    	int colNum = row.getPhysicalNumberOfCells();
	    	String[] title=new String[colNum];
	    	for (int i = 0; i < colNum; i++) {
	        	if(row.getCell(i)==null)
	        		title[i]="";
	        	else
	        		title[i] = row.getCell(i).getStringCellValue();
	        }
	    	return title;
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			try {
				wb.close();
				fileIn.close();				
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}			
		}
    	return null;    	
    }
    public Map<Integer, Map<Integer,Object>> readExcelXLSXContentBySheet(String filePath,int sheetNumber){  
    	FileInputStream fileIn=null;
    	XSSFWorkbook wb=null;
		try {
			fileIn = new FileInputStream(filePath);
			wb=new XSSFWorkbook(fileIn);
	    	XSSFSheet sheet=wb.getSheetAt(sheetNumber);
	    	XSSFRow row=sheet.getRow(0);
	    	
	    	Map<Integer, Map<Integer,Object>> content = new HashMap<Integer, Map<Integer,Object>>();    
	        // 得到总行数  
	        int rowNum = sheet.getLastRowNum();  
	        row = sheet.getRow(0);  
	        int colNum = row.getPhysicalNumberOfCells();  
	        // 正文内容应该从第二行开始,第一行为表头的标题  
	        for (int i = 1; i <= rowNum; i++) {  
	            row = sheet.getRow(i);  
	            int j = 0;  
	            Map<Integer,Object> cellValue = new HashMap<Integer, Object>();  
	            while (j < colNum) {  
	                Object obj = getCellFormatValue(row.getCell(j));  
	                cellValue.put(j, obj);  
	                j++;  
	            }  
	            content.put(i, cellValue);  
	        }  
	        return content;
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			if(wb!=null)
			try {
				wb.close();	
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			if(fileIn!=null)
				try {
					fileIn.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
		}
    	return null;
    }
    public void writeExcelXLSXByCell(String filePath,int sheetNumber,int rowNumber,int cellNumber,String value){   	
    	FileInputStream fileIn=null;
    	XSSFWorkbook wb=null;
    	FileOutputStream fileOut=null;
    	try {
			fileIn = new FileInputStream(filePath);
			wb=new XSSFWorkbook(fileIn);
	    	XSSFSheet sheet=wb.getSheetAt(sheetNumber);
	    	XSSFRow row=sheet.getRow(rowNumber);
	    	XSSFCell cell=row.createCell(cellNumber);
	    	if(value.length()>32767)
	    		cell.setCellValue(value.substring(0, 32767));
	    	else
	    		cell.setCellValue(value);
			
			fileOut=new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			if(wb!=null)
			try {
				wb.close();				
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			if(fileIn!=null)
				try {
					fileIn.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if(fileOut!=null){
				try {
					fileOut.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}    	
    }
    public void writeExcelXLSXByCell(String filePath,int sheetNumber,int rowNumber,int cellNumber,String value,int c){
    	FileInputStream fileIn=null;
    	XSSFWorkbook wb=null;
    	FileOutputStream fileOut=null;
    	try {
			fileIn = new FileInputStream(filePath);
			wb=new XSSFWorkbook(fileIn);
	    	XSSFSheet sheet=wb.getSheetAt(sheetNumber);
	    	XSSFRow row=sheet.getRow(rowNumber);
	    	XSSFCell cell=row.createCell(cellNumber);
	    	
	    	//CellStyle cellStyle=cell.getCellStyle();
			//cellStyle.setAlignment(CellStyle.ALIGN_FILL);
			
	    	if(value.length()>32767){
	    		cell.setCellValue(value.substring(0, 32767));	    		
		        //cell.setCellStyle(cellStyle);
	    	}
	    	else{
	    		cell.setCellValue(value);}
			
			fileOut=new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			if(wb!=null)
			try {
				wb.close();				
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			if(fileIn!=null)
				try {
					fileIn.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if(fileOut!=null){
				try {
					fileOut.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
    }
    @SuppressWarnings("deprecation")
	private Object getCellFormatValue(Cell cell) {  
        Object cellvalue = "";  
        if (cell != null) {  
            // 判断当前Cell的Type  
            switch (cell.getCellType()) {  
            case Cell.CELL_TYPE_NUMERIC:// 如果当前Cell的Type为NUMERIC  
            case Cell.CELL_TYPE_FORMULA: {  
                // 判断当前的cell是否为Date  
                if (DateUtil.isCellDateFormatted(cell)) {  
                    // 如果是Date类型则，转化为Data格式  
                    // data格式是带时分秒的：2013-7-10 0:00:00  
                    // cellvalue = cell.getDateCellValue().toLocaleString();  
                    // data格式是不带带时分秒的：2013-7-10  
                    Date date = cell.getDateCellValue();  
                    cellvalue = date;  
                } else {// 如果是纯数字  
  
                    // 取得当前Cell的数值  
                    cellvalue = String.valueOf(cell.getNumericCellValue());  
                }  
                break;  
            }  
            case Cell.CELL_TYPE_STRING:// 如果当前Cell的Type为STRING  
                // 取得当前的Cell字符串  
                cellvalue = cell.getRichStringCellValue().getString();  
                break;  
            default:// 默认的Cell值  
                cellvalue = "";  
            }  
        } else {  
            cellvalue = "";  
        }  
        return cellvalue;  
    }    

}

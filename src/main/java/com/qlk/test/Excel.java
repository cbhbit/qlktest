package com.qlk.test;

import java.io.File;
import java.io.FileInputStream;  
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;  
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;  
import java.util.HashMap;  
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.DateUtil;  
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.ss.usermodel.Sheet;  
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
import org.slf4j.Logger;  
import org.slf4j.LoggerFactory;

public class Excel{
	private String filePath="C:\\Users\\cbhbit\\Desktop\\ODC.xlsx";
	private Logger logger = LoggerFactory.getLogger(Excel.class);  
    private Workbook wb=null;  
    private Sheet sheet=null;  
    private Row row=null;
    private Column col=null;
    private Cell cell=null;
    
    public Workbook getWorkBook(String filepath){
    	if(filepath==null){  
            return null;  
        }   	
        String ext = filepath.substring(filepath.lastIndexOf("."));  
        try {  
            InputStream is = new FileInputStream(filepath);  
            if(".xls".equals(ext)){  
                wb = new HSSFWorkbook(is);  
            }else if(".xlsx".equals(ext)){
                wb = new XSSFWorkbook(is);  
            }else{  
                wb=null;  
            }  
        } catch (FileNotFoundException e) {  
            logger.error("FileNotFoundException", e);  
        } catch (IOException e) {  
            logger.error("IOException", e);  
        }
    	return wb;
    }
    public int getRowNumber(int sheetNumber){
    	if(wb==null){  
            try {
				throw new Exception("Workbook对象为空！");
			} catch (Exception e) {
				return -1;
			}  
        }
    	int rowNumber=wb.getSheetAt(sheetNumber).getLastRowNum();
    	return rowNumber;
    }
    public int getColumNumber(int sheetNumber,int rowNumber){
    	if(wb==null){  
            try {
				throw new Exception("Workbook对象为空！");
			} catch (Exception e) {
				return -1;
			}  
        }
    	sheet = wb.getSheetAt(sheetNumber);  
        row = sheet.getRow(rowNumber); 
        int columNumber = row.getPhysicalNumberOfCells();
    	return columNumber;
    }
    
    public Excel(String filepath) {
    	this.filePath=filepath;
        wb=getWorkBook(filepath);
    }
    /** 
     * 读取Excel表格表头的内容 
     *  
     * @param InputStream 
     * @return String 表头内容的数组 
     * @author zijing 
     */  
    public String[] readExcelTitle(int sheetNumber,int rowNumber) throws Exception{  
        
        // 标题总列数  
        int colNum = getColumNumber(sheetNumber,rowNumber);  
        //System.out.println("colNum:" + colNum);  
        String[] title = new String[colNum];  
        for (int i = 0; i < colNum; i++) {
        	if(row.getCell(i)==null)
        		title[i]="";
        	else
        		title[i] = row.getCell(i).getStringCellValue();
        }  
        return title;  
    }
    /** 
     * 读取Excel数据内容 
     *  
     * @param InputStream 
     * @return Map 包含单元格数据内容的Map对象 
     * @author zijing 
     */  
    public Map<Integer, Map<Integer,Object>> readExcelContent() throws Exception{  
        if(wb==null){  
            throw new Exception("Workbook对象为空！");  
        }  
        Map<Integer, Map<Integer,Object>> content = new HashMap<Integer, Map<Integer,Object>>();  
          
        sheet = wb.getSheetAt(0);  
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
    }  
  
    /** 
     *  
     * 根据Cell类型设置数据 
     *  
     * @param cell 
     * @return 
     * @author zijing 
     */  
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
    
    public void writeExcel(int sheetNumber,int rowNumber,int colNumber,String value) throws IOException{
    	FileInputStream fis = new FileInputStream(filePath);
		wb=getWorkBook(filePath);
		sheet = wb.getSheetAt(sheetNumber);
		cell=sheet.getRow(rowNumber).getCell(colNumber);
		//cell = sheet.createRow(rowNumber).createCell(colNumber);
		cell.setCellValue(value);
		

		fis.close();// 关闭文件输入流
    	FileOutputStream fos = new FileOutputStream(filePath);
		wb.write(fos);
		fos.close();// 关闭文件输出流
    }  

}

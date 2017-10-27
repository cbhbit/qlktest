package com.qlk.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Excel {
	private Logger logger = LoggerFactory.getLogger(Excel.class);

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
					cell.setCellType(Cell.CELL_TYPE_STRING);
					// 取得当前Cell的数值
					cellvalue = cell.getStringCellValue();//String.valueOf(cell.getNumericCellValue());
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

	public List<Map<Integer, Map<Integer, Object>>> readExcelXLSContent(String filePath) {
		FileInputStream fileIn = null;
		HSSFWorkbook wb = null;
		List<Map<Integer, Map<Integer, Object>>> content = new ArrayList<Map<Integer, Map<Integer, Object>>>();
		try {
			fileIn = new FileInputStream(filePath);
			wb = new HSSFWorkbook(fileIn);

			for (int k = 0; k < wb.getNumberOfSheets(); k++) {
				HSSFSheet sheet = wb.getSheetAt(k);
				HSSFRow row = sheet.getRow(0);
				Map<Integer, Map<Integer, Object>> sheetMap = new HashMap<Integer, Map<Integer, Object>>();
				// 得到总行数
				int rowNum = sheet.getLastRowNum();
				int colNum;
				if (rowNum > 0)
					colNum = row.getPhysicalNumberOfCells();
				else
					colNum = 0;

				for (int i = 0; i <= rowNum; i++) {
					row = sheet.getRow(i);
					int j = 0;
					Map<Integer, Object> cellValue = new HashMap<Integer, Object>();
					while (j < colNum) {
						Object obj = getCellFormatValue(row.getCell(j));
						cellValue.put(j, obj);
						j++;
					}
					sheetMap.put(i, cellValue);
				}
				content.add(sheetMap);
			}
			return content;
		} catch (FileNotFoundException e) {
			logger.error("FileNotFound!", e);
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if (wb != null)
				try {
					wb.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if (fileIn != null)
				try {
					fileIn.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
		}
		return null;
	}

	public void writeExcelXLSByColumn(String[] value, String filePath, int sheetNumber, int columnNumber) {
		FileInputStream fileIn = null;
		HSSFWorkbook wb = null;
		FileOutputStream fileOut = null;
		try {
			fileIn = new FileInputStream(filePath);
			wb = new HSSFWorkbook(fileIn);
			HSSFSheet sheet = wb.getSheetAt(sheetNumber);
			for (int i = 0; i < value.length; i++) {
				HSSFRow row = sheet.getRow(i);
				HSSFCell cell = row.createCell(columnNumber);
				CellStyle cellStyle = wb.createCellStyle();
				cellStyle.setAlignment(HorizontalAlignment.FILL);
				cell.setCellStyle(cellStyle);

				if (value[i].length() > 32767)
					cell.setCellValue(value[i].substring(0, 32767));
				else
					cell.setCellValue(value[i]);
			}

			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if (wb != null)
				try {
					wb.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if (fileIn != null)
				try {
					fileIn.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if (fileOut != null) {
				try {
					fileOut.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
	}

	public void writeExcelXLSByColumnPattern(String[] value, String filePath, int sheetNumber, int columnNumber) {
		FileInputStream fileIn = null;
		HSSFWorkbook wb = null;
		FileOutputStream fileOut = null;
		try {
			fileIn = new FileInputStream(filePath);
			wb = new HSSFWorkbook(fileIn);
			HSSFSheet sheet = wb.getSheetAt(sheetNumber);
			for (int i = 0; i < value.length; i++) {
				HSSFRow row = sheet.getRow(i);
				HSSFCell cell = row.createCell(columnNumber);
				CellStyle cellStyle = wb.createCellStyle();
				if (value[i].equals("Pass")) {
					cellStyle.setFillForegroundColor(IndexedColors.GREEN.index);
					cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				} else if (value[i].equals("Fail")) {
					cellStyle.setFillForegroundColor(IndexedColors.RED.index);
					cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				}
				cell.setCellStyle(cellStyle);

				cell.setCellValue(value[i]);
			}

			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if (wb != null)
				try {
					wb.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if (fileIn != null)
				try {
					fileIn.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if (fileOut != null) {
				try {
					fileOut.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
	}

	public List<Map<Integer, Map<Integer, Object>>> readExcelXLSXContent(String filePath) {
		FileInputStream fileIn = null;
		XSSFWorkbook wb = null;
		List<Map<Integer, Map<Integer, Object>>> content = new ArrayList<Map<Integer, Map<Integer, Object>>>();
		try {
			fileIn = new FileInputStream(filePath);
			wb = new XSSFWorkbook(fileIn);

			for (int k = 0; k < wb.getNumberOfSheets(); k++) {
				XSSFSheet sheet = wb.getSheetAt(k);
				XSSFRow row = sheet.getRow(0);
				Map<Integer, Map<Integer, Object>> sheetMap = new HashMap<Integer, Map<Integer, Object>>();
				// 得到总行数
				int rowNum = sheet.getLastRowNum();
				int colNum;
				if (rowNum > 0)
					colNum = row.getPhysicalNumberOfCells();
				else
					colNum = 0;

				for (int i = 0; i <= rowNum; i++) {
					row = sheet.getRow(i);
					int j = 0;
					Map<Integer, Object> cellValue = new HashMap<Integer, Object>();
					while (j < colNum) {
						Object obj = getCellFormatValue(row.getCell(j));
						cellValue.put(j, obj);
						j++;
					}
					sheetMap.put(i, cellValue);
				}
				content.add(sheetMap);
			}
			return content;
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if (wb != null)
				try {
					wb.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if (fileIn != null)
				try {
					fileIn.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
		}
		return null;
	}

	public void writeExcelXLSXByColumn(String[] value, String filePath, int sheetNumber, int columnNumber) {
		FileInputStream fileIn = null;
		XSSFWorkbook wb = null;
		FileOutputStream fileOut = null;
		try {
			fileIn = new FileInputStream(filePath);
			wb = new XSSFWorkbook(fileIn);
			XSSFSheet sheet = wb.getSheetAt(sheetNumber);
			for (int i = 0; i < value.length; i++) {
				XSSFRow row = sheet.getRow(i);
				XSSFCell cell = row.createCell(columnNumber);
				CellStyle cellStyle = wb.createCellStyle();
				cellStyle.setAlignment(HorizontalAlignment.FILL);
				cell.setCellStyle(cellStyle);

				if (value[i].length() > 32767)
					cell.setCellValue(value[i].substring(0, 32767));
				else
					cell.setCellValue(value[i]);
			}

			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if (wb != null)
				try {
					wb.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if (fileIn != null)
				try {
					fileIn.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if (fileOut != null) {
				try {
					fileOut.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
	}

	public void writeExcelXLSXByColumnPattern(String[] value, String filePath, int sheetNumber, int columnNumber) {
		FileInputStream fileIn = null;
		XSSFWorkbook wb = null;
		FileOutputStream fileOut = null;
		try {
			fileIn = new FileInputStream(filePath);
			wb = new XSSFWorkbook(fileIn);
			XSSFSheet sheet = wb.getSheetAt(1);
			for (int i = 0; i < value.length; i++) {
				XSSFRow row = sheet.getRow(i);
				XSSFCell cell = row.createCell(columnNumber);

				CellStyle cellStyle = wb.createCellStyle();
				cell.setCellValue(value[i]);
				if (cell.getStringCellValue().equals("Pass")) {
					cellStyle.setFillForegroundColor(IndexedColors.GREEN.index);
					cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					cell.setCellStyle(cellStyle);
				} else if (value[i].equals("Fail")) {
					cellStyle.setFillForegroundColor(IndexedColors.RED.index);
					cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				}
				cell.setCellStyle(cellStyle);
			}

			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if (wb != null)
				try {
					wb.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if (fileIn != null)
				try {
					fileIn.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if (fileOut != null) {
				try {
					fileOut.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
	}
	
	public void writeExcelXLSXResultPass(String result,String value, String filePath, int sheetNumber,int rowNumber, int columnNumber){
		FileInputStream fileIn = null;
		XSSFWorkbook wb = null;
		FileOutputStream fileOut = null;
		try {
			fileIn = new FileInputStream(filePath);
			wb = new XSSFWorkbook(fileIn);
			XSSFSheet sheet = wb.getSheetAt(1);

				XSSFRow row = sheet.getRow(rowNumber);
				XSSFCell cell = row.createCell(columnNumber);
				cell.setCellValue(result);
				cell = row.createCell(columnNumber+1);
				CellStyle cellStyle = wb.createCellStyle();
				cell.setCellValue(value);
				if (cell.getStringCellValue().equals("Pass")) {
					cellStyle.setFillForegroundColor(IndexedColors.GREEN.index);
					cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					cell.setCellStyle(cellStyle);
				} else if (value.equals("Fail")) {
					cellStyle.setFillForegroundColor(IndexedColors.RED.index);
					cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				}
				cell.setCellStyle(cellStyle);
			//}

			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if (wb != null)
				try {
					wb.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if (fileIn != null)
				try {
					fileIn.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if (fileOut != null) {
				try {
					fileOut.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
	}
	
	public void writeExcelXLSResultPass(String result,String value, String filePath, int sheetNumber,int rowNumber, int columnNumber){
		FileInputStream fileIn = null;
		HSSFWorkbook wb = null;
		FileOutputStream fileOut = null;
		try {
			fileIn = new FileInputStream(filePath);
			wb = new HSSFWorkbook(fileIn);
			HSSFSheet sheet = wb.getSheetAt(1);

				HSSFRow row = sheet.getRow(rowNumber);
				HSSFCell cell = row.createCell(columnNumber);
				cell.setCellValue(result);
				cell = row.createCell(columnNumber+1);
				CellStyle cellStyle = wb.createCellStyle();
				cell.setCellValue(value);
				if (cell.getStringCellValue().equals("Pass")) {
					cellStyle.setFillForegroundColor(IndexedColors.GREEN.index);
					cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					cell.setCellStyle(cellStyle);
				} else if (value.equals("Fail")) {
					cellStyle.setFillForegroundColor(IndexedColors.RED.index);
					cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				}
				cell.setCellStyle(cellStyle);
			//}

			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if (wb != null)
				try {
					wb.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if (fileIn != null)
				try {
					fileIn.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if (fileOut != null) {
				try {
					fileOut.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
	}
}

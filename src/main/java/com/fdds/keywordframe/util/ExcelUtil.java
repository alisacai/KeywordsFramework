package com.fdds.keywordframe.util;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import com.fdds.keywordframe.configuration.Constant;
import com.fdds.keywordframe.testScript.TestSuiteByExcel;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
	public static XSSFWorkbook workbook;
	public static XSSFSheet sheet;
	public static XSSFRow row;
	public static XSSFCell cell;

	public static void setExcelFile(String filepath) {

		FileInputStream input;
		try {
			input = new FileInputStream(filepath);
			workbook = new XSSFWorkbook(input);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			TestSuiteByExcel.testResult = false;
			Log.info("Excel路径设定失败");
			e.printStackTrace();
		}
	}

	public static void setExcelFile(String filepath, String sheetName) {

		FileInputStream input;
		try {
			input = new FileInputStream(filepath);
			workbook = new XSSFWorkbook(input);
			sheet = workbook.getSheet(sheetName);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			TestSuiteByExcel.testResult = false;
			Log.info("Excel路径设定失败");
			e.printStackTrace();
		}
	}

	public static String getCellData(String sheetName, int rowNum, int colNum) {
		try {
			sheet = workbook.getSheet(sheetName);
			row = sheet.getRow(rowNum);
			cell = row.getCell(colNum);
			//CellType.STRING
			String cellData = cell.getCellType() == CellType.STRING?
					cell.getStringCellValue() + "" :
					String.valueOf(Math.round(cell.getNumericCellValue()));
			return cellData;
		} catch (Exception e) {
			e.printStackTrace();
			return "";
		}
	}

	public static int getRowCount(String sheetName) {
		sheet = workbook.getSheet(sheetName);
		int lastRowNum = sheet.getLastRowNum();
		return lastRowNum;
	}

	public static void setCellData(String sheetName, int rowNum, int colNum, String result) throws IOException {
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(rowNum);
		cell = row.getCell(colNum, Row.RETURN_BLANK_AS_NULL);
		if (cell == null) {
			cell = row.createCell(colNum);
			cell.setCellValue(result);
		}
		cell.setCellValue(result);
		FileOutputStream output = new FileOutputStream(Constant.Path_ExcelFile);
		workbook.write(output);
		output.flush();
		output.close();
	}

	public static int getFirstRowContainsTestCaseID(String sheetName, String TestCaseID, int colNum) {
		int rowNumber;
		try {
			sheet = workbook.getSheet(sheetName);
			int rowCount = ExcelUtil.getRowCount(sheetName);
			for (rowNumber = 1; rowNumber <= rowCount; rowNumber++) {
				row = sheet.getRow(rowNumber);
				String testCaseName = getCellData(sheetName, rowNumber, colNum);
				if (testCaseName.equals(TestCaseID)) {
					break;
				}
			}
			return rowNumber;
		} catch (Exception e) {
			TestSuiteByExcel.testResult = false;
			e.printStackTrace();
			return 0;
		}
	}

	public static int getLastRowContainsTestCaseID(String sheetName, String testCaseID, int colNum) {
		try {
			sheet = workbook.getSheet(sheetName);
			int rowCount = ExcelUtil.getRowCount(sheetName);
			int testStep = ExcelUtil.getFirstRowContainsTestCaseID(sheetName, testCaseID, colNum);
			for (; testStep < rowCount; testStep++) {
				String testCaseName = ExcelUtil.getCellData(sheetName, testStep, colNum);
				if (!testCaseName.equals(testCaseID)) {
					break;
				}
			}
			return testStep - 1;
		} catch (Exception e) {
			e.printStackTrace();
			return 0;
		}
	}
}
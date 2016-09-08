package com.amada.love.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiExcelTest {

	private Workbook wb;

	private Sheet sheet;
	private Row row;

	public PoiExcelTest(String filepath) {
		if (filepath == null) {
			return;
		}
		String ext = filepath.substring(filepath.lastIndexOf("."));
		try {
			InputStream is = new FileInputStream(filepath);
			if (".xls".equals(ext)) {
				wb = new HSSFWorkbook(is);
			} else if (".xlsx".equals(ext)) {
				wb = new XSSFWorkbook(is);
			} else {
				wb = null;
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public void readExcelContent() {
		if (wb == null) {
			throw new RuntimeException("Workbook对象为空！");
		}
		sheet = wb.getSheetAt(0);
		// 得到总行数,从0 开始
		int rowNum = sheet.getLastRowNum();

		// row = sheet.getRow(0);
		// 得到一行的列数
		// int colNum = row.getLastCellNum();

		// System.out.println("rowNum=" + rowNum + ",colNum=" +
		// colNum+",row[0][8]="+row.getCell(9).getStringCellValue());

		for (int r = 0; r <= rowNum; r++) {
			row = sheet.getRow(r);
			int colNum = row.getLastCellNum();
			for (int c = 0; c <= colNum; c++) {
				System.out.print(getCellFormatValue(row.getCell(c))+"  ");
			}
			System.out.println();
		}

	}

	/**
	 * 
	 * 根据Cell类型设置数据
	 * 
	 * @param cell
	 * @return
	 * @author zengwendong
	 */
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

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		String filepath = "files\\美国Top10000 Publisher添加需求新.xlsx";
		PoiExcelTest poiTest = new PoiExcelTest(filepath);
		poiTest.readExcelContent();
	}

}

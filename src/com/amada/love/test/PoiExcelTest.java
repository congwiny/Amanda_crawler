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
			throw new RuntimeException("Workbook����Ϊ�գ�");
		}
		sheet = wb.getSheetAt(0);
		// �õ�������,��0 ��ʼ
		int rowNum = sheet.getLastRowNum();

		// row = sheet.getRow(0);
		// �õ�һ�е�����
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
	 * ����Cell������������
	 * 
	 * @param cell
	 * @return
	 * @author zengwendong
	 */
	private Object getCellFormatValue(Cell cell) {
		Object cellvalue = "";
		if (cell != null) {
			// �жϵ�ǰCell��Type
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:// �����ǰCell��TypeΪNUMERIC
			case Cell.CELL_TYPE_FORMULA: {
				// �жϵ�ǰ��cell�Ƿ�ΪDate
				if (DateUtil.isCellDateFormatted(cell)) {
					// �����Date������ת��ΪData��ʽ
					// data��ʽ�Ǵ�ʱ����ģ�2013-7-10 0:00:00
					// cellvalue = cell.getDateCellValue().toLocaleString();
					// data��ʽ�ǲ�����ʱ����ģ�2013-7-10
					Date date = cell.getDateCellValue();
					cellvalue = date;
				} else {// ����Ǵ�����

					// ȡ�õ�ǰCell����ֵ
					cellvalue = String.valueOf(cell.getNumericCellValue());
				}
				break;
			}
			case Cell.CELL_TYPE_STRING:// �����ǰCell��TypeΪSTRING
				// ȡ�õ�ǰ��Cell�ַ���
				cellvalue = cell.getRichStringCellValue().getString();
				break;
			default:// Ĭ�ϵ�Cellֵ
				cellvalue = "";
			}
		} else {
			cellvalue = "";
		}
		return cellvalue;
	}

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		String filepath = "files\\����Top10000 Publisher���������.xlsx";
		PoiExcelTest poiTest = new PoiExcelTest(filepath);
		poiTest.readExcelContent();
	}

}

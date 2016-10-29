package com.tnz.util;

import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;

public class ReadExcelTest {

	public static void main(String[] args) {
		ExcelUtil eu = new ExcelUtil();
		eu.setPrintMsg(true);
		eu.setStartReadPos(0);// 从第一行开始读取
		eu.setOnlyReadOneSheet(true);
		String src_xlspath = "E:/java_workspace/POI-Utils/test.xls";
		List<Row> rowList;
		try {
			rowList = eu.readExcel(src_xlspath);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}

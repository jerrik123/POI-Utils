package com.tnz.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * 根据模板 在设置值
 * @author YangJie
 *
 */
public class WriteExcelTest {

	public static void main(String[] args) throws IOException {
		ExcelUtil eu = new ExcelUtil();
		eu.setPrintMsg(true);
		eu.setStartReadPos(0);// 从第一行开始读取

		File fi = new File("template.xls");
		POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(fi));
		// 读取excel模板
		HSSFWorkbook wb = new HSSFWorkbook(fs);
		// 读取了模板内所有sheet内容
		HSSFSheet sheet = wb.getSheetAt(0);
		Row firstRow = sheet.createRow(1);
		
		Cell cell = firstRow.createCell(0);
		cell.setCellValue("201634234234234");
		
		Cell cell1 = firstRow.createCell(1);
		cell1.setCellValue(new Date());
		
		Cell cell2 = firstRow.createCell(2);
		cell2.setCellValue("jerrik");
		
		Cell cell3 = firstRow.createCell(3);
		cell3.setCellValue("深圳高新园");
		
		Cell cell4 = firstRow.createCell(4);
		cell4.setCellValue("3500.00");
		// 修改模板内容导出新模板
		FileOutputStream out = new FileOutputStream("D:/export.xls");
		wb.write(out);
		out.close();
	}

}

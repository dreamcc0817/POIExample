package com.dreamcc.demo.controller;

import com.dreamcc.demo.entity.IOCEntity;
import com.dreamcc.demo.entity.MyRowMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import java.io.*;
import java.util.List;

/**
 * @Title: demo
 * @Package: com.dreamcc.demo.controller
 * @Description:
 * @Author: dreamcc
 * @Date: 2018-10-26 20:56
 * @Version: V1.0
 */
@Controller
@RequestMapping("/demo")
public class DemoController {
	@Autowired
	JdbcTemplate jdbcTemplate;

	@RequestMapping("/create")
	public void creatFolder(){
		String sql = "select * from lzioc";

		/**
		 *  把从数据库查出的信息放到一个列表里面，循环列表新建文件夹和文件，
		 *  将列信息并且根据'，'进行分割
		 */
		List<IOCEntity> list = jdbcTemplate.query(sql,new MyRowMapper());

		for (int i = 0; i < list.size(); i++) {
			//新建文件夹
			File dirfile = new File("D:\\lzioc\\"+list.get(i).getDept());
			if(!dirfile.exists()){
				dirfile.mkdir();
			}
			//新建excle文件
			Workbook workbook = new XSSFWorkbook();
			//格式
			CellStyle style = workbook.createCellStyle();
			Font font = workbook.createFont();
			font.setBold(true);
			font.setFontHeightInPoints((short) 14);
			style.setFont(font);
			//style.setAlignment(HorizontalAlignment.CENTER);

			CellStyle style1 = workbook.createCellStyle();
			Font font1 = workbook.createFont();
			font1.setColor(Font.COLOR_RED);
			style1.setFont(font1);
			style1.setVerticalAlignment(VerticalAlignment.CENTER);
			String[] infos = list.get(i).getInfo().split(",");
			//新建sheet
			Sheet sheet = workbook.createSheet("Sheet1");
			sheet.autoSizeColumn(1,true);
			//新建行
			sheet.addMergedRegion(new CellRangeAddress(1,1,2,infos.length+2));
			sheet.addMergedRegion(new CellRangeAddress(2,2,2,infos.length+2));
//			sheet.addMergedRegion(new CellRangeAddress(20,20,2,infos.length+2));
//			sheet.addMergedRegion(new CellRangeAddress(21,21,2,infos.length+2));
//			sheet.addMergedRegion(new CellRangeAddress(22,22,2,infos.length+2));
//			sheet.addMergedRegion(new CellRangeAddress(23,23,2,infos.length+2));
//			sheet.addMergedRegion(new CellRangeAddress(24,24,2,infos.length+2));

			Row row = sheet.createRow(1);
			Cell cell = row.createCell(2);
			cell.setCellValue(list.get(i).getResource());
			cell.setCellStyle(style);


			Row row1 = sheet.createRow(3);
			//循环列信息，建立列
			row1.createCell(0).setCellValue("");
			row1.createCell(1).setCellValue("");
			for (int j = 0; j < infos.length; j++) {
				row1.createCell(j+2).setCellValue(infos[j]);
			}
//
//			Row row2 = sheet.createRow(20);
//			Cell cell1 = row2.createCell(2);
//			cell1.setCellValue(new XSSFRichTextString("注：\r\n" +
//					"1、请勿更改或删除表头；\r\n" +
//					"2、请提供2014-2018年的数据；\r\n" +
//					"3、请务必保证数据的准确性；\r\n" +
//					"4、如表单内容存在客观因素无法填写，请说明原因。"));
//			cell1.setCellStyle(style1);

			Row row11 = sheet.createRow(20);
			Cell cell11 = row11.createCell(2);
			cell11.setCellValue(new XSSFRichTextString("注："));
			cell11.setCellStyle(style1);

			Row row12 = sheet.createRow(21);
			Cell cell12 = row12.createCell(2);
			cell12.setCellValue(new XSSFRichTextString(
					"1、请勿更改或删除表头；"));
			cell12.setCellStyle(style1);

			Row row13 = sheet.createRow(22);
			Cell cell13 = row13.createCell(2);
			cell13.setCellValue(new XSSFRichTextString(
					"2、请提供2014-2018年的数据；" ));
			cell13.setCellStyle(style1);

			Row row14 = sheet.createRow(23);
			Cell cell14 = row14.createCell(2);
			cell14.setCellValue(new XSSFRichTextString(
					"3、请务必保证数据的准确性；"));
			cell14.setCellStyle(style1);

			Row row15 = sheet.createRow(24);
			Cell cell15 = row15.createCell(2);
			cell15.setCellValue(new XSSFRichTextString(
					"4、如表单内容存在客观因素无法填写，请说明原因。"));
			cell15.setCellStyle(style1);

			//保存为excle文件
			try (OutputStream fileOut = new FileOutputStream("D:\\lzioc\\"+list.get(i).getDept()+"\\"+list.get(i).getResource()+".xlsx")) {
				workbook.write(fileOut);
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

}

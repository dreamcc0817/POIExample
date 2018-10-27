package com.dreamcc.demo.controller;

import com.dreamcc.demo.entity.IOCEntity;
import com.dreamcc.demo.entity.MyRowMapper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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

		List<IOCEntity> list = jdbcTemplate.query(sql,new MyRowMapper());

		for (int i = 0; i < list.size(); i++) {
			File dirfile = new File("D:\\lzioc\\"+list.get(i).getDept());
			if(!dirfile.exists()){
				dirfile.mkdir();
			}

			Workbook workbook = new XSSFWorkbook();
			String[] infos = list.get(i).getInfo().split(",");
			Sheet sheet = workbook.createSheet("Sheet1");
			Row row = sheet.createRow(0);
			for (int j = 0; j < infos.length; j++) {
				row.createCell(j).setCellValue(infos[j]);
			}
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

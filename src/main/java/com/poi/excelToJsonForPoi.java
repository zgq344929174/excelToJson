package com.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.alibaba.fastjson.JSON;

public class excelToJsonForPoi {

	@SuppressWarnings("resource")
	public static void main(String[] args) {
		// d盘下的xls目录
		File dir = new File("D:\\xls");
		// 获取xls、xlsx目录下的所有文件列表
		File[] list = dir.listFiles();

		for (File file : list) {
			String fileName = file.getName();
			String suffix = fileName.substring(fileName.lastIndexOf("."));
			File dest = new File("D:\\ToJson\\json_" + fileName + ".json");
			if (dest.exists()) {
				dest.delete();
			}
			try {
				FileInputStream fis = new FileInputStream(file);
				Workbook wk = null;
				if (".xls".equals(suffix)) {
					wk = new HSSFWorkbook(fis);
				} else if (".xlsx".equals(suffix)) {
					wk = new XSSFWorkbook(fis);
				} else {
					continue;
				}
				Sheet sheetAt = wk.getSheetAt(0);
				System.out.println(file.getName());
				Row header = sheetAt.getRow(0);
				for (int i = 1; i < sheetAt.getLastRowNum()+1; i++) {
					Map<String, String> hashMap = new LinkedHashMap<String, String>();
					for (int j = 0; j < sheetAt.getRow(i).getLastCellNum(); j++) {
						hashMap.put(header.getCell(j).getStringCellValue(),
								sheetAt.getRow(i).getCell(j).getStringCellValue());
					}
					String json = JSON.toJSONString(hashMap);
					FileUtils.writeStringToFile(dest, json + "\n", "UTF-8", true);
				}
			} catch (IOException e) {
				e.printStackTrace();
			}

		}
	}
}
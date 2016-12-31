package com.excelToJson;

import java.io.File;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.commons.io.FileUtils;

import com.alibaba.fastjson.JSONObject;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;  
  
public class ExcelToJSON {  
  
    public static void main(String[] args) {  
        // d���µ�xlsĿ¼  
        File dir = new File("D:\\xls");  
        // �����ļ�����  
        FilenameFilter searchSuffix = new FilenameFilter() { 
        	
            public boolean accept(File dir, String name) {  
                return name.endsWith(".xls");  
            }  
        };  
        // ��ȡxlsĿ¼�µ������ļ��б�  
        File[] list = dir.listFiles();  
        for (File file : list) {  
            File dest = new File("D:\\ToJson\\json_"  
                    + file.getName().substring(0,  
                            file.getName().lastIndexOf(".")) + ".json");  
            if (dest.exists()) {  
                dest.delete();  
            }  
            // ��ȡxlsĿ¼�µ��ļ���  
            String fileName = file.getName();  
            // ��ȡ�ļ���׺  
            String suffix = fileName.substring(fileName.lastIndexOf("."));  
            // ���������xlsΪ��β���ļ�����  
            if (!searchSuffix.accept(file, suffix)) {  
                continue;  
            }  
            try {  
                Workbook wb = Workbook.getWorkbook(file); // ���ļ����л�ȡExcel����������WorkBook��  
                Sheet sheet = wb.getSheet(0); // �ӹ�������ȡ��ҳ��Sheet��  
                Cell[] header = sheet.getRow(0);  
                System.out.println(file.getName());
                System.out.println("sheet.getRows()="+sheet.getRows());
                System.out.println("sheet.getColumns()="+sheet.getColumns());
                for (int i = 1; i < sheet.getRows(); i++) { // ѭ����ӡExcel���е�����  
//                    Map hashMap = new HashMap(); 
                	Map hashMap = new LinkedHashMap();
                    for (int j = 0; j < sheet.getColumns(); j++) {  
                        Cell cell = sheet.getCell(j, i);  
                        hashMap.put(header[j].getContents(), cell.getContents());  
                    }  
                    // ���json�ַ�������������Ҫ�ģ�ʵ��Ӧ���п���ֱ�ӷ��ظ��ַ���  
                    String json = JSONObject.toJSONString(hashMap);  
                    // ��ת�����json�ַ���д���ļ�����  
                    FileUtils.writeStringToFile(dest, json + "\n", "UTF-8",true);  
                }  
            } catch (BiffException e) {  
                e.printStackTrace();  
            } catch (IOException e) {  
                e.printStackTrace();  
            }  
  
        }  
    }  
}  
package com.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by z00390414 on 2017/7/3.
 *
 * @version 1.0
 */
public class ExcelImport {
    public List<String> loadScoreInfo(String xlsPath) throws IOException {
        List temp = new ArrayList();

        FileInputStream fileIn = new FileInputStream(xlsPath);
        Workbook wb0 = createWorkbook(fileIn, xlsPath);
//根据指定的文件输入流导入Excel从而产生Workbook对象
//        Workbook wb0 = new XSSFWorkbook(fileIn);
//获取Excel文档中的第一个表单
        Sheet sht0 = wb0.getSheetAt(0);
//对Sheet中的每一行进行迭代
        for (Row r : sht0) {
            //如果当前行的行号（从0开始）未达到2（第三行）则从新循环
//            if (r.getRowNum() < 1) {
//                continue;
//            }
//创建实体类
            String info;
//取出当前行第1个单元格数据，并封装在info实体stuName属性上
//            info.setStuName(r.getCell(0).getStringCellValue());
//            info.setClassName(r.getCell(1).getStringCellValue());
//            info.setRscore(r.getCell(2).getNumericCellValue());
//            info.setLscore(r.getCell(3).getNumericCellValue());
//            temp.add(info);
            info = r.getCell(0).getStringCellValue();
            System.out.println(info);
            temp.add(info);
        }
        fileIn.close();
        return temp;
    }

    public static void main(String[] arg) {

        ExcelImport excelImport = new ExcelImport();
        try {
            excelImport.loadScoreInfo("test.xlsx");
            excelImport.loadScoreInfo("test.xls");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public Workbook createWorkbook(InputStream is, String excelFileName) throws IOException {
        if (excelFileName.endsWith(".xls")) {
            return new HSSFWorkbook(is);
        } else if (excelFileName.endsWith(".xlsx")) {
            return new XSSFWorkbook(is);
        }
        return null;
    }
}

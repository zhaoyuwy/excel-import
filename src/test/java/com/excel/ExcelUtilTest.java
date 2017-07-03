package com.excel;

import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.List;

import static org.junit.Assert.*;

/**
 * Created by z00390414 on 2017/7/3.
 *
 * @version 1.0
 */
public class ExcelUtilTest {
    @Test
    public void isIE() throws Exception {
        System.out.println("This is a test");
    }
    @Test
    public void Util(){

        String xlsPath = "test.xlsx";
        ExcelUtil excelUtil=new ExcelUtil();
        FileInputStream fileIn = null;
        try {
            fileIn = new FileInputStream(xlsPath);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        Object vo = new Object();
        List<Object> objects = excelUtil.importDataFromExcel(vo,fileIn,xlsPath);

        System.out.println(objects);
    }

}
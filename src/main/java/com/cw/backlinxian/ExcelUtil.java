package com.cw.backlinxian;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ExcelUtil {

    public static Workbook readExcel(File file) {
        Workbook wb = null;
        String extString = file.getPath().substring(file.getPath().lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(file.getPath());
            if(".xls".equals(extString)){
                return wb = new HSSFWorkbook(is);
//                return wb = WorkbookFactory.create(is);
            }else if(".xlsx".equals(extString)){
                return wb = new XSSFWorkbook(is);
            }else{
                return wb = null;
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }
}

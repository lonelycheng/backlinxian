package com.cw.backlinxian;

import com.aspose.cells.*;

import java.io.File;
import java.io.FileOutputStream;

public class AsposeUtil {

    public static void convertToImage(String excelPath) {
        Workbook book = null;
        try {
            book = new Workbook("C:\\Users\\99543\\Desktop\\tmp\\1报表生成\\" + excelPath);
            // Get the first worksheet
            //Worksheet sheet = book.getWorksheets().get(0);
            Worksheet sheet = book.getWorksheets().get(0);
            sheet.getPageSetup().setLeftMargin(-20);
            sheet.getPageSetup().setRightMargin(0);
            sheet.getPageSetup().setBottomMargin(0);
            sheet.getPageSetup().setTopMargin(0);

            // Define ImageOrPrintOptions
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            // Specify the image format
            imgOptions.setImageFormat(ImageFormat.getJpeg());
            imgOptions.setCellAutoFit(true);
            imgOptions.setOnePagePerSheet(true);
            //imgOptions.setDesiredSize(1000,800);
            // Render the sheet with respect to specified image/print options
            SheetRender render = new SheetRender(sheet, imgOptions);

            // Render the image for the sheet
            //render.toImage(0, dataDir + "SheetImage.jpg");
            render.toImage(0, "C:\\Users\\99543\\Desktop\\tmp\\1报表生成\\" + excelPath.replace("xlsx", "jpg"));
        } catch (Exception e) {
            e.printStackTrace();
        }

    }


    public static void convertToPdf(String excelPath) {
        try {
            Workbook wb = new Workbook(excelPath);// 原始excel路径

            FileOutputStream fileOS = new FileOutputStream(excelPath.replace("xlsx", "pdf"));
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.setOnePagePerSheet(true);
            int[] autoDrawSheets={3};
            //当excel中对应的sheet页宽度太大时，在PDF中会拆断并分页。此处等比缩放。
            autoDraw(wb,autoDrawSheets);
            int[] showSheets={0};
            //隐藏workbook中不需要的sheet页。
            printSheetPage(wb,showSheets);
            wb.save(fileOS, pdfSaveOptions);
            fileOS.flush();
            fileOS.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    /**
     * 设置打印的sheet 自动拉伸比例
     * @param wb
     * @param page 自动拉伸的页的sheet数组
     */
    public static void autoDraw(Workbook wb,int[] page){
        if(null!=page&&page.length>0){
            for (int i = 0; i < page.length; i++) {
                wb.getWorksheets().get(i).getHorizontalPageBreaks().clear();
                wb.getWorksheets().get(i).getVerticalPageBreaks().clear();
            }
        }
    }

    /**
     * 隐藏workbook中不需要的sheet页。
     * @param wb
     * @param page 显示页的sheet数组
     */
    public static void printSheetPage(Workbook wb,int[] page){
        for (int i= 1; i < wb.getWorksheets().getCount(); i++)  {
            wb.getWorksheets().get(i).setVisible(false);
        }
        if(null==page||page.length==0){
            wb.getWorksheets().get(0).setVisible(true);
        }else{
            for (int i = 0; i < page.length; i++) {
                wb.getWorksheets().get(i).setVisible(true);
            }
        }
    }

    public static void main(String[] args) {
        convertToPdf("C:\\Users\\99543\\Desktop\\tmp\\0报表模板\\6临泉镇解码模板带签章.xlsx");
    }

}
package com.cw.backlinxian;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ConvertToImage {

    public static void ConvertToImage(String excelPath) {
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

//    public static void main(String[] args) {
//        ConvertToImage();
//    }
}
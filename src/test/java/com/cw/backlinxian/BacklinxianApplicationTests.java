package com.cw.backlinxian;

import com.cw.backlinxian.vo.BackPersonVo;
import com.cw.backlinxian.vo.DailyCountVo;
import org.apache.commons.lang.time.DateUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.RowMapper;
import org.springframework.util.ResourceUtils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Array;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

@SpringBootTest
class BacklinxianApplicationTests {

    /*
    * 返乡类型，1-太原，2-省内除太原，3-省外。
    * 2022年4月21日17:34:16，不满足现有的统计算法，新增类型4，从类型2中筛选出一只在临县的。
    * 4- 一直在临县
    * */

    @Autowired
    JdbcTemplate jdbcTemplate;

    // =================init读取疫情防控表，入库当日数据。delete & refresh。======================
    @Test
    void contextLoads() {
        // 读取excel
        Workbook wb2 = ExcelUtil.readExcel(new File("C:\\Users\\99543\\Desktop\\tmp\\1临泉镇新冠疫情防控返临人员排查表4月.xlsx"));
        Sheet sheet2 = wb2.getSheetAt(0);

        jdbcTemplate.update("delete from cw.back_linxian_people /*where uploaddate = '2022.04.22'*/;"); // 清空表
        for (int i = 1; i<sheet2.getPhysicalNumberOfRows(); i++) {
            Row row = sheet2.getRow(i);
            int j = 1; // 从name开始解析，放到arr里下标-1
            String insertSql = "INSERT INTO cw.back_linxian_people\n" +
                    "(name, id, tel, address, homeaddr, backdesc, backtime, xingchengma, jiankangma, testresult, yimiao, temp, uploadvillage, wanggeyuan, uploaddate, `type` ,`method`,person1,person2,person3,person4)\n" +
                    "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);";
            String[] arr = new String[21];
            while (j<18) {
                if(null!=row.getCell(j)) {
                    arr[j-1] = Util.getCellValueByCell(row.getCell(j)); // 处理不同格式的值
                } else {
                    arr[j-1] = "";
                }
                j++;
            }
            String[] strings = Util.wubaoyiMap.get(arr[12]);
            if(null!=strings) {
                arr[j++ - 1] = strings[0];
                arr[j++ - 1] = strings[1];
                arr[j++ - 1] = strings[2];
                arr[j++ - 1] = strings[3];
            }

            jdbcTemplate.update(insertSql, arr);
        }
    }

    // 根据日期（date eg. 2022.04.15）查询所有疫情名单或者所有（ALL）
    List<BackPersonVo> queryBackPerson(String date) {
        String sql = "";
        if("ALL" .equals(date)) {
            Object[] objects = {};
            sql = "select * from cw.back_linxian_people;";
            List<BackPersonVo> result = jdbcTemplate.query(sql, objects, new BackPersonMapper());
            return result;
        } else {
            Object[] objects = {date};
            sql = "select * from cw.back_linxian_people where uploaddate = ?;";
            List<BackPersonVo> result = jdbcTemplate.query(sql, objects, new BackPersonMapper());
            return result;
        }
    }

    /**
     * 0、生成太原、省内、省外的五包一名单
     * 生成完直接发给他们，不需要调整格式
     * @param date
     * @param result
     * @param type
     */
    void generateWubaoyiExcel(String date, List<BackPersonVo> result, String type) {
        // 读取excel
        Workbook wb2 = ExcelUtil.readExcel(new File("C:\\Users\\99543\\Desktop\\tmp\\0报表模板\\0五包一临泉镇返临人员排查表模板.xlsx"));
        Sheet sheet2 = wb2.getSheetAt(0);
        int rowNum = 2; // 从第3行开始创建
        int count = 0; // 序号
        for (BackPersonVo vo : result) {
            Row row = sheet2.createRow(rowNum++);
            int i = 0;
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(++count);
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getName());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getUploadvillage());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue("");
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getId());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getBackdesc());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getBacktime());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getTel());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getPerson1());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getPerson2());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getPerson3());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getPerson4());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getWanggeyuan());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue("阴性");
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue("");
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue("");
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue("");
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue("");
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue("临泉镇人民政府");
        }
        String typeName = "";
        if("1".equals(type)) {
            typeName = "太原";
        }
        if("2".equals(type)) {
            typeName = "省内";
        }
        if("3".equals(type)) {
            typeName = "省外";
        }
        try {
            FileOutputStream output=new FileOutputStream("C:\\Users\\99543\\Desktop\\tmp\\1报表生成\\"+date+"_0"+typeName+"五包一临泉镇返临人员排查表.xlsx");
            wb2.write(output);
            output.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // 1、生成五包一名单 -- 只是当日数据，我们自己打印留存
    void generateWubaoyiExcel(String date, List<BackPersonVo> result) {
        // 读取excel
        Workbook wb2 = ExcelUtil.readExcel(new File("C:\\Users\\99543\\Desktop\\tmp\\0报表模板\\1五包一临泉镇返临人员排查表模板.xlsx"));
        Sheet sheet2 = wb2.getSheetAt(0);
        int rowNum = 3; // 从第4行开始创建
        int count = 0; // 序号 & 总数
        int type1Count = 0; // 太原
        int type2Count = 0; // 省内非太原
        int type3Count = 0; // 省外
        int jujiaCount = 0; // 居家隔离人数
        int jizhongCount = 0; // 集中隔离人数

        // 设置一下字体，方便打印在一张上
        Font font = wb2.createFont();
        font.setFontHeightInPoints((short)8); // 小号字体
        CellStyle style=wb2.createCellStyle();
        style.setFont(font);
        for (BackPersonVo vo : result) {
            // 类型数量
            if("1".equals(vo.getType())) {
                type1Count++;
            }
            if("2".equals(vo.getType()) || "4".equals(vo.getType())) { // 省内和临县的属于省内非太原
                type2Count++;
            }
            if("3".equals(vo.getType())) {
                type3Count++;
            }
            // 隔离数量
            if(vo.getMethod().contains("居家")) {
                jujiaCount++;
            }
            if(vo.getMethod().contains("集中")) {
                jizhongCount++;
            }
            Row row = sheet2.createRow(rowNum++);
            int i = 0;
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(++count);
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getName());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getUploadvillage());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue("");
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getId());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getBackdesc());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getBacktime());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getTel());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getPerson1());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getPerson2());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getPerson3());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getPerson4());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getWanggeyuan());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue("阴性");
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue("");
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue("");
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue("");
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue("");
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue("临泉镇人民政府");
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getType());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getUploaddate());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getMethod());
        }
        Cell cell = sheet2.getRow(1).getCell(0); // 汇总信息的cell
        String totalMsg = date + " 共排查"+count+"人，其中太原"+type1Count+"人，省内除太原"+type2Count+"人，省外"+type3Count+"人。采取防控措施："+jujiaCount+"人居家隔离，"+jizhongCount+"人集中隔离。";
        cell.setCellValue(totalMsg);
        try {
            FileOutputStream output=new FileOutputStream("C:\\Users\\99543\\Desktop\\tmp\\1报表生成\\"+date+"_1五包一临泉镇返临人员排查表.xlsx");
            wb2.write(output);
            output.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    // 2、生成卫健委所需的报表数据，每日和累计报表 date "2022.04.22"
    void generateWeijianwei(String date, List<BackPersonVo> result, List<BackPersonVo> resultAll) throws ParseException {
        Workbook wb = ExcelUtil.readExcel(new File("C:\\Users\\99543\\Desktop\\tmp\\0报表模板\\2县卫健委每日报表.xlsx"));
        Sheet sheet = wb.getSheetAt(0);
        ///1、===============当日数据===============
        int type1Count = 0; // 太原
        int type2Count = 0; // 省内
        int type3Count = 0; // 省外
        // 集中隔离人数
        int type1JizhongCount = 0; // 太原
        int type2JizhongCount = 0; // 省内
        int type3JizhongCount = 0; // 省外
        // 居家隔离
        int type1JujiaCount = 0; // 太原
        int type2JujiaCount = 0; // 省内
        int type3JujiaCount = 0; // 省外
        // 解除隔离 - 当日解除隔离的按七天前计算
        int type1JiechuCount = 0; // 太原
        int type2JiechuCount = 0; // 省内
        int type3JiechuCount = 0; // 省外
        // 失联
        int type1ShilianCount = 0; // 太原
        int type2ShilianCount = 0; // 省内
        int type3ShilianCount = 0; // 省外
        for (BackPersonVo vo : result) {
            // 类型数量
            if("1".equals(vo.getType())) {
                type1Count++;
                // 隔离数量
                if(vo.getMethod().contains("居家")) {
                    type1JujiaCount++;
                }
                if(vo.getMethod().contains("集中")) {
                    type1JizhongCount++;
                }
            }
            if("2".equals(vo.getType()) || "4".equals(vo.getType())) { // 省内和临县的属于省内非太原
                type2Count++;
                // 隔离数量
                if(vo.getMethod().contains("居家")) {
                    type2JujiaCount++;
                }
                if(vo.getMethod().contains("集中")) {
                    type2JizhongCount++;
                }
            }
            if("3".equals(vo.getType())) {
                type3Count++;
                // 隔离数量
                if(vo.getMethod().contains("居家")) {
                    type3JujiaCount++;
                }
                if(vo.getMethod().contains("集中")) {
                    type3JizhongCount++;
                }
            }
        }
        // 2、===============累计数据===============
        int type1CountAll = 0; // 太原
        int type2CountAll = 0; // 省内
        int type3CountAll = 0; // 省外
        // 集中隔离人数
        int type1JizhongCountAll = 0; // 太原
        int type2JizhongCountAll = 0; // 省内
        int type3JizhongCountAll = 0; // 省外
        // 居家隔离
        int type1JujiaCountAll = 0; // 太原
        int type2JujiaCountAll = 0; // 省内
        int type3JujiaCountAll = 0; // 省外
        // 解除隔离
        int type1JiechuCountAll = 0; // 太原
        int type2JiechuCountAll = 0; // 省内
        int type3JiechuCountAll = 0; // 省外
        // 失联
        int type1ShilianCountAll = 77; // 太原
        int type2ShilianCountAll = 0; // 省内
        int type3ShilianCountAll = 0; // 省外

        SimpleDateFormat sdf = new SimpleDateFormat("yyyy.MM.dd");
        Date d = sdf.parse(date);
        Date d2 = DateUtils.addDays(d, -6);
        Date d3 = DateUtils.addDays(d, -7);
        for (BackPersonVo vo : resultAll) {
            // 类型数量
            if("1".equals(vo.getType())) {
                type1CountAll++;
                // 隔离数量
                if(vo.getMethod().contains("居家")) {
                    type1JujiaCountAll++;
                }
                if(vo.getMethod().contains("集中")) {
                    type1JizhongCountAll++;
                }
                if(sdf.parse(vo.getUploaddate()).before(d2)) { // 上报超过七天的，都算为总的解除隔离的数量
                    type1JiechuCountAll++;
                }
                if(sdf.parse(vo.getUploaddate()).equals(d3)) { // 正好是七天前的日期，认为是当前解除隔离的数量
                    type1JiechuCount++;
                }
            }
            if("2".equals(vo.getType()) || "4".equals(vo.getType())) { // 省内和临县的属于省内非太原
                type2CountAll++;
                // 隔离数量
                if(vo.getMethod().contains("居家")) {
                    type2JujiaCountAll++;
                }
                if(vo.getMethod().contains("集中")) {
                    type2JizhongCountAll++;
                }
                if(sdf.parse(vo.getUploaddate()).before(d2)) {
                    type2JiechuCountAll++;
                }
                if(sdf.parse(vo.getUploaddate()).equals(d3)) { // 正好是七天前的日期，认为是当前解除隔离的数量
                    type2JiechuCount++;
                }
            }
            if("3".equals(vo.getType())) {
                type3CountAll++;
                // 隔离数量
                if(vo.getMethod().contains("居家")) {
                    type3JujiaCountAll++;
                }
                if(vo.getMethod().contains("集中")) {
                    type3JizhongCountAll++;
                }
                if(sdf.parse(vo.getUploaddate()).before(d2)) {
                    type3JiechuCountAll++;
                }
                if(sdf.parse(vo.getUploaddate()).equals(d3)) { // 正好是七天前的日期，认为是当前解除隔离的数量
                    type3JiechuCount++;
                }
            }
        }

        // 将生成好的数字填到excel中生成当日报表。
        int i = 1;
        Row type1Row = sheet.getRow(1);
        type1Row.getCell(i++).setCellValue(type1Count);
        type1Row.getCell(i++).setCellValue(type1Count);
        type1Row.getCell(i++).setCellValue(type1JizhongCount);
        type1Row.getCell(i++).setCellValue(type1JujiaCount);
        type1Row.getCell(i++).setCellValue(type1JiechuCount);
        type1Row.getCell(i++).setCellValue(type1ShilianCount);
        i = 1; // 重置
        Row type1RowAll = sheet.getRow(2);
        type1RowAll.getCell(i++).setCellValue(type1CountAll);
        type1RowAll.getCell(i++).setCellValue(type1CountAll);
        type1RowAll.getCell(i++).setCellValue(type1JizhongCountAll);
        type1RowAll.getCell(i++).setCellValue(type1JujiaCountAll);
        type1RowAll.getCell(i++).setCellValue(type1JiechuCountAll - type1ShilianCountAll); // 这个和检委的那个逻辑不一样，居家隔离人数就是太原应该隔离的，也就是总人数，但是解除隔离的应该要把失联的这部分减除掉， 解除隔离人数+失联人数= 总共七天以前的人数。
        type1RowAll.getCell(i++).setCellValue(type1ShilianCountAll);
        i = 1; // 重置
        Row type2Row = sheet.getRow(3);
        type2Row.getCell(i++).setCellValue(type2Count);
        type2Row.getCell(i++).setCellValue(type2Count);
        type2Row.getCell(i++).setCellValue(type2JizhongCount);
        type2Row.getCell(i++).setCellValue(type2JujiaCount);
        type2Row.getCell(i++).setCellValue(type2JiechuCount);
        type2Row.getCell(i++).setCellValue(type2ShilianCount);
        i = 1; // 重置
        Row type2RowAll = sheet.getRow(4);
        type2RowAll.getCell(i++).setCellValue(type2CountAll);
        type2RowAll.getCell(i++).setCellValue(type2CountAll);
        type2RowAll.getCell(i++).setCellValue(type2JizhongCountAll);
        type2RowAll.getCell(i++).setCellValue(type2JujiaCountAll);
        type2RowAll.getCell(i++).setCellValue(type2JiechuCountAll);
        type2RowAll.getCell(i++).setCellValue(type2ShilianCountAll);
        i = 1; // 重置
        Row type3Row = sheet.getRow(6);
        type3Row.getCell(i++).setCellValue(type3Count);
        type3Row.getCell(i++).setCellValue(type3Count);
        type3Row.getCell(i++).setCellValue(type3JizhongCount);
        type3Row.getCell(i++).setCellValue(type3JujiaCount);
        i++;
        type3Row.getCell(i++).setCellValue(type3JiechuCount);
        type3Row.getCell(i++).setCellValue(type3ShilianCount);
        i = 1; // 重置
        Row type3RowAll = sheet.getRow(7);
        type3RowAll.getCell(i++).setCellValue(type3CountAll);
        type3RowAll.getCell(i++).setCellValue(type3CountAll);
        type3RowAll.getCell(i++).setCellValue(type3JizhongCountAll);
        type3RowAll.getCell(i++).setCellValue(type3JujiaCountAll);
        i++;
        type3RowAll.getCell(i++).setCellValue(type3JiechuCountAll);
        type3RowAll.getCell(i++).setCellValue(type3ShilianCountAll);
        try {
            FileOutputStream output=new FileOutputStream("C:\\Users\\99543\\Desktop\\tmp\\1报表生成\\"+date+"_2县卫健委每日报表.xlsx");
            wb.write(output);
            output.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // 3、生成监督委员会需要的异常码人员名单和累计报表
    void generateJianwei(String date, List<BackPersonVo> result, List<BackPersonVo> resultAll) throws ParseException {
        // 1.生成黄码和红码人员名单 - 当日
        // 读取excel
        Workbook wb = ExcelUtil.readExcel(new File("C:\\Users\\99543\\Desktop\\tmp\\0报表模板\\3红码黄码名单模板.xlsx"));
        Sheet sheet = wb.getSheetAt(0);
        int rowNum = 1; // 从第1行开始创建
        int count = 0; // 序号 & 总数
        for (BackPersonVo vo : result) {
            if(vo.getJiankangma().contains("黄") || vo.getJiankangma().contains("红")|| vo.getJiankangma().contains("否")) {
                // 认为健康码是异常的
                Row row = sheet.createRow(rowNum++);
                int i = 0;
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(++count); // 序号
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getName());
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getId());
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getTel());
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getAddress());
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getHomeaddr());
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getBackdesc());
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getBacktime());
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getXingchengma());
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getJiankangma());
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getTestresult());
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getYimiao());
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getTemp());
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getUploadvillage());
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getWanggeyuan());
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getUploaddate());
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getType());
                row.createCell(i).setCellType(CellType.STRING);row.getCell(i++).setCellValue(vo.getMethod());
            }
        }
        try {
            FileOutputStream output=new FileOutputStream("C:\\Users\\99543\\Desktop\\tmp\\1报表生成\\"+date+"_3红码黄码名单模板.xlsx");
            wb.write(output);
            output.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }

        // 2.生成日报表 - 累计,查询所有记录
        // 读取excel
        Workbook wb2 = ExcelUtil.readExcel(new File("C:\\Users\\99543\\Desktop\\tmp\\0报表模板\\4工作簿模板.xlsx"));
        Sheet sheet2 = wb2.getSheetAt(0);
        // 统计截止日期
        sheet2.getRow(1).getCell(0).setCellValue("截止" + date);
        int shengwaiCount = 0;
        int shengneiCount = 0;
        int jizhongCount = 0;
        int jujiaCount = 0;
        int fanchengCount = 77; // 未隔离返程，这个数字是固定的，且只针对当初太原的，在统计居家隔离数字的时候把这项删掉。
        int jiechuCount = 0;
        // 红码外来人员
        int hongmaOut = 0;
        // 红码在临未外出
        int hongmaIn = 0;
        // 黄码外来人员
        int huangmaOut = 0;
        // 黄码在临未外出
        int huangmaIn = 0;
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy.MM.dd");
        Date d = sdf.parse(date);
        Date d2 = DateUtils.addDays(d, -6);
        for (BackPersonVo vo : resultAll) {
            if("3".equals(vo.getType())) {
                shengwaiCount++;
            } else {
                shengneiCount++;
            }
            // 隔离数量
            if(vo.getMethod().contains("居家")) {
                jujiaCount++;
            }
            if(vo.getMethod().contains("集中")) {
                jizhongCount++;
            }
            if(vo.getJiankangma().contains("红")) {
                // 健康码是红码
                if("4".equals(vo.getType())) {
                    // 临县未外出
                    hongmaIn++;
                } else {
                    // 入临
                    hongmaOut++;
                }
            } else if(vo.getJiankangma().contains("黄") || vo.getJiankangma().contains("否")) {
                // 健康码是黄码
                if("4".equals(vo.getType())) {
                    // 临县未外出
                    huangmaIn++;
                } else {
                    // 入临
                    huangmaOut++;
                }
            }
            // 接触隔离人数的计算，为了省事，计算七天之前的人数
            if(sdf.parse(vo.getUploaddate()).before(d2)) {
                jiechuCount++;
            }
        }
        int i = 1; // 从第一列开始填入数据
        sheet2.getRow(4).getCell(i++).setCellValue(shengwaiCount);
        sheet2.getRow(4).getCell(i++).setCellValue(shengneiCount);
        i++; // 跳过一个
        sheet2.getRow(4).getCell(i++).setCellValue(jizhongCount);
        sheet2.getRow(4).getCell(i++).setCellValue(jujiaCount - fanchengCount); // 居家隔离的要把逃出去的减掉，否则总数会超总人数，逻辑问题
        sheet2.getRow(4).getCell(i++).setCellValue(fanchengCount);
        i++; // 跳过一个
        sheet2.getRow(4).getCell(i++).setCellValue(jiechuCount);
        sheet2.getRow(4).getCell(i++).setCellValue(hongmaOut);
        sheet2.getRow(4).getCell(i++).setCellValue(hongmaIn);
        sheet2.getRow(4).getCell(i++).setCellValue(huangmaOut);
        sheet2.getRow(4).getCell(i++).setCellValue(huangmaIn);
        try {
            FileOutputStream output2=new FileOutputStream("C:\\Users\\99543\\Desktop\\tmp\\1报表生成\\"+date+"_4工作簿模板.xlsx");
            wb2.write(output2);
            output2.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // 4、生成当日的日报(包含每日和累计数据) - 给镇领导看
    void ribao(String today) {
        List<String> uploaddate = jdbcTemplate.query("select distinct uploaddate from back_linxian_people;", new RowMapper<String>() {
            @Override
            public String mapRow(ResultSet resultSet, int i) throws SQLException {
                return resultSet.getString("uploaddate");
            }
        });
        ArrayList<DailyCountVo> dailyCountVos = new ArrayList<>();
        // 根据每一天的日报返回
        for (String date : uploaddate) {
            List<BackPersonVo> backPersonVos = queryBackPerson(date);
            System.out.println(date + "->" +backPersonVos.size());
            DailyCountVo dailyCountVo = new DailyCountVo();
            int type1Count = 0;
            int type2Count = 0;
            int type3Count = 0;
            int totalCount = backPersonVos.size();
            int homeCount = 0;
            int groupCount = 0;
            for (BackPersonVo vo : backPersonVos) {
                // 类型数量
                if("1".equals(vo.getType())) {
                    type1Count++;
                }
                if("2".equals(vo.getType()) || "4".equals(vo.getType())) { // 省内和临县的属于省内非太原
                    type2Count++;
                }
                if("3".equals(vo.getType())) {
                    type3Count++;
                }
                // 隔离数量
                if(vo.getMethod().contains("居家")) {
                    homeCount++;
                }
                if(vo.getMethod().contains("集中")) {
                    groupCount++;
                }
            }
            dailyCountVo.setDate(date);
            dailyCountVo.setType1Count(type1Count);
            dailyCountVo.setType2Count(type2Count);
            dailyCountVo.setType3Count(type3Count);
            dailyCountVo.setTotalCount(totalCount);
            dailyCountVo.setHomeCount(homeCount);
            dailyCountVo.setGroupCount(groupCount);
            dailyCountVos.add(dailyCountVo);
        }

        // 分开处理是为了后续拆分接口方便，不然可以直接写一起了
        // 读取模板文件
        Workbook wb = ExcelUtil.readExcel(new File("C:\\Users\\99543\\Desktop\\tmp\\0报表模板\\5每日统计最新模板.xlsx"));
        Sheet sheet = wb.getSheetAt(0);
        // 设置一下字体，居中美观
        Font font = wb.createFont();
        font.setFontHeightInPoints((short)12);
        CellStyle style=wb.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        int rowNum = 2; // 从第3行开始创建
        // 累计的数据
        int type1CountAll = 0;
        int type2CountAll = 0;
        int type3CountAll = 0;
        int totalCountAll = 0;
        int homeCountAll = 0;
        int groupCountAll = 0;
        // 输出日报所需数据
        for (DailyCountVo vo : dailyCountVos) {
            Row row = sheet.createRow(rowNum++);
            int i = 0;
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getDate());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getType1Count());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getType2Count());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getType3Count());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getTotalCount());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getHomeCount());
            row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(vo.getGroupCount());
            type1CountAll+= vo.getType1Count();
            type2CountAll+= vo.getType2Count();
            type3CountAll+= vo.getType3Count();
            totalCountAll+= vo.getTotalCount();
            homeCountAll+= vo.getHomeCount();
            groupCountAll+= vo.getGroupCount();
        }
        // 加一行累计
        Row row = sheet.createRow(rowNum++);
        int i = 0;
        row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue("累计");
        row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(type1CountAll);
        row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(type2CountAll);
        row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(type3CountAll);
        row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(totalCountAll);
        row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(homeCountAll);
        row.createCell(i).setCellType(CellType.STRING);row.getCell(i).setCellStyle(style);row.getCell(i++).setCellValue(groupCountAll);
        try {
            FileOutputStream output=new FileOutputStream("C:\\Users\\99543\\Desktop\\tmp\\1报表生成\\"+today+"_5每日统计最新模板.xlsx");
            wb.write(output);
            output.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // 主方法
    @Test
    void main() throws ParseException {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy.MM.dd");
        String date = sdf.format(new Date());
        List<BackPersonVo> result = queryBackPerson(date);
        List<BackPersonVo> resultAll = queryBackPerson("ALL");
        // 1、卫健委
        // 生成分的五包一文件
        List<BackPersonVo> type1List = new ArrayList<>();
        List<BackPersonVo> type2List = new ArrayList<>();
        List<BackPersonVo> type3List = new ArrayList<>();
        for (BackPersonVo vo : result) {
            if(vo.getMethod().contains("居家")) {
                // 居家隔离的生成五包一名单
                if("1".equals(vo.getType())) {
                    type1List.add(vo);
                }
                if("2".equals(vo.getType()) || "4".equals(vo.getType())) {
                    type2List.add(vo);
                }
                if("3".equals(vo.getType())) {
                    type3List.add(vo);
                }
            }
        }
        // 太原、省内、省外
        generateWubaoyiExcel(date, type1List, "1");
        generateWubaoyiExcel(date, type2List, "2");
        generateWubaoyiExcel(date, type3List, "3");

        // 总的打印的五包一文件
        generateWubaoyiExcel(date, result);
        generateWeijianwei(date, result, resultAll);
        // 2022年4月22日19:31:45 检委会的这两个表又说不用报了。
        // 2、检委会
//        generateJianwei(date, result, resultAll);
        // 3、当日镇日报
        ribao(date);
    }

}

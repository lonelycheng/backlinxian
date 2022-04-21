package com.cw.backlinxian;

import com.cw.backlinxian.vo.BackPersonVo;
import org.apache.commons.lang.time.DateUtils;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.RowMapper;

import java.io.File;
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

@SpringBootTest
class BacklinxianApplicationTests {

    /*
    * 返乡类型，1-太原，2-省内除太原，3-省外。
    * 2022年4月21日17:34:16，不满足现有的统计算法，新增类型4，从类型2中筛选出一只在临县的。
    * 4- 一直在临县
    * */

    @Autowired
    JdbcTemplate jdbcTemplate;

    // 读取疫情防控表，入库当日数据。delete & refresh
    @Test
    void contextLoads() {
        // 读取excel
        Workbook wb2 = ExcelUtil.readExcel(new File("C:\\Users\\99543\\Desktop\\tmp\\1临泉镇新冠疫情防控返临人员排查表4月.xlsx"));
        Sheet sheet2 = wb2.getSheetAt(0);

        jdbcTemplate.update("delete from cw.back_linxian_people /*where uploaddate = '2022.04.21'*/;"); // 清空表
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
    private List<BackPersonVo> queryBackPerson(String date) {
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

    // 生成五包一名单
    @Test
    void generateWubaoyiExcel() {
        String date = "2022.04.21";
        List<BackPersonVo> result = queryBackPerson(date);
        System.out.println("共解析到:" + result.size() + "条记录");

        // 读取excel
        Workbook wb2 = ExcelUtil.readExcel(new File("C:\\Users\\99543\\Desktop\\tmp\\3五包一返临\\五包一临泉镇返临人员排查表模板.xlsx"));
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
            FileOutputStream output=new FileOutputStream("C:\\Users\\99543\\Desktop\\tmp\\3五包一返临\\五包一临泉镇返临人员排查表" + date +".xlsx");
            wb2.write(output);
            output.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    @Test
    // 生成监督委员会需要的异常码人员名单和日报表
    void generateJianwei() throws ParseException {
        // 1.生成黄码和红码人员名单 - 当日
        String date = "2022.04.21";
        List<BackPersonVo> result = queryBackPerson(date);
        System.out.println("共解析到:" + result.size() + "条记录");
        // 读取excel
        Workbook wb = ExcelUtil.readExcel(new File("C:\\Users\\99543\\Desktop\\tmp\\4报备\\红码黄码名单模板.xlsx"));
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
            FileOutputStream output=new FileOutputStream("C:\\Users\\99543\\Desktop\\tmp\\4报备\\红码黄码名单模板" + date +".xlsx");
            wb.write(output);
            output.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }

        // 2.生成日报表 - 累计,查询所有记录
        List<BackPersonVo> resultAll = queryBackPerson("ALL");
        System.out.println("共解析到:" + resultAll.size() + "条记录");
        // 读取excel
        Workbook wb2 = ExcelUtil.readExcel(new File("C:\\Users\\99543\\Desktop\\tmp\\4报备\\工作簿模板.xlsx"));
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
            FileOutputStream output2=new FileOutputStream("C:\\Users\\99543\\Desktop\\tmp\\4报备\\工作簿模板" + date +".xlsx");
            wb2.write(output2);
            output2.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    void test() throws ParseException {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy.MM.dd");
        Date date = sdf.parse("2022.04.21");
        Date date1 = DateUtils.addDays(date, -7);
        System.out.println(sdf.format(date1));
    }

    // 生成当日的日报
    @Test
    void ribao() {
        List<String> uploaddate = jdbcTemplate.query("select distinct uploaddate from back_linxian_people;", new RowMapper<String>() {
            @Override
            public String mapRow(ResultSet resultSet, int i) throws SQLException {
                return resultSet.getString("uploaddate");
            }
        });
        // 根据每一天的日报返回
        for (String date : uploaddate) {
            List<BackPersonVo> backPersonVos = queryBackPerson(date);
            System.out.println(date + "->" +backPersonVos.size());
        }
    }

}

package com.cw.backlinxian;

import com.cw.backlinxian.vo.BackPersonVo;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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
import java.util.ArrayList;
import java.util.List;

@SpringBootTest
class BacklinxianApplicationTests {

    @Autowired
    JdbcTemplate jdbcTemplate;

    // 读取疫情防控表，入库当日数据。delete & refresh
    @Test
    void contextLoads() {
        // 读取excel
        Workbook wb2 = ExcelUtil.readExcel(new File("C:\\Users\\99543\\Desktop\\tmp\\1临泉镇新冠疫情防控返临人员排查表4月.xlsx"));
        Sheet sheet2 = wb2.getSheetAt(0);

        jdbcTemplate.update("delete from cw.back_linxian_people where uploaddate = '2022.04.19';"); // 清空表
        for (int i = 2347; i<sheet2.getPhysicalNumberOfRows(); i++) {
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
        String date = "2022.04.19";
        List<BackPersonVo> result = queryBackPerson(date);
        System.out.println("共解析到:" + result.size() + "条记录");

        // 读取excel
        Workbook wb2 = ExcelUtil.readExcel(new File("C:\\Users\\99543\\Desktop\\tmp\\3五包一返临\\五包一临泉镇太原返临人员排查表模板.xlsx"));
        Sheet sheet2 = wb2.getSheetAt(0);
        int rowNum = 1; // 从第二行开始创建
        int count = 0; // 序号
        for (BackPersonVo vo : result) {
            Row row = sheet2.createRow(rowNum++);
            int i = 0;
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue(++count);
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue(vo.getName());
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue(vo.getUploadvillage());
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue("");
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue(vo.getId());
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue(vo.getBackdesc());
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue(vo.getBacktime());
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue(vo.getTel());
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue(vo.getPerson1());
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue(vo.getPerson2());
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue(vo.getPerson3());
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue(vo.getPerson4());
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue(vo.getWanggeyuan());
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue("阴性");
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue("");
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue("");
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue("");
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue("");
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue("临泉镇人民政府");
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue(vo.getType());
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue(vo.getUploaddate());
            row.createCell(i).setCellType(CellType.STRING);row.createCell(i++).setCellValue(vo.getMethod());
        }

        try {
            FileOutputStream output=new FileOutputStream("C:\\Users\\99543\\Desktop\\tmp\\3五包一返临\\五包一临泉镇太原返临人员排查表" + date +".xlsx");
            wb2.write(output);
            output.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }

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

package com.cw.backlinxian;

import com.cw.backlinxian.vo.BackPersonVo;
import org.springframework.jdbc.core.RowMapper;

import java.sql.ResultSet;
import java.sql.SQLException;

public class BackPersonMapper implements RowMapper<BackPersonVo> {
    @Override
    public BackPersonVo mapRow(ResultSet resultSet, int i) throws SQLException {
        BackPersonVo backPersonVo = new BackPersonVo();
        backPersonVo.setName(resultSet.getString("name"));
        backPersonVo.setId(resultSet.getString("id"));
        backPersonVo.setTel(resultSet.getString("tel"));
        backPersonVo.setAddress(resultSet.getString("address"));
        backPersonVo.setHomeaddr(resultSet.getString("homeaddr"));
        backPersonVo.setBackdesc(resultSet.getString("backdesc"));
        backPersonVo.setBacktime(resultSet.getString("backtime"));
        backPersonVo.setXingchengma(resultSet.getString("xingchengma"));
        backPersonVo.setJiankangma(resultSet.getString("jiankangma"));
        backPersonVo.setTestresult(resultSet.getString("testresult"));
        backPersonVo.setYimiao(resultSet.getString("yimiao"));
        backPersonVo.setTemp(resultSet.getString("temp"));
        backPersonVo.setUploadvillage(resultSet.getString("uploadvillage"));
        backPersonVo.setWanggeyuan(resultSet.getString("wanggeyuan"));
        backPersonVo.setUploaddate(resultSet.getString("uploaddate"));
        backPersonVo.setType(resultSet.getString("type"));
        backPersonVo.setMethod(resultSet.getString("method"));
        backPersonVo.setPerson1(resultSet.getString("person1"));
        backPersonVo.setPerson2(resultSet.getString("person2"));
        backPersonVo.setPerson3(resultSet.getString("person3"));
        backPersonVo.setPerson4(resultSet.getString("person4"));
        return backPersonVo;
    }
}

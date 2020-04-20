package com.wzp.poi;

import com.wzp.poi.bean.User;
import com.wzp.poi.mapper.UserMapper;
import com.wzp.poi.utils.ExcelUtil;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.autoconfigure.security.SecurityProperties;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;
import sun.security.util.AuthResources_it;

import javax.annotation.Resource;
import java.util.ArrayList;
import java.util.List;

@RunWith(SpringRunner.class)
@SpringBootTest
public class PoiApplicationTests {

    @Resource
    UserMapper userMapper;


    @Test
    public void findAll() {
        System.out.println(userMapper.findAllColumnName());
    }

    @Test
    public void test() throws NoSuchFieldException, IllegalAccessException {
        List list = userMapper.findAllUser();
        List<String> columnName = userMapper.findAllColumnName();
        ExcelUtil.export(list, User.class, "学生信息表", columnName, "User");
    }

    @Test
    public void excelTemplate() {
        ExcelUtil.excelTemplate();
    }

    @Test
    public void intoDatabase() throws Exception {
        String str = "ceshi_123456789.mp3";
        System.out.println(str.substring(str.lastIndexOf("_")));

    }

}

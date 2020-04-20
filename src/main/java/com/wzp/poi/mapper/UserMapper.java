package com.wzp.poi.mapper;

import com.wzp.poi.bean.User;
import org.apache.ibatis.annotations.Param;

import java.util.List;
import java.util.Map;


public interface UserMapper {

    Integer addUser(User user);

    List<User> findAllUser();

    User findUser(Long id);

    Integer addUserByList(List userList);

    List<String> findAllColumnName();

    Map<String, String> getToMap(@Param("id") String id);


}

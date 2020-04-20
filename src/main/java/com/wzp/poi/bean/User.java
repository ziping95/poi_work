package com.wzp.poi.bean;

import lombok.Data;

import java.io.Serializable;

@Data
public class User implements Serializable {

    private static final long serialVersionUID = 1L;

    private String id;
    private String name;
    private String sex;
    private String address;
    private Long phone;
    private Long createDate;
    private Long lastLoginDate;

}

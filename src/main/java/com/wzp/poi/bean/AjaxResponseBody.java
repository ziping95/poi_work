package com.wzp.poi.bean;

import lombok.Data;

import java.io.Serializable;

@Data
public class AjaxResponseBody<T> implements Serializable {

    private String status;
    private String msg;
    private T result;

}

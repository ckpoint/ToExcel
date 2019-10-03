package com.github.ckpoint.toexcel.workbook.model;

import com.github.ckpoint.toexcel.annotation.ExcelHeader;
import lombok.Builder;
import lombok.ToString;

@ToString
@Builder
public class UserModel {

    @ExcelHeader(headerName = "name", headerNames ={"닉네임","이메일", "email"} , priority = 0)
    private String name;
    @ExcelHeader(headerName = "age", priority = 1)
    private Integer age;
    @ExcelHeader(headerName = "gender", priority = 2)
    private String gender;
}

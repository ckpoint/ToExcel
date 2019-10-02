package com.github.ckpoint.toexcel.workbook.model;

import com.github.ckpoint.toexcel.annotation.ExcelHeader;
import lombok.Builder;
import lombok.ToString;

@ToString
@Builder
public class UserModel {

    @ExcelHeader(headerName = "이름", headerNames ={"닉네임","이메일", "email"} , priority = 0)
    private String name;
    @ExcelHeader(headerName = "나이", priority = 1)
    private Integer age;
    @ExcelHeader(headerName = "성별", priority = 2)
    private String gender;
}

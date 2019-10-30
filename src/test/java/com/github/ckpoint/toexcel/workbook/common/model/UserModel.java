package com.github.ckpoint.toexcel.workbook.common.model;

import com.github.ckpoint.toexcel.annotation.ExcelHeader;
import com.github.ckpoint.toexcel.util.ToExcelCsvConverter;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.ToString;

@ToString
@Getter
@NoArgsConstructor
public class UserModel extends CommonModel implements ToExcelCsvConverter {

    @ExcelHeader(headerName = "name", headerNames = {"닉네임", "이메일", "email"}, priority = 0)
    private String name;
    @ExcelHeader(headerName = "age", priority = 1)
    private Integer age;
    @ExcelHeader(headerName = "gender", priority = 2)
    private String gender;

    @Builder
    public UserModel(String name, Integer age, String gender) {
        this.setId(name + "/" + age);
        this.name = name;
        this.age = age;
        this.gender = gender;
    }
}

package com.github.ckpoint.toexcel.workbook;

import com.github.ckpoint.toexcel.annotation.ExcelHeader;
import com.github.ckpoint.toexcel.core.ToWorkBook;
import com.github.ckpoint.toexcel.core.ToWorkSheet;
import lombok.ToString;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.List;

public class ExcelMapTest {

    @Test
    public void readExcelTest() throws IOException {

        ToWorkBook workBook = new ToWorkBook(new File("d:/excel/source.xls"));
        ToWorkSheet sheet = workBook.getSheetAt(0);
        List<UserModel> userModelList = sheet.map(UserModel.class);

        for (UserModel userModel : userModelList) {
           System.out.println(userModel.toString());
        }
    }

}

@ToString
class UserModel{

    @ExcelHeader(headerName = "name", headerNames = {"희섭", "이름"})
    private String name;
    @ExcelHeader(headerName = "나이")
    private Integer age;
}
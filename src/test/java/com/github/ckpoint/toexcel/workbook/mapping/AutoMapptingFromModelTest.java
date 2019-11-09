package com.github.ckpoint.toexcel.workbook.mapping;

import com.github.ckpoint.toexcel.core.ToWorkBook;
import com.github.ckpoint.toexcel.core.ToWorkSheet;
import com.github.ckpoint.toexcel.core.type.ToWorkBookType;
import com.github.ckpoint.toexcel.workbook.common.model.UserModel;
import org.junit.Assert;
import org.junit.Test;

import java.io.IOException;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class AutoMapptingFromModelTest {

    @Test
    public void userModelsToExcelWriteTest_01() throws IOException {
        List<UserModel> userModelList =
                IntStream.range(0, 100).mapToObj(i ->
                        UserModel.builder().name("tester" + i).age(i).gender("man").build()).collect(Collectors.toList());

        ToWorkBook workBook = new ToWorkBook(ToWorkBookType.XSSF);
        ToWorkSheet sheet = workBook.createSheet();
        sheet.from(userModelList);

        Assert.assertTrue(sheet.getCell(0, 0).getStringCellValue().equalsIgnoreCase("name"));
        Assert.assertTrue(sheet.getCell(0, 1).getStringCellValue().equalsIgnoreCase("age"));
        Assert.assertTrue(sheet.getCell(0, 2).getStringCellValue().equalsIgnoreCase("gender"));

        Assert.assertTrue(sheet.getCell(1, 0).getStringCellValue().equalsIgnoreCase("tester0"));
        Assert.assertTrue(sheet.getCell(1, 1).getStringCellValue().equalsIgnoreCase("0"));
        Assert.assertTrue(sheet.getCell(1, 2).getStringCellValue().equalsIgnoreCase("man"));

        Assert.assertTrue(sheet.getRowCount() == 101);


        workBook.write("target/excel/map/write_test_1");

    }
}

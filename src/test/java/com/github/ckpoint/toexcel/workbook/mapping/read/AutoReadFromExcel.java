package com.github.ckpoint.toexcel.workbook.mapping.read;

import com.github.ckpoint.toexcel.core.ToWorkBook;
import com.github.ckpoint.toexcel.core.ToWorkSheet;
import com.github.ckpoint.toexcel.core.type.WorkBookType;
import com.github.ckpoint.toexcel.workbook.common.model.UserModel;
import org.junit.Assert;
import org.junit.FixMethodOrder;
import org.junit.Test;
import org.junit.runners.MethodSorters;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

@FixMethodOrder(MethodSorters.NAME_ASCENDING)
public class AutoReadFromExcel {

    @Test
    public void a_userModelsToExcelWriteTest_01() throws IOException {
        List<UserModel> userModelList =
                IntStream.range(0, 100).mapToObj(i ->
                        UserModel.builder().name("tester" + i).age(i).gender("man").build()).collect(Collectors.toList());

        ToWorkBook workBook = new ToWorkBook(WorkBookType.XSSF);
        ToWorkSheet sheet = workBook.createSheet();
        sheet.from(userModelList);

        Assert.assertTrue(sheet.getCell(0, 0).getStringCellValue().equalsIgnoreCase("name"));
        Assert.assertTrue(sheet.getCell(0, 1).getStringCellValue().equalsIgnoreCase("age"));
        Assert.assertTrue(sheet.getCell(0, 2).getStringCellValue().equalsIgnoreCase("gender"));

        Assert.assertTrue(sheet.getCell(1, 0).getStringCellValue().equalsIgnoreCase("tester0"));
        Assert.assertTrue(sheet.getCell(1, 1).getStringCellValue().equalsIgnoreCase("0"));
        Assert.assertTrue(sheet.getCell(1, 2).getStringCellValue().equalsIgnoreCase("man"));

        workBook.write("target/excel/map/read_test_1");
    }

    @Test
    public void b_excelReadToModelTest_01() throws IOException {
        ToWorkBook toWorkBook = new ToWorkBook(new File("target/excel/map/read_test_1.xlsx"));
        ToWorkSheet toWorkSheet = toWorkBook.getSheetAt(0);
        List<UserModel> userModels = toWorkSheet.map(UserModel.class);

        Assert.assertTrue(userModels.size() == 100);
        Assert.assertTrue(userModels.get(0).getName().equals("tester0") );
        Assert.assertTrue(userModels.get(0).getAge().equals(0) );
        Assert.assertTrue(userModels.get(0).getGender().equals("man") );

    }

}

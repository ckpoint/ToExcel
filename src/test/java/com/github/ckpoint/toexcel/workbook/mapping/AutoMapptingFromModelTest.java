package com.github.ckpoint.toexcel.workbook.mapping;

import com.github.ckpoint.toexcel.core.ToWorkBook;
import com.github.ckpoint.toexcel.core.ToWorkSheet;
import com.github.ckpoint.toexcel.core.type.WorkBookType;
import com.github.ckpoint.toexcel.workbook.model.UserModel;
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

        ToWorkBook workBook = new ToWorkBook(WorkBookType.XSSF);
        ToWorkSheet sheet = workBook.createSheet();
        sheet.from(userModelList);

        workBook.writeFile("target/excel/map/write_test_1");

    }
}

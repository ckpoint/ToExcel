package com.github.ckpoint.toexcel.workbook.manual.normal;

import com.github.ckpoint.toexcel.core.ToWorkBook;
import com.github.ckpoint.toexcel.core.ToWorkSheet;
import com.github.ckpoint.toexcel.core.type.ToWorkBookType;
import com.github.ckpoint.toexcel.exception.NotFoundExtException;
import com.github.ckpoint.toexcel.workbook.common.model.UserModel;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.junit.Assert;
import org.junit.FixMethodOrder;
import org.junit.Test;
import org.junit.runners.MethodSorters;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static org.junit.Assert.fail;

@FixMethodOrder(MethodSorters.NAME_ASCENDING)
public class CommonWorkbookTest {

    @Test
    public void a_userModelsToExcelWriteTest_01() throws IOException {
        List<UserModel> userModelList =
                IntStream.range(0, 100).mapToObj(i ->
                        UserModel.builder().name("tester" + i).age(i).gender("man").build()).collect(Collectors.toList());

        ToWorkBook workBook = new ToWorkBook(ToWorkBookType.XSSF);
        ToWorkSheet sheet = workBook.createSheet();
        sheet.from(userModelList);
        workBook.write("target/excel/map/read_test_1");
    }

    @Test
    public void create_workbook_with_type_and_fis_test() throws IOException {

        File file = new File("target/excel/map/read_test_1.xlsx");

        ToWorkBook workBook = new ToWorkBook(ToWorkBookType.XSSF, new FileInputStream(file));
        ToWorkSheet sheet = workBook.getSheetAt(0);

        Assert.assertEquals(sheet.getRowCount(), 101);
    }

    @Test(expected = OfficeXmlFileException.class)
    public void create_workbook_when_type_and_fis_not_equal_test() throws IOException {

        File file = new File("target/excel/map/read_test_1.xlsx");

        ToWorkBook workBook = new ToWorkBook(ToWorkBookType.HSSF, new FileInputStream(file));

        fail("No equals Type file");
    }

    @Test(expected = FileNotFoundException.class)
    public void create_fail_when_no_existed_file_workbook_test() throws IOException {

        File file = new File("no/existed/file.txt");

        ToWorkBook workBook = new ToWorkBook(ToWorkBookType.XSSF,file);

        fail("No existed file");
    }

    @Test(expected = FileNotFoundException.class)
    public void create_fail_when_file_is_Directory_workbook_test() throws IOException {

        File file = new File("no/existed/");

        ToWorkBook workBook = new ToWorkBook(ToWorkBookType.XSSF, file);

        fail("No existed file");
    }
}

package com.github.ckpoint.toexcel.workbook.manual.normal;

import com.github.ckpoint.toexcel.core.ToWorkBook;
import com.github.ckpoint.toexcel.core.ToWorkSheet;
import com.github.ckpoint.toexcel.core.type.WorkBookType;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import static org.junit.Assert.fail;

public class CommonWorkbookTest {

    @Test
    public void create_workbook_with_type_and_fis_test() throws IOException {

        File file = new File("target/excel/map/read_test_1.xlsx");

        ToWorkBook workBook = new ToWorkBook(WorkBookType.XSSF, new FileInputStream(file));
        ToWorkSheet sheet = workBook.getSheetAt(0);

        Assert.assertEquals(sheet.getRowCount(), 101);
    }

    @Test(expected = OfficeXmlFileException.class)
    public void create_workbook_when_type_and_fis_not_equal_test() throws IOException {

        File file = new File("target/excel/map/read_test_1.xlsx");

        ToWorkBook workBook = new ToWorkBook(WorkBookType.HSSF, new FileInputStream(file));

        fail("No equals Type file");
    }

    @Test(expected = FileNotFoundException.class)
    public void create_fail_when_no_existed_file_workbook_test() throws IOException {

        File file = new File("no/existed/file.txt");

        ToWorkBook workBook = new ToWorkBook(file);

        fail("No existed file");
    }

    @Test(expected = FileNotFoundException.class)
    public void create_fail_when_file_is_Directory_workbook_test() throws IOException {

        File file = new File("no/existed/");

        ToWorkBook workBook = new ToWorkBook(file);

        fail("No existed file");
    }
}

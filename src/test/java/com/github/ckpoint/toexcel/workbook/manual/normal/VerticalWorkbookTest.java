package com.github.ckpoint.toexcel.workbook.manual.normal;

import com.github.ckpoint.toexcel.core.ToWorkBook;
import com.github.ckpoint.toexcel.core.ToWorkSheet;
import com.github.ckpoint.toexcel.core.type.SheetDirection;
import com.github.ckpoint.toexcel.core.type.WorkBookType;
import org.junit.Assert;
import org.junit.Test;

import java.io.IOException;
import java.util.stream.IntStream;

public class VerticalWorkbookTest {

    @Test
    public void vertical_test_01() throws IOException {
        ToWorkBook workBook = new ToWorkBook(WorkBookType.XSSF);
        ToWorkSheet sheet = workBook.createSheet().updateDirection(SheetDirection.VERTICAL);

        sheet.createTitleCell(2, "name", "age");
        IntStream.range(0, 2).forEach(i ->{
            sheet.createCellToNewline("hsim" + i , 1 + i);
        });

        Assert.assertTrue(sheet.getCell(0, 0).getStringCellValue().equals("name"));
        Assert.assertTrue(sheet.getCell(1, 0).getStringCellValue().equals("age"));

        Assert.assertTrue(sheet.getCell(0, 1).getStringCellValue().equals("hsim0"));
        Assert.assertTrue(sheet.getCell(1, 1).getStringCellValue().equals("1"));

        workBook.write("target/excel/manual/normal/vertical_test_01");


    }
}

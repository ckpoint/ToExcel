package com.github.ckpoint.toexcel.workbook.manual;

import com.github.ckpoint.toexcel.core.ToWorkBook;
import com.github.ckpoint.toexcel.core.ToWorkSheet;
import com.github.ckpoint.toexcel.core.type.SheetDirection;
import com.github.ckpoint.toexcel.core.type.WorkBookType;
import org.junit.Assert;
import org.junit.Test;

import java.io.IOException;
import java.util.stream.IntStream;

public class SkipCellTest {

    @Test
    public void skipCellTest_horizon_01() throws IOException {
        ToWorkBook workBook = new ToWorkBook(WorkBookType.XSSF);
        ToWorkSheet sheet = workBook.createSheet().updateDirection(SheetDirection.HORIZON);

        sheet.createTitleCell(2, "이름", "나이");
        sheet.newLine();
        IntStream.range(0, 2).forEach(i -> {
            sheet.skip(1);
            sheet.createCell("희섭" + i, 1 + i);
            sheet.newLine();
        });

        Assert.assertTrue(sheet.getCell(1, 0).getStringCellValue().isEmpty());
        Assert.assertTrue(sheet.getCell(2, 0).getStringCellValue().isEmpty());
        Assert.assertTrue(sheet.getCell(1, 1).getStringCellValue().equals("희섭0"));

        Assert.assertTrue(sheet.getCell(1, 2) != null);
        Assert.assertTrue(sheet.getCell(2, 2) != null);


        workBook.write("target/excel/manual/skip/skip_horizon_01.xls");
    }

    @Test
    public void skipCellTest_vertical_01() throws IOException {
        ToWorkBook workBook = new ToWorkBook(WorkBookType.XSSF);
        ToWorkSheet sheet = workBook.createSheet().updateDirection(SheetDirection.VERTICAL);

        sheet.createTitleCell(2, "이름", "나이");
        sheet.newLine();
        IntStream.range(0, 2).forEach(i -> {
            sheet.skip(1);
            sheet.createCell("희섭" + i, 1 + i);
            sheet.newLine();
        });

        Assert.assertTrue(sheet.getCell(0, 1).getStringCellValue().isEmpty());
        Assert.assertTrue(sheet.getCell(0, 2).getStringCellValue().isEmpty());
        Assert.assertTrue(sheet.getCell(1, 1).getStringCellValue().equals("희섭0"));
        Assert.assertTrue(sheet.getCell(1, 2).getStringCellValue().equals("희섭1"));

        Assert.assertTrue(sheet.getCell(0, 1) != null);
        Assert.assertTrue(sheet.getCell(0, 2) != null);

        workBook.write("target/excel/manual/skip/skip_vertical_01.xls");
    }

}

package com.github.ckpoint.toexcel.workbook.manual.custom;

import com.github.ckpoint.toexcel.core.ToWorkBook;
import com.github.ckpoint.toexcel.core.ToWorkSheet;
import com.github.ckpoint.toexcel.core.style.ToWorkBookStyle;
import com.github.ckpoint.toexcel.core.type.SheetDirection;
import com.github.ckpoint.toexcel.core.type.WorkBookType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.Assert;
import org.junit.Test;

import java.io.IOException;
import java.util.stream.IntStream;

public class StyleCustomTest {

    @Test
    public void test_01() throws IOException {
        ToWorkBook workBook = new ToWorkBook(WorkBookType.XSSF);
        ToWorkSheet sheet = workBook.createSheet().updateDirection(SheetDirection.VERTICAL);

        sheet.createTitleCell(2, "이름", "나이");
        sheet.newLine();

        ToWorkBookStyle toWorkBookStyle = new ToWorkBookStyle();
        toWorkBookStyle.setFillForegroundColor(IndexedColors.BLUE);
        toWorkBookStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

        IntStream.range(0, 2).forEach(i -> {
            sheet.createCell(toWorkBookStyle, "희섭" + i, 1 + i);
            sheet.newLine();
        });

        Cell cell = sheet.getCell(0, 1);
        Assert.assertTrue(cell.getStringCellValue().equals("희섭0"));
        Assert.assertTrue(cell.getCellStyle().getFillForegroundColor() == IndexedColors.BLUE.getIndex());

        workBook.write("target/excel/manual/custom/custom_style_01.xls");


    }
}

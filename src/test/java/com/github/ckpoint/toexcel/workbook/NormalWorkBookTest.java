package com.github.ckpoint.toexcel.workbook;

import com.github.ckpoint.toexcel.core.ToWorkBook;
import com.github.ckpoint.toexcel.core.ToWorkSheet;
import org.junit.Test;

import java.io.IOException;
import java.util.stream.IntStream;

public class NormalWorkBookTest {

    @Test
    public void test_01() throws IOException {
        ToWorkBook workBook = new ToWorkBook();
        ToWorkSheet sheet = workBook.createSheet();

        sheet.createTitleCell(2, "이름", "나이");
        sheet.next();
        IntStream.range(0, 5000).forEach(i ->{
            sheet.createCell("희섭" + i , 1 + i);
            sheet.next();
        });

        workBook.writeFile("d:/excel/test.xls");

    }
}

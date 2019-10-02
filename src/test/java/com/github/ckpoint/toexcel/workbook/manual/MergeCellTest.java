package com.github.ckpoint.toexcel.workbook.manual;

import com.github.ckpoint.toexcel.core.ToWorkBook;
import com.github.ckpoint.toexcel.core.ToWorkSheet;
import com.github.ckpoint.toexcel.core.type.SheetDirection;
import com.github.ckpoint.toexcel.core.type.WorkBookType;
import org.junit.Test;

import java.io.IOException;
import java.util.stream.IntStream;

public class MergeCellTest {

    @Test
    public void mergeCellTest_horizon_01() throws IOException {
        ToWorkBook workBook = new ToWorkBook(WorkBookType.HSSF);
        ToWorkSheet sheet = workBook.createSheet().updateDirection(SheetDirection.HORIZON);

        sheet.createTitleCell(2, "이름", "나이");
        sheet.merge(2, 3);
        workBook.writeFile("target/excel/manual/merge/merge_horizon_01.xls");
    }

    @Test
    public void mergeCellTest_vertical_01() throws IOException {
        ToWorkBook workBook = new ToWorkBook(WorkBookType.XSSF);
        ToWorkSheet sheet = workBook.createSheet().updateDirection(SheetDirection.VERTICAL);

        sheet.createTitleCell(2, "이름", "나이");
        sheet.merge(2, 3);

        workBook.writeFile("target/excel/manual/merge/merge_vertical_01.xls");
    }

}

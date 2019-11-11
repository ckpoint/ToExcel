package com.github.ckpoint.toexcel.workbook.manual;

import com.github.ckpoint.toexcel.core.ToWorkBook;
import com.github.ckpoint.toexcel.core.ToWorkSheet;
import com.github.ckpoint.toexcel.core.type.SheetDirection;
import com.github.ckpoint.toexcel.core.type.ToWorkBookType;
import org.junit.Assert;
import org.junit.Test;

import java.io.IOException;

public class MergeCellTest {

    @Test
    public void mergeCellTest_horizon_01() throws IOException {
        ToWorkBook workBook = new ToWorkBook(ToWorkBookType.XSSF);
        ToWorkSheet sheet = workBook.createSheet().updateDirection(SheetDirection.HORIZON);

        sheet.createTitleCell(2, "name", "age", "contactNumber");
        sheet.merge(2, 1); // 2(width) X 1(height) [][]
        sheet.createCellToNewline("sharky", "36", "010-1234-0000", "02-1111-1234");
        sheet.createCellToNewline("melpis", "36", "010-1111-1234", "02-4221-1234");
        sheet.createCellToNewline("heeseob", "32", "010-0000-1234", "-");
        sheet.createCellToNewline("dongjun", "31", "010-4324-1234", "031-4121-1234");

        Assert.assertTrue(sheet.getMergedRegions().get(0).getFirstRow() == 0 );
        Assert.assertTrue(sheet.getMergedRegions().get(0).getLastRow() == 0);
        Assert.assertTrue(sheet.getMergedRegions().get(0).getFirstColumn() == 2);
        Assert.assertTrue(sheet.getMergedRegions().get(0).getLastColumn() == 3);


        workBook.write("target/excel/manual/merge/merge_horizon_01");
    }

    @Test
    public void mergeCellTest_horizon_02() throws IOException {
        ToWorkBook workBook = new ToWorkBook(ToWorkBookType.XSSF);
        ToWorkSheet sheet = workBook.createSheet().updateDirection(SheetDirection.HORIZON);

        sheet.createTitleCell(2, "name");
        sheet.merge(1, 2);

        sheet.createTitleCell(2, "age");
        sheet.merge(1, 2);

        sheet.createTitleCell(2, "contactNumber");
        sheet.merge(2, 1);
        sheet.newLine();

        sheet.createTitleCell(2,"phone", "home");

        sheet.createCellToNewline("sharky", "36", "010-1234-0000", "02-1111-1234");
        sheet.createCellToNewline("melpis", "36", "010-1111-1234", "02-4221-1234");
        sheet.createCellToNewline("heeseob", "32", "010-0000-1234", "-");
        sheet.createCellToNewline("dongjun", "31", "010-4324-1234", "031-4121-1234");

        Assert.assertTrue(sheet.getMergedRegions().size() == 3);

        Assert.assertTrue(sheet.getMergedRegions().get(0).getFirstRow() == 0 && sheet.getMergedRegions().get(0).getLastRow() == 1);
        Assert.assertTrue(sheet.getMergedRegions().get(0).getFirstColumn() == 0 && sheet.getMergedRegions().get(0).getLastColumn() == 0);
        Assert.assertTrue(sheet.getMergedRegions().get(1).getFirstRow() == 0 && sheet.getMergedRegions().get(1).getLastRow() == 1);
        Assert.assertTrue(sheet.getMergedRegions().get(1).getFirstColumn() == 1 && sheet.getMergedRegions().get(1).getLastColumn() == 1);

        Assert.assertTrue(sheet.getMergedRegions().get(2).getFirstRow() == 0 && sheet.getMergedRegions().get(2).getLastRow() == 0);
        Assert.assertTrue(sheet.getMergedRegions().get(2).getFirstColumn() == 2 && sheet.getMergedRegions().get(2).getLastColumn() == 3);

        workBook.write("target/excel/manual/merge/merge_horizon_02");
    }

}

package com.github.ckpoint.toexcel.util;

import com.github.ckpoint.toexcel.annotation.ExcelHeader;
import org.apache.poi.hssf.record.formula.functions.T;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

public interface TitleRowHelper extends ExcelHeaderHelper {


    default Row findTitleRow(Class type, HSSFSheet sheet) {

        Field[] fields = type.getDeclaredFields();
        List<String> titles = new ArrayList<>();
        Arrays.stream(fields).map(fd -> fd.getAnnotation(ExcelHeader.class)).filter(Objects::nonNull).map(this::headerList).forEach(titles::addAll);

        int maxSearchRow = sheet.getLastRowNum() < 100 ? sheet.getLastRowNum() : 100;
        int maxCnt = 0;

        Row titleRow = null;
        for (int i = 0; i < maxSearchRow; i++) {
            Row tmpRow = sheet.getRow(i);
            int cnt = countMatchTitles(tmpRow, titles);
            if (cnt > maxCnt) {
                titleRow = tmpRow;
            }
        }
        return titleRow;
    }

    default int countMatchTitles(Row row, List<String> titles) {
        int count = 0;
        if( row == null){ return 0; }
        for (short i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if( cell == null){ continue; }
            String cellStr = cell.getStringCellValue();
            if(cellStr == null ){ continue; }
            count += titles.contains(cellStr.trim()) ? 1 : 0;
        }
        return count;
    }
}

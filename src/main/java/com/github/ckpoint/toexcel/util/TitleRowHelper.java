package com.github.ckpoint.toexcel.util;

import com.github.ckpoint.toexcel.annotation.ExcelHeader;
import com.github.ckpoint.toexcel.core.converter.ExcelHeaderConverter;
import lombok.NonNull;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/**
 * The interface Title row helper.
 */
public interface TitleRowHelper extends ExcelHeaderHelper {


    /**
     * Find title row row.
     *
     * @param type  the type
     * @param sheet the sheet
     * @return the row
     */
    default Row findTitleRow(@NonNull Class type, @NonNull Sheet sheet, @NonNull ExcelHeaderConverter excelHeaderConverter) {

        Field[] fields = type.getDeclaredFields();
        List<String> titles = new ArrayList<>();
        Arrays.stream(fields).map(fd -> fd.getAnnotation(ExcelHeader.class)).filter(Objects::nonNull).map( fd -> headerList(fd, excelHeaderConverter)).forEach(titles::addAll);

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
        return titleRow == null ? sheet.getRow(0) : titleRow;
    }

    /**
     * Count match titles int.
     *
     * @param row    the row
     * @param titles the titles
     * @return the int
     */
    default int countMatchTitles(Row row, List<String> titles) {
        int count = 0;
        if (row == null) {
            return 0;
        }
        for (short i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell == null) {
                continue;
            }
            String cellStr = cell.getStringCellValue();
            if (cellStr == null) {
                continue;
            }
            count += titles.contains(cellStr.trim()) ? 1 : 0;
        }
        return count;
    }
}

package com.github.ckpoint.toexcel.util;

import com.github.ckpoint.toexcel.annotation.ExcelHeader;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * The interface Excel header helper.
 */
public interface ExcelHeaderHelper {

    /**
     * Gets cell style.
     *
     * @param workbook  the workbook
     * @param option    the option
     * @param fieldType the field type
     * @return the cell style
     */
    default CellStyle getCellStyle(Workbook workbook, ExcelHeader option, Class fieldType) {

        CellStyle style = workbook.createCellStyle();

        if (option != null) {
            style.setAlignment((short) option.alignment().ordinal());
            style.setVerticalAlignment((short) option.verticalAlignment().ordinal());
        } else {
            style.setAlignment(CellStyle.ALIGN_CENTER);
            style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        }

        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setBorderBottom(CellStyle.BORDER_THIN);

        style.setDataFormat(option.format() >= 0 ? option.format() : TypeExcelFormatMap.findFormat(fieldType));

        return style;
    }

    /**
     * Header list list.
     *
     * @param header the header
     * @return the list
     */
    default List<String> headerList(ExcelHeader header) {
        List<String> headers = new ArrayList<>();
        headers.add(header.headerName());
        headers.addAll(Arrays.asList(header.headerNames()));

        return headers.stream().filter(Objects::nonNull).map(String::trim).filter(hd -> !hd.isEmpty()).collect(Collectors.toList());
    }

}

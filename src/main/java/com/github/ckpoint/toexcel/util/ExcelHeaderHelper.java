package com.github.ckpoint.toexcel.util;

import com.github.ckpoint.toexcel.annotation.ExcelHeader;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

public interface ExcelHeaderHelper {

    default HSSFCellStyle getCellStyle(HSSFWorkbook workbook, ExcelHeader option, Class fieldType) {

        HSSFCellStyle style = workbook.createCellStyle();

        if (option != null) {
            style.setAlignment((short) option.alignment().ordinal());
            style.setVerticalAlignment((short) option.verticalAlignment().ordinal());
        } else {
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        }

        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);

        style.setDataFormat(option.format() >= 0 ? option.format() : TypeExcelFormatMap.findFormat(fieldType));

        return style;
    }

    default List<String> headerList(ExcelHeader header) {
        List<String> headers = new ArrayList<>();
        headers.add(header.headerName());
        headers.addAll(Arrays.asList(header.headerNames()));

        return headers.stream().filter(Objects::nonNull).map(String::trim).filter(hd -> !hd.isEmpty()).collect(Collectors.toList());
    }

}

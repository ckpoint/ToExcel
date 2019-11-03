package com.github.ckpoint.toexcel.core.converter;

import com.github.ckpoint.toexcel.annotation.ExcelHeader;

/**
 * The interface Excel header converter.
 */
public interface ExcelHeaderConverter {

    /**
     * Header key converter string.
     *
     * @param excelHeader the excel header
     * @return the string
     */
    String headerKeyConverter(ExcelHeader excelHeader);
}

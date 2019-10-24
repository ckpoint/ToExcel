package com.github.ckpoint.toexcel.core.converter;

import com.github.ckpoint.toexcel.annotation.ExcelHeader;

public interface ExcelHeaderConverter {

    String headerKeyConverter(ExcelHeader excelHeader);
}

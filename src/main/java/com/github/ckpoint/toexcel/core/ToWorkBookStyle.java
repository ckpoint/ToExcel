package com.github.ckpoint.toexcel.core;

import lombok.EqualsAndHashCode;
import lombok.Getter;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.Calendar;
import java.util.Date;

@Getter
@EqualsAndHashCode
public class ToWorkBookStyle {

    private HorizontalAlignment alignment = HorizontalAlignment.CENTER;
    private VerticalAlignment verticalAlignment = VerticalAlignment.CENTER;
    private BorderStyle borderTop = BorderStyle.THIN;
    private BorderStyle borderBottom = BorderStyle.THIN;
    private BorderStyle borderLeft = BorderStyle.THIN;
    private BorderStyle borderRight = BorderStyle.THIN;

    private short format = 0x0;


    public ToWorkBookStyle(Object obj) {
        if (obj == null) {
            return;
        }
        if (obj instanceof Long || obj instanceof Integer || obj instanceof Short) {
            updateNumberType();
        } else if (obj instanceof Double || obj instanceof Float) {
            updateNumberType();
        } else if (obj instanceof Date || obj instanceof Calendar) {
            updateDateType();
        } else {
            return;
        }
    }

    public ToWorkBookStyle updateNumberType() {
        updateformat("#,##0");
        this.alignment = HorizontalAlignment.RIGHT;
        return this;
    }

    public ToWorkBookStyle updateDateType() {
        this.format = 0xe;
        return this;
    }

    public ToWorkBookStyle updateTitleType() {
        return this;
    }


    public void updateformat(String format) {
        this.format = HSSFDataFormat.getBuiltinFormat(format);
    }

    public void updateformat(short format) {
        this.format = format;
    }

    public CellStyle convertHssfStyle(HSSFWorkbook wb) {
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment((short) this.alignment.ordinal());
        cellStyle.setVerticalAlignment((short) this.verticalAlignment.ordinal());
        cellStyle.setBorderTop((short) this.borderTop.ordinal());
        cellStyle.setBorderBottom((short) this.borderBottom.ordinal());
        cellStyle.setBorderLeft((short) this.borderLeft.ordinal());
        cellStyle.setBorderRight((short) this.borderRight.ordinal());
        cellStyle.setDataFormat(this.format);
        return cellStyle;
    }

}

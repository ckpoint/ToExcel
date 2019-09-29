package com.github.ckpoint.toexcel.core.style;

import lombok.EqualsAndHashCode;
import lombok.Getter;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

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

    private Short fillPattern;
    private IndexedColors fillForegroundColor;

    private ToWorkBookFont font;

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
        this.fillForegroundColor = IndexedColors.GREY_25_PERCENT;
        this.fillPattern = CellStyle.SOLID_FOREGROUND;

        this.font = ToWorkBookFont.builder().boldWeight(Font.BOLDWEIGHT_BOLD).fontHeightInPoints((short) 8).build();
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

        if (this.fillForegroundColor != null) {
            cellStyle.setFillForegroundColor(this.fillForegroundColor.getIndex());
        }
        if (this.fillPattern != null) {
            cellStyle.setFillPattern(this.fillPattern);
        }
        if(this.font != null){
            cellStyle.setFont(this.font.convertFont(wb));
        }

        return cellStyle;
    }

}

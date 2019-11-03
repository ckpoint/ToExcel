package com.github.ckpoint.toexcel.core.model;

import com.github.ckpoint.toexcel.annotation.ExcelHeader;
import com.github.ckpoint.toexcel.core.converter.ExcelHeaderConverter;
import com.github.ckpoint.toexcel.util.ExcelHeaderHelper;
import lombok.Getter;
import lombok.NonNull;

import java.lang.reflect.Field;
import java.util.List;

/**
 * The type To title key.
 */
@Getter
public class ToTitleKey implements ExcelHeaderHelper, Comparable<ToTitleKey> {

    private String key;
    private String viewName;
    private ExcelHeader header;
    private Field field;
    private int priority;

    /**
     * Instantiates a new To title key.
     *
     * @param field                the field
     * @param fieldIdx             the field idx
     * @param excelHeaderConverter the excel header converter
     */
    public ToTitleKey(@NonNull Field field, int fieldIdx, ExcelHeaderConverter excelHeaderConverter) {
        this.field = field;
        this.key = field.getName();
        this.header = field.getAnnotation(ExcelHeader.class);
        this.viewName = excelHeaderConverter.headerKeyConverter(this.header);
        this.priority = fieldIdx + (priority * 1000);
    }

    /**
     * Instantiates a new To title key.
     *
     * @param field                the field
     * @param titles               the titles
     * @param excelHeaderConverter the excel header converter
     */
    public ToTitleKey(Field field, List<String> titles, ExcelHeaderConverter excelHeaderConverter) {
        this.field = field;
        this.key = field.getName();
        ExcelHeader header = field.getAnnotation(ExcelHeader.class);
        List<String> headerStrs = headerList(header, excelHeaderConverter);
        this.viewName = headerStrs.stream().filter(titles::contains).findFirst().orElse(key);
    }

    /**
     * Is my name boolean.
     *
     * @param title the title
     * @return the boolean
     */
    public boolean isMyName(String title) {
        if (title == null) {
            return false;
        }

        return title.trim().equalsIgnoreCase(viewName.trim());
    }


    @Override
    public int compareTo(ToTitleKey o) {

        if (o.priority != this.priority) {
            return priority - o.priority;
        }
        return this.viewName.compareTo(o.viewName);
    }
}

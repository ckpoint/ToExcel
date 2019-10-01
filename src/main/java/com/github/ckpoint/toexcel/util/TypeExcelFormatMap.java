package com.github.ckpoint.toexcel.util;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;

import java.util.*;

/**
 * The enum Type excel format map.
 */
public enum TypeExcelFormatMap {

    /**
     * The String.
     */
    STRING(new HashSet<>(Collections.singletonList(String.class)), (short) 0),
    /**
     * The Number.
     */
    NUMBER(new HashSet<>(Arrays.asList(Long.class, long.class, Integer.class, Integer.class, Short.class, short.class, Double.class, double.class, Float.class, float.class)),
            HSSFDataFormat.getBuiltinFormat("#,##0")),
    /**
     * The Date.
     */
    DATE(new HashSet<>(Collections.singletonList(Date.class)), (short) 0xe);

    private final Set<Class> sourceFieldTypes;
    private final short targetFormat;

    TypeExcelFormatMap(Set<Class> sourceFieldTypes, short targetFormat) {
        this.sourceFieldTypes = sourceFieldTypes;
        this.targetFormat = targetFormat;
    }

    /**
     * Find format short.
     *
     * @param fieldType the field type
     * @return the short
     */
    public static short findFormat(Class fieldType) {
        return Arrays.stream(TypeExcelFormatMap.values()).filter(map -> map.sourceFieldTypes.contains(fieldType)).findFirst().orElse(STRING).targetFormat;
    }


}

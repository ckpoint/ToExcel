package com.github.ckpoint.toexcel.util.mapper;

import com.github.ckpoint.toexcel.util.DateConverter;
import org.modelmapper.AbstractConverter;

import java.util.Date;

/**
 * The type String to date converter.
 */
public class StringToDateConverter extends AbstractConverter<String, Date> implements DateConverter {

    private final String[] formats = {
            "yyyy-MM-dd'T'HH:mm:ss.SSS",
            "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'",
            "yyyy-MM-dd'T'HH:mm:ss.SSSXXX",
            "yyyy-MM-dd'T'HH:mm:ssXXX",
            "yyyy-MM-dd'T'HH:mm:ss",
            "yyyy-MM-dd",
            "yyyyMMddZ",
            "yyyyMMdd",
    };

    @Override
    protected Date convert(String s) {
        if (s == null || s.isEmpty()) {
            return null;
        }
        for (String format : formats) {
            Date convertDate = convert(s, format);
            if (convertDate != null) {
                return convertDate;
            }
        }
        return null;
    }
}

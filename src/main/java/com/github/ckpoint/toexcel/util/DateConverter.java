package com.github.ckpoint.toexcel.util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * The interface Date converter.
 */
public interface DateConverter {

    /**
     * Convert date.
     *
     * @param dateStr the date str
     * @param format  the format
     * @return the date
     */
    default Date convert(String dateStr, String format) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(format);
        try {
            return simpleDateFormat.parse(dateStr);
        } catch (ParseException e) {
            return null;
        }
    }
}

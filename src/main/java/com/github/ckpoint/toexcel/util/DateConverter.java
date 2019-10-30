package com.github.ckpoint.toexcel.util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public interface DateConverter {

    default Date convert(String dateStr, String format) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(format);
        try {
            return simpleDateFormat.parse(dateStr);
        } catch (ParseException e) {
            return null;
        }
    }
}

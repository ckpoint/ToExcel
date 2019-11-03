package com.github.ckpoint.toexcel.util.mapper;

import com.github.ckpoint.toexcel.util.DateConverter;
import org.modelmapper.AbstractConverter;

import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;

/**
 * The type Double to date converter.
 */
public class DoubleToDateConverter extends AbstractConverter<Double, Date> implements DateConverter {

    @Override
    protected Date convert(Double numberDate) {
        GregorianCalendar geGregorianCalendar = new GregorianCalendar(1900, Calendar.JANUARY, 1);
        geGregorianCalendar.add(Calendar.DATE, numberDate.intValue() - 1);
        return geGregorianCalendar.getTime();
    }
}

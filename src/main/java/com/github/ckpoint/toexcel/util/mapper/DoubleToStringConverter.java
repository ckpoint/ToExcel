package com.github.ckpoint.toexcel.util.mapper;

import org.modelmapper.AbstractConverter;

/**
 * The type Double to string converter.
 */
public class DoubleToStringConverter extends AbstractConverter<Double, String> {

    @Override
    protected String convert(Double dv) {
        if (dv.longValue() == dv) {
            return String.valueOf(dv.longValue());
        }

        return String.valueOf(dv);
    }
}

package com.github.ckpoint.toexcel.util.mapper;

import org.modelmapper.AbstractConverter;

public class DoubleToStringConverter extends AbstractConverter<Double, String> {

    @Override
    protected String convert(Double dv) {
        if (dv.longValue() == dv) {
            return String.valueOf(dv.longValue());
        }

        return String.valueOf(dv);
    }
}

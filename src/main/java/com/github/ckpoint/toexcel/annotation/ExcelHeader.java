package com.github.ckpoint.toexcel.annotation;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.*;


@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelHeader {

    String headerName();

    String[] headerNames() default {""};

    short format() default -1;

    HorizontalAlignment alignment() default HorizontalAlignment.CENTER;

    VerticalAlignment verticalAlignment() default VerticalAlignment.CENTER;

}

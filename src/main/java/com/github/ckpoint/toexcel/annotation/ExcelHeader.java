package com.github.ckpoint.toexcel.annotation;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.*;


/**
 * The interface Excel header.
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelHeader {

    /**
     * Header name string.
     *
     * @return the string
     */
    String headerName();

    /**
     * Header names string [ ].
     *
     * @return the string [ ]
     */
    String[] headerNames() default {""};

    /**
     * Priority int.
     *
     * @return the int
     */
    int priority() default 0;

    /**
     * Format short.
     *
     * @return the short
     */
    short format() default -1;

    /**
     * Alignment horizontal alignment.
     *
     * @return the horizontal alignment
     */
    HorizontalAlignment alignment() default HorizontalAlignment.CENTER;

    /**
     * Vertical alignment vertical alignment.
     *
     * @return the vertical alignment
     */
    VerticalAlignment verticalAlignment() default VerticalAlignment.CENTER;

}

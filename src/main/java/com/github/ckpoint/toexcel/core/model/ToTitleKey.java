package com.github.ckpoint.toexcel.core.model;

import com.github.ckpoint.toexcel.annotation.ExcelHeader;
import com.github.ckpoint.toexcel.util.ExcelHeaderHelper;
import lombok.Getter;

import java.lang.reflect.Field;
import java.util.List;

@Getter
public class ToTitleKey implements ExcelHeaderHelper {

    private String key;
    private String viewName;

    public ToTitleKey(Field field, List<String> titles) {
        this.key = field.getName();
        ExcelHeader header = field.getAnnotation(ExcelHeader.class);
        List<String> headerStrs = headerList(header);
        this.viewName = headerStrs.stream().filter(titles::contains).findFirst().orElse(key);
    }

    public boolean isMyName(String title){
        if( title == null){
            return false;
        }

        return title.trim().equalsIgnoreCase(viewName.trim());
    }
}

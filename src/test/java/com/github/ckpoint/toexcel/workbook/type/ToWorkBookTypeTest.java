package com.github.ckpoint.toexcel.workbook.type;

import com.github.ckpoint.toexcel.core.type.ToWorkBookType;
import org.junit.Assert;
import org.junit.Test;

public class ToWorkBookTypeTest {

    @Test
    public void translate_hssf_ext_test(){
        String ext = ToWorkBookType.HSSF.translateFileName("user.xl");
        Assert.assertTrue(ext.equals("user.xls"));
    }

    @Test
    public void translate_xssf_ext_test(){
        String ext = ToWorkBookType.XSSF.translateFileName("user.xl");
        Assert.assertTrue(ext.equals("user.xlsx"));
    }
}

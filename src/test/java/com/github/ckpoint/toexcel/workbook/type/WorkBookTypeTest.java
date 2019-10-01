package com.github.ckpoint.toexcel.workbook.type;

import com.github.ckpoint.toexcel.core.type.WorkBookType;
import org.junit.Assert;
import org.junit.Test;

public class WorkBookTypeTest {

    @Test
    public void translate_hssf_ext_test(){
        String ext = WorkBookType.HSSF.translateFileName("user.xl");
        Assert.assertTrue(ext.equals("user.xls"));
    }

    @Test
    public void translate_xssf_ext_test(){
        String ext = WorkBookType.XSSF.translateFileName("user.xl");
        Assert.assertTrue(ext.equals("user.xlsx"));
    }
}

package com.github.ckpoint.toexcel.core.type;

public enum ToWorkCellType {

    TITLE, VALUE;

    public boolean isTitle(){
        return TITLE.equals(this);
    }
}

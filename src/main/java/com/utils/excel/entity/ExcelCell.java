package com.utils.excel.entity;

import java.io.Serializable;

public class ExcelCell implements Serializable {
    private Object cell;

    public Object getCell() {
    
        return cell;
    }

    public void setCell(Object cell) {
    
        this.cell = cell;
    }
    

}

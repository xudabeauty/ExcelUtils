package com.utils.excel.entity;

import java.io.Serializable;
import java.util.List;

public class ExcelSheet implements Serializable {
    private String name;
    private List<ExcelRow> rows;
    public String getName() {
    
        return name;
    }
    public void setName(String name) {
    
        this.name = name;
    }
    public List<ExcelRow> getRows() {
    
        return rows;
    }
    public void setRows(List<ExcelRow> rows) {
    
        this.rows = rows;
    }
    
    

}

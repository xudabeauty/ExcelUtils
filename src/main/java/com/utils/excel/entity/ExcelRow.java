package com.utils.excel.entity;

import java.io.Serializable;
import java.util.List;

public class ExcelRow implements Serializable{
private List<ExcelCell> cells;

public List<ExcelCell> getCells() {

    return cells;
}

public void setCells(List<ExcelCell> cells) {

    this.cells = cells;
}

    
}

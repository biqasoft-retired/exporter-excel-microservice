/*
 * Copyright (c) 2016. com.biqasoft
 */

package com.biqasoft.exporter.excel.dataobject;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * Created by Nikita Bakaev, ya@nbakaev.ru on 3/21/2016.
 * All Rights Reserved
 */
public class CellProcessingInterceptor {
    private int rowNum;
    private int colNum;
    private Object object;
    private Cell cell;
    private Row row;

    public Row getRow() {
        return row;
    }

    public void setRow(Row row) {
        this.row = row;
    }

    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public int getColNum() {
        return colNum;
    }

    public void setColNum(int colNum) {
        this.colNum = colNum;
    }

    public Object getObject() {
        return object;
    }

    public void setObject(Object object) {
        this.object = object;
    }

    public Cell getCell() {
        return cell;
    }

    public void setCell(Cell cell) {
        this.cell = cell;
    }
}

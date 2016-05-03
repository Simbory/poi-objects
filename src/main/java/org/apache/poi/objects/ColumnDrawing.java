package org.apache.poi.objects;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;

public class ColumnDrawing {

    public String getColumnName() {
        return columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }

    public int getColumnIndex() {
        return columnIndex;
    }

    public void setColumnIndex(int columnIndex) {
        this.columnIndex = columnIndex;
    }

    public int getColumnWidth() {
        return columnWidth;
    }

    public void setColumnWidth(int columnWidth) {
        this.columnWidth = columnWidth;
    }

    public boolean isHasAlternate() {
        return hasAlternate;
    }

    public void setHasAlternate(boolean hasAlternate) {
        this.hasAlternate = hasAlternate;
    }

    public void setHeaderStyle(CellStyle headerStyle) {
        this.headerStyle = headerStyle;
    }

    public CellStyle getHeaderStyle() {
        return headerStyle;
    }

    private String columnName;

    private int columnIndex;

    private int columnWidth;

    private boolean hasAlternate;

    private CellStyle headerStyle;
}
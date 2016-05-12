package org.apache.poi.objects;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;

import java.lang.reflect.Method;

public class ColumnDrawing {

    private Font cellFont;

    private String columnName;

    private int columnIndex;

    private int columnWidth;

    private boolean alternate;

    private CellStyle headerStyle;

    private CellStyle cellStyle;
    private CellStyle alternateCellStyle;
    private Font alternateCellFont;
    private Method method;

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

    public boolean getAlternate() {
        return alternate;
    }

    public void setAlternate(boolean alternate) {
        this.alternate = alternate;
    }

    public void setHeaderStyle(CellStyle headerStyle) {
        this.headerStyle = headerStyle;
    }

    public CellStyle getHeaderStyle() {
        return headerStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellFont(Font cellFont) {
        this.cellFont = cellFont;
    }

    public Font getCellFont() {
        return cellFont;
    }

    public void setAlternateCellStyle(CellStyle alternateCellStyle) {
        this.alternateCellStyle = alternateCellStyle;
    }

    public CellStyle getAlternateCellStyle() {
        return alternateCellStyle;
    }

    public void setAlternateCellFont(Font alternateCellFont) {
        this.alternateCellFont = alternateCellFont;
    }

    public Font getAlternateCellFont() {
        return alternateCellFont;
    }

    public void setMethod(Method method) {
        this.method = method;
    }

    public Method getMethod() {
        return method;
    }
}
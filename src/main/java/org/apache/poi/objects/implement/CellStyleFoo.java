package org.apache.poi.objects.implement;

import org.apache.poi.objects.AlternateCell;
import org.apache.poi.objects.Cell;
import org.apache.poi.objects.HeaderCell;

public class CellStyleFoo {
    public int height;

    public String textColor;

    public String backgroundColor;

    public String foregroundColor;

    public short textAlign;

    public short verticalAlign;

    public short fillPattern;

    public short fontWeight;

    public String fontFamily;

    public short fontSize;

    public boolean italic;

    public CellStyleFoo(Cell style) {
        height = style.height();
        textColor = style.textColor();
        backgroundColor = style.backgroundColor();
        foregroundColor = style.foregroundColor();
        textAlign = style.textAlign();
        verticalAlign = style.verticalAlign();
        fillPattern = style.fillPattern();
        fontWeight = style.fontWeight();
        fontFamily = style.fontFamily();
        fontSize = style.fontSize();
        italic = style.italic();
    }

    public CellStyleFoo(AlternateCell style) {
        height = style.height();
        textColor = style.textColor();
        backgroundColor = style.backgroundColor();
        foregroundColor = style.foregroundColor();
        textAlign = style.textAlign();
        verticalAlign = style.verticalAlign();
        fillPattern = style.fillPattern();
        fontWeight = style.fontWeight();
        fontFamily = style.fontFamily();
        fontSize = style.fontSize();
        italic = style.italic();
    }

    public CellStyleFoo(HeaderCell style) {
        height = style.height();
        textColor = style.textColor();
        backgroundColor = style.backgroundColor();
        foregroundColor = style.foregroundColor();
        textAlign = style.textAlign();
        verticalAlign = style.verticalAlign();
        fillPattern = style.fillPattern();
        fontWeight = style.fontWeight();
        fontFamily = style.fontFamily();
        fontSize = style.fontSize();
        italic = style.italic();
    }
}

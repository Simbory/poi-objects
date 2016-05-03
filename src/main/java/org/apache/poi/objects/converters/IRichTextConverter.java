package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.Cell;

public interface IRichTextConverter {
    public String convert(Cell cell);
}
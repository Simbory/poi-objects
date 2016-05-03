package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.Cell;

public interface IBooleanConverter {
    public boolean convert(Cell cell);
}
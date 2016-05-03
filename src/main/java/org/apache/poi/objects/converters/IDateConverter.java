package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.Cell;

import java.util.Date;

public interface IDateConverter {
    public Date convert(Cell cell);
}

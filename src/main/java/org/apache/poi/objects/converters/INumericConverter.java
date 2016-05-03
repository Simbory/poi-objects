package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.Cell;

public interface INumericConverter {
    public double convert(Cell cell);
}

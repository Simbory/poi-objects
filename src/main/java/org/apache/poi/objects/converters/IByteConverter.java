package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.Cell;

public interface IByteConverter {
    public byte convert(Cell cell);
}
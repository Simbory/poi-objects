package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.Cell;

public interface ICharConverter {
    public char convert(Cell cell);
}
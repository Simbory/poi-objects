package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.Cell;

public interface IUnknownTypeConverter {
    public Object convert(Cell cell, Class<?> classType);
}
package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.Cell;

public interface IUnknowTypeConverter {
    public Object convert(Cell cell, Class<?> classType);
}
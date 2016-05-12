package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.Cell;

public class DefaultUnknownTypeConverter implements IUnknownTypeConverter {
    public Object convert(Cell cell, Class<?> classType) {
        return null;
    }
}
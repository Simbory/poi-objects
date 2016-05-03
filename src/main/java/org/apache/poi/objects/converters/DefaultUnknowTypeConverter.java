package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.Cell;

public class DefaultUnknowTypeConverter implements IUnknowTypeConverter {
    public Object convert(Cell cell, Class<?> classType) {
        return null;
    }
}
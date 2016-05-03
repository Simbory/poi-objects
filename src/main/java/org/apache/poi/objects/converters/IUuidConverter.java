package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.Cell;

import java.util.UUID;

public interface IUuidConverter {
    public UUID convert(Cell cell);
}

package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.Cell;

import java.util.UUID;

public class DefaultUuidConverter implements IUuidConverter {
    public UUID convert(Cell cell) {
        try {
            if (cell.getCellType() == Cell.CELL_TYPE_STRING){
                String str = cell.getStringCellValue();
                if (str != null && str.length() > 0) {
                    return UUID.fromString(str);
                }
            }
            return null;
        } catch (Exception ex) {
            return null;
        }
    }
}
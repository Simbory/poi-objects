package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.Cell;

public class DefaultCharConverter implements ICharConverter {
    public char convert(Cell cell) {
        try{
            if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                String str = cell.getStringCellValue();
                return str != null && str.length() > 0 ? str.charAt(0) : 0;
            }
            return 0;
        } catch (Exception ex) {
            return 0;
        }
    }
}
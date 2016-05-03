package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.Cell;

public class DefaultByteConverter implements IByteConverter {
    public byte convert(Cell cell) {
        try{
            return cell.getErrorCellValue();
        } catch (Exception ex) {
            return 0;
        }
    }
}

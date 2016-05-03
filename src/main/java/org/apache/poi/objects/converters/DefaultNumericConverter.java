package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.Cell;

public class DefaultNumericConverter implements INumericConverter {
    public double convert(Cell cell) {
        try{
            return cell.getNumericCellValue();
        } catch (Exception ex) {
            return 0;
        }
    }
}
package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.Cell;

public class DefaultBooleanConverter implements IBooleanConverter {
    public boolean convert(Cell cell) {
        try{
            return cell.getBooleanCellValue();
        } catch (Exception ex){
            return false;
        }
    }
}
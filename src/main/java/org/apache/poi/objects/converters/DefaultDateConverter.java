package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.Cell;

import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;

public class DefaultDateConverter implements IDateConverter {
    public Date convert(Cell cell) {
        try{
            return cell.getDateCellValue();
        } catch (Exception ex){
            Calendar cal = new GregorianCalendar(0,0,0);
            return cal.getTime();
        }
    }
}
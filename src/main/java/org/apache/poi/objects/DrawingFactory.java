package org.apache.poi.objects;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.objects.implement.CellStyleFoo;
import org.apache.poi.objects.implement.POIObjectFoo;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.Color;
import java.io.*;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public abstract class DrawingFactory<T> {

    protected Workbook workbook;

    protected OutputStream outputStream;

    protected boolean isOutStream;

    protected boolean useTemplate;

    protected String excelPath;

    protected ExcelType excelType;

    protected Class<T> entityClass;

    public Workbook getWorkbook() {
        return workbook;
    }

    public OutputStream getOutputStream() {
        return outputStream;
    }

    public boolean isOutStream() {
        return isOutStream;
    }

    public boolean isUseTemplate() {
        return useTemplate;
    }

    public String getExcelPath() {
        return excelPath;
    }

    public ExcelType getExcelType() {
        return excelType;
    }

    public DrawingFactory(String path) throws IllegalArgumentException, IOException {
        if (path == null || path.length() < 1) {
            throw new IllegalArgumentException("the argument \"path\" cannot be empty.");
        }
        excelPath = path;
        String ext = Utils.getFileExtension(path);
        if (!".xls".equals(ext) && !".xlsx".equals(ext)) {
            throw new FileNotFoundException("the file extension is invalid");
        }
        String dir = Utils.getDirectory(path);
        if (dir!= null && dir.length() > 0 && !Utils.dirExists(dir)) {
            Utils.createDir(dir);
        }
        File f = new File(path);
        outputStream = new FileOutputStream(f);
        if (".xls".equals(ext)) {
            workbook = new HSSFWorkbook();
            excelType = ExcelType.EXCEL_2003;
            isOutStream = false;
        } else if (".xlsx".equals(ext)) {
            workbook = new XSSFWorkbook();
            excelType = ExcelType.EXCEL_2007;
            isOutStream = false;
        }
    }

    public DrawingFactory(OutputStream stream, ExcelType excelType) {
        if (stream == null) {
            throw new IllegalArgumentException("the stream cannot be empty");
        }
        outputStream = stream;
        this.excelType = excelType;
        if (excelType == ExcelType.EXCEL_2003) {
            workbook = new HSSFWorkbook();
        } else {
            workbook = new XSSFWorkbook();
        }
        isOutStream = true;
    }

    public DrawingFactory(OutputStream stream) {
        this(stream, ExcelType.EXCEL_2003);
    }

    public void draw(String sheetName, T[] objects){
        if (objects == null || objects.length < 1) {
            throw new IllegalArgumentException("objects cannot be empty");
        }
        Class<T> type = getEntryClass();
        POIObject objAttr = type.getAnnotation(POIObject.class);
        if (objAttr == null) {
            objAttr = new POIObjectFoo();
        }
        ColumnDrawing[] drawings;
        try {
            drawings = getColumnDrawings(type);
        } catch (Exception e) {
            e.printStackTrace();
            return;
        }
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            sheet = workbook.createSheet(sheetName);
            drawHeader(drawings, sheet, objAttr.headerRowIndex());
        }
        for (int i = 0; i < objects.length; i++) {
            T obj = objects[i];
            drawRow(drawings, sheet, objAttr.startIndex() + i, obj);
        }
    }

    protected Class<T> getEntryClass() {
        Class<T> entityClass = null;
        Type t = getClass().getGenericSuperclass();
        if(t instanceof ParameterizedType){
            Type[] p = ((ParameterizedType)t).getActualTypeArguments();
            entityClass = (Class<T>)p[0];
        }
        return entityClass;
    }

    public void draw(String sheetName, List<T> list){
        if (list == null || list.size() < 1) {
            throw new IllegalArgumentException("objects cannot be empty");
        }
        Class<T> type = getEntryClass();
        POIObject objAttr = type.getAnnotation(POIObject.class);
        if (objAttr == null) {
            objAttr = new POIObjectFoo();
        }
        ColumnDrawing[] drawings;
        try {
            drawings = getColumnDrawings(type);
        } catch (Exception e) {
            e.printStackTrace();
            return;
        }
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            sheet = workbook.createSheet(sheetName);
            drawHeader(drawings, sheet, objAttr.headerRowIndex());
        }
        for (int i = 0; i < list.size(); i++) {
            T obj = list.get(i);
            drawRow(drawings, sheet, objAttr.startIndex() + i, obj);
        }
    }

    public void close() throws IOException {
        if (workbook != null) {
            workbook.write(outputStream);
        }
        if (!isOutStream() && outputStream != null) {
            outputStream.close();
        }
    }

    private void drawRow(ColumnDrawing[] drawings, Sheet sheet, int rowIndex, T obj) {
        Row row = sheet.createRow(rowIndex);
        for (ColumnDrawing drawing : drawings)
        {
            org.apache.poi.ss.usermodel.Cell cell = row.createCell(drawing.getColumnIndex());
            Object value = null;
            try {
                value = drawing.getMethod().invoke(obj);
            } catch (Exception e) {
                e.printStackTrace();
                value = null;
            }
            drawCellValue(cell, value);
            drawCellFontAndStyle(cell, drawing, rowIndex%2 == 1 && drawing.getAlternate());
        }
    }

    private void drawCellFontAndStyle(org.apache.poi.ss.usermodel.Cell cell, ColumnDrawing drawing, boolean alternate) {
        if (useTemplate)
            return;
        CellStyle style;
        Font font;
        if (alternate)
        {
            style = drawing.getAlternateCellStyle();
            font = drawing.getAlternateCellFont();
        }
        else
        {
            style = drawing.getCellStyle();
            font = drawing.getCellFont();
        }
        if (style == null)
            style = cell.getSheet().getWorkbook().getCellStyleAt(0);
        cell.setCellStyle(style);
        if (font != null)
            cell.getCellStyle().setFont(font);
    }

    private void drawCellValue(org.apache.poi.ss.usermodel.Cell cell, Object value) {
        if (value == null)
        {
            cell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BLANK);
            return;
        }
        if (value instanceof String)
        {
            String strValue = (String) value;
            cell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING);
            cell.setCellValue(strValue);
            return;
        }
        if (value instanceof char[])
        {
            char[] charValue = (char[]) value;
            cell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING);
            cell.setCellValue(new String(charValue));
            return;
        }
        if (value instanceof Boolean)
        {
            cell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BOOLEAN);
            cell.setCellValue((Boolean) value);
            return;
        }
        if (value instanceof Integer || value instanceof Long || value instanceof Short || value instanceof Float || value instanceof Double)
        {
            cell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC);
            cell.setCellValue(Double.parseDouble(value.toString()));
            return;
        }
        if (value instanceof Byte)
        {
            cell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_ERROR);
            cell.setCellValue((Byte)value);
            return;
        }
        if (value instanceof Date)
        {
            cell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA);
            cell.setCellValue((Date) value);
            return;
        }
        cell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING);
        cell.setCellValue(value.toString());
    }

    private void drawHeader(ColumnDrawing[] drawings, Sheet sheet, int index) {
        Row headerRow = sheet.createRow(index);
        for (ColumnDrawing drawing : drawings) {
            org.apache.poi.ss.usermodel.Cell headerCell = headerRow.createCell(drawing.getColumnIndex());
            drawHeaderFontAndStyle(headerCell, drawing);
            if (drawing.getColumnWidth() > 255) {
                drawing.setColumnWidth(255);
            }
            sheet.setColumnWidth(drawing.getColumnIndex(), drawing.getColumnWidth() * 256);
        }
    }

    private void drawHeaderFontAndStyle(org.apache.poi.ss.usermodel.Cell headerCell, ColumnDrawing drawing) {
        headerCell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING);
        headerCell.setCellValue(drawing.getColumnName());
        if (useTemplate) {
            return;
        }
        if (drawing.getHeaderStyle() != null) {
            headerCell.setCellStyle(drawing.getHeaderStyle());
        }
    }

    private ColumnDrawing[] getColumnDrawings(Class<T> entityType) throws Exception {
        ArrayList<ColumnDrawing> cellList = new ArrayList<ColumnDrawing>();
        Method[] classProperties = entityType.getMethods();
        if (classProperties == null || classProperties.length < 1) {
            return new ColumnDrawing[0];
        }
        ArrayList<Integer> columnIndexes = new ArrayList<Integer>();
        for (Method m : classProperties) {
            if (!m.getName().startsWith("get") || m.getReturnType() == Void.TYPE) {
                continue;
            }
            DrawingIgnore drawingIgnore = m.getAnnotation(DrawingIgnore.class);
            if (drawingIgnore != null && drawingIgnore.ignore()) {
                continue;
            }
            POIColumn propAttr = m.getAnnotation(POIColumn.class);
            if (propAttr == null) {
                continue;
            }

            ColumnDrawing cellInfo = new ColumnDrawing();
            if (propAttr.index() < 0) {
                throw new Exception("Column Index is out of range.\r\nThe value must not be smaller than 0.");
            }
            if (columnIndexes.contains(propAttr.index())) {
                throw new Exception("Duplicate column index " + propAttr.index());
            }
            cellInfo.setColumnIndex(propAttr.index());
            columnIndexes.add(propAttr.index());

            String name = propAttr.name();
            if (name.length() < 1) {
                name = m.getName().replaceFirst("^get", "");
            }
            cellInfo.setColumnName(name);

            HeaderCell headerStyleAttr = m.getAnnotation(HeaderCell.class);
            CellStyle headerStyle = fillCellStyle(new CellStyleFoo(headerStyleAttr));
            if (headerStyle != null) {
                cellInfo.setHeaderStyle(headerStyle);
            }
            if (headerStyleAttr != null) {
                cellInfo.setColumnWidth(headerStyleAttr.columnWidth());
            }

            Cell cellAttr = m.getAnnotation(Cell.class);
            if (cellAttr != null) {
                CellStyle cellStyle = fillCellStyle(new CellStyleFoo(cellAttr));
                if (cellStyle != null) {
                    cellInfo.setCellStyle(cellStyle);
                }
                Font cellFont = fillFont(new CellStyleFoo(cellAttr));
                if (cellFont != null) {
                    cellInfo.setCellFont(cellFont);
                }
            }

            AlternateCell alternateCellStyleAttr = m.getAnnotation(AlternateCell.class);
            if (alternateCellStyleAttr != null) {
                cellInfo.setAlternate(true);
                CellStyle cellStyle = fillCellStyle(new CellStyleFoo(alternateCellStyleAttr));
                if (cellStyle != null) {
                    cellInfo.setAlternateCellStyle(cellStyle);
                }
                Font cellFont = fillFont(new CellStyleFoo(alternateCellStyleAttr));
                if (cellFont != null) {
                    cellInfo.setAlternateCellFont(cellFont);
                }
            } else {
                cellInfo.setAlternate(false);
            }
            cellInfo.setMethod(m);
            cellList.add(cellInfo);
        }

        return cellList.toArray(new ColumnDrawing[cellList.size()]);
    }

    private CellStyle fillCellStyle(CellStyleFoo attr) {
        if (attr == null) {
            return null;
        }
        CellStyle style;
        if (getExcelType() == ExcelType.EXCEL_2003) {
            style = fillCellStyle2003(attr);
        } else {
            style = fillCellStyle2007(attr);
        }
        style.setAlignment(attr.textAlign);
        style.setVerticalAlignment(attr.verticalAlign);
        Font font = fillFont(attr);
        if (font != null) {
            style.setFont(font);
        }
        return style;
    }

    private CellStyle fillCellStyle2003(CellStyleFoo attr) {
        HSSFCellStyle style = ((HSSFWorkbook) workbook).createCellStyle();

        if (attr.backgroundColor != null && attr.backgroundColor.length() > 0) {
            Color color = Utils.convertStr2Color(attr.backgroundColor);
            if (color != null) {
                style.setFillBackgroundColor(getColor2003(color));
                style.setFillPattern(attr.fillPattern);
            }
        }
        if (attr.foregroundColor != null && attr.foregroundColor.length() > 0) {
            Color color = Utils.convertStr2Color(attr.foregroundColor);
            if (color != null) {
                style.setFillForegroundColor(getColor2003(color));
                style.setFillPattern(attr.fillPattern);
            }
        }
        return style;
    }

    private CellStyle fillCellStyle2007(CellStyleFoo attr) {
        XSSFCellStyle style = ((XSSFWorkbook) workbook).createCellStyle();

        if (attr.backgroundColor != null && attr.backgroundColor.length() > 0) {
            Color color = Utils.convertStr2Color(attr.backgroundColor);
            if (color != null) {
                style.setFillBackgroundColor(getColor2007(color));
                style.setFillPattern(attr.fillPattern);
            }
        }
        if (attr.foregroundColor != null && attr.foregroundColor.length() > 0) {
            Color color = Utils.convertStr2Color(attr.foregroundColor);
            if (color != null) {
                style.setFillForegroundColor(getColor2007(color));
                style.setFillPattern(attr.fillPattern);
            }
        }
        return style;
    }

    private short getColor2003(java.awt.Color color) {
        HSSFWorkbook book = (HSSFWorkbook) workbook;
        if (book != null) {
            HSSFPalette palette = book.getCustomPalette();
            HSSFColor workbookColor = palette.findColor((byte) color.getRed(), (byte) color.getGreen(), (byte) color.getBlue());
            if (workbookColor != null) {
                return workbookColor.getIndex();
            }
            try{
                workbookColor = palette.addColor((byte) color.getRed(), (byte) color.getGreen(), (byte) color.getBlue());
                return workbookColor.getIndex();
            } catch (Exception ex) {
                return palette.findSimilarColor(color.getRed(), color.getGreen(), color.getBlue()).getIndex();
            }
        }
        return new HSSFColor().getIndex();
    }

    private XSSFColor getColor2007(java.awt.Color color) {
        return new XSSFColor(color);
    }

    private Font fillFont(CellStyleFoo attr) {
        if (attr == null) {
            return null;
        }
        if (attr.fontWeight > 0
                || (attr.fontFamily!= null && attr.fontFamily.length() > 0)
                || attr.fontSize > 0
                || attr.italic
                || (attr.textColor != null && attr.textColor.length() > 0)) {
            Font font;
            if (excelType == ExcelType.EXCEL_2003) {
                font = fillFont2003(attr);
            } else {
                font = fillFont2007(attr);
            }
            if (attr.fontWeight > 0) {
                font.setBoldweight(attr.fontWeight);
            }
            if (attr.fontFamily!= null && attr.fontFamily.length() > 0) {
                font.setFontName(attr.fontFamily);
            }
            if (attr.fontSize > 0) {
                font.setFontHeightInPoints(attr.fontSize);
            }
            if (attr.italic) {
                font.setItalic(true);
            }
            return font;
        }
        return null;
    }

    private Font fillFont2007(CellStyleFoo attr) {
        XSSFFont font = ((XSSFWorkbook)workbook).createFont();
        if (attr.textColor != null && attr.textColor.length() > 0) {
            Color color = Utils.convertStr2Color(attr.textColor);
            if (color != null) {
                font.setColor(getColor2007(color));
            }
        }
        return font;
    }

    private Font fillFont2003(CellStyleFoo attr) {
        HSSFFont font = ((HSSFWorkbook)workbook).createFont();
        if (attr.textColor != null && attr.textColor.length() > 0) {
            Color color = Utils.convertStr2Color(attr.textColor);
            if (color != null) {
                font.setColor(getColor2003(color));
            }
        }
        return font;
    }
}
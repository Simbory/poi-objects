package org.apache.poi.objects;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.objects.implement.POIObjectFoo;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.crypto.interfaces.PBEKey;
import java.io.*;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.ArrayList;

public class DrawingFactory<T> {

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

    public Class<T> getEntityClass() {
        if (entityClass == null) {
            Type t = getClass().getGenericSuperclass();
            if(t instanceof ParameterizedType){
                Type[] p = ((ParameterizedType)t).getActualTypeArguments();
                entityClass = (Class<T>)p[0];
            }
        }
        return entityClass;
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
        if (!f.exists()) {
            f.createNewFile();
        }
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

    public void draw(int sheetIndex, String sheetName, T[] objects){
        Class<T> tClass = getEntityClass();
        POIObject objDef = tClass.getAnnotation(POIObject.class);
        if (objDef == null) {
            objDef = new POIObjectFoo();
        }
    }

    private ColumnDrawing[] getColumnDrawings(Class<T> entityType) throws Exception {
        ArrayList<ColumnDrawing> list = new ArrayList<ColumnDrawing>();

        Method[] methods = entityType.getMethods();
        ArrayList<Integer> indexList = new ArrayList<Integer>();
        if (methods != null) {
            for (int i = 0; i < methods.length; i++) {
                Method m = methods[i];
                if (!m.getName().startsWith("get") || m.getReturnType() == Void.TYPE){
                    continue;
                }
                DrawingIgnore drawingIgnore = m.getAnnotation(DrawingIgnore.class);
                if (drawingIgnore == null || drawingIgnore.ignore()) {
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
                if (indexList.contains(propAttr.index())) {
                    throw new Exception("Duplicate column index " + propAttr.index());
                }
                cellInfo.setColumnIndex(propAttr.index());
                indexList.add(propAttr.index());

                cellInfo.setColumnName(propAttr.name());

                if (propAttr.name().length() < 1) {
                    cellInfo.setColumnName(m.getName().replaceAll("^get", ""));
                }

                HeaderStyle headerStyleAttr = m.getAnnotation(HeaderStyle.class);
                CellStyle headerStyle = fillCellStyle(headerStyleAttr);
                if (headerStyle != null){
                    cellInfo.setHeaderStyle(headerStyle);
                }
            }
        }

        return (ColumnDrawing[]) list.toArray();
    }

    protected CellStyle fillCellStyle(Object attr) {
        if (attr == null) {
            return null;
        }
        Class<?> attrClass = attr.getClass();
        if (attrClass != HeaderStyle.class && attrClass != CellStyle.class && attrClass != AlternateCellStyle.class){
            return null;
        }
        CellStyle style;
        if (getExcelType() == ExcelType.EXCEL_2003) {
            style = fillCellStyle2003(attr);
        } else {
            style = fillCellStyle2007(attr);
        }
        if (attrClass == HeaderStyle.class){
            HeaderStyle value = (HeaderStyle) attr;
            style.setAlignment(value.textAlign());
            style.setVerticalAlignment(value.verticalAlign());
        }
        if (attrClass == org.apache.poi.objects.CellStyle.class) {
            org.apache.poi.objects.CellStyle value = (org.apache.poi.objects.CellStyle) attr;
            style.setAlignment(value.textAlign());
            style.setVerticalAlignment(value.verticalAlign());
        }
        if (attrClass == AlternateCellStyle.class){
            AlternateCellStyle value = (AlternateCellStyle) attr;
            style.setAlignment(value.textAlign());
            style.setVerticalAlignment(value.verticalAlign());
        }
        Font font = fillFont(attr);
        if (font != null) {
            style.setFont(font);
        }
        return style;
    }

    private Font fillFont(Object attr) {
        // todo
        return null;
    }

    private CellStyle fillCellStyle2007(Object attr) {
        // todo
        return null;
    }

    private CellStyle fillCellStyle2003(Object attr) {
        // todo
        return null;
    }
}
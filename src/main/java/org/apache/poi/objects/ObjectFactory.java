package org.apache.poi.objects;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.objects.converters.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public abstract class ObjectFactory<T> {
    protected Workbook workbook;

    protected InputStream excelStream;

    protected boolean needClose;

    public String excelPath;

    public ExcelType excelType;

    public IRichTextConverter richTextConverter;

    public IBooleanConverter booleanConverter;

    public INumericConverter numericConverter;

    public IDateConverter dateConverter;

    public IByteConverter byteConverter;

    public ICharConverter charConverter;

    public IUuidConverter uuidConverter;

    public IUnknownTypeConverter unknownTypeConverter;

    public ObjectFactory(String path) throws IOException {
        excelPath = path;
        String ext = Utils.getFileExtension(path);
        if (ext == null || ext.length() < 1 || (!ext.toLowerCase().equals(".xls") && !ext.toLowerCase().equals(".xlsx"))){
            throw new IllegalArgumentException("invalid excel path");
        }
        excelType = ext.toLowerCase().equals(".xls") ? ExcelType.EXCEL_2003 : ExcelType.EXCEL_2007;
        excelStream = new FileInputStream(path);
        init();
    }

    public ObjectFactory(InputStream stream, ExcelType excelType) throws IOException {
        this.excelType = excelType;
        excelStream = stream;
        needClose = false;
        init();
    }

    private void init() throws IOException {
        workbook = excelType == ExcelType.EXCEL_2003
                ? new HSSFWorkbook(excelStream)
                : new XSSFWorkbook(excelStream);
        needClose = true;
        richTextConverter = new DefaultRichTextConverter();
        booleanConverter = new DefaultBooleanConverter();
        numericConverter = new DefaultNumericConverter();
        dateConverter = new DefaultDateConverter();
        byteConverter = new DefaultByteConverter();
        charConverter = new DefaultCharConverter();
        uuidConverter = new DefaultUuidConverter();
        unknownTypeConverter = new DefaultUnknownTypeConverter();
    }
}
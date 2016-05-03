package org.apache.poi.objects;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface AlternateCellStyle{

    public int height() default 0;

    public String textColor() default "";

    public String backgroundColor() default "";

    public String foregroundColor() default "";

    public short textAlign() default CellStyle.ALIGN_GENERAL;

    public short verticalAlign() default CellStyle.VERTICAL_TOP;

    public FillPatternType fillPattern() default FillPatternType.SOLID_FOREGROUND;

    public short fontWeight() default 0;

    public String fontFamily() default "";

    public short fontSize() default -1;

    public boolean italic() default false;
}
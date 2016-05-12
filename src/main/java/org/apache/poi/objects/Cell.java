package org.apache.poi.objects;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Cell {

    int height() default 0;

    String textColor() default "";

    String backgroundColor() default "";

    String foregroundColor() default "";

    short textAlign() default org.apache.poi.ss.usermodel.CellStyle.ALIGN_GENERAL;

    short verticalAlign() default org.apache.poi.ss.usermodel.CellStyle.VERTICAL_TOP;

    short fillPattern() default org.apache.poi.ss.usermodel.CellStyle.SOLID_FOREGROUND;

    short fontWeight() default 0;

    String fontFamily() default "";

    short fontSize() default -1;

    boolean italic() default false;
}
package org.apache.poi.objects;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface POIObject {
    public int headerRowIndex() default 0;

    public int startIndex() default 1;

    public int endIndex() default -1;
}

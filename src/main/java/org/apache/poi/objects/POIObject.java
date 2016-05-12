package org.apache.poi.objects;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface POIObject {
    int headerRowIndex() default 0;

    int startIndex() default 1;

    int endIndex() default -1;
}
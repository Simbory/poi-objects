package org.apache.poi.objects.implement;

import org.apache.poi.objects.POIObject;

import java.lang.annotation.Annotation;

public class POIObjectFoo implements POIObject {
    public int headerRowIndex() {
        return 0;
    }

    public int startIndex() {
        return 1;
    }

    public int endIndex() {
        return -1;
    }

    public Class<? extends Annotation> annotationType() {
        return POIObject.class;
    }
}
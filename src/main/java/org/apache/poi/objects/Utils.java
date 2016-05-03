package org.apache.poi.objects;

import java.io.File;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;

final class Utils {
    static String getFileExtension(String filePath) {
        if (filePath == null || filePath.length() < 1) {
            return null;
        }
        int pos = filePath.lastIndexOf('.');
        if (pos == -1) {
            return null;
        }
        return filePath.substring(pos, filePath.length()).toLowerCase();
    }

    static String getDirectory(String filePath) {
        File f = new File(filePath);
        return f.getParent();
    }

    static boolean dirExists(String path) {
        return new File(path).exists();
    }

    static void createDir(String path) {
        new File(path).mkdirs();
    }
}
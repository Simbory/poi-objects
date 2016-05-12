package org.apache.poi.objects;

import java.awt.*;
import java.io.File;

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

    static Color convertStr2Color(String hex) {
        if (hex == null || hex.length() < 1) {
            return null;
        }
        if (hex.startsWith("#")) {
            try{
                return new Color(Integer.parseInt(hex.substring(1), 16));
            } catch (Exception ex) {
                return null;
            }
        } else {
            return Color.getColor(hex);
        }
    }
}
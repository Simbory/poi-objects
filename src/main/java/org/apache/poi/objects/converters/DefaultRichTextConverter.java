package org.apache.poi.objects.converters;

import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.Arrays;

public class DefaultRichTextConverter implements IRichTextConverter {

    private class RichStyleString {
        private Font currentFont;

        private ArrayList<Character> charList;

        Font getCurrentFont() {
            return currentFont;
        }

        void setCurrentFont(Font currentFont) {
            this.currentFont = currentFont;
        }

        ArrayList<Character> getCharList() {
            return charList;
        }

        RichStyleString() {
            charList = new ArrayList<Character>();
        }

        boolean isCurrentStyle(Font font) {
            return !(getCurrentFont() == null || font == null) && font.getIndex() == getCurrentFont().getIndex();
        }

        String toHtml(){
            if (getCharList() == null || getCharList().isEmpty()) {
                return "";
            }
            StringBuilder builder = new StringBuilder();

            if (getCurrentFont() != null){
                builder.append("<span ");
                builder.append(String.format("style=\"font-weight:%d;font-style:%s;font-family:'%s'\""
                        ,getCurrentFont().getBoldweight()
                        ,getCurrentFont().getItalic() ? "italic" : "normal"
                        ,getCurrentFont().getFontName()));
                builder.append(">");
            }
            builder.append(Arrays.toString(getCharList().toArray()));
            if (getCurrentFont() != null) {
                builder.append("</span>");
            }
            return builder.toString();
        }
    }

    public String convert(Cell cell) {
        RichTextString richStr;
        try{
            richStr = cell.getRichStringCellValue();
        } catch (Exception ex) {
            return null;
        }
        StringBuilder builder = new StringBuilder();
        builder.append("<p>");
        RichStyleString styleChars = new RichStyleString();
        for (int i = 0; i < richStr.length(); i++) {
            char chara = richStr.getString().charAt(i);
            if (chara == '\r') {
                continue;
            }
            if (chara == '\n') {
                builder.append("<br/>");
                continue;
            }
            Font font;
            try {
                int fontIndex = richStr.getIndexOfFormattingRun(i);
                if (fontIndex >= 0) {
                    font = cell.getSheet().getWorkbook().getFontAt((short) fontIndex);
                } else {
                    font = null;
                }
            } catch (Exception ex) {
                font = null;
            }
            if (font == null) {
                builder.append(chara);
                continue;
            }
            if (styleChars.isCurrentStyle(font)) {
                styleChars.getCharList().add(chara);
            } else {
                builder.append(styleChars.toHtml());
                styleChars.setCurrentFont(font);
                styleChars.getCharList().clear();
                styleChars.getCharList().add(chara);
            }
        }
        if (styleChars.getCharList().size() > 0) {
            builder.append(styleChars.toHtml());
        }
        builder.append("</p>");
        return builder.toString();
    }
}
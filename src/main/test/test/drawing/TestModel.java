package test.drawing;

import org.apache.poi.objects.*;

@POIObject(headerRowIndex = 1, startIndex = 2)
public class TestModel {
    private int num;
    private String name;

    @Cell
    @HeaderCell(textColor = "#FF0000", foregroundColor = "#000000", fontFamily = "微软雅黑", italic = true, fontSize = 30, fontWeight = 600, columnWidth = 30)
    @POIColumn(index = 0, name = "number")
    @AlternateCell(foregroundColor = "#000000", textColor = "#FFFFFF")
    public int getNum() {
        return num;
    }

    public void setNum(int num) {
        this.num = num;
    }

    @Cell(foregroundColor = "#000000", textColor = "#FFFFFF")
    @AlternateCell(foregroundColor = "#FFFFFF", textColor = "#000000")
    @HeaderCell(textColor = "#FF0000", foregroundColor = "#000000", fontFamily = "微软雅黑", italic = true, fontSize = 30, fontWeight = 600, columnWidth = 30)
    @POIColumn(index = 1, name = "name")
    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}
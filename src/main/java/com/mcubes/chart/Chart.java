package com.mcubes.chart;

import org.apache.poi.xssf.usermodel.XSSFSheet;

public abstract class Chart {

    private String title;
    private int startColumn;
    private int startRow;
    private int colspan;
    private int rowspan;

    public Chart(String title, int startColumn, int startRow, int colspan, int rowspan) {
        this.title = title;
        this.startColumn = startColumn;
        this.startRow = startRow;
        this.colspan = colspan;
        this.rowspan = rowspan;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public int getStartColumn() {
        return startColumn;
    }

    public void setStartColumn(int startColumn) {
        this.startColumn = startColumn;
    }

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }

    public int getColspan() {
        return colspan;
    }

    public void setColspan(int colspan) {
        this.colspan = colspan;
    }

    public int getRowspan() {
        return rowspan;
    }

    public void setRowspan(int rowspan) {
        this.rowspan = rowspan;
    }

    public abstract  void draw(XSSFSheet sheet);
}
